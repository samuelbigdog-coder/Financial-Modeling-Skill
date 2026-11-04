"""Independent verifier for a filled, recalculated DCF template.

It reads only the input cells (Inputs, Hist Financials inputs, Cases, Comps),
rebuilds every headline result in plain Python, and compares it with the
workbook's calculated values. It never reads a workbook total as an input.

Usage:
    python verify_dcf.py Model.xlsx
    python verify_dcf.py Model.xlsx --recalc "python /path/to/recalc.py {file}"   # adds mutation tests

The workbook must already be recalculated (Excel, or LibreOffice headless).
Exit code 0 means every comparison and package check passed.
"""
import argparse
import calendar
import datetime as dt
import shlex
import shutil
import subprocess
import sys
import tempfile
import zipfile
from statistics import mean, median

from openpyxl import load_workbook

from layout import (CASE_COLS, CASE_NAMES, DCF_ROWS, DRIVERS, FC_COLS, HIST_COLS, HIST_ROWS, IN_ROW,
                    MODEL_ROWS, N_FC, SHEETS, WACC_ROWS, case_row)

TOL = 1e-6


def num(v):
    return 0.0 if v is None or v == "" else float(v)


def edate(d, months):
    m = d.month - 1 + months
    y, m = d.year + m // 12, m % 12 + 1
    return d.replace(year=y, month=m, day=min(d.day, calendar.monthrange(y, m)[1]))


def quartile_inc(xs, q):
    xs = sorted(xs)
    if not xs:
        return 0.0
    pos = (len(xs) - 1) * q
    lo = int(pos)
    hi = min(lo + 1, len(xs) - 1)
    return xs[lo] + (xs[hi] - xs[lo]) * (pos - lo)


# ---------------------------------------------------------------- read inputs only
def read_inputs(wb):
    ws = wb["Inputs"]
    inp = {k: ws[f"C{r}"].value for k, r in IN_ROW.items()}
    hs = wb["Hist Financials"]
    hist = {k: [num(hs[f"{c}{HIST_ROWS[k]}"].value) for c in HIST_COLS]
            for k in ("revenue", "cogs", "sga", "rnd", "other_opex", "op_income", "da", "sbc", "nonrecurring",
                      "pretax", "tax", "net_income", "ar", "inventory", "other_ca", "ap", "accrued", "def_rev",
                      "ppe", "gw_intang", "capex", "cfo")}
    cs = wb["Cases"]
    cases = {c: {d: [num(cs[f"{col}{case_row(c, d)}"].value) for col in CASE_COLS] for d, _ in DRIVERS}
             for c in CASE_NAMES}
    cp = wb["Comps"]
    comps = []
    for r in range(6, 18):
        comps.append({k: num(cp[f"{col}{r}"].value) for col, k in zip("DEFHJKLM", ["inc", "price", "shares", "claims",
                                                                                  "rev", "ebitda", "ebit", "ni"])})
    return inp, hist, cases, comps


# ---------------------------------------------------------------- the oracle
class Oracle:
    def __init__(self, inp, hist, cases, comps):
        self.i, self.h, self.cases, self.comps = inp, hist, cases, comps
        g = lambda k: num(inp[k])
        self.g, self.ronic, self.mult = g("g"), g("ronic"), g("exit_mult")
        self.method, self.mid, self.sbc_cash = int(g("tv_method")), int(g("midyear")), g("sbc_cash")
        self.case = int(g("case"))
        h = hist
        self.adj_ebit = [o + n for o, n in zip(h["op_income"], h["nonrecurring"])]
        self.nwc_h = [a + b + c - d - e - f for a, b, c, d, e, f in
                      zip(h["ar"], h["inventory"], h["other_ca"], h["ap"], h["accrued"], h["def_rev"])]
        self.rev0, self.nwc0 = h["revenue"][-1], self.nwc_h[-1]
        # dates and discounting
        fye, val = inp["last_fye"], inp["valuation_date"]
        self.periods, self.stub = [], 0.0
        if isinstance(fye, dt.datetime) and isinstance(val, dt.datetime):
            fye1 = edate(fye, 12)
            span = (fye1 - fye).days
            self.stub = 0.0 if span == 0 else max(0.0, min(1.0, (fye1 - val).days / span))
        for t in range(1, N_FC + 1):
            if t == 1:
                self.periods.append(self.stub / 2 if self.mid == 1 else self.stub)
            else:
                self.periods.append(self.stub + (t - 2) + (0.5 if self.mid == 1 else 1.0))
        self.tv_period = self.periods[-1] + (0.5 if self.mid == 1 else 0.0)
        # shares and bridge
        price, strike = g("price"), g("strike")
        opt = g("options") * (1 - strike / price) if price > 0 and price > strike else 0.0
        self.diluted = g("basic") + g("rsus") + opt + g("other_dil")
        self.bridge = (g("cash") + g("st_inv") + g("eq_inv") - g("debt") - g("leases") * g("inc_leases")
                       - g("preferred") - g("minority") - g("other_dl"))
        self.price = price
        self.wacc = self._wacc()

    def _wacc(self):
        g = lambda k: num(self.i[k])
        rf, erp, bu, tax = g("rf"), g("erp"), g("beta_u"), g("tax_m")
        ksu = rf + bu * erp
        kd = g("kd_override") if g("kd_override") > 0 else rf + g("spread")
        kd_at = kd * (1 - tax)
        mkt = self.price * self.diluted
        debt = g("debt") + g("leases") * g("inc_leases") + g("preferred")
        de_o = self.i["de_override"]
        de_m = float(de_o) if isinstance(de_o, (int, float)) else (0.0 if mkt <= 0 else debt / mkt)
        be = g("book_equity")
        de_b = de_m if be <= 0 else debt / be

        def w(de):
            ksl = ksu + de * (1 - tax) * (ksu - kd)
            wd = de / (1 + de)
            return (1 - wd) * ksl + wd * kd_at
        self.wacc_m, self.wacc_b = w(de_m), w(de_b)
        return self.wacc_m if int(g("wacc_method")) == 1 else (self.wacc_m + self.wacc_b) / 2

    def forecast(self, drv, gshift=0.0, mshift=0.0):
        rev, nwc_prev, out = self.rev0, self.nwc0, []
        for t in range(N_FC):
            rev = rev * (1 + drv["growth"][t] + gshift)
            m = drv["ebit_m"][t] + mshift
            ebit = rev * m
            nopat = ebit - max(0.0, ebit * drv["tax"][t])
            nwc = rev * drv["nwc"][t]
            sbc_ded = -rev * drv["sbc"][t] * self.sbc_cash
            fcf = nopat + rev * drv["da"][t] - rev * drv["capex"][t] - (nwc - nwc_prev) + sbc_ded
            out.append({"rev": rev, "ebit": ebit, "nopat": nopat, "ebitda": ebit + rev * drv["da"][t],
                        "sbc_ded": sbc_ded, "fcf": fcf, "nwc": nwc})
            nwc_prev = nwc
        return out

    def tv(self, fc, w, g, mult, method):
        last = fc[-1]
        if method == 1:
            if w - g <= 0:
                return 0.0
            reinv = 0.0 if self.ronic == 0 else g / self.ronic
            return (last["nopat"] * (1 + g) * (1 - reinv) + last["sbc_ded"] * (1 + g)) / (w - g)
        if method == 2:
            return 0.0 if w - g <= 0 else last["fcf"] * (1 + g) / (w - g)
        return last["ebitda"] * mult

    def value(self, fc, w=None, g=None, mult=None, method=None):
        w = self.wacc if w is None else w
        g = self.g if g is None else g
        mult = self.mult if mult is None else mult
        method = self.method if method is None else method
        cfs = [f["fcf"] for f in fc]
        cfs[0] *= self.stub
        pv = sum(cf / (1 + w) ** p for cf, p in zip(cfs, self.periods))
        tv = self.tv(fc, w, g, mult, method)
        pv_tv = tv / (1 + w) ** self.tv_period
        ev = pv + pv_tv
        vps = 0.0 if self.diluted == 0 else (ev + self.bridge) / self.diluted
        return {"pv": pv, "tv": tv, "pv_tv": pv_tv, "ev": ev, "vps": vps,
                "tv_share": 0.0 if ev == 0 else pv_tv / ev}

    def active(self):
        return self.cases[CASE_NAMES[self.case - 1]]

    @staticmethod
    def interpolate(xs, ys, target, ascending=True):
        pairs = list(zip(xs, ys))
        for (x0, y0), (x1, y1) in zip(pairs, pairs[1:]):
            lo, hi = (y0, y1) if ascending else (y1, y0)
            if lo <= target <= hi and y1 != y0:
                return x0 + (target - y0) / (y1 - y0) * (x1 - x0)
        return None


# ---------------------------------------------------------------- comparisons
def compare(wb, o, results):
    def chk(name, sheet, ref, expected, tol=TOL):
        got = wb[sheet][ref].value
        if isinstance(expected, str) or expected is None:
            ok = got == expected
        else:
            g = num(got) if isinstance(got, (int, float)) else None
            ok = g is not None and abs(g - expected) <= tol * max(1.0, abs(expected))
        results.append((ok, f"{name}: workbook {got!r} vs oracle {expected!r}"))

    act = o.forecast(o.active())
    base = o.value(act)
    chk("WACC (market weights)", "WACC", f"C{WACC_ROWS['wacc_m']}", o.wacc_m)
    chk("WACC (book weights)", "WACC", f"C{WACC_ROWS['wacc_b']}", o.wacc_b)
    chk("WACC used", "WACC", f"C{WACC_ROWS['wacc']}", o.wacc)
    chk("Stub fraction", "DCF", f"I{DCF_ROWS['stub']}", o.stub)
    for t, col in enumerate(FC_COLS):
        chk(f"Revenue {col}", "Model", f"{col}{MODEL_ROWS['revenue']}", act[t]["rev"])
        chk(f"Unlevered FCF {col}", "Model", f"{col}{MODEL_ROWS['ufcf']}", act[t]["fcf"])
        chk(f"Discount period {col}", "DCF", f"{col}{DCF_ROWS['period']}", o.periods[t])
    for k, key in (("PV of forecast", "pv"), ("Terminal value used", "tv"), ("PV of terminal value", "pv_tv"),
                   ("Enterprise value", "ev"), ("TV share", "tv_share"), ("Value per share", "vps")):
        row = {"pv": "sum_pv", "tv": "tv", "pv_tv": "pv_tv", "ev": "ev", "tv_share": "tv_share", "vps": "vps"}[key]
        chk(k, "DCF", f"C{DCF_ROWS[row]}", base[key])
    chk("Diluted shares", "DCF", f"C{DCF_ROWS['diluted']}", o.diluted)
    chk("Equity value", "DCF", f"C{DCF_ROWS['equity']}", base["ev"] + o.bridge)
    for m_, row in ((1, "tv_vd"), (2, "tv_gordon"), (3, "tv_exit")):
        chk(f"TV method {m_}", "DCF", f"C{DCF_ROWS[row]}", o.tv(act, o.wacc, o.g, o.mult, m_))
    scen = {}
    for i, c in enumerate(CASE_NAMES):
        scen[c] = o.value(o.forecast(o.cases[c]))["vps"]
        chk(f"Scenario {c}", "Scenarios", f"C{6+i}", scen[c])
    # reverse DCF grids
    ws = wb["Reverse DCF"]
    for top, key, step in ((6, "g", num(o.i["step_gshift"])), (33, "m", num(o.i["step_mshift"]))):
        xs, ys = [], []
        for k in range(-10, 11):
            r = top + 11 + k
            shift = k * step
            fc = o.forecast(o.active(), gshift=shift if key == "g" else 0.0, mshift=shift if key == "m" else 0.0)
            v = o.value(fc)["vps"]
            xs.append(shift); ys.append(v)
            chk(f"Reverse {key} shift {k:+d}", "Reverse DCF", f"C{r}", v)
        imp = o.interpolate(xs, ys, o.price, ascending=True)
        got = ws[f"C{top + 24}"].value
        if imp is None:
            results.append((got == "Outside grid", f"Reverse {key} implied: workbook {got!r} vs oracle outside grid"))
        else:
            chk(f"Reverse {key} implied shift", "Reverse DCF", f"C{top + 24}", imp)
    xs, ys = [], []
    for k in range(-12, 13):
        w = o.wacc + k * num(o.i["step_wacc"])
        v = "n/a" if (o.method < 3 and w <= o.g) else o.value(act, w=w)["vps"]
        xs.append(w); ys.append(v)
        chk(f"Implied-rate grid {k:+d}", "Reverse DCF", f"C{73 + k}", v)
    num_pairs = [(x, y) for x, y in zip(xs, ys) if isinstance(y, float)]
    imp = o.interpolate([p[0] for p in num_pairs], [p[1] for p in num_pairs], o.price, ascending=False)
    got = ws["C88"].value
    if imp is None:
        results.append((got == "Outside grid", f"Implied discount rate: workbook {got!r} vs oracle outside grid"))
    else:
        chk("Market-implied discount rate", "Reverse DCF", "C88", imp)
    # sensitivity grids
    meth_pg = 1 if o.method == 3 else o.method
    for i, kw in enumerate(range(-3, 4)):
        w = o.wacc + kw * num(o.i["step_wacc"])
        for j, kg in enumerate(range(-2, 3)):
            g = o.g + kg * num(o.i["step_g"])
            v = "n/a" if w <= g else o.value(act, w=w, g=g, method=meth_pg)["vps"]
            chk(f"Sensitivity WACC x g [{kw:+d},{kg:+d}]", "Sensitivity", f"{'CDEFG'[j]}{7+i}", v)
            mu = o.mult + kg * num(o.i["step_mult"])
            chk(f"Sensitivity WACC x multiple [{kw:+d},{kg:+d}]", "Sensitivity", f"{'CDEFG'[j]}{18+i}",
                o.value(act, w=w, mult=mu, method=3)["vps"])
    # comps
    used = {k: [] for k in ("rev", "ebitda", "ebit", "pe")}
    for p in o.comps:
        mcap = p["price"] * p["shares"]
        ev = mcap + p["claims"]
        for k, denom, numer in (("rev", p["rev"], ev), ("ebitda", p["ebitda"], ev), ("ebit", p["ebit"], ev),
                                ("pe", p["ni"], mcap)):
            mult = 0.0 if (denom <= 0 or numer <= 0) else numer / denom
            if p["inc"] == 1 and mult > 0:
                used[k].append(mult)
    for col, k in zip("PQRS", ("rev", "ebitda", "ebit", "pe")):
        chk(f"Comps median {k}", "Comps", f"{col}19", median(used[k]) if used[k] else 0.0)
        chk(f"Comps 25th {k}", "Comps", f"{col}20", quartile_inc(used[k], 0.25))
        chk(f"Comps 75th {k}", "Comps", f"{col}21", quartile_inc(used[k], 0.75))
        chk(f"Comps mean {k}", "Comps", f"{col}22", mean(used[k]) if used[k] else 0.0)
    metric = {"rev": act[0]["rev"], "ebitda": act[0]["ebitda"], "ebit": act[0]["ebit"]}
    for r, k in zip((28, 29, 30), ("rev", "ebitda", "ebit")):
        med = median(used[k]) if used[k] else 0.0
        chk(f"Comps implied per share at median {k}", "Comps", f"H{r}",
            0.0 if o.diluted == 0 else (metric[k] * med + o.bridge) / o.diluted)
    return scen, base, act


def package_checks(path, wb, results):
    results.append((wb.sheetnames == SHEETS, f"Sheet order: {wb.sheetnames}"))
    merged = sum(len(ws.merged_cells.ranges) for ws in wb.worksheets)
    results.append((merged == 0, f"Merged ranges: {merged}"))
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
    results.append((not any("externalLink" in n for n in names), "No external links"))
    results.append((not any("vbaProject" in n for n in names), "No macros"))
    errors, blanks = [], 0
    wf = load_workbook(path)
    for ws in wb.worksheets:
        fws = wf[ws.title]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#") and c.value.rstrip("!?0/") in (
                        "#DIV", "#REF", "#NAME", "#VALUE", "#N/A", "#NUM", "#NULL"):
                    errors.append(f"{ws.title}!{c.coordinate}")
                f = fws[c.coordinate].value
                if isinstance(f, str) and f.startswith("=") and c.value is None:
                    # Comps helper columns return "" for excluded peers by design.
                    if not (ws.title == "Comps" and c.column_letter in "VWXY"):
                        blanks += 1
    results.append((not errors, f"Formula error values: {errors[:10]}"))
    results.append((blanks == 0, f"Formulas without a calculated value (recalculate first): {blanks}"))
    results.append((wb["Checks"]["D30"].value == "ALL OK", f"Checks sheet overall: {wb['Checks']['D30'].value!r}"))


def direction_tests(o, results):
    act = o.forecast(o.active())
    v0 = o.value(act)["vps"]
    results.append((o.value(act, w=o.wacc + 0.01)["vps"] < v0, "Higher WACC lowers value"))
    results.append((o.value(act, method=2, g=o.g - 0.005)["vps"] < o.value(act, method=2)["vps"],
                    "Lower terminal growth lowers value (perpetuity on FCF)"))
    results.append((o.value(act, method=3, mult=o.mult - 1)["vps"] < o.value(act, method=3)["vps"],
                    "Lower exit multiple lowers value"))
    results.append((o.value(act, w=o.g - 0.001, method=2)["tv"] == 0.0, "WACC at or below g gives no perpetuity value"))
    flat = {k: [0.0] * N_FC if k == "growth" else v for k, v in o.active().items()}
    results.append((o.value(o.forecast(flat))["vps"] < v0 if sum(o.active()["growth"]) > 0 else True,
                    "Zero-growth case is below the active case"))


def mutation_tests(path, cmd, inp, hist, cases, comps, results):
    tmp = tempfile.mkdtemp()
    for label_, edits in (("ERP +1 point", {"erp": num(inp["erp"]) + 0.01}), ("Bear case selected", {"case": 1})):
        p = f"{tmp}/mut.xlsx"
        shutil.copy(path, p)
        wb = load_workbook(p)
        for k, v in edits.items():
            wb["Inputs"][f"C{IN_ROW[k]}"] = v
        wb.save(p)
        subprocess.run(shlex.split(cmd.format(file=p)), check=True, capture_output=True)
        got = load_workbook(p, data_only=True)["DCF"][f"C{DCF_ROWS['vps']}"].value
        inp2 = {**inp, **edits}
        o2 = Oracle(inp2, hist, cases, comps)
        exp = o2.value(o2.forecast(o2.active()))["vps"]
        ok = isinstance(got, (int, float)) and abs(got - exp) <= TOL * max(1, abs(exp))
        results.append((ok, f"Mutation {label_}: workbook {got} vs oracle {exp}"))
    shutil.rmtree(tmp, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--recalc", help='command that recalculates a file in place, e.g. "python recalc.py {file}"')
    a = ap.parse_args()
    wb = load_workbook(a.workbook, data_only=True)
    inp, hist, cases, comps = read_inputs(wb)
    o = Oracle(inp, hist, cases, comps)
    results = []
    scen, base, act = compare(wb, o, results)
    package_checks(a.workbook, wb, results)
    direction_tests(o, results)
    if a.recalc:
        mutation_tests(a.workbook, a.recalc, inp, hist, cases, comps, results)
    fails = [m for ok, m in results if not ok]
    print(f"Checks run: {len(results)}  passed: {len(results) - len(fails)}  failed: {len(fails)}")
    for m in fails[:40]:
        print("FAIL", m)
    print(f"Oracle: WACC {o.wacc:.4%}  EV {base['ev']:,.1f}  value per share {base['vps']:,.2f}  "
          f"Bear {scen['Bear']:,.2f}  Base {scen['Base']:,.2f}  Bull {scen['Bull']:,.2f}")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
