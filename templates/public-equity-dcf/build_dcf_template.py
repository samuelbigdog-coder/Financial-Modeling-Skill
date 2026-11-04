"""Build the public-equity DCF template (and optionally a filled copy).

Usage:
    python build_dcf_template.py Public_Equity_DCF_Template.xlsx
    python build_dcf_template.py Example.xlsx --example        # synthetic company

Every calculated cell is an Excel formula. Recalculate the workbook in Excel or
LibreOffice before reading values (openpyxl writes no cached results).
"""
import argparse
import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from layout import (
    BRIDGE_SUM, CASE_BAND, CASE_COLS, CASE_NAMES, D, DCF_BANDS, DCF_ROWS, DILUTED, DRIVERS,
    FC_COLS, H, HIST_BANDS, HIST_COLS, HIST_INPUT_KEYS, HIST_LABELS, HIST_ROWS, I, IN_ROW,
    INPUT_SECTIONS, M, MODEL_BANDS, MODEL_ROWS, N_FC, SHEETS, WACC, WACC_BANDS, WACC_ROWS,
    case_row,
)

# ---------------------------------------------------------------- style (STARWAR system)
NAVY, BLUE_IN, YELLOW, GREEN, GRAY = "002060", "0000FF", "FFFFCC", "008000", "595959"
FONT = "Calibri"
FMT = {
    "amt": '#,##0.0_);(#,##0.0);"-"_)',
    "pct": '0.0%_);(0.0%);"-"_)',
    "pct2": '0.00%_);(0.00%);"-"_)',
    "mult": '0.0"x"_);(0.0"x");"-"_)',
    "dec2": '0.00_);(0.00);"-"_)',
    "ps": '#,##0.00_);(#,##0.00);"-"_)',
    "int": "0",
    "date": "mm-dd-yy",
    "yearA": '0000"A"',
    "yearE": '0000"E"',
    "text": "@",
}
THIN = Side(style="thin", color="000000")
HAIR = Side(style="thin", color="BFBFBF")


def f_title(ws, text, units):
    ws["B2"] = text
    ws["B2"].font = Font(name=FONT, size=12, bold=True)
    ws["B3"] = units
    ws["B3"].font = Font(name=FONT, size=10, italic=True, color=GRAY)
    ws.sheet_view.showGridLines = False


def f_band(ws, row, c1, c2, text=None):
    fill = PatternFill("solid", fgColor=NAVY)
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
    if text is not None:
        ws.cell(row=row, column=c1, value=text)


def f_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True)
        cell.border = Border(bottom=THIN)
        cell.alignment = Alignment(horizontal="right" if c > c1 else "left")


def put(ws, ref, value, fmt=None, kind="calc", bold=False, align=None):
    cell = ws[ref]
    cell.value = value
    color = {"calc": "000000", "link": GREEN, "input": BLUE_IN, "label": "000000"}[kind]
    cell.font = Font(name=FONT, color=color, bold=bold)
    if kind == "input":
        cell.fill = PatternFill("solid", fgColor=YELLOW)
        cell.border = Border(left=HAIR, right=HAIR, top=HAIR, bottom=HAIR)
    if fmt:
        cell.number_format = FMT.get(fmt, fmt)
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


def label(ws, ref, text, bold=False, indent=0):
    cell = ws[ref]
    cell.value = text
    cell.font = Font(name=FONT, bold=bold)
    if indent:
        cell.alignment = Alignment(indent=indent)
    return cell


def total_rule(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.border = Border(top=THIN)
        color = cell.font.color if cell.font is not None and cell.font.color is not None else None
        rgb = color.rgb if color is not None and isinstance(color.rgb, str) else "FF000000"
        cell.font = Font(name=FONT, bold=True, color=rgb)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- formula helpers
def compact_engine(ws, row, sc, drv, gshift="0", mshift="0"):
    """One row that reruns the full FCFF and valuation for a set of drivers.

    drv(t, key) returns a cell reference for driver `key` in forecast year t (1..10).
    Columns: revenue x10, NWC x10, FCF x10, NOPAT_N, EBITDA_N, SBC_N, TV, PV of TV, EV, value per share.
    Returns a dict of column letters for the summary cells.
    """
    rc = [L(sc + i) for i in range(N_FC)]
    nc = [L(sc + 10 + i) for i in range(N_FC)]
    fc = [L(sc + 20 + i) for i in range(N_FC)]
    c_nopat, c_ebitda, c_sbc, c_tv, c_pvtv, c_ev, c_vps = (L(sc + 30 + i) for i in range(7))
    r = row
    sbc_cash, g, ronic, method, mult = I("sbc_cash"), I("g"), I("ronic"), I("tv_method"), I("exit_mult")
    for t in range(1, N_FC + 1):
        prev_rev = M("revenue", "$H").replace("Model!$H", "Model!$H$") if t == 1 else f"{rc[t-2]}{r}"
        prev_rev = "Model!$H$8" if t == 1 else prev_rev
        prev_nwc = "Model!$H$22" if t == 1 else f"{nc[t-2]}{r}"
        R, N = f"{rc[t-1]}{r}", f"{nc[t-1]}{r}"
        m = f"({drv(t, 'ebit_m')}+{mshift})"
        put(ws, R, f"={prev_rev}*(1+{drv(t, 'growth')}+{gshift})", "amt")
        put(ws, N, f"={R}*{drv(t, 'nwc')}", "amt")
        put(ws, f"{fc[t-1]}{r}",
            f"={R}*{m}-MAX(0,{R}*{m}*{drv(t, 'tax')})+{R}*{drv(t, 'da')}-{R}*{drv(t, 'capex')}"
            f"-({N}-{prev_nwc})-{R}*{drv(t, 'sbc')}*{sbc_cash}", "amt")
    RN = f"{rc[-1]}{r}"
    mN = f"({drv(N_FC, 'ebit_m')}+{mshift})"
    put(ws, f"{c_nopat}{r}", f"={RN}*{mN}-MAX(0,{RN}*{mN}*{drv(N_FC, 'tax')})", "amt")
    put(ws, f"{c_ebitda}{r}", f"={RN}*{mN}+{RN}*{drv(N_FC, 'da')}", "amt")
    put(ws, f"{c_sbc}{r}", f"=-{RN}*{drv(N_FC, 'sbc')}*{sbc_cash}", "amt")
    put(ws, f"{c_tv}{r}",
        f"=CHOOSE({method},IF({WACC}-{g}<=0,0,({c_nopat}{r}*(1+{g})*(1-IF({ronic}=0,0,{g}/{ronic}))"
        f"+{c_sbc}{r}*(1+{g}))/({WACC}-{g})),IF({WACC}-{g}<=0,0,{fc[-1]}{r}*(1+{g})/({WACC}-{g})),"
        f"{c_ebitda}{r}*{mult})", "amt")
    put(ws, f"{c_pvtv}{r}", f"={c_tv}{r}/(1+{WACC})^{D('tv_period')}", "amt")
    put(ws, f"{c_ev}{r}",
        f"={fc[0]}{r}*DCF!$I$8*DCF!$I$11+SUMPRODUCT({fc[1]}{r}:{fc[-1]}{r},DCF!$J$11:$R$11)+{c_pvtv}{r}", "amt")
    put(ws, f"{c_vps}{r}", f"=IF({DILUTED}=0,0,({c_ev}{r}+{BRIDGE_SUM})/{DILUTED})", "ps")
    return {"ev": c_ev, "vps": c_vps, "pvtv": c_pvtv, "start": rc[0], "end": c_vps}


def closed_form_vps(w, g, mult, method):
    """Value per share for a given WACC, terminal growth and exit multiple, using the active-case FCF."""
    ronic = I("ronic")
    pv = f"SUMPRODUCT(DCF!$I$9:$R$9,(1+{w})^(-DCF!$I$10:$R$10))"
    tv = (f"CHOOSE({method},IF({w}-{g}<=0,0,(Model!$R$17*(1+{g})*(1-IF({ronic}=0,0,{g}/{ronic}))"
          f"+Model!$R$26*(1+{g}))/({w}-{g})),IF({w}-{g}<=0,0,Model!$R$27*(1+{g})/({w}-{g})),Model!$R$14*{mult})")
    ev = f"({pv}+{tv}/(1+{w})^{D('tv_period')})"
    return f'=IF(AND({method}<3,{w}<={g}),"n/a",IF({DILUTED}=0,0,({ev}+{BRIDGE_SUM})/{DILUTED}))'


def case_driver(case):
    return lambda t, key: f"Cases!${CASE_COLS[t-1]}${case_row(case, key)}"


# ---------------------------------------------------------------- sheets
def sheet_inputs(ws):
    f_title(ws, "Inputs", "Currency in millions, except per-share data, multiples and rates")
    widths(ws, {"A": 2, "B": 58, "C": 34, "D": 11, "E": 60, "F": 48})
    r = 5
    for title, items in INPUT_SECTIONS:
        r += 1
        f_band(ws, r, 2, 6, title)
        ws.cell(row=r, column=3, value="Value")
        ws.cell(row=r, column=4, value="Unit")
        ws.cell(row=r, column=5, value="Rationale")
        ws.cell(row=r, column=6, value="Source and as-of")
        for key, text, unit, default, fmt, kind in items:
            r += 1
            assert IN_ROW[key] == r, key
            label(ws, f"B{r}", text)
            fm = {"pct": "pct", "pct2": "pct2", "mult": "mult", "dec2": "dec2", "ps": "ps", "amt": "amt",
                  "0": "int", "@": "text", "mm-dd-yy": "date"}[fmt]
            put(ws, f"C{r}", default, fm, kind="input", align="left" if kind == "text" else "right")
            put(ws, f"D{r}", unit, kind="label")
            for col in "EF":
                c = ws[f"{col}{r}"]
                c.font = Font(name=FONT, color=BLUE_IN)
                c.alignment = Alignment(wrap_text=False)
        r += 1
    dv_case = DataValidation(type="whole", operator="between", formula1="1", formula2="3", allow_blank=False)
    dv_bin = DataValidation(type="whole", operator="between", formula1="0", formula2="1", allow_blank=False)
    dv_m = DataValidation(type="whole", operator="between", formula1="1", formula2="2", allow_blank=False)
    ws.add_data_validation(dv_case)
    ws.add_data_validation(dv_bin)
    ws.add_data_validation(dv_m)
    dv_case.add(f"C{IN_ROW['case']}")
    dv_case.add(f"C{IN_ROW['tv_method']}")
    for k in ("midyear", "sbc_cash", "inc_leases"):
        dv_bin.add(f"C{IN_ROW[k]}")
    dv_m.add(f"C{IN_ROW['wacc_method']}")
    ws.sheet_view.zoomScale = 85


def sheet_cases(ws):
    f_title(ws, "Cases", "Forecast drivers by case; the Active block feeds the Model through the case selector")
    widths(ws, {"A": 2, "B": 34, "C": 12, **{c: 10 for c in CASE_COLS}})
    label(ws, "B5", "Driver", bold=True)
    label(ws, "C5", "Last actual", bold=True)
    for t, col in enumerate(CASE_COLS, start=1):
        put(ws, f"{col}5", f"=Model!{FC_COLS[t-1]}5", "yearE", kind="link", bold=True)
    f_header(ws, 5, 2, 3 + N_FC)
    last = {"growth": "r_growth", "ebit_m": "r_ebit_m", "da": "r_da", "capex": "r_capex",
            "nwc": "r_nwc", "sbc": "r_sbc", "tax": "r_tax"}
    for case in CASE_NAMES + ["Active"]:
        band = CASE_BAND[case]
        name = f"{case} case" if case != "Active" else "Active case (selected on Inputs)"
        f_band(ws, band, 2, 3 + N_FC, name)
        for key, text in DRIVERS:
            r = case_row(case, key)
            label(ws, f"B{r}", text, indent=1)
            put(ws, f"C{r}", f"={H(last[key], 'H')}", "pct", kind="link")
            for t, col in enumerate(CASE_COLS, start=1):
                if case == "Active":
                    put(ws, f"{col}{r}",
                        f"=CHOOSE({I('case')},{col}{case_row('Bear', key)},{col}{case_row('Base', key)},"
                        f"{col}{case_row('Bull', key)})", "pct")
                else:
                    put(ws, f"{col}{r}", None, "pct", kind="input")
    ws.sheet_view.zoomScale = 85


def sheet_hist(ws):
    f_title(ws, "Hist Financials", "Reported figures in millions as filed (signs: expenses and outflows positive); see Source Register")
    widths(ws, {"A": 2, "B": 52, "C": 2, **{c: 12 for c in HIST_COLS}, "I": 11})
    for k, r in HIST_ROWS.items():
        if k in ("fy", "period_end"):
            label(ws, f"B{r}", HIST_LABELS[k], bold=True)
            for col in HIST_COLS:
                put(ws, f"{col}{r}", None, "int" if k == "fy" else "date", kind="input", align="right")
            continue
    f_header(ws, 7, 2, 9)
    for row, text in HIST_BANDS.items():
        f_band(ws, row, 2, 9, text)
    ws["I52"] = "3Y avg"
    for k, r in HIST_ROWS.items():
        if k in ("fy", "period_end"):
            continue
        label(ws, f"B{r}", HIST_LABELS[k], indent=0 if k in ("gross_profit", "adj_ebit", "op_nwc", "fcf") else 1)
        for i, col in enumerate(HIST_COLS):
            prev = HIST_COLS[i - 1] if i else None
            if k in HIST_INPUT_KEYS:
                put(ws, f"{col}{r}", None, "amt", kind="input")
                continue
            hr = lambda key, c=col: f"{c}{HIST_ROWS[key]}"
            formula, fmt = None, "pct"
            if k == "gross_profit":
                formula, fmt = f"={hr('revenue')}-{hr('cogs')}", "amt"
            elif k == "op_residual":
                formula, fmt = f"={hr('op_income')}-({hr('gross_profit')}-{hr('sga')}-{hr('rnd')}-{hr('other_opex')})", "amt"
            elif k == "adj_ebit":
                formula, fmt = f"={hr('op_income')}+{hr('nonrecurring')}", "amt"
            elif k == "op_nwc":
                formula, fmt = (f"={hr('ar')}+{hr('inventory')}+{hr('other_ca')}-{hr('ap')}-{hr('accrued')}"
                                f"-{hr('def_rev')}"), "amt"
            elif k == "fcf":
                formula, fmt = f"={hr('cfo')}-{hr('capex')}", "amt"
            elif k == "r_growth":
                formula = None if prev is None else f"=IF({prev}{HIST_ROWS['revenue']}=0,0,{hr('revenue')}/{prev}{HIST_ROWS['revenue']}-1)"
            elif k == "r_gm":
                formula = f"=IF({hr('revenue')}=0,0,{hr('gross_profit')}/{hr('revenue')})"
            elif k == "r_ebit_m":
                formula = f"=IF({hr('revenue')}=0,0,{hr('adj_ebit')}/{hr('revenue')})"
            elif k == "r_da":
                formula = f"=IF({hr('revenue')}=0,0,{hr('da')}/{hr('revenue')})"
            elif k == "r_capex":
                formula = f"=IF({hr('revenue')}=0,0,{hr('capex')}/{hr('revenue')})"
            elif k == "r_nwc":
                formula = f"=IF({hr('revenue')}=0,0,{hr('op_nwc')}/{hr('revenue')})"
            elif k == "r_sbc":
                formula = f"=IF({hr('revenue')}=0,0,{hr('sbc')}/{hr('revenue')})"
            elif k == "r_tax":
                formula = f"=IF({hr('pretax')}<=0,0,{hr('tax')}/{hr('pretax')})"
            elif k == "r_roic":
                ic = lambda c: f"({c}{HIST_ROWS['op_nwc']}+{c}{HIST_ROWS['ppe']}+{c}{HIST_ROWS['gw_intang']})"
                avg = ic(col) if prev is None else f"(({ic(col)}+{ic(prev)})/2)"
                formula = f"=IF({avg}=0,0,{hr('adj_ebit')}*(1-{I('tax_m')})/{avg})"
            elif k == "r_fcf_conv":
                formula = f"=IF({hr('net_income')}=0,0,{hr('fcf')}/{hr('net_income')})"
            if formula:
                put(ws, f"{col}{r}", formula, fmt)
            elif fmt:
                ws[f"{col}{r}"].number_format = FMT[fmt]
        if k.startswith("r_"):
            put(ws, f"I{r}", f"=IFERROR(AVERAGE(F{r}:H{r}),0)", "pct")
    for k in ("gross_profit", "adj_ebit", "op_nwc", "fcf"):
        total_rule(ws, HIST_ROWS[k], 2, 8)
    ws.sheet_view.zoomScale = 85


def sheet_model(ws):
    f_title(ws, "Model", "Currency in millions; historical columns link to Hist Financials, forecast columns run on the Active case")
    widths(ws, {"A": 2, "B": 38, "C": 2, **{c: 11 for c in HIST_COLS + FC_COLS}})
    label(ws, "B5", "Fiscal year", bold=True)
    for i, col in enumerate(HIST_COLS):
        put(ws, f"{col}5", f"={H('fy', col)}", "yearA", kind="link", bold=True)
    for t, col in enumerate(FC_COLS, start=1):
        prev = "H" if t == 1 else FC_COLS[t - 2]
        put(ws, f"{col}5", f"={prev}5+1", "yearE", bold=True)
    f_header(ws, 5, 2, 18)
    for row, text in MODEL_BANDS.items():
        f_band(ws, row, 2, 18, text)
    labels = {
        "revenue": "Revenue", "growth": "Revenue growth", "ebit_m": "Adjusted EBIT margin",
        "adj_ebit": "Adjusted EBIT", "da_pct": "D&A % of revenue", "da": "Depreciation and amortization",
        "ebitda": "Adjusted EBITDA", "tax_rate": "Cash tax rate on EBIT", "taxes": "Taxes on EBIT",
        "nopat": "NOPAT", "capex_pct": "Capex % of revenue", "capex": "Capital expenditures",
        "nwc_pct": "Operating NWC % of revenue", "nwc": "Operating NWC", "d_nwc": "Change in NWC (cash)",
        "sbc_pct": "Stock comp % of revenue", "sbc": "Stock-based compensation",
        "sbc_ded": "Stock comp treated as cash cost", "ufcf": "Unlevered free cash flow",
        "ppe": "Net PP&E", "gw": "Goodwill and intangibles", "ic": "Invested capital",
        "roic": "ROIC", "net_reinv": "Net reinvestment", "reinv_rate": "Reinvestment rate",
        "fund_growth": "Fundamental growth (ROIC x reinvestment rate)",
    }
    pct_rows = {"growth", "ebit_m", "da_pct", "tax_rate", "capex_pct", "nwc_pct", "sbc_pct", "roic",
                "reinv_rate", "fund_growth"}
    drv = {"growth": "growth", "ebit_m": "ebit_m", "da_pct": "da", "capex_pct": "capex", "nwc_pct": "nwc",
           "sbc_pct": "sbc", "tax_rate": "tax"}
    histmap = {"revenue": "revenue", "growth": "r_growth", "ebit_m": "r_ebit_m", "adj_ebit": "adj_ebit",
               "da_pct": "r_da", "da": "da", "tax_rate": "r_tax", "capex_pct": "r_capex",
               "nwc_pct": "r_nwc", "nwc": "op_nwc", "sbc_pct": "r_sbc", "sbc": "sbc", "ppe": "ppe",
               "gw": "gw_intang"}
    for k, r in MODEL_ROWS.items():
        if k == "year":
            continue
        label(ws, f"B{r}", labels[k], bold=k in ("revenue", "nopat", "ufcf", "ebitda"),
              indent=0 if k in ("revenue", "nopat", "ufcf", "ebitda", "ic") else 1)
        cols = HIST_COLS + FC_COLS
        for j, col in enumerate(cols):
            hist = j < len(HIST_COLS)
            prev = cols[j - 1] if j else None
            t = j - len(HIST_COLS) + 1
            mr = lambda key, c=col: f"{c}{MODEL_ROWS[key]}"
            fmt = "pct" if k in pct_rows else "amt"
            kind = "calc"
            if hist and k in histmap:
                if k == "growth" and j == 0:
                    continue
                f = f"={H(histmap[k], col)}"
                if k == "capex":
                    f = f"=-{H('capex', col)}"
                kind = "link"
            elif hist and k == "capex":
                f, kind = f"=-{H('capex', col)}", "link"
            elif not hist and k in drv:
                f, kind = f"=Cases!{CASE_COLS[t-1]}{case_row('Active', drv[k])}", "link"
            elif k == "revenue":
                f = f"={prev}{r}*(1+{mr('growth')})"
            elif k == "adj_ebit":
                f = f"={mr('revenue')}*{mr('ebit_m')}"
            elif k == "da":
                f = f"={mr('revenue')}*{mr('da_pct')}"
            elif k == "ebitda":
                f = f"={mr('adj_ebit')}+{mr('da')}"
            elif k == "taxes":
                f = f"=-MAX(0,{mr('adj_ebit')}*{mr('tax_rate')})"
            elif k == "nopat":
                f = f"={mr('adj_ebit')}+{mr('taxes')}"
            elif k == "capex":
                f = f"=-{mr('revenue')}*{mr('capex_pct')}"
            elif k == "nwc":
                f = f"={mr('revenue')}*{mr('nwc_pct')}"
            elif k == "d_nwc":
                f = "=0" if j == 0 else f"=-({mr('nwc')}-{prev}{MODEL_ROWS['nwc']})"
            elif k == "sbc":
                f = f"={mr('revenue')}*{mr('sbc_pct')}"
            elif k == "sbc_ded":
                f = f"=-{mr('sbc')}*{I('sbc_cash')}"
            elif k == "ufcf":
                f = f"={mr('nopat')}+{mr('da')}+{mr('capex')}+{mr('d_nwc')}+{mr('sbc_ded')}"
            elif k == "ppe":
                f = f"={prev}{r}-{mr('capex')}-{mr('da')}"
            elif k == "gw":
                f = f"={prev}{r}"
            elif k == "ic":
                f = f"={mr('nwc')}+{mr('ppe')}+{mr('gw')}"
            elif k == "roic":
                avg = mr("ic") if j == 0 else f"(({mr('ic')}+{prev}{MODEL_ROWS['ic']})/2)"
                f = f"=IF({avg}=0,0,{mr('nopat')}/{avg})"
            elif k == "net_reinv":
                f = f"=-{mr('capex')}-{mr('da')}-{mr('d_nwc')}"
            elif k == "reinv_rate":
                f = f"=IF({mr('nopat')}<=0,0,{mr('net_reinv')}/{mr('nopat')})"
            elif k == "fund_growth":
                f = f"={mr('roic')}*{mr('reinv_rate')}"
            put(ws, f"{col}{r}", f, fmt, kind=kind)
    for k in ("adj_ebit", "ebitda", "nopat", "ufcf", "ic"):
        total_rule(ws, MODEL_ROWS[k], 2, 18)
    for col in FC_COLS:
        for r in range(5, 36):
            c = ws[f"{col}{r}"]
            if r not in MODEL_BANDS and c.fill.fgColor.rgb in (None, "00000000"):
                c.fill = PatternFill("solid", fgColor="F2F2F2")
    ws.sheet_view.zoomScale = 85


def sheet_wacc(ws):
    f_title(ws, "WACC", "Rates as decimals; course convention averages market-value and book-value weighted WACC")
    widths(ws, {"A": 2, "B": 52, "C": 14})
    for row, text in WACC_BANDS.items():
        f_band(ws, row, 2, 3, text)
    W = lambda k: f"$C${WACC_ROWS[k]}"
    debt = f"({I('debt')}+{I('leases')}*{I('inc_leases')}+{I('preferred')})"
    rows = [
        ("rf", "Risk-free rate", f"={I('rf')}", "pct2", "link"),
        ("erp", "Equity risk premium", f"={I('erp')}", "pct2", "link"),
        ("beta_u", "Unlevered beta", f"={I('beta_u')}", "dec2", "link"),
        ("ksu", "Unlevered cost of equity", f"={W('rf')}+{W('beta_u')}*{W('erp')}", "pct2", "calc"),
        ("kd", "Pre-tax cost of debt", f"=IF({I('kd_override')}>0,{I('kd_override')},{W('rf')}+{I('spread')})", "pct2", "calc"),
        ("tax", "Marginal tax rate", f"={I('tax_m')}", "pct", "link"),
        ("kd_at", "After-tax cost of debt", f"={W('kd')}*(1-{W('tax')})", "pct2", "calc"),
        ("mkt_equity", "Market value of equity", f"={D('mkt_cap')}", "amt", "link"),
        ("debt_used", "Debt and debt-like claims", f"={debt}", "amt", "calc"),
        ("book_equity", "Book equity", f"={I('book_equity')}", "amt", "link"),
        ("de_m", "Debt / equity (market)",
         f"=IF(ISNUMBER({I('de_override')}),{I('de_override')},IF({W('mkt_equity')}<=0,0,{W('debt_used')}/{W('mkt_equity')}))",
         "dec2", "calc"),
        ("de_b", "Debt / equity (book)", f"=IF({W('book_equity')}<=0,{W('de_m')},{W('debt_used')}/{W('book_equity')})", "dec2", "calc"),
        ("ksl_m", "Levered cost of equity", f"={W('ksu')}+{W('de_m')}*(1-{W('tax')})*({W('ksu')}-{W('kd')})", "pct2", "calc"),
        ("wd_m", "Debt weight", f"={W('de_m')}/(1+{W('de_m')})", "pct", "calc"),
        ("we_m", "Equity weight", f"=1-{W('wd_m')}", "pct", "calc"),
        ("wacc_m", "WACC (market weights)", f"={W('we_m')}*{W('ksl_m')}+{W('wd_m')}*{W('kd_at')}", "pct2", "calc"),
        ("ksl_b", "Levered cost of equity", f"={W('ksu')}+{W('de_b')}*(1-{W('tax')})*({W('ksu')}-{W('kd')})", "pct2", "calc"),
        ("wd_b", "Debt weight", f"={W('de_b')}/(1+{W('de_b')})", "pct", "calc"),
        ("we_b", "Equity weight", f"=1-{W('wd_b')}", "pct", "calc"),
        ("wacc_b", "WACC (book weights)", f"={W('we_b')}*{W('ksl_b')}+{W('wd_b')}*{W('kd_at')}", "pct2", "calc"),
        ("wacc", "WACC", f"=IF({I('wacc_method')}=1,{W('wacc_m')},AVERAGE({W('wacc_m')},{W('wacc_b')}))", "pct2", "calc"),
        ("beta_l", "Implied levered beta (market weights)", f"=IF({W('erp')}=0,0,({W('ksl_m')}-{W('rf')})/{W('erp')})", "dec2", "calc"),
        ("spread_g", "WACC less terminal growth", f"={W('wacc')}-{I('g')}", "pct2", "calc"),
    ]
    for key, text, f, fmt, kind in rows:
        r = WACC_ROWS[key]
        label(ws, f"B{r}", text, bold=key in ("wacc", "wacc_m", "wacc_b"))
        put(ws, f"C{r}", f, fmt, kind=kind, bold=key == "wacc")
    total_rule(ws, WACC_ROWS["wacc"], 2, 3)


def sheet_dcf(ws):
    f_title(ws, "DCF", "Currency in millions, except per share; stub period from the valuation date to the next fiscal year end")
    widths(ws, {"A": 2, "B": 50, "C": 14, **{c: 2 for c in "DEFGH"}, **{c: 11 for c in FC_COLS}})
    label(ws, "B5", "Fiscal year", bold=True)
    for col in FC_COLS:
        put(ws, f"{col}5", f"=Model!{col}5", "yearE", kind="link", bold=True)
    f_header(ws, 5, 2, 18)
    rows = {"ufcf": "Unlevered free cash flow", "period_end": "Period end", "stub": "Share of year in valuation window",
            "cf": "Cash flow counted", "period": "Discount period (years)", "df": "Discount factor", "pv": "Present value"}
    for k, text in rows.items():
        label(ws, f"B{DCF_ROWS[k]}", text, bold=k == "pv")
    mid = I("midyear")
    for t, col in enumerate(FC_COLS, start=1):
        prev = FC_COLS[t - 2] if t > 1 else None
        R = lambda k, c=col: f"{c}{DCF_ROWS[k]}"
        put(ws, R("ufcf"), f"=Model!{col}{MODEL_ROWS['ufcf']}", "amt", kind="link")
        put(ws, R("period_end"), f"=EDATE({I('last_fye')},12)" if t == 1 else f"=EDATE({prev}{DCF_ROWS['period_end']},12)", "date")
        if t == 1:
            put(ws, R("stub"), f"=IF({R('period_end')}-{I('last_fye')}=0,0,MAX(0,MIN(1,({R('period_end')}-{I('valuation_date')})/({R('period_end')}-{I('last_fye')}))))", "pct")
            put(ws, R("period"), f"=IF({mid}=1,{R('stub')}/2,{R('stub')})", "dec2")
        else:
            put(ws, R("stub"), "=1", "pct")
            put(ws, R("period"), f"=$I${DCF_ROWS['stub']}+{t-2}+IF({mid}=1,0.5,1)", "dec2")
        put(ws, R("cf"), f"={R('ufcf')}*{R('stub')}", "amt")
        put(ws, R("df"), f"=1/(1+{WACC})^{R('period')}", "dec2")
        ws[R("df")].number_format = "0.0000"
        put(ws, R("pv"), f"={R('cf')}*{R('df')}", "amt")
    total_rule(ws, DCF_ROWS["pv"], 9, 18)
    for row, text in DCF_BANDS.items():
        f_band(ws, row, 2, 3, text)
    g, ronic = I("g"), I("ronic")
    S = lambda k: f"C{DCF_ROWS[k]}"
    summary = [
        ("sum_pv", "PV of forecast cash flows", f"=SUM(I{DCF_ROWS['pv']}:R{DCF_ROWS['pv']})", "amt"),
        ("nopat_n1", "Terminal-year NOPAT (year N+1)", f"=Model!R{MODEL_ROWS['nopat']}*(1+{g})", "amt"),
        ("sbc_n1", "Terminal-year stock comp cost (year N+1)", f"=Model!R{MODEL_ROWS['sbc_ded']}*(1+{g})", "amt"),
        ("fcf_vd", "Terminal FCF, value driver (NOPAT x (1 - g/RONIC))",
         f"={S('nopat_n1')}*(1-IF({ronic}=0,0,{g}/{ronic}))+{S('sbc_n1')}", "amt"),
        ("tv_vd", "Terminal value: value driver", f"=IF({WACC}-{g}<=0,0,{S('fcf_vd')}/({WACC}-{g}))", "amt"),
        ("tv_gordon", "Terminal value: perpetuity on final-year FCF",
         f"=IF({WACC}-{g}<=0,0,R{DCF_ROWS['ufcf']}*(1+{g})/({WACC}-{g}))", "amt"),
        ("tv_exit", "Terminal value: exit EV/EBITDA", f"=Model!R{MODEL_ROWS['ebitda']}*{I('exit_mult')}", "amt"),
        ("tv", "Terminal value used", f"=CHOOSE({I('tv_method')},{S('tv_vd')},{S('tv_gordon')},{S('tv_exit')})", "amt"),
        ("tv_period", "Terminal discount period (years)", f"=R{DCF_ROWS['period']}+IF({I('midyear')}=1,0.5,0)", "dec2"),
        ("pv_tv", "PV of terminal value", f"={S('tv')}/(1+{WACC})^{S('tv_period')}", "amt"),
        ("ev", "Enterprise value", f"={S('sum_pv')}+{S('pv_tv')}", "amt"),
        ("tv_share", "Terminal value share of EV", f"=IF({S('ev')}=0,0,{S('pv_tv')}/{S('ev')})", "pct"),
        ("implied_mult", "Implied exit EV/EBITDA of terminal value used",
         f"=IF(Model!R{MODEL_ROWS['ebitda']}=0,0,{S('tv')}/Model!R{MODEL_ROWS['ebitda']})", "mult"),
        ("implied_g", "Implied perpetual growth in exit-multiple value",
         f"=IF({S('tv_exit')}+R{DCF_ROWS['ufcf']}=0,0,({S('tv_exit')}*{WACC}-R{DCF_ROWS['ufcf']})/({S('tv_exit')}+R{DCF_ROWS['ufcf']}))", "pct"),
        ("b_ev", "Enterprise value", f"={S('ev')}", "amt"),
        ("b_cash", "Plus cash and equivalents", f"={I('cash')}", "amt"),
        ("b_stinv", "Plus short-term investments", f"={I('st_inv')}", "amt"),
        ("b_eqinv", "Plus non-operating equity investments", f"={I('eq_inv')}", "amt"),
        ("b_debt", "Less total debt", f"=-{I('debt')}", "amt"),
        ("b_leases", "Less operating leases (if treated as debt)", f"=-{I('leases')}*{I('inc_leases')}", "amt"),
        ("b_pref", "Less preferred equity", f"=-{I('preferred')}", "amt"),
        ("b_minority", "Less noncontrolling interests", f"=-{I('minority')}", "amt"),
        ("b_other", "Less other debt-like items", f"=-{I('other_dl')}", "amt"),
        ("equity", "Equity value", f"=SUM(C{DCF_ROWS['b_ev']}:C{DCF_ROWS['b_other']})", "amt"),
        ("s_basic", "Basic shares", f"={I('basic')}", "amt"),
        ("s_rsus", "Unvested RSUs and PSUs", f"={I('rsus')}", "amt"),
        ("s_options", "Option shares (treasury method)",
         f"=IF(AND({I('price')}>0,{I('price')}>{I('strike')}),{I('options')}*(1-{I('strike')}/{I('price')}),0)", "amt"),
        ("s_other", "Other dilutive shares", f"={I('other_dil')}", "amt"),
        ("diluted", "Diluted shares", f"=SUM(C{DCF_ROWS['s_basic']}:C{DCF_ROWS['s_other']})", "amt"),
        ("vps", "Intrinsic value per share", f"=IF({S('diluted')}=0,0,{S('equity')}/{S('diluted')})", "ps"),
        ("price", "Current share price", f"={I('price')}", "ps"),
        ("upside", "Upside (downside) to current price", f"=IF({S('price')}=0,0,{S('vps')}/{S('price')}-1)", "pct"),
        ("mkt_cap", "Market capitalization", f"={S('price')}*{S('diluted')}", "amt"),
        ("mkt_ev", "Market enterprise value", f"={S('mkt_cap')}-SUM(C{DCF_ROWS['b_cash']}:C{DCF_ROWS['b_other']})", "amt"),
        ("ev_ebitda_dcf", "EV / next-year EBITDA at intrinsic value", f"=IF(Model!I{MODEL_ROWS['ebitda']}=0,0,{S('ev')}/Model!I{MODEL_ROWS['ebitda']})", "mult"),
        ("ev_ebitda_mkt", "EV / next-year EBITDA at market", f"=IF(Model!I{MODEL_ROWS['ebitda']}=0,0,{S('mkt_ev')}/Model!I{MODEL_ROWS['ebitda']})", "mult"),
        ("ev_ebit_dcf", "EV / next-year EBIT at intrinsic value", f"=IF(Model!I{MODEL_ROWS['adj_ebit']}=0,0,{S('ev')}/Model!I{MODEL_ROWS['adj_ebit']})", "mult"),
        ("ev_ebit_mkt", "EV / next-year EBIT at market", f"=IF(Model!I{MODEL_ROWS['adj_ebit']}=0,0,{S('mkt_ev')}/Model!I{MODEL_ROWS['adj_ebit']})", "mult"),
        ("fcf_yield", "Next-year FCF yield on market EV", f"=IF({S('mkt_ev')}=0,0,Model!I{MODEL_ROWS['ufcf']}/{S('mkt_ev')})", "pct"),
    ]
    for key, text, f, fmt in summary:
        r = DCF_ROWS[key]
        label(ws, f"B{r}", text, bold=key in ("ev", "equity", "diluted", "vps"), indent=0 if key in ("ev", "equity", "diluted", "vps", "b_ev") else 1)
        put(ws, f"C{r}", f, fmt, kind="link" if f.startswith("=Inputs") or f.startswith("=-Inputs") or f.startswith("=Model") else "calc",
            bold=key in ("ev", "equity", "vps"))
    for k in ("ev", "equity", "diluted", "vps"):
        total_rule(ws, DCF_ROWS[k], 2, 3)
    ws.sheet_view.zoomScale = 85


def sheet_scenarios(ws):
    f_title(ws, "Scenarios", "Each case rerun through the full FCFF and valuation at the same WACC and terminal settings")
    widths(ws, {"A": 2, "B": 18, "C": 14, "D": 14, "E": 14, "F": 14, "G": 3})
    for col, text in zip("BCDEF", ["Case", "Value per share", "Enterprise value", "vs current price", "TV share of EV"]):
        label(ws, f"{col}5", text, bold=True)
    f_header(ws, 5, 2, 6)
    out = {}
    for i, case in enumerate(CASE_NAMES):
        r = 6 + i
        label(ws, f"B{r}", case)
        cols = compact_engine(ws, r, 8, case_driver(case))
        put(ws, f"C{r}", f"={cols['vps']}{r}", "ps", bold=True)
        put(ws, f"D{r}", f"={cols['ev']}{r}", "amt")
        put(ws, f"E{r}", f"=IF({I('price')}=0,0,C{r}/{I('price')}-1)", "pct")
        put(ws, f"F{r}", f"=IF(D{r}=0,0,{cols['pvtv']}{r}/D{r})", "pct")
        out[case] = r
    f_band(ws, 4, 8, 8 + 36, "Engine: revenue, NWC and FCF by forecast year, then terminal and per-share values")
    hdr = ([f"Revenue Y{t}" for t in range(1, 11)] + [f"NWC Y{t}" for t in range(1, 11)] +
           [f"FCF Y{t}" for t in range(1, 11)] + ["NOPAT N", "EBITDA N", "Stock comp N", "TV", "PV of TV", "EV", "Per share"])
    for j, text in enumerate(hdr):
        c = ws.cell(row=5, column=8 + j, value=text)
        c.font = Font(name=FONT, bold=True)
        c.border = Border(bottom=THIN)
        ws.column_dimensions[L(8 + j)].width = 11
    label(ws, "B9", "Active case", bold=True)
    put(ws, "C9", f"=CHOOSE({I('case')},C6,C7,C8)", "ps", bold=True)
    total_rule(ws, 9, 2, 6)
    return out


def sheet_reverse(ws):
    f_title(ws, "Reverse DCF", "What today's price implies: shifts applied to every forecast year of the Active case")
    widths(ws, {"A": 2, "B": 34, "C": 16, "D": 3, "E": 3})
    price = I("price")
    blocks = {}
    for name, top, shift_step, kind in (("Revenue growth shift", 6, I("step_gshift"), "g"),
                                        ("EBIT margin shift", 33, I("step_mshift"), "m")):
        f_band(ws, top - 1, 2, 3, name)
        label(ws, f"B{top}", "Shift (all years)", bold=True)
        label(ws, f"C{top}", "Value per share", bold=True)
        f_header(ws, top, 2, 3)
        first, last = top + 1, top + 21
        for k in range(-10, 11):
            r = top + 11 + k
            put(ws, f"B{r}", f"={k}*{shift_step}", "pct")
            if kind == "g":
                cols = compact_engine(ws, r, 6, case_driver("Active"), gshift=f"$B{r}")
            else:
                cols = compact_engine(ws, r, 6, case_driver("Active"), mshift=f"$B{r}")
            put(ws, f"C{r}", f"={cols['vps']}{r}", "ps")
        rb = last + 2
        label(ws, f"B{rb}", "Grid bracket")
        put(ws, f"C{rb}", f"=IFERROR(MATCH({price},C{first}:C{last},1),0)", "int")
        label(ws, f"B{rb+1}", f"Implied {name.lower()}", bold=True)
        m = f"C{rb}"
        put(ws, f"C{rb+1}",
            f'=IF(OR({price}<MIN(C{first}:C{last}),{price}>MAX(C{first}:C{last}),{m}=0),"Outside grid",'
            f"IF({m}=21,INDEX(B{first}:B{last},21),INDEX(B{first}:B{last},{m})+({price}-INDEX(C{first}:C{last},{m}))"
            f"/(INDEX(C{first}:C{last},{m}+1)-INDEX(C{first}:C{last},{m}))*(INDEX(B{first}:B{last},{m}+1)-INDEX(B{first}:B{last},{m}))))",
            "pct2", bold=True)
        blocks[kind] = {"first": first, "last": last, "zero": top + 11, "implied": rb + 1}
    # Market-implied discount rate
    top = 60
    f_band(ws, top - 1, 2, 3, "Market-implied discount rate")
    label(ws, f"B{top}", "Discount rate", bold=True)
    label(ws, f"C{top}", "Value per share", bold=True)
    f_header(ws, top, 2, 3)
    first, last = top + 1, top + 25
    method = I("tv_method")
    for k in range(-12, 13):
        r = top + 13 + k
        put(ws, f"B{r}", f"={WACC}+{k}*{I('step_wacc')}", "pct2")
        put(ws, f"C{r}", closed_form_vps(f"$B{r}", I("g"), I("exit_mult"), method), "ps")
    rb = last + 2
    label(ws, f"B{rb}", "Grid bracket")
    put(ws, f"C{rb}", f"=IFERROR(MATCH({price},C{first}:C{last},-1),0)", "int")
    label(ws, f"B{rb+1}", "Market-implied discount rate", bold=True)
    m = f"C{rb}"
    put(ws, f"C{rb+1}",
        f'=IFERROR(IF(OR({m}=0,{m}=25),"Outside grid",INDEX(B{first}:B{last},{m})+({price}-INDEX(C{first}:C{last},{m}))'
        f"/(INDEX(C{first}:C{last},{m}+1)-INDEX(C{first}:C{last},{m}))*(INDEX(B{first}:B{last},{m}+1)-INDEX(B{first}:B{last},{m}))),\"Outside grid\")",
        "pct2", bold=True)
    label(ws, f"B{rb+2}", "Spread to model WACC")
    put(ws, f"C{rb+2}", f'=IF(ISNUMBER(C{rb+1}),C{rb+1}-{WACC},"n/a")', "pct2")
    blocks["r"] = {"first": first, "last": last, "implied": rb + 1}
    f_band(ws, 4, 6, 6 + 36, "Engine")
    return blocks


def sheet_sensitivity(ws):
    f_title(ws, "Sensitivity", "Value per share on the Active case; grid centers equal the live model")
    widths(ws, {"A": 2, "B": 26, **{c: 12 for c in "CDEFGH"}})
    method_pg = f"IF({I('tv_method')}=3,1,{I('tv_method')})"
    grids = {}
    for name, top, colvar, step, meth in (
        ("WACC (rows) x terminal growth (columns), perpetuity method", 5, I("g"), I("step_g"), method_pg),
        ("WACC (rows) x exit EV/EBITDA multiple (columns)", 16, I("exit_mult"), I("step_mult"), "3"),
    ):
        f_band(ws, top, 2, 7, name)
        hr = top + 1
        label(ws, f"B{hr}", "WACC", bold=True)
        for j, k in enumerate(range(-2, 3)):
            col = L(3 + j)
            put(ws, f"{col}{hr}", f"={colvar}+{k}*{step}", "pct" if top == 5 else "mult", bold=True)
        f_header(ws, hr, 2, 7)
        for i, k in enumerate(range(-3, 4)):
            r = hr + 1 + i
            put(ws, f"B{r}", f"={WACC}+{k}*{I('step_wacc')}", "pct2", bold=True)
            for j in range(5):
                col = L(3 + j)
                if top == 5:
                    f = closed_form_vps(f"$B{r}", f"{col}${hr}", I("exit_mult"), meth)
                else:
                    f = closed_form_vps(f"$B{r}", I("g"), f"{col}${hr}", meth)
                put(ws, f"{col}{r}", f, "ps")
                if k == 0 and j == 2:
                    ws[f"{col}{r}"].fill = PatternFill("solid", fgColor="DCE6F1")
                    ws[f"{col}{r}"].font = Font(name=FONT, bold=True)
        grids[top] = {"center": f"E{hr+4}", "range": f"C{hr+1}:G{hr+7}"}
    return grids


def sheet_comps(ws):
    f_title(ws, "Comps", "Peer multiples on next-twelve-month estimates; Include = 1 keeps a peer in the statistics")
    heads = ["Company", "Ticker", "Include", "Price", "Diluted shares (m)", "Market cap", "Net debt and other claims",
             "Enterprise value", "Revenue NTM", "EBITDA NTM", "EBIT NTM", "Net income NTM", "Revenue growth NTM",
             "EBIT margin NTM", "EV / Revenue", "EV / EBITDA", "EV / EBIT", "P / E", "Source and as-of"]
    widths(ws, {"A": 2, "B": 28, "C": 9, "D": 8, **{L(c): 11 for c in range(5, 20)}, "T": 40, "U": 2,
                "V": 10, "W": 10, "X": 10, "Y": 10})
    for j, text in enumerate(heads):
        label(ws, f"{L(2+j)}5", text, bold=True)
    for j, text in enumerate(["EV/Rev used", "EV/EBITDA used", "EV/EBIT used", "P/E used"]):
        label(ws, f"{L(22+j)}5", text, bold=True)
    f_header(ws, 5, 2, 25)
    for r in range(6, 18):
        for col in "BCDEFHJKLMN":
            fmt = {"B": "text", "C": "text", "D": "int", "E": "ps", "N": "pct"}.get(col, "amt")
            put(ws, f"{col}{r}", None, fmt, kind="input")
        put(ws, f"T{r}", None, "text", kind="input")
        put(ws, f"G{r}", f"=E{r}*F{r}", "amt")
        put(ws, f"I{r}", f"=G{r}+H{r}", "amt")
        put(ws, f"O{r}", f"=IF(J{r}=0,0,L{r}/J{r})", "pct")
        for col, num in (("P", "J"), ("Q", "K"), ("R", "L")):
            put(ws, f"{col}{r}", f"=IF(OR({num}{r}<=0,I{r}<=0),0,I{r}/{num}{r})", "mult")
        put(ws, f"S{r}", f"=IF(OR(M{r}<=0,G{r}<=0),0,G{r}/M{r})", "mult")
        for j, src in enumerate("PQRS"):
            put(ws, f"{L(22+j)}{r}", f'=IF(AND(D{r}=1,{src}{r}>0),{src}{r},"")', "mult")
    stats = [("Median", "MEDIAN"), ("25th percentile", "Q1"), ("75th percentile", "Q3"), ("Mean", "AVERAGE")]
    for i, (text, fn) in enumerate(stats):
        r = 19 + i
        label(ws, f"O{r}", text, bold=True)
        for j, col in enumerate("PQRS"):
            rng = f"{L(22+j)}6:{L(22+j)}17"
            if fn == "Q1":
                f = f"=IFERROR(QUARTILE({rng},1),0)"
            elif fn == "Q3":
                f = f"=IFERROR(QUARTILE({rng},3),0)"
            else:
                f = f"=IFERROR({fn}({rng}),0)"
            put(ws, f"{col}{r}", f, "mult", bold=fn == "MEDIAN")
    label(ws, "O23", "Peers included", bold=True)
    for j, col in enumerate("PQRS"):
        put(ws, f"{col}23", f"=COUNT({L(22+j)}6:{L(22+j)}17)", "int")
    label(ws, "O24", "Subject at market", bold=True)
    put(ws, "P24", f"=IF(Model!I{MODEL_ROWS['revenue']}=0,0,{D('mkt_ev')}/Model!I{MODEL_ROWS['revenue']})", "mult")
    put(ws, "Q24", f"={D('ev_ebitda_mkt')}", "mult", kind="link")
    put(ws, "R24", f"={D('ev_ebit_mkt')}", "mult", kind="link")
    total_rule(ws, 19, 15, 19)
    f_band(ws, 26, 2, 9, "Implied value for the subject")
    for j, text in enumerate(["Metric", "Subject NTM", "EV at 25th", "EV at median", "EV at 75th",
                              "Per share at 25th", "Per share at median", "Per share at 75th"]):
        label(ws, f"{L(2+j)}27", text, bold=True)
    f_header(ws, 27, 2, 9)
    implied = {}
    for i, (text, col, mrow) in enumerate((("EV / Revenue", "P", "revenue"), ("EV / EBITDA", "Q", "ebitda"),
                                           ("EV / EBIT", "R", "adj_ebit"))):
        r = 28 + i
        label(ws, f"B{r}", text)
        put(ws, f"C{r}", f"=Model!I{MODEL_ROWS[mrow]}", "amt", kind="link")
        for j, srow in enumerate((20, 19, 21)):
            put(ws, f"{L(4+j)}{r}", f"=C{r}*{col}{srow}", "amt")
            put(ws, f"{L(7+j)}{r}", f"=IF({DILUTED}=0,0,({L(4+j)}{r}+{BRIDGE_SUM})/{DILUTED})", "ps")
        implied[text] = r
    label(ws, "B32", "Implied EBITDA margin of peer medians (EV/Revenue / EV/EBITDA)")
    put(ws, "C32", "=IF(Q19=0,0,P19/Q19)", "pct")
    label(ws, "B33", "Subject next-year EBITDA margin")
    put(ws, "C33", f"=IF(Model!I{MODEL_ROWS['revenue']}=0,0,Model!I{MODEL_ROWS['ebitda']}/Model!I{MODEL_ROWS['revenue']})", "pct")
    ws.sheet_view.zoomScale = 85
    return implied


def sheet_football(ws, scen, grids, comps_rows):
    f_title(ws, "Football Field", "Value per share ranges by method against the current price")
    widths(ws, {"A": 2, "B": 40, "C": 12, "D": 12, "E": 12})
    for col, text in zip("BCDE", ["Method", "Low", "High", "Range"]):
        label(ws, f"{col}5", text, bold=True)
    f_header(ws, 5, 2, 5)
    sens = f"IF({I('tv_method')}=3,Sensitivity!C18:G24,Sensitivity!C7:G13)"
    rows = [
        ("52-week trading range", f"={I('low52')}", f"={I('high52')}"),
        ("Comps EV/Revenue (25th to 75th)", f"=Comps!G{comps_rows['EV / Revenue']}", f"=Comps!I{comps_rows['EV / Revenue']}"),
        ("Comps EV/EBITDA (25th to 75th)", f"=Comps!G{comps_rows['EV / EBITDA']}", f"=Comps!I{comps_rows['EV / EBITDA']}"),
        ("Comps EV/EBIT (25th to 75th)", f"=Comps!G{comps_rows['EV / EBIT']}", f"=Comps!I{comps_rows['EV / EBIT']}"),
        ("DCF sensitivity (WACC and terminal grid)", f"=MIN({sens})", f"=MAX({sens})"),
        ("DCF scenarios (Bear to Bull)", f"=MIN(Scenarios!C6:C8)", f"=MAX(Scenarios!C6:C8)"),
        ("Current share price (marker)", f"={I('price')}*0.995", f"={I('price')}*1.005"),
    ]
    for i, (text, lo, hi) in enumerate(rows):
        r = 6 + i
        label(ws, f"B{r}", text)
        put(ws, f"C{r}", lo, "ps", kind="link")
        put(ws, f"D{r}", hi, "ps", kind="link")
        put(ws, f"E{r}", f"=MAX(0,D{r}-C{r})", "ps")
    label(ws, "B13", "Current share price", bold=True)
    put(ws, "C13", f"={I('price')}", "ps", kind="link", bold=True)
    label(ws, "B14", "Intrinsic value per share (active case)", bold=True)
    put(ws, "C14", f"={D('vps')}", "ps", kind="link", bold=True)
    chart = BarChart()
    chart.type = "bar"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Valuation range per share"
    chart.legend = None
    data = Reference(ws, min_col=3, max_col=3, min_row=5, max_row=12)
    width = Reference(ws, min_col=5, max_col=5, min_row=5, max_row=12)
    cats = Reference(ws, min_col=2, min_row=6, max_row=12)
    chart.add_data(data, titles_from_data=True)
    chart.add_data(width, titles_from_data=True)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties = GraphicalProperties(noFill=True)
    chart.series[0].graphicalProperties.line.noFill = True
    chart.series[1].graphicalProperties = GraphicalProperties(solidFill=NAVY)
    from openpyxl.chart.marker import DataPoint
    pt = DataPoint(idx=6)
    pt.graphicalProperties = GraphicalProperties(solidFill="C0504D")
    chart.series[1].dPt.append(pt)
    chart.y_axis.title = "Per share"
    chart.y_axis.majorGridlines = None
    chart.x_axis.scaling.orientation = "maxMin"
    chart.height, chart.width = 8.5, 18
    ws.add_chart(chart, "G5")


def sheet_exec(ws, rev):
    f_title(ws, "Exec Summary", "Currency in millions, except per share")
    widths(ws, {"A": 2, "B": 44, "C": 30, **{L(4 + i): 10 for i in range(N_FC)}})
    f_band(ws, 5, 2, 3, "Valuation")
    items = [
        ("Company", f"={I('company')}", "text"), ("Ticker", f"={I('ticker')}", "text"),
        ("Model status", f"={I('status')}", "text"), ("Checks", "=Checks!D30", "text"),
        ("Valuation date", f"={I('valuation_date')}", "date"), ("Current share price", f"={I('price')}", "ps"),
        ("Active case", f'=CHOOSE({I("case")},"Bear","Base","Bull")', "text"),
        ("WACC", f"={WACC}", "pct2"),
        ("Terminal method", f'=CHOOSE({I("tv_method")},"Value driver","Perpetuity on FCF","Exit EV/EBITDA")', "text"),
        ("Terminal growth", f"={I('g')}", "pct"),
        ("Enterprise value", f"={D('ev')}", "amt"), ("Equity value", f"={D('equity')}", "amt"),
        ("Diluted shares", f"={D('diluted')}", "amt"), ("Intrinsic value per share", f"={D('vps')}", "ps"),
        ("Upside (downside)", f"={D('upside')}", "pct"), ("Terminal value share of EV", f"={D('tv_share')}", "pct"),
        ("Implied exit EV/EBITDA", f"={D('implied_mult')}", "mult"),
        ("Implied perpetual growth in exit-multiple value", f"={D('implied_g')}", "pct"),
    ]
    for i, (text, f, fmt) in enumerate(items):
        r = 6 + i
        label(ws, f"B{r}", text, bold=text in ("Intrinsic value per share", "Checks"))
        put(ws, f"C{r}", f, fmt, kind="link", bold=text == "Intrinsic value per share", align="right")
    r = 6 + len(items) + 1
    f_band(ws, r, 2, 3, "Scenarios")
    for i, case in enumerate(CASE_NAMES):
        label(ws, f"B{r+1+i}", f"{case} value per share")
        put(ws, f"C{r+1+i}", f"=Scenarios!C{6+i}", "ps", kind="link")
    r += 5
    f_band(ws, r, 2, 3, "What the price assumes")
    for i, (text, ref, fmt) in enumerate((
        ("Implied revenue growth shift (all years)", f"='Reverse DCF'!C{rev['g']['implied']}", "pct2"),
        ("Implied EBIT margin shift (all years)", f"='Reverse DCF'!C{rev['m']['implied']}", "pct2"),
        ("Market-implied discount rate", f"='Reverse DCF'!C{rev['r']['implied']}", "pct2"),
        ("EV / next-year EBITDA at market", f"={D('ev_ebitda_mkt')}", "mult"),
        ("Next-year FCF yield on market EV", f"={D('fcf_yield')}", "pct"),
    )):
        label(ws, f"B{r+1+i}", text)
        put(ws, f"C{r+1+i}", ref, fmt, kind="link", align="right")
    r += 7
    f_band(ws, r, 2, 3 + N_FC, "Forecast (Active case)")
    for t in range(N_FC):
        put(ws, f"{L(4+t)}{r+1}", f"=Model!{FC_COLS[t]}5", "yearE", kind="link", bold=True)
    f_header(ws, r + 1, 2, 3 + N_FC)
    for i, (text, key, fmt) in enumerate((("Revenue", "revenue", "amt"), ("Revenue growth", "growth", "pct"),
                                          ("Adjusted EBIT margin", "ebit_m", "pct"), ("Unlevered FCF", "ufcf", "amt"),
                                          ("ROIC", "roic", "pct"))):
        rr = r + 2 + i
        label(ws, f"B{rr}", text)
        for t in range(N_FC):
            put(ws, f"{L(4+t)}{rr}", f"=Model!{FC_COLS[t]}{MODEL_ROWS[key]}", fmt, kind="link")


def sheet_register(ws):
    f_title(ws, "Source Register", "One row per sourced figure: filing, XBRL concept and period")
    heads = ["Item", "Fiscal year", "Value (m)", "XBRL concept", "Unit", "Form", "Accession", "Filed",
             "Period start", "Period end", "Source"]
    widths(ws, {"A": 2, "B": 36, "C": 10, "D": 12, "E": 44, "F": 8, "G": 8, "H": 22, "I": 11, "J": 11, "K": 11, "L": 70})
    for j, text in enumerate(heads):
        label(ws, f"{L(2+j)}5", text, bold=True)
    f_header(ws, 5, 2, 12)


def sheet_checks(ws, scen, rev, grids):
    f_title(ws, "Checks", "Every status must read OK before the model is used for a decision")
    widths(ws, {"A": 2, "B": 52, "C": 14, "D": 14, "E": 22})
    for col, text in zip("BCDE", ["Check", "Value", "Status", "Pass condition"]):
        label(ws, f"{col}5", text, bold=True)
    f_header(ws, 5, 2, 5)
    case_inputs = "+".join(f"COUNTBLANK(Cases!D{CASE_BAND[c]+1}:M{CASE_BAND[c]+7})" for c in CASE_NAMES)
    hist_keys = ("revenue", "op_income", "da", "capex", "cfo", "pretax", "tax")
    hist_blank = "+".join(f"COUNTBLANK('Hist Financials'!D{HIST_ROWS[k]}:H{HIST_ROWS[k]})" for k in hist_keys)
    vps = D("vps")
    center = f"IF({I('tv_method')}=3,Sensitivity!{grids[16]['center']},Sensitivity!{grids[5]['center']})"
    g0, gl = rev["g"]["first"], rev["g"]["last"]
    errs = ("SUMPRODUCT(--ISERROR(Model!D8:R35))+SUMPRODUCT(--ISERROR(DCF!C5:R59))+SUMPRODUCT(--ISERROR(WACC!C6:C36))"
            "+SUMPRODUCT(--ISERROR(Scenarios!C6:F9))+SUMPRODUCT(--ISERROR('Reverse DCF'!C7:C30))")
    checks = [
        ("Case selector is 1, 2 or 3", f"={I('case')}", '=IF(AND(C{r}>=1,C{r}<=3,INT(C{r})=C{r}),"OK","FAIL")', "1 to 3", "int"),
        ("Terminal method is 1, 2 or 3", f"={I('tv_method')}", '=IF(AND(C{r}>=1,C{r}<=3,INT(C{r})=C{r}),"OK","FAIL")', "1 to 3", "int"),
        ("WACC above terminal growth", f"={WACC}-{I('g')}", '=IF(C{r}>0,"OK","FAIL")', "> 0", "pct2"),
        ("Terminal reinvestment rate g / RONIC", f"=IF({I('ronic')}=0,0,{I('g')}/{I('ronic')})",
         '=IF(AND(C{r}>0,C{r}<1),"OK","FAIL")', "between 0 and 1", "pct"),
        ("Stub share of first forecast year", f"=DCF!I{DCF_ROWS['stub']}", '=IF(AND(C{r}>0,C{r}<=1),"OK","FAIL")', "0 to 1", "pct"),
        ("Blank driver cells across the three cases", f"={case_inputs}", '=IF(C{r}=0,"OK","INCOMPLETE")', "0", "int"),
        ("Blank key historical cells", f"={hist_blank}", '=IF(C{r}=0,"OK","INCOMPLETE")', "0", "int"),
        ("Operating income bridge residual (sum of absolute)",
         f"=SUMPRODUCT(ABS('Hist Financials'!D{HIST_ROWS['op_residual']}:H{HIST_ROWS['op_residual']}))",
         '=IF(C{r}<0.5,"OK","REVIEW")', "< 0.5", "amt"),
        ("Scenario engine equals DCF (active case)", f"=ABS(Scenarios!C9-{vps})", '=IF(C{r}<0.005,"OK","FAIL")', "< 0.005", "dec2"),
        ("Reverse DCF growth grid at zero equals DCF", f"=ABS('Reverse DCF'!C{rev['g']['zero']}-{vps})",
         '=IF(C{r}<0.005,"OK","FAIL")', "< 0.005", "dec2"),
        ("Reverse DCF margin grid at zero equals DCF", f"=ABS('Reverse DCF'!C{rev['m']['zero']}-{vps})",
         '=IF(C{r}<0.005,"OK","FAIL")', "< 0.005", "dec2"),
        ("Sensitivity grid center equals DCF", f"=IFERROR(ABS({center}-{vps}),1)", '=IF(C{r}<0.005,"OK","FAIL")', "< 0.005", "dec2"),
        ("Bear at or below Base at or below Bull", "=IF(AND(Scenarios!C6<=Scenarios!C7,Scenarios!C7<=Scenarios!C8),1,0)",
         '=IF(C{r}=1,"OK","FAIL")', "1", "int"),
        ("Terminal value share of EV", f"={D('tv_share')}", '=IF(C{r}<=0.75,"OK","WATCH")', "75% or less", "pct"),
        ("Reverse growth grid rises with growth", f"=SUMPRODUCT(--('Reverse DCF'!C{g0+1}:C{gl}>'Reverse DCF'!C{g0}:C{gl-1}))",
         '=IF(C{r}=20,"OK","REVIEW")', "20", "int"),
        ("Price inside reverse growth grid", f"=IF(ISNUMBER('Reverse DCF'!C{rev['g']['implied']}),1,0)",
         '=IF(C{r}=1,"OK","WATCH")', "1", "int"),
        ("Diluted shares at or above basic", f"={D('diluted')}-{D('s_basic')}", '=IF(C{r}>=0,"OK","FAIL")', ">= 0", "amt"),
        ("Forecast net PP&E stays positive", "=MIN(Model!I29:R29)", '=IF(C{r}>=0,"OK","REVIEW")', ">= 0", "amt"),
        ("Formula errors in model sheets", f"={errs}", '=IF(C{r}=0,"OK","FAIL")', "0", "int"),
    ]
    for i, (text, val, status, cond, fmt) in enumerate(checks):
        r = 6 + i
        label(ws, f"B{r}", text)
        put(ws, f"C{r}", val, fmt)
        put(ws, f"D{r}", status.format(r=r), "text", bold=True, align="center")
        label(ws, f"E{r}", cond)
    last = 6 + len(checks) - 1
    label(ws, "B30", "Overall", bold=True)
    put(ws, "C30", f'=COUNTIF(D6:D{last},"<>OK")', "int", bold=True)
    put(ws, "D30", '=IF(C30=0,"ALL OK",C30&" to review")', "text", bold=True, align="center")
    total_rule(ws, 30, 2, 5)


# ---------------------------------------------------------------- data writer (shared with the SEC filler)
def write_data(wb, data):
    """Write input values only. Never touches formula cells."""
    ws = wb["Inputs"]
    for key, val in data.get("inputs", {}).items():
        ws[f"C{IN_ROW[key]}"] = val
    for key, note in data.get("rationale", {}).items():
        ws[f"E{IN_ROW[key]}"] = note
    for key, note in data.get("sources", {}).items():
        ws[f"F{IN_ROW[key]}"] = note
    hs = wb["Hist Financials"]
    for key, vals in data.get("hist", {}).items():
        for col, v in zip(HIST_COLS, vals):
            if v is not None:
                hs[f"{col}{HIST_ROWS[key]}"] = v
    cs = wb["Cases"]
    for case, drivers in data.get("cases", {}).items():
        for key, vals in drivers.items():
            for col, v in zip(CASE_COLS, vals):
                cs[f"{col}{case_row(case, key)}"] = v
    cp = wb["Comps"]
    for i, peer in enumerate(data.get("comps", [])[:12]):
        r = 6 + i
        for col, key in zip("BCDEFHJKLMN", ["name", "ticker", "include", "price", "shares", "net_claims",
                                            "rev", "ebitda", "ebit", "ni", "growth"]):
            cp[f"{col}{r}"] = peer.get(key)
        cp[f"T{r}"] = peer.get("source")
    sr = wb["Source Register"]
    for i, row in enumerate(data.get("register", [])):
        for j, v in enumerate(row):
            c = sr.cell(row=6 + i, column=2 + j, value=v)
            c.font = Font(name=FONT, color=BLUE_IN)
            if j == 2 and isinstance(v, (int, float)):
                c.number_format = FMT["amt"]


def build(path, data=None):
    wb = Workbook()
    wb.remove(wb.active)
    ws = {name: wb.create_sheet(name) for name in SHEETS}
    sheet_inputs(ws["Inputs"])
    sheet_cases(ws["Cases"])
    sheet_hist(ws["Hist Financials"])
    sheet_model(ws["Model"])
    sheet_wacc(ws["WACC"])
    sheet_dcf(ws["DCF"])
    scen = sheet_scenarios(ws["Scenarios"])
    rev = sheet_reverse(ws["Reverse DCF"])
    grids = sheet_sensitivity(ws["Sensitivity"])
    comps_rows = sheet_comps(ws["Comps"])
    sheet_football(ws["Football Field"], scen, grids, comps_rows)
    sheet_exec(ws["Exec Summary"], rev)
    sheet_register(ws["Source Register"])
    sheet_checks(ws["Checks"], scen, rev, grids)
    for name in ("Output -->", "Sources--->"):
        ws[name].sheet_view.showGridLines = False
    colors = {"Inputs": "1F497D", "Model": "1F497D", "WACC": "1F497D", "DCF": "1F497D", "Reverse DCF": "1F497D",
              "Cases": "C0504D", "Output -->": "FFFFFF", "Sources--->": "FFFFFF",
              "Exec Summary": "FF0000", "Football Field": "FF0000", "Sensitivity": "FF0000", "Scenarios": "FF0000",
              "Hist Financials": "F79646", "Comps": "F79646", "Source Register": "F79646", "Checks": "808080"}
    for name, color in colors.items():
        ws[name].sheet_properties.tabColor = color
    for sheet in wb.worksheets:
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.print_options.horizontalCentered = True
        sheet.page_margins.left = sheet.page_margins.right = 0.4
        sheet.sheet_view.selection[0].activeCell = "A1"
        sheet.sheet_view.selection[0].sqref = "A1"
        for row in sheet.iter_rows():
            for cell in row:
                if cell.font is None or cell.font.name != FONT:
                    f = cell.font
                    cell.font = Font(name=FONT, bold=f.bold, italic=f.italic, color=f.color, size=f.size)
    wb.active = wb.sheetnames.index("Exec Summary")
    if data:
        write_data(wb, data)
    wb.save(path)
    return path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("out")
    ap.add_argument("--example", action="store_true", help="fill with the synthetic Harborline Instruments case")
    a = ap.parse_args()
    data = None
    if a.example:
        from example_data import EXAMPLE
        data = EXAMPLE
    print(build(a.out, data))
