"""Synthetic example: Harborline Instruments (fictional).

Every company, peer, price and figure here is invented. It exists only to show
the expected format and to let the verifier exercise every formula.
"""
import datetime as dt

YEARS = [2021, 2022, 2023, 2024, 2025]
REV = [3120.0, 3405.0, 3690.0, 3955.0, 4260.0]


def _hist():
    h = {k: [] for k in [
        "revenue", "cogs", "sga", "rnd", "other_opex", "op_income", "da", "sbc", "nonrecurring", "interest",
        "pretax", "tax", "net_income", "dil_shares", "cash", "st_inv", "ar", "inventory", "other_ca", "ap",
        "accrued", "def_rev", "ppe", "gw_intang", "debt", "leases", "equity", "cfo", "capex", "acquisitions",
        "buybacks", "dividends"]}
    gm = [0.452, 0.458, 0.461, 0.466, 0.470]
    sga = [0.168, 0.165, 0.163, 0.160, 0.158]
    rnd = [0.092, 0.093, 0.094, 0.095, 0.095]
    other = [4.0, 6.0, 38.0, 5.0, 7.0]           # FY2023 includes a 31.0 restructuring charge
    nonrec = [0.0, 0.0, 31.0, 0.0, 0.0]
    for i, rev in enumerate(REV):
        cogs = round(rev * (1 - gm[i]), 1)
        s, r_ = round(rev * sga[i], 1), round(rev * rnd[i], 1)
        op = round(rev - cogs - s - r_ - other[i], 1)
        da = round(rev * 0.041, 1)
        sbc = round(rev * 0.021, 1)
        interest = [52.0, 49.0, 55.0, 58.0, 54.0][i]
        pretax = round(op - interest + 6.0, 1)
        tax = round(pretax * 0.205, 1)
        ni = round(pretax - tax, 1)
        h["revenue"].append(rev); h["cogs"].append(cogs); h["sga"].append(s); h["rnd"].append(r_)
        h["other_opex"].append(other[i]); h["op_income"].append(op); h["da"].append(da); h["sbc"].append(sbc)
        h["nonrecurring"].append(nonrec[i]); h["interest"].append(interest); h["pretax"].append(pretax)
        h["tax"].append(tax); h["net_income"].append(ni)
        h["dil_shares"].append([126.8, 125.9, 124.7, 123.6, 122.4][i])
        h["cash"].append([610.0, 655.0, 590.0, 702.0, 748.0][i]); h["st_inv"].append([120.0, 95.0, 110.0, 130.0, 145.0][i])
        h["ar"].append(round(rev * 0.162, 1)); h["inventory"].append(round(rev * 0.118, 1))
        h["other_ca"].append(round(rev * 0.031, 1)); h["ap"].append(round(rev * 0.094, 1))
        h["accrued"].append(round(rev * 0.071, 1)); h["def_rev"].append(round(rev * 0.028, 1))
        h["ppe"].append([1180.0, 1232.0, 1291.0, 1338.0, 1402.0][i])
        h["gw_intang"].append([2210.0, 2195.0, 2260.0, 2244.0, 2231.0][i])
        h["debt"].append([1450.0, 1420.0, 1510.0, 1480.0, 1395.0][i]); h["leases"].append([210.0, 218.0, 226.0, 231.0, 240.0][i])
        h["equity"].append([2680.0, 2790.0, 2905.0, 3060.0, 3240.0][i])
        h["cfo"].append(round(ni + da + sbc - rev * 0.012, 1)); h["capex"].append(round(rev * 0.052, 1))
        h["acquisitions"].append([0.0, 0.0, 96.0, 0.0, 0.0][i]); h["buybacks"].append([150.0, 180.0, 200.0, 210.0, 240.0][i])
        h["dividends"].append([95.0, 101.0, 108.0, 114.0, 121.0][i])
    return h


def _path(start, end, n=10):
    return [round(start + (end - start) * t / (n - 1), 4) for t in range(n)]


CASES = {
    "Bear": {"growth": _path(0.035, 0.02), "ebit_m": _path(0.205, 0.19), "da": [0.041] * 10, "capex": _path(0.055, 0.045),
             "nwc": [0.12] * 10, "sbc": [0.021] * 10, "tax": [0.21] * 10},
    "Base": {"growth": _path(0.065, 0.03), "ebit_m": _path(0.225, 0.235), "da": [0.041] * 10, "capex": _path(0.052, 0.044),
             "nwc": [0.118] * 10, "sbc": [0.02] * 10, "tax": [0.21] * 10},
    "Bull": {"growth": _path(0.09, 0.035), "ebit_m": _path(0.235, 0.26), "da": [0.041] * 10, "capex": _path(0.05, 0.043),
             "nwc": [0.115] * 10, "sbc": [0.019] * 10, "tax": [0.205] * 10},
}

COMPS = [
    {"name": "Arcadia Sensor Systems", "ticker": "ARSS", "include": 1, "price": 64.2, "shares": 210.0, "net_claims": 1800.0,
     "rev": 6100.0, "ebitda": 1420.0, "ebit": 1150.0, "ni": 780.0, "growth": 0.058, "source": "Synthetic"},
    {"name": "Briarwood Metrology", "ticker": "BWM", "include": 1, "price": 118.5, "shares": 64.0, "net_claims": 900.0,
     "rev": 2900.0, "ebitda": 690.0, "ebit": 560.0, "ni": 390.0, "growth": 0.071, "source": "Synthetic"},
    {"name": "Calloway Controls", "ticker": "CLWC", "include": 1, "price": 42.8, "shares": 330.0, "net_claims": 2600.0,
     "rev": 8800.0, "ebitda": 1650.0, "ebit": 1290.0, "ni": 860.0, "growth": 0.041, "source": "Synthetic"},
    {"name": "Delmont Analytics", "ticker": "DLMA", "include": 1, "price": 205.0, "shares": 41.0, "net_claims": 350.0,
     "rev": 2350.0, "ebitda": 640.0, "ebit": 545.0, "ni": 400.0, "growth": 0.094, "source": "Synthetic"},
    {"name": "Easton Precision", "ticker": "EPRC", "include": 1, "price": 77.4, "shares": 98.0, "net_claims": 1100.0,
     "rev": 3700.0, "ebitda": 780.0, "ebit": 610.0, "ni": 410.0, "growth": 0.052, "source": "Synthetic"},
    {"name": "Fairhaven Optics", "ticker": "FHO", "include": 1, "price": 33.1, "shares": 150.0, "net_claims": 420.0,
     "rev": 1650.0, "ebitda": 330.0, "ebit": 255.0, "ni": 170.0, "growth": 0.063, "source": "Synthetic"},
    {"name": "Greyson Test and Measurement", "ticker": "GYTM", "include": 1, "price": 151.0, "shares": 88.0, "net_claims": 2100.0,
     "rev": 5200.0, "ebitda": 1240.0, "ebit": 1020.0, "ni": 700.0, "growth": 0.066, "source": "Synthetic"},
    {"name": "Holloway Industrial Software", "ticker": "HWIS", "include": 0, "price": 288.0, "shares": 52.0, "net_claims": -400.0,
     "rev": 1900.0, "ebitda": 620.0, "ebit": 560.0, "ni": 430.0, "growth": 0.16, "source": "Synthetic; excluded: software mix"},
]

EXAMPLE = {
    "inputs": {
        "company": "Harborline Instruments (fictional)", "ticker": "HBLN (synthetic)", "currency": "USD",
        "last_fye": dt.datetime(2025, 12, 31), "valuation_date": dt.datetime(2026, 11, 4), "price": 72.5,
        "price_date": dt.datetime(2026, 11, 3), "low52": 58.4, "high52": 84.9, "status": "ILLUSTRATIVE",
        "case": 2, "midyear": 1, "tv_method": 1, "g": 0.025, "ronic": 0.12, "exit_mult": 12.0, "sbc_cash": 1,
        "rf": 0.0491, "erp": 0.05, "beta_u": 0.95, "spread": 0.013, "kd_override": 0, "tax_m": 0.23,
        "wacc_method": 2, "de_override": None,
        "bs_date": dt.datetime(2026, 9, 30), "cash": 690.0, "st_inv": 150.0, "debt": 1420.0, "leases": 244.0,
        "inc_leases": 1, "preferred": 0.0, "minority": 18.0, "eq_inv": 40.0, "other_dl": 35.0, "book_equity": 3330.0,
        "basic": 121.6, "rsus": 1.9, "options": 3.2, "strike": 51.0, "other_dil": 0.0,
        "step_wacc": 0.005, "step_g": 0.005, "step_mult": 1.0, "step_gshift": 0.01, "step_mshift": 0.005,
    },
    "rationale": {"rf": "10-year Treasury", "erp": "House convention", "beta_u": "Synthetic industry beta",
                  "ronic": "Synthetic: modest premium to WACC in perpetuity"},
    "sources": {"rf": "Synthetic example value", "price": "Synthetic"},
    "hist": {**_hist(), "fy": YEARS, "period_end": [dt.datetime(y, 12, 31) for y in YEARS]},
    "cases": CASES,
    "comps": COMPS,
    "register": [["Revenue", 2025, 4260.0, "us-gaap:RevenueFromContractWithCustomerExcludingAssessedTax", "USD", "10-K",
                  "synthetic", "synthetic", "2025-01-01", "2025-12-31", "Synthetic example, not a filing"]],
}
