"""Shared cell map for the public-equity DCF template.

The builder, the SEC filler and the verifier all import this module, so a
row or column moves in one place only.
"""
from openpyxl.utils import get_column_letter

N_HIST = 5
N_FC = 10

# Sheet names, in workbook order.
SHEETS = [
    "Inputs", "Cases", "Model", "WACC", "DCF", "Reverse DCF", "Output -->",
    "Exec Summary", "Football Field", "Sensitivity", "Scenarios", "Sources--->",
    "Hist Financials", "Comps", "Source Register", "Checks",
]


def q(sheet):
    """Sheet name as used inside a formula reference."""
    return f"'{sheet}'" if (" " in sheet or "-" in sheet or ">" in sheet) else sheet


# Historical columns (Hist Financials and Model) and forecast columns (Model, DCF).
HIST_COLS = ["D", "E", "F", "G", "H"]
FC_COLS = [get_column_letter(9 + i) for i in range(N_FC)]          # I..R
CASE_COLS = [get_column_letter(4 + i) for i in range(N_FC)]        # D..M on Cases

# ---------------------------------------------------------------- Inputs
# (key, label, unit, default, number format, kind)
# kind: "text", "date", "num", "pct", "int", "mult", "ps" (per share), "sel" (selector)
INPUT_SECTIONS = [
    ("Company and market", [
        ("company", "Company name", "", None, "@", "text"),
        ("ticker", "Ticker and exchange", "", None, "@", "text"),
        ("currency", "Reporting currency", "", "USD", "@", "text"),
        ("last_fye", "Last fiscal year end (latest 10-K)", "date", None, "mm-dd-yy", "date"),
        ("valuation_date", "Valuation date", "date", None, "mm-dd-yy", "date"),
        ("price", "Current share price", "per share", None, "ps", "ps"),
        ("price_date", "Price date and time", "date", None, "mm-dd-yy", "date"),
        ("low52", "52-week low", "per share", None, "ps", "ps"),
        ("high52", "52-week high", "per share", None, "ps", "ps"),
        ("status", "Model status", "", "ILLUSTRATIVE", "@", "text"),
    ]),
    ("Case and terminal value", [
        ("case", "Active case (1 Bear, 2 Base, 3 Bull)", "selector", 2, "0", "sel"),
        ("midyear", "Mid-year discounting (1 on, 0 off)", "selector", 1, "0", "sel"),
        ("tv_method", "Terminal method (1 value driver, 2 perpetuity on FCF, 3 exit EV/EBITDA)", "selector", 1, "0", "sel"),
        ("g", "Terminal growth", "%", 0.025, "pct", "pct"),
        ("ronic", "Return on new invested capital in terminal period", "%", None, "pct", "pct"),
        ("exit_mult", "Exit EV/EBITDA multiple (cross-check)", "x", None, "mult", "mult"),
        ("sbc_cash", "Deduct stock comp as a cash cost (1 yes, 0 no)", "selector", 1, "0", "sel"),
    ]),
    ("Cost of capital", [
        ("rf", "Risk-free rate (10-year Treasury, FRED DGS10)", "%", None, "pct2", "pct"),
        ("erp", "Equity risk premium", "%", 0.05, "pct2", "pct"),
        ("beta_u", "Unlevered beta (industry)", "x", None, "dec2", "num"),
        ("spread", "Credit spread over risk-free", "%", None, "pct2", "pct"),
        ("kd_override", "Pre-tax cost of debt override (bond YTM, 0 = use spread)", "%", 0, "pct2", "pct"),
        ("tax_m", "Marginal tax rate", "%", 0.21, "pct", "pct"),
        ("wacc_method", "Weights (1 market, 2 average of market and book)", "selector", 2, "0", "sel"),
        ("de_override", "Target debt/equity override (blank = observed)", "x", None, "dec2", "num"),
    ]),
    ("Equity bridge (latest balance sheet)", [
        ("bs_date", "Balance sheet date", "date", None, "mm-dd-yy", "date"),
        ("cash", "Cash and equivalents", "m", None, "amt", "num"),
        ("st_inv", "Short-term investments", "m", None, "amt", "num"),
        ("debt", "Total debt (current and long-term)", "m", None, "amt", "num"),
        ("leases", "Operating lease liabilities", "m", None, "amt", "num"),
        ("inc_leases", "Treat leases as debt (1 yes, 0 no)", "selector", 1, "0", "sel"),
        ("preferred", "Preferred equity", "m", None, "amt", "num"),
        ("minority", "Noncontrolling interests", "m", None, "amt", "num"),
        ("eq_inv", "Non-operating equity investments", "m", None, "amt", "num"),
        ("other_dl", "Other debt-like items (pension deficit, legal)", "m", None, "amt", "num"),
        ("book_equity", "Book equity (latest)", "m", None, "amt", "num"),
    ]),
    ("Share count", [
        ("basic", "Basic shares outstanding (10-Q cover)", "m", None, "amt", "num"),
        ("rsus", "Unvested RSUs and PSUs", "m", None, "amt", "num"),
        ("options", "Options outstanding", "m", None, "amt", "num"),
        ("strike", "Options weighted average strike", "per share", None, "ps", "ps"),
        ("other_dil", "Other dilutive shares (convertibles if-converted)", "m", None, "amt", "num"),
    ]),
    ("Sensitivity steps", [
        ("step_wacc", "WACC step", "%", 0.005, "pct", "pct"),
        ("step_g", "Terminal growth step", "%", 0.005, "pct", "pct"),
        ("step_mult", "Exit multiple step", "x", 1.0, "mult", "mult"),
        ("step_gshift", "Reverse DCF growth shift step", "%", 0.01, "pct", "pct"),
        ("step_mshift", "Reverse DCF margin shift step", "%", 0.005, "pct", "pct"),
    ]),
]


def input_rows():
    rows, r = {}, 5
    for _title, items in INPUT_SECTIONS:
        r += 1                     # band row
        for item in items:
            r += 1
            rows[item[0]] = r
        r += 1                     # spacer
    return rows


IN_ROW = input_rows()


def I(key, absolute=True):
    r = IN_ROW[key]
    return f"Inputs!$C${r}" if absolute else f"Inputs!C{r}"


# ---------------------------------------------------------------- Cases
DRIVERS = [
    ("growth", "Revenue growth"),
    ("ebit_m", "Adjusted EBIT margin"),
    ("da", "D&A % of revenue"),
    ("capex", "Capex % of revenue"),
    ("nwc", "Operating NWC % of revenue"),
    ("sbc", "Stock comp % of revenue"),
    ("tax", "Cash tax rate on EBIT"),
]
CASE_NAMES = ["Bear", "Base", "Bull"]
CASE_BAND = {"Bear": 7, "Base": 16, "Bull": 25, "Active": 34}


def case_row(case, driver):
    idx = [d[0] for d in DRIVERS].index(driver)
    return CASE_BAND[case] + 1 + idx


# ---------------------------------------------------------------- Hist Financials
HIST_ROWS = {
    "fy": 6, "period_end": 7,
    "revenue": 10, "cogs": 11, "gross_profit": 12, "sga": 13, "rnd": 14, "other_opex": 15,
    "op_income": 16, "op_residual": 17, "da": 18, "sbc": 19, "nonrecurring": 20, "adj_ebit": 21,
    "interest": 22, "pretax": 23, "tax": 24, "net_income": 25, "dil_shares": 26,
    "cash": 29, "st_inv": 30, "ar": 31, "inventory": 32, "other_ca": 33, "ap": 34, "accrued": 35,
    "def_rev": 36, "op_nwc": 37, "ppe": 38, "gw_intang": 39, "debt": 40, "leases": 41, "equity": 42,
    "cfo": 45, "capex": 46, "acquisitions": 47, "buybacks": 48, "dividends": 49, "fcf": 50,
    "r_growth": 53, "r_gm": 54, "r_ebit_m": 55, "r_da": 56, "r_capex": 57, "r_nwc": 58, "r_sbc": 59,
    "r_tax": 60, "r_roic": 61, "r_fcf_conv": 62,
}
HIST_INPUT_KEYS = [
    "revenue", "cogs", "sga", "rnd", "other_opex", "op_income", "da", "sbc", "nonrecurring",
    "interest", "pretax", "tax", "net_income", "dil_shares",
    "cash", "st_inv", "ar", "inventory", "other_ca", "ap", "accrued", "def_rev", "ppe", "gw_intang",
    "debt", "leases", "equity", "cfo", "capex", "acquisitions", "buybacks", "dividends",
]
HIST_LABELS = {
    "fy": "Fiscal year", "period_end": "Period end",
    "revenue": "Revenue", "cogs": "Cost of revenue", "gross_profit": "Gross profit",
    "sga": "Selling, general and administrative", "rnd": "Research and development",
    "other_opex": "Other operating expense (income)", "op_income": "Operating income (reported)",
    "op_residual": "Operating income bridge residual", "da": "Depreciation and amortization",
    "sbc": "Stock-based compensation", "nonrecurring": "Non-recurring items in operating income (pre-tax, add back)",
    "adj_ebit": "Adjusted EBIT", "interest": "Interest expense", "pretax": "Pre-tax income",
    "tax": "Income tax expense", "net_income": "Net income", "dil_shares": "Diluted weighted shares (m)",
    "cash": "Cash and equivalents", "st_inv": "Short-term investments", "ar": "Accounts receivable",
    "inventory": "Inventory", "other_ca": "Other current operating assets", "ap": "Accounts payable",
    "accrued": "Accrued and other current operating liabilities", "def_rev": "Deferred revenue (current)",
    "op_nwc": "Operating net working capital", "ppe": "Net PP&E", "gw_intang": "Goodwill and intangibles",
    "debt": "Total debt", "leases": "Operating lease liabilities", "equity": "Total equity",
    "cfo": "Cash from operations", "capex": "Capital expenditures", "acquisitions": "Acquisitions, net of cash",
    "buybacks": "Share repurchases", "dividends": "Dividends paid", "fcf": "Free cash flow (CFO less capex)",
    "r_growth": "Revenue growth", "r_gm": "Gross margin", "r_ebit_m": "Adjusted EBIT margin",
    "r_da": "D&A % of revenue", "r_capex": "Capex % of revenue", "r_nwc": "Operating NWC % of revenue",
    "r_sbc": "Stock comp % of revenue", "r_tax": "Effective tax rate", "r_roic": "ROIC (after marginal tax)",
    "r_fcf_conv": "FCF conversion (FCF / net income)",
}
HIST_BANDS = {9: "Income statement", 28: "Balance sheet", 44: "Cash flow", 52: "Ratios"}


def H(key, col):
    return f"'Hist Financials'!{col}{HIST_ROWS[key]}"


# ---------------------------------------------------------------- Model
MODEL_ROWS = {
    "year": 5, "revenue": 8, "growth": 9, "ebit_m": 10, "adj_ebit": 11, "da_pct": 12, "da": 13,
    "ebitda": 14, "tax_rate": 15, "taxes": 16, "nopat": 17,
    "capex_pct": 19, "capex": 20, "nwc_pct": 21, "nwc": 22, "d_nwc": 23, "sbc_pct": 24, "sbc": 25,
    "sbc_ded": 26, "ufcf": 27,
    "ppe": 29, "gw": 30, "ic": 31, "roic": 32, "net_reinv": 33, "reinv_rate": 34, "fund_growth": 35,
}
MODEL_BANDS = {7: "Operating build", 18: "Reinvestment and free cash flow", 28: "Capital and returns"}


def M(key, col):
    return f"Model!{col}{MODEL_ROWS[key]}"


# ---------------------------------------------------------------- WACC
WACC_ROWS = {
    "rf": 6, "erp": 7, "beta_u": 8, "ksu": 9,
    "kd": 11, "tax": 12, "kd_at": 13,
    "mkt_equity": 15, "debt_used": 16, "book_equity": 17, "de_m": 18, "de_b": 19,
    "ksl_m": 22, "wd_m": 23, "we_m": 24, "wacc_m": 25,
    "ksl_b": 28, "wd_b": 29, "we_b": 30, "wacc_b": 31,
    "wacc": 34, "beta_l": 35, "spread_g": 36,
}
WACC_BANDS = {5: "Cost of equity", 10: "Cost of debt", 14: "Capital structure",
              21: "Market-value weights", 27: "Book-value weights", 33: "WACC used"}
WACC = "WACC!$C$34"

# ---------------------------------------------------------------- DCF
DCF_ROWS = {
    "year": 5, "ufcf": 6, "period_end": 7, "stub": 8, "cf": 9, "period": 10, "df": 11, "pv": 12,
    "sum_pv": 15, "nopat_n1": 16, "sbc_n1": 17, "fcf_vd": 18, "tv_vd": 19, "tv_gordon": 20,
    "tv_exit": 21, "tv": 22, "tv_period": 23, "pv_tv": 24, "ev": 25, "tv_share": 26,
    "implied_mult": 27, "implied_g": 28,
    "b_ev": 31, "b_cash": 32, "b_stinv": 33, "b_eqinv": 34, "b_debt": 35, "b_leases": 36,
    "b_pref": 37, "b_minority": 38, "b_other": 39, "equity": 40,
    "s_basic": 43, "s_rsus": 44, "s_options": 45, "s_other": 46, "diluted": 47,
    "vps": 50, "price": 51, "upside": 52, "mkt_cap": 53, "mkt_ev": 54,
    "ev_ebitda_dcf": 55, "ev_ebitda_mkt": 56, "ev_ebit_dcf": 57, "ev_ebit_mkt": 58, "fcf_yield": 59,
}
DCF_BANDS = {14: "Enterprise value", 30: "Equity bridge", 42: "Diluted shares (treasury method at current price)",
             49: "Value per share and implied multiples"}


def D(key):
    return f"DCF!$C${DCF_ROWS[key]}"


BRIDGE_SUM = "SUM(DCF!$C$32:$C$39)"
DILUTED = "DCF!$C$47"
