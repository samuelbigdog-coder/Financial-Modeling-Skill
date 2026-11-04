"""Independently reconstruct the fictional operating and sponsor cases.

Run with python -m pip install openpyxl if the read-only workbook reader is absent.
"""
import math
import sys
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

SFG_ORDER = (
    "Inputs", "Mini Model", "Cases", "M&A Schedule", "Output -->", "Equity Bridge",
    "Fan of Outcomes", "CF Walk", "Returns Scenarios", "LR Performance", "Segment Margins",
    "Exec Summary", "Sources--->", "P&L Detailed", "DS Backlog", "M&A Targets", "EM Backlog",
    "Hist BS Detailed", "SOFR", "Contr. Margin", "LR Data", "Margin Explored", "Comps",
    "Deal Comps", "Capex", "Checks",
)
CASES = {
    "Base": ((.065, .075, .09), (.39, .49, .30)),
    "Upside": ((.105, .11, .14), (.42, .52, .34)),
    "Downside": ((-.04, .015, -.06), (.32, .43, .22)),
    "Management": ((.09, .095, .12), (.41, .51, .33)),
}
PLANS = {
    "Selective": (0, 1, 0, 1, 0),
    "Accelerated": (1, 1, 1, 1, 1),
    "Limited": (0, 0, 1, 0, 0),
    "None": (0, 0, 0, 0, 0),
}
BENCHMARKS = (.041, .037, .035, .035, .036)


def irr(cash_flows):
    """Solve the one-sign-change annual NPV equation independently of workbook IRR."""
    if cash_flows[-1] <= 0:
        return -1
    low, high = -.999, 1.0
    while sum(cf / (1 + high) ** t for t, cf in enumerate(cash_flows)) > 0:
        high *= 2
    for _ in range(100):
        candidate = (low + high) / 2
        npv = sum(cf / (1 + candidate) ** t for t, cf in enumerate(cash_flows))
        if npv > 0:
            low = candidate
        else:
            high = candidate
    return (low + high) / 2


def simulate(case, plan, entry_multiple, exit_multiple=8.5):
    """Solve using a separate ledger and project absolute values from fictional inputs."""
    growth, margins = CASES[case]
    revenue = (55, 31, 24)
    fixed_cost = 10.1
    starting_ebitda = 110 - 67 - 20 + .7
    price = starting_ebitda * entry_multiple
    fee = .018 * price
    original_loan = min(3.25 * starting_ebitda, .48 * price)
    initial_equity = price + fee + 3 - original_loan
    assets = {"cash": 3, "ar": 14.5, "inventory": 8.5, "ppe": 14, "goodwill": price - 30.3}
    liabilities = {"ap": 6.7, "term": original_loan, "revolver": 0}
    book_equity = initial_equity - fee
    accumulated_addons = 0
    equity_flows = [-initial_equity]
    forecast = []
    for year, (new_addons, sofr) in enumerate(zip(PLANS[plan], BENCHMARKS, strict=True)):
        recovery = year * .008 if case == "Downside" else 0
        revenue = tuple(sales * (1 + rate + recovery) for sales, rate in zip(revenue, growth, strict=True))
        gross_profit = sum(sales * (margin + recovery) for sales, margin in zip(revenue, margins, strict=True))
        fixed_cost *= 1.035
        organic_sales = sum(revenue)
        organic_profit = gross_profit - fixed_cost - organic_sales * .09
        accumulated_addons += new_addons
        addon_profit = accumulated_addons * 1.4
        addon_sales = addon_profit / .23
        sales = organic_sales + addon_sales
        ebitda = organic_profit + addon_profit
        cost = sales - gross_profit - addon_profit
        opening_term = liabilities["term"]
        amortization = min(opening_term, .07 * original_loan)
        liabilities["term"] -= amortization
        opening_revolver = liabilities["revolver"]
        interest = (opening_term - amortization / 2) * (sofr + .045) + opening_revolver * (sofr + .055)
        depreciation = assets["ppe"] * .15
        pretax = ebitda - depreciation - interest
        net_income = pretax - max(0, pretax) * .26
        capex = .024 * sales + .012 * organic_sales
        acquisition = new_addons * 1.4 * 6.1
        opening_ar, opening_inventory, opening_ap = assets["ar"], assets["inventory"], liabilities["ap"]
        assets["ar"], assets["inventory"], liabilities["ap"] = sales * 48 / 365, cost * 47 / 365, cost * 37 / 365
        working_capital_investment = (
            assets["ar"] - opening_ar + assets["inventory"] - opening_inventory - liabilities["ap"] + opening_ap
        )
        assets["cash"] += net_income + depreciation - working_capital_investment - capex - acquisition - amortization
        assets["ppe"] += capex - depreciation + acquisition * .2
        assets["goodwill"] += acquisition * .8
        cash_gap = max(0, 3 - assets["cash"])
        available = max(0, min(16, 3.75 * ebitda - liabilities["term"]) - opening_revolver)
        draw = min(cash_gap, available)
        repay = min(opening_revolver, max(0, assets["cash"] - 3))
        equity_injection = cash_gap - draw
        liabilities["revolver"] += draw - repay
        assets["cash"] += draw - repay + equity_injection
        book_equity += net_income + equity_injection
        assert math.isclose(sum(assets.values()), sum(liabilities.values()) + book_equity, abs_tol=1e-9)
        net_debt = liabilities["term"] + liabilities["revolver"] - assets["cash"]
        proceeds = max(0, ebitda * exit_multiple - net_debt)
        equity_flows.append(-equity_injection)
        forecast.append({
            25: organic_sales, 28: sales, 33: cost, 37: organic_profit, 38: ebitda,
            40: depreciation, 42: interest, 45: net_income, 51: assets["cash"],
            52: assets["ar"], 53: assets["inventory"], 54: assets["ppe"], 55: assets["goodwill"],
            56: sum(assets.values()), 57: liabilities["ap"], 58: liabilities["term"],
            59: liabilities["revolver"], 60: book_equity, 66: working_capital_investment,
            67: capex, 68: acquisition, 76: amortization, 79: draw, 80: repay, 88: equity_injection,
            95: net_debt, 96: proceeds, 99: proceeds / -sum(equity_flows),
            100: irr([*equity_flows[:-1], equity_flows[-1] + proceeds]),
        })
    return forecast


def equal(actual, expected, location):
    assert isinstance(actual, (float, int)), (location, actual, expected)
    assert math.isclose(actual, expected, rel_tol=2e-9, abs_tol=1e-8), (location, actual, expected)


def verify(path):
    raw = load_workbook(path)
    cached = load_workbook(path, data_only=True)
    assert tuple(raw.sheetnames) == SFG_ORDER
    assert len(raw["SOFR"]._charts) == 1
    selectors = {str(rule.sqref): rule.formula1 for rule in raw["Inputs"].data_validations.dataValidation}
    assert selectors == {
        "E8": '"Base,Upside,Downside,Management"',
        "E9": '"Selective,Accelerated,Limited,None"',
    }
    with ZipFile(path) as zf:
        assert not any("externalLinks/" in name or "vbaProject" in name for name in zf.namelist())
    formula_count = 0
    for sheet in raw:
        assert not sheet.merged_cells.ranges, sheet.title
        for rows in sheet:
            for cell in rows:
                evaluated = cached[sheet.title][cell.coordinate]
                assert evaluated.data_type != "e", (sheet.title, cell.coordinate, evaluated.value)
                if cell.data_type == "f":
                    formula_count += 1
                    assert evaluated.value is not None or (sheet.title == "Deal Comps" and cell.coordinate == "J13"), (
                        sheet.title, cell.coordinate, cell.value,
                    )
                    assert "'Checks'!" not in cell.value or sheet.title == "Checks"
    checks = cached["Checks"]
    assert [checks[f"{c}26"].value for c in "EFGHIJKL"] == ["PASS"] * 8
    model = cached["Mini Model"]
    returns = cached["Returns Scenarios"]
    active = simulate("Base", "Selective", 8.25)
    for col, predicted in zip("HIJKL", active, strict=True):
        for row, value in predicted.items():
            if row != 100 or col == "L":
                equal(model[f"{col}{row}"].value, value, f"Mini Model!{col}{row}")
        equal(returns[f"{col}13"].value, predicted[100], f"Returns Scenarios!{col}13")
    for idx, multiple in enumerate((7.5, 8, 8.5, 9, 9.5), start=17):
        predicted = simulate("Base", "Selective", 8.25, multiple)
        for col, year in zip("HIJKL", predicted, strict=True):
            equal(returns[f"{col}{idx}"].value, year[99], f"Returns Scenarios!{col}{idx}")
    fan = cached["Fan of Outcomes"]
    for row in range(10, 17):
        case, plan, multiple = (fan[f"{col}{row}"].value for col in "EFG")
        predicted = simulate(case, plan, multiple)
        exit_ = predicted[-1]
        expected = (exit_[38], exit_[51], sum(period[88] for period in predicted), exit_[99], exit_[100])
        for col, amount in zip("HIJKL", expected, strict=True):
            equal(fan[f"{col}{row}"].value, amount, f"Fan of Outcomes!{col}{row}")
    assert fan["L11"].value > fan["L10"].value > fan["L12"].value
    assert fan["J12"].value > 0, "Downside must exercise equity funding"
    assert fan["L15"].value > fan["L10"].value > fan["L16"].value
    assert fan["H14"].value < fan["H10"].value
    print(f"PASS: {len(SFG_ORDER)} sheets, {formula_count} evaluated formulas, 8 reconciliation statuses, "
          "5 independently modeled years, 25 exit sensitivities and 7 scenario captures.")


if __name__ == "__main__":
    workbook = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("Northstar_Motion_MiniModel.xlsx")
    verify(workbook)


