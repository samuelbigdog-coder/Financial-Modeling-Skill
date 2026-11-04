"""Write a synthetic SEC companyfacts JSON from the Harborline example, to test fill_from_sec.py offline.

    python sec_fixture.py fixture_companyfacts.json
"""
import json
import sys

from example_data import EXAMPLE, YEARS

CONCEPTS = {
    "revenue": "RevenueFromContractWithCustomerExcludingAssessedTax", "cogs": "CostOfRevenue",
    "sga": "SellingGeneralAndAdministrativeExpense", "rnd": "ResearchAndDevelopmentExpense",
    "op_income": "OperatingIncomeLoss", "da": "DepreciationDepletionAndAmortization", "sbc": "ShareBasedCompensation",
    "interest": "InterestExpense",
    "pretax": "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    "tax": "IncomeTaxExpenseBenefit", "net_income": "NetIncomeLoss", "cfo": "NetCashProvidedByUsedInOperatingActivities",
    "capex": "PaymentsToAcquirePropertyPlantAndEquipment", "acquisitions": "PaymentsToAcquireBusinessesNetOfCashAcquired",
    "buybacks": "PaymentsForRepurchaseOfCommonStock", "dividends": "PaymentsOfDividendsCommonStock",
}
INSTANTS = {
    "cash": "CashAndCashEquivalentsAtCarryingValue", "st_inv": "ShortTermInvestments", "ar": "AccountsReceivableNetCurrent",
    "inventory": "InventoryNet", "other_ca": "PrepaidExpenseAndOtherAssetsCurrent", "ap": "AccountsPayableCurrent",
    "accrued": "AccruedLiabilitiesCurrent", "def_rev": "ContractWithCustomerLiabilityCurrent",
    "ppe": "PropertyPlantAndEquipmentNet", "leases": "OperatingLeaseLiability", "equity": "StockholdersEquity",
    "debt": "LongTermDebt",
}


def build():
    h = EXAMPLE["hist"]
    gaap = {}

    def add(concept, unit, fact):
        gaap.setdefault(concept, {"units": {}})["units"].setdefault(unit, []).append(fact)

    for i, y in enumerate(YEARS):
        base = {"fy": y + 1, "fp": "FY", "form": "10-K", "filed": f"{y+1}-02-20", "accn": f"0000000000-{y+1-2000:02d}-000001"}
        for k, c in CONCEPTS.items():
            add(c, "USD", {**base, "start": f"{y}-01-01", "end": f"{y}-12-31", "val": h[k][i] * 1e6})
        add("WeightedAverageNumberOfDilutedSharesOutstanding", "shares",
            {**base, "start": f"{y}-01-01", "end": f"{y}-12-31", "val": h["dil_shares"][i] * 1e6})
        for k, c in INSTANTS.items():
            add(c, "USD", {**base, "end": f"{y}-12-31", "val": h[k][i] * 1e6})
        add("Goodwill", "USD", {**base, "end": f"{y}-12-31", "val": h["gw_intang"][i] * 0.7e6})
        add("IntangibleAssetsNetExcludingGoodwill", "USD", {**base, "end": f"{y}-12-31", "val": h["gw_intang"][i] * 0.3e6})
        # a quarterly fact that must be ignored for annual values
        add(CONCEPTS["revenue"], "USD", {**base, "form": "10-Q", "fp": "Q3", "start": f"{y}-07-01", "end": f"{y}-09-30",
                                         "val": h["revenue"][i] * 0.26e6})
    # a later restatement of FY2024 revenue, used by default; --as-filed keeps the original
    add(CONCEPTS["revenue"], "USD", {"fy": 2026, "fp": "FY", "form": "10-K", "filed": "2026-02-20",
                                     "accn": "0000000000-26-000001", "start": "2024-01-01", "end": "2024-12-31",
                                     "val": 3960.0e6})
    inp = EXAMPLE["inputs"]
    q = {"fy": 2026, "fp": "Q3", "form": "10-Q", "filed": "2026-10-30", "accn": "0000000000-26-000099", "end": "2026-09-30"}
    for k, c in (("cash", "CashAndCashEquivalentsAtCarryingValue"), ("st_inv", "ShortTermInvestments"),
                 ("debt", "LongTermDebt"), ("leases", "OperatingLeaseLiability"), ("book_equity", "StockholdersEquity")):
        add(c, "USD", {**q, "val": inp[k] * 1e6})
    add("MinorityInterest", "USD", {**q, "val": inp["minority"] * 1e6})
    dei = {"EntityCommonStockSharesOutstanding": {"units": {"shares": [{**q, "end": "2026-10-25", "val": inp["basic"] * 1e6}]}}}
    return {"cik": 1, "entityName": "Harborline Instruments (fictional)", "facts": {"us-gaap": gaap, "dei": dei}}


if __name__ == "__main__":
    json.dump(build(), open(sys.argv[1], "w"))
    print(sys.argv[1])
