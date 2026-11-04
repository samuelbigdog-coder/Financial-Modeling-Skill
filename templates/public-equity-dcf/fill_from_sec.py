"""Fill the DCF template's historical and balance-sheet inputs from SEC XBRL company facts.

Usage:
    python fill_from_sec.py --cik <CIK> --template Public_Equity_DCF_Template.xlsx --out TICKER_DCF.xlsx \
        --user-agent "Research assistant (your-contact)"
    python fill_from_sec.py --cik <CIK> --facts companyfacts.json --template ... --out ...

What it fills: Hist Financials (five fiscal years from 10-K filings), the latest balance-sheet
items for the equity bridge, basic shares, company name, fiscal year end, and one Source Register
row per figure (concept, form, accession, filed date, period). Values are in millions.

What it never fills: price, rates, beta, cases, comps, non-recurring adjustments or other operating
items. Those need research and judgment. Anything it cannot find stays blank and is listed at the end.
US GAAP filers only; IFRS filers (20-F) need manual entry.
"""
import argparse
import datetime as dt
import json
import sys
import urllib.request

from openpyxl import load_workbook

from build_dcf_template import write_data

ANNUAL_FORMS = ("10-K", "10-K/A")
# key: list of recipes; a recipe is a list of concept groups; a group is a list of alternatives (first found wins).
DURATION = {
    "revenue": [[["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet",
                  "RevenueFromContractWithCustomerIncludingAssessedTax"]]],
    "cogs": [[["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold"]]],
    "sga": [[["SellingGeneralAndAdministrativeExpense"]],
            [["GeneralAndAdministrativeExpense"], ["SellingAndMarketingExpense"]]],
    "rnd": [[["ResearchAndDevelopmentExpense", "ResearchAndDevelopmentExpenseExcludingAcquiredInProcessCost"]]],
    "op_income": [[["OperatingIncomeLoss"]]],
    "da": [[["DepreciationDepletionAndAmortization", "DepreciationAndAmortization",
             "DepreciationAmortizationAndAccretionNet"]]],
    "sbc": [[["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"]]],
    "interest": [[["InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt"]]],
    "pretax": [[["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
                 "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"]]],
    "tax": [[["IncomeTaxExpenseBenefit"]]],
    "net_income": [[["NetIncomeLoss"]]],
    "cfo": [[["NetCashProvidedByUsedInOperatingActivities"]]],
    "capex": [[["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"]]],
    "acquisitions": [[["PaymentsToAcquireBusinessesNetOfCashAcquired"]]],
    "buybacks": [[["PaymentsForRepurchaseOfCommonStock"]]],
    "dividends": [[["PaymentsOfDividendsCommonStock", "PaymentsOfDividends"]]],
}
SHARES_DURATION = {"dil_shares": [[["WeightedAverageNumberOfDilutedSharesOutstanding"]]]}
INSTANT = {
    "cash": [[["CashAndCashEquivalentsAtCarryingValue"]]],
    "st_inv": [[["ShortTermInvestments", "AvailableForSaleSecuritiesDebtSecuritiesCurrent", "MarketableSecuritiesCurrent"]]],
    "ar": [[["AccountsReceivableNetCurrent"]]],
    "inventory": [[["InventoryNet"]]],
    "other_ca": [[["PrepaidExpenseAndOtherAssetsCurrent", "OtherAssetsCurrent"]]],
    "ap": [[["AccountsPayableCurrent"]]],
    "accrued": [[["AccruedLiabilitiesCurrent"]]],
    "def_rev": [[["ContractWithCustomerLiabilityCurrent", "DeferredRevenueCurrent"]]],
    "ppe": [[["PropertyPlantAndEquipmentNet"]]],
    "gw_intang": [[["Goodwill"], ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet"]]],
    "debt": [[["LongTermDebt"], ["ShortTermBorrowings", "CommercialPaper"]],
             [["LongTermDebtNoncurrent"], ["LongTermDebtCurrent"], ["ShortTermBorrowings", "CommercialPaper"]]],
    "leases": [[["OperatingLeaseLiability"]]],
    "equity": [[["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest", "StockholdersEquity"]]],
}
BRIDGE = {"cash": "cash", "st_inv": "st_inv", "debt": "debt", "leases": "leases", "book_equity": "equity"}
BRIDGE_EXTRA = {"minority": [[["MinorityInterest"]]], "preferred": [[["PreferredStockValue"]]]}


def parse(d):
    return dt.date.fromisoformat(d)


def facts_for(gaap, concept, unit):
    return gaap.get(concept, {}).get("units", {}).get(unit, [])


def pick(rows, end, annual, as_filed):
    """Rows for one period end; annual durations need a FY 10-K fact about a year long."""
    out = []
    for f in rows:
        if f.get("end") != end:
            continue
        if annual:
            if f.get("form") not in ANNUAL_FORMS or "start" not in f:
                continue
            if not 350 <= (parse(f["end"]) - parse(f["start"])).days <= 380:
                continue
        elif "start" in f:
            continue
        out.append(f)
    if not out:
        return None
    out.sort(key=lambda f: f.get("filed", ""))
    return out[0] if as_filed else out[-1]


def resolve(gaap, recipes, end, annual, unit, as_filed):
    """Return (value, [(concept, fact)]) using the first recipe whose first group is found."""
    for recipe in recipes:
        total, used, first_found = 0.0, [], False
        for gi, group in enumerate(recipe):
            hit = None
            for concept in group:
                f = pick(facts_for(gaap, concept, unit), end, annual, as_filed)
                if f is not None:
                    hit = (concept, f)
                    break
            if hit is None:
                if gi == 0:
                    break
                continue
            first_found = first_found or gi == 0
            total += hit[1]["val"]
            used.append(hit)
        if first_found:
            return total, used
    return None, []


def latest_instant_end(gaap, concepts):
    ends = [f["end"] for c in concepts for f in facts_for(gaap, c, "USD") if "start" not in f]
    return max(ends) if ends else None


def fill(facts, cik, as_filed=False, years=5):
    gaap = facts.get("facts", {}).get("us-gaap", {})
    dei = facts.get("facts", {}).get("dei", {})
    rev_concepts = DURATION["revenue"][0][0]
    ends = sorted({f["end"] for c in rev_concepts for f in facts_for(gaap, c, "USD")
                   if f.get("form") in ANNUAL_FORMS and "start" in f
                   and 350 <= (parse(f["end"]) - parse(f["start"])).days <= 380})
    ends = ends[-years:]
    if not ends:
        raise SystemExit("No annual revenue facts found; check the CIK or enter data manually.")
    pad = [None] * (years - len(ends))
    hist, register, missing = {}, [], []
    url = lambda accn: f"https://www.sec.gov/Archives/edgar/data/{int(cik)}/{accn.replace('-', '')}/"

    def record(item, fy, val, used, unit):
        for concept, f in used:
            register.append([item, fy, round(f["val"] / 1e6, 3), f"us-gaap:{concept}", unit, f.get("form"),
                             f.get("accn"), f.get("filed"), f.get("start"), f.get("end"), url(f.get("accn", ""))])

    for key, recipes, annual, unit in ([(k, v, True, "USD") for k, v in DURATION.items()] +
                                       [(k, v, True, "shares") for k, v in SHARES_DURATION.items()] +
                                       [(k, v, False, "USD") for k, v in INSTANT.items()]):
        vals = []
        for end in ends:
            v, used = resolve(gaap, recipes, end, annual, unit, as_filed)
            if v is None:
                missing.append(f"{key} {end}")
                vals.append(None)
            else:
                vals.append(round(v / 1e6, 3))
                record(key, parse(end).year, v, used, unit)
        hist[key] = pad + vals
    hist["fy"] = pad + [parse(e).year for e in ends]
    hist["period_end"] = pad + [dt.datetime.combine(parse(e), dt.time()) for e in ends]

    inputs, sources = {}, {}
    name = facts.get("entityName")
    if name:
        inputs["company"] = name
    inputs["currency"] = "USD"
    inputs["last_fye"] = dt.datetime.combine(parse(ends[-1]), dt.time())
    sources["last_fye"] = "SEC XBRL companyfacts, latest 10-K fiscal year end"
    bs_end = latest_instant_end(gaap, ["CashAndCashEquivalentsAtCarryingValue"])
    if bs_end:
        inputs["bs_date"] = dt.datetime.combine(parse(bs_end), dt.time())
        for in_key, hkey in BRIDGE.items():
            v, used = resolve(gaap, INSTANT[hkey], bs_end, False, "USD", False)
            if v is None:
                missing.append(f"bridge {in_key} {bs_end}")
                continue
            inputs[in_key] = round(v / 1e6, 3)
            sources[in_key] = f"SEC XBRL {used[0][1].get('form')} {used[0][1].get('accn')}, as of {bs_end}"
            record(f"bridge {in_key}", parse(bs_end).year, v, used, "USD")
        for in_key, recipes in BRIDGE_EXTRA.items():
            v, used = resolve(gaap, recipes, bs_end, False, "USD", False)
            inputs[in_key] = round(v / 1e6, 3) if v is not None else 0.0
            if used:
                record(f"bridge {in_key}", parse(bs_end).year, v, used, "USD")
    so = [f for f in dei.get("EntityCommonStockSharesOutstanding", {}).get("units", {}).get("shares", [])]
    if so:
        f = max(so, key=lambda x: (x.get("end", ""), x.get("filed", "")))
        inputs["basic"] = round(f["val"] / 1e6, 3)
        sources["basic"] = f"SEC cover page {f.get('form')} {f.get('accn')}, as of {f.get('end')}"
        register.append(["basic shares", parse(f["end"]).year, inputs["basic"], "dei:EntityCommonStockSharesOutstanding",
                         "shares", f.get("form"), f.get("accn"), f.get("filed"), None, f.get("end"), url(f.get("accn", ""))])
    else:
        missing.append("basic shares")
    return {"inputs": inputs, "sources": sources, "hist": hist, "register": register}, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cik", required=True)
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--facts", help="local companyfacts JSON (skips the download)")
    ap.add_argument("--user-agent", help="SEC requires a descriptive User-Agent for downloads")
    ap.add_argument("--as-filed", action="store_true", help="use the first filing of each figure instead of the latest restatement")
    a = ap.parse_args()
    if a.facts:
        facts = json.load(open(a.facts))
    else:
        if not a.user_agent:
            sys.exit("Give --user-agent (SEC fair-access policy) or --facts.")
        req = urllib.request.Request(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{int(a.cik):010d}.json",
                                     headers={"User-Agent": a.user_agent})
        facts = json.load(urllib.request.urlopen(req, timeout=60))
    data, missing = fill(facts, a.cik, as_filed=a.as_filed)
    wb = load_workbook(a.template)
    write_data(wb, data)
    wb.save(a.out)
    print(f"Wrote {a.out}: {len(data['register'])} sourced figures.")
    if missing:
        print("Not found (enter manually and source in the Source Register):")
        for m in missing:
            print("  ", m)
    print("Still to research: price, 52-week range, rates and beta, cases, comps, non-recurring items, "
          "other operating items, RSUs and options. Then recalculate and run verify_dcf.py.")


if __name__ == "__main__":
    main()
