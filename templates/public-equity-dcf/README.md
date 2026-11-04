# Public equity DCF template

A formula-driven, banker-style DCF workbook for one public company, with a script that loads five years of SEC filing data, and an independent verifier that rebuilds every result in Python. It follows the financial-modeling skill's build sequence, methodology and verification steps, and the STARWAR visual system (navy header bands, blue inputs on pale yellow, green cross-sheet links, no merged cells).

| File | What it is |
| --- | --- |
| `Public_Equity_DCF_Template.xlsx` | Blank template. Edit only blue cells on pale yellow. |
| `Example_Harborline_Instruments_DCF.xlsx` | The same template filled with a fictional company, recalculated, every check OK. Use it to see the expected format. |
| `build_dcf_template.py` | Rebuilds both workbooks from code (`--example` for the filled copy). |
| `fill_from_sec.py` | Loads Hist Financials, the latest equity bridge, basic shares and a Source Register from SEC XBRL company facts. |
| `verify_dcf.py` | Independent oracle, package scan, direction tests and optional mutation tests. |
| `layout.py` | Shared cell map used by all three scripts. |
| `example_data.py`, `sec_fixture.py` | Synthetic data and an offline companyfacts fixture for testing. |

## Workflow

1. **Load the filings.** `python fill_from_sec.py --cik <CIK> --template Public_Equity_DCF_Template.xlsx --out <TICKER>_DCF.xlsx --user-agent "<name and contact>"` (or `--facts <file>.json` with a saved companyfacts download). Add `--as-filed` to keep each figure as first reported rather than restated. The script lists anything it could not find.
2. **Research the rest.** On Inputs: price and 52-week range (dated close, two sources), 10-year Treasury (FRED DGS10), equity risk premium, unlevered beta, credit spread or bond YTM, RSUs and options, terminal settings. On Hist Financials: non-recurring items (pre-tax, from the filing reconciliation) and other operating items until the operating income bridge residual is zero. On Cases: ten years of Bear, Base and Bull drivers. On Comps: up to twelve peers on next-twelve-month figures. Put a rationale and a dated source beside every input.
3. **Recalculate** in Excel, or with LibreOffice headless.
4. **Verify.** `python verify_dcf.py <TICKER>_DCF.xlsx --recalc "<command that recalculates {file} in place>"`. It must report zero failures, and the Checks sheet must read ALL OK.

## Sheets

| Tab | Role |
| --- | --- |
| Inputs | Every assumption with unit, rationale and source; case, terminal-method, mid-year, SBC and lease switches |
| Cases | Bear, Base and Bull driver banks by year; the Active block follows the case selector |
| Model | Five historical years and ten forecast years: revenue, adjusted EBIT, D&A, taxes, NOPAT, capex, working capital, stock comp, unlevered FCF, PP&E and invested capital, ROIC, reinvestment rate and fundamental growth |
| WACC | CAPM on a relevered industry beta (course formula), cost of debt from YTM or spread, market-weight and book-weight WACC and their average |
| DCF | Stub period from the valuation date, mid-year option, three terminal methods (value driver NOPAT × (1 − g/RONIC), perpetuity on final-year FCF, exit EV/EBITDA) with implied multiple and implied growth cross-checks, full EV-to-equity bridge, treasury-method dilution, implied multiples at intrinsic value and at market |
| Reverse DCF | What the price implies: uniform revenue-growth shift, uniform EBIT-margin shift and the market-implied discount rate, each from a rerun of the full engine |
| Exec Summary | One-page answer, scenario values, price-implied expectations, forecast strip, check status |
| Football Field | Ranges from 52-week trading, comps (25th to 75th), DCF sensitivity and scenarios, with a chart and a price marker |
| Sensitivity | WACC × terminal growth and WACC × exit multiple grids; centers tie to the live model |
| Scenarios | All three cases valued side by side through the same engine |
| Hist Financials | Reported figures as filed, operating-income bridge, adjusted EBIT, operating NWC, ratios and three-year averages |
| Comps | Peer multiples with an include switch, median and quartiles, implied value per share, implied-margin check |
| Source Register | One row per sourced figure: concept, form, accession, filed date and period |
| Checks | Nineteen pass/fail tests, from selector validity and WACC above growth to engine ties, grid monotonicity and formula errors |

## Conventions and simplifications

- Currency in millions except per-share figures. Historical expenses and cash outflows are entered as positive numbers; the Model applies the signs.
- Stock comp is treated as a cash cost by default (switchable), so it is never added back while dilution is ignored.
- Leases are treated as debt by default (switchable); keep the treatment consistent with the comps.
- Option dilution uses the treasury method at the current price, which avoids a circular reference.
- The forecast is an operating model with PP&E and working-capital roll-forwards, not a full three-statement model. Companies where financing, leverage or working-capital swings drive value need a three-statement build on top of this.
- Banks, insurers and REITs need industry methods (excess return on equity, embedded value, FFO or NAV) and are out of scope.
- `fill_from_sec.py` reads US GAAP company facts only. IFRS filers need manual entry. It was tested against the offline fixture; check its output against the filing before relying on it.

All data in the example and the fixture is invented. No proprietary workbook content is included.
