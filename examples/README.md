# Northstar Motion Systems mini-model

Northstar Motion Systems is a fictional industrial platform with three operating segments. The workbook includes the 25 reference sheet names in order, plus a terminal Checks sheet. Source data, companies, rates, targets, prices and case results are synthetic. The workbook opens in Base / Selective and uses USD millions except where marked.

## Use the workbook

- Select one operating case and one acquisition plan in Inputs. Blue type on yellow cells identifies editable assumptions. Cases and M&A Schedule hold the driver banks. Exec Summary gives the operating and sponsor results.
- The three forecast statements and debt waterfall share a single active build on Mini Model. The model assumes a year-end entry, annual forecast periods, full-year bolt-on EBITDA, opening-balance depreciation and revolver interest, average term debt interest, and no dividends. The simplified purchase allocates value to operating NWC, PP&E and goodwill. Fees are expensed at closing, the new debt is recorded at par, and cash taxes exclude NOLs and deferred tax.
- Returns Scenarios uses the live financing and cash flow path to show exit-year returns and exit-multiple sensitivity. Fan of Outcomes holds dated release captures from full-model recalculation, including operating and acquisition scenarios and entry-multiple changes. The capture sheet is static between builder runs. This preserves one calculation engine and keeps releases reproducible.
- The investment analysis is a sponsor buyout and exit-multiple case. It does not calculate standalone intrinsic DCF value.

## Reproduce and check

Run the Node builder with the Codex spreadsheet runtime. The default path points to that bundled runtime under the user home directory. Override it with an artifact-tool module file URL in the FIN_MODEL_ARTIFACT_TOOL environment variable where needed.

    node examples/build-mini-model.mjs
    python examples/verify_mini_model.py

The builder runs the cases through the active workbook, perturbs the exit multiple, recalculates and exports populated formula caches. The verifier requires Python 3.10+ and openpyxl and only reads the workbook. Its separate forecast ledger and IRR root solver compare the released results, while package checks scan for missing caches, errors, merged cells, macros and external links. The openpyxl reader is not used to write or calculate the workbook. The bundled calculation engine was used for recalc and mutation tests. Desktop Excel COM activation was unavailable in the session.

