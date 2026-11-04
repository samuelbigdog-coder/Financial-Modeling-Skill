# Verification and handoff

## Run before delivery

1. Verify sheet order, labels, units, dates, formula consistency, editable input provenance, and the absence of unauthorized external links and merged ranges. Inspect relevant regions visually. If editing, compare intended and actual changes against the original.
2. Recalculate with an engine supporting the actual features used. Reopen formulas and cached values independently. Scan for formula errors, missing numeric caches and unexpected blanks. Record engine, version and any feature limitation.
3. Reconcile raw sources to historicals, balance sheet parity, cash rollforward, financing sources and uses, debt and equity schedules, EV-to-equity and returns. Calculate residuals from distinct paths and compare with explicit absolute tolerances expressed in model units. Never round away material differences.
4. Build an independent oracle from raw inputs and assumptions. Avoid reading the workbook's computed totals as oracle inputs. Verify headline values, intermediate checkpoints and agreement across relevant scenarios.
5. Run expected-direction and adverse tests. Examples include higher discount rate at positive cash flows, lower multiple, zero acquisitions, constrained revolver and cash shortfall, no tax benefit on losses where assumed, and invalid terminal growth. Assert the grid center matches the live result and compute selected corners independently. Restore the delivery case, recalculate and reopen again.
6. Report files, scenario tested, check results, simplifications and unresolved limitations. Deliver supporting scripts and instructions when reproducibility is required.

## Skill evaluation prompts

| Prompt | Pass criteria |
| --- | --- |
| Repair a model with a capped revolver and a liquidity gap | The gap stays visible and cash, financing and debt schedules reconcile. |
| Extend a three-statement model with an acquisition | Acquired and organic growth separate, consideration reconciles, and opening balances and noncash changes are classified correctly. |
| Build a DCF with reverse DCF and a WACC × growth grid | Center ties, invalid perpetuities are flagged, terminal assumptions align with reinvestment and implied price ties in the forward model. |
| Audit a sponsor model with interim dividends and PIK | Debt and sponsor distributions reconcile, and independently computed dated NPV is near zero at reported IRR. |
| Recreate the STARWAR layout with synthetic data | Every requested sheet exists in order, no merged ranges, editable inputs are distinct and the deliverable contains no proprietary workbook content. |

These prompts define acceptance criteria for future agent evaluations. Workbook test results belong with the executable example and should be distinguished from skill-trigger or agent evaluations.
