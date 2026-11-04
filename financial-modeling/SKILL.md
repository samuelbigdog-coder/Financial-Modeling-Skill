---
name: financial-modeling
description: Build, audit, repair, or study Excel financial models, including operating forecasts, three-statement models, DCF and reverse DCF, LBO, M&A, debt, unit economics, and scenario analysis. Use when the task requires linked financial logic, valuation sensitivities, workbook reconciliation, or a STARWAR-style financial model.
---

# Financial modeling

Frame the decision in terms of the price today, the cash flows and proceeds over time, and the operating and financing assumptions required to earn the target return. Identify which assumptions could overturn the decision.

## Scope and references

- Read [methodology](references/methodology.md) for the applicable modules and [verification](references/verification.md) before delivery.
- Read the [STARWAR style specification](references/starwar-style.md) for new STARWAR-style workbooks. Preserve existing workbook conventions when editing unless the user requests a redesign. Use no merged cells in newly authored workbooks.
- Use the host spreadsheet tools for authoring, calculation and inspection. Confirm the target application, required formula features and actual calculation engine before selecting a workflow.
- Treat any user-supplied workbook as confidential. Keep its numbers, formulas, source notes, charts and metadata out of distributable examples. Use invented data with clear provenance. Refer to the private style reference only as STARWAR. Before public distribution, check filenames, document text, workbook contents and metadata, and archive entries for private identifiers.

## Build sequence

1. Establish the valuation or close date, currency and units, period conventions, entity perimeter, reporting basis and decision outputs. Define historical, forecast and stub periods. Use consistent period columns within comparable schedules.
2. Preserve raw historical inputs and map them to normalized model lines. Reconcile to the sources before forecasting. Show reported to adjusted bridges with one row per material adjustment.
3. Put editable assumptions in Inputs with rationale and source/as-of columns. Preserve historical data on source sheets. Put driver paths in Cases and use a documented selector, with a separate M&A selector if the dimensions can vary independently. Flag unsupported case IDs.
4. Forecast the operating drivers and statements. Use beginning/additions/subtractions/ending schedules for working capital, fixed assets, acquired intangibles, equity and debt as needed. Reconcile cash and debt before valuation. Use opening balances for interest when appropriate, or document and test a deliberate circularity.
5. Build the decision outputs. Reuse the operating model for DCF, LBO or deal returns. Show an executive summary, major return drivers and downside, with auditable sensitivities. Separate organic and acquired growth, and isolate timing, leverage and multiple effects when material.
6. Calculate, reopen, and verify. Compare headline outputs to an independent calculation from raw inputs and assumptions. Exercise adverse cases, verify all checks across periods, and restore the intended delivery case.

## Formula and layout rules

- Use short formulas and labeled intermediate rows. Put the key driver early in the expression when it improves tracing. Calculate totals and repeated links from their canonical source.
- Put economic assumptions and thresholds in cells. Structural literals such as zero, one and sign conversions are acceptable when self-explanatory. Give period bases and conventions explicit labels where they can vary.
- Use direct references and supported functions. Avoid volatile indirection, undocumented external links and formulas that preserve a prior answer through self-reference. Use explicit scenario formulas or a tested Excel data table with a reproducible refresh procedure.
- Match number formats to row units. Keep input, calculation and cross-sheet font roles consistent. Put source and rationale in the data area or in a separate source register. Keep explanatory workbook prose to the requested minimum and place delivery instructions outside the workbook.
- For public equity, usually start with an operating model, DCF and reverse DCF. Add an LBO for sponsor economics or financing analysis when relevant. Scale the number of tabs to the decision unless a reference layout explicitly requires more.
