# STARWAR worksheet style specification

The local STARWAR mini-model was inspected read-only across all 25 worksheets. This specification records reusable layout and formatting only. Rebuild from a blank workbook with synthetic labels, narrative and financial data. Leave the reference workbook, its embedded picture, notes, names, comments, links, chart series and underlying data outside distribution.

## Shared visual system

| Element | STARWAR convention | Implementation |
| --- | --- | --- |
| Canvas | White, with compact bands of data and intentional whitespace | Hide gridlines on 22 worksheets. The three backlog/target exceptions appear below. Use a 15 pt default row height, 5 pt section spacers on dense output tabs and 30 pt rows for wrapped rationale. |
| Type | Calibri 11 body and 12 bold title, with a larger 14 pt CF Walk heading | Preserve Arial 8 source-data blocks on detailed financial statements and Century Gothic 10 on specialist source extracts. Use Calibri if Century Gothic is unavailable and adjust column widths for fit. |
| Major header | White bold type on deep navy `#002060`; CF Walk and SOFR use `#1F497D` | Apply the fill across the entire logical block. Use `#4472C4` for selected executive summary bands and `#156297` or `#00B0F0` for select source subheaders. |
| Inputs | Blue `#0000FF` type on pale yellow `#FFFFCC` with thin outline | Reserve the boxed yellow treatment for editable assumptions and selectors. Many raw source values are blue with plain background, and some imported sections retain black. Use bright yellow `#FFFF00` sparingly for source follow-up cues. |
| Formulas | Local calculations black; in-workbook links green | The template uses `#008000` and `#00B050`. Prefer `#008000` for compact linked text and values. White header text takes precedence where necessary for contrast. |
| Notes and detail | Gray italic helper text, lighter type for ratios, and indented subordinate lines | Use `#595959` or `#808080` where legible. Keep a visible unit line near the heading. Distinguish operating assumptions, linked actuals and calculated output. |
| Accent fills | Soft gray `#F2F2F2`, pale blue `#DCE6F1` and `#DDEBF7`, muted lavender `#E6E0EC` | Apply in the same roles as the tab-level notes below. Pale green and orange are occasional forecast or scenario accents. Keep white space dominant. |
| Lines | Thin black top rules for totals and thin bottom rules under headings | Use a stronger bottom rule selectively in imported/source table headings. Prefer open tables over a full grid. |
| Dimensions | Narrow gutters separate broad blocks | Common gutters 2–3 Excel width units; many label columns 28–38 units and annual numeric columns 9–12 units. Use a wider numeric column where dates or large numbers would clip. |
| Tab colors | Dark blue `#1F497D` for Inputs and Mini Model; muted red `#C0504D` for Cases and M&A Schedule | White `#FFFFFF` for the two dividers, bright red `#FF0000` for the seven output tabs, and orange `#F79646` for the twelve source tabs. |
| Views | Visible worksheets with no frozen panes in the reference | Use reference zoom where given. Add panes only as an explicit usability adaptation for a longer sheet. |

### Number formats

Right-align numeric values and align periods horizontally across related sections. Use true numeric values for periods, percentages and multiples so formulas can reference them. Example Excel formats use the period/decimal convention of the reference.

| Purpose | Format |
| --- | --- |
| Engine and output amounts | `#,##0.0_);(#,##0.0);"-"_)` |
| Source whole amounts | `#,##0_);(#,##0);"-"_)` |
| Percentage | `0.0%_);(0.0%);"-"_)` |
| Multiple | `0.0"x"` or `0.00"x"` |
| Dates and rates | `mmm-yy` for backlog columns, `mm-dd-yy` for curve/deal dates, and `0.00%` for rates |
| Year headers | `0000"A"` and `0000"E"` when period status is carried in a display suffix |

Use parentheses for negative amounts and dashes for displayed monetary zero. A valid zero and an unavailable input require distinct representations. Use precise units on every block that changes scale, such as source amounts in thousands and model amounts in millions. Where a check is diagnostically meaningful, display its unrounded result at suitable precision.

### Merged-header replacements

The reference contains 12 merges across three sheets. Preserve each header's visual grouping and use an ordinary anchor cell with blank neighbors. For a centered band in Excel, apply **Center Across Selection** (OOXML `centerContinuous`) to the anchor and each adjacent blank cell. Fill the band and apply its exterior borders cell by cell. If the target engine does not preserve center-across, align the heading left and retain the continuous fill and group boundaries. Keep child column headings independent.

| Sheet | Source bands to replace |
| --- | --- |
| Mini Model | I26:M26, I81:M81, I109:M109, I118:M118, I140:M140, I190:M190 |
| Cases | D3:H3 |
| EM Backlog | G6:N6, O6:AD6, AF6:AH6, H42:I42, K42:M42 |

## Tab-by-tab map

Coordinates refer to the reference's layout. Preserve the order, role and relative block geometry in a synthetic mini-model. Shorten histories and sample tables where the synthetic case does not need the original span. Very distant formatted cells in the source are outside the functional layout.

| # | Tab | Placement and style | Content role for a clean rebuild |
| ---: | --- | --- | --- |
| 1 | Inputs | B2 title. B5/H5 navy headings, B about 34, D about 12, F about 14, G about 102, H about 38, with A/C/E about 2. Blue type on boxed pale yellow cells. Pale blue explanation bands and wrapped rationale/source text. 80% zoom. | Group deal, financing, operating and exit assumptions with an editable value, rationale and provenance. Identify synthetic inputs and units in adjacent columns. |
| 2 | Mini Model | D5 title, D6 italic units, D about 34, E:M about 11, N about 2, O about 30 for model notes. H9:H10 hold list selectors in the reference. Navy bands, isolated pale yellow selectors, gray subtotals and forecast-specific light accents. 85% zoom. | Put close mechanics above annual income statement, balance sheet, cash flow, debt and returns. Keep historical and forecast columns visually distinct. Replace the reference picture with an original text heading if needed. |
| 3 | Cases | B2 title and B about 34. Four five-column period groups starting at D, J, P and immediately after the final gutter. I/O/U gutters. Navy band, boxed pale yellow input grid and a narrow validation-highlight area. | Label the four operating cases above their driver grids. Drive active model assumptions through the authoritative selector. |
| 4 | M&A Schedule | B2 title and B about 34. Four side-by-side blocks start D, L, S and Z. Year columns about 9. Use small gutter pairs and navy blocks, with a small boxed editable region. | Lay out acquired EBITDA, associated sales and cost, and explicit timing by acquisition case. |
| 5 | Output --> | White tab and blank grid-free sheet. | Navigation divider. |
| 6 | Equity Bridge | B2 title and B about 30. C:P mostly 10–19 wide. White canvas, bottom rules for column headings and pale lavender assumption/stage highlights. 95% zoom. | Separate beginning equity, staged value drivers and ending equity. Include the bridge tie as a diagnostic. |
| 7 | Fan of Outcomes | C2 title; C about 30 and D:I about 15. Thin bottom rules and mostly white fill. | Use horizontal case columns for comparable EBITDA, exit equity, MOIC and IRR. Supply calculated results or an explicitly documented refresh workflow. |
| 8 | CF Walk | B2 title in 14 pt, B about 28. Dark blue year/group bands, numeric groups beginning around E, K and Q separated by very narrow gutters, and 5 pt spacer rows. 85% zoom. | Show year-by-year cash conversion and a distinct comparison/variance panel. Use defined signs consistently across blocks. |
| 9 | Returns Scenarios | B2 title; B about 20; C:G about 12; separate right-hand support blocks. Navy bands, boxed pale yellow input cells and limited warm accent. | Separate active returns, exit-year progression and the two-way sensitivity matrix. Format percent and multiples distinctly. |
| 10 | LR Performance | B2 title; B about 26; C:X about 8.4; years form an uninterrupted horizontal strip. White body with top and bottom rules and smaller notes. | Show long-run performance and growth/margin trends with traceable historical links. |
| 11 | Segment Margins | B2 title, B about 34, C:F about 12; H begins a second display region. White blocks, underlined headers and restrained purple annotation accents. | Show segment economics and margin build with consistent segment order and a consolidated tie. |
| 12 | Exec Summary | B2 title, B about 28, C:K about 10.7, L about 3, M about 21. Navy and medium-blue group bands, soft gray totals and approximately 5 pt section spacers. | Put the operating trajectory on the left and a compact returns/transaction summary on the right. Use only traceable links and concise labels. |
| 13 | Sources---> | White tab and blank grid-free sheet. | Navigation divider. |
| 14 | P&L Detailed | C3 italic introduction, C about 38; navy band around row 6; D:G main financial block, with right-hand source/adjustment blocks and gutters. Calibri model view and Arial 8 source detail. | Distinguish raw reported data, adjustment classes and adjusted values. Make the reported-to-adjusted bridge explicit. |
| 15 | DS Backlog | Gridlines visible by default; source panel starts around G2 with G about 33. Century Gothic 10, cyan band and gray total rows. | Show backlog roll-forward/coverage by customer or market grouping using fully synthetic records. |
| 16 | M&A Targets | A1 title; gray field header around row 6 and taller data rows. A:F about 33/15/25/38/17/26. Century Gothic 10, visible gridlines by default, yellow follow-up cells. | Use a short synthetic target list with financial assumptions kept separately traceable. |
| 17 | EM Backlog | B2 title; B about 24, D about 18, E about 22; G:N and O:AD dense numeric blocks with header bands, additional block after AE gutter. Compact 12.75 pt default, Calibri/Century Gothic 10, visible gridlines by default, blue group accents and stronger bottom rules. | Use a compact project/date matrix with aggregated backlog. Retain month and amount formats. Replace each source merge with a flat band. |
| 18 | Hist BS Detailed | C3 italic introduction, C about 38; navy band around row 6; D:G history, with a gutter then detail blocks starting I and O. Calibri model view, Arial 8 source detail, gray and ruled totals. | Show historical accounts, subtotals and supporting source splits with a balance proof. |
| 19 | SOFR | Date/rate labels around D6:E6 and a native line chart anchored near F6. Dark blue period row, white canvas and two-decimal rate formatting. | Use a synthetic date/rate curve and annual inputs derived from the specified aggregation. Keep the chart data local. |
| 20 | Contr. Margin | B2 title, compact parallel tables, blue/cyan group headers, gray totals, mixed Century Gothic 10 and Calibri 11. | Show representative job-level contribution margins with weighted margin from aggregate revenue and contribution. |
| 21 | LR Data | B2 title, B about 30 and a continuous 9-unit period strip from C; compact white data block and simple rules. | Store traceable synthetic long-run revenue and earnings history for LR Performance. |
| 22 | Margin Explored | B2 and L2 act as left/right block anchors; B about 27, L about 26, Q about 1 gutter. Mixed Calibri/Century Gothic, gray rows and navy/teal headers. | Compare multiple segment-margin views with distinct underlying cohorts or assumptions. |
| 23 | Comps | B1 title, B about 40, C about 2 and D:I about 10–14. Navy header near row 3, pale gray/light blue repeated rows, broad notes region on the right. | Show synthetic company peer multiples with segment grouping, comparability notes and summaries. |
| 24 | Deal Comps | B2 title, B about 46, date C about 11, D:E about 13, narrative G/H about 36/31, short metric columns I:R. White filterable table, thin header rules and a small boxed yellow criteria block. | Show synthetic precedent transactions, explicit inclusion criteria and filtered statistics. Keep the native AutoFilter on the record table. |
| 25 | Capex | Navy heading near B3; B about 14, C about 2, D:H about 9 and a wide right-hand notes/supplementary block beginning O. Calibri core, Century Gothic detail, thin period rules, blue source inputs and green links. | Distinguish maintenance, growth and total capex with an auditable history-to-forecast link. |

## Release acceptance

- Check the 25 required tab names and relative order, blank dividers, tab colors and zero merged ranges in the delivered workbook.
- Inspect header legibility and widths at the selected zoom. Check the main amount/percentage formats, negative signs, historical/forecast boundary and treatment of active assumptions.
- Use workbook styling as a navigation aid. Rebuild calculations independently and test linkage, reconciliation, cases and sensitivities. Historical layout choices carry no authority over model mechanics.
- Confirm that the shipped workbook contains only synthetic records, original text and internally generated charts. Exclude proprietary data, original images, imported comments, external links and reference-workbook caches.

