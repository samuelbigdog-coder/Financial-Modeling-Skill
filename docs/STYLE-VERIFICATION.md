# Synthetic mini-model style verification

Reviewed `examples/Northstar_Motion_MiniModel.xlsx` against the [SFG style specification](../skills/financial-modeling/references/sfg-style.md). The review covered all worksheets structurally, representative rendered regions, and the exported Excel package. The SHA-256 of the reviewed workbook was `37732F40D2A2DD45B2B6A9F40A2D1072AFCB5113CA678331752031653EA26D3C`.

## Results

| Check | Observed |
| --- | --- |
| Workbook map | All 25 reference tabs retain their exact names and order, followed by one new gray Checks tab. Both white dividers have no populated cells. |
| View and tab palette | Reference group colors match across the 25 tabs. DS Backlog, M&A Targets and EM Backlog show gridlines. The other tabs suppress them. No frozen panes or merged cells in the workbook. |
| Style system | Calibri 11 body, compact gutters, 12 pt titles, 14 pt CF Walk title, navy and medium-blue header hierarchy, green internal links, yellow and blue editable cells, bold ruled totals, gray and lavender accents. P&L Detailed and Hist BS Detailed have Arial 8 raw blocks. Backlog, target and contribution tables use Century Gothic 10 where applicable. |
| Control panel | The exported file has two native list validations on Inputs E8:E9. The rendered control rows show boxed pale yellow cells, wrapped pale blue explanation/source areas and two decimal places for fractional input multiples. |
| SOFR chart | One native line chart begins at F13 and reads the local synthetic cells C15:D19. The imported-workbook render through row 30 showed the title, five year labels, a percent axis and the plotted series. |
| Package inspection | No workbook cell comments, hyperlinks, embedded images or external-link parts were found. The package scan found no selected proprietary source identifiers in worksheet, relationship or package XML. The source workbook was read independently and was not used as the output file template. |

All 26 sheets were rendered with the artifact tool after importing the saved workbook. Review regions covered the first 24 rows on most tabs, Inputs through row 18 and SOFR through row 30. These regions showed the intended group hierarchy, number formats and legible main tables. The sheet previews are a sample of the working area, and the package inspection covered all stored cells and worksheet metadata.

## Layout differences

- The example uses a compact shared period grid. Inputs retains the original title position and three-part controls/rationale/provenance pattern, with labels and controls shifted into columns C and E. The rationale column is 76 units wide versus approximately 102 in the source. Mini Model retains its D5 title and approximately 11-unit period columns, with row labels in C and a shorter three-year historical block.
- Cases and M&A Schedule stack case rows on one period grid. The reference places the four case blocks side by side. This preserves the example's active driver and scenario logic and shortens horizontal scrolling.
- CF Walk uses a single forecast cash walk. Exec Summary places transaction and return sections beneath the main operating grid. LR Performance and LR Data use a short synthetic history. The original wide scenario panels, right-hand supplements and long source exports are represented by compact teaching tables.
- The four reference-specific zoom values are absent. Deal Comps keeps a visible inclusion field and filtered statistics, with no native AutoFilter in the delivered sample. The SOFR chart sits below the synthetic rate assumptions, while the reference anchors its chart near F6.
- The reference picture was excluded. Header bands and the two divider tabs are built from ordinary cells. The original merged anchors are documented in the specification for future full-size implementations.
