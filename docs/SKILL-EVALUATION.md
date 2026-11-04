# Skill application evaluation

November 3, 2026. A separate agent manually applied the revised financial-modeling SKILL.md and its methodology and verification references to two synthetic adversarial prompts. The cases were graded against the pass criteria already in verification.md. This is a text-based application evaluation. Workbook authoring, Excel recalculation, visual conformance and Cowork installation or trigger behavior are outside its scope.

## Capped revolver and liquidity gap

**Prompt.** Repair a one-period liquidity model, in USD millions. Opening cash is 10, opening term debt is 50 and opening revolver is 12 against a commitment of 20. The minimum cash target is 5. Operating and investing cash flow totals negative 15 before 2 of mandatory term amortization. Compare against the same case with a 30 commitment. Assume the entire unused commitment is available and no other flows.

| Schedule line | Formula or calculation | Commitment 20 | Commitment 30 |
| --- | --- | ---: | ---: |
| Cash before revolver | Opening cash + operating and investing cash flows − mandatory amortization | −7 | −7 |
| Available commitment | MAX(0, commitment − opening revolver) | 8 | 18 |
| Required draw | MAX(0, minimum cash − cash before revolver) | 12 | 12 |
| Actual draw | MIN(available commitment, required draw) | 8 | 12 |
| Closing cash | Cash before revolver + actual draw | 1 | 5 |
| Liquidity gap | MAX(0, minimum cash − closing cash) | 4 | 0 |
| Closing term debt | Opening term debt − mandatory amortization | 48 | 48 |
| Closing revolver | Opening revolver + actual draw | 20 | 24 |
| Closing gross debt | Closing term debt + closing revolver | 68 | 72 |
| Closing net debt | Closing gross debt − closing cash | 67 | 67 |

Independent check. Opening net debt is 52. With no other economic flows, the negative 15 of operating and investing cash flow increases net debt by 15 to 67. Direct financing cash flows are +6 and +10 respectively, and opening cash plus all flows produces closing cash of 1 and 5. Closing debt independently equals opening gross debt 62 plus those same financing flows. The stressed draw reaches commitment while the 4 cash gap remains visible.

**Grade.** Pass for visible shortfall, cash tie, debt roll-forward and capped financing. The instructions are sufficient for the stated contractual terms. An additional edge case would test a borrowing base or letters of credit that reduce usable availability below the commitment, and a severe case that produces negative modeled cash. A release should flag the latter as an unfunded forecast and require an explicit resolution or exception.

## DCF, reverse DCF and invalid growth

**Prompt.** Build a one-year DCF with a WACC-by-growth grid in USD millions and solve for the growth that produces a target common equity value of 216.493506. Year-one unlevered cash flow is 10, year-one NOPAT is 20 and steady-state incremental ROIC is 20%. Use year-end discounting, zero other equity adjustments and net debt of 20. Base WACC is 10%, base growth is 3% and the reverse DCF holds WACC at 10%. Use grid WACCs of 8%, 10% and 12%, with growth columns of 2%, 3% and 4%. Add a 10% growth stress point.

Build normalized year-two NOPAT as 20 × (1 + g), reinvestment as year-two NOPAT × g / 20%, and year-two cash flow as their difference. If WACC is above growth, discount both the year-one cash flow and year-one terminal value to time zero. Otherwise return an explicit invalid status before dividing.

| WACC | EV at 2% g | EV at 3% g | EV at 4% g | EV at 10% g, stress |
| --- | ---: | ---: | ---: | ---: |
| 8% | 292.592593 | 333.518519 | 394.444444 | Invalid |
| 10% | 217.727273 | 236.493506 | 261.212121 | Invalid |
| 12% | 172.857143 | 182.638889 | 194.642857 | 500.000000 |

The base case gives year-two NOPAT of 20.60, reinvestment of 3.09, year-two free cash flow of 17.51, and terminal value at year one of 250.142857. Its present value is 227.402597, about 96.2% of total EV, and its addition to 9.090909 of explicit-period present value ties to the center EV of 236.493506. Subtracting 20 of net debt produces the target equity value.

Independent reverse check. An enterprise value of 236.493506 implies a terminal value at year one of approximately 250.142857 after compounding and subtracting the year-one cash flow. At 2% and 4% growth, the forward equity values of 197.727273 and 241.212121 bracket the target. Solving within that interval for 20 × (1 + g) × (1 − g/20%) / (10% − g) equal to the implied terminal value yields 3.00% growth. Feeding that value forward produces the target within 0.000001 million. At positive cash flows, the 3% column declines as WACC rises. At 10% growth, WACC below or equal to growth is invalid. The valid 12% WACC stress result is mechanically high and highlights sensitivity to a narrow spread.

**Grade.** Pass for a reconciled center, invalid domain handling, growth-funded terminal reinvestment and an implied growth that ties when run forward. The references call for defensible bounds and a forward tie. A stronger acceptance test would require the agent to confirm monotonicity over the chosen bracket, or disclose multiple admissible solutions and report an unbracketed target as unresolved. This evaluation uses an economically bounded local bracket of 2% to 4%.

