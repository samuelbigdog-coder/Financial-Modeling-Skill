# Modeling methodology

## Historical data and operations

Record source URL or internal reference, filing date, period end, currency, units and any conversion. Align fiscal calendars and treatment of trailing and stub periods before combining sources. Preserve reported figures and build separate adjustments for discontinued operations, acquisitions, accounting changes and nonrecurring items. Reconcile mapping totals and investigate gaps, duplicate keys and inconsistencies. Treat missing values explicitly.

Choose the smallest driver tree that explains the business. Decompose sales into volume, price and mix or locations and sales per location. Separate recurring and one-time streams and organic growth from acquired revenue. Bridge backlog to revenue using conversion, cancellations and timing. Benchmark contribution margin and long-run profitability. Separate fixed expenses, indexed with a relevant cost driver, from expenses that vary with volume. Backtest a historical period against a simple forecast before using a more elaborate approach.

## Statements and financing

- Use accounts receivable days on credit sales and inventory/payable days on an appropriate cost or purchases base. Explain any proxy or seasonal distortion. In cash flow, increases in operating assets consume cash and increases in operating liabilities provide cash.
- Roll PP&E from opening balance through capital expenditure, depreciation, dispositions and acquisition effects. Separate maintenance and growth investment where material. Roll intangibles, goodwill and deferred taxes separately when needed.
- Reconcile retained earnings to opening balance, net income, dividends and adjustments. Model share issuance, repurchases and dilution under the applicable convention when EPS or value per share matters.
- Compute book taxes and cash taxes separately when NOLs, deferred taxes, jurisdiction or interest deductibility is material. Include a simple current-tax floor only as a documented simplifying assumption.
- Calculate cash before financing, contractual amortization, facility draws to the minimum cash target subject to availability, and repayment in contractual priority. Use eligible borrowing base, commitment, existing draws and other restrictions to calculate usable headroom. Flag cash and covenant shortfalls explicitly. If modeled cash drops below zero, report the unfunded amount and the resulting infeasible case. Include PIK, OID, fees, maturities and rate floors where relevant.
- Use opening or known time-weighted debt for interest as a transparent default. If cash sweep and average debt must be solved together, prefer an algebraic solution or isolate the intentional loop and test settings, convergence, reset and off-state behavior.
- Bridge net debt from opening to closing using operating and investing cash flows, distributions and external equity, plus noncash debt changes. Define included debt and unrestricted cash consistently. Reconcile cash, debt and equity movements separately.

## DCF and reverse DCF

Forecast EBIT and cash taxes, add back noncash charges, subtract capex and incremental operating working capital, and isolate other material operating cash items. Match nominal or real cash flows, currency, discount rate and timing convention. Compute WACC from market-value weights, risk-consistent cost of equity and after-tax marginal borrowing cost. If financing changes materially, consider APV or time-varying rates.

For a perpetuity, use normalized year N+1 cash flows and require WACC above terminal growth. Test whether reinvestment supports terminal growth at the assumed incremental return on capital. For an exit multiple, match the multiple to the terminal metric and timing and benchmark implied growth or cash yield. Show terminal value share and both methods where meaningful. Bridge EV to common equity with cash, debt, leases, minority interests, dilution and other claims consistently with cash flow and comparables. In reverse DCF, vary one identifiable driver over defensible bounds, hold other assumptions explicit, and verify the implied price using the forward model. State the search interval, report when it does not bracket a solution, and identify materially different plausible roots if the response is nonmonotonic.

## LBO and M&A

Reconcile sources and uses at close. Include fees, refinanced debt, required cash, financing constraints and sponsor equity. Stress low and high entry prices and check minimum equity, leverage and facility caps. Separate sponsor equity from management rollover and minority interests.

Model acquisitions by close date and distinguish organic and acquired EBITDA, synergies and integration costs, acquired working capital, transaction consideration and financing. Include opening acquired balances and a purchase-price allocation when material. Reconcile uses, goodwill, cash and debt independently.

Model exit EBITDA and multiple, cash, debt and other equity claims at the exit date. Calculate MOIC from dated equity contributions and distributions, and verify XIRR by discounting flows at the result using the same day-count convention. Inspect multiple sign changes, invalid domains and absence of a root. Test exit-year steps, target return entry price, and entry × exit and growth sensitivities. In a return bridge, document sequencing and tie contributions to the headline change.

## Scenarios and comparables

Assign economic narratives to base, upside and downside paths. Model correlated driver changes within a case and keep independent dimensions separate. Distinguish a scenario from a one-factor or two-factor sensitivity that holds other assumptions fixed. Label sensitivity axes, units, step sizes and the exact anchor case. Guard invalid points and reconcile the center to the live output. Refresh data tables in a verified engine or write direct formulas.

Choose comparable companies and transactions by economic similarity, business mix, geography, size, growth and profitability. Align date, EV bridge and measurement period. Normalize lease, SBC, acquisition and unusual-item treatment with the numerator and denominator. Show exclusions and explain spread and outliers before relying on a median.

## Specialized modules

- For sale-leaseback, model sale proceeds, any gain, lease obligations and recurring cash rent separately. Verify the applicable sale and lease accounting, and align lease treatment between EBITDA, cash flows, leverage and enterprise value.
- For floating-rate debt, distinguish spot overnight rates, compounded historical rates and forward tenor assumptions. Model contractual floor, spread, reset dates and day count. Stress rates and operating performance together where plausible.
- For public filing data, retain accession, concept, dimensions, fiscal period and units. Verify custom tags and segment facts against the filing. Keep as-filed and restated histories distinct when backtesting.
- If a DCF includes acquired EBITDA, model the acquisition reinvestment required to support it. For a standalone organic DCF, exclude both future purchases and their associated economics.
