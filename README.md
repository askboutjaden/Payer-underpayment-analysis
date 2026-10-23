# Payer Underpayment Analysis (Excel)

Identifies claims where payers paid less than the contracted rate, then quantifies underpayment by payer, CPT code, and service month. Built in Excel with Power Query, XLOOKUP, SUMIFS/COUNTIFS, PivotTables, PivotCharts, and slicers.

**Headline finding: 1,459 paid claims (33.8%) were underpaid by a total of $107,894.87, or 3.6% of expected revenue.**

![Dashboard](dashboard.png)

---

## Business question
Revenue cycle teams need to know which payers are paying below contract, on which services, and whether the problem is getting better or worse. Underpayments are often missed because the claim shows as "Paid."

## Data
Synthetic dataset (same source as my SQL + Power BI claims/denials dashboard):

| Table | Rows | Key fields |
|---|---|---|
| claims | 6,094 | claim_id, payer_id, cpt_code, units, service_date, allowed_amount, paid_amount, claim_status |
| payers | 8 | payer_id, payer_name, payer_type |
| cpt_codes | 12 | cpt_code, description, standard_charge |
| contract (built) | 8 | payer_id, contract_rate |

**Assumed fee schedule:** contract rates are expressed as a % of standard charge (e.g., Medicare Part B = 36%, Florida Medicaid = 27%, Self-Pay = 100%). The payer table did not include rates, so I built a separate contract table, the way fee schedules are kept apart from the payer master in real RCM systems.

## Method
1. **Load and validate.** Imported CSVs with Power Query and set data types (codes as text so J-codes such as J0702 aren't broken). Built a Checks sheet that reconciles row counts, status counts, and dollar totals to the source files to the cent (17 checks, all passing).
2. **Expected payment.**
   - `std_charge`: XLOOKUP of CPT standard charge
   - `contract_rate`: XLOOKUP of payer contract rate
   - `expected_payment = ROUND(std_charge × units × contract_rate, 2)`
3. **Variance.** `expected_payment − allowed_amount`, for **Paid claims only**.
   - Compared against *allowed* (not paid), because paid excludes patient responsibility. Using paid would falsely flag every claim with coinsurance.
   - Denied and pending claims are excluded; they have not been adjudicated for payment.
4. **Underpaid flag.** Variance greater than **5% of expected payment**. The tolerance keeps rounding noise out of the work queue.
5. **Analysis.** PivotTables of underpaid claim count and dollars by payer, CPT code, and service month.
6. **Dashboard.** KPI block (SUMIFS/COUNTIFS), three combo PivotCharts (dollars as bars, claim count as a line), and clinic and payer slicers connected to all charts.

## Key findings
| KPI | Value |
|---|---|
| Paid claims | 4,321 |
| Expected payment (paid claims) | $2,968,252.76 |
| Underpaid claims | 1,459 (33.8%) |
| Underpaid dollars | $107,894.87 |
| Average underpayment per claim | $73.95 |
| Revenue leakage | 3.6% of expected |

1. **One drug drives most of the dollars.** Infliximab (J1745) accounts for **$80,360 (74%)** of underpaid dollars across only 80 claims. Auditing high-cost specialty drug claims gives the best recovery per hour of work.
2. **Commercial payers lead.** Summit Health Commercial ($33,646) and Blue Horizon Commercial ($29,896) make up **59%** of underpayment. Blue Horizon has more underpaid claims (376), while Summit's underpayments are larger per claim.
3. **Low-dollar, high-volume codes are a different problem.** 99214 has the most underpaid claims (288) but only $2,628 in variance. That suggests an automated fix (a contract load error) rather than claim-by-claim appeals.
4. **The recent drop is not real improvement.** September 2026 shows only $414 underpaid because most recent claims are still pending, not because payers improved.

## Recommendations
- Prioritize underpayment appeals for J1745 claims with Summit Health and Blue Horizon.
- Verify that commercial contract rates for J1745 are loaded correctly in the payer systems.
- Re-run the analysis monthly; exclude the most recent 60 days to avoid pending-claim bias.

## Workbook tour
| Sheet | Purpose |
|---|---|
| Dashboard | KPIs, charts, slicers |
| Claims | Claim data with the calculated columns |
| Payers / CPT / Contract | Lookup tables |
| Pivot_Payer / Pivot_CPT / Pivot_month | PivotTables behind the charts |
| Checks | Data validation and reconciliation |

## Tools
Excel (Microsoft 365): Power Query · XLOOKUP · SUMIFS / COUNTIFS · PivotTables · PivotCharts · Slicers · Conditional formatting

## Related projects
- SQL + Power BI Claims & Denials Dashboard (same dataset)
