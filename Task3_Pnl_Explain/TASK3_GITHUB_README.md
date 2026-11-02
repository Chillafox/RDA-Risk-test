# Task 03 — P&L Explain & Exception Monitoring

**Fixed Income Risk Dashboard · Rates Desk Analytics Portfolio**

> **Business question:** Can a risk-based P&L estimate explain the desk's reported daily P&L, and which residuals deserve investigation relative to current interest-rate risk?

This module extends the [RiskHub Fixed Income Risk Dashboard Python Lab](https://riskhub.org/python-labs/fixed-income-risk-dashboard). It builds a carry-plus-market P&L explain from yesterday's bond portfolio and an eleven-vertex yield curve, estimates the unexplained residual and assesses fixed versus risk-scaled exception thresholds. All market and trading records are **simulated**.

## Results

| KPI | Result |
|:--|--:|
| Reporting mornings | 518 (258 in 2023; 260 in 2024) |
| 2023 reported P&L RMS | USD 517,683 |
| 2023 unexplained RMS — market-only prediction | USD 74,011 |
| 2023 unexplained RMS — carry + market prediction | **USD 16,615** |
| 2023 unexplained RMS — full-revaluation proxy | USD 16,161 |
| Fitted vertex-level noise, using 2023 | **0.0963 bp** |
| 2024 fixed-limit flags at 3.5 sigma | **49** |
| 2024 risk-scaled flags at 3.5 sigma | **34** |
| Largest absolute 2024 Z-score | **78.0, 2024-11-20** |

The 15 fewer scaled flags reflect a different exposure-aware screening model; they are **not** proof that the corresponding observations were false alarms.

## Methodology

1. Rebuild the previous-day positions from amended trade tickets.
2. Compute carry using the aged portfolio value and interim bond cash payments.
3. Estimate first-order market P&L from yesterday's Key-Rate DV01 and today's curve changes.
4. Compute unexplained P&L as reported minus predicted P&L.
5. Estimate vertex-level mark noise using 2023, then evaluate standardized residuals for 2024.
6. Deliver detailed attribution and exception review through a C#-generated Excel workbook.

**Two complementary explanations:** The full-revaluation proxy helps separate the reported-versus-model **data gap** from the full-revaluation-versus-linear **model gap**. Its 2023 RMS differences are about USD 16,161 and USD 3,128 respectively; RMS components are not additive.

## Deliverables

| File | Purpose |
|:--|:--|
| [pnl_explain.ipynb](pnl_explain.ipynb) | Portfolio notebook with segmented commentary, formulas, tests and the 2024 Z-score chart |
| [DeskRiskPnL_Task3.cs](../Excel_Template_Design/DeskRiskPnL_Task3.cs) | C# / ClosedXML Excel generator |
| `Excel_Template_Output/task3_snapshot.json` | Generated intermediate JSON, excluded from Git |
| `Excel_Template_Output/RatesDesk_PnLExplain_Task3_2024-12-31.xlsx` | Generated workbook, excluded from Git |

Excel worksheets: **01 PnL Overview**, **02 Daily Explain**, **03 Morning Drilldown**, **04 Tenor Attribution**, **05 Exception Monitor**, **06 Methodology**, and a hidden data sheet. The Morning Drilldown supports an editable reporting date and sigma threshold. A static Python image does not change when the Excel threshold changes.

## Run locally

- Python 3 with `numpy`, `pandas`, `matplotlib`, `jupyter`, and `ipython`.
- Obtain the authorized educational feed from the original lab and place it at `DataSource/rates_desk_feed.csv` in the project root. It is omitted from the repository pending redistribution permission.
- Open `Task3_Pnl_Explain/pnl_explain.ipynb` and run all cells. It searches upward from the working directory for `DataSource`; `RAD_RISK_ROOT` can also specify the root explicitly.
- To generate the Excel workbook, install .NET 10, retain `Excel_Template_Design/DeskRiskPnL_Task3.cs`, then run the optional final code cell. The C# generator has environment-specific output settings; check these before running elsewhere.

## Important limitations

- Simulated fixed-income desk: not real positions, certified prices or an official P&L ledger.
- First-order DV01 market P&L is not full nonlinear repricing.
- Noise calibration assumes independent vertex-level daily mark differences; an extreme Z-score is a **review trigger**, not proof of a specific incident.
- Does not explicitly explain trade execution, funding, liquidity, credit-spread, accounting or transaction-cost P&L.
- Tasks 4–5 provide incident forensics and policy evaluation; they are not completed by this Task 3 screening model alone.

**Stack:** Python, NumPy, pandas, Matplotlib, C#/.NET 10, ClosedXML, Excel.
