# Task 02 — Daily Risk Attribution & VaR Walk

**Fixed Income Risk Dashboard · Simulated Rates Desk · Python + C# + Excel**

> **Business question:** Yesterday's Value at Risk differs from today's. Was the change driven by portfolio aging, market moves, changes in the historical VaR window, or trades and booking amendments?

## At a glance

| Metric | Result |
|:--|--:|
| Risk reports reviewed | **518 mornings** (2023–2024) |
| Closing Net DV01 (2024-12-31) | **+&#36;47,234 per bp** |
| Closing 1-day 99% Historical VaR | **&#36;939,306** |
| Change in VaR during 2024 | **−&#36;497,829** |
| Risk attribution reconciliation | **PASS** |

## How it works

1. Reconstruct the daily bond book from the amended trade blotter and daily zero-rate curves.
2. Build four Key-Rate DV01 ladders to separate **Time**, **Market** and **Trades**.
3. Recalculate Historical Simulation VaR through five states, additionally isolating **Window Roll**.
4. Reconcile daily opening risk and the sum of attributed changes against closing risk.
5. Export daily, annual and tenor-level results to JSON and generate an interactive workbook using C# / ClosedXML.

**2024 VaR bridge:** Time/Aging **−&#36;444,035** · Market **+&#36;17,383** · Window **−&#36;104,467** · Trades **+&#36;33,290**.

## Files

| File | Purpose |
|:--|:--|
| [`risk_walk.ipynb`](risk_walk.ipynb) | **Start here:** calculations, visuals, validated results and investigation cases |
| [`Excel_Template_Design/DeskRiskAttribution_Task2_V2.cs`](../Excel_Template_Design/DeskRiskAttribution_Task2_V2.cs) | Generates the Excel dashboard with date drilldown, VaR bridge and reconciliation checks |
| `Excel_Template_Output/task2_snapshot.json` | Local generated interface between Python and C#; not committed |
| `Excel_Template_Output/*.xlsx` | Local generated reports; not committed |

### Running locally

- Install Python with `numpy`, `pandas`, `matplotlib`, and Jupyter.
- Place the authorized dataset at `DataSource/rates_desk_feed.csv`.
- Run Jupyter from the project root or Task 2 folder; the notebook searches parent folders for the dataset. Alternatively, set `RAD_RISK_ROOT` to the repository root.
- Run the notebook from top to bottom to calculate risk and export the JSON snapshot.
- **Optional Excel generation:** install .NET 10 SDK, then run the final C# publishing cell. Verify/configure the existing C# output directory on a new machine.

## Practical interpretation

- **A large VaR change is an investigation signal, not automatically a trading loss or a control breach.**
- The **2024-04-16** `T019` amendment produces a **−&#36;1.355m** Trades contribution to VaR.
- The **2024-11-20** historical window roll contributes **+&#36;66,023** to VaR without a new trade effect.
- Reconciliation `PASS` validates model arithmetic, not independent market data or trade-book completeness.

## Scope and attribution

Simulated government bond book; 11 zero-rate vertices (3M–30Y); 500 historical daily curve changes; fifth-worst scenario used for empirical 99% VaR. The linear DV01 approximation excludes explicit convexity, spread, liquidity and funding risk. VaR attribution is sequential and depends on the state order.

Adapted and extended from [RiskHub — Fixed Income Risk Dashboard](https://riskhub.org/python-labs/fixed-income-risk-dashboard). This is an educational analytics and reporting project, not a production risk system or investment recommendation. Review source-data and starter-code redistribution rights before publishing.
