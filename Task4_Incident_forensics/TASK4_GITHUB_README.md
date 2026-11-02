# Task 4 — Incident Forensics & Risk Controls

**Fixed Income Risk Dashboard · Python / C# / Excel · Portfolio module**

> **Business question:** How can a rates desk flag suspect pricing or booking inputs, investigate discrepancies against reported P&L, and measure the effect of confirmed corrections on risk?

This module builds on Task 1 (portfolio valuation / DV01 / VaR), Task 2 (daily risk attribution), and Task 3 (P&L explain). It presents screening as a **prospective** control and confirmed incident evidence as a **separate retrospective** evaluation.

## At a glance

| 2024 report mornings | Retrospective incident cases | Fixed / scaled P&L flag days | Fully corrected year-end VaR |
|:--|--:|:--|--:|
| **260** | **6** | **49 / 34** | **USD 1,460,314** |

The as-received VaR was **USD 939,306**. Rebuilding risk with corrected quotes and trade history produces a higher year-end estimate of **USD 1,460,314** (35.7% understatement relative to corrected VaR). This is a *simulated retrospective correction* rather than an observed trading loss.

## Screening logic

| Screen | Input | Alert criterion | 2024 flagged days |
|:--|:--|:--|--:|
| `PNL_FIXED` | Reported minus predicted P&L | More than 3.5 × 2023 unexplained-P&L standard deviation | 49 |
| `PNL_SCALED` | Residual divided by risk-scaled expected noise | Absolute Z-score greater than 3.5 | 34 |
| `STALE` | Daily movements of each zero-rate vertex | At least three consecutive movements below 0.02 bp | 1 |
| `KINK` | Vertex relative to adjacent maturities | Change in local curve kink greater than 8 bp | 2 |
| `LATE` | First booked date versus traded date | First booking occurs after the trade date | 1 |

**Count distinction:** The `KINK` rule finds **six vertex-level jumps on two dates**; `STALE` identifies one run that triggers one alert date. Alert counts are not interchangeable with incident counts.

## Investigation evidence

The six simulated operations-review cases cover wrong-side booking, incorrect notional, late booking, completely missed trading activity, stale yield data, and a bad market print. The notebook compares first-hit dates and inferred DV01 gaps with the retrospective case records. **Case labels, missing tickets and corrected vendor quotes are not used to construct detection flags.**

## Workbook deliverable

Python exports `task4_snapshot.json`; C# / .NET 10 and ClosedXML generate the Excel report.

- `01 Overview` — summary and detection counts
- `02 Daily Screens` — report-date screening grid
- `03 Investigation` — date selection and observable evidence
- `04 Evidence Register` — market-data and booking exceptions
- `05 Case Review` — retrospective incident windows and first-hit delays
- `06 VaR Impact` — risk after verified feed / book corrections
- `07 Methodology` — definitions and limitations

The C# source is stored at [../Excel_Template_Design/DeskRiskForensics_Task4.cs](../Excel_Template_Design/DeskRiskForensics_Task4.cs). Generated JSON and workbooks belong in `../Excel_Template_Output/` and are not committed.

## Repository structure

```text
Fixed Income Risk Dashboard/
├── DataSource/                           # local, not committed
├── Task4_Incident_Forensics/
│   ├── incident_forensics.ipynb
│   └── TASK4_GITHUB_README.md
├── Excel_Template_Design/
│   └── DeskRiskForensics_Task4.cs
└── Excel_Template_Output/                # generated, gitignored
```

## Reproduce the analysis

1. Provide the authorized practice input file as `DataSource/rates_desk_feed.csv` at the repository root; source data is not redistributed by this project.
2. Run the notebook from the repository or `Task4_Incident_Forensics` directory. It searches parent folders for the data; alternatively set `RAD_RISK_ROOT` to the project root.
3. Execute the Python cells in order to compute prospective alerts, retrospective analysis, figures and JSON output.
4. Install .NET 10 SDK and run the final C# cell to generate the Excel workbook. The Python analytics work without .NET.

## Mathematical conventions

- Zero yields are continuously compounded percentages; curve differences are converted to basis points.
- Unexplained P&L is reported P&L less carry plus first-order market P&L prediction.
- Risk-scaled P&L uses a 2023-fitted noise factor and prior-day Key-Rate DV01 norm.
- End-of-day 99% Historical Simulation VaR uses the fifth-worst of 500 linearly approximated one-day curve-shock P&Ls.

For formulas, diagnostic results and figures, see [incident_forensics.ipynb](incident_forensics.ipynb).

## Limitations & attribution

The work uses simulated data and a predefined retrospective incident log. Alerts may be benign and cases may be missed. Detection coverage in the six-case exercise is not evidence of generalization to live markets. No execution or automated trade blocking is performed.

Educational exercise based on the [RiskHub Fixed Income Risk Dashboard Python Lab](https://riskhub.org/python-labs/fixed-income-risk-dashboard), expanded with forensic screening, operational investigation reporting and a C# Excel delivery layer.
