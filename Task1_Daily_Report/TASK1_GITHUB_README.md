# Task 01 — Daily Rates Desk Risk Report

**Fixed Income Risk Dashboard · Python / NumPy / pandas / C# / ClosedXML / Excel**

> **Question:** What is the desk holding, what is it worth, and how much could the portfolio lose under historical interest-rate moves?

This module reconstructs a simulated government-bond portfolio from a trade blotter, values each remaining coupon/principal cash flow against an eleven-vertex zero-coupon curve, calculates analytical Key-Rate DV01, and estimates **1-day 99% Historical Simulation VaR** using 500 daily curve shocks.

The full implementation and **executed charts** are in [`daily_report.ipynb`](daily_report.ipynb). The interactive Excel reporting generator is [`Excel_Template_Design/DeskRiskExcel_V3.cs`](../Excel_Template_Design/DeskRiskExcel_V3.cs).

## Closing Snapshot · 2024-12-31

| Metric | Result |
|:--|--:|
| Live bond positions | **17** |
| Net marked bond value | **+&#36;324,335,641** |
| Net Key-Rate DV01 | **+&#36;47,234 per bp** |
| Gross *position* DV01 (absolute sum) | **&#36;600,193 per bp** |
| 1-day 99% Historical Simulation VaR | **&#36;939,306** |
| Worst modeled historical scenario | **−&#36;1,424,133** (2023-08-10 curve move) |
| Parallel-only scenario counterfactual VaR | **&#36;755,236** |

**Interpretation:** Net DV01 is much smaller than gross position risk, so a parallel-rate summary hides offsetting positions. The 7Y and 3Y curve vertices contain large opposing sensitivities; historical non-parallel shocks are therefore important.

## Design & Calculation

| Component | Method |
|:--|:--|
| Book reconstruction | Last booked version per ticket replaces earlier versions; aggregate by bond |
| Cash flows | Fixed coupons paid semiannually plus principal at maturity |
| Present value | Continuously compounded zero-coupon discount factors, ACT/365 time to payment |
| Yield interpolation | Linear between 11 tenors (3M–30Y); flat extrapolation |
| Key-Rate DV01 | Analytical PV sensitivity to a 1bp decrease in one curve vertex |
| Historical VaR | Apply 500 historical curve moves to the current 11-element DV01 ladder; take fifth-worst P&L |
| Validation | Independent ±1bp central-difference repricing, position-to-ladder reconciliation |

**Model equations** (yields enter the PV calculation in decimal form):

$$
PV=\sum_i CF_i e^{-z(\tau_i)\tau_i},
\qquad
P\&L_s\approx-\sum_k DV01_k\,\Delta y_{s,k}^{bp}.
$$

The workbook adds a **macro-free, formula-driven Stress Lab** for editable tenor shocks, but it is **not** a full repricer and does **not** recompute Historical VaR when the user changes a stress scenario.

## Repository Layout

```text
RDA Risk test/
├── Task1_Daily_Report/
│   ├── daily_report.ipynb        # Notebook to read on GitHub
│   └── TASK1_GITHUB_README.md
├── Task2_Risk_Walk/
│   ├── risk_walk.ipynb           # Task 2: Daily risk attribution
│   └── TASK2_GITHUB_README.md
├── Excel_Template_Design/
│   └── DeskRiskExcel_V3.cs        # C# report generator
├── DataSource/
│   └── rates_desk_feed.csv        # Local input; not distributed
└── Excel_Template_Output/
    ├── task1_snapshot.json       # Generated
    └── RatesDesk_RiskWorkstation_V3_2024-12-31.xlsx  # Generated
```

The `DataSource/` and `Excel_Template_Output/` directories can be excluded from version control. The notebook stores its analytical charts as outputs so the results remain visible in GitHub's notebook viewer.

## Running Locally

1. Use Python 3.10+ with `numpy`, `pandas`, `matplotlib`, `ipython`, and `jupyter` installed.
2. Place an authorized copy of the simulated lab input at `DataSource/rates_desk_feed.csv`. The source feed is **not** distributed in this repository.
3. Open `daily_report.ipynb` and **Run All**. The notebook resolves the project directory from the current working directory or the optional `RAD_RISK_ROOT` environment variable.
4. The JSON reporting interface writes `Excel_Template_Output/task1_snapshot.json`.
5. To build Excel, install .NET 10 SDK and run `DeskRiskExcel_V3.cs` via the final notebook cell or directly with `dotnet run --file <path-to-cs>`. The supplied C# generator may need its `folder` setting updated when cloning to a different machine.

```bash
# From the repository root after placing the authorized dataset:
pip install numpy pandas matplotlib ipython jupyter
jupyter notebook Task1_Daily_Report/daily_report.ipynb
```

## Validation & Controls

- Reconciles the sum of bond-level DV01s against the aggregated eleven-vertex ladder.
- Bumps each curve vertex by ±1bp and checks the analytical sensitivity against an independent central-difference repricing.
- Confirms exactly 500 historical daily moves are used for each current-window VaR estimate.
- Uses the **fifth-worst** historical P&L from that window to obtain 99% VaR.

**Observed test result:** Normalized maximum analytical vs bumped DV01 gap approximately **2.57 × 10⁻⁷**, within the required `1e-6` tolerance for this simulated test case.

## Limitations & Next Steps

- Educational, simulated portfolio—not a live or production-grade trading system.
- Simplified coupon schedules and discounting; no settlement, accrual, holiday/calendar, or curve-bootstrapping engine.
- DV01-based scenario P&L is linear; convexity, credit spreads, funding, liquidity, and transaction costs are not modeled.
- Reported P&L is not reconciled here; **Task 2** explains changes in risk and **Task 3** adds P&L explanation and exception checks.

## Attribution

Adapted from the educational [RiskHub Fixed Income Risk Dashboard Python Lab](https://riskhub.org/python-labs/fixed-income-risk-dashboard), with notebook documentation, risk validation, visualization, and a separate C# Excel reporting interface. Dataset redistribution should be checked against the lab's terms before publishing data.
