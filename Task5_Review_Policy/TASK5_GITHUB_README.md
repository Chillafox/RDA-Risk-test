# Task 05 — Alert Policy Review & Optimization

**Fixed Income Risk Dashboard | Python · NumPy · pandas · Matplotlib · C#/.NET 10 · ClosedXML · Excel**

## Business problem

A fixed-income risk manager needs a monitoring policy that identifies trading-book and market-data errors while they are active **without creating too many false alerts**. This project evaluates five candidate alert policies under a fixed acceptance rule instead of choosing the policy with the highest raw detection rate.

> **Decision:** `SCALED+FEED+LATE` is the only candidate meeting the simulation's acceptance constraints (93.8% overall detection, at least 87.0% for each incident type, and 1 clean alert across two benchmark years).

## How the experiment works

1. **Reconstruct a clean baseline** from the lab's corrected historical yield curves and trade book. These corrections are retrospective and exist for offline benchmarking only.
2. **Calibrate P&L noise** using clean 2023 observations; also construct stale-feed, curve-kink, late-booking and persistence signals.
3. **Predefine acceptance criteria:** detect at least 90% of simulated incidents overall, at least 80% of each incident type, and flag at most five clean mornings **in each year** (2023, 2024).
4. **Inject 1,200 isolated synthetic incidents:** 200 each of wrong-side, notional, late, missed, stale-feed and bad-print errors. Use random seed `2024`.
5. **Score all five policies** on the same trials; rank eligible candidates by clean-alert count, then median detection delay.
6. **Export an auditable JSON snapshot** and publish the results using a separate C# Excel generator.

## Policy comparison

| Candidate | Overall detection | Clean 2023 | Clean 2024 | Passes? |
|:--|--:|--:|--:|:--:|
| `PNL_FIXED` | 83.9% | 1 | 5 | No |
| `PNL_SCALED` | 80.8% | 1 | 0 | No |
| `SCALED+FEED` | 87.7% | 1 | 0 | No |
| **`SCALED+FEED+LATE`** | **93.8%** | **1** | **0** | **Yes** |
| `ALL+PERSIST` | 94.8% | 6 | 3 | No |

**Why the highest detection rate does not win:** `ALL+PERSIST` exceeds the false-alert ceiling in the 2023 clean period. The qualifying combination improves on P&L-only detection by adding feed-integrity and booking-timing checks.

## Risk-size effect

On the subset of **600 simulated booking errors**, a standalone scaled-P&L detector catches **13.9%** of errors with absolute mismatch DV01 below USD 2,500/bp and **100.0%** above USD 40,000/bp. This finding explains why complementary checks are necessary.

## Repository navigation

| Artifact | Location |
|:--|:--|
| Documented, executable Notebook | [`review_policy.ipynb`](review_policy.ipynb) |
| C# Excel review generator | [`../Excel_Template_Design/DeskRiskPolicy_Task5.cs`](../Excel_Template_Design/DeskRiskPolicy_Task5.cs) |
| Historical Task 4 investigation | [`../Task4_Incident_Forensics/incident_forensics.ipynb`](../Task4_Incident_Forensics/incident_forensics.ipynb) |

The generated `task5_snapshot.json` and Excel output are written under `Excel_Template_Output/` (normally excluded by `.gitignore`). The simulation input file `DataSource/rates_desk_feed.csv` may not be redistributed; supply it locally to reproduce the full run.

## Run locally

1. Install Python with `numpy`, `pandas`, `matplotlib`, and Jupyter.
2. Place the lab feed at `DataSource/rates_desk_feed.csv` in the project root, or set `RAD_RISK_ROOT` to the root directory.
3. Open [`review_policy.ipynb`](review_policy.ipynb) and run all cells. The Python experiment builds the policy scorecard and JSON export.
4. **Optional Excel step:** with .NET 10 installed, the final code cell runs `dotnet run --file ../Excel_Template_Design/DeskRiskPolicy_Task5.cs` through the resolved absolute path. The C# generator uses its own configured input/output path, which may need to be adapted for other machines.

## Validation and limitations

- The Monte Carlo-style injections are **seeded and reproducible**, but do not constitute independent evidence of real-world detection performance.
- Retrospective corrections are used to construct the clean test baseline. They are **not permitted as inputs to real-time alerting**.
- The 2023 reference period is also used for noise calibration, so its false-alarm estimate is in-sample.
- Detection delay is reported among detected cases; missed incidents have no detection delay.
- No simultaneous incidents, spread-risk engine, operational queues or live feeds are modeled.
- The "sign-off" is the outcome of the **exercise policy-selection rule**, not an authorized production model or a regulatory approval.

## Source acknowledgment

Educational work based on the [RiskHub Fixed Income Risk Dashboard Python Lab](https://riskhub.org/python-labs/fixed-income-risk-dashboard), reorganized and extended with narrative documentation, audit checks, JSON exchange and C# Excel review reporting. Simulated data; not live trading risk.
