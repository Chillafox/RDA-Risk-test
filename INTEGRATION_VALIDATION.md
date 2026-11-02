# Final Workstation — Integration Validation Results

**Snapshot as-of:** 2024-12-31  
**Validation type:** Notebook-level numerical and data contract checks (not .NET compilation).  
**Results:** 13 / 13 PASS

| Control | Status | Details |
|:--|:--|:--|
| As-of dates | PASS | 2024-12-31, 2024-12-31, 2024-12-31, 2024-12-31, 2024-12-31 |
| Task 1 scenario count | PASS | scenarios=500, tenors=11 |
| Unified date coverage | PASS | 2024 dates=260, Task 2/3/4/5 aligned |
| History coverage | PASS | Task 2, Task 3, Task 5 each contain 518 reporting dates |
| Closing Net DV01 | PASS | Task 1=47,234.457739 vs Task 2=47,234.457739 USD/bp |
| Closing VaR | PASS | Task 1=939,305.680 vs Task 2=939,305.680 USD |
| Incident as-received VaR | PASS | Task 4 as-received=939,305.680 |
| Closing reported P&L | PASS | Task 1=286,426.72, Task 3=286,426.72 |
| Daily risk walks | PASS | max DV01 residual=4.46e-11; max VaR residual=0 |
| P&L explain | PASS | max explain residual=0; max tenor residual=3.49e-10 |
| Task 3 to Task 4 observed P&L | PASS | max U residual=0 USD |
| Task 5 simulation scope | PASS | 1,200 seeded trials; five candidate policies |
| Precommitted policy sign-off | PASS | winner=SCALED+FEED+LATE; clean flags tie-break preserved |

## Source coverage

- Reporting scope: 260 aligned 2024 report dates.
- Portfolio: 17 live closing positions, 11 key-rate tenors.
- 99% HS VaR: 500 historical scenarios.
- Policy evaluation: 1200 trials, 5 candidate policies.

## Compilation boundary

The C# source was reviewed for compatibility with the previously used ClosedXML 0.105.1 APIs, but the local creation environment does not have .NET 10 SDK and therefore did not compile or run the final C# application. Run the source on Windows using `dotnet run --file .\Excel_Template_Design\DeskRiskWorkstation_Final.cs`. Perform the desktop-Excel UI checks in `RUN_FINAL.md`.
