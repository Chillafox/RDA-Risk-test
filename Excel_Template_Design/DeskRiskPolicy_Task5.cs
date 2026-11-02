#:package ClosedXML@0.105.1
#:property PublishAot=false
#:property PublishTrimmed=false

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;

// ============================================================================
// TASK 5 | ALERT POLICY REVIEW & OPTIMIZATION
// .NET 10 single-file program | ClosedXML 0.105.1 | Macro-free workbook
// Input: Excel_Template_Output/task5_snapshot.json
// Upstream engine: review_policy.ipynb (1,200 seeded simulated incidents)
//
// DESIGN CONTROL:
//   Python performs incident injection, screen evaluation and sign-off.
//   C# independently audits input counts/rates, builds Excel and the UI.
//   The precommitted acceptance criteria are NOT edited by Excel users.
// ============================================================================

string root = Environment.GetEnvironmentVariable("RAD_RISK_ROOT")
              ?? Directory.GetCurrentDirectory();
string outputDir = Path.Combine(root, "Excel_Template_Output");
string input = Path.Combine(outputDir, "task5_snapshot.json");
if (!File.Exists(input))
    throw new FileNotFoundException("Run Task 5's Python JSON export cell first.", input);

var jsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
var d = JsonSerializer.Deserialize<PolicySnapshot>(File.ReadAllText(input), jsonOptions)
        ?? throw new InvalidOperationException("Cannot read Task 5 JSON snapshot.");

string[] expectedPolicies = {
    "PNL_FIXED", "PNL_SCALED", "SCALED+FEED", "SCALED+FEED+LATE", "ALL+PERSIST"
};
string[] expectedKinds = {
    "WRONG_SIDE", "NOTIONAL", "LATE", "MISSED", "STALE", "BAD_PRINT"
};

// Fail early if the simulated evidence or policy outcomes do not reconcile.
if (d.IncidentsPerKind != 200 || d.SimulationSeed != 2024 ||
    d.Policies.Count != 5 || d.KindNames.Count != 6 ||
    d.TrialRecords.Count != 1200 || d.CleanDays.Count != 518 ||
    d.NoiseBp <= 0 || d.FixedThresholdUsd <= 0 ||
    d.AsOf != "2024-12-31")
    throw new InvalidOperationException("Unexpected Task 5 experiment metadata.");
if (!d.PolicyNames.SequenceEqual(expectedPolicies) ||
    !d.KindNames.SequenceEqual(expectedKinds) ||
    !d.Policies.Select(x => x.Name).SequenceEqual(expectedPolicies))
    throw new InvalidOperationException("Five policies or six incident kinds are missing/reordered.");
if (d.CleanDays.Count(x => x.Year == 2023) != 258 ||
    d.CleanDays.Count(x => x.Year == 2024) != 260 ||
    d.CleanDays.Select(x => x.Date).Distinct().Count() != 518)
    throw new InvalidOperationException("Clean-year daily records are incomplete.");
if (d.TrialRecords.Select(x => x.TrialId).Distinct().Count() != 1200 ||
    d.TrialRecords.Any(x => x.Results.Count != expectedPolicies.Length))
    throw new InvalidOperationException("Trial IDs or policy results are incomplete.");
foreach (var kind in expectedKinds)
    if (d.TrialRecords.Count(x => x.Kind == kind) != 200)
        throw new InvalidOperationException($"Incorrect trial count for {kind}.");

var byPolicy = d.Policies.ToDictionary(x => x.Name);
foreach (string name in expectedPolicies)
{
    var p = byPolicy[name];
    if (p.ByKind.Count != 6 || p.ByKind.Any(k => !expectedKinds.Contains(k.Kind)))
        throw new InvalidOperationException($"Incomplete incident coverage for {name}.");
    int caught = 0;
    var delays = new List<int>();
    foreach (string kind in expectedKinds)
    {
        var trials = d.TrialRecords.Where(x => x.Kind == kind)
            .Select(x => x.Results.Single(z => z.Policy == name)).ToArray();
        int hit = trials.Count(x => x.Caught);
        if (trials.Any(x => x.Caught != x.Delay.HasValue ||
                            (x.Delay.HasValue && x.Delay.Value < 0)))
            throw new InvalidOperationException($"Invalid hit/delay pair for {name} / {kind}.");
        var row = p.ByKind.Single(x => x.Kind == kind);
        if (hit != row.Caught || Math.Abs(row.Rate - hit / 200.0) > 1e-9)
            throw new InvalidOperationException($"Hit rate mismatch for {name} / {kind}.");
        caught += hit;
        delays.AddRange(trials.Where(x => x.Caught).Select(x => x.Delay!.Value));
    }
    double median = Median(delays);
    if (Math.Abs(p.Overall - caught / 1200.0) > 1e-9 ||
        Math.Abs(p.MedianDelay - median) > 1e-9)
        throw new InvalidOperationException($"Overall hit rate/delay mismatch for {name}.");
    int clean23 = d.CleanDays.Count(x => x.Year == 2023 && x.PolicyFlags[name]);
    int clean24 = d.CleanDays.Count(x => x.Year == 2024 && x.PolicyFlags[name]);
    bool accepted = p.Overall >= 0.90 && p.ByKind.All(x => x.Rate >= 0.80) &&
                    clean23 <= 5 && clean24 <= 5;
    if (clean23 != p.Clean2023 || clean24 != p.Clean2024 || accepted != p.Passes)
        throw new InvalidOperationException($"Precommitted pass/fail mismatch for {name}.");
}
string chosen = d.Policies.Where(x => x.Passes)
    .OrderBy(x => x.Clean2023 + x.Clean2024)
    .ThenBy(x => x.MedianDelay)
    .Select(x => x.Name).FirstOrDefault() ?? "";
if (chosen != d.Winner)
    throw new InvalidOperationException("Signed-off policy differs from the precommitted rule.");
if (d.Dv01Bands.Sum(x => x.Total) != 600 ||
    d.Dv01Bands.Any(x => x.Caught > x.Total ||
                         Math.Abs(x.Rate - (double)x.Caught / x.Total) > 1e-9))
    throw new InvalidOperationException("Booking-error sensitivity does not reconcile to 600 trials.");

using var wb = new XLWorkbook();
var overview = wb.Worksheets.Add("01 Policy Overview");
var scorecard = wb.Worksheets.Add("02 Policy Scorecard");
var matrix = wb.Worksheets.Add("03 Detection Matrix");
var clean = wb.Worksheets.Add("04 Clean Timeline");
var drill = wb.Worksheets.Add("05 Policy Drilldown");
var sensitivity = wb.Worksheets.Add("06 DV01 Sensitivity");
var trialsSheet = wb.Worksheets.Add("07 Trial Register");
var guide = wb.Worksheets.Add("08 Methodology");
string[] pages = wb.Worksheets.Select(x => x.Name).ToArray();
foreach (var ws in wb.Worksheets)
    PolicyTheme.Setup(ws, d.AsOf, pages);

const int FIRST = 9;
int trialLast = FIRST + d.TrialRecords.Count - 1;
int cleanLast = FIRST + d.CleanDays.Count - 1;

// ============================================================================
// 07 | SIMULATED TRIAL REGISTER (POPULATE BEFORE THE SUMMARY)
// Each row is one of 1,200 re-drawn injected incidents; each policy sees it.
// ============================================================================
PolicyTheme.Banner(trialsSheet, "SIMULATED INCIDENT TRIALS / EVALUATION REGISTER");
PolicyTheme.Section(trialsSheet, "B7:Q7", "200 TRIALS PER KIND | HIT AND FIRST-DETECTION DELAY PER CANDIDATE POLICY");
string[] trialHeaders = {
    "TRIAL ID", "INCIDENT KIND",
    "FIXED HIT", "FIXED DELAY", "SCALED HIT", "SCALED DELAY",
    "FEED HIT", "FEED DELAY", "FEED+LATE HIT", "FEED+LATE DELAY",
    "PERSIST HIT", "PERSIST DELAY"
};
for (int k = 0; k < trialHeaders.Length; k++)
    trialsSheet.Cell(8, 2 + k).Value = trialHeaders[k];
PolicyTheme.Header(trialsSheet, "B8:M8");
trialsSheet.Columns("B:C").Width = 21;
trialsSheet.Columns("D:M").Width = 18;
for (int i = 0; i < d.TrialRecords.Count; i++)
{
    int r = FIRST + i;
    var tr = d.TrialRecords[i];
    trialsSheet.Cell(r, 2).Value = tr.TrialId;
    trialsSheet.Cell(r, 3).Value = tr.Kind;
    for (int j = 0; j < expectedPolicies.Length; j++)
    {
        var x = tr.Results.Single(z => z.Policy == expectedPolicies[j]);
        int col = 4 + j * 2;
        trialsSheet.Cell(r, col).Value = x.Caught ? 1 : 0;
        if (x.Delay.HasValue) trialsSheet.Cell(r, col + 1).Value = x.Delay.Value;
        PolicyTheme.Flag(trialsSheet.Cell(r, col), x.Caught);
    }
    PolicyTheme.Stripe(trialsSheet, r, 2, 13, i);
}
trialsSheet.Range(8, 2, trialLast, 13).SetAutoFilter();
trialsSheet.SheetView.FreezeRows(8);
PolicyTheme.Note(trialsSheet, 1213, "HIT=1 means the policy fired within the injected incident's live window; DELAY is measured in report mornings.");
PolicyTheme.Note(trialsSheet, 1215, "Each trial contains one incident at a time, on the corrected clean-world base. No real incidents are claimed.");

// ============================================================================
// 04 | CLEAN WORLD MONITOR | 258 + 260 clean mornings
// No injected incidents here: counts measure false alert burden per policy.
// ============================================================================
PolicyTheme.Banner(clean, "CLEAN-WORLD DAILY FLAGS / FALSE-ALERT AUDIT");
PolicyTheme.Section(clean, "B7:Q7", "CORRECTED BASELINE | SIX RAW SIGNALS | FIVE CANDIDATE POLICY FLAGS");
string[] cleanHeaders = {
    "REPORT DATE", "YEAR", "UNEXPLAINED (USD)", "Z-SCORE",
    "PNL_FIXED", "PNL_SCALED", "STALE", "KINK", "LATE", "PERSIST",
    expectedPolicies[0], expectedPolicies[1], expectedPolicies[2],
    expectedPolicies[3], expectedPolicies[4], "SELECTED FLAG"
};
for (int k = 0; k < cleanHeaders.Length; k++)
    clean.Cell(8, 2 + k).Value = cleanHeaders[k];
PolicyTheme.Header(clean, "B8:Q8");
clean.Columns("B:C").Width = 17;
clean.Column("D").Width = 25;
clean.Column("E").Width = 14;
clean.Columns("F:K").Width = 15;
clean.Columns("L:P").Width = 21;
clean.Column("Q").Width = 18;
for (int i = 0; i < d.CleanDays.Count; i++)
{
    int r = FIRST + i;
    var x = d.CleanDays[i];
    clean.Cell(r, 2).Value = x.Date;
    clean.Cell(r, 3).Value = x.Year;
    clean.Cell(r, 4).Value = x.Unexplained;
    clean.Cell(r, 5).Value = x.ZScore;
    clean.Cell(r, 4).Style.NumberFormat.Format = PolicyTheme.SignedMoney;
    clean.Cell(r, 5).Style.NumberFormat.Format = "+0.00;-0.00;0.00";
    string[] names = { "pnl_fixed", "pnl_scaled", "stale", "kink", "late", "persist" };
    for (int j = 0; j < names.Length; j++)
    {
        bool value = x.Signals.TryGetValue(names[j], out bool hit) && hit;
        clean.Cell(r, 6 + j).Value = value ? 1 : 0;
        PolicyTheme.Flag(clean.Cell(r, 6 + j), value);
    }
    for (int j = 0; j < expectedPolicies.Length; j++)
    {
        bool hit = x.PolicyFlags[expectedPolicies[j]];
        clean.Cell(r, 12 + j).Value = hit ? 1 : 0;
        PolicyTheme.Flag(clean.Cell(r, 12 + j), hit);
    }
    // Dynamic flag for whichever policy is selected in the interactive sheet.
    clean.Cell(r, 17).FormulaA1 =
        $"IFERROR(INDEX(L{r}:P{r},1,MATCH('05 Policy Drilldown'!$C$9,$L$8:$P$8,0)),0)";
    PolicyTheme.Stripe(clean, r, 2, 17, i);
}
clean.Range(8, 2, cleanLast, 17).SetAutoFilter();
clean.SheetView.FreezeRows(8);
PolicyTheme.Note(clean, cleanLast + 3, "Clean = corrected market feed and corrected true book in the simulation; not a contemporaneous production system.");
PolicyTheme.Note(clean, cleanLast + 5, "A flagged CLEAN day is a false-alarm burden, not an injected incident detection.");

// ============================================================================
// 02 | POLICY SCORECARD | acceptance rule is fixed, not user tunable
// ============================================================================
PolicyTheme.Banner(scorecard, "POLICY EVALUATION SCORECARD / PRECOMMITTED DECISION RULE");
PolicyTheme.Section(scorecard, "B7:Q7", "1,200 SYNTHETIC INCIDENTS | SIX TYPES | FIVE CANDIDATE POLICIES");
string[] scoreHeaders = {
    "POLICY", "OVERALL", "WRONG_SIDE", "NOTIONAL", "LATE", "MISSED",
    "STALE", "BAD_PRINT", "MIN TYPE", "CLEAN 2023", "CLEAN 2024",
    "CLEAN TOTAL", "MEDIAN DELAY", "RULE STATUS", "ENGINE STATUS"
};
for (int k = 0; k < scoreHeaders.Length; k++)
    scorecard.Cell(8, 2 + k).Value = scoreHeaders[k];
PolicyTheme.Header(scorecard, "B8:P8");
scorecard.Column("B").Width = 27;
scorecard.Columns("C:J").Width = 16;
scorecard.Columns("K:P").Width = 18;
for (int i = 0; i < d.Policies.Count; i++)
{
    int r = FIRST + i;
    var p = d.Policies[i];
    scorecard.Cell(r, 2).Value = p.Name;
    scorecard.Cell(r, 3).Value = p.Overall;
    for (int k = 0; k < expectedKinds.Length; k++)
        scorecard.Cell(r, 4 + k).Value = p.ByKind.Single(x => x.Kind == expectedKinds[k]).Rate;
    scorecard.Cell(r, 10).FormulaA1 = $"MIN(D{r}:I{r})";
    scorecard.Cell(r, 11).Value = p.Clean2023;
    scorecard.Cell(r, 12).Value = p.Clean2024;
    scorecard.Cell(r, 13).FormulaA1 = $"K{r}+L{r}";
    scorecard.Cell(r, 14).Value = p.MedianDelay;
    scorecard.Cell(r, 15).FormulaA1 =
        $"IF(AND(C{r}>=0.9,J{r}>=0.8,K{r}<=5,L{r}<=5),\"PASS\",\"FAIL\")";
    scorecard.Cell(r, 16).Value = p.Passes ? "PASS" : "FAIL";
    scorecard.Range(r, 3, r, 10).Style.NumberFormat.Format = "0.0%";
    PolicyTheme.Stripe(scorecard, r, 2, 16, i);
    for (int c = 4; c <= 9; c++)
        PolicyTheme.Coverage(scorecard.Cell(r, c), p.ByKind[c - 4].Rate);
    if (p.Name == chosen)
    {
        scorecard.Range(r, 2, r, 16).Style.Font.Bold = true;
        scorecard.Cell(r, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#DCF3EF");
        scorecard.Cell(r, 2).Style.Font.FontColor = XLColor.FromHtml("#086F69");
    }
}
PolicyTheme.Section(scorecard, "B17:P17", "LOCKED SIGN-OFF TEST");
PolicyTheme.Note(scorecard, 19, "PASS only if: overall >= 90%; EACH of six types >= 80%; clean flags <= 5 in 2023 AND <= 5 in 2024.");
PolicyTheme.Note(scorecard, 21, "If multiple policies pass, select the fewest combined clean flags; tie-break by lower median detection delay.");
PolicyTheme.Note(scorecard, 23, "ENGINE STATUS is audited against the Python simulation; RULE STATUS is calculated with visible Excel formulas.");
PolicyTheme.Note(scorecard, 25, "The policy names and sensitivity thresholds are precommitted; editing displayed coverage does not rerun incidents.");

// ============================================================================
// 03 | DETECTION COVERAGE MATRIX | by incident type
// ============================================================================
PolicyTheme.Banner(matrix, "INCIDENT-TYPE DETECTION / POLICY COVERAGE HEATMAP");
PolicyTheme.Section(matrix, "B7:Q7", "DETECTION RATE WHILE INCIDENT IS LIVE | 200 INCIDENTS PER TYPE");
matrix.Column("B").Width = 25;
matrix.Column("C").Width = 15;
matrix.Columns("D:H").Width = 24;
matrix.Column("I").Width = 22;
matrix.Cell("B9").Value = "INCIDENT TYPE";
matrix.Cell("C9").Value = "TRIALS";
for (int j = 0; j < expectedPolicies.Length; j++)
    matrix.Cell(9, 4 + j).Value = expectedPolicies[j];
matrix.Cell("I9").Value = "SIGNED-OFF COVERAGE";
PolicyTheme.Header(matrix, "B9:I9");
for (int i = 0; i < expectedKinds.Length; i++)
{
    int r = 10 + i;
    string kind = expectedKinds[i];
    matrix.Cell(r, 2).Value = kind;
    matrix.Cell(r, 3).Value = d.IncidentsPerKind;
    for (int j = 0; j < expectedPolicies.Length; j++)
    {
        double pct = byPolicy[expectedPolicies[j]].ByKind.Single(x => x.Kind == kind).Rate;
        matrix.Cell(r, 4 + j).Value = pct;
        matrix.Cell(r, 4 + j).Style.NumberFormat.Format = "0.0%";
        PolicyTheme.Coverage(matrix.Cell(r, 4 + j), pct);
    }
    matrix.Cell(r, 9).Value = chosen == "" ? "NO PASS" :
        byPolicy[chosen].ByKind.Single(x => x.Kind == kind).Rate;
    if (chosen != "") matrix.Cell(r, 9).Style.NumberFormat.Format = "0.0%";
    PolicyTheme.Stripe(matrix, r, 2, 9, i);
}
PolicyTheme.Section(matrix, "B19:O19", "READING THE MATRIX");
PolicyTheme.Note(matrix, 21, "Green coverage >= 80%. Red coverage < 80%: fails the minimum incident-type requirement.");
PolicyTheme.Note(matrix, 23, "A high overall hit rate cannot compensate for poor coverage in a specific operational incident type.");
PolicyTheme.Note(matrix, 25, "LATE signal catches delayed ticket booking only when it becomes observable; it cannot reveal never-booked trades.");

// ============================================================================
// 06 | BOOKING ERROR SENSITIVITY (PNL_SCALED ONLY)
// ============================================================================
PolicyTheme.Banner(sensitivity, "BOOKING-ERROR SENSITIVITY / IMPLIED DV01 SIZE");
PolicyTheme.Section(sensitivity, "B7:Q7", "600 SYNTHETIC WRONG_SIDE / NOTIONAL / MISSED INCIDENTS | PNL_SCALED");
sensitivity.Columns("B:F").Width = 25;
string[] sensitivityHeaders = { "ABS ERROR DV01 BAND", "INCIDENTS", "CAUGHT", "DETECTION RATE", "MISSED" };
for (int j = 0; j < sensitivityHeaders.Length; j++)
    sensitivity.Cell(9, 2 + j).Value = sensitivityHeaders[j];
PolicyTheme.Header(sensitivity, "B9:F9");
int sensitivityRow = 10;
foreach (var b in d.Dv01Bands)
{
    sensitivity.Cell(sensitivityRow, 2).Value = b.Band;
    sensitivity.Cell(sensitivityRow, 3).Value = b.Total;
    sensitivity.Cell(sensitivityRow, 4).Value = b.Caught;
    sensitivity.Cell(sensitivityRow, 5).FormulaA1 = $"IFERROR(D{sensitivityRow}/C{sensitivityRow},0)";
    sensitivity.Cell(sensitivityRow, 6).FormulaA1 = $"C{sensitivityRow}-D{sensitivityRow}";
    sensitivity.Cell(sensitivityRow, 5).Style.NumberFormat.Format = "0.0%";
    PolicyTheme.Stripe(sensitivity, sensitivityRow, 2, 6, sensitivityRow - 10);
    sensitivityRow++;
}
sensitivity.Cell("B16").Value = "TOTAL";
sensitivity.Cell("C16").FormulaA1 = $"SUM(C10:C{sensitivityRow - 1})";
sensitivity.Cell("D16").FormulaA1 = $"SUM(D10:D{sensitivityRow - 1})";
sensitivity.Cell("E16").FormulaA1 = "IFERROR(D16/C16,0)";
sensitivity.Cell("F16").FormulaA1 = "C16-D16";
sensitivity.Cell("E16").Style.NumberFormat.Format = "0.0%";
PolicyTheme.Total(sensitivity, 16, 2, 6);
PolicyTheme.Section(sensitivity, "B20:P20", "MODEL SENSITIVITY / INTERPRETATION");
PolicyTheme.Note(sensitivity, 22, "This analysis covers only 600 injected BOOKING errors and the scaled P&L screen, NOT all 1,200 trials.");
PolicyTheme.Note(sensitivity, 24, "A small missing DV01 produces a small expected P&L gap and may be difficult to distinguish from marking noise.");
PolicyTheme.Note(sensitivity, 26, "The aggregate policy improves coverage with feed and late-booking controls; it is not proof of universal detection.");

// ============================================================================
// 05 | INTERACTIVE POLICY + CLEAN-DAY DRILLDOWN (NO VBA)
// ============================================================================
PolicyTheme.Banner(drill, "POLICY DRILLDOWN / INTERACTIVE REVIEW");
PolicyTheme.Section(drill, "B7:Q7", "BLUE INPUTS: SELECT POLICY AND CLEAN DATE | WORKBOOK FORMULAS UPDATE THE VIEW");
drill.Column("B").Width = 30;
drill.Column("C").Width = 30;
drill.Columns("F:I").Width = 22;
drill.Columns("J:Q").Width = 15;
drill.Cell("B9").Value = "SELECT POLICY";
drill.Cell("C9").Value = chosen != "" ? chosen : expectedPolicies[0];
PolicyTheme.Input(drill.Cell("C9"));
for (int i = 0; i < expectedPolicies.Length; i++) drill.Cell(9 + i, 26).Value = expectedPolicies[i];
drill.Cell("C9").CreateDataValidation().List(drill.Range("Z9:Z13"));

drill.Cell("F11").Value = "OVERALL DETECTION";
drill.Cell("F12").Value = "MIN TYPE COVERAGE";
drill.Cell("F13").Value = "CLEAN FLAGS / 2023";
drill.Cell("F14").Value = "CLEAN FLAGS / 2024";
drill.Cell("F15").Value = "MEDIAN DELAY";
drill.Cell("F16").Value = "RULE STATUS";
string PolicyLookup(string col) =>
    $"IFERROR(INDEX('02 Policy Scorecard'!${col}$9:${col}$13,MATCH($C$9,'02 Policy Scorecard'!$B$9:$B$13,0)),\"\")";
drill.Cell("G11").FormulaA1 = PolicyLookup("C");
drill.Cell("G12").FormulaA1 = PolicyLookup("J");
drill.Cell("G13").FormulaA1 = PolicyLookup("K");
drill.Cell("G14").FormulaA1 = PolicyLookup("L");
drill.Cell("G15").FormulaA1 = PolicyLookup("N");
drill.Cell("G16").FormulaA1 = PolicyLookup("O");
drill.Cell("G11").Style.NumberFormat.Format = "0.0%";
drill.Cell("G12").Style.NumberFormat.Format = "0.0%";

PolicyTheme.Section(drill, "B18:Q18", "CLEAN-DAY INSPECTION | SCREEN INPUTS ARE FROM THE CORRECTED BASELINE");
drill.Cell("B20").Value = "SELECT CLEAN DATE";
drill.Cell("C20").Value = d.CleanDays[^1].Date;
PolicyTheme.Input(drill.Cell("C20"));
for (int i = 0; i < d.CleanDays.Count; i++) drill.Cell(9 + i, 27).Value = d.CleanDays[i].Date;
drill.Cell("C20").CreateDataValidation().List(drill.Range($"AA9:AA{cleanLast}"));
string DayLookup(string col) =>
    $"IFERROR(INDEX('04 Clean Timeline'!${col}$9:${col}${cleanLast},MATCH($C$20,'04 Clean Timeline'!$B$9:$B${cleanLast},0)),\"\")";
string[] dayLabels = {
    "YEAR", "UNEXPLAINED P&L", "Z-SCORE", "PNL FIXED", "PNL SCALED",
    "STALE", "KINK", "LATE", "PERSIST", "SELECTED POLICY FLAG"
};
string[] dailyCols = { "C", "D", "E", "F", "G", "H", "I", "J", "K", "Q" };
for (int i = 0; i < dayLabels.Length; i++)
{
    int r = 23 + i;
    drill.Cell(r, 2).Value = dayLabels[i];
    drill.Cell(r, 7).FormulaA1 = DayLookup(dailyCols[i]);
}
drill.Cell("G24").Style.NumberFormat.Format = PolicyTheme.SignedMoney;
drill.Cell("G25").Style.NumberFormat.Format = "+0.00;-0.00;0.00";
drill.Columns("Z:AA").Hide();
PolicyTheme.Section(drill, "B36:Q36", "IMPORTANT: REVIEW A POLICY, NOT AN IN-SHEET RE-OPTIMIZATION");
PolicyTheme.Note(drill, 38, "The dropdown changes which FIXED historical result is displayed. It does not retrain, rerun incidents or change the approved rule.");
PolicyTheme.Note(drill, 40, "Clean-day flags are retrospective false-alarm examples; trial detection rates are evaluated separately on injected incidents.");
PolicyTheme.Note(drill, 42, "Approval is a simulated research sign-off, not an operational compliance approval or production rollout.");

// ============================================================================
// 01 | EXECUTIVE POLICY REVIEW
// ============================================================================
PolicyTheme.Banner(overview, "RATES DESK / ALERT POLICY REVIEW BOARD");
PolicyTheme.Section(overview, "B7:Q7", "RECOMMENDED POLICY | REPRODUCIBLE SIMULATION | PRECOMMITTED ACCEPTANCE TEST");
PolicyTheme.Kpi(overview, "B9:F9", "B10:F12", "SELECTED POLICY", chosen != "" ? chosen : "NO POLICY PASSED");
PolicyTheme.Kpi(overview, "H9:K9", "H10:K12", "OVERALL DETECTION", chosen == "" ? "N/A" : byPolicy[chosen].Overall.ToString("P1", CultureInfo.InvariantCulture));
PolicyTheme.Kpi(overview, "M9:Q9", "M10:Q12", "CLEAN FLAGS / 2 YEARS", chosen == "" ? "N/A" :
    (byPolicy[chosen].Clean2023 + byPolicy[chosen].Clean2024).ToString(CultureInfo.InvariantCulture));
PolicyTheme.Section(overview, "B15:Q15", "CANDIDATE POLICY COMPARISON");
string[] executiveHeaders = { "POLICY", "HIT RATE", "WORST KIND", "CLEAN 2023", "CLEAN 2024", "DELAY", "DECISION" };
int[] executiveCols = { 2, 6, 8, 10, 12, 14, 16 };
for (int i = 0; i < executiveHeaders.Length; i++)
    overview.Cell(17, executiveCols[i]).Value = executiveHeaders[i];
PolicyTheme.Header(overview, "B17:Q17");
for (int i = 0; i < d.Policies.Count; i++)
{
    int r = 18 + i;
    var p = d.Policies[i];
    overview.Cell(r, 2).Value = p.Name;
    overview.Cell(r, 6).Value = p.Overall;
    overview.Cell(r, 8).Value = p.ByKind.Min(x => x.Rate);
    overview.Cell(r, 10).Value = p.Clean2023;
    overview.Cell(r, 12).Value = p.Clean2024;
    overview.Cell(r, 14).Value = p.MedianDelay;
    overview.Cell(r, 16).Value = p.Passes ? (p.Name == chosen ? "SELECTED" : "ELIGIBLE") : "REJECTED";
    overview.Cell(r, 6).Style.NumberFormat.Format = "0.0%";
    overview.Cell(r, 8).Style.NumberFormat.Format = "0.0%";
    PolicyTheme.Stripe(overview, r, 2, 17, i);
    if (p.Name == chosen)
    {
        overview.Range(r, 2, r, 17).Style.Font.Bold = true;
        overview.Cell(r, 16).Style.Font.FontColor = XLColor.FromHtml("#06736E");
    }
}
PolicyTheme.Section(overview, "B26:Q26", "WHY THIS POLICY / WHY NOT THE HIGHEST RAW HIT RATE?");
PolicyTheme.Note(overview, 28, "Approval requires: >=90% overall detection, >=80% on each of six incident types, and <=5 clean flags in EACH year.");
PolicyTheme.Note(overview, 30, "ALL+PERSIST has a higher raw hit rate but fails the 2023 clean false-alert ceiling; broader detection is not automatically better.");
PolicyTheme.Note(overview, 32, "The selected policy combines scaled P&L, stale feed, kink and late-booking signals without a persistence add-on.");
PolicyTheme.Section(overview, "B36:Q36", "RUN METADATA & CONTROL BOUNDARIES");
overview.Cell("B38").Value = "Simulation seed";
overview.Cell("H38").Value = d.SimulationSeed;
overview.Cell("B39").Value = "Synthetic incidents";
overview.Cell("H39").Value = d.TrialRecords.Count;
overview.Cell("B40").Value = "Incidents per type";
overview.Cell("H40").Value = d.IncidentsPerKind;
overview.Cell("B41").Value = "Clean-year observations";
overview.Cell("H41").Value = d.CleanDays.Count;
overview.Cell("B42").Value = "2023 fitted noise (bp)";
overview.Cell("H42").Value = d.NoiseBp;
overview.Cell("H42").Style.NumberFormat.Format = "0.0000";
overview.Cell("B43").Value = "3.5-sigma fixed P&L limit";
overview.Cell("H43").Value = d.FixedThresholdUsd;
overview.Cell("H43").Style.NumberFormat.Format = PolicyTheme.Money;
PolicyTheme.Note(overview, 46, "EDUCATIONAL SIMULATION / NOT A LIVE DESK / NO LIVE MARKET DATA / NOT A PRODUCTION CONTROL APPROVAL.");

// ============================================================================
// 08 | METHODOLOGY AND INTERPRETATION
// ============================================================================
PolicyTheme.Banner(guide, "METHODOLOGY / EXPERIMENT DESIGN & DECISION GOVERNANCE");
guide.Column("B").Width = 30;
guide.Column("C").Width = 34;
guide.Column("D").Width = 105;
PolicyTheme.Section(guide, "B7:Q7", "EXPERIMENT PROCESS / BUSINESS PURPOSE");
guide.Cell("B9").Value = "STEP";
guide.Cell("C9").Value = "SOURCE";
guide.Cell("D9").Value = "RULE / PURPOSE";
PolicyTheme.Header(guide, "B9:D9");
var steps = new (string Label, string Source, string Note)[] {
    ("01 | Clean World", "Corrected feed + true book", "Construct a corrected baseline free of the known Task 4 booking and market-feed incidents; use it to measure clean-year false-alarm burden."),
    ("02 | Calibrate", "2023 clean unexplained P&L", "Fit noise scale s, and lock the fixed 3.5-sigma threshold before policy comparison."),
    ("03 | Generate", "Seeded random simulation", "Use seed 2024; inject 200 distinct episodes of each of six incident types, one incident at a time on the 2024 baseline."),
    ("04 | Detect", "Five candidate policies", "Compute scaled/fixed P&L, feed, late and persistence flags from observable simulated morning data. The retrospective incident kind labels are evaluation-only."),
    ("05 | Score", "Incident live window", "A trial is caught if a selected policy fires on or after incident first and no later than incident last; delay counts mornings from first eligible report to first hit."),
    ("06 | False Alarms", "2023 + 2024 clean years", "Count policy flags in clean data without injected incidents. The <=5 limit applies separately to EACH year."),
    ("07 | Select", "Precommitted acceptance rule", "Require >=90% overall detection and >=80% per incident kind; choose minimum combined clean flags among passes, then shortest median delay."),
    ("08 | Publish", "Python JSON -> C# / Excel", "Publish an auditable scorecard, per-kind coverage, clean-day flags and the 1,200-trial register; no VBA macros or real-time service calls.")
};
for (int i = 0; i < steps.Length; i++)
{
    int r = 10 + i;
    guide.Cell(r, 2).Value = steps[i].Label;
    guide.Cell(r, 3).Value = steps[i].Source;
    guide.Cell(r, 4).Value = steps[i].Note;
    guide.Range(r, 2, r, 4).Style.Alignment.WrapText = true;
    guide.Row(r).Height = 40;
    PolicyTheme.Stripe(guide, r, 2, 4, i);
}
PolicyTheme.Section(guide, "B21:Q21", "DETECTION SIGNALS / EXACT POLICY MEMBERSHIP");
guide.Cell("B23").Value = "POLICY";
guide.Cell("C23").Value = "SIGNALS";
PolicyTheme.Header(guide, "B23:D23");
for (int i = 0; i < d.Policies.Count; i++)
{
    int r = 24 + i;
    guide.Cell(r, 2).Value = d.Policies[i].Name;
    guide.Cell(r, 3).Value = string.Join(" OR ", d.Policies[i].Signals);
    guide.Range(r, 3, r, 4).Merge();
    guide.Row(r).Height = 25;
    PolicyTheme.Stripe(guide, r, 2, 4, i);
}
PolicyTheme.Section(guide, "B32:Q32", "CRITICAL LIMITATIONS");
PolicyTheme.Note(guide, 34, "Simulated hit rates are conditional on the injected incident sizes, timing, type mix and seed; they do not estimate real-world prevalence or predictive accuracy.");
PolicyTheme.Note(guide, 36, "Policy OR-combination may catch an incident by a coincidental alert; a hit is not equivalent to successful root-cause identification.");
PolicyTheme.Note(guide, 38, "The threshold and policy membership are LOCKED for this evaluation. What-if modifications require rerunning the Python simulation.");
PolicyTheme.Note(guide, 40, "Shorter detection delay and lower clean false-alarm burden matter in addition to headline coverage. No financial materiality weighting is applied.");
PolicyTheme.Note(guide, 42, "No live feeds, trade execution, production compliance sign-off, persistence monitoring process or delegated user approvals are implemented.");
PolicyTheme.Note(guide, 45, "Model lineage: RiskHub fixed income training scenario; Task 1 -> Task 2 -> Task 3 -> Task 4 -> Task 5 simulation.");

// Save the workbook; Excel recalculates formula-driven drilldown on opening.
Directory.CreateDirectory(outputDir);
string destination = Path.Combine(outputDir, $"RatesDesk_PolicyReview_Task5_{d.AsOf}.xlsx");
wb.SaveAs(destination);
Console.WriteLine("SUCCESS: Task 5 Policy Review Excel generated");
Console.WriteLine("Excel: " + destination);
Console.WriteLine("Trials: " + d.TrialRecords.Count);
Console.WriteLine("Clean 2023 / 2024: 258 / 260");
Console.WriteLine("Signed-off policy: " + (chosen == "" ? "NONE" : chosen));
if (chosen != "")
    Console.WriteLine($"Detection: {byPolicy[chosen].Overall:P1} | Clean flags: {byPolicy[chosen].Clean2023 + byPolicy[chosen].Clean2024}");

static double Median(List<int> values)
{
    if (values.Count == 0) return double.NaN;
    values.Sort();
    int n = values.Count;
    return n % 2 == 1 ? values[n / 2] : (values[n / 2 - 1] + values[n / 2]) / 2.0;
}

public sealed class PolicySnapshot
{
    public string AsOf { get; set; } = "";
    public string Experiment { get; set; } = "";
    public string PrecommittedRule { get; set; } = "";
    public int SimulationSeed { get; set; }
    public int IncidentsPerKind { get; set; }
    public double NoiseBp { get; set; }
    public double FixedThresholdUsd { get; set; }
    public string Winner { get; set; } = "";
    public List<string> KindNames { get; set; } = new();
    public List<string> PolicyNames { get; set; } = new();
    public List<PolicyResult> Policies { get; set; } = new();
    public List<CleanDay> CleanDays { get; set; } = new();
    public List<IncidentTrial> TrialRecords { get; set; } = new();
    public List<Dv01Band> Dv01Bands { get; set; } = new();
}
public sealed class PolicyResult
{
    public string Name { get; set; } = "";
    public List<string> Signals { get; set; } = new();
    public double Overall { get; set; }
    public List<KindResult> ByKind { get; set; } = new();
    public double MedianDelay { get; set; }
    public int Clean2023 { get; set; }
    public int Clean2024 { get; set; }
    public bool Passes { get; set; }
}
public sealed class KindResult
{
    public string Kind { get; set; } = "";
    public double Rate { get; set; }
    public int Caught { get; set; }
}
public sealed class CleanDay
{
    public string Date { get; set; } = "";
    public int Year { get; set; }
    public double Unexplained { get; set; }
    public double ZScore { get; set; }
    public Dictionary<string, bool> Signals { get; set; } = new();
    public Dictionary<string, bool> PolicyFlags { get; set; } = new();
}
public sealed class IncidentTrial
{
    public string TrialId { get; set; } = "";
    public string Kind { get; set; } = "";
    public List<TrialPolicyResult> Results { get; set; } = new();
}
public sealed class TrialPolicyResult
{
    public string Policy { get; set; } = "";
    public bool Caught { get; set; }
    public int? Delay { get; set; }
}
public sealed class Dv01Band
{
    public string Band { get; set; } = "";
    public int Total { get; set; }
    public int Caught { get; set; }
    public double Rate { get; set; }
}

public static class PolicyTheme
{
    static readonly XLColor Navy = XLColor.FromHtml("#14263D");
    static readonly XLColor Blue = XLColor.FromHtml("#244767");
    static readonly XLColor Pale = XLColor.FromHtml("#EDF3F8");
    static readonly XLColor Muted = XLColor.FromHtml("#607389");
    static readonly XLColor Teal = XLColor.FromHtml("#078C86");
    static readonly XLColor Red = XLColor.FromHtml("#C54D5F");
    public const string Money = "\"USD \"#,##0;\"USD \"(#,##0);\"USD 0\"";
    public const string SignedMoney = "+\"USD \"#,##0;-\"USD \"#,##0;\"USD 0\"";

    public static void Setup(IXLWorksheet ws, string asOf, string[] pages)
    {
        ws.ShowGridLines = false;
        ws.Style.Font.FontName = "Aptos";
        ws.Style.Font.FontSize = 10;
        ws.Style.Font.FontColor = Navy;
        ws.Column("A").Width = 3;
        ws.Columns("B:Q").Width = 13;
        ws.Row(2).Height = 25;
        ws.Row(3).Height = 25;
        ws.Row(5).Height = 26;
        ws.Cell("B4").Value = $"AS OF {asOf} | TASK 05 | SIMULATED POLICY REVIEW / NOT LIVE";
        ws.Cell("B4").Style.Font.FontColor = Muted;
        ws.Cell("B4").Style.Font.FontSize = 9;
        for (int i = 0; i < pages.Length; i++)
        {
            int c = 2 + 2 * i;
            var area = ws.Range(5, c, 5, c + 1);
            area.Merge();
            string name = pages[i];
            area.FirstCell().FormulaA1 =
                $"HYPERLINK(\"#'{name}'!B2\",\"{(i + 1):00} {ShortNav(name)}\")";
            area.Style.Fill.BackgroundColor = Pale;
            area.Style.Font.FontColor = Blue;
            area.Style.Font.Bold = true;
            area.Style.Font.FontSize = 8;
            area.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            area.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }
    }
    static string ShortNav(string name) => name.Substring(3) switch
    {
        "Policy Overview" => "OVERVIEW",
        "Policy Scorecard" => "SCORECARD",
        "Detection Matrix" => "COVERAGE",
        "Clean Timeline" => "CLEAN FLAGS",
        "Policy Drilldown" => "DRILLDOWN",
        "DV01 Sensitivity" => "DV01 BANDS",
        "Trial Register" => "TRIALS",
        _ => "METHOD"
    };
    public static void Banner(IXLWorksheet ws, string title)
    {
        var area = ws.Range("B2:Q3");
        area.Merge();
        area.FirstCell().Value = title;
        area.Style.Fill.BackgroundColor = Navy;
        area.Style.Font.FontColor = XLColor.White;
        area.Style.Font.FontSize = 18;
        area.Style.Font.Bold = true;
        area.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Section(IXLWorksheet ws, string location, string title)
    {
        var area = ws.Range(location);
        area.Merge();
        area.FirstCell().Value = title;
        area.Style.Fill.BackgroundColor = Blue;
        area.Style.Font.FontColor = XLColor.White;
        area.Style.Font.Bold = true;
        area.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        ws.Row(area.RangeAddress.FirstAddress.RowNumber).Height = 25;
    }
    public static void Kpi(IXLWorksheet ws, string header, string body, string label, string value)
    {
        var top = ws.Range(header);
        top.Merge();
        top.FirstCell().Value = label;
        top.Style.Fill.BackgroundColor = Pale;
        top.Style.Font.FontColor = Muted;
        top.Style.Font.Bold = true;
        var val = ws.Range(body);
        val.Merge();
        val.FirstCell().Value = value;
        val.Style.Fill.BackgroundColor = Pale;
        val.Style.Font.FontColor = Navy;
        val.Style.Font.FontSize = value.Length > 18 ? 13 : 21;
        val.Style.Font.Bold = true;
        val.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        val.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }
    public static void Header(IXLWorksheet ws, string address)
    {
        var area = ws.Range(address);
        area.Style.Fill.BackgroundColor = Blue;
        area.Style.Font.FontColor = XLColor.White;
        area.Style.Font.Bold = true;
        area.Style.Alignment.WrapText = true;
        ws.Row(area.RangeAddress.FirstAddress.RowNumber).Height = 32;
    }
    public static void Stripe(IXLWorksheet ws, int row, int col1, int col2, int n)
    {
        if (n % 2 == 0)
            ws.Range(row, col1, row, col2).Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F3F7FA");
    }
    public static void Flag(IXLCell cell, bool active)
    {
        if (active)
        {
            cell.Style.Font.FontColor = Red;
            cell.Style.Font.Bold = true;
        }
        else cell.Style.Font.FontColor = Muted;
    }
    public static void Coverage(IXLCell cell, double value)
    {
        cell.Style.Font.FontColor = value >= 0.80 ? Teal : Red;
        cell.Style.Font.Bold = true;
    }
    public static void Input(IXLCell cell)
    {
        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#E3EEFF");
        cell.Style.Font.FontColor = XLColor.FromHtml("#1255B0");
        cell.Style.Font.Bold = true;
        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        cell.Style.Border.OutsideBorderColor = XLColor.FromHtml("#6B96D5");
    }
    public static void Total(IXLWorksheet ws, int row, int col1, int col2)
    {
        var area = ws.Range(row, col1, row, col2);
        area.Style.Fill.BackgroundColor = Pale;
        area.Style.Font.Bold = true;
        ws.Row(row).Height = 28;
    }
    public static void Note(IXLWorksheet ws, int row, string message)
    {
        var cell = ws.Cell(row, 2);
        cell.Value = message;
        cell.Style.Font.FontColor = Muted;
        cell.Style.Font.FontSize = 10;
        ws.Row(row).Height = 24;
    }
}
