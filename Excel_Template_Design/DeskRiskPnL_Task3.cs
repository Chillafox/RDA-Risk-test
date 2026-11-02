#:package ClosedXML@0.105.1
#:property PublishAot=false
#:property PublishTrimmed=false

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

// ============================================================================
// TASK 3 | FIXED INCOME P&L EXPLAIN + NOISE-AWARE EXCEPTION MONITOR
// .NET 10 single-file app / ClosedXML / No VBA / simulated desk data
// Input: Excel_Template_Output/task3_snapshot.json (from pnl_explain.ipynb)
// Run from project root: dotnet run --file Excel_Template_Design/DeskRiskPnL_Task3.cs
// ============================================================================

string root = Environment.GetEnvironmentVariable("RAD_RISK_ROOT")
    ?? Directory.GetCurrentDirectory();
string outputDir = Path.Combine(root, "Excel_Template_Output");
string source = Path.Combine(outputDir, "task3_snapshot.json");
if (!File.Exists(source))
    throw new FileNotFoundException("Run the Task 3 Jupyter JSON exporter first.", source);

var data = JsonSerializer.Deserialize<PnlSnapshot>(File.ReadAllText(source))
    ?? throw new InvalidOperationException("Invalid Task 3 JSON snapshot.");
var days = data.Daily.OrderBy(x => x.Date, StringComparer.Ordinal).ToList();
string[] tenors = {"3M", "6M", "1Y", "2Y", "3Y", "5Y", "7Y", "10Y", "15Y", "20Y", "30Y"};

if (days.Count == 0 || data.Tenors.Count != tenors.Length ||
    !data.Tenors.SequenceEqual(tenors) || data.FitNoiseBp <= 0 ||
    data.TrainingStdUnexplained <= 0)
    throw new InvalidOperationException("Invalid record count, tenor list, or 2023 calibration metadata.");

int trainN = days.Count(x => x.Date.StartsWith("2023-", StringComparison.Ordinal));
int assessN = days.Count(x => x.Date.StartsWith("2024-", StringComparison.Ordinal));
if (trainN != data.TrainingDays || assessN != data.AssessmentDays)
    throw new InvalidOperationException("2023/2024 date counts do not match calibration metadata.");
if (days.Any(x => x.Date.Length != 10 || x.TenorPnl.Count != tenors.Length ||
    x.TenorDv01.Count != tenors.Length || x.CurveMoveBp.Count != tenors.Length))
    throw new InvalidOperationException("Every daily record must have ISO date and 11 tenor inputs.");
if (days.Select(x => x.Date).Distinct().Count() != days.Count)
    throw new InvalidOperationException("Duplicate report dates in Task 3 JSON.");

foreach (var d in days)
{
    if (!double.IsFinite(d.ReportedPnl) || !double.IsFinite(d.Carry) ||
        !double.IsFinite(d.MarketPnl) || !double.IsFinite(d.PredictedPnl) ||
        !double.IsFinite(d.FullRevaluation) || !double.IsFinite(d.Unexplained) ||
        !double.IsFinite(d.NoiseSd) || d.NoiseSd <= 0 || !double.IsFinite(d.ZScore))
        throw new InvalidOperationException($"Invalid metric on {d.Date}.");
    if (Math.Abs(d.Carry + d.MarketPnl - d.PredictedPnl) > 0.0001 ||
        Math.Abs(d.ReportedPnl - d.PredictedPnl - d.Unexplained) > 0.0001 ||
        Math.Abs(d.TenorPnl.Sum() - d.MarketPnl) > 0.0001 ||
        Math.Abs(d.Unexplained / d.NoiseSd - d.ZScore) > 0.0001)
        throw new InvalidOperationException($"Python Task 3 metrics fail reconciliation on {d.Date}.");
    for (int i = 0; i < tenors.Length; i++)
        if (Math.Abs(-d.TenorDv01[i] * d.CurveMoveBp[i] - d.TenorPnl[i]) > 0.0001)
            throw new InvalidOperationException($"Curve contribution mismatch on {d.Date} / {tenors[i]}.");
}

int fixedAtDefault = days.Count(x => x.Date.StartsWith("2024-") &&
    Math.Abs(x.Unexplained) > data.DefaultSigma * data.TrainingStdUnexplained);
int scaledAtDefault = days.Count(x => x.Date.StartsWith("2024-") &&
    Math.Abs(x.ZScore) > data.DefaultSigma);
if (fixedAtDefault != data.DefaultFixedCount || scaledAtDefault != data.DefaultScaledCount)
    throw new InvalidOperationException("Baseline 2024 flag counts differ from exported calibration results.");

using var wb = new XLWorkbook();
var overview = wb.Worksheets.Add("01 PnL Overview");
var daily = wb.Worksheets.Add("02 Daily Explain");
var drill = wb.Worksheets.Add("03 Morning Drilldown");
var tenor = wb.Worksheets.Add("04 Tenor Attribution");
var exceptions = wb.Worksheets.Add("05 Exception Monitor");
var method = wb.Worksheets.Add("06 Methodology");
var tenorData = wb.Worksheets.Add("07 Tenor Data");
string[] pages = {overview.Name, daily.Name, drill.Name, tenor.Name, exceptions.Name, method.Name};
foreach (var sheet in wb.Worksheets)
    PnlTheme.Setup(sheet, data.AsOf, pages);

const int FIRST = 9;
int LAST = FIRST + days.Count - 1;
int first2024 = FIRST + trainN;
int last2024 = LAST;

// ============================================================================
// 02 | ALL DAILY P&L EXPLAIN RECORDS (SOURCE OF INTERACTIVE WORKBOOK FORMULAS)
// Displayed amounts are in USD. 2023 = calibration, 2024 = assessment.
// ============================================================================
PnlTheme.Banner(daily, "DAILY P&L EXPLAIN / AUDITABLE RECONCILIATION");
PnlTheme.Section(daily, "B7:T7", "P&L MODEL, RECONCILIATION AND EXCEPTION FLAGS | 2023 BASELINE / 2024 ASSESSMENT");
string[] dayHeaders = {
    "DATE", "REPORTED", "CARRY", "MARKET / LINEAR", "PREDICTED", "FULL REVALUE",
    "UNEXPLAINED", "DATA GAP", "MODEL GAP", "WALK RESIDUAL", "EXPECTED NOISE SD",
    "Z-SCORE", "ABS Z", "FIXED FLAG", "SCALED FLAG", "COMPARISON",
    "RECON FLAG", "MONTH", "YEAR"
};
for (int i = 0; i < dayHeaders.Length; i++) daily.Cell(8, i + 2).Value = dayHeaders[i];
PnlTheme.Header(daily, "B8:T8");
daily.Column("B").Width = 16;
daily.Columns("C:L").Width = 19;
daily.Column("M").Width = 13;
daily.Column("N").Width = 13;
daily.Columns("O:R").Width = 19;
daily.Columns("S:T").Width = 11;

for (int i = 0; i < days.Count; i++)
{
    int row = FIRST + i;
    var d = days[i];
    daily.Cell(row, 2).Value = DateTime.ParseExact(d.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    daily.Cell(row, 2).Style.DateFormat.Format = "yyyy-mm-dd";
    daily.Cell(row, 3).Value = d.ReportedPnl;
    daily.Cell(row, 4).Value = d.Carry;
    daily.Cell(row, 5).Value = d.MarketPnl;
    daily.Cell(row, 6).FormulaA1 = $"D{row}+E{row}";
    daily.Cell(row, 7).Value = d.FullRevaluation;
    daily.Cell(row, 8).FormulaA1 = $"C{row}-F{row}";
    daily.Cell(row, 9).FormulaA1 = $"C{row}-G{row}";
    daily.Cell(row, 10).FormulaA1 = $"G{row}-F{row}";
    daily.Cell(row, 11).FormulaA1 = $"H{row}-I{row}-J{row}";
    daily.Cell(row, 12).Value = d.NoiseSd;
    daily.Cell(row, 13).FormulaA1 = $"IF(L{row}>0,H{row}/L{row},0)";
    daily.Cell(row, 14).FormulaA1 = $"ABS(M{row})";
    daily.Cell(row, 15).FormulaA1 =
        $"IF(T{row}=2024,IF(ABS(H{row})>'03 Morning Drilldown'!$C$10*'03 Morning Drilldown'!$C$11,\"RED\",\"OK\"),\"BASELINE\")";
    daily.Cell(row, 16).FormulaA1 =
        $"IF(T{row}=2024,IF(ABS(M{row})>'03 Morning Drilldown'!$C$10,\"RED\",\"OK\"),\"BASELINE\")";
    daily.Cell(row, 17).FormulaA1 =
        $"IF(T{row}<>2024,\"BASELINE\",IF(AND(O{row}=\"RED\",P{row}=\"RED\"),\"BOTH\",IF(O{row}=\"RED\",\"FIXED ONLY\",IF(P{row}=\"RED\",\"SCALED ONLY\",\"CLEAR\"))))";
    daily.Cell(row, 18).FormulaA1 = $"IF(ABS(K{row})<0.01,\"PASS\",\"CHECK\")";
    daily.Cell(row, 19).Value = int.Parse(d.Date.Substring(5, 2), CultureInfo.InvariantCulture);
    daily.Cell(row, 20).Value = int.Parse(d.Date.Substring(0, 4), CultureInfo.InvariantCulture);
    daily.Range(row, 3, row, 12).Style.NumberFormat.Format = PnlTheme.Money;
    daily.Cell(row, 11).Style.NumberFormat.Format = PnlTheme.Precise;
    daily.Cell(row, 13).Style.NumberFormat.Format = PnlTheme.Zformat;
    daily.Cell(row, 14).Style.NumberFormat.Format = PnlTheme.Zformat;
    PnlTheme.Stripe(daily, row, 2, 20, i);
    daily.Row(row).Height = 21;
}
daily.Range(8, 2, LAST, 20).SetAutoFilter();
daily.SheetView.FreezeRows(8);

// ============================================================================
// 07 | RAW TENOR INPUTS (HIDDEN) - FULLY REPRODUCIBLE COMPONENT OF EACH DAY
// B date; C:M prior ladder; N:X curve moves in bp; Y:AI linear curve P&L.
// ============================================================================
tenorData.Cell("B7").Value = "BOND RISK ENGINE OUTPUT / TASK 3";
tenorData.Cell("B8").Value = "DATE";
for (int k = 0; k < 11; k++)
{
    tenorData.Cell(8, 3 + k).Value = $"DV01 {tenors[k]}";
    tenorData.Cell(8, 14 + k).Value = $"MOVE {tenors[k]}";
    tenorData.Cell(8, 25 + k).Value = $"P&L {tenors[k]}";
}
for (int i = 0; i < days.Count; i++)
{
    int row = FIRST + i;
    tenorData.Cell(row, 2).Value = days[i].Date;
    for (int k = 0; k < 11; k++)
    {
        tenorData.Cell(row, 3 + k).Value = days[i].TenorDv01[k];
        tenorData.Cell(row, 14 + k).Value = days[i].CurveMoveBp[k];
        tenorData.Cell(row, 25 + k).Value = days[i].TenorPnl[k];
    }
}

// ============================================================================
// 03 | INTERACTIVE MORNING DRILLDOWN
// C9: ISO DATE dropdown. C10: user-editable sigma threshold.
// C11: 2023 std(U); C12: fitted per-vertex noise in bp.
// ============================================================================
PnlTheme.Banner(drill, "MORNING RECONCILIATION / SELECT A REPORT DATE");
PnlTheme.Section(drill, "B7:L7", "BLUE CELLS ARE INTERACTIVE | SELECT DATE, ADJUST THE SIGMA THRESHOLD");
drill.Column("B").Width = 27;
drill.Column("C").Width = 19;
drill.Column("E").Width = 26;
drill.Columns("F:H").Width = 20;
drill.Column("I").Width = 17;
drill.Column("J").Width = 22;
for (int k = 0; k < days.Count; k++)
    drill.Cell(FIRST + k, 52).Value = days[k].Date; // AZ is column 52

drill.Cell("B9").Value = "REPORT DATE";
drill.Cell("C9").Value = days[^1].Date;
drill.Cell("C9").CreateDataValidation().List(drill.Range($"AZ{FIRST}:AZ{LAST}"));
PnlTheme.Input(drill.Cell("C9"));
drill.Cell("B10").Value = "ALERT THRESHOLD / SIGMA";
drill.Cell("C10").Value = data.DefaultSigma;
drill.Cell("C10").Style.NumberFormat.Format = "0.0";
PnlTheme.Input(drill.Cell("C10"));
// Restrict the input to a reasonable positive range. ClosedXML validation
// is optional here; Excel uses the editable numeric cell as the threshold.
drill.Cell("B11").Value = "2023 STD / UNEXPLAINED";
drill.Cell("C11").Value = data.TrainingStdUnexplained;
drill.Cell("C11").Style.NumberFormat.Format = PnlTheme.Money;
drill.Cell("B12").Value = "FITTED NOISE / BP";
drill.Cell("C12").Value = data.FitNoiseBp;
drill.Cell("C12").Style.NumberFormat.Format = "0.0000";
drill.Cell("E9").Value = "SELECTED REPORT #";
drill.Cell("G9").FormulaA1 = $"MATCH($C$9,$AZ${FIRST}:$AZ${LAST},0)";
drill.Cell("E10").Value = "FIXED USD CUTOFF";
drill.Cell("H10").FormulaA1 = "$C$10*$C$11";
drill.Cell("H10").Style.NumberFormat.Format = PnlTheme.Money;
drill.Cell("E11").Value = "SELECTED FLAG";
drill.Cell("H11").FormulaA1 = Lookup("Q");
drill.Cell("E12").Value = "RECONCILIATION";
drill.Cell("H12").FormulaA1 = Lookup("R");
drill.Column("AZ").Hide();

string Lookup(string col) => $"INDEX('02 Daily Explain'!${col}${FIRST}:${col}${LAST},$G$9)";
PnlTheme.Section(drill, "B15:L15", "REPORTED VS EXPECTED | EXPLAIN THE UNEXPLAINED");
drill.Cell("B17").Value = "METRIC";
drill.Cell("F17").Value = "SELECTED DATE / VALUE";
PnlTheme.Header(drill, "B17:L17");
var entries = new (string Label, string Column, int Row)[] {
    ("Reported desk P&L", "C", 18),
    ("Carry / time passage", "D", 19),
    ("Linear curve P&L", "E", 20),
    ("Predicted P&L", "F", 21),
    ("Full-revaluation P&L", "G", 22),
    ("Unexplained P&L", "H", 23),
    ("Data gap: reported - full", "I", 24),
    ("Model gap: full - predicted", "J", 25),
    ("Expected noise standard deviation", "L", 26),
    ("Noise-scaled Z-score", "M", 27)
};
for (int i = 0; i < entries.Length; i++)
{
    var x = entries[i];
    drill.Cell(x.Row, 2).Value = x.Label;
    drill.Cell(x.Row, 6).FormulaA1 = Lookup(x.Column);
    drill.Cell(x.Row, 6).Style.NumberFormat.Format = x.Row == 27 ? PnlTheme.Zformat : PnlTheme.Money;
    PnlTheme.Stripe(drill, x.Row, 2, 12, i);
    drill.Row(x.Row).Height = 24;
}
PnlTheme.Total(drill, 23, 2, 12);
PnlTheme.Section(drill, "B30:L30", "INTERPRETATION / INVESTIGATION PRIORITIES");
string[] drillNotes = {
    "Predicted P&L = Carry + (- previous-day key-rate DV01 × observed curve move in bp).",
    "Unexplained P&L = Reported P&L − Predicted P&L.",
    "Data gap = Reported P&L − Full-revaluation P&L; Model gap = Full − Predicted.",
    "Z-score = Unexplained P&L / expected noise SD, calibrated from 2023.",
    "Changing the sigma threshold updates 2024 flags across the workbook; it does not refit 2023.",
    "2023 is the calibration period; 2024 is out-of-sample evaluation, not a trading-limit breach."
};
for (int i = 0; i < drillNotes.Length; i++)
{
    drill.Range(32 + i, 2, 32 + i, 12).Merge();
    drill.Cell(32 + i, 2).Value = drillNotes[i];
    drill.Cell(32 + i, 2).Style.Font.FontColor = PnlTheme.Muted;
    drill.Row(32 + i).Height = 24;
}

// ============================================================================
// 04 | DYNAMIC KEY-RATE MARKET P&L ATTRIBUTION
// Reference 07 Tenor Data; date selector shared with Morning Drilldown.
// ============================================================================
PnlTheme.Banner(tenor, "KEY-RATE P&L / 11-TENOR CONTRIBUTION");
PnlTheme.Section(tenor, "B7:L7", "MARKET P&L FOR THE SELECTED DATE | PREVIOUS-DAY LADDER × CURVE MOVE");
tenor.Cell("B9").Value = "REPORT DATE";
tenor.Cell("D9").FormulaA1 = "'03 Morning Drilldown'!C9";
tenor.Cell("F9").Value = "SELECTED REPORT #";
tenor.Cell("H9").FormulaA1 = "'03 Morning Drilldown'!G9";
tenor.Columns("B:B").Width = 18;
tenor.Column("C").Width = 4;
tenor.Columns("D:F").Width = 23;
tenor.Column("G").Width = 18;
tenor.Column("H").Width = 38;
var tenorHeaders = new[] {"TENOR", "PREV DV01 (USD/BP)", "CURVE MOVE (BP)", "MARKET P&L (USD)", "ABS SHARE", "MAGNITUDE"};
var tenorCols = new[] {2, 4, 5, 6, 7, 8};
for (int i = 0; i < tenorHeaders.Length; i++) tenor.Cell(11, tenorCols[i]).Value = tenorHeaders[i];
PnlTheme.Header(tenor, "B11:H11");
for (int k = 0; k < 11; k++)
{
    int row = 12 + k;
    string dvCol = ExcelColumn(3 + k);
    string moveCol = ExcelColumn(14 + k);
    string pnlCol = ExcelColumn(25 + k);
    tenor.Cell(row, 2).Value = tenors[k];
    tenor.Cell(row, 4).FormulaA1 = $"INDEX('07 Tenor Data'!${dvCol}${FIRST}:${dvCol}${LAST},$H$9)";
    tenor.Cell(row, 5).FormulaA1 = $"INDEX('07 Tenor Data'!${moveCol}${FIRST}:${moveCol}${LAST},$H$9)";
    tenor.Cell(row, 6).FormulaA1 = $"-D{row}*E{row}";
    string denom = string.Join(",", Enumerable.Range(12, 11).Select(r => $"ABS($F${r})"));
    tenor.Cell(row, 7).FormulaA1 = $"IF(SUM({denom})=0,0,ABS(F{row})/SUM({denom}))";
    // The positive/negative sign remains in F; bars represent ABS magnitude only.
    tenor.Cell(row, 8).FormulaA1 = $"REPT(\"■\",MIN(26,ROUND(26*G{row},0)))";
    tenor.Cell(row, 8).Style.Font.FontColor = PnlTheme.Teal;
    tenor.Cell(row, 4).Style.NumberFormat.Format = PnlTheme.Money;
    tenor.Cell(row, 5).Style.NumberFormat.Format = "+0.00;[Red](0.00);-";
    tenor.Cell(row, 6).Style.NumberFormat.Format = PnlTheme.Money;
    tenor.Cell(row, 7).Style.NumberFormat.Format = "0.0%";
    PnlTheme.Stripe(tenor, row, 2, 8, k);
}
tenor.Cell("B24").Value = "MARKET P&L TOTAL";
tenor.Cell("F24").FormulaA1 = "SUM(F12:F22)";
tenor.Cell("F24").Style.NumberFormat.Format = PnlTheme.Money;
PnlTheme.Total(tenor, 24, 2, 8);
tenor.Cell("B26").Value = "RECONCILIATION VS DAILY REPORT";
tenor.Cell("F26").FormulaA1 = "F24-'03 Morning Drilldown'!F20";
tenor.Cell("F26").Style.NumberFormat.Format = PnlTheme.Precise;
tenor.Cell("G26").FormulaA1 = "IF(ABS(F26)<0.01,\"PASS\",\"CHECK\")";
PnlTheme.Section(tenor, "B29:L29", "RISK NOTE");
tenor.Range("B31:L31").Merge();
tenor.Cell("B31").Value = "Linear DV01 attribution; omits convexity and non-parallel nonlinear repricing effects.";

// ============================================================================
// 01 | EXECUTIVE SNAPSHOT + CUMULATIVE EXCEPTION COUNTS
// All user-adjustable alert totals are Excel formulas, not hardcoded values.
// ============================================================================
PnlTheme.Banner(overview, "RATES DESK / DAILY P&L EXPLAIN MONITOR");
PnlTheme.Section(overview, "B7:L7", "EOD P&L & RECONCILIATION | CLOSE OF " + data.AsOf);
PnlTheme.KpiFormula(overview, "B8:F8", "B9:F11", "REPORTED P&L / USD", $"'02 Daily Explain'!C{LAST}");
PnlTheme.KpiFormula(overview, "H8:L8", "H9:L11", "PREDICTED P&L / USD", $"'02 Daily Explain'!F{LAST}");
PnlTheme.KpiFormula(overview, "B13:F13", "B14:F16", "UNEXPLAINED P&L / USD", $"'02 Daily Explain'!H{LAST}");
PnlTheme.KpiFormula(overview, "H13:L13", "H14:L16", "NOISE-ADJUSTED Z-SCORE", $"'02 Daily Explain'!M{LAST}", PnlTheme.Zformat);
PnlTheme.Section(overview, "B18:L18", "2024 EXCEPTION SCREENING | CURRENT SIGMA THRESHOLD");
overview.Cell("B20").Value = "CURRENT SIGMA THRESHOLD";
overview.Cell("E20").FormulaA1 = "'03 Morning Drilldown'!C10";
overview.Cell("H20").Value = "FITTED NOISE / BP";
overview.Cell("K20").Value = data.FitNoiseBp;
overview.Cell("K20").Style.NumberFormat.Format = "0.0000";
overview.Cell("B22").Value = "FIXED USD FLAGS";
overview.Cell("E22").FormulaA1 = $"COUNTIF('02 Daily Explain'!O{first2024}:O{last2024},\"RED\")";
overview.Cell("H22").Value = "RISK-SCALED FLAGS";
overview.Cell("K22").FormulaA1 = $"COUNTIF('02 Daily Explain'!P{first2024}:P{last2024},\"RED\")";
overview.Cell("B23").Value = "FIXED ONLY";
overview.Cell("E23").FormulaA1 = $"COUNTIF('02 Daily Explain'!Q{first2024}:Q{last2024},\"FIXED ONLY\")";
overview.Cell("H23").Value = "SCALED ONLY";
overview.Cell("K23").FormulaA1 = $"COUNTIF('02 Daily Explain'!Q{first2024}:Q{last2024},\"SCALED ONLY\")";
overview.Cell("B24").Value = "BASELINE (2023)";
overview.Cell("E24").Value = trainN;
overview.Cell("H24").Value = "ASSESSMENT (2024)";
overview.Cell("K24").Value = assessN;
PnlTheme.Info(overview, "B26:L26", "These are investigation flags, not risk-limit breaches. Change the threshold at 03 Morning Drilldown / C10.");
PnlTheme.Section(overview, "B29:L29", "2024 NOISE-ADJUSTED UNEXPLAINED P&L | STATIC 3.5σ REFERENCE CHART");
if (!string.IsNullOrWhiteSpace(data.ZChart))
{
    using var chart = new MemoryStream(Convert.FromBase64String(data.ZChart));
    overview.AddPicture(chart, XLPictureFormat.Png).MoveTo(overview.Cell("B30")).Scale(0.58);
}
PnlTheme.Section(overview, "B57:L57", "2023 CALIBRATION / MODEL QUALITY");
overview.Cell("B59").Value = "2023 RMS REPORTED P&L";
overview.Cell("E59").Value = data.TrainingRmsReported;
overview.Cell("B60").Value = "2023 RMS UNEXPLAINED";
overview.Cell("E60").Value = data.TrainingRmsUnexplained;
overview.Cell("B61").Value = "2023 RMS FULL-REVAL GAP";
overview.Cell("E61").Value = data.TrainingRmsFullGap;
overview.Cell("B62").Value = "2023 MEAN DAILY CARRY";
overview.Cell("E62").Value = data.TrainingMeanCarry;
foreach (int row in new[] {20, 22, 23, 24})
{
    overview.Range(row, 2, row, 4).Merge();
    overview.Range(row, 8, row, 10).Merge();
}
for (int row = 59; row <= 62; row++)
{
    overview.Range(row, 2, row, 4).Merge();
    overview.Cell(row, 5).Style.NumberFormat.Format = PnlTheme.Money;
}
PnlTheme.Info(overview, "B65:L65", "Model: carry + linear curve DV01. Sigma fitted on 2023; assessment in 2024. Simulated data / not live risk.");

// ============================================================================
// 05 | 20 BIGGEST 2024 NOISE-ADJUSTED MOVES + MONTHLY FLAGS
// Ranking remains |z|-based, while classifications respond to user threshold.
// ============================================================================
PnlTheme.Banner(exceptions, "EXCEPTION MONITOR / OUT-OF-SAMPLE 2024");
PnlTheme.Section(exceptions, "B7:M7", "TOP 20 DAYS BY ABSOLUTE Z-SCORE | THRESHOLD-RESPONSIVE CLASSIFICATION");
exceptions.Columns("B:B").Width = 16;
exceptions.Columns("C:J").Width = 21;
exceptions.Column("K").Width = 22;
exceptions.Columns("L:M").Width = 17;
exceptions.Cell("B9").Value = "THRESHOLD";
exceptions.Cell("D9").FormulaA1 = "'03 Morning Drilldown'!C10";
exceptions.Cell("F9").Value = "2024 ASSESSMENT DAYS";
exceptions.Cell("J9").Value = assessN;
var top = days.Select((d, i) => new {Day=d, N=i+1})
    .Where(x => x.Day.Date.StartsWith("2024-", StringComparison.Ordinal))
    .OrderByDescending(x => Math.Abs(x.Day.ZScore)).ThenBy(x => x.Day.Date)
    .Take(20).ToList();
string[] rankHeaders = {"DATE", "REPORTED P&L", "UNEXPLAINED", "EXPECTED SD", "ABS Z", "FIXED", "SCALED", "CLASS", "RECON", "RANK", "REPORT #"};
for (int i = 0; i < rankHeaders.Length; i++) exceptions.Cell(11, i + 2).Value = rankHeaders[i];
PnlTheme.Header(exceptions, "B11:L11");
foreach (var rec in top.Select((x, idx) => new {x.Day, x.N, Rank=idx+1}))
{
    int row = 11 + rec.Rank;
    exceptions.Cell(row, 2).Value = DateTime.ParseExact(rec.Day.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    exceptions.Cell(row, 2).Style.DateFormat.Format = "yyyy-mm-dd";
    exceptions.Cell(row, 12).Value = rec.N;
    exceptions.Cell(row, 3).FormulaA1 = $"INDEX('02 Daily Explain'!$C${FIRST}:$C${LAST},$L{row})";
    exceptions.Cell(row, 4).FormulaA1 = $"INDEX('02 Daily Explain'!$H${FIRST}:$H${LAST},$L{row})";
    exceptions.Cell(row, 5).FormulaA1 = $"INDEX('02 Daily Explain'!$L${FIRST}:$L${LAST},$L{row})";
    exceptions.Cell(row, 6).FormulaA1 = $"INDEX('02 Daily Explain'!$N${FIRST}:$N${LAST},$L{row})";
    exceptions.Cell(row, 7).FormulaA1 = $"INDEX('02 Daily Explain'!$O${FIRST}:$O${LAST},$L{row})";
    exceptions.Cell(row, 8).FormulaA1 = $"INDEX('02 Daily Explain'!$P${FIRST}:$P${LAST},$L{row})";
    exceptions.Cell(row, 9).FormulaA1 = $"INDEX('02 Daily Explain'!$Q${FIRST}:$Q${LAST},$L{row})";
    exceptions.Cell(row, 10).FormulaA1 = $"INDEX('02 Daily Explain'!$R${FIRST}:$R${LAST},$L{row})";
    exceptions.Cell(row, 11).Value = rec.Rank;
    exceptions.Range(row, 3, row, 5).Style.NumberFormat.Format = PnlTheme.Money;
    exceptions.Cell(row, 6).Style.NumberFormat.Format = PnlTheme.Zformat;
    PnlTheme.Stripe(exceptions, row, 2, 12, rec.Rank);
}
exceptions.Range(11, 2, 11+top.Count, 12).SetAutoFilter();
exceptions.SheetView.FreezeRows(11);
PnlTheme.Section(exceptions, "B35:L35", "MONTHLY 2024 RED FLAGS / FIXED VS SCALED");
exceptions.Cell("B37").Value = "MONTH";
exceptions.Cell("D37").Value = "FIXED";
exceptions.Cell("F37").Value = "SCALED";
exceptions.Cell("H37").Value = "FIXED ONLY";
PnlTheme.Header(exceptions, "B37:J37");
for (int m = 1; m <= 12; m++)
{
    int row = 37 + m;
    exceptions.Cell(row, 2).Value = new DateTime(2024, m, 1).ToString("MMM", CultureInfo.InvariantCulture);
    exceptions.Cell(row, 4).FormulaA1 = $"COUNTIFS('02 Daily Explain'!$S${first2024}:$S${last2024},{m},'02 Daily Explain'!$O${first2024}:$O${last2024},\"RED\")";
    exceptions.Cell(row, 6).FormulaA1 = $"COUNTIFS('02 Daily Explain'!$S${first2024}:$S${last2024},{m},'02 Daily Explain'!$P${first2024}:$P${last2024},\"RED\")";
    exceptions.Cell(row, 8).FormulaA1 = $"COUNTIFS('02 Daily Explain'!$S${first2024}:$S${last2024},{m},'02 Daily Explain'!$Q${first2024}:$Q${last2024},\"FIXED ONLY\")";
    PnlTheme.Stripe(exceptions, row, 2, 10, m);
}
exceptions.Cell("B51").Value = "TOTAL";
foreach (int col in new[] {4,6,8})
{
    string l = ExcelColumn(col);
    exceptions.Cell(51, col).FormulaA1 = $"SUM({l}38:{l}49)";
}
PnlTheme.Total(exceptions, 51, 2, 10);
PnlTheme.Info(exceptions, "B54:L54", "Flags identify investigation candidates. This workbook has no approved trading limits, workflow routing or live feed.");

// ============================================================================
// 06 | BUSINESS AND TECHNICAL METHODOLOGY
// ============================================================================
PnlTheme.Banner(method, "METHODOLOGY / P&L EXPLAIN AND NOISE MODEL");
method.Column("B").Width = 28;
method.Column("C").Width = 70;
method.Column("D").Width = 48;
PnlTheme.Section(method, "B7:D7", "HOW THE ENGINE AND WORKBOOK FIT TOGETHER");
var notes = new (string Label, string Detail)[] {
    ("Source", "Simulated government bond portfolio, amended ticket blotter and daily zero yield curves."),
    ("History", "2023: 258 report days to fit noise. 2024: 260 out-of-sample report days."),
    ("Reported desk P&L", "External PNL_REPORTED series. Does not necessarily equal the model's daily P&L."),
    ("Carry / time P&L", "Yesterday's positions multiplied by (today's aged price on prior curve + paid cashflows - yesterday's price)."),
    ("Market / linear P&L", "Negative dot-product of yesterday's USD/bp key-rate DV01 ladder with today's 11 zero-curve moves in bp."),
    ("Predicted P&L", "Carry plus Market / linear P&L. New trade P&L is not separately attributed in this Task 3 model."),
    ("Full revaluation", "Carry plus yesterday's positions multiplied by today's price minus the aged price on yesterday's curve."),
    ("Unexplained P&L", "Reported P&L minus predicted P&L."),
    ("Data gap", "Reported P&L minus full-revaluation P&L."),
    ("Model gap", "Full-revaluation P&L minus predicted P&L (linearization / modeling gap)."),
    ("Exact reconciliation", "Unexplained P&L equals Data Gap plus Model Gap. The Excel arithmetic residual is checked to USD 0.01."),
    ("2023 noise calibration", "s = std_2023(Unexplained) / sqrt(2 × mean_2023(sum(previous DV01 squared))). Uses population std (ddof=0)."),
    ("Expected noise SD", "sqrt(2) × s × sqrt(sum(previous-day DV01 squared)) for each report day."),
    ("Scaled Z-score", "Unexplained P&L / expected noise SD. Positive or negative extreme values are candidates for investigation."),
    ("Fixed flags", "2024: abs(Unexplained) > user sigma threshold × 2023 population std of unexplained."),
    ("Scaled flags", "2024: abs(Z-score) > user sigma threshold. The scale adapts to daily portfolio risk."),
    ("Editable control", "03 Morning Drilldown!C10 changes flag thresholds, counts and classifications. Fitted 2023 noise s remains fixed."),
    ("Top 20", "05 Exception Monitor ranks 2024 dates by absolute Z-score; rank is independent of user threshold."),
    ("11-tenor attribution", "04 Tenor Attribution recomputes linear market P&L from per-tenor previous DV01 and yield move."),
    ("Chart note", "Overview time-series image is a static 3.5-sigma reference, not a chart linked to adjustable threshold."),
    ("Risk limitations", "First-order DV01, simplified pricing, no convexity, explicit credit, liquidity, funding, fees or trade-day P&L attribution."),
    ("Non-production", "Educational simulated data; screening flags are not validated risk limits, adjudicated incidents or live controls."),
    ("Deployment", "Python / NumPy / pandas calculates -> JSON snapshot -> .NET 10 / ClosedXML -> macro-free Excel workbook.")
};
method.Cell("B9").Value = "TOPIC";
method.Cell("C9").Value = "METHOD / BUSINESS MEANING";
PnlTheme.Header(method, "B9:D9");
for (int i = 0; i < notes.Length; i++)
{
    int row = 10 + i;
    method.Cell(row, 2).Value = notes[i].Label;
    method.Range(row, 3, row, 4).Merge();
    method.Cell(row, 3).Value = notes[i].Detail;
    method.Range(row, 2, row, 4).Style.Alignment.WrapText = true;
    method.Row(row).Height = 39;
    PnlTheme.Stripe(method, row, 2, 4, i);
}
method.SheetView.FreezeRows(9);

// Hide the reproducible tenor source table. Interactive reports reference it.
tenorData.Hide();
Directory.CreateDirectory(outputDir);
string destination = Path.Combine(outputDir, $"RatesDesk_PnLExplain_Task3_{data.AsOf}.xlsx");
wb.SaveAs(destination);
Console.WriteLine("SUCCESS / TASK 3 P&L EXPLAIN");
Console.WriteLine($"Workbook: {destination}");
Console.WriteLine($"Calibration 2023: {trainN} days | Assessment 2024: {assessN} days");
Console.WriteLine($"Fitted noise s: {data.FitNoiseBp:F4} bp | threshold: {data.DefaultSigma:F1} sigma");
Console.WriteLine($"2024 baseline fixed flags: {fixedAtDefault} | scaled flags: {scaledAtDefault}");
Console.WriteLine("Use 03 Morning Drilldown C9 for date and C10 for editable screening threshold.");

static string ExcelColumn(int n)
{
    string result = "";
    for (; n > 0; n = (n - 1) / 26)
        result = (char)('A' + (n - 1) % 26) + result;
    return result;
}

// ============================================================================
// JSON MODEL CONTRACT - PROVIDED BY THE JUPYTER TASK 3 EXPORT CELL
// ============================================================================
public sealed class PnlSnapshot
{
    [JsonPropertyName("asOf")] public string AsOf { get; set; } = "";
    [JsonPropertyName("tenors")] public List<string> Tenors { get; set; } = new();
    [JsonPropertyName("daily")] public List<PnlDay> Daily { get; set; } = new();
    [JsonPropertyName("fitNoiseBp")] public double FitNoiseBp { get; set; }
    [JsonPropertyName("trainingStdUnexplained")] public double TrainingStdUnexplained { get; set; }
    [JsonPropertyName("defaultSigma")] public double DefaultSigma { get; set; }
    [JsonPropertyName("trainingDays")] public int TrainingDays { get; set; }
    [JsonPropertyName("assessmentDays")] public int AssessmentDays { get; set; }
    [JsonPropertyName("defaultFixedCount")] public int DefaultFixedCount { get; set; }
    [JsonPropertyName("defaultScaledCount")] public int DefaultScaledCount { get; set; }
    [JsonPropertyName("trainingRmsReported")] public double TrainingRmsReported { get; set; }
    [JsonPropertyName("trainingRmsUnexplained")] public double TrainingRmsUnexplained { get; set; }
    [JsonPropertyName("trainingRmsFullGap")] public double TrainingRmsFullGap { get; set; }
    [JsonPropertyName("trainingMeanCarry")] public double TrainingMeanCarry { get; set; }
    [JsonPropertyName("zChart")] public string ZChart { get; set; } = "";
}
public sealed class PnlDay
{
    [JsonPropertyName("date")] public string Date { get; set; } = "";
    [JsonPropertyName("reportedPnl")] public double ReportedPnl { get; set; }
    [JsonPropertyName("carry")] public double Carry { get; set; }
    [JsonPropertyName("marketPnl")] public double MarketPnl { get; set; }
    [JsonPropertyName("predictedPnl")] public double PredictedPnl { get; set; }
    [JsonPropertyName("fullRevaluation")] public double FullRevaluation { get; set; }
    [JsonPropertyName("unexplained")] public double Unexplained { get; set; }
    [JsonPropertyName("noiseSd")] public double NoiseSd { get; set; }
    [JsonPropertyName("zScore")] public double ZScore { get; set; }
    [JsonPropertyName("tenorDv01")] public List<double> TenorDv01 { get; set; } = new();
    [JsonPropertyName("curveMoveBp")] public List<double> CurveMoveBp { get; set; } = new();
    [JsonPropertyName("tenorPnl")] public List<double> TenorPnl { get; set; } = new();
}

public static class PnlTheme
{
    public const string Money = "+$#,##0;[Red]($#,##0);-";
    public const string Precise = "+#,##0.0000;[Red](#,##0.0000);-";
    public const string Zformat = "+0.00;[Red](0.00);-";
    public static readonly XLColor Navy = XLColor.FromHtml("#14263D");
    public static readonly XLColor Teal = XLColor.FromHtml("#078C86");
    public static readonly XLColor Muted = XLColor.FromHtml("#64748B");
    static readonly XLColor Blue = XLColor.FromHtml("#244767");
    static readonly XLColor Pale = XLColor.FromHtml("#EDF3F8");

    public static void Setup(IXLWorksheet ws, string date, string[] tabs)
    {
        ws.ShowGridLines = false;
        ws.Style.Font.FontName = "Aptos";
        ws.Style.Font.FontSize = 10;
        ws.Style.Font.FontColor = Navy;
        ws.Column(1).Width = 3;
        ws.Columns("B:L").Width = 14;
        ws.Row(2).Height = 29;
        ws.Row(3).Height = 26;
        ws.Cell("B4").Value = $"AS OF {date}  |  FIXED INCOME  |  SIMULATED DESK  |  TASK 3";
        ws.Cell("B4").Style.Font.FontColor = Muted;
        string[] tabCells = {"B5", "D5", "F5", "H5", "J5", "L5"};
        string[] labels = {"OVERVIEW", "DAILY", "DRILLDOWN", "TENORS", "EXCEPTIONS", "METHOD"};
        for (int i = 0; i < tabs.Length; i++)
        {
            var cell = ws.Cell(tabCells[i]);
            cell.FormulaA1 = $"HYPERLINK(\"#'{tabs[i]}'!B2\",\"{labels[i]}\")";
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = Blue;
            cell.Style.Fill.BackgroundColor = Pale;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
    }
    public static void Banner(IXLWorksheet ws, string text)
    {
        var r = ws.Range("B2:L3");
        r.Merge(); r.FirstCell().Value = text;
        r.Style.Fill.BackgroundColor = Navy;
        r.Style.Font.FontColor = XLColor.White;
        r.Style.Font.FontSize = 18;
        r.Style.Font.Bold = true;
        r.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Section(IXLWorksheet ws, string range, string title)
    {
        var r = ws.Range(range); r.Merge(); r.FirstCell().Value = title;
        r.Style.Fill.BackgroundColor = Blue;
        r.Style.Font.FontColor = XLColor.White;
        r.Style.Font.Bold = true;
        r.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Header(IXLWorksheet ws, string range)
    {
        var r = ws.Range(range);
        r.Style.Fill.BackgroundColor = Blue;
        r.Style.Font.FontColor = XLColor.White;
        r.Style.Font.Bold = true;
        r.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void KpiFormula(IXLWorksheet ws, string label, string valueArea,
                                  string name, string formula, string numberFormat = Money)
    {
        var a = ws.Range(label); a.Merge(); a.FirstCell().Value = name;
        a.Style.Fill.BackgroundColor = Pale;
        a.Style.Font.FontColor = Muted;
        a.Style.Font.Bold = true;
        var b = ws.Range(valueArea); b.Merge(); b.FirstCell().FormulaA1 = formula;
        b.Style.NumberFormat.Format = numberFormat;
        b.Style.Font.FontSize = 22;
        b.Style.Font.Bold = true;
        b.Style.Font.FontColor = Navy;
        b.Style.Fill.BackgroundColor = Pale;
        b.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        b.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Info(IXLWorksheet ws, string range, string note)
    {
        var r = ws.Range(range); r.Merge(); r.FirstCell().Value = note;
        r.Style.Font.FontColor = Muted;
        r.Style.Font.FontSize = 9;
    }
    public static void Stripe(IXLWorksheet ws, int row, int left, int right, int n)
    {
        if (n % 2 == 0)
            ws.Range(row, left, row, right).Style.Fill.BackgroundColor = XLColor.FromHtml("#F4F7FA");
    }
    public static void Total(IXLWorksheet ws, int row, int left, int right)
    {
        var r = ws.Range(row, left, row, right);
        r.Style.Fill.BackgroundColor = Pale;
        r.Style.Font.FontColor = Navy;
        r.Style.Font.Bold = true;
        r.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(row).Height = 27;
    }
    public static void Input(IXLCell cell)
    {
        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#E5F0FF");
        cell.Style.Font.FontColor = XLColor.FromHtml("#1255B0");
        cell.Style.Font.Bold = true;
        cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
    }
}
