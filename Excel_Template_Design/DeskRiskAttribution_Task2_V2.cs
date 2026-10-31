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

// =============================================================
// TASK 2 V2 / INTERACTIVE RATES DESK RISK ATTRIBUTION WORKSTATION
// Independent .NET 10 single-file application; no VBA required.
// Python computes Task 2 numbers, this code presents and checks them.
// =============================================================

string outputDirectory =
    @"C:\Users\franklinpeng\Desktop\RDA Risk test\Excel_Template_Output";
string sourcePath = Path.Combine(outputDirectory, "task2_snapshot.json");

if (!File.Exists(sourcePath))
    throw new FileNotFoundException(
        "Missing Task 2 snapshot. Run the export cell in risk_walk.ipynb first.",
        sourcePath);

var snapshot = JsonSerializer.Deserialize<Task2Snapshot>(
    File.ReadAllText(sourcePath))
    ?? throw new InvalidOperationException("Unable to read Task 2 snapshot.");

if (snapshot.Daily.Count == 0 || snapshot.Years.Count == 0 ||
    snapshot.Ladder2024.Count != 11)
    throw new InvalidOperationException(
        "Task 2 snapshot must contain daily rows, annual data and 11 curve vertices.");

var dailyRows = snapshot.Daily.OrderBy(x => x.Date).ToList();
var annualRows = snapshot.Years.OrderBy(x => x.Year).ToList();
var year24 = annualRows.FirstOrDefault(x => x.Year == 2024)
    ?? throw new InvalidOperationException("Missing 2024 annual attribution.");

// Fail fast if the provided Task 2 daily walks don't reconcile.
// A VaR walk is an exact telescoping decomposition of its five states.
foreach (var r in dailyRows)
{
    double dvEnd = r.DvStart + r.DvTime + r.DvMarket + r.DvTrades;
    double varEnd = r.VarStart + r.VarTime + r.VarMarket + r.VarWindow + r.VarTrades;
    if (Math.Abs(dvEnd - r.DvEnd) > 0.001 ||
        Math.Abs(varEnd - r.VarEnd) > 0.001)
        throw new InvalidOperationException($"Unbalanced walk on {r.Date}.");
}

if (snapshot.ChainMaxGap > 0.001)
    throw new InvalidOperationException(
        $"Previous-day ladder continuity failed: {snapshot.ChainMaxGap}");

using var wb = new XLWorkbook();
var overview = wb.Worksheets.Add("01 Attribution Overview");
var walk = wb.Worksheets.Add("02 Daily Walk");
var detail = wb.Worksheets.Add("03 Morning Drilldown");
var annual = wb.Worksheets.Add("04 Annual Attribution");
var method = wb.Worksheets.Add("05 Methodology");
var movers = wb.Worksheets.Add("06 Movers & Waterfall");

string[] sheetNames = {
    "01 Attribution Overview", "02 Daily Walk", "03 Morning Drilldown",
    "04 Annual Attribution", "05 Methodology", "06 Movers & Waterfall"
};

foreach (var ws in wb.Worksheets)
    RiskStyle.Setup(ws, snapshot.AsOf, sheetNames);

// =============================================================
// 01 / ATTRIBUTION OVERVIEW
// What moved DV01 and VaR over 2024? Does the walk reconcile?
// =============================================================
RiskStyle.Banner(overview, "RATES DESK / RISK ATTRIBUTION");
RiskStyle.Section(overview, "B7:L7", "2024 YEAR-END | RECONCILED DV01 & 99% HISTORICAL VaR");

RiskStyle.Kpi(overview, "B8:F8", "B9:F11", "CLOSING NET DV01 / USD PER BP", year24.DvEnd);
RiskStyle.Kpi(overview, "H8:L8", "H9:L11", "CLOSING 1-DAY 99% VaR / USD", year24.VarEnd, true);
RiskStyle.Kpi(overview, "B13:F13", "B14:F16", "2024 CHANGE IN NET DV01 / USD PER BP",
    year24.DvEnd - year24.DvStart);
RiskStyle.Kpi(overview, "H13:L13", "H14:L16", "2024 CHANGE IN VaR / USD",
    year24.VarEnd - year24.VarStart, true);

RiskStyle.Info(overview, "B17:L17",
    $"2024: {year24.ReportDays} report dates  |  " +
    $"DV01 chain gap: {snapshot.ChainMaxGap:G3} USD/bp  |  " +
    "Model: linear key-rate DV01, 500 historical moves");

RiskStyle.Section(overview, "B19:L19", "CUMULATIVE VaR ATTRIBUTION | 2023–2024");
if (!string.IsNullOrWhiteSpace(snapshot.CumulativeChart))
{
    using var pic = new MemoryStream(Convert.FromBase64String(snapshot.CumulativeChart));
    overview.AddPicture(pic, XLPictureFormat.Png)
        .MoveTo(overview.Cell("B20")).Scale(0.56);
}

RiskStyle.Section(overview, "B42:L42", "2024 ANNUAL ATTRIBUTION | CONTRIBUTION BY DRIVER");
string[] overviewColumns = { "DRIVER", "DV01 CHANGE / USD/BP", "VaR CHANGE / USD" };
for (int k = 0; k < overviewColumns.Length; k++)
    overview.Cell(44, new[] { 2, 5, 9 }[k]).Value = overviewColumns[k];
RiskStyle.Header(overview, "B44:L44");

var drivers = new (string Label, double Dv, double Var)[]
{
    ("Time / Aging", year24.DvTime, year24.VarTime),
    ("Market / Repricing", year24.DvMarket, year24.VarMarket),
    ("Historical Window Roll", 0, year24.VarWindow),
    ("Trades / Amendments", year24.DvTrades, year24.VarTrades)
};
for (int k = 0; k < drivers.Length; k++)
{
    int r = 45 + k;
    overview.Cell(r, 2).Value = drivers[k].Label;
    overview.Cell(r, 5).Value = drivers[k].Dv;
    overview.Cell(r, 9).Value = drivers[k].Var;
    overview.Cell(r, 5).Style.NumberFormat.Format = RiskStyle.Num;
    overview.Cell(r, 9).Style.NumberFormat.Format = RiskStyle.Money;
    RiskStyle.Stripe(overview, r, 2, 12, k);
}
overview.Cell("B50").Value = "TOTAL CHANGE";
overview.Cell("E50").FormulaA1 = "SUM(E45:E48)";
overview.Cell("I50").FormulaA1 = "SUM(I45:I48)";
overview.Cell("E50").Style.NumberFormat.Format = RiskStyle.Num;
overview.Cell("I50").Style.NumberFormat.Format = RiskStyle.Money;
RiskStyle.Total(overview, 50, 2, 12);
overview.Cell("B52").Value = "Reconciliation check";
overview.Cell("E52").FormulaA1 = $"E50-({year24.DvEnd.ToString("R", CultureInfo.InvariantCulture)}-{year24.DvStart.ToString("R", CultureInfo.InvariantCulture)})";
overview.Cell("I52").FormulaA1 = $"I50-({year24.VarEnd.ToString("R", CultureInfo.InvariantCulture)}-{year24.VarStart.ToString("R", CultureInfo.InvariantCulture)})";
overview.Cell("E52").Style.NumberFormat.Format = RiskStyle.Decimal;
overview.Cell("I52").Style.NumberFormat.Format = RiskStyle.Decimal;
RiskStyle.Info(overview, "B55:L55",
    "Each driver is a sequential change in a risk measure, not standalone cash P&L. " +
    "VaR attribution depends on the stated order of steps.");

// =============================================================
// 02 / DAILY WALK / ALL 518 REPORT DATES
// Formula checks recalculate inside Excel, independently of JSON.
// =============================================================
RiskStyle.Banner(walk, "DAILY RISK WALK / AUDITABLE STATES");
RiskStyle.Section(walk, "B7:S7", "DV01 (USD/BP) AND VaR (USD) | PASS = WALK RESIDUAL WITHIN 0.01");
string[] dailyHeaders = {
    "REPORT DATE", "DV01 START", "TIME", "MARKET", "TRADES", "DV01 END",
    "CALC END", "RESIDUAL", "VaR START", "TIME", "MARKET", "WINDOW",
    "TRADES", "VaR END", "CALC END", "RESIDUAL", "DV01 FLAG", "VaR FLAG"
};
for (int k = 0; k < dailyHeaders.Length; k++)
    walk.Cell(8, k + 2).Value = dailyHeaders[k];
RiskStyle.Header(walk, "B8:S8");
walk.Columns("B:B").Width = 16;
walk.Columns("C:Q").Width = 17;
walk.Columns("R:S").Width = 15;

for (int k = 0; k < dailyRows.Count; k++)
{
    int r = 9 + k;
    var v = dailyRows[k];
    walk.Cell(r, 2).Value = DateTime.ParseExact(v.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    walk.Cell(r, 2).Style.DateFormat.Format = "yyyy-mm-dd";
    foreach (var x in new (int Col, double Value)[] {
        (3,v.DvStart),(4,v.DvTime),(5,v.DvMarket),(6,v.DvTrades),(7,v.DvEnd),
        (10,v.VarStart),(11,v.VarTime),(12,v.VarMarket),(13,v.VarWindow),
        (14,v.VarTrades),(15,v.VarEnd)
    })
    {
        walk.Cell(r, x.Col).Value = x.Value;
        walk.Cell(r, x.Col).Style.NumberFormat.Format = RiskStyle.Num;
    }
    // Excel-native formulas for auditability (no static hardcoded checks).
    walk.Cell(r, 8).FormulaA1 = $"SUM(C{r}:F{r})";
    walk.Cell(r, 9).FormulaA1 = $"G{r}-H{r}";
    walk.Cell(r, 16).FormulaA1 = $"SUM(J{r}:N{r})";
    walk.Cell(r, 17).FormulaA1 = $"O{r}-P{r}";
    // Distinguish an arithmetic reconciliation warning from a risk-limit breach.
    walk.Cell(r, 18).FormulaA1 = $"IF(ABS(I{r})<=0.01,\"PASS\",\"CHECK\")";
    walk.Cell(r, 19).FormulaA1 = $"IF(ABS(Q{r})<=0.01,\"PASS\",\"CHECK\")";
    foreach (int c in new[] { 8, 9, 16, 17 })
        walk.Cell(r, c).Style.NumberFormat.Format = RiskStyle.Decimal;
    RiskStyle.Stripe(walk, r, 2, 19, k);
    walk.Row(r).Height = 21;
}
int firstDailyRow = 9;
int lastDailyRow = firstDailyRow + dailyRows.Count - 1;
walk.Range(8, 2, lastDailyRow, 19).SetAutoFilter();
walk.SheetView.FreezeRows(8);

// =============================================================
// 03 / INTERACTIVE MORNING DRILLDOWN
// The selected report date is an ISO text value validated against a
// same-sheet helper range. MATCH converts it to the 1-based row number.
// This avoids locale-dependent date parsing in Excel and needs no VBA.
// =============================================================
RiskStyle.Banner(detail, "MORNING DRILLDOWN / FIVE-STATE VaR WALK");
RiskStyle.Section(detail, "B7:L7", "SELECT A DATE IN THE BLUE C9 DROPDOWN | VaR BRIDGE UPDATES AUTOMATICALLY");
// Excel validation lists are most reliable from a range on the SAME sheet.
// AZ is an invisible source list of report-date labels (YYYY-MM-DD).
for (int n = 0; n < dailyRows.Count; n++)
    detail.Cell($"AZ{firstDailyRow + n}").Value = dailyRows[n].Date;
detail.Cell("B9").Value = "REPORT DATE";
detail.Cell("C9").Value = dailyRows[^1].Date;
detail.Cell("C9").Style.Fill.BackgroundColor = XLColor.FromHtml("#E5F0FF");
detail.Cell("C9").Style.Font.FontColor = XLColor.FromHtml("#1255B0");
detail.Cell("C9").Style.Font.Bold = true;
detail.Column("C").Width = 19;
// The validated string is matched to the source list, not to locale-dependent dates.
detail.Cell("C9").CreateDataValidation()
    .List(detail.Range($"AZ{firstDailyRow}:AZ{lastDailyRow}"));
detail.Column("AZ").Hide();
detail.Cell("E9").Value = "MATCHED REPORT #";
detail.Cell("G9").FormulaA1 = $"MATCH($C$9,$AZ${firstDailyRow}:$AZ${lastDailyRow},0)";
detail.Columns("B:B").Width = 25;
detail.Columns("C:C").Width = 19;
detail.Columns("D:D").Width = 17;
detail.Columns("F:F").Width = 23;
detail.Columns("H:H").Width = 19;
detail.Columns("J:J").Width = 22;

string[] drillHeaders = { "WALK STATE", "DV01 DELTA", "DV01 RUNNING", "VaR DELTA", "VaR RUNNING" };
int[] drillCols = { 2, 4, 6, 8, 10 };
for (int k = 0; k < drillHeaders.Length; k++)
    detail.Cell(12, drillCols[k]).Value = drillHeaders[k];
RiskStyle.Header(detail, "B12:K12");

// INDEX reads each metric using the report index resolved from C9.
string Lookup(string col) =>
    $"INDEX('02 Daily Walk'!${col}${firstDailyRow}:${col}${lastDailyRow},$G$9)";

var labels = new[] {
    "Yesterday's report", "Time / one day older", "Market / new curve",
    "Historical window roll", "Trades / amendments", "Today's report"
};
for (int k = 0; k < labels.Length; k++)
{
    int r = 13 + k;
    detail.Cell(r, 2).Value = labels[k];
    RiskStyle.Stripe(detail, r, 2, 11, k);
    detail.Row(r).Height = 29;
}
// Yesterday's closing state; successive rows add one driver.
detail.Cell("D13").Value = 0;
detail.Cell("F13").FormulaA1 = Lookup("C");
detail.Cell("H13").Value = 0;
detail.Cell("J13").FormulaA1 = Lookup("J");

string[] dvChange = { Lookup("D"), Lookup("E"), "0", Lookup("F"), "0" };
string[] varChange = { Lookup("K"), Lookup("L"), Lookup("M"), Lookup("N"), "0" };
for (int k = 0; k < 5; k++)
{
    int r = 14 + k;
    detail.Cell(r, 4).FormulaA1 = dvChange[k];
    detail.Cell(r, 6).FormulaA1 = $"F{r-1}+D{r}";
    detail.Cell(r, 8).FormulaA1 = varChange[k];
    detail.Cell(r, 10).FormulaA1 = $"J{r-1}+H{r}";
}
foreach (int col in new[] { 4, 6, 8, 10 })
    detail.Range(13, col, 18, col).Style.NumberFormat.Format = RiskStyle.Num;
RiskStyle.Total(detail, 18, 2, 11);

detail.Cell("B21").Value = "END-STATE RECONCILIATION";
detail.Cell("F21").FormulaA1 = $"{Lookup("G")}-F18";
detail.Cell("J21").FormulaA1 = $"{Lookup("O")}-J18";
detail.Cell("F21").Style.NumberFormat.Format = RiskStyle.Decimal;
detail.Cell("J21").Style.NumberFormat.Format = RiskStyle.Decimal;
// Selected-date reconciliation status, recalculated inside Excel.
detail.Cell("B22").Value = "WALK RECONCILIATION";
detail.Cell("F22").FormulaA1 = "IF(AND(ABS(F21)<=0.01,ABS(J21)<=0.01),\"PASS\",\"CHECK\")";
detail.Cell("H22").Value = "Tolerance: 0.01";
detail.Cell("F22").Style.Font.Bold = true;
detail.Cell("F22").Style.Fill.BackgroundColor = XLColor.FromHtml("#E7F5EF");
RiskStyle.Section(detail, "B24:L24", "HOW TO READ THE WALK");
string[] explanatory = {
    "Time: same yesterday positions and yesterday curve, aged one business day.",
    "Market: reprice yesterday positions on the new curve.",
    "Window: change the historical 500-day sample for VaR only.",
    "Trades: add today's booked trades and amendments.",
    "VaR steps are sequential and are not independent risk measures.",
    "The report date represents the market close reviewed the following morning."
};
for (int k = 0; k < explanatory.Length; k++)
{
    detail.Range(26 + k, 2, 26 + k, 12).Merge();
    detail.Cell(26 + k, 2).Value = explanatory[k];
}
RiskStyle.Section(detail, "B34:L34", "INTERACTIVE VISUAL & TOP 10 RISK MOVERS");
detail.Range("B36:L36").Merge();
detail.Cell("B36").Value =
    "Open '06 Movers & Waterfall' to inspect the formula-driven floating VaR bridge and largest daily moves.";
detail.Cell("B36").Style.Font.FontColor = XLColor.FromHtml("#64748B");

// =============================================================
// 04 / ANNUAL ATTRIBUTION AND 2024 LADDER PER VERTEX
// =============================================================
RiskStyle.Banner(annual, "ANNUAL ATTRIBUTION / 2023–2024");
RiskStyle.Section(annual, "B7:N7", "START + TIME + MARKET + WINDOW + TRADES = END");
string[] yearHeaders = {
    "YEAR", "DAYS", "DV01 START", "TIME", "MARKET", "TRADES", "DV01 END",
    "VaR START", "TIME", "MARKET", "WINDOW", "TRADES", "VaR END"
};
for (int k = 0; k < yearHeaders.Length; k++)
    annual.Cell(8, k + 2).Value = yearHeaders[k];
RiskStyle.Header(annual, "B8:N8");
annual.Columns("B:B").Width = 12;
annual.Columns("C:C").Width = 10;
annual.Columns("D:N").Width = 18;
for (int k = 0; k < annualRows.Count; k++)
{
    int r = 9 + k;
    var a = annualRows[k];
    annual.Cell(r, 2).Value = a.Year;
    annual.Cell(r, 3).Value = a.ReportDays;
    double[] values = {
        a.DvStart, a.DvTime, a.DvMarket, a.DvTrades, a.DvEnd,
        a.VarStart, a.VarTime, a.VarMarket, a.VarWindow, a.VarTrades, a.VarEnd
    };
    for (int j = 0; j < values.Length; j++)
    {
        annual.Cell(r, 4 + j).Value = values[j];
        annual.Cell(r, 4 + j).Style.NumberFormat.Format = RiskStyle.Num;
    }
    RiskStyle.Stripe(annual, r, 2, 14, k);
}

RiskStyle.Section(annual, "B14:G14", "2024 KEY-RATE DV01 | DECOMPOSITION BY TENOR");
string[] ladderHeaders = { "TENOR", "START", "TIME", "MARKET", "TRADES", "END" };
for (int k = 0; k < ladderHeaders.Length; k++)
    annual.Cell(15, k + 2).Value = ladderHeaders[k];
RiskStyle.Header(annual, "B15:G15");
for (int k = 0; k < snapshot.Ladder2024.Count; k++)
{
    int r = 16 + k;
    var x = snapshot.Ladder2024[k];
    annual.Cell(r, 2).Value = x.Tenor;
    annual.Cell(r, 3).Value = x.Start;
    annual.Cell(r, 4).Value = x.Time;
    annual.Cell(r, 5).Value = x.Market;
    annual.Cell(r, 6).Value = x.Trades;
    annual.Cell(r, 7).Value = x.End;
    for (int col = 3; col <= 7; col++)
        annual.Cell(r, col).Style.NumberFormat.Format = RiskStyle.Num;
    RiskStyle.Stripe(annual, r, 2, 7, k);
}
annual.Cell("B28").Value = "TOTAL NET DV01";
for (int col = 3; col <= 7; col++)
{
    string l = ((char)('A' + col - 1)).ToString();
    annual.Cell(28, col).FormulaA1 = $"SUM({l}16:{l}26)";
    annual.Cell(28, col).Style.NumberFormat.Format = RiskStyle.Num;
}
RiskStyle.Total(annual, 28, 2, 7);
RiskStyle.Section(annual, "B31:N31", "2024 LADDER | VERTEX ATTRIBUTION");
if (!string.IsNullOrWhiteSpace(snapshot.LadderChart))
{
    using var pic = new MemoryStream(Convert.FromBase64String(snapshot.LadderChart));
    annual.AddPicture(pic, XLPictureFormat.Png)
        .MoveTo(annual.Cell("B32")).Scale(0.60);
}

// =============================================================
// 06 / RISK MOVERS & DYNAMIC VaR WATERFALL
// A floating-cell waterfall, not a raster chart: each block is an
// ordinary Excel formula, so selecting another date updates the bridge.
// The axis is auto-zoomed to the selected morning's VaR range; it is
// deliberately NOT a zero-based risk-limit chart.
// =============================================================
RiskStyle.Banner(movers, "RISK MOVERS / INTERACTIVE VaR WATERFALL");
RiskStyle.Section(movers, "B7:AM7", "DATE-DRIVEN VaR BRIDGE | SELECT A DATE ON 03 MORNING DRILLDOWN");
movers.Cell("B9").Value = "SELECTED DATE";
movers.Cell("E9").FormulaA1 = "'03 Morning Drilldown'!C9";
movers.Cell("H9").Value = "WALK CHECK";
movers.Cell("J9").FormulaA1 = "'03 Morning Drilldown'!F22";
movers.Cell("J9").Style.Font.Bold = true;

RiskStyle.Section(movers, "B12:AM12", "FLOATING VaR WATERFALL | RELATIVE SCALE, USD");
movers.Cell("B14").Value = "STATE / DRIVER";
movers.Cell("E14").Value = "CHANGE";
movers.Cell("G14").Value = "RUNNING VaR";
movers.Range("H14:AM14").Merge();
movers.Cell("H14").Value = "FLOATING CONTRIBUTION BAR →";
RiskStyle.Header(movers, "B14:AM14");
movers.Column("B").Width = 27;
movers.Column("C").Width = 4;
movers.Column("D").Width = 4;
movers.Column("E").Width = 17;
movers.Column("F").Width = 4;
movers.Column("G").Width = 21;
movers.Columns("H:AM").Width = 2.8;

string[] stateLabels = {
    "Opening VaR", "Time / Aging", "Market / Curve",
    "Historical Window Roll", "Trades / Amendments", "Closing VaR"
};
string[] runningCells = { "J13", "J14", "J15", "J16", "J17", "J18" };
string[] deltaCells = { "", "H14", "H15", "H16", "H17", "" };
string[] barColors = { "#14263D", "#B7791F", "#244767", "#64748B", "#078C86", "#14263D" };

for (int k = 0; k < stateLabels.Length; k++)
{
    int row = 15 + k;
    movers.Cell(row, 2).Value = stateLabels[k];
    movers.Cell(row, 7).FormulaA1 = $"'03 Morning Drilldown'!{runningCells[k]}";
    if (k == 0) movers.Cell(row, 5).Value = 0;
    else if (k == 5) movers.Cell(row, 5).FormulaA1 = "G20-G15";
    else movers.Cell(row, 5).FormulaA1 = $"'03 Morning Drilldown'!{deltaCells[k]}";
    movers.Cell(row, 5).Style.NumberFormat.Format = RiskStyle.Money;
    movers.Cell(row, 7).Style.NumberFormat.Format = RiskStyle.Money;
    movers.Row(row).Height = 26;
    RiskStyle.Stripe(movers, row, 2, 39, k);
    for (int col = 8; col <= 39; col++)
    {
        // Midpoint of each of 32 equal-width bins on the auto-zoomed VaR axis.
        string px = $"($AO$2+(COLUMN()-COLUMN($H${row})+0.5)*($AP$2-$AO$2)/32)";
        string formula;
        if (k == 0 || k == 5)
        {
            // Starting and closing VaR are point markers, not misleading 0-based bars.
            formula = $"IF(ABS({px}-$G${row})<=($AP$2-$AO$2)/64,\"◆\",\"\")";
        }
        else
        {
            // True floating bridge: color only the interval from prior to new VaR.
            formula = $"IF(AND({px}>=MIN($G${row-1},$G${row})," +
                      $"{px}<MAX($G${row-1},$G${row})),\"■\",\"\")";
        }
        var block = movers.Cell(row, col);
        block.FormulaA1 = formula;
        block.Style.Font.FontColor = XLColor.FromHtml(barColors[k]);
        block.Style.Font.FontSize = 12;
        block.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }
}
// Scale is based on this date's six states, with 10% dynamic margin.
// Hidden helper cells are also useful for inspecting/chart troubleshooting.
movers.Cell("AO2").FormulaA1 =
    "MIN(G15:G20)-MAX(1,MAX(G15:G20)-MIN(G15:G20))*0.1";
movers.Cell("AP2").FormulaA1 =
    "MAX(G15:G20)+MAX(1,MAX(G15:G20)-MIN(G15:G20))*0.1";
movers.Columns("AO:AP").Hide();
movers.Cell("B22").Value = "BRIDGE RECONCILIATION";
movers.Cell("G22").FormulaA1 = "G15+SUM(E16:E19)-G20";
movers.Cell("G22").Style.NumberFormat.Format = RiskStyle.Decimal;
RiskStyle.Info(movers, "B24:AM24",
    "Bar lengths are shown on a relative auto-zoomed VaR axis, not a zero-based scale. Step order matters.");
RiskStyle.Info(movers, "B25:AM25",
    "A large daily VaR move is an investigation candidate, not evidence of an erroneous trade or a breached limit.");

var top10 = dailyRows.Select((r, index) => new
    {
        ReportNumber = index + 1,
        Date = r.Date,
        Start = r.VarStart,
        End = r.VarEnd,
        Delta = r.VarEnd - r.VarStart,
        Time = r.VarTime,
        Market = r.VarMarket,
        Window = r.VarWindow,
        Trades = r.VarTrades
    })
    .OrderByDescending(r => Math.Abs(r.Delta))
    .ThenBy(r => r.Date)
    .Take(10)
    .ToList();

RiskStyle.Section(movers, "B28:AM28", "TOP 10 ABSOLUTE DAILY VaR CHANGES | ALL REPORT MORNINGS");
movers.Cell("B30").Value = "DATE";
movers.Cell("E30").Value = "REPORT #";
movers.Cell("G30").Value = "DELTA VaR";
movers.Range("H30:P30").Merge();
movers.Cell("H30").Value = "LARGEST DRIVER";
RiskStyle.Header(movers, "B30:P30");
for (int k = 0; k < top10.Count; k++)
{
    int row = 31 + k;
    var r = top10[k];
    movers.Cell(row, 2).Value = DateTime.ParseExact(r.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    movers.Cell(row, 2).Style.DateFormat.Format = "yyyy-mm-dd";
    movers.Cell(row, 5).Value = r.ReportNumber;
    movers.Cell(row, 7).Value = r.Delta;
    movers.Cell(row, 7).Style.NumberFormat.Format = RiskStyle.Money;
    var contributions = new (string Name, double Amount)[] {
        ("TIME",r.Time),("MARKET",r.Market),("WINDOW",r.Window),("TRADES",r.Trades)
    };
    movers.Range(row, 8, row, 16).Merge();
    movers.Cell(row, 8).Value = contributions
        .OrderByDescending(x => Math.Abs(x.Amount)).First().Name;
    RiskStyle.Stripe(movers, row, 2, 16, k);
}
movers.Range("B43:AM43").Merge();
movers.Cell("B43").Value =
    "Select one of these dates in 03 Morning Drilldown to inspect the dynamic waterfall.";
movers.Cell("B43").Style.Font.FontColor = XLColor.FromHtml("#64748B");

// Top daily movers also appear on Overview as a compact investigation queue.
RiskStyle.Section(overview, "B58:L58", "TOP 10 DAILY VaR MOVES | INVESTIGATION QUEUE, NOT LIMIT BREACHES");
overview.Cell("B60").Value = "DATE";
overview.Cell("E60").Value = "DELTA VaR";
overview.Cell("I60").Value = "LARGEST DRIVER";
RiskStyle.Header(overview, "B60:L60");
for (int k = 0; k < top10.Count; k++)
{
    int row = 61 + k;
    var r = top10[k];
    overview.Cell(row, 2).Value = DateTime.ParseExact(r.Date, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    overview.Cell(row, 2).Style.DateFormat.Format = "yyyy-mm-dd";
    overview.Cell(row, 5).Value = r.Delta;
    overview.Cell(row, 5).Style.NumberFormat.Format = RiskStyle.Money;
    var largest = new (string Name, double Amount)[] {
        ("TIME",r.Time),("MARKET",r.Market),("WINDOW",r.Window),("TRADES",r.Trades)
    };
    overview.Cell(row, 9).Value = largest
        .OrderByDescending(x => Math.Abs(x.Amount)).First().Name;
    RiskStyle.Stripe(overview, row, 2, 12, k);
}

// Cross-sheet, Excel-native reconciliations; no VBA or external refresh required.
RiskStyle.Section(overview, "B73:L73", "DAILY WALK RECONCILIATION | ARITHMETIC CONTROL");
overview.Cell("B75").Value = "DV01 DAILY CHECKS";
overview.Cell("E75").FormulaA1 =
    $"COUNTIF('02 Daily Walk'!R{firstDailyRow}:R{lastDailyRow},\"CHECK\")";
overview.Cell("H75").Value = "VaR DAILY CHECKS";
overview.Cell("K75").FormulaA1 =
    $"COUNTIF('02 Daily Walk'!S{firstDailyRow}:S{lastDailyRow},\"CHECK\")";
overview.Cell("B77").Value = "OVERALL CONTROL";
overview.Cell("E77").FormulaA1 =
    "IF(SUM(E75,K75)=0,\"PASS\",\"CHECK\")";
overview.Cell("E77").Style.Font.Bold = true;
RiskStyle.Info(overview, "B79:L79",
    "PASS tests only the algebra of the DV01/VaR decomposition; it does NOT validate source trades, curve marks or limits.");

// =============================================================
// 05 / METHODOLOGY & CONTROL NOTES
// =============================================================
RiskStyle.Banner(method, "RISK ATTRIBUTION / METHODOLOGY");
method.Column("B").Width = 27;
method.Column("C").Width = 28;
method.Column("D").Width = 92;
RiskStyle.Section(method, "B7:D7", "HOW THE FIVE VaR STATES ARE BUILT");
(string stage, string definition, string reason)[] explanations = {
    ("V0 / Yesterday", "Yesterday book, yesterday ladder/window", "Starting VaR as of the previous close."),
    ("V1 / Time", "Yesterday book, aged ladder, old window", "Isolates bond aging and time passage."),
    ("V2 / Market", "Yesterday book, current curve, old window", "Isolates curve repricing of the aged book."),
    ("V3 / Window", "Same ladder, shifted 500-day window", "Isolates historical window roll in the VaR distribution."),
    ("V4 / Trades", "Current book, current curve/window", "Isolates today's booked trades and amendments.")
};
for (int k = 0; k < explanations.Length; k++)
{
    int r = 9 + k;
    method.Cell(r, 2).Value = explanations[k].stage;
    method.Cell(r, 3).Value = explanations[k].definition;
    method.Cell(r, 4).Value = explanations[k].reason;
    method.Row(r).Height = 38;
    method.Range(r, 2, r, 4).Style.Alignment.WrapText = true;
    RiskStyle.Stripe(method, r, 2, 4, k);
}
RiskStyle.Section(method, "B16:D16", "DEFINITIONS, VALIDATION AND LIMITATIONS");
(string key, string note)[] notes = {
    ("Net DV01 / USD per bp", "Sum of 11 key-rate DV01s; gain for a parallel 1bp rate fall."),
    ("Time / Aging", "Same book, but one day closer to its future coupon and maturity cash flows."),
    ("Market / Curve", "Change in the current bond sensitivity as the yield curve moves."),
    ("Historical Window", "Changing the 500-day historical move distribution may change VaR even without new trades."),
    ("Trades / Amendments", "Change in booked face amounts since the previous close."),
    ("Exact waterfall", "Daily step differences telescope to close minus start; the order chosen affects attribution."),
    ("Daily audit checks", "02 Daily Walk: I and Q are numeric walk residuals; R/S show PASS or CHECK at a 0.01 tolerance."),
    ("Morning selector", "03 Morning Drilldown: blue C9 is an ISO report date selected from a validated dropdown."),
    ("Dynamic waterfall", "06 Movers & Waterfall visualizes the selected date's six VaR states on a relative auto-zoomed scale."),
    ("Risk movers", "Top 10 dates are ranked by absolute daily VaR change; this is investigative triage, NOT a risk-limit breach."),
    ("As-of convention", "A report for close of date t is reviewed on the next reporting morning."),
    ("Model limitation", "Historical VaR uses first-order key-rate DV01; it excludes full convexity repricing."),
    ("Risk status", "Simulated educational project; not live production limits, P&L or trade execution."),
    ("Data source", "risk_walk.ipynb exports task2_snapshot.json from simulated rates_desk_feed.csv and trade blotter.")
};
for (int k = 0; k < notes.Length; k++)
{
    int r = 18 + k;
    method.Cell(r, 2).Value = notes[k].key;
    method.Range(r, 3, r, 4).Merge();
    method.Cell(r, 3).Value = notes[k].note;
    method.Row(r).Height = 34;
    method.Range(r, 2, r, 4).Style.Alignment.WrapText = true;
    RiskStyle.Stripe(method, r, 2, 4, k);
}

// =============================================================
// EXPORT / EXISTING FOLDER CONVENTIONS
// =============================================================
Directory.CreateDirectory(outputDirectory);
string destination = Path.Combine(outputDirectory,
    $"RatesDesk_RiskAttribution_Task2_V2_{snapshot.AsOf}.xlsx");
wb.SaveAs(destination);
Console.WriteLine("SUCCESS / TASK 2 V2 INTERACTIVE ATTRIBUTION");
Console.WriteLine($"Excel: {destination}");
Console.WriteLine($"Report days: {dailyRows.Count}");
Console.WriteLine("Interactive date dropdown, VaR bridge, top 10 movers, reconciliation flags ready.");
Console.WriteLine($"2024 Closing DV01: {year24.DvEnd:+#,##0;-#,##0;0} USD/bp");
Console.WriteLine($"2024 Closing VaR: ${year24.VarEnd:N0}");
Console.WriteLine($"Max previous-day ladder gap: {snapshot.ChainMaxGap:G4}");

// =============================================================
// JSON CONTRACT / EXACT FIELD NAMES USED BY THE PYTHON EXPORTER
// =============================================================
public sealed class Task2Snapshot
{
    [JsonPropertyName("asOf")] public string AsOf { get; set; } = "";
    [JsonPropertyName("chainMaxGap")] public double ChainMaxGap { get; set; }
    [JsonPropertyName("cumulativeChart")] public string CumulativeChart { get; set; } = "";
    [JsonPropertyName("ladderChart")] public string LadderChart { get; set; } = "";
    [JsonPropertyName("daily")] public List<DailyWalk> Daily { get; set; } = new();
    [JsonPropertyName("years")] public List<YearWalk> Years { get; set; } = new();
    [JsonPropertyName("ladder2024")] public List<LadderStep> Ladder2024 { get; set; } = new();
}
public sealed class DailyWalk
{
    [JsonPropertyName("date")] public string Date { get; set; } = "";
    [JsonPropertyName("dvStart")] public double DvStart { get; set; }
    [JsonPropertyName("dvTime")] public double DvTime { get; set; }
    [JsonPropertyName("dvMarket")] public double DvMarket { get; set; }
    [JsonPropertyName("dvTrades")] public double DvTrades { get; set; }
    [JsonPropertyName("dvEnd")] public double DvEnd { get; set; }
    [JsonPropertyName("varStart")] public double VarStart { get; set; }
    [JsonPropertyName("varTime")] public double VarTime { get; set; }
    [JsonPropertyName("varMarket")] public double VarMarket { get; set; }
    [JsonPropertyName("varWindow")] public double VarWindow { get; set; }
    [JsonPropertyName("varTrades")] public double VarTrades { get; set; }
    [JsonPropertyName("varEnd")] public double VarEnd { get; set; }
}
public sealed class YearWalk
{
    [JsonPropertyName("year")] public int Year { get; set; }
    [JsonPropertyName("reportDays")] public int ReportDays { get; set; }
    [JsonPropertyName("dvStart")] public double DvStart { get; set; }
    [JsonPropertyName("dvTime")] public double DvTime { get; set; }
    [JsonPropertyName("dvMarket")] public double DvMarket { get; set; }
    [JsonPropertyName("dvTrades")] public double DvTrades { get; set; }
    [JsonPropertyName("dvEnd")] public double DvEnd { get; set; }
    [JsonPropertyName("varStart")] public double VarStart { get; set; }
    [JsonPropertyName("varTime")] public double VarTime { get; set; }
    [JsonPropertyName("varMarket")] public double VarMarket { get; set; }
    [JsonPropertyName("varWindow")] public double VarWindow { get; set; }
    [JsonPropertyName("varTrades")] public double VarTrades { get; set; }
    [JsonPropertyName("varEnd")] public double VarEnd { get; set; }
}
public sealed class LadderStep
{
    [JsonPropertyName("tenor")] public string Tenor { get; set; } = "";
    [JsonPropertyName("start")] public double Start { get; set; }
    [JsonPropertyName("time")] public double Time { get; set; }
    [JsonPropertyName("market")] public double Market { get; set; }
    [JsonPropertyName("trades")] public double Trades { get; set; }
    [JsonPropertyName("end")] public double End { get; set; }
}

// =============================================================
// CONSISTENT TRADING-DESK STYLE (MIRRORS TASK 1 V3)
// =============================================================
public static class RiskStyle
{
    public const string Num = "+#,##0;[Red](#,##0);-";
    public const string Money = "$#,##0;[Red]($#,##0);-";
    public const string Decimal = "+#,##0.0000;[Red](#,##0.0000);-";
    static readonly XLColor Navy = XLColor.FromHtml("#14263D");
    static readonly XLColor Blue = XLColor.FromHtml("#244767");
    static readonly XLColor Pale = XLColor.FromHtml("#EDF3F8");
    static readonly XLColor Muted = XLColor.FromHtml("#64748B");
    static readonly XLColor White = XLColor.White;
    public static void Setup(IXLWorksheet ws, string date, string[] tabs)
    {
        ws.ShowGridLines = false;
        ws.Style.Font.FontName = "Aptos";
        ws.Style.Font.FontSize = 10;
        ws.Style.Font.FontColor = Navy;
        ws.Column(1).Width = 3;
        ws.Columns("B:Q").Width = 14;
        ws.Row(2).Height = 27;
        ws.Row(3).Height = 25;
        ws.Cell("B4").Value = $"AS OF {date} | FIXED INCOME | SIMULATED / EOD";
        ws.Cell("B4").Style.Font.FontColor = Muted;
        string[] links = { "B5", "D5", "F5", "H5", "J5", "L5" };
        string[] labels = { "OVERVIEW", "DAILY WALK", "DRILLDOWN", "ANNUAL", "METHOD", "MOVERS" };
        for (int k = 0; k < tabs.Length; k++)
        {
            var cell = ws.Cell(links[k]);
            cell.FormulaA1 = $"HYPERLINK(\"#'{tabs[k]}'!B2\",\"{labels[k]}\")";
            cell.Style.Fill.BackgroundColor = Pale;
            cell.Style.Font.FontColor = Blue;
            cell.Style.Font.Bold = true;
        }
    }
    public static void Banner(IXLWorksheet ws, string text)
    {
        var box = ws.Range("B2:L3");
        box.Merge();
        box.FirstCell().Value = text;
        box.Style.Fill.BackgroundColor = Navy;
        box.Style.Font.FontColor = White;
        box.Style.Font.Bold = true;
        box.Style.Font.FontSize = 18;
        box.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Section(IXLWorksheet ws, string address, string text)
    {
        var box = ws.Range(address);
        box.Merge();
        box.FirstCell().Value = text;
        box.Style.Fill.BackgroundColor = Blue;
        box.Style.Font.FontColor = White;
        box.Style.Font.Bold = true;
        box.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Header(IXLWorksheet ws, string address)
    {
        var b = ws.Range(address);
        b.Style.Fill.BackgroundColor = Blue;
        b.Style.Font.FontColor = White;
        b.Style.Font.Bold = true;
        b.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Kpi(IXLWorksheet ws, string label, string valueArea,
                           string name, double value, bool money = false)
    {
        var a = ws.Range(label); a.Merge();
        a.FirstCell().Value = name;
        a.Style.Fill.BackgroundColor = Pale;
        a.Style.Font.FontColor = Muted;
        a.Style.Font.Bold = true;
        var b = ws.Range(valueArea); b.Merge();
        b.FirstCell().Value = value;
        b.Style.Fill.BackgroundColor = Pale;
        b.Style.Font.FontColor = Navy;
        b.Style.Font.FontSize = 22;
        b.Style.Font.Bold = true;
        b.Style.NumberFormat.Format = money ? Money : Num;
        b.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        b.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    }
    public static void Info(IXLWorksheet ws, string address, string text)
    {
        var b = ws.Range(address); b.Merge();
        b.FirstCell().Value = text;
        b.Style.Font.FontColor = Muted;
        b.Style.Font.FontSize = 9;
    }
    public static void Stripe(IXLWorksheet ws, int row, int first, int last, int index)
    {
        if (index % 2 == 0)
            ws.Range(row, first, row, last).Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F4F7FA");
    }
    public static void Total(IXLWorksheet ws, int row, int first, int last)
    {
        var b = ws.Range(row, first, row, last);
        b.Style.Fill.BackgroundColor = Pale;
        b.Style.Font.FontColor = Navy;
        b.Style.Font.Bold = true;
        b.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        ws.Row(row).Height = 27;
    }
}
