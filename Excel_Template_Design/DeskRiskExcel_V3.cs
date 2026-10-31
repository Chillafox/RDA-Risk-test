
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

// ====================================================
// V3 CONFIGURATION - reads the existing Task 1 Jupyter snapshot.
// The source JSON schema has not changed: you can reuse your V2 exporter.
// ====================================================

string folder =
    @"C:\Users\franklinpeng\Desktop\RDA Risk test\Excel_Template_Output";

string input = Path.Combine(folder, "task1_snapshot.json");

if (!File.Exists(input))
    throw new FileNotFoundException(
        "Run the Jupyter snapshot export first.", input);

var data = JsonSerializer.Deserialize<DeskSnapshot>(
    File.ReadAllText(input)
) ?? throw new Exception("Invalid snapshot JSON.");

if (data.Positions.Count == 0 ||
    data.Ladder.Count != 11 ||
    data.Scenarios.Count != 500)
{
    throw new Exception(
        "Expected positions, 11 key rates, and 500 scenarios.");
}

var ranked = data.Scenarios
    .OrderBy(x => x.Pnl)
    .ToList();

double var99 = -ranked[4].Pnl;
double netDv01 = data.Ladder.Sum(x => x.Dv01);
double grossDv01 = data.Positions.Sum(x => Math.Abs(x.Dv01));
double netValue = data.Positions.Sum(x => x.Value);

if (Math.Abs(var99 - data.Var99) > 1.0)
    throw new Exception("VaR reconciliation failed.");

using var wb = new XLWorkbook();

var overview = wb.Worksheets.Add("01 Overview");
var positions = wb.Worksheets.Add("02 Positions");
var keyrates = wb.Worksheets.Add("03 Key Rates");
var scenarios = wb.Worksheets.Add("04 VaR Scenarios");
var stress = wb.Worksheets.Add("05 Stress Lab");

foreach (var ws in wb.Worksheets)
    Theme.Setup(ws, data.AsOf);

// ====================================================
// 01 OVERVIEW - top metrics, largest risk positions and audit status.
// ====================================================

Theme.Banner(
    overview,
    "RATES DESK  /  DAILY RISK MONITOR"
);

Theme.Section(
    overview, "B6:L6",
    "DESK PERFORMANCE & MARKET RISK"
);

Theme.Kpi(
    overview,
    "B7:F7", "B8:F10",
    "REPORTED DAILY P&L",
    data.ReportedPnl,
    true
);

Theme.Kpi(
    overview,
    "H7:L7", "H8:L10",
    "1-DAY 99% HISTORICAL VaR",
    var99,
    true
);

Theme.Info(
    overview, "B11:F11",
    "P&L SOURCE: REPORTED / UNRECONCILED"
);

Theme.Info(
    overview, "H11:L11",
    $"DoD VaR: {Theme.Signed(var99 - data.PreviousVar, true)}"
);

Theme.Kpi(
    overview,
    "B13:F13", "B14:F16",
    "NET KEY-RATE DV01  /  USD/BP",
    netDv01,
    false
);

Theme.Kpi(
    overview,
    "H13:L13", "H14:L16",
    "GROSS POSITION DV01  /  USD/BP",
    grossDv01,
    false
);

Theme.Info(
    overview, "B17:F17",
    $"DoD DV01: {Theme.Signed(netDv01 - data.PreviousNetDv01, false)}"
);

Theme.Info(
    overview, "H17:L17",
    $"NET BOOK VALUE: ${netValue:N0}"
);

Theme.Section(
    overview, "B19:L19",
    "KEY-RATE DV01  /  2023 VS 2024"
);

// Actual risk exposure chart from the Python engine
using var curveStream = new MemoryStream(
    Convert.FromBase64String(data.CurveImage)
);

overview.AddPicture(curveStream, XLPictureFormat.Png)
    .MoveTo(overview.Cell("B20"))
    .Scale(0.57);

Theme.Section(
    overview, "B36:L36",
    "TOP FIVE POSITION RISK CONTRIBUTORS"
);

overview.Cell("B38").Value = "BOND";
overview.Cell("D38").Value = "FACE (MM)";
overview.Cell("F38").Value = "DV01 (USD/BP)";
overview.Cell("I38").Value = "DIRECTION";

Theme.Header(overview, "B38:L38");

var topPositions = data.Positions
    .OrderByDescending(x => Math.Abs(x.Dv01))
    .Take(5)
    .ToList();

for (int n = 0; n < topPositions.Count; n++)
{
    int r = 39 + n;
    var p = topPositions[n];

    overview.Cell(r, 2).Value = p.Bond;
    overview.Cell(r, 4).Value = p.FaceM;
    overview.Cell(r, 6).Value = p.Dv01;
    overview.Cell(r, 9).Value =
        p.Dv01 >= 0 ? "LONG RATES" : "SHORT RATES";

    overview.Cell(r, 4).Style.NumberFormat.Format =
        Theme.NumberFormat;

    overview.Cell(r, 6).Style.NumberFormat.Format =
        Theme.NumberFormat;

    Theme.Stripe(overview, r, 2, 12, n);

    Theme.RiskColor(overview.Cell(r, 6), p.Dv01);
}

Theme.Section(
    overview, "B46:L46",
    "RISK CONTROL & REPORT STATUS"
);

overview.Cell("B48").Value = "Valuation Model";
overview.Cell("F48").Value = "ZERO CURVE / CONTINUOUS";

overview.Cell("B49").Value = "Historical Scenarios";
overview.Cell("F49").Value = data.Scenarios.Count;

overview.Cell("B50").Value = "VaR Confidence";
overview.Cell("F50").Value = "99% / ONE DAY";

overview.Cell("B51").Value = "Risk Limits";
overview.Cell("F51").Value = "NOT CONFIGURED";

overview.Cell("B52").Value = "P&L Reconciliation";
overview.Cell("F52").Value = "TASK 3 REQUIRED";
overview.Cell("B53").Value = "Interactive Stress";
overview.Cell("F53").FormulaA1 =
    "HYPERLINK(\"#'05 Stress Lab'!B2\",\"OPEN STRESS LAB\")";

Theme.Info(
    overview, "B56:L56",
    "SIMULATED DESK  |  STATIC EOD REPORT  |  NOT LIVE MARKET RISK"
);

// ====================================================
// 02 POSITIONS
// ====================================================

Theme.Banner(
    positions,
    "BOND POSITIONS  /  CLOSING BOOK"
);

string[] positionHeaders = {
    "BOND",
    "FACE (USD MM)",
    "MATURITY (YRS)",
    "VALUE (USD)",
    "DV01 (USD/BP)",
    "SIDE",
    "GROSS DV01 SHARE"
};

for (int j = 0; j < positionHeaders.Length; j++)
    positions.Cell(7, j + 2).Value = positionHeaders[j];

Theme.Header(positions, "B7:H7");

positions.Columns("B:H").Width = 21;
positions.Column("E").Width = 25;
positions.Column("F").Width = 23;

for (int n = 0; n < data.Positions.Count; n++)
{
    int r = 8 + n;
    var p = data.Positions[n];

    positions.Cell(r, 2).Value = p.Bond;
    positions.Cell(r, 3).Value = p.FaceM;
    positions.Cell(r, 4).Value = p.Years;
    positions.Cell(r, 5).Value = p.Value;
    positions.Cell(r, 6).Value = p.Dv01;

    positions.Cell(r, 7).Value =
        p.FaceM > 0 ? "LONG" : "SHORT";

    positions.Cell(r, 8).Value =
        grossDv01 == 0
            ? 0
            : Math.Abs(p.Dv01) / grossDv01;

    positions.Cell(r, 3).Style.NumberFormat.Format =
        Theme.NumberFormat;

    positions.Cell(r, 4).Style.NumberFormat.Format =
        "0.00";

    positions.Cell(r, 5).Style.NumberFormat.Format =
        Theme.MoneyFormat;

    positions.Cell(r, 6).Style.NumberFormat.Format =
        Theme.NumberFormat;

    positions.Cell(r, 8).Style.NumberFormat.Format =
        "0.0%";

    Theme.Stripe(positions, r, 2, 8, n);
    Theme.RiskColor(positions.Cell(r, 6), p.Dv01);

    positions.Row(r).Height = 23;
}

int positionTotal = 8 + data.Positions.Count;

positions.Cell(positionTotal, 2).Value = "BOOK TOTAL";
positions.Cell(positionTotal, 5).Value = netValue;
positions.Cell(positionTotal, 6).Value = netDv01;
positions.Cell(positionTotal, 8).Value = 1.0;

positions.Cell(positionTotal, 5).Style.NumberFormat.Format =
    Theme.MoneyFormat;

positions.Cell(positionTotal, 6).Style.NumberFormat.Format =
    Theme.NumberFormat;

positions.Cell(positionTotal, 8).Style.NumberFormat.Format =
    "0.0%";

Theme.Total(positions, positionTotal, 2, 8);

positions.Range(7, 2, positionTotal - 1, 8)
    .SetAutoFilter();

positions.SheetView.FreezeRows(7);

// ====================================================
// 03 KEY RATES
// ====================================================

Theme.Banner(
    keyrates,
    "KEY-RATE DV01  /  TERM STRUCTURE"
);

string[] keyHeaders = {
    "TENOR",
    "2023 DV01",
    "2024 DV01",
    "CHANGE",
    "ABS RISK SHARE",
    "DIRECTION"
};

for (int j = 0; j < keyHeaders.Length; j++)
    keyrates.Cell(7, j + 2).Value = keyHeaders[j];

Theme.Header(keyrates, "B7:G7");

keyrates.Columns("B:G").Width = 23;

double grossKeyRate = data.Ladder.Sum(
    x => Math.Abs(x.Dv01)
);

for (int n = 0; n < data.Ladder.Count; n++)
{
    int r = 8 + n;
    var item = data.Ladder[n];

    keyrates.Cell(r, 2).Value = item.Tenor;
    keyrates.Cell(r, 3).Value = item.YearAgo;
    keyrates.Cell(r, 4).Value = item.Dv01;
    keyrates.Cell(r, 5).Value =
        item.Dv01 - item.YearAgo;

    keyrates.Cell(r, 6).Value =
        grossKeyRate == 0
            ? 0
            : Math.Abs(item.Dv01) / grossKeyRate;

    keyrates.Cell(r, 7).Value =
        item.Dv01 >= 0
            ? "LONG RATES"
            : "SHORT RATES";

    for (int col = 3; col <= 5; col++)
        keyrates.Cell(r, col).Style.NumberFormat.Format =
            Theme.NumberFormat;

    keyrates.Cell(r, 6).Style.NumberFormat.Format =
        "0.0%";

    Theme.Stripe(keyrates, r, 2, 7, n);
    Theme.RiskColor(keyrates.Cell(r, 4), item.Dv01);

    keyrates.Row(r).Height = 24;
}

keyrates.Cell("B20").Value = "NET DV01";
keyrates.Cell("C20").Value =
    data.Ladder.Sum(x => x.YearAgo);
keyrates.Cell("D20").Value = netDv01;
keyrates.Cell("E20").Value =
    netDv01 - data.Ladder.Sum(x => x.YearAgo);

Theme.Total(keyrates, 20, 2, 7);

foreach (string address in new[] {"C20", "D20", "E20"})
    keyrates.Cell(address).Style.NumberFormat.Format =
        Theme.NumberFormat;

Theme.Section(
    keyrates, "B23:G23",
    "RISK INTERPRETATION"
);

keyrates.Cell("B25").Value =
    "Gross key-rate exposure is the sum of absolute vertex DV01s.";

keyrates.Cell("B26").Value =
    "Offsetting tenors can conceal curve risk when net DV01 is small.";

keyrates.Cell("B27").Value =
    "Positive DV01 = gain from a 1bp fall at that vertex.";

// ====================================================
// 04 HISTORICAL VaR SCENARIOS
// ====================================================

Theme.Banner(
    scenarios,
    "HISTORICAL SIMULATION  /  500 CURVE SHOCKS"
);

scenarios.Columns("B:B").Width = 20;
scenarios.Columns("C:C").Width = 25;
scenarios.Columns("D:E").Width = 16;
scenarios.Columns("F:F").Width = 24;

string[] scenarioHeaders = {
    "SHOCK DATE",
    "P&L (USD)",
    "RANK",
    "TAIL STATUS",
    "LOSS MAGNITUDE"
};

for (int j = 0; j < scenarioHeaders.Length; j++)
    scenarios.Cell(7, j + 2).Value = scenarioHeaders[j];

Theme.Header(scenarios, "B7:F7");

for (int n = 0; n < ranked.Count; n++)
{
    int r = 8 + n;
    var s = ranked[n];

    scenarios.Cell(r, 2).Value =
        DateTime.ParseExact(
            s.Date,
            "yyyy-MM-dd",
            CultureInfo.InvariantCulture
        );

    scenarios.Cell(r, 3).Value = s.Pnl;
    scenarios.Cell(r, 4).Value = n + 1;
    scenarios.Cell(r, 5).Value =
        n < 5 ? "VaR TAIL" : "";

    scenarios.Cell(r, 6).Value =
        Math.Max(0, -s.Pnl);

    scenarios.Cell(r, 2).Style.DateFormat.Format =
        "yyyy-mm-dd";

    scenarios.Cell(r, 3).Style.NumberFormat.Format =
        Theme.MoneyFormat;

    scenarios.Cell(r, 6).Style.NumberFormat.Format =
        Theme.MoneyFormat;

    Theme.Stripe(scenarios, r, 2, 6, n);

    if (n < 5)
    {
        var tailRow = scenarios.Range(r, 2, r, 6);
        tailRow.Style.Fill.BackgroundColor =
            XLColor.FromHtml("#FCE7E9");

        tailRow.Style.Font.FontColor =
            XLColor.FromHtml("#A82F45");

        tailRow.Style.Font.Bold = true;
    }
}

scenarios.Range(7, 2, 7 + ranked.Count, 6)
    .SetAutoFilter();

scenarios.SheetView.FreezeRows(7);

Theme.Section(
    scenarios, "H7:P7",
    "FIVE LARGEST HISTORICAL LOSSES"
);

using var lossStream = new MemoryStream(
    Convert.FromBase64String(data.LossImage)
);

scenarios.AddPicture(lossStream, XLPictureFormat.Png)
    .MoveTo(scenarios.Cell("H9"))
    .Scale(0.55);

Theme.Section(
    scenarios, "H28:P28",
    "HISTORICAL VaR METHODOLOGY"
);

scenarios.Cell("H30").Value = "Lookback Window";
scenarios.Cell("M30").Value = "500 DAYS";

scenarios.Cell("H31").Value = "Confidence";
scenarios.Cell("M31").Value = "99%";

scenarios.Cell("H32").Value = "VaR Order Statistic";
scenarios.Cell("M32").Value = "5TH WORST";

scenarios.Cell("H33").Value = "Current VaR (USD)";
scenarios.Cell("M33").Value = var99;

scenarios.Cell("M33").Style.NumberFormat.Format =
    Theme.MoneyFormat;

scenarios.Cell("H35").Value =
    "P&L = - historical curve move (bp) x current DV01 ladder.";


// ====================================================
// 05 STRESS LAB - NATIVE EXCEL WHAT-IF (NO VBA)
// ====================================================
// A rates trader changes the scenario dropdown (E9) or the blue custom
// shock cells (D12:D22). Excel updates the applied curve shocks and the
// first-order P&L estimate. The current risk ladder is from Task 1.
// No external price feed, trade execution, recalculated VaR or convexity.

Theme.Banner(stress, "STRESS LAB  /  YIELD CURVE WHAT-IF");
Theme.Section(stress, "B7:N7", "INTERACTIVE SCENARIOS  |  EDIT BLUE CELLS  |  USD PER BASIS POINT");
stress.Columns("B:B").Width = 16;
stress.Columns("C:C").Width = 23;
stress.Columns("D:E").Width = 20;
stress.Column("F").Width = 24;
stress.Column("G").Width = 20;
stress.Column("H").Width = 3;
stress.Columns("I:N").Width = 21;

stress.Cell("B9").Value = "ACTIVE SCENARIO";
stress.Cell("B9").Style.Font.Bold = true;
stress.Cell("E9").Value = "Parallel +10bp";
stress.Cell("E9").Style.Fill.BackgroundColor = XLColor.FromHtml("#DBEAFE");
stress.Cell("E9").Style.Font.FontColor = XLColor.FromHtml("#1455A3");
stress.Cell("E9").Style.Font.Bold = true;
stress.Cell("I9").Value = "REPORT DATE";
stress.Cell("K9").Value = data.AsOf;

string[] shockHeaders = { "TENOR", "DV01 (USD/BP)", "CUSTOM (BP)", "APPLIED (BP)", "P&L IMPACT (USD)", "RISK WEIGHT" };
for (int k = 0; k < shockHeaders.Length; k++)
    stress.Cell(11, k + 2).Value = shockHeaders[k];
Theme.Header(stress, "B11:G11");

// The scenario library is kept as transparent worksheet inputs, not hidden
// program logic. Preset values can be reviewed or edited below the main table.
Theme.Section(stress, "I27:N27", "SCENARIO LIBRARY  /  EDITABLE PRESET SHOCKS (BP)");
string[] presetNames = {
    "Parallel +10bp", "Parallel -10bp", "Bear Steepener",
    "Bull Flattener", "Front-End +25bp", "Custom"
};
for (int k = 0; k < presetNames.Length; k++)
    stress.Cell(29, k + 9).Value = presetNames[k];
Theme.Header(stress, "I29:N29");

// 11 tenor-specific shocks. "Custom" uses the D12:D22 cells instead.
double[][] presetShock = {
    new double[] {10,10,10,10,10,10,10,10,10,10,10},
    new double[] {-10,-10,-10,-10,-10,-10,-10,-10,-10,-10,-10},
    new double[] {0,0,0,0,2,5,10,15,20,20,20},
    new double[] {-20,-20,-18,-15,-12,-10,-8,-6,-5,-5,-5},
    new double[] {25,25,25,25,12.5,0,0,0,0,0,0},
    new double[] {0,0,0,0,0,0,0,0,0,0,0}
};

// E9 dropdown refers to a valid on-sheet range. This avoids the Excel
// cross-sheet list validation limitation and requires no macros.
stress.Cell("E9").CreateDataValidation().List(stress.Range("I29:N29"));
stress.Range("D12:D22").CreateDataValidation().Decimal.Between(-300, 300);

for (int k = 0; k < data.Ladder.Count; k++)
{
    int row = 12 + k;
    int libraryRow = 30 + k;
    var item = data.Ladder[k];

    // SOURCE: Task 1 Key-Rate DV01 in USD for a 1bp yield FALL.
    stress.Cell(row, 2).Value = item.Tenor;
    stress.Cell(row, 3).Value = item.Dv01;
    stress.Cell(row, 3).Style.NumberFormat.Format = Theme.NumberFormat;
    Theme.RiskColor(stress.Cell(row, 3), item.Dv01);

    // INPUT: trader's custom per-tenor shock, in basis points.
    stress.Cell(row, 4).Value = 0.0;
    stress.Cell(row, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#DBEAFE");
    stress.Cell(row, 4).Style.Font.FontColor = XLColor.FromHtml("#1455A3");
    stress.Cell(row, 4).Style.NumberFormat.Format = "+0.0;-0.0;0.0";

    // CALCULATION: choose the selected preset, or a custom input.
    // INDEX/MATCH references the preset table; no C# or Python rerun needed.
    stress.Cell(row, 5).FormulaA1 =
        $"IF($E$9=\"Custom\",D{row},INDEX($I${libraryRow}:$N${libraryRow},1,MATCH($E$9,$I$29:$N$29,0)))";
    stress.Cell(row, 5).Style.NumberFormat.Format = "+0.0;-0.0;0.0";

    // P&L sign convention: rising yields -> negative P&L for positive DV01.
    stress.Cell(row, 6).FormulaA1 = $"-C{row}*E{row}";
    stress.Cell(row, 6).Style.NumberFormat.Format = Theme.MoneyFormat;
    stress.Cell(row, 7).Value =
        grossKeyRate == 0 ? 0 : Math.Abs(item.Dv01) / grossKeyRate;
    stress.Cell(row, 7).Style.NumberFormat.Format = "0.0%";
    Theme.Stripe(stress, row, 2, 7, k);
    stress.Row(row).Height = 24;

    // Preset source curves (editable by an analyst for testing).
    for (int p = 0; p < presetNames.Length; p++)
    {
        stress.Cell(libraryRow, 9 + p).Value = presetShock[p][k];
        stress.Cell(libraryRow, 9 + p).Style.NumberFormat.Format = "+0.0;-0.0;0.0";
        stress.Cell(libraryRow, 9 + p).Style.Fill.BackgroundColor =
            XLColor.FromHtml("#F0F5FA");
    }
}

stress.Cell("B24").Value = "TOTAL STRESS P&L";
stress.Cell("C24").Value = netDv01;
stress.Cell("C24").Style.NumberFormat.Format = Theme.NumberFormat;
stress.Cell("F24").FormulaA1 = "SUM(F12:F22)";
stress.Cell("F24").Style.NumberFormat.Format = Theme.MoneyFormat;
Theme.Total(stress, 24, 2, 7);

// Scenario selection information stays distinct from computed historical VaR.
stress.Cell("B26").Value = "HOW TO USE";
stress.Cell("B27").Value = "1. Select a scenario in E9.";
stress.Cell("B28").Value = "2. For custom shocks, select Custom and edit the blue D12:D22 cells.";
stress.Cell("B29").Value = "3. Read each tenor's estimated P&L in column F and the total in F24.";
stress.Cell("B31").Value = "MODEL BOUNDARY";
stress.Cell("B32").Value = "Linear DV01 only; no convexity or full scenario repricing.";
stress.Cell("B33").Value = "The stress result does not change the reported daily P&L or historical VaR.";
stress.Cell("B34").Value = "This workbook is a simulated EOD training tool, not a live trading application.";

// ====================================================
// 06 METHODOLOGY follows (existing V2 content retained)
// ====================================================

// ====================================================
// 06 METHODOLOGY & DATA DEFINITIONS
// ====================================================

var guide = wb.Worksheets.Add("06 Methodology");

guide.ShowGridLines = false;
guide.Style.Font.FontName = "Aptos";
guide.Style.Font.FontSize = 11;

guide.Column("A").Width = 3;
guide.Column("B").Width = 25;
guide.Column("C").Width = 34;
guide.Column("D").Width = 85;

var guideTitle = guide.Range("B2:D3");
guideTitle.Merge();
guideTitle.FirstCell().Value =
    "RATES DESK | METHODOLOGY & DATA GUIDE";
guideTitle.Style.Fill.BackgroundColor =
    XLColor.FromHtml("#14263D");
guideTitle.Style.Font.FontColor = XLColor.White;
guideTitle.Style.Font.FontSize = 18;
guideTitle.Style.Font.Bold = true;

guide.Cell("B5").Value =
    "Purpose: Explain how the daily risk report is calculated.";
guide.Cell("B5").Style.Font.FontColor =
    XLColor.FromHtml("#64748B");

// Processing workflow
guide.Cell("B7").Value = "STEP";
guide.Cell("C7").Value = "INPUT";
guide.Cell("D7").Value = "PROCESS / BUSINESS PURPOSE";

var steps = new (string Step, string Input, string Purpose)[]
{
    (
        "01 | Market Data",
        "Daily zero yield curves",
        "Load 11 yield-curve vertices from 3M to 30Y. Rates are continuously compounded and expressed in percent."
    ),
    (
        "02 | Position Book",
        "Trade blotter",
        "Apply the last booked amendment for each ticket and aggregate closing face amounts by bond."
    ),
    (
        "03 | Cash Flows",
        "Coupon and maturity",
        "Generate future semiannual coupon payments and principal repayment at maturity."
    ),
    (
        "04 | Bond Valuation",
        "Cash flows + zero curve",
        "Interpolate zero rates and discount each future payment to calculate the bond's present value."
    ),
    (
        "05 | Key-Rate DV01",
        "Discounted cash flows",
        "Calculate the USD sensitivity to a 1bp yield decrease at each of the 11 curve vertices."
    ),
    (
        "06 | Risk Aggregation",
        "Bond-level DV01",
        "Aggregate position sensitivities to construct the portfolio's key-rate ladder."
    ),
    (
        "07 | Model Validation",
        "Analytical DV01",
        "Compare analytical sensitivities with central-difference repricing using +1bp and -1bp shocks."
    ),
    (
        "08 | Historical VaR",
        "500 daily curve moves",
        "Estimate scenario P&L using the current ladder and select the fifth-worst observation for 99% VaR."
    ),
    (
        "09 | Report Output",
        "Positions and risk results",
        "Generate the dashboard, position details, key-rate ladder and historical scenario tables."
    )
};

for (int n = 0; n < steps.Length; n++)
{
    int row = 8 + n;
    var step = steps[n];

    guide.Cell(row, 2).Value = step.Step;
    guide.Cell(row, 3).Value = step.Input;
    guide.Cell(row, 4).Value = step.Purpose;

    guide.Row(row).Height = 40;
    guide.Range(row, 2, row, 4)
        .Style.Alignment.Vertical =
        XLAlignmentVerticalValues.Center;

    guide.Range(row, 2, row, 4)
        .Style.Alignment.WrapText = true;

    if (n % 2 == 0)
    {
        guide.Range(row, 2, row, 4)
            .Style.Fill.BackgroundColor =
            XLColor.FromHtml("#F1F5F9");
    }
}

// Metric definitions
guide.Cell("B20").Value = "RISK METRIC";
guide.Cell("C20").Value = "UNIT";
guide.Cell("D20").Value = "INTERPRETATION";

var definitions = new (string Metric, string Unit, string Meaning)[]
{
    (
        "Net DV01",
        "USD / 1bp",
        "Approximate portfolio gain when all curve vertices fall by 1bp."
    ),
    (
        "Gross Position DV01",
        "USD / 1bp",
        "Sum of absolute bond-level DV01 values before offsets."
    ),
    (
        "Key-Rate DV01",
        "USD / 1bp",
        "Sensitivity to a 1bp fall at one specific yield-curve vertex."
    ),
    (
        "99% Historical VaR",
        "USD",
        "Fifth-worst estimated portfolio P&L across 500 historical scenarios, expressed as a positive loss measure."
    ),
    (
        "Reported P&L",
        "USD",
        "P&L received from the external reporting feed; not reconciled by Task 1."
    ),
    (
        "Portfolio Value",
        "USD",
        "Net discounted value of all live long and short bond positions."
    )
};

for (int n = 0; n < definitions.Length; n++)
{
    int row = 21 + n;
    var item = definitions[n];

    guide.Cell(row, 2).Value = item.Metric;
    guide.Cell(row, 3).Value = item.Unit;
    guide.Cell(row, 4).Value = item.Meaning;

    guide.Row(row).Height = 35;
    guide.Range(row, 2, row, 4)
        .Style.Alignment.WrapText = true;
}

// Assumptions & limitations
guide.Cell("B29").Value = "MODEL ASSUMPTIONS";
guide.Cell("B30").Value = "Day Count";
guide.Cell("D30").Value = "ACT/365";

guide.Cell("B31").Value = "Discounting";
guide.Cell("D31").Value =
    "Continuously compounded zero rates.";

guide.Cell("B32").Value = "Interpolation";
guide.Cell("D32").Value =
    "Linear between vertices; flat extrapolation outside the curve.";

guide.Cell("B33").Value = "VaR Method";
guide.Cell("D33").Value =
    "Linear key-rate DV01 approximation, not full scenario repricing.";

guide.Cell("B34").Value = "Model Scope";
guide.Cell("D34").Value =
    "Excludes explicit credit spread, funding, liquidity and transaction-cost models.";

guide.Cell("B35").Value = "Data Classification";
guide.Cell("D35").Value =
    "Simulated training dataset. Not live production risk.";

foreach (int headerRow in new[] {7, 20})
{
    var range = guide.Range(headerRow, 2, headerRow, 4);
    range.Style.Fill.BackgroundColor =
        XLColor.FromHtml("#244767");
    range.Style.Font.FontColor = XLColor.White;
    range.Style.Font.Bold = true;
}

guide.Range("B29:D29").Style.Fill.BackgroundColor =
    XLColor.FromHtml("#244767");
guide.Range("B29:D29").Style.Font.FontColor = XLColor.White;
guide.Range("B29:D29").Style.Font.Bold = true;

guide.Range("D30:D35").Style.Alignment.WrapText = true;
guide.SheetView.FreezeRows(7);

// V3: separate historical risk from an editable first-order stress estimate.
guide.Cell("B38").Value = "INTERACTIVE STRESS LAB";
guide.Range("B38:D38").Style.Fill.BackgroundColor = XLColor.FromHtml("#244767");
guide.Range("B38:D38").Style.Font.FontColor = XLColor.White;
guide.Range("B38:D38").Style.Font.Bold = true;
guide.Cell("B39").Value = "Input";
guide.Cell("D39").Value = "Curve shock at each tenor, measured in basis points.";
guide.Cell("B40").Value = "Scenario P&L";
guide.Cell("D40").Value = "Negative sum of (Key-Rate DV01 x yield shock).";
guide.Cell("B41").Value = "Implementation";
guide.Cell("D41").Value = "Excel dropdown + editable custom cells + INDEX/MATCH formulas; no VBA.";
guide.Cell("B42").Value = "Limitation";
guide.Cell("D42").Value = "Linear approximation only. No convexity, changing positions, new market prices or VaR re-estimation.";
guide.Range("D39:D42").Style.Alignment.WrapText = true;
guide.Rows(39, 42).Height = 32;

// ====================================================
// SAVE WORKBOOK
// ====================================================

Directory.CreateDirectory(folder);

string destination = Path.Combine(
    folder,
    $"RatesDesk_RiskWorkstation_V3_{data.AsOf}.xlsx"
);

wb.SaveAs(destination);

Console.WriteLine("SUCCESS: V3 Risk Workstation (macro-free interactive Stress Lab)");
Console.WriteLine($"Excel: {destination}");
Console.WriteLine($"Net DV01: {netDv01:+#,##0;-#,##0;0}");
Console.WriteLine($"Gross DV01: {grossDv01:N0}");
Console.WriteLine($"Historical VaR: ${var99:N0}");
Console.WriteLine($"Positions: {data.Positions.Count}");
Console.WriteLine($"Scenarios: {data.Scenarios.Count}");
Console.WriteLine("Stress Lab: 6 selectable scenarios, 11 editable tenor shocks");


// ====================================================
// CLASSES
// ====================================================

public sealed class DeskSnapshot
{
    [JsonPropertyName("asOf")]
    public string AsOf { get; set; } = "";

    [JsonPropertyName("previousAsOf")]
    public string PreviousAsOf { get; set; } = "";

    [JsonPropertyName("reportedPnl")]
    public double? ReportedPnl { get; set; }

    [JsonPropertyName("previousVar")]
    public double PreviousVar { get; set; }

    [JsonPropertyName("previousNetDv01")]
    public double PreviousNetDv01 { get; set; }

    [JsonPropertyName("var99")]
    public double Var99 { get; set; }

    [JsonPropertyName("curveImage")]
    public string CurveImage { get; set; } = "";

    [JsonPropertyName("lossImage")]
    public string LossImage { get; set; } = "";

    [JsonPropertyName("positions")]
    public List<BondPosition> Positions { get; set; } = new();

    [JsonPropertyName("ladder")]
    public List<KeyRate> Ladder { get; set; } = new();

    [JsonPropertyName("scenarios")]
    public List<Scenario> Scenarios { get; set; } = new();
}

public sealed class BondPosition
{
    [JsonPropertyName("bond")]
    public string Bond { get; set; } = "";

    [JsonPropertyName("face_m")]
    public double FaceM { get; set; }

    [JsonPropertyName("years")]
    public double Years { get; set; }

    [JsonPropertyName("value")]
    public double Value { get; set; }

    [JsonPropertyName("dv01")]
    public double Dv01 { get; set; }
}

public sealed class KeyRate
{
    [JsonPropertyName("tenor")]
    public string Tenor { get; set; } = "";

    [JsonPropertyName("dv01")]
    public double Dv01 { get; set; }

    [JsonPropertyName("yearAgo")]
    public double YearAgo { get; set; }
}

public sealed class Scenario
{
    [JsonPropertyName("date")]
    public string Date { get; set; } = "";

    [JsonPropertyName("pnl")]
    public double Pnl { get; set; }
}

// ====================================================
// STYLING / LAYOUT UTILITIES
// ====================================================

public static class Theme
{
    public const string NumberFormat =
        "+#,##0;[Red](#,##0);-";

    public const string MoneyFormat =
        "$#,##0;[Red]($#,##0);-";

    static readonly XLColor Navy =
        XLColor.FromHtml("#14263D");

    static readonly XLColor Blue =
        XLColor.FromHtml("#244767");

    static readonly XLColor Teal =
        XLColor.FromHtml("#078C86");

    static readonly XLColor Red =
        XLColor.FromHtml("#CD5260");

    static readonly XLColor Pale =
        XLColor.FromHtml("#EDF3F8");

    static readonly XLColor Muted =
        XLColor.FromHtml("#63768A");

    static readonly XLColor White =
        XLColor.White;

    public static void Setup(IXLWorksheet ws, string date)
    {
        ws.ShowGridLines = false;

        ws.Style.Font.FontName = "Aptos";
        ws.Style.Font.FontSize = 10;
        ws.Style.Font.FontColor = Navy;

        ws.Column(1).Width = 3;
        ws.Columns("B:P").Width = 12;

        ws.Row(2).Height = 27;
        ws.Row(3).Height = 25;
        ws.Row(5).Height = 28;

        ws.Cell("B4").Value =
            $"AS OF {date}  |  FIXED INCOME  |  SIMULATED DESK";

        ws.Cell("B4").Style.Font.FontColor = Muted;
        ws.Cell("B4").Style.Font.FontSize = 9;

        Navigation(ws);
    }

    public static void Navigation(IXLWorksheet ws)
    {
        // Navigation looks like buttons, but is entirely macro-free.
        // HYPERLINK() takes the user to the selected worksheet.
        string[] names = {
            "01 Overview", "02 Positions", "03 Key Rates",
            "04 VaR Scenarios", "05 Stress Lab", "06 Methodology"
        };
        string[] labels = {
            "OVERVIEW", "POSITIONS", "KEY RATES",
            "VaR SCENARIOS", "STRESS LAB", "METHODOLOGY"
        };
        for (int k = 0; k < names.Length; k++)
        {
            int left = 2 + 2 * k;
            var region = ws.Range(5, left, 5, left + 1);
            region.Merge();
            var cell = region.FirstCell();
            cell.FormulaA1 =
                $"HYPERLINK(\"#'{names[k]}'!B2\",\"{labels[k]}\")";
            region.Style.Font.Bold = true;
            region.Style.Font.FontColor = Blue;
            region.Style.Fill.BackgroundColor = Pale;
            region.Style.Alignment.Horizontal =
                XLAlignmentHorizontalValues.Center;
            region.Style.Alignment.Vertical =
                XLAlignmentVerticalValues.Center;
        }
    }

    public static void Banner(IXLWorksheet ws, string text)
    {
        var r = ws.Range("B2:M3");
        r.Merge();

        r.FirstCell().Value = text;
        r.Style.Fill.BackgroundColor = Navy;
        r.Style.Font.FontColor = White;
        r.Style.Font.FontSize = 19;
        r.Style.Font.Bold = true;
        r.Style.Alignment.Vertical =
            XLAlignmentVerticalValues.Center;
    }

    public static void Section(
        IXLWorksheet ws,
        string address,
        string text)
    {
        var r = ws.Range(address);
        r.Merge();

        r.FirstCell().Value = text;
        r.Style.Fill.BackgroundColor = Blue;
        r.Style.Font.FontColor = White;
        r.Style.Font.Bold = true;
        r.Style.Font.FontSize = 10;
        r.Style.Alignment.Vertical =
            XLAlignmentVerticalValues.Center;
    }

    public static void Header(
        IXLWorksheet ws,
        string address)
    {
        var r = ws.Range(address);

        r.Style.Fill.BackgroundColor = Blue;
        r.Style.Font.FontColor = White;
        r.Style.Font.Bold = true;
        r.Style.Alignment.Vertical =
            XLAlignmentVerticalValues.Center;
    }

    public static void Kpi(
        IXLWorksheet ws,
        string labelAddress,
        string valueAddress,
        string name,
        double? value,
        bool money)
    {
        var label = ws.Range(labelAddress);
        label.Merge();

        label.FirstCell().Value = name;
        label.Style.Fill.BackgroundColor = Pale;
        label.Style.Font.FontColor = Muted;
        label.Style.Font.Bold = true;

        var body = ws.Range(valueAddress);
        body.Merge();

        if (value.HasValue)
        {
            body.FirstCell().Value = value.Value;
            body.Style.NumberFormat.Format =
                money ? MoneyFormat : NumberFormat;
        }
        else
        {
            body.FirstCell().Value = "NOT AVAILABLE";
        }

        body.Style.Fill.BackgroundColor = Pale;
        body.Style.Font.FontColor = Navy;
        body.Style.Font.FontSize = 23;
        body.Style.Font.Bold = true;

        body.Style.Alignment.Horizontal =
            XLAlignmentHorizontalValues.Right;

        body.Style.Alignment.Vertical =
            XLAlignmentVerticalValues.Center;
    }

    public static void Info(
        IXLWorksheet ws,
        string address,
        string value)
    {
        var r = ws.Range(address);
        r.Merge();

        r.FirstCell().Value = value;
        r.Style.Font.FontColor = Muted;
        r.Style.Font.FontSize = 9;
    }

    public static void Stripe(
        IXLWorksheet ws,
        int row,
        int firstCol,
        int lastCol,
        int index)
    {
        if (index % 2 == 0)
        {
            ws.Range(row, firstCol, row, lastCol)
                .Style.Fill.BackgroundColor =
                XLColor.FromHtml("#F4F7FA");
        }
    }

    public static void RiskColor(
        IXLCell cell,
        double value)
    {
        cell.Style.Font.FontColor =
            value >= 0 ? Teal : Red;

        cell.Style.Font.Bold = true;
    }

    public static void Total(
        IXLWorksheet ws,
        int row,
        int firstCol,
        int lastCol)
    {
        var r = ws.Range(row, firstCol, row, lastCol);

        r.Style.Fill.BackgroundColor = Pale;
        r.Style.Font.FontColor = Navy;
        r.Style.Font.Bold = true;

        ws.Row(row).Height = 28;
    }

    public static string Signed(
        double value,
        bool money)
    {
        if (money)
            return value.ToString(
                "+$#,##0;-$#,##0;$0",
                CultureInfo.InvariantCulture
            );

        return value.ToString(
            "+#,##0;-#,##0;0",
            CultureInfo.InvariantCulture
        );
    }
}
