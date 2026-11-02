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
using ClosedXML.Excel.Drawings;

// =====================================================================
// FIXED INCOME RISK WORKSTATION | V4 COMMAND CENTER
// .NET 10 file-based C# | ClosedXML 0.105.1 | no VBA
// Input: Excel_Template_Output/workstation_snapshot.json
// Prepared + validated by Final.ipynb, not manually edited.
// =====================================================================

string projectRoot = Environment.GetEnvironmentVariable("RAD_RISK_ROOT")
                     ?? Directory.GetCurrentDirectory();
string outDir = Path.Combine(projectRoot, "Excel_Template_Output");
string source = Path.Combine(outDir, "workstation_snapshot.json");
if (!File.Exists(source))
    throw new FileNotFoundException("Run Final.ipynb to validate and export the workstation snapshot.", source);

using var doc = JsonDocument.Parse(File.ReadAllText(source));
JsonElement d = doc.RootElement;
if (J.N(d, "schemaVersion") != 1)
    throw new InvalidOperationException("Incompatible workstation contract schema version.");
string asOf = J.S(d, "asOf");
JsonElement risk = d.GetProperty("risk");
JsonElement forensic = d.GetProperty("forensic");
JsonElement policy = d.GetProperty("policy");
JsonElement pnlCalibration = d.GetProperty("pnlCalibration");
JsonElement validation = d.GetProperty("validation");
var dates = d.GetProperty("dateRows").EnumerateArray().ToList();
var positions = risk.GetProperty("positions").EnumerateArray().ToList();
var ladder = risk.GetProperty("ladder").EnumerateArray().ToList();
var scenarios = risk.GetProperty("scenarios").EnumerateArray().ToList();
var policies = policy.GetProperty("policies").EnumerateArray().ToList();
var cases = forensic.GetProperty("casesEvaluationOnly").EnumerateArray().ToList();
var qa = validation.GetProperty("checks").EnumerateArray().ToList();

if (asOf != "2024-12-31" || dates.Count != 260 || ladder.Count != 11 ||
    positions.Count == 0 || scenarios.Count != 500 || policies.Count != 5 ||
    cases.Count != 6 || qa.Count == 0 || qa.Any(c => J.S(c, "status") != "PASS"))
    throw new InvalidOperationException("Final.ipynb QA must pass, with expected 2024 dataset coverage.");
var isoDates = dates.Select(x => J.S(x,"date")).ToList();
if (isoDates.Distinct().Count() != dates.Count ||
    !isoDates.SequenceEqual(isoDates.OrderBy(x => x, StringComparer.Ordinal)) ||
    !isoDates.All(x => x.StartsWith("2024-", StringComparison.Ordinal)))
    throw new InvalidOperationException("Report dates must be unique, sorted ISO 2024 strings.");

// The final snapshot contains independently checked 2024 risk and P&L rows.
foreach (JsonElement r in dates)
{
    if (Math.Abs(J.N(r,"dvEnd") - (J.N(r,"dvStart")+J.N(r,"dvTime")+
        J.N(r,"dvMarket")+J.N(r,"dvTrades"))) > 0.01)
        throw new InvalidOperationException("DV01 walk residual on " + J.S(r,"date"));
    if (Math.Abs(J.N(r,"varEnd") - (J.N(r,"varStart")+J.N(r,"varTime")+
        J.N(r,"varMarket")+J.N(r,"varWindow")+J.N(r,"varTrades"))) > 0.01)
        throw new InvalidOperationException("VaR walk residual on " + J.S(r,"date"));
    if (Math.Abs(J.N(r,"reportedPnl") - J.N(r,"predictedPnl") - J.N(r,"unexplained")) > 0.01)
        throw new InvalidOperationException("P&L reconciliation residual on " + J.S(r,"date"));
    if (r.GetProperty("tenorPnl").GetArrayLength() != 11)
        throw new InvalidOperationException("Missing P&L contributions for all eleven tenors.");
}
var finalRow = dates[^1];
if (Math.Abs(J.N(finalRow,"varEnd")-J.N(risk,"var99")) > 1 ||
    Math.Abs(J.N(finalRow,"dvEnd")-J.N(risk,"netDv01")) > 0.01 ||
    J.S(policy,"winner") != "SCALED+FEED+LATE")
    throw new InvalidOperationException("Final 2024 book risk or offline sign-off is inconsistent.");

Directory.CreateDirectory(outDir);
string destination = Path.Combine(outDir, $"RatesDesk_RiskWorkstation_FINAL_{asOf}.xlsx");
using var wb = new XLWorkbook();
var command = wb.Worksheets.Add("01 Command Center");
var portfolio = wb.Worksheets.Add("02 Portfolio Risk");
var attribution = wb.Worksheets.Add("03 Risk Attribution");
var pnl = wb.Worksheets.Add("04 PnL Explain");
var incident = wb.Worksheets.Add("05 Incident Monitor");
var governance = wb.Worksheets.Add("06 Policy Governance");
var stress = wb.Worksheets.Add("07 Stress Lab");
var guide = wb.Worksheets.Add("08 Methodology");
var hidden = wb.Worksheets.Add("_DATA");
var trendData = wb.Worksheets.Add("_TREND");

string[] sheetNames = {
    "01 Command Center", "02 Portfolio Risk", "03 Risk Attribution",
    "04 PnL Explain", "05 Incident Monitor", "06 Policy Governance",
    "07 Stress Lab", "08 Methodology"
};
foreach (var ws in wb.Worksheets)
{
    Ui.Base(ws, asOf);
    if (ws != hidden && ws != trendData) Ui.Nav(ws, sheetNames);
}

// ---------------------------------------------------------------------
// HIDDEN DATA / one row per 2024 reporting morning
// Column A is intentionally a text ISO date for stable MATCH formulas.
// ---------------------------------------------------------------------
string[] internalFields = {
    "date", "dvStart", "dvTime", "dvMarket", "dvTrades", "dvEnd",
    "varStart", "varTime", "varMarket", "varWindow", "varTrades", "varEnd",
    "reportedPnl", "carry", "marketPnl", "predictedPnl", "unexplained",
    "noiseSd", "zScore", "PNL_FIXED", "PNL_SCALED", "STALE", "KINK", "LATE",
    "signalCount", "staleEvidence", "kinkEvidence", "lateTickets", "reserved"
};
for (int j=0; j<internalFields.Length; j++) hidden.Cell(1,j+1).Value=internalFields[j];
for (int k=0; k<11; k++) hidden.Cell(1,30+k).Value="tenorPnl_"+J.S(ladder[k],"tenor");

int rowNumber=2;
foreach (var r in dates)
{
    hidden.Cell(rowNumber,1).Value=J.S(r,"date");
    string[] numericFields = {
        "dvStart","dvTime","dvMarket","dvTrades","dvEnd",
        "varStart","varTime","varMarket","varWindow","varTrades","varEnd",
        "reportedPnl","carry","marketPnl","predictedPnl","unexplained","noiseSd","zScore"
    };
    for (int c=0;c<numericFields.Length;c++)
        hidden.Cell(rowNumber,c+2).Value=J.N(r,numericFields[c]);
    JsonElement flags=r.GetProperty("screens");
    string[] flagNames={"PNL_FIXED","PNL_SCALED","STALE","KINK","LATE"};
    int signals=0;
    for (int c=0;c<flagNames.Length;c++)
    {
        bool triggered=J.B(flags,flagNames[c]);
        if (triggered) signals++;
        hidden.Cell(rowNumber,20+c).Value=triggered ? "FLAG" : "CLEAR";
    }
    hidden.Cell(rowNumber,25).Value=signals;
    hidden.Cell(rowNumber,26).Value=J.S(r,"staleEvidence");
    hidden.Cell(rowNumber,27).Value=J.S(r,"kinkEvidence");
    hidden.Cell(rowNumber,28).Value=J.S(r,"lateTickets");
    var perTenor = r.GetProperty("tenorPnl").EnumerateArray().ToList();
    for (int c=0;c<11;c++) hidden.Cell(rowNumber,30+c).Value=perTenor[c].GetDouble();
    rowNumber++;
}
int last = rowNumber-1;

// ---------------------------------------------------------------------
// V4 TREND DATA | Excel-calculated glyph charts (no static screenshots)
// Both timelines end on the selected Command date. 32 uniformly spaced
// samples preserve the shape of the selected 20D/60D/YTD window.
// The full original 260-day series is retained in _DATA.
// ---------------------------------------------------------------------
trendData.Cell("B2").FormulaA1 =
    $"IFERROR(MATCH('01 Command Center'!$C$6,'_DATA'!$A$2:$A${last},0),{dates.Count})";
trendData.Cell("B3").FormulaA1 =
    "IF('01 Command Center'!$C$20=\"20D\",20,"+
    "IF('01 Command Center'!$C$20=\"60D\",60,260))";
trendData.Cell("B4").FormulaA1 = "MAX(1,B2-B3+1)";
trendData.Cell("B5").FormulaA1 =
    $"MIN(INDEX('_DATA'!$L$2:$L${last},$B$4):INDEX('_DATA'!$L$2:$L${last},$B$2))";
trendData.Cell("B6").FormulaA1 =
    $"MAX(INDEX('_DATA'!$L$2:$L${last},$B$4):INDEX('_DATA'!$L$2:$L${last},$B$2))";
trendData.Cell("B7").FormulaA1 =
    $"MIN(INDEX('_DATA'!$F$2:$F${last},$B$4):INDEX('_DATA'!$F$2:$F${last},$B$2))";
trendData.Cell("B8").FormulaA1 =
    $"MAX(INDEX('_DATA'!$F$2:$F${last},$B$4):INDEX('_DATA'!$F$2:$F${last},$B$2))";
const int sparkPoints = 32;
for (int k = 0; k < sparkPoints; k++)
{
    string col = Ui.ExcelColumn(3+k);
    trendData.Cell($"{col}11").FormulaA1 =
        $"MAX(1,ROUND($B$4+($B$2-$B$4)*{k}/{sparkPoints-1},0))";
    trendData.Cell($"{col}12").FormulaA1 =
        $"INDEX('_DATA'!$L$2:$L${last},{col}11)";
    trendData.Cell($"{col}13").FormulaA1 =
        $"UNICHAR(9601+ROUND(7*({col}12-$B$5)/MAX(0.000001,$B$6-$B$5),0))";
    trendData.Cell($"{col}14").FormulaA1 =
        $"INDEX('_DATA'!$F$2:$F${last},{col}11)";
    trendData.Cell($"{col}15").FormulaA1 =
        $"UNICHAR(9601+ROUND(7*({col}14-$B$7)/MAX(0.000001,$B$8-$B$7),0))";
}


// DATE UX V2: the Command date is the global default. Each daily analysis
// worksheet has a LOCAL override in F9, with its effective date in L9.
// Excel cannot bidirectionally update the same input cell on several sheets
// without macros/add-ins. This one-way-inheritance model is deliberate.
string ActiveDateRef(IXLWorksheet? ws = null) =>
    ws != null && (ws == attribution || ws == pnl || ws == incident)
        ? "$L$9"
        : "'01 Command Center'!$C$6";
string DataLookup(string col, IXLWorksheet? ws = null) =>
    $"IFERROR(INDEX('_DATA'!${col}$2:${col}${last},"+
    $"MATCH({ActiveDateRef(ws)},'_DATA'!$A$2:$A${last},0)),\"N/A\")";
void SetLookup(IXLWorksheet ws,string target,string col,string fmt="")
{
    var cell = ws.Cell(target);
    cell.FormulaA1=DataLookup(col,ws);
    if (!string.IsNullOrEmpty(fmt)) cell.Style.NumberFormat.Format=fmt;
}
void AddLocalDateSelector(IXLWorksheet ws)
{
    // Excel Data Validation is local to each sheet. A separate hidden local
    // list avoids cross-sheet list-reference errors in different Excel clients.
    const string followGlobal = "FOLLOW GLOBAL";
    ws.Cell("Z8").Value=followGlobal;
    for (int i=0;i<isoDates.Count;i++) ws.Cell(9+i,26).Value=isoDates[i];
    ws.Column("Z").Hide();

    ws.Cell("B9").Value="LOCAL DATE";
    ws.Cell("F9").Value=followGlobal;
    ws.Cell("F9").Style.NumberFormat.Format="@"; // keep ISO dates as text, matching _DATA
    ws.Cell("F9").CreateDataValidation().List(ws.Range($"Z8:Z{8+isoDates.Count}"));
    ws.Cell("F9").Style.Fill.BackgroundColor=Ui.Editable;
    ws.Cell("F9").Style.Font.FontColor=Ui.Blue;
    ws.Cell("F9").Style.Font.Bold=true;
    ws.Cell("F9").Style.Alignment.Horizontal=XLAlignmentHorizontalValues.Center;
    ws.Column("F").Width=22;

    ws.Cell("J9").Value="ACTIVE DATE";
    ws.Cell("L9").FormulaA1=
        "IF($F$9=\"FOLLOW GLOBAL\",'01 Command Center'!$C$6,$F$9)";
    ws.Cell("L9").Style.Fill.BackgroundColor=Ui.Pale;
    ws.Cell("L9").Style.Font.FontColor=Ui.Navy;
    ws.Cell("L9").Style.Font.Bold=true;
    ws.Column("L").Width=19;
    ws.Cell("N9").FormulaA1=
        "IF($F$9=\"FOLLOW GLOBAL\",\"GLOBAL\",\"LOCAL\")";
    ws.Cell("N9").Style.Font.FontColor=Ui.Teal;
    ws.Cell("N9").Style.Font.Bold=true;

    ws.Cell("B10").Value=
        "F9: choose a report date here, or FOLLOW GLOBAL to inherit Command Center C6.";
    ws.Cell("B10").Style.Font.FontColor=Ui.Muted;
    ws.Cell("B10").Style.Font.FontSize=9;
}

// ---------------------------------------------------------------------
// 01 COMMAND CENTER | Trading desk command center V4
// 2024 VaR/DV01 trend glyphs are native recalculating Excel formulas,
// NOT fixed PNG images. Closing 11-vertex Key-Rate chart is static 2024 EOD.
// ---------------------------------------------------------------------
Ui.Banner(command,"RATES DESK  /  RISK COMMAND CENTER");
command.Cell("B6").Value="REPORT DATE";
command.Cell("C6").Value=asOf;
command.Cell("C6").Style.NumberFormat.Format="@";
command.Cell("C6").Style.Fill.BackgroundColor=Ui.Editable;
command.Cell("C6").Style.Font.FontColor=Ui.Blue;
command.Cell("C6").Style.Font.Bold=true;
command.Cell("C6").Style.Alignment.Horizontal=XLAlignmentHorizontalValues.Center;
command.Column("C").Width=20;
command.Cell("F6").Value="2024 EOD  |  GLOBAL REPORT DATE";
command.Cell("F6").Style.Font.FontColor=Ui.Muted;
command.Cell("B8").Value="Select a date here; Analysis pages 03-05 may override locally using their F9 dropdown.";
command.Cell("B8").Style.Font.FontColor=Ui.Muted;
for (int n=0;n<isoDates.Count;n++) command.Cell(10+n,26).Value=isoDates[n];
command.Cell("C6").CreateDataValidation().List(command.Range($"Z10:Z{9+isoDates.Count}"));
command.Column("Z").Hide();

Ui.Section(command,"B10:P10","SELECTED DATE  |  RISK & P&L SNAPSHOT");
Ui.Kpi(command,"B11:D11","B12:D14","NET DV01  /  USD PER BP",DataLookup("F"),Ui.Number);
Ui.Kpi(command,"F11:H11","F12:H14","1-DAY 99% HISTORICAL VaR",DataLookup("L"),Ui.Money);
Ui.Kpi(command,"J11:L11","J12:L14","REPORTED DAILY P&L",DataLookup("M"),Ui.Money);
Ui.Kpi(command,"N11:P11","N12:P14","UNEXPLAINED P&L",DataLookup("Q"),Ui.Money);

// Reuse the underlying dated rows, not yesterday's calendar date.
// For the first report day, DoD change is shown as N/A.
string PrevDeltaFormula(string col)
{
    string rank=$"MATCH($C$6,'_DATA'!$A$2:$A${last},0)";
    string range=$"'_DATA'!${col}$2:${col}${last}";
    return $"IFERROR(IF({rank}=1,\"N/A\",INDEX({range},{rank})-INDEX({range},{rank}-1)),\"N/A\")";
}
command.Cell("B15").Value="DoD DV01 Δ";
command.Cell("C15").FormulaA1=PrevDeltaFormula("F");
command.Cell("C15").Style.NumberFormat.Format=Ui.Number;
command.Cell("F15").Value="DoD VaR Δ";
command.Cell("G15").FormulaA1=PrevDeltaFormula("L");
command.Cell("G15").Style.NumberFormat.Format=Ui.Money;
command.Cell("J15").Value="DoD P&L Δ";
command.Cell("K15").FormulaA1=PrevDeltaFormula("M");
command.Cell("K15").Style.NumberFormat.Format=Ui.Money;
command.Cell("N15").Value="Z-SCORE";
SetLookup(command,"O15","S","+0.00;-0.00;0.00");
command.Cell("O15").Style.Font.FontColor=Ui.Red;
command.Cell("O15").Style.Font.Bold=true;
foreach (string address in new[]{"B15","F15","J15","N15"})
    command.Cell(address).Style.Font.FontColor=Ui.Muted;

// Chart range selector. Every visible sparkline recalculates with report date
// and lookback; the source remains the entire 2024 report-day series.
Ui.Section(command,"B18:P18","RISK TREND MONITOR  |  SELECTED PERIOD VS STATIC CLOSING BOOK");
command.Cell("B20").Value="LOOKBACK";
command.Cell("C20").Value="60D";
command.Cell("Z400").Value="20D";
command.Cell("Z401").Value="60D";
command.Cell("Z402").Value="YTD";
command.Cell("C20").CreateDataValidation().List(command.Range("Z400:Z402"));
command.Cell("C20").Style.Fill.BackgroundColor=Ui.Editable;
command.Cell("C20").Style.Font.FontColor=Ui.Blue;
command.Cell("C20").Style.Font.Bold=true;
command.Cell("E20").Value="20D / 60D / YTD end on report date";
command.Cell("E20").Style.Font.FontColor=Ui.Muted;
Ui.Section(command,"B22:I22","HISTORICAL VaR  |  SELECTED WINDOW");
Ui.Section(command,"J22:P22","KEY-RATE DV01  |  STATIC 2024-12-31 BOOK");
// Visually useful formula-based sparklines; do not represent them as Excel
// chart objects or pretend the closing-book bars change with report date.
var varGraphic=command.Range("B24:I26");varGraphic.Merge();
varGraphic.FirstCell().FormulaA1="TEXTJOIN(\"\",TRUE,'_TREND'!C13:AH13)";
varGraphic.Style.Font.FontName="Segoe UI Symbol";
varGraphic.Style.Font.FontSize=24;
varGraphic.Style.Font.FontColor=Ui.Teal;
varGraphic.Style.Alignment.Horizontal=XLAlignmentHorizontalValues.Center;
varGraphic.Style.Alignment.Vertical=XLAlignmentVerticalValues.Center;
varGraphic.Style.Fill.BackgroundColor=Ui.Pale;
command.Cell("B27").Value="WINDOW LOW / HIGH";
command.Cell("D27").FormulaA1="'_TREND'!B5";
command.Cell("F27").FormulaA1="'_TREND'!B6";
command.Cell("D27").Style.NumberFormat.Format="\"USD \"0.00,,\"M\"";
command.Cell("F27").Style.NumberFormat.Format="\"USD \"0.00,,\"M\"";
command.Cell("B27").Style.Font.FontColor=Ui.Muted;
Ui.Section(command,"B29:I29","NET DV01  |  SELECTED WINDOW");
var dvGraphic=command.Range("B31:I33");dvGraphic.Merge();
dvGraphic.FirstCell().FormulaA1="TEXTJOIN(\"\",TRUE,'_TREND'!C15:AH15)";
dvGraphic.Style.Font.FontName="Segoe UI Symbol";
dvGraphic.Style.Font.FontSize=24;
dvGraphic.Style.Font.FontColor=Ui.Blue;
dvGraphic.Style.Alignment.Horizontal=XLAlignmentHorizontalValues.Center;
dvGraphic.Style.Alignment.Vertical=XLAlignmentVerticalValues.Center;
dvGraphic.Style.Fill.BackgroundColor=Ui.Pale;
command.Cell("B34").Value="WINDOW LOW / HIGH";
command.Cell("D34").FormulaA1="'_TREND'!B7";
command.Cell("F34").FormulaA1="'_TREND'!B8";
command.Cell("D34").Style.NumberFormat.Format=Ui.Number;
command.Cell("F34").Style.NumberFormat.Format=Ui.Number;
command.Cell("B34").Style.Font.FontColor=Ui.Muted;
command.Cell("B36").Value="Trend glyphs sample 32 points from the chosen period; exact daily values remain in _DATA.";
command.Cell("B36").Style.Font.FontColor=Ui.Muted;
command.Cell("B36").Style.Font.FontSize=9;

string curveBase64Command=J.S(risk,"curveImage");
if (!string.IsNullOrWhiteSpace(curveBase64Command))
{
    using var curveStreamCommand=new MemoryStream(Convert.FromBase64String(curveBase64Command));
    command.AddPicture(curveStreamCommand,XLPictureFormat.Png)
        .MoveTo(command.Cell("J24")).Scale(0.47);
}
command.Cell("J36").Value="2024-12-31 FIXED BOOK (NOT DATE-LINKED)";
command.Cell("J36").Style.Font.FontColor=Ui.Muted;

// Daily VaR walk makes the move actionable rather than simply displaying its size.
Ui.Section(command,"B39:P39","DAILY RISK DRIVERS  |  EXCEPTION REVIEW");
Ui.Section(command,"B41:I41","VaR CHANGE ATTRIBUTION / USD");
string[] walkTitles={"Opening VaR","Time / Aging","Market / Curve","Window Roll","Trades / Amendments","Closing VaR"};
string[] walkFields={"G","H","I","J","K","L"};
for(int i=0;i<walkTitles.Length;i++)
{
    int rr=42+i;
    command.Cell(rr,2).Value=walkTitles[i];
    SetLookup(command,$"H{rr}",walkFields[i],Ui.Money);
    Ui.Stripe(command,rr,2,9,i);
}
command.Cell("B49").Value="TOTAL VaR CHANGE";
command.Cell("H49").FormulaA1="H47-H42";
command.Cell("H49").Style.NumberFormat.Format=Ui.Money;
command.Cell("H49").Style.Font.Bold=true;
command.Cell("B50").FormulaA1="HYPERLINK(\"#'03 Risk Attribution'!B11\",\"INVESTIGATE RISK DRIVERS >\")";
command.Cell("B50").Style.Font.FontColor=Ui.Blue;

Ui.Section(command,"J41:P41","EXCEPTION QUEUE | NOT CONFIRMED INCIDENTS");
command.Cell("J42").Value="Z-score";
SetLookup(command,"N42","S","+0.00;-0.00;0.00");
command.Cell("J43").Value="Risk-scaled P&L screen";
SetLookup(command,"N43","U");
command.Cell("J44").Value="Other triggered screens";
command.Cell("N44").FormulaA1=$"IFERROR({DataLookup("Y")}-IF({DataLookup("U")}=\"FLAG\",1,0),\"N/A\")";
command.Cell("J45").Value="Unexplained P&L";
SetLookup(command,"N45","Q",Ui.Money);
command.Cell("J46").Value="Review threshold";
command.Cell("N46").Value=J.N(pnlCalibration,"defaultSigma");
command.Cell("N46").Style.NumberFormat.Format="0.0";
command.Cell("J47").Value="MORNING STATUS";
command.Cell("N47").FormulaA1="IF(ABS(N42)>N46,\"REVIEW P&L\",IF(N44>0,\"REVIEW FEED / BOOK\",\"CLEAR\"))";
command.Cell("N47").Style.Font.Bold=true;
command.Cell("N47").Style.Font.FontColor=Ui.Red;
command.Cell("J49").FormulaA1="HYPERLINK(\"#'04 PnL Explain'!B12\",\"OPEN P&L EXPLAIN >\")";
command.Cell("J50").FormulaA1="HYPERLINK(\"#'05 Incident Monitor'!B12\",\"OPEN INCIDENT SCREENS >\")";
command.Cell("J49").Style.Font.FontColor=Ui.Blue;
command.Cell("J50").Style.Font.FontColor=Ui.Blue;

// Single, exact bp input for ad-hoc first-order P&L.
// Different from historical stress recalculation; user retains full Stress Lab.
Ui.Section(command,"B53:P53","QUICK WHAT-IF | STATIC CLOSING BOOK / FIRST-ORDER ONLY");
command.Cell("B55").Value="PARALLEL YIELD SHOCK (BP)";
command.Cell("F55").Value=10.0;
command.Cell("F55").Style.Fill.BackgroundColor=Ui.Editable;
command.Cell("F55").Style.Font.FontColor=Ui.Blue;
command.Cell("F55").Style.Font.Bold=true;
command.Cell("F55").Style.NumberFormat.Format="+0.0;-0.0;0.0";
command.Cell("J55").Value="EST. SHOCK P&L / USD";
command.Cell("N55").FormulaA1="-F55*SUM('02 Portfolio Risk'!L14:L24)";
command.Cell("N55").Style.NumberFormat.Format=Ui.Money;
command.Cell("N55").Style.Font.Bold=true;
command.Cell("B56").Value="Pricing is not rerun; convexity and live trading positions are excluded.";
command.Cell("B56").Style.Font.FontColor=Ui.Muted;
command.Cell("J57").FormulaA1="HYPERLINK(\"#'07 Stress Lab'!B9\",\"OPEN 11-VERTEX STRESS LAB >\")";
command.Cell("J57").Style.Font.FontColor=Ui.Blue;
Ui.Note(command,60,"Simulation only. Date-linked P&L and VaR use 2024 reporting data; the Key-Rate chart and Quick Stress use the fixed 2024-12-31 closing book. An alert is not a confirmed incident or limit breach.");

// ---------------------------------------------------------------------
// 02 PORTFOLIO RISK | STATIC as-of closing book
// ---------------------------------------------------------------------
Ui.Banner(portfolio,"PORTFOLIO RISK  /  CLOSING BOOK");
Ui.Section(portfolio,"B7:P7","STATIC BOOK  /  2024-12-31  |  DOES NOT CHANGE WITH REPORT DATE");
portfolio.Cell("B9").Value="NET PORTFOLIO VALUE (USD)";
portfolio.Cell("F9").Value=J.N(risk,"netValue");
portfolio.Cell("F9").Style.NumberFormat.Format=Ui.Money;
portfolio.Cell("J9").Value="GROSS POSITION DV01";
portfolio.Cell("N9").Value=J.N(risk,"grossPositionDv01");
portfolio.Cell("N9").Style.NumberFormat.Format=Ui.Number;
Ui.Section(portfolio,"B12:H12","POSITIONS  /  CLOSING BOOK");
string[] positionHeader={"BOND","FACE (MM)","MATURITY Y","VALUE USD","DV01 USD/BP","SIDE","RISK SHARE"};
for(int n=0;n<positionHeader.Length;n++) portfolio.Cell(14,2+n).Value=positionHeader[n];
Ui.Head(portfolio,"B14:H14");
portfolio.Columns("B:H").Width=18;
portfolio.Column("E").Width=23;
portfolio.Column("F").Width=20;
double gross=J.N(risk,"grossPositionDv01");
for (int n=0;n<positions.Count;n++)
{
    int r=15+n; var p=positions[n];double dv=J.N(p,"dv01");
    portfolio.Cell(r,2).Value=J.S(p,"bond");
    portfolio.Cell(r,3).Value=J.N(p,"face_m");
    portfolio.Cell(r,4).Value=J.N(p,"years");
    portfolio.Cell(r,4).Style.NumberFormat.Format="0.00";
    portfolio.Cell(r,5).Value=J.N(p,"value");
    portfolio.Cell(r,6).Value=dv;
    portfolio.Cell(r,7).Value=J.N(p,"face_m")>=0?"LONG":"SHORT";
    portfolio.Cell(r,8).Value=gross==0?0:Math.Abs(dv)/gross;
    portfolio.Cell(r,8).Style.NumberFormat.Format="0.0%";
    portfolio.Cell(r,5).Style.NumberFormat.Format=Ui.Money;
    portfolio.Cell(r,6).Style.NumberFormat.Format=Ui.Number;
    Ui.Stripe(portfolio,r,2,8,n);
    Ui.Sign(portfolio.Cell(r,6),dv);
}
portfolio.Range(14,2,14+positions.Count,8).SetAutoFilter();
portfolio.SheetView.FreezeRows(14);
portfolio.Cell("J13").Value="TENOR";
portfolio.Cell("L13").Value="DV01 USD/BP";
portfolio.Cell("N13").Value="2023 DV01";
Ui.Head(portfolio,"J13:P13");
for(int i=0;i<ladder.Count;i++)
{
    int r=14+i;var a=ladder[i];
    portfolio.Cell(r,10).Value=J.S(a,"tenor");
    portfolio.Cell(r,12).Value=J.N(a,"dv01");
    portfolio.Cell(r,14).Value=J.N(a,"yearAgo");
    portfolio.Cell(r,12).Style.NumberFormat.Format=Ui.Number;
    portfolio.Cell(r,14).Style.NumberFormat.Format=Ui.Number;
    Ui.Sign(portfolio.Cell(r,12),J.N(a,"dv01"));
}
portfolio.Cell("J26").Value="VaR (99%, 500 scenarios)";
portfolio.Cell("N26").Value=J.N(risk,"var99");
portfolio.Cell("N26").Style.NumberFormat.Format=Ui.Money;
Ui.Section(portfolio,"J29:P29","WORST HISTORICAL P&L SCENARIOS");
portfolio.Column("M").Width=23; // prevents ######## for large negative scenario P&L
var worst=scenarios.OrderBy(x=>J.N(x,"pnl")).Take(5).ToList();
for(int i=0;i<worst.Count;i++)
{
    portfolio.Cell(31+i,10).Value=J.S(worst[i],"date");
    portfolio.Cell(31+i,13).Value=J.N(worst[i],"pnl");
    portfolio.Cell(31+i,13).Style.NumberFormat.Format=Ui.Money;
}
Ui.Note(portfolio,39,"The year-end position sheet is static; switch dates for historical DV01 and VaR on Command Center.");
Ui.Section(portfolio,"B42:P42","STATIC AS-OF 2024-12-31 | RISK LADDER AND HISTORICAL TAIL");
string curveBase64=J.S(risk,"curveImage");
string lossBase64=J.S(risk,"lossImage");
if (!string.IsNullOrWhiteSpace(curveBase64))
{
    using var curveStream=new MemoryStream(Convert.FromBase64String(curveBase64));
    portfolio.AddPicture(curveStream,XLPictureFormat.Png)
        .MoveTo(portfolio.Cell("B44")).Scale(0.53);
}
if (!string.IsNullOrWhiteSpace(lossBase64))
{
    using var lossStream=new MemoryStream(Convert.FromBase64String(lossBase64));
    portfolio.AddPicture(lossStream,XLPictureFormat.Png)
        .MoveTo(portfolio.Cell("J44")).Scale(0.47);
}


// ---------------------------------------------------------------------
// 03 RISK ATTRIBUTION | dynamic full risk walk
// ---------------------------------------------------------------------
Ui.Banner(attribution,"DAILY RISK ATTRIBUTION  /  DV01 & VaR");
Ui.Section(attribution,"B7:P7","DAILY RISK WALK | F9 LOCAL DATE OVERRIDE OR FOLLOW GLOBAL");
AddLocalDateSelector(attribution);
attribution.Cell("B11").Value="DRIVER";
attribution.Cell("F11").Value="DV01 (USD/BP)";
attribution.Cell("K11").Value="VaR (USD)";
Ui.Head(attribution,"B11:P11");
string[] drivers={"Opening","Time / Aging","Market / Curve","Window Roll","Trades / Amendments","Closing"};
string[] dvCols={"B","C","D","", "E","F"};
string[] varCols={"G","H","I","J","K","L"};
for(int i=0;i<drivers.Length;i++)
{
    int r=12+i;
    attribution.Cell(r,2).Value=drivers[i];
    if (dvCols[i]!="") SetLookup(attribution,$"F{r}",dvCols[i],Ui.Number);
    else attribution.Cell(r,6).Value="N/A";
    SetLookup(attribution,$"K{r}",varCols[i],Ui.Money);
    Ui.Stripe(attribution,r,2,16,i);
}
Ui.Section(attribution,"B20:P20","TELESCOPING WALK RECONCILIATION");
attribution.Cell("B22").Value="DV01: Close - Open - Time - Market - Trades";
attribution.Cell("K22").FormulaA1="ROUND(F17-F12-F13-F14-F16,2)";
attribution.Cell("K22").Style.NumberFormat.Format=Ui.Number;
attribution.Cell("B23").Value="VaR: Close - Open - Time - Market - Window - Trades";
attribution.Cell("K23").FormulaA1="ROUND(K17-SUM(K12:K16),2)";
attribution.Cell("K23").Style.NumberFormat.Format=Ui.Money;
attribution.Cell("B25").Value="RECONCILIATION";
attribution.Cell("K25").FormulaA1="IF(AND(ABS(K22)<0.01,ABS(K23)<0.01),\"PASS\",\"CHECK\")";
Ui.Note(attribution,29,"A change in VaR is a sequential model attribution, not evidence of an error or realized P&L.");

// ---------------------------------------------------------------------
// 04 P&L EXPLAIN | dynamic daily residuals and 11 tenor contributions
// ---------------------------------------------------------------------
Ui.Banner(pnl,"P&L EXPLAIN  /  DAILY RECONCILIATION");
Ui.Section(pnl,"B7:P7","DAILY P&L EXPLAIN | F9 LOCAL DATE OVERRIDE OR FOLLOW GLOBAL");
AddLocalDateSelector(pnl);
string[] pnlNames={"Reported P&L","Carry / Aging","Market P&L (11 key rates)",
                   "Predicted P&L","Unexplained P&L","Noise SD (expected)","Z-score"};
string[] pnlCols={"M","N","O","P","Q","R","S"};
for(int i=0;i<pnlNames.Length;i++)
{
    int r=12+i;
    pnl.Cell(r,2).Value=pnlNames[i];
    SetLookup(pnl,$"F{r}",pnlCols[i],i==6?"0.00":Ui.Money);
    Ui.Stripe(pnl,r,2,9,i);
}
pnl.Cell("B21").Value="P&L RECONCILIATION / USD";
pnl.Cell("F21").FormulaA1="F12-F15-F16";
pnl.Cell("F21").Style.NumberFormat.Format=Ui.Money;
pnl.Cell("B22").Value="Scaled exception threshold (sigma)";
pnl.Cell("F22").Value=J.N(pnlCalibration,"defaultSigma");
pnl.Cell("F22").Style.Fill.BackgroundColor=Ui.Editable;
pnl.Cell("F22").Style.Font.FontColor=Ui.Blue;
pnl.Cell("B23").Value="Risk-scaled exception?";
pnl.Cell("F23").FormulaA1="IF(ABS(F18)>F22,\"REVIEW\",\"CLEAR\")";
Ui.Section(pnl,"J11:P11","TENOR CONTRIBUTION / USD");
for(int i=0;i<ladder.Count;i++)
{
    int r=12+i;
    pnl.Cell(r,10).Value=J.S(ladder[i],"tenor");
    string col=Ui.ExcelColumn(30+i);
    SetLookup(pnl,$"N{r}",col,Ui.Money);
}
pnl.Cell("J24").Value="TOTAL TENOR MARKET P&L";
pnl.Cell("N24").FormulaA1="SUM(N12:N22)";
pnl.Cell("N24").Style.NumberFormat.Format=Ui.Money;
pnl.Cell("J25").Value="TENOR P&L RESIDUAL";
pnl.Cell("N25").FormulaA1="N24-F14";
pnl.Cell("N25").Style.NumberFormat.Format=Ui.Money;
Ui.Note(pnl,29,"Risk-scaled thresholds are local investigation settings; they do not retrain the Task 5 policy.");

// ---------------------------------------------------------------------
// 05 INCIDENT MONITOR | observed screening separate from forensic truth
// ---------------------------------------------------------------------
Ui.Banner(incident,"INCIDENT MONITOR  /  MARKET DATA & BOOKING");
Ui.Section(incident,"B7:P7","OBSERVED 2024 SCREENS | F9 LOCAL DATE OVERRIDE OR FOLLOW GLOBAL");
AddLocalDateSelector(incident);
string[] flagLabels={"P&L fixed threshold","P&L risk-scaled","Stale curve vertex",
                     "Curve kink jump","Late booking ticket"};
string[] flagCols={"T","U","V","W","X"};
for(int i=0;i<flagLabels.Length;i++)
{
    int r=12+i;
    incident.Cell(r,2).Value=flagLabels[i];
    SetLookup(incident,$"J{r}",flagCols[i]);
    Ui.Stripe(incident,r,2,13,i);
}
incident.Cell("B18").Value="TOTAL TRIGGERED SCREENS";
SetLookup(incident,"J18","Y",Ui.Number);
Ui.Section(incident,"B21:P21","SOURCE EVIDENCE ON SELECTED DATE");
incident.Cell("B23").Value="Stale run evidence";
SetLookup(incident,"F23","Z");
incident.Cell("B24").Value="Kink jump evidence";
SetLookup(incident,"F24","AA");
incident.Cell("B25").Value="Late booking evidence";
SetLookup(incident,"F25","AB");
incident.Columns("F:P").Width=16;
incident.Column("F").Width=24;
Ui.Section(incident,"B29:P29","RETROSPECTIVE CASE REGISTER | EVALUATION ONLY");
string[] caseHead={"CASE","KIND","FIRST","LAST","INSTRUMENT","COMMENT"};
int[] caseCol={2,4,6,8,10,12};
for(int i=0;i<caseHead.Length;i++) incident.Cell(31,caseCol[i]).Value=caseHead[i];
Ui.Head(incident,"B31:P31");
for(int i=0;i<cases.Count;i++)
{
    int r=32+i;var c=cases[i];
    incident.Cell(r,2).Value=J.S(c,"id");
    incident.Cell(r,4).Value=J.S(c,"kind");
    incident.Cell(r,6).Value=J.S(c,"first");
    incident.Cell(r,8).Value=J.S(c,"last");
    incident.Cell(r,10).Value=J.S(c,"bond");
    incident.Cell(r,12).Value=J.S(c,"what");
    incident.Range(r,12,r,16).Merge();
    incident.Row(r).Height=36;
    incident.Range(r,12,r,16).Style.Alignment.WrapText=true;
    Ui.Stripe(incident,r,2,16,i);
}
Ui.Note(incident,40,"Cases are a hindsight evaluation key from the training lab, not a live risk screen or confirmed operational incident.");

// ---------------------------------------------------------------------
// 06 POLICY GOVERNANCE | offline five-policy experiment
// ---------------------------------------------------------------------
Ui.Banner(governance,"POLICY GOVERNANCE  /  OFFLINE REVIEW");
Ui.Section(governance,"B7:P7","PRECOMMITTED SELECTION RULE | 1,200 SEEDED SYNTHETIC INCIDENTS");
governance.Cell("B9").Value="SELECTED POLICY";
governance.Cell("F9").Value=J.S(policy,"winner");
governance.Cell("F9").Style.Font.FontColor=Ui.Teal;
governance.Cell("F9").Style.Font.Bold=true;
governance.Cell("B10").Value="ACCEPT: overall >= 90%; each kind >= 80%; <=5 clean flags / year";
governance.Cell("B11").Value="Tie-break: fewest clean flags, then shorter median detection delay";
governance.Cell("B11").Style.Font.FontColor=Ui.Muted;
string[] policyHead={"POLICY","OVERALL","MIN KIND","CLEAN 23","CLEAN 24","DELAY","PASS/FAIL"};
int[] policyCols={2,5,7,9,11,13,15};
for(int i=0;i<policyHead.Length;i++) governance.Cell(13,policyCols[i]).Value=policyHead[i];
Ui.Head(governance,"B13:P13");
for(int i=0;i<policies.Count;i++)
{
    int r=14+i;var p=policies[i];
    double minKind=p.GetProperty("byKind").EnumerateArray().Min(x=>J.N(x,"rate"));
    governance.Cell(r,2).Value=J.S(p,"name");
    governance.Cell(r,5).Value=J.N(p,"overall");
    governance.Cell(r,5).Style.NumberFormat.Format="0.0%";
    governance.Cell(r,7).Value=minKind;
    governance.Cell(r,7).Style.NumberFormat.Format="0.0%";
    governance.Cell(r,9).Value=J.N(p,"clean2023");
    governance.Cell(r,11).Value=J.N(p,"clean2024");
    governance.Cell(r,13).Value=J.N(p,"medianDelay");
    governance.Cell(r,15).Value=J.B(p,"passes")?"PASS":"FAIL";
    Ui.Stripe(governance,r,2,16,i);
    if (J.B(p,"passes"))
    {
        governance.Range(r,2,r,16).Style.Fill.BackgroundColor=XLColor.FromHtml("#DDF5EB");
        governance.Range(r,2,r,16).Style.Font.Bold=true;
    }
}
Ui.Section(governance,"B21:P21","SIX INCIDENT CATEGORIES | DETECTION COVERAGE");
string[] kinds={"WRONG_SIDE","NOTIONAL","LATE","MISSED","STALE","BAD_PRINT"};
for(int i=0;i<kinds.Length;i++) governance.Cell(23,i+5).Value=kinds[i];
Ui.Head(governance,"B23:P23");
for(int i=0;i<policies.Count;i++)
{
    int r=24+i;var p=policies[i];
    governance.Cell(r,2).Value=J.S(p,"name");
    foreach(var row in p.GetProperty("byKind").EnumerateArray())
    {
        int col=Array.IndexOf(kinds,J.S(row,"kind"))+5;
        if(col>=5)
        {
            governance.Cell(r,col).Value=J.N(row,"rate");
            governance.Cell(r,col).Style.NumberFormat.Format="0.0%";
        }
    }
    Ui.Stripe(governance,r,2,16,i);
}
Ui.Note(governance,33,"All policy scores come from seeded Task 5 simulation; selection is not a production approval or risk limit.");

// ---------------------------------------------------------------------
// 07 STRESS LAB | static 2024-12-31 portfolio, editable shocks
// ---------------------------------------------------------------------
Ui.Banner(stress,"STRESS LAB  /  YIELD CURVE WHAT-IF");
Ui.Section(stress,"B7:P7","CLOSING BOOK 2024-12-31 | LINEAR KEY-RATE DV01 ONLY");
stress.Cell("B9").Value="SCENARIO";
stress.Cell("C9").Value="Parallel +10bp";
stress.Cell("C9").Style.Fill.BackgroundColor=Ui.Editable;
stress.Cell("C9").Style.Font.FontColor=Ui.Blue;
stress.Column("C").Width=22;
string[] scenarioNames={"Parallel +10bp","Parallel -10bp","Bear Steepener",
                        "Bull Flattener","Front-End +25bp","Custom"};
for(int i=0;i<scenarioNames.Length;i++) stress.Cell(9+i,17).Value=scenarioNames[i];
stress.Cell("C9").CreateDataValidation().List(stress.Range("Q9:Q14"));
stress.Column("Q").Hide();

string[] stressHeads={"TENOR","DV01 USD/BP","CUSTOM BP","APPLIED BP","STRESS P&L"};
for(int i=0;i<stressHeads.Length;i++) stress.Cell(11,2+i).Value=stressHeads[i];
string[] presetHeads={"Parallel +10bp","Parallel -10bp","Bear Steepener","Bull Flattener","Front-End +25bp"};
for(int i=0;i<presetHeads.Length;i++) stress.Cell(11,8+i).Value=presetHeads[i];
Ui.Head(stress,"B11:L11");
for(int i=0;i<ladder.Count;i++)
{
    int r=12+i;
    double tenorYears = new double[]{0.25,0.5,1,2,3,5,7,10,15,20,30}[i];
    stress.Cell(r,2).Value=J.S(ladder[i],"tenor");
    stress.Cell(r,3).Value=J.N(ladder[i],"dv01");
    stress.Cell(r,3).Style.NumberFormat.Format=Ui.Number;
    stress.Cell(r,4).Value=0.0;
    stress.Cell(r,4).Style.Fill.BackgroundColor=Ui.Editable;
    stress.Cell(r,4).Style.Font.FontColor=Ui.Blue;
    stress.Cell(r,8).Value=10;
    stress.Cell(r,9).Value=-10;
    stress.Cell(r,10).Value=tenorYears<2?0:(tenorYears<5?2:(tenorYears<10?5:(tenorYears<15?15:20)));
    stress.Cell(r,11).Value=tenorYears<2?-20:(tenorYears<5?-15:(tenorYears<10?-10:-5));
    stress.Cell(r,12).Value=tenorYears<=2?25:0;
    stress.Cell(r,5).FormulaA1=$"IF($C$9=\"Custom\",D{r},INDEX($H{r}:$L{r},1,MATCH($C$9,$H$11:$L$11,0)))";
    stress.Cell(r,6).FormulaA1=$"-C{r}*E{r}";
    stress.Cell(r,6).Style.NumberFormat.Format=Ui.Money;
    Ui.Stripe(stress,r,2,12,i);
    // Highlight editable cells after striping.
    stress.Cell(r,4).Style.Fill.BackgroundColor=Ui.Editable;
}
stress.Cell("B25").Value="TOTAL STRESS P&L";
stress.Cell("F25").FormulaA1="SUM(F12:F22)";
stress.Cell("F25").Style.NumberFormat.Format=Ui.Money;
Ui.Total(stress,"B25:F25");
Ui.Note(stress,28,"Editable CUSTOM BP cells apply only when Scenario = Custom. Does not recalculate historical VaR or full pricing.");

// ---------------------------------------------------------------------
// 08 METHODOLOGY | ownership, control semantics and QA evidence
// ---------------------------------------------------------------------
Ui.Banner(guide,"METHOD / VALIDATION / DATA LINEAGE");
Ui.Section(guide,"B7:P7","REPORT OWNERSHIP & CONTROL BOUNDARIES");
(string module,string purpose)[] methodRows={
    ("Task 1 | Valuation","Bond cash flows, zero-curve pricing, 11-vertex DV01; 500-scenario 99% HS VaR."),
    ("Task 2 | Risk Walk","Time, Market, Window, Trades; exact telescoping DV01 and VaR reconciliation."),
    ("Task 3 | P&L Explain","Reported = Predicted + Unexplained; 2023 noise calibration; daily Z-score."),
    ("Task 4 | Observed Screens","P&L Fixed/Scaled, Stale, Kink, Late; cases kept as retrospective evidence."),
    ("Task 5 | Policy Review","1,200 synthetic incidents, 5 candidate policies, precommitted selection."),
    ("Report date controls","Command C6 sets global default; Task 2-4 F9 local dropdown overrides only its own page; L9 shows active date."),
    ("Closing book / Stress","Task 1 book is fixed on 2024-12-31; first-order shock P&L, not full reprice."),
    ("V4 command trends","20D / 60D / YTD sampling; glyph-based Excel formula trends use 32 sampled points. Full 2024 daily series is stored in _DATA."),
    ("Quick parallel shock","Command F55 shock uses the static 2024-12-31 Key-Rate sum; formula P&L = -Net DV01 x shock (bp)."),
    ("Control scope","No approved desk limits, live feed, independent P&L source certification or real approvals.")
};
for(int i=0;i<methodRows.Length;i++)
{
    int r=9+i;
    guide.Cell(r,2).Value=methodRows[i].module;
    guide.Range(r,6,r,16).Merge();
    guide.Cell(r,6).Value=methodRows[i].purpose;
    guide.Range(r,6,r,16).Style.Alignment.WrapText=true;
    guide.Row(r).Height=32;
    Ui.Stripe(guide,r,2,16,i);
}
Ui.Section(guide,"B20:P20","CROSS-MODULE QA | VERIFIED IN FINAL.IPYNB");
for(int i=0;i<qa.Count;i++)
{
    int r=22+i;
    guide.Cell(r,2).Value=J.S(qa[i],"name");
    guide.Cell(r,8).Value=J.S(qa[i],"status");
    guide.Range(r,10,r,16).Merge();
    guide.Cell(r,10).Value=J.S(qa[i],"detail");
    guide.Range(r,10,r,16).Style.Alignment.WrapText=true;
    guide.Row(r).Height=27;
    Ui.Stripe(guide,r,2,16,i);
    guide.Cell(r,8).Style.Font.FontColor=Ui.Teal;
}
Ui.Note(guide,39,"Date controls: FOLLOW GLOBAL syncs pages with Command C6; local F9 overrides apply to one worksheet only.");
guide.Cell("B40").Value="Educational RiskHub extension; simulated data, not production risk.";
guide.Cell("B41").Value="RiskHub source: https://riskhub.org/python-labs/fixed-income-risk-dashboard";

hidden.Hide();
trendData.Hide();
wb.SaveAs(destination);
Console.WriteLine("SUCCESS: " + destination);
Console.WriteLine($"Global 2024 reporting days: {dates.Count}; positions: {positions.Count}; scenarios: {scenarios.Count}");
Console.WriteLine($"Closing VaR: USD {J.N(risk,"var99"):N0}; Net DV01: {J.N(risk,"netDv01"):N0} USD/bp");
Console.WriteLine($"Offline policy: {J.S(policy,"winner")}; Python QA checks: {qa.Count} PASS");

// =====================================================================
// HELPERS: simple JSON access and consistent Excel presentation
// =====================================================================
public static class J
{
    public static string S(JsonElement e,string key)
    {
        if (!e.TryGetProperty(key,out var v) || v.ValueKind==JsonValueKind.Null)
            return "";
        return v.ValueKind==JsonValueKind.String ? (v.GetString()??"") : v.ToString();
    }
    public static double N(JsonElement e,string key) => e.GetProperty(key).GetDouble();
    public static bool B(JsonElement e,string key) => e.GetProperty(key).GetBoolean();
}
public static class Ui
{
    public const string Money="\"USD \"#,##0;[Red](\"USD \"#,##0);\"USD 0\"";
    public const string Number="+#,##0;[Red](#,##0);-";
    public static readonly XLColor Navy=XLColor.FromHtml("#14263D");
    public static readonly XLColor Blue=XLColor.FromHtml("#244767");
    public static readonly XLColor Teal=XLColor.FromHtml("#078C86");
    public static readonly XLColor Red=XLColor.FromHtml("#CD5260");
    public static readonly XLColor Pale=XLColor.FromHtml("#EDF3F8");
    public static readonly XLColor Editable=XLColor.FromHtml("#DFEBFF");
    public static readonly XLColor Muted=XLColor.FromHtml("#63768A");
    public static void Base(IXLWorksheet ws,string asOf)
    {
        ws.ShowGridLines=false;
        ws.Style.Font.FontName="Aptos";
        ws.Style.Font.FontSize=10;
        ws.Style.Font.FontColor=Navy;
        ws.Column("A").Width=3;
        ws.Columns("B:P").Width=13;
        ws.Columns("B:D").Width=17;
        ws.Row(2).Height=27;
        ws.Row(3).Height=25;
        ws.Row(5).Height=25;
        ws.Cell("B4").Value=$"FIXED INCOME | 2024 EOD | SIMULATED | Closing book: {asOf}";
        ws.Cell("B4").Style.Font.FontColor=Muted;
        ws.Cell("B4").Style.Font.FontSize=9;
    }
    public static void Banner(IXLWorksheet ws,string label)
    {
        var r=ws.Range("B2:P3");r.Merge();r.FirstCell().Value=label;
        r.Style.Fill.BackgroundColor=Navy;
        r.Style.Font.FontColor=XLColor.White;
        r.Style.Font.FontSize=18;
        r.Style.Font.Bold=true;
        r.Style.Alignment.Vertical=XLAlignmentVerticalValues.Center;
    }
    public static void Nav(IXLWorksheet ws,string[] sheets)
    {
        string[] cols={"B","D","F","H","J","L","N","P"};
        string[] labels={"COMMAND","BOOK","ATTRIBUTION","P&L","INCIDENTS","POLICY","STRESS","METHOD"};
        for(int i=0;i<sheets.Length;i++)
        {
            var c=ws.Cell(cols[i]+"5");
            c.FormulaA1=$"HYPERLINK(\"#'{sheets[i]}'!B2\",\"{labels[i]}\")";
            c.Style.Fill.BackgroundColor=Pale;
            c.Style.Font.Bold=true;
            c.Style.Font.FontColor=Blue;
            c.Style.Alignment.Horizontal=XLAlignmentHorizontalValues.Center;
        }
    }
    public static void Section(IXLWorksheet ws,string range,string label)
    {
        var r=ws.Range(range);r.Merge();r.FirstCell().Value=label;
        r.Style.Fill.BackgroundColor=Blue;
        r.Style.Font.FontColor=XLColor.White;
        r.Style.Font.Bold=true;
        r.Style.Alignment.Vertical=XLAlignmentVerticalValues.Center;
    }
    public static void Head(IXLWorksheet ws,string range)
    {
        var r=ws.Range(range);
        r.Style.Fill.BackgroundColor=Blue;
        r.Style.Font.FontColor=XLColor.White;
        r.Style.Font.Bold=true;
        r.Style.Alignment.Vertical=XLAlignmentVerticalValues.Center;
        ws.Row(r.RangeAddress.FirstAddress.RowNumber).Height=27;
    }
    public static void Kpi(IXLWorksheet ws,string labelAddress,string valueAddress,
                           string label,string expression,string fmt)
    {
        var heading=ws.Range(labelAddress);heading.Merge();heading.FirstCell().Value=label;
        heading.Style.Fill.BackgroundColor=Pale;
        heading.Style.Font.FontColor=Muted;
        heading.Style.Font.Bold=true;
        heading.Style.Font.FontSize=9;
        var value=ws.Range(valueAddress);value.Merge();value.FirstCell().FormulaA1=expression;
        value.Style.Fill.BackgroundColor=Pale;
        value.Style.Font.FontColor=Navy;
        value.Style.Font.FontSize=18;
        value.Style.Font.Bold=true;
        value.Style.Alignment.Horizontal=XLAlignmentHorizontalValues.Center;
        value.Style.Alignment.Vertical=XLAlignmentVerticalValues.Center;
        value.Style.NumberFormat.Format=fmt;
    }
    public static void Stripe(IXLWorksheet ws,int row,int c1,int c2,int index)
    {
        if(index%2==0)ws.Range(row,c1,row,c2).Style.Fill.BackgroundColor=XLColor.FromHtml("#F4F7FA");
    }
    public static void Sign(IXLCell cell,double v)
    {
        cell.Style.Font.FontColor=v<0?Red:Teal;
        cell.Style.Font.Bold=true;
    }
    public static void Total(IXLWorksheet ws,string address)
    {
        var r=ws.Range(address);
        r.Style.Fill.BackgroundColor=Pale;
        r.Style.Font.FontColor=Navy;
        r.Style.Font.Bold=true;
        r.Style.Font.FontSize=13;
    }
    public static void Note(IXLWorksheet ws,int row,string note)
    {
        var r=ws.Range(row,2,row,16);
        r.Merge();r.FirstCell().Value=note;
        r.Style.Fill.BackgroundColor=Pale;
        r.Style.Font.FontColor=Muted;
        r.Style.Font.FontSize=9;
        r.Style.Alignment.WrapText=true;
        ws.Row(row).Height=32;
    }
    public static string ExcelColumn(int col)
    {
        string answer="";
        while(col>0){col--;answer=(char)('A'+col%26)+answer;col/=26;}
        return answer;
    }
}
