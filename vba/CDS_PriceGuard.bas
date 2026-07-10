Attribute VB_Name = "CDS_PriceGuard"
Option Explicit

' ============================================================
' CDS Price Staleness Guard
'
' Compares each holding's FMV/Quantity "implied price" already on the
' report against a fresh quote pulled through Excel's Stocks linked
' data type (Data > Stocks), and flags any ticker whose live price has
' drifted more than Settings!DriftAlertPct (default 1.0%) away from
' the report's implied price.
'
' RefreshLivePrices = FetchLivePrices + ApplyDriftAlerts:
'   FetchLivePrices  - builds/clears the "CDS Live Prices" helper
'                       sheet: one row per eligible holding with its
'                       Ticker, Class, Implied Price (report FMV/Qty,
'                       written as a value, never a live formula
'                       against the report), and a fresh Live Price
'                       pulled via ConvertToLinkedDataType + FIELDVALUE.
'                       Never writes to the report sheet.
'   ApplyDriftAlerts - reads the helper sheet, computes Drift % per
'                       row, colors the report's ticker cell (col C)
'                       red when |drift| exceeds the threshold, clears
'                       it when it no longer does (idempotent across
'                       reruns), and writes a one-line summary to the
'                       Sell Workbench status block (row 9, same
'                       column as "Plan Status") when a workbench
'                       exists on the sheet.
' Kept as two public procedures so tests can seed the helper sheet
' directly and call ApplyDriftAlerts without needing live network
' connectivity.
'
' Skips: class CASH and SHORT (money market - price is $1 and drift is
' meaningless), tickers IsCUSIPSettings() identifies as bond CUSIPs (no
' tradeable public quote), and any row with no/zero Quantity (no
' implied price to compare against).
'
' Excel Stocks linked data types require Microsoft 365 with an active
' connection. Range.ConvertToLinkedDataType is invoked through a late-
' bound Object reference (not a typed Range) specifically so a build
' of Excel that does not expose the member at all still compiles and
' runs this module - it just fails the call at runtime, which is
' caught like any other error. When that happens, FetchLivePrices
' notes every row "linked data types unavailable", shows one MsgBox,
' and returns False; RefreshLivePrices then skips ApplyDriftAlerts
' entirely so the report is not touched. When the call succeeds but an
' individual ticker never resolves within the ~10s bounded wait (slow
' or missing network), that single row is simply "n/a - not resolved"
' - not treated as a whole-run failure.
' ============================================================

Private Const PRICE_SHEET_NAME As String = "CDS Live Prices"
Private Const STOCKS_SERVICE_ID As Long = 268435456
Private Const NOTE_UNRESOLVED As String = "n/a - not resolved"
Private Const NOTE_UNAVAILABLE As String = "linked data types unavailable"
Private Const FETCH_TIMEOUT_SECONDS As Double = 10

' ============================================================
' PUBLIC ENTRY POINT
' ============================================================
Public Sub RefreshLivePrices()
    If Not FetchLivePrices() Then Exit Sub
    ApplyDriftAlerts
End Sub

' ============================================================
' FETCH: build the helper sheet, pull live quotes, freeze to values.
' Never writes to the report sheet. Returns False only on a guard
' failure or when ConvertToLinkedDataType itself is unusable.
' ============================================================
Public Function FetchLivePrices() As Boolean
    Dim prevScreenUpdating As Boolean
    prevScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Function

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Open the client's workbook first.", vbCritical
        Exit Function
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim headerRow As Long, dataStart As Long, dataEnd As Long, totalRow As Long
    If Not GetReportStatePG(ws, headerRow, dataStart, dataEnd, totalRow) Then
        MsgBox "Run ProcessCDSHoldings first.", vbExclamation
        Exit Function
    End If

    Application.ScreenUpdating = False

    Dim priceWs As Worksheet
    Set priceWs = GetOrCreatePriceSheetPG(ws)
    priceWs.Cells.Clear

    priceWs.Cells(1, 1).Value = "Checked: " & Format(Now, "yyyy-mm-dd hh:mm")
    priceWs.Cells(1, 1).Font.Bold = True

    WritePriceHeaderRowPG priceWs

    Const COL_TICKER As Long = 1
    Const COL_CLASS As Long = 2
    Const COL_IMPLIED As Long = 3
    Const COL_LIVE As Long = 4
    Const COL_DRIFT As Long = 5
    Const COL_NOTE As Long = 6
    Const COL_HELPER As Long = 7

    Dim r As Long, outRow As Long
    Dim ticker As String, cls As String
    Dim qty As Variant, fmv As Variant

    outRow = 3

    For r = dataStart To dataEnd
        cls = UCase(Trim(CStr(ws.Cells(r, 1).Value)))
        ticker = Trim(CStr(ws.Cells(r, 3).Value))
        qty = ws.Cells(r, 11).Value
        fmv = ws.Cells(r, 5).Value

        If cls = "CASH" Or cls = "SHORT" Then
            ' money market / cash - price is $1 and meaningless, skip
        ElseIf ticker = "" Then
            ' nothing to look up
        ElseIf IsCUSIPSettings(ticker) Then
            ' bond CUSIP - no tradeable public quote
        ElseIf Not IsNumeric(qty) Then
            ' no quantity, no implied price
        ElseIf CDbl(qty) = 0 Then
            ' zero quantity, no implied price
        Else
            priceWs.Cells(outRow, COL_TICKER).Value = ticker
            priceWs.Cells(outRow, COL_CLASS).Value = cls
            priceWs.Cells(outRow, COL_IMPLIED).Value = SafeDoublePG(fmv) / CDbl(qty)
            priceWs.Cells(outRow, COL_IMPLIED).NumberFormat = "#,##0.0000"
            outRow = outRow + 1
        End If
    Next r

    Dim dataLastRow As Long
    dataLastRow = outRow - 1

    ApplyPriceSheetFormattingPG priceWs, dataLastRow

    If dataLastRow < 3 Then
        ' No eligible holdings (all CASH/SHORT/CUSIP/no-qty) - still a valid run.
        FetchLivePrices = True
        GoTo Done
    End If

    ' --- Live price mechanics ---
    Dim i As Long
    For i = 3 To dataLastRow
        priceWs.Cells(i, COL_LIVE).Value = priceWs.Cells(i, COL_TICKER).Value
    Next i

    Dim liveRange As Range
    Set liveRange = priceWs.Range(priceWs.Cells(3, COL_LIVE), priceWs.Cells(dataLastRow, COL_LIVE))

    If Not TryConvertToLinkedDataTypePG(liveRange) Then
        For i = 3 To dataLastRow
            priceWs.Cells(i, COL_LIVE).ClearContents
            priceWs.Cells(i, COL_NOTE).Value = NOTE_UNAVAILABLE
        Next i
        MsgBox "Live prices are unavailable on this machine (Excel Stocks linked data " & _
               "types require Microsoft 365 with an internet connection). Implied prices " & _
               "are still listed below; the report itself has not been changed.", _
               vbInformation, "CDS Price Guard"
        FetchLivePrices = False
        GoTo Done
    End If

    ' Helper column resolves to a plain number once the linked record populates.
    For i = 3 To dataLastRow
        priceWs.Cells(i, COL_HELPER).Formula = "=IFERROR(FIELDVALUE(" & _
            priceWs.Cells(i, COL_LIVE).Address(False, False) & ",""Price""),"""")"
    Next i

    Dim startT As Double
    startT = Timer
    Do
        DoEvents
        Application.Calculate
        If AllPriceHelperCellsFilledPG(priceWs, COL_HELPER, 3, dataLastRow) Then Exit Do
    Loop While (Timer - startT) < FETCH_TIMEOUT_SECONDS And (Timer - startT) >= 0

    ' Freeze: overwrite Live Price with the resolved number (or blank it), and
    ' drop the linked-entity conversion so the sheet holds plain values only.
    Dim helperVal As Variant
    For i = 3 To dataLastRow
        helperVal = priceWs.Cells(i, COL_HELPER).Value
        If IsNumeric(helperVal) And Trim(CStr(helperVal)) <> "" Then
            priceWs.Cells(i, COL_LIVE).Value = CDbl(helperVal)
        Else
            priceWs.Cells(i, COL_LIVE).ClearContents
            priceWs.Cells(i, COL_NOTE).Value = NOTE_UNRESOLVED
        End If
    Next i

    priceWs.Range(priceWs.Cells(3, COL_HELPER), priceWs.Cells(dataLastRow, COL_HELPER)).ClearContents

    ' Drift % (a value, not a formula) wherever both prices are numeric.
    Dim impliedV As Variant, liveV As Variant
    For i = 3 To dataLastRow
        impliedV = priceWs.Cells(i, COL_IMPLIED).Value
        liveV = priceWs.Cells(i, COL_LIVE).Value
        If IsNumericValuePG(impliedV) And IsNumericValuePG(liveV) And CDbl(impliedV) <> 0 Then
            priceWs.Cells(i, COL_DRIFT).Value = (CDbl(liveV) - CDbl(impliedV)) / CDbl(impliedV) * 100
        End If
    Next i

    priceWs.Range(priceWs.Cells(3, COL_LIVE), priceWs.Cells(dataLastRow, COL_LIVE)).NumberFormat = "#,##0.0000"
    priceWs.Range(priceWs.Cells(3, COL_DRIFT), priceWs.Cells(dataLastRow, COL_DRIFT)).NumberFormat = "0.00"

    FetchLivePrices = True

Done:
    Application.ScreenUpdating = prevScreenUpdating
    Exit Function

ErrHandler:
    MsgBox "RefreshLivePrices failed while fetching prices: " & Err.Number & " - " & Err.Description, vbExclamation
    FetchLivePrices = False
    Resume Done
End Function

' ============================================================
' APPLY: color/clear report ticker cells from whatever is currently in
' the helper sheet, and write the workbench status line. Deliberately
' re-derives the report row for each ticker by matching col C rather
' than relying on stored row numbers, so this can be called standalone
' against seeded helper-sheet values (e.g. in tests) without having
' just run FetchLivePrices in the same session.
' ============================================================
Public Sub ApplyDriftAlerts()
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    On Error GoTo ErrHandler

    If ActiveWorkbook Is Nothing Then Exit Sub

    Dim priceWs As Worksheet
    On Error Resume Next
    Set priceWs = ActiveWorkbook.Worksheets(PRICE_SHEET_NAME)
    On Error GoTo 0
    If priceWs Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Or ws.Name = PRICE_SHEET_NAME Then
        Set ws = FindProcessedReportSheetPG(ActiveWorkbook)
        If ws Is Nothing Then Exit Sub
    End If

    Dim headerRow As Long, dataStart As Long, dataEnd As Long, totalRow As Long
    If Not GetReportStatePG(ws, headerRow, dataStart, dataEnd, totalRow) Then Exit Sub

    Dim driftThreshold As Double
    driftThreshold = GetSettingNum("DriftAlertPct", 1#)

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    PrepareCDSWorksheetForMacro ws

    Dim lastPriceRow As Long
    lastPriceRow = 2
    Do While Trim(CStr(priceWs.Cells(lastPriceRow + 1, 1).Value)) <> ""
        lastPriceRow = lastPriceRow + 1
    Loop

    Dim i As Long, reportRow As Long
    Dim tk As String
    Dim impliedV As Variant, liveV As Variant
    Dim driftNumeric As Boolean, driftPct As Double
    Dim totalTickers As Long, drifted As Long
    Dim maxAbsDrift As Double, maxSignedDrift As Double, maxTicker As String

    totalTickers = 0
    drifted = 0
    maxAbsDrift = -1

    For i = 3 To lastPriceRow
        tk = Trim(CStr(priceWs.Cells(i, 1).Value))
        If tk <> "" Then
            totalTickers = totalTickers + 1

            impliedV = priceWs.Cells(i, 3).Value
            liveV = priceWs.Cells(i, 4).Value
            driftNumeric = False

            If IsNumericValuePG(impliedV) And IsNumericValuePG(liveV) And CDbl(impliedV) <> 0 Then
                driftPct = (CDbl(liveV) - CDbl(impliedV)) / CDbl(impliedV) * 100
                priceWs.Cells(i, 5).Value = driftPct
                driftNumeric = True
            End If

            reportRow = FindReportRowByTickerPG(ws, dataStart, dataEnd, tk)

            If reportRow > 0 Then
                If driftNumeric And Abs(driftPct) > driftThreshold Then
                    ws.Cells(reportRow, 3).Interior.Color = RGB(255, 120, 120)
                    drifted = drifted + 1
                    If Abs(driftPct) > maxAbsDrift Then
                        maxAbsDrift = Abs(driftPct)
                        maxSignedDrift = driftPct
                        maxTicker = tk
                    End If
                Else
                    ws.Cells(reportRow, 3).Interior.Pattern = xlNone
                End If
            End If
        End If
    Next i

    Application.Calculate

    WritePriceStatusLinePG ws, totalTickers, drifted, driftThreshold, maxTicker, maxSignedDrift

Done:
    On Error Resume Next
    If Not ws Is Nothing Then ProtectCDSWorksheet ws
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Exit Sub

ErrHandler:
    MsgBox "ApplyDriftAlerts failed: " & Err.Number & " - " & Err.Description, vbExclamation
    Resume Done
End Sub

' ============================================================
' LIVE-FETCH HELPERS
' ============================================================

' Invoked through a late-bound Object (not a typed Range) so this
' module still compiles and imports on an Excel build whose Range type
' library does not define ConvertToLinkedDataType at all - the member
' lookup then happens at runtime and any failure lands in Err, same as
' a normal runtime error, instead of a VBE compile error that would
' block the whole project from loading.
Private Function TryConvertToLinkedDataTypePG(rng As Range) As Boolean
    Dim obj As Object
    Set obj = rng

    On Error Resume Next
    Err.Clear
    obj.ConvertToLinkedDataType ServiceID:=STOCKS_SERVICE_ID, LanguageCulture:="en-US"
    TryConvertToLinkedDataTypePG = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function AllPriceHelperCellsFilledPG(ws As Worksheet, col As Long, r1 As Long, r2 As Long) As Boolean
    Dim i As Long

    AllPriceHelperCellsFilledPG = True
    For i = r1 To r2
        If Not IsNumeric(ws.Cells(i, col).Value) Then
            AllPriceHelperCellsFilledPG = False
            Exit Function
        End If
    Next i
End Function

' ============================================================
' HELPER SHEET BUILD
' ============================================================
Private Function GetOrCreatePriceSheetPG(reportWs As Worksheet) As Worksheet
    Dim wb As Workbook
    Set wb = reportWs.Parent

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(PRICE_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=reportWs)
        ws.Name = PRICE_SHEET_NAME
    End If

    Set GetOrCreatePriceSheetPG = ws
End Function

Private Sub WritePriceHeaderRowPG(ws As Worksheet)
    Dim headers As Variant
    headers = Array("Ticker", "Class", "Implied Price", "Live Price", "Drift %", "Note")

    Dim i As Long
    For i = 0 To 5
        With ws.Cells(2, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Interior.Color = RGB(189, 215, 238)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    Next i
End Sub

Private Sub ApplyPriceSheetFormattingPG(ws As Worksheet, lastRow As Long)
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 10
    ws.Columns("F").ColumnWidth = 26

    If lastRow >= 3 Then
        With ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 6)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
    End If
End Sub

' ============================================================
' WORKBENCH STATUS LINE
' ============================================================
Private Sub WritePriceStatusLinePG(ws As Worksheet, totalTickers As Long, drifted As Long, _
                                   driftThreshold As Double, maxTicker As String, maxSignedDrift As Double)
    Dim modeCell As Range
    Set modeCell = FindCellExactPG(ws, "Sell Mode")
    If modeCell Is Nothing Then Exit Sub ' no sell workbench on this sheet

    Dim statusCol As Long
    statusCol = FindPlanStatusColPG(ws)
    If statusCol = 0 Then Exit Sub

    Dim checkedTime As String
    checkedTime = Format(Now, "hh:mm")

    Dim msg As String
    If drifted > 0 Then
        msg = "Prices checked " & checkedTime & " - " & drifted & " of " & totalTickers & _
              " drifted >" & FormatPctPG(driftThreshold) & "% (max: " & maxTicker & " " & _
              FormatSignedPctPG(maxSignedDrift) & ")"
    Else
        msg = "Prices checked " & checkedTime & " - no drift >" & FormatPctPG(driftThreshold) & "%"
    End If

    ws.Cells(9, statusCol).Value = msg
End Sub

Private Function FindPlanStatusColPG(ws As Worksheet) As Long
    Dim c As Range
    Set c = FindCellExactPG(ws, "Plan Status")
    If Not c Is Nothing Then FindPlanStatusColPG = c.Column
End Function

Private Function FormatPctPG(v As Double) As String
    FormatPctPG = Format(v, "0.0")
End Function

Private Function FormatSignedPctPG(v As Double) As String
    If v >= 0 Then
        FormatSignedPctPG = "+" & Format(v, "0.0") & "%"
    Else
        FormatSignedPctPG = Format(v, "0.0") & "%"
    End If
End Function

' ============================================================
' REPORT LOOKUP HELPERS
' ============================================================
Private Function FindReportRowByTickerPG(ws As Worksheet, dataStart As Long, dataEnd As Long, _
                                         tickerText As String) As Long
    Dim r As Long
    Dim tk As String
    tk = UCase(Trim(tickerText))

    For r = dataStart To dataEnd
        If UCase(Trim(CStr(ws.Cells(r, 3).Value))) = tk Then
            FindReportRowByTickerPG = r
            Exit Function
        End If
    Next r
End Function

Private Function FindProcessedReportSheetPG(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If UCase(Trim(CStr(ws.Cells(2, 1).Value))) = "ASSET CLASS" Then
            Set FindProcessedReportSheetPG = ws
            Exit Function
        End If
    Next ws
End Function

Private Function GetReportStatePG(ws As Worksheet, ByRef headerRow As Long, ByRef dataStart As Long, _
                                  ByRef dataEnd As Long, ByRef totalRow As Long) As Boolean
    headerRow = FindHeaderRowPG(ws)
    If headerRow = 0 Then Exit Function

    dataStart = headerRow + 1
    totalRow = FindTotalRowPG(ws, dataStart)
    If totalRow = 0 Then Exit Function

    dataEnd = totalRow - 1
    GetReportStatePG = True
End Function

Private Function FindHeaderRowPG(ws As Worksheet) As Long
    Dim r As Long
    For r = 1 To 10
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRowPG = r
            Exit Function
        End If
    Next r
End Function

Private Function FindTotalRowPG(ws As Worksheet, dataStart As Long) As Long
    Dim r As Long
    For r = dataStart To dataStart + 500
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And Trim(CStr(ws.Cells(r, 5).Value)) <> "" Then
            FindTotalRowPG = r
            Exit Function
        End If
    Next r
End Function

Private Function FindCellExactPG(ws As Worksheet, textValue As String) As Range
    On Error Resume Next
    Set FindCellExactPG = ws.Cells.Find(What:=textValue, _
                                        LookIn:=xlValues, _
                                        LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, _
                                        SearchDirection:=xlNext, _
                                        MatchCase:=False)
    On Error GoTo 0
End Function

Private Function SafeDoublePG(v As Variant) As Double
    If IsNumeric(v) Then
        SafeDoublePG = CDbl(v)
    Else
        SafeDoublePG = 0
    End If
End Function

' IsNumeric(Empty) returns True in VBA (Empty coerces to 0), which would
' silently treat a blank Live Price cell as a real $0 quote and produce a
' bogus -100% drift. Guard explicitly against Empty/Null before trusting
' IsNumeric.
Private Function IsNumericValuePG(v As Variant) As Boolean
    IsNumericValuePG = Not IsEmpty(v) And Not IsNull(v) And IsNumeric(v)
End Function
