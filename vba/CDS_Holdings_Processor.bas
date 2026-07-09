Attribute VB_Name = "CDS_Holdings_Processor"
Option Explicit

' ============================================================
' CDS Holdings Processor v12.0
'
' v12 changes vs v11:
'   - Unknown tickers are written to an on-sheet review box.
'   - User can classify unknowns from dropdowns.
'   - SaveUnknownsAndRefresh can push classifications to CDS_Settings.
'   - Old popup now points user to the review workflow.
'
' Prior v11 behavior retained:
'   - Ticker classification via CDS_Settings
'   - Freeze panes location via CDS_Settings ("FreezePanesAt")
'   - Column K Quantity hidden by default
'   - PERSONAL.XLSB safety guard
'
' USAGE:
'   1. Open raw CDS Holdings CSV.
'   2. Alt+F8 > ProcessCDSHoldings > Run.
'   3. If unknown tickers appear, classify them in review box.
'   4. Alt+F8 > SaveUnknownsAndRefresh > Run.
'   5. Then run scenario macros if needed.
' ============================================================

Sub ProcessCDSHoldings()
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    ' --- SAFETY GUARD: don't process PERSONAL.XLSB itself ---
    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Open the client's CSV first, then run.", vbCritical
        GoTo Done
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dollarFmt As String
    dollarFmt = "$#,##0;($#,##0)"

    ' --- ALL DECLARATIONS ---
    Dim headerRow As Long, i As Long, r As Long, j As Long
    Dim tk As String, tempStr As String, tempLong As Long
    Dim acctName As String, acctNum As String
    Dim scanEnd As Long, col As Long, lr As Long
    Dim holdCount As Long, idx As Long
    Dim hDesc() As String, hTicker() As String
    Dim hFMV() As Double, hGL() As Double, hBasis() As Double
    Dim hIncome() As Double, hYield() As Double, hQty() As String
    Dim hClass() As String, unknowns As String
    Dim unkTicker() As String, unkDesc() As String, unkCount As Long
    Dim mainIdx() As Long, shortIdx() As Long
    Dim mainCount As Long, shortCount As Long
    Dim mi As Long, si As Long
    Dim sortKeys() As String
    Dim dataStart As Long, dataEnd As Long, totRow As Long
    Dim shortStart As Long, shortEnd As Long
    Dim nextRow As Long
    Dim c As Variant
    Dim bRng As Range
    Dim allCols As Variant

    allCols = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K")

    ' --- FIND HEADER ---
    headerRow = 0

    For i = 1 To 10
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = "accountname" Or _
           LCase(Trim(CStr(ws.Cells(i, 3).Value))) = "description" Then
            headerRow = i
            Exit For
        End If
    Next i

    If headerRow = 0 Then
        MsgBox "Header row not found. Is this a raw CDS Holdings CSV?", vbExclamation
        GoTo Done
    End If

    acctName = Trim(CStr(ws.Cells(headerRow + 1, 1).Value))
    acctNum = Trim(CStr(ws.Cells(headerRow + 1, 2).Value))

    ' --- SCAN HOLDINGS ---
    scanEnd = 1

    For col = 1 To 12
        lr = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If lr > scanEnd Then scanEnd = lr
    Next col

    holdCount = 0

    For i = headerRow + 1 To scanEnd
        If Trim(CStr(ws.Cells(i, 3).Value)) <> "" Then
            holdCount = holdCount + 1
        End If
    Next i

    If holdCount = 0 Then
        MsgBox "No holdings found.", vbExclamation
        GoTo Done
    End If

    ReDim hDesc(1 To holdCount)
    ReDim hTicker(1 To holdCount)
    ReDim hFMV(1 To holdCount)
    ReDim hGL(1 To holdCount)
    ReDim hBasis(1 To holdCount)
    ReDim hIncome(1 To holdCount)
    ReDim hYield(1 To holdCount)
    ReDim hQty(1 To holdCount)

    idx = 0

    For i = headerRow + 1 To scanEnd
        If Trim(CStr(ws.Cells(i, 3).Value)) <> "" Then
            idx = idx + 1

            hDesc(idx) = Trim(CStr(ws.Cells(i, 3).Value))
            hTicker(idx) = Trim(CStr(ws.Cells(i, 4).Value))
            hFMV(idx) = SafeDouble(ws.Cells(i, 6).Value)
            hGL(idx) = SafeDouble(ws.Cells(i, 7).Value)
            hBasis(idx) = SafeDouble(ws.Cells(i, 9).Value)
            hIncome(idx) = SafeDouble(ws.Cells(i, 10).Value)
            hYield(idx) = SafeDouble(ws.Cells(i, 11).Value)
            hQty(idx) = Trim(CStr(ws.Cells(i, 12).Value))
        End If
    Next i

    ' --- ASSIGN CLASSES via CDS_Settings ---
    ReDim hClass(1 To holdCount)
    ReDim unkTicker(1 To holdCount)
    ReDim unkDesc(1 To holdCount)

    unknowns = ""
    unkCount = 0

    Dim cls As String

    For i = 1 To holdCount
        cls = ClassifyTickerWithFallback(hTicker(i))

        If cls <> "" Then
            hClass(i) = cls
        Else
            hClass(i) = "???"

            unknowns = unknowns & UCase(Trim(hTicker(i))) & " (" & Left(hDesc(i), 40) & ")" & vbCrLf

            unkCount = unkCount + 1
            unkTicker(unkCount) = UCase(Trim(hTicker(i)))
            unkDesc(unkCount) = hDesc(i)
        End If
    Next i

    ' --- SEPARATE INTO MAIN + SHORT ---
    mainCount = 0
    shortCount = 0

    For i = 1 To holdCount
        If hClass(i) = "SHORT" Then
            shortCount = shortCount + 1
        Else
            mainCount = mainCount + 1
        End If
    Next i

    If mainCount > 0 Then ReDim mainIdx(1 To mainCount)
    If shortCount > 0 Then ReDim shortIdx(1 To shortCount)

    mi = 0
    si = 0

    For i = 1 To holdCount
        If hClass(i) = "SHORT" Then
            si = si + 1
            shortIdx(si) = i
        Else
            mi = mi + 1
            mainIdx(mi) = i
        End If
    Next i

    ' --- PRE-SORT MAIN: CASH first, then A-Z ---
    If mainCount > 0 Then
        ReDim sortKeys(1 To mainCount)

        For i = 1 To mainCount
            If hClass(mainIdx(i)) = "CASH" Then
                sortKeys(i) = "!" & hClass(mainIdx(i))
            Else
                sortKeys(i) = hClass(mainIdx(i))
            End If
        Next i

        For i = 1 To mainCount - 1
            For j = i + 1 To mainCount
                If sortKeys(i) > sortKeys(j) Then
                    tempStr = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempStr

                    tempLong = mainIdx(i)
                    mainIdx(i) = mainIdx(j)
                    mainIdx(j) = tempLong
                End If
            Next j
        Next i
    End If

    ' ============================================================
    ' OUTPUT SECTION
    ' ============================================================

    ws.Cells.Clear

    ' --- Column widths ---
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 55
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 30
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").Hidden = True
    ws.Columns("L").ColumnWidth = 4
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 13
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10

    ' --- Row 1: BLACK fill, WHITE bold text ---
    ws.Cells(1, 1).Value = acctNum
    ws.Cells(1, 2).Value = acctName

    For Each c In allCols
        With ws.Range(c & "1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 0, 0)
            .Font.Color = RGB(255, 255, 255)
        End With
    Next c

    ' --- Row 2: Headers ---
    Dim hdrs As Variant
    hdrs = Array("ASSET CLASS", "Description", "TICKER", "%", "FMV", "G/L", "% G/L", "COST BASIS", "INCOME", "YIELD", "Quantity")

    For i = 0 To 10
        With ws.Cells(2, i + 1)
            .Value = hdrs(i)
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    Next i

    ' --- WRITE MAIN HOLDINGS ---
    r = 3

    For i = 1 To mainCount
        idx = mainIdx(i)

        ws.Cells(r, 1).Value = hClass(idx)
        ws.Cells(r, 2).Value = hDesc(idx)
        ws.Cells(r, 3).Value = hTicker(idx)

        ws.Cells(r, 5).Value = hFMV(idx)
        ws.Cells(r, 5).NumberFormat = dollarFmt

        If hGL(idx) <> 0 Then
            ws.Cells(r, 6).Value = hGL(idx)
            ws.Cells(r, 6).NumberFormat = dollarFmt
        End If

        ws.Cells(r, 8).Value = hBasis(idx)
        ws.Cells(r, 8).NumberFormat = dollarFmt

        If hIncome(idx) <> 0 Then
            ws.Cells(r, 9).Value = hIncome(idx)
            ws.Cells(r, 9).NumberFormat = dollarFmt
        End If

        If hYield(idx) <> 0 Then
            ws.Cells(r, 10).Value = hYield(idx) / 100
            ws.Cells(r, 10).NumberFormat = "0.00%"
        End If

        ws.Cells(r, 11).Value = CleanQty(hQty(idx))

        If IsNumeric(ws.Cells(r, 11).Value) Then
            If ws.Cells(r, 11).Value <> Int(ws.Cells(r, 11).Value) Then
                ws.Cells(r, 11).NumberFormat = "#,##0.00"
            Else
                ws.Cells(r, 11).NumberFormat = "#,##0"
            End If
        End If

        For Each c In allCols
            With ws.Range(c & r)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        Next c

        r = r + 1
    Next i

    dataStart = 3
    dataEnd = dataStart + mainCount - 1
    totRow = dataEnd + 1

    ' --- % and %G/L formulas ---
    For i = dataStart To dataEnd
        ws.Cells(i, 4).Formula = "=E" & i & "/$E$" & totRow
        ws.Cells(i, 4).NumberFormat = "0%"

        ws.Cells(i, 7).Formula = "=IF(E" & i & "=0,0,F" & i & "/E" & i & ")"
        ws.Cells(i, 7).NumberFormat = "0%"
    Next i

    ' --- Total row ---
    ws.Cells(totRow, 4).Formula = "=IF(E" & totRow & "=0,0,1)"
    ws.Cells(totRow, 4).NumberFormat = "0%"

    ws.Cells(totRow, 5).Formula = "=SUM(E" & dataStart & ":E" & dataEnd & ")"
    ws.Cells(totRow, 5).NumberFormat = dollarFmt

    ws.Cells(totRow, 6).Formula = "=SUM(F" & dataStart & ":F" & dataEnd & ")"
    ws.Cells(totRow, 6).NumberFormat = dollarFmt

    ws.Cells(totRow, 7).Formula = "=IF(E" & totRow & "=0,0,F" & totRow & "/E" & totRow & ")"
    ws.Cells(totRow, 7).NumberFormat = "0%"

    ws.Cells(totRow, 8).Formula = "=SUM(H" & dataStart & ":H" & dataEnd & ")"
    ws.Cells(totRow, 8).NumberFormat = dollarFmt

    ws.Cells(totRow, 9).Formula = "=SUM(I" & dataStart & ":I" & dataEnd & ")"
    ws.Cells(totRow, 9).NumberFormat = dollarFmt

    ws.Cells(totRow, 10).Formula = "=IF(E" & totRow & "=0,0,I" & totRow & "/E" & totRow & ")"
    ws.Cells(totRow, 10).NumberFormat = "0.00%"

    For Each c In allCols
        With ws.Range(c & totRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    Next c

    Set bRng = ws.Range("A1:K" & totRow)

    With bRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    bRng.Borders(xlInsideVertical).Weight = xlThin
    bRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    bRng.Borders(xlInsideHorizontal).Weight = xlThin

    ' --- SHORT BLOCK ---
    nextRow = totRow + 2

    If shortCount > 0 Then
        shortStart = nextRow

        For i = 1 To shortCount
            idx = shortIdx(i)

            ws.Cells(nextRow, 1).Value = hClass(idx)
            ws.Cells(nextRow, 2).Value = hDesc(idx)
            ws.Cells(nextRow, 3).Value = hTicker(idx)
            ws.Cells(nextRow, 5).Value = hFMV(idx)
            ws.Cells(nextRow, 5).NumberFormat = dollarFmt

            For Each c In Array("A", "B", "C", "E")
                With ws.Range(c & nextRow)
                    .Interior.Color = RGB(198, 239, 206)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                End With
            Next c

            nextRow = nextRow + 1
        Next i

        shortEnd = nextRow - 1

        Set bRng = ws.Range("A" & shortStart & ":E" & shortEnd)

        With bRng.Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With

        bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        bRng.Borders(xlInsideVertical).Weight = xlThin

        nextRow = shortEnd + 2
    End If

    ' --- TAX SECTION ---
    Dim taxStart As Long
    Dim taxEnd As Long
    Dim currentYear As Long
    Dim taxLabels As Variant

    taxStart = nextRow
    currentYear = Year(Date)

    taxLabels = Array("MONTHLY DISTRIBUTION", _
                      currentYear & " REALIZED GAINS", _
                      currentYear & " TAXABLE DIVIDENDS", _
                      currentYear & " NON-TAXABLE DIVIDENDS")

    For i = 0 To UBound(taxLabels)
        r = taxStart + i

        ws.Cells(r, 9).Value = taxLabels(i)
        ws.Cells(r, 9).HorizontalAlignment = xlRight
        ws.Cells(r, 9).Font.Bold = True
        ws.Cells(r, 9).Interior.Color = RGB(252, 195, 185)
        ws.Cells(r, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(r, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Cells(r, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Cells(r, 9).Borders(xlEdgeRight).LineStyle = xlContinuous

        ws.Cells(r, 10).Value = 0
        ws.Cells(r, 10).NumberFormat = dollarFmt
        ws.Cells(r, 10).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(r, 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Cells(r, 10).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Cells(r, 10).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next i

    taxEnd = taxStart + 3

    Set bRng = ws.Range("I" & taxStart & ":J" & taxEnd)

    With bRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    bRng.Borders(xlInsideVertical).Weight = xlThin
    bRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    bRng.Borders(xlInsideHorizontal).Weight = xlThin

    Application.Calculate

    ' --- PIVOT TABLE at M2 ---
    Dim srcRange As Range
    Set srcRange = ws.Range("A2:D" & dataEnd)

    Dim pvtCache As PivotCache
    Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)

    Dim pvt As PivotTable
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=ws.Range("M2"), _
        TableName:="AllocPivot")

    With pvt.PivotFields("ASSET CLASS")
        .Orientation = xlRowField
        .Position = 1
    End With

    With pvt.PivotFields("%")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "0%"
        .Name = "Sum of %"
    End With

    pvt.RowAxisLayout xlTabularRow
    pvt.ShowTableStyleRowStripes = False
    pvt.TableStyle2 = ""

    Dim pvtRange As Range
    Set pvtRange = pvt.TableRange1

    Dim pvtHeaderRow As Long
    Dim pvtLastRow As Long

    pvtHeaderRow = pvtRange.Row
    pvtLastRow = pvtRange.Row + pvtRange.Rows.Count - 1

    ws.Range("M" & pvtHeaderRow & ":N" & pvtHeaderRow).Interior.Color = RGB(0, 176, 240)
    ws.Range("M" & pvtHeaderRow & ":N" & pvtHeaderRow).Font.Bold = True

    Dim pr As Long

    For pr = pvtHeaderRow + 1 To pvtLastRow - 1
        ws.Range("M" & pr & ":N" & pr).Interior.Color = RGB(221, 235, 247)
    Next pr

    ws.Range("M" & pvtLastRow & ":N" & pvtLastRow).Interior.Color = RGB(0, 176, 80)
    ws.Range("M" & pvtLastRow & ":N" & pvtLastRow).Font.Bold = True
    ws.Range("M" & pvtLastRow & ":N" & pvtLastRow).Font.Color = RGB(255, 255, 255)

    Set bRng = ws.Range("M" & pvtHeaderRow & ":N" & pvtLastRow)

    With bRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    bRng.Borders(xlInsideVertical).Weight = xlThin
    bRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    bRng.Borders(xlInsideHorizontal).Weight = xlThin

    ' --- UNKNOWN TICKER REVIEW BOX ---
    If unkCount > 0 Then
        WriteUnknownTickerReviewBox ws, unkTicker, unkDesc, unkCount
    End If

    ' --- CONTEXT TRACKING HELPER CELL ---
    With ws.Cells(1, 12)
        .Value = dataStart
        .Font.Color = RGB(255, 255, 255)
    End With

    ' --- FREEZE PANES ---
    Dim freezeAt As String
    freezeAt = GetSetting("FreezePanesAt", "D3")

    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.FreezePanes = False
    ws.Range(freezeAt).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select

    ' --- PRINT SETUP ---
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .PrintArea = "A1:N" & taxEnd
        .CenterHorizontally = True
    End With

    ' --- ALERT UNKNOWNS ---
    If unknowns <> "" Then
        MsgBox "Unknown tickers were added to the review box below the allocation pivot." & vbCrLf & vbCrLf & _
               "Choose asset classes from the dropdowns, then run:" & vbCrLf & _
               "Alt+F8 > SaveUnknownsAndRefresh", _
               vbInformation, "New Tickers Found"
    End If

    GoTo Done

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation

Done:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
End Sub

Private Function SafeDouble(v As Variant) As Double
    If IsEmpty(v) Or v = "" Or v = "-" Then
        SafeDouble = 0
        Exit Function
    End If

    Dim s As String
    s = CStr(v)

    s = Replace(s, ",", "")
    s = Replace(s, "$", "")
    s = Replace(s, "(", "-")
    s = Replace(s, ")", "")
    s = Trim(s)

    If IsNumeric(s) Then
        SafeDouble = CDbl(s)
    Else
        SafeDouble = 0
    End If
End Function

Private Function CleanQty(v As String) As Variant
    If v = "" Then
        CleanQty = ""
        Exit Function
    End If

    Dim s As String
    s = Replace(v, ",", "")
    s = Replace(s, """", "")
    s = Trim(s)

    If IsNumeric(s) Then
        CleanQty = CDbl(s)
    Else
        CleanQty = v
    End If
End Function


Public Sub ProcessCDSHoldings_Lite()
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    ' --- SAFETY GUARD: don't process PERSONAL.XLSB itself ---
    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Open the client's CSV first, then run.", vbCritical
        GoTo Done
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dollarFmt As String
    dollarFmt = "$#,##0;($#,##0)"

    ' --- ALL DECLARATIONS ---
    Dim headerRow As Long, i As Long, r As Long, j As Long
    Dim tempStr As String, tempLong As Long
    Dim acctName As String, acctNum As String
    Dim scanEnd As Long, col As Long, lr As Long
    Dim holdCount As Long, idx As Long
    Dim hDesc() As String, hTicker() As String
    Dim hFMV() As Double, hGL() As Double, hBasis() As Double
    Dim hIncome() As Double, hYield() As Double, hQty() As String
    Dim hClass() As String, unknowns As String
    Dim unknownCount As Long
    Dim mainIdx() As Long, shortIdx() As Long
    Dim mainCount As Long, shortCount As Long
    Dim mi As Long, si As Long
    Dim sortKeys() As String
    Dim dataStart As Long, dataEnd As Long, totRow As Long
    Dim shortStart As Long, shortEnd As Long
    Dim nextRow As Long
    Dim c As Variant
    Dim bRng As Range
    Dim allCols As Variant

    allCols = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K")

    ' --- FIND HEADER ---
    headerRow = 0

    For i = 1 To 10
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = "accountname" Or _
           LCase(Trim(CStr(ws.Cells(i, 3).Value))) = "description" Then
            headerRow = i
            Exit For
        End If
    Next i

    If headerRow = 0 Then
        MsgBox "Header row not found. Is this a raw CDS Holdings CSV?", vbExclamation
        GoTo Done
    End If

    acctName = Trim(CStr(ws.Cells(headerRow + 1, 1).Value))
    acctNum = Trim(CStr(ws.Cells(headerRow + 1, 2).Value))

    ' --- SCAN HOLDINGS ---
    scanEnd = 1

    For col = 1 To 12
        lr = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If lr > scanEnd Then scanEnd = lr
    Next col

    holdCount = 0

    For i = headerRow + 1 To scanEnd
        If Trim(CStr(ws.Cells(i, 3).Value)) <> "" Then
            holdCount = holdCount + 1
        End If
    Next i

    If holdCount = 0 Then
        MsgBox "No holdings found.", vbExclamation
        GoTo Done
    End If

    ReDim hDesc(1 To holdCount)
    ReDim hTicker(1 To holdCount)
    ReDim hFMV(1 To holdCount)
    ReDim hGL(1 To holdCount)
    ReDim hBasis(1 To holdCount)
    ReDim hIncome(1 To holdCount)
    ReDim hYield(1 To holdCount)
    ReDim hQty(1 To holdCount)
    ReDim hClass(1 To holdCount)

    idx = 0

    For i = headerRow + 1 To scanEnd
        If Trim(CStr(ws.Cells(i, 3).Value)) <> "" Then
            idx = idx + 1

            hDesc(idx) = Trim(CStr(ws.Cells(i, 3).Value))
            hTicker(idx) = Trim(CStr(ws.Cells(i, 4).Value))
            hFMV(idx) = SafeDouble(ws.Cells(i, 6).Value)
            hGL(idx) = SafeDouble(ws.Cells(i, 7).Value)
            hBasis(idx) = SafeDouble(ws.Cells(i, 9).Value)
            hIncome(idx) = SafeDouble(ws.Cells(i, 10).Value)
            hYield(idx) = SafeDouble(ws.Cells(i, 11).Value)
            hQty(idx) = Trim(CStr(ws.Cells(i, 12).Value))
        End If
    Next i

    ' --- ASSIGN CLASSES using ORIGINAL hidden/silent classifier ---
    unknowns = ""
    unknownCount = 0

    Dim cls As String

    For i = 1 To holdCount
        cls = ClassifyTickerWithFallback(hTicker(i))

        If cls <> "" Then
            hClass(i) = cls
        Else
            hClass(i) = "???"
            unknownCount = unknownCount + 1
            unknowns = unknowns & UCase(Trim(hTicker(i))) & " (" & Left(hDesc(i), 40) & ")" & vbCrLf
        End If
    Next i

    ' --- SEPARATE INTO MAIN + SHORT ---
    mainCount = 0
    shortCount = 0

    For i = 1 To holdCount
        If hClass(i) = "SHORT" Then
            shortCount = shortCount + 1
        Else
            mainCount = mainCount + 1
        End If
    Next i

    If mainCount > 0 Then ReDim mainIdx(1 To mainCount)
    If shortCount > 0 Then ReDim shortIdx(1 To shortCount)

    mi = 0
    si = 0

    For i = 1 To holdCount
        If hClass(i) = "SHORT" Then
            si = si + 1
            shortIdx(si) = i
        Else
            mi = mi + 1
            mainIdx(mi) = i
        End If
    Next i

    ' --- PRE-SORT MAIN: CASH first, BOND second, ??? third, then A-Z ---
    If mainCount > 0 Then
        ReDim sortKeys(1 To mainCount)

        For i = 1 To mainCount
            Select Case UCase(Trim(hClass(mainIdx(i))))
                Case "CASH"
                    sortKeys(i) = "000_CASH"
                Case "BOND"
                    sortKeys(i) = "001_BOND"
                Case "???"
                    sortKeys(i) = "002_UNKNOWN"
                Case Else
                    sortKeys(i) = "100_" & UCase(Trim(hClass(mainIdx(i))))
            End Select
        Next i

        For i = 1 To mainCount - 1
            For j = i + 1 To mainCount
                If sortKeys(i) > sortKeys(j) Then
                    tempStr = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempStr

                    tempLong = mainIdx(i)
                    mainIdx(i) = mainIdx(j)
                    mainIdx(j) = tempLong
                End If
            Next j
        Next i
    End If

    ' ============================================================
    ' OUTPUT SECTION
    ' ============================================================

    ws.Cells.Clear

    ' --- Column widths ---
    ws.Columns("A").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 55
    ws.Columns("C").ColumnWidth = 9
    ws.Columns("D").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 9
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 30
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").Hidden = True
    ws.Columns("L").ColumnWidth = 4
    ws.Columns("M").ColumnWidth = 14
    ws.Columns("N").ColumnWidth = 13
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10

    ' --- Row 1: BLACK fill, WHITE bold text ---
    ws.Cells(1, 1).Value = acctNum
    ws.Cells(1, 2).Value = acctName

    For Each c In allCols
        With ws.Range(c & "1")
            .Font.Bold = True
            .Interior.Color = RGB(0, 0, 0)
            .Font.Color = RGB(255, 255, 255)
        End With
    Next c

    ' --- Row 2: Headers ---
    Dim hdrs As Variant
    hdrs = Array("ASSET CLASS", "Description", "TICKER", "%", "FMV", "G/L", "% G/L", "COST BASIS", "INCOME", "YIELD", "Quantity")

    For i = 0 To 10
        With ws.Cells(2, i + 1)
            .Value = hdrs(i)
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    Next i

    ' --- WRITE MAIN HOLDINGS ---
    r = 3

    For i = 1 To mainCount
        idx = mainIdx(i)

        ws.Cells(r, 1).Value = hClass(idx)
        ws.Cells(r, 2).Value = hDesc(idx)
        ws.Cells(r, 3).Value = hTicker(idx)

        ws.Cells(r, 5).Value = hFMV(idx)
        ws.Cells(r, 5).NumberFormat = dollarFmt

        If hGL(idx) <> 0 Then
            ws.Cells(r, 6).Value = hGL(idx)
            ws.Cells(r, 6).NumberFormat = dollarFmt
        End If

        ws.Cells(r, 8).Value = hBasis(idx)
        ws.Cells(r, 8).NumberFormat = dollarFmt

        If hIncome(idx) <> 0 Then
            ws.Cells(r, 9).Value = hIncome(idx)
            ws.Cells(r, 9).NumberFormat = dollarFmt
        End If

        If hYield(idx) <> 0 Then
            ws.Cells(r, 10).Value = hYield(idx) / 100
            ws.Cells(r, 10).NumberFormat = "0.00%"
        End If

        ws.Cells(r, 11).Value = CleanQty(hQty(idx))

        If IsNumeric(ws.Cells(r, 11).Value) Then
            If ws.Cells(r, 11).Value <> Int(ws.Cells(r, 11).Value) Then
                ws.Cells(r, 11).NumberFormat = "#,##0.00"
            Else
                ws.Cells(r, 11).NumberFormat = "#,##0"
            End If
        End If

        For Each c In allCols
            With ws.Range(c & r)
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End With
        Next c

        r = r + 1
    Next i

    dataStart = 3
    dataEnd = dataStart + mainCount - 1
    totRow = dataEnd + 1

    ' --- % and %G/L formulas ---
    For i = dataStart To dataEnd
        ws.Cells(i, 4).Formula = "=E" & i & "/$E$" & totRow
        ws.Cells(i, 4).NumberFormat = "0%"

        ws.Cells(i, 7).Formula = "=IF(E" & i & "=0,0,F" & i & "/E" & i & ")"
        ws.Cells(i, 7).NumberFormat = "0%"
    Next i

    ' --- Total row ---
    ws.Cells(totRow, 4).Formula = "=IF(E" & totRow & "=0,0,1)"
    ws.Cells(totRow, 4).NumberFormat = "0%"

    ws.Cells(totRow, 5).Formula = "=SUM(E" & dataStart & ":E" & dataEnd & ")"
    ws.Cells(totRow, 5).NumberFormat = dollarFmt

    ws.Cells(totRow, 6).Formula = "=SUM(F" & dataStart & ":F" & dataEnd & ")"
    ws.Cells(totRow, 6).NumberFormat = dollarFmt

    ws.Cells(totRow, 7).Formula = "=IF(E" & totRow & "=0,0,F" & totRow & "/E" & totRow & ")"
    ws.Cells(totRow, 7).NumberFormat = "0%"

    ws.Cells(totRow, 8).Formula = "=SUM(H" & dataStart & ":H" & dataEnd & ")"
    ws.Cells(totRow, 8).NumberFormat = dollarFmt

    ws.Cells(totRow, 9).Formula = "=SUM(I" & dataStart & ":I" & dataEnd & ")"
    ws.Cells(totRow, 9).NumberFormat = dollarFmt

    ws.Cells(totRow, 10).Formula = "=IF(E" & totRow & "=0,0,I" & totRow & "/E" & totRow & ")"
    ws.Cells(totRow, 10).NumberFormat = "0.00%"

    For Each c In allCols
        With ws.Range(c & totRow)
            .Font.Bold = True
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    Next c

    Set bRng = ws.Range("A1:K" & totRow)

    With bRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    bRng.Borders(xlInsideVertical).Weight = xlThin
    bRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    bRng.Borders(xlInsideHorizontal).Weight = xlThin

    ' --- SHORT BLOCK ---
    nextRow = totRow + 2

    If shortCount > 0 Then
        shortStart = nextRow

        For i = 1 To shortCount
            idx = shortIdx(i)

            ws.Cells(nextRow, 1).Value = hClass(idx)
            ws.Cells(nextRow, 2).Value = hDesc(idx)
            ws.Cells(nextRow, 3).Value = hTicker(idx)
            ws.Cells(nextRow, 5).Value = hFMV(idx)
            ws.Cells(nextRow, 5).NumberFormat = dollarFmt

            For Each c In Array("A", "B", "C", "E")
                With ws.Range(c & nextRow)
                    .Interior.Color = RGB(198, 239, 206)
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                End With
            Next c

            nextRow = nextRow + 1
        Next i

        shortEnd = nextRow - 1

        Set bRng = ws.Range("A" & shortStart & ":E" & shortEnd)

        With bRng.Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With

        bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        bRng.Borders(xlInsideVertical).Weight = xlThin

        nextRow = shortEnd + 2
    End If

    ' --- TAX SECTION ---
    Dim taxStart As Long
    Dim taxEnd As Long
    Dim currentYear As Long
    Dim taxLabels As Variant

    taxStart = nextRow
    currentYear = Year(Date)

    taxLabels = Array("MONTHLY DISTRIBUTION", _
                      currentYear & " REALIZED GAINS", _
                      currentYear & " TAXABLE DIVIDENDS", _
                      currentYear & " NON-TAXABLE DIVIDENDS")

    For i = 0 To UBound(taxLabels)
        r = taxStart + i

        ws.Cells(r, 9).Value = taxLabels(i)
        ws.Cells(r, 9).HorizontalAlignment = xlRight
        ws.Cells(r, 9).Font.Bold = True
        ws.Cells(r, 9).Interior.Color = RGB(252, 195, 185)
        ws.Cells(r, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(r, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Cells(r, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Cells(r, 9).Borders(xlEdgeRight).LineStyle = xlContinuous

        ws.Cells(r, 10).Value = 0
        ws.Cells(r, 10).NumberFormat = dollarFmt
        ws.Cells(r, 10).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(r, 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Cells(r, 10).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Cells(r, 10).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next i

    taxEnd = taxStart + 3

    Set bRng = ws.Range("I" & taxStart & ":J" & taxEnd)

    With bRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    bRng.Borders(xlInsideVertical).Weight = xlThin
    bRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    bRng.Borders(xlInsideHorizontal).Weight = xlThin

    Application.Calculate

    ' --- PIVOT TABLE at M2 ---
    If mainCount > 0 Then
        Dim srcRange As Range
        Set srcRange = ws.Range("A2:D" & dataEnd)

        Dim pvtCache As PivotCache
        Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=srcRange)

        Dim pvt As PivotTable
        Dim pvtName As String
        pvtName = "AllocPivotLite_" & Format(Now, "hhmmss")

        Set pvt = pvtCache.CreatePivotTable( _
            TableDestination:=ws.Range("M2"), _
            TableName:=pvtName)

        With pvt.PivotFields("ASSET CLASS")
            .Orientation = xlRowField
            .Position = 1
        End With

        With pvt.PivotFields("%")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "0%"
            .Name = "Sum of %"
        End With

        pvt.RowAxisLayout xlTabularRow
        pvt.ShowTableStyleRowStripes = False
        pvt.TableStyle2 = ""

        ' --- FORCE PIVOT ORDER ---
        ' Desired top order:
        '   CASH
        '   BOND
        '   ???
        '
        ' Move in reverse order because the last item sent to Position 1 becomes top.
        On Error Resume Next
        With pvt.PivotFields("ASSET CLASS")
            .AutoSort xlManual, "ASSET CLASS"
            .PivotItems("???").Position = 1
            .PivotItems("BOND").Position = 1
            .PivotItems("CASH").Position = 1
        End With
        On Error GoTo ErrHandler

        Dim pvtRange As Range
        Set pvtRange = pvt.TableRange1

        Dim pvtHeaderRow As Long
        Dim pvtLastRow As Long
        Dim pr As Long

        pvtHeaderRow = pvtRange.Row
        pvtLastRow = pvtRange.Row + pvtRange.Rows.Count - 1

        ws.Range("M" & pvtHeaderRow & ":N" & pvtHeaderRow).Interior.Color = RGB(0, 176, 240)
        ws.Range("M" & pvtHeaderRow & ":N" & pvtHeaderRow).Font.Bold = True

        For pr = pvtHeaderRow + 1 To pvtLastRow - 1
            ws.Range("M" & pr & ":N" & pr).Interior.Color = RGB(221, 235, 247)
        Next pr

        ws.Range("M" & pvtLastRow & ":N" & pvtLastRow).Interior.Color = RGB(0, 176, 80)
        ws.Range("M" & pvtLastRow & ":N" & pvtLastRow).Font.Bold = True
        ws.Range("M" & pvtLastRow & ":N" & pvtLastRow).Font.Color = RGB(255, 255, 255)

        Set bRng = ws.Range("M" & pvtHeaderRow & ":N" & pvtLastRow)

        With bRng.Borders
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With

        bRng.Borders(xlInsideVertical).LineStyle = xlContinuous
        bRng.Borders(xlInsideVertical).Weight = xlThin
        bRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        bRng.Borders(xlInsideHorizontal).Weight = xlThin
    End If

    ' --- CONTEXT TRACKING HELPER CELL ---
    With ws.Cells(1, 12)
        .Value = dataStart
        .Font.Color = RGB(255, 255, 255)
    End With

    ' --- NO FREEZE PANES / NO SPLIT WINDOW ---
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 0
    End With
    ws.Range("A1").Select

    ' --- PRINT SETUP ---
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .PrintArea = "A1:N" & taxEnd
        .CenterHorizontally = True
    End With

    ' --- ALERT UNKNOWNS WITHOUT CREATING REVIEW TABLE ---
    If unknownCount > 0 Then
        MsgBox unknownCount & " holding(s) were marked as ??? in Column A." & vbCrLf & vbCrLf & _
               "Manually overwrite ??? directly in Column A as needed." & vbCrLf & _
               "Then right-click the allocation pivot and choose Refresh.", _
               vbInformation, "Unknown Asset Classes"
    End If

    GoTo Done

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation

Done:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
End Sub

