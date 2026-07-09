Attribute VB_Name = "CDS_Unknowns"
Option Explicit

' ============================================================
' CDS Unknown Ticker Review v1.1
'
' Fix vs v1.0:
'   - Sort helper moved from column O to column L.
'   - Sort range is now A:L, avoiding M:N allocation pivot.
'   - Prevents Excel 1004 merged-cell / same-size sort error.
'   - Clears any leftover helper values from O if prior failed run
'     wrote sort keys there.
'
' Workflow:
'   1. ProcessCDSHoldings writes unknown tickers to a review box.
'   2. User selects asset class from dropdowns.
'   3. SaveUnknownsAndRefresh writes choices to CDS_Settings.
'   4. Report reclassifies, sorts, recalculates, and refreshes pivot.
'
' Important:
'   - Run SaveUnknownsAndRefresh BEFORE adding raise-cash scenarios.
'   - If scenarios already exist, this macro refuses to run because
'     re-sorting holdings would desync scenario columns.
' ============================================================

Private Const UNKNOWN_TITLE As String = "UNKNOWN TICKERS - ASSIGN ASSET CLASS"
Private Const UNKNOWN_START_COL As Long = 13 ' M
Private Const UNKNOWN_WIDTH As Long = 3

Private Const VALID_CLASSES As String = "CASH,SHORT,BOND,GLOBAL,INTL,LARGE,MIXED,MVOL,REIT,SECT,SMID,STOCK"

' ============================================================
' PUBLIC: Called by ProcessCDSHoldings
' ============================================================

Public Sub WriteUnknownTickerReviewBox(ws As Worksheet, unknownTickers As Variant, unknownDescs As Variant, unknownCount As Long)
    On Error GoTo ErrHandler

    If unknownCount <= 0 Then Exit Sub

    Dim startRow As Long
    startRow = FindUnknownBoxStartRow(ws)

    ' Clear prior review box area.
    ws.Range(ws.Cells(startRow, UNKNOWN_START_COL), _
             ws.Cells(startRow + Application.Max(50, unknownCount + 5), UNKNOWN_START_COL + UNKNOWN_WIDTH - 1)).Clear

    ' Title row.
    On Error Resume Next
    ws.Range(ws.Cells(startRow, UNKNOWN_START_COL), ws.Cells(startRow, UNKNOWN_START_COL + 2)).UnMerge
    On Error GoTo ErrHandler

    With ws.Range(ws.Cells(startRow, UNKNOWN_START_COL), ws.Cells(startRow, UNKNOWN_START_COL + 2))
        .Merge
        .Value = UNKNOWN_TITLE
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(192, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    ' Header row.
    ws.Cells(startRow + 1, UNKNOWN_START_COL).Value = "Ticker"
    ws.Cells(startRow + 1, UNKNOWN_START_COL + 1).Value = "Description"
    ws.Cells(startRow + 1, UNKNOWN_START_COL + 2).Value = "Asset Class"

    With ws.Range(ws.Cells(startRow + 1, UNKNOWN_START_COL), ws.Cells(startRow + 1, UNKNOWN_START_COL + 2))
        .Font.Bold = True
        .Interior.Color = RGB(0, 176, 240)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With

    Dim i As Long
    Dim r As Long

    For i = 1 To unknownCount
        r = startRow + 1 + i

        ws.Cells(r, UNKNOWN_START_COL).Value = UCase(Trim(CStr(unknownTickers(i))))
        ws.Cells(r, UNKNOWN_START_COL + 1).Value = CStr(unknownDescs(i))
        ws.Cells(r, UNKNOWN_START_COL + 2).Value = ""

        With ws.Range(ws.Cells(r, UNKNOWN_START_COL), ws.Cells(r, UNKNOWN_START_COL + 2))
            .Interior.Color = RGB(255, 242, 204)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .VerticalAlignment = xlTop
        End With

        ws.Cells(r, UNKNOWN_START_COL).Font.Bold = True
        ws.Cells(r, UNKNOWN_START_COL + 2).Font.Color = RGB(0, 0, 255)
        ws.Cells(r, UNKNOWN_START_COL + 2).Font.Bold = True

        With ws.Cells(r, UNKNOWN_START_COL + 2).Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:=VALID_CLASSES
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Asset Class"
            .InputMessage = "Choose the asset class for this ticker."
            .ErrorTitle = "Invalid Asset Class"
            .ErrorMessage = "Choose one of the valid asset classes from the dropdown."
        End With
    Next i

    ws.Columns(UNKNOWN_START_COL).ColumnWidth = 12
    ws.Columns(UNKNOWN_START_COL + 1).ColumnWidth = 42
    ws.Columns(UNKNOWN_START_COL + 2).ColumnWidth = 16

    Exit Sub

ErrHandler:
    MsgBox "Error writing unknown ticker review box: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ============================================================
' PUBLIC: User runs this from Alt+F8
' ============================================================

Public Sub SaveUnknownsAndRefresh()
    On Error GoTo ErrHandler

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB. Open the processed client workbook first.", vbCritical
        Exit Sub
    End If

    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = ActiveSheet

    PrepareCDSWorksheetForMacro ws

    If ScenariosAlreadyExist(ws) Then
        MsgBox "Scenarios already exist on this sheet." & vbCrLf & vbCrLf & _
               "SaveUnknownsAndRefresh needs to sort the holdings rows, which would desync scenario columns." & vbCrLf & vbCrLf & _
               "Recommended workflow:" & vbCrLf & _
               "1. Run ProcessCDSHoldings." & vbCrLf & _
               "2. Fill the unknown ticker review box." & vbCrLf & _
               "3. Run SaveUnknownsAndRefresh." & vbCrLf & _
               "4. Then run AddRaiseCashScenarios.", _
               vbExclamation, "Run Before Scenarios"
        GoTo Done
    End If

    Dim startRow As Long
    startRow = FindUnknownBoxStartRow(ws)

    If UCase(Trim(CStr(ws.Cells(startRow, UNKNOWN_START_COL).Value))) <> UNKNOWN_TITLE Then
        MsgBox "No unknown ticker review box found on this sheet.", vbExclamation
        GoTo Done
    End If

    Dim r As Long
    Dim savedCount As Long
    Dim missingClass As String
    Dim tk As String
    Dim cls As String
    Dim lastUnknownRow As Long

    r = startRow + 2

    Do While Trim(CStr(ws.Cells(r, UNKNOWN_START_COL).Value)) <> ""
        tk = UCase(Trim(CStr(ws.Cells(r, UNKNOWN_START_COL).Value)))
        cls = UCase(Trim(CStr(ws.Cells(r, UNKNOWN_START_COL + 2).Value)))
        lastUnknownRow = r

        If cls = "" Then
            missingClass = missingClass & tk & vbCrLf
        ElseIf Not IsValidAssetClass(cls) Then
            missingClass = missingClass & tk & " (invalid class: " & cls & ")" & vbCrLf
        End If

        r = r + 1
    Loop

    If missingClass <> "" Then
        MsgBox "These tickers still need an asset class:" & vbCrLf & vbCrLf & missingClass, vbExclamation
        GoTo Done
    End If

    If lastUnknownRow = 0 Then
        MsgBox "No ticker classifications were saved.", vbInformation
        GoTo Done
    End If

    For r = startRow + 2 To lastUnknownRow
        tk = UCase(Trim(CStr(ws.Cells(r, UNKNOWN_START_COL).Value)))
        cls = UCase(Trim(CStr(ws.Cells(r, UNKNOWN_START_COL + 2).Value)))

        AddTickerClassToSettings tk, cls
        savedCount = savedCount + 1
    Next r

    RefreshTickerMapCache

    ' Clear the merged review box before refreshing the allocation pivot.
    ' Otherwise a pivot resize can collide with the M:O review block and throw 1004.
    ClearUnknownTickerReviewBox ws, startRow, r

    ReclassifySortAndRefreshReport ws

    MsgBox savedCount & " ticker classification(s) saved to CDS_Settings." & vbCrLf & _
           "Report reclassified, sorted, recalculated, and allocation pivot refreshed.", _
           vbInformation, "Unknowns Saved"

Done:
    Application.ScreenUpdating = prevScreenUpdating
    Application.Calculation = prevCalculation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Resume Done
End Sub

' ============================================================
' INTERNAL HELPERS
' ============================================================

Private Function FindUnknownBoxStartRow(ws As Worksheet) As Long
    Dim pvt As PivotTable

    On Error Resume Next
    Set pvt = ws.PivotTables("AllocPivot")
    On Error GoTo 0

    If Not pvt Is Nothing Then
        FindUnknownBoxStartRow = pvt.TableRange1.Row + pvt.TableRange1.Rows.Count + 2
    Else
        FindUnknownBoxStartRow = 2
    End If
End Function

Private Function ScenariosAlreadyExist(ws As Worksheet) As Boolean
    On Error GoTo FallbackCheck

    Dim startCol As Long
    startCol = ScenStartCol()

    If InStr(1, CStr(ws.Cells(2, startCol).Value), "Raise $", vbTextCompare) > 0 Then
        ScenariosAlreadyExist = True
    Else
        ScenariosAlreadyExist = False
    End If

    Exit Function

FallbackCheck:
    ' Default scenario start column is P / 16 in current settings design.
    If InStr(1, CStr(ws.Cells(2, 16).Value), "Raise $", vbTextCompare) > 0 Then
        ScenariosAlreadyExist = True
    Else
        ScenariosAlreadyExist = False
    End If
End Function

Private Function IsValidAssetClass(ByVal assetClass As String) As Boolean
    Dim validItems As Variant
    Dim i As Long

    validItems = Split(VALID_CLASSES, ",")

    For i = LBound(validItems) To UBound(validItems)
        If UCase(Trim(CStr(validItems(i)))) = UCase(Trim(assetClass)) Then
            IsValidAssetClass = True
            Exit Function
        End If
    Next i
End Function

Private Sub ClearUnknownTickerReviewBox(ByVal ws As Worksheet, ByVal startRow As Long, ByVal firstBlankRow As Long)
    Dim clearEndRow As Long

    clearEndRow = Application.Max(startRow + 50, firstBlankRow + 5)

    With ws.Range(ws.Cells(startRow, UNKNOWN_START_COL), _
                  ws.Cells(clearEndRow, UNKNOWN_START_COL + UNKNOWN_WIDTH - 1))
        On Error Resume Next
        .UnMerge
        On Error GoTo 0
        .Clear
    End With
End Sub

Private Sub ReclassifySortAndRefreshReport(ws As Worksheet)
    Dim headerRow As Long
    Dim dataStart As Long
    Dim dataEnd As Long
    Dim totRow As Long

    headerRow = FindHeaderRowLocal(ws)

    If headerRow = 0 Then
        Err.Raise vbObjectError + 100, , "Could not find ASSET CLASS header row."
    End If

    dataStart = headerRow + 1

    totRow = FindTotalRowLocal(ws, dataStart)

    If totRow = 0 Then
        Err.Raise vbObjectError + 101, , "Could not find total row."
    End If

    dataEnd = totRow - 1

    Dim r As Long
    Dim tk As String
    Dim cls As String

    ' Clear any leftover helper values from failed v1.0 run.
    ws.Range(ws.Cells(dataStart, 15), ws.Cells(dataEnd, 15)).Clear

    ' Reclassify every main holding from centralized settings.
    For r = dataStart To dataEnd
        tk = Trim(CStr(ws.Cells(r, 3).Value))
        cls = ClassifyTickerWithFallback(tk)

        If cls = "" Then cls = "???"

        ws.Cells(r, 1).Value = cls
    Next r

    ' Helper sort column.
    ' Column L is intentionally used because M:N hold the allocation pivot.
    Dim helperCol As Long
    helperCol = 12 ' L

    For r = dataStart To dataEnd
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "CASH" Then
            ws.Cells(r, helperCol).Value = "!CASH"
        Else
            ws.Cells(r, helperCol).Value = UCase(Trim(CStr(ws.Cells(r, 1).Value)))
        End If
    Next r

    ' Sort only the main holdings table plus helper column.
    ' Do NOT include M:N because that is where AllocPivot lives.
    With ws.Sort
        .SortFields.Clear

        .SortFields.Add key:=ws.Range(ws.Cells(dataStart, helperCol), ws.Cells(dataEnd, helperCol)), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        .SortFields.Add key:=ws.Range(ws.Cells(dataStart, 2), ws.Cells(dataEnd, 2)), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal

        .SetRange ws.Range(ws.Cells(dataStart, 1), ws.Cells(dataEnd, helperCol))
        .Header = xlNo
        .Apply
    End With

    ws.Range(ws.Cells(dataStart, helperCol), ws.Cells(dataEnd, helperCol)).Clear

    ' Rebuild row-level formulas after sort.
    For r = dataStart To dataEnd
        ws.Cells(r, 4).Formula = "=E" & r & "/$E$" & totRow
        ws.Cells(r, 4).NumberFormat = "0%"

        ws.Cells(r, 7).Formula = "=IF(E" & r & "=0,0,F" & r & "/E" & r & ")"
        ws.Cells(r, 7).NumberFormat = "0%"
    Next r

    Application.Calculate

    ' Refresh base allocation pivot.
    On Error Resume Next
    ws.PivotTables("AllocPivot").RefreshTable
    On Error GoTo 0
End Sub

Private Function FindHeaderRowLocal(ws As Worksheet) As Long
    Dim i As Long

    FindHeaderRowLocal = 0

    For i = 1 To 5
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRowLocal = i
            Exit Function
        End If
    Next i
End Function

Private Function FindTotalRowLocal(ws As Worksheet, dataStart As Long) As Long
    Dim i As Long

    FindTotalRowLocal = 0

    For i = dataStart To dataStart + 300
        If Trim(CStr(ws.Cells(i, 1).Value)) = "" And ws.Cells(i, 5).Value <> "" Then
            FindTotalRowLocal = i
            Exit Function
        End If
    Next i
End Function

