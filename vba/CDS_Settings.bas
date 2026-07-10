Attribute VB_Name = "CDS_Settings"
Option Explicit

' ============================================================
' CDS Settings v1.1
'
' v1.1 fix: BuildSeedTickers split into chunks because VBA caps
' line continuations at 24 per statement. Original v1.0 packed
' ~150 tickers into one Array() call and failed to compile.
'
' Manages a hidden CDS_Settings sheet inside PERSONAL.XLSB that
' holds all user-customizable values + the ticker classification
' map. Other modules call:
'   GetSetting("name", fallback)
'   GetSettingNum("name", fallback)
'   ClassifyTickerWithFallback(ticker)
'
' SHEET LAYOUT:
'   Cols A-B: Named settings (key/value pairs)
'   Cols D-E: Ticker classification map
'   Col G:    Instructions
'
' MAINTENANCE:
'   Alt+F8 > OpenSettings  -> unhides sheet, jumps to it
'   Alt+F8 > CloseSettings -> hides sheet, saves PERSONAL.XLSB,
'                             clears ticker cache
' ============================================================

Public Const SETTINGS_SHEET_NAME As String = "CDS_Settings"
Private cachedTickerMap As Object
Private g_PreviousWorkbookName As String

Public Function GetSetting(settingName As String, Optional defaultVal As String = "") As String
    On Error GoTo ReturnFallback
    Dim ws As Worksheet
    Set ws = GetOrCreateSettingsSheet()

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If LCase(Trim(CStr(ws.Cells(i, 1).Value))) = LCase(Trim(settingName)) Then
            GetSetting = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i

ReturnFallback:
    GetSetting = defaultVal
End Function

Public Function GetSettingNum(settingName As String, Optional defaultVal As Double = 0) As Double
    Dim s As String
    s = GetSetting(settingName, "")
    If IsNumeric(s) Then
        GetSettingNum = CDbl(s)
    Else
        GetSettingNum = defaultVal
    End If
End Function

Public Sub RefreshTickerMapCache()
    Set cachedTickerMap = Nothing
End Sub

Public Sub AddTickerClassToSettings(ticker As String, assetClass As String)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSettingsSheet()

    Dim tk As String
    Dim cls As String

    tk = UCase(Trim(CStr(ticker)))
    cls = UCase(Trim(CStr(assetClass)))

    If tk = "" Or cls = "" Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    If lastRow < 2 Then lastRow = 1

    Dim i As Long

    ' If ticker already exists, update its class.
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 4).Value))) = tk Then
            ws.Cells(i, 5).Value = cls
            RefreshTickerMapCache
            Exit Sub
        End If
    Next i

    ' Otherwise append new ticker/class pair.
    ws.Cells(lastRow + 1, 4).Value = tk
    ws.Cells(lastRow + 1, 5).Value = cls

    RefreshTickerMapCache
    Exit Sub

ErrHandler:
    MsgBox "Error adding ticker to settings: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Public Function GetTickerClass(ticker As String) As String
    On Error GoTo UseFallback

    Dim tk As String
    tk = UCase(Trim(ticker))
    If tk = "" Then GetTickerClass = "": Exit Function
    If tk = "CASH" Then GetTickerClass = "CASH": Exit Function

    If cachedTickerMap Is Nothing Then
        Set cachedTickerMap = LoadTickerMapFromSheet()
    End If

    If cachedTickerMap.Exists(tk) Then
        GetTickerClass = cachedTickerMap(tk)
    Else
        GetTickerClass = ""
    End If
    Exit Function

UseFallback:
    GetTickerClass = ""
End Function
Private Function LoadTickerMapFromSheet() As Object
    Dim ws As Worksheet
    Set ws = GetOrCreateSettingsSheet()

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    Dim i As Long, tk As String, cls As String
    For i = 2 To lastRow
        tk = UCase(Trim(CStr(ws.Cells(i, 4).Value)))
        cls = UCase(Trim(CStr(ws.Cells(i, 5).Value)))
        If tk <> "" And cls <> "" Then
            map(tk) = cls
        End If
    Next i

    Set LoadTickerMapFromSheet = map
End Function

Public Function ClassifyTickerWithFallback(ticker As String) As String
    Dim tk As String
    tk = UCase(Trim(ticker))
    If tk = "" Then ClassifyTickerWithFallback = "": Exit Function
    If tk = "CASH" Then ClassifyTickerWithFallback = "CASH": Exit Function

    Dim cls As String
    cls = GetTickerClass(tk)
    If cls <> "" Then
        ClassifyTickerWithFallback = cls
    ElseIf IsCUSIPSettings(tk) Then
        ClassifyTickerWithFallback = "BOND"
    Else
        ClassifyTickerWithFallback = ""
    End If
End Function

Public Function IsCUSIPSettings(ticker As String) As Boolean
    IsCUSIPSettings = False
    ticker = Trim(CStr(ticker))

    ' Excel's normal CSV open can coerce all-numeric CUSIPs and strip a leading zero.
    ' Treat 8-digit numeric values as CUSIPs so bonds do not become ???.
    If Len(ticker) = 8 And IsNumeric(ticker) Then
        IsCUSIPSettings = True
        Exit Function
    End If

    If Len(ticker) <> 9 Then Exit Function

    Dim jj As Long, ch As String
    For jj = 1 To 9
        ch = Mid(ticker, jj, 1)
        If Not (ch >= "0" And ch <= "9") And _
           Not (ch >= "A" And ch <= "Z") And _
           Not (ch >= "a" And ch <= "z") Then Exit Function
    Next jj
    ch = Right(ticker, 1)
    If ch >= "0" And ch <= "9" Then IsCUSIPSettings = True
End Function

Public Function GetOrCreateSettingsSheet() As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Sheets.Add
        ws.Name = SETTINGS_SHEET_NAME
        SeedSettingsSheet ws
        ws.Visible = xlSheetHidden
    End If

    Set GetOrCreateSettingsSheet = ws
End Function

Private Sub SeedSettingsSheet(ws As Worksheet)
    ws.Cells.Clear

    ws.Cells(1, 1).Value = "Setting"
    ws.Cells(1, 2).Value = "Value"
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(0, 0, 0)
    ws.Range("A1:B1").Font.Color = RGB(255, 255, 255)

    ' Settings rows (no continuations, one pair at a time)
    Dim r As Long
    r = 2
    WriteSetting ws, r, "DefaultMM", "CJTXX": r = r + 1
    WriteSetting ws, r, "DefaultCGRate", "0.371": r = r + 1
    WriteSetting ws, r, "TradeDeskEmail", "": r = r + 1
    WriteSetting ws, r, "Salutation", "TT,": r = r + 1
    WriteSetting ws, r, "Signoff", "Thanks,": r = r + 1
    WriteSetting ws, r, "LiquidationThresh", "0.95": r = r + 1
    WriteSetting ws, r, "MinTradeAmt", "1": r = r + 1
    WriteSetting ws, r, "BuyPlanRows", "10": r = r + 1
    WriteSetting ws, r, "ScenarioStartCol", "16": r = r + 1
    WriteSetting ws, r, "ScenarioStride", "6": r = r + 1
    WriteSetting ws, r, "FreezePanesAt", "D3": r = r + 1
    WriteSetting ws, r, "AutomationMode", "0": r = r + 1

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 30

    ws.Cells(1, 4).Value = "Ticker"
    ws.Cells(1, 5).Value = "Asset Class"
    ws.Range("D1:E1").Font.Bold = True
    ws.Range("D1:E1").Interior.Color = RGB(0, 0, 0)
    ws.Range("D1:E1").Font.Color = RGB(255, 255, 255)

    ' Ticker map - chunked seed loaders, each well under 24 continuations
    Dim tickerRow As Long
    tickerRow = 2
    SeedShortTickers ws, tickerRow
    SeedBondTickers ws, tickerRow
    SeedGlobalTickers ws, tickerRow
    SeedIntlTickers ws, tickerRow
    SeedLargeTickers ws, tickerRow
    SeedMixedTickers ws, tickerRow
    SeedMvolTickers ws, tickerRow
    SeedReitTickers ws, tickerRow
    SeedSectTickers ws, tickerRow
    SeedSmidTickers ws, tickerRow
    SeedStockTickersA ws, tickerRow
    SeedStockTickersB ws, tickerRow
    SeedStockTickersC ws, tickerRow

    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 14

    ws.Cells(1, 7).Value = "INSTRUCTIONS"
    ws.Cells(1, 7).Font.Bold = True
    ws.Cells(1, 7).Interior.Color = RGB(0, 0, 0)
    ws.Cells(1, 7).Font.Color = RGB(255, 255, 255)
    ws.Cells(2, 7).Value = "1. Edit values in column B to change defaults."
    ws.Cells(3, 7).Value = "2. Add tickers in cols D:E (one per row, ticker + asset class)."
    ws.Cells(4, 7).Value = "3. Valid asset classes: CASH, SHORT, BOND, GLOBAL, INTL, LARGE,"
    ws.Cells(5, 7).Value = "   MIXED, MVOL, REIT, SECT, SMID, STOCK"
    ws.Cells(6, 7).Value = "4. Run CloseSettings when done (hides this sheet, saves)."
    ws.Cells(7, 7).Value = "5. Changes apply on next macro run."
    ws.Columns("G").ColumnWidth = 70
End Sub

Private Sub WriteSetting(ws As Worksheet, r As Long, key As String, val As String)
    ws.Cells(r, 1).Value = key
    ws.Cells(r, 2).Value = val
End Sub

Private Sub WriteTicker(ws As Worksheet, ByRef r As Long, ticker As String, cls As String)
    ws.Cells(r, 4).Value = ticker
    ws.Cells(r, 5).Value = cls
    r = r + 1
End Sub

' ============================================================
' SEED CHUNKS
' ============================================================

Private Sub SeedShortTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "CJTXX", "SHORT"
    WriteTicker ws, r, "LUBYX", "SHORT"
    WriteTicker ws, r, "RJMXX", "SHORT"
End Sub

Private Sub SeedBondTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "AGMHX", "BOND"
    WriteTicker ws, r, "FAFTX", "BOND"
    WriteTicker ws, r, "FCAVX", "BOND"
    WriteTicker ws, r, "FRCZX", "BOND"
    WriteTicker ws, r, "FVCAX", "BOND"
    WriteTicker ws, r, "GUIRX", "BOND"
    WriteTicker ws, r, "LBNYX", "BOND"
    WriteTicker ws, r, "LFLIX", "BOND"
    WriteTicker ws, r, "TGBAX", "BOND"
    WriteTicker ws, r, "VCADX", "BOND"
End Sub

Private Sub SeedGlobalTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "SGIIX", "GLOBAL"
End Sub

Private Sub SeedIntlTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "AEPFX", "INTL"
    WriteTicker ws, r, "AVDV", "INTL"
    WriteTicker ws, r, "AVEM", "INTL"
    WriteTicker ws, r, "CIVIX", "INTL"
    WriteTicker ws, r, "DODFX", "INTL"
    WriteTicker ws, r, "EFA", "INTL"
    WriteTicker ws, r, "FENI", "INTL"
    WriteTicker ws, r, "FIIIX", "INTL"
    WriteTicker ws, r, "GEMIX", "INTL"
    WriteTicker ws, r, "GIRMX", "INTL"
    WriteTicker ws, r, "IEFA", "INTL"
    WriteTicker ws, r, "IEMG", "INTL"
    WriteTicker ws, r, "MQGIX", "INTL"
    WriteTicker ws, r, "MWNIX", "INTL"
    WriteTicker ws, r, "MYSIX", "INTL"
    WriteTicker ws, r, "ODVYX", "INTL"
    WriteTicker ws, r, "SCHC", "INTL"
    WriteTicker ws, r, "SEQFX", "INTL"
    WriteTicker ws, r, "TGVIX", "INTL"
    WriteTicker ws, r, "WDIV", "INTL"
End Sub

Private Sub SeedLargeTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "BKDV", "LARGE"
    WriteTicker ws, r, "CGGR", "LARGE"
    WriteTicker ws, r, "DODBX", "LARGE"
    WriteTicker ws, r, "DODGX", "LARGE"
    WriteTicker ws, r, "DVY", "LARGE"
    WriteTicker ws, r, "EITVX", "LARGE"
    WriteTicker ws, r, "FDVLX", "LARGE"
    WriteTicker ws, r, "FDYZX", "LARGE"
    WriteTicker ws, r, "FEQTX", "LARGE"
    WriteTicker ws, r, "FFOG", "LARGE"
    WriteTicker ws, r, "FMILX", "LARGE"
    WriteTicker ws, r, "FRDAX", "LARGE"
    WriteTicker ws, r, "FXAIX", "LARGE"
    WriteTicker ws, r, "IVE", "LARGE"
    WriteTicker ws, r, "IVV", "LARGE"
    WriteTicker ws, r, "IWF", "LARGE"
    WriteTicker ws, r, "MNHIX", "LARGE"
    WriteTicker ws, r, "PAXIX", "LARGE"
    WriteTicker ws, r, "PRSGX", "LARGE"
    WriteTicker ws, r, "SCHD", "LARGE"
    WriteTicker ws, r, "SDY", "LARGE"
    WriteTicker ws, r, "SFLNX", "LARGE"
    WriteTicker ws, r, "SPY", "LARGE"
    WriteTicker ws, r, "SPYG", "LARGE"
    WriteTicker ws, r, "SPYV", "LARGE"
    WriteTicker ws, r, "STLYX", "LARGE"
    WriteTicker ws, r, "TBCIX", "LARGE"
    WriteTicker ws, r, "TCHP", "LARGE"
    WriteTicker ws, r, "TRBCX", "LARGE"
    WriteTicker ws, r, "VDIGX", "LARGE"
    WriteTicker ws, r, "VFIAX", "LARGE"
    WriteTicker ws, r, "VIG", "LARGE"
    WriteTicker ws, r, "VOO", "LARGE"
    WriteTicker ws, r, "VPMAX", "LARGE"
    WriteTicker ws, r, "VWENX", "LARGE"
    WriteTicker ws, r, "VWNAX", "LARGE"
End Sub

Private Sub SeedMixedTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "FRIAX", "MIXED"
    WriteTicker ws, r, "VWIAX", "MIXED"
End Sub

Private Sub SeedMvolTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "JHEQX", "MVOL"
    WriteTicker ws, r, "JHQDX", "MVOL"
    WriteTicker ws, r, "USMV", "MVOL"
    WriteTicker ws, r, "BUFP", "MVOL"
    WriteTicker ws, r, "JEPI", "MVOL"
End Sub

Private Sub SeedReitTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "FRESX", "REIT"
    WriteTicker ws, r, "ICF", "REIT"
    WriteTicker ws, r, "IYR", "REIT"
    WriteTicker ws, r, "PHRIX", "REIT"
    WriteTicker ws, r, "PURZX", "REIT"
End Sub

Private Sub SeedSectTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "THISX", "SECT"
End Sub

Private Sub SeedSmidTickers(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "FCDIX", "SMID"
    WriteTicker ws, r, "FESM", "SMID"
    WriteTicker ws, r, "FGSIX", "SMID"
    WriteTicker ws, r, "FTHSX", "SMID"
    WriteTicker ws, r, "FTSIX", "SMID"
    WriteTicker ws, r, "IJH", "SMID"
    WriteTicker ws, r, "IJR", "SMID"
    WriteTicker ws, r, "IMCB", "SMID"
    WriteTicker ws, r, "IWM", "SMID"
    WriteTicker ws, r, "IWN", "SMID"
    WriteTicker ws, r, "IWO", "SMID"
    WriteTicker ws, r, "IWS", "SMID"
    WriteTicker ws, r, "JISGX", "SMID"
    WriteTicker ws, r, "NOSGX", "SMID"
    WriteTicker ws, r, "PRJIX", "SMID"
    WriteTicker ws, r, "PRSVX", "SMID"
    WriteTicker ws, r, "SLYV", "SMID"
    WriteTicker ws, r, "STMPX", "SMID"
    WriteTicker ws, r, "VBK", "SMID"
    WriteTicker ws, r, "VBR", "SMID"
    WriteTicker ws, r, "VIMAX", "SMID"
    WriteTicker ws, r, "VO", "SMID"
    WriteTicker ws, r, "LSGRX", "SMID"
End Sub

Private Sub SeedStockTickersA(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "AAPL", "STOCK"
    WriteTicker ws, r, "ABBV", "STOCK"
    WriteTicker ws, r, "ABT", "STOCK"
    WriteTicker ws, r, "ACN", "STOCK"
    WriteTicker ws, r, "ADBE", "STOCK"
    WriteTicker ws, r, "ADP", "STOCK"
    WriteTicker ws, r, "AMAT", "STOCK"
    WriteTicker ws, r, "AMD", "STOCK"
    WriteTicker ws, r, "AMGN", "STOCK"
    WriteTicker ws, r, "AMZN", "STOCK"
    WriteTicker ws, r, "AON", "STOCK"
    WriteTicker ws, r, "APH", "STOCK"
    WriteTicker ws, r, "AVGO", "STOCK"
    WriteTicker ws, r, "AXP", "STOCK"
    WriteTicker ws, r, "BA", "STOCK"
    WriteTicker ws, r, "BAC", "STOCK"
    WriteTicker ws, r, "BLK", "STOCK"
    WriteTicker ws, r, "BKNG", "STOCK"
    WriteTicker ws, r, "BMY", "STOCK"
    WriteTicker ws, r, "BRK.B", "STOCK"
    WriteTicker ws, r, "C", "STOCK"
    WriteTicker ws, r, "CAT", "STOCK"
    WriteTicker ws, r, "CB", "STOCK"
    WriteTicker ws, r, "CCL", "STOCK"
    WriteTicker ws, r, "CHD", "STOCK"
    WriteTicker ws, r, "CI", "STOCK"
    WriteTicker ws, r, "CL", "STOCK"
    WriteTicker ws, r, "CME", "STOCK"
    WriteTicker ws, r, "CMG", "STOCK"
    WriteTicker ws, r, "COIN", "STOCK"
    WriteTicker ws, r, "COP", "STOCK"
    WriteTicker ws, r, "COST", "STOCK"
    WriteTicker ws, r, "CRM", "STOCK"
    WriteTicker ws, r, "CRWD", "STOCK"
    WriteTicker ws, r, "CSCO", "STOCK"
    WriteTicker ws, r, "CVX", "STOCK"
End Sub

Private Sub SeedStockTickersB(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "DE", "STOCK"
    WriteTicker ws, r, "DIS", "STOCK"
    WriteTicker ws, r, "ETN", "STOCK"
    WriteTicker ws, r, "F", "STOCK"
    WriteTicker ws, r, "FI", "STOCK"
    WriteTicker ws, r, "GD", "STOCK"
    WriteTicker ws, r, "GE", "STOCK"
    WriteTicker ws, r, "GEHC", "STOCK"
    WriteTicker ws, r, "GILD", "STOCK"
    WriteTicker ws, r, "GM", "STOCK"
    WriteTicker ws, r, "GOOG", "STOCK"
    WriteTicker ws, r, "GOOGL", "STOCK"
    WriteTicker ws, r, "GS", "STOCK"
    WriteTicker ws, r, "HCA", "STOCK"
    WriteTicker ws, r, "HD", "STOCK"
    WriteTicker ws, r, "HON", "STOCK"
    WriteTicker ws, r, "IBM", "STOCK"
    WriteTicker ws, r, "ICE", "STOCK"
    WriteTicker ws, r, "INTC", "STOCK"
    WriteTicker ws, r, "ISRG", "STOCK"
    WriteTicker ws, r, "JNJ", "STOCK"
    WriteTicker ws, r, "JPM", "STOCK"
    WriteTicker ws, r, "KD", "STOCK"
    WriteTicker ws, r, "KLAC", "STOCK"
    WriteTicker ws, r, "KO", "STOCK"
    WriteTicker ws, r, "LLY", "STOCK"
    WriteTicker ws, r, "LMT", "STOCK"
    WriteTicker ws, r, "LOW", "STOCK"
    WriteTicker ws, r, "MA", "STOCK"
    WriteTicker ws, r, "MAR", "STOCK"
    WriteTicker ws, r, "MCD", "STOCK"
    WriteTicker ws, r, "MCO", "STOCK"
    WriteTicker ws, r, "MDLZ", "STOCK"
    WriteTicker ws, r, "META", "STOCK"
    WriteTicker ws, r, "MO", "STOCK"
    WriteTicker ws, r, "MRK", "STOCK"
    WriteTicker ws, r, "MS", "STOCK"
    WriteTicker ws, r, "MSFT", "STOCK"
End Sub

Private Sub SeedStockTickersC(ws As Worksheet, ByRef r As Long)
    WriteTicker ws, r, "NCLH", "STOCK"
    WriteTicker ws, r, "NEE", "STOCK"
    WriteTicker ws, r, "NFLX", "STOCK"
    WriteTicker ws, r, "NOC", "STOCK"
    WriteTicker ws, r, "NOW", "STOCK"
    WriteTicker ws, r, "NVDA", "STOCK"
    WriteTicker ws, r, "ORCL", "STOCK"
    WriteTicker ws, r, "PANW", "STOCK"
    WriteTicker ws, r, "PEP", "STOCK"
    WriteTicker ws, r, "PFE", "STOCK"
    WriteTicker ws, r, "PG", "STOCK"
    WriteTicker ws, r, "PLD", "STOCK"
    WriteTicker ws, r, "PLTR", "STOCK"
    WriteTicker ws, r, "PM", "STOCK"
    WriteTicker ws, r, "PYPL", "STOCK"
    WriteTicker ws, r, "QCOM", "STOCK"
    WriteTicker ws, r, "RCL", "STOCK"
    WriteTicker ws, r, "REGN", "STOCK"
    WriteTicker ws, r, "RTX", "STOCK"
    WriteTicker ws, r, "SCHW", "STOCK"
    WriteTicker ws, r, "SO", "STOCK"
    WriteTicker ws, r, "SPGI", "STOCK"
    WriteTicker ws, r, "SPWR", "STOCK"
    WriteTicker ws, r, "SYK", "STOCK"
    WriteTicker ws, r, "T", "STOCK"
    WriteTicker ws, r, "TGT", "STOCK"
    WriteTicker ws, r, "TSLA", "STOCK"
    WriteTicker ws, r, "TXN", "STOCK"
    WriteTicker ws, r, "UBER", "STOCK"
    WriteTicker ws, r, "UNH", "STOCK"
    WriteTicker ws, r, "UNP", "STOCK"
    WriteTicker ws, r, "UPS", "STOCK"
    WriteTicker ws, r, "V", "STOCK"
    WriteTicker ws, r, "VZ", "STOCK"
    WriteTicker ws, r, "WAB", "STOCK"
    WriteTicker ws, r, "WM", "STOCK"
    WriteTicker ws, r, "WMT", "STOCK"
    WriteTicker ws, r, "XOM", "STOCK"
End Sub

' ============================================================
' OPEN / CLOSE settings sheet
' ============================================================
Public Sub OpenSettings()
Attribute OpenSettings.VB_ProcData.VB_Invoke_Func = "T\n14"
    ' Remember which workbook was active before opening settings,
    ' so CloseSettings can return to it.
    On Error Resume Next
    g_PreviousWorkbookName = ActiveWorkbook.Name
    On Error GoTo 0

    Dim ws As Worksheet
    Set ws = GetOrCreateSettingsSheet()
    ws.Visible = xlSheetVisible

    ' Make PERSONAL.XLSB itself visible so the user can see the tab
    On Error Resume Next
    Application.Windows("PERSONAL.XLSB").Visible = True
    On Error GoTo 0

    ws.Activate
    ws.Range("A2").Select
    MsgBox "Settings open. Edit cells, then run CloseSettings to hide and save.", _
           vbInformation, "CDS Settings"
End Sub

Public Sub CloseSettings()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No settings sheet found.", vbExclamation
        Exit Sub
    End If

    ws.Visible = xlSheetHidden
    RefreshTickerMapCache

    ' Save PERSONAL.XLSB before hiding (Excel can be cranky about saving
    ' a hidden workbook in some versions)
    On Error Resume Next
    ThisWorkbook.Save
    Dim saveErrNum As Long
    Dim saveErrDesc As String
    saveErrNum = Err.Number
    saveErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0

    ' Switch back to the workbook the user was on before opening settings.
    ' If we can't find it (closed, never set), fall back to first non-PERSONAL workbook.
    Dim wbName As String
    wbName = g_PreviousWorkbookName
    Dim switched As Boolean
    switched = False

    If wbName <> "" And LCase(wbName) <> "personal.xlsb" Then
        On Error Resume Next
        Workbooks(wbName).Activate
        If Err.Number = 0 Then switched = True
        Err.Clear
        On Error GoTo 0
    End If

    If Not switched Then
        ' Find any other open workbook to switch to
        Dim wb As Workbook
        For Each wb In Application.Workbooks
            If LCase(wb.Name) <> "personal.xlsb" Then
                wb.Activate
                switched = True
                Exit For
            End If
        Next wb
    End If

    ' Now safe to hide PERSONAL.XLSB
    On Error Resume Next
    Application.Windows("PERSONAL.XLSB").Visible = False
    On Error GoTo 0

    g_PreviousWorkbookName = ""

    If saveErrNum <> 0 Then
        MsgBox "Settings were updated for this Excel session, but PERSONAL.XLSB did not save." & vbCrLf & vbCrLf & _
               "Save error " & saveErrNum & ": " & saveErrDesc, _
               vbExclamation, "CDS Settings"
    Else
        MsgBox "Settings saved. Changes will apply on next macro run.", _
               vbInformation, "CDS Settings"
    End If
End Sub
