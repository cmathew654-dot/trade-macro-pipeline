Attribute VB_Name = "CDS_Trade_Email"
Option Explicit

' ============================================================
' CDS Trade Email Generator v4.2
'
' v4.2 changes vs v4.1:
'   - All defaults via CDS_Settings (DefaultMM, Salutation, Signoff,
'     LiquidationThresh, MinTradeAmt, ScenarioStartCol, ScenarioStride,
'     BuyPlanRows, TradeDeskEmail)
'   - PERSONAL.XLSB safety guard
'   - Pre-fills To: address from settings if available
'
' Reads scenario sells + buy plan, generates Outlook HTML email.
' Falls back to clipboard if Outlook isn't running.
' ============================================================

' Constants kept as fallbacks; runtime uses settings values
Private Const SCEN_START_COL_DEFAULT As Long = 16
Private Const SCEN_STRIDE_DEFAULT As Long = 6
Private Const FULL_LIQ_THRESHOLD_DEFAULT As Double = 0.95
Private Const MIN_TRADE_AMT_DEFAULT As Double = 1#
Private Const DEFAULT_SHORT_TICKER_DEFAULT As String = "CJTXX"
Private Const BUY_PLAN_ROWS_DEFAULT As Long = 10
Private Const MAX_SHORT_POSITIONS As Long = 20
Private Const RESIDUAL_TOLERANCE As Double = 0.5
Private Const PLAN_SCENARIO_NUM As Long = 2
Private Const ROUTING_DATA_ROWS As Long = 5

Sub GenerateTradeEmail()
Attribute GenerateTradeEmail.VB_ProcData.VB_Invoke_Func = "E\n14"
    On Error GoTo ErrHandler

    If LCase(ActiveWorkbook.Name) = "personal.xlsb" Then
        MsgBox "Cannot run on PERSONAL.XLSB.", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim scenStartColLocal As Long, scenStrideLocal As Long
    scenStartColLocal = CLng(GetSettingNum("ScenarioStartCol", SCEN_START_COL_DEFAULT))
    scenStrideLocal = CLng(GetSettingNum("ScenarioStride", SCEN_STRIDE_DEFAULT))

    Dim minTradeAmt As Double, fullLiqThresh As Double
    minTradeAmt = GetSettingNum("MinTradeAmt", MIN_TRADE_AMT_DEFAULT)
    fullLiqThresh = GetSettingNum("LiquidationThresh", FULL_LIQ_THRESHOLD_DEFAULT)

    Dim buyPlanRowsLocal As Long
    buyPlanRowsLocal = CLng(GetSettingNum("BuyPlanRows", BUY_PLAN_ROWS_DEFAULT))

    Dim defaultMM As String, salutation As String, signoff As String, deskEmail As String
    defaultMM = GetSetting("DefaultMM", DEFAULT_SHORT_TICKER_DEFAULT)
    salutation = GetSetting("Salutation", "TT,")
    signoff = GetSetting("Signoff", "Thanks,")
    deskEmail = GetSetting("TradeDeskEmail", "")

    Dim automationMode As Boolean
    automationMode = (GetSetting("AutomationMode", "0") = "1")

    ' --- VALIDATE SHEET ---
    Dim headerRow As Long, dataStart As Long, dataEnd As Long, totRow As Long
    headerRow = FindHeaderRow(ws)
    If headerRow = 0 Then
        MsgBox "Run ProcessCDSHoldings first.", vbExclamation
        Exit Sub
    End If
    dataStart = headerRow + 1
    totRow = FindTotalRow(ws, dataStart)
    If totRow = 0 Then
        MsgBox "Could not find totals row.", vbExclamation
        Exit Sub
    End If
    dataEnd = totRow - 1

    Dim scenCount As Long
    scenCount = CountScenarios(ws, scenStartColLocal, scenStrideLocal)
    If scenCount = 0 Then
        MsgBox "No scenarios found. Run AddRaiseCashScenarios first.", vbExclamation
        Exit Sub
    End If

    ' --- PICK SCENARIO ---
    Dim scenNum As Long
    If scenCount = 1 Then
        scenNum = 1
    Else
        Dim sn As String, defaultScen As String
        defaultScen = CStr(scenCount)
        sn = InputBox("Which scenario to email? (1 to " & scenCount & ")", _
                      "Generate Trade Email", defaultScen)
        If sn = "" Then Exit Sub
        If Not IsNumeric(sn) Then MsgBox "Not a number.", vbExclamation: Exit Sub
        scenNum = CLng(sn)
        If scenNum < 1 Or scenNum > scenCount Then
            MsgBox "S" & scenNum & " does not exist. Available: 1 to " & scenCount, vbExclamation
            Exit Sub
        End If
    End If

    Dim scenCol As Long
    scenCol = scenStartColLocal + (scenNum - 1) * scenStrideLocal

    If InStr(1, CStr(ws.Cells(2, scenCol).Value), "Raise $", vbTextCompare) = 0 Then
        MsgBox "Column " & ColLetter(scenCol) & " row 2 does not say 'Raise $'." & vbCrLf & _
               "Macro 2 layout may have shifted. Aborting.", vbExclamation
        Exit Sub
    End If

    Dim acctNum As String, acctName As String
    acctNum = Trim(CStr(ws.Cells(1, 1).Value))
    acctName = Trim(CStr(ws.Cells(1, 2).Value))

    ' Spec-faithful sell narration only applies to the plan scenario (S2)
    ' when a Sell Workbench (Amt Type / Amount columns) exists on the
    ' sheet. Other scenarios and legacy sheets keep the plain dollar
    ' phrasing untouched.
    Dim useSpecPhrasing As Boolean, modeColFound As Long
    Dim specTypeColFound As Long, specAmtColFound As Long
    useSpecPhrasing = False
    If scenNum = PLAN_SCENARIO_NUM Then
        modeColFound = FindModeColEmail(ws)
        If modeColFound > 0 Then
            specTypeColFound = modeColFound + 1
            specAmtColFound = modeColFound + 2
            useSpecPhrasing = True
        End If
    End If

    ' --- COLLECT SELLS ---
    Dim r As Long, i As Long
    Dim ticker As String, raiseAmt As Double, fmv As Double
    Dim totalRaised As Double, tradeCount As Long
    Dim sellLinesHTML As String, sellLinesPlain As String
    Dim fullLiqHTML As String, fullLiqPlain As String
    Dim amtType As String, specAmtVal As Double, isForcedAll As Boolean

    For r = dataStart To dataEnd
        raiseAmt = 0
        If IsNumeric(ws.Cells(r, scenCol).Value) Then raiseAmt = CDbl(ws.Cells(r, scenCol).Value)
        If raiseAmt >= minTradeAmt Then
            ticker = Trim(CStr(ws.Cells(r, 3).Value))
            fmv = 0
            If IsNumeric(ws.Cells(r, 5).Value) Then fmv = CDbl(ws.Cells(r, 5).Value)
            tradeCount = tradeCount + 1
            totalRaised = totalRaised + raiseAmt

            amtType = "": specAmtVal = 0
            If useSpecPhrasing Then
                amtType = Trim(CStr(ws.Cells(r, specTypeColFound).Value))
                If IsNumeric(ws.Cells(r, specAmtColFound).Value) Then specAmtVal = CDbl(ws.Cells(r, specAmtColFound).Value)
            End If

            ' ALL always forces the marker, even if rounding puts the
            ' computed dollar amount a hair under the liquidation threshold.
            isForcedAll = (fmv > 0 And raiseAmt >= fmv * fullLiqThresh) Or (amtType = "ALL")

            fullLiqHTML = "": fullLiqPlain = ""
            If isForcedAll Then
                fullLiqHTML = " <i>*SELLING ALL of this position*</i>"
                fullLiqPlain = " *SELLING ALL of this position*"
            End If

            Select Case amtType
                Case "Shares"
                    sellLinesHTML = sellLinesHTML & "<li>Sell " & FormatSpecAmtEmail(specAmtVal) & _
                                    " shares of <b>" & ticker & "</b> (~$" & Format(raiseAmt, "#,##0") & _
                                    ")" & fullLiqHTML & "</li>" & vbCrLf
                    sellLinesPlain = sellLinesPlain & "   * Sell " & FormatSpecAmtEmail(specAmtVal) & _
                                     " shares of " & ticker & " (~$" & Format(raiseAmt, "#,##0") & _
                                     ")" & fullLiqPlain & vbCrLf
                Case "% Pos"
                    sellLinesHTML = sellLinesHTML & "<li>Sell " & FormatSpecAmtEmail(specAmtVal) & _
                                    "% of <b>" & ticker & "</b> (~$" & Format(raiseAmt, "#,##0") & _
                                    ")" & fullLiqHTML & "</li>" & vbCrLf
                    sellLinesPlain = sellLinesPlain & "   * Sell " & FormatSpecAmtEmail(specAmtVal) & _
                                     "% of " & ticker & " (~$" & Format(raiseAmt, "#,##0") & _
                                     ")" & fullLiqPlain & vbCrLf
                Case "% Acct"
                    sellLinesHTML = sellLinesHTML & "<li>Sell " & FormatSpecAmtEmail(specAmtVal) & _
                                    "% of account value from <b>" & ticker & "</b> (~$" & Format(raiseAmt, "#,##0") & _
                                    ")" & fullLiqHTML & "</li>" & vbCrLf
                    sellLinesPlain = sellLinesPlain & "   * Sell " & FormatSpecAmtEmail(specAmtVal) & _
                                     "% of account value from " & ticker & " (~$" & Format(raiseAmt, "#,##0") & _
                                     ")" & fullLiqPlain & vbCrLf
                Case Else
                    ' "$" spec type, Pool rows, ALL (marker already forced
                    ' above), and legacy sheets without the spec columns.
                    sellLinesHTML = sellLinesHTML & "<li>$" & Format(raiseAmt, "#,##0") & _
                                    "&#9;<b>" & ticker & "</b>" & fullLiqHTML & "</li>" & vbCrLf
                    sellLinesPlain = sellLinesPlain & "   * $" & Format(raiseAmt, "#,##0") & _
                                     vbTab & ticker & fullLiqPlain & vbCrLf
            End Select
        End If
    Next r

    If tradeCount = 0 Then
        MsgBox "No trades found in S" & scenNum & "." & vbCrLf & _
               "Enter Raise $ amounts in the scenario column first.", vbInformation
        Exit Sub
    End If

    ' Income Lost from scenario block
    Dim incomeLost As Double
    incomeLost = 0
    If IsNumeric(ws.Cells(totRow, scenCol + 4).Value) Then
        incomeLost = CDbl(ws.Cells(totRow, scenCol + 4).Value)
    End If

    ' --- READ BUY PLAN ---
    Dim buyTickers() As String, buyAmounts() As Double, buyClasses() As String
    Dim buyCount As Long, totalBought As Double, residual As Double
    Dim hasBuyPlan As Boolean, isV12Plan As Boolean
    Dim incomeGained As Double
    ReDim buyTickers(0 To buyPlanRowsLocal - 1)
    ReDim buyAmounts(0 To buyPlanRowsLocal - 1)
    ReDim buyClasses(0 To buyPlanRowsLocal - 1)

  Dim pivotEnd As Long
Dim buyPlanHeaderRow As Long

pivotEnd = FindPivotGrandTotalRow(ws, scenCol, totRow)

If pivotEnd > 0 Then
    buyPlanHeaderRow = FindBuyPlanHeaderRow_Email(ws, scenCol, pivotEnd)

    If buyPlanHeaderRow > 0 Then
        hasBuyPlan = True
        isV12Plan = (Trim(CStr(ws.Cells(buyPlanHeaderRow + 1, scenCol + 2).Value)) = "Asset Class")
    End If
End If

    Dim unclassifiedTickers As String
    unclassifiedTickers = ""

    If hasBuyPlan Then
    Dim inputStart As Long, inputEnd As Long, totalRow As Long, incomeRow As Long

    inputStart = buyPlanHeaderRow + 2
    inputEnd = inputStart + buyPlanRowsLocal - 1
    totalRow = inputEnd + 1
    incomeRow = totalRow + 1

        For r = inputStart To inputEnd
            Dim t As String, a As Double, cls As String
            t = UCase(Trim(CStr(ws.Cells(r, scenCol).Value)))
            a = 0
            If IsNumeric(ws.Cells(r, scenCol + 1).Value) Then
                a = CDbl(ws.Cells(r, scenCol + 1).Value)
            End If
            cls = ""
            If isV12Plan Then
                cls = UCase(Trim(CStr(ws.Cells(r, scenCol + 2).Value)))
            End If

            If t <> "" And a >= minTradeAmt Then
                buyTickers(buyCount) = t
                buyAmounts(buyCount) = a
                buyClasses(buyCount) = cls
                buyCount = buyCount + 1
                totalBought = totalBought + a
                If isV12Plan And cls = "" Then
                    unclassifiedTickers = unclassifiedTickers & t & ", "
                End If
            ElseIf t = "" And a >= minTradeAmt Then
                MsgBox "Buy plan row " & r & " has amount $" & Format(a, "#,##0") & _
                       " but no ticker. Fix and re-run.", vbExclamation
                Exit Sub
            ElseIf t <> "" And a < minTradeAmt Then
                MsgBox "Buy plan row " & r & " has ticker '" & t & _
                       "' but no amount. Fill in or clear the ticker.", vbExclamation
                Exit Sub
            End If
        Next r

        If isV12Plan Then
            If IsNumeric(ws.Cells(incomeRow, scenCol + 1).Value) Then
                incomeGained = CDbl(ws.Cells(incomeRow, scenCol + 1).Value)
            End If
        End If
    End If

    residual = totalRaised - totalBought

    If residual < -RESIDUAL_TOLERANCE Then
        MsgBox "Buy plan total $" & Format(totalBought, "#,##0") & _
               " exceeds sells $" & Format(totalRaised, "#,##0") & _
               " by $" & Format(-residual, "#,##0") & ".", vbExclamation
        Exit Sub
    End If

    If unclassifiedTickers <> "" Then
        If Right(unclassifiedTickers, 2) = ", " Then unclassifiedTickers = Left(unclassifiedTickers, Len(unclassifiedTickers) - 2)
        Dim resp As VbMsgBoxResult
        resp = MsgBox("These buy tickers have no asset class:" & vbCrLf & "    " & unclassifiedTickers & vbCrLf & vbCrLf & _
                      "Post-rebalance allocation pivot will be incomplete for these buys." & vbCrLf & _
                      "Continue anyway?", vbYesNo + vbQuestion, "Unclassified Tickers")
        If resp = vbNo Then Exit Sub
    End If

    ' --- PROCEEDS ROUTING (S2 only, when a PROCEEDS ROUTING block exists) ---
    ' Replaces the legacy residual-destination prompt: the routing block
    ' already states where every dollar of proceeds goes, so there is
    ' nothing left to ask interactively. A routing block that does not
    ' reconcile ("Routing OK") blocks the email outright - never email a
    ' plan whose routing does not reconcile.
    Dim usesRouting As Boolean
    Dim routingDestCol As Long, routingDataStartRow As Long
    usesRouting = False
    If scenNum = PLAN_SCENARIO_NUM Then
        Dim routingTitleCell As Range
        Set routingTitleCell = FindRoutingTitleCellEmail(ws)
        If Not routingTitleCell Is Nothing Then
            routingDestCol = routingTitleCell.Column
            routingDataStartRow = routingTitleCell.Row + 2
            usesRouting = True
        End If
    End If

    Dim routingParaHTML As String, routingParaPlain As String

    ' --- RESIDUAL DESTINATION ---
    Dim residualTicker As String, hasResidual As Boolean
    hasResidual = (residual >= RESIDUAL_TOLERANCE)

    If usesRouting Then
        Dim routingStatusRow As Long, routingStatusText As String
        routingStatusRow = routingDataStartRow + ROUTING_DATA_ROWS
        routingStatusText = Trim(CStr(ws.Cells(routingStatusRow, routingDestCol).Value))
        If routingStatusText <> "Routing OK" Then
            MsgBox "Proceeds routing does not reconcile (" & routingStatusText & ")." & vbCrLf & _
                   "Fix the PROCEEDS ROUTING block before emailing this plan.", vbExclamation
            Exit Sub
        End If
        BuildRoutingNarrative ws, routingDestCol, routingDataStartRow, totalRaised, routingParaHTML, routingParaPlain
        hasResidual = False
    ElseIf hasResidual Then
        residualTicker = PromptResidualDestination(ws, totRow, residual, (buyCount > 0), defaultMM)
        If residualTicker = "" Then Exit Sub
    End If

    ' --- SUMMARY CONFIRMATION ---
    Dim summaryText As String
    summaryText = "ACCOUNT: " & acctName & " - " & acctNum & vbCrLf & _
                  "SCENARIO: S" & scenNum & vbCrLf & vbCrLf & _
                  "SELL TOTAL: $" & Format(totalRaised, "#,##0") & "  (" & tradeCount & " positions)" & vbCrLf & _
                  "    Income Lost / yr: -$" & Format(incomeLost, "#,##0") & vbCrLf

    If buyCount > 0 Then
        summaryText = summaryText & vbCrLf & "BUY:" & vbCrLf
        For i = 0 To buyCount - 1
            summaryText = summaryText & "    $" & Format(buyAmounts(i), "#,##0") & "    " & buyTickers(i)
            If isV12Plan And buyClasses(i) <> "" Then
                summaryText = summaryText & "    [" & buyClasses(i) & "]"
            End If
            summaryText = summaryText & vbCrLf
        Next i
        summaryText = summaryText & "    Buys subtotal: $" & Format(totalBought, "#,##0") & vbCrLf
        If isV12Plan Then
            summaryText = summaryText & "    Income Gained / yr: +$" & Format(incomeGained, "#,##0") & vbCrLf
        End If
    End If

    If hasResidual Then
        summaryText = summaryText & vbCrLf & "RESIDUAL: $" & Format(residual, "#,##0") & " to " & residualTicker & " (money market)" & vbCrLf
    ElseIf usesRouting And routingParaPlain <> "" Then
        summaryText = summaryText & vbCrLf & "ROUTING: " & routingParaPlain & vbCrLf
    End If

    If isV12Plan And buyCount > 0 Then
        Dim netIncome As Double
        netIncome = incomeGained - incomeLost
        Dim netLabel As String
        If netIncome >= 0 Then netLabel = "+$" Else netLabel = "-$"
        summaryText = summaryText & vbCrLf & "NET INCOME CHANGE: " & netLabel & Format(Abs(netIncome), "#,##0") & "/yr"
    End If

    summaryText = summaryText & vbCrLf & vbCrLf & "Generate email?"

    If Not automationMode Then
        If MsgBox(summaryText, vbYesNo + vbQuestion, "Confirm Trade Email - S" & scenNum) = vbNo Then Exit Sub
    End If

    ' --- BUILD MESSAGE ---
    Dim totalK As String
    totalK = "~$" & Format(totalRaised / 1000, "#,##0") & "K"

    Dim subjectLine As String
    subjectLine = "Trade Instructions - " & acctName & " - " & acctNum

    Dim buyLinesHTML As String, buyLinesPlain As String
    For i = 0 To buyCount - 1
        buyLinesHTML = buyLinesHTML & "<li>$" & Format(buyAmounts(i), "#,##0") & _
                       "&#9;<b>" & buyTickers(i) & "</b></li>" & vbCrLf
        buyLinesPlain = buyLinesPlain & "   * $" & Format(buyAmounts(i), "#,##0") & _
                        vbTab & buyTickers(i) & vbCrLf
    Next i

    Dim htmlBody As String
    htmlBody = "<p style='font-family:Calibri;font-size:11pt;'>" & salutation & "</p>" & vbCrLf
    htmlBody = htmlBody & "<p style='font-family:Calibri;font-size:11pt;'>"
    htmlBody = htmlBody & "Within the " & acctName & " - " & acctNum & ", please execute the following:</p>" & vbCrLf
    htmlBody = htmlBody & "<p style='font-family:Calibri;font-size:11pt;margin-bottom:4px;'><b>SELL:</b></p>" & vbCrLf
    htmlBody = htmlBody & "<ul style='font-family:Calibri;font-size:11pt;margin-top:0;'>" & vbCrLf & sellLinesHTML & "</ul>" & vbCrLf

    If buyCount > 0 Then
        htmlBody = htmlBody & "<p style='font-family:Calibri;font-size:11pt;margin-bottom:4px;'><b>BUY:</b></p>" & vbCrLf
        htmlBody = htmlBody & "<ul style='font-family:Calibri;font-size:11pt;margin-top:0;'>" & vbCrLf & buyLinesHTML & "</ul>" & vbCrLf
    End If

    If hasResidual Then
        htmlBody = htmlBody & "<p style='font-family:Calibri;font-size:11pt;'>"
        If buyCount > 0 Then
            htmlBody = htmlBody & "Remaining $" & Format(residual, "#,##0") & " of proceeds"
        Else
            htmlBody = htmlBody & "This will result in " & totalK & " of proceeds"
        End If
        htmlBody = htmlBody & " &ndash; please place into <b>" & residualTicker & "</b> (money market).</p>" & vbCrLf
    ElseIf usesRouting And routingParaHTML <> "" Then
        htmlBody = htmlBody & "<p style='font-family:Calibri;font-size:11pt;'>" & routingParaHTML & "</p>" & vbCrLf
    End If

    htmlBody = htmlBody & "<p style='font-family:Calibri;font-size:11pt;'>" & signoff & "</p>"

    Dim plainText As String
    plainText = salutation & vbCrLf & vbCrLf
    plainText = plainText & "Within the " & acctName & " - " & acctNum & ", please execute the following:" & vbCrLf & vbCrLf
    plainText = plainText & "SELL:" & vbCrLf & sellLinesPlain
    If buyCount > 0 Then plainText = plainText & vbCrLf & "BUY:" & vbCrLf & buyLinesPlain
    If hasResidual Then
        plainText = plainText & vbCrLf
        If buyCount > 0 Then
            plainText = plainText & "Remaining $" & Format(residual, "#,##0") & " of proceeds"
        Else
            plainText = plainText & "This will result in " & totalK & " of proceeds"
        End If
        plainText = plainText & " - please place into " & residualTicker & " (money market)." & vbCrLf
    ElseIf usesRouting And routingParaPlain <> "" Then
        plainText = plainText & vbCrLf & routingParaPlain & vbCrLf
    End If
    plainText = plainText & vbCrLf & signoff

    ' --- AUTOMATION MODE PREVIEW ---
    ' Skips the confirm dialog (already skipped above) and Outlook entirely;
    ' writes subject + full plain-text body to a preview sheet instead. Lets
    ' the headless harness (and any advisor who wants to eyeball the exact
    ' text first) see the generated email without Outlook running.
    If automationMode Then
        WriteEmailPreviewSheet ws.Parent, ws, subjectLine, plainText
        OfferSaveSnapshot
        Exit Sub
    End If

    ' --- OPEN IN OUTLOOK ---
    Dim olApp As Object, olMail As Object
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo ErrHandler

    If olApp Is Nothing Then
        Dim cb As Object
        Set cb = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        cb.SetText plainText
        cb.PutInClipboard
        Set cb = Nothing
        MsgBox "Outlook not available. Trade email copied to clipboard as plain text.", vbInformation
        Exit Sub
    End If

    Set olMail = olApp.CreateItem(0)
    With olMail
        .Subject = subjectLine
        If deskEmail <> "" Then .To = deskEmail
        .htmlBody = htmlBody
        .Display
    End With

    Set olMail = Nothing
    Set olApp = Nothing
    OfferSaveSnapshot
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation
End Sub

' ============================================================
' RESIDUAL DESTINATION PROMPT
' ============================================================

Private Function PromptResidualDestination(ws As Worksheet, totRow As Long, _
                                           residualAmt As Double, hasOtherBuys As Boolean, _
                                           defaultMM As String) As String
    Dim shortTickers() As String, shortFMVs() As Double, shortCount As Long
    ReDim shortTickers(0 To MAX_SHORT_POSITIONS - 1)
    ReDim shortFMVs(0 To MAX_SHORT_POSITIONS - 1)
    shortCount = 0

    Dim i As Long
    For i = totRow + 1 To totRow + 30
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = "SHORT" Then
            If shortCount < MAX_SHORT_POSITIONS Then
                shortTickers(shortCount) = Trim(CStr(ws.Cells(i, 3).Value))
                If IsNumeric(ws.Cells(i, 5).Value) Then shortFMVs(shortCount) = CDbl(ws.Cells(i, 5).Value)
                shortCount = shortCount + 1
            End If
        End If
    Next i

    Dim j As Long, tmpStr As String, tmpDbl As Double
    For i = 0 To shortCount - 2
        For j = i + 1 To shortCount - 1
            If shortFMVs(j) > shortFMVs(i) Then
                tmpDbl = shortFMVs(i): shortFMVs(i) = shortFMVs(j): shortFMVs(j) = tmpDbl
                tmpStr = shortTickers(i): shortTickers(i) = shortTickers(j): shortTickers(j) = tmpStr
            End If
        Next j
    Next i

    Dim residualK As String
    residualK = "$" & Format(residualAmt, "#,##0")

    Dim leadIn As String
    If hasOtherBuys Then
        leadIn = "After buy plan, " & residualK & " of proceeds remain." & vbCrLf & "Where should the residual go?"
    Else
        leadIn = "Routing " & residualK & " of proceeds to money market."
    End If

    Dim promptText As String, defaultDest As String
    If shortCount = 0 Then
        defaultDest = defaultMM
        promptText = leadIn & vbCrLf & vbCrLf & _
                     "No existing money market positions on this sheet." & vbCrLf & vbCrLf & _
                     "Type ticker, or accept default."
    ElseIf shortCount = 1 Then
        defaultDest = shortTickers(0)
        promptText = leadIn & vbCrLf & vbCrLf & _
                     "Existing position on this sheet:" & vbCrLf & _
                     "    " & shortTickers(0) & "   $" & Format(shortFMVs(0), "#,##0") & vbCrLf & vbCrLf & _
                     "Press OK to consolidate, or type a different ticker."
    Else
        defaultDest = shortTickers(0)
        promptText = leadIn & vbCrLf & vbCrLf & _
                     "Existing positions on this sheet (largest first):" & vbCrLf
        For i = 0 To shortCount - 1
            promptText = promptText & "    " & shortTickers(i) & "   $" & Format(shortFMVs(i), "#,##0") & vbCrLf
        Next i
        promptText = promptText & vbCrLf & "Default: " & defaultDest & " (largest). Press OK or type different."
    End If

    Dim destInput As String
    destInput = InputBox(promptText, "Money Market Destination (" & residualK & ")", defaultDest)
    If destInput = "" Then
        PromptResidualDestination = ""
    Else
        PromptResidualDestination = UCase(Trim(destInput))
    End If
End Function

' ============================================================
' HELPERS
' ============================================================

Private Function FindHeaderRow(ws As Worksheet) As Long
    Dim i As Long
    FindHeaderRow = 0
    For i = 1 To 5
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRow = i: Exit Function
        End If
    Next i
End Function

Private Function FindTotalRow(ws As Worksheet, dataStart As Long) As Long
    Dim i As Long
    FindTotalRow = 0
    For i = dataStart To dataStart + 200
        If Trim(CStr(ws.Cells(i, 1).Value)) = "" And ws.Cells(i, 5).Value <> "" Then
            FindTotalRow = i: Exit Function
        End If
    Next i
End Function

Private Function FindPivotGrandTotalRow(ws As Worksheet, startCol As Long, totRow As Long) As Long
    Dim r As Long
    FindPivotGrandTotalRow = 0
    For r = totRow + 2 To totRow + 50
        If UCase(Trim(CStr(ws.Cells(r, startCol).Value))) = "GRAND TOTAL" Then
            FindPivotGrandTotalRow = r: Exit Function
        End If
    Next r
End Function

Private Function CountScenarios(ws As Worksheet, startColIn As Long, strideIn As Long) As Long
    Dim i As Long, col As Long
    CountScenarios = 0
    For i = 0 To 30
        col = startColIn + i * strideIn
        If InStr(1, CStr(ws.Cells(2, col).Value), "Raise $", vbTextCompare) > 0 Then
            CountScenarios = i + 1
        Else
            Exit Function
        End If
    Next i
End Function
Private Function FindBuyPlanHeaderRow_Email(ws As Worksheet, startCol As Long, pivotEnd As Long) As Long
    Dim r As Long

    FindBuyPlanHeaderRow_Email = 0

    For r = pivotEnd + 1 To pivotEnd + 120
        If UCase(Trim(CStr(ws.Cells(r, startCol).Value))) = "BUY PLAN" Then
            FindBuyPlanHeaderRow_Email = r
            Exit Function
        End If
    Next r
End Function
Private Function ColLetter(colNum As Long) As String
    ColLetter = Split(Columns(colNum).Address(, False), ":")(0)
End Function

' ============================================================
' SPEC-FAITHFUL SELL NARRATION + PROCEEDS ROUTING (W10)
' ============================================================

' VBA's Format(v, "#,##0.##") leaves a trailing "." on whole numbers
' (e.g. 50 -> "50."). Drop the decimal places entirely when the spec
' amount (share count or percentage) has no fractional part.
Private Function FormatSpecAmtEmail(v As Double) As String
    If v = Int(v) Then
        FormatSpecAmtEmail = Format(v, "#,##0")
    Else
        FormatSpecAmtEmail = Format(v, "#,##0.##")
    End If
End Function

Private Function FindModeColEmail(ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Cells.Find(What:="Sell Mode", _
                          LookIn:=xlValues, _
                          LookAt:=xlWhole, _
                          SearchOrder:=xlByRows, _
                          SearchDirection:=xlNext, _
                          MatchCase:=False)
    On Error GoTo 0
    If Not c Is Nothing Then FindModeColEmail = c.Column
End Function

Private Function FindRoutingTitleCellEmail(ws As Worksheet) As Range
    On Error Resume Next
    Set FindRoutingTitleCellEmail = ws.Cells.Find(What:="PROCEEDS ROUTING", _
                                                  LookIn:=xlValues, _
                                                  LookAt:=xlWhole, _
                                                  SearchOrder:=xlByRows, _
                                                  SearchDirection:=xlNext, _
                                                  MatchCase:=False)
    On Error GoTo 0
End Function

' Narrates each PROCEEDS ROUTING row with Routed $ > tolerance into a single
' paragraph (HTML and plain-text in parallel). Buy Plan introduces the
' existing BUY: list ("the purchases below"); Money Market/Transfer
' Out/Hold in Cash each get their own sentence.
Private Sub BuildRoutingNarrative(ws As Worksheet, destCol As Long, dataStartRow As Long, _
                                  totalRaisedAmt As Double, ByRef paraHTML As String, ByRef paraPlain As String)
    Dim detailCol As Long, routedCol As Long
    detailCol = destCol + 1
    routedCol = destCol + 4

    Dim rr As Long, dest As String, detail As String, routed As Double, sentence As String
    paraHTML = ""
    paraPlain = ""

    For rr = dataStartRow To dataStartRow + ROUTING_DATA_ROWS - 1
        dest = Trim(CStr(ws.Cells(rr, destCol).Value))
        detail = Trim(CStr(ws.Cells(rr, detailCol).Value))
        routed = 0
        If IsNumeric(ws.Cells(rr, routedCol).Value) Then routed = CDbl(ws.Cells(rr, routedCol).Value)

        If routed > RESIDUAL_TOLERANCE Then
            sentence = ""
            Select Case dest
                Case "Buy Plan"
                    sentence = "Of the ~$" & Format(totalRaisedAmt, "#,##0") & " raised: $" & _
                               Format(routed, "#,##0") & " funds the purchases below."
                Case "Money Market"
                    sentence = "$" & Format(routed, "#,##0") & " remains in " & detail & " (money market)."
                Case "Transfer Out"
                    sentence = "$" & Format(routed, "#,##0") & " to be transferred out (" & detail & ")."
                Case "Hold in Cash"
                    sentence = "$" & Format(routed, "#,##0") & " holds as cash pending further instruction."
            End Select

            If sentence <> "" Then
                If paraPlain <> "" Then paraPlain = paraPlain & " "
                paraPlain = paraPlain & sentence
                If paraHTML <> "" Then paraHTML = paraHTML & " "
                paraHTML = paraHTML & sentence
            End If
        End If
    Next rr
End Sub

' AutomationMode=1 hook: writes subject (A1) + full plain-text body (A2, one
' cell) to a "CDS Email Preview" sheet instead of creating an Outlook item,
' then re-activates the source sheet.
Private Sub WriteEmailPreviewSheet(wb As Workbook, sourceWs As Worksheet, subjectLine As String, bodyText As String)
    Dim previewWs As Worksheet
    On Error Resume Next
    Set previewWs = wb.Sheets("CDS Email Preview")
    On Error GoTo 0

    If previewWs Is Nothing Then
        Set previewWs = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        previewWs.Name = "CDS Email Preview"
    Else
        previewWs.Cells.Clear
    End If

    previewWs.Cells(1, 1).Value = subjectLine
    previewWs.Cells(2, 1).Value = bodyText

    sourceWs.Activate
End Sub

' ============================================================
' POST-EMAIL SNAPSHOT OFFER (W9)
' ============================================================

' Fires after both the Outlook .Display success path and the
' AutomationMode preview success path. A snapshot is a values-only
' archive of the plan that was just emailed - useful months later
' ("how did we raise the 200k last time"). Never blocking: declining
' just skips it. Headless shims answer vbYes, so tests exercise the
' snapshot path here too.
Private Sub OfferSaveSnapshot()
    If MsgBox("Save a snapshot of this plan for the record?", vbYesNo + vbQuestion, "CDS Snapshot") = vbYes Then
        SaveCDSSnapshot
    End If
End Sub
