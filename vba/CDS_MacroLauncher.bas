Attribute VB_Name = "CDS_MacroLauncher"
Option Explicit

' ============================================================
' State-aware launcher/dispatcher for Ribbon and UserForm entrypoints.
' Keep macro safety rules in one place so every UI surface behaves alike.
' ============================================================

Private Const STATE_NO_WORKBOOK As String = "no_workbook"
Private Const STATE_MACRO_ACTIVE As String = "macro_active"
Private Const STATE_RAW As String = "raw"
Private Const STATE_PROCESSED_UNKNOWNS As String = "processed_unknowns"
Private Const STATE_PROCESSED_CLEAN As String = "processed_clean"
Private Const STATE_SCENARIOS As String = "scenarios"
Private Const STATE_BUY_PLANS As String = "buy_plans"
Private Const STATE_UNKNOWN As String = "unknown"

Public Function RunMacroById(ByVal macroId As String) As Boolean
    On Error GoTo ErrHandler

    Dim id As String
    id = NormalizeMacroId(macroId)

    Dim reason As String
    If Not CanRunMacroById(id, reason) Then
        MsgBox "Cannot run " & MacroLabelById(id) & "." & vbCrLf & vbCrLf & _
               reason, vbExclamation, "CDS Trade Assistant"
        Exit Function
    End If

    Dim previewText As String
    previewText = GetMacroPreviewById(id)

    If MsgBox(previewText & vbCrLf & vbCrLf & "Run this action now?", _
              vbYesNo + vbQuestion, MacroLabelById(id)) = vbNo Then
        Exit Function
    End If

    Select Case id
        Case "process_holdings"
            ProcessCDSHoldings

        Case "save_unknowns"
            SaveUnknownsAndRefresh

        Case "add_scenarios"
            PrepareCDSWorksheetForMacro ActiveSheet
            AddRaiseCashScenarios
            ApplyScenarioUXRulesToSheet ActiveSheet

        Case "plan_sells"
            PrepareCDSWorksheetForMacro ActiveSheet
            BuildSellWorkbench

        Case "build_routing"
            PrepareCDSWorksheetForMacro ActiveSheet
            BuildRoutingBlock

        Case "spawn_scenario"
            PrepareCDSWorksheetForMacro ActiveSheet
            SpawnScenario
            ApplyScenarioUXRulesToSheet ActiveSheet

        Case "remove_scenario"
            PrepareCDSWorksheetForMacro ActiveSheet
            RemoveScenario
            ApplyScenarioUXRulesToSheet ActiveSheet

        Case "add_buy_plans"
            PrepareCDSWorksheetForMacro ActiveSheet
            AddBuyPlans
            ApplyScenarioUXRulesToSheet ActiveSheet

        Case "cash_only_buy_plan"
            AddCashOnlyBuyPlan

        Case "generate_email"
            GenerateTradeEmail

        Case "refresh_prices"
            RefreshLivePrices

        Case "save_snapshot"
            SaveCDSSnapshot

        Case "export_snapshot"
            ExportSnapshotToFile

        Case "new_session"
            StartNewSession

        Case "open_settings"
            OpenSettings

        Case "close_settings"
            CloseSettings

        Case Else
            MsgBox "Unknown macro id: " & macroId, vbExclamation, "CDS Trade Assistant"
            Exit Function
    End Select

    RunMacroById = True
    Exit Function

ErrHandler:
    MsgBox MacroLabelById(macroId) & " failed: " & Err.Number & " - " & Err.Description, _
           vbExclamation, "CDS Trade Assistant"
End Function

Public Function CanRunMacroById(ByVal macroId As String, ByRef reason As String) As Boolean
    On Error GoTo FailClosed

    Dim id As String
    id = NormalizeMacroId(macroId)
    reason = ""

    If id = "open_settings" Or id = "close_settings" Then
        CanRunMacroById = True
        Exit Function
    End If

    Dim ws As Worksheet
    If Not ActiveClientSheet(ws, reason) Then Exit Function

    Dim state As String
    state = WorkflowStateCode(ws)

    Select Case id
        Case "process_holdings"
            If state = STATE_RAW Then
                CanRunMacroById = True
            Else
                reason = "Open or activate a raw CDS holdings export first."
            End If

        Case "save_unknowns"
            If state = STATE_PROCESSED_UNKNOWNS And CountScenariosLocal(ws) = 0 Then
                CanRunMacroById = True
            ElseIf CountScenariosLocal(ws) > 0 Then
                reason = "Scenarios already exist. Unknowns must be saved before scenarios are added."
            Else
                reason = "This sheet does not have unsaved unknown ticker classifications."
            End If

        Case "add_scenarios"
            If state = STATE_PROCESSED_CLEAN Then
                CanRunMacroById = True
            ElseIf CountUnknownRowsLocal(ws) > 0 Then
                reason = "Classify and save unknown tickers before adding scenarios."
            ElseIf CountScenariosLocal(ws) > 0 Then
                reason = "Scenarios already exist on this sheet."
            Else
                reason = "Process a CDS holdings export first."
            End If

        Case "plan_sells"
            If (state = STATE_PROCESSED_CLEAN Or state = STATE_SCENARIOS Or state = STATE_BUY_PLANS) And CountUnknownRowsLocal(ws) = 0 Then
                CanRunMacroById = True
            ElseIf CountUnknownRowsLocal(ws) > 0 Then
                reason = "Classify and save unknown tickers before planning sells."
            Else
                reason = "Process a CDS holdings export first."
            End If

        Case "spawn_scenario", "remove_scenario", "add_buy_plans", "generate_email"
            If (state = STATE_SCENARIOS Or state = STATE_BUY_PLANS) And CountUnknownRowsLocal(ws) = 0 Then
                CanRunMacroById = True
            ElseIf CountUnknownRowsLocal(ws) > 0 Then
                reason = "Classify and save unknown tickers before scenario or email actions."
            Else
                reason = "Add raise-cash scenarios first."
            End If

        Case "build_routing"
            If SellWorkbenchExistsLocal(ws) Then
                CanRunMacroById = True
            Else
                reason = "Create the sell workbench (Plan Sells) first."
            End If

        Case "cash_only_buy_plan"
            If IsProcessedReportLocal(ws) And CountUnknownRowsLocal(ws) = 0 Then
                CanRunMacroById = True
            ElseIf CountUnknownRowsLocal(ws) > 0 Then
                reason = "Classify and save unknown tickers before creating a cash-only buy plan."
            Else
                reason = "Process a CDS holdings export first."
            End If

        Case "refresh_prices"
            If state = STATE_PROCESSED_CLEAN Or state = STATE_SCENARIOS Or state = STATE_BUY_PLANS Then
                CanRunMacroById = True
            Else
                reason = "Process a CDS holdings export first."
            End If

        Case "save_snapshot"
            If Left$(ws.Name, 5) = "SNAP " Then
                reason = "This sheet is already a snapshot. Activate the live report sheet first."
            ElseIf IsProcessedReportLocal(ws) Then
                CanRunMacroById = True
            Else
                reason = "Activate a processed CDS report sheet first."
            End If

        Case "export_snapshot"
            If Left$(ws.Name, 5) = "SNAP " Then
                CanRunMacroById = True
            Else
                reason = "Activate a SNAP snapshot sheet first."
            End If

        Case "new_session"
            ' No workflow-state restriction: this is how a client file is
            ' born, so it must be runnable on an empty/raw/processed sheet
            ' alike. ActiveClientSheet above already enforced "a client
            ' workbook is active and it isn't PERSONAL.XLSB."
            CanRunMacroById = True

        Case Else
            reason = "Unknown macro id: " & macroId
    End Select

    Exit Function

FailClosed:
    reason = "The workbook state could not be checked safely: " & Err.Description
End Function

Public Function GetMacroPreviewById(ByVal macroId As String) As String
    Dim id As String
    id = NormalizeMacroId(macroId)

    Dim summary As String
    summary = GetWorkflowStateSummary()

    Select Case id
        Case "process_holdings"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: rebuild the active raw export into the standard CDS holdings report."

        Case "save_unknowns"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: validate every unknown ticker class, save valid mappings, then refresh the report."

        Case "add_scenarios"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: add the S1 pro-rata raise-cash scenario and scenario summary."

        Case "plan_sells"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: create or refresh a Sell Workbench where you type one target raise amount, mark rows Pool/Manual/Exclude, and review the proposed S2 sell plan."

        Case "build_routing"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: Adds/rebuilds the PROCEEDS ROUTING table for the plan scenario: declare how much of the raise funds buys, stays in money market, transfers out, or holds as cash."

        Case "spawn_scenario"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: add the next manual scenario block for custom sells."

        Case "remove_scenario"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: prompt for a scenario number, remove it, and renumber remaining scenarios."

        Case "add_buy_plans"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: add or refresh buy-plan blocks while preserving existing ticker/amount inputs."

        Case "cash_only_buy_plan"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: create or refresh the separate Cash Buy Plan worksheet."

        Case "generate_email"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: validate scenario trades and show the trade-email confirmation before Outlook."

        Case "refresh_prices"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: Checks live quotes against the export's implied prices on a CDS Live Prices sheet and flags tickers drifted past the alert threshold. Never modifies holdings values."

        Case "save_snapshot"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: Freezes this plan to a values-only, protected SNAP sheet and logs it in the CDS Snapshots index for future reference."

        Case "export_snapshot"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: Saves the active snapshot sheet as a standalone .xlsx under Snapshots\."

        Case "new_session"
            GetMacroPreviewById = summary & vbCrLf & vbCrLf & _
                "Action: Freezes the current report as a dated snapshot, then imports a fresh Client Center CSV and processes it as the new live report."

        Case "open_settings"
            GetMacroPreviewById = "Action: open the CDS settings sheet for ticker mappings and defaults."

        Case "close_settings"
            GetMacroPreviewById = "Action: hide settings, refresh cached mappings, and save the macro workbook."

        Case Else
            GetMacroPreviewById = "Unknown action."
    End Select
End Function

Public Function GetWorkflowStateSummary() As String
    On Error GoTo Fallback

    If Not HasActiveWorkbook() Then
        GetWorkflowStateSummary = "Workbook: none" & vbCrLf & "State: no workbook"
        Exit Function
    End If

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet

    Dim buyPlanText As String
    If BuyPlansExistLocal(ws) Then
        buyPlanText = "present"
    Else
        buyPlanText = "not present"
    End If

    GetWorkflowStateSummary = "Workbook: " & wb.Name & vbCrLf & _
                              "Sheet: " & ws.Name & vbCrLf & _
                              "Account: " & GetAccountLabelLocal(ws) & vbCrLf & _
                              "State: " & WorkflowStateLabel(WorkflowStateCode(ws)) & vbCrLf & _
                              "Unknowns: " & CountUnknownRowsLocal(ws) & vbCrLf & _
                              "Scenarios: " & CountScenariosLocal(ws) & vbCrLf & _
                              "Buy Plans: " & buyPlanText
    Exit Function

Fallback:
    GetWorkflowStateSummary = "State check failed: " & Err.Description
End Function

Public Function GetRecommendedMacroId() As String
    On Error GoTo Fallback

    If Not HasActiveWorkbook() Then Exit Function

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Select Case WorkflowStateCode(ws)
        Case STATE_RAW
            GetRecommendedMacroId = "process_holdings"
        Case STATE_PROCESSED_UNKNOWNS
            GetRecommendedMacroId = "save_unknowns"
        Case STATE_PROCESSED_CLEAN
            GetRecommendedMacroId = "plan_sells"
        Case STATE_SCENARIOS
            GetRecommendedMacroId = "add_buy_plans"
        Case STATE_BUY_PLANS
            If RoutingBlockExistsLocal(ws) Then
                GetRecommendedMacroId = "generate_email"
            Else
                GetRecommendedMacroId = "build_routing"
            End If
    End Select

Fallback:
End Function

Public Function MacroLabelById(ByVal macroId As String) As String
    Select Case NormalizeMacroId(macroId)
        Case "process_holdings": MacroLabelById = "Process Holdings"
        Case "save_unknowns": MacroLabelById = "Save Unknowns + Refresh"
        Case "add_scenarios": MacroLabelById = "Add Scenarios"
        Case "plan_sells": MacroLabelById = "Plan Sells"
        Case "build_routing": MacroLabelById = "Proceeds Routing"
        Case "spawn_scenario": MacroLabelById = "Spawn Scenario"
        Case "remove_scenario": MacroLabelById = "Remove Scenario"
        Case "add_buy_plans": MacroLabelById = "Add Buy Plans"
        Case "cash_only_buy_plan": MacroLabelById = "Cash-Only Buy Plan"
        Case "generate_email": MacroLabelById = "Generate Trade Email"
        Case "refresh_prices": MacroLabelById = "Refresh Live Prices"
        Case "save_snapshot": MacroLabelById = "Save Snapshot"
        Case "export_snapshot": MacroLabelById = "Export Snapshot"
        Case "new_session": MacroLabelById = "New Session (Import CSV)"
        Case "open_settings": MacroLabelById = "Open Settings"
        Case "close_settings": MacroLabelById = "Close Settings"
        Case Else: MacroLabelById = "Macro"
    End Select
End Function

Private Function ActiveClientSheet(ByRef ws As Worksheet, ByRef reason As String) As Boolean
    On Error GoTo Fail

    If Not HasActiveWorkbook() Then
        reason = "Open a client workbook first."
        Exit Function
    End If

    If IsMacroWorkbook(ActiveWorkbook) Then
        reason = "The macro workbook is active. Click into the client workbook first."
        Exit Function
    End If

    Set ws = ActiveSheet
    ActiveClientSheet = True
    Exit Function

Fail:
    reason = "Could not read the active workbook: " & Err.Description
End Function

Private Function WorkflowStateCode(ByVal ws As Worksheet) As String
    If Not HasActiveWorkbook() Then
        WorkflowStateCode = STATE_NO_WORKBOOK
    ElseIf IsMacroWorkbook(ActiveWorkbook) Then
        WorkflowStateCode = STATE_MACRO_ACTIVE
    ElseIf IsRawCDSHoldingsLocal(ws) Then
        WorkflowStateCode = STATE_RAW
    ElseIf IsProcessedReportLocal(ws) Then
        If CountUnknownRowsLocal(ws) > 0 Then
            WorkflowStateCode = STATE_PROCESSED_UNKNOWNS
        ElseIf BuyPlansExistLocal(ws) Then
            WorkflowStateCode = STATE_BUY_PLANS
        ElseIf CountScenariosLocal(ws) > 0 Then
            WorkflowStateCode = STATE_SCENARIOS
        Else
            WorkflowStateCode = STATE_PROCESSED_CLEAN
        End If
    Else
        WorkflowStateCode = STATE_UNKNOWN
    End If
End Function

Private Function WorkflowStateLabel(ByVal stateCode As String) As String
    Select Case stateCode
        Case STATE_NO_WORKBOOK: WorkflowStateLabel = "No workbook"
        Case STATE_MACRO_ACTIVE: WorkflowStateLabel = "Macro workbook active"
        Case STATE_RAW: WorkflowStateLabel = "Raw CDS export"
        Case STATE_PROCESSED_UNKNOWNS: WorkflowStateLabel = "Processed with unknowns"
        Case STATE_PROCESSED_CLEAN: WorkflowStateLabel = "Processed clean"
        Case STATE_SCENARIOS: WorkflowStateLabel = "Scenarios ready"
        Case STATE_BUY_PLANS: WorkflowStateLabel = "Buy plans ready"
        Case Else: WorkflowStateLabel = "Unrecognized sheet"
    End Select
End Function

Private Function IsRawCDSHoldingsLocal(ByVal ws As Worksheet) As Boolean
    On Error GoTo Done

    Dim r As Long
    For r = 1 To 10
        If LCase(Trim(CStr(ws.Cells(r, 1).Value))) = "accountname" Then
            IsRawCDSHoldingsLocal = True
            Exit Function
        End If

        If LCase(Trim(CStr(ws.Cells(r, 3).Value))) = "description" Then
            IsRawCDSHoldingsLocal = True
            Exit Function
        End If
    Next r

Done:
End Function

Private Function IsProcessedReportLocal(ByVal ws As Worksheet) As Boolean
    IsProcessedReportLocal = (FindHeaderRowLocal(ws) > 0)
End Function

Private Function CountUnknownRowsLocal(ByVal ws As Worksheet) As Long
    On Error GoTo Done

    If Not IsProcessedReportLocal(ws) Then GoTo Done

    Dim dataStart As Long
    Dim totalRow As Long
    Dim r As Long

    dataStart = FindHeaderRowLocal(ws) + 1
    totalRow = FindTotalRowLocal(ws, dataStart)

    If dataStart <= 1 Or totalRow = 0 Then GoTo Done

    For r = dataStart To totalRow - 1
        If Trim(CStr(ws.Cells(r, 1).Value)) = "???" Then
            CountUnknownRowsLocal = CountUnknownRowsLocal + 1
        End If
    Next r

Done:
End Function

Private Function CountScenariosLocal(ByVal ws As Worksheet) As Long
    On Error GoTo Fallback

    Dim i As Long
    Dim col As Long

    For i = 0 To 30
        col = ScenStartCol() + i * ScenStride()

        If InStr(1, CStr(ws.Cells(2, col).Value), "Raise $", vbTextCompare) > 0 Then
            CountScenariosLocal = i + 1
        Else
            Exit Function
        End If
    Next i

    Exit Function

Fallback:
    CountScenariosLocal = 0
End Function

Private Function BuyPlansExistLocal(ByVal ws As Worksheet) As Boolean
    On Error GoTo Done

    Dim dataStart As Long
    Dim totalRow As Long
    Dim r As Long
    Dim c As Long
    Dim txt As String

    If IsProcessedReportLocal(ws) Then
        dataStart = FindHeaderRowLocal(ws) + 1
        totalRow = FindTotalRowLocal(ws, dataStart)

        If totalRow > 0 Then
            For r = totalRow + 1 To totalRow + 220
                For c = 1 To 90
                    txt = UCase(Trim(CStr(ws.Cells(r, c).Value)))
                    If txt = "BUY PLAN" Or txt = "CASH-ONLY BUY PLAN" Then
                        BuyPlansExistLocal = True
                        Exit Function
                    End If
                Next c
            Next r
        End If
    End If

    Dim planWs As Worksheet
    On Error Resume Next
    Set planWs = ActiveWorkbook.Worksheets("Cash Buy Plan")
    On Error GoTo Done

    If Not planWs Is Nothing Then
        BuyPlansExistLocal = (UCase(Trim(CStr(planWs.Range("A1").Value))) = "CASH-ONLY BUY PLAN")
    End If

Done:
End Function

Private Function GetAccountLabelLocal(ByVal ws As Worksheet) As String
    On Error GoTo Fallback

    If IsProcessedReportLocal(ws) Then
        GetAccountLabelLocal = Trim(CStr(ws.Cells(1, 2).Value)) & " / " & Trim(CStr(ws.Cells(1, 1).Value))
        If Replace(GetAccountLabelLocal, "/", "") <> "" Then Exit Function
    End If

Fallback:
    GetAccountLabelLocal = "-"
End Function

Private Function SellWorkbenchExistsLocal(ByVal ws As Worksheet) As Boolean
    SellWorkbenchExistsLocal = Not (FindCellExactLocal(ws, "Sell Mode") Is Nothing)
End Function

Private Function RoutingBlockExistsLocal(ByVal ws As Worksheet) As Boolean
    RoutingBlockExistsLocal = Not (FindCellExactLocal(ws, "PROCEEDS ROUTING") Is Nothing)
End Function

Private Function FindCellExactLocal(ByVal ws As Worksheet, ByVal textValue As String) As Range
    On Error Resume Next
    Set FindCellExactLocal = ws.Cells.Find(What:=textValue, _
                                           LookIn:=xlValues, _
                                           LookAt:=xlWhole, _
                                           SearchOrder:=xlByRows, _
                                           SearchDirection:=xlNext, _
                                           MatchCase:=False)
    On Error GoTo 0
End Function

Private Function FindHeaderRowLocal(ByVal ws As Worksheet) As Long
    Dim r As Long

    For r = 1 To 5
        If UCase(Trim(CStr(ws.Cells(r, 1).Value))) = "ASSET CLASS" Then
            FindHeaderRowLocal = r
            Exit Function
        End If
    Next r
End Function

Private Function FindTotalRowLocal(ByVal ws As Worksheet, ByVal dataStart As Long) As Long
    Dim r As Long

    For r = dataStart To dataStart + 300
        If Trim(CStr(ws.Cells(r, 1).Value)) = "" And ws.Cells(r, 5).Value <> "" Then
            FindTotalRowLocal = r
            Exit Function
        End If
    Next r
End Function

Private Function NormalizeMacroId(ByVal macroId As String) As String
    NormalizeMacroId = LCase(Trim(CStr(macroId)))
End Function

Private Function HasActiveWorkbook() As Boolean
    On Error GoTo Nope

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    HasActiveWorkbook = Not wb Is Nothing
    Exit Function

Nope:
    HasActiveWorkbook = False
End Function

Private Function IsMacroWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo Nope

    If wb Is ThisWorkbook Then
        IsMacroWorkbook = True
    ElseIf LCase(wb.Name) = "personal.xlsb" Then
        IsMacroWorkbook = True
    End If

Nope:
End Function
