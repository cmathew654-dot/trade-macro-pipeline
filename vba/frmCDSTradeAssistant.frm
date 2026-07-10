VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCDSTradeAssistant 
   Caption         =   "CDS Trade Assistant"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   OleObjectBlob   =   "frmCDSTradeAssistant.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCDSTradeAssistant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' State-aware modeless UI for the CDS Trade Assistant.
' All macro execution runs through CDS_MacroLauncher.RunMacroById.
' ============================================================

Private mBuilt As Boolean
Private mHandlers As Collection
Private mPrimaryMacroId As String

Private Sub UserForm_Initialize()
    Me.Caption = "CDS Trade Assistant"
    Me.Width = 535
    Me.Height = 763
    Me.BackColor = RGB(22, 27, 31)

    Set mHandlers = New Collection

    BuildAssistantUI
    RefreshAssistantStatus
End Sub

Private Sub BuildAssistantUI()
    If mBuilt Then Exit Sub
    mBuilt = True

    AddLabel "lblTitle", "CDS Trade Assistant", 16, 12, 485, 24, True, 15, RGB(255, 255, 255)
    AddLabel "lblSubtitle", "One controlled path from raw holdings to trade email.", 16, 36, 485, 18, False, 9, RGB(180, 197, 207)

    AddStatusLabel "lblStatus", "Status loading...", 16, 66, 485, 98
    AddLabel "lblNext", "Next: -", 16, 174, 485, 20, True, 10, RGB(255, 222, 128)

    AddButton "cmdPrimary", "Run Next Step", "primary", 16, 202, 485, 42, True
    AddButton "cmdRefresh", "Refresh Status", "refresh", 16, 254, 485, 24, False

    AddLabel "lblStart", "START", 16, 292, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdProcess", "Process Holdings", "process_holdings", 16, 316, 235, 30, False

    AddLabel "lblUnknowns", "UNKNOWNS", 266, 292, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdSaveUnknowns", "Save Unknowns + Refresh", "save_unknowns", 266, 316, 235, 30, False

    AddLabel "lblScenarios", "SCENARIOS", 16, 360, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdPlanSells", "Plan Sells", "plan_sells", 16, 384, 235, 28, False
    AddButton "cmdBuildRouting", "Proceeds Routing", "build_routing", 16, 418, 235, 28, False
    AddButton "cmdAddScenarios", "Add Scenarios", "add_scenarios", 16, 452, 112, 28, False
    AddButton "cmdSpawnScenario", "Spawn Scenario", "spawn_scenario", 139, 452, 112, 28, False
    AddButton "cmdRemoveScenario", "Remove Scenario", "remove_scenario", 16, 486, 235, 28, False

    AddLabel "lblBuyPlans", "BUY PLANS", 266, 360, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdAddBuyPlans", "Add Buy Plans", "add_buy_plans", 266, 384, 235, 28, False
    AddButton "cmdCashOnly", "Cash-Only Plan", "cash_only_buy_plan", 266, 418, 235, 28, False

    AddLabel "lblEmail", "EMAIL", 16, 536, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdEmail", "Generate Trade Email", "generate_email", 16, 560, 235, 30, False

    AddLabel "lblSettings", "SETTINGS", 266, 536, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdOpenSettings", "Open Settings", "open_settings", 266, 560, 112, 30, False
    AddButton "cmdCloseSettings", "Close Settings", "close_settings", 389, 560, 112, 30, False

    AddLabel "lblReference", "REFERENCE", 266, 602, 235, 18, True, 8, RGB(180, 197, 207)
    AddButton "cmdRefreshPrices", "Refresh Live Prices", "refresh_prices", 266, 626, 235, 28, False
    AddButton "cmdSaveSnapshot", "Save Snapshot", "save_snapshot", 266, 660, 112, 28, False
    AddButton "cmdExportSnapshot", "Export Snapshot", "export_snapshot", 389, 660, 112, 28, False

    AddButton "cmdClose", "Close", "close", 16, 702, 485, 28, False
End Sub

Private Sub AddStatusLabel(controlName As String, captionText As String, _
                           leftPos As Single, topPos As Single, _
                           widthVal As Single, heightVal As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", controlName, True)

    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .ForeColor = RGB(233, 240, 244)
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(32, 41, 48)
        .WordWrap = True
    End With
End Sub

Private Sub AddLabel(controlName As String, captionText As String, _
                     leftPos As Single, topPos As Single, _
                     widthVal As Single, heightVal As Single, _
                     isBold As Boolean, fontSize As Single, foreColor As Long)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", controlName, True)

    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Name = "Segoe UI"
        .Font.Bold = isBold
        .Font.Size = fontSize
        .ForeColor = foreColor
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
    End With
End Sub

Private Sub AddButton(controlName As String, captionText As String, actionName As String, _
                      leftPos As Single, topPos As Single, _
                      widthVal As Single, heightVal As Single, _
                      isBold As Boolean)
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", controlName, True)

    With btn
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Name = "Segoe UI"
        .Font.Bold = isBold
        .Font.Size = 9
    End With

    Dim h As CDS_ButtonHandler
    Set h = New CDS_ButtonHandler
    Set h.btn = btn
    h.actionName = actionName
    Set h.ParentForm = Me
    mHandlers.Add h
End Sub

Public Sub HandleButtonClick(actionName As String)
    Select Case LCase(actionName)
        Case "refresh"
            RefreshAssistantStatus

        Case "primary"
            ExecuteMacroFromUI mPrimaryMacroId

        Case "close"
            Unload Me

        Case Else
            ExecuteMacroFromUI actionName
    End Select
End Sub

Private Sub ExecuteMacroFromUI(ByVal macroId As String)
    If Trim(macroId) = "" Then
        MsgBox "No safe next action is available for the active workbook.", vbInformation, "CDS Trade Assistant"
        RefreshAssistantStatus
        Exit Sub
    End If

    RunMacroById macroId
    RefreshAssistantStatus
End Sub

Private Sub RefreshAssistantStatus()
    On Error GoTo ErrHandler

    Me.Controls("lblStatus").Caption = GetWorkflowStateSummary()

    mPrimaryMacroId = GetRecommendedMacroId()

    If mPrimaryMacroId <> "" Then
        Me.Controls("lblNext").Caption = "Next: " & MacroLabelById(mPrimaryMacroId)
        Me.Controls("cmdPrimary").Caption = "Run: " & MacroLabelById(mPrimaryMacroId)
    Else
        Me.Controls("lblNext").Caption = "Next: choose an enabled action below"
        Me.Controls("cmdPrimary").Caption = "No Safe Next Step"
    End If

    UpdateButtonState "cmdPrimary", mPrimaryMacroId
    UpdateButtonState "cmdProcess", "process_holdings"
    UpdateButtonState "cmdSaveUnknowns", "save_unknowns"
    UpdateButtonState "cmdPlanSells", "plan_sells"
    UpdateButtonState "cmdBuildRouting", "build_routing"
    UpdateButtonState "cmdAddScenarios", "add_scenarios"
    UpdateButtonState "cmdSpawnScenario", "spawn_scenario"
    UpdateButtonState "cmdRemoveScenario", "remove_scenario"
    UpdateButtonState "cmdAddBuyPlans", "add_buy_plans"
    UpdateButtonState "cmdCashOnly", "cash_only_buy_plan"
    UpdateButtonState "cmdEmail", "generate_email"
    UpdateButtonState "cmdOpenSettings", "open_settings"
    UpdateButtonState "cmdCloseSettings", "close_settings"
    UpdateButtonState "cmdRefreshPrices", "refresh_prices"
    UpdateButtonState "cmdSaveSnapshot", "save_snapshot"
    UpdateButtonState "cmdExportSnapshot", "export_snapshot"

    StyleUtilityButton "cmdRefresh"
    StyleUtilityButton "cmdClose"
    Exit Sub

ErrHandler:
    Me.Controls("lblStatus").Caption = "Status check failed: " & Err.Description
    Me.Controls("lblNext").Caption = "Next: refresh status"
End Sub

Private Sub UpdateButtonState(ByVal controlName As String, ByVal macroId As String)
    On Error Resume Next

    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls(controlName)

    If Trim(macroId) = "" Then
        btn.Enabled = False
        StyleDisabledButton btn
        Exit Sub
    End If

    Dim reason As String
    Dim canRun As Boolean
    canRun = CanRunMacroById(macroId, reason)

    btn.Enabled = canRun

    If canRun Then
        StyleActionButton btn, (controlName = "cmdPrimary")
    Else
        StyleDisabledButton btn
    End If
End Sub

Private Sub StyleActionButton(ByVal btn As MSForms.CommandButton, ByVal isPrimary As Boolean)
    If isPrimary Then
        btn.BackColor = RGB(0, 139, 139)
        btn.ForeColor = RGB(255, 255, 255)
        btn.Font.Bold = True
    Else
        btn.BackColor = RGB(44, 62, 73)
        btn.ForeColor = RGB(235, 243, 247)
    End If
End Sub

Private Sub StyleUtilityButton(ByVal controlName As String)
    On Error Resume Next

    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls(controlName)
    btn.Enabled = True
    btn.BackColor = RGB(67, 77, 86)
    btn.ForeColor = RGB(245, 245, 245)
End Sub

Private Sub StyleDisabledButton(ByVal btn As MSForms.CommandButton)
    btn.BackColor = RGB(52, 55, 58)
    btn.ForeColor = RGB(138, 145, 150)
End Sub
