Attribute VB_Name = "CDS_RibbonCallbacks"
Option Explicit

Public Sub RunTradeAssistant(ctl As IRibbonControl)
    OpenCDSTradeAssistant
End Sub

Public Sub RunProcess(ctl As IRibbonControl)
    RunMacroById "process_holdings"
End Sub

Public Sub RunSaveUnknowns(ctl As IRibbonControl)
    RunMacroById "save_unknowns"
End Sub

Public Sub RunAddScenarios(ctl As IRibbonControl)
    RunMacroById "add_scenarios"
End Sub

Public Sub RunPlanSells(ctl As IRibbonControl)
    RunMacroById "plan_sells"
End Sub

Public Sub RunBuildRouting(ctl As IRibbonControl)
    RunMacroById "build_routing"
End Sub

Public Sub RunSpawn(ctl As IRibbonControl)
    RunMacroById "spawn_scenario"
End Sub

Public Sub RunRemove(ctl As IRibbonControl)
    RunMacroById "remove_scenario"
End Sub

Public Sub RunBuyPlans(ctl As IRibbonControl)
    RunMacroById "add_buy_plans"
End Sub

Public Sub RunCashOnlyBuyPlan(ctl As IRibbonControl)
    RunMacroById "cash_only_buy_plan"
End Sub

Public Sub RunEmail(ctl As IRibbonControl)
    RunMacroById "generate_email"
End Sub

Public Sub RunRefreshPrices(ctl As IRibbonControl)
    RunMacroById "refresh_prices"
End Sub

Public Sub RunSaveSnapshot(ctl As IRibbonControl)
    RunMacroById "save_snapshot"
End Sub

Public Sub RunExportSnapshot(ctl As IRibbonControl)
    RunMacroById "export_snapshot"
End Sub

Public Sub RunOpenSettings(ctl As IRibbonControl)
    RunMacroById "open_settings"
End Sub

Public Sub RunCloseSettings(ctl As IRibbonControl)
    RunMacroById "close_settings"
End Sub
