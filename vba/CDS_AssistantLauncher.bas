Attribute VB_Name = "CDS_AssistantLauncher"
Option Explicit

' ============================================================
' CDS Assistant Launcher
'
' Opens the CDS Trade Assistant UserForm.
' Run from:
'   Alt+F8 > OpenCDSTradeAssistant
' ============================================================

Public Sub OpenCDSTradeAssistant()
    frmCDSTradeAssistant.Show vbModeless
End Sub
