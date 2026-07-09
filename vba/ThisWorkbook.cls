VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Workbook_Open()
    Set App = Application
End Sub

Public Sub EnableContextHandler()
    Set App = Application
    MsgBox "Context handler enabled. Click any row in a processed sheet to update scenario titles.", vbInformation, "CDS Context"
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim prevEnableEvents As Boolean
    prevEnableEvents = Application.EnableEvents

    On Error GoTo SafeExit
    
    ' Only act on sheets that have been processed by Macro 1
    If UCase(Trim(CStr(Sh.Cells(2, 1).Value))) <> "ASSET CLASS" Then Exit Sub
    
    ' Only react to clicks in plausible data rows (avoid headers, totals row, etc.)
    Dim activeRow As Long
    activeRow = Target.Row
    If activeRow < 3 Or activeRow > 200 Then Exit Sub
    
    ' Skip if the user clicked an empty row (no description)
    If Trim(CStr(Sh.Cells(activeRow, 2).Value)) = "" Then Exit Sub
    
    ' Update L1 helper cell. Suppress recursion.
    Application.EnableEvents = False
    Sh.Cells(1, 12).Value = activeRow
    Application.EnableEvents = prevEnableEvents

    Exit Sub

SafeExit:
    On Error Resume Next
    Application.EnableEvents = prevEnableEvents
End Sub

