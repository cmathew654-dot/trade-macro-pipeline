Attribute VB_Name = "TestShims"
' =====================================================================
' TestShims - headless-run doubles for MsgBox / InputBox.
' Imported ONLY by the test harness (tests/run_pipeline.py); never ship
' this module in the production add-in. A same-project public function
' shadows VBA.Interaction for unqualified calls, so product code runs
' unmodified while dialogs are answered deterministically.
'
' MsgBox  -> logs prompt to TestLog sheet; answers vbYes to Yes/No
'            styles, vbOK otherwise.
' InputBox-> logs prompt; pops the next queued answer from the
'            TestInputs sheet (col A, pointer in B1); falls back to
'            the caller's Default.
' =====================================================================
Option Explicit

Public Function MsgBox(ByVal Prompt As String, _
                       Optional ByVal Buttons As Long = 0, _
                       Optional ByVal Title As String = "") As VbMsgBoxResult
    LogShim "MsgBox", Prompt
    Dim style As Long
    style = Buttons And 7
    If style = 3 Or style = 4 Then          ' vbYesNoCancel / vbYesNo
        MsgBox = vbYes
    ElseIf style = 1 Or style = 5 Then      ' vbOKCancel / vbRetryCancel
        MsgBox = vbOK
    Else
        MsgBox = vbOK
    End If
End Function

Public Function InputBox(ByVal Prompt As String, _
                         Optional ByVal Title As String = "", _
                         Optional ByVal Default As String = "") As String
    LogShim "InputBox", Prompt
    InputBox = NextTestInput(Default)
End Function

Private Function NextTestInput(ByVal fallback As String) As String
    Dim ws As Worksheet, ptr As Long, v As String
    On Error GoTo UseFallback
    Set ws = ThisWorkbook.Worksheets("TestInputs")
    ptr = CLng(Val(ws.Range("B1").Value))
    If ptr < 1 Then ptr = 1
    v = CStr(ws.Cells(ptr, 1).Value)
    If Len(v) = 0 Then GoTo UseFallback
    ws.Range("B1").Value = ptr + 1
    NextTestInput = v
    Exit Function
UseFallback:
    NextTestInput = fallback
End Function

Private Sub LogShim(ByVal kind As String, ByVal Prompt As String)
    On Error Resume Next
    Dim ws As Worksheet, r As Long, prevSheet As Object
    Set prevSheet = ActiveSheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets("TestLog")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "TestLog"
        If Not prevSheet Is Nothing Then prevSheet.Activate
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = kind
    ws.Cells(r, 2).Value = Left$(Prompt, 900)
End Sub
