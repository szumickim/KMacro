Attribute VB_Name = "Module2"
Option Explicit

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare PtrSafe Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal uFlags As Long) As Long

Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Sub progress(ByVal i As Long, ByVal lLast As Long)
 
 Dim pctCompl As Single
 
 pctCompl = Round(i / lLast * 100, 0)
 Progression.Label1.Caption = i & " / " & lLast
 Progression.Text.Caption = pctCompl & "% Completed"
 Progression.Bar.Width = pctCompl * 2

 DoEvents 'update the userform

End Sub
