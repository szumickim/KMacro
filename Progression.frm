VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progression 
   Caption         =   "Progression"
   ClientHeight    =   1392
   ClientLeft      =   -210
   ClientTop       =   -1140
   ClientWidth     =   3495
   OleObjectBlob   =   "Progression.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub USERFORM_Initialize()

    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
    
    Dim ret As Long
    Dim formHWnd As Long
    
    'Get window handle of the userform
    
    formHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, Me.Caption)
    If formHWnd = 0 Then
        Debug.Print Err.LastDllError
    End If

    'Set userform window to 'always on top'
    
    ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    If ret = 0 Then
        Debug.Print Err.LastDllError
    End If
    
End Sub
