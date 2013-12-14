VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBackGround"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" ( _
    ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, _
    ByVal wType As Long) As Long


Public Function ShowAForm()
Dim f As New frmLogin
f.Show

'Set f = Nothing
'added =nothing

AddConsoleText "Form Shown"

End Function

