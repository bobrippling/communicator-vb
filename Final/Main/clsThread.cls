VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
    ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Public Sub NewLoginFormThread() 'ObPtr As Long, MainFormHandle As Long)

'The interesting part is here
'We store the WindowHandle of the MainWindow and a pointer
'to the Main Window Form by using the SetProp function against
'The Hwnd of the New Form as a virtual property

Dim Frm As frmLogin

'AddConsoleText "Setting Frm = New frmLogin..."

Set Frm = New frmLogin

'AddConsoleText "Showing Form..."

Frm.Show 'vbModeless, frmMain

'AddConsoleText "Setting Properties..."
'SetProp Frm.hWnd, "MAINHWND", MainFormHandle
'SetProp Frm.hWnd, "OBPTR", ObPtr

'AddConsoleText "Removing Form..."

Set Frm = Nothing

End Sub
