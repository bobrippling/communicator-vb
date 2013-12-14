VERSION 5.00
Begin VB.Form frmDevCmd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dev Commands"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkHide 
      Caption         =   "Hide my command"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CheckBox chkClear 
      Caption         =   "Clear command box when a command is sent"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CommandButton cmdDevSend 
      Caption         =   "Send"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtDev 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5055
   End
   Begin VB.Frame fraDev 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.ComboBox cboDevCmd 
         Height          =   315
         ItemData        =   "frmDevCmd.frx":0000
         Left            =   2160
         List            =   "frmDevCmd.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtSendTo 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "Send to: "
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
End
Attribute VB_Name = "frmDevCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msDevDelay As Long = 100 'for typing

Private Sub cmdDevSend_Click()
Dim dMsg As String, SendTo As String, CmdNo As String
Static LastTick As Long

If (LastTick + msDevDelay) < GetTickCount() Then
    On Error Resume Next
    SendTo = Trim$(Right$(txtSendTo.Text, Len(txtSendTo.Text) - 9))
    On Error GoTo 0
    
    
    CmdNo = cboDevCmd.ItemData(cboDevCmd.ListIndex)
    'If Left$(cboDevCmd.Text, 1) = "-" Then
        'cmdno = Left$(cboDevCmd.Text, 2)
    'Else
        'CmdNo = 'Left$(cboDevCmd.Text, 1)
    'End If
    
    If LenB(SendTo) = 0 And CmdNo <> NoFilter Then 'Left$(cboDevCmd.Text, 1) <> "0" Then
        'AddText "Please Select a computer to send to", TxtError, True
        lblInfo.Caption = "Send to...?"
        Exit Sub
    End If
    
    frmMain.SendDevCmd CmdNo, SendTo, txtDev.Text, CBool(chkHide.Value)
    
    With txtDev
        If chkClear.Value = 1 Then
            .Text = vbNullString
        Else
            .Selstart = 0
            .Sellength = Len(.Text)
        End If
        SetFocus2 txtDev
    End With
    
    LastTick = GetTickCount()
    
    lblInfo.Caption = vbNullString
Else
    lblInfo.Caption = "Don't spam - Not nice"
End If

End Sub

Private Sub txtDev_Change()

With txtDev
    'cmdDevSend.Enabled = (LenB(.Text) <> 0)
    If LenB(.Text) Then
        cmdDevSend.Default = True
    Else
        cmdDevSend.Default = False
    End If
End With

End Sub

Private Sub txtSendTo_Change() 'Optional ByVal Ignore As Boolean = False)
Const TheCap As String = "Send to: "
Dim Text As String

With txtSendTo
    '(Len(txtSendTo.Text) < 8 Or txtSendTo.SelStart < 8) Then
    On Error Resume Next
    Text = Left$(.Text, 9)
    On Error GoTo 0
    
    If Text <> TheCap Then
        txtSendTo.Text = TheCap
        txtSendTo.Selstart = Len(txtSendTo.Text)
    End If
End With

'Private Sub txtSendTo_KeyPress(KeyAscii As Integer)
'If Len(txtSendTo.Text) <= 8 Then
    'If KeyAscii = 8 Then KeyAscii = 0
'End If
'End Sub

End Sub

Private Sub txtSendTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    If Len(txtSendTo.Text) <= 9 Then KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
Dim i As eDevCmds
Dim iClient As Integer
Dim canDev As Boolean

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetTextBoxBanner txtDev.hWnd, "Send Dev Command"
End If

bDevCmdFormLoaded = True

cboDevCmd.Clear
canDev = modDev.devCanDo(modDev.Dev_Level_Heightened)

For i = 9 To -5 Step -1
    If modDev.DevCmdAllowed(i) Or canDev Then
        cboDevCmd.AddItem modDev.GetDevCommandName(i)
        
        cboDevCmd.ItemData(cboDevCmd.NewIndex) = CInt(i)
    End If
Next i

cboDevCmd.ListIndex = 2 + Abs(canDev)
'                          ^ add 1 since an extra command is in the way

'cboDevCmd.AddItem CStr(eDevCmds.NoFilter) & " - No Filter"
'cboDevCmd.AddItem CStr(eDevCmds.dBeep) & " - Beep"
'cboDevCmd.AddItem CStr(eDevCmds.CmdPrompt) & " - Command Prompt"
'cboDevCmd.AddItem CStr(eDevCmds.ClpBrd) & " - Clipboard"
'cboDevCmd.AddItem CStr(eDevCmds.Visible) & " - Visible"
'cboDevCmd.AddItem CStr(eDevCmds.Shel) & " - Shell"
'cboDevCmd.AddItem CStr(eDevCmds.Name) & " - Name"
'cboDevCmd.AddItem CStr(eDevCmds.Version) & " - Version"
'cboDevCmd.AddItem CStr(eDevCmds.Disco) & " - Disconnect"
'cboDevCmd.AddItem CStr(eDevCmds.CompName) & " - Computer Name"
'cboDevCmd.AddItem CStr(eDevCmds.GameForm) & " - Open Game Window"
'cboDevCmd.AddItem CStr(eDevCmds.Caps) & " - CapsLock"
'cboDevCmd.AddItem CStr(eDevCmds.Script) & " - VBScript"
'cboDevCmd.ListIndex = 6

If canDev Then
    chkHide.Value = 1
    chkHide.Visible = True
Else
    chkHide.Value = 0
End If

lblInfo.Caption = vbNullString


iClient = FindClient(frmMain.GetSelectedClient())
If iClient > -1 Then
    On Error Resume Next
    txtSendTo.Text = "Send to: " & Clients(iClient).sName
End If


Call FormLoad(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

bDevCmdFormLoaded = False

Call FormLoad(Me, True)
End Sub
