VERSION 5.00
Begin VB.Form frmInvite 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invite a Computer..."
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdInvite 
      Caption         =   "Invite"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox lstComputers 
      Height          =   2985
      ItemData        =   "frmInvite.frx":0000
      Left            =   0
      List            =   "frmInvite.frx":0007
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdInvite_Click()
Call Invite
End Sub

Private Sub cmdRefresh_Click()
frmMain.RefreshNetwork True
End Sub

Private Sub Form_Load()
Call FormLoad(Me)
cmdRefresh_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FormLoad Me, True
End Sub

Private Sub lstComputers_Click()
If lstComputers.ListIndex <> -1 Then
    cmdInvite.Enabled = True
Else
    cmdInvite.Enabled = False
End If
End Sub

Public Sub Invite()

Dim Ans As VbMsgBoxResult
Dim Name As String, IPHost As String

If lstComputers.ListIndex = -1 Then
    'Me.Hide
    AddText "Select a computer to invite", TxtError, True
    Unload Me
    Exit Sub
End If

Me.Hide

Name = lstComputers.Text

'On Error GoTo 0

With frmMain
    
    Ans = .Question("Invite " & Name & "?", .mnuFileInvite)
    
    If Ans = vbYes Then
        'If .SockAr(0) Is Nothing Then Load .SockAr(0)
        
        .SockAr(0).Close
        
        .SockAr(0).RemoteHost = Name
        .SockAr(0).RemotePort = RPort
        .SockAr(0).LocalPort = LPort - 1
        
        'On Error Resume Next
        .SockAr(0).Connect
        
        'Pause 1500
        
        If Server Then
            IPHost = .SckLC.LocalHostName
        Else
            IPHost = .SckLC.RemoteHost
        End If
        
        'Pause 10
        
        '.SockAr(0).SendData eCommands.Invite & IPHost
        
        Dim TimeOut As Integer
        
        Do While Not modVars.Closing
            If .SockAr(0).State = sckConnected Then
                Exit Do
            ElseIf .SockAr(0).State = sckError Then
                Exit Do
            ElseIf TimeOut > 300 Then '30 seconds
                Exit Do
            End If
            Pause 10
            TimeOut = TimeOut + 1
        Loop
        
        If TimeOut > 300 Then
            AddText "Invite Connection Timed Out", TxtError, True
        ElseIf .SockAr(0).State = sckError Then
            AddText "Error " & Err.Description, TxtError, True
        Else
            SendData eCommands.Invite & IPHost, 0
            
            AddText "Invited " & Name, , True
            
            Pause 3500
            
        End If
        
        .SockAr(0).Close
        
    End If
    
End With

Unload Me
End Sub
