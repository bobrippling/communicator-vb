VERSION 5.00
Begin VB.Form frmBug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bug Alert!"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5895
   Begin VB.CommandButton cmdSend 
      Caption         =   "Done, Ok, Send!"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtBody 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label lblHelp 
      Caption         =   "Your name will be added to the end. So no spammage"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label lblBody 
      Caption         =   "Put some details in here. Not just any details, those about the bug are sufficent"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sEmailFrom = "microbsoft1@gmail.com", sEmailTo = "microbsoft0@gmail.com", sPass = "abucrabsa"
Const sServer = "smtp.gmail.com"

'### Not to be changed on pain of death ###
Private Function SendMail(ByVal sBody As String, ByVal sSubject As String) As Long
Dim oMsg As Object 'CDO.Message
Dim oConf As Object 'CDO.Configuration
Dim oFlds As Variant
Const Schema = "http://schemas.microsoft.com/cdo/configuration/"


Set oMsg = CreateObject("CDO.Message")
Set oConf = CreateObject("CDO.Configuration")
Set oFlds = oConf.Fields

'send one copy with Google SMTP server (with autentication)
With oFlds
    .Item(Schema & "sendusing") = 2
    .Item(Schema & "smtpserver") = sServer
    .Item(Schema & "smtpserverport") = 25
    .Item(Schema & "smtpauthenticate") = 1
    .Item(Schema & "sendusername") = sEmailFrom
    .Item(Schema & "sendpassword") = sPass
    .Item(Schema & "smtpusessl") = 1
    .Update
End With

With oMsg
    .To = sEmailTo
    .From = sEmailFrom
    .Subject = sSubject
    .HTMLBody = sBody
    
    .Sender = frmMain.LastName
    
    '.ReplyTo = "myemail@mydomain.com"
    
    Set .Configuration = oConf
    
    Sleep 100
    
    SendMail = .Send
End With


SendMail = True
oCleanUp:
Set oFlds = Nothing
Set oConf = Nothing
Set oMsg = Nothing

Exit Function
EH:
lblInfo.Caption = "Error Sending - " & Err.Description

SendMail = False

Resume oCleanUp
End Function
'### Not to be changed on pain of death ###

Private Sub cmdSend_Click()
Dim sBod As String
Dim sSub As String
Const Sep = "       | END OF MESSAGE. Stats: "

cmdSend.Enabled = False
txtBody.Enabled = False

Me.Refresh

sBod = txtBody.Text
If LenB(sBod) Then
    
    lblInfo.Caption = "Give me some time..."
    
    sSub = "Bug Report from " & frmMain.LastName
    sBod = sBod & Sep & GetInfoForMail()
    
    
    If SendMail(sBod, sSub) Then
        lblInfo.Caption = "Bug Reported"
    'Else
        'error added
    End If
    
Else
    lblInfo.Caption = "Enter something, ya crazy man"
    modDisplay.ShowBalloonTip txtBody, "Anyone there?", "Enter some text/a bug to report, then...", TTI_WARNING
    
    txtBody.Enabled = True
End If

End Sub

Private Function GetInfoForMail() As String
Const Sep = "      |      "

GetInfoForMail = "Name: " & frmMain.LastName & Sep & _
        "Local IP: " & modWinsock.LocalIP & Sep & _
        "Remote IP: " & IIf(LenB(modWinsock.RemoteIP), modWinsock.RemoteIP, "?") & Sep & _
        "Version: " & GetVersion()

End Function

Private Sub Form_Load()
lblInfo.Caption = "Loaded Window"
txtBody.Enabled = True
FormLoad Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FormLoad Me, True
End Sub

Private Sub txtBody_Change()
cmdSend.Enabled = CBool(LenB(txtBody.Text))
End Sub
