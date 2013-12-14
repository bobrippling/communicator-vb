VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSpeech 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voice Options"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkVoice 
      Caption         =   "Use Voice (Overrides All Options)  (Not Saved)"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Frame fraScroll 
      Height          =   2655
      Left            =   4800
      TabIndex        =   10
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox picScrolly 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   120
         ScaleHeight     =   2295
         ScaleWidth      =   4095
         TabIndex        =   11
         Top             =   240
         Width           =   4095
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Select Voice"
            Enabled         =   0   'False
            Height          =   735
            Left            =   2640
            TabIndex        =   19
            Top             =   1560
            Width           =   1455
         End
         Begin MSComctlLib.Slider sldrVol 
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            SmallChange     =   5
            Min             =   1
            Max             =   100
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sldrSpeed 
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   4
            SmallChange     =   2
            Min             =   -10
            TickFrequency   =   5
         End
         Begin projMulti.ScrollListBox lstVoices 
            Height          =   735
            Left            =   0
            TabIndex        =   18
            Top             =   1560
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   1296
         End
         Begin MSComctlLib.Slider sldrPitch 
            Height          =   255
            Left            =   2040
            TabIndex        =   17
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   4
            SmallChange     =   2
            Min             =   -10
            TickFrequency   =   5
         End
         Begin VB.Label lblPitch 
            Caption         =   "Pitch: WWW"
            Height          =   255
            Left            =   2040
            TabIndex        =   15
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblVol 
            Caption         =   "Volume: WWW"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   1335
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Speed: WWW"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   840
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   320
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Timer tmrStatus 
      Interval        =   500
      Left            =   0
      Top             =   480
   End
   Begin VB.Frame fraSettings 
      Caption         =   "Settings"
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   2880
      Width           =   9015
      Begin VB.PictureBox picSettings 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   8850
         TabIndex        =   21
         Top             =   240
         Width           =   8850
         Begin VB.CheckBox chkSent 
            Caption         =   "Speak Sent Messages"
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   1200
            Width           =   3735
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "Speak Received Information Messages"
            Height          =   255
            Left            =   4680
            TabIndex        =   31
            Top             =   960
            Width           =   3735
         End
         Begin VB.CheckBox chkQuestions 
            Caption         =   "Speak Questions that Communicator Asks"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   240
            Width           =   3375
         End
         Begin VB.CheckBox chkBalloon 
            Caption         =   "Speak the Contents of Balloon Tips"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   2895
         End
         Begin VB.CheckBox chkReceived 
            Caption         =   "Speak received messages"
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   480
            Width           =   4095
         End
         Begin VB.CheckBox chkHiBye 
            Caption         =   "Say Hello and Goodbye (Program Startup and Close)"
            Height          =   255
            Left            =   4680
            TabIndex        =   23
            Top             =   0
            Width           =   4095
         End
         Begin VB.CheckBox chkSayName 
            Caption         =   "Say the sender's name before their message"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   3855
         End
         Begin VB.CheckBox chkHi 
            Caption         =   "Say Hello"
            Height          =   255
            Left            =   4920
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkBye 
            Caption         =   "Say Goodbye"
            Height          =   255
            Left            =   4920
            TabIndex        =   27
            Top             =   480
            Width           =   1815
         End
         Begin VB.CheckBox chkGameSpeak 
            Caption         =   "Speak Messages, etc, in Game Mode"
            Height          =   255
            Left            =   4680
            TabIndex        =   29
            Top             =   720
            Width           =   3015
         End
         Begin VB.CheckBox chkForegroundOnly 
            Caption         =   "Only speak if Communicator is the foreground window"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   960
            Width           =   4095
         End
      End
   End
   Begin VB.TextBox txtTest 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Speak"
      Enabled         =   0   'False
      Height          =   320
      Left            =   3000
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Details"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Status"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblDesc 
         Alignment       =   2  'Center
         Caption         =   "Desc"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblGender 
         Alignment       =   2  'Center
         Caption         =   "Gender"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label lblAge 
         Alignment       =   2  'Center
         Caption         =   "Age"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkBalloon_Click()
modSpeech.sBalloon = CBool(chkBalloon.Value)
End Sub

Private Sub chkForegroundOnly_Click()
modSpeech.sOnlyForeground = CBool(chkForegroundOnly.Value)
End Sub

Private Sub chkGameSpeak_Click()
modSpeech.sGameSpeak = CBool(chkGameSpeak.Value)
End Sub

Private Sub chkHi_Click()

modSpeech.sHi = CBool(chkHi.Value)

Call HiByeCheck

End Sub

Private Sub chkBye_Click()

modSpeech.sBye = CBool(chkBye.Value)

Call HiByeCheck

End Sub

Private Sub HiByeCheck()

If modSpeech.sHi = False And modSpeech.sBye = False Then
    chkHiBye.Value = 0
End If

End Sub

Private Sub chkHiBye_Click()
modSpeech.sHiBye = CBool(chkHiBye.Value)

If Me.Visible Then
    If modSpeech.sHiBye Then
        chkHi.Value = 1
        chkBye.Value = 1
    Else
        chkHi.Value = 0
        chkBye.Value = 0
    End If
End If

chkHi.Enabled = modSpeech.sHiBye
chkBye.Enabled = modSpeech.sHiBye

End Sub

Private Sub chkInfo_Click()
modSpeech.sSayInfo = CBool(chkInfo.Value)
End Sub

Private Sub chkQuestions_Click()
modSpeech.sQuestions = CBool(chkQuestions.Value)
End Sub

Private Sub chkReceived_Click()
modSpeech.sReceived = CBool(chkReceived.Value)

If Me.Visible Then chkSayName.Value = IIf(modSpeech.sReceived, 1, 0)
chkSayName.Enabled = modSpeech.sReceived

End Sub

Private Sub chkSayName_Click()

modSpeech.sSayName = CBool(chkSayName.Value)

End Sub

Private Sub cmdOverride_Click(Index As Integer)
Dim Ctrl As Control

On Error Resume Next
For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is CheckBox Then
        Ctrl.Value = Index
    End If
Next Ctrl

End Sub

Private Sub chkSent_Click()
modSpeech.sSent = CBool(chkSent.Value)
End Sub

Private Sub chkVoice_Click()
modSpeech.bVoice = CBool(chkVoice.Value)
End Sub

Private Sub cmdSelect_Click()
Dim sVoice As String
Dim i As Integer

cmdSelect.Enabled = False

sVoice = lstVoices.Text

If LenB(sVoice) Then
    
    For i = 0 To UBound(modSpeech.VoiceAr)
        If modSpeech.VoiceAr(i).sDesc = sVoice Then
            modSpeech.SetVoice i
            Exit For
        End If
    Next i
    
    RefreshInfo
End If

End Sub

Private Sub cmdStop_Click()
cmdStop.Enabled = False
modSpeech.StopSpeech
tmrStatus_Timer

modDisplay.ShowBalloonTip txtTest, "Stopped", "How dare you interrupt me?!"
End Sub

Private Sub cmdTest_Click()
cmdTest.Enabled = False
cmdTest.Default = False
modSpeech.Say txtTest.Text, sldrVol.Value, sldrSpeed.Value, True
tmrStatus_Timer

modDisplay.ShowBalloonTip txtTest, "Speaking", "I am speaking, silence please"
End Sub

Private Sub lstVoices_Click()
Dim sTxt As String

sTxt = lstVoices.Text

cmdSelect.Enabled = CBool(LenB(sTxt)) And (sTxt <> modSpeech.Description)

End Sub

Private Sub tmrStatus_Timer()
Dim nStatus As eSpeechStatus

lblStatus.Caption = "Status: " & modSpeech.SpeechStatus

nStatus = modSpeech.nSpeechStatus

If (Len(txtTest.Text) > 0) Then
    cmdTest.Enabled = (nStatus = sIdle)
    cmdTest.Default = cmdTest.Enabled
Else
    cmdTest.Enabled = False
End If

cmdStop.Enabled = (nStatus = sSpeaking)

End Sub

Private Sub Form_Load()
Dim i As Integer

Call RefreshInfo

'list voices
For i = 0 To UBound(modSpeech.VoiceAr)
    lstVoices.AddItem modSpeech.VoiceAr(i).sDesc
    'lstVoices.LetItemData i, lstVoices.ListCount - 1
Next i

sldrVol.Value = modSpeech.Vol
sldrSpeed.Value = modSpeech.Speed
sldrPitch.Value = modSpeech.pitch

sldrVol_Change 'in case the value is already what has been loaded
sldrSpeed_Change
sldrPitch_Change

chkBalloon.Value = IIf(modSpeech.sBalloon, 1, 0)
chkReceived.Value = IIf(modSpeech.sReceived, 1, 0)
chkQuestions.Value = IIf(modSpeech.sQuestions, 1, 0)
chkHiBye.Value = IIf(modSpeech.sHiBye, 1, 0)
chkHi.Value = Abs(modSpeech.sHi And modSpeech.sHiBye)
chkBye.Value = Abs(modSpeech.sBye And modSpeech.sHiBye)
chkSayName.Value = IIf(modSpeech.sSayName, 1, 0)
chkGameSpeak.Value = IIf(modSpeech.sGameSpeak, 1, 0)
chkForegroundOnly.Value = IIf(modSpeech.sOnlyForeground, 1, 0)
chkInfo.Value = Abs(modSpeech.sSayInfo)
chkVoice.Value = Abs(modSpeech.bVoice)
chkSent.Value = Abs(modSpeech.sSent)

'enable/disable
chkHiBye_Click
chkReceived_Click

TurnOffToolTip sldrVol.hWnd
TurnOffToolTip sldrSpeed.hWnd
TurnOffToolTip sldrPitch.hWnd

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdSelect.hWnd, frmMain.GetCommandIconHandle()
End If

Me.Left = frmMain.Left - Me.width
Me.Top = frmMain.Top + frmMain.height / 2 - Me.height / 2
If Me.Left < 10 Then Me.Left = 10
Call FormLoad(Me, , , False)

Show

tmrStatus_Timer

End Sub

Private Sub RefreshInfo()
lblDesc.Caption = "Description: " & modSpeech.Description

lblName.Caption = "Name: " & modSpeech.Attr("Name")
lblGender.Caption = "Gender: " & modSpeech.Attr("Gender")
lblAge.Caption = "Age: " & modSpeech.Attr("Age")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub sldrSpeed_Change()
lblSpeed.Caption = "Speed: " & sldrSpeed.Value + 11
modSpeech.Speed = sldrSpeed.Value
End Sub

Private Sub sldrSpeed_Click()
sldrSpeed_Change
End Sub

Private Sub sldrSpeed_Scroll()
sldrSpeed_Change
End Sub

Private Sub sldrVol_Change()
lblVol.Caption = "Volume: " & sldrVol.Value
modSpeech.Vol = sldrVol.Value
End Sub

Private Sub sldrVol_Click()
sldrVol_Change
End Sub

Private Sub sldrVol_Scroll()
sldrVol_Change
End Sub

Private Sub sldrPitch_Change()
lblPitch.Caption = "Pitch: " & sldrPitch.Value
modSpeech.pitch = sldrPitch.Value
End Sub

Private Sub sldrPitch_Click()
sldrPitch_Change
End Sub

Private Sub sldrPitch_Scroll()
sldrPitch_Change
End Sub

Private Sub txtTest_Change()
tmrStatus_Timer
End Sub

Private Sub txtTest_GotFocus()
txtTest.Selstart = 0
txtTest.Sellength = Len(txtTest.Text)
End Sub
