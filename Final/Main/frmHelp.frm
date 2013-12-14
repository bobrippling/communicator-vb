VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   6885
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11385
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdCmdline 
      Caption         =   "Command line"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "ChangeLog"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "File/Ports"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame fra 
      Height          =   4095
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtCmdline 
         BackColor       =   &H8000000F&
         Height          =   915
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmHelp.frx":0000
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtCL 
         BackColor       =   &H8000000F&
         Height          =   1725
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Text            =   "frmHelp.frx":000A
         Top             =   240
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label lblOptionsMain 
         AutoSize        =   -1  'True
         Caption         =   "Optn"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblFileMain 
         AutoSize        =   -1  'True
         Caption         =   "File"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pHeight As Integer = 8000
Private Const pWidth As Integer = 12240 '8000

Private Sub Form_Load()
Const Indent As Integer = 5

Me.WindowState = vbNormal
Me.height = pHeight
Me.width = pWidth

lblFileMain.Caption = "Save/Load Settings - Does what it says on the label" & vbNewLine & _
                        "(These Settings will be automatically loaded)" & vbNewLine & _
                        "Manual Connect - Connect to Specified IP" & vbNewLine & _
                        "Client Window - View Connected Clients" & vbNewLine & _
                        "Game Mode - Don't show when connected, and no balloon tips or shaking" & vbNewLine & _
                        vbNewLine & _
                        "Ports to Forward - " & modPorts.MainPort & " TCP (Main Connection)" & vbNewLine & _
                        Space$(27) & " - " & modPorts.StickPort & " UDP (Stick Game)" & vbNewLine & _
                        Space$(27) & " - " & modPorts.SpacePort & " UDP (Space Game)" & vbNewLine & _
                        Space$(27) & " - " & modPorts.FTPort & " TCP (File Transfer)" & vbNewLine & _
                        Space$(27) & " - " & modPorts.DPPort & " TCP (Display Pictures)"


lblOptionsMain.Caption = "Window" & vbNewLine & _
                        Space$(Indent) & "> Flash - Flash the Window" & vbNewLine & _
                        "Messaging" & vbNewLine & _
                        Space$(Indent) & "> Balloon - Show 'Balloon Tips'" & vbNewLine & _
                        Space$(Indent) & "> Different Colours - Display each person's chat colour" & vbNewLine & _
                        Space$(Indent) & "> Matrix - Type into the textbox directly" & vbNewLine & _
                        "Advanced" & vbNewLine & _
                        Space$(Indent) & "> Inactivity Timer - Hide the Window and Listen (after one minute)" & vbNewLine & _
                        Space$(Indent) & "> Host Mode - Listen when disconnected" & vbNewLine & _
                        Space$(Indent) & "> Minimize - Hide after disconnection"


'txtSpace.Text = InfoStart & "Fighter" & InfoEnd & vbNewLine & _
    "Average Speed" & vbNewLine & _
    "Average Acceleration" & vbNewLine & _
    "Average Shields" & vbNewLine & _
    "Average Fire Rate" & vbNewLine & vbNewLine & _
    InfoStart & "Behemoth" & InfoEnd & vbNewLine & _
    "Low Speed" & vbNewLine & _
    "Low Acceleration" & vbNewLine & _
    "High Shields" & vbNewLine & _
    "High Fire Rate" & vbNewLine & vbNewLine & _
    InfoStart & "Hornet" & InfoEnd & vbNewLine & _
    "High Speed" & vbNewLine & _
    "High Acceleration" & vbNewLine & _
    "Low Shields" & vbNewLine & _
    "Low Fire Rate" & vbNewLine & vbNewLine & InfoStart & "Mothership" & InfoEnd & vbNewLine & _
    "Very Low Speed" & vbNewLine & "Very Low Acceleration" & vbNewLine & _
    "Very High Shields" & vbNewLine & "5 Second Recharge Time (2 Second Fast Burst Fire)" & vbNewLine & _
    "Bullet Deflection" & vbNewLine & _
    "High Bullet Speed" & vbNewLine & vbNewLine & _
    "To obtain a mothership, you must have two kills with each other ship type." & vbNewLine & _
    "To obtain a wraith, you must have seven kills, five kills for an infiltrator." & vbNewLine & _
    "Middle click/Press Tab to open game settings." & vbNewLine & "Press F1 to view your score for each ship."

txtCL.Text = InfoStart & "ChangeLog" & InfoEnd & vbNewLine & vbNewLine & LoadResText(102) ', vbProperCase)
txtCmdline.Text = LoadResText(104)
txtCmdline.Font.Name = "Courier New"

cmdChange_Click

bModalFormShown = True

Call FormLoad(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bModalFormShown = False
Call FormLoad(Me, True)
End Sub

Private Sub Form_Resize()

With fra
    .width = ScaleWidth - .Left - 100
    .height = ScaleHeight - .Top - 100
    txtCL.width = .width - txtCL.Left - 200
    txtCL.height = .height - txtCL.Top - 200
    'txtSpace.width = txtCL.width
    'txtSpace.height = txtCL.height
    txtCmdline.width = txtCL.width
    txtCmdline.height = txtCL.height
End With

End Sub

Public Sub cmdChange_Click()
Call HideAll
txtCL.Visible = True

cmdChange.Default = True
End Sub

Private Sub cmdFile_Click()
Call HideAll
lblFileMain.Visible = True
End Sub

Private Sub cmdOptions_Click()
Call HideAll
lblOptionsMain.Visible = True
End Sub

'Private Sub cmdSpace_Click()
'Call HideAll
'txtSpace.Visible = True
'End Sub

Private Sub cmdCmdline_Click()
Call HideAll
txtCmdline.Visible = True
End Sub

Private Sub HideAll()
lblFileMain.Visible = False
lblOptionsMain.Visible = False
'txtSpace.Visible = False
txtCL.Visible = False
txtCmdline.Visible = False
End Sub
