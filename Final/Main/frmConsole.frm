VERSION 5.00
Begin VB.Form frmConsole 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Console Commands"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5025
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   4718.735
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblConsole 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lblConsole.Caption = "CommandLine-Only Commands:" & vbNewLine _
    & "/DevForm - Opens the Developer Form" & vbNewLine _
    & "/Startup - Acts as if running at startup" & vbNewLine _
    & "/Host (0 or 1) - Host Mode" & vbNewLine _
    & "/Reset - Resets all settings" & vbNewLine _
    & "/InstancePrompt - Prompt if Previous Instance" & vbNewLine _
    & "/ForceOpen - Opens even if Previous Instance" & vbNewLine _
    & "/Console - Open the console while loading" & vbNewLine _
 _
    & vbNewLine & "Console/Command Line Commands:" _
    & "/Cls - Clears the screen" & vbNewLine _
    & "/SubClass (0 or 1) - (Advanced) SubClasses" & vbNewLine _
    & "/Log (0 or 1) - Enables/Disables Logging" & vbNewLine _
    & "/Dev <Password> - Developer Mode" & vbNewLine _
    & "/GameMode - Game Mode" & vbNewLine _
 _
    & vbNewLine & "For Console Commands, type Help at the Console" ' _
    & "Console-Only Commands:" & vbNewLine _
    & "/Listen or Close or Connect ""IP"" - Listen/Close Connection/Connect" '& vbNewLine _
    & "/Send * - Send * to Clients"  & vbNewLine _
    & "/Status ""ComputerName"" - Get Status" & vbNewLine _
    & "/Disconnect ""ComputerName"" - Disconnect Computer"

Me.Top = frmMain.Top + frmMain.width / 2 - Me.width / 2
Me.Left = frmMain.Left - Me.width
If Me.Left < 10 Then Me.Left = 10

Call FormLoad(Me, , , False)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub
