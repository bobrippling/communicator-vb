VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3060
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   Begin projMulti.VistaProg progLoad 
      Height          =   225
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   397
   End
   Begin VB.Label lblLoading2 
      Alignment       =   2  'Center
      Caption         =   "Hello there"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1620
      Width           =   3735
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   120
      Picture         =   "frmSplash.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "AppTitle"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version: x.y.z"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      Caption         =   "Loading, Please Wait..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label lblDesc 
      Alignment       =   2  'Center
      Caption         =   "File Desc -by Company Name"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   3735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadLoadStatement()
'http://stackoverflow.com/questions/182112/funny-loading-statements-to-keep-users-amused
Dim Statements(0 To 5) As String
Dim iRnd As Integer

Statements(0) = "Pay no attention to the man behind the curtain"
Statements(1) = "Spinning up the hamster..."
Statements(2) = "What do you call a sleeping bull? A bulldozer"
Statements(3) = "Why are pirates called pirates? Cos they arrrr"
Statements(4) = "What is Jaws' Grandfather called? Gums"
Statements(5) = "What's brown and sounds like a bell? Duuuung"


Do
    For iRnd = 0 To UBound(Statements)
        If Rnd() > 0.9 Then
            lblLoading2.Caption = Statements(iRnd)
            Exit Do
        End If
    Next iRnd
Loop


Erase Statements

End Sub

Private Sub Form_Load()
Dim lhWnd As Long, btTrans As Byte
Dim lTick As Long, lGTC As Long

LoadLoadStatement

lblVersion.Caption = "Version: " & modVars.GetVersion() ' App.Major & "." & App.Minor & "." & App.Revision
lblDesc.Caption = App.FileDescription & " - by " & App.CompanyName
Me.Caption = App.Title
lblTitle.Caption = App.Title
lblStatus.Caption = "Loading " & App.Title & "..."

imgLogo.Picture = Me.Icon
imgLogo.ZOrder vbBringToFront

If modVars.bStartup Then
    frmSplash.progLoad.Value = frmSplash.progLoad.Max
End If

modLoadProgram.frmSplash_Loaded = True

DrawBorder Me

If modVars.bStealth = False Then
    Me.Move Screen.width / 2 - Me.width / 2, Screen.height / 2 - Me.height / 2
    Show '- done by setontop()
    Me.MousePointer = MousePointerConstants.vbHourglass
    lhWnd = Me.hWnd
    
    
    'If modVars.bStartup Then
        btTrans = 50
        modDisplay.SetTransparentStyle lhWnd
        
        lTick = GetTickCount()
        
        Do
            modDisplay.SetTransparency lhWnd, btTrans
            
            lGTC = GetTickCount()
            On Error Resume Next
            btTrans = btTrans + 5 * (lGTC - lTick) / 10
            lTick = lGTC
            
            If Err.Number Then Exit Do
            
            Pause 5
        Loop While btTrans <= 250
        
        modDisplay.SetTransparentStyle lhWnd, False
    'End If
Else
    Unload Me
End If

End Sub

Public Sub pSetInfo(ByVal Inf As String)
Const K As String = "Status: "

lblStatus.Caption = K & Inf

Me.Refresh
lblStatus.Refresh
progLoad.Refresh

End Sub

Private Sub Form_Paint()
progLoad.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lhWnd As Long, btTrans As Byte
Dim lTick As Long, lGTC As Long


If modVars.bStartup Or modLoadProgram.frmMain_Loaded Then
    
    lhWnd = Me.hWnd
    btTrans = 255
    
    If Me.Visible Then
        modDisplay.SetTransparentStyle lhWnd
        
        lTick = GetTickCount()
        
        Do
            modDisplay.SetTransparency lhWnd, btTrans
            
            lGTC = GetTickCount()
            On Error Resume Next
            btTrans = btTrans - 5 * (lGTC - lTick) / 10
            lTick = lGTC
            
            If Err.Number Then Exit Do
            
            Me.Refresh
            If modLoadProgram.frmMain_Loaded Then frmMain.Refresh
            Sleep 5
        Loop While btTrans > 5
        
        Me.Hide
    End If
End If


modDisplay.SetTransparentStyle lhWnd, False
modLoadProgram.frmSplash_Loaded = False
End Sub

