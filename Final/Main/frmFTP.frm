VERSION 5.00
Begin VB.Form frmFTP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "FTP Download/Upload"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin projMulti.ucFloodProgBar FloodBar 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
   End
   Begin projMulti.VistaProg progFTP 
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   397
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      Caption         =   "Speed: 0 Bytes/s"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const bDebug_Manual_FTP = False

Private WithEvents cFTP As clsFTP
Attribute cFTP.VB_VarHelpID = -1

Private sInfo As String
Private pbCancel As Boolean, pbDeferToUpdateForm As Boolean

'############
'Speed stuff
Private lLastPacket As Long, lLastCurrentBytes As Long



'############

'#Const bShowModal=0
'#If bShowModal Then
    'Private Const Show_Style = vbModal
'#Else
    'Private Const Show_Style = vbModeless
'#End If


Public Property Let bIsDownload(bV As Boolean)

sInfo = IIf(bV, "Download", "Upload") & " Progress: "

End Property

Public Sub cFTP_FileTransferProgress(ByVal lCurrentBytes As Long, ByVal lTotalBytes As Long, _
    ByRef bCancel As Boolean)

Dim Percent As Single
Dim sPercent As String, sSpeed As String ', sETA As String
Dim lGTC As Long

If modFTP.FTP_StealthMode Then Exit Sub


Percent = 100 * lCurrentBytes / lTotalBytes

'##############################################################################
's = d/t
lGTC = GetTickCount()
sSpeed = CStr((lCurrentBytes - lLastCurrentBytes) / (lGTC - lLastPacket + 1)) '+1 prevents /0
''t=d/s
'sETA = CStr((lTotalBytes - lCurrentBytes) / (CLng(sSpeed) + 1))
lLastPacket = lGTC
lLastCurrentBytes = lCurrentBytes
'##############################################################################

If pbDeferToUpdateForm Then
    frmUpdate.Set_Progress Percent, bCancel
Else
    sPercent = FormatNumber$(Percent, 2, vbTrue, vbFalse, vbFalse) & "%"
    
    lblSpeed.Caption = "Speed: " & FormatNumber$(sSpeed, 2, vbTrue, vbFalse, vbFalse) & " Bytes/s" '& _
                       Space$(5) & "ETA: " & FormatNumber$(sETA, vbTrue, vbFalse, vbFalse) & " seconds"
    
    'If modFTP.bUsingManualMethod Then
        lblInfo.Caption = sInfo & sPercent
        progFTP.Value = Percent
        bCancel = pbCancel
    'Else
        FloodBar.Flood_Show_Percentage Percent, "File Transfer, " & sPercent
    'End If
    
    Me.Refresh
    lblInfo.Refresh
    
    DoEvents
End If

End Sub

Public Sub SetLabelInfo(sMsg As String)

If pbDeferToUpdateForm Then
    frmUpdate.lblState.Caption = sMsg
Else
    lblInfo.Caption = sMsg
End If

End Sub

Private Sub cmdCancel_Click()
If modFTP.bUsingManualMethod Then
    pbCancel = True
Else
    modFTP.bCancelFTP = True
End If

cmdCancel.Enabled = False
Me.Refresh

End Sub

'########################################################################################

Private Sub Form_Load()

lLastPacket = GetTickCount()
lblSpeed.Caption = "If Communicator freezes, wait for it, be patient"

'##########################################
'must be done here, so cFTP is set, in case defered to update form
If modFTP.bUsingManualMethod Then
    Set cFTP = New clsFTP
    
    cFTP.SetMode (Not frmMain.mnuOnlineFTPPassive.Checked)
    
    'cmdCancel.Enabled = True
    pbCancel = False
    
    FloodBar.Visible = False
    'progFTP.Visible = True
Else
    'cmdCancel.Enabled = False
    
    'FloodBar.Top = progFTP.Top
    FloodBar.Visible = True
    'progFTP.Visible = False
    
    FloodBar.Flood_Show_Message "Connecting to Server..."
End If
'##########################################


If modFTP.FTP_DeferToUpdateForm Then
    Me.Visible = False
    pbDeferToUpdateForm = True
Else
    
    pbDeferToUpdateForm = False
    
    'SetOnTop Me.hWnd '+show
    FormLoad Me, , Not modFTP.FTP_StealthMode
    lblInfo.Caption = "Connecting to Server..."
    
    
    If modFTP.FTP_StealthMode = False Then
        Me.Show vbModeless, frmMain
        DrawBorder Me
        If modFTP.bUsingManualMethod Then Pause 100
    End If
End If


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If modFTP.bUsingManualMethod Then
    If modFTP.FTP_StealthMode = False Then
        lblInfo.Caption = "Closing Connection to Server..."
        Me.Refresh
        lblInfo.Refresh
        
        cmdCancel_Click
        
        Pause 150
    End If
    
    'cFTP.CloseConnection
    'done below
    Set cFTP = Nothing
End If

End Sub

'########################################################################################

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    If Me.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessageByLong Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End If

End Sub

Private Sub lblInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub progFTP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub

'########################################################################################

Public Function FTP_Transfer(ByVal ServerName As String, ByVal UName As String, ByVal Pass As String, _
    ByVal RemoteFile As String, ByVal LocalFile As String, ByVal bGet As Boolean, _
    ByRef sError As String) As Boolean

Dim bSuccess As Boolean, bPrintOut As Boolean
Dim f As Integer
Dim sPath As String

bSuccess = False

With cFTP
    If .OpenConnection(ServerName, UName, Pass) Then
        
        If bGet Then
            
            bSuccess = .FTPDownloadFile(LocalFile, RemoteFile)
            
        Else
            
            bSuccess = .FTPUploadFile(LocalFile, RemoteFile)
            
        End If
        
        If bSuccess Then
            sError = vbNullString
            bPrintOut = False
        Else
            sError = .SimpleLastErrorMessage
            
            If FileExists(LocalFile) Then
                On Error Resume Next
                Kill LocalFile
            End If
            progFTP.Value = 0
            
            bPrintOut = True
        End If
        
        .CloseConnection
        
        
        
    Else
        .CloseConnection
        
        sError = .SimpleLastErrorMessage
        bPrintOut = True
        
        lblInfo.Caption = sError
    End If
    
End With


#If bDebug_Manual_FTP Then
    If bPrintOut And modFTP.FTP_StealthMode = False Then
        f = FreeFile()
        
        sPath = frmMain.GetLogPath() & "FTP_Debug.txt"
        
        Open sPath For Append As #f
            Print #f, CStr(Time) & " FTP Error"
            Print #f, "Simple Message: " & cFTP.SimpleLastErrorMessage
            Print #f, "Advanced Message: " & cFTP.LastErrorMessage
            Print #f, vbNewLine;
        Close #f
        
        
        'Pause 500
    End If
#End If

FTP_Transfer = bSuccess

End Function
