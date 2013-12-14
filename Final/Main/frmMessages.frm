VERSION 5.00
Begin VB.Form frmMessages 
   Caption         =   "Messages"
   ClientHeight    =   2850
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5955
   Begin VB.CommandButton cmdDate 
      Caption         =   "Insert Date"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "Insert Time"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox txtMessage 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmMessages.frx":0000
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status: Idle"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   3555
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSelect 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const minX = 453, minY = 244

'pasting etc
Private Enum eClipTypes
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
End Enum

'Private Const WM_CUT = &H300
'Private Const WM_COPY = &H301
'Private Const WM_PASTE = &H302
'Private Const WM_CLEAR = &H303 'aka del
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Any) As Long
'end

Private Const FName As String = "Messages." & FileExt
Private RFile As String
Private LFile As String '= rootdrive & fname

Private Sub DoClipboard(ByVal lType As eClipTypes)

SendMessageByLong Me.txtMessage.hWnd, lType, 0&, 0&

End Sub

Private Sub cmdDevEnable_Click()
txtMessage.Enabled = True
cmdUpload.Enabled = True
cmdDownload.Enabled = True
End Sub

Private Sub cmdDate_Click()
If txtMessage.Enabled Then
    txtMessage.SelText = Date$
End If
End Sub

Private Sub cmdDownload_Click()
Dim f As Integer
Dim Total As String, sError As String
Dim ErrorOccured As Boolean

SetStatus "Downloading Messages"

cmdDownload.Enabled = False

Me.Refresh

If modFTP.DownloadFTPFile(LFile, RFile, cmdDownload, sError, True) Then
    If FileExists(LFile) Then
        
        f = FreeFile()
        Open LFile For Binary Access Read As #f
            Total = Space$(LOF(1))
            Get #f, , Total
        Close #f
        
        'Total = Mid$(Total, 3) 'get rid of beginning newline
        
        txtMessage.Text = Trim$(Total) 'get rid of trailing newline
        txtMessage.Enabled = True
        cmdTime.Enabled = True
        cmdDate.Enabled = True
        cmdDownload.Enabled = False
        
        On Error Resume Next
        Kill LFile
        
        SetStatus "Downloaded Messages Successfully"
        
    Else
        ErrorOccured = True
    End If
Else
    ErrorOccured = True
End If

If ErrorOccured Then
    SetStatus "Error Downloading File" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
    cmdDownload.Enabled = True
    cmdUpload.Enabled = False
End If

End Sub

Private Sub cmdTime_Click()
If txtMessage.Enabled Then
    txtMessage.SelText = Time$
End If
End Sub

Private Sub cmdUpload_Click()
Dim f As Integer
Dim Str As String, sError As String

Str = txtMessage.Text
f = FreeFile()

If FileExists(LFile) Then
    On Error Resume Next
    Kill LFile
End If

If LenB(Str) = 0 Then Str = vbSpace

Open LFile For Output As #f
    Print #f, Str;
Close #f

cmdUpload.Enabled = False
txtMessage.Enabled = False
cmdDownload.Enabled = False
cmdTime.Enabled = False
cmdDate.Enabled = False

SetStatus "Uploading..."
Me.Refresh

If modFTP.UploadFTPFile(LFile, RFile, cmdUpload, sError, True) Then
    SetStatus "Uploaded Successfully"
    cmdDownload.Enabled = True
    txtMessage.Text = "Messages entered here."
Else
    SetStatus "Error Uploading" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
    cmdUpload.Enabled = True
    txtMessage.Enabled = True
End If

On Error Resume Next
Kill LFile
End Sub

Private Sub Form_Load()
LFile = modSettings.GetTmpFilePath() & "Messages." & modVars.FileExt 'Left$(App.Path, 3) & FName

'If bDevMode Then
'    lblStatus.Top = 1080
'    cmdDevEnable.Visible = True
'End If

RFile = modFTP.FTP_Root_Location & "/Messages/" & FName

Me.width = ScaleX(minX, vbPixels, vbTwips)
Me.height = ScaleY(minY, vbPixels, vbTwips)

If modLoadProgram.bIsIDE = False Then
    modSubClass.SubclassAuto Me
End If

Call FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If modLoadProgram.bIsIDE = False Then
    modSubClass.SubclassAuto Me, False
End If

Call FormLoad(Me, True)
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Dim MMI As MINMAXINFO


If uMsg = WM_GETMINMAXINFO Then
    
    CopyMemory MMI, ByVal lParam, LenB(MMI)
    
    'set the MINMAXINFO data to the
    'minimum and maximum values set
    'by the option choice
    
    With MMI
        .ptMinTrackSize.X = minX
        .ptMinTrackSize.Y = minY
        '.ptMaxTrackSize.X = maxX
        '.ptMaxTrackSize.Y = maxY
    End With
    
    CopyMemory ByVal lParam, MMI, LenB(MMI)
    WindowProc = 0
Else
    WindowProc = modSubClass.CallWindowProc(GetProp(hWnd, WndProcStr), hWnd, uMsg, wParam, lParam)
End If

End Function

Private Sub Form_Resize()
On Error Resume Next
txtMessage.width = ScaleWidth
txtMessage.height = ScaleHeight - txtMessage.Top
'fra.Left = Me.ScaleWidth / 2 - fra.width / 2

lblStatus.width = Me.ScaleWidth

cmdDownload.Left = Me.ScaleWidth / 5 - cmdDownload.width
cmdUpload.Left = Me.ScaleWidth * 2 / 5 - cmdUpload.width
cmdTime.Left = Me.ScaleWidth * 3 / 5 - cmdTime.width
cmdDate.Left = Me.ScaleWidth * 4 / 5 - cmdDate.width - 120
End Sub

Private Sub SetStatus(ByVal T As String)
lblStatus.Caption = "Status: " & T
lblStatus.Refresh
End Sub

Private Sub mnuEditCopy_Click()
DoClipboard WM_COPY
End Sub

Private Sub mnuEditCut_Click()
DoClipboard WM_CUT
End Sub

Private Sub mnuEditPaste_Click()
DoClipboard WM_PASTE
End Sub

Private Sub mnuEditSelect_Click()
With txtMessage
    .Selstart = 0
    .Sellength = Len(.Text)
End With
End Sub

Private Sub txtMessage_Change()
If txtMessage.Enabled Then cmdUpload.Enabled = True
End Sub
