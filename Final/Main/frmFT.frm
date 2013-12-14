VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Upload/Download"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvFiles 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh File List"
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload a File"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   8775
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private RemoteLocation As String
Private LocalLocation As String '"c:\"

Private Property Get RemoteLocation() As String
RemoteLocation = modFTP.FTP_Root_Location & "/Files"
'in case FTP settings are changed while this form is loaded
End Property

Private Function getSelectedFileName() As String
If Not (lvFiles.SelectedItem Is Nothing) Then
    getSelectedFileName = lvFiles.SelectedItem.Text
End If
End Function

Private Sub cmdDel_Click()
Dim FName As String
Dim Ans As VbMsgBoxResult
Dim ErrorOccured As Boolean
Dim sError As String

FName = getSelectedFileName()
If LenB(FName) Then
    
    
    cmdDel.Enabled = False
    cmdRefresh.Enabled = False
    cmdUpload.Enabled = False
    cmdDownload.Enabled = False
    
    Ans = MsgBoxEx("Delete '" & FName & "'" & vbNewLine & _
        "Are You Sure?", "Once this file's gone, it's not comming back...", _
        vbQuestion + vbYesNo, "Download", , , frmMain.Icon)
    
    If Ans = vbYes Then
        
        SetStatus "Deleting File..."
        
        If modFTP.DelFTPFile(RemoteLocation & "/" & FName, sError) Then
            cmdRefresh_Click
            SetStatus "File Deleted"
        Else
            ErrorOccured = True
            cmdRefresh_Click
        End If
    End If
Else
    SetStatus "Please Select a File to Delete"
End If

If ErrorOccured Then
    SetStatus "Error Deleting File - " & sError
End If

cmdRefresh.Enabled = True

End Sub

Private Sub cmdDownload_Click()
Dim FName As String, LFile As String, RFile As String, sError As String
Dim Ans As VbMsgBoxResult
Dim ErrorOccured As Boolean, CDlError As Boolean
Dim IDir As String, sFileExt As String

FName = getSelectedFileName()

cmdDel.Enabled = False
cmdRefresh.Enabled = False
cmdUpload.Enabled = False
cmdDownload.Enabled = False

If LenB(FName) Then
    'Ans = MsgBox("Download '" & FName & "'" & vbNewLine & _
        "Are You Sure?", vbQuestion + vbYesNo, "Download")
    
    'If Ans = vbYes Then
    
    LFile = LocalLocation & "\" & FName
    sFileExt = GetFileExtension(FName)
    
    IDir = AppPath() & "Received Files"
    If FileExists(IDir, vbDirectory) = False Then
        IDir = LocalLocation
    End If
    
    frmMain.CommonDPath LFile, CDlError, "Upload File", _
        "'." & sFileExt & "' File|*." & sFileExt & "|All Files (*.*)|*.*", _
        IDir
    
    
    If CDlError = False Then
        If LenB(LFile) Then
            
            RFile = RemoteLocation & "/" & FName
            
            SetStatus "Downloading File..."
            
            If modFTP.DownloadFTPFile(LFile, RFile, cmdDownload, sError, True) Then
                If FileExists(LFile) Then
                    
                    'Shell "explorer.exe " & Left$(LFile, _
                        InStrRev(LFile, "\", , vbTextCompare)), vbNormalFocus
                    
                    OpenFolder vbNormalFocus, , LFile
                    
                    cmdRefresh_Click
                    
                    SetStatus "Downloaded File Successfully"
                Else
                    cmdRefresh_Click
                    ErrorOccured = True
                End If
            Else
                ErrorOccured = True
            End If
        Else
            ErrorOccured = False
            SetStatus vbNullString
        End If
    End If
    'End If
Else
    SetStatus "Please Select a File"
End If

If ErrorOccured Then
    SetStatus "Error Downloading File" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
End If

cmdRefresh.Enabled = True

End Sub

Private Sub cmdRefresh_Click()
SetStatus "Refreshing List"

'Me.Refresh
Call RefreshList

End Sub

Private Sub cmdUpload_Click()
Dim LFile As String, RFile As String, FName As String, sError As String
Dim Ans As VbMsgBoxResult
Dim ErrorOccured As Boolean

cmdDel.Enabled = False
cmdRefresh.Enabled = False
cmdUpload.Enabled = False
cmdDownload.Enabled = False

frmMain.CommonDPath LFile, ErrorOccured, "Upload File", "All Files (*.*)|*.*", LocalLocation, True

If ErrorOccured = False Then
    If LenB(LFile) Then
        Ans = MsgBoxEx("Upload '" & LFile & "'" & vbNewLine & _
            "Are You Sure?", "This may take a while, depending on the file's size...", _
            vbQuestion + vbYesNo, "Download", , , frmMain.Icon)
        
        If Ans = vbYes Then
            
            FName = Right$(LFile, Len(LFile) - InStrRev(LFile, "\", , vbTextCompare))
            RFile = RemoteLocation & "/" & FName
            
            SetStatus "Uploading File..."
            
            If modFTP.UploadFTPFile(LFile, RFile, cmdUpload, sError, True) Then
                SetStatus "Uploaded File Successfully"
            Else
                ErrorOccured = True
            End If
        End If
    Else
        SetStatus "Please Select a File"
    End If
Else
    SetStatus vbNullString
    cmdUpload.Enabled = True
    Exit Sub
End If

cmdRefresh_Click

If ErrorOccured Then
    SetStatus "Error Uploading File" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
End If

End Sub

Private Sub Form_Load()

With lvFiles.ColumnHeaders
    .Add , , "Name"
    .Add , , "Size"
    .Add , , "Last Written"
End With
lvFiles.ColumnHeaders(1).width = lvFiles.width / 3 - ScaleX(2, vbPixels, vbTwips)
lvFiles.ColumnHeaders(2).width = lvFiles.ColumnHeaders(1).width
lvFiles.ColumnHeaders(3).width = lvFiles.ColumnHeaders(1).width



LocalLocation = Trim$(modPaths.SavedFilesPath)

If LenB(LocalLocation) = 0 Or FileExists(LocalLocation, vbDirectory) = False Then
    LocalLocation = modVars.RootDrive & "\"
End If

Call FormLoad(Me)
SetStatus "Idle"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub RefreshList()
Dim FileList() As ptFTPFile
Dim sError As String
Dim i As Integer


cmdDel.Enabled = False
cmdRefresh.Enabled = False
cmdUpload.Enabled = False
cmdDownload.Enabled = False

lvFiles.ListItems.Clear

Me.Refresh

If modFTP.ListDir(RemoteLocation, FileList, sError) Then
    
    If FileArrayDimensioned(FileList) = False Then
        SetStatus "Obtained File List for " & frmMain.GetServerName() & " (No Files Present)"
    Else
        For i = LBound(FileList) To UBound(FileList)
    '        j = InStr(1, Files(i), "|", vbTextCompare) - 1
    '        Name = Trim$(Left$(Files(i), j))
    '        Size = Trim$(Mid$(Files(i), j + 2, InStr(j + 3, Files(i), "|", vbTextCompare) - j - 2))
    '        Dt = Trim$(Right$(Files(i), Len(Files(i)) - InStrRev(Files(i), "|", , vbTextCompare)))
    '
    '        On Error Resume Next
    '        lSize = CLng(Left$(Size, InStrRev(Size, Space$(1), , vbTextCompare) - 1)) / 1024
    '
    '        Size = CStr(lSize) & " KB"
            
            With FileList(i)
                AddToList Trim$(.sName), .lFileSize, .dDateLastWritten
            End With
        Next i
        
        SetStatus "Obtained File List for " & frmMain.GetServerName()
        
        'cmdDownload.Enabled = True
        'cmdDel.Enabled = True
    End If
    cmdUpload.Enabled = True
    
Else
    SetStatus "Error - " & sError
End If

cmdRefresh.Enabled = True

End Sub

Private Sub SetStatus(ByVal S As String)
lblStatus.Caption = "Status: " & S
lblStatus.Refresh
End Sub
Private Sub AddToList(sName As String, lSize As Long, dDateLastWrite As Date)


lvFiles.ListItems.Add , , sName

With lvFiles.ListItems(lvFiles.ListItems.Count)
    .SubItems(1) = CStr(Round(lSize / 1024, 2)) & "KB"
    .SubItems(2) = CStr(dDateLastWrite)
End With

End Sub

Private Sub lvFiles_Click()
cmdDownload.Enabled = LenB(getSelectedFileName())
cmdDel.Enabled = cmdDownload.Enabled
End Sub
