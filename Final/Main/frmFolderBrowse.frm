VERSION 5.00
Begin VB.Form frmFolderBrowse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse for Folder"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4425
   Begin VB.CommandButton cmdNewFolder 
      Caption         =   "New Folder"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeleteFolder 
      Caption         =   "Delete Folder"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.DirListBox oDir 
      Height          =   2115
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3975
   End
   Begin VB.DriveListBox oDrive 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label lblPath 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      Caption         =   "Choose a folder."
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmFolderBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Property Let InitDir(nVal As String)
SetDir nVal
End Property

Private Sub cmdCancel_Click()
modVars.pBrowse_FolderPath = vbNullString
Unload Me
End Sub

Private Sub cmdDeleteFolder_Click()
Dim Ans As Integer
Dim i As Integer
Dim Path As String

Path = Left$(oDir.Path, Len(Path) - InStrRev(Path, "\", , vbTextCompare))


Ans = MsgBox("Delete '" & oDir.Path & "'" & vbCrLf & "Are You Sure?", vbYesNo + vbQuestion, "Delete Folder")
If Ans = vbNo Then Exit Sub

On Error Resume Next
Kill oDir.Path & "\*"

RmDir oDir.Path
oDir.Path = Path
oDir.Refresh
End Sub

Private Sub cmdNewFolder_Click()
Dim Name As String
Name = InputBox("Enter a Name for the folder", "New Folder", "New Folder")

If LenB(Name) Then
    
    If Dir$(oDir.Path & "\" & Name, vbDirectory) <> "" Then
        MsgBox "Folder Already Exists", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    MkDir oDir.Path & "\" & Name
    
    oDir.Refresh
End If

End Sub

Private Sub cmdOK_Click()
modVars.pBrowse_FolderPath = oDir.Path
Unload Me
End Sub

Private Sub oDir_Change()
SetDir oDir.Path
End Sub

Private Sub oDrive_Change()
On Error GoTo ErrH

SetDir oDrive.Drive

Exit Sub
ErrH:
MsgBox "Error: " & Err.Description, vbOKOnly + vbExclamation, "Error"
SetDir oDir.Path
End Sub

Private Sub SetDir(Path As String)
Dim SubF As String
Dim i As Integer, j As Integer
Dim nBackSlashes As Integer

oDir.Path = Path
oDrive.Drive = Left$(Path, 2)

For i = 1 To Len(Path)
    If Mid$(Path, i, 1) = "\" Then
        nBackSlashes = nBackSlashes + 1
    End If
Next i

If nBackSlashes > 3 Then 'more than 3 folders deep, add dots
    j = InStrRev(Path, "\")
    
    i = InStrRev(Path, "\", j - 1)
    If i = 0 Then
        i = j
    Else
        i = InStrRev(Path, "\", i - 1)
        If i = 0 Then i = j
    End If
    
    SubF = Mid$(Path, i)
    lblPath.Caption = Left$(Path, 3) & "..." & SubF
Else
    lblPath.Caption = oDir.Path
End If

End Sub

Private Sub Form_Load()
oDir_Change
'Call FormLoad(Me)
End Sub
