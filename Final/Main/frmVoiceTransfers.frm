VERSION 5.00
Begin VB.Form frmVoiceTransfers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voice Transfers"
   ClientHeight    =   3585
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDebug 
      Caption         =   "Debug"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdDebugReset 
      Caption         =   "Reset Socket"
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtDebug 
      Height          =   2415
      Left            =   6600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   4695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtCurrent 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin projMulti.ScrollListBox lstTransfers 
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4260
   End
   Begin VB.Label lblDebug 
      Alignment       =   2  'Center
      Caption         =   "Debug Info"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7320
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   4695
   End
   Begin VB.Menu mnuVoice 
      Caption         =   "Voice Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuVoicePlay 
         Caption         =   "Play Recording"
      End
      Begin VB.Menu mnuVoiceRemove 
         Caption         =   "Remove From List"
      End
      Begin VB.Menu mnuVoiceSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVoiceFolder 
         Caption         =   "Open Recording Folder"
      End
      Begin VB.Menu mnuVoiceDel 
         Caption         =   "Delete Recording"
      End
   End
End
Attribute VB_Name = "frmVoiceTransfers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ptVoiceTransfer
    sFile As String
    bSent As Boolean
End Type

Private Transfers() As ptVoiceTransfer
Private nTransfers As Integer

Public bRecordingCanceled As Boolean

Private Sub chkDebug_Click()
Form_Resize
End Sub

Private Sub cmdClear_Click()
Erase Transfers
nTransfers = 0
RefreshList
End Sub

Private Sub Form_Load()
Me.Visible = False
chkDebug.Value = 0
Me.mnuVoice.Visible = False
txtCurrent.Enabled = False
cmdCancel.Enabled = False
Me.lstTransfers.ToolTipText = "Right Click for a Menu"

modLoadProgram.frmVoiceTransfers_Loaded = True
nTransfers = 0
End Sub

Public Sub ShowForm(Optional bShow As Boolean = True)

If bShow Then
    lblInfo.Caption = vbNullString
    
    If modDev.bDevMode() Then
        txtDebug.Visible = True
        lblDebug.Visible = True
        cmdDebugReset.Visible = True
        cmdDebugReset.Caption = "Reset Socket"
        Me.width = 11415
    Else
        txtDebug.Visible = False
        lblDebug.Visible = False
        cmdDebugReset.Visible = False
        Me.width = 6705
    End If
    
    FormLoad Me
    Me.Show vbModeless, frmMain
    RefreshList
Else
    FormLoad Me, True
    Me.Hide
End If

End Sub

Private Sub cmdDebugReset_Click()
frmMain.ucVoiceTransfer.Disconnect
cmdDebugReset.Caption = "Reset Socket... Done"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = vbFormControlMenu Then
    Cancel = 1
    ShowForm False
Else
    Erase Transfers
    nTransfers = 0
    modLoadProgram.frmVoiceTransfers_Loaded = False
End If

End Sub

Private Sub RemoveVoiceTransfer(Index As Integer)
Dim i As Integer

If nTransfers = 1 Then
    Erase Transfers
    nTransfers = 0
Else
    For i = Index To nTransfers - 2
        Transfers(i) = Transfers(i + 1)
    Next i
    
    nTransfers = nTransfers - 1
    ReDim Preserve Transfers(nTransfers - 1)
    
End If

RefreshList
End Sub

Public Sub AddVoiceTransfer(sFileName As String, bSent As Boolean)

ReDim Preserve Transfers(nTransfers)
With Transfers(nTransfers)
    .sFile = sFileName
    .bSent = bSent
End With
nTransfers = nTransfers + 1

RefreshList
End Sub

Private Sub RefreshList()
Dim i As Integer

lstTransfers.Clear

For i = 0 To nTransfers - 1
    lstTransfers.AddItem GetFileName(Transfers(i).sFile) '& _
        vbTab & vbTab & "[" & IIf(Transfers(i).bSent, "Sent", "Received") & "]"
    
Next i

cmdClear.Enabled = nTransfers > 0

End Sub

Private Sub Form_Resize()
Dim b As Boolean
b = Not CBool(chkDebug.Value)
If b Then
    Me.width = 6705
    lstTransfers.width = Me.width - lstTransfers.Left
Else
    Me.width = 11415
End If
txtDebug.Enabled = b
cmdDebugReset.Enabled = b
End Sub

Private Sub lstTransfers_DblClick()
mnuVoicePlay_Click
End Sub

Private Sub lstTransfers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then
    If LenB(lstTransfers.Text) Then
        PopupMenu mnuVoice, , , , mnuVoicePlay
    End If
End If

End Sub

Private Function getSelectedFile() As String
'Dim sFile As String
'Dim i As Integer

'sFile = lstTransfers.Text

'i = InStr(1, sFile, vbTab)
'getSelectedFile = Left$(sFile, i - 1)

getSelectedFile = Trim$(lstTransfers.Text)

End Function
Private Function getSelectedFileName() As String
Dim sFile As String
Dim i As Integer

sFile = getSelectedFile()

For i = 0 To nTransfers - 1
    If GetFileName(Transfers(i).sFile) = sFile Then
        getSelectedFileName = Transfers(i).sFile
        Exit Function
    End If
Next i

getSelectedFileName = vbNullString

End Function
Private Function getSelectedIndex() As String
Dim sFile As String
Dim i As Integer

sFile = getSelectedFile()

For i = 0 To nTransfers - 1
    If GetFileName(Transfers(i).sFile) = sFile Then
        getSelectedIndex = i
        Exit Function
    End If
Next i

getSelectedIndex = -1

End Function

Private Sub mnuVoicePlay_Click()
Dim sFileName As String

sFileName = getSelectedFileName()

If LenB(sFileName) Then
    lblInfo.Caption = "Playing/Played " & GetFileName(sFileName)
    modAudio.PlayFileNameSound sFileName
End If

End Sub

Private Sub mnuVoiceRemove_Click()
Dim i As Integer

i = getSelectedIndex()

If i > -1 Then
    RemoveVoiceTransfer i
    'list is refreshed
    
    lblInfo.Caption = "Remove From List"
End If

End Sub

Private Sub mnuVoiceDel_Click()
Dim i As Integer
Dim sFileName As String

i = getSelectedIndex()

If i > -1 Then
    If MsgBoxEx("Delete Recording, Are You Sure?", "Do you want to delete the recording file?", _
            vbQuestion + vbYesNo, "Delete Recording", _
            Me.ScaleX(Me.Left, vbTwips, vbPixels), _
            Me.ScaleY(Me.Top, vbTwips, vbPixels), , , Me.hWnd) = vbYes Then
        
        
        sFileName = Transfers(i).sFile
        
        
        RemoveVoiceTransfer i
        'list is refreshed
        
        
        'delete it
        On Error GoTo EH
        Kill sFileName
        lblInfo.Caption = "File Deleted"
    End If
End If

Exit Sub
EH:
lblInfo.Caption = "Error Deleting - " & Err.Description
End Sub

Private Sub mnuVoiceFolder_Click()
Dim sFileName As String

sFileName = getSelectedFileName()

If LenB(sFileName) Then
    OpenFolder vbNormalFocus, , sFileName
    lblInfo.Caption = "Folder Opened"
End If
End Sub

Public Sub updateCurrent(percentProgress As Byte, fileName As String, bSending As Boolean)

If LenB(fileName) = 0 Then
    txtCurrent.Text = "No Transfer in Progress"
    cmdCancel.Enabled = False
Else
    If bSending Then
        txtCurrent.Text = "Sending"
    Else
        txtCurrent.Text = "Receiving"
    End If
    
    txtCurrent.Selstart = Len(txtCurrent.Text)
    txtCurrent.SelText = " '" & Mid$(fileName, InStr(1, fileName, vbSpace) + 1) & _
        "'... (" & CStr(percentProgress) & "%)"
    
    cmdCancel.Enabled = True
End If

End Sub

Private Sub cmdCancel_Click()
cmdCancel.Enabled = False
bRecordingCanceled = True
frmMain.ucVoiceTransfer.Disconnect
End Sub

Public Sub addToDebug(S As String)

txtDebug.Selstart = Len(txtDebug.Text)
txtDebug.SelText = vbNewLine & S

End Sub
