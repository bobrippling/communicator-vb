VERSION 5.00
Begin VB.Form frmManual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect to IP"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   1425
   End
   Begin projMulti.ucListEdit ucListEditIP 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Width           =   1500
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1500
   End
   Begin VB.ComboBox cboIP 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "cboIP"
      Top             =   120
      Width           =   4695
   End
   Begin VB.CheckBox chkRetry 
      Caption         =   "Auto Retry, if connection fails"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Label lblF2Help 
      Alignment       =   2  'Center
      Caption         =   "Click on the item to edit it"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label lblF2 
      Caption         =   "Press F2 to Edit"
      Height          =   255
      Left            =   3705
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kCaption = "Connect to ", kDash = " - "

'################################################################
Private Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, cbInfo As COMBOBOXINFO) As Long

Private Const EM_LIMITTEXT = &HC5, _
              Max_Len = 80

Private Const NormHeight = 1980, ExtHeight = 5055

'################################################################

Private Sub cboIP_Change()
Dim sTxt As String
Dim bEn As Boolean

sTxt = cboIP.Text
cboIP.Text = Trim$(sTxt)

bEn = CBool(LenB(sTxt))
cmdConnect.Enabled = bEn

'cmdConnect.Default = cmdConnect.Enabled

cboIP_Click

End Sub

Private Sub showEdit(bShow As Boolean)
Dim i As Integer

Me.ucListEditIP.Visible = bShow
cmdDel.Visible = bShow

If bShow Then
    Me.height = ExtHeight
Else
    Me.height = NormHeight
End If

End Sub

Private Sub cboIP_KeyDown(KeyCode As Integer, Shift As Integer)

cboIP_Scroll

End Sub

Private Sub cboIP_Scroll()
cboIP_Change
End Sub

Public Sub cboIP_Click()
Dim sIP As String

sIP = GetCurIP()

If LenB(sIP) Then
    Me.Caption = kCaption & sIP
Else
    Me.Caption = "Manual Connect"
End If


End Sub

Private Function GetCurIP() As String
Dim i As Integer
Dim sTxt As String

sTxt = cboIP.Text

If LenB(sTxt) Then
    'On Error Resume Next
    'sTxt = cboIP.List(0)
    i = InStr(1, sTxt, vbSpace)
    
    If i Then
        GetCurIP = Left$(sTxt, i - 1)
    Else
        GetCurIP = sTxt
    End If
End If

End Function

Private Sub cmdCancel_Click()
modVars.bRetryConnection = False
Unload Me
End Sub

Private Sub cmdClear_Click()

If MsgBoxEx("Are you sure you want to clear the list of IPs?", _
        "Clearing the list will remove all the IPs from Communicator's memory", vbQuestion Or vbYesNo, _
        "Clear IP List", , , , , Me.hWnd) = vbYes Then
    
    cmdClear.Enabled = False
    
    cboIP.Clear
    Me.ucListEditIP.listViewObject.ListItems.Clear
    
    ReDim UsedIPs(0)
End If

End Sub

Private Sub cmdConnect_Click()
Dim IP As String
Dim bRetry As Boolean
Dim i As Integer, port As Integer

IP = Trim$(GetCurIP())

i = InStr(1, IP, ":")
If i Then
    On Error GoTo EH
    port = CInt(Mid$(IP, i + 1))
    If Not (1 <= port And port <= 10000) Then
        GoTo EH
    End If
    IP = Left$(IP, i - 1)
    If Status <> Idle Then
        frmMain.CleanUp Status = Connected
    End If
    modPorts.MainPort = port
End If
    

bRetry = CBool(chkRetry.Value)

If bRetry Then
    'LastAutoRetry = GetTickCount()
    'AddText "You can hide/close Communicator and it'll pop up when it connects", , True
    AddText "Connecting (with Auto-Retry) to " & IP & ":" & MainPort & "...", , True
End If


Unload Me
DoEvents

modVars.bRetryConnection = bRetry
modVars.bRetryConnection_Static = bRetry
frmMain.Connect IP 'Cmds() will turn off bRetryConnection, so force it back on
modVars.bRetryConnection = bRetry

Exit Sub
EH:
MsgBoxEx "A socket address format is IP:Port, port must be a number between 1 and 10,000", "Get rid of the colon", vbExclamation, _
    "IP Address", , , , , Me.hWnd
End Sub

Private Sub PopulateList()
Dim i As Integer

cboIP.Clear

For i = 0 To UBound(UsedIPs)
    If LenB(UsedIPs(i).sIP) Then
        
        AddToList UsedIPs(i).sIP, UsedIPs(i).sName
        
        
    End If
Next i

If LenB(modMessaging.UsedIPs(0).sIP) Then
    cboIP.ListIndex = 0
     'Set_cbo_Text UsedIPs(0).sIP, UsedIPs(0).sName
End If

End Sub

Private Sub AddToList(sIP As String, sName As String)
If LenB(sName) Then
    cboIP.AddItem sIP & kDash & sName
Else
    cboIP.AddItem sIP
End If

cmdClear.Enabled = True
End Sub
Private Sub Set_cbo_Text(sIP As String, sName As String)
If LenB(sName) Then
    cboIP.Text = sIP & kDash & sName
Else
    cboIP.Text = sIP
End If
End Sub

Private Sub cmdDel_Click()
Dim i As Integer

'cmdDel.Enabled = False

With Me.ucListEditIP.listViewObject
    If Not (.SelectedItem Is Nothing) Then
        i = .SelectedItem.Index
        .ListItems.Remove i
        
        If .ListItems.Count > 0 Then
            If i > 1 Then
                .ListItems(i - 1).Selected = True
            Else
                .ListItems(1).Selected = True
            End If
        End If
    End If
End With
ucListEditIP.HideEditBox


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    showEdit Not Me.ucListEditIP.Visible
    cmdConnect.Default = False
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim b As Boolean
Dim cbInfo As COMBOBOXINFO
Const cchMax = Max_Len

modVars.bModalFormShown = True
chkRetry.Value = Abs(modVars.bRetryConnection_Static)

'###########################
cbInfo.cbSize = Len(cbInfo)
GetComboBoxInfo cboIP.hWnd, cbInfo

SendMessageByLong cbInfo.hwndEdit, EM_LIMITTEXT, cchMax, 0
'###############################################################

With Me.ucListEditIP.listViewObject
    .ColumnHeaders.Add , , "IP"
    .ColumnHeaders.Add , , "Name"
    
    .Checkboxes = False
End With
Me.ucListEditIP.Visible = False
cmdDel.Visible = False
Me.ucListEditIP.addColumnOfIntrest 0
Me.ucListEditIP.addColumnOfIntrest 1

Me.ucListEditIP.listViewObject.ListItems.Clear

For i = 0 To UBound(UsedIPs)
    If LenB(UsedIPs(i).sIP) Then
        
        ucListEditIP.listViewObject.ListItems.Add , , UsedIPs(i).sIP
        
        With ucListEditIP.listViewObject.ListItems(ucListEditIP.listViewObject.ListItems.Count)
            .SubItems(1) = UsedIPs(i).sName
            '.Checked = True
        End With
    End If
Next i


Me.ucListEditIP.UserControl_Resize
'###############################################################


cmdClear.Enabled = False

Me.Caption = kCaption & "IP"
Me.height = 1920
Set_cbo_Text "127.0.0.1", "localhost"

PopulateList

If LenB(modMessaging.LastIP) Then
    For i = 0 To UBound(UsedIPs)
        If UsedIPs(i).sIP = modMessaging.LastIP Then
            'cboIP.Text = modMessaging.UsedIPs(i).sIP & kDash & UsedIPs(i).sName
            'Set_cbo_Text UsedIPs(i).sIP, UsedIPs(i).sName
            cboIP.ListIndex = i
            Exit For
        End If
    Next i
End If
cboIP_Click


b = CBool(LenB(cboIP.Text))
cmdConnect.Enabled = b
cmdCancel.Default = Not b
cmdConnect.Default = b


If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdConnect.hWnd, frmMain.GetCommandIconHandle()
End If

Call FormLoad(Me)

'bValidIP = cboIP.HasValidIP

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer

With Me.ucListEditIP.listViewObject
    ReDim modMessaging.UsedIPs(0)
    'modMessaging.CurIPIndex = 0 ... should sort itself out?
    
    For i = 1 To .ListItems.Count
        'If .ListItems(i).Checked Then
            modMessaging.AddUsedIP .ListItems(i).Text, .ListItems(i).ListSubItems(1)
        'End If
    Next i
End With

FormLoad Me, True
modVars.bModalFormShown = False
End Sub
