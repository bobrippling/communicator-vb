VERSION 5.00
Begin VB.Form frmIPs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IPs"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin projMulti.ScrollListBox lstDates 
      Height          =   2295
      Left            =   3360
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   4048
   End
   Begin VB.Frame fraDev 
      Caption         =   "Dev Commands"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   6495
      Begin VB.CheckBox chkOnline 
         Height          =   255
         Left            =   5520
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox picCmdContainer 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   6255
         TabIndex        =   16
         Top             =   960
         Width           =   6255
         Begin VB.CommandButton cmdAddRec 
            Caption         =   "Add Record"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   1575
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove Selected Record"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1920
            TabIndex        =   18
            Top             =   120
            Width           =   2535
         End
         Begin VB.CommandButton cmdEnable 
            Caption         =   "Enable All"
            Height          =   375
            Left            =   4680
            TabIndex        =   19
            Top             =   120
            Width           =   1455
         End
      End
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblInfo2 
         Caption         =   "IP                                    Name                           Date                    Online"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Myself"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Upload IPs"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Download IPs"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin projMulti.ScrollListBox lstNames 
      Height          =   2295
      Left            =   1680
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   4048
   End
   Begin projMulti.ScrollListBox lstIPs 
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4048
   End
   Begin projMulti.ScrollListBox lstOnline 
      Height          =   2295
      Left            =   5040
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   4048
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label lblInfo 
      Caption         =   "IPs                           Names                           Dates                           Online"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   5415
   End
End
Attribute VB_Name = "frmIPs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Sep1 As String = "@", Sep2 As String = "#"
'ip record:
'xxx.xxx.xxx.xxx@Name@Date

Private Const NormHeight As Integer = 4350
Private Const DevHeight As Integer = 6210 ' NormHeight + 2000

Private CanClose As Boolean
Private AddOnlineStatus As Boolean 'was online status added (not uploaded yet)

Private Sub cmdAdd_Click()
Dim tRIP As String, Tmp As String
Dim i As Integer
Dim Ans As VbMsgBoxResult

tRIP = modWinsock.RemoteIP

If LenB(tRIP) Then
    For i = 0 To lstIPs.ListCount - 1
        If lstIPs.List(i) = tRIP Then
            
            Tmp = "'" & tRIP & "' is already in the list"
            
            Ans = MsgBoxEx(Tmp & vbNewLine & _
                "Replace it?", "IPs must be unique, otherwise everyone will get confused", _
                vbQuestion + vbYesNo, "Replace Name?", , , frmMain.Icon)
            
            If Ans = vbYes Then
                Call RemoveFromLists(i)
                Exit For 'and add to lists..
            Else
                SetStatus Tmp
                'cmdAdd.Enabled = False
                Exit Sub
            End If
            
            
        ElseIf lstNames.List(i) = frmMain.LastName Then
            
            Tmp = "'" & frmMain.LastName & "' is already in the list"
            
            Ans = MsgBoxEx(Tmp & vbNewLine & _
                "Replace it?", "You can't have two IPs, that's just greedy", _
                vbQuestion + vbYesNo, "Replace Name?", , , frmMain.Icon)
            
            If Ans = vbYes Then
                Call RemoveFromLists(i)
                Exit For 'and add to lists..
            Else
                SetStatus Tmp
                'cmdAdd.Enabled = False
                Exit Sub
            End If
            
            
        End If
    Next i
    
    Ans = MsgBoxEx("Set Online Status to True?", "Do you want to show other people that you're online? Or be a sneaky begger, and hide?", _
        vbYesNo + vbQuestion, "Online?") ', , , frmMain.Icon)
    
    AddOnlineStatus = (Ans = vbYes)
    
    AddToLists tRIP, frmMain.LastName, Date, (Ans = vbYes)
    cmdAdd.Enabled = False
    cmdUp.Enabled = True
    SetStatus "Added to List"
Else
    'MsgBox "Error - External IP Not obtained" & vbNewLine & _
            "Right click on the main window's status bar to obtain it.", _
            vbExclamation, "Error"
    SetStatus "External IP not obtained. Right click the status bar to obtain it"
    'If frmMain.HasFocus = False Then
        'FlashWin True
    'End If
End If

End Sub

Private Sub cmdAddRec_Click()
Dim Tmp As String

txtIP.Text = Trim$(txtIP.Text)
txtName.Text = Trim$(txtName.Text)
txtDate.Text = Trim$(txtDate.Text)

If LenB(txtName.Text) Then
    If LenB(txtIP.Text) Then
        If LenB(txtDate.Text) Then
            If IsDate(txtDate.Text) Then
                If IsIP(txtIP.Text) Then
                    
                    AddToLists txtIP.Text, txtName.Text, txtDate.Text, CBool(chkOnline.Value)
                    
                    
                    ValidateIPs Tmp
                    
                    AddIPs Tmp
                    
                    
                    cmdUp.Enabled = True
                Else
                    SetStatus "Please Enter a valid IP"
                End If
            Else
                SetStatus "Please Enter a valid Date"
            End If
        Else
            SetStatus "Please Enter a Date"
        End If
    Else
        SetStatus "Please Enter an IP"
    End If
Else
    SetStatus "Please Enter a Name"
End If
End Sub

Private Sub cmdConnect_Click()
Dim IP As String

IP = lstIPs.Text

Unload Me

frmMain.Connect IP
End Sub

Private Sub cmdDown_Click()
Dim RawFile As String
Dim OldIPs As String, NewIPs As String, sError As String
Dim Erro As Boolean, Changed As Boolean

cmdDown.Enabled = False
CanClose = False

SetStatus "Downloading IPs..."
modFTP.DownloadIPs RawFile, Erro, cmdDown, sError, True

'78.150.185.15@Lennis-II#19/01/2008@1
'86.1.244.76@Rob#19/01/2008@0

If Erro Then
    SetStatus "Error Downloading IPs" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
    cmdDown.Enabled = True
Else
    
    Call AddIPs(RawFile)
    
    Call ValidateIPs(RawFile, Changed)
    
    
    If Changed Then
        Call AddIPs(RawFile)
        SetStatus "Downloaded and Removed Old IPs"
        cmdUp.Enabled = True
    Else
        SetStatus "Downloaded IPs"
    End If
    
    cmdAdd.Enabled = True
End If

'cmdDown.Enabled = True
CanClose = True

End Sub

Private Sub AddIPs(ByVal RawFile As String)
Dim i As Integer, J As Integer, K As Integer
Dim Tmp As String
Dim IPs() As String

Call ClearList

IPs = Split(RawFile, vbNewLine, , vbTextCompare)

For i = LBound(IPs) To UBound(IPs)
    
    Tmp = Trim$(IPs(i))
    
    If LenB(Tmp) Then
        
        J = InStr(1, Tmp, Sep1, vbTextCompare)
        K = InStr(1, Tmp, Sep2, vbTextCompare)
        
        AddToLists Left$(Tmp, J - 1), _
                Mid$(Tmp, 1 + J, K - J - 1), _
                Mid$(Tmp, 1 + K, InStr(J + 1, Tmp, Sep1, vbTextCompare) - K - 1), _
                CBool(Right$(Tmp, 1))
        
    End If
Next i

End Sub

Private Sub cmdEnable_Click()
cmdUp.Enabled = True
cmdDown.Enabled = True
cmdAdd.Enabled = True
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer

i = lstIPs.ListIndex

Call RemoveFromLists(i)

End Sub

Private Sub cmdUp_Click()
Dim f As Integer ', i As Integer
Dim RawFile As String, sError As String
'Dim dt As Date

cmdUp.Enabled = False

ValidateIPs RawFile

'f = FreeFile()
'
'Open modFTP.FTP_IPLocal_File For Output As #f
'    Print #f, RawFile; '; prevents newline
'Close #f
'
'SetStatus "Uploading IPs..."
'Me.Refresh

'B = modFTP.UploadIPs(cmdUp, sError)


SetStatus "Uploading IPs..."
Me.Refresh


If modFTP.UploadIPs(cmdUp, sError, RawFile) Then
    SetStatus "Uploaded IPs Successfully"
    modFTP.OnlineStatusIs = AddOnlineStatus
    cmdDown.Enabled = True
    cmdUp.Enabled = False
Else
    SetStatus "Error Uploading" & IIf(LenB(sError), " (" & sError & ")", vbNullString)
    cmdUp.Enabled = True
End If

'On Error Resume Next
'Kill modFTP.FTP_IPLocal_File
'On Error GoTo 0

End Sub

Private Sub Form_Load()
Call ClearList

lblStatus.Caption = vbNullString
fraDev.Enabled = bDevMode
Me.height = IIf(bDevMode, DevHeight, NormHeight)
fraDev.Visible = bDevMode

If frmMain.mnuOptionsAdvDisplayVistaControls.Checked Then
    modDisplay.SetButtonIcon cmdConnect.hWnd, frmMain.GetCommandIconHandle()
End If

CanClose = True
Call FormLoad(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If CanClose Then
    Call FormLoad(Me, True)
Else
    Cancel = True
End If
End Sub

Private Sub lstDates_Click()
Call ListClick(lstDates.ListIndex)
End Sub

Private Sub lstIPs_Click()
If LenB(lstIPs.Text) Then
    cmdConnect.Enabled = True
End If
Call ListClick(lstIPs.ListIndex)
End Sub

Private Sub ListClick(ByVal i As Integer)

On Error Resume Next
lstIPs.ListIndex = i
lstNames.ListIndex = i
lstDates.ListIndex = i
lstOnline.ListIndex = i

If bDevMode Then
    cmdRemove.Enabled = (i <> -1)
    
    txtIP.Text = lstIPs.Text
    txtName.Text = lstNames.Text
    txtDate.Text = lstDates.Text
    chkOnline.Value = IIf(CBool(lstOnline.Text), 1, 0)
    
End If

On Error GoTo 0
End Sub

Private Sub lstNames_Click()
Call ListClick(lstNames.ListIndex)
End Sub

Private Sub AddToLists(ByVal IP As String, ByVal Name As String, ByVal dDate As Date, ByVal Online As Boolean)

lstIPs.AddItem IP, lstIPs.ListIndex + 1
lstNames.AddItem Name, lstNames.ListIndex + 1
lstDates.AddItem CStr(dDate), lstDates.ListIndex + 1
lstOnline.AddItem CStr(Online), lstOnline.ListIndex + 1

End Sub

Private Sub RemoveFromLists(ByVal i As Integer)
lstIPs.RemoveItem i
lstNames.RemoveItem i
lstDates.RemoveItem i
lstOnline.RemoveItem i
End Sub
 
Private Sub ValidateIPs(ByRef RawFile As String, Optional ByRef InvalidFound As Boolean)
'Dim RawFile As String
Dim i As Integer
Dim Dt As Date
Dim RawFileChanged As Boolean
Dim mRawFile As String
'Dim InvalidFound As Boolean

SetStatus "Validating IPs..."

For i = 0 To lstIPs.ListCount - 1
    Dt = lstDates.List(i)
    
    
    If DateDiff("d", Dt, Date) < 3 And (DateDiff("d", Dt, Date) >= 0) Then 'if the date difference > 2 days
        
        mRawFile = mRawFile & vbNewLine & lstIPs.List(i) & Sep1 & lstNames.List(i) & Sep2 & lstDates.List(i) & Sep1 & IIf(CBool(lstOnline.List(i)), 1, 0)
        
        RawFileChanged = True
        
    Else
        'exclude it
        InvalidFound = True
    End If
Next i

On Error Resume Next
mRawFile = Mid$(mRawFile, 3)

If InvalidFound Then
    SetStatus "Validated IPs - Invalid Record(s) Removed"
Else
    SetStatus "Validated IPs"
End If

RawFile = mRawFile

End Sub

Private Sub SetStatus(ByVal T As String)
lblStatus.Caption = "Status: " & T
lblStatus.Refresh
End Sub

Private Sub ClearList()
lstIPs.Clear
lstNames.Clear
lstDates.Clear
lstOnline.Clear
End Sub

Private Sub lstOnline_Click()
Call ListClick(lstOnline.ListIndex)
End Sub

Private Sub txtDate_LostFocus()
If IsDate(txtDate.Text) Then txtDate.Text = CDate(txtDate.Text)
End Sub
