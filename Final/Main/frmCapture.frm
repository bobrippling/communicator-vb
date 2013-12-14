VERSION 5.00
Begin VB.Form frmCapture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video Capture"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   562
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSend 
      Interval        =   50
      Left            =   4320
      Top             =   2640
   End
   Begin VB.PictureBox picExtraBG 
      Height          =   1815
      Left            =   3360
      ScaleHeight     =   1755
      ScaleWidth      =   3195
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Frame fraExtras 
         Caption         =   "Extras"
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   10
         Width           =   3015
         Begin VB.PictureBox picExtraContainter 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            ScaleHeight     =   1335
            ScaleWidth      =   2775
            TabIndex        =   9
            Top             =   240
            Width           =   2775
            Begin VB.CommandButton cmdCompression 
               Caption         =   "Compression"
               Height          =   375
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   1335
            End
            Begin VB.CommandButton cmdCopy 
               Caption         =   "Copy to Clipboard"
               Height          =   375
               Left            =   0
               TabIndex        =   13
               Top             =   960
               Width           =   2775
            End
            Begin VB.CommandButton cmdFormat 
               Caption         =   "Format"
               Height          =   375
               Left            =   1440
               TabIndex        =   12
               Top             =   0
               Width           =   1335
            End
            Begin VB.CommandButton cmdDisplay 
               Caption         =   "Display"
               Height          =   375
               Left            =   1440
               TabIndex        =   11
               Top             =   480
               Width           =   1335
            End
            Begin VB.CommandButton cmdSource 
               Caption         =   "Source"
               Height          =   375
               Left            =   0
               TabIndex        =   10
               Top             =   480
               Width           =   1335
            End
         End
      End
   End
   Begin VB.Frame fraControls 
      Caption         =   "Controls"
      Height          =   2775
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox picControls 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   3375
         TabIndex        =   2
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox chkViewExtras 
            Caption         =   "View Extras"
            Height          =   255
            Left            =   2040
            TabIndex        =   6
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkCap 
            Caption         =   "Capture"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1440
            Width           =   975
         End
         Begin VB.CommandButton cmdInit 
            Caption         =   "Select Device"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   3
            Top             =   960
            Width           =   3255
         End
         Begin projMulti.ScrollListBox lstDevice 
            Height          =   975
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   3375
            _ExtentX        =   3413
            _ExtentY        =   5106
         End
      End
   End
   Begin VB.PictureBox picCap 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const UseCallbacks = False
Private hCapWnd As Long
Private bRunning As Boolean

Private Devices() As ptDevice
Private nCurrentDriver As Integer

'socket stuff
Private hSocket As Long
Private CamServer As Boolean

Private Function InitSocket(ByVal bHost As Boolean) As Boolean

hSocket = modWinsock.CreateSocket()

If hSocket = WINSOCK_ERROR Then
    InitSocket = False
Else
    CamServer = bHost
    
    If CamServer Then
        If modWinsock.BindSocket(hSocket, Video_Port) <> WINSOCK_ERROR Then
            InitSocket = True
        Else
            InitSocket = False
        End If
    End If
End If

End Function

Private Sub TermWinsock()

modWinsock.DestroySocket hSocket

End Sub

Private Sub SendCurrentImage()
Dim Data As String
Dim CurPic As Image

modCapture.capEditCopy hCapWnd
Picture = Clipboard.GetData()

Data = Picture

modWinsock.SendPacket hSocket, ClientSockAddr, Data

End Sub

'end socket stuff
'#####################################################################


Private Sub ListDevices()
Dim NDevs As Integer, i As Integer

NDevs = modCapture.VBEnumCapDrivers(Devices)

If NDevs Then
    
    For i = 0 To UBound(Devices)
        lstDevice.AddItem Devices(i).sName
        'lstVersion.AddItem Devices(i).sVersion
    Next i
    
    'lstDevice.ListIndex = 0
End If

End Sub

Private Sub Init()
    
'//Create Capture Window
'Call capGetDriverDescription( nDriverIndex,  lpszName, 100, lpszVer, 100  '// Retrieves driver info
hCapWnd = modCapture.capCreateCaptureWindow("Communicator VID CAP WINDOW", WS_CHILD Or WS_VISIBLE, 0, 0, 640, 480, picCap.hWnd, 0)

If hCapWnd Then
    
    If modCapture.ConnectCapDriver(hCapWnd, nCurrentDriver) Then
        #If UseCallbacks Then
            'if we have a valid capwnd we can enable our status callback function
            Call modCapture.capSetCallbackOnStatus(hCapWnd, AddressOf StatusProc)
            'Debug.Print "---Callback set on capture status---"
            
            '// Set the video stream callback function
            'capSetCallbackOnVideoStream lwndC, AddressOf MyVideoStreamCallback
            'capSetCallbackOnFrame lwndC, AddressOf MyFrameCallback
            
        #End If
    Else
        MsgBox "Could not connect to capture driver", vbExclamation, "Error"
    End If
Else
    MsgBox "Could not create capture window", vbExclamation, "Error"
End If

End Sub

Private Sub Terminate()

'unsubclass if necessary
#If UseCallbacks Then
    ' Disable status callback
    Call capSetCallbackOnStatus(hCapWnd, 0&)
    'Debug.Print "---Capture status callback released---"
#End If

'disconnect VFW driver
Call capDriverDisconnect(hCapWnd)

'destroy CapWnd
If hCapWnd <> 0 Then Call DestroyWindow(hCapWnd)

hCapWnd = 0
nCurrentDriver = -1

End Sub

Private Sub EnableCapture(Optional ByVal bEn As Boolean = True)
Call capPreview(hCapWnd, bEn)
End Sub

Private Sub cmdFormat_Click()
Call capDlgVideoFormat(hCapWnd)
'Call ResizeCaptureWindow(hCapWnd)
End Sub

Private Sub cmdCompression_Click()
Call capDlgVideoCompression(hCapWnd)
End Sub

Private Sub cmdCopy_Click()
Call capEditCopy(hCapWnd)
End Sub

Private Sub cmdDisplay_Click()
Call capDlgVideoDisplay(hCapWnd)
End Sub

Private Sub cmdSource_Click()
Call capDlgVideoSource(hCapWnd)
End Sub

Private Sub CapCmds(ByVal bE As Boolean)
chkCap.Enabled = bE
chkViewExtras.Enabled = bE
End Sub

'#################################################################################
'#################################################################################

Private Sub chkCap_Click()
'If chkCap.Value Then
'    If nCurrentDriver > -1 Then
'        StartCap
'    Else
'        chkCap.Value = 0
'        chkCap.Enabled = False
'    End If
'Else
'    EndCap
'End If
EnableCapture CBool(chkCap.Value)
End Sub

Private Sub chkViewExtras_Click()
picExtraBG.Visible = CBool(chkViewExtras.Value)
End Sub

Private Sub cmdInit_Click()

If bRunning Then
    cmdInit.Caption = "Select Device"
    CapCmds False
    
    Terminate
    bRunning = False
    
    lstDevice.ListIndex = -1
Else
    cmdInit.Caption = "Stop Device"
    
    chkCap.Value = 1
    CapCmds True
    
    Init
    bRunning = True
End If

End Sub

Private Sub lstDevice_Click()
nCurrentDriver = lstDevice.ListIndex
cmdInit.Enabled = (nCurrentDriver > -1)
If cmdInit.Enabled Then
    cmdInit.Caption = "Use Device"
End If

'Call ListClick(nCurrentDriver)

End Sub

'Private Sub ListClick(i As Integer)
'On Error Resume Next
'lstDevice.ListIndex = i
''lstVersion.ListIndex = i
'End Sub

Private Sub Form_Load()
nCurrentDriver = -1
picExtraBG.BorderStyle = 0
CapCmds False

ListDevices

Call FormLoad(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Terminate

hCapWnd = 0
bRunning = False
Erase Devices()
nCurrentDriver = -1

Call FormLoad(Me, True)
End Sub

'Private Sub lstVersion_Click()
'Call ListClick(lstVersion.ListIndex)
'End Sub
