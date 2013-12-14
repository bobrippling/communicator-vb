Attribute VB_Name = "modVars"
Option Explicit
'general Constant
Public Const WM_USER = &H400

Public Const vbSpace = " "
Public Const frmPrivateName = "frmPrivate"
Public Const FileExt As String = "mcc"

Public RootDrive As String '="C:"
Public Comm_Safe_Path As String 'i.e. a path safe to write to (in Vista)
                                'has trailing \
Private pPC_Name As String, pUser_Name As String

'used in modconsole and modimplode.moveform and modAlert
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'setwindowpos
Private Const SWP_SHOWWINDOW = &H40, SWP_NOACTIVATE = &H10

'always on top
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

'to update frame:
'const SWP_FRAMECHANGED = whatever
Public Const WM_SETICON = &H80
Public Const ICON_BIG = 1
Public Const ICON_SMALL = 0


Public bCleanedUp As Boolean

'##################################################################################################
'system time
Private Declare Function apiSetSystemTime Lib "kernel32" Alias "SetSystemTime" (lpSystemTime As SYSTEMTIME) As Long

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type
'##################################################################################################

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const lstConnectedNormLeft = 1680
Private Const lstConnectedExtendedWidth = 3015
Public Const EM_GETSCROLLPOS = WM_USER + 221
Public Const EM_SETSCROLLPOS = WM_USER + 222

'for popupmenu in systray
Public bModalFormShown As Boolean

'taskbar height
Private Const SPI_GETWORKAREA = &H30 '48
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long


'dragging
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1


'auto connect stuff
Public bRetryConnection As Boolean, bRetryConnection_Static As Boolean
'Public LastAutoRetry As Long


'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Const WM_SETREDRAW = &HB
Private Const EM_GETEVENTMASK = (WM_USER + 59)
Private Const EM_SETEVENTMASK = (WM_USER + 69)


'sck errors
Public Const WSANO_DATA = 11004
Public Const WSAEADDRINUSE = 10048
Public Const WSAECONNREFUSED = 10061
Public Const WSAHOSTNOTFOUND = 11001
Public Const CustomLagError = -2

Public Const ConnectedListPlaceHolder = "Updating List..."

''file transfer
'Public bFT_AutoAccept As Boolean


'anti-tooltip stuff
Public Const TBM_SETTOOLTIPS = WM_USER + 29


''non breaking space
'Public Const NBSP As String = vbspace
Public Const DefaultFontName = "MS Sans Serif"
Public Const DefaultFontSize = 8

'ip choosing
Public pIP_Choice As String

'browse for folder
Public pBrowse_FolderPath As String

'for foreground window check
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

'window style etc etc
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000

Public Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hWnd As Long, _
  ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hWnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

'Public Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hWnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'end window style


'getwindowsversion
Private Declare Function apiGetVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  'High-order word contains major & minor ver nos.
  PlatformId      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
End Type                          'Win9x: string providing arbitrary additional information


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
'end version

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
            Destination As Any, Source As Any, ByVal Length As Long)

'for ping
Private Declare Function GetRTTAndHopCount Lib "iphlpapi.dll" ( _
    ByVal lDestIPAddr As Long, _
    ByRef lHopCount As Long, _
    ByVal lMaxHops As Long, _
    ByRef lRTT As Long) As Long


Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
'end ping


'for version stuff
Public Const Dot As String = "."


Public Type ptRGB
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Const sOpen As String = "open"
Public Const SW_SHOWNORMAL As Long = 1 'for showing it at front/back
Public Const SW_HIDE As Long = 0
Public Const SW_SHOW = 5
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

'for getlastdllerrstr
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long

'for openurl()
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'flag if a port has been forwarded
Public APortForwarded As Boolean

'end

Public OnlineStatus As Boolean '1 = set to online, 0 = set to offline
'Public CanUseInet As Boolean


Public Const PvtCap As String = "Private Comm Channel - "

'will look at updateurl & updatetxt for new version
'will open updateurl for download
Private puPassword As String

Public nPrivateChats As Integer

Private pbDebug As Boolean
Public bStartup As Boolean
Public bNoInternet As Boolean
Public bStealth As Boolean
Public bDisableAddText As Boolean

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function apiGetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long

'####################################################################################################
'Send Messages
Public Declare Function SendMessageByLong Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
        wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessageByString Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
        wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function SendMessageByAny Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
        wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'####################################################################################################

'Private theTitle As String
'Private theHwnd As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lparam As Long) As Long

'Ports are relative to when connecting

'Public lIP As String
'Public rIP As String

Public ConsoleShown As Boolean

'Public Const SafeConfirm As Byte = 1
'Public Property Get SafeFile() As String
'SafeFile = AppPath() & "Communicator.dat"
'End Property

Public SocketCounter As Long

Public Status As eStatus
Private pServer As Boolean
Public Clients() As ptClient

Private Type ptClient
    iSocket As Integer
    sName As String
    sIP As String
    sStatus As String
    bShownConnection As Boolean 'has the "Timmy Connected" popup been shown?
    
    BlockDrawing As Boolean
    
    sVersion As String
    
    iPing As Integer
    lPingStart As Long
    lLastPing As Long
    
    IPicture As IPictureDisp
    
    'iRequestedDP As Integer
    'iLastDPSent As Integer 'last other client's picture sent
    bSentHostDP As Boolean
    sHasiDPs As String 'lists client [b]indexs[/b] of pictures this client has
                       'e.g. if this client has DPs from iclient=2,-1, it will = "2,-1"
    bDPSet As Boolean
    
    bDPIsGIF As Boolean
End Type

Public WState As FormWindowStateConstants
Public InTray As Boolean

Public Closing As Boolean

Private Const Capt As String = "Communicator"

Public Enum eStatus
    Idle = 0
    Connected = 1
    Connecting = 2
    Listening = 3
End Enum

Public Enum eCommands
    cmdOther = 0
    
    Message = 1
    Draw = 2
    Typing = 3
    ClientList = 4
    'GetName = 5
    'ReplyName = 6

    FileTransferCmd = 5

    SetSocket = 6

    '###New###
    'mPing = 5

    Shake = 7
    matrixMessage = 8
    'Invite = 9
    LobbyCmd = 9

    DevSend = -1
    DevRecieve = -2
    Info = -3
    Drawing = -4 'like Typing
    Prvate = -5
    SetClientVar = -6

    SetTyping = -7

    PingCmd = -8
    
    HostCmd = -9
End Enum

Public Enum eOtherCmds
    SetServerName = 0
    ConnectToServerVoicePort '= 1
End Enum

Public Enum eHostCmds
    RemoveDP = 1
End Enum

Public Enum eClientVarCmds
    SetName = 0
    SetDrawing = 1
    'SetIndex = 2#
    SetVersion = 2
    'SetSocket = 3#
    'SetRequestedDP = 4#
    SetStatus = 3
    SetDPSet = 4
    SetsStatus = 5
End Enum

Public Enum ePingCmds
    aPing = 1
    aPong = 2
End Enum

Public Enum eFTCmds 'file transfer
    FT_SendDPToHost = 0
    'FT_Listen = 1
    FT_Close = 1
    FT_ConnectToHost = 2
End Enum

'Public Enum eLobbyCmds
'    'Add = 0
'    'Remove = 1
'
'    Request = 2
'    Reply = 3
'End Enum

Public Const InfoStart As String = "----- "
Public Const InfoEnd As String = " -----"

Public TxtReceived As Long 'colours
Public TxtSent As Long 'i.e. shakes. For type colour, see txtforeground
Public TxtInfo As Long
Public TxtError As Long
Public TxtUnknown As Long
Public TxtQuestion As Long
Private pTxtForeGround As Long
Private pTxtBackGround As Long
Public Const MGreen As Long = &HBD700 '775936 'RGB(0, 222, 0)
Public Const MOrange As Long = &HDA6F0 '894704 'RGB(247, 174, 0)
Public Const MBrown As Long = &H165185 '1462661
Public Const MPurple As Long = &HFF00FF '16711935 '15921677
Public Const MGrey As Long = &H74778B '7632779
Public Const MLightBlue As Long = &HFFFF00 '16776960
Public Const MSilver As Long = &HC0C0C0


Public Type PointAPI
    X As Long
    Y As Long
End Type

Private pSniperMode As Boolean 'aka pbStealthMode

Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" ( _
    ByVal lpBuffer As String, lpnSize As Long) As Long 'c bool


'########################################################
'File transfer
Public Type ptFileTransfer
    sName As String
    bReceived As Boolean
End Type
Public TransferFilePaths() As ptFileTransfer
Public nFilePaths As Integer
'########################################################


'##################################################################################
'last input
'Private Type LASTINPUTINFO
'    cbSize As Long
'    dwTime As Long
'End Type
'Private Declare Function GetLastInputInfo Lib "user32" (ByVal plii As PLASTINPUTINFO) As Long
'##################################################################################

Public Function FormatTimeElapsed(ByVal lSeconds As Single) As String
Const Sixty As Single = 60, TwentyFour As Single = 24, Seven As Single = 7
Dim MinsUp As Single, HoursUp As Single, DaysUp As Single, WeeksUp As Single
Dim UpTimeText As String

If lSeconds >= Sixty Then
    MinsUp = lSeconds / Sixty
    
    If MinsUp >= Sixty Then
        HoursUp = MinsUp / Sixty
        
        If HoursUp >= TwentyFour Then
            DaysUp = HoursUp / TwentyFour
            
            If DaysUp >= Seven Then
                WeeksUp = DaysUp / Seven
                
                UpTimeText = Round(WeeksUp, 1) & " week" & IIf(WeeksUp <> 1, "s", vbNullString)
                
            Else
                UpTimeText = Round(DaysUp, 1) & " day" & IIf(DaysUp <> 1, "s", vbNullString)
            End If
        Else
            UpTimeText = Round(HoursUp, 1) & " hour" & IIf(HoursUp <> 1, "s", vbNullString)
        End If
    Else
        UpTimeText = Round(MinsUp, 1) & " min" & IIf(MinsUp <> 1, "s", vbNullString)
    End If
Else
    UpTimeText = Round(lSeconds, 2) & " second" & IIf(lSeconds <> 1, "s", vbNullString)
End If

FormatTimeElapsed = UpTimeText

End Function

Public Function SetSystemTime(dDateTime As Date, ByVal msAdj As Long) As Boolean
Dim lR As Long
Dim TimeType As SYSTEMTIME
Dim iSeconds As Integer

With TimeType
    .wYear = Year(dDateTime)
    .wMonth = Month(dDateTime)
    .wDayOfWeek = Day(dDateTime)
    .wHour = Hour(dDateTime)
    .wDay = Day(dDateTime)
    
    
    If msAdj >= 1000 Then
        
        
        .wMilliseconds = msAdj Mod 1000
        iSeconds = Second(dDateTime) + msAdj \ 1000
        
        
        If iSeconds > 60 Then
            .wSecond = iSeconds Mod 60
            .wMinute = Minute(dDateTime) + iSeconds \ 60
        Else
            .wSecond = iSeconds
            .wMinute = Minute(dDateTime)
        End If
    Else
        .wMilliseconds = msAdj
        .wMinute = Minute(dDateTime)
        .wSecond = Second(dDateTime)
    End If
        
        
End With


'Set system time with new data
SetSystemTime = (apiSetSystemTime(TimeType) <> 0)

End Function

Public Function FindClient(ByVal iSock As Integer) As Integer
Dim i As Integer

FindClient = -1

If iSock = 0 Then Exit Function

For i = 0 To UBound(Clients)
    If Clients(i).iSocket = iSock Then
        FindClient = i
        Exit For
    End If
Next i

End Function

Public Sub Vars_Init()

Comm_Safe_Path = GetUserSettingsPath() 'has a \ at the end
RootDrive = Left$(Comm_Safe_Path, 2)
bRetryConnection_Static = True

GetUserInfo

End Sub

Public Property Get PC_Name() As String
PC_Name = pPC_Name
End Property
Public Property Get User_Name() As String
User_Name = pUser_Name
End Property

Private Sub GetUserInfo()
Dim sBuffer As String
Dim nSize As Long
Const nSize_n As Long = 255&

nSize = nSize_n
sBuffer = String$(nSize, 0)
If apiGetComputerName(sBuffer, nSize) Then
    pPC_Name = Left$(sBuffer, nSize)
Else
    AddConsoleText "Error Getting PC Name " & CStr(Err.LastDllError)
End If

nSize = nSize_n
sBuffer = String$(nSize, 0)
If apiGetUserName(sBuffer, nSize) Then
    pUser_Name = Left$(sBuffer, nSize - 1)
Else
    AddConsoleText "Error Getting User Name " & CStr(Err.LastDllError)
End If

End Sub

Public Sub SetMiniInfo(ByVal sTxt As String)

If modLoadProgram.frmMini_Loaded Then
    frmMini.SetInfo sTxt
End If

End Sub

Public Function GetTaskbarHeight() As Integer
Dim lRes As Long
Dim rectVal As RECT

lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)

GetTaskbarHeight = Screen.height - rectVal.Bottom * Screen.TwipsPerPixelY

End Function

Public Property Let bDebug(ByVal bVal As Boolean)

pbDebug = bVal

If modLoadProgram.bLoading = False Then
    frmMain.mnuDevAdvCmdsDebug.Checked = pbDebug
End If

End Property

Public Property Get bDebug() As Boolean

bDebug = pbDebug

End Property

Public Sub GetWindowsVersion( _
      Optional ByRef lMajor As Long = 0, _
      Optional ByRef lMinor As Long = 0, _
      Optional ByRef lRevision As Long = 0, _
      Optional ByRef lBuildNumber As Long = 0, _
      Optional ByRef bIsNt As Boolean = False, _
      Optional ByRef bVistaOrW7OrW7 As Boolean = False)

Static osv As OSVERSIONINFO
Static bDone As Boolean
Static lR As Long

If bDone = False Then
    
    lR = apiGetVersion()
    
    osv.OSVSize = Len(osv)
    
    If GetVersionEx(osv) = 1 Then
        bDone = True
    End If
End If

lMajor = osv.dwVerMajor
lMinor = osv.dwVerMinor
lRevision = (lR And &HFF0000) \ &H10000
'lBuildNumber = osv.dwBuildNumber
'bIsNt = (osv.PlatformID > 4)
'bVistaOrW7 = (osv.PlatformID = 6)


'################################
'old method
'Dim lR As Long
'
'lR = apiGetVersion()
'
lBuildNumber = (lR And &H7F000000) \ &H1000000

If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80

'lRevision = (lR And &HFF0000) \ &H10000
'lMinor = (lR And &HFF00&) \ &H100
'lMajor = (lR And &HFF)

bIsNt = ((lR And &H80000000) = 0)
bVistaOrW7OrW7 = (lMajor = 6 Or lMajor = 7)

End Sub

Public Function GetFileName(sPath As String) As String
GetFileName = Mid$(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFilePath(sPath As String) As String
GetFilePath = Left$(sPath, InStrRev(sPath, "\"))
End Function

Public Sub OpenFolder(ByVal WinStyle As VbAppWinStyle, Optional ByVal Path As String = vbNullString, _
    Optional ByVal SelectedFile As String = vbNullString)

If LenB(Path) = 0 Then
    If LenB(SelectedFile) = 0 Then
        Path = AppPath()
    End If
End If

If LenB(SelectedFile) Then
    On Error Resume Next
    Shell "explorer.exe /select," & SelectedFile, WinStyle
Else
    On Error Resume Next
    Shell "explorer.exe " & Path, WinStyle
End If
'explorer.exe /select,d:\work\My Picture.jpg

End Sub

Public Function OpenNewCommunicator(Optional additionalParams As String = vbNullString) As Boolean
Dim cmd As String
Dim lTmr As Long, l As Long
Const Communicator_Close_Wait = 5000

cmd = Trim$(Command() & vbSpace & additionalParams)

If InStr(1, cmd, "/startup", vbTextCompare) Then
    cmd = Replace$(cmd, "/startup", vbNullString, , , vbTextCompare)
End If
'If InStr(1, cmd, "/killold", vbTextCompare) Then
    'cmd = Replace$(cmd, "/killold", vbNullString, , , vbTextCompare)
'End If

On Error GoTo EH
lTmr = GetTickCount()

Do
    l = Shell(AppPath() & App.EXEName & vbSpace & Trim$(cmd), vbNormalNoFocus)
Loop While (l = 0) And ((lTmr + Communicator_Close_Wait) > GetTickCount())

OpenNewCommunicator = l > 0

Exit Function
EH:
OpenNewCommunicator = False
End Function

Public Sub OpenImage(ByVal Path As String)
'Const Quot = """"
'Shell "rundll32.exe C:\WINDOWS\System32\shimgvw.dll,ImageView_Fullscreen " & Path

ShellExecute 0&, sOpen, Path, 0&, 0&, SW_SHOW

End Sub

Public Property Let StealthMode(ByVal TurnOn As Boolean)
Dim sTxt As String

'Static LastDrawingChecked As Boolean
On Error GoTo EH

pSniperMode = TurnOn
'modMessaging.mMStealth = TurnOn
modVars.bStealth = TurnOn

frmMain.ShowForm Not TurnOn, False 'problem is after, or here
DoSystray Not TurnOn

If TurnOn Then
    'Unload frmStealth
    Load frmStealth
    
    On Error Resume Next
    frmStealth.Left = frmMain.Left + frmMain.width / 2 - frmStealth.width / 2
    frmStealth.Top = frmMain.Top + frmMain.height / 2 - frmStealth.height / 2
    frmStealth.Show
    frmStealth.mnuFileSend_Click
    
    'LastDrawingChecked = frmMain.mnuOptionsMessagingDrawingOff.Checked
    'frmMain.mnuOptionsMessagingDrawingOff.Checked = False
Else
    Unload frmStealth
    'frmMain.mnuOptionsMessagingDrawingOff.Checked = LastDrawingChecked
End If

AddConsoleText "Stealth Mode " & IIf(TurnOn, "A", "Dea") & "ctivated"
modLogging.addToActivityLog "Stealth Mode " & IIf(TurnOn, "A", "Dea") & "ctivated"

'without this, no icon would appear for frmMain... buh?
frmMain.RefreshIcon

Exit Property
EH:
sTxt = "Stealth Mode Transition Error: " & Err.Description

AddConsoleText sTxt
AddText sTxt, TxtError, True
End Property

Public Property Get StealthMode() As Boolean

StealthMode = pSniperMode

End Property

'-------------------------------------------------------

Public Function FileExists(ByVal Path As String, Optional ByVal Attribs As VbFileAttribute = vbNormal) As Boolean

On Error GoTo EH
FileExists = CBool(Len(Dir$(Path, Attribs))) And CBool(Len(Path))
Exit Function
EH:
FileExists = False
End Function

Public Function FNameOnly(sPath As String) As String
FNameOnly = Mid$(sPath, 1 + InStrRev(sPath, "\", , vbTextCompare))
End Function

'Private Function FileExist(ByRef inFile As String) As Boolean
'FileExist = CBool(Len(Dir(inFile)))
'End Function
'
'Private Function FileExist(ByRef inFile As String) As Boolean
'On Error Resume Next
'FileExist = CBool(FileLen(inFile) + 1)
'End Function

Public Function IsFileOpen(ByVal FilePath As String) As Boolean

Dim Fn As Integer, ErrNum As Long

On Error Resume Next
Fn = FreeFile()

'Attempt to open the file and lock it.
Open FilePath For Input Lock Read As #Fn
Close #Fn

ErrNum = Err.Number

Select Case ErrNum
    Case 0
        'No error occurred - File is NOT already open by another user.
        IsFileOpen = False
        
    Case 70
        'Error number for "Permission Denied. - File is opened by another user.
        IsFileOpen = True
        
    Case Else
        ' Yikes some other error occurred.
        IsFileOpen = True
        
End Select

End Function

Public Property Let uPassword(ByVal P As String)

puPassword = P

End Property

'Public Sub LoadInet()
'AddConsoleText "Loading Inet Control...", , True
'Load frmInet
'modVars.CanUseInet = True
'AddConsoleText "Inet Control Loaded Successfully", , , True
'End Sub

Public Function AppPath() As String
Dim Tmp As String

Tmp = App.Path

If bIsIDE Then
    Tmp = "C:\Documents and Settings\Rob\My Documents\Code\Programs VB\winsock\Multi\"
End If

If Right$(Tmp, 1) <> "\" Then
    Tmp = Tmp & "\"
End If
AppPath = Tmp
End Function

Public Function Fill(ByVal Text As String, ByVal N As Integer) As String
Dim Tmp As String

Tmp = Space$(N)

Mid$(Tmp, 1, Len(Text)) = Text

Fill = Tmp

End Function

Public Sub AddText(ByVal Text As String, _
    Optional ByVal colour As Long = (-1), Optional ByVal Info As Boolean = False, _
    Optional ByVal DisableAutoTimeStamp As Boolean = False, _
    Optional sFont As String = vbNullString)

Dim TextToAdd As String, sLeft As String, sRight As String, Tmp As String
'Dim i As Integer, j As Integer

'Dim OldSelStart As Long, OldSelLen As Long
Dim ScrollY As Long, AddedTextHeight As Single, bAtBottom As Boolean

If modLoadProgram.frmMain_Loaded = False Or bDisableAddText Then Exit Sub
'otherwise it'll try to access rtfIn, which loads frmmain, but then exits frmMain, leaving it still unloaded

'OldSelStart = frmMain.rtfIn.Selstart
'OldSelLen = frmMain.rtfIn.Sellength


ScrollY = frmMain.rtfIn.ScrollPosY
bAtBottom = frmMain.rtfIn.ScrollIsAtBottom() '(ScrollY = frmMain.rtfIn.ScrollPosY)

'bAtBottom = (ScrollY >= frmMain.ScaleY(frmMain.TextHeight(frmMain.rtfIn.Text), vbTwips, vbPixels))

'If ScrollY <> frmMain.rtfIn.ScrollPosY Then
    'not at bottom
    'bAtBottom = False



Redraw False

'set colour + pos
frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
If colour = (-1) Then colour = TxtInfo
frmMain.rtfIn.SelColor = colour

If LenB(sFont) > 0 Then
    frmMain.rtfIn.SelFontName = sFont
ElseIf frmMain.rtfIn.SelFontName <> DefaultFontName Then
    frmMain.rtfIn.SelFontName = DefaultFontName
End If

If CInt(frmMain.rtfIn.SelFontSize) <> CInt(frmMain.rtfFontSize) Then
    frmMain.rtfIn.SelFontSize = frmMain.rtfFontSize
End If
If frmMain.rtfIn.SelItalic <> frmMain.rtfItalic Then
    frmMain.rtfIn.SelItalic = frmMain.rtfItalic
End If
If frmMain.rtfIn.SelBold <> frmMain.rtfBold Then
    frmMain.rtfIn.SelBold = frmMain.rtfBold
End If

''add www. if it's a url
'i = InStr(1, Text, ".com", vbTextCompare)
'If i Then
'    If InStr(1, Text, "www.", vbTextCompare) = 0 Then
'
'        For j = i To 1 Step -1
'            If Mid$(Text, j, 1) = vbspace Then
'                sRight = Mid$(Text, i)
'                sLeft = Left$(Text, j)
'                Tmp = Mid$(Text, j + 1, i - j - 1)
'
'                Text = sLeft & "www." & Tmp & sRight
'
'                Exit For
'
'            End If
'        Next j
'
'    End If
'End If


TextToAdd = Text

'add info bits
If Info Then
    TextToAdd = InfoStart & TextToAdd & InfoEnd
End If

'add timestamp
If frmMain.mnuOptionsTimeStamp.Checked And frmMain.mnuOptionsTimeStampInfo.Checked And (Not DisableAutoTimeStamp) Then
    TextToAdd = "[" & FormatDateTime$(Time$, vbLongTime) & "] " & TextToAdd
End If


'stick the text in
Call AddWithTags(vbNewLine & TextToAdd)


'if stealth, add it there
If StealthMode Then
    frmStealth.AddText Text, True
End If

'On Error Resume Next
AddedTextHeight = frmMain.ScaleY(frmMain.TextHeight(Text), vbTwips, vbPixels)

If ScrollY Then
    If bAtBottom Then 'ScrollY >= (frmMain.rtfIn.ScrollPosY - (AddedTextHeight * 6)) Then
        
        frmMain.rtfIn.ScrollPosY = ScrollY + _
            frmMain.ScaleY(frmMain.TextHeight(Text), vbTwips, vbPixels) * frmMain.TextWidth(Text) \ frmMain.rtfIn.width
        'take note: \ (aka DIV)
        
        frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
    Else
        frmMain.rtfIn.ScrollPosY = ScrollY
    End If
End If

Redraw True

'frmMain.rtfIn.Selstart = OldSelStart
'frmMain.rtfIn.Sellength = OldSelLen

End Sub

Private Sub AddWithTags(ByVal TextToAdd As String)
'e.g.
'TextToAdd = "this <i>is <b>a</b></i> reet <u>nice</u> test"
'TextToAdd = "this <i><b>is a</b> re</i>et <u>nice</u> test"

Dim i As Integer
Dim IndexAr(1 To 6) As Integer
Dim Orig_Bold As Boolean, Orig_Italic As Boolean, Orig_UnderLine As Boolean
Dim sTmpTxt As String


IndexAr(1) = InStr(1, TextToAdd, MakeTag(BoldTag), vbTextCompare)
IndexAr(2) = InStr(1, TextToAdd, MakeTag(ItalicTag), vbTextCompare)
IndexAr(3) = InStr(1, TextToAdd, MakeTag(UnderLineTag), vbTextCompare)


If (IndexAr(1) + IndexAr(2) + IndexAr(3)) > 0 Then
    
    IndexAr(4) = InStr(1, TextToAdd, MakeTag(BoldTag, True), vbTextCompare)
    IndexAr(5) = InStr(1, TextToAdd, MakeTag(ItalicTag, True), vbTextCompare)
    IndexAr(6) = InStr(1, TextToAdd, MakeTag(UnderLineTag, True), vbTextCompare)
    
    With frmMain.rtfIn
        Orig_Bold = .Font.Bold
        Orig_Italic = .Font.Italic
        Orig_UnderLine = .Font.Underline
        
        .EnableSmiles = False
        
        For i = 1 To Len(TextToAdd)
            
Top:
            If i = IndexAr(1) Then
                .SelBold = True
                i = i + 3
                GoTo Top
            End If
            If i = IndexAr(2) Then
                .SelItalic = True
                i = i + 3
                GoTo Top
            End If
            If i = IndexAr(3) Then
                .SelUnderLine = True
                i = i + 3
                GoTo Top
            End If
            
            
            If i = IndexAr(4) Then
                .SelBold = Orig_Bold
                i = i + 4
                GoTo Top
            End If
            If i = IndexAr(5) Then
                .SelItalic = Orig_Italic
                i = i + 4
                GoTo Top
            End If
            If i = IndexAr(6) Then
                .SelUnderLine = Orig_UnderLine
                i = i + 4
                GoTo Top
            End If
            
            
            'elseifs can't be used - a tag may be missed by i=i+x
            'goto top needed, in case the order is wrong
            
            
            .SelText = Mid$(TextToAdd, i, 1)
            
            
        Next i
        
        'make sure no changes are left
        .SelBold = Orig_Bold
        .SelItalic = Orig_Italic
        .SelUnderLine = Orig_UnderLine
        'Debug.Print "restored: ob: " & Orig_Bold & ", oi: " & Orig_Italic & ", ou: " & Orig_Italic
        
        .EnableSmiles = frmMain.mnuOptionsMessagingDisplaySmiliesComm.Checked Or _
                        frmMain.mnuOptionsMessagingDisplaySmiliesMSN.Checked
        
        .Process_Smiles
        
    End With
    
Else
    frmMain.rtfIn.SelText = TextToAdd
End If


End Sub

Private Function Smallest_Int(ParamArray Vars() As Variant) As Integer
Dim i As Integer, Smallest As Integer
Dim Cur As Integer

Cur = 32767
Smallest = -1

For i = 0 To UBound(Vars)
    If Vars(i) < Cur And Vars(i) > 0 Then
        Cur = Vars(i)
        Smallest = i
    End If
Next i


Smallest_Int = Smallest

End Function

Public Function MakeTag(sBaseTag As String, Optional bCloseTag As Boolean = False) As String
Dim i As Integer

i = InStr(1, sBaseTag, vbSpace)

If bCloseTag And i > 0 Then
    MakeTag = "</" & Left$(sBaseTag, i - 1) & ">"
Else
    MakeTag = "<" & IIf(bCloseTag, "/", vbNullString) & sBaseTag & ">"
End If

End Function

Private Sub Redraw(bTurnOn As Boolean)
Static EventMask As Long
Dim hWnd As Long

hWnd = frmMain.rtfIn.hWnd

If bTurnOn Then
    'events
    SendMessageByLong hWnd, EM_SETEVENTMASK, 0, EventMask
    
    'redrawing
    SendMessageByLong hWnd, WM_SETREDRAW, 1, 0
    
    frmMain.rtfIn.Refresh
Else
    'redrawing
    SendMessageByLong hWnd, WM_SETREDRAW, 0, 0
    
    'events
    EventMask = SendMessageByLong(hWnd, EM_GETEVENTMASK, 0, 0)
End If


End Sub

Public Sub MidText(ByVal Text As String, Optional ByVal colour As Long = (-1))

frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)

If colour = (-1) Then colour = TxtInfo

frmMain.rtfIn.SelColor = colour

If Asc(Text) <> vbKeyBack Then
    frmMain.rtfIn.SelText = Text
Else
    frmMain.rtfIn.Selstart = frmMain.rtfIn.Selstart - 1
    frmMain.rtfIn.Sellength = 1
    frmMain.rtfIn.SelText = vbNullString
End If

frmMain.rtfIn.Refresh

End Sub

Public Sub ErrorHandler(ByVal Desc As String, ByVal ErrNo As Long, _
    Optional ByVal ShowError As Boolean = True, Optional ByVal ForceBalloon As Boolean = False) ', _
    Optional ByVal RetryListen As Boolean = False, Optional ByVal RetryConnect As Boolean = False)

If ErrNo = 0 Then
    Exit Sub
ElseIf LenB(Desc) = 0 Then
    Exit Sub
End If

AddConsoleText "Main Socket Error - " & Desc '& vbNewLine & _
    Space$(modConsole.IndentLevel) & "Retry Listen:" & CStr(RetryListen) & vbNewLine & _
    Space$(modConsole.IndentLevel) & "Retry Connect:" & CStr(RetryConnect)

If ShowError Then
    If ErrNo = WSAEADDRINUSE Then
        AddText String$(5, "-") & vbNewLine & "Error: Address in use - Another Communicator is already listening/connected" & vbNewLine & String$(5, "-"), TxtError
    'ElseIf ErrNo = WSANO_DATA Then
        'AddText String$(5, "-") & vbNewLine & "Error: Address in use - Another Communicator is already listening/connected" & vbNewLine & String$(5, "-"), TxtError
    Else
        AddText "Error: " & Desc, TxtError, True
    End If
End If

'modLogging.LogEvent "Socket Error - " & Desc, LogError
Call frmMain.CleanUp(False)

'If RetryListen Then
'    Call RetryLC(True, ErrNo)
'ElseIf RetryConnect Then
'    Call RetryLC(False, ErrNo)
'End If

If ForceBalloon Or (Not modVars.IsForegroundWindow()) Then  'frmMain.Visible = False Or ForceBalloon Then
    frmSystray.ShowBalloonTip "Socket Error - " & Desc, , NIIF_ERROR, , True
End If

End Sub

'Public Sub RetryLC(ByVal bListen As Boolean, ByVal ErrNo As Long)
'Dim Ans As VbMsgBoxResult
'Dim b As Boolean
'
'Const AddrInUse As Integer = 10048
'Const AddrInUse2 As Integer = 362
'
'If (ErrNo = AddrInUse Or ErrNo = AddrInUse2) Then
'    Ans = frmMain.Question("Retry with a different port? (Current Port: " & RPort & ")", _
'                                IIf(bListen, frmMain.cmdListen, frmMain.cmdAdd))
'
'    If Ans = vbYes Then
'        'B = (LPort > RPort)
'
'        'Do
'            'If B Then
''        Do
''
''            If LPort >= 2860 Then
''                LPort = 2850
''            Else
''                LPort = LPort + 1
''            End If
''
''        Loop While LPort = RPort
'
'        RPort = RPort + 1
'
'            'Else
'                'LPort = LPort + 1
'            'End If 'keep lport nearish rport
'
'        'Loop While LPort = RPort
'
'        If bListen Then
'            Call frmMain.Listen
'        Else
'            Call frmMain.Connect(LastIP)
'        End If
'
'        'lport = rport+1
'
'    End If
'
'End If
'
'End Sub

Public Sub FormLoad(ByVal Frm As Form, Optional ByVal bReverse As Boolean = False, _
    Optional ByVal bImplode As Boolean = True, Optional ByVal bSetLR As Boolean = True, _
    Optional ByVal bConnectedIcon As Boolean = False)

'Dim pFrm As Form
'Dim bIsForm As Boolean
'
'
'If Frm.hWnd <> frmMain.hWnd Then
'    If InTray Then
'        If Frm.hWnd <> frmSystray.hWnd Then
'            bIsForm = True
'        End If
'    Else
'        bIsForm = (Frm.Name <> "frmSystray")
'    End If
'End If
'
'If bIsForm Then
'    If Reverse = False Then
'        'loading a form
'        AFormLoaded = False
'    Else
'        AFormLoaded = False
'
'        For Each pFrm In Forms
'            If Frm.hWnd <> frmMain.hWnd Then
'                If Frm.Name <> "frmSystray" Then
'                    AFormLoaded = True
'                    Exit For
'                End If
'            End If
'        Next pFrm
'
'    End If
'End If


If Not (Frm Is Nothing) Then
    On Error Resume Next
    
    If Not bReverse Then
        If bConnectedIcon Then
            Frm.Icon = frmMain.ConnectedIcon
        Else
            Frm.Icon = frmMain.IdleIcon 'frmMain.imglstIcons.ListImages(eStatus.Idle + 1).Picture
        End If
        
        'Frm.Top = frmMain.Top + frmMain.Height / 4
        'Frm.Left = frmMain.Left + frmMain.Width / 4
        'Frm.Top = frmMain.Top + (Abs(frmMain.Height - Frm.Height) / 2)
        'Frm.Left = frmMain.Left + (Abs(frmMain.Width - Frm.Width) / 2)
        If bSetLR Then
            Frm.Left = frmMain.Left + frmMain.width / 2 - Frm.width / 2
            Frm.Top = frmMain.Top + frmMain.height / 2 - Frm.height / 2
            
            If (Frm.Left + Frm.width) > Screen.width Then
                Frm.Left = Screen.width - Frm.width
            End If
            If Frm.Left < 10 Then Frm.Left = 10
            
        End If
    'Else
        'Frm.StartUpPosition = vbStartUpManual
        'Frm.Visible = False
        'bI = True
    End If
    
    If bImplode Then
        'ImplodeFormToMouse Frm.hWnd, Not Reverse ', True
        modImplode.AnimateAWindow Frm.hWnd, aRandom, bReverse
    End If
End If

End Sub

Public Sub LostFocus(ByRef TB As Control)

TB.Text = Trim$(TB.Text)

End Sub

Public Function GetStatus(Optional ByVal tStatus As eStatus = -1, _
                        Optional ByVal Serv As Boolean = False) As String
Dim Tmp As String

Select Case IIf(tStatus = -1, Status, tStatus)
    Case eStatus.Idle
        Tmp = "Idle"
    Case eStatus.Connecting
        Tmp = "Connecting"
    Case eStatus.Connected
        If tStatus = -1 Then
            Serv = Server
        End If
        
        If Serv Then
            Tmp = "Connected - Host"
        Else
            Tmp = "Connected - Client"
        End If
        
    Case eStatus.Listening
        Tmp = "Listening"
End Select

GetStatus = Tmp

End Function

Public Sub Cmds(ByVal sStatus As eStatus)
Dim strStatus As String

'If Status = sStatus Then Exit Sub

bCleanedUp = False

Status = sStatus
With frmMain
    
    .ResetTxtOutHeight
    
    .lstConnected.Clear
    
    If Status <> Connected Then
        .cmdSmile.Enabled = False
        '.cmdSlash.Enabled = False
        '.mnuFileInvite.Enabled = False
        .cmdShake.Enabled = False
        '.mnuFileClient.Enabled = False
        .mnuFileSaveDraw.Enabled = False
        .mnuFileRefresh.Enabled = True
        
        .picDraw.Enabled = False
        .EnableCmdCls False
        '.cmdDevSend.Enabled = False
        '.txtDev.Enabled = False
        .mnuOptionsMatrix.Enabled = False
        
        .EnableCmd 3, False
        '.txtOut.Enabled = False
        .cmdSend.Enabled = False
        .EnableCmd 5, False 'private
        
        If modVars.bStealth Then frmStealth.mnuGame.Enabled = True
        
        .lstConnected.Enabled = False
        .mnuOptionsMessagingDrawingOff.Enabled = False
        .mnuOptionsMessagingDrawingOff.Checked = False
        
        '##################################################
        '.mnuOptionsMessagingLobby.Enabled = False
        '.mnuOptionsMessagingPrivate.Enabled = False
        '.mnuOptionsMessagingWindowsFT.Enabled = False
        .mnuOptionsMessagingWindows.Enabled = False
        '##################################################
        
        .mnuOptionsDPReset.Enabled = False
        .mnuOptionsDPSet.Enabled = False
        '.mnuOptionsDPRefresh.Enabled = False
        .mnuOptionsDPClear.Enabled = False
        
        .mnuDevFormsCmds.Enabled = False
        
        '.tmrAnim.Enabled = False
        .OLEDropMode = 0
        .rtfIn.OLEDropMode = vbOLEDropNone
        
        .mnuOptionsMessagingServerMsg.Enabled = False
        
        .Set_lstConnected_Left
        
        .lstConnected.width = .lstComputers.width
        
        .mnuHelpBugReport.Enabled = False
        
        .DP_OLEDragDrop False
        
        .mnuOptionsMessagingWindowsRecord.Enabled = False
        '.mnuOptionsMessagingWindowsRecordings.Enabled = False
        
        If modLoadProgram.frmVoiceTransfers_Loaded Then
            Unload frmVoiceTransfers
        End If
    End If
    
    If sStatus <> Connected Then
        modMessaging.bReceivedWelcomeMessage = False
        'unset it here, in case, while connected, a `cmds connected` is run
    End If
    
    Select Case sStatus
        Case eStatus.Idle
            .EnableCmd 0, True
            .EnableCmd 1, False
            .lstComputers.Enabled = True
            .EnableCmd 2
            '.mnuFileIPs.Enabled = True

            
        Case eStatus.Connected
            .EnableCmd 0, False
            .EnableCmd 1
            '.txtOut.Enabled = True
            '.cmdSend.Enabled = True
            .lstComputers.Enabled = False
            '.cmdRemove.Enabled = True 'IIf(Server, True, False)
            .EnableCmd 2, False ' was true, but use invite instead 'IIf(Server, True, False)
            .picDraw.Enabled = True
            .EnableCmdCls True
            '.cmdDevSend.Enabled = True 'so that a nofilter or msg that requires no args can be done
            '.txtDev.Enabled = True
            .mnuOptionsMatrix.Enabled = True
            '.mnuFileInvite.Enabled = True
            .cmdShake.Enabled = .mnuOptionsMessagingShake.Checked
            '.mnuFileClient.Enabled = True 'Server
            .mnuFileSaveDraw.Enabled = True
            .cmdSmile.Enabled = frmMain.mnuOptionsMessagingDisplaySmiliesComm.Checked Or frmMain.mnuOptionsMessagingDisplaySmiliesMSN.Checked
            .mnuFileRefresh.Enabled = False
            
            If modVars.bStealth Then frmStealth.mnuGame.Enabled = True
            
            '.mnuOptionsMessagingStick.Enabled = True
            '.mnuFileIPs.Enabled = False
            '.mnuOptionsMessagingPrivate.Enabled = True
            '.cmdPrivate.Enabled = True
            .lstConnected.Enabled = True
            '.cmdSlash.Enabled = True
            
            .mnuOptionsMessagingDrawingOff.Enabled = Not Server
            
            
            .mnuOptionsDPReset.Enabled = True
            .mnuOptionsDPSet.Enabled = True
            '.mnuOptionsDPRefresh.Enabled = True
            '.mnuOptionsDPClear.Enabled = False
            
            
            .lstConnected.Clear
            .lstConnected.AddItem ConnectedListPlaceHolder
            
            '.tmrAnim.Enabled = True
            .mnuDevFormsCmds.Enabled = True
            .OLEDropMode = vbOLEDropManual
            .rtfIn.OLEDropMode = vbOLEDropManual
            .mnuOptionsMessagingServerMsg.Enabled = Server
            .lstConnected.Left = .lstComputers.Left
            .lstConnected.width = lstConnectedExtendedWidth
            
            
            '.mnuOptionsMessagingWindowsFT.Enabled = True
            '.mnuOptionsMessagingLobby.Enabled = True
            'don't enable private chat here
            .mnuOptionsMessagingPrivate.Enabled = False
            .mnuOptionsMessagingWindows.Enabled = True
            
            .mnuHelpBugReport.Enabled = True
            
            .DP_OLEDragDrop True
            
            .mnuOptionsMessagingWindowsRecord.Enabled = True
            '.mnuOptionsMessagingWindowsRecordings.Enabled = True
            
            If Not modLoadProgram.frmVoiceTransfers_Loaded Then
                Load frmVoiceTransfers
            End If
            If Len(frmMain.txtOut.Text) > 0 Then
                frmMain.cmdSend.Enabled = True
            End If
            
        Case eStatus.Connecting
            
            .EnableCmd 1, True
            .lstComputers.Enabled = False
            .EnableCmd 3, False
            .EnableCmd 0, False
            .EnableCmd 2, False
            '.mnuFileIPs.Enabled = False
            
        Case eStatus.Listening
            
            .EnableCmd 0, False
            .EnableCmd 1, True
            .lstComputers.Enabled = False
            .EnableCmd 2, True 'allow manual connect, even when listening
            '.mnuFileIPs.Enabled = True
            .mnuOptionsMessagingServerMsg.Enabled = Server
            
    End Select
    
    frmSystray.mnuPopupCloseC.Enabled = .cmdArray(1).Enabled
    frmSystray.mnuPopupHost.Enabled = .cmdArray(0).Enabled
    
    Call GetTrayText '+set
    
    strStatus = GetStatus()
    AddConsoleText "Setting Status: " & strStatus
    modLogging.addToActivityLog "Setting Status: " & strStatus
    If modLoadProgram.frmMini_Loaded Then
        frmMini.SetInfo strStatus
    End If
    
    If InTray Then
        frmMain.SetIcon sStatus
    End If
    
    
    If Not modVars.bRetryConnection Then
        If modVars.IsForegroundWindow() Then
            SetFocus2 .rtfIn
        End If
    End If
End With

End Sub

Public Function GetTrayText(Optional ByVal SetText As Boolean = True) As String

Const CaptExt As String = " - "
Const kSystrayText As String = "Communicator"
Dim MainText As String
Dim sStatus As String, rIP As String

'Dim i As Integer

'On Error Resume Next

sStatus = GetStatus()
frmMain.Caption = Capt & CaptExt & sStatus
rIP = modWinsock.RemoteIP


MainText = kSystrayText & CaptExt & GetVersion() & vbNewLine & _
        "Status: " & sStatus & vbNewLine & _
        "Local IP: " & modWinsock.LocalIP & vbNewLine & _
        IIf(LenB(rIP), "Remote IP: " & rIP, "[Remote IP Not Obtained]") & vbNewLine _
        & GetMode()


'i = Len(TopLine) - Len(BottomLine)
'
'If i < 0 Then
'    i = 0
'    TopLine = Space$(2) & TopLine
'End If
'
'BottomLine = Space$(i) & BottomLine & Space$(i)


If SetText Then
    frmSystray.ToolTip = MainText
End If

GetTrayText = MainText
End Function

Private Function GetMode() As String

If frmMain.mnuFileGameMode.Checked Then
    GetMode = "Game Mode"
    
ElseIf bDevMode Then
    
    'If bUberDevMode Then
        'If bHeightenedDev Then
            'GetMode = "Heightened Mode"
        'Else
            'GetMode = "Uber-Dev Mode"
        'End If
    'Else
        'GetMode = "Dev Mode"
    'End If
    
    GetMode = modDev.getDevLevelName()
    
Else
    GetMode = "Normal Mode"
End If


End Function

Public Function GetVersion() As String
GetVersion = CStr(App.Major & Dot & App.Minor & Dot & App.Revision)
End Function

Public Sub Pause(ByVal HowLong As Long)
Dim Tick As Long
Tick = GetTickCount()
Do
  DoEvents
Loop Until Tick + HowLong < GetTickCount()
End Sub

Public Sub SetOnTop(ByVal hWnd As Long, Optional ByVal bOnTop As Boolean = True, _
    Optional ByVal bShow As Boolean = True)

'http://www.xtremevbtalk.com/showthread.php?t=28299
Dim f As Long

f = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

If bShow Then
    f = f Or SWP_SHOWWINDOW
End If

If bOnTop Then
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, f
Else
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, f
End If

End Sub

Public Sub DoSystray(ByVal TurnOn As Boolean)
Static Added As Boolean

If TurnOn Then
    Load frmSystray
    
    If Not Added Then
        Added = True
        SetSplashInfo "Adding Main Form Script Object..."
        frmMain.SC.AddObject "frmSystray", frmSystray, True
    End If
    
    Call GetTrayText
Else
    'frmmain.SC.remove (?)
    Unload frmSystray
End If
End Sub

Public Function GetState(ByVal i As Integer) As String
Dim T As String

Select Case i
    Case 0
        T = "SckClosed"
    Case 1
        T = "SckOpen"
    Case 2
        T = "SckListening"
    Case 6
        T = "SckConnecting"
    Case 7
        T = "SckConnected"
    Case 9
        T = "SckError"
    Case 8
        T = "SckClosing"
    Case 3
        T = "SckConnectionPending"
    Case 4
        T = "SckResolvingHost"
    Case 5
        T = "SckHostResolved"
End Select

GetState = T

End Function

Public Sub SetDefaultColours()
TxtError = vbRed
TxtInfo = MGreen
TxtReceived = MOrange
TxtSent = vbBlue
TxtUnknown = vbBlack
TxtBackGround = vbWhite
TxtQuestion = MBrown
TxtForeGround = MPurple
End Sub

Public Property Let TxtForeGround(ByVal Col As Long)
pTxtForeGround = Col
'If frmMain.mnuOptionsMessagingColours.Checked Then
    frmMain.txtOut.ForeColor = pTxtForeGround
'Else
    'frmMain.txtOut.ForeColor = TxtSent
'End If
End Property

Public Property Let TxtBackGround(ByVal Col As Long)
pTxtBackGround = Col
frmMain.rtfIn.BackColor = pTxtBackGround
End Property

Public Property Get TxtForeGround() As Long

TxtForeGround = pTxtForeGround

End Property

Public Property Get TxtBackGround() As Long

TxtBackGround = pTxtBackGround

End Property

'Public Function GetApphWnd(ByVal Text As String) As Long
'
'theTitle = Text
'
'EnumWindows AddressOf WindowEnumerator, 0
'
'theTitle = vbNullString
'
'GetApphWnd = theHwnd
'theHwnd = 0
'
'End Function
'
'' Return False to stop the enumeration.
'Private Function WindowEnumerator(ByVal app_hwnd As Long, ByVal lparam As Long) As Long
'Dim buf As String * 256
'Dim title As String
'Dim length As Long
'
'    ' Get the window's title.
'    length = GetWindowText(app_hwnd, buf, Len(buf))
'    title = Left$(buf, length)
'
'    ' See if the title contains the target.
'    If InStr(title, theTitle) <> 0 Then
'        ' Save the hwnd and end the enumeration.
'        theHwnd = app_hwnd
'        WindowEnumerator = False
'    Else
'        ' Continue the enumeration.
'        WindowEnumerator = True
'    End If
'End Function

Public Function Password(ByVal Prompt As String, Owner As Form, _
    Optional ByVal sTitle As String = vbNullString, Optional ByVal sDefault As String = vbNullString, _
    Optional ByVal UsePassChar As Boolean = True, Optional ByVal iMaxLen As Integer = 0, _
    Optional ByVal bNumeric As Boolean = False) As String

Load frmPassword

With frmPassword
    .Prompt = Prompt
    
    'SetWindowPos .hWnd, 0, .ScaleX(frmMain.width \ 2 - .width \ 2, vbTwips, vbPixels), _
                           .ScaleY(frmMain.height \ 2 - .height \ 2, vbTwips, vbPixels), _
                           .ScaleX(.width, vbTwips, vbPixels), _
                           .ScaleY(.height, vbTwips, vbPixels), _
                           0
    
    .Icon = frmMain.Icon
    
    If LenB(sTitle) Then .Caption = sTitle
    If UsePassChar Then
        .PassChar = "*"
    Else
        .PassChar = vbNullString
    End If
    If iMaxLen Then .MaxLen = iMaxLen
    .Default = sDefault
    .Numeric = bNumeric
    
    
    'ShowWindow .hWnd, SW_SHOWNORMAL
    'ShowWindow .hWnd, SW_HIDE
    'set pos
    '.Show
    '.Hide
    'Owner.Refresh
    
    Call FormLoad(frmPassword, , False, False)
    
    
    'ShowWindow .hWnd, SW_HIDE
    
    If LenB(sDefault) Then
        .txtPassword.Selstart = 0
        .txtPassword.Sellength = Len(.txtPassword.Text)
    End If
    
    .Left = Owner.Left + Owner.width / 2 - .width / 2
    .Top = Owner.Top + Owner.height / 2 - .height / 2
    
    .Show vbModal, Owner
    
End With

Password = puPassword
puPassword = vbNullString

End Function

Public Function BrowseForFolder(sInitDir As String, Title As String, Owner As Form) As String

Load frmFolderBrowse
With frmFolderBrowse
    .Left = frmMain.width \ 2 - .width \ 2
    .Top = frmMain.height \ 2 - .height \ 2
    .Caption = Title
    
    .InitDir = sInitDir
    
    FormLoad frmFolderBrowse
    
    .Show vbModal, Owner
    ' "freezes" here
    
End With

BrowseForFolder = pBrowse_FolderPath
pBrowse_FolderPath = vbNullString

End Function

Public Function IPChoice(Owner As Form, Optional bChooseSocket As Boolean = False, _
    Optional sCaption As String = "Select an IP...") As String


Load frmIPChooser

With frmIPChooser
    
    .Move Owner.Left + Owner.width \ 2 - .width \ 2, _
          Owner.Top + Owner.height \ 2 - .height \ 2
    
    FormLoad frmIPChooser, , , False, True
    
    .bChooseSocket = bChooseSocket
    
    .lblInfo.Caption = sCaption
    
    .Show vbModal, Owner
    ' "freezes" here
    
End With

IPChoice = pIP_Choice
pIP_Choice = vbNullString

End Function

Public Property Get Server() As Boolean
Server = pServer
End Property

Public Property Let Server(ByVal S As Boolean)

pServer = S
frmMain.mnuDevCmdsServer.Checked = S
If S Then
    modMessaging.MySocket = -1
Else
    modMessaging.MySocket = 0
End If

'AddConsoleText "My Socket: " & CStr(modMessaging.MySocket)

End Property

Public Function IsForegroundWindow() As Boolean 'ByVal hWnd As Long) As Boolean
Dim hWndForeground As Long
Dim Frm As Form

hWndForeground = apiGetForegroundWindow()

For Each Frm In Forms
    If Frm.hWnd = hWndForeground Then
        IsForegroundWindow = True
        Exit For
    End If
Next Frm

End Function

Public Function IshWndForegroundWindow(ByVal hWnd As Long) As Boolean

IshWndForegroundWindow = (apiGetForegroundWindow() = hWnd)

End Function

Public Function GetForegroundWindow() As Long
GetForegroundWindow = apiGetForegroundWindow()
End Function

Public Function OpenURL(ByVal sSite As String) As Long
OpenURL = ShellExecute(0&, sOpen, sSite, vbNullString, vbNullString, SW_SHOW)
End Function

'Public Function DllErrorDescription(Optional ByVal lLastDLLError As Long) As String
'Dim sBuff As String * 256
'Dim lCount As Long
''Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
''Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
''Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
''Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
'
'If lLastDLLError = 0 Then
'    'Use Err object to get dll error number
'    lLastDLLError = Err.LastDllError
'End If
'
'lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
'    0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
'
'If lCount Then
'    ErrorDescriptionDLL = Left$(sBuff, lCount - 2)    'Remove line feeds
'End If
'
'End Function
Public Function DllErrorDescription(Optional ByVal LastErr As Long = -1) As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Const LANG_NEUTRAL = &H0
Dim sBuff As String
Dim lRet As Long

If LastErr = -1 Then LastErr = Err.LastDllError
sBuff = String$(256, 0)

lRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
    ByVal 0&, LastErr, LANG_NEUTRAL, sBuff, Len(sBuff), ByVal 0&)


DllErrorDescription = TrimNewLine(Left$(sBuff, lRet))
End Function

Public Function TrimNewLine(ByVal sTxt As String) As String

Do While Left$(sTxt, 2) = vbNewLine
    sTxt = Mid$(sTxt, 3)
Loop

Do While Right$(sTxt, 2) = vbNewLine
    sTxt = Left$(sTxt, Len(sTxt) - 2)
Loop

TrimNewLine = sTxt

End Function

'Public Function TranslateWindowsVer( _
'    ByVal iMaj As Long, _
'    ByVal iMin As Long, _
'    ByVal bNt As Boolean, _
'    ByVal BVista As Boolean) As String
'
'Dim Tmp As String
'
'Tmp = "Window's Version: " & iMaj & Dot & iMin & "," '& Dot & iRev & ","
'
'If bNt Then
'    Tmp = Tmp & " Windows NT,"
'End If
'
'If BVista Then
'    Tmp = Tmp & " Windows Vista,"
'End If
'
'TranslateWindowsVer = Left$(Tmp, Len(Tmp) - 1)
'
'End Function

'Public Function SimplePing(ByVal sIP As String) As Long
'
'' Based on an article on 'Codeguru' by 'Bill Nolde'
'' Thx to this guy! It 's simple and great!
'
'' Implemented for VB in November 2002 by G. Wirth, Ulm,  Germany
'' Enjoy!
'
'Dim lIPadr As Long, lHopsCount As Long, lRTT As Long, lMaxHops As Long, lResult As Long
''                                         ^ Round Trip Time aka Ping
'Const Success = 1
'
'lMaxHops = 20 'number of routers, etc the packet can pass through
''like time to live, but jumps, rather than time
'
'lIPadr = inet_addr(sIP)
'
'If lIPadr <> -1 Then
'    lResult = GetRTTAndHopCount(lIPadr, lHopsCount, lMaxHops, lRTT)
'
'    If lResult = Success Then
'        SimplePing = lRTT
'    Else
'        SimplePing = -1
'    End If
'Else
'    SimplePing = -1
'End If
'
'End Function

'Public Sub AddToRightClick(ByVal Ext As String, ByVal MenuTitle As String, _
'    ByVal PathToShell As String)
'
'Dim ShellPath As String
'
'If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, Ext) = False Then
'    modRegistry.regCreate_A_Key HKEY_CLASSES_ROOT, Ext
'End If
'
'ShellPath = Ext & "\shell"
'If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, ShellPath) = False Then
'    modRegistry.regCreate_A_Key HKEY_CLASSES_ROOT, ShellPath
'End If
'
'ShellPath = Ext & "\shell\" & MenuTitle
'If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, ShellPath) = False Then
'    modRegistry.regCreate_A_Key HKEY_CLASSES_ROOT, ShellPath
'End If
'
'ShellPath = Ext & "\shell\" & MenuTitle & "\command"
'modRegistry.regCreate_A_Key HKEY_CLASSES_ROOT, ShellPath
'
'modRegistry.regCreate_Key_Value HKEY_CLASSES_ROOT, _
'    ShellPath, vbNullString, PathToShell
'
'End Sub
'
'Public Sub RemoveFromRightClick(ByVal Ext As String, ByVal MenuTitle As String)
'
'Dim ShellPath As String
'
'If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, Ext) Then
'
'    ShellPath = Ext & "\shell"
'    If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, ShellPath) Then
'
'        If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, ShellPath & "\" & MenuTitle) Then
'
'            modRegistry.regDelete_A_Key HKEY_CLASSES_ROOT, ShellPath & "\" & MenuTitle, "command"
'            modRegistry.regDelete_A_Key HKEY_CLASSES_ROOT, ShellPath, MenuTitle
'
'        End If
'    End If
'End If
'
'End Sub

Public Function InRightClickMenu(ByVal Ext As String, ByVal MenuTitle As String) As Boolean

Dim ShellPath As String

If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, Ext) Then
    
    ShellPath = Ext & "\shell"
    If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, ShellPath) Then
        
        ShellPath = ShellPath & "\" & MenuTitle
        
        If modRegistry.regDoes_Key_Exist(HKEY_CLASSES_ROOT, ShellPath) Then
            InRightClickMenu = True
        End If
    End If
End If


End Function

Public Function IsIP(ByVal IP As String) As Boolean
Dim IPs() As String
Dim i As Integer

IP = LCase$(IP)

If IP = vbNullString Then GoTo NotIP

If IP = "localhost" Then
    IsIP = True
    Exit Function
End If

IPs = Split(IP, ".", , vbTextCompare)

If UBound(IPs) <> 3 Then GoTo NotIP
If LBound(IPs) Then GoTo NotIP

For i = 0 To 3
    If Len(IPs(i)) > 3 Then GoTo NotIP
    If Len(IPs(i)) < 1 Then GoTo NotIP
    If Not IsNumeric(IPs(i)) Then GoTo NotIP
    If IPs(i) > 255 Then GoTo NotIP
    If IPs(i) < 0 Then GoTo NotIP
Next i

IsIP = True

ENDOFFUNCTION:
Erase IPs
Exit Function

NotIP:
IsIP = False
GoTo ENDOFFUNCTION
End Function

Public Function ClearFolder(sRoot As String, Optional bRemoveFolder As Boolean = True) As Boolean
Dim dr As String

On Error GoTo EH2

dr = Dir$(sRoot & "*.*")

Do While LenB(dr)
    If dr <> "." And dr <> ".." Then
        On Error GoTo EH
        Kill sRoot & dr
    End If
    
    dr = Dir$()
    
Loop

If bRemoveFolder Then
    On Error GoTo EH
    RmDir sRoot
End If

ClearFolder = True

Exit Function
EH:
ClearFolder = False
Exit Function
EH2:
End Function

Public Function GetDate() As String
GetDate = Format(Date, "yyyy.mm.dd")
End Function

Public Sub DrawBorder(ByRef Frm As Form)
Dim w As Single, H As Single

w = Frm.ScaleWidth - 10
H = Frm.ScaleHeight - 10

Frm.Line (0, 0)-(w, 0), vbBlack
Frm.Line -(w, H), vbBlack
Frm.Line -(0, H), vbBlack
Frm.Line -(0, 0), vbBlack
End Sub

Public Function IntRand(ByVal Low As Integer, ByVal High As Integer) As Integer
IntRand = Int((High - Low + 0.5) * Rnd()) + Low
End Function

Public Function Trim0(sTxt As String) As String
Dim i As Integer

i = InStr(1, sTxt, vbNullChar)

If i Then
    Trim0 = Left$(sTxt, i - 1)
Else
    Trim0 = sTxt
End If

End Function

Public Function FileArrayDimensioned(ar() As ptFTPFile) As Boolean

On Error GoTo EH
FileArrayDimensioned = ((UBound(ar()) = 0) Or True)

EH:
End Function

Public Function SetFocus2(Obj As Object) As Boolean
Err.Clear
On Error Resume Next
Obj.SetFocus
SetFocus2 = (Err.Number = 0)
End Function
Public Function SetFocus2txtOut() As Boolean
SetFocus2txtOut = SetFocus2(frmMain.txtOut)
End Function

Public Function ObjFromPtr(ByVal lPtr As Long) As Object
Dim Obj As Object

CopyMemory Obj, lPtr, 4&
Set ObjFromPtr = Obj
CopyMemory Obj, 0&, 4&

End Function

Public Function RemoveChars(Before As String) As String
Dim NameOp As String

NameOp = Before

If InStr(NameOp, "@") Then NameOp = Trim$(Replace$(NameOp, "@", vbNullString, , , vbTextCompare))
If InStr(NameOp, "#") Then NameOp = Trim$(Replace$(NameOp, "#", vbNullString, , , vbTextCompare))
If InStr(NameOp, modMessaging.MsgEncryptionFlag) Then NameOp = Trim$(Replace$(NameOp, modMessaging.MsgEncryptionFlag, vbNullString, , , vbTextCompare))
If InStr(NameOp, ":") Then NameOp = Trim$(Replace$(NameOp, ":", vbNullString, , , vbTextCompare)) 'for chat in game
If InStr(NameOp, modSpaceGame.mPacketSep) Then NameOp = Trim$(Replace$(NameOp, modSpaceGame.mPacketSep, vbNullString, , , vbTextCompare))
If InStr(NameOp, modSpaceGame.UpdatePacketSep) Then NameOp = Trim$(Replace$(NameOp, modSpaceGame.UpdatePacketSep, vbNullString, , , vbTextCompare))

RemoveChars = NameOp

End Function

Public Function IntoPixelsX(w As Single) As Long
IntoPixelsX = w / Screen.TwipsPerPixelX
End Function
Public Function IntoPixelsY(H As Single) As Long
IntoPixelsY = H / Screen.TwipsPerPixelY
End Function

Public Function LoadResText(ID As Integer) As String
LoadResText = StrConv(LoadResData(ID, "TEXT"), vbUnicode)
End Function

Public Function GetFileExtension(sFileName As String) As String

GetFileExtension = Right$(sFileName, Len(sFileName) - InStrRev(sFileName, Dot))

End Function

Public Function RemoveFileExt(sFileName As String) As String
Dim i As Integer

sFileName = Trim$(sFileName)

i = InStrRev(sFileName, Dot)

If i Then
    RemoveFileExt = Left$(sFileName, i - 1)
Else
    RemoveFileExt = sFileName
End If

End Function

Public Function GetFileSize_Bytes(sFileName As String) As Long

On Error Resume Next
GetFileSize_Bytes = FileLen(sFileName)

End Function

Public Function KillOldVersion() As Boolean
On Error Resume Next
Pause 250 'in case just restarted

On Error Resume Next
Kill modFTP.FTP_Comm_Exe_File 'kill .zip in d/l location

Kill AppPath() & "Communicator Old.exe"

If Err.Number Then
    AddConsoleText "Error Killing Old Communicator -  " & Err.Description
    KillOldVersion = False
Else
    AddConsoleText "Old Communicator Killed"
    KillOldVersion = True
End If

End Function

'Public Sub DoPing(ByVal Address As String, Optional ByVal addConsole As Boolean = False)
'Dim Png As Long
'
'If Status = Connected Or addConsole Then
'
'    If mnuOptionsAdvPing.Checked Then
'        Address = modWinsock.GetIPFromHostName(Address)
'
'        DoEvents
'
'        Png = modWinsock.Ping(Address) 'address needs resolving
'
'        If addConsole Then
'            AddConsoleText "Pinged " & Address & " Time: " & CStr(Png)
'        End If
'
'        If addConsole = False Then
'            sbMain.Panels(2).Text = "Ping: " & Png & "ms"
'        End If
'    End If
'
'End If
'
'End Sub
