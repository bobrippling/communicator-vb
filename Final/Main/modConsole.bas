Attribute VB_Name = "modConsole"
'http://visualbasic.about.com/od/learnvb6/l/bldykvb6dosa.htm

Option Explicit

Private pConsoleText As String
Private ConsolehWnd As Long
Private pIndentLevel As Integer
Private Const IndentNo = 3

Public frmMainhWnd As Long

'Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, _
    ByVal lpRect As Long, ByVal bErase As Long) As Long

'Public Enum eConsoleColours
'    Normal = FOREGROUND_RED + FOREGROUND_GREEN + FOREGROUND_BLUE
'    Red = FOREGROUND_RED
'    Blue = FOREGROUND_BLUE
'    Green = FOREGROUND_GREEN
'End Enum
'-----------

Private Declare Function AllocConsole _
    Lib "kernel32" () _
    As Long
Private Declare Function FreeConsole _
    Lib "kernel32" () _
    As Long
    
Private Declare Function GetStdHandle _
    Lib "kernel32" ( _
    ByVal nStdHandle As Long) _
    As Long
    
Private Declare Function ReadConsole _
    Lib "kernel32" Alias "ReadConsoleA" ( _
    ByVal hConsoleInput As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfCharsToRead As Long, _
    lpNumberOfCharsRead As Long, _
    lpReserved As Any) _
    As Long

Private Declare Function SetConsoleMode _
    Lib "kernel32" ( _
    ByVal hConsoleOutput As Long, _
    dwMode As Long) _
    As Long
    
Private Declare Function SetConsoleTextAttribute _
    Lib "kernel32" ( _
    ByVal hConsoleOutput As Long, _
    ByVal wAttributes As Long) _
    As Long
    
Private Declare Function SetConsoleTitle _
    Lib "kernel32" Alias "SetConsoleTitleA" ( _
    ByVal lpConsoleTitle As String) _
    As Long
    
Private Declare Function WriteConsole _
    Lib "kernel32" Alias "WriteConsoleA" ( _
    ByVal hConsoleOutput As Long, _
    ByVal lpBuffer As Any, _
    ByVal nNumberOfCharsToWrite As Long, _
    lpNumberOfCharsWritten As Long, _
    lpReserved As Any) _
    As Long

'#########################
Private Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" ( _
    ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, dwWriteCoord As Long, _
    lpNumberOfCharsWritten As Long) As Long
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, _
    ByVal dwCursorPosition As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, _
    lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long

Private Type COORD
    X As Integer
    Y As Integer
End Type
Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type
Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Private Declare Function FlushConsoleInputBuffer Lib "kernel32" (ByVal hConsoleInput As Long) As Long
'Private Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
'Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal dwMode As Long) As Long

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&

'SetConsoleTextAttribute color values
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_INTENSITY = &H80
'SetConsoleMode (input)
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
'SetConsoleMode (output)
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

'Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

' Global Variables
Private hConsoleIn As Long ' console input handle
Private hConsoleOut As Long ' console output handle
Private hConsoleErr As Long ' console error handle


'Forecolor constants collected into an enumarated type for a simplier use
'Public Enum ConsoleForeGroundAttributes
'   fBlack = &H0
'   fDBlue = &H1
'   fDGreen = &H2
'   fDCyan = &H3
'   fDRed = &H4
'   fDMagenta = &H5
'   fDYellow = &H6
'   fGrey = &H7
'   fDGrey = &H8
'   fBlue = &H1 Or &H8
'   fGreen = &H2 Or &H8
'   fCyan = &H3 Or &H8
'   fRed = &H4 Or &H8
'   fMagenta = &H5 Or &H8
'   fYellow = &H6 Or &H8
'   fWhite = &H7 Or &H8
'End Enum
''Backcolor constants
'Public Enum ConsoleBackGroundAttributes
'   bBlack = &H0
'   bDBlue = &H10
'   bDGreen = &H20
'   bDCyan = &H30
'   bDRed = &H40
'   bDMagenta = &H50
'   bDYellow = &H60
'   bGrey = &H70
'   bDGrey = &H80
'   bBlue = &H10 Or &H80
'   bGreen = &H20 Or &H80
'   bCyan = &H30 Or &H80
'   bRed = &H40 Or &H80
'   bMagenta = &H50 Or &H80
'   bYellow = &H60 Or &H80
'   bWhite = &H70 Or &H80
'End Enum

Public Enum eConsoleRunBatErrors
    Success = -1
    ConsoleNotOpen = -2
    FileNotFound = -3
End Enum

Public Sub RunBat(ByVal Path As String, ByRef Er As eConsoleRunBatErrors, ByRef pID As Long)
Dim H As Double

If ConsoleShown = False Then
    Er = ConsoleNotOpen
Else
    'If ConsoleShown Then
    On Error GoTo EH
    H = Shell(Path)
    'should print to console
    
    pID = H
    Er = Success
End If

Exit Sub
EH:
If Err.Number = 53 Then
    Er = FileNotFound
End If
End Sub

Public Property Get ConsoleText() As String
ConsoleText = pConsoleText
End Property

Public Property Get IndentLevel() As Integer
IndentLevel = pIndentLevel * IndentNo
End Property

Public Sub Indent(Optional ByVal Increase As Boolean = True)

If Increase Then
    pIndentLevel = pIndentLevel + 1
Else
    pIndentLevel = pIndentLevel - 1
End If

If pIndentLevel < 0 Then pIndentLevel = 0

End Sub

Public Sub ShowConsole(Optional ByVal ShowIt As Boolean = True, Optional ByVal AtStartup As Boolean = False)

Dim UserInput As String, Command As String, Param As String
Dim hSysMenu As Long, menuCount As Long
'Dim Nt As Boolean
Dim Tick As Long
Const Title As String = "Communicator Console"

If ShowIt Then
    'Create an instance of the Win32 console window
    If Not ConsoleShown Then
        AllocConsole
        SetConsoleTitle Title
        
        'Get handles
        hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
        hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
        hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)
        
        'remove the close menu
        On Error Resume Next
        
        'Pause 150 'otherwise findwindow fails
        
        ConsolehWnd = 0
        Tick = GetTickCount()
        
        Do
            ConsolehWnd = FindWindow(vbNullString, Title)
            If ConsolehWnd Then
                ShowWindow ConsolehWnd, SW_HIDE
                Exit Do
            End If
        Loop While ((Tick + 1000) > GetTickCount()) And Not modVars.Closing
        
        'hide, set pos, implode, then show
        SetPos ConsolehWnd
        
        'ImplodeForm ConsolehWnd, True
        modImplode.AnimateAWindow ConsolehWnd, aImplode, , , True
        
        ShowWindow ConsolehWnd, SW_SHOWNORMAL
        'SetFocus ConsolehWnd
        
        SetForegroundWindow ConsolehWnd
        
        'If frmMainhWnd <> 0 Then SetForegroundWindow frmMainhWnd
        
        hSysMenu = GetSystemMenu(ConsolehWnd, 0)
        menuCount = GetMenuItemCount(hSysMenu)
        RemoveMenu hSysMenu, menuCount - 4, MF_BYPOSITION 'prevent user from closing t'console
                                                            'by removing close menu
        '##############
        'Set Attributes
        '##############
        'SetConsoleMode hConsoleOut, ENABLE_MOUSE_INPUT
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED + FOREGROUND_GREEN + FOREGROUND_BLUE
        
        'modVars.GetWindowsVersion , , , , Nt
        
        'If Nt Then
            'SetConsoleCtrlHandler 0&, CLng(True) 'Ignore CTRL-C
            'Call HandleConsoleEvents
        'End If
        
        
        
        If LenB(pConsoleText) Then
            ConsolePrint pConsoleText
        End If
        
        'force it to redraw (so redraw the close menu)
        'UpdateWindow ConsolehWnd
        'InvalidateRect ConsolehWnd, 0, 1 <-- doesn't work
        'DoEvents
    End If
    
    
    ConsoleShown = True
Else
    If ConsoleShown Then
        'ImplodeForm ConsolehWnd
        modImplode.AnimateAWindow ConsolehWnd, aImplode, True, , True
        
        'Call HandleConsoleEvents '(False)
        
        CloseHandle hConsoleOut
        CloseHandle hConsoleIn
        FreeConsole
        
    End If
    ConsoleShown = False
    hConsoleOut = 0
End If

'If Not AtStartup Then
'    frmMain.mnuConsole.Visible = ConsoleShown
'End If

End Sub

'Public Property Get IndentLevel() As Integer
'IndentLevel = pIndentLevel
'End Property
'
'Public Property Let IndentLevel(ByVal L As Integer)
'
'pIndentLevel = L
'
'End Property

Public Sub AddConsoleText(ByVal Text As String, _
        Optional ByVal ClearThenAdd As Boolean = False, _
        Optional ByVal IncreaseIndent As Boolean = False, _
        Optional ByVal DecreaseIndent As Boolean = False, _
        Optional ByVal NewLineBefore As Boolean = False) 'Colour As eConsoleColours = Normal)

Dim Tmp As String

If DecreaseIndent Then modConsole.Indent False

Tmp = IIf(NewLineBefore, vbNewLine, vbNullString) & Space$(pIndentLevel * IndentNo) & Text & vbNewLine

If Not ClearThenAdd Then
    pConsoleText = pConsoleText & Tmp
Else
    'ShowConsole False
    Call Clear
    pConsoleText = Text
    'ShowConsole
End If

If LenB(Tmp) Then
    ConsolePrint Tmp
End If

If modVars.bDebug Then
    pAddtoFile Tmp
End If

If IncreaseIndent Then modConsole.Indent

End Sub

Private Sub pAddtoFile(ByVal Text As String)
Dim f As Integer

f = FreeFile()

On Error Resume Next
Open GetDebugPath() For Append As #f
    Print #f, Text;
Close #f
On Error GoTo 0

End Sub

Public Function GetDebugPath() As String
Static sPath As String

If LenB(sPath) = 0 Then
    sPath = AppPath() & "Comm_Debug.txt"
End If

GetDebugPath = sPath
End Function

Private Sub ConsolePrint(ByVal MsgOut As String)

'If NewLine Then MsgOut = MsgOut & vbNewLine
If hConsoleOut Then
    WriteConsole hConsoleOut, MsgOut, Len(MsgOut), vbNull, vbNull
End If

End Sub


'Public Sub ProcessConsoleCommand(Optional ByVal DoLoop As Boolean = False)
'Dim StopLoop As Boolean
'
'If DoLoop Then
'    Do
'        Call pProcessConsoleCmd(StopLoop)
'        DoEvents '- repaint etc
'    Loop Until StopLoop
'Else
'    Call pProcessConsoleCmd(StopLoop)
'End If
'
'End Sub
'
'Private Sub pProcessConsoleCmd(ByRef StopLoop As Boolean)
'
'Dim Text As String, Tmp As String, tStr As String
'Dim i As Integer
'Dim SvrList As ListOfServer
'Const CmdLine As String = vbNewLine & "</"
'
'
'pConsoleText = pConsoleText & CmdLine 'only place where console print used, becuase don't want newline
'ConsolePrint CmdLine
'
''AddConsoleText CmdLine
'
''clear text
'Call FlushConsoleInputBuffer(hConsoleOut)
'
'Text = Trim$(ConsoleRead())
'
'pConsoleText = pConsoleText & Text
'
''PrintConsole Text - not needed - added itself
'
'Select Case LCase$(Text) 'do it here, so when text = whatever, it is proper case
'    Case "exit", "quit"
'        AddConsoleText vbNullString
'        ShowConsole False
'        StopLoop = True
'
'    Case "stop", vbNullString
'        SetForegroundWindow frmMainhWnd
'        StopLoop = True
'
'    Case "cls"
'        'frmMain.ClearRtfIn
'        AddConsoleText vbNullString, True
'
'    Case "cleartype"
'        frmMain.mnuOptionsMessagingClearTypeList_Click
'        AddConsoleText "Typing List Cleared"
'
'    Case "subclass"
'        modSubClass.SubClass frmMainhWnd, Not modSubClass.bSubClassing
'        'AddConsoleText "Subclassing: " & CStr(modSubClass.bSubClassing)
'
'    Case "log"
'        With frmMain.mnuOptionsMessagingLoggingConv
'            .Checked = Not .Checked
'            AddConsoleText "Logging: " & CStr(.Checked)
'        End With
'
'    Case "listen"
'        frmMain.Listen False
'
'    Case "close"
'        frmMain.CleanUp (True)
'        AddText "All Connections Closed", , True
'
'    Case "mysocket"
'
'        Tmp = "Socket: " & CStr(modMessaging.MySocket)
'
'        AddConsoleText Tmp
'
'    Case "status"
'        With frmMain
'           Tmp = "Status: " & GetStatus() & vbNewLine & _
'                    "Main Socket State: " & Mid$(GetState(.SckLC.State), 4)
'
'            AddConsoleText Tmp
'        End With
'
''    Case "listcomputers"
''
''        DoEvents
''
''        SvrList = modNetwork.EnumServer(modNetwork.SRV_TYPE_ALL)
''
''        With SvrList
''            If .Init Then
''                For i = LBound(.List) To UBound(.List)
''                    Tmp = Tmp & "Name: " & .List(i).ServerName & Space$(3) & "Comment: " & .List(i).Comment & vbNewLine
''                Next i
''            Else
''                Tmp = "Error in Listing Computers - '" & Str$(SvrList.LastErr) & "'"
''            End If
''
''        End With
''
''        AddConsoleText Tmp
'
'    Case "debug"
'
'        modVars.bDebug = Not modVars.bDebug
'
'        AddConsoleText "Debug: " & CStr(modVars.bDebug)
'
'    Case "help"
'
'        AddConsoleText "Stop - Exit the command loop" & vbNewLine & _
'                        "Cls - Clear Screen" & vbNewLine & _
'                        "Subclass - Toggle Subclassing" & vbNewLine & _
'                        "Log - Toggle Logging" & vbNewLine & _
'                        "Listen - Listen on Connection" & vbNewLine & _
'                        "Close - Close Connection" & vbNewLine & _
'                        "Connect * - Connect to *" & vbNewLine & _
'                        "Dev * - Dev + Password" & vbNewLine & _
'                        "Status - Show Current Status" & vbNewLine & _
'                        "Send * - Send * as Data" & vbNewLine & _
'                        "ClearType - Clears Typing List" & vbNewLine & _
'                        "Exec * - Execute a VBScript Command" & vbNewLine & _
'                        "Bat * - Execute a Batch Script" & vbNewLine & _
'                        "Debug - Toggle Console PrintOut (To File)" & vbNewLine & _
'                        "MySocket - Show my socket number" & vbNewLine & _
'                        "Help - Show this"
'
'
'    Case Else
'
'        tStr = Left$(Text, InStr(1, Text, Space$(1), vbTextCompare))
'
'        If LenB(tStr) = 0 Then tStr = Text
'
'        If Left$(tStr, 4) = "exec" Then
'            On Error Resume Next
'            Tmp = Trim$(Mid$(Text, 5))
'            On Error GoTo 0
'
'            If LenB(Tmp) Then
'                On Error GoTo ScriptError
'                frmMain.SC.ExecuteStatement Tmp
'                AddConsoleText "'" & Tmp & "' was executed"
'            Else
'                AddConsoleText "Please enter a command"
'            End If
'
'        'ElseIf Left$(tStr, 11) = "udpsendinfo" Then
'
'            'frmUDP.SendToSingle ip, frmUDP.UDPInfo & "Invite to  " & Me.LastName & " Rejected", False
'            'frmUDP.UDPListen
'
'        ElseIf InStr(1, tStr, "bat", vbTextCompare) Then
'
'            Dim Er As eConsoleRunBatErrors
'            Dim l As Long
'
'            On Error Resume Next
'            Tmp = Trim$(Mid$(Text, 4))
'            On Error GoTo 0
'
'            RunBat Tmp, Er, l
'
'            If Er = Success Then
'                AddConsoleText vbNewLine & "Batch File Ran (pID: " & CStr(l) & ")"
'            ElseIf Er = FileNotFound Then
'                AddConsoleText "Batch File Not Found"
'            Else
'                AddConsoleText "Unknown Error: " & Err.Description
'            End If
'
'        ElseIf InStr(1, tStr, "connect", vbTextCompare) Then
'            On Error Resume Next
'            Tmp = Trim$(Mid$(Text, 8))
'            On Error GoTo 0
'
'            'AddConsoleText "Returning Focus to Communicator"
'
'            frmMain.Connect Tmp
'
'            SetForegroundWindow frmMainhWnd
'
'            'need to otherwise, it hangs, and can't connect
'            StopLoop = True
'
'        ElseIf InStr(1, tStr, "dev", vbTextCompare) Then
'            On Error Resume Next
'            Tmp = Trim$(Mid$(Text, 4))
'            On Error GoTo 0
'
'            If bDevMode Then
'                modDev.setDevLevel modDev.Dev_Level_None, vbNullString
'            Else
'                If modDev.devLogin(Tmp) = False Then
'                    AddConsoleText "Incorrect DevMode Password"
'                End If
'            End If
'
'        ElseIf InStr(1, Text, "send", vbTextCompare) Then
'            On Error Resume Next
'            Tmp = Trim$(Mid$(Text, 5))
'            On Error GoTo 0
'
'            If Server Then
'                DistributeMsg Tmp, -1
'            Else
'                SendData Tmp
'            End If
'
'            AddConsoleText "Sent: '" & Tmp & "'"
'
'            DoEvents
'        Else
'            AddConsoleText "Error - Command Not Recognised"
'        End If
'
''        ElseIf InStr(1, tStr, "ping", vbTextCompare) Then
''            On Error Resume Next
''            Tmp = Trim$(Mid$(Text, 5))
''            On Error GoTo 0
''
''
''            AddConsoleText "Pinging..."
''
''            'Tmp = modWinsock.GetIPFromHostName(Tmp)
''
''            If LenB(Tmp) = 0 Then
''                If Server Then
''                    For i = 1 To frmMain.SockAr.Count - 1
''                        If frmMain.SockAr(i).State = sckConnected Then
''                            Exit For
''                        End If
''                    Next i
''
''                    If i = frmMain.SockAr.Count Then i = 1
''
''                    On Error Resume Next
''                    Tmp = frmMain.SockAr(i).RemoteHostIP
''                    On Error GoTo 0
''
''
''                Else
''                    Tmp = frmMain.SckLC.RemoteHostIP
''                End If
''            End If
''
''            If Tmp = vbNullString Then Tmp = "localhost"
''
''            AddConsoleText "Unsure who to ping - pinged myself"
''
''            Call frmMain.DoPing(Tmp, True)
'
'End Select
'
'Exit Sub
'ScriptError:
'AddConsoleText "Error - " & Err.Description
'End Sub

'---------------------


'Private Sub SetConsoleText(ByVal Text As String)
'
'ShowConsole False
'ShowConsole
'
'ConsolePrint Text
'
'End Sub


Private Function ConsoleRead() As String
Dim MsgIn As String * 256

SetForegroundWindow ConsolehWnd

Call ReadConsole(hConsoleIn, MsgIn, Len(MsgIn), vbNull, vbNull)

On Error Resume Next
ConsoleRead = Left$(MsgIn, InStr(MsgIn, Chr$(0)) - 3)
On Error GoTo 0

End Function

Private Sub SetPos(ByVal hWnd As Long)

SetWindowPos hWnd, 0, 50, 50, 700, 340, 0

End Sub

Public Sub Clear()
Dim Ret As Long

FillConsoleOutputCharacter hConsoleOut, 32, WindowWidth * WindowHeight, ByVal 0&, Ret
MoveCursor 0, 0

End Sub

'METHOD MoveCursor
'Moves the cursor to a specific location
Public Sub MoveCursor(X As Integer, Y As Integer)

Dim Crd As Long
CopyMemory Crd, X, 2
CopyMemory ByVal (VarPtr(Crd) + 2), Y, 2
SetConsoleCursorPosition hConsoleOut, Crd

End Sub

'PROPERTY WindowWidth
'Retrieves the width of the window
Public Property Get WindowWidth() As Long

Dim SBI As CONSOLE_SCREEN_BUFFER_INFO ', Ret As Long
GetConsoleScreenBufferInfo hConsoleOut, SBI
WindowWidth = SBI.srWindow.Right - SBI.srWindow.Left + 1

End Property

'PROPERTY WindowHeight
'Retrieves the height of the window
Public Property Get WindowHeight() As Long

Dim SBI As CONSOLE_SCREEN_BUFFER_INFO
GetConsoleScreenBufferInfo hConsoleOut, SBI
WindowHeight = SBI.srWindow.Bottom - SBI.srWindow.Top + 1

End Property

'Private Sub HandleConsoleEvents() 'Optional ByVal bHandle As Boolean = True)

'If bHandle Then
    'SetConsoleCtrlHandler AddressOf ConsoleHandler, CLng(True)
'Else
    'SetConsoleCtrlHandler ByVal 0&, 0&
'End If

'End Sub

'Private Function ConsoleHandler(CEvent As Long) As Long
'
''Select Case CEvent
''    Case CTRL_C_EVENT
''
''    Case CTRL_BREAK_EVENT
''
''    Case CTRL_CLOSE_EVENT
''
''    Case CTRL_LOGOFF_EVENT
''
''    Case CTRL_SHUTDOWN_EVENT
''
''End Select
'
'ConsoleHandler = 1
'
'End Function

'Public Function ReadChar() As Byte   'KeyAscii
'Dim Mode As Long
'Dim Char As Byte
'Dim CharsRead As Long
'
'' Flush input buffer.
'Call FlushConsoleInputBuffer(hConsoleOut)
'
'' Cache existing mode, so it can be restored.
'Call GetConsoleMode(hConsoleOut, Mode)
'
'' Set mode to not wait for an Enter key before returning.
'' No echo of character, either.
'Call SetConsoleMode(hConsoleOut, ByVal 0&)
'
'' Wait for a single keystroke.
'Call ReadConsole(hConsoleOut, Char, 1&, CharsRead, ByVal 0&)
'
'' Restore original mode.
'Call SetConsoleMode(hConsoleOut, Mode)
'
'' Return KeyAscii value of the key user pressed.
'ReadChar = Char
'End Function
