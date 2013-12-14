Attribute VB_Name = "modConsole"
'http://visualbasic.about.com/od/learnvb6/l/bldykvb6dosa.htm

Option Explicit

Private pConsoleText As String
Private ConsolehWnd As Long

Public frmMainhWnd As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


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

' Global Variables
Private hConsoleIn As Long ' console input handle
Private hConsoleOut As Long ' console output handle
Private hConsoleErr As Long ' console error handle

Public Sub AddConsoleText(ByVal Text As String, Optional ByVal ClearThenAdd As Boolean = False, _
                                                        Optional ByVal OverideReload As Boolean = False, _
                                                        Optional ByVal NewLine As Boolean = True)

If ClearThenAdd = False Then
    pConsoleText = pConsoleText & IIf(NewLine, vbNewLine, vbNullString) & Text
Else
    pConsoleText = IIf(Len(Text) > 0, Text, vbNullString)
End If

If ConsoleShown And Not OverideReload Then SetConsoleText vbNullString 'pConsoleText

If OverideReload Then ConsolePrint Text, NewLine

End Sub

Public Sub ShowConsole(Optional ByVal ShowIt As Boolean = True, Optional ByVal TypeIn As Boolean = False)

Dim UserInput As String, Command As String, Param As String
Dim hSysMenu As Long, menuCount As Long
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
        
        Pause 10 'otherwise findwindow fails
        
        ConsolehWnd = FindWindow(vbNullString, Title)
        SetPos ConsolehWnd
        SetForegroundWindow frmMainhWnd
        hSysMenu = GetSystemMenu(ConsolehWnd, 0)
        menuCount = GetMenuItemCount(hSysMenu)
        RemoveMenu hSysMenu, menuCount - 4, MF_BYPOSITION 'prevent user from closing t'console
                                                            'by removing close menu
        
        SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED + FOREGROUND_GREEN + FOREGROUND_BLUE
        
        
        
        If Len(pConsoleText) > 0 Then
            ConsolePrint pConsoleText, False
        End If
        
    End If
    
    If TypeIn Then
        DoEvents
        ProcessConsoleCommand
    End If
    
    ConsoleShown = True
Else
    If ConsoleShown Then
        FreeConsole
    End If
    ConsoleShown = False
End If

frmMain.mnuConsole.Visible = ConsoleShown

End Sub

Public Sub ProcessConsoleCommand()
Dim Text As String, Tmp As String
Const CmdLine As String = vbNewLine & "</"

'pConsoleText = pConsoleText & CmdLine
'
'ConsolePrint CmdLine, False

pConsoleText = pConsoleText & vbNewLine

AddConsoleText CmdLine, , True, False

Text = LCase$(Trim$(ConsoleRead()))

'AddConsoleText Text, , True, False

pConsoleText = pConsoleText & Text

Select Case Text
    Case "exit", "quit"
        ShowConsole False
        
    Case vbNullString
        SetForegroundWindow frmMainhWnd
        
    Case "cls"
        'frmMain.ClearRtfIn
        AddConsoleText vbNullString, True, , False
        
        
    Case "subclass"
        modSubClass.SubClass frmMainhWnd, Not modSubClass.bSubClassing
        ConsolePrint "Subclassing: " & CStr(modSubClass.bSubClassing)
        
    Case "log"
        With frmMain.mnuOptionsMessagingLog
            .Checked = Not .Checked
            ConsolePrint "Logging: " & CStr(.Checked)
        End With
        
    Case "listen"
        frmMain.Listen False
        
    Case "close"
        frmMain.CleanUp
        
        
    Case Else
        If InStr(1, Text, "connect", vbTextCompare) Then
            Tmp = Trim$(Mid$(Text, 8))
            frmMain.Connect Tmp
            
        ElseIf InStr(1, Text, "dev", vbTextCompare) Then
            Tmp = Trim$(Mid$(Text, 4))
            
            If Tmp = "5ae" Then
                DevMode Not bDevMode
            Else
                AddConsoleText "Incorrect DevMode Password", , True
            End If
        Else
            AddConsoleText "Error - Command Not Recognised", , True
        End If
End Select


End Sub

Private Sub ConsolePrint(ByVal MsgOut As String, Optional ByVal NewLine As Boolean = True)

If NewLine Then MsgOut = MsgOut & vbNewLine

WriteConsole hConsoleOut, MsgOut, Len(MsgOut), vbNull, vbNull

End Sub

Private Sub SetConsoleText(ByVal Text As String)

ShowConsole False
ShowConsole

ConsolePrint Text

End Sub

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

