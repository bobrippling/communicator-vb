Attribute VB_Name = "modSettings"
Option Explicit

Private pbLoadingSettings As Boolean

Private pLastUpdate As Date

Public Const RegPath As String = "Software\MicRobSoft" 'Rob's Programs"
Public Const RegKeyPath As String = RegPath & "\Communicator"
Public Const MessagePath As String = RegKeyPath & "\Message Settings"
Public Const ColourPath As String = RegKeyPath & "\Colour Settings"
Public Const OtherPath As String = RegKeyPath & "\Other Settings"
Public Const SpeechPath As String = RegKeyPath & "\Speech Settings"
Public Const PosPath As String = RegKeyPath & "\Position Settings"
Public Const GraphicsPath As String = RegKeyPath & "\Graphics Settings"

Private Const Sep As String = "@"

Public Property Get bLoadingSettings() As Boolean
bLoadingSettings = pbLoadingSettings
End Property

Private Sub AddToFile(ByRef sFileContents As String, ByVal sParam As String)
sFileContents = sFileContents & sParam & vbNewLine
End Sub
Private Sub GetFromFile(sFileContents As String, sHeader As String, ByRef sParam() As String)
Dim i As Integer, j As Integer

i = InStr(1, sFileContents, sHeader)  '2=vbnewline
If i Then
    i = i + Len(sHeader) + 2
    'i = start of section
    'j = end of section
    
    j = InStr(i, sFileContents, "]")
    
    If j = 0 Then
        j = Len(sFileContents)
        'last section, should have newline after, but it might not
    End If
    
    
    sParam = Split(Mid$(sFileContents, i, j - i + 1), vbNewLine)
    
    'trim
    For i = UBound(sParam) To 0 Step -1
        If LenB(sParam(i)) = 0 Then
            ReDim Preserve sParam(0 To i - 1)
            Exit For
        End If
    Next i
    
End If

End Sub

Public Function ImportSettings(ByVal Path As String, bAddInfoText As Boolean, bAddErrorText As Boolean) As Boolean

Dim iError As Integer, f As Integer: f = FreeFile()
Dim sFile As String, sCurrentBlock As String
Const ERR_TYPE_MISMATCH = 13, ERR_SUBSCRIPT_OUT_OF_RANGE = 9

pbLoadingSettings = True

On Error GoTo EH
Open Path For Input As #f
    sFile = input(LOF(f), #f)
Close #f


sCurrentBlock = "Colour"
ProcessColour sFile, False
sCurrentBlock = "Window"
ProcessWindow sFile, False
sCurrentBlock = "Display"
ProcessDisplay sFile, False
sCurrentBlock = "Logging"
ProcessLogging sFile, False
sCurrentBlock = "FTP"
ProcessFTP sFile, False
sCurrentBlock = "Speech"
ProcessSpeech sFile, False
sCurrentBlock = "User"
ProcessUser sFile, False
sCurrentBlock = "Messaging"
ProcessMessaging sFile, False
sCurrentBlock = "Advanced"
ProcessAdvanced sFile, False

LoadBlockedIPs Path


If bAddInfoText Then
    AddText "Loaded Settings from '" & GetFileName(Path) & "'", TxtInfo, True
End If

pbLoadingSettings = False
ImportSettings = True

Exit Function
EH:
iError = Err.Number

'Debug.Assert False

If iError = ERR_TYPE_MISMATCH Or iError = ERR_SUBSCRIPT_OUT_OF_RANGE Then
    'Could be a type mismatch, if it was trying:
    '   tim.checked = sParam(i)
    '   [sParam(i) = ""]
    iError = 0
    
    
    'print to console/txtMain
    sCurrentBlock = "Error Loading " & sCurrentBlock & " settings block (Continuing with other blocks...)"
    AddConsoleText sCurrentBlock
    If bAddErrorText Then
        AddText sCurrentBlock, TxtError, True
    End If
    
    
    Resume Next '(in this procedure)
    'attempt to load next block
Else
    iError = 0
    
    If bAddErrorText Then
        AddText "Error Loading Settings - " & Err.Description, TxtError, True
        'AddText "Save the settings to overwrite the current settings file", TxtError, True
    End If
    Close #f
    ImportSettings = False
End If

pbLoadingSettings = False

End Function

Public Sub ExportSettings(ByVal Path As String, Optional bAddText As Boolean = True)
Dim f As Integer
Dim sFile As String

pbLoadingSettings = True

ProcessColour sFile, True
ProcessWindow sFile, True
ProcessDisplay sFile, True
ProcessLogging sFile, True
ProcessFTP sFile, True
ProcessSpeech sFile, True
ProcessUser sFile, True
ProcessMessaging sFile, True
ProcessAdvanced sFile, True

f = FreeFile()
Open Path For Output As #f
    Print #f, sFile;
Close #f


SaveBlockedIPs Path


If bAddText Then
    AddText "Exported Settings to '" & GetFileName(Path) & "'", , True
End If

pbLoadingSettings = False

Exit Sub
EH:
If bAddText Then
    AddText "Error Exporting Settings - " & Err.Description, , True
End If
Close #f
pbLoadingSettings = False
End Sub

Private Sub SaveBlockedIPs(sFile As String)
Dim sOutputPath As String
Dim f As Integer, i As Integer

sOutputPath = GetBlockedIPsPath(sFile)

On Error GoTo EH
f = FreeFile()
Open sOutputPath For Output As #f
    For i = LBound(modMessaging.BlockedIPs) To UBound(modMessaging.BlockedIPs)
        If LenB(modMessaging.BlockedIPs(i)) Then
            Print #f, modMessaging.BlockedIPs(i)
        End If
    Next i
Close #f


EH:
End Sub
Private Sub LoadBlockedIPs(sFile As String)
Dim sInputPath As String
Dim f As Integer, i As Integer
Dim List() As String

'Blocked bool is loaded from other config file

sInputPath = GetBlockedIPsPath(sFile)
i = 0

On Error GoTo EH
f = FreeFile()
Open sInputPath For Input As #f
    
    Do While Not EOF(f)
        ReDim Preserve List(i)
        Line Input #f, List(i)
        i = i + 1
    Loop
    
Close #f

i = i - 1
'i is now the ubound

If i > -1 Then
    'we have some
    ReDim modMessaging.BlockedIPs(i)
    
    For i = 0 To UBound(modMessaging.BlockedIPs)
        modMessaging.BlockedIPs(i) = List(i)
    Next i
    
Else
    ReDim modMessaging.BlockedIPs(0)
End If


EH:
End Sub

Private Function GetBlockedIPsPath(sFile As String) As String
GetBlockedIPsPath = GetFilePath(sFile) & "Blocked.cfg"
End Function

Private Sub ProcessUser(sFile As String, bOut As Boolean)
Const kStr = "[USER]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, frmMain.LastName
    AddToFile sFile, modPorts.MainPort
    AddToFile sFile, modPorts.StickPort
    AddToFile sFile, modPorts.SpacePort
    AddToFile sFile, modPorts.FTPort
    AddToFile sFile, modPorts.DPPort
    AddToFile sFile, modPorts.VoicePort
    AddToFile sFile, modMessaging.LastIP
    AddToFile sFile, frmMain.mnuFileMini.Checked
    AddToFile sFile, frmMain.mnuOptionsDPSaveAll.Checked
    AddToFile sFile, frmMain.mnuOnlineFTPServerMsg.Checked
    '-------------------
    If frmMain.mnuOptionsAdvInactive.Checked Then
        AddToFile sFile, frmMain.Inactive_Interval
    Else
        AddToFile sFile, "0"
    End If
    
    If frmMain.mnuOptionsAdvDisplayConnF(0).Checked Then
        AddToFile sFile, "0"
    Else
        AddToFile sFile, "1"
    End If
    '-------------------
    '------------------------------------------
    AddToFile sFile, modMessaging.bAllBlocked
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    frmMain.LastName = sParam(0)
    modPorts.MainPort = sParam(1)
    modPorts.StickPort = sParam(2)
    modPorts.SpacePort = sParam(3)
    modPorts.FTPort = sParam(4)
    modPorts.DPPort = sParam(5)
    modPorts.VoicePort = sParam(6)
    modMessaging.LastIP = sParam(7)
    frmMain.mnuFileMini.Checked = sParam(8)
    frmMain.mnuOptionsDPSaveAll.Checked = sParam(9)
    frmMain.mnuOnlineFTPServerMsg.Checked = sParam(10)
    
    frmMain.SetInactiveInterval CInt(sParam(11))
    frmMain.mnuOptionsAdvDisplayConnF_Click CInt(sParam(12))
    modMessaging.bAllBlocked = CBool(sParam(13))
End If

Erase sParam

End Sub
Private Sub ProcessMessaging(sFile As String, bOut As Boolean)
Const kStr = "[MESSAGING]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, frmMain.mnuOptionsMessagingReplaceQ.Checked
    AddToFile sFile, frmMain.rtfFontSize
    AddToFile sFile, frmMain.rtfFontName
    AddToFile sFile, frmMain.picColour.BackColor 'drawing
    AddToFile sFile, frmMain.cboWidth.Text
    AddToFile sFile, frmMain.cboRubber.Text
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    frmMain.mnuOptionsMessagingReplaceQ.Checked = sParam(0)
    frmMain.rtfFontSize = sParam(1)
    frmMain.rtfFontName = sParam(2)
    frmMain.picColour.BackColor = sParam(3) 'drawing
    frmMain.cboWidth.Text = sParam(4)
    frmMain.cboRubber.Text = sParam(5)
End If

Erase sParam

End Sub

Private Sub ProcessAdvanced(sFile As String, bOut As Boolean)
Const kStr = "[ADVANCED]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, frmMain.mnuOptionsHost.Checked
    AddToFile sFile, frmMain.mnuOptionsAdvHostMin.Checked
    AddToFile sFile, modPaths.SavedFilesPath
    AddToFile sFile, frmMain.mnuOptionsAdvAutoUpdate.Checked
    AddToFile sFile, frmMain.mnuOptionsAdvNoStandby.Checked
    AddToFile sFile, frmMain.mnuOptionsAdvNoStandbyConnected.Checked
    AddToFile sFile, frmMain.mnuOptionsAdvInactive.Checked
    AddToFile sFile, frmMain.mnuFileMini.Checked
    AddToFile sFile, modVars.bRetryConnection_Static
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    frmMain.mnuOptionsHost.Checked = sParam(0)
    frmMain.mnuOptionsAdvHostMin.Checked = sParam(1)
    modPaths.SavedFilesPath = sParam(2)
    frmMain.mnuOptionsAdvAutoUpdate.Checked = sParam(3)
    frmMain.mnuOptionsAdvNoStandby.Checked = sParam(4)
    frmMain.mnuOptionsAdvNoStandbyConnected.Checked = sParam(5)
    frmMain.mnuOptionsAdvInactive.Checked = sParam(6)
    frmMain.mnuFileMini.Checked = Not CBool(sParam(7))
    frmMain.mnuFileMini_Click
    modVars.bRetryConnection_Static = CBool(sParam(8))
End If

Erase sParam

End Sub

Private Sub ProcessSpeech(sFile As String, bOut As Boolean)
Const kStr = "[SPEECH]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, modSpeech.sBalloon
    AddToFile sFile, modSpeech.sHiBye
    AddToFile sFile, modSpeech.sQuestions
    AddToFile sFile, modSpeech.sReceived
    AddToFile sFile, modSpeech.Vol
    AddToFile sFile, modSpeech.Speed
    AddToFile sFile, modSpeech.sHi
    AddToFile sFile, modSpeech.sBye
    AddToFile sFile, modSpeech.sSayName
    AddToFile sFile, modSpeech.sGameSpeak
    AddToFile sFile, modSpeech.sOnlyForeground
    AddToFile sFile, modSpeech.sSayInfo
    AddToFile sFile, modSpeech.pitch
    AddToFile sFile, modSpeech.sSent
    AddToFile sFile, modSpeech.bHurgh
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    modSpeech.sBalloon = sParam(0)
    modSpeech.sHiBye = sParam(1)
    modSpeech.sQuestions = sParam(2)
    modSpeech.sReceived = sParam(3)
    modSpeech.Vol = sParam(4)
    modSpeech.Speed = sParam(5)
    modSpeech.sHi = sParam(6)
    modSpeech.sBye = sParam(7)
    modSpeech.sSayName = sParam(8)
    modSpeech.sGameSpeak = sParam(9)
    modSpeech.sOnlyForeground = sParam(10)
    modSpeech.sSayInfo = sParam(11)
    modSpeech.pitch = sParam(12)
    modSpeech.sSent = sParam(13)
    modSpeech.bHurgh = sParam(14): frmMain.mnuOptionsMessagingHurgh.Checked = modSpeech.bHurgh
End If

Erase sParam

End Sub

Private Sub ProcessFTP(sFile As String, bOut As Boolean)
Const kStr = "[FTP]"
Dim sParam() As String
Dim i As Integer

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    For i = frmMain.mnuOnlineFTPDLO.LBound To frmMain.mnuOnlineFTPDLO.UBound
        If frmMain.mnuOnlineFTPDLO(i).Checked Then
            AddToFile sFile, CStr(i)
            Exit For
        End If
    Next i
    If i = frmMain.mnuOnlineFTPDLO.UBound + 1 Then
        'Print #f,"0" 'HTTP by default
        AddToFile sFile, CStr(eFTP_Methods.FTP_Default)
    End If
    
    
    For i = frmMain.mnuOnlineFTPULO.LBound To frmMain.mnuOnlineFTPULO.UBound
        If frmMain.mnuOnlineFTPULO(i).Checked Then
            AddToFile sFile, CStr(i)
            Exit For
        End If
    Next i
    If i = frmMain.mnuOnlineFTPULO.UBound + 1 Then
        'Print #f,"0" 'HTTP by default
        AddToFile sFile, CStr(eFTP_Methods.FTP_Default)
    End If
    
    
    For i = frmMain.mnuOnlineFTPServerAr.LBound To frmMain.mnuOnlineFTPServerAr.UBound
        If frmMain.mnuOnlineFTPServerAr(i).Checked Then
            AddToFile sFile, CStr(i)
            Exit For
        End If
    Next i
    If i = frmMain.mnuOnlineFTPServerAr.UBound + 1 Then
        'Print #f,"0" 'Primary Server by default
        AddToFile sFile, "0" 'server #0
    End If
    
    AddToFile sFile, frmMain.mnuOnlineFTPPassive.Checked
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    frmMain.mnuOnlineFTPDLO_Click CInt(sParam(0))
    frmMain.mnuOnlineFTPULO_Click CInt(sParam(1))
    frmMain.mnuOnlineFTPServerAr_Click CInt(sParam(2))
    frmMain.mnuOnlineFTPPassive.Checked = sParam(3)
End If

Erase sParam

End Sub

Private Sub ProcessLogging(sFile As String, bOut As Boolean)
Const kStr = "[LOGGING]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, frmMain.mnuOptionsMessagingLoggingAutoSave.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingLoggingConv.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingLoggingDrawing.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingLoggingPrivate.Checked
    AddToFile sFile, modPaths.logPath
    AddToFile sFile, frmMain.mnuOptionsMessagingLoggingActivity.Checked
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    frmMain.mnuOptionsMessagingLoggingAutoSave.Checked = sParam(0)
    frmMain.mnuOptionsMessagingLoggingConv.Checked = sParam(1)
    frmMain.mnuOptionsMessagingLoggingDrawing.Checked = sParam(2)
    frmMain.mnuOptionsMessagingLoggingPrivate.Checked = sParam(3)
    modPaths.logPath = sParam(4)
    frmMain.mnuOptionsMessagingLoggingActivity.Checked = sParam(5)
End If

Erase sParam

End Sub

Private Sub ProcessWindow(sFile As String, bOut As Boolean)
Const kStr = "[WINDOW]"
Dim sParam() As String
Dim i As Integer

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, frmMain.mnuOptionsWindow2Animation.Checked 'IIf(frmMain.mnuOptionsWindow2All.Checked, eAnimType.aRandom, _
                        IIf(frmMain.mnuOptionsWindow2Slide.Checked, eAnimType.aSlide, _
                        IIf(frmMain.mnuOptionsWindow2Implode.Checked, eAnimType.aImplode, _
                        IIf(frmMain.mnuOptionsWindow2Fade.Checked, eAnimType.aFade, _
                        eAnimType.None))))
    
    AddToFile sFile, frmMain.mnuOptionsWindow2SingleClick.Checked
    'AddToFile sFile, frmMain.mnuOptionsWindow2BalloonInstance.Checked
    '---------------
    For i = frmMain.mnuOptionsAlertsStyle.LBound To frmMain.mnuOptionsAlertsStyle.UBound
        If frmMain.mnuOptionsAlertsStyle(i).Checked Then
            AddToFile sFile, CStr(i)
            Exit For
        End If
    Next i
    If i = frmMain.mnuOptionsAlertsStyle.UBound + 1 Then
        AddToFile sFile, "0"
    End If
    '---------------
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    'frmMain.AnimClick , CByte(sParam(0))
    frmMain.mnuOptionsWindow2Animation.Checked = sParam(0)
    frmMain.mnuOptionsWindow2SingleClick.Checked = sParam(1)
    'frmMain.mnuOptionsWindow2BalloonInstance.Checked = sParam(2)
    frmMain.mnuOptionsAlertsStyle_Click CInt(sParam(2))
End If

Erase sParam

End Sub

Private Sub ProcessDisplay(ByRef sFile As String, bOut As Boolean)
Const kStr = "[DISPLAY]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplaySmiliesComm.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplaySmiliesMSN.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplaySysUserName.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplayNewLine.Checked
    AddToFile sFile, frmMain.mnuOptionsTimeStamp.Checked
    AddToFile sFile, frmMain.mnuOptionsTimeStampInfo.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingShake.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplayCompact.Checked
    AddToFile sFile, frmMain.mnuOptionsFlashMsg.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplayShowHost.Checked
    AddToFile sFile, frmMain.mnuOptionsMessagingDisplayShowBlocked.Checked
    AddToFile sFile, frmMain.mnuOptionsAdvDisplayGlassBG.Checked
    AddToFile sFile, frmMain.mnuOptionsAdvDisplayVistaControls.Checked
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    frmMain.mnuOptionsMessagingDisplaySmiliesComm.Checked = sParam(0)
    frmMain.mnuOptionsMessagingDisplaySmiliesMSN.Checked = sParam(1)
    frmMain.ApplySmileySettings
    
    frmMain.mnuOptionsMessagingDisplaySysUserName.Checked = sParam(2)
    frmMain.mnuOptionsMessagingDisplayNewLine.Checked = sParam(3)
    frmMain.mnuOptionsTimeStamp.Checked = sParam(4)
    frmMain.mnuOptionsTimeStampInfo.Checked = sParam(5)
    frmMain.mnuOptionsMessagingShake.Checked = sParam(6)
    frmMain.mnuOptionsMessagingDisplayCompact.Checked = sParam(7)
    frmMain.mnuOptionsFlashMsg.Checked = sParam(8)
    frmMain.mnuOptionsMessagingDisplayShowHost.Checked = sParam(9)
    frmMain.mnuOptionsMessagingDisplayShowBlocked.Checked = sParam(10)
    
    frmMain.mnuOptionsAdvDisplayGlassBG.Checked = Not CBool(sParam(11))
    frmMain.mnuOptionsAdvDisplayGlassBG_Click
    
    frmMain.mnuOptionsAdvDisplayVistaControls.Checked = Not CBool(sParam(12))
    frmMain.mnuOptionsAdvDisplayVistaControls_Click
End If

Erase sParam

End Sub

Private Sub ProcessColour(ByRef sFile As String, bOut As Boolean)
Const kStr = "[COLOUR]"
Dim sParam() As String

If bOut Then
    AddToFile sFile, kStr
    '------------------------------------------
    AddToFile sFile, TxtInfo
    AddToFile sFile, TxtError
    AddToFile sFile, TxtReceived
    AddToFile sFile, TxtSent
    AddToFile sFile, TxtUnknown
    AddToFile sFile, TxtQuestion
    AddToFile sFile, TxtBackGround
    AddToFile sFile, TxtForeGround
    '------------------------------------------
    AddToFile sFile, vbNullString 'newline
Else
    GetFromFile sFile, kStr, sParam
    
    TxtInfo = sParam(0)
    TxtError = sParam(1)
    TxtReceived = sParam(2)
    TxtSent = sParam(3)
    TxtUnknown = sParam(4)
    TxtQuestion = sParam(5)
    TxtBackGround = sParam(6)
    TxtForeGround = sParam(7)
End If

Erase sParam

End Sub

'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################

Public Sub LoadUsedIPs()
Dim FName As String, sTxt As String
Dim f As Integer, i As Integer

FName = UsedIPsFName()
f = FreeFile()

On Error GoTo EH

Open FName For Input As #f
    
    Do While Not EOF(f)
        sTxt = vbNullString
        Line Input #f, sTxt
        
        i = InStr(1, sTxt, "@")
        
        If i Then
            modMessaging.AddUsedIP Left$(sTxt, i - 1), Mid$(sTxt, i + 1)
        Else
            modMessaging.AddUsedIP sTxt
        End If
        
    Loop
    
Close #f


Exit Sub
EH:
Close #f
AddConsoleText "Error Loading Used IPs: " & Err.Description
End Sub
Public Sub SaveUsedIPs()
Dim FName As String
Dim i As Integer, f As Integer

FName = UsedIPsFName()
f = FreeFile()

On Error GoTo EH

Open FName For Output As #f
    For i = 0 To UBound(UsedIPs)
        If LenB(UsedIPs(i).sIP) Then
            Print #f, UsedIPs(i).sIP & "@" & UsedIPs(i).sName
        End If
    Next i
Close #f

Exit Sub
EH:
Close #f
AddConsoleText "Error Saving Used IPs: " & Err.Description
End Sub
Private Function UsedIPsFName() As String
UsedIPsFName = GetUserSettingsPath() & "IPs.cfg" '& modVars.FileExt
End Function

'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################

'todo list
Public Sub ProcessTodoList()

ProcessTodoListContents GetTodoListContents()
On Error Resume Next
Kill GetTodoListPath()

End Sub

Public Function TodoItemPresent(sItem As String) As Boolean
Dim arContents() As String
Dim i As Integer

arContents = Split(GetTodoListContents(), vbNewLine)

For i = LBound(arContents) To UBound(arContents)
    If LCase$(arContents(i)) = sItem Then
        TodoItemPresent = True
        Exit For
    End If
Next i

Erase arContents

End Function

Public Function AddToTodoList(sItem As String) As Boolean
Dim sFilePath As String
Dim f As Integer

sFilePath = GetTodoListPath()

f = FreeFile()
On Error GoTo EH
Open sFilePath For Append As #f
    Print #f, sItem
Close #f


AddToTodoList = True
Exit Function
EH:
Close #f
End Function

Private Function GetTodoListContents() As String
Dim sFilePath As String, sContents As String
Dim f As Integer

sFilePath = GetTodoListPath()

If FileExists(sFilePath) Then
    
    f = FreeFile()
    On Error GoTo EH
    Open sFilePath For Input As #f
        sContents = input(LOF(f), f)
    Close #f
End If

GetTodoListContents = sContents

Exit Function
EH:
Close #f
End Function

Private Sub ProcessTodoListContents(sContents As String)
Dim arContents() As String
Dim i As Integer

arContents = Split(sContents, vbNewLine)

For i = LBound(arContents) To UBound(arContents)
    Select Case LCase$(arContents(i))
        Case "killold"
            modVars.KillOldVersion
        
    End Select
Next i

Erase arContents

End Sub

Public Function GetTodoListPath() As String
GetTodoListPath = GetUserSettingsPath() & "Mem.ini"
End Function

'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'update stuff

Public Property Let LastUpdate(nVal As Date)
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, modSettings.OtherPath, "Last Update", nVal
pLastUpdate = nVal
SavePLastUpdate
End Property

Public Property Get LastUpdate() As Date

Dim sFile As String, sTmp As String
Dim f As Integer

sFile = GetExtrasFile()

If LenB(sFile) Then
    If FileExists(sFile) Then
        
        f = FreeFile()
        Open sFile For Input As #f
            Line Input #f, sTmp
        Close #f
        
        On Error Resume Next
        pLastUpdate = CDate(sTmp)
    Else
        pLastUpdate = 0
    End If
    
End If


'###############

If pLastUpdate = 0 Then
    'take 6 days off today
    pLastUpdate = Date 'DateAdd("d", -6, Date)
End If

LastUpdate = pLastUpdate

Exit Property
EH:
LastUpdate = Date '"01/01/01"
End Property

Private Sub SavePLastUpdate()
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Last Update", pLastUpdate
Dim sFile As String
Dim f As Integer

sFile = GetExtrasFile()

If LenB(sFile) Then
    f = FreeFile()
    Open sFile For Output As #f
        Print #f, CStr(pLastUpdate)
    Close #f
End If

End Sub

'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################

Public Function GetTmpFilePath() As String
GetTmpFilePath = GetUserSettingsPath()
End Function

Public Function GetUserSettingsPath() As String
Dim sE As String
Dim MRS_Path As String, Comm_Path As String
'microbsoft

sE = Environ$("APPDATA")

MRS_Path = sE & IIf(Right$(sE, 1) <> "\", "\", vbNullString) & "MicRobSoft\"
If FileExists(MRS_Path, vbDirectory) = False Then
    On Error GoTo EH
    MkDir MRS_Path
End If

Comm_Path = MRS_Path & "Communicator\"
If FileExists(Comm_Path, vbDirectory) = False Then
    On Error GoTo EH
    MkDir Comm_Path
End If

GetUserSettingsPath = Comm_Path

EH:
End Function

Public Function GetExtrasFile() As String
Dim sTmp As String

sTmp = GetUserSettingsPath()

If LenB(sTmp) Then
    GetExtrasFile = sTmp & "Communicator_Extras.cfg"
End If

End Function

Public Function GetSettingsFile() As String
Dim sTmp As String

sTmp = GetUserSettingsPath()

If LenB(sTmp) Then
    GetSettingsFile = sTmp & "Communicator.cfg"
End If

End Function

'############################################################################################################################
'############################################################################################################################
'############################################################################################################################
'############################################################################################################################

Public Function LoadSettings() As Boolean
Dim TmpRPort As Integer, TmpLPort As Integer ', TmpN As Integer, TmpDrawHeight As Integer
Dim bSystray As Boolean
Dim TmpStr As String
Dim TmpI As Integer
Dim TmpB As Boolean

LoadSettings = True

If modRegistry.regDoes_Key_Exist(HKEY_CURRENT_USER, RegKeyPath) = False Then
    LoadSettings = False
    AddConsoleText "Error Loading Settings"
    Exit Function
End If

On Error Resume Next

TmpStr = Trim$(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Name"))

If LenB(TmpStr) Then
    'frmMain.txtName.Text = DefaultName
    frmMain.Rename TmpStr
End If

'Colour---------------------------
TxtError = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "Error"))
TxtInfo = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "Info"))
TxtReceived = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "Recieved"))
TxtSent = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "Sent"))
TxtUnknown = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "Unknown"))
TxtBackGround = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "BackGround"))
TxtQuestion = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "Question"))
TxtForeGround = val(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, ColourPath, "ForeGround"))

If TxtError = 0 Then
    If TxtInfo = 0 Then
        If TxtReceived = 0 Then
            If TxtSent = 0 Then
                If TxtUnknown = 0 Then
                    If TxtBackGround = 0 Then
                        If TxtQuestion = 0 Then 'corrupt/wrong etc
                            If TxtForeGround = 0 Then modVars.SetDefaultColours
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
            
With frmMain
    'Message--------------------------
    .mnuOptionsTimeStamp.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "TimeStamp")
    .mnuOptionsTimeStampInfo.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "TimeStamp All")
    .mnuOptionsTimeStampInfo.Enabled = .mnuOptionsTimeStamp.Checked
    .mnuOptionsFlashMsg.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Flash Form")
    '.mnuOptionsFlashInvert.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Flash Form Invert")
    '.mnuOptionsMessagingColours.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Coloured Text")
    .mnuOptionsMessagingShake.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Shake")
    .picColour.BackColor = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Draw Colour")
    .cboWidth.Text = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Width")
    
    .cboRubber.Text = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Rubber Width")
    If LenB(Trim$(.cboRubber.Text)) = 0 Then
        .cboRubber.Text = 20
    ElseIf val(.cboRubber.Text) < 1 Then
        .cboRubber.Text = 1
    ElseIf val(.cboRubber.Text) > 100 Then
        .cboRubber.Text = 100
    End If
    
    .mnuOptionsMessagingLoggingConv.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Log")
    
    '.mnuOptionsMessagingDisplaySmiliesEnable.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Smilies")
    '.rtfIn.EnableSmiles = .mnuOptionsMessagingDisplaySmiliesEnable.Checked
    
    'Other----------------------------
    TmpRPort = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "RPort")
    TmpLPort = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "LPort")
    
    '.mnuOptionsBalloonMessages.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Balloon")
    .mnuOptionsHost.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Host")
    .mnuOptionsAdvHostMin.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "AutoMinimize")
    .mnuOptionsStartup.Checked = modStartup.WillRunAtStartup(App.EXEName)
    .mnuOptionsAdvDisplayStyles.Checked = modDisplay.VisualStyle 'modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "XP")
    
    TmpI = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Animation Type")
    'frmMain.AnimClick , TmpI
    frmMain.mnuOptionsWindow2Animation.Checked = TmpI
    
    '.mnuOptionsAdvInactive.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Inactive")
    
    
    .mnuOptionsWindow2SingleClick.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Tray Single Click")
    frmSystray.mnuPopupSingleClick.Checked = .mnuOptionsWindow2SingleClick.Checked
    
    .mnuDevShowAll.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "DevShowAll")
    .mnuDevShowCmds.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "DevShowDevCmds")
    
    'TmpN = .DrawHeight
    'TmpDrawHeight = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Drawing Height")
    
    'If TmpN <> TmpDrawHeight Then .Height = .Height + (TmpDrawHeight - TmpN)
    
    '.DrawHeight = TmpDrawHeight
    TmpStr = Trim$(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Font"))
    
    If LenB(TmpStr) Then
        '.rtfIn.FontName = TmpStr
        .rtfFontName = TmpStr
    End If
    
    TmpI = val(Trim$(modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Size")))
    
    If TmpI Then
        If TmpI >= MinFont Then
            If TmpI <= MaxFont Then
                '.rtfIn.Font.Size = TmpI
                .rtfFontSize = TmpI
            End If
        End If
    End If
    
    
    .mnuOptionsMessagingReplaceQ.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Replace Question")
    .mnuOptionsMessagingEncrypt.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Encrypt")
    .mnuOptionsMessagingDisplayShowBlocked.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Show Blocked")
    .mnuOptionsMessagingLoggingAutoSave.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, MessagePath, "Auto Save")
End With


'Rest-----------------------------
modMessaging.colour = frmMain.picColour.BackColor
'frmMain.mnuOptionsFlashInvert.Enabled = frmMain.mnuOptionsFlashMsg.Checked

modPaths.SavedFilesPath = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Save Location")

If CBool(TmpRPort) And Status = Idle Then
    MainPort = TmpRPort
End If

'If TmpLPort <> 0 And Status = Idle Then
    'LPort = TmpLPort
'End If

'frmSystray.mnuPopupAnim.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Animate Tray")
frmMain.mnuOptionsMessagingDisplaySysUserName.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, _
    OtherPath, "System Username")

'bSystray = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Systray")
'If bSystray Then
    'Call DoSystray(True)
'Else
    'Call DoSystray(False)
'End If


'speech settings
modSpeech.sBalloon = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Balloon")
modSpeech.sHiBye = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "HiBye")
modSpeech.sQuestions = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Questions")
modSpeech.sReceived = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Received")
modSpeech.sHi = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Hi")
modSpeech.sBye = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Bye")
modSpeech.sSayName = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Say Name")
modSpeech.Vol = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Volume")
modSpeech.Speed = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Speed")
modSpeech.sGameSpeak = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Game")
modSpeech.sOnlyForeground = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, SpeechPath, "Foreground")

'frmMain.mnuOptionsWindow2BalloonInstance.Checked = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, OtherPath, "Balloon Second Instance")

'pos settings
TmpI = -1
TmpI = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, PosPath, "X")
modImplode.fmX = TmpI
TmpI = -1
TmpI = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, PosPath, "Y")
modImplode.fmY = TmpI

If modLoadProgram.bVistaOrW7 Then
    TmpB = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, GraphicsPath, "Glass Border")
    
    If TmpB Then
        frmMain.mnuOptionsAdvDisplayGlassBG.Checked = False 'will be reversed
        frmMain.mnuOptionsAdvDisplayGlassBG_Click
    End If
    
    TmpB = modRegistry.regQuery_A_Key(HKEY_CURRENT_USER, GraphicsPath, "Vista Controls")
    
    If TmpB Then
        If modDisplay.VisualStyle() Then
            frmMain.mnuOptionsAdvDisplayVistaControls.Checked = False 'will be reversed
            frmMain.mnuOptionsAdvDisplayVistaControls_Click
        End If
    End If
End If


AddConsoleText "Loaded Settings"

End Function

Public Function SaveSettings()

Call DelSettings

SaveSettings = True
On Error GoTo EH

modRegistry.regCreate_A_Key HKEY_CURRENT_USER, RegKeyPath

modRegistry.regCreate_A_Key HKEY_CURRENT_USER, ColourPath
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "Error", TxtError
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "Info", TxtInfo
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "Recieved", TxtReceived
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "Sent", TxtSent
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "Unknown", TxtUnknown
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "Question", TxtQuestion
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "BackGround", TxtBackGround
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, ColourPath, "ForeGround", TxtForeGround

modRegistry.regCreate_A_Key HKEY_CURRENT_USER, MessagePath
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Flash Form", frmMain.mnuOptionsFlashMsg.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Flash Form Invert", frmMain.mnuOptionsFlashInvert.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "TimeStamp", frmMain.mnuOptionsTimeStamp.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "TimeStamp All", frmMain.mnuOptionsTimeStampInfo.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Name", frmMain.LastName
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Draw Colour", frmMain.picColour.BackColor
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Width", frmMain.cboWidth.Text
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Rubber Width", frmMain.cboRubber.Text
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Coloured Text", frmMain.mnuOptionsMessagingColours.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Log", frmMain.mnuOptionsMessagingLoggingConv.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Smilies", frmMain.mnuOptionsMessagingDisplaySmiliesEnable.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Shake", frmMain.mnuOptionsMessagingShake.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Size", frmMain.rtfFontSize 'frmMain.rtfIn.Font.Size
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Font", frmMain.rtfFontName 'frmMain.rtfIn.Font.Name
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Replace Question", frmMain.mnuOptionsMessagingReplaceQ.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Encrypt", frmMain.mnuOptionsMessagingEncrypt.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Show Blocked", frmMain.mnuOptionsMessagingDisplayShowBlocked.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, MessagePath, "Auto Save", frmMain.mnuOptionsMessagingLoggingAutoSave.Checked

modRegistry.regCreate_A_Key HKEY_CURRENT_USER, OtherPath
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Host", frmMain.mnuOptionsHost.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "RPort", MainPort
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "LPort", LPort
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Balloon", frmMain.mnuOptionsBalloonMessages.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Startup", frmMain.mnuOptionsStartup.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Animation Type", frmMain.mnuOptionsWindow2Animation.Checked
              'IIf(frmMain.mnuOptionsWindow2All.Checked, eAnimType.aRandom, _
              IIf(frmMain.mnuOptionsWindow2Slide.Checked, eAnimType.aSlide, _
              IIf(frmMain.mnuOptionsWindow2Implode.Checked, eAnimType.aImplode, _
              IIf(frmMain.mnuOptionsWindow2Fade.Checked, eAnimType.aFade, _
              eAnimType.None))))

modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Tray Single Click", frmMain.mnuOptionsWindow2SingleClick.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Inactive", frmMain.mnuOptionsAdvInactive.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Drawing Height", frmMain.DrawHeight
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "AutoMinimize", frmMain.mnuOptionsAdvHostMin.Checked

modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "DevShowAll", IIf(bDevMode, frmMain.mnuDevShowAll.Checked, False)
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "DevShowDevCmds", IIf(bDevMode, frmMain.mnuDevShowCmds.Checked, False)
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Save Location", modPaths.SavedFilesPath
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Animate Tray", frmSystray.mnuPopupAnim.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Balloon Second Instance", frmMain.mnuOptionsWindow2BalloonInstance.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "System Username", frmMain.mnuOptionsMessagingDisplaySysUserName.Checked
SavePLastUpdate


modRegistry.regCreate_A_Key HKEY_CURRENT_USER, SpeechPath
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Balloon", modSpeech.sBalloon
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "HiBye", modSpeech.sHiBye
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Questions", modSpeech.sQuestions
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Received", modSpeech.sReceived
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Volume", modSpeech.Vol
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Speed", modSpeech.Speed
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Hi", modSpeech.sHi
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Bye", modSpeech.sBye
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Say Name", modSpeech.sSayName
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Game", modSpeech.sGameSpeak
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, SpeechPath, "Foreground", modSpeech.sOnlyForeground

modRegistry.regCreate_A_Key HKEY_CURRENT_USER, PosPath
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, PosPath, "Y", frmMain.Top
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, PosPath, "X", frmMain.Left

modRegistry.regCreate_A_Key HKEY_CURRENT_USER, GraphicsPath
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, GraphicsPath, "Glass Border", frmMain.mnuOptionsAdvDisplayGlassBG.Checked
modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, GraphicsPath, "Vista Controls", frmMain.mnuOptionsAdvDisplayVistaControls.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, GraphicsPath, "Auto Save/Load", frmMain.mnuOptionsAdvDisplaySL.Checked

AddConsoleText "Saved Settings"

'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "XP", frmMain.mnuOptionsXP.Checked
'modRegistry.regCreate_Key_Value HKEY_CURRENT_USER, OtherPath, "Systray", InTray
'modstartup.SetRunAtStartup(app.EXEName,app.Path, frmmain.mnuOptionsStartup.Checked)

Exit Function
EH:
SaveSettings = False
End Function

Public Sub DelSettings()

modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegKeyPath, "Message Settings"
modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegKeyPath, "Other Settings"
modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegKeyPath, "Colour Settings"
modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegKeyPath, "Speech Settings"
modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegKeyPath, "Position Settings"
modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegKeyPath, "Graphics Settings"
modRegistry.regDelete_A_Key HKEY_CURRENT_USER, RegPath, "Communicator"

AddConsoleText "Deleted Settings"

End Sub

