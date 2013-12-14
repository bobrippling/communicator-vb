Attribute VB_Name = "modLoadProgram"
Option Explicit

Private Declare Sub apiExitProcess Lib "kernel32" Alias "ExitProcess" (ByVal uExitCode As Long)

'prev instance stuff
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Declare Sub api_InitCommonControls Lib "comctl32.dll" Alias "InitCommonControls" ()

Public bSlow As Boolean
Public bLoading As Boolean
Public bIsCommT As Boolean
Public bJustUpdated As Boolean
Public bVistaOrW7 As Boolean
Public bAprilFools As Boolean
Public bSafeMode As Boolean
Public bAllowXPButtons As Boolean
Private bSkipDLLCheck As Boolean

Public bIsIDE As Boolean
Public bEnableKeyboardHook As Boolean

Public frmSplash_Loaded As Boolean
Public frmMain_Loaded As Boolean
Public frmMini_Loaded As Boolean, bLoadMiniAtStartup As Boolean
Public frmInfo_Loaded As Boolean
Public frmThumbNail_Loaded As Boolean
Public frmManualFT_Loaded As Boolean
Public frmDev_Loaded As Boolean
Public frmVoiceTransfers_Loaded As Boolean
Public LoadStart As Long


'check other communicator responsiveness
Private Const SMTO_BLOCK = &H1
Private Const SMTO_ABORTIFHUNG = &H2
Private Const WM_NULL = &H0
'Private Const HWND_BROADCAST As Long = &HFFFF&
'send to all Windows

Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" _
    (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
    ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

'Private Declare Function IsHungAppWindow Lib "user32" (ByVal hWnd As Long) As Long

'threading
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Enum eExit_Codes
    e_Normal_Unload = 0
    e_PrevInstance = 1
    e_Load_Error = 2
End Enum

Public Const Systray_Caption As String = "Communicator Systray - MicRobSoft"

Public Sub ExitProcess(ByRef lExitCode As eExit_Codes)

Debug.Assert False 'breakpoint

AddConsoleText "ExitProcess() Called - Exit Code: " & CStr(lExitCode)

If Not bIsIDE Then
    apiExitProcess lExitCode
End If

End Sub

Sub Main()

'Dim ProcessID As Long, curProcessID As Long
'Dim bCan As Boolean
'Dim hWnd As Long
'
''hWnd = FindWindow(App.ProductName, vbNullString) 'vbNullString, "Communicator")
'hWnd = FindWindow(vbNullString, "Systray Communicator - Robco")
'
'bCan = False
'
'AddConsoleText "hWnd: " & CStr(hWnd)
'
'If hWnd <> 0 Then
'    'find it's pID
'    GetWindowThreadProcessId hWnd, ProcessID
'    'find my pID
'    curProcessID = GetCurrentProcessId()
'
'    AddConsoleText "Other Window Found, pID: " & CStr(ProcessID) & Space$(3) & "My pID: " & CStr(curProcessID)
'
'    If curProcessID <> ProcessID Then
'        bCan = True 'If it's a different process..
'    End If
'Else
'    'no other window found
'    AddConsoleText "No Other Window Found"
'    bCan = True
'End If
'
'AddConsoleText "Main Load Procedure (bCan): " & CStr(bCan)
'AddConsoleText vbNullString
'
'If bCan Then
    '---------
    Dim bConsole As Boolean ', bStealth As Boolean
    Dim CmdLine As String, IPText As String, DebugPath As String
    Dim Ans As VbMsgBoxResult
    Dim Tmp As String, sTxt As String
    Dim sLoadTime As Single
    
    Const Obtain_Str = "Obtaining Remote IP"
    
    'for win ver
    Dim iMaj As Long, iMin As Long, iRev As Long, bNt As Boolean, bVista As Boolean
    
    '###################################################################################################
    LoadStart = GetTickCount()
    bLoading = True
    
    CmdLine = LCase$(Command$) 'IDE cmdline: "/subclass 0 /dev :Comm: /vista 1 /internet 1 /logall 0"
    
    bSlow = CommandLinePresent("slow", CmdLine)
    bConsole = CommandLinePresent("console", CmdLine)
    bIsCommT = (App.CompanyName = "CptNeutral")
    modVars.bStealth = CommandLinePresent("stealth", CmdLine) Or bIsCommT
    modVars.bStartup = CommandLinePresent("startup", CmdLine)
    modVars.bDebug = CommandLinePresent("debug", CmdLine)
    bSafeMode = CommandLinePresent("safemode", CmdLine)
    modVars.bNoInternet = CommandLinePresent("internet", CmdLine) 'Or bSafeMode
    bJustUpdated = CommandLinePresent("killold", CmdLine) Or TodoItemPresent("killold")
    bSkipDLLCheck = CommandLinePresent("skipdll", CmdLine)
    
    
    bAprilFools = CBool(Left$(Format$(Date$, "dd/mm/yyyy"), 5) = "1/04/") 'commandline checked in MainWindow_load
    modLoadProgram.bIsIDE = IsIDE()
    
    modImplode.fmX = -1
    modImplode.fmY = -1
    
    If modVars.bDebug Then
        If FileExists(modConsole.GetDebugPath()) Then
            On Error Resume Next
            Kill modConsole.GetDebugPath()
        End If
    End If
    
    '###################################################################################################
    
    InitCommonControls
    AddConsoleText "Initialised Common Controls"
    
    
    modWinsock.InitWinsock
    'AddConsoleText "Initialised Winsock"
    
    modLogging.bLogging = True
    AddConsoleText "Started Logging"
    
    
    modSubClass.SubClass_Init
    modVars.Vars_Init
    modFTP.FTP_Init
    modDev.initDev
    modVars.bDisableAddText = False
    
    Randomize CSng(Right$(Timer, 1)) + CSng(Right$(GetTickCount(), 1)) + CSng(Right$(Time$, 1))
    'Randomize Int(Right$(Time$, 1)) + Int(Right$(Time$, 2))
    
    
    If Not bStealth Then
        Load frmSplash '########################################################
        AddConsoleText "Loaded Splash Window", , True
        frmSplash.Refresh
        
        If bStartup Then
            'frmSplash.lblStatus.Caption = "Loading Communicator..."
            SetOnTop frmSplash.hWnd '+Show
            Pause 750
            Unload frmSplash
            AddConsoleText "Unloaded Splash Window", , , True
        End If
    End If
    
    SetSplashProgress 5
    
    
    If modVars.bDebug Then
        DebugPath = AppPath() & "Debug.txt"
        
        If FileExists(DebugPath) Then
            On Error Resume Next
            Kill DebugPath
            On Error GoTo 0
        End If
    End If
    
    
    'If modVars.CanUseInet Then
        'modVars.LoadInet
    'End If
    
    If bConsole And Not bStealth Then
        modLoadProgram.SetSplashInfo "Loading Console..."
        ShowConsole , True
    End If
    
    
    modLoadProgram.SetSplashInfo "Checking OS Version..."
    'getwindowsver console stuff
    modVars.GetWindowsVersion iMaj, iMin, iRev, , bNt, bVistaOrW7
    AddConsoleText "Window's Version: " & iMaj & Dot & iMin & Dot & iRev, , , , True
    AddConsoleText "Windows NT: " & CStr(bNt)
    AddConsoleText "Windows Vista: " & CStr(bVistaOrW7)
    AddConsoleText vbNullString
    bAllowXPButtons = Not modLoadProgram.bVistaOrW7
    
    
    modAlert.Init
    
    
    If bSkipDLLCheck = False Then 'bSafeMode Then
        'check dlls
        AddConsoleText "Checking for vital DLLs...", , True, , True
        modLoadProgram.SetSplashInfo "Checking for vital DLLs..."
        
        Tmp = DllsExist()
        
        If LenB(Tmp) Then
            
            sTxt = "Vital DLLs don't exist (" & Tmp & ")"
            
            AddConsoleText sTxt
            modLogging.LogEvent sTxt, eLogEventTypes.LogWarning
            
            Ans = LP_MsgBoxEx("Vital Dlls are missing (" & Tmp & ")" & vbNewLine & _
                "Communicator may be able to run without them" & vbNewLine & vbNewLine & _
                 "Continue Anyway?", "Communicator requires certain files to work properly. However, Communicator may be able to run without those files. The choice is yours...", _
                 vbYesNo Or vbExclamation Or vbDefaultButton2, "Warning")
            
            AddConsoleText "User chose to " & IIf(Ans = vbYes, "exit the program", "continue"), , , True
            
            If Ans = vbNo Then
                Unload frmSplash
                Exit Sub
            Else
                LP_MsgBoxEx "If communicator freezes, it is due to missing Dll files," & vbNewLine & _
                    "Please contact " & App.CompanyName & " or whoever gave you this program, preferably with a large stick." & vbNewLine & _
                    "If you don't know how to contact " & App.CompanyName & " then it's tough luck.", _
                    "These DLL files are needed. Communicator can freeze, if they are not present. Attempt to contact " & App.CompanyName, _
                    vbInformation, "Communicator"
                
            End If
        Else
            AddConsoleText "All Dlls present", , , True
        End If
        
    Else
        AddConsoleText "Dlls check skipped", , , True
    End If
    
    SetSplashProgress 10
    
    
    
    If CheckPrevInstance(CmdLine) Then
        bLoading = False
        SetSplashProgress 100
        
        AddConsoleText "Exiting Program... (Previous Instance)"
        
        'Unload frmSplash <-- done below
        'AddConsoleText "Unloaded Splash Window"
        
        ExitProgram
        'AddConsoleText "ExitProgram() Called"
        
        ExitProcess e_PrevInstance
    Else
        
        'If modLoadProgram.bslo Then
            AddConsoleText Obtain_Str
            modLoadProgram.SetSplashInfo Obtain_Str
            modWinsock.ObtainRemoteIP
        'End If
        
        modLoadProgram.SetSplashInfo "Loading Main Window..."
        AddConsoleText "Loading Main Window...", , True, , True
        
        If Not bStartup And Not bStealth Then frmSplash.ZOrder vbBringToFront
        
        SetSplashProgress 20
        On Error GoTo frmMain_LoadEH 'in case previnstance and they unload
        Load frmMain '#############################################################
        
        On Error GoTo EndOfSub
        frmMain.ZOrder vbSendToBack
        
        
    '    If bIsIDE Then <-- done in form_load()
    '        On Error Resume Next
    '        frmMain.rtfIn.DisableURLHook
    '    End If
        
        
        
        SetSplashProgress 85
        
        If bConsole And Not bStealth Then ShowConsole True 'show the mnuconsole on frmmain
        
        
        modLoadProgram.SetSplashInfo "Initialising Network Broadcast..."
        AddConsoleText "Initialising Network Broadcast...", , True, , True
        Load frmUDP
        frmUDP.Visible = False
        AddConsoleText "Initialised Network Broadcast", , , True
        AddConsoleText vbNullString
        modLoadProgram.SetSplashInfo "Initialised Network Broadcast"
        
        SetSplashProgress 95
        
        frmMain.ZOrder vbBringToFront
        
        SetFocus2 frmMain
        On Error GoTo EndOfSub
        
        sLoadTime = (GetTickCount() - LoadStart) / 1000
        If Not bSafeMode Then
'            Select Case sLoadTime
'                Case Is < 0.5
'                    sTxt = "Smooth criminal"
'                Case Is < 2
'                    sTxt = "Whippy"
'                Case Is < 5
'                    sTxt = "Average..."
'                Case Is < 10
'                    sTxt = "...with the speed of a striking slug!"
'                Case Is < 15
'                    sTxt = "Call yourself a computer?" '"What type of chip you got in there? McCain's ovenfry?"
'                Case Else
'                    sTxt = "That was so fast i almost lost count of how long it took" '& vbNewLine & _
'                        "Either that, or you're using a Mac."
'            End Select
            
            
            'settings have been loaded <=> iif(timestamp
            'AddText Trim$(InfoStart) & vbNewLine & _
                    "Load Time: " & FormatNumber$(sLoadTime, 2, vbTrue, vbFalse, vbFalse) & " seconds - " & sTxt & _
                    IIf(frmMain.mnuOptionsTimeStamp.Checked, _
                        vbNewLine & FormatDateTime$(Time$, vbLongTime) & " - " & FormatDateTime$(Date$, vbLongDate), _
                        vbNullString) & vbNewLine & _
                    Trim$(InfoEnd), , , True
            
            If CommandLinePresent("cls", CmdLine) = False Then
                AddText "Load Time: " & FormatTimeElapsed(sLoadTime) & ", Communicator Version: " & modVars.GetVersion(), , True
            End If
            
        Else
            If CommandLinePresent("cls", CmdLine) = False Then
                AddText "Communicator has started in Safe Mode (Load Time: " & FormatTimeElapsed(sLoadTime) & ")", TxtUnknown, True
            End If
        End If
        
        
        AddConsoleText "Load Time: " & CStr(sLoadTime)
        
        If Not bStartup And modLoadProgram.frmSplash_Loaded Then
            frmSplash.Refresh
            Unload frmSplash
            AddConsoleText "Unloaded Splash Window", , , True
        End If
        
        SetFocus2 frmMain.txtOut
        frmMain.tmrHost.Enabled = True
        frmMain.tmrLP_Timer
    
EndOfSub:
        If bStealth Then
            StealthMode = True
        End If
    
    End If
    
ExitLP:
    bLoading = False
    SetSplashProgress 100
    
    Exit Sub
frmMain_LoadEH:
    
    sTxt = Err.Description
    If LenB(sTxt) Then
        sTxt = sTxt & vbNewLine & _
               "Error No.:" & Err.Number & vbNewLine & _
               "DLL Error: " & Err.LastDllError
    Else
        sTxt = "Error No.:" & Err.Number & vbNewLine & _
               "DLL Error: " & Err.LastDllError
    End If
    
    AddConsoleText "Error Loading Main Window: " & sTxt
    
    LP_MsgBoxEx "Error Loading Main Window. What do you think you're playing at?" & vbNewLine & _
        "Here's the error: " & sTxt & vbNewLine & "Contact MicRobSoft or whoever gave you this program. Preferably with a large stick", _
        "Yo yo, you need some more files on your system. Probably. Don't take my word for it though", _
        vbCritical, "Error Loading"
    
    
    ExitProgram
    ExitProcess e_Load_Error
    
End Sub

Public Sub SetSplashInfo(ByVal Inf As String)

If modLoadProgram.frmSplash_Loaded And Not bStartup Then 'And Not bStealth Then
    frmSplash.pSetInfo Inf
End If

End Sub
Public Sub SetSplashProgress(TotalProgress As Integer)

If frmSplash_Loaded Then
    With frmSplash.progLoad
        
        If modLoadProgram.bAprilFools Then
            On Error Resume Next
            .Value = Rnd() * .Max
        Else
            On Error Resume Next
            .Value = TotalProgress
        End If
        
        .Refresh
    End With
End If

End Sub

Private Function IsIDE() As Boolean
On Error GoTo EH
IsIDE = False
Debug.Assert 1 / 0
Exit Function
EH:
IsIDE = True
End Function

Private Function DllsExist() As String
Const Sys32Path As String = "c:\windows\system32\"

If FileExists(Sys32Path & "msvbvm60.dll") = False Then
    DllsExist = "msvbvm60.dll"
ElseIf FileExists(Sys32Path & "oleaut32.dll") = False Then
    DllsExist = "oleaut32.dll"
ElseIf FileExists(Sys32Path & "olepro32.dll") = False Then
    DllsExist = "olepro32.dll"
ElseIf FileExists(Sys32Path & "asycfilt.dll") = False Then
    DllsExist = "asycfilt.dll"
ElseIf FileExists(Sys32Path & "stdole2.tlb") = False Then
    DllsExist = "stdole2.tlb"
ElseIf FileExists(Sys32Path & "COMCAT.dll") = False Then
    DllsExist = "COMCAT.dll"
ElseIf FileExists(Sys32Path & "wininet.dll") = False Then
    DllsExist = "wininet.dll"
ElseIf FileExists(Sys32Path & "mswinsck.ocx") = False Then
    DllsExist = "MSWINSCK.ocx"
ElseIf FileExists(Sys32Path & "RICHTX32.ocx") = False Then
    DllsExist = "RICHTX32.ocx"
ElseIf FileExists(Sys32Path & "MSCOMCTL.ocx") = False Then
    DllsExist = "MSCOMCTL.ocx"
ElseIf FileExists(Sys32Path & "COMDLG32.ocx") = False Then
    DllsExist = "COMDLG32.ocx"
ElseIf FileExists(Sys32Path & "msscript.ocx") = False Then
    DllsExist = "msscript.ocx"
ElseIf FileExists(Sys32Path & "hnetcfg.dll") = False Then
    DllsExist = "hnetcfg.dll"
Else
    DllsExist = vbNullString
End If

End Function

'Private Function AquireIP() As String
'Dim IPText As String, rIP As String
'Dim brIP As Boolean
'
''If LenB(lIP) = 0 Then lIP = frmMain.SckLC.LocalIP
''If LenB(rIP) = 0 Then rIP = frmMain.GetIP()
'
'rIP = modWinsock.RemoteIP
'brIP = Not CBool(InStr(1, rIP, "Error:", vbTextCompare))
'
'If Len(rIP) > Len("xxx.xxx.xxx.xxx") Then brIP = False
'If brIP = False Then rIP = vbNullString
'
'IPText = "Local" & IIf(brIP, " and Remote", vbNullString)
'
'IPText = IPText & " IP" & IIf(brIP, "s", vbNullString) & " aquired"
'
'IPText = IPText & vbNewLine & Space$(modConsole.IndentLevel) & "Internal IP: " & lIP & _
'                IIf(brIP, vbNewLine & Space$(modConsole.IndentLevel) & _
'                "External IP: " & rIP, vbNullString)
'
'AquireIP = IPText
'
'End Function

Public Sub ExitProgram()

If modFTP.bFTP_Doing Then
    modFTP.bCancelFTP = True
End If

modVars.Closing = True


If modLoadProgram.frmSplash_Loaded Then Unload frmSplash


If frmMain_Loaded Then
    frmMain.CleanUp True
    Unload frmMain
    
    If modVars.Closing Then 'might have been canceled
        
        Close 'close all open files
        
        If frmMain_Loaded Then
            'frmMain.Form_Terminate
            'Unload frmMain
            Set frmMain = Nothing 'calls frmMain_Terminate
        End If
    End If
End If

End Sub

Private Function CheckPrevInstance(CmdLn As String) As Boolean
Dim SystrayHandle As Long, CmdHandle As Long ', frmMain_Handle As Long

modLoadProgram.SetSplashInfo "Checking for Other Communicator..."


If App.PrevInstance Then
    If InStr(1, CmdLn, "/killold", vbTextCompare) = 0 Then
        If Not modVars.bStealth Then
            Dim Ans As VbMsgBoxResult
            
            
            If InStr(1, CmdLn, "/forceopen", vbTextCompare) Then
                Ans = vbNo
            ElseIf InStr(1, Command$(), "/instanceprompt", vbTextCompare) Then
                Ans = LP_MsgBoxEx("Another Communicator is Already Running." & vbNewLine & _
                               "Switch to It?", "Another Communicator is active on your system, do you want to go to it, instead of loading this?", _
                               vbYesNo + vbQuestion, "Communicator")
            Else
                Ans = vbYes
            End If
            
            
            If Ans = vbYes Then
                
                modLoadProgram.SetSplashInfo "Showing Other Communicator..."
                
                CheckPrevInstance = True
                
                'frmMain_Handle = FindWindow(vbNullString, "Communicator -*")
                
                SystrayHandle = FindWindow(vbNullString, Systray_Caption)
                If SystrayHandle = 0 Then
                    CheckPrevInstance = False
                    
                ElseIf WindowResponsive(SystrayHandle) Then
                    If SystrayHandle Then
                        CmdHandle = FindWindowEx(SystrayHandle, 0&, vbNullString, "Show")
                        
                        If CmdHandle Then
                            Pause 500
                            
                            
                            SendMessageByLong CmdHandle, WM_LBUTTONDOWN, 0&, 0&
                            SendMessageByLong CmdHandle, WM_LBUTTONUP, 0&, 0&
                            
                            CheckPrevInstance = True
                            Exit Function
                        End If
                    End If
                    
                Else
                    Ans = LP_MsgBoxEx("Another Communicator is running, but is not responding" & vbNewLine & _
                                   "Continue to load this instance of Communicator?", _
                                   "The other Communicator already running has frozen/hung. Do you want to load this instance of Communicator as well?", _
                                   vbYesNo Or vbQuestion, "Communicator")
                    
                    
                    If Ans = vbNo Then
                        CheckPrevInstance = True
                    Else
                        CheckPrevInstance = False
                    End If
                    
                    
                End If
                
                'MsgBox "The other program is in the system tray," & vbNewLine & _
                    "near the clock", vbInformation, "Communicator"
                
                
            End If
        End If
    End If
End If

End Function

Private Function WindowResponsive(hWnd As Long) As Boolean
Dim lRet As Long, lResult As Long

lRet = SendMessageTimeout(hWnd, WM_NULL, 0&, 0&, SMTO_ABORTIFHUNG Or SMTO_BLOCK, 200, lResult)

WindowResponsive = CBool(lRet)

'WindowResponsive = Not CBool(IsHungAppWindow(hWnd))

End Function

Private Function LP_MsgBoxEx(ByVal Prompt As String, _
                ByVal sContent As String, _
                Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                Optional ByVal Title As String = vbNullString) As VbMsgBoxResult

If frmSplash_Loaded Then
    
    LP_MsgBoxEx = MsgBoxEx(Prompt, sContent, Buttons, Title, _
        frmSplash.ScaleX(frmSplash.Left, frmSplash.ScaleMode, vbPixels), _
        frmSplash.ScaleY(frmSplash.Top, frmSplash.ScaleMode, vbPixels), _
        frmSplash.Icon, , frmSplash.hWnd)
    
Else
    LP_MsgBoxEx = MsgBoxEx(Prompt, sContent, Buttons, Title)
End If

End Function

Private Function CommandLinePresent(sSwitch As String, sCommandLine As String) As Boolean
Dim i As Integer
Dim S As String

i = InStr(1, sCommandLine, "/" & sSwitch, vbTextCompare)

If i Then
    S = Mid$(sCommandLine, (1 + i + Len(sSwitch)), 1)
    If LenB(S) > 0 Then
        If S <> vbSpace And S <> "/" Then
            CommandLinePresent = False
        Else
            CommandLinePresent = True
        End If
    Else
        CommandLinePresent = True
    End If
Else
    CommandLinePresent = False
End If


End Function
'Private Function CommandLinePresent_Switch(sSwitch As String, sCommandLine As String) As VbTriState
'Dim i As Integer
'Dim S As String
'
'On Error GoTo EH
'i = InStr(1, sCommandLine, "/" & sSwitch, vbTextCompare)
'
'If i Then
'    S = Mid$(sCommandLine, (1 + i + Len(sSwitch)), 1)
'    If LenB(S) > 0 Then
'        If S = vbSpace Then
'            If Mid$(sCommandLine, (2 + i + Len(sSwitch)), 2) = "1" Then
'                CommandLinePresent_Switch = vbTrue
'            Else
'                CommandLinePresent_Switch = vbFalse
'            End If
'        ElseIf S = "/" Then
'            CommandLinePresent_Switch = vbUseDefault
'        Else
'            CommandLinePresent_Switch = vbUseDefault
'        End If
'    Else
'        CommandLinePresent_Switch = vbUseDefault
'    End If
'Else
'    CommandLinePresent_Switch = vbUseDefault
'End If
'
'Exit Function
'EH:
'CommandLinePresent_Switch = False
'End Function

Private Sub InitCommonControls()

If bSafeMode Then
    If modDisplay.VisualStyle() Then
        modDisplay.setVisualStyle False
    End If
End If

Call api_InitCommonControls

End Sub
