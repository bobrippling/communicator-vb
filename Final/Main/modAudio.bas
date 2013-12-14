Attribute VB_Name = "modAudio"
Option Explicit

'sound
Private Declare Function apiPlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Private Const SND_ALIAS = &H10000        'lpszName is a string identifying the name of the system event sound to play.
Private Const SND_ALIAS_ID = &H110000    'lpszName is a string identifying the name of the predefined sound identifier to play.
Private Const SND_APPLICATION = &H80     'lpszName is a string identifying the application-specific event association sound to play.
Private Const SND_FILENAME = &H20000     'lpszName is a string identifying the filename of the .wav file to play.
Private Const SND_RESOURCE = &H4004      'lpszName is the numeric resource identifier of the sound stored in an application. hModule must be specified as that application's module handle.

Private Const SND_ASYNC = &H1            'Play the sound asynchronously -- return immediately after beginning to play the sound and have it play in the background.
Private Const SND_NODEFAULT = &H2        'If the specified sound cannot be found, terminate the function with failure instead of playing the SystemDefault sound. If this flag is not specified, the SystemDefault sound will play if the specified sound cannot be located and the function will return with success.
Private Const SND_NOSTOP = &H10          'If a sound is already playing, do not prematurely stop that sound from playing and instead return with failure. If this flag is not specified, the playing sound will be terminated and the sound specified by the function will play instead.
Private Const SND_PURGE = &H40           'Stop playback of any waveform sound. lpszName must be an empty string.
Private Const SND_NOWAIT = &H2000        'If a sound is already playing, do not wait for the currently playing sound to stop and instead return with failure.

Private Const SND_SYNC = &H0             'Play the sound synchronously -- do not return until the sound has finished playing.
Private Const SND_LOOP = &H8             'Continue looping the sound until this function is called again ordering the looped playback to stop. SND_ASYNC must also be specified.
Private Const SND_MEMORY = &H4           'lpszName is a numeric pointer refering to the memory address of the image of the waveform sound loaded into RAM.




Public Enum eSysSounds
    sys_SystemAsterisk = 0    'Asterisk
    sys_Default = 1           'Default Beep
    sys_EmptyRecycleBin = 2   'when recycle bin is emptied
    sys_SystemExclamation = 3 'when windows shows a warning
    sys_SystemExit = 4        'when Windows shuts down
    sys_Maximize = 5          'when a program is maximized
    sys_MenuCommand = 6       'when a menu item is clicked on
    sys_MenuPopup = 7         'when a (sub)menu pops up
    sys_Minimize = 8          'when a program is minimized to taskbar
    sys_MailBeep = 9          'when email is received
    sys_Open = 10             'when a program is opened
    sys_SystemHand = 11       'when a critical stop occurs
    sys_AppGPFault = 12       'when a program causes an error
    sys_SystemQuestion = 13   'when a system question occurs
    sys_RestoreDown = 14      'when a program is restored to normal size
    sys_RestoreUp = 15        'when a program is restored to normal size from taskbar
    sys_SystemStart = 16      'when Windows starts up
    sys_Close = 17            'when program is closed
    sys_Ringout = 18          'when (fax) call is made outbound and the line is ringing
    sys_RingIn = 19           'incoming (fax) call
End Enum

Private BugReportPath As String, HurghPath As String

Public bDXSoundEnabled As Boolean
Public bDXSoundInited As Boolean

Public WeaponPath(0 To eWeaponTypes.Chopper) As String
Public ReloadPath(0 To eWeaponTypes.RPG) As String
Public NadeExplosionPath As String, NadeBouncePath As String, NadeThrowPath As String

Public NadeBGPath As String
Public RifleBGPath As String ', LastBGRifle As Long
Public SilencedPath As String, Silenced2Path As String, MedKitPath As String

Public DeathNoisePath(1 To 3) As String
Public RoundStartPath As String, ToastyPath As String
Public RicochetPath(1 To 7) As String
Public LightSaberPath(1 To 3) As String
Public WeaponPickupPath As String, LandSound As String, WilhelmPath As String
Public TickPath As String

Public Const iStickSoundBase As Long = 31 '24 '25 '23 '21
Public Const StickTickSoundIndex As Long = iStickSoundBase + 27

Public Function PlayHurgh() As Long
PlayHurgh = PlayFileNameSound(HurghPath)
End Function

Public Function PlayBugReport() As Long
PlayBugReport = PlayFileNameSound(BugReportPath)
End Function

'##############################################################################################################

'Public Sub PlayWeaponSound(vWeapon As eWeaponTypes, lPan as long)
'Dim i As Integer: i = CInt(vWeapon)
'
'SetDXPan i, sPan
'PlayDXSound i
'End Sub
Public Sub PlayWeaponSound_Panned(vWeapon As eWeaponTypes, lPan As Long)
PlayDXSound_Panned CInt(vWeapon), lPan
End Sub
Public Sub PlayReloadSound(vWeapon As eWeaponTypes) 'no need for panning
PlayDXSound CInt(vWeapon + eWeaponTypes.Chopper + 1)
End Sub
Public Sub StopWeaponReloadSound(vWeapon As eWeaponTypes) 'no need for panning
If bDXSoundEnabled Then
    On Error Resume Next
    modDXSound.StopSound CInt(vWeapon + eWeaponTypes.Chopper + 1)
End If
End Sub
Public Function SoundPlaying(vWeapon As eWeaponTypes) As Boolean
If bDXSoundEnabled Then
    On Error Resume Next
    SoundPlaying = (modDXSound.GetSoundState(CInt(vWeapon)) = s_Playing)
End If
End Function

Public Sub PlayNadeExplosion(lPan As Long)
Const i As Integer = iStickSoundBase + 1

PlayDXSound_Panned i, lPan
End Sub
Public Sub PlayNadeBounce(lPan As Long)
Const i As Integer = iStickSoundBase + 2

PlayDXSound_Panned i, lPan
End Sub
Public Sub PlayNadeThrow(lPan As Long)
Const i As Integer = iStickSoundBase + 3

PlayDXSound_Panned i, lPan
End Sub

Public Sub PlayBackGroundNade(lPan As Long)
Const i As Integer = iStickSoundBase + 5

PlayDXSound_Panned i, lPan
End Sub

Public Sub PlaySilencedSound(lPan As Long, iSound As Integer)
'                                           ^ 0 or 1
Const i As Integer = iStickSoundBase + 6
'################################
'THIS IS 135 IN THE RESOURCE FILE
'################################

PlayDXSound_Panned i + iSound, lPan
End Sub

'Public Sub PlayBackGroundShot(Pan As Long)
'If LastBGRifle + 1500 < GetTickCount() Then
'    'PlayFileNameFast RifleBGPath
'
'    SetDXPan iStickSoundBase + 5, Pan '26
'    PlayDXSound iStickSoundBase + 5
'
'    LastBGRifle = GetTickCount()
'End If
'End Sub

Public Sub PlayMedKit() 'no need for pan
'PlayFileNameFast MedKitPath
PlayDXSound iStickSoundBase + 8
End Sub

Public Sub PlayDeathNoise()
PlayDXSound IntRand(iStickSoundBase + 9, iStickSoundBase + 11) '31
End Sub

Public Sub PlayNewRoundSound()
PlayDXSound iStickSoundBase + 12 '32
End Sub

Public Sub PlayToasty() 'no need for pan
PlayDXSound iStickSoundBase + 13
End Sub

Public Sub PlayRicochet(lPan As Long)
Const i As Integer = iStickSoundBase + 14, j As Integer = iStickSoundBase + 20

PlayDXSound_Panned IntRand(i, j), lPan
End Sub

Public Sub PlayLightSaberSound(lPan As Long)
Const i As Integer = iStickSoundBase + 21, j As Integer = iStickSoundBase + 23

PlayDXSound_Panned IntRand(i, j), lPan
End Sub

Public Sub PlayWeaponPickUpSound()
PlayDXSound iStickSoundBase + 24
End Sub

Public Sub PlayLandSound(lPan As Long)
Const i As Integer = iStickSoundBase + 25

PlayDXSound_Panned i, lPan
End Sub

Public Function PlayWilhelm() As Long
Const i As Integer = iStickSoundBase + 26

PlayDXSound i
End Function

Public Sub PlayTickSound()
PlayDXSound StickTickSoundIndex 'iStickSoundBase + 27
End Sub

'#############################################

Private Sub PlayDXSound_Panned(i As Integer, lPan As Long)
If bDXSoundEnabled Then
    On Error Resume Next
    modDXSound.SetPan i, lPan
    modDXSound.PlaySound i
End If
End Sub
Private Sub PlayDXSound(i As Integer)
If bDXSoundEnabled Then
    On Error Resume Next
    modDXSound.PlaySound i
End If
End Sub
Private Sub SetDXPan(i As Integer, Pan As Long)
If bDXSoundEnabled Then
    On Error Resume Next
    modDXSound.SetPan i, Pan
End If
End Sub

'##############################################################################################################

'Private Function PlayFileNameFast(ByVal sName As String) As Long
'
'PlayFileNameFast = apiPlaySound(sName, 0, SND_FILENAME Or SND_NODEFAULT Or SND_ASYNC)
'
'End Function

Public Function PlayFileNameSound(ByVal sName As String) As Long

PlayFileNameSound = apiPlaySound(sName, 0, SND_FILENAME Or SND_NOSTOP Or SND_NODEFAULT Or SND_ASYNC)

End Function

Public Function PlaySysSound(ByVal vName As eSysSounds) As Long

PlaySysSound = apiPlaySound(GetSysName(vName), 0, SND_ASYNC Or SND_ALIAS Or SND_NODEFAULT)

End Function

Private Function GetSysName(vName As eSysSounds) As String

Select Case vName
    Case sys_MailBeep
        GetSysName = "MailBeep"
    Case sys_SystemAsterisk
        GetSysName = "SystemAsterisk"
    Case sys_Default
        GetSysName = "Default"
    Case sys_EmptyRecycleBin
        GetSysName = "EmptyRecycleBin"
    Case sys_SystemExclamation
        GetSysName = "SystemExclamation"
    Case sys_SystemExit
        GetSysName = "SystemExit"
    Case sys_Maximize
        GetSysName = "Maximize"
    Case sys_MenuCommand
        GetSysName = "MenuCommand"
    Case sys_MenuPopup
        GetSysName = "MenuPopup"
    Case sys_Minimize
        GetSysName = "Minimize"
    Case sys_Open
        GetSysName = "Open"
    Case sys_SystemHand
        GetSysName = "SystemHand"
    Case sys_AppGPFault
        GetSysName = "AppGPFault"
    Case sys_SystemQuestion
        GetSysName = "SystemQuestion"
    Case sys_RestoreDown
        GetSysName = "RestoreDown"
    Case sys_RestoreUp
        GetSysName = "RestoreUp"
    Case sys_SystemStart
        GetSysName = "SystemStart"
    Case sys_Close
        GetSysName = "Close"
    Case sys_Ringout
        GetSysName = "Ringout"
    Case sys_RingIn
        GetSysName = "RingIn"
End Select

End Function

Public Sub StopSound()

apiPlaySound vbNullString, 0, SND_PURGE

End Sub

Public Sub InitSounds()
Dim Root As String
Dim f As Integer

Root = modSettings.GetUserSettingsPath() 'has trailing \
f = FreeFile()

On Error GoTo EH

BugReportPath = Root & "Bug.wav"
Open BugReportPath For Output As #f
    Print #f, StrConv(LoadResData(101, "WAVE"), vbUnicode);
Close #f

HurghPath = Root & "Hurgh.wav"
Open HurghPath For Output As #f
    Print #f, StrConv(LoadResData(1, "WAVE"), vbUnicode);
Close #f

Exit Sub
EH:
AddText "Sound Initialisation Error: " & Err.Description, TxtError, True
End Sub

Public Function InitStickSounds() As Boolean
Dim f As Integer, i As Integer, NextFreeRes As Integer
Dim Root As String
Const SoundExt = ".STICK", sWave = "WAVE"
Const iStart_Res = 102, iTotal_Res = 156 - iStart_Res

Root = modSettings.GetUserSettingsPath() & "Stick_SFX\"

If FileExists(Root, vbDirectory) = False Then
    'error caught above
    On Error Resume Next
    MkDir Root
    
    If Err.Number > 0 Then
        Err.Raise vbObjectError, "InitStickSounds", "Couldn't Create SoundFX Directory"
        InitStickSounds = False
        Exit Function
    End If
End If

On Error GoTo EH

f = FreeFile()


'################################################################################
'rifles, etc
NextFreeRes = iStart_Res
For i = 0 To eWeaponTypes.Chopper
    
    WeaponPath(i) = Root & modStickGame.GetWeaponName(CInt(i)) & SoundExt
    
    
    Open WeaponPath(i) For Output As #f
        Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
    Close #f
    
    NextFreeRes = NextFreeRes + 1
    UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res
Next i


'NextFreeRes = 13

'reload sounds
For i = 0 To eWeaponTypes.RPG
    
    If i = USP Then
        ReloadPath(i) = Root & modStickGame.GetWeaponName(DEagle) & SoundExt & "R"
    ElseIf i = AWM Then
        ReloadPath(i) = Root & modStickGame.GetWeaponName(M82) & SoundExt & "R"
    ElseIf i = G3 Then
        ReloadPath(i) = Root & modStickGame.GetWeaponName(AK) & SoundExt & "R"
    ElseIf i = SPAS Then
        ReloadPath(i) = Root & modStickGame.GetWeaponName(W1200) & SoundExt & "R"
    ElseIf i = Mac10 Then
        ReloadPath(i) = Root & modStickGame.GetWeaponName(XM8) & SoundExt & "R"
    Else
        ReloadPath(i) = Root & modStickGame.GetWeaponName(CInt(i)) & SoundExt & "R"
        
        
        Open ReloadPath(i) For Output As #f
            Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
        Close #f
        
        NextFreeRes = NextFreeRes + 1
        UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res
    End If
    
Next i

'NextFreeRes = 22

'################################################################################
'nades - explosion, throw, bounce, pullout

NadeExplosionPath = Root & "Nade_Boom" & SoundExt
Open NadeExplosionPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

NadeBouncePath = Root & "Nade_Bounce" & SoundExt
Open NadeBouncePath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

NadeThrowPath = Root & "Nade_Throw" & SoundExt
Open NadeThrowPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

'NextFreeRes = 25
'################################################################################
'off screen sounds
NadeBGPath = Root & "Nade_Background" & SoundExt
Open NadeBGPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res


'For i = 1 To 4
'    RifleBGPath(i) = Root & "Rifle_Background" & CStr(i) & SoundExt
'    Open RifleBGPath(i) For Output As #f
'        Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
'    Close #f
'    NextFreeRes = NextFreeRes + 1
'Next i
RifleBGPath = Root & "Rifle_Background" & SoundExt
Open RifleBGPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

'NextFreeRes = 27
'################################################################################
'misc
SilencedPath = Root & "Silenced" & SoundExt
Open SilencedPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

Silenced2Path = Root & "Silenced2" & SoundExt
Open Silenced2Path For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res


MedKitPath = Root & "Medkit" & SoundExt
Open MedKitPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res


'NextFreeRes = 29
'################################################################################
'death
For i = 1 To 3
    DeathNoisePath(i) = Root & "Death" & CStr(i) & SoundExt
    
    Open DeathNoisePath(i) For Output As #f
        Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
    Close #f
    
    NextFreeRes = NextFreeRes + 1
    UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res
Next i


'NextFreeRes = 32
'################################################################################
RoundStartPath = Root & "RoundStart" & SoundExt
Open RoundStartPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

'NextFreeRes = 33
'################################################################################
ToastyPath = Root & "Toasty" & SoundExt
Open ToastyPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

'NextFreeRes = 34
'################################################################################
For i = 1 To 7
    RicochetPath(i) = Root & "Ricochet" & CInt(i) & SoundExt
    Open RicochetPath(i) For Output As #f
        Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
    Close #f
    NextFreeRes = NextFreeRes + 1
    UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res
Next i

'NextFreeRes = 41

For i = 1 To 3
    LightSaberPath(i) = Root & "Lightsaber" & CStr(i) & SoundExt
    Open LightSaberPath(i) For Output As #f
        Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
    Close #f
    NextFreeRes = NextFreeRes + 1
    UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res
Next i



WeaponPickupPath = Root & "WeaponPickup" & SoundExt
Open WeaponPickupPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res


LandSound = Root & "Land" & SoundExt
Open LandSound For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, sWave), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res


WilhelmPath = Root & "Wilhelm" & SoundExt '"Wilhelm.wav"
Open WilhelmPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, "WAVE"), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res


'NextFreeRes = 47
'################################################################################

TickPath = Root & "Tick" & SoundExt
Open TickPath For Output As #f
    Print #f, StrConv(LoadResData(NextFreeRes, "WAVE"), vbUnicode);
Close #f
NextFreeRes = NextFreeRes + 1
UpdateStickSoundInitProg NextFreeRes, iStart_Res, iTotal_Res

InitStickSounds = True

Exit Function
EH:
Stop
Resume

Close #f
InitStickSounds = False
'Err.Raise vbObjectError, "modAudio", Err.Description
End Function

Private Sub UpdateStickSoundInitProg(ByVal iCurrent As Integer, iStart As Integer, iTotal As Integer)
Const Line_Len = 1750

With frmStickGame
    'no need to .Cls
    
    .picMain.Line (.ConnectingkX, .ConnectingkY)-( _
        .ConnectingkX + Line_Len * CLng(iCurrent - iStart) / iTotal, .ConnectingkY), vbRed
    '                                                       ^ iStart has already been subbed from it
    
    
    .BltToForm
    .Refresh
End With

End Sub
