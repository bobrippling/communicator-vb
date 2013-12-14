Attribute VB_Name = "modStickGame"
Option Explicit

'cursor blink rate
Public Declare Function GetCursorBlinkTime Lib "user32" Alias "GetCaretBlinkTime" () As Long
'Private Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long

Public Const Stick_Edit_Zoom = 0.3, Map_Ext = "map"
Public StickMapPath As String 'full path


Public Enum eWeaponTypes
    AK = 0 '102
    XM8 '=1 '103
    AUG 'etc '104
    G3 '105
    
    W1200 '106
    SPAS 'SEE AWM NOTE '107
    
    M82 '108
    AWM 'SEE USP NOTE '109
    
    M249 '110
    
    MP5 '111
    Mac10 '112
    
    DEagle '113
    USP '114 'NOTE: USP IS LEFT OUT OF RELOAD SOUNDS - IT STEALS DEAGLE'S
    
    FlameThrower '115
    RPG '116
    Knife
    Chopper
End Enum
Public Enum eStickPerks
    pNone = 0
    pJuggernaut  '- Increased health
    pSleightOfHand '- Faster reloading
    pStoppingPower  '- Increased bullet damage
    pBombSquad  '- duh
    pConditioning  '- Faster Movement
    pSniper 'pStealth '-Almost invisible when prone''''+ESP
    pFocus '-Sniper scope for all weapons
    pMartyrdom '-drop live nade on death
    'pResistance 'resistance to flash + fire
    pSteadyAim '- steady aim
    pMechanic '-Rapid fire
    pDeepImpact
    pZombie
    pSpy '-take someone else's name + colour
End Enum

Public Enum eFireModes
    'Default = 0
    Auto = 0
    Semi_3
    Semi_2
    Single_Shot
End Enum

Public Enum eNadeTypes
    nFrag = 0
    nFlash
    'nSmoke
    nTime
    nGravity
    nEMP
End Enum


'cvars
'Public cg_Smoke As Boolean
'Public cg_Explosions As Boolean
Public cg_Blood As Boolean
Public cg_Casing As Boolean
'Public cg_RPGFlame As Boolean
Public cg_AutoCamera As Boolean
Public cg_DrawFPS As Boolean
Public cg_LaserSight As Boolean
Public cg_DeadSticks As Boolean
Public cg_Magazines As Boolean
Public cg_Sparks As Boolean
Public cg_BGColour As Long
Public cg_SimpleStaticWeapons As Boolean
Public cg_WallMarks As Boolean
Public cg_HolsteredWeap As Boolean
Public cg_Smoke As Boolean
Public cg_ShowBulletTrails As Boolean
Public cg_ExSmoke As Boolean

Public cl_Subclass As Boolean
Public cl_DamageTick As Boolean
Public cl_SpecSpeed As Single '2 = normal
Public cl_MiddleMineDrop As Boolean
Public cl_SniperScope As Boolean
Public cl_ToggleCrouch As Boolean
Public cl_StickBotChat As Boolean

Public cg_DisplayMode As Long '=vbNotSrcCopy or vbSrcCopy
Public Const cg_DisplayMode_Normal = vbSrcCopy, _
             cg_DisplayMode_Invert = vbNotSrcCopy

'server vars
Public Enum eStickGameTypes
    gDeathMatch = 0
    gElimination '= 1
    gCoOp '= 2
End Enum

Public sv_StickGameSpeed As Single
Public sv_Hardcore As Boolean
Public sv_HPBonus As Boolean
Public sv_AIMove As Boolean
Public sv_AIShoot As Boolean
Public sv_AIMine As Boolean
Public sv_AIHeliRocket As Boolean
'Public sv_2Weapons As Boolean
Private Const Default_Win_Score = 12
Public sv_WinScore As Integer
Public sv_GameType As eStickGameTypes
Public sv_AI_Rotation_Rate As Single, sv_AI_pi2LessRotRate As Single
Public sv_AIUseFlashBangs As Boolean
Public sv_BulletsThroughWalls As Boolean
Public sv_Spawn_Delay As Long 'ms
Public sv_Draw_Nade_Time As Boolean
Private Const Def_Spawn_Delay = 3000
Public sv_Damage_Factor As Single
Public sv_AllowedWeapons(0 To eWeaponTypes.Chopper) As Boolean
Public sv_SpawnWithShields As Boolean

Public Const SleightOfHandReloadDecrease = 3

Public StickServer As Boolean
Public StickServerIP As String

Public StickFormLoaded As Boolean
Public StickOptionFormLoaded As Boolean
'Public StickSettingsFormLoaded As Boolean
'Public StickTeamFormLoaded As Boolean

Public bStickEditing As Boolean

'camera and stuff
Public StickCentreX As Single, StickCentreY As Single
'old - Public Const StickCentreX = 7100, StickCentreY = 4600

'virtual width + height
Public Const StickGameWidth As Single = 50000, StickGameHeight As Single = 20000 '14000
'old - Public Const StickGameWidth = 28000, StickGameHeight = 12000


'########################################################################
'Platforms
'Public Enum ePlatformTypes
'    pNormal = 0
'    pSpikes
'End Enum

Public Type ptStickPlatform
    Left As Single
    Top As Single
    width As Single
    height As Single
    'iType As ePlatformTypes
End Type
Public Type ptStickBox
    Left As Single
    Top As Single
    width As Single
    height As Single
    bInUse As Boolean
End Type

Public ubdPlatforms As Integer '= 7
Public Platform() As ptStickPlatform

Public ubdBoxes As Integer '= 10
Public Box() As ptStickBox

Public ubdtBoxes As Integer '= 8
Public tBox() As ptStickPlatform
'########################################################################


'commands
'Public Const sKicks As String * 1 = "K" 'borrow spacegame one
Public Const sExits As String * 1 = "E"
Public Const sNewMaps As String * 1 = "Z"
Public Const sMapRequests As String * 1 = "Q", sMapNames As String * 1 = "I"


'camera pos
Public Type ptPoint
    X As Single
    Y As Single
End Type
Public cg_sCamera As ptPoint
Public cg_sZoom As Single

'fps stuff
Private Const Stick_Main_FPS = 72 'FPS the game runs at - changing this changes the speed
Public Const Stick_Ms_Delay = 1000 / Stick_Main_FPS '14
Public Const Stick_Required_FPS = 120 'Any FPS - adjustments are made to meet this - essentially Max_FPS
Public Const Stick_Ms_Required_Delay = 1000 / Stick_Required_FPS
Public Const Frame_Const As Long = 34 'DO NOT CHANGE - DELAYS RELY ON THIS
Public StickElapsedTime As Long
Public StickTimeFactor As Single

Public Const AI_Delay As Long = 500

'types
Public Stick() As ptStick
Public NumSticks As Integer

'#######################
'player settings
Public cl_StartWeapon1 As eWeaponTypes
Public cl_StartWeapon2 As eWeaponTypes
Public cl_StartPerk As eStickPerks
'#######################


Public Type ptStick
    ID As Integer
    LastPacket As Long
    LastPacketID As Long
    
    LastSlowPacketID As Long
    
    Name As String * 15
    Facing As Single
    ActualFacing As Single
    Heading As Single
    X As Single
    Y As Single
    
    'JumpStartY As Single
    
    LastBullet As Long
    LastNade As Long
    LastMine As Long
    NadeStart As Long 'for clientside only, delay in nade throwing
    LastLoudBullet As Long
    
    Health As Integer
    state As Integer
    
    Shield As Single
    LastShieldHitTime As Long
    ShieldCharging As Boolean
    
    SockAddr As ptSockAddr
    
    colour As Long
    
    Speed As Single
    
    LegWidth As Integer
    LegBigger As Boolean
    LastMuzzleFlash As Long
    
    'StartJumpTime As Long
    
    GunPoint As ptPoint
    CasingPoint As ptPoint
    
    WeaponType As eWeaponTypes
    bLightSaber As Boolean
    
    bOnSurface As Boolean
    bTouchedSurface As Boolean
    LastGravity As Long
    iCurrentPlatform As Integer
    
    BulletsFired As Integer
    BulletsFired2 As Integer 'for AUG - burst fire
    ReloadStart As Long
    LastRoundIn As Long 'for shotgun etc
    Burst_Bullets As Integer 'for weapon fire mode
    Burst_Delay As Long      '"
    
    
    RecoilLeft As Boolean
    
    '#####################
    IsBot As Boolean
    'Temp (?)
    'AIDir As Boolean
    LastAI As Long
    AINadeDelay As Long
    AILastMineAttempt As Long
    AICurrentTarget As Integer
    AI_AngleToTarget As Single
    AILastFacingAdjust As Long
    LastFlashBang As Long 'flashbang
    AIPickedNade As Boolean
    'AIWantToFace As Single
    AI_Targets_Seen As String
    '#####################
    
    iKills As Integer
    iDeaths As Integer
    iKillsInARow As Integer
    LastSpawnTime As Long
    lDeathTime As Long
    sgTimeZone As Single
    
    bHadMag As Boolean 'added mag for this stick's reload? (foreign sticks)
    CurrentWeapons(1 To 2) As eWeaponTypes
    LastWeaponSwitch As Long 'delay for switching weapons
    'PrevWeapon As eWeaponTypes
    
    Team As eTeams
    bAlive As Boolean
    
    bSilenced As Boolean
    bTyping As Boolean
    bFlashed As Boolean
    bOnFire As Boolean
    iNadeType As eNadeTypes
    
    LastFlameTouch As Long
    LastFlameTouchOwnerID As Integer
    LastFlameDamage As Long
    bFlameIsFromTag As Boolean
    
    Perk As eStickPerks
    MaskID As Integer
    
    'chopper stuff
    TailRotorFacing As Single
    RotorWidth As Integer
    RotorDir As Boolean
    ChopperFacingAmount As Single 'gradual increase of facing
    
    
    LastChatMsg As Long
    curChatMsg As String
End Type

Public Enum eStickStates
    STICK_NONE = 0
    STICK_RIGHT = 1
    STICK_LEFT = 2
    
    STICK_JUMP = 4
    STICK_CROUCH = 8
    STICK_PRONE = 16
    
    STICK_FIRE = 32
    STICK_RELOAD = 64
    STICK_NADE = 128
    STICK_MINE = 256
End Enum

Private kWeaponName(0 To eWeaponTypes.Chopper) As String


'line + circle stuff
Private Declare Function MoveToEx Lib "gdi32" ( _
    ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, _
    ByRef lpPoint As Any) As Long 'POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" ( _
    ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function apiRectangle Lib "gdi32" Alias "Rectangle" ( _
    ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function Ellipse Lib "gdi32" ( _
    ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function Pie Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal nLeftRect As Long, ByVal nTopRect As Long, _
    ByVal nRightRect As Long, ByVal nBottomRect As Long, _
    ByVal nXRadial1 As Long, ByVal nYRadial1 As Long, _
    ByVal nXRadial2 As Long, ByVal nYRadial2 As Long) As Long


'####################################################################################
Public Function MakeSquareNumber() As Integer
Dim i As Integer

i = Round(Rnd() * 101)
MakeSquareNumber = i * i

End Function

Public Function IsSquare(sN As Single) As Boolean
Dim sRoot As Single

sRoot = Sqr(sN)

IsSquare = (Int(sRoot) = sRoot)

End Function
'####################################################################################

Public Function GetStickMapPath() As String
Dim sPath As String

sPath = AppPath() & "Stick Maps"

If FileExists(sPath, vbDirectory) = False Then
    On Error Resume Next
    MkDir sPath
End If

GetStickMapPath = sPath & "\"
End Function

Public Sub Stick_FormLoad(Frm As Form, Optional bReverse As Boolean = False)

FormLoad Frm, bReverse, False, False, True

If bReverse Then
    SetFocus2 frmStickGame
End If
End Sub

Public Sub StickMotion(ByRef sngX As Single, ByRef sngY As Single, ByVal sngSpeed As Single, ByVal sngHeading As Single)
Dim f As Single

''adjust for slower computers
'F = sticktimefactor

sngX = sngX + sngSpeed * Sine(sngHeading) * StickTimeFactor
sngY = sngY - sngSpeed * CoSine(sngHeading) * StickTimeFactor
                                          

End Sub

'Public Sub StickAddVectors(sngMag1 As Single, sngDir1 As Single, sngMag2 As Single, sngDir2 As Single, _
'Optional ByRef sngMagResult As Single, Optional ByRef sngDirResult As Single)
'
'Dim sngXComp As Single
'Dim sngYComp As Single
'
''Determine the components of the resultant vector
'sngXComp = (sngMag1 * sine(sngDir1) + sngMag2 * sine(sngDir2))
'
'sngYComp = (sngMag1 * cosine(sngDir1) + sngMag2 * cosine(sngDir2))
'
'
''Determine the resultant magnitude
'sngMagResult = Sqr(sngXComp ^ 2 + sngYComp ^ 2)
'
''Calculate the resultant direction, and adjust for atngent by adding Pi if necessary
'If sngYComp > 0 Then
'    sngDirResult = atn(sngXComp / sngYComp)
'ElseIf sngYComp < 0 Then
'    sngDirResult = atn(sngXComp / sngYComp) + pi
'End If
'
'End Sub

Public Sub HostStickGame(ByVal IPToDist As String, Optional ByVal MapPath As String)
'                                                   ^ needs to be optional for Dev Cmd Host

Dim Str As String, sDef_Map_Path As String
Dim bDef_Map As Boolean

Str = eCommands.LobbyCmd & eLobbyCmds.Add & frmMain.LastName & "#" & IPToDist & "S"

If Server Then
    'modMessaging.DistributeMsg Str, -1
    Call DataArrival(Str) 'this'll distribute it
Else
    SendData Str
End If


modStickGame.StickServer = True
modStickGame.StickServerIP = IPToDist
modStickGame.bStickEditing = False

sDef_Map_Path = modStickGame.GetStickMapPath() & "Default." & Map_Ext

If LenB(MapPath) = 0 Then
    MapPath = sDef_Map_Path
End If


bDef_Map = (MapPath = sDef_Map_Path)


If bDef_Map Or FileExists(MapPath) Then
    modStickGame.StickMapPath = MapPath
    
    'DoEvents 'refresh screen
    
    On Error Resume Next
    Load frmStickGame
Else
    AddText "Error - Map Doesn't Exist", TxtError, True
End If

End Sub

Public Sub JoinStickGame(ByVal IP As String)

modStickGame.StickServer = False
modStickGame.StickServerIP = IP
modStickGame.bStickEditing = False
modStickGame.StickMapPath = vbNullString

On Error Resume Next
Load frmStickGame

End Sub

Public Sub EditStickGame(sMapPath As String)

modStickGame.StickServer = False
modStickGame.StickServerIP = vbNullString
modStickGame.bStickEditing = True
If LenB(sMapPath) = 0 Then
    modStickGame.StickMapPath = modStickGame.GetStickMapPath() & "Default." & Map_Ext
Else
    modStickGame.StickMapPath = sMapPath
End If

On Error Resume Next
Load frmStickGame 'takes care of frmStickEdit

End Sub

Public Sub InitVars()

modStickGame.cl_Subclass = True
modStickGame.cl_SpecSpeed = 1
modStickGame.cl_DamageTick = True
modStickGame.cl_MiddleMineDrop = True

modStickGame.cg_Smoke = False
modStickGame.cg_ExSmoke = True
modStickGame.cg_WallMarks = True
modStickGame.cg_HolsteredWeap = True
modStickGame.cg_ShowBulletTrails = True
modStickGame.cg_BGColour = &HFF8080
modStickGame.cg_DeadSticks = True
modStickGame.cg_Magazines = True
modStickGame.cg_Sparks = True
modStickGame.cg_Blood = True
modStickGame.cg_Casing = True
modStickGame.cg_sZoom = 1
modStickGame.cg_DisplayMode = cg_DisplayMode_Normal

modStickGame.sv_WinScore = Default_Win_Score
modStickGame.sv_BulletsThroughWalls = True
modStickGame.sv_Spawn_Delay = Def_Spawn_Delay
modStickGame.sv_AIMove = True
modStickGame.sv_AIShoot = True
modStickGame.sv_AIUseFlashBangs = True
modStickGame.sv_AIMine = True
modStickGame.sv_Draw_Nade_Time = True
modStickGame.sv_Damage_Factor = 1
modStickGame.sv_StickGameSpeed = 1
modStickGame.sv_SpawnWithShields = False

modStickGame.cl_SniperScope = True
modStickGame.cl_StartWeapon1 = AK
modStickGame.cl_StartWeapon2 = USP
modStickGame.cl_StartPerk = pNone
modStickGame.cl_ToggleCrouch = True
modStickGame.cl_StickBotChat = True


modAudio.bDXSoundEnabled = True

UpdateBotRotationRate Pi / 20

MakeWeaponNameArray
End Sub

Public Sub UpdateBotRotationRate(sRate As Single)

sv_AI_Rotation_Rate = sRate
sv_AI_pi2LessRotRate = Pi2 - sv_AI_Rotation_Rate 'fixed

End Sub

Public Function FindAngle_Actual(intX1 As Single, intY1 As Single, intX2 As Single, intY2 As Single) As Single

Dim sngXComp As Single
Dim sngYComp As Single

'Find the angle between the 2 coords
sngXComp = intX2 - intX1
sngYComp = intY1 - intY2

If sngYComp > 0 Then
    FindAngle_Actual = Atn(sngXComp / sngYComp)
    
ElseIf sngYComp < 0 Then
    FindAngle_Actual = Atn(sngXComp / sngYComp) + Pi
    
ElseIf sngXComp > 0 Then
    FindAngle_Actual = piD2
    'Debug.Print "pid2"
Else
    FindAngle_Actual = pi3D2
    'Debug.Print "pi3d2"
End If

End Function

'#################################################################################
'DRAWING##########################################################################
'#################################################################################

Public Sub PrintStickText(ByVal Text As String, X As Single, Y As Single, colour As Long)

'frmStickGame.ForeColor = Colour
'frmStickGame.CurrentX = x
'frmStickGame.CurrentY = y
'
'frmStickGame.Print Text

Dim lhDC As Long

lhDC = frmStickGame.picMain.hDC

Call SetBkColor(lhDC, 0)
Call SetTextColor(lhDC, colour)
Call TextOut(lhDC, _
    frmStickGame.ScaleX(X * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
    Text, Len(Text))

End Sub

Public Sub PrintStickFormText(Text As String, X As Single, Y As Single, colour As Long)

'With frmStickGame
'    .CurrentX = x
'    .CurrentY = y
'End With
'
'frmStickGame.Print Str

Dim lhDC As Long

lhDC = frmStickGame.picMain.hDC

Call SetBkColor(lhDC, 0)
Call SetTextColor(lhDC, colour)
Call TextOut(lhDC, _
    frmStickGame.ScaleX(X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y, vbTwips, vbPixels), _
    Text, Len(Text))

End Sub

Public Sub sCircle(X As Single, Y As Single, Radius As Single, colour As Long)

'frmStickGame.picMain.Circle (X * cg_sZoom - cg_sCamera.X, _
                Y * cg_sZoom - cg_sCamera.Y), _
                Radius * cg_sZoom, Colour

frmStickGame.picMain.ForeColor = colour

'convert centre to points, etc
Ellipse frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX((X - Radius) * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY((Y - Radius) * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
    frmStickGame.ScaleX((X + Radius) * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY((Y + Radius) * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)

End Sub

Public Sub sCircleAspect(X As Single, Y As Single, Radius As Single, colour As Long, sAspect As Single)


frmStickGame.picMain.Circle (X * cg_sZoom - cg_sCamera.X, _
                Y * cg_sZoom - cg_sCamera.Y), _
                Radius * cg_sZoom, colour, , , sAspect

End Sub

'Public Sub sCircleSE(X As Single, Y As Single, Radius As Single, Colour As Long, _
    sgAngle1 As Single, sgAngle2 As Single)
Public Sub sCircleSE(X As Single, Y As Single, Radius As Single, colour As Long, _
    sgStart As Single, sgEnd As Single)

frmStickGame.picMain.Circle (X * cg_sZoom - cg_sCamera.X, _
                Y * cg_sZoom - cg_sCamera.Y), _
                Radius * cg_sZoom, colour, sgStart, sgEnd

'Dim R1 As Long, R2 As Long, R3 As Long, R4 As Long
'Dim Ra1 As Long, Ra2 As Long, Ra3 As Long, Ra4 As Long
'Dim Wd2 As Long, Hd2 As Long
'Dim Start1 As Long, Start2 As Long
'
'Wd2 = Radius * cg_sZoom \ 2
'Hd2 = Wd2 'Radius \ 2
'
'R1 = X - Wd2: R2 = Y - Hd2
'R3 = X + Wd2: R4 = Y + Hd2
'
'
'Start1 = R1 + Wd2
'Start2 = R2 + Hd2
'
'
'Ra1 = Start1 + Sin(sgAngle1) '(R3 - R1) * Sin(Angle1)
'Ra2 = Start2 - Cos(sgAngle1) '(R4 - R2) * Cos(Angle1)
'
'Ra3 = Start1 + Sin(sgAngle2) '(R3 - R1) * Sin(Angle2)
'Ra4 = Start2 - Cos(sgAngle2) '(R4 - R2) * Cos(Angle2)
'
'frmStickGame.ForeColor = Colour
'Pie frmStickGame.hDC, frmStickGame.ScaleX(R1 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), frmStickGame.ScaleY(R2 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
'                      frmStickGame.ScaleX(R3 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), frmStickGame.ScaleY(R4 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
'                      frmStickGame.ScaleX(Ra1 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), frmStickGame.ScaleY(Ra2 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
'                      frmStickGame.ScaleX(Ra3 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), frmStickGame.ScaleY(Ra4 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)

End Sub

Public Sub sLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) ', Optional Colour As Long = -1)

'If Colour = -1 Then
'    Colour = frmStickGame.ForeColor
'End If

'old
'frmStickGame.picMain.Line (x1 * cg_sZoom - cg_sCamera.X, _
              y1 * cg_sZoom - cg_sCamera.Y) _
            -(X2 * cg_sZoom - cg_sCamera.X, _
              Y2 * cg_sZoom - cg_sCamera.Y), _
              Colour


'frmStickGame.picMain.ForeColor = IIf(Colour = -1, frmStickGame.ForeColor, Colour)

MoveToEx frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX(X1 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y1 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), ByVal 0&

LineTo frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX(X2 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y2 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)


End Sub

Public Sub sLine_SetBegin(X1 As Single, Y1 As Single)

MoveToEx frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX(X1 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y1 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), ByVal 0&

End Sub
Public Sub sLine_FromLast(X2 As Single, Y2 As Single)

LineTo frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX(X2 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y2 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)

End Sub

Public Sub sBox(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, colour As Long)

'frmStickGame.picMain.Line (X1 * cg_sZoom - cg_sCamera.X, _
              Y1 * cg_sZoom - cg_sCamera.Y) _
            -(X2 * cg_sZoom - cg_sCamera.X, _
              Y2 * cg_sZoom - cg_sCamera.Y), _
              Colour, B

frmStickGame.picMain.ForeColor = colour
apiRectangle frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX(X1 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y1 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
    frmStickGame.ScaleX(X2 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y2 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)

End Sub

Public Sub sBoxFilled(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, colour As Long)


frmStickGame.picMain.ForeColor = colour
frmStickGame.picMain.FillStyle = vbFSSolid
frmStickGame.picMain.FillColor = colour
apiRectangle frmStickGame.picMain.hDC, _
    frmStickGame.ScaleX(X1 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y1 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels), _
    frmStickGame.ScaleX(X2 * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y2 * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)

frmStickGame.picMain.FillStyle = vbFSTransparent
'frmStickGame.picMain.Line (X1 * cg_sZoom - cg_sCamera.X, _
              Y1 * cg_sZoom - cg_sCamera.Y) _
            -(X2 * cg_sZoom - cg_sCamera.X, _
              Y2 * cg_sZoom - cg_sCamera.Y), _
              Colour, BF

End Sub

Public Sub sPoly(Pts() As PointAPI, lFillCol As Long)
Dim j As Integer

For j = LBound(Pts) To UBound(Pts)
    Pts(j).X = frmStickGame.ScaleX(Pts(j).X * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels)
    Pts(j).Y = frmStickGame.ScaleY(Pts(j).Y * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)
Next j

modGDI.DrawPoly Pts, frmStickGame.picMain.hDC, lFillCol

End Sub

Public Sub sPoly_NoOutline(Pts() As PointAPI, lFillCol As Long)
Dim j As Integer

For j = LBound(Pts) To UBound(Pts)
    Pts(j).X = frmStickGame.ScaleX(Pts(j).X * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels)
    Pts(j).Y = frmStickGame.ScaleY(Pts(j).Y * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)
Next j

modGDI.DrawPoly_NoOutline Pts, frmStickGame.picMain.hDC, lFillCol

End Sub

Public Sub sHatchCircle(ByVal X As Single, ByVal Y As Single, lFillCol As Long, ByVal iSize As Integer)

modGDI.HatchCircle frmStickGame.picMain.hDC, lFillCol, iSize, _
    frmStickGame.ScaleX(X * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels), _
    frmStickGame.ScaleY(Y * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)

End Sub

Public Function GetWeaponName(vWeapon As eWeaponTypes) As String

GetWeaponName = kWeaponName(vWeapon)

End Function

Public Function WeaponNameToInt(sWeaponName As String) As eWeaponTypes
Dim i As eWeaponTypes

For i = 0 To eWeaponTypes.Chopper
    If GetWeaponName(i) = sWeaponName Then
        WeaponNameToInt = i
        Exit Function
    End If
Next i

WeaponNameToInt = -1

End Function

Public Sub MakeWeaponNameArray()
Dim i As Integer

For i = 0 To eWeaponTypes.Chopper
    If i = XM8 Then
        kWeaponName(i) = "XM-8" 'FN-XM8"
    ElseIf i = AK Then
        kWeaponName(i) = "AK-47"
    ElseIf i = M249 Then
        kWeaponName(i) = "M249 SAW"
    ElseIf i = M82 Then
        kWeaponName(i) = "M-107 Sniper"
    ElseIf i = RPG Then
        kWeaponName(i) = "RPG-7"
    ElseIf i = W1200 Then
        kWeaponName(i) = "W1200"
    ElseIf i = DEagle Then
        kWeaponName(i) = "Desert Eagle"
    ElseIf i = FlameThrower Then
        kWeaponName(i) = "FlameThrower"
    ElseIf i = Knife Then
        kWeaponName(i) = "Sword"
    ElseIf i = AUG Then
        kWeaponName(i) = "AUG"
    ElseIf i = Chopper Then
        kWeaponName(i) = "Chopper"
    ElseIf i = USP Then
        kWeaponName(i) = "USP"
    ElseIf i = AWM Then
        kWeaponName(i) = "AWM Sniper"
    ElseIf i = MP5 Then
        kWeaponName(i) = "HK MP5"
    ElseIf i = Mac10 Then
        kWeaponName(i) = "Mac 10"
    ElseIf i = SPAS Then
        kWeaponName(i) = "SPAS"
    ElseIf i = G3 Then
        kWeaponName(i) = "G3"
    Else
        kWeaponName(i) = "<NAME>"
    End If
Next i

End Sub

'Public Sub sHatchPoly(Pts() As POINTAPI, lFillCol As Long)
'Dim j As Integer
'
'For j = LBound(Pts) To UBound(Pts)
'    Pts(j).X = frmStickGame.ScaleX(Pts(j).X * cg_sZoom - cg_sCamera.X, vbTwips, vbPixels)
'    Pts(j).Y = frmStickGame.ScaleY(Pts(j).Y * cg_sZoom - cg_sCamera.Y, vbTwips, vbPixels)
'Next j
'
'modGDI.HatchPoly Pts, frmStickGame.picMain.hdc, lFillCol
'
'End Sub
'#################################################################################
'END DRAWING######################################################################
'#################################################################################

