VERSION 5.00
Begin VB.Form frmStickGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "Stick Shooter"
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   15915
   Begin VB.PictureBox picHandle 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Start!"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cboWeapon 
      Height          =   315
      Left            =   6360
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox picToasty 
      BorderStyle     =   0  'None
      Height          =   2040
      Left            =   2880
      Picture         =   "frmStickGame.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   1635
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   1640
   End
   Begin VB.PictureBox picBlank 
      Height          =   255
      Left            =   4800
      Picture         =   "frmStickGame.frx":0ED6
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4800
      Top             =   960
   End
   Begin VB.Shape shHealthPack 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   3240
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   855
      Index           =   0
      Left            =   1800
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   0
      Left            =   1800
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   360
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmStickGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Note: When Compiling, uncomment _AmbientChanged in ucButtons
#Const bTimeAdjust = True

#Const Clip_X_Camera = False
#Const Clip_Y_Camera = False

#Const Hack_All = True
#Const Hack_ForceOff = False
#Const Hack_AimBot = Not Hack_ForceOff And (Hack_All Or False)
#Const Hack_Recoil = Not Hack_ForceOff And (Hack_All Or False)
#Const Hack_Ammo = Not Hack_ForceOff And (Hack_All Or False)
#Const Hack_Shield = Not Hack_ForceOff And (Hack_All Or False)
#Const Hack_AIShield = False


'############################################################
'EDIT STUFF
'Private Const WS_THICKFRAME = &H40000
'Private Const WS_MAXIMIZEBOX = &H10000
'Private Const WS_MINIMIZEBOX = &H20000
'Private Const WS_DLGFRAME = &H400000 - no frame AT ALL
'Private Const WS_BORDER = &H800000

Private Const Edit_Width = 15135, Edit_Height = 7100 '5100
Private Const MapSep = "|"

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Const NULL_BRUSH = 5
Private Const PS_SOLID = 0
Private Const R2_NOT = 6

Private Enum ControlState
    StateNothing = 0
    StateDragging '= 1
    StateSizing '= 2
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New cRect
Private m_DragPoint As PointAPI

Private map_Changed As Boolean
'############################################################

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long


'kills
Private Enum eKillTypes
    kNormal = 0
    kHead '= 1
    kNade '= 2
    kRPG '= 3
    kKnife '= 4
    kMine '= 5
    kChoppered '= 6
    kFlame '= 7
    kBurn '= 8
    kSilenced '= 9
    kCrushed '= 10
    kFlameTag '= 11
    kLightSaber '= 12
    kBarrel '= 13
    kFall '=14
    kMartyrdom '=15
    kAirMine '=16
    kCeiling '=17
    kSpikes
End Enum
Private Enum eMagTypes
    mAK = 0
    mXM8 '= 1
    mSniper '= 2
    mPistol '= 3
    mFlameThrower '= 4
    mAUG '= 5
End Enum


Private Type ptBullet
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    'Facing As Single
    'Decay As Long
    OwnerIndex As Integer
    'Colour As Long
    Damage As Integer 'Single
    LastDiffract As Long
    
    'bShowSniperBullet As Boolean
    bSniperBullet As Boolean
    bShotgunBullet As Boolean
    bChopperBullet As Boolean
    bDEagleBullet As Boolean
    
    bHeadingChanged As Boolean
    bHadCircleBlast As Boolean
    
    bTracer As Boolean
    
    LastGravity As Long
    
    bSilenced As Boolean
    LastSmoke As Long
    
    
'    NumTrails As Long
'    Trail() As ptBulletTrail
'    LastTrailDir As Single
End Type


Private Type ptBlood
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    Decay As Long
End Type

'Private Type ptNadeTrail
'    X As Single
'    Y As Single
'    'lColour As Long
'    'iSize As Single
'    lCreation As Long
'End Type

Private Type ptNade
    X As Single
    Y As Single
    
    'Decay As Long 'Boom
    Start_Time As Long
    
    Heading As Single
    Speed As Single
    
    OwnerID As Integer
    
    IsRPG As Boolean
    
    LastSmoke As Long
    LastGravity As Long
    
    colour As Long
    iType As eNadeTypes
    
    bIsMartyrdomNade As Boolean
    
    'LastNadeTrail As Long
    'NadeTrail() As ptNadeTrail
    'NumNadeTrails As Integer
End Type

Private Type ptMine
    X As Single
    Y As Single
    OwnerID As Integer
    colour As Long
    
    ID As Integer
    
    LastGravity As Long
    bOnSurface As Boolean
    
    Speed As Single
    Heading As Single
End Type

Private Type ptCasing
    X As Single
    Y As Single
    Decay As Single
    Heading As Single
    Facing As Single
    Speed As Single
    LastGravity As Long
    bSniperCasing As Boolean
End Type

Private Type ptMagazine
    X As Single
    Y As Single
    Decay As Single
    Heading As Single
    Speed As Single
    LastGravity As Long
    
    bOnSurface As Boolean
    iMagType As eMagTypes
End Type

Private Type ptDeadStick
    X As Single
    Y As Single
    colour As Long
    Decay As Long
    bOnSurface As Boolean
    
    Speed As Single
    Heading As Single
    
    LastGravity As Long
    
    bFacingRight As Single
    bFlamed As Boolean
    bIsMe As Boolean
End Type

Private Type ptDeadChopper
    X As Single
    Y As Single
    colour As Long
    Decay As Long
    bOnSurface As Boolean
    
    Speed As Single
    Heading As Single
    
    LastGravity As Long
    
    LastSmoke As Long
    
    iOwner As Integer
End Type

Private Type ptFlame
    X As Single
    Y As Single
    Heading As Single
    Speed As Single
    
    OwnerID As Integer
    Decay As Long
    
    Size As Single
    
    LastGravity As Long
End Type

Private Type ptSpark
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    Decay As Long
    LastReduction As Long
End Type

Private Type ptStaticWeapon
    X As Single
    Y As Single
    iWeapon As eWeaponTypes
    
    bOnSurface As Boolean
    LastGravity As Long
    Speed As Single
    Heading As Single
End Type

'Private Type ptSmallSmoke
'    'X As Single
'    'Y As Single
'    'Heading As Single
'    'Speed As Single
'    AngleFromMain As Single
'    DistanceFromMain As Single
'
'    sAspect As Single
'    AspectDir As Integer
'
'    DistanceFromMainInc As Single
'End Type
'Private Type ptLargeSmoke
'    CentreX As Single
'    CentreY As Single
'
'    SingleSmoke(1 To 10) As ptSmallSmoke
'
'    iSize As Single
'    iDirection As Integer
'
'    'pPoly(1 To 10) As POINTAPI
'End Type

'Private Type ptAmmoPack
'    X As Single
'    Y As Single
'
'    bOnSurface As Boolean
'    LastGravity As Long
'    Speed As Single
'    Heading As Single
'End Type

'explosions
Private Type ptCirc
    X As Single
    Y As Single
    
    
    currentRadius As Single
    MaxRadius As Single
    
    ExpandSpeed As Single
    
    colour As Long
    sgDirection As Single 'so i can do "*sgDirection" [1 & -1]
End Type

Private Type ptWallMark
    X As Single
    Y As Single
    Radius As Single
    Decay As Long
End Type

Private Type ptSmokeBlast
    X As Single
    Y As Single
    
    Heading As Single
    sOffset As Single
    
    sLength As Single
    'sWidth As Single
    
    'iDir As Integer
End Type

Private Type ptExplosiveBarrel
    X As Single
    Y As Single
    
    LastTouchID As Integer
    'nBulletsHit As Integer
    iHealth As Integer
    
    ID As Integer
End Type

Private Type ptTimeZone
    X As Single
    Y As Single
    TimeAdjust As Single
    'Decay As Long
    sSize As Single
    bShrinking As Boolean
End Type
Private Type ptGravityZone
    X As Single
    Y As Single
    
    sSize As Single
    bShrinking As Boolean
End Type

Private Type ptGrass
    X As Single
    iPlatform As Integer
    
    RndK1 As Single
    RndK2 As Single
    RndK3 As Single
    RndK4 As Single
End Type

Private Type ptCircleBlast
    sgSize As Single
    iDirection As Single
    
    X As Single
    Y As Single
    
    bFading As Boolean
End Type

Private Type ptBulletTrail
    Heading As Single
    Speed As Single
    'Speed_Accel As Single
    
    X As Single
    Y As Single
    
    SpawnTime As Long
    
    sgLength As Single
    
    bTracer As Boolean
End Type

Private Type ptNadeTrail
    Heading As Single
    Speed As Single
    
    X As Single
    Y As Single
    
    SpawnTime As Long
    LastSmoke As Long
    LastGravity As Long
    lColour As Long
End Type
Private Type ptGravitySmoke
    X As Single
    Y As Single
    
    Heading As Single
    Speed As Single
    
    SpawnTime As Long
    lColour As Long
    
    sgSize As Single
End Type

Private Type ptPath
    XStart As Single
    YStart As Single
    XEnd As Single
    YEnd As Single
End Type

Private Type ptAttentionGrabber
    Decay As Long
    X As Single
    Y As Single
    lColour As Long
End Type

Private Type ptHead
    Decay As Long
    LastGravity As Long
    
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    
    lColour As Long
End Type

Private Type ptShieldWave
    X As Single
    Y As Single
    Facing As Single
    Size As Single 'arc size
    colour As Long
End Type

Private TimeZoneCircs() As ptCirc, ScreenCircs() As ptCirc
Private NumTimeZoneCircs As Integer, NumScreenCircs As Integer

Private Const MAXBULLETS As Long = 750
Private NumBullets As Long
Private Bullet(MAXBULLETS) As ptBullet

Private Const Max_Sparks As Long = 500 'number of sparks allowed in game to prevent visual lag
Private Const MAXSPARKS As Long = Max_Sparks + 1
Private NumSparks As Long
Private Spark(MAXSPARKS) As ptSpark

Private Const Max_Casings As Long = 750
Private Const MAXCASINGS As Long = Max_Casings + 1
Private NumCasings As Long
Private Casing(MAXCASINGS) As ptCasing

Private NumBlood As Long
Private Blood() As ptBlood

Private NumNades As Long
Private Nade() As ptNade

Private NumMines As Long
Private Mine() As ptMine

Private NumDeadSticks As Long
Private DeadStick() As ptDeadStick

Private NumMags As Long
Private Mag() As ptMagazine

Private NumDeadChoppers As Long
Private DeadChopper() As ptDeadChopper

Private NumFlames As Long
Private Flame() As ptFlame

Private NumStaticWeapons As Long
Private StaticWeapon() As ptStaticWeapon

'Private NumLargeSmokes As Integer
'Private LargeSmoke() As ptLargeSmoke

Private NumWallMarks As Long
Private WallMark() As ptWallMark

Private NumSmokeBlasts As Long
Private SmokeBlast() As ptSmokeBlast

Private NumBarrels As Long
Private Barrel() As ptExplosiveBarrel

Private NumTimeZones As Long
Private TimeZone() As ptTimeZone

Private NumGravityZones As Long
Private GravityZone() As ptGravityZone

Private NumGrass As Long
Private Grass() As ptGrass

Private NumCircleBlasts As Long
Private CircleBlast() As ptCircleBlast

Private NumBulletTrails As Long
Private BulletTrail() As ptBulletTrail

Private NumNadeTrails As Long
Private NadeTrail() As ptNadeTrail

Private NumGravitySmokes As Long
Private GravitySmoke() As ptGravitySmoke

Private NumAttentions As Long
Private Attention() As ptAttentionGrabber

Private NumHeads As Long
Private Head() As ptHead

Private NumShieldWaves As Long
Private ShieldWave() As ptShieldWave

Private NumFires As Integer
Private Fire() As ptFlame
'############ ADD TO RESETVARS() ##################


'optimization stuff
Private NumSticksM1 As Integer


'angle stuff
Private Const SmallAngle = Pi / 4
'end angle


'stats
Private Const Health_Start As Integer = 100
Private Current_Health_Start As Integer

Private Const Accel = 4
Private Const Max_Speed = 70 '112
Private Const JumpMultiple = 120 'move stick up by Accel*JumpMultiple
Private Const NadeMultiple = JumpMultiple * 1.2 'force stick away by accel*nademultiple

Private Const Gravity_Strength As Single = 14, _
              Gravity_Zone_Strength As Single = Gravity_Strength * -4, _
              Bullet_Gravity_Strength As Single = 2, _
              Gravity_Zone_Direction As Single = Pi


Private Const Gravity_Direction = Pi
Private Const Gravity_Delay = 100
Private Const GravityZone_Time = 30000, GravityZone_Radius = 6000, GravityZone_Colour = MGrey
'Private Const JumpTime = 100

Private Const SmokeOutline As Long = &HCCCCCC '&HDDDDDD
Private Const SmokeFill As Long = &HE1E1E1
'Private Const SmokeOutline = &H777777
'Private Const SmokeFill = &HFDFDFD

Private Const BoxCol As Long = &HC0C0C0

Private Const Attention_Time As Long = 1500, _
              Head_Time = 10000
Private Const Head_Bounce_Reduction = 2 / 3

Private Const BulletTrail_Smoke_Delay As Long = Frame_Const, BulletTrail_Time As Long = 300&, _
    def_BulletTrailLen As Single = 500, BulletTrail_StartSpeed As Single = 20

Private Const Blood_Time = 750
Private Const Casing_Time = 10000
Private Const Casing_Len = 25

Private Const DeadStickTime = 40000, _
    MAX_DeadStick_And_StaticWeap_Speed = 250

Private Const Nade_Arm_Time = 300

Private Const Lim = 50
Private Const Left_Indent = 15000 'for bot coop positioning
'Private Const Winner_Colour As Long = MSilver
'Private Const Winner_DrawMode As Long = vbInvert
'vbXorPen
'vbNotMaskPen
'vbMaskPen
'vbNotXorPen
'end stats


'######################################################################
Private Const StickSize As Integer = 800
Private Const HeadRadius As Integer = StickSize \ 10
Private Const BodyLen As Integer = HeadRadius * 4
Private Const ArmLen = HeadRadius * 2
Private Const ArmNeckDist As Integer = 250
Private Const LegHeight As Integer = StickSize / 3.2
Private Const MaxLegWidth As Integer = 90
Private Const StickHeight As Integer = 1100

Private Const Bullet_Radius As Integer = 5
Private Const Bullet_Decay As Integer = 5000
Private Const Bullet_Damage As Integer = 3
Private Const BULLET_SPEED As Single = 410 '250  'Max_Speed * 2 '=222
Private Const BULLET_LEN  As Single = StickHeight \ 16
Private Const Bullet_Min_Speed  As Single = 35
Private Const Bullet_Diffract_Delay As Single = 400
Private Const Bullet_Wall_Diffract_Delay As Single = 50
Private Const Bullet_Silenced_Damage_Factor  As Single = 0.75
Private Const HeadShot_Damage_Factor As Single = 3
'######################################################################

'Weapon stats/optimisation
Private AmmoFired(0 To eWeaponTypes.Chopper) As Integer
Private kBulletDelay(0 To eWeaponTypes.Chopper) As Integer
Private kMaxRounds(0 To eWeaponTypes.Chopper) As Integer
Private kReloadTime(0 To eWeaponTypes.Chopper) As Integer
Private kPerkName(0 To eStickPerks.pSpy) As String
Private kTeamColour(0 To eTeams.Spec) As Long
Private kGameType(0 To eStickGameTypes.gCoOp) As String
Private kRecoverAmount(0 To eWeaponTypes.Chopper) As Single
Private kRecoilAmount(0 To eWeaponTypes.Chopper) As Single
Private kRecoilTime(0 To eWeaponTypes.Chopper) As Long
Private kRecoilForce(0 To eWeaponTypes.Chopper) As Single
Private kNadeName(0 To eNadeTypes.nEMP) As String
Private kSilencable(0 To eWeaponTypes.Chopper) As Boolean
Private kBurstBullets(0 To eWeaponTypes.Chopper) As Integer, kBurstDelay(0 To eWeaponTypes.Chopper) As Long
Private kBulletDamage(0 To eWeaponTypes.Chopper) As Single, kBulletSpeed(0 To eWeaponTypes.Chopper) As Single, _
    kWeapon_Special(0 To eWeaponTypes.Chopper) As Boolean, kSprayAngle(0 To eWeaponTypes.Chopper) As Single, _
    kSilencedOffset(0 To eWeaponTypes.Chopper) As Integer
Private StickIndexIDMap() As Integer

'WHEN ADDING A WEAPON:
'Add to eWeaponTypes
'Add to MakeWeaponNameArray()
'Make the constants
'Make the Draw<Weap_Name>, Draw<Weap_Name>2 and DrawStatic<Weap_Name> procedures
'
'Alter:
'MakeBulletDelayArray()
'MakeMaxRoundsArray()
'InitWeaponStats()
'MakeReloadTimeArray()
'GetTotalMags()
'
'Add to DrawCrossHairPoint()
'Add to DrawStick() [weapon]
'Add to WeaponSilencable() [if needed]
'Add to WeaponSupportsFireMode()
'
'Add to FireShot() ?


Private Const W1200_Gauge = 13
Private Const W1200_Spray_Angle = Pi / 17
Private Const W1200_Recoil_Time = 800
Private Const W1200_SingleRecoil_Angle = SmallAngle
Private Const W1200_Recover_Amount = W1200_SingleRecoil_Angle / 50
Private Const W1200_Bullet_Delay = 800
Private Const W1200_Bullets = 6
Private Const W1200_Reload_Time = 1600, W1200_Round_Reload_Delay = W1200_Reload_Time \ W1200_Bullets
Private Const W1200_Bullet_Damage = 120 / W1200_Gauge 'was 7 - will kill if all shots hit
Private Const W1200_RecoilForce = 10
Private Const W1200_Mags As Long = 1.8 * W1200_Gauge

Private Const SPAS_Gauge = 8
Private Const SPAS_Spray_Angle = Pi / 50
Private Const SPAS_Recoil_Time = 500
Private Const SPAS_SingleRecoil_Angle = SmallAngle / 2
Private Const SPAS_Recover_Amount = SPAS_SingleRecoil_Angle / 25
Private Const SPAS_Bullet_Delay = 600
Private Const SPAS_Bullets = 8
Private Const SPAS_Reload_Time = 2200, SPAS_Round_Reload_Delay = SPAS_Reload_Time \ SPAS_Bullets
Private Const SPAS_Bullet_Damage = 101 / SPAS_Gauge 'was 7 - will kill if all shots hit
Private Const SPAS_RecoilForce = 6
Private Const SPAS_Mags = 3 * SPAS_Gauge

Private Const AK_Spray_Angle = Pi / 75
Private Const AK_Recoil_Time = 50
Private Const AK_SingleRecoil_Angle = SmallAngle / 90
Private Const AK_Recover_Amount = AK_SingleRecoil_Angle / AK_Recoil_Time
Private Const AK_Bullet_Delay = 100 '600 rpm = 10 rps, delay = 1/10 = 0.1 = 100ms
Private Const AK_Bullets = 30
Private Const AK_Reload_Time = 1800
Private Const AK_Bullet_Damage = 21 '2010 joules
Private Const AK_Mags = 4

Private Const AUG_Spray_Angle = Pi / 350
Private Const AUG_Recoil_Time = 10
Private Const AUG_SingleRecoil_Angle = SmallAngle / 350
Private Const AUG_Recover_Amount = AUG_SingleRecoil_Angle / AUG_Recoil_Time
Private Const AUG_Bullet_Delay = 300 '175 is about the time taken for the game to remove the stick's fire state
Private Const AUG_Single_Bullet_Delay = 88 '#####THIS IS THE ONE TO SET##### 680 rpm = 11+1/3 rps, delay = 1/(11+1/3) = 0.088 = 88ms
Private Const AUG_Bullets = 30
Private Const AUG_Burst_Bullets = 3
Private Const AUG_Reload_Time = 1800
Private Const AUG_Bullet_Damage = 20
'1775 joules - more due to burst fire - (AUG_Bullet_Damage*3)*2>100
Private Const AUG_Mags = 3

Private Const G3_Recoil_Time = 70
Private Const G3_SingleRecoil_Angle = SmallAngle / 5
Private Const G3_Recover_Amount = G3_SingleRecoil_Angle / 3
Private Const G3_Single_Bullet_Delay = 120 '500 rpm = 8+1/3 rps, delay = 1/(8+1/3) = 0.12 = 120ms
Private Const G3_Bullet_Delay = 250
Private Const G3_Bullets = 20
Private Const G3_Reload_Time = 2000
Private Const G3_Bullet_Damage As Long = 100 / HeadShot_Damage_Factor 'i.e. 1 headshot to kill
Private Const G3_Sniper_Damage_Factor As Single = 1.5
Private Const G3_Mags = 4
Private Const G3_Burst_Bullets = 2
Private Const G3_Speed As Single = BULLET_SPEED * 1.5

Private Const M82_Recoil_Time = 460
Private Const M82_SingleRecoil_Angle = SmallAngle / 1.5
Private Const M82_Recover_Amount = M82_SingleRecoil_Angle / 30 '15
Private Const M82_Bullet_Delay = 500
Private Const M82_Bullets = 6
Private Const M82_Reload_Time = 3000 '1500
Private Const M82_Bullet_Damage = 140
Private Const M82_RecoilForce = 25
Private Const M82_Mags = 2
Private Const M82_Speed As Single = BULLET_SPEED * 1.4
Private Const M82_Silent_Recoil_Reduction = 3

Private Const AWM_Recoil_Time = 1000
Private Const AWM_SingleRecoil_Angle = SmallAngle / 1.2
Private Const AWM_Recover_Amount = AWM_SingleRecoil_Angle / 65
Private Const AWM_Bullet_Delay = 1300
Private Const AWM_Bullets = 5
Private Const AWM_Reload_Time = 3000
Private Const AWM_Bullet_Damage = 125
Private Const AWM_RecoilForce = 12
Private Const AWM_Mags = 4
Private Const AWM_Speed As Single = BULLET_SPEED * 1.8

Private Const XM8_Spray_Angle = Pi / 200
Private Const XM8_Recoil_Time = 15
Private Const XM8_SingleRecoil_Angle = SmallAngle / 300
Private Const XM8_Recover_Amount = XM8_SingleRecoil_Angle / XM8_Recoil_Time
Private Const XM8_Bullet_Delay = 80 '750 rpm = 12.5 rps, delay = 1/12.5 = 0.08 = 80ms
Private Const XM8_Bullets = 30
Private Const XM8_Reload_Time = 1000
Private Const XM8_Bullet_Damage = 19 '1775 joules
Private Const XM8_Mags = 3

Private Const MP5_Spray_Angle = Pi / 250
Private Const MP5_Recoil_Time = 50
Private Const MP5_SingleRecoil_Angle = SmallAngle / 275
Private Const MP5_Recover_Amount = MP5_SingleRecoil_Angle / MP5_Recoil_Time
Private Const MP5_Bullet_Delay = 80 '750 rpm = 12.5 rps, delay = 1/12.5 = 0.08 = 80ms
Private Const MP5_Bullets = 30
Private Const MP5_Reload_Time = 1800
Private Const MP5_Bullet_Damage = 15
Private Const MP5_Mags = 3

Private Const Mac10_Spray_Angle = Pi / 150
Private Const Mac10_Recoil_Time = 13
Private Const Mac10_SingleRecoil_Angle = SmallAngle / 200 'visually recoil, but not via mouse coords
Private Const Mac10_Recover_Amount = Mac10_SingleRecoil_Angle / Mac10_Recoil_Time
Private Const Mac10_Bullet_Delay = 54 '1100 rpm = 18.33 rps, delay = 1/18.33 = 0.05454 = 54ms
Private Const Mac10_Bullets = 32
Private Const Mac10_Reload_Time = 600
Private Const Mac10_Bullet_Damage = 15
Private Const Mac10_Mags = 4

Private Const RPG_Recoil_Time = 1500
Private Const RPG_SingleRecoil_Angle = piD4
Private Const RPG_Recover_Amount = RPG_SingleRecoil_Angle / 90 '60
Private Const RPG_Bullet_Delay = 1500 'for Rocket_Spam, set this to 0, and _
                                       don't inc .BulletsFired in AddNade(), _
                                       or set _Bullets to however many
Private Const RPG_Bullets As Long = 1
Private Const RPG_Reload_Time As Long = 2500
Private Const RPG_Smoke_Delay As Long = 20
Private Const RPG_RecoilForce As Single = 12
Private Const RPG_Speed As Single = 250
Private Const RPG_Mags As Integer = 4

Private Const M249_Spray_Angle = -Pi / 50
Private Const M249_Recoil_Time = 50
Private Const M249_SingleRecoil_Angle = SmallAngle / 90
Private Const M249_Recover_Amount = M249_SingleRecoil_Angle / M249_Recoil_Time
Private Const M249_Bullet_Delay = 60 '1000 rpm = 16 rps, delay = 1/16 = 0.06 = 60ms
Private Const M249_Bullets = 180
Private Const M249_Reload_Time = 5000
Private Const M249_Bullet_Damage = 13
Private Const M249_RecoilForce = 7 '''''''''Doesn't give game a chance to stop stick moving
Private Const M249_Mags = 2

Private Const DEagle_Spray_Angle = Pi / 50
Private Const DEagle_Recoil_Time = 1200
Private Const DEagle_SingleRecoil_Angle = SmallAngle * 1.4
Private Const DEagle_Recover_Amount = DEagle_SingleRecoil_Angle / 64 '17
Private Const DEagle_Bullet_Delay = 1200
Private Const DEagle_Bullets = 7
Private Const DEagle_Reload_Time = 1000
Private Const DEagle_Bullet_Damage = 70
Private Const DEagle_RecoilForce = 7
Private Const DEagle_Mags = 2

Private Const Flame_Speed = 90
Private Const Flame_Bullet_Delay = 50
Private Const Flame_Bullets = 30
Private Const Flame_Reload_Time = 1000
Private Const Flame_Time = 2500 'time flames last
Private Const Flame_Max_Radius = 350
Private Const Flame_Damage = 15 'damage from standing in the flame-line
Private Const Flame_Impact_Delay = 150 'do above damage every Flame_Impact_Delay seconds, for each flame touching
Private Const Flame_Burn_Time = 20000 'time that a flame'll burn after touch
Private Const Flame_Burn_Damage = 1 'totaldamage (after Flame_Burn_Time seconds) = _
                                     Flame_Burn_Damage * Flame_Burn_Time / Flame_Burn_Damage_Time 'damage to apply to a burn
Private Const Flame_Burn_Damage_Time = 500 'apply above damage every x milliseconds
Private Const Flame_Burn_Radius = 100 'stick on fire
Private Const Flame_Inertia_Reduction = 3
Private Const Flamethrower_Mags = 6
Private Const Fire_Smoke_Delay As Long = 500, _
              Fire_Time As Long = 10000 'seconds to leave a fire going for


Private Const USP_Recoil_Time = 300
Private Const USP_SingleRecoil_Angle = SmallAngle
Private Const USP_Recover_Amount = USP_SingleRecoil_Angle / 20 '17
Private Const USP_Bullet_Delay = 400 'Frame_Const
Private Const USP_Bullets = 12
Private Const USP_Reload_Time = 800
Private Const USP_Bullet_Damage = 34
'Private Const USP_Burst_Bullets = 2
Private Const USP_Mags = 4


Private Const Knife_Delay = 100
Private Const Throwing_Strength = 125

Private Const Nade_Explode_Radius = 2300
Private Const Nade_Radius = 50
Private Const Nade_Time = 2000 'time until BOOM
Private Const Nade_Delay = 5000 'time until can throw next nade
Private Const Nade_Bounce_Reduction = 1.3 'non-elastic
'Private Const Nade_Bullet_Invul_Time = Nade_Time / 2 'Time that nade can't be shot
'--------------------------------------------------------------------------
Private Const Mine_Radius = 4
Private Const Mine_Explode_Radius = 3000, Mine_Damage = 250000
Private Const Mine_Delay = 10000
Private Const Mine_Y_Increase As Single = 570 '570.06 'BodyLen + HeadRadius * 3.5
'Private Const Mine_Hold_Time = 1000
Private Mine_StickLim As Long ', Mine_StickLimY As Long



Private Const Casing_Bounce_Reduction = 2 / 3
Private Const MFlash_Time = Frame_Const * 4 'muzzle flash

Private Const Nade_Release_Delay = 1000 'time until state is removed from said stick
Private Const Bullet_Release_Delay As Long = 200, KeyPressDelay As Long = 150 'make sure the shot gets through/sent via packet
Private FireKeyUpTime As Long
Private Const AutoReload_Delay = 400

Private Const SwitchWeaponDelay = 260
'Private Const UseKeyReleaseDelay = 300
Private Const StaticWeaponSendDelay = 2000, Max_Static_Weapons = 25, Min_Static_Weapons = 10

Private Const Sniper_Smoke_Delay = 10 '25

Private Const Hardcore_Damage_Amp = 2 '1.5
'end weapon stats
'##################################################################
'my stats
'Private Const FullRadar_Time = 30000
'Private Const RowKillsForFullRadar = 4
'Private FullRadarStartTime As Long
'Private bHadFullRadar As Boolean 'for displaying "Radar Expired"
'############## ResetVars() and StopPlay() need adding to ############

Public ChopperAvail As Boolean
Private Const RowKillsForChopper = 6

Private Const RowFlameKillsForToasty = 3

'Private KillsInARow As Integer
Private FlamesInARow As Integer

Private KnifesInARow As Integer
Private Const KnivesForSaber = 3

Private Const RowKillsForShield = 3
'Private Const Shield_Colour = MSilver 'vbBlack
Private Const Max_Shield As Integer = 100, shieldChargeDist As Single = 300, _
    ShieldDamageDec As Single = 2, Shield_Recharge_Delay As Long = 1500
Private Const Radar_Bullet_ShowTime = 4000


Private FireMode_Current As eFireModes, FireMode_2ndWeapon As eFireModes
Private b2ndWeaponSilenced As Boolean

Private LastNadeSwitch As Long, LastFireModeSwitch As Long, _
       LastProneSwitch As Long, LastWeaponSwitch As Long, _
       LastCrouchToggle As Long


'end my stats
'##################################################################
'other stuff
Private Const FlashBang_Time = 5000

Private Const BarrelWidth = 200, BarrelHeight = 500
Private Const Barrel_Explode_Radius = Nade_Explode_Radius * 2, _
    BarrelMultiple = NadeMultiple / 2

Private Const def_TimeAdjust As Single = 0.2
Private Const TimeZone_Time = 30000, TimeZone_Radius = 6000, TimeZone_Colour = MSilver

Private Const Tracer_Col = vbRed
Private Tracer_DefCol As Long
'end other stuff
'##################################################################

'chopper stats
Private Const Chopper_Max_Speed = 75
Private Const Chopper_Lift = Accel * 2
Private Const Chopper_RPG_Delay = 2500
Private Const DeadChopperTime = 30000

Private Const DeadChopper_Smoke_Delay As Long = Frame_Const

Private Const Chopper_Bullet_Damage = 13
Private Const Chopper_Bullet_Delay = Frame_Const + 3
Private Const Chopper_Spray_Angle As Single = modSpaceGame.piD40
Private Const Chopper_Impact_Speed_Dec = 1.5
Private Const Chopper_Damage_Reduction = 5
'Private Const Chopper_Casing_Reduction As Single = 4
Private Const ChopperLen = 3500, _
    CLD2 = ChopperLen / 2, _
    CLD3 = ChopperLen / 3, _
    CLD4 = ChopperLen / 4, _
    CLD6 = ChopperLen / 6, _
    CLD8 = ChopperLen / 8, _
    CLD10 = ChopperLen / 10
'end chopper stats

Private Const Max_Health As Integer = 100

'bullet stuff above

Private Const Mag_Decay = 8000

Private Const GunLen = 213 'BodyLen / 1.5

Private Const Spark_Time As Long = 2500
Private Const Spark_Diffraction As Single = piD3, Spark_Speed As Long = 50 '<-- for groups
Private Const Spark_Min_Speed As Long = 1, Spark_Speed_Reduction As Single = 1.08
Private Const Spark_Speed_Reduction_Delay As Long = 40 '35

Private Const WallMark_Time As Long = 60000
Private Const WallMark_Bullet_Radius As Long = 30, WallMark_Explosion_Radius As Long = WallMark_Bullet_Radius * 3
Private Const CircleBlast_MaxSize As Single = 300

Private Const Degrees10 As Single = Pi * 1 / 18
Private Const ProneRightLimit As Single = piD2 + Degrees10, ProneLeftLimit As Single = pi3D2 - Degrees10

Private Const Chat_Round_Offset As Single = 9, Chat_X_Offset As Single = 10, _
    Chat_Chat_Offset As Single = 1000, Chat_Kills_Offset As Single = 3750

'end drawing

'Private LastSpawnTime As Long
Private Const Spawn_Invul_Time As Long = 750 'time they are invulnerable for

Private Const Max_Chat As Long = 24
Private FPS As Integer

Private Const mPacket_LAG_TOL = 1000  'Milliseconds to wait before rendering a stick motionless
Private Const mPacket_LAG_KILL = 7000    'Milliseconds to wait before removing a stick due to lack of info
Private Const StickServer_RETRY_FREQ = 2000   'Milliseconds between attempts to connect to StickServer
Private Const StickServer_NUM_RETRIES = 5
Private Const ServerVarSendDelay = 3000
Private Const BoxInfoDelay = 7000
Private Const LagOut_Delay = mPacket_LAG_TOL * 3 'time to lag out
Private Const NameCheckDelay = 10000
Private LastUpdatePacket As Long 'Are we lagging out?

Private WindowClosing As Boolean

'Public MyID As Integer 'Which Stick are we?
Private LastScoreCheck As Long


'########################################################################
Private LastDamageTick As Long
Private Const DamageTickTime As Long = 750

'########################################################################

Private Const Grass_Col As Long = &H68000 'RGB(0,128,6)
Private LastGrassRefresh As Long, LastFireRefresh As Long

'########################################################################

'packet stuff

'Private Type ptPacketToSend
'    sToSend As String
'    lDecay As Long
'    lLastSend As Long
'End Type
'Private PacketsToSend() As ptPacketToSend
'Private NumPacketsToSend As Integer
'Private Const PacketsToSend_Time = 2000, PacketsToSendRetry = 300


Private Type ptStickPacket
    ID As Integer
    PacketID As Long
    
    ActualFacing As Single
    Facing As Single
    Heading As Single
    Speed As Single
    X As Single
    Y As Single
    state As Integer
    
    WeaponType As eWeaponTypes
    'PrevWeapon As eWeaponTypes
    Health As Integer
    Shield As Integer
    
    iNadeType As eNadeTypes
End Type

Private Type ptStickSlowPacket
    Name As String * 15
    ID As Integer
    colour As Long
    
    iKills As Integer
    iDeaths As Integer
    iKillsInARow As Integer
    
    Team As eTeams
    bAlive As Boolean
    Perk As eStickPerks
    MaskID As Integer
    
    bSilenced As Boolean
    bTyping As Boolean
    bFlashed As Boolean
    bOnFire As Boolean
    
    bLightSaber As Boolean
    CurrentWeap1 As eWeaponTypes
    CurrentWeap2 As eWeaponTypes
    
    
    Burst_Bullets As Integer
    Burst_Delay As Long
    
    
    PacketID As Long
End Type

Private mPacket As ptStickPacket
Private msPacket As ptStickSlowPacket

Private Const mPacket_SEND_DELAY = 65 'Milliseconds between update packets
Private Const msPacket_SEND_DELAY = 250 'Milliseconds between slow update packets
Private Const Client_Nade_Delay = mPacket_SEND_DELAY + 25
'########################################################################

Public bRunning As Boolean  'Is the render loop running?
Private bPlaying As Boolean 'In the middle of a game?
Private PacketTimer As Long 'Time at which last mPacket was sent
Private ServerSockAddr As ptSockAddr   'StickServer's sock addr
Public lSocket As Long       'Socket with which we'll send/receive essages

Private Type CHATTYPE
    Text As String           'The string to display
    Decay As Long            'When is this one removed?
    colour As Long
    bChatMessage As Boolean
    
    sTextHeight As Single
    sTextWidth As Single
End Type

Private Chat() As CHATTYPE       'Our chat array
Private NumChat As Long          'How many chat messages are there currently?
Private Const CHAT_DECAY = 15000        'How long before chat messages disappear?

'big message(s)
Private Type ptMainMessage
    Text As String           'The string to display
    Decay As Long            'When is this one removed?
    colour As Long
End Type

Private MainMessages() As ptMainMessage       'Our chat array
Private NumMainMessages As Long          'How many chat messages are there currently?
Private Const MainMessageDecay = 5000        'How long before chat messages disappear?

'Smoke ########################################################################

Private Type ptSmoke
    X As Single
    Y As Single
    
    Size As Single
    
    Direction As Integer '1=grow (2x rate), -1=shrink
    
    Speed As Single
    Heading As Single
    
    bLongTime As Boolean
    
End Type

Private Smoke() As ptSmoke
Private NumSmoke As Integer


'##############################################################################

Private Type ptServerVars
    'bAllowRockets As Boolean
    'bAllowFlameThrowers As Boolean
    'bAllowChoppers As Boolean
    sAllowedWeapons As String
    
    'bShootNades As Boolean
    bDrawNadeTime As Boolean
    bHPBonus As Boolean
    
    sgGameSpeed As Single
    sgDamageFactor As Single
    
    iGameType As eGameTypes
    bHardCore As Boolean
    'b2Weapons As Boolean
    bBulletsThroughWalls As Boolean
    bSpawnWithShield As Boolean
    
    iSpawnDelay As Integer
    iScoreToWin As Integer
    
    iSequenceNo As Long
End Type
Private LastServerSettingVar As Long

'##############################################################################

Private Type ptHealthPack
    X As Single
    Y As Single
    bActive As Boolean
    LastUsed As Long
End Type

Private HealthPack As ptHealthPack
Private Const HealthPack_Radius = 100
Private Const HealthPackDelay = 15000

Public HealthPackX As Long, HealthPackY As Long


'ammo pickups
Private Const NumAmmoPickUpsM1 = 4 - 1
Private Const AmmoPickUp_Spawn_Delay = 30000
Private AmmoPickup(0 To NumAmmoPickUpsM1) As ptHealthPack
Private TotalMags(0 To eWeaponTypes.Knife - 1) As Byte

'##############################################################################
'chat
Private ChatFlashDelay As Long 'Const ChatFlashDelay = 300 'for the _ thing
Private LastFlash As Long
Private bChatCursor As Boolean 'for the _ thing
Private bChatActive As Boolean   'chatting?
Private strChat As String 'current chat string

'##############################################################################
'perks
Private Const JuggernautDamageReduction As Single = 3
Private Const StoppingPowerIncrease As Single = 2.4

Private Const ConditioningAccelInc As Single = 5
Private Const ConditioningMaxSpeedInc As Single = 2

Private Const SteadyAim_Reduction As Single = 3

Private Const Sniper_Damage_Inc As Single = 3
Private Const Sniper_Max_Speed_Dec As Single = 1.2

Private Const Mechanic_Bullet_Inc As Single = 2
'Private Const StealthESPDist = StickGameWidth \ 2

Private Const Zombie_Health As Integer = 1500, Zombie_Mine_Weakness As Integer = 10, ZombieMaxSpeedDec As Single = 0.1
Private Const Zombie_Col As Long = &H42AE   '174,66,0

'##############################################################################

'Round Stuff
Private Const RoundInfoSendDelay = 2000
Private Const ScoreCheckDelay = 1000
Private Const RoundWaitTime = 10000
Private Const PresenceSendDelay = 500
Private RoundWinnerID As Integer
Private RoundPausedAtThisTime As Long

'##############################################################################
'resize 'constants'
Private RadarLeft As Single ': RadarLeft = Me.width - RadarWidth - 100
Private PlayingX As Single ': PlayingX = StickCentreX - 600
Public ConnectingkX As Single ': kX = StickCentreX - 900
Public ConnectingkY As Single ': kY = StickCentreY + 650

Private Const RadarWidth = 2000

'##############################################################################
Private Type ptBotTaunt
    sTaunt As String
    bAddName As Boolean
End Type

Private BotTaunts() As ptBotTaunt, kBotNames() As String

'##############################################################################

Private Declare Function IntersectRect Lib "user32" ( _
    lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long


'##############################################################################

Private MouseX As Single, MouseY As Single, bHasFocus As Boolean
Private StunnedMouseX As Single, StunnedMouseY As Single

'##############################################################################

Private Const sAccepts As String * 1 = "A"
Private Const sBoxInfos As String * 1 = "B"
Private Const sChats As String * 1 = "C"
Private Const sKillAndDeathInfos As String * 1 = "D"
'Private Const sExits As String * 1 = "E"
Private Const sExplodeMines As String * 1 = "F"
Private Const sGrassRefreshs As String * 1 = "G"
Private Const sHealthPacks As String * 1 = "H"
'Public Const sMapNames As String * 1 = "I"
Private Const sJoins As String * 1 = "J"
'Public Const sKicks As String * 1 = "K" 'borrow spacegame one
Private Const sSlowUpdates As String * 1 = "L"
Private Const sMineRefreshs As String * 1 = "M"
Private Const sBarrelRefreshs As String * 1 = "N"
Private Const sExplodeBarrels As String * 1 = "O"
Private Const sPresences As String * 1 = "P"
'Public Const sMapRequests As String * 1 = "Q"
Private Const sRoundInfos As String * 1 = "R"
Private Const sServerVarss As String * 1 = "S"
Private Const sTimeZoneRefreshs As String * 1 = "T"
Private Const sUpdates As String * 1 = "U"
Private Const sGravityZoneRefreshs As String * 1 = "V"
Private Const sStaticWeaponUpdates As String * 1 = "W"
Private Const sDamageTicks As String * 1 = "X"
Private Const sWeaponSwapInfos As String * 1 = "Y"
'Public Const sNewMaps As String * 1 = "Z"
'Private Const sFireRefreshs As String * 1 = "f"

'Letters left: abcdeghijklmnopqrstuvwxyz
'##############################################################################

Private Const Ally_Colour = vbGreen, Enemy_Colour = vbBlack 'text


'Private LeftKey As Boolean, RightKey As Boolean, JumpKey As Boolean, CrouchKey As Boolean, ProneKey As Boolean, _
    ReloadKey As Boolean, MineKey As Boolean
Private ShowScoresKey As Boolean
Private UseKey As Boolean, FireKey As Boolean, CrouchKey As Boolean


Private SpecUp As Boolean, SpecDown As Boolean, SpecLeft As Boolean, SpecRight As Boolean


'Private Const ControlKey = 17
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private WeaponKey As eWeaponTypes
'Private Scroll_WeaponKey As eWeaponTypes
'Private LastScrollWeaponSwitch As Long
'Private Const Scroll_Delay = 750

'zoom
Private LastZoomPress As Long
Private Const ZoomShowTime = 750
Private Const ZoomInc = 0.05
Private Const MinZoom = 0.6 + ZoomInc
Private Const MaxZoom = 2
Private bFullScreened As Boolean

'Private Sub AddPacketsToSend(sString As String)
'Dim GTC As Long
'
'ReDim Preserve PacketsToSend(NumPacketsToSend)
'
'GTC = GetTickCount()
'
'With PacketsToSend(NumPacketsToSend)
'    .lDecay = GTC + PacketsToSend_Time
'    .sToSend = sString
'    .lLastSend = GTC
'End With
'
'NumPacketsToSend = NumPacketsToSend + 1
'
'End Sub
'
'Private Sub RemovePacketsToSend(Index As Integer)
'
'Dim i as integer
'
'If NumPacketsToSend = 1 Then
'    Erase PacketsToSend
'    NumPacketsToSend = 0
'Else
'    'Remove the bullet
'    For i = Index To NumPacketsToSend - 2
'        PacketsToSend(i) = PacketsToSend(i + 1)
'    Next i
'
'    'Resize the array
'    ReDim Preserve PacketsToSend(NumPacketsToSend - 2)
'    NumPacketsToSend = NumPacketsToSend - 1
'End If
'
'End Sub
'
'Private Sub ProcessPacketsToSend()
'Dim i as integer
'
'Do While i < NumPacketsToSend
'
'    If PacketsToSend(i).lLastSend + PacketsToSendRetry < GetTickCount() Then
'        modWinsock.SendPacket lSocket, ServerSockAddr, PacketsToSend(i).sToSend
'
'        PacketsToSend(i).lLastSend = GetTickCount()
'    End If
'
'
'    If PacketsToSend(i).lDecay < GetTickCount() Then
'        RemovePacketsToSend i
'        i = i - 1
'    End If
'    i = i + 1
'Loop
'
'
'End Sub

Public Sub SetSticksWeapon(iStick As Integer, vWeapon As eWeaponTypes, Optional bSwapFireMode As Boolean = True)
Dim vSwap As eFireModes

With Stick(iStick)
    .WeaponType = vWeapon
    .Burst_Bullets = kBurstBullets(vWeapon)
    .Burst_Delay = kBurstDelay(vWeapon)
End With

If iStick = 0 And bSwapFireMode Then
    'switch FireMode vars
    vSwap = FireMode_Current
    FireMode_Current = FireMode_2ndWeapon
    FireMode_2ndWeapon = vSwap
End If

If WeaponSupportsFireMode(vWeapon, FireMode_Current) Then
    SetFireMode iStick, FireMode_Current
Else
    Set_Default_FireMode iStick
End If

End Sub

Private Sub Set_Default_FireMode(iStick As Integer)
SetFireMode iStick, CInt(-1)
End Sub

Private Sub Make_Weapon_Semi_Auto(iStick As Integer, iBullets_Burst As Integer)

'kBurstBullets(vWeapon) = iBullets_Burst
'kBurstDelay(vWeapon) = 240
With Stick(iStick)
    .Burst_Bullets = iBullets_Burst
    .Burst_Delay = 300
End With

End Sub
Private Sub Make_Weapon_Single_Shot(iStick As Integer)

'kBurstBullets(vWeapon) = 1
'kBurstDelay(vWeapon) = 350
With Stick(iStick)
    .Burst_Bullets = 1
    .Burst_Delay = 400
End With

End Sub
Private Sub Make_Weapon_Auto(iStick As Integer)

'kBurstBullets(vWeapon) = 0
'kBurstDelay(vWeapon) = 0
With Stick(iStick)
    .Burst_Bullets = 0 'skip the burst code, and rely on the weapon's default delays
    .Burst_Delay = 0
End With

End Sub
Public Sub Make_Weapon_Default_FireMode(iStick As Integer)

With Stick(iStick)
    .Burst_Bullets = kBurstBullets(Stick(iStick).WeaponType)
    .Burst_Delay = kBurstDelay(Stick(iStick).WeaponType)
End With

End Sub
Public Sub SetFireMode(iStick As Integer, vFireMode As eFireModes)

If iStick = 0 Then
    FireMode_Current = vFireMode Mod (eFireModes.Single_Shot + 1)
End If


If vFireMode = Auto Then
    Make_Weapon_Auto iStick
ElseIf vFireMode = Semi_3 Then
    Make_Weapon_Semi_Auto iStick, 3
ElseIf vFireMode = Semi_2 Then
    Make_Weapon_Semi_Auto iStick, 2
ElseIf vFireMode = Single_Shot Then
    Make_Weapon_Single_Shot iStick
Else
    Make_Weapon_Default_FireMode iStick
    
    If iStick = 0 Then
        Select Case Stick(0).Burst_Bullets
            Case 0
                FireMode_Current = Auto
            Case 1
                FireMode_Current = Single_Shot
            Case 2
                FireMode_Current = Semi_2
            Case 3
                FireMode_Current = Semi_3
            Case Else
                FireMode_Current = Auto
        End Select
    End If
End If

End Sub
Public Function GetFireModeName(vFireMode As eFireModes) As String

If vFireMode = Auto Then
    GetFireModeName = "Automatic"
ElseIf vFireMode = Semi_3 Then
    GetFireModeName = "Three Shot Burst"
ElseIf vFireMode = Semi_2 Then
    GetFireModeName = "Two Shot Burst"
ElseIf vFireMode = Single_Shot Then
    GetFireModeName = "Single Shot"
Else
    GetFireModeName = "Unknown Fire Mode" '"Default"
End If

End Function
Public Function FireModeNameToInt(sName As String) As eFireModes
Dim i As Integer

For i = 0 To eFireModes.Single_Shot
    If GetFireModeName(CInt(i)) = sName Then
        FireModeNameToInt = i
        Exit Function
    End If
Next i

FireModeNameToInt = -1

End Function
Public Function WeaponBurstable(vWeapon As eWeaponTypes) As Boolean

If WeaponIsSniper(vWeapon) Then
     Exit Function
ElseIf WeaponIsShotgun(vWeapon) Then
    Exit Function
ElseIf WeaponIsPistol(vWeapon) Then
    Exit Function
ElseIf vWeapon = RPG Then
    Exit Function
ElseIf vWeapon = FlameThrower Then
    Exit Function
ElseIf vWeapon = M249 Then
    Exit Function
End If

WeaponBurstable = True
End Function
Public Function WeaponSupportsFireMode(vWeapon As eWeaponTypes, vFireMode As eFireModes) As Boolean

'If vFireMode = Default Then
'    WeaponSupportsFireMode = True
'
'Else
    'firemode is either Auto or Single_Shot Or Semi_X
    
    Select Case vWeapon
        Case XM8
            WeaponSupportsFireMode = True 'all modes
            
        Case G3, MP5
            WeaponSupportsFireMode = (vFireMode <> Semi_3)
            
        Case AK, M249
            WeaponSupportsFireMode = (vFireMode = Auto Or vFireMode = Single_Shot)
            
        Case W1200, M82, AWM
            WeaponSupportsFireMode = (vFireMode = Single_Shot)
                
        Case SPAS
            WeaponSupportsFireMode = (vFireMode = Auto)
            
        Case Mac10, AUG
            WeaponSupportsFireMode = (vFireMode <> Semi_2)
            
            
            
        Case Else
            'DEagle
            'USP
            
            'FlameThrower
            'RPG
            
            'Knife
            'Chopper
            
            'All of the above weapons should be in Form_KeyPress to prevent infinite loop
            '(for keyascii = weaponfiremodekey)
            WeaponSupportsFireMode = False '(vfiremode = Auto)
    End Select
    
'End If


End Function

'################################################################################################

Private Sub AccurateShot(TargetX As Single, TargetY As Single, TargetSpeed As Single, _
    TargetHeading As Single, SourceX As Single, SourceY As Single, SourceSpeed As Single, _
    SourceHeading As Single, ProjectileSpeed As Single, ByRef AccurateSpeed As Single, _
    ByRef AccurateHeading As Single)

Dim DeltaX As Single
Dim DeltaY As Single
Dim DeltaSpeed As Single
Dim DeltaHeading As Single
Dim ResultX As Single
Dim ResultY As Single
Dim TResult As Single
Dim bPossible As Boolean

Dim A As Single
Dim b As Single
Dim C As Single
Dim sq As Single
Dim t1 As Single
Dim t2 As Single

'Assume it's possible
bPossible = True

'Determine the relative location of the target
DeltaX = TargetX - SourceX
DeltaY = TargetY - SourceY

'Subtract the velocity vectors to find the relative velocity
AddVectors TargetSpeed, TargetHeading, SourceSpeed, SourceHeading + Pi, DeltaSpeed, DeltaHeading

'Set up the quadratic equation's variables
A = (ProjectileSpeed ^ 2 - DeltaSpeed ^ 2)
b = -(2 * DeltaSpeed * (DeltaX * Sine(DeltaHeading) - DeltaY * CoSine(DeltaHeading)))
C = -(DeltaX ^ 2 + DeltaY ^ 2)

'Ensure there's no problem with the square root, and no divide by zero
sq = (b ^ 2) - (4 * A * C)
If (sq < 0) Or (A = 0) Then
    bPossible = False
Else
    'We're good to go, get the two results of the quadratic
    t1 = (-b - Sqr(sq)) / (2 * A)
    t2 = (-b + Sqr(sq)) / (2 * A)
    
    'Is the first Time value the optimal one?
    If t1 > 0 And t1 < t2 Then
        TResult = t1
    ElseIf t2 > 0 Then
        TResult = t2
    Else
        bPossible = False
    End If
End If


'Is there a solution?
If bPossible Then
    'Where will the target be, in TResult seconds?
    ResultX = TargetX + TargetSpeed * Sine(TargetHeading) * TResult
    ResultY = TargetY - TargetSpeed * CoSine(TargetHeading) * TResult
    
    'Return the angle to hit the target
    AccurateHeading = FindAngle_Actual(SourceX, SourceY, ResultX, ResultY)
    
    'Return the speed of the bullet (have to add the source's speed vector)
    AddVectors SourceSpeed, SourceHeading, ProjectileSpeed, AccurateHeading, AccurateSpeed
Else
    'It's not possible, just shoot straight at 'em
    AddVectors SourceSpeed, SourceHeading, ProjectileSpeed, FindAngle_Actual(SourceX, SourceY, TargetX, TargetY), _
        AccurateSpeed, AccurateHeading
End If

End Sub

Private Function RectCollision(Rect1 As RECT, Rect2 As RECT) As Boolean
Dim TempRect As RECT
RectCollision = CBool(IntersectRect(TempRect, Rect1, Rect2))
End Function
Private Function PointToRect(X As Single, Y As Single) As RECT
With PointToRect
    .Left = X
    .Right = .Left + 1
    .Top = Y
    .Bottom = .Top + 1
End With
End Function

'Private Sub ToggleFullScreen()
'
'FullScreenMode Not bFullScreened
'
'End Sub
'Private Sub FullScreenMode(Optional bFullScreen As Boolean = True)
'
'Me.WindowState = vbNormal
'
'Me.BorderStyle = IIf(bFullScreen, vbBSNone, vbSizable)
'Me.Caption = Me.Caption
'
'If bFullScreen Then
'    Me.Left = 0
'    Me.Top = 0
'    Me.width = Screen.width
'    Me.height = Screen.height
'    Me.BorderStyle = vbMaximized
'End If
'
'bFullScreened = bFullScreen
'
'End Sub

'####################################################################################

Private Sub AddNadeTrail_Simple(X As Single, Y As Single)

'ensure it's facing up, left or right regardless
If Rnd() > 0.5 Then
    AddNadeTrail X, Y, Rnd() * piD2
Else
    AddNadeTrail X, Y, pi3D2 + Rnd() * piD2
End If

End Sub

Private Sub AddNadeTrail(X As Single, Y As Single, Heading As Single)

'If modStickGame.cg_Smoke Then
    ReDim Preserve NadeTrail(NumNadeTrails)
    
    With NadeTrail(NumNadeTrails)
        .X = X
        .Y = Y
        
        .Speed = 25 + PM_Rnd() * 10
        .Heading = Heading
        
        
        '##############################
        'CHANGE THIS
        '##############################
        .lColour = RandomRGBBetween(150, 200)
        
        .SpawnTime = GetTickCount() + Rnd() * 1000
    End With
    
    NumNadeTrails = NumNadeTrails + 1
'End If

End Sub

Private Sub RemoveNadeTrail(Index As Integer)

Dim i As Integer

If NumNadeTrails = 1 Then
    Erase NadeTrail
    NumNadeTrails = 0
Else
    For i = Index To NumNadeTrails - 2
        NadeTrail(i) = NadeTrail(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve NadeTrail(NumNadeTrails - 2)
    NumNadeTrails = NumNadeTrails - 1
End If

End Sub

Private Sub ProcessNadeTrails()
Dim i As Integer

Const Speed_DeGrowth As Single = 0.8, Smoke_Speed_DeGrowth = 0.9, _
    Gravity_StrengthDX = Gravity_Strength / 4

Dim Adj As Single


Do While i < NumNadeTrails
    
    Adj = GetTimeZoneAdjust(NadeTrail(i).X, NadeTrail(i).Y)
    
    If NadeTrail(i).SpawnTime + 1000 / Adj < GetTickCount() Then
        RemoveNadeTrail i
        i = i - 1
    Else
        With NadeTrail(i)
            MotionStickObject .X, .Y, .Speed, .Heading
            ApplyGravityVector .LastGravity, Adj, .Speed, .Heading, .X, .Y, Gravity_StrengthDX
            
            
            If .LastSmoke + (20 + 40 * Rnd()) / Adj < GetTickCount() Then
                AddGravitySmoke .X, .Y, 100, .Heading, .lColour
                .LastSmoke = GetTickCount()
            End If
            
        End With
    End If
    
    
    i = i + 1
Loop

'##########################################
'processgravitysmoke
i = 0
Do While i < NumGravitySmokes
    'If GravitySmoke(i).SpawnTime + 2000 / GetTimeZoneAdjust(GravitySmoke(i).X, GravitySmoke(i).Y) < GetTickCount() Then
        'RemoveGravitySmoke i
        'i = i - 1
    'End If
    
    GravitySmoke(i).sgSize = GravitySmoke(i).sgSize - modStickGame.StickTimeFactor * modStickGame.sv_StickGameSpeed / 2
    
    If GravitySmoke(i).sgSize < 1 Or GravitySmoke(i).Y > StickGameHeight Then
        RemoveGravitySmoke i
        i = i - 1
    End If
    
    i = i + 1
Loop

End Sub

Private Sub DrawGravitySmokes()
Dim i As Integer

Const GravitySmokeCol = &HE0E0E0
'                         B G R

picMain.FillStyle = vbFSSolid

For i = 0 To NumGravitySmokes - 1
    With GravitySmoke(i)
        picMain.FillColor = .lColour
        modStickGame.sCircle .X, .Y, .sgSize, .lColour
    End With
Next i

picMain.FillStyle = vbFSTransparent

End Sub

Private Sub AddGravitySmoke(X As Single, Y As Single, Speed As Single, Heading As Single, lColour As Long)

If modStickGame.cg_ExSmoke Then
    ReDim Preserve GravitySmoke(NumGravitySmokes)
    
    With GravitySmoke(NumGravitySmokes)
        .X = X
        .Y = Y
        
        .Speed = Speed
        .Heading = Heading
        
        .lColour = lColour
        
        .sgSize = 50
        
        .SpawnTime = GetTickCount() + Rnd() * 700
    End With
    
    NumGravitySmokes = NumGravitySmokes + 1
End If

End Sub

Private Sub RemoveGravitySmoke(Index As Integer)

Dim i As Integer

If NumGravitySmokes = 1 Then
    Erase GravitySmoke
    NumGravitySmokes = 0
Else
    For i = Index To NumGravitySmokes - 2
        GravitySmoke(i) = GravitySmoke(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve GravitySmoke(NumGravitySmokes - 2)
    NumGravitySmokes = NumGravitySmokes - 1
End If

End Sub

'####################################################################################

Private Sub AddBulletTrail(X As Single, Y As Single, Heading As Single, RelSpeed As Single, _
    bTracer As Boolean)

ReDim Preserve BulletTrail(NumBulletTrails)

With BulletTrail(NumBulletTrails)
    .X = X
    .Y = Y
    
    .sgLength = def_BulletTrailLen * RelSpeed
    
    .bTracer = bTracer
    
    .Speed = BulletTrail_StartSpeed * Sgn(PM_Rnd()) * RelSpeed
    .Heading = Heading
    
    .SpawnTime = GetTickCount()
End With

NumBulletTrails = NumBulletTrails + 1

End Sub

Private Sub RemoveBulletTrail(Index As Integer)

Dim i As Integer

If NumBulletTrails = 1 Then
    Erase BulletTrail
    NumBulletTrails = 0
Else
    For i = Index To NumBulletTrails - 2
        BulletTrail(i) = BulletTrail(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve BulletTrail(NumBulletTrails - 2)
    NumBulletTrails = NumBulletTrails - 1
End If

End Sub

Private Sub ProcessBulletTrails()
Dim i As Integer

Const Speed_Growth As Single = 0.5
Dim Adj As Single



Do While i < NumBulletTrails
    
    Adj = GetTimeZoneAdjust(BulletTrail(i).X, BulletTrail(i).Y)
    
    If BulletTrail(i).SpawnTime + BulletTrail_Time / Adj < GetTickCount() Then
        RemoveBulletTrail i
        i = i - 1
    Else
        With BulletTrail(i)
            MotionStickObject .X, .Y, .Speed, .Heading
            
            .Speed = .Speed * Adj * Speed_Growth '* modStickGame.StickTimeFactor
        End With
    End If
    
    
    i = i + 1
Loop


End Sub

Public Sub SetBulletTrail_defCol()
Dim vRGB As ptRGB
Dim lDefCol As Long

vRGB = modSpaceGame.RGBDecode(modStickGame.cg_BGColour)

With vRGB
    .Red = .Red - 25
    .Blue = .Blue - 25
    .Green = .Green - 25
    
    Tracer_DefCol = RGB(.Red, .Green, .Blue)
End With

End Sub

Private Sub DrawBulletTrails()
Dim i As Integer


picMain.DrawWidth = 1
For i = 0 To NumBulletTrails - 1
    
    With BulletTrail(i)
        
        If .bTracer Then
            picMain.ForeColor = Tracer_Col
        Else
            picMain.ForeColor = Tracer_DefCol
        End If
        
        modStickGame.sLine .X, .Y, _
            .X + .sgLength * Sine(.Heading + piD2), _
            .Y - .sgLength * CoSine(.Heading + piD2)
    
    End With
    
Next i


End Sub

'####################################################################################

Private Sub AddCircleBlast(X As Single, Y As Single, IDir As Integer)

If modStickGame.cg_Smoke Then
    ReDim Preserve CircleBlast(NumCircleBlasts)
    
    With CircleBlast(NumCircleBlasts)
        .X = X
        .Y = Y
        
        .sgSize = 1
        .iDirection = IDir
    End With
    
    NumCircleBlasts = NumCircleBlasts + 1
End If

End Sub

Private Sub RemoveCircleBlast(Index As Integer)

Dim i As Integer

If NumCircleBlasts = 1 Then
    Erase CircleBlast
    NumCircleBlasts = 0
Else
    For i = Index To NumCircleBlasts - 2
        CircleBlast(i) = CircleBlast(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve CircleBlast(NumCircleBlasts - 2)
    NumCircleBlasts = NumCircleBlasts - 1
End If

End Sub

Private Sub ProcessCircleBlasts()
Dim i As Integer

Const Growth_Amount As Single = 40, DeGrowth_Amount As Single = 30

Do While i < NumCircleBlasts
    
    If CircleBlast(i).bFading = False Then
        With CircleBlast(i)
            .sgSize = .sgSize + modStickGame.StickTimeFactor * GetTimeZoneAdjust(.X, .Y) * Growth_Amount
            'f = modStickGame.StickTimeFactor * GetTimeZoneAdjust(.X, .Y)
            '.Size = .Size + 1 * f
            
            If .sgSize > CircleBlast_MaxSize Then
                .bFading = True
                .sgSize = CircleBlast_MaxSize
            End If
            
        End With
    Else
        With CircleBlast(i)
            .sgSize = .sgSize - modStickGame.StickTimeFactor * GetTimeZoneAdjust(.X, .Y) * DeGrowth_Amount
        End With
        
        If CircleBlast(i).sgSize <= 1 Then
            RemoveCircleBlast i
            i = i - 1
        End If
    End If
    
    
    i = i + 1
Loop



End Sub

Private Sub DrawCircleBlasts()
Dim i As Integer, j As Integer

picMain.FillColor = SmokeFill
picMain.FillStyle = vbFSSolid 'vbopaque
picMain.DrawWidth = 1

For i = 0 To NumCircleBlasts - 1
    With CircleBlast(i)
        If .bFading Then
            For j = 0 To 4
                modStickGame.sCircleAspect .X, .Y + .iDirection * CircleBlast_MaxSize * j / 5, .sgSize / 10 * j + 1, SmokeOutline, 0.75
            Next j
        Else
            For j = 0 To 4
                modStickGame.sCircleAspect .X, .Y + .iDirection * .sgSize * j / 5, 40 * j + 1, SmokeOutline, 0.75
            Next j
        End If
    End With
Next i

picMain.FillStyle = vbFSTransparent

End Sub

'####################################################################################

Private Sub AddGrass(X As Single, iPlatform As Integer)
Const Grass_Step = 237&, kFactor = 0.2, Max_Grass = 24
Dim Grass_Num As Integer
Dim i As Integer
Dim Platform_Right As Single

Grass_Num = Int((Rnd() + kFactor) * Max_Grass)
NumGrass = NumGrass + Grass_Num

Platform_Right = Platform(iPlatform).Left + Platform(iPlatform).width


ReDim Preserve Grass(NumGrass - 1)

For i = NumGrass - Grass_Num To NumGrass - 1  'so we get 0*Grass_Step below
    With Grass(i)
        .X = X + (i - NumGrass + Grass_Num) * Grass_Step * (Rnd() + 0.5)
        
        .iPlatform = iPlatform
        
        .RndK1 = Rnd() + kFactor
        .RndK2 = Rnd() + kFactor
        .RndK3 = Rnd() + kFactor
        .RndK4 = Rnd() + kFactor
    End With
    
    
    If Grass(i).X > Platform_Right Then
        Grass(i).X = Platform_Right
        ReDim Preserve Grass(i)
        NumGrass = i + 1
        Exit For
    End If
Next i

End Sub

Private Sub AddSingleGrass(X As Single, iPlatform As Integer)
Const kFactor = 0.2

ReDim Preserve Grass(NumGrass)

With Grass(NumGrass)
    .X = X
    
    .iPlatform = iPlatform
    
    .RndK1 = Rnd() + kFactor
    .RndK2 = Rnd() + kFactor
    .RndK3 = Rnd() + kFactor
    .RndK4 = Rnd() + kFactor
    
End With

NumGrass = NumGrass + 1

End Sub

Private Sub RemoveGrass(ByVal Index As Integer)
Dim i As Integer

If NumGrass = 1 Then
    Erase Grass
    NumGrass = 0
Else
    For i = Index To NumGrass - 2
        Grass(i) = Grass(i + 1)
    Next i
    
    ReDim Preserve Grass(NumGrass - 2)
    NumGrass = NumGrass - 1
End If

End Sub

Private Sub DrawGrass()
Dim i As Integer

Const Grass_Sep1 = 50, Grass_Sep2 = -150, Grass_Sep3 = 90, Grass_Sep4 = -40
Const Grass_Height1 = 190, Grass_Height2 = 110, Grass_Height3 = 86, Grass_Height4 = 150

picMain.ForeColor = Grass_Col
picMain.DrawWidth = 2

For i = 0 To NumGrass - 1
    With Grass(i)
        modStickGame.sLine .X, Platform(.iPlatform).Top, _
                           .X + Grass_Sep1 * .RndK1, Platform(.iPlatform).Top - Grass_Height1 * .RndK4
        
        modStickGame.sLine .X, Platform(.iPlatform).Top, _
                           .X + Grass_Sep2 * .RndK2, Platform(.iPlatform).Top - Grass_Height2 * .RndK3
                
        modStickGame.sLine .X, Platform(.iPlatform).Top, _
                           .X + Grass_Sep3 * .RndK3, Platform(.iPlatform).Top - Grass_Height3 * .RndK2
        
        modStickGame.sLine .X, Platform(.iPlatform).Top, _
                           .X + Grass_Sep4 * .RndK4, Platform(.iPlatform).Top - Grass_Height4 * .RndK1
        
    End With
Next i


End Sub

Private Sub MakeGrass()
Dim iPlatform As Integer
Dim i As Integer


For iPlatform = 0 To ubdPlatforms

    For i = 1 To 1 + Rnd() * 1.4 * Platform(iPlatform).width / Platform(0).width
        AddGrass RandomXOnPlatform(iPlatform), iPlatform
    Next i

Next iPlatform


End Sub

'####################################################################################

Private Sub ProcessAmmoPickups()
Dim i As Integer, bTold As Boolean

GenerateAmmoPickup

If Stick(0).WeaponType < Knife Then
    
    For i = 0 To NumAmmoPickUpsM1
        If AmmoPickup(i).bActive Then
            If CoOrdInStick(AmmoPickup(i).X, AmmoPickup(i).Y, 0) Then
                If TotalMags(Stick(0).WeaponType) < GetTotalMags(Stick(0).WeaponType) Then
                    
                    If StickiHasState(0, STICK_CROUCH) Then
                        
                        TotalMags(Stick(0).WeaponType) = GetTotalMags(Stick(0).WeaponType) + IIf(TotalMags(Stick(0).WeaponType) = 0 And Stick(0).BulletsFired = GetMaxRounds(Stick(0).WeaponType), 1, 0)
                        
                        modAudio.PlayWeaponPickUpSound
                        
                        AmmoPickup(i).LastUsed = GetTickCount()
                        AmmoPickup(i).bActive = False
                        
                    ElseIf bTold = False Then
                        
                        PrintStickText "Crouch to Pick Up " & GetMagName(Stick(0).WeaponType), Stick(0).X + 750, Stick(0).Y - 500, vbRed
                        
                        bTold = True
                    End If
                    
                    
                End If
            End If
        End If
    Next i
    
End If


End Sub

Private Sub GenerateAmmoPickup(Optional bForce As Boolean = False)
Dim i As Integer, iPlatform As Integer
Dim GTC As Long
Dim bCan As Boolean
Const LR_Lim = 500

GTC = GetTickCount()

For i = 0 To NumAmmoPickUpsM1
    bCan = False
    
    If bForce Then
        bCan = True
    ElseIf AmmoPickup(i).bActive = False Then
        If AmmoPickup(i).LastUsed + AmmoPickUp_Spawn_Delay < GTC Then
            bCan = True
        End If
    End If
    
    
    If bCan Then
        'Do
            iPlatform = GetRandomPlatform()
        'Loop While iPlatform = 0 And modStickGame.ubdPlatforms > 0
        
        
        AmmoPickup(i).Y = YOnPlatform(iPlatform)
        AmmoPickup(i).X = RandomXOnPlatform(iPlatform)
        
        If AmmoPickup(i).X < Platform(iPlatform).Left + LR_Lim Then
            AmmoPickup(i).X = Platform(iPlatform).Left + LR_Lim
        ElseIf AmmoPickup(i).X > Platform(iPlatform).Left + Platform(iPlatform).width Then
            AmmoPickup(i).X = Platform(iPlatform).Left + Platform(iPlatform).width - LR_Lim
        End If
        
        
        AmmoPickup(i).bActive = True
        AmmoPickup(i).LastUsed = GTC
    End If
Next i

End Sub

Private Sub DrawAmmoPickups()
Dim i As Integer
Const AmmoPickup_Radius = HealthPack_Radius * 2


picMain.FillStyle = vbFSSolid
picMain.FillColor = vbBlack
picMain.DrawWidth = 2
For i = 0 To NumAmmoPickUpsM1
    If AmmoPickup(i).bActive Then
        
        modStickGame.sBox AmmoPickup(i).X, AmmoPickup(i).Y, _
            AmmoPickup(i).X + AmmoPickup_Radius, AmmoPickup(i).Y + HealthPack_Radius, _
            vbBlack
        
        
        modStickGame.PrintStickText "Ammo Pickup", AmmoPickup(i).X - 300, AmmoPickup(i).Y - 200, vbBlack
        
    End If
Next i

picMain.FillStyle = vbFSTransparent

End Sub

Private Function GetTotalMags(vWeapon As eWeaponTypes) As Byte
'code can be slow here

Select Case vWeapon
    Case AK
        GetTotalMags = AK_Mags '5
    Case DEagle
        GetTotalMags = DEagle_Mags '8
    Case FlameThrower
        GetTotalMags = Flamethrower_Mags '5
    Case M249
        GetTotalMags = M249_Mags '3
    Case M82
        GetTotalMags = M82_Mags '6
    Case RPG
        GetTotalMags = RPG_Mags '4
    Case AUG
        GetTotalMags = AUG_Mags '5
    Case W1200
        GetTotalMags = W1200_Mags '12
    Case XM8
        GetTotalMags = XM8_Mags '5
    Case USP
        GetTotalMags = USP_Mags
    Case AWM
        GetTotalMags = AWM_Mags
    Case MP5
        GetTotalMags = MP5_Mags
    Case Mac10
        GetTotalMags = Mac10_Mags
    Case SPAS
        GetTotalMags = SPAS_Mags
    Case G3
        GetTotalMags = G3_Mags
End Select

End Function

Private Function GetMagName(vWeapon As eWeaponTypes) As String

If vWeapon = RPG Then
    GetMagName = "Rockets"
ElseIf vWeapon = FlameThrower Then
    GetMagName = "Canisters"
ElseIf WeaponIsShotgun(vWeapon) Then
    GetMagName = "Shells"
Else
    GetMagName = "Magazines"
End If

End Function

Private Function PM_Rnd() As Single
PM_Rnd = (Rnd() - Rnd())
End Function

Private Sub DrawWallMarks()
Dim i As Integer
Dim GTC As Long

picMain.FillStyle = vbFSSolid
picMain.FillColor = modStickGame.cg_BGColour 'WallMark_Colour

GTC = GetTickCount()

Do While i < NumWallMarks
    
    modStickGame.sCircle WallMark(i).X, WallMark(i).Y, WallMark(i).Radius, modStickGame.cg_BGColour
    
    
    If WallMark(i).Decay < GTC Then
        RemoveWallMark i
        i = i - 1
    End If
    
    
    i = i + 1
Loop
picMain.FillStyle = vbFSTransparent

End Sub

Private Sub AddWallMark(X As Single, Y As Single, Radius As Single)
Const MaxWM = 300

If modStickGame.cg_WallMarks Then
    
    If NumWallMarks > MaxWM Then
        RemoveWallMark 0
    End If
    
    
    ReDim Preserve WallMark(NumWallMarks)
    
    With WallMark(NumWallMarks)
        .Decay = GetTickCount() + WallMark_Time / modStickGame.sv_StickGameSpeed
        .X = X
        .Y = Y
        .Radius = Radius
    End With
    
    NumWallMarks = NumWallMarks + 1
End If

End Sub

Private Sub RemoveWallMark(Index As Integer)

Dim i As Integer

If NumWallMarks = 1 Then
    Erase WallMark
    NumWallMarks = 0
Else
    For i = Index To NumWallMarks - 2
        WallMark(i) = WallMark(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve WallMark(NumWallMarks - 2)
    NumWallMarks = NumWallMarks - 1
End If

End Sub

Private Sub SpawnHealthPack(ByVal X As Single, ByVal Y As Single)
HealthPack.X = X
HealthPack.Y = Y
HealthPack.bActive = True
End Sub

Private Sub GenerateHealthPack()

If modStickGame.sv_GameType <> gCoOp Then
    If HealthPack.LastUsed + HealthPackDelay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        
        SpawnHealthPack HealthPackX, HealthPackY
        
        SendBroadcast sHealthPacks & CStr(HealthPack.X) & "|" & CStr(HealthPack.Y)
        
        HealthPack.LastUsed = GetTickCount()
    End If
End If

End Sub

Private Sub DisplayHealthPack()
Const HealthPack_RadiusX2 = HealthPack_Radius * 2

If HealthPack.bActive Then
    
    If modStickGame.sv_GameType = gCoOp Then
        HealthPack.bActive = False
    Else
        picMain.FillStyle = vbFSSolid
        picMain.FillColor = vbWhite
        sBox HealthPack.X, HealthPack.Y, HealthPack.X + HealthPack_RadiusX2, HealthPack.Y + HealthPack_Radius, vbRed
        picMain.FillStyle = vbFSTransparent
        
        
        PrintStickText "Health Pack!", HealthPack.X - 400, HealthPack.Y - 200, vbBlack
    End If
    
End If

End Sub

Public Function StickInGame(ByVal iStick As Integer) As Boolean
StickInGame = Stick(iStick).Team <> Spec And Stick(iStick).bAlive
End Function

Private Function IsAlly(ByVal t1 As eTeams, ByVal t2 As eTeams) As Boolean
IsAlly = (Not (t1 = Neutral Or t2 = Neutral)) And (t1 = t2)
End Function

Private Function ForeignStick(i As Integer) As Boolean
If Stick(i).IsBot = False Then
    ForeignStick = i > 0
End If
End Function

Private Function GetTeamColour(ByVal vTeam As eTeams) As Long
GetTeamColour = kTeamColour(vTeam)
End Function

Private Sub MakeTeamColourArray()
Dim i As Integer

For i = 0 To eTeams.Spec
    Select Case i
        Case eTeams.Neutral
            kTeamColour(i) = MGrey
        Case eTeams.Red
            kTeamColour(i) = vbRed
        Case eTeams.Blue
            kTeamColour(i) = vbBlue
        Case eTeams.Spec
            kTeamColour(i) = MGrey
    End Select
Next i

End Sub

Private Sub CheckKillsInARow()

If Stick(0).iKillsInARow = RowKillsForShield Then
    'add shield
    If Stick(0).Shield = 0 And Stick(0).WeaponType <> Chopper Then
        Stick(0).Shield = 1 'start charging
        AddMainMessage CStr(RowKillsForShield) & " kills - Shield Acquired", False, vbBlack
    End If
    
'ElseIf Stick(0).iKillsInARow = RowKillsForRadar Then
'    FullRadarStartTime = GetTickCount()
'    bHadFullRadar = True
'    AddMainMessage "Radar Active for 30 Seconds"
    
ElseIf Stick(0).iKillsInARow = RowKillsForChopper Then
    If modStickGame.sv_GameType <> gCoOp Then
        If Stick(0).WeaponType <> Chopper Then
            ChopperAvail = True
            AddMainMessage "Chopper Available, Press 0", True, vbBlack
        End If
    End If
    
End If

If FlamesInARow = RowFlameKillsForToasty Then
    
    If Stick(0).WeaponType = FlameThrower Then
        AddMainMessage "TOASTY! (Not 3D)", True
        'picToasty.Visible = True
        ShowToasty
    End If
    
ElseIf FlamesInARow >= RowFlameKillsForToasty Then
    FlamesInARow = 0
    'picToasty.Visible = False
End If

If Stick(0).WeaponType = Knife Then
    KnifesInARow = KnifesInARow + 1
    
    If KnifesInARow = KnivesForSaber Then
        If Stick(0).Perk <> pZombie Then
            If Stick(0).bLightSaber = False Then
                Stick(0).bLightSaber = True
                AddMainMessage "Lightsaber Acquired. Hold Vertically to Block Bullets", True
            End If
        End If
    End If
    
Else
    KnifesInARow = 0
End If

End Sub

Public Sub BltToForm()

BitBlt Me.hDC, 0, 0, ScaleX(Me.width, vbTwips, vbPixels), ScaleY(Me.height, vbTwips, vbPixels), _
    Me.picMain.hDC, 0, 0, modStickGame.cg_DisplayMode

'BitBlt Me.hdc, 0, 0, ScaleX(StickGameWidth, vbTwips, vbPixels), ScaleY(StickGameHeight, vbTwips, vbPixels), _
    Me.picMain.hdc, 0, 0, modStickGame.cg_DisplayMode

'vbNotSrcCopy
'vbSrcCopy

'RasterOpConstants
End Sub

Public Sub SwitchWeapon(ByVal vWeapon As eWeaponTypes, Optional bSwitchSilenced As Boolean = True)
Const RPG_Bullet_DelayX3 = RPG_Bullet_Delay * 3
Static bTold As Boolean
Dim bSilenced As Boolean

If Stick(0).Perk = pZombie Then
    If Stick(0).WeaponType <> Knife Then
        Stick(0).WeaponType = Knife
    End If
    Exit Sub
End If

If Stick(0).WeaponType <> vWeapon Then
    
    If vWeapon <> -1 Then
        If vWeapon <> Chopper Then
            
            If Stick(0).WeaponType <= Knife Then
                AmmoFired(Stick(0).WeaponType) = Stick(0).BulletsFired
            End If
            
            
            'Stick(0).PrevWeapon = Stick(0).WeaponType
            
            On Error GoTo EH
            Stick(0).BulletsFired = AmmoFired(vWeapon)
            
        Else
            'ChopperAvail = False
            modStickGame.cg_LaserSight = False
            Stick(0).Health = Health_Start
            Stick(0).Shield = 0
        End If
        
        
        If StickiHasState(0, STICK_RELOAD) Then
            frmStickGame.SubStickiState 0, STICK_RELOAD
        End If
        If StickiHasState(0, STICK_FIRE) Then
            frmStickGame.SubStickiState 0, STICK_FIRE
        End If
        
        modAudio.StopWeaponReloadSound Stick(0).WeaponType
        
        If bSwitchSilenced Then
            'switch silenced bools
            bSilenced = Stick(0).bSilenced
            Stick(0).bSilenced = b2ndWeaponSilenced
            b2ndWeaponSilenced = bSilenced
        End If
        
        'SWITCH IS HERE #################################################
        'Stick(0).WeaponType = vWeapon
        SetSticksWeapon 0, vWeapon, bSwitchSilenced
        'SWITCH IS HERE #################################################
        
        
        'Scroll_WeaponKey = vWeapon
        WeaponKey = -1
        
        If WeaponIsSniper(vWeapon) Then
            If Stick(0).Perk <> pSniper Then
                If Not bTold Then
                    AddMainMessage "You must be crouched or prone to snipe", True
                    bTold = True
                End If
            End If
        End If
        
        
        Stick(0).LastBullet = GetTickCount() - 1000000 '/ GetMyTimeZone()
        Stick(0).LastMuzzleFlash = Stick(0).LastBullet - 1000 'turn it off
        
        If modStickGame.StickOptionFormLoaded Then
            frmStickOptions.chkShh.Value = IIf(WeaponSilencable(vWeapon), IIf(Stick(0).bSilenced, 1, 0), 0)
        End If
        
        If picToasty.Visible Then
            If vWeapon <> FlameThrower Then
                picToasty.Visible = False
            End If
        End If
        
        Stick(0).LastWeaponSwitch = GetTickCount()
        
    End If
End If

EH:
End Sub

Public Sub MakeZombie(iStick As Integer, Optional bZombie As Boolean = True)
'already set perk

If bZombie Then
    If iStick = 0 Then
        SwitchWeapon Knife
        
        Current_Health_Start = Zombie_Health
    Else
        Stick(iStick).WeaponType = Knife
    End If
    
    Stick(iStick).Health = Zombie_Health * (Stick(iStick).Health / 100)
    'force it to do the division first, otherwise OVERFLOW
    
    Stick(iStick).Shield = 0
    
    Stick(iStick).Perk = pZombie 'just in case
Else
    If iStick = 0 Then
        Current_Health_Start = Health_Start
    Else
        Stick(iStick).WeaponType = AK
    End If
    
    Stick(iStick).Health = 100 * (Stick(iStick).Health / Zombie_Health)
    'force it to do the division first, otherwise OVERFLOW
End If

End Sub

Private Sub CheckCapsLock()

If modKeys.Caps() Then
    AddMainMessage "(Caps Lock is on - Some keys won't work)", True
End If

End Sub

'Private Sub PrepareWeaponSelection()
'Dim i as integer
'
'For i = 0 To eWeaponTypes.Knife
'    cboWeapon.AddItem GetWeaponName(CInt(i))
'Next i
'
'cboWeapon.ListIndex = 0
'
'End Sub
'
'Private Sub cmdGo_Click()
'Dim i as integer
'
'For i = 0 To eWeaponTypes.Knife
'    If GetWeaponName(CInt(i)) = cboWeapon.Text Then
'        SwitchWeapon i
'        Exit For
'    End If
'Next i
'
'cmdGo.Visible = False
'cboWeapon.Visible = False
'
'tmrMain.Enabled = True
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim CentreX As Single, CentreY As Single


On Error GoTo EH

If KeyCode = 20 Then
    CheckCapsLock
End If

If bChatActive Then Exit Sub

If KeyCode = vbKeyF1 Then
    ShowScoresKey = True
Else
    If StickInGame(0) And bPlaying Then
        
        If Stick(0).WeaponType = Chopper Then
            If KeyCode = vbKeySpace Then
                Stick(0).Speed = Stick(0).Speed / 2
                Exit Sub
            End If
        End If
        
        
        Select Case KeyCode
            Case vbKeySpace, vbKeyW
                'JumpKey = Stick(0).OnSurface
                
                If Stick(0).bOnSurface Then
                    If Stick(0).LastMine + 500 < GetTickCount() Then
                        'don't let them jump immediatly - let clients place the mine too
                        
                        'If Stick(0).StartJumpTime + JumpTime < GetTickCount() Then
                            AddStickiState 0, STICK_JUMP
                            
                            If Stick(0).WeaponType <> Chopper Then
                                Stick(0).Y = Stick(0).Y - 50
                            End If
                            'Stick(0).StartJumpTime = GetTickCount()
                            
                            'JumpKey = False
                        'End If
                    End If
                End If
                
                
            Case vbKeyA
                If Stick(0).bOnSurface Then
                    AddStickiState 0, STICK_LEFT
                End If
                
            Case vbKeyD
                If Stick(0).bOnSurface Then
                    AddStickiState 0, STICK_RIGHT
                End If
                
                
            Case vbKeyControl, vbKeyS
                
                If Stick(0).bOnSurface Then
                    
                    If LastCrouchToggle + 200 < GetTickCount() Then
                        If StickiHasState(0, STICK_PRONE) Then
                            SubStickiState 0, STICK_PRONE
                        End If
                        
                        If modStickGame.cl_ToggleCrouch Then
                            If StickiHasState(0, STICK_CROUCH) Then
                                SubStickiState 0, STICK_CROUCH
                                CrouchKey = False
                            Else
                                AddStickiState 0, STICK_CROUCH
                                CrouchKey = True
                            End If
                        Else
                            AddStickiState 0, STICK_CROUCH
                            CrouchKey = True
                        End If
                        LastCrouchToggle = GetTickCount()
                    End If
                    
                End If
                
            Case vbKeyR
                
                If Stick(0).WeaponType < Knife Then
                    If TotalMags(Stick(0).WeaponType) > 0 Then
                        If StickiHasState(0, STICK_RELOAD) = False Then
                            If Stick(0).BulletsFired > 0 Then
                                'Debug.Print "Reload Pressed"
                                StartReload 0
                            End If
                        End If
                    End If
                End If
                
                
            Case vbKeyK, vbKeyV
                
                SwitchWeapon Knife
                
        End Select
        
        
        CheckM82Fire
    Else
        Select Case KeyCode
            Case vbKeyW
                SpecUp = True
                
            Case vbKeyA
                SpecLeft = True
                
            Case vbKeyD
                SpecRight = True
                
            Case vbKeyS
                SpecDown = True
                                
            Case vbKeyAdd
                
                If cg_sZoom < MaxZoom Then
                    CentreX = modStickGame.cg_sCamera.X + frmStickGame.ScaleWidth / cg_sZoom
                    CentreY = modStickGame.cg_sCamera.Y + frmStickGame.ScaleHeight / cg_sZoom
                    
                    cg_sZoom = Round(cg_sZoom + ZoomInc, 2)
                    
                    modStickGame.cg_sCamera.X = CentreX - frmStickGame.ScaleWidth / cg_sZoom
                    modStickGame.cg_sCamera.Y = CentreY - frmStickGame.ScaleHeight / cg_sZoom
                End If
                LastZoomPress = GetTickCount()
            
            Case vbKeySubtract
                
                If cg_sZoom >= MinZoom Then
                    CentreX = modStickGame.cg_sCamera.X + frmStickGame.ScaleWidth / cg_sZoom
                    CentreY = modStickGame.cg_sCamera.Y + frmStickGame.ScaleHeight / cg_sZoom
                    
                    cg_sZoom = Round(cg_sZoom - ZoomInc, 2)
                    
                    modStickGame.cg_sCamera.X = CentreX - frmStickGame.ScaleWidth / cg_sZoom
                    modStickGame.cg_sCamera.Y = CentreY - frmStickGame.ScaleHeight / cg_sZoom
                End If
                LastZoomPress = GetTickCount()
                
            Case vbKeyMultiply
                
                cg_sZoom = 1
                LastZoomPress = GetTickCount()
                
        End Select
    End If
End If

EH:
End Sub

Private Sub CheckM82Fire()
If Stick(0).WeaponType = M82 Then
    If Stick(0).Perk <> pSniper Then
        If StickiHasState(0, STICK_PRONE) = False Then
            If StickiHasState(0, STICK_CROUCH) = False Then
                SubStickiState 0, STICK_FIRE
            End If
        End If
    End If
End If
End Sub

Private Function GetNadeTypeName() As String
GetNadeTypeName = kNadeName(Stick(0).iNadeType)
End Function

Private Sub MakeNadeNameArray()
Dim i As Integer

For i = 0 To UBound(kNadeName)
    If i = nFrag Then
        kNadeName(i) = "Frag"
    ElseIf i = nFlash Then
        kNadeName(i) = "Flash Bang"
    ElseIf i = nTime Then
        kNadeName(i) = "Time"
    ElseIf i = nGravity Then
        kNadeName(i) = "Gravity"
    Else
        kNadeName(i) = "EMP"
    End If
Next i

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim bCan As Boolean
Dim i As Integer
Dim vMode As eFireModes


On Error GoTo EH

If KeyAscii = vbKeyTab Then
    
    If NumSticks Then
        Unload frmStickOptions
        Load frmStickOptions
        frmStickOptions.Show vbModeless, frmStickGame
    End If
    
Else
    
    bCan = (bChatActive = False And StickInGame(0))
    
    Select Case True
        
        'Case KeyAscii = 114
            'AddFlame Stick(0).X, Stick(0).Y, Stick(0).Facing, 30, 0
        
        
        'Case KeyAscii = 112 And bCan
            'AddLargeSmoke Stick(0).X, Stick(0).Y
        
        
        Case (KeyAscii = vbKey2 Or KeyAscii = vbKey1) And bCan
            If LastWeaponSwitch + KeyPressDelay < GetTickCount() Then
                If Stick(0).WeaponType <> Chopper Then
                    
                    If Stick(0).WeaponType = Stick(0).CurrentWeapons(1) Then
                        i = 2
                    Else
                        i = 1
                    End If
                    
                    
                    If modStickGame.sv_AllowedWeapons(Stick(0).CurrentWeapons(i)) Then
                        SwitchWeapon Stick(0).CurrentWeapons(i)
                    End If
                End If
                
                LastWeaponSwitch = GetTickCount()
            End If
            
        Case KeyAscii = vbKey0 And bCan
            
            If ChopperAvail Then
                If modStickGame.sv_AllowedWeapons(eWeaponTypes.Chopper) Then
                    SubStickiState 0, STICK_CROUCH
                    CrouchKey = False
                    
                    SwitchWeapon Chopper
                Else
                    ChopperAvail = False
                End If
            End If
            
        
        Case KeyAscii = 101 And bCan
            'E - vbKeyE
            For i = 0 To NumStaticWeapons - 1
                If StickNearStaticWeapon(0, i) Then
                    UseKey = True
                    Exit For
                End If
            Next i
            
            
        'vbKeyV
        Case (KeyAscii = 98 Or KeyAscii = vbKey3) And Not bChatActive
            'switch nade type
            
            If LastNadeSwitch + KeyPressDelay < GetTickCount() Then
                Stick(0).iNadeType = CInt(Stick(0).iNadeType) + 1
                If Stick(0).iNadeType > nEMP Then
                    Stick(0).iNadeType = nFrag
                End If
                
                AddMainMessage "Grenade Type: " & GetNadeTypeName(), True
                
                LastNadeSwitch = GetTickCount()
            End If
            
            
            
        Case (KeyAscii = vbKey4) And Not bChatActive
            'switch firemode
            
            If Not (Stick(0).WeaponType = DEagle Or Stick(0).WeaponType = USP Or _
                    Stick(0).WeaponType = FlameThrower Or Stick(0).WeaponType = RPG Or _
                    Stick(0).WeaponType = Chopper Or Stick(0).WeaponType = Knife) Then
                
                
                If LastFireModeSwitch + KeyPressDelay < GetTickCount() Then
                    
                    vMode = FireMode_Current
                    
                    Do
                        vMode = (vMode + 1) Mod (eFireModes.Single_Shot + 1)
                    Loop While WeaponSupportsFireMode(Stick(0).WeaponType, vMode) = False
                    
                    SetFireMode 0, vMode
                    
                    AddMainMessage "Fire Mode: " & GetFireModeName(FireMode_Current), False
                    LastFireModeSwitch = GetTickCount()
                End If
            End If
            
            
            
        
        'vbKeyP
        Case (KeyAscii = 112) And modStickGame.StickServer And Not bChatActive
            
            modStickGame.sv_AIShoot = Not modStickGame.sv_AIShoot
            AddMainMessage "AI can" & IIf(modStickGame.sv_AIShoot, vbNullString, "'t") & " shoot", True
            
        Case (KeyAscii = 111) And Not bChatActive And modDev.getDevLevel() >= modDev.Dev_Level_Heightened
            'O
            Stick(0).ShieldCharging = False
            Stick(0).Shield = 1
            Stick(0).LastShieldHitTime = 0
        
        Case (KeyAscii = 105) And Not bChatActive And modDev.getDevLevel() >= modDev.Dev_Level_Heightened
            'I
            RemoveSticksShield 0
            
            
        Case (KeyAscii = 117) And Not bChatActive And modDev.getDevLevel() >= modDev.Dev_Level_Heightened
            'I
            Stick(0).LastMine = 0
        
        '109 = vbkeyM
        Case KeyAscii = 109 And bCan 'vbKeyZ
            AddStickiState 0, STICK_MINE
            
            
        
        '################################ PRONE ################################
        Case (KeyAscii = 99 Or KeyAscii = vbKeyC Or KeyAscii = 102 Or KeyAscii = vbKeyF) And bCan    'vbKeyC
            If Stick(0).WeaponType <> Chopper Then
                
                If Stick(0).bOnSurface Then
                    If LastProneSwitch + KeyPressDelay < GetTickCount() Then
                        If StickiHasState(0, STICK_PRONE) Then
                            SubStickiState 0, STICK_PRONE
                        Else
                            AddStickiState 0, STICK_PRONE
                            
                            CrouchKey = False
                            SubStickiState 0, STICK_CROUCH
                            
                        End If
                        
                        LastProneSwitch = GetTickCount()
                    End If
                End If
            End If
        '############################## END PRONE ##############################
            
            
        Case (KeyAscii = vbKeyZ Or KeyAscii = 122 Or KeyAscii = 26) And bCan
            modStickGame.cg_LaserSight = Not modStickGame.cg_LaserSight
            
            If modStickGame.StickOptionFormLoaded Then
                frmStickOptions.chkLaserSight.Value = Abs(modStickGame.cg_LaserSight)
            End If
            
            
        Case (KeyAscii = vbKeyQ Or KeyAscii = 113 Or KeyAscii = 17) And bCan
            
            If StickiHasState(0, STICK_FIRE) = False Then
                If WeaponSilencable(Stick(0).WeaponType) Then
                    Stick(0).bSilenced = Not Stick(0).bSilenced
                    
                    'Stick(0).LastMuzzleFlash = GetTickCount() - MFlash_Time / 0.001 - 1
                    ResetTimeLong Stick(0).LastMuzzleFlash, MFlash_Time
                    
                    If modStickGame.StickOptionFormLoaded Then
                        frmStickOptions.chkShh.Value = Abs(Stick(0).bSilenced)
                    End If
                End If
            End If
            
            
            
        Case KeyAscii = vbKeySpace And bPlaying = False And _
                modStickGame.StickServer And bChatActive = False
            
            StopPlay False
            
            
        '#########################################
        'Chat handling
        'Escape kills the chat
        Case KeyAscii = vbKeyEscape
            bChatActive = False
            Stick(0).bTyping = False
            strChat = vbNullString
            
            
            'T TO TALK
        Case ((KeyAscii = 116) Or (KeyAscii = 84)) And (bChatActive = False)
            '116=t
            bChatActive = True
            Stick(0).bTyping = True
            
            'disable movement
            SubStickiState 0, STICK_RIGHT
            SubStickiState 0, STICK_LEFT
            SubStickiState 0, STICK_JUMP
            
            
        Case KeyAscii = vbKeyBack
            If LenB(strChat) Then
                strChat = Left$(strChat, Len(strChat) - 1)
            End If
            
            'Return finishes and sends the chat
        Case KeyAscii = vbKeyReturn
            
            If bChatActive Then
                'Send it!
                
                strChat = Trim$(RemoveChars(strChat))
                
                If Left$(strChat, 1) = "/" Then
                    ProcessConsoleCmd Mid$(strChat, 2)
                ElseIf LenB(strChat) Then
                    SendChatPacket Trim$(Stick(0).Name) & modMessaging.MsgNameSeparator & strChat, Stick(0).colour
                End If
                
                'Reset
                bChatActive = False
                Stick(0).bTyping = False
                strChat = vbNullString
            End If
            
        Case Else
            'If chat is on, add keystroke to chat text
            If bChatActive Then
                If KeyAscii > 31 Then
                    If LenB(strChat) < 150 Then
                        strChat = strChat & Chr$(KeyAscii)
                    End If
                End If
            End If
        '#########################################
    End Select
End If

EH:
End Sub

Private Sub ProcessConsoleCmd(ByVal sCmd As String)
Dim sCommand As String, sParam As String
Dim i As Integer, sing As Single

On Error Resume Next
i = InStr(1, sCmd, vbSpace)
If i Then
    sCommand = Left$(sCmd, i - 1)
    sParam = Mid$(sCmd, i + 1)
Else
    sCommand = sCmd
End If


Select Case LCase$(sCommand)
    Case "disconnect"
        Unload Me
        
    Case "team"
        Select Case LCase$(sParam)
            Case "red"
                Stick(0).Team = Red
            Case "blue"
                Stick(0).Team = Blue
            Case "neutral"
                Stick(0).Team = Neutral
            Case "spec", "spectator"
                Stick(0).Team = Spec
        End Select
        
    Case "gamespeed"
        If modStickGame.StickServer Then
            sing = val(sParam)
            If Err.Number = 0 Then
                
                If sing >= 0.1 And sing <= 1.2 Then
                    StickGameSpeedChanged modStickGame.sv_StickGameSpeed, sing
                    modStickGame.sv_StickGameSpeed = sing
                Else
                    AddMainMessage "Game Speed must be between 0.1 and 1.2", True
                End If
                
            Else
                AddMainMessage "Game Speed must be a number", True
            End If
        Else
            AddMainMessage "Only the server can change the game speed", True
        End If
        
    Case Else
        AddMainMessage "Please enter a valid console command", True
End Select

End Sub

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF1 Then
    ShowScoresKey = False
Else
    
    'If KeyCode = vbKeyF11 Then
        'ToggleFullScreen
    'Else
        If StickInGame(0) And bPlaying Then
            Select Case KeyCode
                Case vbKeySpace, vbKeyW
                    SubStickiState 0, STICK_JUMP
                    
                Case vbKeyA
                    'If Stick(0).OnSurface Then
                        SubStickiState 0, STICK_LEFT
                    'End If
                    
                Case vbKeyD
                    'If Stick(0).OnSurface Then
                        SubStickiState 0, STICK_RIGHT
                    'End If
                    
                Case vbKeyControl, vbKeyS
                    If Not modStickGame.cl_ToggleCrouch Or Stick(0).WeaponType = Chopper Then
                        SubStickiState 0, STICK_CROUCH
                        CrouchKey = False
                    End If
                    
            End Select
            
            CheckM82Fire
        Else
            Select Case KeyCode
                Case vbKeyW
                    SpecUp = False
                    
                Case vbKeyA
                    SpecLeft = False
                    
                Case vbKeyD
                    SpecRight = False
                    
                Case vbKeyS
                    SpecDown = False
            End Select
        End If
    'End If
End If

End Sub

Private Sub SetStickiState(i As Integer, state As eStickStates)

Stick(i).state = state

End Sub

'Private Sub AddStickState(ID As Integer, State As eStickStates)
'Dim i as integer
'
'i = FindStick(ID)
'
''Find the specified Stick and add to his state
'Stick(i).State = (Stick(i).State Or State)
'
'End Sub
'
'Public Sub SubStickState(ID As Integer, State As eStickStates)
'Dim i as integer
'
'i = FindStick(ID)
'
''Find the specified Stick and subtract from his state
'Stick(i).State = (Stick(i).State And (Not State))
'
'End Sub

Private Sub AddStickiState(i As Integer, state As eStickStates)
'Dim ls As eStickStates
'ls = Stick(i).State

Stick(i).state = (Stick(i).state Or state)

'If ls <> Stick(i).State Then
'    If i = 0 Then
'        Debug.Print "State Added " & Rnd()
'    End If
'End If

End Sub

Public Sub SubStickiState(i As Integer, state As eStickStates)

Stick(i).state = (Stick(i).state And (Not state))

End Sub

'Private Function StickHasState(ID As Integer, vState As eStickStates) As Boolean
'
'StickHasState = CBool((Stick(FindStick(ID)).State And vState))
'
'End Function

Private Function StickiHasState(Index As Integer, vState As eStickStates) As Boolean

StickiHasState = CBool((Stick(Index).state And vState) = vState)

End Function

Private Sub Form_Load()

Dim i As Integer

If modVars.Closing Then
    Unload Me
    Exit Sub
End If

modStickGame.StickFormLoaded = True

Me.Left = 500  'Screen.width / 2 - Me.width / 2
Me.Top = Screen.height / 2 - Me.height / 2
'Me.Top = frmMain.Top + frmMain.height / 2 - Me.height / 2
Me.BackColor = modStickGame.cg_BGColour

picToasty.Left = Me.width / 2 - picToasty.width / 2
picToasty.Visible = False

modStickGame.sv_StickGameSpeed = 1
WindowClosing = False
MouseX = 15915
MouseY = 3435
picMain.BackColor = Me.BackColor
picMain.Visible = False

If Not modStickGame.bStickEditing Then Me.WindowState = vbMaximized

If modStickGame.bStickEditing Then
    Call FormLoad(Me, , , False)
    StartEdit
    Me.Caption = "Stick Map Editor"
    'Show
    Me.ZOrder vbBringToFront
Else
    Me.Caption = "Stick Shooter"
    
    
    Call FormLoad(Me, , , False, True)
    
    'Display the form
    Show 'vbModeless, frmMain
    PrintLoadingText "Loading Stick Game..."
    Me.ZOrder vbBringToFront
    
    If modStickGame.cl_Subclass Then
        If bIsIDE Or bDevMode Then
            modStickGame.cl_Subclass = False
        Else
            modSubClass.SubClassStick Me.hWnd
        End If
    End If
    
    
    'Call PrepareWeaponSelection
    If InitVariables() Then
        tmrMain.Enabled = True
    Else
        Unload Me
    End If
End If


End Sub

Private Sub Form_Resize()

If modStickGame.bStickEditing Then
    Me.WindowState = vbNormal
    Me.height = Edit_Height
    Me.width = Edit_Width
Else
    picMain.width = Me.width
    picMain.height = Me.height
    
    StickCentreX = Me.width \ 2 - 500
    StickCentreY = Me.height \ 2 - 500
    
    'sort out constants
    RadarLeft = Me.width - RadarWidth - 100
    PlayingX = StickCentreX - 600
    ConnectingkX = StickCentreX - 900
    ConnectingkY = StickCentreY + 650
    
'    If Me.WindowState = vbMaximized Then
'        Me.BorderStyle = vbBSNone
'    Else
'        Me.BorderStyle = vbSizable
'    End If
    
End If

End Sub

Private Sub StartEdit()
'Dim hSysMenu As Long 'lStyle As Long

Me.height = Edit_Height '5100
Me.width = Edit_Width '15135

tmrMain.Enabled = False
Me.BorderStyle = vbFixedSingle
Me.Caption = Me.Caption

map_Changed = False

'#################################################
'lStyle = GetWindowLong(Me.hWnd, GWL_STYLE)
'lStyle = lStyle And Not WS_THICKFRAME
'lStyle = lStyle And Not WS_MAXIMIZEBOX
'SetWindowLong Me.hWnd, GWL_STYLE, lStyle

'hSysMenu = GetSystemMenu(Me.hWnd, 0)
'RemoveMenu hSysMenu, 4, MF_BYPOSITION

Me.Show
'#################################################

'-----------------------------------------
'object init

oPlatform(0).BackColor = &H808080

oPlatform(0).Visible = True
oBox(0).Visible = True
otBox(0).Visible = True


otBox(0).width = 375 * Stick_Edit_Zoom
otBox(0).height = 495 * Stick_Edit_Zoom
otBox(0).Top = oPlatform(0).Top - otBox(0).height

oBox(0).width = 375 * Stick_Edit_Zoom
oBox(0).Top = otBox(0).Top - oBox(0).height
oBox(0).Left = otBox(0).Left

shHealthPack.Visible = True
'HealthPackX = 49200: HealthPackY = 4800
HealthPackX = 49200: HealthPackY = 12000
Show_shpHealthPack_Pos
'-----------------------------------------


Load frmStickEdit
frmStickEdit.Show vbModeless, Me

DragInit

End Sub

Public Sub ResetEditPlatforms()
Dim i As Integer

For i = 1 To oPlatform.UBound
    Unload oPlatform(i)
Next i

For i = 1 To oBox.UBound
    Unload oBox(i)
Next i

For i = 1 To otBox.UBound
    Unload otBox(i)
Next i

With frmStickEdit
    .cmdRemoveBox.Enabled = False
    .cmdRemovePlatform.Enabled = False
    .cmdRemovetBox.Enabled = False
End With

End Sub

Public Sub Show_shpHealthPack_Pos()
shHealthPack.Left = HealthPackX * Stick_Edit_Zoom - shHealthPack.width / 2
shHealthPack.Top = HealthPackY * Stick_Edit_Zoom - shHealthPack.height / 2
End Sub

Private Sub MainLoop()
Dim Timer As Long
Dim LastFullSecond As Long
Dim nFrames As Integer
Dim newTick As Long
Dim khWnd As Long: khWnd = Me.hWnd

bRunning = True
bPlaying = True
Timer = GetTickCount() 'prevent elapsed time from being huge

#If bTimeAdjust = False Then
    StickTimeFactor = 1
#End If

Do While bRunning
    
    newTick = GetTickCount()
    If Timer + modStickGame.Stick_Ms_Required_Delay < newTick Then
    
        
        #If bTimeAdjust Then
            modStickGame.StickElapsedTime = newTick - Timer
            StickTimeFactor = StickElapsedTime / Stick_Ms_Delay '* modStickGame.sv_StickGameSpeed
        #End If
        
        
        nFrames = nFrames + 1
        If LastFullSecond + 1000 < newTick Then
            FPS = nFrames
            nFrames = 1
            LastFullSecond = newTick
        End If
        
        Timer = newTick 'GetTickCount()
        
        
        On Error GoTo EH
        
        If GetPacket() = False Then Exit Do
        
        
        SendUpdatePacket
        SendSlowPacket
        
        
        If Stick(0).bFlashed Then
            If ((GetTickCount() - Stick(0).LastFlashBang) / FlashBang_Time * PM_Rnd()) > 0.25 Then
                picMain.Cls
            End If
        Else
            picMain.Cls
        End If
        
        
        
        If bPlaying Then
            ProcessCamera 'must be at the very beginning, so we don't get any glitch-movey things
            
            'DRAWING
            DrawBulletTrails
            DrawBullets
            DrawBlood
            DisplaySticks
            DrawDeadSticks
            '---
            DrawLaserSight
            '---
            DrawMagazines
            DrawMuzzleFlashes
            '---
            DrawNames
            DrawDeadChoppers
            '---
            DrawPlatforms
            DrawBoxes
            DrawMines
            DrawtBoxes
            DrawCasings '+ remove old ones
            DrawExplosiveBarrels
            DrawWallMarks '+remove old ones
            DrawStaticWeapons 'only the images
            DrawGrass
            DrawHeads
            '---
            DrawSmokeBlasts '+remove old
            DrawCircleBlasts
            DrawGravitySmokes
            '---
            DrawTimeZones
            DrawGravityZones
            DrawAmmoPickups
            '---
            DrawTimeZoneCircs
            DrawNades
            DrawSmoke
            DrawSparks
            ProcessStaticWeaponPickup 'draw "Pick up AK-47" bit ONLY
            DisplayHealthPack
            DrawFlames '+sticks that are on fire
            DrawShieldWaves
            DrawRadar
            DrawMainMessages
            ShowChatEntry
            DrawDamageTick
            '---------
            DrawScreenCircs
            DisplayChat
            DisplayHUD
            DrawAttentions
            DrawCrosshair
            
            
            
            
            'PROCESSING
            Physics
            ProcessBlood
            ProcessDeadSticks
            ProcessMagazines
            ProcessDeadChoppers
            ProcessAllCircs
            ProcessSmoke
            ProcessSparks
            ProcessStaticWeapons
            ProcessFlames
            '---------
            ProcessKeys
            ProcessAllAI
            ProcessNades
            ProcessMines
            ProcessCasings
            ProcessMainMessages
            ProcessToasty 'technically draw, but...
            ProcessAmmoPickups 'draw "Crouch to..."
            ProcessExplosiveBarrels
            ProcessTimeZones
            ProcessGravityZones
            ProcessCircleBlasts
            ProcessRespawn '+draw "Respawn in..."
            ProcessBulletTrails
            ProcessNadeTrails
            ValidateWeapons
            ProcessAttentions
            ProcessHeads
            ProcessShields
        Else
            StaticPhysics
            'DrawBullets
            
            ProcessCamera 'must be at the very beginning, so we don't get any glitch-movey things
            
            DrawBulletTrails
            ProcessBlood
            DrawBlood
            DisplaySticks
            DrawNames
            
            DrawDeadSticks
            ProcessMagazines '+draw
            ProcessStaticWeapons '+draw
            DrawDeadChoppers
            
            
            DrawPlatforms
            DrawBoxes
            DrawtBoxes
            
            
            
            
            DrawExplosiveBarrels
            DrawWallMarks
            DrawStaticWeapons 'only the images
            DrawGrass
            DrawHeads
            DrawSmokeBlasts
            DrawGravitySmokes
            
            DrawTimeZones
            DrawGravityZones
            
            DrawCasings
            ProcessAllCircs '+ draw "Pick up AK-47" bit ONLY
            ProcessNades
            ProcessCasings
            DrawNades
            DrawSmoke
            DrawTimeZoneCircs
            ProcessSmoke '+draw
            ProcessSparks
            DrawShieldWaves
            'DrawNadeTrails
            
            'SetMyStickFacing
            'DrawCrosshair
            
            DisplayChat
            DrawAttentions
            DrawScreenCircs
            
            'show scoreboard, etc
            ProcessEndRound
            
            
            ShowChatEntry
            
            ProcessToasty
            ProcessCircleBlasts
            'ProcessNadeTrails
            ProcessKeys
            ProcessBulletTrails
            ProcessNadeTrails
            ProcessDeadSticks
            ProcessDeadChoppers
            SetMyStickFacing
            ProcessAttentions
            ProcessHeads
            ProcessShieldWaves
        End If
        
        
        'draw on form
        BltToForm
        
        
        CheckStickNames
        SendServerVarPacket
        SendRoundInfo
        
        
        If modStickGame.StickServer Then
            If bPlaying Then
                SendBoxInfo
                
                SendMineRefresh
                SendBarrelRefresh
                SendTimeZoneRefresh
                SendGravityZoneRefresh
                SendGrassRefresh
                'SendFireRefresh
                
                GenerateHealthPack
                SendStaticWeaponsPacket
                CheckMaxScore
                
                If modStickGame.sv_GameType = gElimination Or modStickGame.sv_GameType = gCoOp Then
                    ProcessElimination
                End If
                
                
            End If
        End If
        
        
        
        bHasFocus = modVars.IshWndForegroundWindow(khWnd)
    End If
    
    
EH:
    DoEvents
Loop


'On Error GoTo 0
'SavePicture picMain.Image, AppPath() & "\test.bmp"
'Stop

End Sub

Private Sub DisplayScoreBoard()
Const ScoreBoardWidth = 2150, Y = 1050
Dim sTxt As String
Dim i As Integer
Dim X As Single, tempY As Single

X = Me.width - ScoreBoardWidth

'If RadarStartTime + Radar_Time > GetTickCount() Then
    'Y = 1500
'Else
    'Y = 10
'End If

'On Error Resume Next
BorderedBox X, Y, X + ScoreBoardWidth, Y + 195 * CSng(NumSticks), BoxCol


'tempY = Y + 195 * GetHighestScorer_i()
'X = Me.width - ScoreBoardWidth
'picMain.DrawMode = Winner_DrawMode
'picmain.DrawStyle
'picMain.Line (X + 15, tempY)-(X + ScoreBoardWidth - 15, tempY + 200), vbWhite, BF
'picMain.DrawMode = vbCopyPen 'normal


X = X + 100
For i = 0 To NumSticksM1
    sTxt = Trim$(Stick(i).Name) & modMessaging.MsgNameSeparator & CStr(Stick(i).iKills - Stick(i).iDeaths)
    
    PrintStickFormText sTxt, X, i * TextHeight(sTxt) + Y, Stick(i).colour
Next i


End Sub

Private Sub BangFlash(iNade As Integer)
Dim i As Integer
Const FlashDist = 10000
Const CircDist = 10000

Stick(0).LastFlashBang = GetTickCount()
Stick(0).bFlashed = True

StunnedMouseX = MouseX + 1000 * PM_Rnd()
StunnedMouseY = MouseY + 1000 * PM_Rnd()

modStickGame.sBoxFilled -FlashDist, -FlashDist, StickGameWidth + FlashDist, StickGameHeight + FlashDist, vbWhite

For i = 0 To 100
    
    AddCirc Stick(0).X + PM_Rnd() * CircDist, Stick(0).Y + PM_Rnd() * CircDist, 750, 1, RandomRGBColour(), 30, False
    
Next i

AddCirc Nade(iNade).X, Nade(iNade).Y, 2000, 1, vbYellow, 100, True  'main explosion of nade

End Sub

Private Sub StaticPhysics()
Dim i As Integer

For i = 0 To NumSticksM1
    If Stick(i).Speed > 0 Then
        Stick(i).Speed = 0
        Stick(i).state = STICK_NONE
    'Else
        'Stick(i).Speed = Stick(i).Speed * modStickGame.StickTimeFactor / 2
    End If
    
'    If StickInGame(i) Then
'        ApplyGravity i
'
'        Motion Stick(i).X, Stick(i).Y, Stick(i).Speed, Stick(i).Heading
'    End If
    
Next i


End Sub

Private Sub ProcessToasty()
If picToasty.Visible Then
    
    picToasty.Left = picToasty.Left + modStickGame.StickTimeFactor * 30
    
    If picToasty.Left > Me.ScaleWidth Then
        picToasty.Visible = False
    End If
    
End If
End Sub

Private Sub ShowToasty()
picToasty.Left = Me.ScaleWidth / 2
picToasty.Top = Me.height - picToasty.height - 10

picToasty.Visible = True
modAudio.PlayToasty
End Sub

Private Sub BorderedBox(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, lColour As Long)
picMain.DrawWidth = 1
picMain.Line (X1, Y1)-(X2, Y2), lColour, BF
picMain.Line (X1, Y1)-(X2, Y2), vbBlack, B
End Sub

Private Sub ProcessElimination()
Const ScoreCheckDelayXK = ScoreCheckDelay * 3
Dim NumAlive As Integer, i As Integer
Static LastCheck As Long

Dim RedPresent As Boolean, BluePresent As Boolean, NeutralPresent As Boolean
Dim bRoundEnded As Boolean
Dim BiggestScore As Integer
Dim BestID As Integer

If LastCheck + ScoreCheckDelayXK < GetTickCount() Then
    
    If NumSticks > 1 Then
        
        For i = 0 To NumSticksM1
            If StickInGame(i) Then
                NumAlive = NumAlive + 1
                
                If Stick(i).Team = Blue Then
                    BluePresent = True
                ElseIf Stick(i).Team = Red Then
                    RedPresent = True
                Else
                    NeutralPresent = True
                End If
                
            End If
        Next i
        
        
        If NumAlive <= 1 Then
            
            RoundWinnerID = -1
            For i = 0 To NumSticksM1
                If StickInGame(i) Then
                    RoundWinnerID = Stick(i).ID
                    bRoundEnded = True
                    StopPlay True
                    Exit For
                End If
            Next i
            
            If RoundWinnerID = -1 Then
                'happens
                RoundWinnerID = -1 'Stick(0).ID
                bRoundEnded = True
                StopPlay True
            End If
            
            
        ElseIf Not bRoundEnded Then
            
            
            'no single Stick is alive, but a team could be alive
            If RedPresent And BluePresent = False And NeutralPresent = False Then
                'find best red Stick + end
                
                BiggestScore = -1
                
                For i = 0 To NumSticksM1
                    If StickInGame(i) Then
                        If Stick(i).Team = Red Then
                            If Stick(i).iKills > BiggestScore Then
                                BiggestScore = Stick(i).iKills
                                BestID = Stick(i).ID
                            End If
                        End If
                    End If
                Next i
                
                RoundWinnerID = BestID
                bRoundEnded = True
                StopPlay True
                
            ElseIf BluePresent And RedPresent = False And NeutralPresent = False Then
                'find best blue Stick + end
                
                BiggestScore = -1
                
                For i = 0 To NumSticksM1
                    If StickInGame(i) Then
                        If Stick(i).Team = Blue Then
                            If Stick(i).iKills > BiggestScore Then
                                BiggestScore = Stick(i).iKills
                                BestID = Stick(i).ID
                            End If
                        End If
                    End If
                Next i
                
                RoundWinnerID = BestID
                bRoundEnded = True
                StopPlay True
                
            End If
            
            
            
        End If
        
    ElseIf Stick(0).bAlive = False Then
        Stick(0).bAlive = True
        
    End If
    
    
    LastCheck = GetTickCount()
End If


End Sub

Private Function GetGameType() As String
GetGameType = kGameType(modStickGame.sv_GameType)
End Function

Private Sub MakeGameTypeArray()
Dim i As Integer

For i = 0 To eStickGameTypes.gCoOp
    Select Case i
        Case eStickGameTypes.gDeathMatch
            kGameType(i) = "DeathMatch"
        Case eStickGameTypes.gElimination
            kGameType(i) = "Elimination"
        Case eStickGameTypes.gCoOp
            kGameType(i) = "Co-Op"
    End Select
Next i

End Sub

'Private Sub ProcessPerk()
'Const ESP_Print_Len = 4500
'Const ESP_Print_LenDX = ESP_Print_Len * 1.1
'Const ESP_Y_Offset = 500
'Const ESP_Col = vbBlue

'Dim i As Integer
'Dim tDist As Single, tAng As Single


'If Stick(0).Perk = pSniper And StickInGame(0) Then
'    'draw esp map
'
'    picMain.DrawWidth = 1
'
'    PrintStickText "Stealth Awareness", Stick(0).X - 600, Stick(0).Y + ESP_Y_Offset - ESP_Print_Len - 250, ESP_Col
'
'    For i = 1 To NumSticksM1
'        If StickInGame(i) Then
'            If CanSeeStick(i) Then
'                tDist = GetDist(Stick(0).X, Stick(0).Y, Stick(i).X, Stick(i).Y)
'
'                If tDist < StealthESPDist Then
'                    tAng = FindAngle(Stick(0).X, Stick(0).Y, Stick(i).X, Stick(i).Y - 1)
'
'                    'picMain.Font.Size = 9 - tDist / 10000
'
'                    PrintStickText Trim$(Stick(i).Name), _
'                        Stick(0).X + ESP_Print_Len * Sine(tAng), _
'                        Stick(0).Y + ESP_Y_Offset - ESP_Print_Len * CoSine(tAng), _
'                        Stick(i).Colour
'                End If
'            End If
'        End If
'    Next i
'
'    modStickGame.sCircle Stick(0).X, Stick(0).Y + ESP_Y_Offset, ESP_Print_LenDX, ESP_Col
'    picMain.Font.Size = 8
'End If

'End Sub

Public Sub StickGameSpeedChanged(oldSpeed As Single, newSpeed As Single)
'Dim i As Integer
'Dim K As Long
'Const def_Freq = 22050

Erase Flame: NumFlames = 0
Erase Smoke: NumSmoke = 0

If StickInGame(0) = False Then
    SetSoundFreq newSpeed
    Stick(0).sgTimeZone = newSpeed
ElseIf Stick(0).sgTimeZone = newSpeed Then
    Stick(0).sgTimeZone = newSpeed + 0.1 'ensure sound_freq gets changed
End If


'If oldSpeed > -1 Then
'    K = oldSpeed / newSpeed
'    For i = 0 To NumNades - 1
'        Nade(i).Decay = (Nade(i).Decay - GetTickCount()) * K + GetTickCount()
'        '(Nade(i).Decay - GetTickCount()) * oldSpeed / newSpeed + GetTickCount()
'        '(n-g)o/n + g
'        'o-g(o/n + 1)   ?
'    Next i
'End If

End Sub

Private Sub CheckChopperCollisions()
Dim i As Integer, j As Integer, K As Integer

For i = 0 To NumSticksM1
    If StickInGame(i) Then
        If Stick(i).WeaponType = Chopper Then
            For j = 0 To NumSticksM1
                If StickInGame(j) Then
                    If j <> i Then
                        If Not IsAlly(Stick(j).Team, Stick(i).Team) Then
                            'If Stick(j).LastSpawnTime + Spawn_Invul_Time / GetTimeZoneAdjust < GetTickCount() Then
                            If StickInvul(j) = False Then
                                If CoOrdInStick(Stick(j).X, Stick(j).Y, i) Then
                                    
                                    'If Stick(j).Perk <> pZombie Then
                                        For K = 1 To 30
                                            'splatter!
                                            AddBlood Stick(j).X, Stick(j).Y, Rnd() * Pi2
                                        Next K
                                        
                                        If j = 0 Or Stick(j).IsBot Then
                                            Call Killed(j, i, kChoppered)
                                        End If
                                    'End If
                                    
                                End If 'co-ord endif
                            End If 'invul endif
                        End If 'team endif
                    End If 'j<>i endif
                End If 'stickingame endif
            Next j
        End If 'ischopper endif
    End If 'stickingame endif
Next i


End Sub

Private Sub DrawRadar()
Dim i As Integer
Dim C As Long 'yes needed
Dim pX As Single, pY As Single
Const RadarTop = 10
Const RadarHeight = 1000
Dim bCan As Boolean


'If RadarStartTime + Radar_Time > GetTickCount() Then
    
    picMain.DrawWidth = 2
    picMain.ForeColor = vbBlack
    picMain.Line (RadarLeft, RadarTop)-(RadarLeft + RadarWidth, RadarTop + RadarHeight), vbBlue, BF
    
    'modStickGame.PrintStickFormText "Time Left: " & _
        CStr(Round(((Radar_Time + RadarStartTime - GetTickCount()) / 1000))), _
        RadarLeft + 200, RadarTop + RadarHeight, vbBlack
    
    
    picMain.FillStyle = vbFSSolid
    picMain.FillColor = vbBlack
    For i = 0 To UBound(AmmoPickup)
        If AmmoPickup(i).bActive Then
            
            picMain.Circle (RadarLeft + RadarWidth * AmmoPickup(i).X / StickGameWidth, _
                RadarTop + RadarHeight * AmmoPickup(i).Y / StickGameHeight), 10, vbBlack
            
        End If
    Next i
    
    
    picMain.DrawWidth = 1
    For i = 0 To modStickGame.ubdPlatforms
        
        picMain.Line (RadarLeft + RadarWidth * Platform(i).Left / StickGameWidth, _
                      RadarTop + RadarHeight * Platform(i).Top / StickGameHeight)- _
                     (RadarLeft + RadarWidth * (Platform(i).Left + Platform(i).width) / StickGameWidth, _
                     RadarTop + RadarHeight * (Platform(i).Top + Platform(i).height) / StickGameHeight), BoxCol, BF
        
        
    Next i
    
    
    
    picMain.FillStyle = vbFSTransparent
    For i = 0 To NumSticksM1
        
        If StickInGame(i) Then
            
            If i = 0 Then
                bCan = True
            ElseIf Stick(i).WeaponType = Chopper Then
                bCan = True
            ElseIf IsAlly(Stick(0).Team, Stick(i).Team) Then
                bCan = True
            Else
                bCan = (Stick(i).LastLoudBullet + Radar_Bullet_ShowTime > GetTickCount())
            End If
            
            
            
            If bCan Then
                'If Stick(i).Perk <> pBombSquad Then
                    
                    pX = RadarLeft + RadarWidth * Stick(i).X / StickGameWidth
                    pY = RadarTop + RadarHeight * Stick(i).Y / StickGameHeight
                    
                    C = Stick(i).colour 'GetTeamColour(Stick(i).Team)
                    picMain.FillColor = C
                    picMain.Circle (pX, pY), 60, C
                    
                    If i = 0 Then
                        'draw an X on me
                        DrawX pX, pY
                    End If
                'End If
            End If
        End If
    Next i
    
    
    
'ElseIf bHadRadar Then
    'AddMainMessage "Radar Expired"
    'bHadRadar = False
'End If

End Sub

Private Sub DrawX(ByVal pX As Single, ByVal pY As Single)

Const CrossWidth = 75

picMain.Line (pX - CrossWidth, pY + CrossWidth)-(pX + CrossWidth, pY - CrossWidth), Stick(0).colour
picMain.Line (pX + CrossWidth, pY + CrossWidth)-(pX - CrossWidth, pY - CrossWidth), Stick(0).colour

End Sub

Private Sub DrawMainMessages()
'Const WO2 = 3935 'Width \ 2 - 100
'Const HO2 = 3235 'Height \ 2 - 100
Dim i As Integer
Dim Tmp As String
Const TextH = 480

If NumMainMessages Then
    'If Not F1Pressed Then  'And ShowMainMsg Then
    
    picMain.Font.Size = 18
    picMain.ForeColor = Stick(0).colour 'MGrey
    
    
    For i = 0 To NumMainMessages - 1
        Tmp = MainMessages(i).Text
        
        PrintStickFormText Tmp, _
            StickCentreX - picMain.TextWidth(Tmp) / 2, _
            i * TextH + StickCentreY + 1000, _
            MainMessages(i).colour
        
        'can't const StickCentreY + 1000, since StickCentreY isn't const
        
    Next i
    
    picMain.Font.Size = 8
End If

End Sub

Private Sub ProcessMainMessages()
Dim i As Integer
Dim GTC As Long

GTC = GetTickCount()
    
Do While i < NumMainMessages
    If MainMessages(i).Decay < GTC Then
        RemoveMainMessage i
        i = i - 1
    End If
    
    i = i + 1
Loop

End Sub

Public Sub AddMainMessage(ChatText As String, bAllowSkipping As Boolean, Optional ByVal lColour As Long = -1)

If lColour = -1 Then lColour = Stick(0).colour


If NumMainMessages > 0 Then
    If MainMessages(NumMainMessages - 1).Text = ChatText Then
        If bAllowSkipping Then
            With MainMessages(NumMainMessages - 1)
                .Decay = GetTickCount() + MainMessageDecay
                
                .colour = IIf(.colour = lColour, vbBlack, lColour)
            End With
            
            Exit Sub
        'Else
            'was a "Timmy WAS killed etc..." message, let it stay
        End If
    End If
End If

'Add this value to the chat text array
ReDim Preserve MainMessages(NumMainMessages)

MainMessages(NumMainMessages).Decay = GetTickCount() + MainMessageDecay
MainMessages(NumMainMessages).Text = ChatText
MainMessages(NumMainMessages).colour = lColour

NumMainMessages = NumMainMessages + 1

End Sub

Private Sub RemoveMainMessage(Index As Integer)

Dim i As Long

'Remove the specified chat text
For i = Index To NumMainMessages - 2
    MainMessages(i) = MainMessages(i + 1)
Next i

'Resize the array
If NumMainMessages = 1 Then
    Erase MainMessages
    NumMainMessages = 0
Else
    ReDim Preserve MainMessages(NumMainMessages - 2)
    NumMainMessages = NumMainMessages - 1
End If
    
End Sub

Private Sub SetStickOnFire(i As Integer)

If Stick(i).bOnFire Then
    If StickInvul(i) Then
        Stick(i).bOnFire = False
    Else
        If Stick(i).LastFlameTouch + Flame_Burn_Time / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            Stick(i).bOnFire = False
            ResetStickFire 0
        End If
    End If
End If

End Sub
Private Sub SetStickFlashed(i As Integer)

If Stick(i).bFlashed Then
    If StickInvul(i) Then
        Stick(i).bFlashed = False
    Else
        If Stick(i).LastFlashBang + FlashBang_Time / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            Stick(i).bFlashed = False
            ResetStickFlash 0
        End If
    End If
End If

End Sub

Private Sub ProcessKeys()
Dim i As Integer

'##########################################
SetStickFlashed 0
SetStickOnFire 0
'##########################################


If StickInGame(0) And bPlaying Then
    
    
    'THIS IS THE BIT THAT'LL RESET W1200/SNIPER FACING AFTER A BULLET
    DoMyStickFacing
    
    
    If FireKey Then
        AddStickiState 0, STICK_FIRE
    End If
    
    If CrouchKey Then
        If Stick(0).bOnSurface Then
            AddStickiState 0, STICK_CROUCH
        End If
    End If
    
Else
    
    Const SpecCamInc = 250
    
    For i = NumDeadSticks - 1 To 0 Step -1 'go backwards, so we get the most recent dead stick of me
        If DeadStick(i).bIsMe Then
            If DeadStick(i).bOnSurface = False Then
                
                CentreCameraOnPoint DeadStick(i).X, DeadStick(i).Y - 250
                
                'i = -1
                'Exit For
                Exit Sub
            End If
        End If
    Next i
    
    
    'If i > -1 Then
        If SpecUp Then
            MoveCameraY modStickGame.cg_sCamera.Y - SpecCamInc * modStickGame.cl_SpecSpeed
        ElseIf SpecDown Then
            MoveCameraY modStickGame.cg_sCamera.Y + SpecCamInc * modStickGame.cl_SpecSpeed
        End If
        
        
        If SpecLeft Then
            MoveCameraX modStickGame.cg_sCamera.X - SpecCamInc * modStickGame.cl_SpecSpeed
        ElseIf SpecRight Then
            MoveCameraX modStickGame.cg_sCamera.X + SpecCamInc * modStickGame.cl_SpecSpeed
        End If
    'End If
    
    
End If

End Sub

Public Function WeaponSilencable(vWeapon As eWeaponTypes) As Boolean
'Select Case vWeapon
    'Case AK, M82, XM8, AUG, USP, MP5, AWM, Mac10, SPAS, G3
        'WeaponSilencable = True
'End Select

WeaponSilencable = kSilencable(vWeapon)

End Function
Public Function WeaponIsSniper(vWeapon As eWeaponTypes) As Boolean
'Select Case vWeapon
'    Case M82, AWM
'        WeaponIsSniper = True
'End Select

WeaponIsSniper = ((vWeapon = M82) Or (vWeapon = AWM))

End Function
Public Function WeaponIsPistol(vWeapon As eWeaponTypes) As Boolean
'Select Case vWeapon
'    Case M82, AWM
'        WeaponIsSniper = True
'End Select

WeaponIsPistol = ((vWeapon = DEagle) Or (vWeapon = USP))

End Function

Private Sub ProcessAllAI()
Dim i As Integer


'If modStickGame.sv_BotAI Then
    For i = 1 To NumSticksM1
        If Stick(i).IsBot Then
            If StickInGame(i) Then
                ProcessAI i
            End If
        End If
    Next i
'Else
    'For i = 0 To NumSticksM1
        'If Stick(i).IsBot Then
            'If Stick(i).State <> Stick_None Then
                'Stick(i).State = Stick_None
            'End If
        'End If
    'Next i
'End If

End Sub

Private Sub ProcessAI(i As Integer)
Dim Dist As Single, AngleToTarget As Single
Dim iTarget As Integer, j As Integer

Const AI_Facing_Adjust_Delay = AI_Delay / 6
Const BotHardcoreScanInc = Pi / 6

If i > -1 Then
    
    SetStickFlashed i
    SetStickOnFire i
    
    If Stick(i).bFlashed Then
        
        If Rnd() > 0.9 Then
            If StickiHasState(i, STICK_LEFT) Then
                SubStickiState i, STICK_LEFT
                AddStickiState i, STICK_RIGHT
            Else
                SubStickiState i, STICK_RIGHT
                AddStickiState i, STICK_LEFT
            End If
            
            
            If Rnd() > 0.9 Then
                Stick(i).ActualFacing = Stick(i).ActualFacing + PM_Rnd()
                If StickiHasState(i, STICK_FIRE) Then
                    SubStickiState i, STICK_FIRE
                Else
                    AddStickiState i, STICK_FIRE
                End If
            End If
            
        End If
    ElseIf Stick(i).AICurrentTarget > -1 And Stick(i).AICurrentTarget <= NumSticksM1 Then
        If StickInGame(Stick(i).AICurrentTarget) Then
            If Stick(i).AILastFacingAdjust + AI_Facing_Adjust_Delay < GetTickCount() Then
                
                DoAIFacing i, Stick(i).AICurrentTarget, AngleToTarget
                Stick(i).AI_AngleToTarget = AngleToTarget
                
                Stick(i).AILastFacingAdjust = GetTickCount()
                
            End If
        Else
            Stick(i).AICurrentTarget = -1
        End If
    ElseIf Stick(i).AICurrentTarget > NumSticksM1 Then
        Stick(i).AICurrentTarget = -1
    End If
    
    
    
    If Not Stick(i).bFlashed Then
        If Stick(i).LastAI + AI_Delay < GetTickCount() Then
            
            iTarget = ClosestTargetI(i, Dist)
            If Stick(i).AICurrentTarget <> iTarget Then
                Stick(i).AICurrentTarget = iTarget
                SubStickiState i, STICK_FIRE
                'stop firing between locating targets
            End If
            
            
            If Stick(i).AICurrentTarget > -1 Then
                If Stick(i).WeaponType = Chopper Then
                    ProcessChopperAI i, Stick(i).AICurrentTarget, Dist
                Else
                    ProcessStickAI i, Stick(i).AICurrentTarget, Dist, Stick(i).AI_AngleToTarget
                End If
                
            Else
                
                If modStickGame.sv_Hardcore Then
                    'attempt to locate someone
                    Stick(i).ActualFacing = Stick(i).ActualFacing + BotHardcoreScanInc
                    Stick(i).Facing = Stick(i).ActualFacing
                    
                    For j = 0 To NumSticksM1
                        If (GetTickCount() - Stick(j).LastBullet) < 100 Then
                            Stick(i).AICurrentTarget = j
                            Exit For
                        End If
                    Next j
                    
                End If
                
                
                Stick_Still i
            End If
            
            
            Stick(i).LastAI = GetTickCount()
            
            
        End If
    End If
    
End If

End Sub

Private Sub Stick_Still(i As Integer)
Dim bReloading As Boolean
Dim State_To_Be As eStickStates

bReloading = ((Stick(i).state And STICK_RELOAD) = STICK_RELOAD)

If bReloading Then
    State_To_Be = STICK_RELOAD
Else
    State_To_Be = STICK_NONE
End If

If Stick(i).WeaponType <> Chopper Then
    State_To_Be = State_To_Be Or STICK_CROUCH
End If


If Stick(i).state <> State_To_Be Then
    SetStickiState i, State_To_Be
    
    CheckAIReload i
End If

'If StickiHasState(i, Stick_Left) Then
'    SubStickiState i, Stick_Left
'End If
'If StickiHasState(i, Stick_Right) Then
'    SubStickiState i, Stick_Right
'End If
'If StickiHasState(i, Stick_Jump) Then
'    SubStickiState i, Stick_Jump
'End If
'If StickiHasState(i, Stick_Crouch) Then
'    If Stick(i).WeaponType = Chopper Then
'        SubStickiState i, Stick_Crouch
'    End If
'ElseIf Stick(i).WeaponType <> Chopper Then
'    AddStickiState i, Stick_Crouch
'End If


End Sub

Private Sub CheckAIReload(i As Integer)
If Stick(i).BulletsFired > 5 Then
    If Stick(i).WeaponType <> M249 Then
        If StickiHasState(i, STICK_RELOAD) = False Then
            StartReload i
        End If
    End If
End If
End Sub


Private Sub DoAIFacing(iAI As Integer, iTarget As Integer, AngleToTarget As Single)
Dim FixedAngle As Single

If Stick(iTarget).WeaponType = Chopper Then
    AngleToTarget = FindAngle_Actual(Stick(iAI).GunPoint.X, Stick(iAI).GunPoint.Y, Stick(iTarget).X, GetStickY(iTarget) - CLD6)
Else
    AngleToTarget = FindAngle_Actual(Stick(iAI).X, Stick(iAI).Y, Stick(iTarget).X, GetStickY(iTarget)) '+ HeadRadius)
End If


'Adjust facing
FixedAngle = FixAngle(AngleToTarget - Stick(iAI).ActualFacing)
If FixedAngle <= modStickGame.sv_AI_Rotation_Rate Or FixedAngle >= modStickGame.sv_AI_pi2LessRotRate Then
    
    Stick(iAI).ActualFacing = AngleToTarget
    
    
ElseIf FixedAngle >= Pi Then
    'mudtFleet2(i).sngFacing = mudtFleet2(i).sngFacing - FLEET2_ROTATION_RATE * pi / 180
    Stick(iAI).ActualFacing = Stick(iAI).ActualFacing - _
        modStickGame.sv_AI_Rotation_Rate * Stick(iAI).sgTimeZone
    
ElseIf FixedAngle < Pi Then
    'mudtFleet2(i).sngFacing = mudtFleet2(i).sngFacing + FLEET2_ROTATION_RATE * pi / 180
    Stick(iAI).ActualFacing = Stick(iAI).ActualFacing + _
        modStickGame.sv_AI_Rotation_Rate * Stick(iAI).sgTimeZone
    
End If


If StickiHasState(iAI, STICK_FIRE) = False Then Stick(iAI).Facing = Stick(iAI).ActualFacing


End Sub

Private Sub ProcessStickAI(iAI As Integer, iTarget As Integer, DistToTarget As Single, AngleToTarget As Single)
Const BulletRange = StickGameWidth / 6, LevelGap = 1800, BulletRangeX2 = BulletRange * 2
Const AI_Mine_Attempt_Delay = Mine_Delay * 3

'decision vars
Dim yDist As Single, xDist As Single, bDontMove As Boolean
'act-on vars
Dim bJump As Boolean, IDir As Integer, bCanShoot As Boolean, bCloseToRange As Boolean, i_xDirection As Integer
'                                                                                     1 = right, -1 = left, 0 = stay


yDist = Stick(iAI).Y - Stick(iTarget).Y
xDist = Stick(iAI).X - Stick(iTarget).X

If Stick(iAI).Perk = pZombie Then
    bCloseToRange = True
    bCanShoot = True
Else
    If DistToTarget < BulletRange Then
        bCanShoot = True
    ElseIf WeaponIsSniper(Stick(iAI).WeaponType) Or Stick(iTarget).WeaponType = Chopper Then
        If DistToTarget < BulletRangeX2 Then
            bCanShoot = True
        Else
            bCloseToRange = True
            CheckAIReload iAI
        End If
    Else
        'close to BulletRange
        bCloseToRange = True
        
        CheckAIReload iAI
    End If
End If


If yDist > LevelGap Then
    'target above
    
    If Abs(xDist) < IIf(Stick(iTarget).WeaponType = Chopper, 7000, 5500) Then
        bJump = True
        bCanShoot = False
    Else
        bCloseToRange = True
        bJump = False
    End If
    
ElseIf yDist < -LevelGap Then
    'target 1+ level(s) down
    
    ''always go right
    'not anymore
    i_xDirection = IIf(xDist > 0, 1, -1) 'changable
    
End If



If bCloseToRange Then
    i_xDirection = IIf(xDist > 0, -1, 1) 'do not change
    
    SubStickiState iAI, STICK_CROUCH
End If

If bCanShoot Then
    
    'shoot+nade
    If Stick(iAI).WeaponType = FlameThrower Or Stick(iAI).WeaponType = RPG Then
        Stick(iAI).Facing = Stick(iAI).ActualFacing
    End If
    
    
    
    If modStickGame.sv_AIShoot Then
        
        If Stick(iAI).Perk = pSniper Then 'WeaponIsSniper(Stick(iAI).WeaponType) Then
            
            If Stick(iAI).bOnSurface Then
                If AngleToTarget <= ProneRightLimit Or AngleToTarget >= ProneLeftLimit Then
                    'If StickiHasState(iAI, Stick_Prone) = False Then
                    AddStickiState iAI, STICK_PRONE
                    'End If
                    SubStickiState iAI, STICK_CROUCH
                Else
                    SubStickiState iAI, STICK_PRONE
                    
                    'If StickiHasState(iAI, Stick_Crouch) = False Then
                    AddStickiState iAI, STICK_CROUCH
                    'End If
                End If
                
                
                If StickiHasState(iAI, STICK_CROUCH) = False Then
                    bCanShoot = StickiHasState(iAI, STICK_PRONE)
                'Else
                    'bcanshoot=false
                End If
            Else
                bCanShoot = False
            End If
            
            
        ElseIf StickiHasState(iAI, STICK_CROUCH) Then
            SubStickiState iAI, STICK_CROUCH
        End If
        
        
        If bCanShoot Then
            If StickInvul(iTarget) = False Then
                If AnglesRoughlyEqual(Stick(iAI).ActualFacing, AngleToTarget) Then
                    If StickiHasState(iAI, STICK_FIRE) = False Then
                        AddStickiState iAI, STICK_FIRE
                    End If
                ElseIf StickiHasState(iAI, STICK_FIRE) Then
                    SubStickiState iAI, STICK_FIRE
                End If
            End If
        Else
            SubStickiState iAI, STICK_FIRE
        End If
        
'        If Stick(iAi).WeaponType = M82 Or Stick(iAi).WeaponType = W1200 Then
'            'delay between shots
'
'            If Stick(iAi).LastBullet + 300 > GetTickCount() Then
'                If StickiHasState(iAi, Stick_Fire) Then
'                    SubStickiState iAi, Stick_Fire
'                End If
'            End If
'
'        End If
        
        If Stick(iAI).Perk <> pZombie Then
            If Stick(iAI).LastNade + Stick(iAI).AINadeDelay / GetSticksTimeZone(iAI) < GetTickCount() Then
                AddStickiState iAI, STICK_NADE
                
                SetAINadeDelay iAI
                Stick(iAI).AIPickedNade = False
                
            ElseIf Stick(iAI).AIPickedNade = False Then
                
                Stick(iAI).AIPickedNade = True
                
                If sv_AIUseFlashBangs Then
                    Stick(iAI).iNadeType = IIf(Rnd() > 0.75, eNadeTypes.nFlash, eNadeTypes.nFrag)
                Else
                    Stick(iAI).iNadeType = nFrag
                End If
                
            End If
            
            If modStickGame.sv_AIMine Then
                If Stick(iAI).AILastMineAttempt + AI_Mine_Attempt_Delay < GetTickCount() Then
                    If ShouldPlantMine(iAI) Then
                        AddStickiState iAI, STICK_MINE
                        
                        Stick(iAI).AILastMineAttempt = GetTickCount()
                    Else
                        
                        Stick(iAI).AILastMineAttempt = GetTickCount() - AI_Mine_Attempt_Delay * Rnd()
                    End If
                End If
            End If
        End If
        
    ElseIf StickiHasState(iAI, STICK_FIRE) Then
        SubStickiState iAI, STICK_FIRE
    End If
End If


ProcessAI_LeftRight iAI, i_xDirection


If bJump Then
    If modStickGame.sv_AIMove Then
        If Stick(iTarget).bOnSurface Then
            If Stick(iAI).Speed < 40 Or Stick(iAI).Perk = pZombie Or Stick(iAI).bOnSurface = False Then
                
                If Stick(iAI).bOnSurface Then
                    AddStickiState iAI, STICK_JUMP
                    Stick(iAI).Y = Stick(iAI).Y - 50
                Else
                    SubStickiState iAI, STICK_JUMP
                End If
                
            End If
        End If
    End If
End If


End Sub

Private Function ShouldPlantMine(iStick As Integer) As Boolean
Dim i As Integer

For i = 0 To NumSticksM1
    If GetDist(Stick(iStick).X, Stick(iStick).Y, _
               Stick(i).X, Stick(i).Y) < Mine_StickLim * 1.5 Then
        
        If i <> iStick Then
            ShouldPlantMine = False
            Exit Function
        End If
        
    End If
Next i

ShouldPlantMine = True

End Function

Private Sub ProcessAI_LeftRight(iAI As Integer, i_xDirection As Integer)

If i_xDirection And Stick(iAI).bOnSurface Then
    'move to range
    If modStickGame.sv_AIMove Then
        If i_xDirection = -1 Then
            If StickiHasState(iAI, STICK_LEFT) = False Then
                AddStickiState iAI, STICK_LEFT
            End If
            If StickiHasState(iAI, STICK_RIGHT) Then
                SubStickiState iAI, STICK_RIGHT
            End If
        Else
            If StickiHasState(iAI, STICK_RIGHT) = False Then
                AddStickiState iAI, STICK_RIGHT
            End If
            If StickiHasState(iAI, STICK_LEFT) Then
                SubStickiState iAI, STICK_LEFT
            End If
        End If
    End If
Else
    If StickiHasState(iAI, STICK_LEFT) Then
        SubStickiState iAI, STICK_LEFT
    ElseIf StickiHasState(iAI, STICK_RIGHT) Then
        SubStickiState iAI, STICK_RIGHT
    End If
End If

End Sub

Private Sub ProcessChopperAI(iAI As Integer, iTarget As Integer, DistToTarget As Single)
Const ChopperMinYDist = 2000, ChopperMinYDist2 = 800, ChopperMinXDist = 7000, StickGameWidthD2 = StickGameWidth / 2
Dim yDist As Long, xDist As Long, IDir As Integer

If modStickGame.sv_AIMove Then
    'If modStickGame.sv_GameType <> gCoOp Then
        'SPAWNED ABOVE CHOPPERLEN IN KILLED()
            'AddStickiState iAI, STICK_JUMP
'            'Stick(iAI).Speed = 0
'        ElseIf StickiHasState(iAI, STICK_JUMP) Then
'            SubStickiState iAI, STICK_JUMP
'            'ResetYComp iAI
'        End If
'
        
        'Stick(iAI).Speed = 0
        
        'SubStickiState iAI, STICK_CROUCH
        
        'yDist = Stick(iAI).Y - Stick(iTarget).Y
        
'        'up + down
'        If yDist > ChopperMinYDist2 Then
'            'they're above me
'
'            If StickiHasState(iAI, stick_Jump) = False Then
'                AddStickiState iAI, stick_Jump
'            ElseIf StickiHasState(iAI, Stick_Crouch) Then
'                SubStickiState iAI, Stick_Crouch
'            End If
'
'
'        ElseIf yDist < -ChopperMinYDist Then
'            If StickiHasState(iAI, Stick_Crouch) = False Then
'                AddStickiState iAI, Stick_Crouch
'            ElseIf StickiHasState(iAI, stick_Jump) Then
'                SubStickiState iAI, stick_Jump
'            End If
'
'
'        Else
'
'            'on their level
'            If StickiHasState(iAI, stick_Jump) Then
'                SubStickiState iAI, stick_Jump
'            ElseIf StickiHasState(iAI, Stick_Crouch) Then
'                SubStickiState iAI, Stick_Crouch
'            End If
'
'            If Stick(iAI).Speed > 0 Then
'                Stick(iAI).Speed = Stick(iAI).Speed / 3
'            End If
'
'        End If
        
    If modStickGame.sv_GameType = gCoOp Then
        If Stick(iAI).Facing < piD2 Then
            If Stick(iAI).Facing > pi3D2 Then
                'can't face up
                Stick(iAI).Facing = IIf(Stick(iAI).Facing < Pi, piD2, pi3D2)
            End If
        End If
    End If
    
    
    xDist = Stick(iAI).X - Stick(iTarget).X
    If xDist > ChopperMinXDist Then
        IDir = -1
    ElseIf xDist < -ChopperMinXDist Then
        IDir = 1
    End If
End If

If DistToTarget < StickGameWidthD2 Then
    If modStickGame.sv_AIShoot Then
        AddStickiState iAI, STICK_FIRE
        If Stick(iAI).Speed > 0 Then
            Stick(iAI).Speed = Stick(iAI).Speed / 2
        End If
    Else
        SubStickiState iAI, STICK_FIRE
    End If
Else
    SubStickiState iAI, STICK_FIRE
End If

ProcessAI_LeftRight iAI, IDir


If modStickGame.sv_AIHeliRocket Then
    If modStickGame.sv_AIShoot Then
        If StickiHasState(iAI, STICK_NADE) = False Then
            AddStickiState iAI, STICK_NADE
        End If
    End If
End If


End Sub

Private Function AnglesRoughlyEqual(A1 As Single, A2 As Single) As Boolean
Const lAccuracy As Long = 1

AnglesRoughlyEqual = (Round(FixAngle(A1), lAccuracy) = Round(FixAngle(A2), lAccuracy))

End Function

Private Sub SetAINadeDelay(i As Integer)

Stick(i).AINadeDelay = Nade_Delay * 5 * Rnd()

End Sub

Private Function ClosestTargetI(iSource As Integer, DistToTarget As Single) As Integer

Dim GTC As Long
Dim iTarget As Integer
Dim Dist As Single, TestDist As Single
Dim iCurrent As Integer 'current stick with least dist to iSource
Dim jSpy As Integer
Dim bCanTestStick As Boolean
Const AI_Bullet_Wait_Time = 1000, AI_Box_Fire_Wait_Time = 3000 'AI can see you for 1 second after you shoot


Dist = StickGameWidth + 100
iCurrent = -1
GTC = GetTickCount()


For iTarget = 0 To NumSticksM1
    If iTarget <> iSource Then
        If StickInGame(iTarget) Then
            
            
            If Stick(iTarget).Perk = pSpy Then
                
                If Stick(iTarget).WeaponType <> Chopper Then
                    jSpy = FindStick(Stick(iTarget).MaskID)
                    If jSpy = -1 Then jSpy = iTarget
                Else
                    jSpy = iTarget
                End If
                
                
                If jSpy = iSource Then
                    bCanTestStick = True
                    jSpy = iTarget
                ElseIf jSpy = iTarget Then
                    bCanTestStick = True
                Else
                    bCanTestStick = False
                End If
                
                
            ElseIf Stick(iTarget).Perk = pSniper Then
                jSpy = iTarget
                
                bCanTestStick = True
                If StickiHasState(iTarget, STICK_PRONE) Then
                    If Stick(iTarget).Speed = 0 Then
                        If Stick(iTarget).LastBullet + AI_Bullet_Wait_Time < GTC Then
                            bCanTestStick = StickSeenStick(iSource, iTarget)
                            
                        ElseIf Stick(iTarget).bSilenced Then
                            'they have shot in the part AI_Bullet_Wait_Time seconds
                            
                            If Not StickiHasState(iTarget, STICK_FIRE) Then
                                'if target isn't shooting, it depends on whether we've seen them
                                bCanTestStick = StickSeenStick(iSource, iTarget)
                            End If
                            
                        End If
                        
                    End If
                End If
                
                
            Else
                jSpy = iTarget
                bCanTestStick = True
            End If
            
            
            If bCanTestStick Then
                If Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team) Then
                    If StickCanSeeStick(iSource, iTarget) Then
                        'If (StickInSmoke(iTarget) = False) Then
                            'If Stick(iTarget).Speed < 20 Then
                                'If StickiHasState(iTarget, Stick_Fire) = False Then
                                    If StickSeenStick(iSource, iTarget) Then
                                        bCanTestStick = True
                                    ElseIf StickInTBox(iTarget) = False Or _
                                            ( _
                                            ((GetTickCount() - Stick(iTarget).LastBullet) < AI_Box_Fire_Wait_Time) And _
                                            Stick(iTarget).bSilenced = False _
                                            ) Then
                                        
                                        bCanTestStick = True
                                    Else
                                        bCanTestStick = False
                                    End If
                                    
                                    
                                    If bCanTestStick Then
                                        'they are not in a box, OR
                                        'they are in a box and have shot in the past AI_Box_Fire_Wait_Time seconds, not silently
                                        
                                        TestDist = GetDist(Stick(iSource).X, Stick(iSource).Y, Stick(iTarget).X, Stick(iTarget).Y)
                                        
                                        
                                        If TestDist < Dist Then
                                            iCurrent = iTarget
                                            
                                            'If Stick(iTarget).bSilenced Then
                                                'Dist = TestDist * 1.5 'harder to see
                                            'Else
                                                Dist = TestDist
                                            'End If
                                            
                                        End If
                                        
                                    End If 'in tBox
                                'End If 'fire state
                            'End If 'speed<20
                        'End If 'stickinsmoke
                    End If 'stickcanseestick
                End If 'isally
            End If
            
            
        End If
    End If
Next iTarget


If iCurrent > -1 Then
    If StickSeenStick(iSource, iCurrent) = False Then
        If PointVisibleOnSticksScreen(Stick(iCurrent).X, Stick(iCurrent).Y, iSource) Then
            AddStickToBotsTargets iSource, iCurrent
        End If
    End If
End If


ClosestTargetI = iCurrent
DistToTarget = Dist

End Function

'Private Function StickInSmoke(iStick As Integer) As Boolean
'Dim iTarget As Integer
'Const MinDist = 50, Inc = 0.5
'
'
'If Stick(iStick).WeaponType <> Chopper Then
'
'    For iTarget = 0 To NumLargeSmokes - 1
'
'        If GetDist(Stick(iStick).X, Stick(iStick).Y, LargeSmoke(iTarget).CentreX, LargeSmoke(iTarget).CentreY) < _
'            (MinDist + Inc * LargeSmoke(iTarget).iSize) Then
'
'            StickInSmoke = True
'            Exit For
'
'        End If
'    Next iTarget
'End If
'
'
'End Function

Private Sub AddStickToBotsTargets(iSource As Integer, iTarget As Integer)

Stick(iSource).AI_Targets_Seen = Stick(iSource).AI_Targets_Seen & "," & CStr(iTarget)

End Sub

Private Sub RemoveStickFromBotsTargets(iSource As Integer, iTarget As Integer)
Dim i As Integer
Dim Targets() As String

'rebuild it, excluding the target to remove

Targets = Split(Stick(iSource).AI_Targets_Seen, ",")
Stick(iSource).AI_Targets_Seen = vbNullString

For i = LBound(Targets) To UBound(Targets)
    If LenB(Targets(i)) Then
        If Targets(i) <> CStr(iTarget) Then
            Stick(iSource).AI_Targets_Seen = Stick(iSource).AI_Targets_Seen & "," & CStr(Targets(i))
        End If
    End If
Next i

Erase Targets

End Sub

Private Function StickSeenStick(iSee_er As Integer, iTarget As Integer) As Boolean

StickSeenStick = InStr(1, Stick(iSee_er).AI_Targets_Seen, CStr(iTarget))

End Function

Private Function StickInTBox(iStick As Integer) As Boolean
Dim i As Integer
Dim sY As Single
Dim rcStick As RECT
Dim rcTBox As RECT

sY = GetStickY(iStick)
If Stick(iStick).Perk = pSniper Then
    If StickiHasState(iStick, STICK_CROUCH) Then
        sY = sY + ArmLen
    End If
End If
rcStick = PointToRect(Stick(iStick).X, sY)


For i = 0 To modStickGame.ubdtBoxes
    With rcTBox
        .Top = tBox(i).Top
        .Bottom = .Top + tBox(i).height
        .Left = tBox(i).Left
        .Right = .Left + tBox(i).width
    End With
    
    If RectCollision(rcStick, rcTBox) Then
        StickInTBox = True
        Exit For
    End If
Next i

End Function

Private Sub DrawPlatforms()
Dim i As Integer, j As Single

'Const Spike_Width = 500, Spike_Height = 500

picMain.FillStyle = vbFSTransparent
picMain.DrawWidth = 1

For i = 0 To ubdPlatforms
    modStickGame.sBoxFilled Platform(i).Left, Platform(i).Top, _
        Platform(i).Left + Platform(i).width, Platform(i).Top + Platform(i).height, _
        BoxCol 'IIf(Platform(i).iType = pNormal, BoxCol, vbRed)
    
    
    'PrintStickText "Platform " & CStr(i), Platform(i).Left, Platform(i).Top, vbBlack
    
    
    
'    If Platform(i).iType = pSpikes Then
'        For j = Platform(i).Left To Platform(i).Left + Platform(i).width Step 500
'            DrawTriangle CSng(j), Platform(i).Top, Spike_Width, Spike_Height
'        Next j
'    End If
Next i

picMain.DrawWidth = 5
modStickGame.sBox 0, 0, StickGameWidth, StickGameHeight, BoxCol
'modStickGame.sLine 0, 0, StickGameWidth, 0, Me.BackColor

picMain.DrawWidth = 1
End Sub

'Private Sub DrawTriangle(LeftX As Single, LeftY As Single, tWidth As Single, tHeight As Single)
'Dim TopX As Single, TopY As Single
'
'TopX = LeftX + tWidth / 2
'TopY = LeftY - tHeight
'
'modStickGame.sLine LeftX, LeftY, TopX, TopY
''modStickGame.sLine TopX, TopY, LeftX + width, LeftY
'
'
'End Sub

Private Sub DrawBoxes()
Dim i As Integer

picMain.DrawWidth = 2
For i = 0 To ubdBoxes
    If Box(i).bInUse Then
        modStickGame.sBox Box(i).Left, Box(i).Top, _
            Box(i).Left + Box(i).width, Box(i).Top + Box(i).height, BoxCol
        
        'PrintStickText "Box " & CStr(i), Box(i).Left, Box(i).Top, vbBlack
    End If
Next i
End Sub

Private Sub DrawtBoxes()
Dim i As Integer


picMain.DrawWidth = 2
For i = 0 To ubdtBoxes
    modStickGame.sBoxFilled tBox(i).Left, tBox(i).Top, _
        tBox(i).Left + tBox(i).width, tBox(i).Top + tBox(i).height, BoxCol
    
    'PrintStickText "tBox " & CStr(i), tBox(i).Left, tBox(i).Top, vbBlack
Next i

End Sub

Private Sub DrawCrosshair()

If StickInGame(0) Then
    If Stick(0).bFlashed = False Then
        
        picMain.DrawWidth = 2
        
        'If modStickGame.cg_LaserSight Then
            'picMain.Line (MouseX, Mousey - 20)-(MouseX, Mousey + 20)
            'picMain.Line (MouseX + 20, Mousey)-(MouseX - 20, Mousey)
            'picMain.Circle (MouseX, MouseY), 20, vbBlack
            
        'Else
            
            'If Stick(0).bFlashed Then
                'StunnedMouseX = StunnedMouseX + 10 * Rnd()
                'StunnedMouseY = StunnedMouseY + 10 * Rnd()
                
                'MouseX = StunnedMouseX
                'Mousey = StunnedMouseY
            'Else
                'MouseX = MouseX
                'Mousey = MouseY
            'End If
            
            
            'Red if you can't fire - in the middle of a burst
            If Stick(0).BulletsFired2 = 0 Then
                picMain.ForeColor = vbBlack
            Else
                picMain.ForeColor = vbRed
            End If
            
            DrawCrossHairPoint MouseX, MouseY
            
        'End If
        
        
        
        'Me.modstickgame.sCircle (MouseX, Mousey), 150, vbRed
        'picMain.Line (MouseX - 3, Mousey - 100,MouseX - 3, Mousey + 100), Stick(0).Colour
        'picMain.Line (MouseX - 100, Mousey - 3,MouseX + 100, Mousey - 3), Stick(0).Colour
    End If
End If

End Sub

Private Sub DrawCrossHairPoint(pX As Single, pY As Single)

Select Case Stick(0).WeaponType
    Case eWeaponTypes.W1200, eWeaponTypes.FlameThrower, eWeaponTypes.SPAS
        picMain.Circle (pX, pY), 200
        
    Case eWeaponTypes.AK, eWeaponTypes.XM8, eWeaponTypes.M249, _
            eWeaponTypes.Chopper, eWeaponTypes.AUG, eWeaponTypes.MP5
        
        picMain.Circle (pX, pY), 90
        
        picMain.Line (pX, pY - 150)-(pX, pY - 50)
        picMain.Line (pX, pY + 150)-(pX, pY + 50)
        
        picMain.Line (pX - 150, pY)-(pX - 50, pY)
        picMain.Line (pX + 150, pY)-(pX + 50, pY)
        
        
    Case eWeaponTypes.DEagle, eWeaponTypes.USP, eWeaponTypes.Mac10
        picMain.Circle (pX, pY), 90
        
        picMain.Line (pX, pY - 150)-(pX, pY - 75)
        picMain.Line (pX, pY + 150)-(pX, pY + 75)
        
        picMain.Line (pX - 150, pY)-(pX - 75, pY)
        picMain.Line (pX + 150, pY)-(pX + 75, pY)
        
    Case eWeaponTypes.M82, eWeaponTypes.AWM, eWeaponTypes.G3
        
        picMain.Circle (pX, pY), 75
        
        picMain.DrawWidth = 1
        picMain.Line (pX, pY - 150)-(pX, pY + 150)
        picMain.Line (pX + 150, pY)-(pX - 150, pY)
        
    Case eWeaponTypes.RPG
        picMain.Circle (pX, pY + 50), 100
        picMain.Line (pX + 150, pY + 100)-(pX - 150, pY + 100)
        picMain.Line (pX + 100, pY + 150)-(pX - 100, pY + 150)
        picMain.Line (pX + 50, pY + 200)-(pX - 50, pY + 200)
        
        
    Case eWeaponTypes.Knife
        
        picMain.Line (pX, pY - 150)-(pX, pY + 150)
        picMain.Line (pX + 150, pY)-(pX - 150, pY)
        
End Select

End Sub

Private Sub DisplayChat()

Dim i As Integer, iChat As Integer, iKill As Integer
Dim GTC As Long
Dim sngY As Single
Dim nChat As Integer, nKill As Integer


'################################################################################################
'Stick's personal chat, [u]below[/u] main chat
Const Chat_Show_Time = 10000

picMain.Font.Name = "Times New Roman"
picMain.Font.Size = 10
For i = 0 To NumSticksM1
    If StickInGame(i) Then
        If Stick(i).LastChatMsg + Chat_Show_Time > GetTickCount() Then
            If LenB(Stick(i).curChatMsg) Then
                If CanSeeStick(i) Then
                    If Not (StickiHasState(i, STICK_CROUCH) Or StickiHasState(i, STICK_PRONE)) Then
                        PrintStickText Stick(i).curChatMsg, Stick(i).X - TextWidth(Stick(i).curChatMsg) / 2, Stick(i).Y - 1500, vbBlack
                    End If
                End If
            End If
        End If
    End If
Next i
picMain.Font.Size = 8
picMain.Font.Name = DefaultFontName

'################################################################################################


picMain.Font.Size = 6

'Check if any chat texts have decayed
i = 0
GTC = GetTickCount()
Do While i < NumChat
    'Is it decay time?
    If Chat(i).Decay < GTC Then
        RemoveChatText i
        i = i - 1
    End If
    
    'Increment the counter
    i = i + 1
Loop

'If NumChat > Max_Chat Then
'    'remove numchat - max_chat from the beginning
'    'j = NumChat - Max_Chat + 1
'    'For i = 0 To j
'        'RemoveChatText i
'    'Next i
'    RemoveChatText LBound(Chat)
'End If

'ichatmax = 13
'ikillmax = 30

'CalcChat nChat, nKill

'########################################################
For i = 0 To NumChat - 1
    If Chat(i).bChatMessage Then
        nChat = nChat + 1
    Else
        nKill = nKill + 1
    End If
Next i

'remove first one
If nChat > 13 Then
    For i = 0 To NumChat - 1 'To 0 Step -1
        If Chat(i).bChatMessage Then
            RemoveChatText i
            Exit For
        End If
    Next i
End If
If nKill > 30 Then
    For i = 0 To NumChat - 1 'To 0 Step -1
        If Chat(i).bChatMessage = False Then
            RemoveChatText i
            Exit For
        End If
    Next i
End If
'########################################################


If Not bPlaying Then
    iChat = Chat_Round_Offset
    iKill = Chat_Round_Offset
Else
    iChat = 0
    iKill = 0
End If

For i = 0 To NumChat - 1
    If Chat(i).bChatMessage Then
        sngY = iChat * Chat(i).sTextHeight + Chat_Chat_Offset
        
        picMain_BoxFilled Chat_X_Offset, sngY, Chat_X_Offset + Chat(i).sTextWidth, sngY + Chat(i).sTextHeight, vbWhite
        '                                                                                   ^ yes, that's right. Makes a rect
        PrintStickFormText Chat(i).Text, Chat_X_Offset, iChat * Chat(i).sTextHeight + 1000, Chat(i).colour
        
        iChat = iChat + 1
    Else
        sngY = iKill * Chat(i).sTextHeight + Chat_Kills_Offset
        
        picMain_BoxFilled Chat_X_Offset, sngY, Chat_X_Offset + Chat(i).sTextWidth, sngY + Chat(i).sTextHeight, vbWhite
        '                                                                                   ^ yes, that's right. Makes a rect
        PrintStickFormText Chat(i).Text, Chat_X_Offset, sngY, Chat(i).colour
        
        iKill = iKill + 1
    End If
Next i



End Sub

Private Sub picMain_BoxFilled(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, lCol As Long)
picMain.Line (X1, Y1)-(X2, Y2), lCol, BF
End Sub

Private Function GetReloadTime(iStick As Integer) As Long

'Select Case Stick(iStick).WeaponType
'    Case eWeaponTypes.AK
'        GetReloadTime = AK_Reload_Time
'    Case eWeaponTypes.M82
'        GetReloadTime = M82_Reload_Time
'    Case eWeaponTypes.DEagle
'        GetReloadTime = DEagle_Reload_Time
'    Case eWeaponTypes.W1200
'        GetReloadTime = W1200_Reload_Time
'    Case eWeaponTypes.XM8
'        GetReloadTime = XM8_Reload_Time
'    Case eWeaponTypes.M249
'        GetReloadTime = M249_Reload_Time
'    Case eWeaponTypes.RPG
'        GetReloadTime = RPG_Reload_Time
'    Case eWeaponTypes.FlameThrower
'        GetReloadTime = Flame_Reload_Time
'    Case eWeaponTypes.Chopper
'        GetReloadTime = 1
'End Select

GetReloadTime = kReloadTime(Stick(iStick).WeaponType) / GetSticksTimeZone(iStick)

If Stick(iStick).Perk = pSleightOfHand Then
    GetReloadTime = GetReloadTime / _
        (SleightOfHandReloadDecrease) '* GetTimeZoneAdjust)
'Else
    'GetReloadTime = GetReloadTime '/ GetTimeZoneAdjust
End If

End Function

Private Sub MakeReloadTimeArray()
Dim i As Integer

For i = 0 To eWeaponTypes.Chopper
    Select Case i
        Case eWeaponTypes.AK
            kReloadTime(i) = AK_Reload_Time
        Case eWeaponTypes.M82
            kReloadTime(i) = M82_Reload_Time
        Case eWeaponTypes.DEagle
            kReloadTime(i) = DEagle_Reload_Time
        Case eWeaponTypes.W1200
            kReloadTime(i) = W1200_Reload_Time
        Case eWeaponTypes.XM8
            kReloadTime(i) = XM8_Reload_Time
        Case eWeaponTypes.M249
            kReloadTime(i) = M249_Reload_Time
        Case eWeaponTypes.RPG
            kReloadTime(i) = RPG_Reload_Time
        Case eWeaponTypes.FlameThrower
            kReloadTime(i) = Flame_Reload_Time
        Case eWeaponTypes.AUG
            kReloadTime(i) = AUG_Reload_Time
        Case eWeaponTypes.Chopper
            kReloadTime(i) = 1
        Case eWeaponTypes.USP
            kReloadTime(i) = USP_Reload_Time
        Case eWeaponTypes.AWM
            kReloadTime(i) = AWM_Reload_Time
        Case eWeaponTypes.MP5
            kReloadTime(i) = MP5_Reload_Time
        Case eWeaponTypes.Mac10
            kReloadTime(i) = Mac10_Reload_Time
        Case eWeaponTypes.SPAS
            kReloadTime(i) = SPAS_Reload_Time
        Case eWeaponTypes.G3
            kReloadTime(i) = G3_Reload_Time
    End Select
Next i

End Sub

Private Function GetHighestScorer_i() As Integer

Dim iTempScore As Integer, iMaxScore As Integer, iMaxScoreOwner As Integer
Dim i As Integer

iMaxScore = Stick(0).iKills - Stick(0).iDeaths
iMaxScoreOwner = 0

For i = 1 To NumSticksM1
    iTempScore = Stick(i).iKills - Stick(i).iDeaths
    If iTempScore > iMaxScore Then
        iMaxScoreOwner = i
        iMaxScore = iTempScore
    End If
Next i

GetHighestScorer_i = iMaxScoreOwner

End Function

Private Sub ShowScores()
Const Sp8 As String * 8 = "        "
Const TopY = CentreY - 3000, Score_Box_Width = 9200, Winner_Radius = 50, Winner_RadiusX2 = Winner_Radius * 2
Const TitleOffset = 290
Dim Txt As String
Dim i As Integer
Dim X As Single, Y As Single

picMain.Font.Size = 11
'picMain.Font.Bold = True

On Error Resume Next
X = StickCentreX - 0.2 * Me.width '#######################THIS NEEDS UPDATING WHEN ADDING NEW COLUMNS###########################
'normally 2000         ^

'Select Case modStickGame.sv_GameType
    'Case eStickGameTypes.gCoOp, eStickGameTypes.gElimination
BorderedBox X, TopY - 300, X + Score_Box_Width, Y + 195 * NumSticks + 1100, BoxCol
    'Case Else
        'BorderedBox X, TopY - 300, X + 7000, Y + 195 * NumSticks + 1100, BoxCol
'End Select

Y = TopY + 195 * GetHighestScorer_i() + 30 + Winner_Radius
picMain.FillStyle = vbFSSolid: picMain.FillColor = vbBlue
picMain.Circle (X + Winner_RadiusX2, Y), Winner_Radius, vbBlue
picMain.FillStyle = vbFSTransparent
'picMain.DrawMode = Winner_DrawMode
'picMain.Line (X + 15, Y)-(X + Score_Box_Width - 15, Y + 220), , BF
''                         ^ = textheight("W")
'picMain.DrawMode = vbCopyPen


'#################################################################################
PrintStickFormText Sp8 & "Name" & Sp8, X, TopY - TitleOffset, vbBlack
'                                 290 = TextHeight(Txt)*1.5

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(Stick(i).Name), 20)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
    'PrintStickFormText CStr(Stick(i).ID), X - 500, TopY + TextHeight(Txt) * i, Stick(i).Colour
Next i
'#################################################################################
X = X + 1500
PrintStickFormText " Score ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(CStr(Stick(i).iKills - Stick(i).iDeaths)), 6)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
Next i

'#################################################################################
X = X + 1000
PrintStickFormText " Kills ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(CStr(Stick(i).iKills)), 6)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
Next i
'#################################################################################
X = X + 1000
PrintStickFormText "Deaths ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(CStr(Stick(i).iDeaths)), 8)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
Next i
'#################################################################################
X = X + 1000
PrintStickFormText " Team ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(GetTeamStr(Stick(i).Team)), 10)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, GetTeamColour(Stick(i).Team)
Next i
'#################################################################################
X = X + 1000
PrintStickFormText " Row Kills ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(CStr(Stick(i).iKillsInARow)), 11)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
Next i
'#################################################################################
X = X + 1000
PrintStickFormText "     Perk   ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(GetPerkName(Stick(i).Perk), 20)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
Next i
'#################################################################################
'Select Case modStickGame.sv_GameType
    'Case gElimination, gCoOp
        X = X + 2000
        PrintStickFormText " Status ", X, TopY - TitleOffset, vbBlack
        
        For i = 0 To NumSticksM1
            Txt = IIf(StickInGame(i), "  Alive", "  Dead")
            PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).colour
        Next i
'End Select



'#################################################################################
'Extra Stuff
'#################################################################################
If modStickGame.sv_GameType = gDeathMatch Then
    PrintStickFormText "Score To Win: " & CStr(modStickGame.sv_WinScore), 6000, 20, vbBlack
End If

PrintStickFormText "Game Type: " & GetGameType(), 9000, 20, vbBlack

'PrintStickFormText "Grenades Shot Down: " & NadesShot, 12600, 20, vbBlack

picMain.Font.Size = 8
'picMain.Font.Bold = False

End Sub

Private Function GetPerkName(vPerk As eStickPerks) As String
GetPerkName = kPerkName(vPerk)
End Function

Private Sub MakePerkNameArray()
Dim i As Integer

For i = 0 To eStickPerks.pSpy
    If i = pConditioning Then
        kPerkName(i) = "Conditioning"
    ElseIf i = pJuggernaut Then
        kPerkName(i) = "Juggernaut"
    ElseIf i = pNone Then
        kPerkName(i) = "None"
    ElseIf i = pBombSquad Then
        kPerkName(i) = "Bomb Squad" '"Jammer + Awareness"
    ElseIf i = pSleightOfHand Then
        kPerkName(i) = "Sleight of Hand"
    ElseIf i = pSpy Then
        kPerkName(i) = "Spy"
    ElseIf i = pSniper Then
        kPerkName(i) = "Sniper/Stealth"
    ElseIf i = pStoppingPower Then
        kPerkName(i) = "Stopping Power"
    ElseIf i = pFocus Then
        kPerkName(i) = "Focus"
    ElseIf i = pMartyrdom Then
        kPerkName(i) = "Martyrdom"
    ElseIf i = pSteadyAim Then
        kPerkName(i) = "Steady Aim"
    ElseIf i = pMechanic Then
        kPerkName(i) = "Rapid Fire"
    ElseIf i = pDeepImpact Then
        kPerkName(i) = "Deep Impact"
    ElseIf i = pZombie Then
        kPerkName(i) = "Zombie"
    Else
        kPerkName(i) = "Forgot This.."
    End If
Next i

End Sub

Private Sub DrawNames()
Dim Txt As String
Dim i As Integer, j As Integer
Dim Col As Long
Dim sName As String
Dim bAlly As Boolean
Dim bIAmSpec As Boolean
Dim bTest As Boolean

'Txt = Trim$(Stick(0).Name) & IIf(StickiHasState(0, Stick_Reload), " (Reloading)", vbNullString)
'PrintStickText Txt, Stick(0).X - TextWidth(Txt) / 2, GetStickY(0) - 250, vbBlack

bIAmSpec = Not StickInGame(0)

For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        If bIAmSpec Then
            bTest = True
        Else
            bTest = CanSeeStick(i)
        End If
        
        If bTest Then
            
            If CBool(i) Then
                bAlly = IsAlly(Stick(0).Team, Stick(i).Team)
                If bAlly Then
                    Col = Ally_Colour
                Else
                    Col = Enemy_Colour
                End If
            Else
                'bally=False
                Col = vbBlack
            End If
            
            If Stick(i).WeaponType <> Chopper Then
                
                If Stick(i).Perk = pSpy Then
                    j = FindStick(Stick(i).MaskID)
                    If j = -1 Then j = 0
                    sName = Trim$(Stick(j).Name)
                Else
                    sName = Trim$(Stick(i).Name)
                End If
                
                
                
                If StickiHasState(i, STICK_PRONE) Then
                    
                    If bIAmSpec Or bAlly Then
                        bTest = True
                    Else
                        bTest = (Stick(i).Perk <> pSniper)
                    End If
                    
                    If bTest Then
                        Txt = sName & ", " & CStr(Stick(i).Health) & IIf(Stick(i).Shield, ", " & CStr(Round(Stick(i).Shield)), vbNullString)
                        
                        PrintStickText Txt, Stick(i).X - TextWidth(Txt) / 2, GetStickY(i) - 250, Col
                    End If
                    
                Else
                    
                    If bIAmSpec Or bAlly Then
                        bTest = True
                    Else
                        bTest = Not (Stick(i).Perk = pSniper And StickiHasState(i, STICK_CROUCH))
                    End If
                    
                    If bTest Then
                        
                        Txt = sName & IIf(StickiHasState(i, STICK_RELOAD), " [Reloading]", vbNullString)
                        PrintStickText Txt, CSng(Stick(i).X - TextWidth(Txt) / 2), GetStickY(i) - 250, Col
                        
                        Txt = CStr(Stick(i).Health) & IIf(Stick(i).Shield, ", " & CStr(Round(Stick(i).Shield)), vbNullString)
                        PrintStickText Txt, CSng(Stick(i).X - TextWidth(Txt) / 2), GetStickY(i) - 500, Col
                        
                        'DrawSemiCircle Stick(i).X, Stick(i).Y - BodyLen, vbGreen, vbRed, Stick(i).Health / Health_Start, 600
                        
                    End If
                    
                End If
            Else
                Txt = Trim$(Stick(i).Name) '& IIf(StickiHasState(i, Stick_Reload), " (Reloading)", vbNullString)
                PrintStickText Txt, Stick(i).X - TextWidth(Txt) / 2, Stick(i).Y - 500, Col
                
                Txt = "Health: " & CStr(Stick(i).Health) '& IIf(Stick(i).Armour > 0, "   Armour: " & CStr(Stick(i).Armour), vbNullString)
                PrintStickText Txt, Stick(i).X - TextWidth(Txt) / 2, Stick(i).Y - 800, Col
            End If
        End If
    End If
Next i

End Sub

Private Sub DisplayHUD()
Dim Txt As String
Dim TimeLeft As Single
Dim X As Single, Y As Single, Adj As Single
Dim mR As Integer, MaxRounds As Integer, Reload_Time As Long
Dim i As Integer
Dim bChopper As Boolean

Dim SemiX As Single, SemiY As Single

'Txt = "Health: " & CStr(Round(Stick(0).Health))
'PrintStickText Txt, Me.width / 2 - TextWidth(Txt) / 2, TextHeight(Txt) + 500, vbblack

If modStickGame.cg_DrawFPS Then
    PrintStickFormText "FPS: " & CStr(FPS), StickCentreX - 300, StickCentreY + 2000, vbBlack
    'PrintStickFormText "Elapsed: " & CStr(modStickGame.StickElapsedTime), 10, 820, vbBlack
    'Debug.Print FPS & ", " & modStickGame.StickElapsedTime
End If


If ShowScoresKey Then
    ShowScores
ElseIf StickInGame(0) = False Then
    DisplayScoreBoard
Else
    'in game, and not showing scores
    PrintStickFormText "Kills in a Row: " & CStr(Stick(0).iKillsInARow), 10, 10, vbBlack
    PrintStickFormText "Score: " & CStr(Stick(i).iKills - Stick(i).iDeaths), 10, 260, vbBlack
End If


If LastZoomPress + ZoomShowTime > GetTickCount() Then
    PrintStickFormText "Zoom: " & FormatNumber$(cg_sZoom, 2, vbTrue, vbFalse, vbFalse), _
        StickCentreX, StickCentreY - 4000, vbBlack
End If

'If Scroll_WeaponKey <> Stick(0).WeaponType Then
    'Txt = "Switching Weapon: " & GetWeaponName(Scroll_WeaponKey) '& Space$(2) & _
        CStr(Format((Scroll_Delay - GetTickCount() + LastScrollWeaponSwitch) / 1000, "0.00"))
    
    'PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY - 1000, vbBlack
'End If


picMain.DrawWidth = 3
If StickInGame(0) Then
    
    bChopper = (Stick(0).WeaponType = Chopper)
    
    SemiY = Me.ScaleHeight
    SemiX = Me.ScaleWidth - 700
    
    If bChopper = False Then
        'HUD box
        picMain.Line (Me.ScaleWidth - 2700, SemiY - 1250)-(Me.ScaleWidth - 50, Me.ScaleHeight), vbBlue, B
    Else
        picMain.Line (Me.ScaleWidth - 1500, SemiY - 880)-(Me.ScaleWidth - 50, Me.ScaleHeight), vbBlue, B
    End If
    
    
    If Stick(0).WeaponType <> Knife Then
        If Stick(0).WeaponType <> Chopper Then
            
            MaxRounds = GetMaxRounds(Stick(0).WeaponType)
            'If Stick(0).WeaponType = W1200 Then
                'mR = (MaxRounds - Stick(0).BulletsFired) / W1200_Gauge
            'Else
                mR = MaxRounds - Stick(0).BulletsFired
            'End If
            
            If StickiHasState(0, STICK_RELOAD) And Not WeaponIsShotgun(Stick(0).WeaponType) Then
                
                Reload_Time = GetReloadTime(0)
                
                TimeLeft = (Reload_Time - GetTickCount() + Stick(0).ReloadStart)
                
                DrawSemiCircle SemiX, SemiY, _
                    vbBlue, vbRed, 1 - TimeLeft / Reload_Time, 600
                
                DrawSemiCircle StickCentreX, StickCentreY - 2000, vbBlue, vbRed, 1 - TimeLeft / Reload_Time, 600
                PrintStickFormText "Reload Progress", StickCentreX - 600, StickCentreY - 1800, vbBlack
                
        '        x = StickCentreX - Reload_Time / 2
        '        y = StickCentreY - 650
        '
        '        picMain.Line (x, y)-(x + Reload_Time, y), vbRed
        '        picMain.Line (x, y)-(x + Reload_Time - TimeLeft, y), vbBlue
                
            Else
                
                DrawSemiCircle SemiX, SemiY, _
                    vbBlue, vbRed, mR / MaxRounds, 600
                
            End If
            
            If WeaponIsShotgun(Stick(0).WeaponType) Then mR = mR / GetGauge(Stick(0).WeaponType)
            
            
            Txt = "Rounds: " & CStr(mR)
            PrintStickFormText Txt, SemiX - TextWidth(Txt) / 2, SemiY - 300, vbBlack
            
            If StickiHasState(0, STICK_RELOAD) = False Then
                If Stick(0).WeaponType <> RPG Then
                    If WeaponIsShotgun(Stick(0).WeaponType) Then
                        If mR < (MaxRounds * 0.3 / GetGauge(Stick(0).WeaponType)) Then
                            PrintStickFormText "Low Ammo", StickCentreX - 400, StickCentreY - 750, vbRed
                        End If
                    ElseIf mR < (MaxRounds * 0.3) Then
                        PrintStickFormText "Low Ammo", StickCentreX - 400, StickCentreY - 750, vbRed
                    End If
                    
                ElseIf TotalMags(RPG) = 0 Then
                    
                    PrintStickFormText "Low Ammo", StickCentreX - 400, StickCentreY - 750, vbRed
                End If
            End If
            
            
            Txt = "Total " & GetMagName(Stick(0).WeaponType) & ": " & TotalMags(Stick(0).WeaponType)
            PrintStickFormText Txt, SemiX + 650 - TextWidth(Txt), SemiY - 1500, vbBlack
            
            If TotalMags(Stick(0).WeaponType) = 0 Then
                If Stick(0).WeaponType <> RPG Or Stick(0).BulletsFired = 1 Then
                    Txt = "No " & GetMagName(Stick(0).WeaponType) & " left - Find some ammo"
                    PrintStickFormText Txt, StickCentreX - 1300, CentreY - 900, vbRed
                    'PrintStickFormText Txt, SemiX - TextWidth(Txt) / 2, SemiY - 1750, vbRed
                End If
            End If
            
            
        End If 'weapon type endifs
    End If
    
    Adj = GetMyTimeZone()
    PrintStickFormText "Time Zone: " & FormatNumber$(Adj, 2, vbTrue, vbFalse, vbFalse), SemiX - 500, SemiY - 1750, vbBlack
    
    If Stick(0).WeaponType <> Chopper Then
        
        X = SemiX - 700
        Y = SemiY - 950
        
        picMain.FontBold = True
        picMain.DrawWidth = 3
        
        
        
        If Stick(0).LastNade + Nade_Delay / Adj < GetTickCount() Then
            Txt = GetNadeTypeName() & " Ready" 'IIf(Stick(0).iNadeType = nFrag, "Grenade Ready", "Flash-Bang Ready")
            PrintStickFormText Txt, X - TextWidth(Txt) / 2, Y, vbGreen
            'C = MGreen
        Else
            'Txt = "Grenade Not Ready"
            
            TimeLeft = Nade_Delay + (Stick(0).LastNade - GetTickCount()) * Adj
            
            Me.picMain.Line (SemiX - 1900, SemiY - 850)-(SemiX - 1900 + TimeLeft / 2, SemiY - 850), vbRed
            
        End If
        
        Y = Y - 250
        If Stick(0).LastMine + Mine_Delay / Adj < GetTickCount() Then
            Txt = "Mine Ready"
            PrintStickFormText Txt, X - TextWidth(Txt) / 2, Y, vbGreen
            'C = MGreen
        Else
            'Txt = "GreMine Not Ready"
            
            TimeLeft = Mine_Delay + (Stick(0).LastMine - GetTickCount()) * Adj
            
            Y = Y + 100
            Me.picMain.Line (SemiX - 1900, Y)-(SemiX - 1900 + TimeLeft / 6.2, Y), vbRed
            
        End If
        
        Me.picMain.FontBold = False
        'X = 200 + TextWidth(Txt)
        'Y = SemiY - 150
        'picMain.Line (X - K, Y - K)-(X + K, Y + K), vbBlack
        'picMain.Circle (X, Y), Nade_Radius, vbBlack
        
        
        Txt = "Weapon: " & GetWeaponName(Stick(0).WeaponType)
        PrintStickFormText Txt, X - TextWidth(Txt) / 2, SemiY - 700, vbBlack
    End If
    
    If bChopper Then
        
        Me.picMain.FontBold = True
        Me.picMain.DrawWidth = 3
        
        
        X = SemiX - 730
        Y = SemiY - 700
        
        
        If Stick(0).LastNade + Chopper_RPG_Delay / Adj < GetTickCount() Then
            Txt = "Rocket Ready"
            PrintStickFormText Txt, X + 550 - TextWidth(Txt) / 2, Y - 100, vbGreen
        Else
            TimeLeft = Chopper_RPG_Delay + (Stick(0).LastNade - GetTickCount()) * Adj
            
            Me.picMain.Line (X, Y)-(X + TimeLeft / 2, Y), vbRed
        End If
        
        Me.picMain.FontBold = False
        
        X = Me.ScaleWidth - 750
    Else
        X = Me.ScaleWidth - 2000
        If ChopperAvail Then
            If modStickGame.sv_AllowedWeapons(eWeaponTypes.Chopper) = False Then
                ChopperAvail = False
            Else
                picMain.Font.Bold = True
                PrintStickFormText "Chopper Available - Press 0", 10, 510, vbBlack
                picMain.Font.Bold = False
            End If
        End If
    End If
    Txt = "Health: " & CStr(Stick(0).Health)
    PrintStickFormText Txt, X - TextWidth(Txt) / 2, SemiY - 350, vbBlack
    DrawSemiCircle X, SemiY, MGreen, vbBlue, Stick(0).Health / Current_Health_Start, 600
    
    
    Txt = "Shield: " & CStr(Stick(0).Shield)
    PrintStickFormText Txt, X - TextWidth(Txt) / 2, SemiY - 200, vbBlack
    'Dim sRatio As Single
    'sRatio = Stick(0).Shield / Max_Shield
    'Y = Stick(0).Y - 350
    DrawSemiCircle X, SemiY, MSilver, vbBlue, Stick(0).Shield / Max_Shield, 675
    'modStickGame.sLine Stick(0).X - ArmLen, Y, Stick(0).X + sRatio * 1000, Y
End If


If Not modStickGame.StickServer Then
    If (LastUpdatePacket + LagOut_Delay) < GetTickCount() Then
        If LastUpdatePacket Then
            picMain.Font.Size = 16
            picMain.ForeColor = &HC0C0C0
            
            PrintStickFormText "Connection Interrupted", StickCentreX - 750, StickCentreY - 2000, vbRed '&HC0C0C0
            
            picMain.Font.Size = 8
        End If
    End If
End If


End Sub

Private Sub ShowChatEntry()
Dim sTxt As String
Dim bCan As Boolean

Const PlayingY = 2300, RoundY = 100
Const RoundX = 1500

'###########
'show chat
If bChatActive Then
    'Me.picMain.ForeColor = Stick(0).Colour
    
    If ChatFlashDelay < 0 Then
        bCan = True
    ElseIf (LastFlash + ChatFlashDelay) < GetTickCount() Then
        bCan = True
    End If
    
    If bCan Then
        bChatCursor = Not bChatCursor
        LastFlash = GetTickCount()
    End If
    
    sTxt = Trim$(Stick(0).Name) & modMessaging.MsgNameSeparator & strChat
    
    If bPlaying Then
        PrintStickFormText sTxt & IIf(bChatCursor, "_", vbNullString), StickCentreX - TextWidth(sTxt) / 2, PlayingY + 300, vbBlack 'Stick(0).Colour
        PrintStickFormText "Escape to Exit Chat", PlayingX, PlayingY, vbBlack 'Stick(0).Colour
    Else
        PrintStickFormText sTxt & IIf(bChatCursor, "_", vbNullString), RoundX + 500 - TextWidth(sTxt) / 2, RoundY + 300, vbBlack 'Stick(0).Colour
        PrintStickFormText "Escape to Exit Chat", RoundX, RoundY, vbBlack 'Stick(0).Colour
    End If
    
End If

End Sub

Private Sub DrawSemiCircle(tX As Single, tY As Single, _
    ForeCol As Long, BackCol As Long, _
    sAmountFull As Single, sWidth As Single)

Dim Start As Single
picMain.DrawWidth = 3

If sAmountFull < 1 Then
    picMain.Circle (tX, tY), sWidth, BackCol, 0, Pi, 0.75
End If

If sAmountFull Then
    On Error Resume Next
    Start = Pi * (1 - sAmountFull)
    If Start > (Pi - 0.1) Then 'pi*179 / 180
        Start = Pi - 0.1
    End If
    
    picMain.Circle (tX, tY), sWidth, ForeCol, Abs(Start), Pi, 0.75
End If

End Sub

Private Sub DrawLaserSight()
Dim Facing As Single
Dim tX As Single, tY As Single 'end point
Dim OldtX As Single, OldtY As Single
Dim i As Integer
Dim SF100 As Single
Const nLines = 10
Const Laser_Len = 12000

If modStickGame.cg_LaserSight Then
    If StickInGame(0) Then
        'If Stick(0).bFlashed = False Then
            If StickiHasState(0, STICK_RELOAD) = False Then
                'If StickiHasState(0, Stick_Prone) = False Then
                    If Stick(0).WeaponType <> Knife Then
                        If Stick(0).WeaponType <> Chopper Then
                            
                            'aFacing = FindAngle(CSng(Stick(0).GunPoint.x), CSng(Stick(0).GunPoint.y), MouseX, MouseY)
                            
                            'Facing = FindAngle(CSng(Stick(0).GunPoint.x), CSng(Stick(0).GunPoint.y), MouseX, MouseY) 'Stick(0).Facing
                            '.0349 = pi*2/180 = 2 degrees
                            
                            
                            Facing = Stick(0).Facing + Sine(Stick(0).Facing) / 100 '0.0349
                            
                            OldtX = Stick(0).GunPoint.X '+ SF100
                            OldtY = Stick(0).GunPoint.Y '+ SF100
                            
                            picMain.DrawWidth = 1
                            
                            'gradient'd line
                            For i = 1 To nLines
                                
                                tX = Stick(0).GunPoint.X + Sine(Facing) * Laser_Len * i / nLines '+ SF100
                                tY = Stick(0).GunPoint.Y - CoSine(Facing) * Laser_Len * i / nLines '+ SF100
                                
                                picMain.ForeColor = RGB(255 - i * 10, i * 10, i * 30)
                                modStickGame.sLine OldtX, OldtY, tX, tY
                                
                                OldtX = tX
                                OldtY = tY
                                
                            Next i
                        Else
                            'is chopper
                            modStickGame.cg_LaserSight = False
                        End If 'chopper endif
                    End If 'knife endif
                'End If 'prone endif
            End If 'reload endif
        'End If 'flashed endif
    End If 'ingame endif
End If

End Sub

Private Sub DrawSilencer(X As Single, Y As Single, Facing As Single)
Dim X2 As Single, Y2 As Single
Const SilencerLen = 100

X2 = X + SilencerLen * Sine(Facing)
Y2 = Y - SilencerLen * CoSine(Facing)

picMain.DrawWidth = 2
picMain.ForeColor = vbBlack
modStickGame.sLine X, Y, X2, Y2

End Sub

Private Sub LimitSpeed(i As Integer)
Dim MaxSpeed As Single, XComp As Single, YComp As Single, Max_Y_Speed As Single
Const def_Max_Y_Speed = 170 'never going to get here, but in case a nade goes off under you...

Max_Y_Speed = IIf(Stick(i).WeaponType = Chopper, 100, def_Max_Y_Speed)

MaxSpeed = GetMaxSpeed(i)
XComp = Stick(i).Speed * Sine(Stick(i).Heading)
YComp = Stick(i).Speed * CoSine(Stick(i).Heading)


'only limit x-speed, let the stick fall at whatever speed
If Abs(XComp) > MaxSpeed Then
    XComp = Sgn(XComp) * MaxSpeed
    
    Stick(i).Speed = Sqr(XComp * XComp + YComp * YComp)
    
    If YComp > 0 Then
        Stick(i).Heading = Atn(XComp / YComp)
    ElseIf YComp < 0 Then
        Stick(i).Heading = Atn(XComp / YComp) + Pi
    ElseIf XComp > 0 Then
        Stick(i).Heading = piD2
    Else
        Stick(i).Heading = pi3D2
    End If
    
End If
If Abs(YComp) > Max_Y_Speed Then
    YComp = Sgn(YComp) * Max_Y_Speed
    
    Stick(i).Speed = Sqr(XComp * XComp + YComp * YComp)
    
    If YComp > 0 Then
        Stick(i).Heading = Atn(XComp / YComp)
    ElseIf YComp < 0 Then
        Stick(i).Heading = Atn(XComp / YComp) + Pi
    ElseIf XComp > 0 Then
        Stick(i).Heading = piD2
    Else
        Stick(i).Heading = pi3D2
    End If
End If



End Sub

Private Sub Physics()

Const mPacket_LAG_KILLX2 = mPacket_LAG_KILL * 2

Dim i As Integer, j As Integer
Dim TempMag As Single
Dim TempDir As Single, BaseTempDir As Single
Dim Tmp As Integer
Dim Bullet_Delay As Long
Dim Stick_Moving As Boolean, bLBound As Boolean
Dim XComp As Single, YComp As Single, Adj As Single
Dim bThrow As Boolean

Dim tX As Single, tY As Single


i = 1
Do While i < NumSticks
    If Not Stick(i).IsBot Then
        'Time to remove this Stick?
        If (Stick(i).LastPacket + mPacket_LAG_KILL < GetTickCount()) Then
            'Remove!
            If StickServer Then
                SendChatPacketBroadcast Trim$(Stick(i).Name) & " lagged out", vbRed
            End If
            
            RemoveStick i
            i = i - 1
        End If
    End If
    'Increment counter
    i = i + 1
Loop


'check if i lagged out
If Not modStickGame.StickServer Then
    If (LastUpdatePacket + mPacket_LAG_KILLX2) < GetTickCount() Then
        If LastUpdatePacket Then
            bRunning = False
            AddText "Error - Lagged Out (No Packet Flow)", TxtError, True
            Exit Sub
        End If
    End If
End If


'With Stick(0)
'    modstickgame.sLine .X, .Y + BodyLen,.X + 1000 * sine(.ActualFacing), .Y - 1000 * cosine(.ActualFacing)), vbRed
'End With

'Loop through each Stick and perform physics

If Stick(0).bSilenced Then
    Stick(0).bSilenced = WeaponSilencable(Stick(0).WeaponType)
End If


For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        
        If i = 0 Or Stick(i).IsBot Then
            DoReload i
        End If
        
        
        'Cap Speed
        LimitSpeed i
        
        
        Stick_Moving = False
        bLBound = True 'should be true
        
        'Check lag tol
        If i > 0 Then
            If Not Stick(i).IsBot Then
                If (Stick(i).LastPacket + mPacket_LAG_TOL) < GetTickCount() Then
                    SetStickiState i, STICK_NONE
                End If
            End If
        End If
        
        Stick(i).Facing = FixAngle(Stick(i).Facing)
        Stick(i).ActualFacing = FixAngle(Stick(i).ActualFacing)
        
        
        Adj = GetSticksTimeZone(i)
        
        
        'Firing
        If (Stick(i).state And STICK_RELOAD) = 0 Then
            If Stick(i).state And STICK_FIRE Then
                
                Bullet_Delay = GetBulletDelay(i)
                
                
                If Stick(i).BulletsFired < GetMaxRounds(Stick(i).WeaponType) Or ForeignStick(i) Then
                    '                                                           ^ allow clients/others to keep shooting
                    
                    If Stick(i).LastBullet + Bullet_Delay / Adj < GetTickCount() Then
                        
                        
                        tX = Stick(i).GunPoint.X
                        tY = Stick(i).GunPoint.Y
                        
                        If Stick(i).WeaponType <> Knife Then
                            
                            '###########################
                            'work out BaseMag and BaseDir
                            If Stick(i).WeaponType <> M82 Then
                                If Stick(i).WeaponType <> RPG Then
                                    If Stick(i).WeaponType <> Chopper Then
                                        If Stick(i).WeaponType <> FlameThrower Then
                                            AddVectors Stick(i).Speed / 4, Stick(i).Heading, BULLET_SPEED, Stick(i).ActualFacing, TempMag, BaseTempDir
                                        Else
                                            'flame only
                                            AddVectors Stick(i).Speed / Flame_Inertia_Reduction, Stick(i).Heading, Flame_Speed, Stick(i).ActualFacing, TempMag, BaseTempDir
                                        End If
                                    Else
                                        BaseTempDir = Stick(i).ActualFacing
                                    End If
                                End If
                            End If
                            '###########################
                            
                            
                            'sin because when facing up/down, it is dead on (and sin 0 or sin 180 = 0)
                            If Stick(i).ActualFacing > Pi Then
                                Tmp = -1
                                Stick(i).RecoilLeft = True
                            Else
                                Tmp = 1
                                Stick(i).RecoilLeft = False
                            End If
                            
                            
                            
                            'If kBurstBullets(Stick(i).WeaponType) Then
                            If Stick(i).Burst_Bullets Then
                                'If Stick(i).BulletsFired2 < kBurstBullets(Stick(i).WeaponType) Then
                                If Stick(i).BulletsFired2 < Stick(i).Burst_Bullets Then
                                    
                                    FireShot i, BaseTempDir, TempMag, Tmp, tX, tY, Stick_Moving, bLBound
                                    
                                    Stick(i).LastBullet = GetTickCount()
                                    
                                ElseIf Stick(i).LastBullet + Stick(i).Burst_Delay / Adj < GetTickCount() Then
                                    
                                    Stick(i).BulletsFired2 = 0
                                    
                                End If
                                
                            Else
                                FireShot i, BaseTempDir, TempMag, Tmp, tX, tY, Stick_Moving, bLBound
                                Stick(i).LastBullet = GetTickCount()
                            End If
                            
                            
                            If Stick(i).bSilenced = False Then
                                Stick(i).LastLoudBullet = GetTickCount()
                            End If
                            
                            
                        Else 'If Stick(i).WeaponType = Knife Then 'knife else
                            
                            'have knife and are 'firing'
                            
                            For j = 0 To NumSticksM1
                                If j <> i Then
                                    If StickInGame(j) Then
                                        If Not IsAlly(Stick(j).Team, Stick(i).Team) Then
                                            If StickInvul(j) = False Then
                                                If CoOrdNearStick(Stick(i).GunPoint.X, Stick(i).GunPoint.Y, j) Then
                                                    AddBloodExplosion Stick(i).GunPoint.X, Stick(i).GunPoint.Y
                                                    
                                                    If PointHearableOnSticksScreen(Stick(j).X, Stick(j).Y, 0) Then
                                                        If Stick(i).bLightSaber Then
                                                            modAudio.PlayLightSaberSound GetRelPan(Stick(j).X)
                                                        Else
                                                            modAudio.PlayWeaponSound_Panned Knife, GetRelPan(Stick(j).X)
                                                        End If
                                                    End If
                                                    
                                                    If j = 0 Or Stick(j).IsBot Then
                                                        Call Killed(j, i, IIf(Stick(i).bLightSaber, kLightSaber, kKnife))
                                                    End If
                                                    
                                                    Exit For
                                                End If 'co-ord endif
                                            End If 'invul endif
                                        End If 'team endif
                                    End If 'stickingame endif
                                End If 'j<>i endif
                            Next j
                            
                        End If 'knife endif
                        
                    End If 'bullet_delay
                    
                    
                    
                    If i = 0 Then
                        If Not FireKey Then
                            If FireKeyUpTime + Bullet_Release_Delay < GetTickCount() Then
                                SubStickiState i, STICK_FIRE
                                Stick(i).BulletsFired2 = 0
                            End If
                        End If
                    End If
                    
                    
                End If 'bullets fired
                
            End If 'state_fire
            
        End If 'state_reload
        
        
        
        If Stick(i).Perk = pConditioning Then
            TempMag = Accel * ConditioningAccelInc
        'ElseIf Stick(i).Perk = pZombie Then
            'TempMag = Accel * ZombieAccelDec
        Else
            TempMag = Accel
        End If
        
        TempMag = TempMag * modStickGame.StickTimeFactor / Adj
        
        'If Stick(i).OnSurface Then
            If Stick(i).state And STICK_LEFT Then
                'If Stick(i).OnSurface Then
                    'Apply acceleration
                    AddVectors Stick(i).Speed, Stick(i).Heading, TempMag, CSng(pi3D2), Stick(i).Speed, Stick(i).Heading
                'End If
                
                Stick_Moving = True
                
            ElseIf Stick(i).state And STICK_RIGHT Then
                'If Stick(i).OnSurface Then
                    'Apply reverse acceleration
                    AddVectors Stick(i).Speed, Stick(i).Heading, TempMag, CSng(piD2), Stick(i).Speed, Stick(i).Heading
                'End If
                
                Stick_Moving = True
                
            End If
        'End If
        
        '################################################################################
        
        
        'stickmotion
        Dim sticklasty As Single: sticklasty = Stick(i).Y
        MotionStickObject Stick(i).X, Stick(i).Y, Stick(i).Speed, Stick(i).Heading
        
        '################################################################################
        
        
        If Stick(i).WeaponType <> Chopper Then
            Call DoRecoil(i, Stick_Moving, bLBound)
            
            
            If Stick(i).state And STICK_NADE Then
                
                If Stick(i).LastNade + Nade_Delay / GetSticksTimeZone(i) < GetTickCount() Then
                    
                    If i = 0 Then
                        
                        If Stick(0).NadeStart + Client_Nade_Delay < GetTickCount() Then
                            bThrow = True
                        End If
                        
                    Else
                        bThrow = True
                    End If
                    
                    
                    If bThrow Then
                        'needs to be .actualfacing so recoil doesn't affect it
                        AddVectors Stick(i).Speed, Stick(i).Heading, Throwing_Strength, _
                            Stick(i).ActualFacing + IIf(Stick(i).ActualFacing > Pi, piD6, -piD6) * Abs(Sine(Stick(i).ActualFacing)), _
                            TempMag, TempDir
                        
                        
                        AddNade Stick(i).X, Stick(i).Y, TempDir, TempMag, i, Stick(i).colour, Stick(i).iNadeType
                        
                        'If Stick(i).iNadeType = nSmoke Then
                            'Stick(i).LastNade = GetTickCount() + Nade_Time / 2
                        'Else
                        Stick(i).LastNade = GetTickCount()
                        'End If
                        
                    End If
                    
                ElseIf Stick(i).LastNade + Nade_Release_Delay < GetTickCount() Then
                    SubStickiState i, STICK_NADE
                    
                End If
                
            ElseIf Stick(i).state And STICK_MINE Then
                
                If Stick(i).LastMine + Mine_Delay / Adj < GetTickCount() Then
                    
                    'If Stick(i).OnSurface Then
                    AddMine Stick(i).X, Stick(i).Y + Mine_Y_Increase, Stick(i).ID, Stick(i).colour, Stick(i).Heading, Stick(i).Speed
                    
                    Stick(i).LastMine = GetTickCount()
                    'End If
                    
                ElseIf Stick(i).LastMine + Nade_Release_Delay < GetTickCount() Then
                    SubStickiState i, STICK_MINE
                    
                End If
                
            End If
            
            
            If Stick(i).state And STICK_JUMP Then
                
                Stick_Moving = True
                
                If Stick(i).bOnSurface = False Then 'Stick(i).StartJumpTime + JumpTime < GetTickCount() Then
                    
                    If Stick(i).IsBot = False Then
                        SubStickiState i, STICK_JUMP
                    End If
                    
                    'Debug.Print "Sub " & Rnd & vbNewLine
                Else
                    'If StickiHasState(i, stick_Left) Then SubStickiState i, stick_Left
                    'If StickiHasState(i, stick_Right) Then SubStickiState i, stick_Right
                    
                    AddVectors Stick(i).Speed, Stick(i).Heading, JumpMultiple, 0, Stick(i).Speed, Stick(i).Heading
                    
                    
                    'Stick(i).JumpStartY = Stick(i).Y
                End If
                
            End If
            
            
            If Stick(i).state And STICK_RELOAD Then
                If Stick(i).bHadMag = False Then
                    If Not (i = 0 Or Stick(i).IsBot) Then
                        'reset a remote stick's bullets fired
                        Stick(i).BulletsFired = 0
                    End If
                    
                    AddMagForStick i
                    Stick(i).bHadMag = True
                End If
            ElseIf Stick(i).bHadMag Then
                Stick(i).bHadMag = False
            End If
            
            
            ApplyGravity i, sticklasty
            
            
            If Stick_Moving = False Then
                'friction
                
                XComp = Stick(i).Speed * Sine(Stick(i).Heading)
                
                If Abs(XComp) > 4 Then
                    XComp = XComp / 1.2
                    
                    YComp = Stick(i).Speed * CoSine(Stick(i).Heading)
                    
                    Stick(i).Speed = Sqr(XComp * XComp + YComp * YComp)
                    
                End If
            End If
            
            
        Else
            
            Stick(i).bOnSurface = True
            
            If StickiHasState(i, STICK_JUMP) Then
                'If (Stick(i).StartJumpTime + JumpTime) < GetTickCount() Then
                    'If i = 0 Then JumpKey = False
                    'SubStickiState i, Stick_Jump
                'Else
                    Stick_Moving = True
                    AddVectors Stick(i).Speed, Stick(i).Heading, Chopper_Lift * 2, 0, Stick(i).Speed, Stick(i).Heading
                    If Stick(i).IsBot = False Then SubStickiState i, STICK_JUMP
                'End If
            ElseIf StickiHasState(i, STICK_CROUCH) Then
                Stick_Moving = True
                AddVectors Stick(i).Speed, Stick(i).Heading, Chopper_Lift, Pi, Stick(i).Speed, Stick(i).Heading
                If Stick(i).IsBot = False Then SubStickiState i, STICK_CROUCH
            End If
            
'            'chopper physics
'            If Stick(i).Y + CLD2 > 7000 Then
'                Stick(i).Y = 7000 - CLD2
'                Stick(i).Speed = 0
'                'ReverseYComp Stick(i).Speed, Stick(i).Heading
'            ElseIf Stick(i).X > 41000 Then
'                Stick(i).X = 41000
'                Stick(i).Speed = 0
'            End If
            
            If Stick(i).Y > StickGameHeight - CLD2 Then
                Stick(i).Y = StickGameHeight - CLD2
                Stick(i).Speed = 0
            End If
            
            
            
            
            'chopper rockets
            
            
            If StickiHasState(i, STICK_NADE) Then
                If Stick(i).LastNade + Chopper_RPG_Delay / GetSticksTimeZone(i) < GetTickCount() Then
                    
                    AddNade Stick(i).GunPoint.X, Stick(i).GunPoint.Y, Stick(i).ActualFacing, RPG_Speed, _
                        i, Stick(i).colour, nFrag, True
                    
                    
                    Stick(i).LastNade = GetTickCount()
                     
                ElseIf Stick(i).LastNade + Nade_Release_Delay < GetTickCount() Then
                    SubStickiState i, STICK_NADE
                End If
            End If
            
            
        End If
        
        
        'health pack
        Call CheckStickHealthPack(i)
        
        
        'STICKMOTION MOVED TO TOP
        
        'Wrap edges + speed
        ClipStick i, bLBound
        
    End If 'stickingame endif
    
Next i


Call CheckChopperCollisions


ProcessBullets


EH:
End Sub

Private Sub ProcessBullets()
Dim i As Integer, j As Integer
Dim lastPoint As ptPoint

'Loop through each bullet and perform physics
i = 0
Do While i < NumBullets
    
    With Bullet(i)
        lastPoint.X = .X
        lastPoint.Y = .Y
        
        MotionStickObject .X, .Y, .Speed, .Heading
        
        ApplyGravityVector .LastGravity, GetTimeZoneAdjust(.X, .Y), .Speed, .Heading, .X, .Y, Bullet_Gravity_Strength
    End With
    
    
    'Wrap edges
    If ClipBullet(i) = False Then
        
        'If FindStick(Bullet(i).OwnerIndex) <> -1 Then
        'Check for collisions
            For j = 0 To NumSticksM1
                If StickInGame(j) Then
                    If j <> Bullet(i).OwnerIndex Then
                        If Not IsAlly(Stick(j).Team, Stick(Bullet(i).OwnerIndex).Team) Then
                            'If BulletInStick(j, i) Then 'GetDist(Stick(j).X, Stick(j).Y, Bullet(i).X, Bullet(i).Y) < (Bullet_Radius + BodyLen / 2) Then
                            
                            On Error GoTo Bullet_EH
                            If BulletPassedThroughStick(j, i, lastPoint) Then
                                
                                If BulletHitStick(i, j, CInt(Bullet(i).OwnerIndex)) Then
                                    'only exit, if bullet removed, otherwise, check the other sticks
                                    Exit For
                                End If
                                
                            End If 'bulletinstick endif
Bullet_EH:
                            
                        End If 'ally endif
                    End If 'bulletID endif
                End If 'stickingame endif
            Next j
        'Else
            'RemoveBullet i, False 'owner left the game
            'i = i - 1
        'End If 'ownerindex <> -1 endif
    Else
        i = i - 1
    End If 'clip endif
    
    'Increment counter
    i = i + 1
Loop

End Sub

Private Function BulletPassedThroughStick(iStick As Integer, iBullet As Integer, BulletStart As ptPoint) As Boolean
Dim rcStick As RECT, rcBulletPath As RECT
Dim kArmLen As Single, sngSwap As Single
Const ArmLenExtended = ArmLen * 1.2, ArmLenX3 = ArmLen * 3
Const HeadRadiusX2 = HeadRadius * 2, BodyLenX2 = BodyLen * 2

If Stick(iStick).WeaponType = Chopper Then
    BulletPassedThroughStick = CoOrdInChopper(Bullet(iBullet).X, Bullet(iBullet).Y, iStick)
Else
    
    If StickiHasState(iStick, STICK_PRONE) Then
        kArmLen = ArmLenX3 'body/leg hit
    Else
        kArmLen = ArmLenExtended
    End If
    
    
    With rcStick
        .Left = Stick(iStick).X - kArmLen
        .Right = Stick(iStick).X + kArmLen
        .Top = GetStickY(iStick)
        .Bottom = .Top + IIf(StickiHasState(iStick, STICK_PRONE), HeadRadiusX2, BodyLenX2)
    End With
    With rcBulletPath
        .Left = BulletStart.X
        .Top = BulletStart.Y
        .Right = Bullet(iBullet).X
        .Bottom = Bullet(iBullet).Y
        
        If .Left > .Right Then
            sngSwap = .Left
            .Left = .Right
            .Right = sngSwap
        End If
        If .Top > .Bottom Then
            sngSwap = .Top
            .Top = .Bottom
            .Bottom = sngSwap
        End If
        
        'ensure there's a rect collision, if possible
        If .Top = .Bottom Then
            .Bottom = .Bottom + 1
        End If
        If .Left = .Right Then
            .Right = .Right + 1
        End If
    End With
    
    
    
    If RectCollision(rcStick, rcBulletPath) Then
        'determine if the bullet hit
        
        BulletPassedThroughStick = SimultaneousSolvable(iStick, rcStick, iBullet, BulletStart)
        
    End If
End If


End Function

Private Function SimultaneousSolvable(iStick As Integer, rcStick As RECT, iBullet As Integer, BulletStart As ptPoint) As Boolean
Dim sngY As Single, dx As Single, Bullet_Y_Intercept As Single, Theta As Single
Dim Bullet_Path As ptPath

With Bullet_Path
    .XStart = BulletStart.X
    .YStart = BulletStart.Y
    .XEnd = Bullet(iBullet).X
    .YEnd = Bullet(iBullet).Y
    
    dx = .XStart - .XEnd
    
    
    If dx <> 0 Then
        Theta = FixAngle(Bullet(iBullet).Heading - piD2)
        
        If Theta <> 0 And Theta <> Pi Then
            '      y = mx + c
            '       c          =       y           -     m          *        x
            Bullet_Y_Intercept = Bullet(iBullet).Y - Tangent(Theta) * Bullet(iBullet).X
            
            
            ' y  =            x    *                 m      +        c
            sngY = Stick(iStick).X * (.YStart - .YEnd) / dx + Bullet_Y_Intercept
            'sngY = the y-co-ord of where the bullet crossed the stick's vertical line
            
            
            '<Stan> Not entierly suer [sure] on this... It works though </Stan>
            If rcStick.Top <= sngY Then
                SimultaneousSolvable = True
            ElseIf sngY <= rcStick.Bottom Then
                SimultaneousSolvable = True
            End If
            
            Exit Function
        'Else
            'dealt with below
        End If
    'Else
        'dealt with below
    End If
    
    
    'bullet is heading up or down - compare y values
    Select Case Bullet(iBullet).Heading 'Round(Bullet(iBullet).Heading * 180 / Pi)
        Case 0
            'heading up
            If .XStart > rcStick.Top Then
                SimultaneousSolvable = (.XEnd < rcStick.Bottom)
            End If
            
        Case Else '180
            'heading down
            If .XStart < rcStick.Bottom Then
                 SimultaneousSolvable = (.XEnd > rcStick.Top)
            End If
            
        'Case Else
            'Debug.Print "Uh oh " & Rnd() & " - [" & Bullet(iBullet).Heading / Pi * 180 & "]"
            
    End Select
End With

End Function

Private Function StickPlatCollision(iPlat As Integer, rcPlat As RECT, iStick As Integer, stickstart As ptPoint) As Boolean
Dim sngY As Single, dx As Single, Bullet_Y_Intercept As Single, Theta As Single
Dim Bullet_Path As ptPath

With Bullet_Path
    .XStart = stickstart.X
    .YStart = stickstart.Y
    .XEnd = Stick(iStick).X
    .YEnd = Stick(iStick).Y
    
    dx = .XStart - .XEnd
    
    
    If dx <> 0 Then
        Theta = -piD2
        Bullet_Y_Intercept = Stick(iStick).Y - Tangent(Theta) * Stick(iStick).X
        
        
        ' y  =            x    *                 m      +        c
        sngY = rcPlat.Left * (.YStart - .YEnd) / dx + Bullet_Y_Intercept
        'sngY = the y-co-ord of where the bullet crossed the stick's vertical line
        
        
        '<Stan> Not entierly suer [sure] on this... It works though </Stan>
        If rcPlat.Top <= sngY Then
            StickPlatCollision = True
        ElseIf sngY <= rcPlat.Bottom Then
            StickPlatCollision = True
        End If
    End If
End With
End Function

'Private Function BulletInStick(Sticki As Integer, Bulleti As Integer) As Boolean
'Dim kArmLen As Single
'Const ArmLenExtended = ArmLen * 1.2
'Const HeadRadiusX2 = HeadRadius * 2, BodyLenX2 = BodyLen * 2
'Dim rcStick As RECT
'
'If Stick(Sticki).WeaponType = Chopper Then
'
'    BulletInStick = CoOrdInChopper(Bullet(Bulleti).X, Bullet(Bulleti).Y, Sticki)
'
''    If Bullet(Bulleti).X > (Stick(Sticki).X - CLDx) Then
''        If Bullet(Bulleti).X < (Stick(Sticki).X + CLD4) Then
''            If Bullet(Bulleti).Y > Stick(Sticki).Y - CLD8 Then
''                If Bullet(Bulleti).Y < Stick(Sticki).Y + CLD6 Then
''                    BulletInStick = True
''                End If
''            End If
''
''        End If
''    End If
'
'Else
'
'    If StickiHasState(Sticki, Stick_Prone) Then
'        kArmLen = ArmLen * 3 'body/leg hit
'    Else
'        kArmLen = ArmLenExtended
'    End If
'
'
'    With rcStick
'        .Left = Stick(Sticki).X - kArmLen
'        .Right = Stick(Sticki).X + kArmLen
'        .Top = GetStickY(Sticki)
'        .Bottom = .Top + IIf(StickiHasState(Sticki, Stick_Prone), HeadRadiusX2, BodyLenX2)
'    End With
'
'    BulletInStick = RectCollision(rcStick, PointToRect(Bullet(Bulleti).X, Bullet(Bulleti).Y))
'
'
''    If Abs(Bullet(Bulleti).X - Stick(Sticki).X) < kArmLen Then
''        sY = GetStickY(Sticki)
''
''        If Bullet(Bulleti).Y > sY Then '(Stick(Sticki).y) Then
''            BulletInStick = (Bullet(Bulleti).Y < sY + IIf(StickiHasState(Sticki, Stick_Prone), HeadRadiusX2, BodyLenX2))
''        End If
''    End If
'
'
'End If
'
'End Function

Private Function GetMaxSpeed(i As Integer) As Single

If Stick(i).WeaponType = Chopper Then
    GetMaxSpeed = Chopper_Max_Speed
Else
    If Stick(i).bOnSurface Then
        If StickiHasState(i, STICK_CROUCH) Then
            GetMaxSpeed = Max_Speed / 4
        ElseIf StickiHasState(i, STICK_PRONE) Then
            GetMaxSpeed = Max_Speed / 8
        Else
            If Stick(i).Perk = pConditioning Then
                GetMaxSpeed = Max_Speed * ConditioningMaxSpeedInc
            Else
                GetMaxSpeed = Max_Speed
            End If
        End If
    Else
        If Stick(i).Perk = pConditioning Then
            GetMaxSpeed = Max_Speed * ConditioningMaxSpeedInc
        Else
            GetMaxSpeed = Max_Speed
        End If
    End If
    
    If Stick(i).Perk = pSniper Then GetMaxSpeed = GetMaxSpeed / Sniper_Max_Speed_Dec
    If Stick(i).Perk = pZombie Then GetMaxSpeed = Max_Speed * ZombieMaxSpeedDec
End If

End Function

Private Function BulletHitStick(i As Integer, j As Integer, BulletOwnerIndex As Integer) As Boolean
Dim f As Single, BHeading As Single
Dim bHeadShot As Boolean
Dim kType As eKillTypes
Dim bCan As Boolean
Const Pi2Mabit As Single = Pi2 - 0.01

bCan = True

'if saber, deflect, else damage
If Stick(j).bLightSaber And Stick(j).WeaponType = Knife Then
    f = FixAngle(Stick(j).Facing)
    If f = 0 Or f > Pi2Mabit Then
        bCan = False
    End If
    'BHeading = FixAngle(Bullet(i).Heading)
    
    
    
'    If BHeading > Pi Then
'        If f = 0 Or f = Pi Then
'            'deflect!
'            bCan = False
'        End If
'    Else
'        If f > Pi And f < pi5D4 Or f > pi7D4 And f < Pi2 Then
'            'deflect!
'            bCan = False
'        End If
'    End If
    
End If


If bCan Then
    With Bullet(i)
        If Stick(j).Shield > 0 Then
            AddShieldWave .X, .Y, .Heading - Pi
        ElseIf Stick(j).WeaponType <> Chopper Then
            AddBlood .X, .Y, .Heading 'add blood before removing t'bullet
            If StickiHasState(j, STICK_PRONE) = False Then
                AddVectors Stick(j).Speed, Stick(j).Heading, .Damage * .Speed / 100, .Heading, Stick(j).Speed, Stick(j).Heading
            End If
        Else
            'chopper
            AddBulletExplosion Bullet(i).X, Bullet(i).Y
            AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading - Pi
        End If
    End With
Else
    'saber...
    AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading - Pi
End If


If (0 = j Or Stick(j).IsBot) And bCan Then 'And modStickGame.StickServer) Then
  '^FindStick(MyID)
  'If this is our Stick...
    
    'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
    If StickInvul(j) = False Then
        
        'If Bullet(i).bShotgunBullet Then
            'AlterW1200BulletDamage i
        'End If
        
        
        If Stick(j).WeaponType <> Chopper Then
            bHeadShot = BulletInHead(j, i)
            
            If bHeadShot Then
                DamageStick Bullet(i).Damage * HeadShot_Damage_Factor, j, BulletOwnerIndex
            Else
                DamageStick Bullet(i).Damage, j, BulletOwnerIndex
            End If
        ElseIf Bullet(i).bSniperBullet Then
            'sniper hitting a chopper
            DamageStick 33 * Chopper_Damage_Reduction / modStickGame.sv_Damage_Factor, j, BulletOwnerIndex
        Else
            'normal bullet hitting chopper
            DamageStick Bullet(i).Damage / modStickGame.sv_Damage_Factor, j, BulletOwnerIndex
        End If
        
        If bHeadShot Then
            kType = kHead
            
            With Bullet(i)
                If Stick(j).Shield Then
                    AddShieldWave .X, .Y, .Heading - Pi
                Else
                    AddBlood .X, .Y, .Heading + piD10
                    AddBlood .X, .Y, .Heading
                    AddBlood .X, .Y, .Heading - piD10
                    
                    AddBlood .X, .Y, .Heading - Pi
                End If
            End With
            
        ElseIf BulletOwnerIndex <> -1 Then
            If Stick(BulletOwnerIndex).bSilenced Then
                kType = kSilenced
            Else
                kType = kNormal
            End If
        Else
            kType = kNormal
        End If
        
'        If kType = kHead Then
'             will never get to this point
'            If Stick(BulletOwnerIndex).WeaponType = Chopper Then
'                kType = kNormal
'            End If
'        End If
        
        If Stick(j).Health < 1 Then
            Call Killed(j, BulletOwnerIndex, kType)
        End If
        
    End If 'spawn invul endif
End If 'stick(j) = me endif


If Stick(j).WeaponType = Chopper Or Not Bullet(i).bSniperBullet Then
    RemoveBullet i, False, bCan
    BulletHitStick = True
End If

End Function

Private Sub FireShot(i As Integer, BaseTempDir As Single, TempMag As Single, _
    iDirection As Integer, tX As Single, tY As Single, Stick_Moving As Boolean, bLBound As Boolean)

Const AccuracyRedux As Single = 2000 'bigger, the more accurate
Const Flame_Sound_Delay As Long = 300
Const Crouch_Aim_Reduction As Single = 6, Prone_Aim_Reduction As Single = Crouch_Aim_Reduction * 2
Const GunLenDx = GunLen / 3


Dim bDrawWeaponFlash As Boolean
Dim vPoint As PointAPI
Dim AimRedux As Single
Dim j As Integer

'#################################
If Stick(i).WeaponType < Chopper Then
#If Hack_Recoil = False Then
    If i = 0 Then
        If WeaponIsSniper(Stick(i).WeaponType) = False Then
            If Stick(i).WeaponType <> Mac10 Then
                GetCursorPos vPoint
                
                AimRedux = 10 * (kRecoilAmount(Stick(i).WeaponType) + 1)
                
                If Stick(i).bSilenced Then AimRedux = AimRedux / 3
                If Stick(i).Perk = pSteadyAim Then AimRedux = AimRedux / SteadyAim_Reduction
                If StickiHasState(i, STICK_PRONE) Then
                    AimRedux = AimRedux / Prone_Aim_Reduction
                ElseIf StickiHasState(i, STICK_CROUCH) Then
                    AimRedux = AimRedux / Crouch_Aim_Reduction
                End If
                
                
                vPoint.Y = vPoint.Y - AimRedux
                
                
                SetCursorPos vPoint.X, vPoint.Y
            End If
        End If
    End If
#End If
    
    'VISIBLE RECOIL IS DONE HERE
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * kRecoilAmount(Stick(i).WeaponType)
End If
'#################################


BaseTempDir = BaseTempDir + PM_Rnd * Stick(i).Speed / AccuracyRedux




If Stick(i).WeaponType = FlameThrower Then
    AddFlame Stick(i).GunPoint.X, Stick(i).GunPoint.Y, _
        BaseTempDir + IIf(Stick(i).Facing > Pi, 0.05, -0.05), TempMag, Stick(i).ID, i
    
    Stick(i).BulletsFired = Stick(i).BulletsFired + 1
    
    
ElseIf Stick(i).WeaponType = RPG Then
    
    If ForeignStick(i) Then
        AddNade tX, tY + GunLenDx * Tangent(kRecoilAmount(Stick(i).WeaponType)), Stick(i).ActualFacing, RPG_Speed, i, Stick(i).colour, nFrag, True
        'don't alter the value of tY
    Else
        AddNade tX, tY, Stick(i).ActualFacing, RPG_Speed, i, Stick(i).colour, nFrag, True
    End If
    
    
    'rear point
    tX = tX - GunLen * 4 * Sine(Stick(i).ActualFacing)
    tY = tY + GunLen * 4 * CoSine(Stick(i).ActualFacing)
    
    AddExplosion tX, tY, 500
    AddExplosion tX, tY, 500
    
    
    For j = 1 To 10
        AddSmokeGroup tX, tY, 4, 20 + 30 * Rnd(), Stick(i).ActualFacing - Pi
    Next j
    
    'recoil-force
    AddVectors Stick(i).Speed, Stick(i).Heading, _
               RPG_RecoilForce, FixAngle(Stick(i).Facing - Pi), _
               Stick(i).Speed, Stick(i).Heading
    
    
    Stick_Moving = True
    bLBound = False
    
    If i = 0 Then
        'keep the firekey held
        FireKeyUpTime = GetTickCount()
    End If
    
Else
    FireWeapon i, tX, tY, BaseTempDir, TempMag, iDirection, Stick_Moving, bLBound
End If




If Stick(i).WeaponType <> FlameThrower Then
    If Stick(i).bSilenced Then
        If Stick(i).WeaponType <> Mac10 Then
            If i = 0 Then
                modAudio.PlaySilencedSound 0, kSilencedOffset(Stick(i).WeaponType)
            ElseIf PointHearableOnSticksScreen(tX, tY, 0) Then
                modAudio.PlaySilencedSound GetRelPan(tX), kSilencedOffset(Stick(i).WeaponType)
            End If
        End If
    Else
        
        If i = 0 Then
            modAudio.PlayWeaponSound_Panned Stick(i).WeaponType, 0
            
        ElseIf PointHearableOnSticksScreen(tX, tY, 0) Then
            modAudio.PlayWeaponSound_Panned Stick(i).WeaponType, GetRelPan(tX)
        End If
        
        
    End If
ElseIf Stick(i).LastBullet + Flame_Sound_Delay < GetTickCount() Then
    If i = 0 Then
        modAudio.PlayWeaponSound_Panned Stick(i).WeaponType, 0
    ElseIf PointHearableOnSticksScreen(Stick(i).X, Stick(i).Y, 0) Then
        modAudio.PlayWeaponSound_Panned Stick(i).WeaponType, GetRelPan(Stick(i).X)
    End If
End If

End Sub

Private Sub FireWeapon(i As Integer, MuzzleX As Single, MuzzleY As Single, _
                        Bullet_Dir As Single, Bullet_Mag As Single, _
                        iRecoil_Dir As Integer, _
                        ByRef Stick_Moving As Boolean, ByRef bLBound As Boolean)

Dim j As Integer
Dim sngSprayAngle As Single, Shot_Gauge As Integer
Const Crouch_Recoil_Reduction As Long = 4, GunLenDx = GunLen / 3

If kRecoilForce(Stick(i).WeaponType) Then
    
    If StickiHasState(i, STICK_PRONE) = False Then
        Stick_Moving = True
        bLBound = False
        
        
        AddVectors Stick(i).Speed, Stick(i).Heading, _
            kRecoilForce(Stick(i).WeaponType) / IIf(StickiHasState(i, STICK_CROUCH), Crouch_Recoil_Reduction, 1), FixAngle(Stick(i).Facing - Pi), _
            Stick(i).Speed, Stick(i).Heading
        
    End If
    
End If


'adjust for lag - reset Stick(i).GunPoint.Y, so the bullet is added at the same point
If ForeignStick(i) Then
    MuzzleY = MuzzleY + GunLenDx * Tangent(kRecoilAmount(Stick(i).WeaponType))
End If


If WeaponIsShotgun(Stick(i).WeaponType) Then
    
    '################################################################################
    
    
    'adjust since bullets will be created from stick.actualfacing + ...
    
    If Stick(i).WeaponType = SPAS Then
        Shot_Gauge = SPAS_Gauge
        sngSprayAngle = SPAS_Spray_Angle / Shot_Gauge
        
        If StickiHasState(i, STICK_CROUCH) Then
            Bullet_Dir = Bullet_Dir - Sine(Stick(i).ActualFacing) / 32
        End If
        
        Stick(i).LastMuzzleFlash = GetTickCount()
        
    Else
        Shot_Gauge = W1200_Gauge
        sngSprayAngle = W1200_Spray_Angle / Shot_Gauge
        
        Bullet_Dir = Bullet_Dir - Sine(Stick(i).ActualFacing) / _
                           IIf(StickiHasState(i, STICK_CROUCH), 10, 25)
        
        
        AddExplosion Stick(i).GunPoint.X, Stick(i).GunPoint.Y, 100
    End If
    
    
    
    For j = 1 To Shot_Gauge
        Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * sngSprayAngle
        
        
        AddBullet MuzzleX, MuzzleY, Bullet_Mag + Rnd() * 100, Bullet_Dir, kBulletDamage(Stick(i).WeaponType), i
    Next j
    
    AddCasing Stick(i).CasingPoint.X, Stick(i).CasingPoint.Y, Bullet_Dir, False, i
    '################################################################################
    
Else
    If kSprayAngle(Stick(i).WeaponType) Then
        Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * kSprayAngle(Stick(i).WeaponType)
    Else
        Bullet_Dir = Stick(i).ActualFacing
    End If
    
    AddBullet MuzzleX, MuzzleY, kBulletSpeed(Stick(i).WeaponType), Bullet_Dir, kBulletDamage(Stick(i).WeaponType), i
    
    If Stick(i).bSilenced = False Then
        Stick(i).LastMuzzleFlash = GetTickCount()
    End If
End If



'###############################################################
'###############################################################
'###############################################################

'If Stick(i).WeaponType = Chopper Then
'
'    AddBullet Stick(i).GunPoint.X, Stick(i).GunPoint.Y, BULLET_SPEED, _
'        Stick(i).ActualFacing + PM_Rnd() * Chopper_Spray_Angle, Chopper_Bullet_Damage, i
'
'ElseIf Stick(i).WeaponType = g3 Then
'
'    Bullet_Dir = Stick(i).ActualFacing
'
'    AddBullet MuzzleX, MuzzleY, G3_Speed, Bullet_Dir, G3_Bullet_Damage, i
'
'ElseIf Stick(i).WeaponType = SPAS Then
'
'    If StickiHasState(i, STICK_CROUCH) Then
'        Bullet_Dir = Bullet_Dir + Sine(Stick(i).Facing) / 10
'    End If
'
'    'adjust since bullets aren't sent for stick(i).pointapi
'    Bullet_Dir = Bullet_Dir - Sine(Stick(i).ActualFacing) / IIf(StickiHasState(i, STICK_CROUCH), 8, 25)
'
'    For j = 1 To SPAS_Gauge
'        Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * SPAS_Spray_Angle
'        AddBullet MuzzleX, MuzzleY, Bullet_Mag + Rnd() * 100, Bullet_Dir, Stick(i).ID, SPAS_Bullet_Damage, i
'        'AddBullet MuzzleX, MuzzleY, TempMag, Bullet_Dir, Stick(i).ID, W1200_Bullet_Damage, i
'    Next j
'
'    AddCasing Stick(i).CasingPoint.X, Stick(i).CasingPoint.Y, Bullet_Dir, False, i
'
'
'    'recoil
'    AddVectors Stick(i).Speed, Stick(i).Heading, SPAS_RecoilForce, FixAngle(Stick(i).Facing - Pi), Stick(i).Speed, Stick(i).Heading
'
'ElseIf Stick(i).WeaponType = Mac10 Then
'
'
'    Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * Mac10_Spray_Angle
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, Mac10_Bullet_Damage, i
'
'ElseIf Stick(i).WeaponType = W1200 Then
'
'
'    'adjust since bullets aren't sent for stick(i).pointapi
'    Bullet_Dir = Bullet_Dir - Sine(Stick(i).ActualFacing) / IIf(StickiHasState(i, STICK_CROUCH), 8, 25)
'
'    For j = 1 To W1200_Gauge
'        Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * W1200_Spray_Angle
'        AddBullet MuzzleX, MuzzleY, Bullet_Mag + Rnd() * 100, Bullet_Dir, Stick(i).ID, W1200_Bullet_Damage, i
'    Next j
'
'    AddCasing Stick(i).CasingPoint.X, Stick(i).CasingPoint.Y, Bullet_Dir, False, i
'
'
'    'recoil
'    AddVectors Stick(i).Speed, Stick(i).Heading, W1200_RecoilForce, FixAngle(Stick(i).Facing - Pi), Stick(i).Speed, Stick(i).Heading
'
'    AddExplosion MuzzleX, MuzzleY, 300, 0.1, 30, Bullet_Dir
'
'ElseIf Stick(i).WeaponType = AK Then
'
'
'    Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * AK_Spray_Angle
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, AK_Bullet_Damage, i
'
'
'ElseIf Stick(i).WeaponType = M249 Then
'
'
'
'    Bullet_Dir = Bullet_Dir - Sine(Stick(i).Facing) / 25 + iRecoil_Dir * Rnd() * M249_Spray_Angle
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, M249_Bullet_Damage, i
'
'    AddVectors Stick(i).Speed, Stick(i).Heading, M249_RecoilForce, FixAngle(Stick(i).Facing - Pi), Stick(i).Speed, Stick(i).Heading
'
'ElseIf Stick(i).WeaponType = MP5 Then
'
'
'    Bullet_Dir = Bullet_Dir + iRecoil_Dir * Rnd() * MP5_Spray_Angle
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, MP5_Bullet_Damage, i
'
'
'ElseIf Stick(i).WeaponType = DEagle Then
'
'
'    If StickiHasState(i, STICK_CROUCH) Then
'        Bullet_Dir = Bullet_Dir - Sine(Stick(i).Facing) / 200
'    Else
'        Bullet_Dir = Bullet_Dir + Sine(Stick(i).Facing) / 50
'    End If
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, DEagle_Bullet_Damage, i
'    AddSmokeGroup MuzzleX + Bullet_Mag * Sine(Bullet_Dir), MuzzleY, 4, Rnd() * Bullet_Mag / 10, Bullet_Dir
'
'
'    'recoil
'    AddVectors Stick(i).Speed, Stick(i).Heading, DEagle_RecoilForce, FixAngle(Stick(i).Facing - Pi), Stick(i).Speed, Stick(i).Heading
'
'ElseIf Stick(i).WeaponType = USP Then
'
'
'    If StickiHasState(i, STICK_CROUCH) Then
'        Bullet_Dir = Bullet_Dir - Sine(Stick(i).Facing) / 200
'    Else
'        Bullet_Dir = Bullet_Dir + Sine(Stick(i).Facing) / 50
'    End If
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, USP_Bullet_Damage, i
'    AddSmokeGroup MuzzleX + Bullet_Mag * Sine(Bullet_Dir), MuzzleY, 4, Rnd() * Bullet_Mag / 10, Bullet_Dir
'
'
'ElseIf Stick(i).WeaponType = M82 Then
'
'
'    'due to recoil, by the time others find out i'm shooting,
'    'i've moved back from where i was, so if it's a foreign stick, adjust the angle more
'    'screw ^^
'    'deadly accurate ftw
'
'    Bullet_Dir = Stick(i).ActualFacing
'
'
'    AddBullet MuzzleX, MuzzleY, M82_Speed, Bullet_Dir, Stick(i).ID, M82_Bullet_Damage, i
'
'    AddVectors Stick(i).Speed, Stick(i).Heading, M82_RecoilForce / IIf(Stick(i).bSilenced, M82_Silent_Recoil_Reduction, 1), FixAngle(Stick(i).Facing - Pi), Stick(i).Speed, Stick(i).Heading
'
'ElseIf Stick(i).WeaponType = AWM Then
'
'
'    Bullet_Dir = Stick(i).ActualFacing
'
'    AddBullet MuzzleX, MuzzleY, AWM_Speed, Bullet_Dir, Stick(i).ID, AWM_Bullet_Damage, i
'
'
'    AddVectors Stick(i).Speed, Stick(i).Heading, _
'        AWM_RecoilForce, _
'        FixAngle(Stick(i).Facing - Pi), Stick(i).Speed, Stick(i).Heading
'
'
'
'ElseIf Stick(i).WeaponType = XM8 Then
'
'    Bullet_Dir = Bullet_Dir + Sine(Stick(i).Facing) / IIf(StickiHasState(i, STICK_CROUCH), -60, 50) - iRecoil_Dir * Rnd() * XM8_Spray_Angle
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, XM8_Bullet_Damage, i
'
'ElseIf Stick(i).WeaponType = AUG Then
'
'    If StickiHasState(i, STICK_CROUCH) Then
'        Bullet_Dir = Bullet_Dir - Sine(Stick(i).Facing) / 50
'    Else
'        Bullet_Dir = Bullet_Dir '- Sine(Stick(i).Facing) / 50
'    End If
'
'
'    AddBullet MuzzleX, MuzzleY, Bullet_Mag, Bullet_Dir, Stick(i).ID, AUG_Bullet_Damage, i
'End If

End Sub


'Private Sub AlterW1200BulletDamage(i As Integer)
''time left = B(i).Decay - GTC()
'
'Bullet(i).Damage = Bullet(i).Damage * modStickGame.StickTimeFactor * (Bullet(i).Decay - GetTickCount()) / Bullet_Decay
'
'End Sub

Private Sub ProcessBlood()
Dim i As Integer

i = 0
Do While i < NumBlood
    'Is this one decayed?
    If Blood(i).Decay < GetTickCount() Then
        'Kill it!
        RemoveBlood i
        'Decrement the counter
        i = i - 1
    End If
    'Increment the counter
    i = i + 1
Loop


i = 0
Do While i < NumBlood
    MotionStickObject Blood(i).X, Blood(i).Y, Blood(i).Speed, Blood(i).Heading
    i = i + 1
Loop

End Sub

Private Sub DoRecoil(i As Integer, ByRef Stick_Moving As Boolean, ByRef bLBound As Boolean)
Dim Adj As Single

Adj = GetSticksTimeZone(i) '* modStickGame.sv_StickGameSpeed


If Stick(i).LastBullet + kRecoilTime(Stick(i).WeaponType) / Adj > GetTickCount() Then
    
    Stick(i).Facing = Stick(i).Facing + _
        IIf(Stick(i).RecoilLeft, -1, 1) * kRecoverAmount(Stick(i).WeaponType) * Adj
    
    
    If kRecoilForce(Stick(i).WeaponType) Then
        
        If Stick(i).WeaponType <> RPG Then
            Stick_Moving = True
            bLBound = False
        ElseIf Stick(i).LastBullet + RPG_Recoil_Time / (3 * Adj) > GetTickCount() Then
            Stick_Moving = True
            bLBound = False
        End If
        
    End If
End If


'If Stick(i).WeaponType = W1200 Then
'    If Stick(i).LastBullet + W1200_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * W1200_Recover_Amount * GetTimeZoneAdjust
'
'        'allow recoil to push back
'        Stick_Moving = True
'        bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = AK Then
'    If Stick(i).LastBullet + AK_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * AK_Recover_Amount * GetTimeZoneAdjust
'    End If
'ElseIf Stick(i).WeaponType = DEagle Then
'    If Stick(i).LastBullet + DEagle_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * DEagle_Recover_Amount * GetTimeZoneAdjust
'
'        'allow recoil to push back
'        Stick_Moving = True
'        bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = M82 Then
'    If Stick(i).LastBullet + M82_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'
'
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * M82_Recover_Amount * GetTimeZoneAdjust
'
'        'allow recoil to push back
'        Stick_Moving = True
'        bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = XM8 Then
'    If Stick(i).LastBullet + XM8_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * XM8_Recover_Amount * GetTimeZoneAdjust
'    End If
'ElseIf Stick(i).WeaponType = M249 Then
'    If Stick(i).LastBullet + M249_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * M249_Recover_Amount * GetTimeZoneAdjust
'
'        ''allow recoil to push back
'        'Stick_Moving = True
'        'bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = RPG Then
'    If Stick(i).LastBullet + RPG_Recoil_Time / GetTimeZoneAdjust > GetTickCount() Then
'
'        If i = 0 Then 'prevent wobble
'            Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * RPG_Recover_Amount * GetTimeZoneAdjust
'        End If
'
'
'        If Stick(i).LastBullet + RPG_Recoil_Time / (3 * GetTimeZoneAdjust) > GetTickCount() Then
'            'allow recoil to push back
'            Stick_Moving = True
'            bLBound = False
'        End If
'
'    End If
'End If
End Sub

Private Sub InitWeaponStats()
Dim i As Integer

For i = 0 To eWeaponTypes.Knife
    If i = W1200 Then
        kRecoilTime(i) = W1200_Recoil_Time
        kRecoverAmount(i) = W1200_Recover_Amount
        kRecoilForce(i) = W1200_RecoilForce 'True
        kRecoilAmount(i) = W1200_SingleRecoil_Angle
        
    ElseIf i = SPAS Then
        kRecoilTime(i) = SPAS_Recoil_Time
        kRecoverAmount(i) = SPAS_Recover_Amount
        kRecoilForce(i) = SPAS_RecoilForce
        kRecoilAmount(i) = SPAS_SingleRecoil_Angle
        'kSilencable(i) = True
        
    ElseIf i = AK Then
        kRecoilTime(i) = AK_Recoil_Time
        kRecoverAmount(i) = AK_Recover_Amount
        kRecoilAmount(i) = AK_SingleRecoil_Angle
        kSilencable(i) = True
        
    ElseIf i = DEagle Then
        kRecoilTime(i) = DEagle_Recoil_Time
        kRecoverAmount(i) = DEagle_Recover_Amount
        kRecoilForce(i) = DEagle_RecoilForce
        kRecoilAmount(i) = DEagle_SingleRecoil_Angle
        
    ElseIf i = M82 Then
        kRecoilTime(i) = M82_Recoil_Time
        kRecoverAmount(i) = M82_Recover_Amount
        kRecoilForce(i) = M82_RecoilForce
        kRecoilAmount(i) = M82_SingleRecoil_Angle
        
    ElseIf i = AWM Then
        kRecoilTime(i) = AWM_Recoil_Time
        kRecoverAmount(i) = AWM_Recover_Amount
        kRecoilForce(i) = AWM_RecoilForce
        kRecoilAmount(i) = AWM_SingleRecoil_Angle
        kSilencable(i) = True
        kSilencedOffset(i) = 1
        
    ElseIf i = XM8 Then
        kRecoilTime(i) = XM8_Recoil_Time
        kRecoverAmount(i) = XM8_Recover_Amount
        kRecoilAmount(i) = XM8_SingleRecoil_Angle
        kSilencable(i) = True
        kSilencedOffset(i) = 1
        
    ElseIf i = M249 Then
        kRecoilTime(i) = M249_Recoil_Time
        kRecoverAmount(i) = M249_Recover_Amount
        kRecoilForce(i) = M249_RecoilForce
        kRecoilAmount(i) = M249_SingleRecoil_Angle
        
    ElseIf i = RPG Then
        kRecoilTime(i) = RPG_Recoil_Time
        kRecoverAmount(i) = RPG_Recover_Amount
        kRecoilForce(i) = RPG_RecoilForce
        kRecoilAmount(i) = RPG_SingleRecoil_Angle
        
    ElseIf i = AUG Then
        kRecoilTime(i) = AUG_Recoil_Time
        kRecoverAmount(i) = AUG_Recover_Amount
        kRecoilAmount(i) = AUG_SingleRecoil_Angle
        kSilencable(i) = True
        kBurstBullets(i) = AUG_Burst_Bullets
        kBurstDelay(i) = AUG_Bullet_Delay
        kSilencedOffset(i) = 1
        
    ElseIf i = USP Then
        kRecoilTime(i) = USP_Recoil_Time
        kRecoverAmount(i) = USP_Recover_Amount
        kRecoilAmount(i) = USP_SingleRecoil_Angle
        kSilencable(i) = True
        
        'kBurstBullets(i) = USP_Burst_Bullets
        'kBurstDelay(i) = USP_Bullet_Delay
        
    ElseIf i = MP5 Then
        kRecoilTime(i) = MP5_Recoil_Time
        kRecoverAmount(i) = MP5_Recover_Amount
        kRecoilAmount(i) = MP5_SingleRecoil_Angle
        kSilencable(i) = True
        
    ElseIf i = Mac10 Then
        kRecoilTime(i) = Mac10_Recoil_Time
        kRecoverAmount(i) = Mac10_Recover_Amount
        kRecoilAmount(i) = Mac10_SingleRecoil_Angle
        kSilencable(i) = True
        
    ElseIf i = G3 Then
        kRecoilTime(i) = G3_Recoil_Time
        kRecoverAmount(i) = G3_Recover_Amount
        kRecoilAmount(i) = G3_SingleRecoil_Angle
        kSilencable(i) = True
        kBurstBullets(i) = G3_Burst_Bullets
        kBurstDelay(i) = G3_Bullet_Delay
        kSilencedOffset(i) = 1
        
    End If
Next i

End Sub

Private Function StickInvul(i As Integer) As Boolean
StickInvul = (Stick(i).LastSpawnTime + Spawn_Invul_Time > GetTickCount())
End Function

Private Sub RemoveSticksShield(i As Integer)
If Stick(i).Shield Then
    Stick(i).Shield = 0
    ShowShieldsDown i
End If
End Sub

Private Sub ShowShieldsDown(i As Integer)
Const shieldSparksToAdd As Integer = 60
With Stick(i)
    AddMoreSparks .X, .Y, shieldSparksToAdd
    addShieldExhaustWave i, 50, vbRed
End With
End Sub

Private Sub DamageStick(ByVal DamageToDo As Integer, iStick As Integer, _
    iDamager As Integer, Optional bDamageShield As Boolean = True, _
    Optional bApplyDamageFactor As Boolean = True)

With Stick(iStick)
    If 0 <= iDamager And iDamager < NumSticks Then
        If .IsBot Then
            'we are host/server[obviously] and stick being damaged is a bot
            
            If iDamager = 0 Then
                'if we're damaging said bot, tell us
                ReceiveDamageTick
            Else
                'we are the server, tell the stick he's damaging our bot
                modWinsock.SendPacket lSocket, Stick(iDamager).SockAddr, sDamageTicks & Stick(iDamager).ID
            End If
            
            
        ElseIf Stick(iDamager).IsBot = False Then
            'no point telling a bot he's doing damage
            
            'we're being damaged, since this <u>procedure</u> is only called if it's us or a bot,
            'and iStick != bot, from the start of the IF
            
            
            'HUMAN BEING DAMAGED BY A HUMAN
            
            If iDamager <> iStick Then
                If modStickGame.StickServer Then
                    modWinsock.SendPacket lSocket, Stick(iDamager).SockAddr, sDamageTicks & Stick(iDamager).ID
                Else
                    modWinsock.SendPacket lSocket, ServerSockAddr, sDamageTicks & Stick(iDamager).ID
                    'server will either receive or redirect
                End If
            End If
            
        End If
    End If
    
    
    '##################################################
    'adjust depending on settings
    If bApplyDamageFactor Then DamageToDo = DamageToDo * modStickGame.sv_Damage_Factor
    If modStickGame.sv_Hardcore Then DamageToDo = DamageToDo * Hardcore_Damage_Amp
    'If modStickGame.sv_GameType = gCoOp Then
    '    If .IsBot = False Then
    '        DamageToDo = DamageToDo \ 2
    '    End If
    'End If
    '##################################################
    
    
    
    '##################################################
    'perks/weapons, etc
    If .WeaponType = Chopper Then
        DamageToDo = DamageToDo \ Chopper_Damage_Reduction
        bDamageShield = False
    End If
    If iStick = 0 Then
        'me, check for low damage perk
        If .Perk = pJuggernaut Then
            DamageToDo = DamageToDo \ JuggernautDamageReduction
        End If
    End If
    If .Perk = pSniper Then
        DamageToDo = DamageToDo * Sniper_Damage_Inc 'double damage for snipers
    End If
    '##################################################
    
    
    
    '##################################################
    'apply
    If .Shield > 0 And bDamageShield Then
        .LastShieldHitTime = GetTickCount()
        .ShieldCharging = False
        
        If DamageToDo <= ShieldDamageDec Then
            .Shield = .Shield - 1
        Else
            .Shield = .Shield - DamageToDo \ ShieldDamageDec
        End If
        
        
        If .Shield < 0 Then
            .Health = .Health + .Shield
            .Shield = 0
        End If
        
        If .Shield = 0 Then
            'shields have been knocked out
            'AddExplosion .X, .Y, 500
            ShowShieldsDown iStick
        End If
    Else
        If DamageToDo = 0 Then
            DamageToDo = 1
        End If
        
        .Health = .Health - DamageToDo
    End If
    '##################################################
End With

End Sub

Private Function GetBulletDelay(i As Integer) As Long


If Stick(i).Perk = pMechanic Then
    GetBulletDelay = kBulletDelay(Stick(i).WeaponType) / Mechanic_Bullet_Inc
Else
    GetBulletDelay = kBulletDelay(Stick(i).WeaponType)
End If


'/ GetSticksTimeZone(i)

'If Stick(i).WeaponType = W1200 Then
'    GetBulletDelay = W1200_Bullet_Delay
'ElseIf Stick(i).WeaponType = AK Then
'    GetBulletDelay = AK_Bullet_Delay
'ElseIf Stick(i).WeaponType = M82 Then
'    GetBulletDelay = M82_Bullet_Delay
'ElseIf Stick(i).WeaponType = XM8 Then
'    GetBulletDelay = XM8_Bullet_Delay
'ElseIf Stick(i).WeaponType = DEagle Then
'    GetBulletDelay = DEagle_Bullet_Delay
'ElseIf Stick(i).WeaponType = M249 Then
'    GetBulletDelay = M249_Bullet_Delay
'ElseIf Stick(i).WeaponType = RPG Then
'    GetBulletDelay = RPG_Bullet_Delay  'needed to prevent spam
'ElseIf Stick(i).WeaponType = Chopper Then
'    GetBulletDelay = Chopper_Bullet_Delay
'ElseIf Stick(i).WeaponType = FlameThrower Then
'    GetBulletDelay = Flame_Bullet_Delay
'Else
'    GetBulletDelay = Knife_Delay
'End If
End Function

Private Sub MakeBulletStatsArray()
Dim i As Integer

'kbulletdelay = array(

For i = 0 To eWeaponTypes.Chopper
    kBulletSpeed(i) = BULLET_SPEED
    
    If i = W1200 Then
        kBulletDelay(i) = W1200_Bullet_Delay
        kBulletDamage(i) = W1200_Bullet_Damage
        kSprayAngle(i) = W1200_Spray_Angle
        
    ElseIf i = AK Then
        kBulletDelay(i) = AK_Bullet_Delay
        kBulletDamage(i) = AK_Bullet_Damage
        kSprayAngle(i) = AK_Spray_Angle
        
    ElseIf i = M82 Then
        kBulletDelay(i) = M82_Bullet_Delay
        kBulletDamage(i) = M82_Bullet_Damage
        kSprayAngle(i) = 0 'm82_Spray_Angle
        kBulletSpeed(i) = M82_Speed
        
    ElseIf i = XM8 Then
        kBulletDelay(i) = XM8_Bullet_Delay
        kBulletDamage(i) = XM8_Bullet_Damage
        kSprayAngle(i) = XM8_Spray_Angle
        
    ElseIf i = DEagle Then
        kBulletDelay(i) = DEagle_Bullet_Delay
        kBulletDamage(i) = DEagle_Bullet_Damage
        kSprayAngle(i) = DEagle_Spray_Angle
        
    ElseIf i = M249 Then
        kBulletDelay(i) = M249_Bullet_Delay
        kBulletDamage(i) = M249_Bullet_Damage
        kSprayAngle(i) = M249_Spray_Angle
        
    ElseIf i = RPG Then
        kBulletDelay(i) = RPG_Bullet_Delay  'needed to prevent spam
        'kBulletDamage(i) = rpg_Bullet_Damage
        'kSprayAngle(i) = _Spray_Angle
        
    ElseIf i = Chopper Then
        kBulletDelay(i) = Chopper_Bullet_Delay
        kBulletDamage(i) = Chopper_Bullet_Damage
        kSprayAngle(i) = Chopper_Spray_Angle
        
    ElseIf i = FlameThrower Then
        kBulletDelay(i) = Flame_Bullet_Delay
        'kBulletDamage(i) = _Bullet_Damage
        'kSprayAngle(i) = flame_Spray_Angle
        
    ElseIf i = AUG Then
        kBulletDelay(i) = AUG_Single_Bullet_Delay 'AUG_Bullet_Delay
        kBulletDamage(i) = AUG_Bullet_Damage
        kSprayAngle(i) = AUG_Spray_Angle
        
    ElseIf i = USP Then
        kBulletDelay(i) = USP_Bullet_Delay
        kBulletDamage(i) = USP_Bullet_Damage
        kSprayAngle(i) = 0 'usp_Spray_Angle
        
    ElseIf i = AWM Then
        kBulletDelay(i) = AWM_Bullet_Delay
        kBulletDamage(i) = AWM_Bullet_Damage
        kSprayAngle(i) = 0 '_Spray_Angle
        kBulletSpeed(i) = AWM_Speed
        
    ElseIf i = MP5 Then
        kBulletDelay(i) = MP5_Bullet_Delay
        kBulletDamage(i) = MP5_Bullet_Damage
        kSprayAngle(i) = MP5_Spray_Angle
        
    ElseIf i = Mac10 Then
        kBulletDelay(i) = Mac10_Bullet_Delay
        kBulletDamage(i) = Mac10_Bullet_Damage
        kSprayAngle(i) = Mac10_Spray_Angle
        
    ElseIf i = SPAS Then
        kBulletDelay(i) = SPAS_Bullet_Delay
        kBulletDamage(i) = SPAS_Bullet_Damage
        kSprayAngle(i) = SPAS_Spray_Angle
        
    ElseIf i = G3 Then
        kBulletDelay(i) = G3_Single_Bullet_Delay
        kBulletDamage(i) = G3_Bullet_Damage
        kSprayAngle(i) = 0 'G3_Spray_Angle
        
    Else
        kBulletDelay(i) = Knife_Delay
    End If
    
Next i

End Sub

Private Sub StartReload(iStick As Integer)
Dim Gauge As Integer
Dim bShotgun As Boolean

With Stick(iStick)
    
    If StickiHasState(iStick, STICK_RELOAD) = False Then
        AddStickiState iStick, STICK_RELOAD
        SubStickiState iStick, STICK_FIRE
        
        bShotgun = WeaponIsShotgun(.WeaponType)
        
        If bShotgun Then
            Gauge = GetGauge(.WeaponType)
            If .BulletsFired = Gauge Then
                .BulletsFired = 2 * Gauge 'add a delay to the single-reload
            End If
        End If
        
        If iStick = 0 Then
            FireKey = False
            
            If Not bShotgun Then
                modAudio.PlayReloadSound .WeaponType
            End If
        End If
        
        '.BulletsFired = 0'done at end
        .BulletsFired2 = 0
        .ReloadStart = GetTickCount()
    End If
End With
End Sub
Private Sub DoReload(iStick As Integer)
Dim nBullets As Integer

With Stick(iStick)
    If .WeaponType = Knife Or .WeaponType = Chopper Then '.WeaponType = W1200
        If .BulletsFired > 0 Then
            .BulletsFired = 0
        End If
        If StickiHasState(iStick, STICK_RELOAD) Then
            SubStickiState iStick, STICK_RELOAD
        End If
        Exit Sub
    ElseIf WeaponIsShotgun(.WeaponType) Then
        DoShotgunReload iStick
        Exit Sub
    End If
    
    nBullets = GetMaxRounds(Stick(iStick).WeaponType)
    
    If StickiHasState(iStick, STICK_RELOAD) Then
        
        If .ReloadStart + GetReloadTime(iStick) < GetTickCount() Then
            SubStickiState iStick, STICK_RELOAD
            
            .BulletsFired = 0
            
            If .WeaponType = RPG Then
                If StickiHasState(iStick, STICK_FIRE) Then
                    SubStickiState iStick, STICK_FIRE
                End If
            End If
            
            If iStick = 0 Then
                If Stick(0).WeaponType < Knife Then
                    If TotalMags(Stick(0).WeaponType) > 0 Then
                        TotalMags(Stick(0).WeaponType) = TotalMags(Stick(0).WeaponType) - 1
                    End If
                End If
            End If
            
        End If
        
        
        'auto reload below
    ElseIf .BulletsFired >= nBullets Then
        
        'SubStickState .ID, Stick_Fire
        'FireKey = False
        If iStick = 0 Then
            If Stick(0).WeaponType < Knife Then
                If TotalMags(Stick(0).WeaponType) = 0 Then
                    If StickiHasState(0, STICK_RELOAD) Then
                        SubStickiState 0, STICK_RELOAD
                    ElseIf StickiHasState(0, STICK_FIRE) Then
                        SubStickiState 0, STICK_FIRE
                        FireKey = False
                    End If
                    
                    Exit Sub
                End If
            End If
            
        End If
        
        
        
        If .LastBullet + AutoReload_Delay / GetSticksTimeZone(iStick) < GetTickCount() Then
            Call StartReload(iStick)
        End If
        
    'ElseIf StickHasState(.ID, Stick_Reload) Then
        'SubStickState .ID, Stick_Reload
    End If
    
End With

End Sub
Private Function WeaponIsShotgun(vWeapon As eWeaponTypes) As Boolean
WeaponIsShotgun = (vWeapon = SPAS) Or (vWeapon = W1200)
End Function
Private Function GetGauge(vWeapon As eWeaponTypes) As Integer
GetGauge = IIf(vWeapon = W1200, W1200_Gauge, SPAS_Gauge)
End Function
Private Sub DoShotgunReload(iStick As Integer)
Dim lDelay As Long
Dim bW1200 As Boolean

bW1200 = (Stick(iStick).WeaponType = W1200)


If StickiHasState(iStick, STICK_RELOAD) Then
    
    If TotalMags(Stick(iStick).WeaponType) > 0 Then
        If Stick(iStick).BulletsFired > 0 Then
            
            lDelay = IIf(bW1200, _
                    W1200_Round_Reload_Delay, SPAS_Round_Reload_Delay) _
                    / GetSticksTimeZone(iStick)
            
            
            
            If Stick(iStick).Perk = pSleightOfHand Then
                lDelay = lDelay / SleightOfHandReloadDecrease
            End If
            
            
            If Stick(iStick).LastRoundIn + lDelay < GetTickCount() Then
                
                Stick(iStick).BulletsFired = Stick(iStick).BulletsFired - GetGauge(Stick(iStick).WeaponType)
                
                If iStick = 0 Then
                    TotalMags(Stick(iStick).WeaponType) = TotalMags(Stick(iStick).WeaponType) - 1
                    
                    modAudio.PlayReloadSound W1200
                End If
                
                Stick(iStick).LastRoundIn = GetTickCount()
            End If
        Else
            SubStickiState iStick, STICK_RELOAD
        End If
    Else
        SubStickiState iStick, STICK_RELOAD
    End If
    
    
ElseIf (Stick(iStick).BulletsFired / GetGauge(Stick(iStick).WeaponType)) >= IIf(bW1200, W1200_Bullets, SPAS_Bullets) Then
    
    If iStick = 0 Then
        If TotalMags(Stick(iStick).WeaponType) = 0 Then
            If StickiHasState(0, STICK_RELOAD) Then
                SubStickiState 0, STICK_RELOAD
            ElseIf StickiHasState(0, STICK_FIRE) Then
                SubStickiState 0, STICK_FIRE
            End If
            
            Exit Sub
        End If
        
    End If
    
    
    If Stick(iStick).LastBullet + AutoReload_Delay / GetSticksTimeZone(iStick) < GetTickCount() Then
        Call StartReload(iStick)
    End If
End If


End Sub

Private Sub CheckStickHealthPack(iStick As Integer)
Const HealthPack_RadiusXX = HealthPack_Radius * 4.5


If HealthPack.bActive Then
    If Stick(iStick).WeaponType <> Chopper Then
        If GetDist(Stick(iStick).X, Stick(iStick).Y, HealthPack.X, HealthPack.Y) < HealthPack_RadiusXX Then
            
            With Stick(iStick)
                .Health = Max_Health
                
                ResetStickFireAndFlash iStick
                
                If .Shield < Max_Shield Then
                    If .Shield = 0 Then .Shield = 1
                    ResetTimeLong .LastShieldHitTime, Shield_Recharge_Delay
                End If
            End With
            
            'AddCirc HealthPack.X, HealthPack.Y, 500, 2, vbGreen
            AddInfoCirc HealthPack.X, HealthPack.Y, vbGreen
            
            HealthPack.bActive = False
            HealthPack.LastUsed = GetTickCount()
            
            If iStick = 0 Then
                modAudio.PlayMedKit
                FillTotalMags
            End If
            
        End If
    End If
End If


End Sub

Private Sub AddInfoCirc(X As Single, Y As Single, VCol As Long)
AddCirc X, Y, 500, 1, VCol, 100, False
End Sub

Private Sub ResetStickFireAndFlash(i As Integer)
ResetStickFire i
ResetStickFlash i
End Sub
Private Sub ResetStickFire(i As Integer)
Stick(i).bOnFire = False
'Stick(i).LastFlameTouch = GetTickCount() - Flame_Burn_Time / 0.0001
ResetTimeLong Stick(i).LastFlameTouch, Flame_Burn_Time
End Sub
Private Sub ResetStickFlash(i As Integer)
Stick(i).bFlashed = False
'Stick(i).LastFlashBang = GetTickCount() - FlashBang_Time / 0.0001
ResetTimeLong Stick(i).LastFlashBang, FlashBang_Time
End Sub

Private Sub AddMagForStick(i As Integer)
Dim vMag As eMagTypes

Select Case Stick(i).WeaponType
    Case AK
        vMag = mAK
    Case XM8, G3
        vMag = mXM8
    Case M82, M249
        vMag = mSniper
    Case DEagle, USP, AWM
        vMag = mPistol
    Case FlameThrower
        vMag = mFlameThrower
    Case AUG, MP5, Mac10
        vMag = mAUG
    Case Else
        'vMag = -1
        Exit Sub
End Select

AddMag Stick(i).CasingPoint.X, Stick(i).CasingPoint.Y, Stick(i).Speed, Stick(i).Heading, vMag

End Sub

Private Function GetMaxRounds(vWeapon As eWeaponTypes) As Integer 'Sticki As Integer) As Integer

GetMaxRounds = kMaxRounds(vWeapon)

'If vWeapon = AK Then
'    GetMaxRounds = AK_Bullets
'ElseIf vWeapon = M82 Then
'    GetMaxRounds = M82_Bullets
'ElseIf vWeapon = W1200 Then
'    GetMaxRounds = W1200_Bullets * W1200_Gauge
'ElseIf vWeapon = XM8 Then
'    GetMaxRounds = XM8_Bullets
'ElseIf vWeapon = DEagle Then
'    GetMaxRounds = DEagle_Bullets
'ElseIf vWeapon = M249 Then
'    GetMaxRounds = M249_Bullets
'ElseIf vWeapon = RPG Then
'    GetMaxRounds = RPG_Bullets
'ElseIf vWeapon = FlameThrower Then
'    GetMaxRounds = Flame_Bullets
'Else
'    GetMaxRounds = 1
'End If

End Function

Private Sub MakeMaxRoundsArray()
Dim i As Integer

For i = 0 To eWeaponTypes.Chopper
    If i = AK Then
        kMaxRounds(i) = AK_Bullets
    ElseIf i = M82 Then
        kMaxRounds(i) = M82_Bullets
    ElseIf i = W1200 Then
        kMaxRounds(i) = W1200_Bullets * W1200_Gauge
    ElseIf i = XM8 Then
        kMaxRounds(i) = XM8_Bullets
    ElseIf i = DEagle Then
        kMaxRounds(i) = DEagle_Bullets
    ElseIf i = M249 Then
        kMaxRounds(i) = M249_Bullets
    ElseIf i = RPG Then
        kMaxRounds(i) = RPG_Bullets
    ElseIf i = FlameThrower Then
        kMaxRounds(i) = Flame_Bullets
    ElseIf i = AUG Then
        kMaxRounds(i) = AUG_Bullets
    ElseIf i = USP Then
        kMaxRounds(i) = USP_Bullets
    ElseIf i = AWM Then
        kMaxRounds(i) = AWM_Bullets
    ElseIf i = MP5 Then
        kMaxRounds(i) = MP5_Bullets
    ElseIf i = Mac10 Then
        kMaxRounds(i) = Mac10_Bullets
    ElseIf i = SPAS Then
        kMaxRounds(i) = SPAS_Bullets * SPAS_Gauge
    ElseIf i = G3 Then
        kMaxRounds(i) = G3_Bullets
    Else
        kMaxRounds(i) = 1
    End If
Next i

End Sub

Private Function CoOrdInStick(X As Single, Y As Single, Sticki As Integer) As Boolean

Const AL1p5 = ArmLen * 1.5
Const K = BodyLen + LegHeight * 2
Const CLDx = ChopperLen / 1.2
Dim rc1 As RECT

If Stick(Sticki).WeaponType = Chopper Then
    
'    If X > (Stick(Sticki).X - CLDx) Then
'        If X < (Stick(Sticki).X + CLD4) Then
'            If Y > Stick(Sticki).Y - CLD8 Then
'                If Y < Stick(Sticki).Y + CLD6 Then
'                    CoOrdInStick = True
'                End If
'            End If
'        End If
'    End If
    
    'modStickGame.sBox Stick(iStick).X - ChopperLen / 1.2, Stick(iStick).Y - CLD4, _
                       Stick(iStick).X + CLD4,             Stick(iStick).Y + CLD3, _
                       vbRed
    
    With rc1
        .Left = Stick(Sticki).X - CLDx
        .Right = Stick(Sticki).X + CLD4
        .Top = Stick(Sticki).Y - CLD4
        .Bottom = Stick(Sticki).Y + CLD3
    End With
Else
    
'    If Abs(X - CLng(Stick(Sticki).X)) < AL1p5 Then
'        sY = GetStickY(Sticki)
'        If Y > sY Then '(Stick(Sticki).y) Then
'            CoOrdInStick = (Y < sY + K)
'        End If
'    End If
    
    With rc1
        .Left = Stick(Sticki).X - AL1p5
        .Right = Stick(Sticki).X + AL1p5
        .Top = GetStickY(Sticki)
        .Bottom = .Top + K
    End With
End If

CoOrdInStick = RectCollision(rc1, PointToRect(X, Y))

End Function

Private Function CoOrdInChopper(X As Single, Y As Single, iChopper As Integer) As Boolean
Const CLDx = ChopperLen / 1.2
Dim rcChopper As RECT

'If X > (Stick(iChopper).X - CLDx) Then
'    If X < (Stick(iChopper).X + CLD4) Then
'        If Y > Stick(iChopper).Y - CLD8 Then
'            If Y < Stick(iChopper).Y + CLD6 Then
'                CoOrdInChopper = True
'            End If
'        End If
'    End If
'End If

With rcChopper
    .Left = Stick(iChopper).X - CLDx
    .Right = Stick(iChopper).X + CLD4
    .Top = Stick(iChopper).Y - CLD4
    .Bottom = Stick(iChopper).Y + CLD3
End With

CoOrdInChopper = RectCollision(rcChopper, PointToRect(X, Y))

End Function

Private Function GetStickY(i As Integer) As Single
Const BodyLenX1p3 = BodyLen * 1.3
Const BodyLenD2 = BodyLen / 2

If StickiHasState(i, STICK_PRONE) Then
    GetStickY = Stick(i).Y + BodyLenX1p3
ElseIf StickiHasState(i, STICK_CROUCH) Then
    GetStickY = Stick(i).Y + BodyLenD2
Else
    GetStickY = Stick(i).Y
End If

End Function

Private Function BulletInHead(Sticki As Integer, Bulleti As Integer) As Boolean

Const HeadRadiusX2 = HeadRadius * 2
Const HeadRadiusX3 = HeadRadius * 3
Const HeadRadiusX1p5 = HeadRadius * 1.5
Dim rcHead As RECT


With rcHead
    .Left = Stick(Sticki).X - HeadRadiusX3
    .Right = Stick(Sticki).X + HeadRadiusX3
    .Top = GetStickY(Sticki) - HeadRadius
    .Bottom = .Top + HeadRadiusX3 'HeadRadiusX2 + Headradius
End With

BulletInHead = RectCollision(rcHead, PointToRect(Bullet(Bulleti).X, Bullet(Bulleti).Y))



'If Stick(Sticki).WeaponType <> Chopper Then
'    If Bullet(Bulleti).X > (Stick(Sticki).X - HeadRadiusX2) Then
'        If Bullet(Bulleti).X < (Stick(Sticki).X + HeadRadiusX2) Then
'            'yes, supposed to by -10 below
'            'If Bullet(Bulleti).Y > GetStickY(Sticki) - 30 Then '(Stick(Sticki).y - 10) Then
'            If Bullet(Bulleti).Y < GetStickY(Sticki) + HeadRadiusX2 Then '(Stick(Sticki).y + HeadRadiusX2) Then
'                BulletInHead = True
'            End If
'        End If
'    End If
'End If

End Function

'Private Function NadeInBullet(Nadei As Integer) As Boolean
'Dim i As Integer
'
'For i = 0 To NumBullets - 1
'    If Bullet(i).bSilenced = False Then
'        If BulletNearNade(Nadei, i) Then
'            NadeInBullet = True
'            Exit For
'        End If
'    End If
'Next i
'
'End Function
'
'Private Function BulletNearNade(Nadei As Integer, Bulleti As Integer) As Boolean
'Const NadeLim = 150
'
'If Bullet(Bulleti).bHeadingChanged = False Or Bullet(Bulleti).bSniperBullet Then
'    If Bullet(Bulleti).X > (Nade(Nadei).X - NadeLim) Then
'        If Bullet(Bulleti).X < (Nade(Nadei).X + NadeLim) Then
'
'            If Bullet(Bulleti).Y > (Nade(Nadei).Y - NadeLim) Then
'                If Bullet(Bulleti).Y < (Nade(Nadei).Y + NadeLim) Then
'                    BulletNearNade = True
'                End If
'            End If
'
'        End If
'    End If
'End If
'
'End Function

Private Sub ResetYComp(iStick As Integer)
With Stick(iStick)
    .Speed = .Speed * Abs(CoSine(.Heading))
    'abs, because it's like Sqr(XComp^2 + YComp^2) = +ve
    '??? or not?
    
    
    .Heading = FixAngle(.Heading)
    .Heading = IIf(0 <= .Heading And .Heading <= Pi, 0, Pi) * IIf(.Speed > 0, 1, -1)
End With
End Sub
Private Sub ResetXComp(iStick As Integer)
With Stick(iStick)
    .Speed = .Speed * Abs(Sine(.Heading))
    'abs, because it's like Sqr(XComp^2 + YComp^2) = +ve
    '??? or not?
    
    
    .Heading = FixAngle(.Heading)
    .Heading = IIf(.Heading < Pi, piD2, pi3D2) * IIf(.Speed > 0, 1, -1)
End With
End Sub

Private Sub ApplyGravity(iStick As Integer, lasty As Single) ', Optional bResetSpeed As Boolean = True)
Dim j As Integer
Dim bOnPlatform As Boolean
Dim ySpeed As Single ', yDiff As Single

Const Inc_Into_Platform As Single = 20 - BodyLen - LegHeight, _
      Damage_Dec As Single = 6, _
      Min_Speed_For_Fall_Damage = -145 ', _
      Fall_Dist_Damage As Single = 3050


'fly mode
'If Not Stick(iStick).OnSurface Then Stick(iStick).OnSurface = True
'Exit Sub

'PrintStickText "YSpeed: " & Round(Stick(iStick).Speed * CoSine(Stick(iStick).Heading), 2), Stick(0).X, Stick(0).Y - 1000, vbBlack


On Error GoTo EH

If Stick(iStick).Y > StickGameHeight Then
    j = 0 'prevent floor bug
          'by forcing them to be on floor 0
    GoTo On_Platform 'mid loop, can't be bothered restructuring code
End If


'If Stick(iStick).Y < Stick(iStick).JumpStartY Then
'    Stick(iStick).JumpStartY = Stick(iStick).JumpStartY
'End If

'If Stick(iStick).OnSurface = False Then
    'If StickiHasState(iStick, stick_Jump) = False Then
        For j = 0 To ubdPlatforms
            If StickPassedThroughSurface(iStick, j, lasty) Then
                
On_Platform:
                bOnPlatform = True
                
                Stick(iStick).iCurrentPlatform = j
                
                Stick(iStick).Y = Platform(j).Top + Inc_Into_Platform
                ySpeed = Stick(iStick).Speed * CoSine(Stick(iStick).Heading)
                
                'If bResetSpeed Then
                    ResetXComp iStick
                'End If
                
                
                
                'yDiff = Stick(iStick).Y - Stick(iStick).JumpStartY
                'Stick(iStick).JumpStartY = Stick(iStick).Y
                
                
                If ySpeed < Min_Speed_For_Fall_Damage Then
                    '     ^ because negative
                    
                    
                    If PointHearableOnSticksScreen(Stick(iStick).X, Stick(iStick).Y, 0) Then
                        modAudio.PlayLandSound GetRelPan(Stick(iStick).X)
                    End If
                    
                    
                    
                    If StickInvul(iStick) Or Not Stick(iStick).bTouchedSurface Then
                        'prevent damage unless.. yeah
                        Stick(iStick).bTouchedSurface = True
                    Else
                        AddBloodExplosion Stick(iStick).X, Stick(iStick).Y
                        
                        If iStick = 0 Or Stick(iStick).IsBot Then
                            
                            
                            DamageStick Abs(ySpeed) / (Damage_Dec * _
                                IIf(Stick(iStick).Perk = pJuggernaut Or Stick(iStick).Perk = pConditioning, 3, 1)), _
                                    iStick, iStick, False, False
                            
                            If Stick(iStick).Health < 1 Then
                                Killed iStick, iStick, kFall
                            End If
                            
                            
                        End If
                    End If
                End If
                
                
                'If Platform(j).iType = pSpikes Then
                    'If iStick = 0 Or Stick(iStick).IsBot Then
                        'Killed iStick, iStick, kSpikes
                    'End If
                'End If
                
                
                Exit For
            End If
        Next j
    'End If
'Else
'    bOnPlatform = True
'End If


If Not bOnPlatform Then
    
    Stick(iStick).iCurrentPlatform = -1
    
    ApplyGravityVector Stick(iStick).LastGravity, Stick(iStick).sgTimeZone, _
        Stick(iStick).Speed, Stick(iStick).Heading, Stick(iStick).X, Stick(iStick).Y
    
    
    If StickiHasState(iStick, STICK_CROUCH) Then
        SubStickiState iStick, STICK_CROUCH
    ElseIf StickiHasState(iStick, STICK_PRONE) Then
        SubStickiState iStick, STICK_PRONE
    End If
'    ElseIf StickiHasState(iStick, STICK_LEFT) Then
'        SubStickiState iStick, STICK_LEFT
'    ElseIf StickiHasState(iStick, STICK_RIGHT) Then
'        SubStickiState iStick, STICK_RIGHT
'    End If
    
End If


Stick(iStick).bOnSurface = bOnPlatform

EH:
End Sub

Private Function StickPassedThroughSurface(iStick As Integer, iPlatform As Integer, lasty As Single) As Boolean
Dim rcStick As RECT

With rcStick
    .Left = Stick(iStick).X
    .Right = .Left + 1 'don't use .X, because .X is a single, and could cause rounding errors
                       'e.g .X = 5.5   .Left = 6, .Right = .X + 1 = 6 ==> rectangle is a line ==> fail
    .Top = Stick(iStick).Y + BodyLen + LegHeight
    .Bottom = .Top + 1
End With

StickPassedThroughSurface = RectCollision(rcStick, PlatformToRect(Platform(iPlatform)))

End Function

Private Function PlatformToRect(vPlatform As ptStickPlatform) As RECT
With PlatformToRect
    .Left = vPlatform.Left
    .Top = vPlatform.Top
    .Right = .Left + vPlatform.width
    .Bottom = .Top + vPlatform.height
End With
End Function

Private Sub SomeoneDied(ByVal DeadStickj As Integer, ByVal iKiller As Integer, ByVal KillType As eKillTypes)
Dim bDeadStickExists As Boolean, bKillerExists As Boolean

On Error GoTo EH
bDeadStickExists = Stick(DeadStickj).bTyping Or True

If iKiller > -1 Then
    bKillerExists = Stick(iKiller).bTyping Or True
End If


If bKillerExists Then
    If Stick(iKiller).IsBot Then
        If Stick(iKiller).bAlive Then
            If DeadStickj <> iKiller Then
                ShowBotTaunt iKiller, DeadStickj
            End If
        End If
    End If
End If

If bDeadStickExists Then
    If Stick(DeadStickj).WeaponType < Knife Then
        AddStaticWeapon Stick(DeadStickj).X, Stick(DeadStickj).Y, Stick(DeadStickj).WeaponType
        
        With StaticWeapon(NumStaticWeapons - 1)
            If Stick(DeadStickj).Speed > MAX_DeadStick_And_StaticWeap_Speed Then
                .Speed = MAX_DeadStick_And_StaticWeap_Speed
            Else
                .Speed = Stick(DeadStickj).Speed
            End If
            
            .Heading = Stick(DeadStickj).Heading
        End With
    End If
    
    If Stick(DeadStickj).WeaponType = Chopper Then
        AddDeadChopper Stick(DeadStickj).X, Stick(DeadStickj).Y, Stick(DeadStickj).colour, DeadStickj
    Else 'If DeadStickj > 0 Then
        'otherwise, add mine in killed()
        AddDeadStick Stick(DeadStickj).X, Stick(DeadStickj).Y, Stick(DeadStickj).colour, _
            (Stick(DeadStickj).Facing < Pi), (KillType = kBurn Or KillType = kFlame), _
            Stick(DeadStickj).Perk = pSniper, _
            Stick(DeadStickj).Speed, Stick(DeadStickj).Heading, DeadStickj = 0
    
    End If
    
    
    Stick(DeadStickj).BulletsFired2 = 0
    'Stick(DeadStickj).LastSpawnTime = GetTickCount() + modStickGame.sv_Spawn_Delay
    Stick(DeadStickj).bTouchedSurface = False
    
    Stick(DeadStickj).bAlive = False
    If Stick(DeadStickj).Perk = pMartyrdom Then
        AddNade Stick(DeadStickj).X, Stick(DeadStickj).Y, 0, 10, DeadStickj, Stick(DeadStickj).colour, nFrag, , True
        
'        For iKiller = 0 To 50
'            AddNade Stick(DeadStickj).X, Stick(DeadStickj).Y, Rnd() * Pi2, Rnd() * 60, DeadStickj, Stick(DeadStickj).Colour, nFrag, , True
'            AddNade Stick(DeadStickj).X, Stick(DeadStickj).Y, Rnd() * Pi2, Rnd() * 60, DeadStickj, Stick(DeadStickj).Colour, nFrag, , True
'            AddNade Stick(DeadStickj).X, Stick(DeadStickj).Y, Rnd() * Pi2, Rnd() * 60, DeadStickj, Stick(DeadStickj).Colour, nFrag, , True
'        Next
    End If
End If

EH:
End Sub

Private Sub ShowBotTaunt(iBot As Integer, Dead_Stick As Integer)
Dim sTaunt As String
Dim i As Integer

If Not modStickGame.cl_StickBotChat Then Exit Sub
If Rnd() > 0.6 Then Exit Sub
'If Stick(iBot).Perk = pSniper Then Exit Sub 'stealth maitatined // doesn't matter
If Stick(iBot).Perk = pZombie Then Exit Sub


i = IntRand(1, UBound(BotTaunts))

If BotTaunts(i).bAddName Then
    sTaunt = BotTaunts(i).sTaunt & Trim$(Stick(Dead_Stick).Name) & "!"
Else
    sTaunt = BotTaunts(i).sTaunt
End If

'always server, no need to check
SendChatPacketBroadcast Trim$(Stick(iBot).Name) & modMessaging.MsgNameSeparator & sTaunt, Stick(iBot).colour

End Sub

Private Sub InitBotInfo()

ReDim BotTaunts(1 To 7)

BotTaunts(1).sTaunt = "You suck, " '& Name
BotTaunts(1).bAddName = True

BotTaunts(2).sTaunt = "Call that a dodge?!"

BotTaunts(3).sTaunt = "Have some of that!"

BotTaunts(4).sTaunt = "Die " ' & Name
BotTaunts(4).bAddName = True

BotTaunts(5).sTaunt = "Take that!"

BotTaunts(6).sTaunt = "I knew I'd win"

BotTaunts(7).sTaunt = "Goodbye, " '& Name
BotTaunts(7).bAddName = True


ReDim kBotNames(0 To 4)
kBotNames(0) = "Timmy"
kBotNames(1) = "Tim"
kBotNames(2) = "J. Randombot"
kBotNames(3) = "Stan"
kBotNames(4) = "Agent Smith"


End Sub

Private Sub RemoveStickFromBotTargetList(DeadStickj As Integer)
Dim i As Integer

For i = 1 To NumSticksM1
    If Stick(i).IsBot Then
        If i <> DeadStickj Then
            If StickSeenStick(i, DeadStickj) Then
                RemoveStickFromBotsTargets i, DeadStickj
            End If
        End If
    End If
Next i


End Sub

Private Sub Killed(ByVal DeadStickj As Integer, ByVal iKiller As Integer, ByVal KillType As eKillTypes)
Dim ChatText As String, FullText As String
Dim i As Integer
Dim bDeadStickExists As Boolean, bKillerExists As Boolean


On Error Resume Next
bDeadStickExists = Stick(DeadStickj).bTyping Or True
bKillerExists = Stick(iKiller).bTyping Or True

On Error GoTo EH
If DeadStickj <> -1 And bDeadStickExists Then
    
    
    'MUST be before stick is re-positioned
    SomeoneDied DeadStickj, iKiller, KillType
    
    If Stick(DeadStickj).Perk = pZombie Then
        If KillType = kHead Then
            AddHead Stick(DeadStickj).X, Stick(DeadStickj).Y, Zombie_Col, Stick(DeadStickj).Speed / 2, Stick(DeadStickj).Heading
        End If
    End If
    'END
    
    
    RemoveStickFromBotTargetList DeadStickj
    
    Stick(DeadStickj).Speed = 0
    Stick(DeadStickj).Heading = 0
    If Stick(DeadStickj).Perk = pZombie Then
        Stick(DeadStickj).Health = Zombie_Health
        Stick(DeadStickj).Shield = 0
    Else
        Stick(DeadStickj).Health = Health_Start
        Stick(DeadStickj).Shield = 0 'IIf(modStickGame.sv_SpawnWithShields, Max_Shield, 0)
    End If
    
    Stick(DeadStickj).BulletsFired = 0
    Stick(DeadStickj).BulletsFired2 = 0
    SetStickiState DeadStickj, STICK_NONE
    Stick(DeadStickj).X = (StickGameWidth - 1000) * Rnd()
    
    If Stick(DeadStickj).IsBot And Stick(DeadStickj).WeaponType = Chopper Then
        Stick(DeadStickj).Y = ChopperLen * Rnd()
    Else
        Stick(DeadStickj).Y = (StickGameHeight - 501) * Rnd()
    End If
    
    Stick(DeadStickj).Facing = Pi2 * Rnd()
    
    Stick(DeadStickj).LastSpawnTime = GetTickCount()
    'SubStickState Stick(DeadStickj).ID, stick_Left
    'SubStickState Stick(DeadStickj).ID, stick_Right
    Stick(DeadStickj).state = STICK_NONE
    Stick(DeadStickj).bOnSurface = False
    
    ResetStickFireAndFlash DeadStickj
    
    Stick(DeadStickj).iDeaths = Stick(DeadStickj).iDeaths + 1
    Stick(DeadStickj).iKillsInARow = 0
    'Stick(DeadStickj).LastFlashBang = GetTickCount() - FlashBang_Time / 0.1 - 1000
    'Stick(DeadStickj).bLightSaber = False
    'ResetJumpStartY DeadStickj
    
    
    Stick(DeadStickj).lDeathTime = Stick(DeadStickj).LastSpawnTime 'aka GTC
    
    
    If DeadStickj = 0 Then '=FindStick(MyID) Then
        
'        If modStickGame.sv_GameType <> gCoOp Then
'            If modStickGame.sv_GameType <> gElimination Then
'                AddCirc Stick(DeadStickj).X, Stick(DeadStickj).Y, 1000, 2, vbGreen
'            End If
'        End If
        
        'ResetAmmoFired
        FillTotalMags
        
        For i = 0 To eWeaponTypes.Knife
            AmmoFired(i) = 0
        Next i
        If Stick(DeadStickj).WeaponType = Chopper Then
            'it's me
            ChopperAvail = False
            SwitchWeapon Stick(0).CurrentWeapons(1)
        End If
        
        
        modAudio.StopWeaponReloadSound Stick(0).WeaponType
        
        If modStickGame.sv_GameType <> gDeathMatch Then
            Stick(0).sgTimeZone = modStickGame.sv_StickGameSpeed
            SetSoundFreq Stick(0).sgTimeZone
        End If
        
        HideCursor False
        
        'LeftKey = False
        'RightKey = False
        'JumpKey = False
        'ProneKey = False
        CrouchKey = False
        UseKey = False
        FireKey = False
        
        SpecLeft = False 'reset, just in case
        SpecRight = False
        SpecUp = False
        SpecDown = False
        
        FlamesInARow = 0
        KnifesInARow = 0
        
        If Rnd() > 0.2 Then
            modAudio.PlayDeathNoise
        Else
            modAudio.PlayWilhelm
        End If
    End If
    
    
    'If modStickGame.sv_GameType = gElimination Or modStickGame.sv_GameType = gCoOp Then
        Stick(DeadStickj).bAlive = False
    'End If
    
    
    If iKiller <> -1 And bKillerExists Then
        
        If DeadStickj = iKiller Then
            
            If KillType = kFall Then
                ChatText = Trim$(Stick(DeadStickj).Name) & " fell to his doom"
            ElseIf KillType = kCeiling Then
                ChatText = Trim$(Stick(DeadStickj).Name) & " cracked his head on the ceiling"
            ElseIf KillType = kSpikes Then
                ChatText = Trim$(Stick(DeadStickj).Name) & " was spiked"
            ElseIf KillType = kMine Or KillType = kAirMine Then
                ChatText = Trim$(Stick(DeadStickj).Name) & " mined himself"
            Else
                ChatText = Trim$(Stick(DeadStickj).Name) & " committed suicide"
            End If
            
        ElseIf KillType = kNormal Then
            
            If WeaponIsSniper(Stick(iKiller).WeaponType) Then
                If WeaponIsSniper(Stick(DeadStickj).WeaponType) Then
                    ChatText = "counter-sniped by " & Trim$(Stick(iKiller).Name)
                Else
                    ChatText = "sniped by " & Trim$(Stick(iKiller).Name)
                End If
            Else
                ChatText = "killed by " & Trim$(Stick(iKiller).Name)
            End If
            
        ElseIf KillType = kNade Then
            ChatText = "grenaded by " & Trim$(Stick(iKiller).Name)
            
        ElseIf KillType = kSilenced Then
            ChatText = "silenced by " & Trim$(Stick(iKiller).Name)
            
        ElseIf KillType = kRPG Then
            ChatText = "rocketed by " & Trim$(Stick(iKiller).Name)
            
        ElseIf KillType = kHead Then
            ChatText = IIf(Stick(iKiller).bSilenced, "stealth ", vbNullString) & "headshotted by " & Trim$(Stick(iKiller).Name)
            
        ElseIf KillType = kMine Then
            ChatText = "mined by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kFlame Then
            ChatText = "fried by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kBurn Then
            ChatText = "toasted by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kChoppered Then
            If Stick(DeadStickj).WeaponType = Chopper Then
                ChatText = "rammed by " & Trim$(Stick(iKiller).Name)
            Else
                ChatText = "diced by " & Trim$(Stick(iKiller).Name)
            End If
        ElseIf KillType = kKnife Then
            If Stick(iKiller).Perk = pZombie Then
                If Rnd() > 0.5 Then
                    ChatText = "mashed by " & Trim$(Stick(iKiller).Name)
                Else
                    ChatText = "ripped apart by " & Trim$(Stick(iKiller).Name)
                End If
            Else
                ChatText = "knifed by " & Trim$(Stick(iKiller).Name)
            End If
        ElseIf KillType = kCrushed Then
            ChatText = "crushed by " & frmMain.FormatApostrophe(Trim$(Stick(iKiller).Name)) & " chopper"
        ElseIf KillType = kFlameTag Then
            ChatText = "flame-tagged by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kLightSaber Then
            ChatText = "Lightsaber'd by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kBarrel Then
            ChatText = "Barrel'd by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kMartyrdom Then
            ChatText = "Martyrdom'd by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kAirMine Then
            ChatText = "Air-Mined by " & Trim$(Stick(iKiller).Name)
        End If
        
        
        If DeadStickj <> iKiller Then
            FullText = Trim$(Stick(DeadStickj).Name) & " was " & ChatText
            Stick(iKiller).iKills = Stick(iKiller).iKills + 1 'INCREASE HERE
            Stick(iKiller).iKillsInARow = Stick(iKiller).iKillsInARow + 1
        Else
            FullText = ChatText
        End If
        
        
        
        
        If StickServer Then
            SendChatPacketBroadcast FullText, Stick(iKiller).colour
        Else
            modWinsock.SendPacket lSocket, ServerSockAddr, sChats & FullText & "#" & CStr(Stick(iKiller).colour)
            
            'AddChatText ChatText, Stick(iKiller).Colour
            'we'll get it back
        End If
        
        
        If DeadStickj = 0 Then 'if we're dead, tell the server to add one to the killer's kills
            
            If DeadStickj <> iKiller Then
                AddMainMessage UCase$(Left$(ChatText, 1)) & Mid$(ChatText, 2), False, Stick(iKiller).colour
            End If
            
        ElseIf Stick(DeadStickj).IsBot Then
            
            Stick(DeadStickj).AI_Targets_Seen = vbNullString
            
            
            If iKiller = 0 Then
                If Stick(iKiller).WeaponType = FlameThrower Then
                    If KillType = kBurn Or KillType = kFlame Then
                        FlamesInARow = FlamesInARow + 1
                    End If
                End If
            End If
            
        End If
        
        
        If iKiller = 0 Then
            If iKiller = DeadStickj Then
                
                If KillType = kMine Or KillType = kAirMine Then
                    AddMainMessage "Suicide Mine!", False
                ElseIf KillType = kFall Then
                    AddMainMessage "Fall damage > You", False
                ElseIf KillType = kSpikes Then
                    AddMainMessage "Spiked!", False
                Else
                    AddMainMessage "Sigh", False
                End If
                
            Else
                Call CheckKillsInARow
            End If
        End If
        
        
        If modStickGame.StickServer = False Then
            'modWinsock.SendPacket lSocket, ServerSockAddr, sDeathInfos & CStr(Stick(DeadStickj).ID) & _
                CStr(Abs(KillType = kFlame Or KillType = kBurn))
            
            'modWinsock.SendPacket lSocket, ServerSockAddr, sKillInfos & CStr(Stick(iKiller).ID)
            modWinsock.SendPacket lSocket, ServerSockAddr, sKillAndDeathInfos & _
                CStr(Stick(iKiller).ID) & "#" & Stick(DeadStickj).ID & _
                CStr(Abs(KillType = kFlame Or KillType = kBurn))
            
            
        Else
            'SendBroadcast sDeathInfos & CStr(Stick(DeadStickj).ID) & _
                CStr(Abs(KillType = kFlame Or KillType = kBurn))
            
            'SendBroadcast sKillInfos & CStr(Stick(iKiller).ID)
            SendBroadcast sKillAndDeathInfos & _
                CStr(Stick(iKiller).ID) & "#" & Stick(DeadStickj).ID & _
                CStr(Abs(KillType = kFlame Or KillType = kBurn))
            
            
        End If
        
        
    End If
End If

EH:
End Sub

Private Sub ProcessRespawn()

Dim Time_Left As Long
Dim i As Integer

If modStickGame.sv_GameType = gDeathMatch Then
    If Stick(0).bAlive = False Then
        
        Time_Left = (Stick(0).lDeathTime + modStickGame.sv_Spawn_Delay - GetTickCount()) / 1000
        'used 2x below
        
        If Time_Left < 1 Then
            Stick(0).Shield = IIf(modStickGame.sv_SpawnWithShields, 1, 0)
            Stick(0).bAlive = True
            Stick(0).LastSpawnTime = GetTickCount()
            HideCursor True
            cg_sZoom = 1
        Else
            
            modStickGame.PrintStickFormText _
                "Respawn in " & CStr(Round(Time_Left)) & " second" & IIf(Time_Left > 1, "s", vbNullString) & "...", _
                StickCentreX - 888, StickCentreY, vbBlack
            
            '888=textwidth("Respawn in x seconds...")/2
        End If
    End If
    
    
    If modStickGame.StickServer Then
        For i = 1 To NumSticksM1
            If Stick(i).IsBot Then
                
                If Not Stick(i).bAlive Then
                    If (Stick(i).lDeathTime + modStickGame.sv_Spawn_Delay - GetTickCount()) < 1000 Then
                        Stick(i).Shield = IIf(modStickGame.sv_SpawnWithShields, 1, 0)
                        Stick(i).bAlive = True
                        Stick(i).LastSpawnTime = GetTickCount()
                    End If
                End If
                
            End If
        Next i
    End If
End If

End Sub

'Private Sub ResetAmmoFired()
'Dim i As eWeaponTypes
'
'For i = 0 To CInt(eWeaponTypes.Knife)
'    AmmoFired(CInt(i)) = GetMaxRounds(i) * 4
'Next i
'
'
'End Sub

Public Sub BackgroundColourChanged()
SetBulletTrail_defCol
End Sub

Private Function InitVariables() As Boolean

Dim Ctrl As Control
Dim sTxt As String
Const Mine_Dist_Inc = 300
Dim i As Integer

For Each Ctrl In Controls
    If Not (TypeOf Ctrl Is Timer) Then
        If Not (TypeOf Ctrl Is Shape) Then
            Ctrl.TabStop = False
        End If
    End If
Next Ctrl


'Add us as a Stick!
AddStick


'Allow all weapons
For i = 0 To eWeaponTypes.Chopper
    modStickGame.sv_AllowedWeapons(i) = True
Next i


'set BG colour for trails
BackgroundColourChanged


WeaponKey = -1

'If we're the host, assign our ID now
If StickServer Then
    Stick(0).ID = 0
    AdjustIDArray
End If
RandomizeMyStickPos
Stick(0).Facing = Pi2 * Rnd()
Stick(0).Health = Health_Start
Stick(0).Name = frmMain.LastName
Stick(0).colour = modVars.TxtForeGround
Current_Health_Start = Health_Start

'######################################################
LastDamageTick = GetTickCount() - DamageTickTime - 1
ChatFlashDelay = GetCursorBlinkTime()

LastWeaponSwitch = GetTickCount()
LastNadeSwitch = LastWeaponSwitch
LastFireModeSwitch = LastWeaponSwitch
LastProneSwitch = LastWeaponSwitch
LastCrouchToggle = LastWeaponSwitch
'######################################################

If InitMap() = False Then
    InitVariables = False
    AddText "Error Loading Map", TxtError, True
Else
    
    MakeBulletStatsArray
    MakeMaxRoundsArray
    MakeReloadTimeArray
    MakePerkNameArray
    'MakeWeaponNameArray
    MakeTeamColourArray
    MakeGameTypeArray
    InitWeaponStats
    MakeNadeNameArray
    
    
    FillTotalMags
    
    
    'increase detection dist, so clients won't drop dead spontaneously
    If modStickGame.StickServer Then
        Mine_StickLim = 1500 + Mine_Dist_Inc
        'Mine_StickLimY = 2000 + Mine_Dist_Inc
    Else
        Mine_StickLim = 1500
        'Mine_StickLimY = 2000
    End If
    
    InitBotInfo
    frmStickGame.SetCurrentWeapons
    
    
    
    If modAudio.bDXSoundEnabled Then
        
        PrintLoadingText "Initialising Sound..."
        
        
        If modAudio.InitStickSounds() Then 'load into files
            sTxt = vbNullString
            If modDXSound.DXSound_Init(Me.hWnd, sTxt) Then
                
                If LoadDXSounds() Then
                    'InitVariables = True
                    modAudio.bDXSoundEnabled = True
                Else
                    'InitVariables = False
                    modAudio.bDXSoundEnabled = False
                End If
            Else
                'InitVariables = False
                modAudio.bDXSoundEnabled = False
                
                AddText "Error Initialising DirectX - " & sTxt, TxtError, True
                
                'AddConsoleText "DX Error - " & Err.Description
            End If
        Else
            'InitVariables = False
            modAudio.bDXSoundEnabled = False
            AddText "Error Loading Sound Files - Couldn't Create Files", TxtError, True
            AddConsoleText "InitStickSounds() Error - " & Err.Description
        End If
        PrintLoadingText "Initialised Sound"
    Else
        PrintLoadingText "Skipped Sound Initialisation..."
    End If
    modAudio.bDXSoundInited = modAudio.bDXSoundEnabled
    
    InitVariables = True
    
    
    If modAudio.bDXSoundEnabled = False Then
        modDXSound.DXSound_Terminate
        PrintLoadingText "Sound Error"
    End If
    
    ''sound
    'modAudio.PlayNadeExplosion
    'modAudio.StopSound
End If


Stick(0).CurrentWeapons(1) = modStickGame.cl_StartWeapon1
Stick(0).CurrentWeapons(2) = modStickGame.cl_StartWeapon2
Stick(0).Perk = modStickGame.cl_StartPerk

'Stick(0).WeaponType = Stick(0).CurrentWeapons(1)
SetSticksWeapon 0, Stick(0).CurrentWeapons(1), False

Set_Default_FireMode 0


End Function

Private Sub PrintLoadingText(sTxt As String)

picMain.Cls
PrintStickFormText sTxt, StickCentreX - TextWidth(sTxt) / 2, StickCentreY - TextHeight(sTxt) * 2, vbBlack
BltToForm
Me.Refresh

End Sub

Private Sub InitVarsForMap()
Dim i As Integer

For i = 0 To NumMags - 1
    Mag(i).bOnSurface = False
Next i
For i = 0 To NumDeadSticks - 1
    DeadStick(i).bOnSurface = False
Next i


Erase StaticWeapon: NumStaticWeapons = 0
MakeStaticWeapons

Erase Barrel: NumBarrels = 0
AddBarrels

Erase WallMark: NumWallMarks = 0
Erase Casing: NumCasings = 0


Erase Grass: NumGrass = 0
If modStickGame.StickServer Then MakeGrass

GenerateAmmoPickup True

End Sub

'########################################################################################
Private Function LoadDXSounds() As Boolean
Dim i As Integer
Const def_Sound_Vol = -1500
Dim bTold As Boolean

If modLoadProgram.bIsIDE Then
    On Error GoTo 0
Else
    On Error GoTo EH
End If

'################################################
For i = 0 To eWeaponTypes.Chopper
    LoadDXSound WeaponPath(i)
Next i

For i = 0 To eWeaponTypes.RPG
    LoadDXSound ReloadPath(i)
Next i

LoadDXSound NadeExplosionPath
LoadDXSound NadeBouncePath
LoadDXSound NadeThrowPath
LoadDXSound NadeBGPath
LoadDXSound RifleBGPath
LoadDXSound SilencedPath
LoadDXSound Silenced2Path
LoadDXSound MedKitPath

For i = 1 To 3
    LoadDXSound DeathNoisePath(i)
Next i

LoadDXSound RoundStartPath
LoadDXSound ToastyPath

For i = 1 To 7
    LoadDXSound RicochetPath(i)
Next i

For i = 1 To 3
    LoadDXSound LightSaberPath(i)
Next i


LoadDXSound WeaponPickupPath
LoadDXSound LandSound
LoadDXSound WilhelmPath

LoadDXSound TickPath


SetDXSoundVol def_Sound_Vol
'################################################

LoadDXSounds = True

Exit Function
EH:
LoadDXSounds = False

If Not bTold Then
    AddConsoleText "DX Sound Load Error - " & Err.Description
    AddText "Error Loading Sounds into DirectX (" & Err.Description & ")", TxtError, True
    AddText "Continuing Anyway...", , True
    
    bTold = True
End If

Resume Next
End Function

Private Function LoadDXSound(sFilePath As String) As Integer
modDXSound.LoadSound sFilePath, True, False, True, True, True, False, Me
End Function

Public Sub SetDXSoundVol(lVol As Long)
Dim i As Integer, iTick As Integer

iTick = modAudio.StickTickSoundIndex

On Error Resume Next
For i = 0 To modDXSound.nSounds
    If i <> iTick Then 'tick sound
        modDXSound.SetVolume i, lVol
    End If
Next i

End Sub

Private Sub SetSoundFreq(nMultiple As Single)
Dim i As Integer, iTick As Integer

If modAudio.bDXSoundEnabled = False Then Exit Sub


iTick = modAudio.StickTickSoundIndex


For i = 0 To modDXSound.nSounds
    If i <> iTick Then
        modDXSound.SetRelativeFrequency i, nMultiple
    End If
Next i

If Stick(0).Perk = pSleightOfHand Then
    For i = 0 To eWeaponTypes.Knife - 1
        If i <> iTick Then
            modDXSound.SetRelativeFrequency CInt(i + eWeaponTypes.Chopper + 1), _
                nMultiple * modStickGame.SleightOfHandReloadDecrease
        End If
    Next i
End If

End Sub

'########################################################################################

Private Sub FillTotalMags()
Dim i As Integer

For i = 0 To eWeaponTypes.Knife - 1
    TotalMags(i) = GetTotalMags(CInt(i))
Next i

End Sub

Public Sub RandomizeMyStickPos()
Stick(0).X = StickGameWidth * Rnd()
Stick(0).Y = (StickGameHeight - 1000) * Rnd()
End Sub

Private Function InitMap() As Boolean
'Dim i As Integer
'Const Def_Height = 375
Dim f As Integer
Dim DefMapPath As String

'############################################################################
'create default map
f = FreeFile()

DefMapPath = modStickGame.GetStickMapPath() & "Default." & Map_Ext

On Error GoTo EH
Open DefMapPath For Output As #f
    Print #f, LoadResText(103);
Close #f
'############################################################################


If modStickGame.StickServer Then
    InitMap = LoadMapEx(modStickGame.StickMapPath)
Else
    'dealt with in RequestMap()
    InitMap = True
End If


Exit Function
EH:
InitMap = False
Close #f
End Function
'######################################################################################################
'ubdPlatforms = 7
'ReDim Platform(ubdPlatforms)
''Platform
'Platform(0).Left = -1000: Platform(0).Top = 13572: Platform(0).width = 52000: Platform(0).height = 855
'Platform(1).Left = 0: Platform(1).Top = 11400: Platform(1).width = 7575: Platform(1).height = Def_Height
'Platform(2).Left = 6240: Platform(2).Top = 8400: Platform(2).width = 25000: Platform(2).height = Def_Height
'Platform(3).Left = 840: Platform(3).Top = 6000: Platform(3).width = 4935: Platform(3).height = Def_Height
'Platform(4).Left = 13000: Platform(4).Top = 5500: Platform(4).width = 10000: Platform(4).height = Def_Height
'Platform(5).Left = 12120: Platform(5).Top = 11400: Platform(5).width = 35000: Platform(5).height = Def_Height
'Platform(6).Left = 44500: Platform(6).Top = 5000: Platform(6).width = 5500: Platform(6).height = Def_Height
'Platform(7).Left = 42000: Platform(7).Top = 7750: Platform(7).width = 500: Platform(7).height = Def_Height
'
'
'ubdtBoxes = 8
'ReDim tBox(ubdtBoxes)
''tBox
'tBox(0).Left = 7200: tBox(0).Top = 10905: tBox(0).width = Def_Height: tBox(0).height = 495
'tBox(1).Left = 35000: tBox(1).Top = Platform(5).Top - 495: tBox(1).width = 1215: tBox(1).height = 495
'tBox(2).Left = 5400: tBox(2).Top = Platform(3).Top - Platform(3).height: tBox(2).width = Def_Height: tBox(2).height = 495
'tBox(3).Left = 9600: tBox(3).Top = 8005: tBox(3).width = 495: tBox(3).height = 495
'tBox(4).Left = Platform(4).Left + Platform(4).width - Def_Height: tBox(4).Top = Platform(4).Top - Platform(4).height: tBox(4).width = Def_Height: tBox(4).height = 495
'tBox(5).Left = Platform(6).Left: tBox(5).height = 900: tBox(5).Top = Platform(6).Top - tBox(5).height: tBox(5).width = 500
'tBox(6).Left = 25000: tBox(6).Top = 8025: tBox(6).width = 1215: tBox(6).height = Def_Height
'tBox(7).Left = 0: tBox(7).Top = 11025: tBox(7).width = 1215: tBox(7).height = Def_Height
'tBox(8).Left = 15000: tBox(8).Top = Platform(4).Top - Platform(4).height: tBox(8).width = 1200: tBox(8).height = Def_Height
'
'
'ubdBoxes = 10
'ReDim Box(ubdBoxes)
''Box
'Box(0).Left = 5587: Box(0).Top = tBox(2).Top - 1095: Box(0).width = 135: Box(0).height = 1095
'Box(1).Left = 6290: Box(1).Top = 7305: Box(1).width = 135: Box(1).height = 1095
'Box(2).Left = tBox(1).Left + 67: Box(2).Top = tBox(1).Top - 1095: Box(2).width = 135: Box(2).height = 1095
'Box(3).Left = tBox(4).Left + 67: Box(3).Top = tBox(4).Top - 1095: Box(3).width = 135: Box(3).height = 1095
'Box(4).Left = 3840: Box(4).Top = 10405: Box(4).width = 135: Box(4).height = 1095
'Box(5).Left = 7200: Box(5).Top = 8775: Box(5).width = Def_Height: Box(5).height = 2505
'Box(6).Left = 28000: Box(6).Top = Platform(2).Top + Platform(2).height - 100: Box(6).width = Def_Height: Box(6).height = Platform(5).Top - Box(6).Top
'Box(7).Left = 7440: Box(7).Top = 11775: Box(7).width = 135: Box(7).height = 1897
'Box(8).Left = Platform(7).Left: Box(8).Top = Platform(7).Top + Platform(7).height: Box(8).width = 495: Box(8).height = Platform(5).Top - Box(8).Top
'Box(9).Left = 15240: Box(9).Top = 8775: Box(9).width = Def_Height: Box(9).height = 2725
'Box(10).Left = 13920: Box(10).Top = 5875: Box(10).width = Def_Height: Box(10).height = 2625
''Box(11).Left = tBox(5).Left + 67: Box(11).Top = tBox(5).Top - 1095: Box(11).width = 135: Box(11).height = 1095
'For i = 0 To ubdBoxes
'    Box(i).bInUse = True
'Next i
'
'HealthPackX = 49200
'HealthPackY = 4800
'######################################################################################################

Public Function RandomRGBColour() As Long

RandomRGBColour = RGB( _
        Int(Rnd() * 256), _
        Int(Rnd() * 256), _
        Int(Rnd() * 256))

End Function

Private Function RandomRGBBetween(btMin As Byte, btMax As Byte) 'rMin As Byte, gMin As Byte, bMin As Byte, _
                                  rMax As Byte, gMax As Byte, bMax As Byte) As Long

Dim btRGB As Byte

btRGB = IntRand(btMin, btMax)

RandomRGBBetween = RGB(btRGB, btRGB, btRGB)

End Function

Public Function AddBot(vWeapon As eWeaponTypes, vTeam As eTeams, Col As Long) As Integer
Dim i As Integer

i = AddStick()

Stick(i).colour = Col
Stick(i).Team = vTeam

Stick(i).Facing = Pi2 * Rnd()

'######################################################################
Stick(i).IsBot = True

Stick(i).Name = GenerateBotName() '"Bot " & NumSticks


If Trim$(Stick(i).Name) = "Bot 4" Then
    SetSticksWeapon i, W1200
    Stick(i).CurrentWeapons(1) = W1200
Else
    'Stick(i).WeaponType = vWeapon
    SetSticksWeapon i, vWeapon
    Stick(i).CurrentWeapons(1) = vWeapon
End If

'Stick(NumSticks).Facing = piD2 'right
'Stick(NumSticks).x = StickGameWidth / 2
Stick(i).X = Rnd() * StickGameWidth
Stick(i).Y = Rnd() * StickGameHeight

Stick(i).Health = Health_Start
Stick(i).Shield = IIf(modStickGame.sv_SpawnWithShields, Max_Shield, 0)

If LCase$(Trim$(Stick(i).Name)) = "agent smith" Then
    Stick(i).colour = vbBlack
Else
    Stick(i).colour = modSpaceGame.RandomRGBColour()
End If
SetAINadeDelay i

Stick(i).AICurrentTarget = -1
'######################################################################

If vWeapon = DEagle Or vWeapon = USP Then
    Stick(i).CurrentWeapons(2) = AK
Else
    Stick(i).CurrentWeapons(2) = USP
End If

If WeaponIsSniper(vWeapon) Then
    Stick(i).Perk = pSniper
End If

AddBot = i

End Function

Public Sub RemoveBot(iBot As Integer)

frmStickGame.SendBroadcast sExits & CStr(Stick(iBot).ID)

SendChatPacketBroadcast "Bot Removed: " & Trim$(Stick(iBot).Name), Stick(iBot).colour
RemoveStick iBot

End Sub

Private Function GenerateBotName() As String
Dim iTestNo As Integer, i As Integer, j As Integer
Dim bNameInUse As Boolean
Dim uNames As Integer: uNames = UBound(kBotNames)

ReDim arTried(0 To uNames) As Boolean


For i = 0 To uNames
    For j = 0 To NumSticksM1
        If Stick(j).IsBot Then
            If Trim$(Stick(j).Name) = kBotNames(i) Then
                arTried(i) = True
                Exit For
            End If
        End If
    Next j
Next i

bNameInUse = True
For i = 0 To uNames
    If arTried(i) = False Then
        bNameInUse = False
    End If
Next i

If Not bNameInUse Then
    Do
        Do
            i = IntRand(0, uNames)
        Loop While arTried(i)
        
        bNameInUse = False
        arTried(i) = True
        
        For j = 1 To NumSticks - 2 'leave out current bot
            If Trim$(Stick(j).Name) = kBotNames(i) Then
                bNameInUse = True
                Exit For
            End If
        Next j
    Loop While bNameInUse And i < uNames
End If


If bNameInUse = False Then
    GenerateBotName = kBotNames(i)
Else
    i = 1
    For i = 0 To NumSticksM1
        If Stick(i).IsBot Then
            j = val(Right$(Stick(i).Name, Len(Stick(i).Name) - InStr(1, Stick(i).Name, vbSpace, vbTextCompare)))
    
            If j = iTestNo Then
                iTestNo = iTestNo + 1
                i = 0 'will = 1 after "Next i" executed
            End If
        End If
    Next i
    
    GenerateBotName = "Bot " & CStr(iTestNo)
End If

End Function
Private Function GenerateStickID() As Integer
Dim iTestNo As Integer, i As Integer

iTestNo = 0 'attempt to assign ID 0
For i = 0 To NumSticks - 1 'NumSticksM1 isn't set yet
    If iTestNo = Stick(i).ID Then
        iTestNo = iTestNo + 1
        i = -1 'will = 0 after "Next i" executed
    End If
Next i

GenerateStickID = iTestNo

End Function
Private Function GenerateMineID() As Integer
Dim iTestNo As Integer, i As Integer

iTestNo = 0 'attempt to assign ID 0
For i = 0 To NumMines - 1
    If iTestNo = Mine(i).ID Then
        iTestNo = iTestNo + 1
        i = -1 'will = 0 after "Next i" executed
    End If
Next i

GenerateMineID = iTestNo

End Function
Private Function GenerateBarrelID()
Dim iTestNo As Integer, i As Integer

iTestNo = 0 'attempt to assign ID 0
For i = 0 To NumBarrels - 1
    If iTestNo = Barrel(i).ID Then
        iTestNo = iTestNo + 1
        i = -1 'will = 0 after "Next i" executed
    End If
Next i

GenerateBarrelID = iTestNo

End Function

Public Function AddStick() As Integer

'Add a Stick onto the array, and return his index
ReDim Preserve Stick(NumSticks)
'ReDim Preserve ScoreList(NumSticks)

'ScoreList(NumSticks).ID = Stick(NumSticks).ID
Stick(NumSticks).LastPacket = GetTickCount()
Stick(NumSticks).LegWidth = 50

Stick(NumSticks).Name = vbNullString
Stick(NumSticks).bAlive = True
Stick(NumSticks).LastPacket = GetTickCount() + mPacket_SEND_DELAY * 4
Stick(NumSticks).LastBullet = GetTickCount() - 10000

ResetTimeLong Stick(NumSticks).LastGravity, Gravity_Delay

'Stick(NumSticks).JumpStartY = StickGameHeight + 100
'ResetJumpStartY NumSticks
Stick(NumSticks).sgTimeZone = 1

Stick(NumSticks).ID = GenerateStickID() 'Stick(NumSticksM1).ID + 1

ResetStickFireAndFlash NumSticks

AddStick = NumSticks
NumSticksM1 = NumSticks

AdjustIDArray

NumSticks = NumSticks + 1

End Function

Public Function FindStick(ID As Integer) As Integer

On Error GoTo EH
FindStick = StickIndexIDMap(ID)
Exit Function
EH:
FindStick = -1

'Dim i As Integer
'
'For i = 0 To NumSticksM1
'    If Stick(i).ID = ID Then
'        FindStick = i
'        Exit Function
'    End If
'Next i
'
'FindStick = -1

End Function

Private Sub AdjustIDArray()
Dim MaxID As Integer, i As Integer

MaxID = -1
For i = 0 To NumSticksM1
    If Stick(i).ID > MaxID Then
        MaxID = Stick(i).ID
    End If
Next i

ReDim Preserve StickIndexIDMap(0 To MaxID)

For i = 0 To MaxID
    StickIndexIDMap(i) = -1
Next i

For i = 0 To NumSticksM1
    StickIndexIDMap(Stick(i).ID) = i
Next i

End Sub

Public Sub RemoveStick(Index As Integer)
Dim i As Integer


If Index > 0 Then 'not removing local stick
    If Stick(0).Perk = pSpy Then
        'if we're a spy...
        If Stick(0).MaskID = Stick(Index).ID Then
            'we were masquerading as the removed stick
            AddMainMessage "Target Stick has left the game (Spy Perk)", False
            Stick(0).MaskID = Stick(0).ID
        End If
    End If
End If


i = 0
Do While i < NumBullets
    If Bullet(i).OwnerIndex = Index Then
        RemoveBullet i, False, False
        i = i - 1
    End If
    i = i + 1
Loop
'reset all the other bullet's indexs
For i = 0 To NumBullets - 1
    If Bullet(i).OwnerIndex > Index Then
        Bullet(i).OwnerIndex = Bullet(i).OwnerIndex - 1
    End If
Next i


i = 0
Do While i < NumNades
    If Nade(i).OwnerID = Stick(Index).ID Then
        RemoveNade i, False
        i = i - 1
    End If
    i = i + 1
Loop
'don't need to reset all the other nades's stick-ids


i = 0
Do While i < NumMines
    If Mine(i).OwnerID = Stick(Index).ID Then
        RemoveMine i
        i = i - 1
    End If
    i = i + 1
Loop


On Error Resume Next
'Remove this Stick from the array
For i = Index To NumSticks - 2
    Stick(i) = Stick(i + 1)
Next i

'Resize the array
ReDim Preserve Stick(NumSticks - 2)
NumSticks = NumSticksM1
NumSticksM1 = NumSticks - 1

AdjustIDArray

End Sub

Private Sub ProcessCameraMovement()

Static xMoveConst As Long ', yMoveConst As Long
Static FacingWasGreaterThanPi As Boolean ', FacingWasGreaterThanPiD2 As Boolean
Const xMaxMovement = 2500 ', yMaxMovement = 1000
Const xMoveInc = xMaxMovement / 10 ', yMoveInc = yMaxMovement / 10
Dim bMovingCam As Boolean

'x smooth motion
bMovingCam = xMoveConst < xMaxMovement
If bMovingCam Then xMoveConst = xMoveConst + xMoveInc * modStickGame.sv_StickGameSpeed

If Stick(0).ActualFacing > Pi Then 'MouseX < Me.width / 2 Then
    If FacingWasGreaterThanPi = False Then
        If bMovingCam = False Then
            xMoveConst = -xMaxMovement
            FacingWasGreaterThanPi = True
        End If
    End If
    
    MoveCameraX Stick(0).X * cg_sZoom - StickCentreX - xMoveConst
    
Else
    If FacingWasGreaterThanPi Then
        If bMovingCam = False Then
            xMoveConst = -xMaxMovement
            FacingWasGreaterThanPi = False
        End If
    End If
    
    MoveCameraX Stick(0).X * cg_sZoom - StickCentreX + xMoveConst
End If

''#########################################################################
''y smooth motion
'
'If yMoveConst < yMaxMovement Then yMoveConst = yMoveConst + yMoveInc
'
'If Stick(0).ActualFacing > pi3D4 And Stick(0).ActualFacing < 5 * pi / 4 Then
'    If FacingWasGreaterThanPiD2 = False Then
'        yMoveConst = -yMaxMovement
'        FacingWasGreaterThanPiD2 = True
'    End If
'
'    MoveCameraY Stick(0).y * cg_sZoom - StickCentreY + yMoveConst
'
'Else
'    If FacingWasGreaterThanPiD2 Then
'        yMoveConst = -yMaxMovement
'        FacingWasGreaterThanPiD2 = False
'    End If
'
'    MoveCameraY Stick(0).y * cg_sZoom - StickCentreY - yMoveConst
'End If

MoveCameraY Stick(0).Y * cg_sZoom - StickCentreY

End Sub

Private Sub ProcessCamera()
Dim bNorm As Boolean
Const Flash_Fluctuation As Single = 50

If StickInGame(0) And bPlaying Then
    
    If Stick(0).bFlashed Then
        'MoveCameraX Stick(0).X * cg_sZoom - StickCentreX + fX
        'MoveCameraY Stick(0).Y * cg_sZoom - StickCentreY + fY
        
        CentreCameraOnPoint Stick(0).X + PM_Rnd() * Flash_Fluctuation, Stick(0).Y + PM_Rnd() * Flash_Fluctuation
        
    Else
        bNorm = True
        If WeaponIsSniper(Stick(0).WeaponType) Then
            If modStickGame.cl_SniperScope Then
                If StickiHasState(0, STICK_CROUCH) Then
                    bNorm = False
                ElseIf StickiHasState(0, STICK_PRONE) Then
                    bNorm = False
                Else
                    bNorm = Not (Stick(0).Perk = pFocus)
                End If
'            Else
'                bNorm = True
            End If
        ElseIf Stick(0).WeaponType = G3 Then
            If Stick(0).Perk = pSniper Then
                If StickiHasState(0, STICK_CROUCH) Then
                    bNorm = Not modStickGame.cl_SniperScope
                ElseIf StickiHasState(0, STICK_PRONE) Then
                    bNorm = Not modStickGame.cl_SniperScope
'                Else
'                    bNorm = True
                End If
            Else
                bNorm = Not Stick(0).Perk = pFocus
            End If
        ElseIf Stick(0).WeaponType = Chopper Then 'THIS MUST BE BEFORE PERK, SINCE IF IT'S A CHOPPER, IT NEEDS TESTAN
            bNorm = False 'Not modStickGame.cl_SniperScope
        ElseIf Stick(0).Perk = pFocus Then
            bNorm = False
'        Else
'            bNorm = True
        End If
        
        
        If bNorm Then
            If modStickGame.cg_AutoCamera Then
                ProcessCameraMovement
            Else
                'MoveCameraX Stick(0).X * cg_sZoom - StickCentreX
                'MoveCameraY Stick(0).Y * cg_sZoom - StickCentreY
                CentreCameraOnPoint Stick(0).X, Stick(0).Y
            End If
        Else
            ProcessSniperCamera
        End If
        
        
    End If
End If

End Sub

Private Sub ProcessSniperCamera()
Dim DistToCursor As Single, Angle As Single

DistToCursor = GetDist(StickCentreX, StickCentreY, MouseX, MouseY)
Angle = FindAngle(StickCentreX, StickCentreY, MouseX, MouseY)

If Stick(0).Perk = pFocus Then
    If WeaponIsSniper(Stick(0).WeaponType) Then DistToCursor = DistToCursor * 1.5
End If

MoveCameraX (Stick(0).X * cg_sZoom - StickCentreX) + DistToCursor * Sine(Angle)
MoveCameraY (Stick(0).Y * cg_sZoom - StickCentreY) - DistToCursor * CoSine(Angle)
'CentreCameraOnPoint (Stick(0).X * cg_sZoom) + DistToCursor * Sine(Angle),
'^ doesn't work right, because of shiz

End Sub

Private Sub DisplaySticks()

Dim i As Integer
Dim Txt As String
Const Invul_Radius = BodyLen * 2


picMain.FillStyle = vbFSTransparent 'transparent


For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        
        Me.picMain.DrawWidth = 2
        
        
        If CanSeeStick(i) Then
            
            If modStickGame.cg_HolsteredWeap Then 'And modStickGame.sv_2Weapons Then
                If Not StickiHasState(i, STICK_PRONE) Then
                    DrawHolsteredWeapon i
                End If
            End If
            
            
            Me.picMain.DrawWidth = 3
            
            If Stick(i).WeaponType <> Chopper Then
                'If Stick(i).LastSpawnTime + Spawn_Invul_Time / GetTimeZoneAdjust > GetTickCount() Then
                If StickInvul(i) Then
                    modStickGame.sCircle Stick(i).X, Stick(i).Y + 250, Invul_Radius, Stick(i).colour
                End If
            End If
            
            
            'PrintStickText "In Smoke: " & StickInSmoke(i), Stick(0).X, Stick(0).Y - 2000, vbBlue
            'PrintStickText "In tBox: " & StickIntBox(i), Stick(0).X, Stick(0).Y - 2000, vbBlue
            
        End If
        
        DrawStick i
    End If
Next i


End Sub

Private Sub DrawHolsteredWeapon(i As Integer)

Dim iWeapToDraw As eWeaponTypes
Dim sDir As Single, sX As Single, sY As Single, kY As Single
Dim bFacingGreaterThanPi As Boolean

Const DistFromBody = ArmLen / 2, DistFromBodyX2 = DistFromBody * 2, DistFromBodyD4 = DistFromBody / 4
Const DistFromBodyD2 = DistFromBody / 2, HeadRadiusD2 = HeadRadius \ 2, DistFromBodyX1p5 = DistFromBody * 1.5
Const bReloading As Boolean = True
Const M82_Inc = BodyLen + ArmLen, AWM_Inc = M82_Inc / 2


If Stick(i).WeaponType = Chopper Then Exit Sub
If Stick(i).Perk = pZombie Then Exit Sub


If Stick(i).WeaponType = Stick(i).CurrentWeapons(1) Then
    'draw #2
    iWeapToDraw = Stick(i).CurrentWeapons(2)
Else
    'draw #1
    iWeapToDraw = Stick(i).CurrentWeapons(1)
End If


Stick(i).Facing = FixAngle(Stick(i).Facing)
bFacingGreaterThanPi = Stick(i).Facing > Pi

If bFacingGreaterThanPi Then
    sDir = Pi
    sX = Stick(i).X + DistFromBody
    kY = -1
Else
    sDir = Pi
    sX = Stick(i).X - DistFromBody
    kY = 1
End If

sY = GetStickY(i)



Select Case iWeapToDraw
    Case W1200
        If bFacingGreaterThanPi Then
            sX = Stick(i).X + DistFromBodyX2
        Else
            sX = Stick(i).X - DistFromBodyX2
        End If
        
        DrawW12002 0, sX, sY + M82_Inc, 0, 0, kY, False, sX, sY, i
        
        
    Case G3
        If bFacingGreaterThanPi Then
            sX = Stick(i).X + DistFromBodyX1p5
        Else
            sX = Stick(i).X - DistFromBodyX1p5
        End If
        
        DrawG32 Pi, sX, sY + HeadRadiusD2, 0, 0, kY, False, sX, sY, i, False
        
        
    Case SPAS
        
        DrawSPAS2 Pi, sX, sY + ArmLen, 0, 0, kY, False, sX, sY, i
        
        
    Case AK
        'If bFacingGreaterThanPi Then
            'sX = Stick(i).X + DistFromBodyX1p5
        'Else
            'sX = Stick(i).X - DistFromBodyX1p5
        'End If
        
        DrawAK2 sDir, sX, sY + ArmLen, 0, 0, kY, False, sX, sY + ArmLen, i, bReloading
        
        
    Case M82
        If bFacingGreaterThanPi Then
            sX = Stick(i).X + DistFromBodyD2
        Else
            sX = Stick(i).X - DistFromBodyD2
        End If
        
        DrawM822 0, sX, sY + M82_Inc, 0, 0, kY, False, sX, sY, i, bReloading, 1, 0, False
        sDir = 0
        
    Case AWM
        If bFacingGreaterThanPi Then
            sX = Stick(i).X + DistFromBodyD2
        Else
            sX = Stick(i).X - DistFromBodyD2
        End If
        
        DrawAWM2 0, sX, sY + AWM_Inc, 0, 0, kY, False, sX, sY, i
        sDir = 0
        
    Case MP5
        
        DrawMP52 sDir, sX, sY + ArmLen, 0, 0, kY, False, sX, sY, i, bReloading
        
    Case DEagle
        sX = Stick(i).X
        DrawDEagle2 sDir, sX, sY + BodyLen, 0, 0, kY, False, sX, sY, i, ArmLen
        
    Case USP
        sX = Stick(i).X
        DrawUSP2 sDir, sX, sY + BodyLen, 0, 0, kY, False, sX, sY, i, ArmLen
        
    Case XM8
        DrawXM82 sDir, sX, sY + DistFromBody, 0, 0, kY, False, sX, sY, i, False
        
        
    Case AUG
        DrawAUG2 sDir, sX, sY + ArmLen, 0, 0, kY, False, sX, sY, i, bReloading
        
        
    Case RPG
        If bFacingGreaterThanPi Then
            sX = Stick(i).X + DistFromBodyD4
        Else
            sX = Stick(i).X - DistFromBodyD4
        End If
        DrawRPG2 sDir, sX, sY + DistFromBody, 0, 0, IIf(kY = 1, -1, 1), False, sX, sY, i, True
        
        
    Case M249
        DrawM2492 sDir, sX, sY + DistFromBody, 0, 0, kY, False, sX, sY, i
        
        
    Case FlameThrower
        DrawFlamethrower2 sDir, sX, sY, 0, 0, kY, False, sX, sY, i, bReloading
        
    Case Mac10
        If bFacingGreaterThanPi Then
            sX = Stick(i).X + DistFromBodyD4
        Else
            sX = Stick(i).X - DistFromBodyD4
        End If
        DrawMac102 sDir, sX, sY + AWM_Inc, 0, 0, kY, False, sX, sY, i, bReloading
        
End Select


If b2ndWeaponSilenced Then 'Stick(i).bSilenced Then
    If WeaponSilencable(GetSticksSecondWeapon(i)) Then
        DrawSticksSilencer i, sDir
    End If
End If

End Sub

Private Sub DrawSticksSilencer(i As Integer, sAngle As Single)
DrawSilencer Stick(i).GunPoint.X, Stick(i).GunPoint.Y, sAngle
End Sub

Private Function CanSeeStick(i As Integer) As Boolean
Const Peripheral_Vision = piD6 '[b]piD3[/b] / 2

Dim Theta As Single
Dim f As Single


If i = 0 Then
    CanSeeStick = True
ElseIf modStickGame.sv_Hardcore = False Then
    CanSeeStick = True
ElseIf StickInGame(0) = False Then
    CanSeeStick = True
ElseIf Stick(i).WeaponType = Chopper Then
    CanSeeStick = True
Else
    
    f = FixAngle(Stick(0).ActualFacing)
    Theta = FixAngle(FindAngle(Stick(0).X, Stick(0).Y, Stick(i).X, GetStickY(i) + 1))
    
    
    'stick(0).ActualFacing -abit < Theta < Stick(0).ActualFacing +abit
    If (f - Peripheral_Vision) < Theta Then
        If Theta < (f + Peripheral_Vision) Then
            CanSeeStick = True
        End If
    End If
    
End If


End Function

Private Function StickCanSeeStick(iSource As Integer, iTarget As Integer) As Boolean
Const Peripheral_Vision = piD6 '[b]piD3[/b] / 2

Dim Theta As Single
Dim f As Single


If iTarget = iSource Then
    StickCanSeeStick = True
ElseIf modStickGame.sv_Hardcore = False Then
    StickCanSeeStick = True
ElseIf StickInGame(iTarget) = False Then
    StickCanSeeStick = False
ElseIf Stick(iTarget).WeaponType = Chopper Then
    StickCanSeeStick = True
ElseIf Stick(iSource).WeaponType = Chopper Then
    StickCanSeeStick = True
Else
    
    f = FixAngle(Stick(iSource).ActualFacing)
    Theta = FixAngle(FindAngle_Actual(Stick(iSource).X, Stick(iSource).Y, Stick(iTarget).X, GetStickY(iTarget)))
    
    
    'stick(0).ActualFacing -abit < Theta < Stick(0).ActualFacing +abit
    If (f - Peripheral_Vision) < Theta Then
        If Theta < (f + Peripheral_Vision) Then
            StickCanSeeStick = True
        End If
    End If
    
End If


End Function

Private Sub DrawGhillie_Norm(iStick As Integer, Y As Single, kY As Single)
Const ArmLenD2 = ArmLen / 2, _
      HeadRadiusD2 = HeadRadius / 2, _
      HeadRadiusX1p3 = HeadRadius * 1.3, _
      armlenx2 = ArmLen * 2, _
      HeadRadiusX2 = HeadRadius * 2, Small_Length = ArmLen / 5


'#######################################################
'head
DrawGhillie_Head iStick, Y, kY

With Stick(iStick)
    
    '(.X, Y) = top of head
    
    '#######################################################
    'body
    modStickGame.sLine .X, Y + HeadRadius, _
                       .X + HeadRadiusD2 * kY, Y + HeadRadius + BodyLen
    
    
    
    '#######################################################
    'legs
    modStickGame.sLine .X - HeadRadiusD2 * kY, Y + BodyLen, _
                       .X + ArmLenD2, Y + BodyLen + HeadRadius
    
    modStickGame.sLine .X + ArmLenD2, Y + BodyLen + HeadRadiusD2, _
                       .X - ArmLenD2, Y + ArmLen
    
    '#######################################################
    'arms/rifle
    
    DrawGhillie_Rifle iStick, Y, kY
    
End With

End Sub
Private Sub DrawGhillie_Rifle(iStick As Integer, Y As Single, kY As Single)
Const HeadRadiusX2 = HeadRadius * 2, _
      ArmLenD2 = ArmLen / 2

With Stick(iStick)
    
    If .WeaponType = Knife Then Exit Sub
    
    modStickGame.sLine .X, Y + HeadRadiusX2, _
        (.GunPoint.X + .X) / 2, (.GunPoint.Y + Y) / 2
    
    modStickGame.sLine .X, Y + HeadRadiusX2, _
        .CasingPoint.X - ArmLenD2 * Sine(.Facing), .CasingPoint.Y + ArmLenD2 * CoSine(.Facing)
    
    modStickGame.sLine .X, Y + HeadRadiusX2, _
        .CasingPoint.X + ArmLenD2 * Sine(.Facing), .CasingPoint.Y - ArmLenD2 * CoSine(.Facing)
    
    
    modStickGame.sLine .X, Y + HeadRadiusX2, _
        .GunPoint.X - ArmLenD2 * Sine(.Facing), .GunPoint.Y + ArmLenD2 * CoSine(.Facing)
    
End With

End Sub
Private Sub DrawGhillie_Head(iStick As Integer, Y As Single, kY As Single)
Const HeadRadiusD2 = HeadRadius / 2, ArmLenD2 = ArmLen / 2

With Stick(iStick)
    modStickGame.sLine .X - HeadRadius * kY, Y - HeadRadiusD2, _
                       .X, Y + HeadRadius
    
    modStickGame.sLine .X - HeadRadius, Y + HeadRadiusD2, _
                       .X + ArmLenD2 * kY, Y
    
    
End With

End Sub
Private Sub DrawGhillie_Prone(iStick As Integer, Y As Single, kY As Single)

DrawGhillie_Head iStick, Y, kY
DrawGhillie_Rifle iStick, Y, kY

End Sub

Private Sub DrawStick(i As Integer)

Dim Crouching As Boolean, Prone As Boolean, bGhillie As Boolean, bZombie As Boolean
Dim X As Single, Y As Single, tX As Single, tY As Single 'stick's co-ords
Dim XComp As Single 'for leg width
Const LegWidthK As Single = 8 'leg width speed
Const HRx1p8 = HeadRadius * 1.8, FlashEffect_Radius = HeadRadius / 1.5
'Const ArmourTop = HeadRadius * 2 + 50
Const HeadRadiusX2 = HeadRadius * 2
Dim Hand1X As Single, Hand1Y As Single 'hand co-ords
Dim Hand2X As Single, Hand2Y As Single
Dim ShoulderY As Single, Adj As Single
Dim A_Facing As Single
Dim GunY As Single
Dim TeamCol As Long, lCol As Long
Dim j As Integer 'for spy mask
Dim bCanSee As Boolean
'zombie consts
Const armlenxy As Single = ArmLen * 1.4

'arm stuff
Dim ThrowTime As Long, GTC As Long, sAngle As Single
Const Throwing_ArmLen = ArmLen * 1.5

If Stick(i).Perk = pSpy Then
    j = FindStick(Stick(i).MaskID)
    If j = -1 Then j = 0
Else
    j = i
End If


If Stick(i).WeaponType = Chopper Then
    DrawChopper i
Else
    Crouching = StickiHasState(i, STICK_CROUCH)
    Prone = StickiHasState(i, STICK_PRONE)
    Stick(i).Facing = FixAngle(Stick(i).Facing)
    'bArmoured = (Stick(i).Armour > 0)
    bZombie = (Stick(i).Perk = pZombie)
    
    If Stick(i).Perk = pSniper Then
        bGhillie = True
        lCol = Grass_Col
    ElseIf bZombie Then
        bGhillie = True
        lCol = Zombie_Col
    Else
        lCol = Stick(j).colour
    End If
    
    '###################################################
    'Find X and Y
    X = Stick(i).X
    Y = GetStickY(i)
    
    If Prone Then
        GunY = Y - 50
    Else
        GunY = Y
    End If
    'If Crouching Then
        'Y = Stick(i).Y + BodyLen / 2
        'GunY = Y
    'ElseIf Prone Then
        'Y = Stick(i).Y + BodyLen * 1.2
        'GunY = Y - 50
    'Else
        'Y = Stick(i).Y
        'GunY = Y
    'End If
    
    '###################################################
    
    'Draw Head + Body
    bCanSee = CanSeeStick(i)
    If bCanSee Then
        picMain.DrawWidth = 2
        'Col = IIf(Stick(i).Armour > 0, Armour_Colour, Stick(i).Colour)
        
        '##########################################################################################
        'head here
        Me.picMain.FillStyle = vbSolid
        If Stick(j).Team > Neutral And Not bGhillie Then
            TeamCol = GetTeamColour(Stick(j).Team)
            Me.picMain.FillColor = TeamCol
            modStickGame.sCircle X, Y + HeadRadius, HeadRadius, Stick(j).colour 'head
        Else
            Me.picMain.FillColor = lCol
            modStickGame.sCircle X, Y + HeadRadius, HeadRadius, lCol 'head
        End If
        Me.picMain.FillStyle = vbFSTransparent
        '##########################################################################################
        
        'If bArmoured Then
            'modStickGame.sCircleSE X, Y + HeadRadius, HeadRadius, Armour_Colour, -(Stick(i).Facing - piD2), Stick(i).Facing - pi3D2
            'modStickGame.sCircleSE X, Y + HeadRadius, HeadRadius, Armour_Colour, -0.01, -Pi
            'modStickGame.sCircleSE X, Y + HeadRadius, HeadRadius, Armour_Colour, piD2, pi3D2
        'End If
        
        
        'body
        If Prone Then
            
            'If bArmoured Then
                'picMain.DrawWidth = 3
                'picMain.ForeColor = Armour_Colour
                'modStickGame.sLine X + IIf(Stick(i).Facing > Pi, HeadRadius, -HeadRadius), Y + HeadRadius, _
                               X + IIf(Stick(i).Facing > Pi, BodyLen, -BodyLen), Y + HRx1p8
                'picMain.DrawWidth = 2
            'Else
                picMain.ForeColor = lCol
                modStickGame.sLine X + IIf(Stick(i).Facing > Pi, HeadRadius, -HeadRadius), Y + HeadRadius, _
                               X + IIf(Stick(i).Facing > Pi, BodyLen, -BodyLen), Y + HRx1p8
            'End If
            
        Else
            'draw legs
            DrawLegs i, X, Y, Crouching, Prone, lCol
            
            If Stick(i).Shield Then
                Dim r As Byte
                r = Stick(i).Shield / Max_Shield
                picMain.ForeColor = RGB(0, 150 + 100 * r, 100) 'lCol Or vbGreen
            Else
                picMain.ForeColor = lCol
            End If
            
            modStickGame.sLine X, Y + HeadRadiusX2, X, Y + BodyLen 'body
            
            'If bArmoured Then
                'modStickGame.sBoxFilled X - 10, Y + ArmourTop, X + 10, Y + BodyLen, Armour_Colour
            'End If
        End If
        
        
        If Stick(i).bTyping Then
            If Not (StickiHasState(i, STICK_PRONE) Or StickiHasState(i, STICK_CROUCH)) Then
                DrawTypeBubble Stick(i).X, Stick(i).Y
            End If
        End If
    End If
    
    
    '###################################################
    
    
    
    If Stick(i).Perk <> pZombie Then
        'MUST DRAW THE WEAPON TO GET UPDATED GUNPOINT CO-ORDS
        Select Case Stick(i).WeaponType
            Case W1200
                DrawW1200 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
            Case Mac10
                DrawMac10 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case SPAS
                DrawSPAS i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case AK
                DrawAK i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case M82
                DrawM82 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, lCol, A_Facing
            Case AWM
                DrawAWM i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case DEagle
                DrawDEagle i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
            Case XM8
                DrawXM8 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case AUG
                DrawAUG i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case MP5
                DrawMP5 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case RPG
                DrawRPG i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
            Case M249
                DrawM249 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
            Case FlameThrower
                DrawFlameThrower i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
            Case USP
                DrawUSP i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case G3
                DrawG3 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, A_Facing
            Case Else
                DrawKnife i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
        End Select
    Else
        Stick(i).GunPoint.X = Stick(i).X
        Stick(i).GunPoint.Y = Stick(i).Y
        If Stick(i).Facing > Pi Then
            'left
            Hand1X = Stick(i).X - armlenxy
            Hand1Y = Stick(i).Y + 250
            Hand2X = Stick(i).X - armlenxy
            Hand2Y = Stick(i).Y + 100
        Else
            'right
            Hand1X = Stick(i).X + armlenxy
            Hand1Y = Stick(i).Y + 250
            Hand2X = Stick(i).X + armlenxy
            Hand2Y = Stick(i).Y + 100
        End If
    End If
    
    
    
    
    If bCanSee Then
        picMain.DrawWidth = 2
        
        If Stick(i).bSilenced Then
            DrawSticksSilencer i, IIf(Stick(i).Facing > Pi, A_Facing + Pi, A_Facing)
            'needs A_Facing to be passed to the Draw_Weapon procedure
        End If
        
        
        If bGhillie Then
            picMain.ForeColor = lCol 'Grass_Col
            If Prone Then
                DrawGhillie_Prone i, Y, IIf(Stick(i).Facing > Pi, 1, -1)
            Else
                DrawGhillie_Norm i, Y, IIf(Stick(i).Facing > Pi, 1, -1)
            End If
        End If
        
        
        'picMain.DrawWidth = 2
        picMain.ForeColor = lCol 'Stick(j).Colour
        ShoulderY = Y + BodyLen / 2
        
        
        GTC = GetTickCount()
        ThrowTime = GTC - Stick(i).LastNade
        Adj = GetSticksTimeZone(i)
        
        
        If ThrowTime < Nade_Arm_Time / Adj Then
            
            sAngle = piD2 * ThrowTime / (Nade_Arm_Time / Adj)
            
            
            modStickGame.sLine X, ShoulderY, _
                               X + Throwing_ArmLen * Sine(sAngle) * IIf(Stick(i).Facing > Pi, -1, 1), _
                               ShoulderY - Throwing_ArmLen * CoSine(sAngle)
            
            
            
            modStickGame.sLine X, ShoulderY, Hand1X, Hand1Y 'arm1
            
        Else
            
            'If Stick(i).WeaponType <> AUG Then
                modStickGame.sLine X, ShoulderY, Hand1X, Hand1Y 'arm1
                modStickGame.sLine X, ShoulderY, Hand2X, Hand2Y 'arm2
            'Else
                'modStickGame.sLine X, ShoulderY, Hand1X, Hand1Y 'arm1
            'End If
            
        End If
        
        
        If Stick(i).bFlashed Then
            modStickGame.sCircle X + PM_Rnd() * HeadRadius, Y + HeadRadius * (1 + PM_Rnd()), FlashEffect_Radius, vbYellow
        End If
        
        
        '##########################
        'move his legs
        If Stick(i).bOnSurface Then
            If StickIsMoving(i) Then
                If Stick(i).LegWidth > MaxLegWidth Then
                    Stick(i).LegBigger = False
                ElseIf Stick(i).LegWidth < -MaxLegWidth Then
                    Stick(i).LegBigger = True
                End If
                
                XComp = Abs(Stick(i).Speed * Sine(Stick(i).Heading))
                
                On Error GoTo LegEH
                If Stick(i).LegBigger Then
                    Stick(i).LegWidth = Stick(i).LegWidth + XComp * Adj / LegWidthK
                Else
                    Stick(i).LegWidth = Stick(i).LegWidth - XComp * Adj / LegWidthK
                End If
            End If
        End If
        
        If Abs(Stick(i).LegWidth) > MaxLegWidth Then Stick(i).LegWidth = MaxLegWidth
        '##########################
        
        
    End If
End If

Exit Sub
LegEH:
Stick(i).LegWidth = 0
End Sub

Private Sub DrawChopper(iStick As Integer)
Dim pt(1 To 11) As PointAPI, ScreenPt(1 To 3) As PointAPI
Dim WheelPtX As Single, WheelPtY As Single
Dim WheelConnectionX As Single, WheelConnectionY As Single

Dim GunPtX As Single, GunPtY As Single
Dim GunTipPtX As Single, GunTipPtY As Single

Dim Rotor1X As Single, Rotor2X As Single, Rotor1Y As Single, Rotor2Y As Single
Dim RotorX As Single, RotorY As Single

Dim TailRotorX As Single, TailRotorY As Single
Dim TailRotor1X As Single, TailRotor2X As Single, TailRotor1Y As Single, TailRotor2Y As Single

Dim Facing As Single
Dim kY As Integer

Const RotorInc = 400
Const TailRotorInc = 0.4
Const TailRotorLen = 200
Const Pi2 = Pi * 2
Const GunLen = 200


If StickiHasState(iStick, STICK_LEFT) Then

'    If Stick(iStick).ChopperFacingAmount < piD3 Then
'        Stick(iStick).ChopperFacingAmount = Stick(iStick).ChopperFacingAmount + 0.05
'    Else
'        Stick(iStick).ChopperFacingAmount = piD3
'    End If
    
    Facing = pi5D12 'Stick(iStick).ChopperFacingAmount
    
    
ElseIf StickiHasState(iStick, STICK_RIGHT) Then
'    If Stick(iStick).ChopperFacingAmount < pi2d3 Then
'        Stick(iStick).ChopperFacingAmount = Stick(iStick).ChopperFacingAmount - 0.05
'    End If
    
    Facing = pi7D12 'Stick(iStick).ChopperFacingAmount
    
Else
    Facing = piD2
End If


pt(1).X = Stick(iStick).X
pt(1).Y = Stick(iStick).Y

pt(2).X = pt(1).X + CLD6 * Sine(Facing + piD6)
pt(2).Y = pt(1).Y - CLD6 * CoSine(Facing + piD6)

pt(3).X = pt(2).X + CLD10 * Sine(Facing + piD3)
pt(3).Y = pt(2).Y - CLD10 * CoSine(Facing + piD3)

pt(4).X = pt(3).X + CLD2 * Sine(Facing - Pi)
pt(4).Y = pt(3).Y - CLD2 * CoSine(Facing - Pi)

pt(5).X = pt(4).X + CLD10 * Sine(Facing - pi3D4)
pt(5).Y = pt(4).Y - CLD10 * CoSine(Facing - pi3D4)

pt(6).X = pt(5).X + CLD3 * Sine(Facing - Pi)
pt(6).Y = pt(5).Y - CLD3 * CoSine(Facing - Pi)

pt(7).X = pt(6).X + CLD6 * Sine(Facing - pi3D4)
pt(7).Y = pt(6).Y - CLD6 * CoSine(Facing - pi3D4)

pt(8).X = pt(7).X + CLD8 * Sine(Facing)
pt(8).Y = pt(7).Y - CLD8 * CoSine(Facing)

pt(9).X = pt(8).X + CLD8 * Sine(Facing + piD4)
pt(9).Y = pt(8).Y - CLD8 * CoSine(Facing + piD4)

pt(10).X = pt(9).X + CLD6 * Sine(Facing)
pt(10).Y = pt(9).Y - CLD6 * CoSine(Facing)

pt(11).X = pt(1).X + CLD8 * Sine(Facing - Pi)
pt(11).Y = pt(1).Y - CLD8 * CoSine(Facing - Pi)



ScreenPt(1).X = pt(1).X + Sine(Facing + piD2) * 50
ScreenPt(1).Y = pt(1).Y - CoSine(Facing + piD2) * 50

ScreenPt(2).X = pt(2).X + Sine(Facing - Pi) * 50
ScreenPt(2).Y = pt(2).Y - CoSine(Facing - Pi) * 50

ScreenPt(3).X = ScreenPt(2).X - CLD6 * Sine(Facing)
ScreenPt(3).Y = ScreenPt(2).Y + CLD6 * CoSine(Facing)



WheelPtX = pt(3).X + CLD6 * Sine(Facing + pi8D9)
WheelPtY = pt(3).Y - CLD6 * CoSine(Facing + pi8D9)

WheelConnectionX = WheelPtX + 250 * Sine(Facing - pi13D18)
WheelConnectionY = WheelPtY - 250 * CoSine(Facing - pi13D18)

GunPtX = pt(4).X + CLD6 * Sine(Facing + piD10)
GunPtY = pt(4).Y - CLD6 * CoSine(Facing + piD10)
GunTipPtX = GunPtX + GunLen * Sine(Stick(iStick).ActualFacing)
GunTipPtY = GunPtY - GunLen * CoSine(Stick(iStick).ActualFacing)


RotorX = pt(1).X + 280 * Sine(Facing - pi3D4)
RotorY = pt(1).Y - 280 * CoSine(Facing - pi3D4)

Rotor1X = RotorX + Stick(iStick).RotorWidth * Sine(Facing)
Rotor1Y = RotorY - Stick(iStick).RotorWidth * CoSine(Facing)

Rotor2X = RotorX - Stick(iStick).RotorWidth * Sine(Facing)
Rotor2Y = RotorY + Stick(iStick).RotorWidth * CoSine(Facing)



TailRotorX = pt(6).X + 350 * Sine(Facing - pi5D9)
TailRotorY = pt(6).Y - 350 * CoSine(Facing - pi5D9)

TailRotor1X = TailRotorX + TailRotorLen * Sine(Stick(iStick).TailRotorFacing)
TailRotor1Y = TailRotorY - TailRotorLen * CoSine(Stick(iStick).TailRotorFacing)

TailRotor2X = TailRotorX - TailRotorLen * Sine(Stick(iStick).TailRotorFacing)
TailRotor2Y = TailRotorY + TailRotorLen * CoSine(Stick(iStick).TailRotorFacing)


Stick(iStick).GunPoint.X = GunTipPtX
Stick(iStick).GunPoint.Y = GunTipPtY

Stick(iStick).CasingPoint.X = GunPtX 'Stick(iStick).GunPoint.X
Stick(iStick).CasingPoint.Y = GunPtY 'Stick(iStick).GunPoint.Y


'MUST BE DONE BEFORE POINTS ARE SCALED INTO PIXELS
'############################## SILVER ############################################
picMain.ForeColor = MSilver
modStickGame.sLine WheelPtX, WheelPtY, _
                   WheelConnectionX, _
                   WheelConnectionY

modStickGame.sLine CSng(pt(4).X), CSng(pt(4).Y), GunPtX, GunPtY

modStickGame.sLine WheelConnectionX, WheelConnectionY, GunPtX, GunPtY

'############################## BLACK ############################################
picMain.ForeColor = vbBlack
modStickGame.sLine Rotor1X, Rotor1Y, Rotor2X, Rotor2Y
modStickGame.sLine RotorX, RotorY, CSng(pt(1).X + 200 * Sine(Facing - Pi)), _
                                   CSng(pt(1).Y - 200 * CoSine(Facing - Pi))

modStickGame.sCircle WheelPtX, WheelPtY, 75, vbBlack


'############################## REST ############################################
picMain.ForeColor = Stick(iStick).colour
modStickGame.sCircle GunPtX, GunPtY, 50, Stick(iStick).colour
modStickGame.sLine GunPtX, GunPtY, GunTipPtX, GunTipPtY


picMain.DrawWidth = 2
'picMain.ForeColor = picMain.BackColor
If Stick(iStick).Team <> Neutral Then
    picMain.ForeColor = GetTeamColour(Stick(iStick).Team)
Else
    picMain.ForeColor = MSilver
End If
modStickGame.sPoly pt, MSilver 'Stick(iStick).Colour

picMain.ForeColor = MSilver
modStickGame.sPoly ScreenPt, Stick(iStick).colour
picMain.DrawStyle = 0
picMain.DrawWidth = 2
picMain.ForeColor = vbBlack
modStickGame.sLine TailRotor1X, TailRotor1Y, TailRotor2X, TailRotor2Y 'must be on top of polygons


''senser/circle on top
'picMain.fillstyle = vbFSSolid
'picMain.FillColor = MSilver
'modStickGame.sCircle RotorX + 200 * sine(Facing - piD2), RotorY - 200 * cosine(Facing - piD2), 200, ChopperCol
'picMain.fillstyle = vbFSTransparent

'On Error Resume Next
'rotor adjust
Facing = GetSticksTimeZone(iStick)

If Stick(iStick).RotorDir Then
    If Stick(iStick).RotorWidth < (RotorInc + 1) * Facing Then
        Stick(iStick).RotorDir = Not Stick(iStick).RotorDir
    Else
        Stick(iStick).RotorWidth = Stick(iStick).RotorWidth - RotorInc * Facing
    End If
Else
    If Stick(iStick).RotorWidth > CLD2 Then
        Stick(iStick).RotorDir = Not Stick(iStick).RotorDir
    Else
        Stick(iStick).RotorWidth = Stick(iStick).RotorWidth + RotorInc * Facing
    End If
End If


Stick(iStick).TailRotorFacing = Stick(iStick).TailRotorFacing + TailRotorInc * Facing
If Stick(iStick).TailRotorFacing > Pi2 Then
    Stick(iStick).TailRotorFacing = FixAngle(Stick(iStick).TailRotorFacing)
End If


End Sub

Private Sub DrawTypeBubble(X As Single, Y As Single) ', Col As Long)

'modStickGame.sCircle X, Y, 100, Col
modStickGame.PrintStickText "Typing", X - 250, Y - 1250, vbRed 'Col

End Sub

Private Sub DrawLegs(i As Integer, X As Single, ByVal Y As Single, _
    Crouching As Boolean, Prone As Boolean, Col As Long)

If Stick(i).bOnSurface Then
    DrawSurfaceLegs i, X, Y, Crouching, Prone, Col
Else
    DrawAirLegs i, X, Y, Col
End If

End Sub

Private Sub DrawAirLegs(i As Integer, X As Single, Y As Single, Col As Long)

Dim nY As Single, iDirection As Integer
Dim Knee1X As Single, Knee1Y As Single 'front knee
Dim Knee2X As Single, Knee2Y As Single

Dim Foot1X As Single, Foot1Y As Single 'front foot
Dim Foot2X As Single, Foot2Y As Single

If Stick(i).Facing > Pi Then
    iDirection = -1
Else
    iDirection = 1
End If
nY = Y + BodyLen



Knee1X = X + ArmLen * iDirection
Knee1Y = nY + HeadRadius

Knee2X = X + ArmLen / 3 * iDirection
Knee2Y = nY + HeadRadius * 1.5

'note: relative to Knee1
Foot1X = Knee1X - ArmLen * iDirection
Foot1Y = Knee2Y + HeadRadius

Foot2X = X - ArmLen / 2 * iDirection
Foot2Y = 2 * Knee2Y - nY - HeadRadius


picMain.ForeColor = Col
modStickGame.sLine X, nY, Knee1X, Knee1Y
modStickGame.sLine X, nY, Knee2X, Knee2Y

modStickGame.sLine Foot1X, Foot1Y, Knee1X, Knee1Y
modStickGame.sLine Foot2X, Foot2Y, Knee2X, Knee2Y

End Sub

Private Sub DrawSurfaceLegs(i As Integer, X As Single, ByVal Y As Single, _
    Crouching As Boolean, Prone As Boolean, Col As Long)

Dim Knee1X As Single, Knee1Y As Single
Dim Knee2X As Single, Knee2Y As Single
Dim iDirection As Integer, LegSgn As Single
Const HRx1p8 = HeadRadius * 1.8

If Stick(i).Facing > Pi Then
    iDirection = -1
Else
    iDirection = 1
End If

picMain.ForeColor = Col


If Crouching Then
    
    
    
''    '############################################ 1st knee
''    'make legs slightly wider
''    Knee1X = X + iDirection * Abs(Stick(i).LegWidth)
''    Knee1Y = Y + BodyLen + LegHeight / 4 * iDirection * Abs(Stick(i).LegWidth) / 90
''
''
''    modstickgame.sLine X, Y + BodyLen,Knee1X, Knee1Y)
''    modstickgame.sLine Knee1X, Knee1Y,Knee1X, Y + BodyLen / 2 + LegHeight)
''
''
''    '############################################ 2nd knee
''    'make legs slightly wider
''    Knee2X = X + iDirection * Abs(Stick(i).LegWidth / 4)
''    Knee2Y = Y + BodyLen + LegHeight / 4
''
''
''    modstickgame.sLine X, Y + BodyLen,Knee2X, Knee2Y)
''    modstickgame.sLine Knee2X, Knee2Y,Knee2X, Y + BodyLen / 2 + LegHeight)
    
    
    
    If Stick(i).Speed > 6 Then

        '############################################ 1st knee
        'make legs slightly wider
        Knee1X = X + iDirection * Abs(Stick(i).LegWidth)
        Knee1Y = Y + BodyLen + LegHeight / 4


        modStickGame.sLine X, Y + BodyLen, Knee1X, Knee1Y
        modStickGame.sLine Knee1X, Knee1Y, Knee1X, Y + BodyLen / 2 + LegHeight


        '############################################ 2nd knee
        'make legs slightly wider
        Knee2X = X + iDirection * Abs(Stick(i).LegWidth / 4)
        Knee2Y = Y + BodyLen + LegHeight / 4


        modStickGame.sLine X, Y + BodyLen, Knee2X, Knee2Y
        modStickGame.sLine Knee2X, Knee2Y, Knee2X, Y + BodyLen / 2 + LegHeight

    Else
        
        If Abs(Stick(i).LegWidth) < 44 Then
            Stick(i).LegWidth = 44 * Sgn(Stick(i).LegWidth)
        End If
        
        '############################################ 1st knee
        Knee1X = X + Stick(i).LegWidth / 2
        Knee1Y = Y + BodyLen / 2 + LegHeight / 2
        
        
        modStickGame.sLine X, Y + BodyLen, Knee1X, Knee1Y
        modStickGame.sLine Knee1X, Knee1Y, Knee1X, Y + BodyLen / 2 + LegHeight

        '############################################ 2nd knee
        Knee2X = X - Stick(i).LegWidth / 2
        Knee2Y = Y + BodyLen / 2 + LegHeight / 2
        
        
        modStickGame.sLine X, Y + BodyLen, Knee2X, Knee2Y
        modStickGame.sLine Knee2X, Knee2Y, Knee2X, Y + BodyLen / 2 + LegHeight
    End If
    
    '#########
    'modstickgame.sLine X + Stick(i).LegWidth, Y + BodyLen + LegHeight / 2,X + Stick(i).LegWidth, Y + BodyLen + LegHeight)
    'modstickgame.sLine X + -Stick(i).LegWidth, Y + BodyLen + LegHeight / 2,X - Stick(i).LegWidth, Y + BodyLen + LegHeight)
    
ElseIf Prone Then
    
    LegSgn = IIf(Stick(i).Facing > Pi, 1, -1)
    
    Y = Y + HRx1p8
    
    modStickGame.sLine X + LegSgn * BodyLen, Y, _
                       X + LegSgn * (BodyLen + LegHeight / 2), Y
    
Else
    modStickGame.sLine X, Y + BodyLen, X + Stick(i).LegWidth, Y + BodyLen + LegHeight 'leg 1
    modStickGame.sLine X, Y + BodyLen, X - Stick(i).LegWidth, Y + BodyLen + LegHeight 'leg 2
End If

End Sub

Private Sub DrawW1200(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If Facing > Pi Then
    Flip = True
    
    If Reloading Then Facing = 5 * Pi / 4
    
    Facing = Facing - Pi
    kY = 1
Else
    If Reloading Then Facing = pi3D4
    kY = -1
End If

'hand position
Hand1X = Stick(i).X - ArmLen / 2

If StickiHasState(i, STICK_CROUCH) Then
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen / 0.8
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 2
    End If
Else
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen / 0.8
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 2
    End If
End If
'Hand1Y = Stick(i).Y + ArmNeckDist + IIf(StickHasState(Stick(i).ID, Stick_Crouch), BodyLen, BodyLen / 2)

DrawW12002 Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i

End Sub
Private Sub DrawW12002(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer)

Dim X(1 To 11) As Single, Y(1 To 11) As Single
Dim j As Integer
Const SAd2 = SmallAngle / 2


X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 2 * Sine(Facing + kY * SmallAngle)
Y(2) = Y(1) - GunLen / 2 * CoSine(Facing + kY * SmallAngle)

X(3) = X(1) + GunLen / 1.5 * Sine(Facing + kY * SmallAngle)
Y(3) = Y(1) - GunLen / 1.5 * CoSine(Facing + kY * SmallAngle)

X(4) = X(1) + GunLen / 1.5 * Sine(Facing + kY * SAd2)
Y(4) = Y(1) - GunLen / 1.5 * CoSine(Facing + kY * SAd2)

X(5) = X(1) + GunLen * Sine(Facing + kY * SAd2)
Y(5) = Y(1) - GunLen * CoSine(Facing + kY * SAd2)

'pump action bit
X(6) = X(1) + GunLen * Sine(Facing + kY * SAd2)
Y(6) = Y(1) - GunLen * CoSine(Facing + kY * SAd2)

X(7) = X(1) + GunLen * 1.5 * Sine(Facing + kY * SmallAngle / 3)
Y(7) = Y(1) - GunLen * 1.5 * CoSine(Facing + kY * SmallAngle / 3)
'end pump action bit

X(8) = X(1) + GunLen * 2 * Sine(Facing + kY * SmallAngle / 3)
Y(8) = Y(1) - GunLen * 2 * CoSine(Facing + kY * SmallAngle / 3)

X(9) = X(1) + GunLen * 2.5 * Sine(Facing + kY * SmallAngle / 3.5)
Y(9) = Y(1) - GunLen * 2.5 * CoSine(Facing + kY * SmallAngle / 3.5)

X(10) = X(9) + GunLen / 6 * Sine(Facing + kY * pi2d3)
Y(10) = Y(9) - GunLen / 6 * CoSine(Facing + kY * pi2d3)

X(11) = X(9) + GunLen / 20 * Sine(Facing + kY * Pi)
Y(11) = Y(9) - GunLen / 20 * CoSine(Facing + kY * Pi)

If Flip Then
    'flip image
    For j = 1 To 11
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    For j = 1 To 11
        Y(j) = 2 * sY - Y(j) + BodyLen * 2.2
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(1)
End If

Hand2X = X(7)
Hand2Y = Y(7)

'end calculation

If CanSeeStick(i) Then
    picMain.DrawWidth = 2
    
    'handle section
    picMain.ForeColor = vbRed
    modStickGame.sLine X(1), Y(1), X(3), Y(3)
    
    picMain.DrawWidth = 2
    picMain.ForeColor = vbBlack
    modStickGame.sLine X(2), Y(2), X(4), Y(4)
    
    picMain.ForeColor = &H555555
    modStickGame.sLine X(2), Y(2), X(8), Y(8)
    modStickGame.sLine X(3), Y(3), X(9), Y(9)
    
    picMain.ForeColor = vbRed
    modStickGame.sLine X(1), Y(1), X(4), Y(4)
    modStickGame.sLine X(6), Y(6), X(7), Y(7)
    
    'picMain.ForeColor = &H555555
    picMain.DrawWidth = 1
    modStickGame.sLine X(10), Y(10), X(11), Y(11)
    
    'modstickgame.sLine X(), Y(),X(), Y())
End If

Stick(i).GunPoint.X = X(9)
Stick(i).GunPoint.Y = Y(9)

Stick(i).CasingPoint.X = X(7)
Stick(i).CasingPoint.Y = Y(7)

picMain.DrawWidth = 1

End Sub

Private Sub DrawAK(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, ByRef A_Facing As Single)

'Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single


A_Facing = FixAngle(Stick(i).Facing)

Reloading = StickiHasState(i, STICK_RELOAD)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = pi3D4
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = piD4
    kY = 1
End If


'hand position
Hand1X = sX + ArmLen / 4

If StickiHasState(i, STICK_CROUCH) Then
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 10
    End If
Else
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen / 1.2
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 4
    End If
End If

DrawAK2 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading



End Sub
Private Sub DrawAK2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    Reloading As Boolean)

Dim X(1 To 18) As Single, Y(1 To 18) As Single
Dim j As Integer
Dim tX As Single, tY As Single

Const SAd2 = SmallAngle / 2
Const SAd4 = SmallAngle / 4
Const SAd8 = SmallAngle / 8

X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 4 * Sine(Facing + kY * 11 * Pi / 18)
Y(2) = Y(1) - GunLen / 4 * CoSine(Facing + kY * 11 * Pi / 18) '90+20deg


X(3) = X(1) + GunLen / 4 * Sine(Facing + kY * piD2)
Y(3) = Y(1) - GunLen / 4 * CoSine(Facing + kY * piD2)

X(4) = X(1) + GunLen / 20 * Sine(Facing)
Y(4) = Y(1) - GunLen / 20 * CoSine(Facing)

X(5) = X(1) + GunLen / 4 * Sine(Facing)
Y(5) = Y(1) - GunLen / 4 * CoSine(Facing)

X(6) = X(1) + GunLen / 3.2 * Sine(Facing - kY * SAd2)
Y(6) = Y(1) - GunLen / 3.2 * CoSine(Facing - kY * SAd2)

X(7) = X(6) + GunLen / 1.5 * Sine(Facing + kY * piD4)
Y(7) = Y(6) - GunLen / 1.5 * CoSine(Facing + kY * piD4)

X(8) = X(7) + GunLen / 4 * Sine(Facing - kY * piD4)
Y(8) = Y(7) - GunLen / 4 * CoSine(Facing - kY * piD4)

X(9) = X(1) + GunLen / 2 * Sine(Facing - kY * SAd2)
Y(9) = Y(1) - GunLen / 2 * CoSine(Facing - kY * SAd2)

X(10) = X(9) + GunLen * Sine(Facing - kY * SAd8)
Y(10) = Y(9) - GunLen * CoSine(Facing - kY * SAd8)

X(11) = X(10) + GunLen / 4 * Sine(Facing - kY * piD2)
Y(11) = Y(10) - GunLen / 4 * CoSine(Facing - kY * piD2)

X(12) = X(11) + GunLen / 4 * Sine(Facing + kY * (piD2 + SmallAngle))
Y(12) = Y(11) - GunLen / 4 * CoSine(Facing + kY * (piD2 + SmallAngle))

X(13) = X(12) + GunLen / 3 * Sine(Facing - kY * Pi)
Y(13) = Y(12) - GunLen / 3 * CoSine(Facing - kY * Pi)

X(14) = X(13) + GunLen / 3 * Sine(Facing - kY * Pi)
Y(14) = Y(13) - GunLen / 3 * CoSine(Facing - kY * Pi)

X(15) = X(14) + GunLen * 0.6 * Sine(Facing + kY * (Pi - SAd4))
Y(15) = Y(14) - GunLen * 0.6 * CoSine(Facing + kY * (Pi - SAd4))

X(16) = X(2) + GunLen / 2 * Sine(Facing - kY * (Pi + SAd4))
Y(16) = Y(2) - GunLen / 2 * CoSine(Facing - kY * (Pi + SAd4))

X(17) = X(16) + GunLen / 4 * Sine(Facing + kY * (piD2 - SAd4))
Y(17) = Y(16) - GunLen / 4 * CoSine(Facing + kY * (Pi / 2 - SAd4))

X(18) = X(1) + GunLen / 8 * Sine(Facing - kY * Pi)
Y(18) = Y(1) - GunLen / 8 * CoSine(Facing - kY * Pi)


If Flip Then
    'flip image
    For j = 1 To 18
        Y(j) = 2 * sY - Y(j) + BodyLen * 1.6
    Next j
    
    For j = 1 To 18
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(1)
End If

Hand2X = X(5)
Hand2Y = Y(5)

'end calculation

'drawing

If CanSeeStick(i) Then
    picMain.DrawWidth = 2
    picMain.ForeColor = &H6AD5
    'handle
    modStickGame.sLine X(1), Y(1), X(2), Y(2)
    modStickGame.sLine X(3), Y(3), X(2), Y(2)
    modStickGame.sLine X(3), Y(3), X(4), Y(4)
    
    picMain.ForeColor = vbBlack
    picMain.DrawWidth = 2
    'handle-mag bit
    modStickGame.sLine X(5), Y(5), X(4), Y(4)
    modStickGame.sLine X(5), Y(5), X(6), Y(6)
    
    'magazine
    picMain.DrawWidth = 2
    If Reloading = False Then
        modStickGame.sLine X(7), Y(7), X(6), Y(6)
        modStickGame.sLine X(7), Y(7), X(8), Y(8)
        modStickGame.sLine X(9), Y(9), X(8), Y(8)
    End If
    
    'magazine top bit
    modStickGame.sLine X(9), Y(9), X(6), Y(6)
    
    'barrel
    picMain.ForeColor = &H6AD5
    modStickGame.sLine X(9), Y(9), X(10), Y(10)
    picMain.ForeColor = vbBlack
    modStickGame.sLine X(11), Y(11), X(10), Y(10) 'iron sight
    modStickGame.sLine X(11), Y(11), X(12), Y(12) 'iron sight
    picMain.ForeColor = &H6AD5
    modStickGame.sLine X(13), Y(13), X(12), Y(12)
    modStickGame.sLine X(13), Y(13), X(14), Y(14)
    picMain.ForeColor = vbBlack
    modStickGame.sLine X(15), Y(15), X(14), Y(14)
    
    'stock
    picMain.ForeColor = &H6AD5
    modStickGame.sLine X(15), Y(15), X(16), Y(16)
    modStickGame.sLine X(17), Y(17), X(16), Y(16)
    modStickGame.sLine X(17), Y(17), X(18), Y(18)
    picMain.ForeColor = vbBlack
    modStickGame.sLine X(18), Y(18), X(1), Y(1)
    
    
    'If Stick(i).bSilenced Then
        'DrawSilencer X(10), Y(10), Facing + IIf(Stick(i).Facing > Pi, Pi, 0)
    'End If
End If

Stick(i).GunPoint.X = X(10)
Stick(i).GunPoint.Y = Y(10)
Stick(i).CasingPoint.X = X(6)
Stick(i).CasingPoint.Y = Y(6)

'modstickgame.sLine X(), Y(),X(), Y())

picMain.DrawWidth = 1
End Sub

Private Sub DrawXM8(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, ByRef A_Facing As Single)

'Dim Facing As Single

Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const ArmLenD4 As Single = ArmLen / 4


A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = pi3D4 '1-below
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = piD4 'below is here
    kY = 1
End If

'hand position
Hand1X = sX + ArmLenD4


If Flip Then
    'If StickiHasState(i, STICK_CROUCH) Then
        'Hand1Y = sY + HeadRadius + BodyLen
    'Else
        Hand1Y = sY + HeadRadius + BodyLen / 1.2
    'End If
Else
    Hand1Y = sY + HeadRadius + BodyLen / 4
End If

DrawXM82 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading



End Sub
Private Sub DrawXM82(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    Reloading As Boolean)

Dim pt(1 To 17) As PointAPI
Dim PtGap(1 To 3) As PointAPI
Dim ptMag(1 To 4) As PointAPI

Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
'Dim Grip1X As Single, Grip1Y As Single, Grip2X As Single, Grip2Y As Single
Dim j As Integer

Const XM8_Col = &H101010, XM8_Mag_Col = &H202020
Dim tX As Single, tY As Single
Dim SinFacing As Single, CosFacing As Single
Dim tSin As Single, tCos As Single
Const BodyLenX1p6 = BodyLen * 1.6, Barrel_Len = GunLen / 3, GunLenD6 = GunLen / 6

SinFacing = Sine(Facing)
CosFacing = CoSine(Facing)

pt(1).X = Hand1X
pt(1).Y = Hand1Y

pt(2).X = pt(1).X + GunLen / 3 * Sine(Facing + kY * pi3D4)
pt(2).Y = pt(1).Y - GunLen / 3 * CoSine(Facing + kY * pi3D4)

pt(3).X = pt(2).X + GunLen / 6 * SinFacing
pt(3).Y = pt(2).Y - GunLen / 6 * CosFacing

pt(4).X = pt(1).X + GunLen / 6 * SinFacing
pt(4).Y = pt(1).Y - GunLen / 6 * CosFacing

pt(5).X = pt(4).X + GunLen / 6 * SinFacing
pt(5).Y = pt(4).Y - GunLen / 6 * CosFacing


tSin = Sine(Facing - kY * piD8)
tCos = CoSine(Facing - kY * piD8)

pt(6).X = pt(5).X + GunLen / 4 * tSin
pt(6).Y = pt(5).Y - GunLen / 4 * tCos

'#######
ptMag(1) = pt(5)

ptMag(2).X = pt(5).X + GunLen / 3 * Sine(Facing + kY * pi4D9)
ptMag(2).Y = pt(5).Y - GunLen / 3 * CoSine(Facing + kY * pi4D9)

ptMag(3).X = ptMag(2).X + GunLen / 4 * tSin
ptMag(3).Y = ptMag(2).Y - GunLen / 4 * tCos

ptMag(4) = pt(6)
'#######

pt(7).X = pt(6).X + GunLen / 5 * tSin
pt(7).Y = pt(6).Y - GunLen / 5 * tCos

'straight bottom part of barrel
pt(8).X = pt(7).X + GunLen / 1.5 * SinFacing
pt(8).Y = pt(7).Y - GunLen / 1.5 * CosFacing

'wedge
pt(9).X = pt(8).X + GunLen / 2.8 * Sine(Facing - kY * pi3D4)
pt(9).Y = pt(8).Y - GunLen / 2.8 * CoSine(Facing - kY * pi3D4)


pt(10).X = pt(9).X + GunLen / 1.4 * Sine(Facing - kY * Pi)
pt(10).Y = pt(9).Y - GunLen / 1.4 * CoSine(Facing - kY * Pi)

pt(11).X = pt(10).X + GunLen / 6 * Sine(Facing - kY * piD2)
pt(11).Y = pt(10).Y - GunLen / 6 * CoSine(Facing - kY * piD2)

pt(12).X = pt(11).X + GunLen / 3 * Sine(Facing - kY * Pi)
pt(12).Y = pt(11).Y - GunLen / 3 * CoSine(Facing - kY * Pi)

pt(13).X = pt(12).X + GunLen / 6 * Sine(Facing + kY * piD2)
pt(13).Y = pt(12).Y - GunLen / 6 * CoSine(Facing + kY * piD2)

pt(14).X = pt(13).X + GunLen / 15 * Sine(Facing + kY * piD2)
pt(14).Y = pt(13).Y - GunLen / 15 * CoSine(Facing + kY * piD2)

'top buttstock
pt(15).X = pt(14).X + GunLen / 2 * Sine(Facing - kY * (Pi * 1.1))
pt(15).Y = pt(14).Y - GunLen / 2 * CoSine(Facing - kY * (Pi * 1.1))

'bottom buttstock
pt(16).X = pt(15).X + GunLen / 3 * Sine(Facing + kY * piD2)
pt(16).Y = pt(15).Y - GunLen / 3 * CoSine(Facing + kY * piD2)

pt(17).X = pt(16).X + GunLen / 4 * Sine(Facing - kY * piD2)
pt(17).Y = pt(16).Y - GunLen / 4 * CoSine(Facing - kY * piD2)


''start of fancy bits
'Pt(20) = Pt(9) + GunLen / 6 * tSin 'F-piD8
'Pt(20) = Pt(9) - GunLen / 6 * tCos
'
'Pt(21) = Pt(20) + GunLen / 2 * SinFacing
'Pt(21) = Pt(20) - GunLen / 2 * CosFacing
'
'Pt(22) = Pt(20) + GunLen / 6 * sine(Facing - kY * piD2)
'Pt(22) = Pt(20) - GunLen / 6 * cosine(Facing - kY * piD2)
'
'Pt(23) = Pt(22) + GunLen / 3 * SinFacing
'Pt(23) = Pt(22) - GunLen / 3 * CosFacing


'#############
'Hole in front of scope
PtGap(1).X = pt(14).X + GunLen / 6 * SinFacing
PtGap(1).Y = pt(14).Y - GunLen / 6 * CosFacing

PtGap(2).X = PtGap(1).X + GunLen / 8 * Sine(Facing + kY * piD2)
PtGap(2).Y = PtGap(1).Y - GunLen / 8 * CoSine(Facing + kY * piD2)

PtGap(3).X = PtGap(2).X + GunLen / 1.5 * Sine(Facing - kY * piD20)
PtGap(3).Y = PtGap(2).Y - GunLen / 1.5 * CoSine(Facing - kY * piD20)

'Pt(22) = Pt(21) + GunLen / 5.2 * sine(Facing - piD4)
'Pt(22) = Pt(21) - GunLen / 5.2 * cosine(Facing - piD4)


'#############
'barrel
Barrel1X = pt(8).X + GunLenD6 * Sine(Facing - kY * pi3D4)
Barrel1Y = pt(8).Y - GunLenD6 * CoSine(Facing - kY * pi3D4)

Barrel2X = Barrel1X + Barrel_Len * SinFacing 'GunLen/x = BarrelLen
Barrel2Y = Barrel1Y - Barrel_Len * CosFacing


'#############
'grip
'Grip1X = Pt(7).X + GunLen / 3 * SinFacing
'Grip1Y = Pt(7).Y - GunLen / 3 * CosFacing

'Grip2X = Grip1X + GunLen / 3 * Sine(Facing + kY * piD2)
'Grip2Y = Grip1Y - GunLen / 3 * CoSine(Facing + kY * piD2)



'Pt(26) = Pt(11) + GunLen / 6 * sine(Facing + pi3D4)
'Pt(26) = Pt(11) - GunLen / 6 * cosine(Facing + pi3D4)
'
'Pt(25) = Pt(26) + GunLen / 2 * SinFacing
'Pt(25) = Pt(26) - GunLen / 2 * CosFacing


If Flip Then
    'flip image
    For j = 1 To 17
        pt(j).Y = 2 * sY - pt(j).Y + BodyLenX1p6
        pt(j).X = 2 * sX - pt(j).X
    Next j
    
    For j = 1 To 3
        PtGap(j).Y = 2 * sY - PtGap(j).Y + BodyLenX1p6
        PtGap(j).X = 2 * sX - PtGap(j).X
    Next j
    
    For j = 1 To 4
        ptMag(j).Y = 2 * sY - ptMag(j).Y + BodyLenX1p6
        ptMag(j).X = 2 * sX - ptMag(j).X
    Next j
    
    Barrel1X = 2 * sX - Barrel1X
    Barrel2X = 2 * sX - Barrel2X
    Barrel1Y = 2 * sY - Barrel1Y + BodyLenX1p6
    Barrel2Y = 2 * sY - Barrel2Y + BodyLenX1p6
    
    'Grip1X = 2 * sX - Grip1X
    'Grip2X = 2 * sX - Grip2X
    'Grip1Y = 2 * sY - Grip1Y + BodyLenX1p6
    'Grip2Y = 2 * sY - Grip2Y + BodyLenX1p6
    
    'For j = 1 To 24
        'Pt(j) = Pt(j) - 2 * (Pt(j) - Stick(i).X)
        'Pt(j) = 2 * sX - Pt(j)
    'Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = pt(1).Y
End If

Hand2X = pt(7).X '(Grip1X + Grip2X) / 2
Hand2Y = pt(7).Y '(Grip1Y + Grip2Y) / 2
'end calculation


Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y

Stick(i).CasingPoint.X = pt(5).X
Stick(i).CasingPoint.Y = pt(5).Y


'drawing
If CanSeeStick(i) Then
    picMain.DrawWidth = 1
    'picMain.ForeColor = &H2F2F2F
    picMain.ForeColor = XM8_Col
    
    'If Stick(i).bSilenced Then
        'DrawSilencer Barrel1X, Barrel1Y, Facing + IIf(Stick(i).Facing > Pi, Pi, 0)
    'End If
    
    
    
    picMain.DrawWidth = 1
    modStickGame.sPoly pt, XM8_Col
    
    picMain.ForeColor = XM8_Col
    picMain.DrawWidth = 2
    'modStickGame.sLine Grip1X, Grip1Y, Grip2X, Grip2Y
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
    modStickGame.sPoly PtGap, modStickGame.cg_BGColour
    
    
    If Not Reloading Then
        picMain.DrawWidth = 1
        modStickGame.sPoly ptMag, XM8_Mag_Col
    End If
    
    
'    modStickGame.sLine Pt(1), Pt(1), Pt(2), Pt(2)
'    modStickGame.sLine Pt(2), Pt(2), Pt(3), Pt(3)
'    modStickGame.sLine Pt(3), Pt(3), Pt(4), Pt(4)
'    modStickGame.sLine Pt(4), Pt(4), Pt(5), Pt(5)
'
'    'magazine
'    If Reloading = False Then
'        modStickGame.sLine Pt(5), Pt(5), Pt(6), Pt(6)
'        modStickGame.sLine Pt(6), Pt(6), Pt(7), Pt(7)
'        modStickGame.sLine Pt(7), Pt(7), Pt(8), Pt(8)
'    End If
'
'    'mag modstickgame.sLine
'    modStickGame.sLine Pt(5), Pt(5), Pt(8), Pt(8)
'
'    modStickGame.sLine Pt(8), Pt(8), Pt(9), Pt(9)
'    modStickGame.sLine Pt(9), Pt(9), Pt(10), Pt(10)
'    modStickGame.sLine Pt(10), Pt(10), Pt(11), Pt(11)
'    modStickGame.sLine Pt(11), Pt(11), Pt(12), Pt(12)
'    modStickGame.sLine Pt(12), Pt(12), Pt(13), Pt(13)
'    modStickGame.sLine Pt(13), Pt(13), Pt(14), Pt(14)
'    modStickGame.sLine Pt(14), Pt(14), Pt(15), Pt(15)
'    modStickGame.sLine Pt(15), Pt(15), Pt(16), Pt(16)
'    modStickGame.sLine Pt(16), Pt(16), Pt(17), Pt(17)
'    modStickGame.sLine Pt(17), Pt(17), Pt(18), Pt(18)
'    modStickGame.sLine Pt(18), Pt(18), Pt(19), Pt(19)
'
'    'hole bit
'    modStickGame.sLine Pt(20), Pt(20), Pt(21), Pt(21)
'    'modstickgame.sLine Pt(22), Pt(22),Pt(21), Pt(21)
'    modStickGame.sLine Pt(20), Pt(20), Pt(22), Pt(22)
'
'    'scope modstickgame.sLine
'    modStickGame.sLine Pt(16), Pt(16), Pt(22), Pt(22)
'
'    'connect stock to handle
'    modStickGame.sLine Pt(1), Pt(1), Pt(18), Pt(18)
'
'    'barrel
'    picMain.DrawWidth = 1
'    modStickGame.sLine Pt(23), Pt(23), Pt(24), Pt(24)
'    'modstickgame.sLine Pt(25), Pt(25),Pt(26), Pt(26)
    
    
    picMain.DrawWidth = 1
End If

End Sub

Private Sub DrawM82(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, _
    StickCol As Long, ByRef A_Facing As Single)

'Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean ', bProne As Boolean
Dim kY As Single
Dim Adj As Single
Dim GTC As Long
Dim BarrelLen As Single '1 = normal


A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)
'bProne = StickiHasState(i, Stick_Prone)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = pi3D5
        'Facing = piD2
        'Facing = pi * 0.2
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then
        A_Facing = pi2D5
        'Facing = piD2
        'Facing = pi / 1.2
    End If
    kY = 1
End If

If i = 0 Then
    GTC = GetTickCount()
    
    Adj = GetSticksTimeZone(i)
    
    If Stick(i).LastBullet + M82_Recoil_Time / Adj > GTC Then
        BarrelLen = (GTC - Stick(0).LastBullet) * Adj / 850 + 0.55
        'change 750 - i think:
        '750 proportionalTo GameSpeed
        
        '(Reload_Time - (GetTickCount() - Stick(0).ReloadStart))
        
        '(Stick(i).LastBullet - GTC + GunLen) / (M82_Recoil_Time - Stick(i).LastBullet)
    Else
        BarrelLen = 1
    End If
Else
    BarrelLen = 1
End If


'hand position
Hand1X = sX - ArmLen / 4
If Flip Then
    'If StickHasState(Stick(i).ID, Stick_Crouch) Then
        'Hand1Y = sY + HeadRadius + BodyLen / 1.4
    'Else
    Hand1Y = sY + HeadRadius + BodyLen / 1.4
    'End If
Else
    Hand1Y = sY + HeadRadius + BodyLen / 2.8
End If


DrawM822 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading, BarrelLen, StickCol, _
    (GetTickCount() - Stick(i).LastNade) > (Nade_Arm_Time / GetSticksTimeZone(i))

End Sub
Private Sub DrawM822(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, Reloading As Boolean, _
    BarrelLenFactor As Single, StickCol As Long, bDoArm As Boolean)

Dim X(1 To 32) As Single, Y(1 To 32) As Single
Dim j As Integer

Const GLd10 = GunLen / 10
Const SAd4 = SmallAngle / 4

Dim SinFacing As Single
Dim CosFacing As Single
Dim SinFacingLess_kYpiD2 As Single, SinFacingLess_kYpiD4 As Single
Dim CosFacingLess_kYpiD2 As Single, CosFacingLess_kYpiD4 As Single


SinFacingLess_kYpiD2 = Sine(Facing - kY * piD2)
CosFacingLess_kYpiD2 = CoSine(Facing - kY * piD2)
SinFacingLess_kYpiD4 = Sine(Facing - kY * piD4)
CosFacingLess_kYpiD4 = CoSine(Facing - kY * piD4)

SinFacing = Sine(Facing)
CosFacing = CoSine(Facing)

X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 4 * Sine(Facing - kY * piD4)
Y(2) = Y(1) - GunLen / 4 * CoSine(Facing - kY * piD4)

X(3) = X(2) + GunLen / 6 * SinFacing
Y(3) = Y(2) - GunLen / 6 * CosFacing

X(4) = X(1) + GunLen / 6 * SinFacing
Y(4) = Y(1) - GunLen / 6 * CosFacing

X(5) = X(4) + GunLen / 6 * SinFacing
Y(5) = Y(4) - GunLen / 6 * CosFacing

X(6) = X(2) + GunLen / 4 * SinFacing
Y(6) = Y(2) - GunLen / 4 * CosFacing

X(7) = X(6) + GunLen / 2 * SinFacing
Y(7) = Y(6) - GunLen / 2 * CosFacing

X(8) = X(5) + GunLen / 3 * SinFacing
Y(8) = Y(5) - GunLen / 3 * CosFacing

X(9) = X(8) + GunLen / 3 * SinFacing
Y(9) = Y(8) - GunLen / 3 * CosFacing

X(10) = X(7) + GunLen / 3 * SinFacing
Y(10) = Y(7) - GunLen / 3 * CosFacing

X(11) = X(10) + GunLen / 2 * SinFacing
Y(11) = Y(10) - GunLen / 2 * CosFacing

X(12) = X(11) + GunLen / 40 * SinFacingLess_kYpiD2
Y(12) = Y(11) - GunLen / 40 * CosFacingLess_kYpiD2

X(13) = X(12) + GunLen * 1.5 * SinFacing * BarrelLenFactor 'BARREL
Y(13) = Y(12) - GunLen * 1.5 * CosFacing * BarrelLenFactor

X(14) = X(12) + GunLen / 8 * Sine(Facing - kY * Pi)
Y(14) = Y(12) - GunLen / 8 * CoSine(Facing - kY * Pi)

X(15) = X(14) + GLd10 * SinFacingLess_kYpiD2
Y(15) = Y(14) - GLd10 * CosFacingLess_kYpiD2

X(16) = X(15) + GunLen / 10 * Sine(Facing - kY * Pi) 'iron sight bottom
Y(16) = Y(15) - GunLen / 10 * CoSine(Facing - kY * Pi)

X(17) = X(16) + GunLen / 10 * SinFacingLess_kYpiD2 'iron sight top
Y(17) = Y(16) - GunLen / 10 * CosFacingLess_kYpiD2

X(18) = X(15) + GunLen / 6 * Sine(Facing - kY * Pi)
Y(18) = Y(15) - GunLen / 6 * CoSine(Facing - kY * Pi)

X(19) = X(18) + GunLen / 2 * Sine(Facing - kY * Pi) 'end of straight top bit
Y(19) = Y(18) - GunLen / 2 * CoSine(Facing - kY * Pi)

X(20) = X(1) + GunLen / 4 * SinFacingLess_kYpiD2
Y(20) = Y(1) - GunLen / 4 * CosFacingLess_kYpiD2

'sight stand
'bottom points
X(21) = X(18) + GunLen / 8 * Sine(Facing - kY * Pi) 'forward bottom
Y(21) = Y(18) - GunLen / 8 * CoSine(Facing - kY * Pi)

X(22) = X(21) + GunLen / 4 * Sine(Facing - kY * Pi) 'rearward bottom
Y(22) = Y(21) - GunLen / 4 * CoSine(Facing - kY * Pi)
'top points
X(23) = X(21) + GunLen / 6 * SinFacingLess_kYpiD2 'forward top
Y(23) = Y(21) - GunLen / 6 * CosFacingLess_kYpiD2

X(24) = X(22) + GunLen / 6 * SinFacingLess_kYpiD2 'rearward top
Y(24) = Y(22) - GunLen / 6 * CosFacingLess_kYpiD2
'modstickgame.sLine from 21->23, 22->24

'scope
X(25) = X(24) + GunLen / 4 * Sine(Facing - kY * Pi) 'rear bottom pt
Y(25) = Y(24) - GunLen / 4 * CoSine(Facing - kY * Pi)

X(26) = X(24) + GunLen / 1.5 * SinFacing 'front bottom pt
Y(26) = Y(24) - GunLen / 1.5 * CosFacing

X(27) = X(25) + GunLen / 6 * SinFacingLess_kYpiD2 'rear top pt
Y(27) = Y(25) - GunLen / 6 * CosFacingLess_kYpiD2

X(28) = X(26) + GunLen / 8 * SinFacingLess_kYpiD2 'front top pt
Y(28) = Y(26) - GunLen / 8 * CosFacingLess_kYpiD2

'If bProne Then
    'bipod
    X(30) = X(12) + GunLen / 2 * SinFacing 'GunLen/x = Stand's Connection
    Y(30) = Y(12) - GunLen / 2 * CosFacing
    
    X(31) = X(30) + GunLen / 2 * Sine(Facing + kY * Pi / 1.8) 'GunLen/x = Height of Stand
    Y(31) = Y(30) - GunLen / 2 * CoSine(Facing + kY * Pi / 1.8)
    
    X(32) = X(31) + GunLen / 4 * SinFacing 'GunLen/x = separation of stands
    Y(32) = Y(31) - GunLen / 4 * CosFacing
'End If

'flash thing
X(29) = X(13) - GunLen / 6 * SinFacing
Y(29) = Y(13) + GunLen / 6 * CosFacing

'X(29) = X(11) + GunLen / 1.6 * sine(Facing) 'GunLen/x = Start Point of Flashy Bit
'Y(29) = Y(11) - GunLen / 1.6 * cosine(Facing)
'
'X(30) = X(29) + GunLen / 4 * sine(Facing) 'GunLen/x = Length of Flashy Bit
'Y(30) = Y(29) - GunLen / 4 * cosine(Facing)
'
'X(31) = X(30) + GunLen / 8 * SinFacingLess_kYpiD2 'GunLen/x = Height of Flashy Bit
'Y(31) = Y(30) - GunLen / 8 * CosFacingLess_kYpiD2
'
'X(32) = X(31) + GunLen / 4 * sine(Facing - kY * pi) 'Must be same as X(30)
'Y(32) = Y(31) - GunLen / 4 * cosine(Facing - kY * pi)


If Flip Then
    'flip image
    For j = 1 To 32
        Y(j) = 2 * sY - Y(j) + BodyLen * 1.6
    Next j
    
    For j = 1 To 32
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(4)
End If

Hand2X = X(7) + GunLen / 3 * Sine(Facing + piD2)
Hand2Y = Y(7) - GunLen / 3 * CoSine(Facing + piD2)

'end calculation

'drawing

'handle
If CanSeeStick(i) Then
    
    'EXTRA HAND BIT
    If bDoArm Then
        picMain.ForeColor = StickCol
        modStickGame.sLine Hand2X, Hand2Y, X(10), Y(10)
    End If
    
    picMain.DrawWidth = 1
    
    'v. d. blue = &H693F3F
    picMain.ForeColor = &H3F3F3F
    modStickGame.sLine X(1), Y(1), X(2), Y(2)
    modStickGame.sLine X(2), Y(2), X(3), Y(3)
    modStickGame.sLine X(3), Y(3), X(4), Y(4)
    modStickGame.sLine X(4), Y(4), X(5), Y(5)
    modStickGame.sLine X(5), Y(5), X(6), Y(6)
    modStickGame.sLine X(6), Y(6), X(7), Y(7)
    If Not Reloading Then
        modStickGame.sLine X(7), Y(7), X(8), Y(8)
        modStickGame.sLine X(8), Y(8), X(9), Y(9)
        modStickGame.sLine X(9), Y(9), X(10), Y(10)
    End If
    modStickGame.sLine X(10), Y(10), X(11), Y(11)
    modStickGame.sLine X(11), Y(11), X(12), Y(12)
    
    picMain.DrawWidth = 1
    picMain.ForeColor = vbBlack
    modStickGame.sLine X(12), Y(12), X(13), Y(13) 'BARREL
    
    'modStickGame.sLine X(13), Y(13), X(14), Y(14)
    'modStickGame.sLine X(14), Y(14), X(15), Y(15)
    modStickGame.sLine X(12), Y(12), X(15), Y(15)
    
    picMain.DrawWidth = 1
    picMain.ForeColor = &H693F3F
    modStickGame.sLine X(15), Y(15), X(16), Y(16)
    modStickGame.sLine X(16), Y(16), X(17), Y(17)
    modStickGame.sLine X(17), Y(17), X(18), Y(18)
    modStickGame.sLine X(18), Y(18), X(19), Y(19)
    modStickGame.sLine X(19), Y(19), X(20), Y(20)
    
    
    'magazine barrier
    modStickGame.sLine X(7), Y(7), X(10), Y(10)
    
    'end of stock
    modStickGame.sLine X(20), Y(20), X(1), Y(1)
    
    'sight stand
    'modstickgame.sLine from 21->23, 22->24
    modStickGame.sLine X(21), Y(21), X(23), Y(23)
    modStickGame.sLine X(22), Y(22), X(24), Y(24)
    
    'scope
    picMain.ForeColor = vbBlack '&H555555
    picMain.DrawWidth = 2
    modStickGame.sLine X(25), Y(25), X(26), Y(26)
    modStickGame.sLine X(26), Y(26), X(28), Y(28)
    modStickGame.sLine X(28), Y(28), X(27), Y(27)
    modStickGame.sLine X(27), Y(27), X(25), Y(25)
    
    ''flash bit
    'modstickgame.sLine X(29), Y(29),X(30), Y(30))
    'modstickgame.sLine X(30), Y(30),X(31), Y(31))
    'modstickgame.sLine X(31), Y(31),X(32), Y(32))
    'modstickgame.sLine X(32), Y(32),X(29), Y(29))
    
    'flash bit
    Me.picMain.FillStyle = vbFSSolid
    Me.picMain.FillColor = vbBlack
    modStickGame.sCircle X(29), Y(29), 15, vbBlack 'flash thing on end of barrel
    Me.picMain.FillStyle = vbFSTransparent
    
    'bipod
    'If bProne Then
        picMain.DrawWidth = 1
        modStickGame.sLine X(30), Y(30), X(31), Y(31)
        modStickGame.sLine X(30), Y(30), X(32), Y(32)
    'End If
    
    'modstickgame.sLine X(), Y(),X(), Y())
    
    'If Stick(i).bSilenced Then
        'DrawSilencer X(13), Y(13), Facing + IIf(Stick(i).Facing > Pi, Pi, 0)
    'End If
End If

Stick(i).GunPoint.X = X(13)
Stick(i).GunPoint.Y = Y(13)
Stick(i).CasingPoint.X = X(7)
Stick(i).CasingPoint.Y = Y(7)

picMain.DrawWidth = 1

End Sub

Private Sub DrawKnife(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Const SaberGreen As Long = 65280

Dim Facing As Single
Dim X(1 To 5) As Single, Y(1 To 5) As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer
Dim tX As Single, tY As Single

Dim f As Single

Facing = FixAngle(Stick(i).Facing)

'hand position
Hand1X = sX + ArmLen

If Facing > Pi Then
    Flip = True
    Facing = Facing - Pi
    
    kY = -1
    
    
    Hand1Y = sY + HeadRadius + BodyLen / 2.8 + 100 * CoSine(Facing)
Else
    kY = 1
    
    Hand1Y = sY + HeadRadius + BodyLen / 2.8 + 100 * CoSine(Facing)
End If


X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + 400 * Sine(Facing)
Y(2) = Y(1) - 400 * CoSine(Facing)

If Stick(i).bLightSaber = False Then
    X(5) = X(1) + 50 * Sine(Facing)
    Y(5) = Y(1) - 50 * CoSine(Facing)
    
    X(3) = X(5) + 65 * Sine(Facing + piD2)
    Y(3) = Y(5) - 65 * CoSine(Facing + piD2)
    
    X(4) = X(5) + 65 * Sine(Facing - piD2)
    Y(4) = Y(5) - 65 * CoSine(Facing - piD2)
End If


If Flip Then
    'flip image
    For j = 1 To 5
        Y(j) = 2 * sY - Y(j) + BodyLen * 1.1
    Next j
    
    For j = 1 To 5
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(1)
End If

Hand2X = 2 * sX - Hand1X
Hand2Y = Hand1Y

'end calculation

'drawing
If CanSeeStick(i) Then
    
    If Stick(i).bLightSaber Then
        picMain.DrawWidth = 2
        'picMain.ForeColor = SaberGreen
        picMain.ForeColor = SaberGreen
        modStickGame.sLine X(1), Y(1), X(2), Y(2)
        'modStickGame.sLine X(3), Y(3), X(4), Y(4)
        
        picMain.FillStyle = vbFSSolid
        picMain.FillColor = MSilver
        modStickGame.sCircle X(1), Y(1), 25, MSilver
        picMain.FillStyle = vbFSTransparent
    Else
        picMain.DrawWidth = 1
        'picMain.ForeColor = &H3F3F3F
        picMain.ForeColor = &H3F3F3F
        modStickGame.sLine X(1), Y(1), X(2), Y(2)
        modStickGame.sLine X(3), Y(3), X(4), Y(4)
        modStickGame.sCircle X(5), Y(5), 25, &H3F3F3F
        'modstickgame.sLine X(), Y(),X(), Y())
    End If
    
End If

Stick(i).GunPoint.X = X(2)
Stick(i).GunPoint.Y = Y(2)

End Sub

Private Sub DrawRPG(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)


If Facing > Pi Then
    Flip = True
    
    Facing = Facing - Pi
    kY = -1
Else
    kY = 1
End If

'hand position
Hand1X = Stick(i).X + ArmLen / 2

If StickiHasState(i, STICK_CROUCH) Then
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen / 0.8
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 2
    End If
Else
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen / 0.8
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 2
    End If
End If
'Hand1Y = Stick(i).Y + ArmNeckDist + IIf(StickHasState(Stick(i).ID, Stick_Crouch), BodyLen, BodyLen / 2)


DrawRPG2 Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading

End Sub
Private Sub DrawRPG2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, Reloading As Boolean)


Dim X(1 To 16) As Single, Y(1 To 16) As Single
Dim j As Integer
Dim tX As Single, tY As Single
Const SAd2 = SmallAngle / 2


X(2) = Hand1X
Y(2) = Hand1Y

X(1) = X(2) + GunLen / 2 * Sine(Facing - kY * piD2)
Y(1) = Y(2) - GunLen / 2 * CoSine(Facing - kY * piD2)

X(3) = X(1) + GunLen / 1.5 * Sine(Facing)
Y(3) = Y(1) - GunLen / 1.5 * CoSine(Facing)

X(4) = X(3) + GunLen / 2 * Sine(Facing + kY * piD2)
Y(4) = Y(3) - GunLen / 2 * CoSine(Facing + kY * piD2)

X(5) = X(3) + GunLen / 1.5 * Sine(Facing)
Y(5) = Y(3) - GunLen / 1.5 * CoSine(Facing)

X(6) = X(5) + GunLen / 4 * Sine(Facing - kY * piD2)
Y(6) = Y(5) - GunLen / 4 * CoSine(Facing - kY * piD2)

X(7) = X(6) + GunLen * 3 * Sine(Facing - kY * Pi) 'rear top point
Y(7) = Y(6) - GunLen * 3 * CoSine(Facing - kY * Pi)

X(8) = X(1) + GunLen * 1.7 * Sine(Facing - kY * Pi) 'rear bottom point
Y(8) = Y(1) - GunLen * 1.7 * CoSine(Facing - kY * Pi)

'rear funnel
X(9) = X(7) + GunLen / 3 * Sine(Facing - kY * pi3D4) 'rear top point
Y(9) = Y(7) - GunLen / 3 * CoSine(Facing - kY * pi3D4)

X(10) = X(8) + GunLen / 3 * Sine(Facing + kY * pi3D4) 'rear bottom point
Y(10) = Y(8) - GunLen / 3 * CoSine(Facing + kY * pi3D4)

'sights
X(11) = X(6) + GunLen / 1.2 * Sine(Facing - kY * Pi)
Y(11) = Y(6) - GunLen / 1.2 * CoSine(Facing - kY * Pi)

X(12) = X(11) + GunLen / 4 * Sine(Facing - kY * piD2)
Y(12) = Y(11) - GunLen / 4 * CoSine(Facing - kY * piD2)

X(13) = X(12) + GunLen / 4 * Sine(Facing - kY * piD4)
Y(13) = Y(12) - GunLen / 4 * CoSine(Facing - kY * piD4)

X(14) = X(13) + GunLen / 4 * Sine(Facing - kY * piD2)
Y(14) = Y(13) - GunLen / 4 * CoSine(Facing - kY * piD2)

X(15) = X(14) + GunLen / 2 * Sine(Facing + kY * pi3D4)
Y(15) = Y(14) - GunLen / 2 * CoSine(Facing + kY * pi3D4)

X(16) = X(15) + GunLen / 4 * Sine(Facing + kY * piD2)
Y(16) = Y(15) - GunLen / 4 * CoSine(Facing + kY * piD2)

If Flip Then
    'flip image
    For j = 1 To 16
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    For j = 1 To 16
        Y(j) = 2 * sY - Y(j) + BodyLen * 2.2
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(2)
End If

Hand2X = X(4)
Hand2Y = Y(4)
'end calculation

'drawing
If CanSeeStick(i) Then
    picMain.ForeColor = vbBlack
    picMain.DrawWidth = 2
    'handles
    modStickGame.sLine X(1), Y(1), X(2), Y(2)
    modStickGame.sLine X(4), Y(4), X(3), Y(3)
    
    picMain.DrawWidth = 1
    modStickGame.sLine X(1), Y(1), X(3), Y(3)
    modStickGame.sLine X(3), Y(3), X(5), Y(5)
    modStickGame.sLine X(6), Y(6), X(7), Y(7)
    
    modStickGame.sLine X(1), Y(1), X(8), Y(8)
    modStickGame.sLine X(7), Y(7), X(9), Y(9) 'funnel
    modStickGame.sLine X(10), Y(10), X(8), Y(8)
    'modstickgame.sLine X(7), Y(7),X(8), Y(8)) 'funnel connection
    
    'sights
    modStickGame.sLine X(11), Y(11), X(12), Y(12)
    modStickGame.sLine X(13), Y(13), X(12), Y(12)
    modStickGame.sLine X(13), Y(13), X(14), Y(14)
    modStickGame.sLine X(15), Y(15), X(14), Y(14)
    modStickGame.sLine X(15), Y(15), X(16), Y(16)
    modStickGame.sLine X(11), Y(11), X(16), Y(16)
    
    
    
    If Reloading = False Then
        
        If Not (TotalMags(eWeaponTypes.RPG) = 0 And Stick(0).BulletsFired = 1) Or i > 0 Then
            
            If Stick(i).LastBullet + AutoReload_Delay / GetSticksTimeZone(i) < GetTickCount() Then
                'prevent from drawing rocket straight after firing
                'and before reload state received
                
                If Flip Then
                    tX = X(6)
                    tY = Y(6)
                Else
                    tX = X(5)
                    tY = Y(5)
                End If
                
                DrawRocket tX + GunLen / 1.2 * Sine(Stick(i).Facing - piD20), _
                           tY - GunLen / 1.2 * CoSine(Stick(i).Facing - piD20), _
                           Stick(i).Facing
                
                
            Else
                'draw barrier
                modStickGame.sLine X(5), Y(5), X(6), Y(6)
            End If
        Else
            'draw barrier
            modStickGame.sLine X(5), Y(5), X(6), Y(6)
        End If
    Else
        'draw barrier
        modStickGame.sLine X(5), Y(5), X(6), Y(6)
    End If
    
    
End If


Stick(i).GunPoint.X = X(5)
Stick(i).GunPoint.Y = Y(5)
Stick(i).CasingPoint.X = X(2)
Stick(i).CasingPoint.Y = Y(2)

End Sub

Private Sub DrawM249(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)
    

Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const top_offset As Single = BodyLen / 3

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If Facing > Pi Then
    Flip = True
    
    If Reloading Then Facing = 5 * Pi / 4
    
    Facing = Facing - Pi
    kY = -1
Else
    If Reloading Then Facing = pi3D4
    kY = 1
End If


'hand position
Hand1X = Stick(i).X + ArmLen / 4

If Flip Then
    Hand1Y = sY + 1.9 * BodyLen
'ElseIf StickiHasState(i, STICK_CROUCH) Then
    'Hand1Y = sY + HeadRadius + BodyLen / 6
Else
    Hand1Y = sY + top_offset
End If


DrawM2492 Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i

End Sub
Private Sub DrawM2492(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer)

Dim X(1 To 20) As Single, Y(1 To 20) As Single
Dim j As Integer
Dim SinFacing As Single, CosFacing As Single

SinFacing = Sine(Facing)
CosFacing = CoSine(Facing)

X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 2 * Sine(Facing + kY * pi3D4)
Y(2) = Y(1) - GunLen / 2 * CoSine(Facing + kY * pi3D4)

X(3) = X(2) + GunLen / 4 * SinFacing
Y(3) = Y(2) - GunLen / 4 * CosFacing

X(4) = X(1) + GunLen / 4 * SinFacing
Y(4) = Y(1) - GunLen / 4 * CosFacing
'end handle

'gap between handle and handy bit
X(5) = X(4) + GunLen / 4 * SinFacing
Y(5) = Y(4) - GunLen / 4 * CosFacing

X(6) = X(5) + GunLen / 6 * Sine(Facing + kY * piD2)
Y(6) = Y(5) - GunLen / 6 * CoSine(Facing + kY * piD2)

X(7) = X(6) + GunLen / 2 * SinFacing
Y(7) = Y(6) - GunLen / 2 * CosFacing

X(8) = X(5) + GunLen / 2 * SinFacing
Y(8) = Y(5) - GunLen / 2 * CosFacing

'bipod
X(9) = X(2) + GunLen * 1.2 * Sine(Facing + kY * piD10)
Y(9) = Y(2) - GunLen * 1.2 * CoSine(Facing + kY * piD10)

X(10) = X(2) + GunLen * 1.5 * Sine(Facing + kY * piD10)
Y(10) = Y(2) - GunLen * 1.5 * CoSine(Facing + kY * piD10)

'barrel
X(11) = X(8) + GunLen / 1.5 * SinFacing
Y(11) = Y(8) - GunLen / 1.5 * CosFacing

'sights
X(12) = X(8) + GunLen / 4 * SinFacing
Y(12) = Y(8) - GunLen / 4 * CosFacing

X(13) = X(12) + GunLen / 4 * Sine(Facing - kY * piD2)
Y(13) = Y(12) - GunLen / 4 * CoSine(Facing - kY * piD2)

'top bit
X(14) = X(8) + GunLen / 10 * Sine(Facing - kY * piD2)
Y(14) = Y(8) - GunLen / 10 * CoSine(Facing - kY * piD2)

'top handle
X(15) = X(14) + GunLen / 4 * Sine(Facing - kY * Pi)
Y(15) = Y(14) - GunLen / 4 * CoSine(Facing - kY * Pi)

X(16) = X(15) + GunLen / 6 * Sine(Facing - kY * piD2)
Y(16) = Y(15) - GunLen / 6 * CoSine(Facing - kY * piD2)

X(17) = X(16) + GunLen / 4 * Sine(Facing - kY * pi3D4)
Y(17) = Y(16) - GunLen / 4 * CoSine(Facing - kY * pi3D4)
'end handle

X(18) = X(15) + GunLen / 4 * Sine(Facing - kY * Pi)
Y(18) = Y(15) - GunLen / 4 * CoSine(Facing - kY * Pi)

X(18) = X(15) + GunLen / 4 * Sine(Facing - kY * Pi)
Y(18) = Y(15) - GunLen / 4 * CoSine(Facing - kY * Pi)

X(19) = X(1) + GunLen / 2 * Sine(Facing - kY * Pi)
Y(19) = Y(1) - GunLen / 2 * CoSine(Facing - kY * Pi)

X(20) = X(19) + GunLen / 4 * Sine(Facing + kY * piD2)
Y(20) = Y(19) - GunLen / 4 * CoSine(Facing + kY * piD2)

If Flip Then
    'flip image
    For j = 1 To 20
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    For j = 1 To 20
        Y(j) = 2 * sY - Y(j) + BodyLen * 2.2
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(1)
End If

Hand2X = X(6)
Hand2Y = Y(6)

'end calculation

If CanSeeStick(i) Then
    picMain.ForeColor = vbBlack
    picMain.DrawWidth = 1
    
    
    modStickGame.sLine X(1), Y(1), X(2), Y(2)
    modStickGame.sLine X(2), Y(2), X(3), Y(3)
    modStickGame.sLine X(3), Y(3), X(4), Y(4)
    modStickGame.sLine X(4), Y(4), X(5), Y(5)
    modStickGame.sLine X(5), Y(5), X(6), Y(6)
    modStickGame.sLine X(6), Y(6), X(7), Y(7)
    modStickGame.sLine X(7), Y(7), X(8), Y(8)
    modStickGame.sLine X(8), Y(8), X(9), Y(9)
    modStickGame.sLine X(8), Y(8), X(10), Y(10)
    modStickGame.sLine X(8), Y(8), X(11), Y(11)
    modStickGame.sLine X(12), Y(12), X(13), Y(13)
    modStickGame.sLine X(8), Y(8), X(14), Y(14)
    modStickGame.sLine X(14), Y(14), X(15), Y(15)
    modStickGame.sLine X(16), Y(16), X(15), Y(15)
    modStickGame.sLine X(18), Y(18), X(15), Y(15)
    modStickGame.sLine X(18), Y(18), X(19), Y(19)
    modStickGame.sLine X(20), Y(20), X(19), Y(19)
    modStickGame.sLine X(20), Y(20), X(1), Y(1)
    
    picMain.DrawWidth = 2 'handle
    modStickGame.sLine X(16), Y(16), X(17), Y(17)
    
    picMain.DrawWidth = 1
End If

Stick(i).GunPoint.X = X(11)
Stick(i).GunPoint.Y = Y(11)

Stick(i).CasingPoint.X = X(6)
Stick(i).CasingPoint.Y = Y(6)

End Sub

Private Sub DrawDEagle(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean ', JustShot As Boolean
Dim ArmLenDist As Single
Dim GTC As Long
Dim kY As Single

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)
'JustShot = (Stick(i).LastBullet + DEagle_Bullet_DelayD2 > getickcount())



If i = 0 Then
    GTC = GetTickCount()
    ArmLenDist = GetSticksTimeZone(i)
    
    If Stick(i).LastBullet + DEagle_Recoil_Time / ArmLenDist > GTC Then
        
        ArmLenDist = (GTC - Stick(0).LastBullet) * ArmLenDist / 8 + 5
        
    Else
        ArmLenDist = ArmLen
    End If
Else
    ArmLenDist = ArmLen
End If



'If JustShot Then
'    If Stick(i).ActualFacing < pi Then
'        facing =
'    Else
'Else
    If Facing > Pi Then
        Flip = True
        
        If Reloading Then Facing = pi5D4
        
        Facing = Facing - Pi
        kY = -1
    Else
        If Reloading Then Facing = pi3D4
        kY = 1
    End If
'End If

'If Reloading Then
'    'hand position
'    Hand1X = Stick(i).X + ArmLen / 2
'
'    If Flip Then
'        Hand1Y = sY + HeadRadius + BodyLen / 1.8
'    Else
'        Hand1Y = sY + HeadRadius2 + BodyLen / 2
'    End If
'Else
    'hand position
Hand1X = Stick(i).X + ArmLenDist * Sine(Facing)

If Flip Then
    Hand1Y = sY + ArmLen * 3 - ArmLenDist * CoSine(Facing)
Else
    Hand1Y = sY + ArmLen / 1.5 - ArmLenDist * CoSine(Facing)
End If
'End If

DrawDEagle2 Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, ArmLenDist

End Sub
Private Sub DrawDEagle2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    ArmLenDist As Single)

Dim X(1 To 10) As Single, Y(1 To 10) As Single
Dim j As Integer
Const HeadRadius2 = HeadRadius * 2 ', DEagle_Bullet_DelayD2 = DEagle_Bullet_Delay / 2

X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 2 * Sine(Facing)
Y(2) = Y(1) - GunLen / 2 * CoSine(Facing)

X(3) = X(2) + GunLen / 6 * Sine(Facing - kY * piD3) '60 deg
Y(3) = Y(2) - GunLen / 6 * CoSine(Facing - kY * piD3)

X(4) = X(3) + GunLen / 12 * Sine(Facing - kY * piD2)
Y(4) = Y(3) - GunLen / 12 * CoSine(Facing - kY * piD2)

X(5) = X(3) + GunLen / 10 * Sine(Facing - kY * Pi)
Y(5) = Y(3) - GunLen / 10 * CoSine(Facing - kY * Pi)

X(6) = X(3) + GunLen / 1.6 * Sine(Facing - kY * Pi)
Y(6) = Y(3) - GunLen / 1.6 * CoSine(Facing - kY * Pi)

X(6) = X(3) + GunLen / 1.6 * Sine(Facing - kY * Pi)
Y(6) = Y(3) - GunLen / 1.6 * CoSine(Facing - kY * Pi)

X(7) = X(6) + GunLen / 4 * Sine(Facing + kY * pi8D9)
Y(7) = Y(6) - GunLen / 4 * CoSine(Facing + kY * pi8D9)

X(8) = X(1) + GunLen / 6 * Sine(Facing - kY * Pi)
Y(8) = Y(1) - GunLen / 6 * CoSine(Facing - kY * Pi)

X(9) = X(8) + GunLen / 3 * Sine(Facing + kY * pi13D18)
Y(9) = Y(8) - GunLen / 3 * CoSine(Facing + kY * pi13D18)

X(10) = X(9) + GunLen / 6 * Sine(Facing)
Y(10) = Y(9) - GunLen / 6 * CoSine(Facing)



If Flip Then
    'flip image
    For j = 1 To 10
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        X(j) = 2 * sX - X(j)
    Next j
    
    For j = 1 To 10
        Y(j) = 2 * sY - Y(j) + BodyLen * 1.8
    Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Y(1)
End If

Hand2X = X(8)
Hand2Y = Y(8)

'end calculation
If CanSeeStick(i) Then
    picMain.ForeColor = MSilver
    modStickGame.sLine X(1), Y(1), X(2), Y(2)
    modStickGame.sLine X(2), Y(2), X(3), Y(3)
    
    modStickGame.sLine X(5), Y(5), X(6), Y(6)
    modStickGame.sLine X(6), Y(6), X(7), Y(7) 'vbYellow
    
    picMain.ForeColor = vbBlack
    modStickGame.sLine X(3), Y(3), X(4), Y(4)
    modStickGame.sLine X(4), Y(4), X(5), Y(5)
    modStickGame.sLine X(7), Y(7), X(8), Y(8)
    modStickGame.sLine X(8), Y(8), X(9), Y(9)
    modStickGame.sLine X(9), Y(9), X(10), Y(10)
    
    modStickGame.sLine X(10), Y(10), X(1), Y(1)
    
    'modstickgame.sLine X(), Y(),X(), Y())
    'If Stick(i).bSilenced Then
        'DrawSilencer X(2), Y(2), Facing + IIf(Stick(i).Facing > Pi, Pi, 0)
    'End If
    picMain.DrawWidth = 1
End If

Stick(i).GunPoint.X = X(2)
Stick(i).GunPoint.Y = Y(2)

Stick(i).CasingPoint.X = X(1)
Stick(i).CasingPoint.Y = Y(1)
End Sub

Private Sub DrawFlameThrower(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single

Const ArmLenDX = ArmLen / 3
Const BodyLenD2 = BodyLen / 2
Const BodyLenX2 = BodyLen * 2

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

sY = GetStickY(i)

If Facing > Pi Then
    Flip = True
    
    If Reloading Then
        Facing = pi3D4
    Else
        Facing = Facing - Pi
    End If
    
    kY = -1
    
    If StickiHasState(i, STICK_PRONE) Then
        sY = sY + BodyLen
    Else
        sY = sY + BodyLenX2
    End If
    
Else
    If Reloading Then Facing = piD4
    kY = 1
End If

Hand1X = Stick(i).X + ArmLenDX
If StickiHasState(i, STICK_PRONE) Then
    If Facing > Pi Then
        Hand1Y = sY - BodyLen
    Else
        Hand1Y = sY + BodyLenD2
    End If
Else
    Hand1Y = sY + BodyLen
End If

DrawFlamethrower2 Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading

End Sub
Private Sub DrawFlamethrower2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, Reloading As Boolean)

Dim j As Integer
Dim MB(1 To 10) As PointAPI
Dim FB(1 To 4) As PointAPI
'mb = MainBarrel
'fb = FuelBox



MB(1).X = Hand1X
MB(1).Y = Hand1Y

MB(2).X = MB(1).X + GunLen / 5 * Sine(Facing)
MB(2).Y = MB(1).Y - GunLen / 5 * CoSine(Facing)

MB(3).X = MB(2).X + GunLen / 3 * Sine(Facing - kY * piD4)
MB(3).Y = MB(2).Y - GunLen / 3 * CoSine(Facing - kY * piD4)

MB(4).X = MB(3).X + GunLen * Sine(Facing)
MB(4).Y = MB(3).Y - GunLen * CoSine(Facing)

MB(5).X = MB(4).X + GunLen / 6 * Sine(Facing - kY * piD4)
MB(5).Y = MB(4).Y - GunLen / 6 * CoSine(Facing - kY * piD4)

MB(6).X = MB(5).X + GunLen / 3 * Sine(Facing - kY * piD6)
MB(6).Y = MB(5).Y - GunLen / 3 * CoSine(Facing - kY * piD6)

MB(7).X = MB(6).X + GunLen / 10 * Sine(Facing - kY * piD2)
MB(7).Y = MB(6).Y - GunLen / 10 * CoSine(Facing - kY * piD2)

MB(8).X = MB(7).X + GunLen / 3 * Sine(Facing - kY * Pi)
MB(8).Y = MB(7).Y - GunLen / 3 * CoSine(Facing - kY * Pi)

MB(9).X = MB(8).X + GunLen / 3 * Sine(Facing + kY * pi3D4)
MB(9).Y = MB(8).Y - GunLen / 3 * CoSine(Facing + kY * pi3D4)

MB(10).X = MB(9).X + GunLen * Sine(Facing - kY * Pi)
MB(10).Y = MB(9).Y - GunLen * CoSine(Facing - kY * Pi)

If Not Reloading Then
    FB(1).X = MB(3).X '+ GunLen / 4 * sine(Facing)
    FB(1).Y = MB(3).Y '- GunLen / 4 * sine(Facing)
    
    FB(2).X = MB(3).X + GunLen / 2 * Sine(Facing) 'glDx = boxlen
    FB(2).Y = MB(3).Y - GunLen / 2 * CoSine(Facing)
    
    FB(3).X = FB(2).X + GunLen / 3 * Sine(Facing + kY * piD2) 'glDx = boxheight
    FB(3).Y = FB(2).Y - GunLen / 3 * CoSine(Facing + kY * piD2)
    
    FB(4).X = FB(3).X + GunLen / 4 * Sine(Facing - kY * Pi)
    FB(4).Y = FB(3).Y - GunLen / 4 * CoSine(Facing - kY * Pi)
End If

If Flip Then
    'flip image
    For j = 1 To 10
        'X(j) = X(j) - 2 * (X(j) - Stick(i).X)
        MB(j).X = 2 * sX - MB(j).X
        MB(j).Y = 2 * sY - MB(j).Y
    Next j
    
    If Not Reloading Then
        For j = 1 To 4
            FB(j).X = 2 * sX - FB(j).X
            FB(j).Y = 2 * sY - FB(j).Y
        Next j
    End If
    
    Hand1X = MB(1).X
    Hand1Y = MB(1).Y
End If

Hand2X = MB(9).X
Hand2Y = MB(9).Y

Stick(i).GunPoint.X = MB(7).X
Stick(i).GunPoint.Y = MB(7).Y
Stick(i).CasingPoint.X = MB(3).X
Stick(i).CasingPoint.Y = MB(3).Y

If CanSeeStick(i) Then
    picMain.ForeColor = vbBlack
    picMain.DrawWidth = 2
    
    modStickGame.sPoly MB, -1
    
    If Not Reloading Then
        modStickGame.sPoly FB, vbRed
    End If
End If
End Sub

Private Sub DrawAUG(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, ByRef A_Facing As Single)

'Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const ArmLenDX As Single = ArmLen / 2

A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = piD4 '1-below
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = pi3D4 'below is here
    kY = 1
End If

'hand position
Hand1X = sX + ArmLenDX '* Sine(A_Facing) '* 2D3
If Flip Then
    'If StickiHasState(i, stick_Crouch) Then
        'Hand1Y = sY + HeadRadius + BodyLen
    'Else
        Hand1Y = sY - HeadRadius * 1.5 - ArmLen * CoSine(A_Facing)
    'End If
Else
    Hand1Y = sY + HeadRadius + BodyLen / 6 - ArmLen * CoSine(A_Facing)
End If



DrawAUG2 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading


End Sub
Private Sub DrawAUG2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, Reloading As Boolean)

Dim j As Integer

Dim sX2 As Single, sY2 As Single
Const kGreen = 32768 '32768=rgb(0,128,0)

Dim pGrip(1 To 4) As PointAPI
Dim ptBarrel(1 To 4) As PointAPI
Dim ptMain(1 To 5) As PointAPI
Dim ptMag(1 To 4) As PointAPI
Dim ptSights(1 To 4) As PointAPI
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
'Dim Stock1X As Single, Stock1Y As Single, Stock2X As Single, Stock2Y As Single
Const BarrelLen As Single = GunLen / 3
Const GrayColour As Long = &H666666


Dim SinFacing As Single, CosFacing As Single
SinFacing = Sine(Facing)
CosFacing = CoSine(Facing)

'grip
pGrip(1).X = Hand1X
pGrip(1).Y = Hand1Y

pGrip(2).X = pGrip(1).X + GunLen / 3 * Sine(Facing + kY * pi3D4) 'grip height
pGrip(2).Y = pGrip(1).Y - GunLen / 3 * CoSine(Facing + kY * pi3D4)

pGrip(3).X = pGrip(2).X + GunLen / 4 * SinFacing
pGrip(3).Y = pGrip(2).Y - GunLen / 4 * CosFacing

pGrip(4).X = pGrip(1).X + GunLen / 4 * SinFacing
pGrip(4).Y = pGrip(1).Y - GunLen / 4 * CosFacing
'end grip

'green barrel part
ptBarrel(1).X = pGrip(4).X
ptBarrel(1).Y = pGrip(4).Y

ptBarrel(2).X = ptBarrel(1).X + GunLen / 2 * SinFacing 'GL/x = Green Len
ptBarrel(2).Y = ptBarrel(1).Y - GunLen / 2 * CosFacing

ptBarrel(3).X = ptBarrel(2).X + GunLen / 5 * Sine(Facing - kY * pi2d3) '100deg
ptBarrel(3).Y = ptBarrel(2).Y - GunLen / 5 * CoSine(Facing - kY * pi2d3)

ptBarrel(4).X = ptBarrel(1).X + GunLen / 4 * Sine(Facing - kY * piD2)
ptBarrel(4).Y = ptBarrel(1).Y - GunLen / 4 * CoSine(Facing - kY * piD2)
'end green barrel

'black barrel
Barrel1X = (ptBarrel(2).X + ptBarrel(3).X) / 2
Barrel1Y = (ptBarrel(2).Y + ptBarrel(3).Y) / 2
Barrel2X = Barrel1X + BarrelLen * SinFacing
Barrel2Y = Barrel1Y - BarrelLen * CosFacing


'main black bit
ptMain(1).X = ptBarrel(4).X
ptMain(1).Y = ptBarrel(4).Y

ptMain(2).X = ptMain(1).X - GunLen * SinFacing 'length that it goes back (to the stock)
ptMain(2).Y = ptMain(1).Y + GunLen * CosFacing

ptMain(3).X = ptMain(2).X + GunLen / 4 * Sine(Facing + kY * piD2)
ptMain(3).Y = ptMain(2).Y - GunLen / 4 * CoSine(Facing + kY * piD2)

ptMain(4).X = pGrip(1).X - GunLen / 3 * SinFacing
ptMain(4).Y = pGrip(1).Y + GunLen / 3 * CosFacing

ptMain(5).X = ptBarrel(1).X
ptMain(5).Y = ptBarrel(1).Y

'magazine
ptMag(1).X = pGrip(1).X - GunLen / 2 * SinFacing 'must be outside the below if block, so casing point is set
ptMag(1).Y = pGrip(1).Y + GunLen / 2 * CosFacing
If Not Reloading Then
    ptMag(2).X = ptMag(1).X - GunLen / 6 * SinFacing 'GL/x = Mag Width
    ptMag(2).Y = ptMag(1).Y + GunLen / 6 * CosFacing
    
    ptMag(3).X = ptMag(2).X + GunLen / 2 * Sine(Facing + kY * piD3)
    ptMag(3).Y = ptMag(2).Y - GunLen / 2 * CoSine(Facing + kY * piD3)
    
    ptMag(4).X = ptMag(1).X + GunLen / 2 * Sine(Facing + kY * piD3)
    ptMag(4).Y = ptMag(1).Y - GunLen / 2 * CoSine(Facing + kY * piD3)
End If

'sights
'bottom right
ptSights(1).X = pGrip(1).X + GunLen / 3 * Sine(Facing - kY * piD2)
ptSights(1).Y = pGrip(1).Y - GunLen / 3 * CoSine(Facing - kY * piD2)

'top right
ptSights(2).X = ptSights(1).X + GunLen / 6 * Sine(Facing - kY * piD4) 'GL/x = sight height
ptSights(2).Y = ptSights(1).Y - GunLen / 6 * CoSine(Facing - kY * piD4)

'top left
ptSights(3).X = ptSights(2).X - GunLen / 2 * SinFacing
ptSights(3).Y = ptSights(2).Y + GunLen / 2 * CosFacing

'bottom left
ptSights(4).X = ptSights(1).X - GunLen / 4 * SinFacing
ptSights(4).Y = ptSights(1).Y + GunLen / 4 * CosFacing




'#############
'Stock1X = CSng(ptMain(2).X)
'Stock1Y = CSng(ptMain(2).Y)
'Stock2X = CSng(ptMain(3).X)
'Stock2Y = CSng(ptMain(3).Y)
'#############


If Flip Then
    sX2 = 2 * sX
    sY2 = 2 * sY
    
    
    'flip image
    For j = 1 To 4
        pGrip(j).X = sX2 - pGrip(j).X
        pGrip(j).Y = sY2 - pGrip(j).Y
    Next j
    For j = 1 To 4
        ptBarrel(j).X = sX2 - ptBarrel(j).X
        ptBarrel(j).Y = sY2 - ptBarrel(j).Y
    Next j
    For j = 1 To 5
        ptMain(j).X = sX2 - ptMain(j).X
        ptMain(j).Y = sY2 - ptMain(j).Y
    Next j
    For j = 1 To 4
        ptMag(j).X = sX2 - ptMag(j).X
        ptMag(j).Y = sY2 - ptMag(j).Y
    Next j
    For j = 1 To 4
        ptSights(j).X = sX2 - ptSights(j).X
        ptSights(j).Y = sY2 - ptSights(j).Y
    Next j
    Barrel1X = sX2 - Barrel1X: Barrel1Y = sY2 - Barrel1Y
    Barrel2X = sX2 - Barrel2X: Barrel2Y = sY2 - Barrel2Y
    'Stock1X = sX2 - Stock1X: Stock1Y = sY2 - Stock1Y
    'Stock2X = sX2 - Stock2X: Stock2Y = sY2 - Stock2Y
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
End If



Hand2X = (ptBarrel(1).X + ptBarrel(2).X) / 2
Hand2Y = (ptBarrel(1).Y + ptBarrel(2).Y) / 2
'end calculation


Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y

Stick(i).CasingPoint.X = ptMag(1).X
Stick(i).CasingPoint.Y = ptMag(1).Y



'drawing
If CanSeeStick(i) Then
    picMain.DrawStyle = vbFSSolid
    picMain.ForeColor = vbBlack
    'picMain.FillColor = vbBlack 'not needed
    picMain.DrawWidth = 2
    
    
    'barrel
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
    
    picMain.DrawWidth = 1
    
    'sight stand
    modStickGame.sLine CLng(ptSights(1).X), _
                       CLng(ptSights(1).Y), _
                       CLng(ptSights(1).X + GunLen / 6 * Sine(Facing + piD2)), _
                       CLng(ptSights(1).Y - GunLen / 6 * CoSine(Facing + piD2))
    
    modStickGame.sLine CLng(ptSights(4).X), _
                       CLng(ptSights(4).Y), _
                       CLng(ptSights(4).X + GunLen / 6 * Sine(Facing + piD2)), _
                       CLng(ptSights(4).Y - GunLen / 6 * CoSine(Facing + piD2))
    
    
    
    
    modStickGame.sPoly pGrip, vbBlack
    modStickGame.sPoly ptSights, vbBlack
    If Not Reloading Then
        modStickGame.sPoly ptMag, vbBlack
    End If
    
    picMain.ForeColor = GrayColour
    modStickGame.sPoly ptBarrel, GrayColour
    modStickGame.sPoly ptMain, GrayColour
    
End If
End Sub

Private Sub DrawUSP(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, ByRef A_Facing As Single)

'Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim ArmLenDist As Single
Dim GTC As Long
Dim kY As Single

A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)


If i = 0 Then
    GTC = GetTickCount()
    ArmLenDist = GetSticksTimeZone(i)
    
    If Stick(i).LastBullet + USP_Recoil_Time / ArmLenDist > GTC Then
        
        ArmLenDist = (GTC - Stick(0).LastBullet) * ArmLenDist / 3 + 50
        
    Else
        ArmLenDist = ArmLen
    End If
Else
    ArmLenDist = ArmLen
End If



If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then A_Facing = pi5D4
    
    A_Facing = A_Facing - Pi
    kY = -1
Else
    If Reloading Then A_Facing = pi3D4
    kY = 1
End If


Hand1X = Stick(i).X + ArmLenDist * Sine(A_Facing)

If Flip Then
    Hand1Y = sY + ArmLen * 3 - ArmLenDist * CoSine(A_Facing)
Else
    Hand1Y = sY + ArmLen / 1.5 - ArmLenDist * CoSine(A_Facing)
End If
'End If

DrawUSP2 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, ArmLenDist

End Sub
Private Sub DrawUSP2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    ArmLenDist As Single)

Dim Pts(1 To 10) As PointAPI
Dim j As Integer
'Const HeadRadius2 = HeadRadius * 2

Pts(1).X = Hand1X
Pts(1).Y = Hand1Y

Pts(2).X = Pts(1).X + GunLen / 3 * Sine(Facing)
Pts(2).Y = Pts(1).Y - GunLen / 3 * CoSine(Facing)

Pts(3).X = Pts(2).X + GunLen / 6 * Sine(Facing - kY * piD3) '60 deg
Pts(3).Y = Pts(2).Y - GunLen / 6 * CoSine(Facing - kY * piD3)

Pts(4).X = Pts(3).X + GunLen / 12 * Sine(Facing - kY * piD2)
Pts(4).Y = Pts(3).Y - GunLen / 12 * CoSine(Facing - kY * piD2)

Pts(5).X = Pts(3).X + GunLen / 10 * Sine(Facing - kY * Pi)
Pts(5).Y = Pts(3).Y - GunLen / 10 * CoSine(Facing - kY * Pi)

Pts(6).X = Pts(3).X + GunLen / 1.6 * Sine(Facing - kY * Pi)
Pts(6).Y = Pts(3).Y - GunLen / 1.6 * CoSine(Facing - kY * Pi)

Pts(6).X = Pts(3).X + GunLen / 1.6 * Sine(Facing - kY * Pi)
Pts(6).Y = Pts(3).Y - GunLen / 1.6 * CoSine(Facing - kY * Pi)

Pts(7).X = Pts(6).X + GunLen / 4 * Sine(Facing + kY * pi8D9)
Pts(7).Y = Pts(6).Y - GunLen / 4 * CoSine(Facing + kY * pi8D9)

Pts(8).X = Pts(1).X + GunLen / 6 * Sine(Facing - kY * Pi)
Pts(8).Y = Pts(1).Y - GunLen / 6 * CoSine(Facing - kY * Pi)

Pts(9).X = Pts(8).X + GunLen / 3 * Sine(Facing + kY * pi13D18)
Pts(9).Y = Pts(8).Y - GunLen / 3 * CoSine(Facing + kY * pi13D18)

Pts(10).X = Pts(9).X + GunLen / 6 * Sine(Facing)
Pts(10).Y = Pts(9).Y - GunLen / 6 * CoSine(Facing)



If Flip Then
    'flip image
    For j = 1 To 10
        'pts(j) = pts(j) - 2 * (pts(j) - Stick(i).X)
        Pts(j).X = 2 * sX - Pts(j).X
        Pts(j).Y = 2 * sY - Pts(j).Y + BodyLen * 1.8
    Next j
    
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Pts(1).Y
End If

Hand2X = Pts(8).X
Hand2Y = Pts(8).Y

Stick(i).GunPoint.X = Pts(2).X
Stick(i).GunPoint.Y = Pts(2).Y

Stick(i).CasingPoint.X = Pts(1).X
Stick(i).CasingPoint.Y = Pts(1).Y

'end calculation
If CanSeeStick(i) Then
    
    'If Stick(i).bSilenced Then
        'DrawSilencer CSng(Pts(2).X), CSng(Pts(2).Y), Facing + IIf(Stick(i).Facing > Pi, Pi, 0)
    'End If
    
    picMain.ForeColor = vbBlack
    modStickGame.sPoly Pts, vbBlack
    
    
    picMain.DrawWidth = 1
End If


End Sub

Private Sub DrawG3(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, _
    ByRef A_Facing As Single)

'Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const ArmLenD3 = ArmLen / 3, HeadRadiusX1p5 = HeadRadius * 1.5


A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = pi2D5  '1-below
    Else
        A_Facing = A_Facing - Pi 'because of the flip
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = pi3D5 'below is here
    kY = 1
End If

'hand position
Hand1X = sX + ArmLenD3
If Flip Then
    Hand1Y = sY - HeadRadiusX1p5
Else
    Hand1Y = sY + HeadRadiusX1p5
End If



DrawG32 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading


End Sub
Private Sub DrawG32(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    bReloading As Boolean)


Dim j As Integer

Dim pGrip(1 To 6) As PointAPI, pMag(1 To 4) As PointAPI, _
    pBarrel(1 To 4) As PointAPI

Dim sX2 As Single, sY2 As Single
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Dim ForesightX As Single, ForesightY As Single
Dim Stock1X As Single, Stock1Y As Single, Stock2X As Single, Stock2Y As Single
Const BarrelLen = GunLen / 2

Const Grip_Width = GunLen * 2 / 3, Grip_Height = GunLen / 4, _
    Grip_Handle_Height = Grip_Height, Grip_Handle_Width = Grip_Width / 4, _
    Grip_Non_Handle_Width = Grip_Width - Grip_Handle_Width
Const Barrel_Width = GunLen, Barrel_Height = Grip_Height / 2
Const Mag_Width = Grip_Width / 3, Mag_Height = Grip_Height * 3 / 2
'Const Foresight_Offset = 30
Const Stock_Len = BarrelLen, Stock_Height = Grip_Height

'###############
Dim SineFacing As Single, CosFacing As Single
Dim sf_M_PiD2 As Single, cf_M_PiD2 As Single
Dim sf_P_PiD2 As Single, cf_P_PiD2 As Single
'   sine facing plus/minus pid2
SineFacing = Sine(Facing)
CosFacing = CoSine(Facing)
sf_M_PiD2 = Sine(Facing - kY * piD2)
cf_M_PiD2 = CoSine(Facing - kY * piD2)
sf_P_PiD2 = Sine(Facing + kY * piD2)
cf_P_PiD2 = CoSine(Facing + kY * piD2)
'###############

pGrip(1).X = Hand1X + Grip_Height * sf_M_PiD2
pGrip(1).Y = Hand1Y - Grip_Height * cf_M_PiD2

pGrip(2).X = pGrip(1).X + Grip_Width * SineFacing
pGrip(2).Y = pGrip(1).Y - Grip_Width * CosFacing

pGrip(3).X = pGrip(2).X + Grip_Height * sf_P_PiD2
pGrip(3).Y = pGrip(2).Y - Grip_Height * cf_P_PiD2

pGrip(4).X = pGrip(3).X - Grip_Non_Handle_Width * SineFacing
pGrip(4).Y = pGrip(3).Y + Grip_Non_Handle_Width * CosFacing

pGrip(5).X = pGrip(4).X + Grip_Handle_Height * sf_P_PiD2
pGrip(5).Y = pGrip(4).Y - Grip_Handle_Height * cf_P_PiD2

pGrip(6).X = pGrip(5).X - Grip_Handle_Width * SineFacing
pGrip(6).Y = pGrip(5).Y + Grip_Handle_Width * CosFacing


MakeSquarePoints pGrip(2).X, pGrip(2).Y, Barrel_Width, Barrel_Height, Facing, pBarrel, kY


If Not bReloading Then
    MakeSquarePoints pGrip(3).X - Mag_Width * SineFacing, _
                     pGrip(3).Y + Mag_Width * CosFacing, _
                     Mag_Width, Mag_Height, Facing, pMag, kY
    
End If


Barrel1X = pBarrel(3).X
Barrel1Y = pBarrel(3).Y
Barrel2X = Barrel1X + SineFacing * BarrelLen
Barrel2Y = Barrel1Y - CosFacing * BarrelLen

ForesightX = pBarrel(2).X '+ Foresight_Offset * sf_M_PiD2
ForesightY = pBarrel(2).Y '- Foresight_Offset * sf_M_PiD2

Stock1X = pGrip(1).X - Stock_Len * SineFacing
Stock1Y = pGrip(1).Y + Stock_Len * CosFacing

Stock2X = Stock1X - Stock_Height * sf_M_PiD2
Stock2Y = Stock1Y + Stock_Height * cf_M_PiD2

If Flip Then
    
    sX2 = sX * 2
    sY2 = sY * 2
    
    'flip image
    For j = 1 To 4
        pMag(j).X = sX2 - pMag(j).X
        pMag(j).Y = sY2 - pMag(j).Y
        
        pBarrel(j).X = sX2 - pBarrel(j).X
        pBarrel(j).Y = sY2 - pBarrel(j).Y
    Next j
    
    For j = 1 To 6
        pGrip(j).X = sX2 - pGrip(j).X
        pGrip(j).Y = sY2 - pGrip(j).Y
    Next j
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
    
    
    Barrel1X = sX2 - Barrel1X
    Barrel1Y = sY2 - Barrel1Y
    Barrel2X = sX2 - Barrel2X
    Barrel2Y = sY2 - Barrel2Y
    
    ForesightX = sX2 - ForesightX
    ForesightY = sY2 - ForesightY
    
    Stock1X = sX2 - Stock1X
    Stock1Y = sY2 - Stock1Y
    Stock2X = sX2 - Stock2X
    Stock2Y = sY2 - Stock2Y
End If


Hand2X = pGrip(3).X
Hand2Y = pGrip(3).Y
'end calculation


Stick(i).CasingPoint.X = pGrip(2).X
Stick(i).CasingPoint.Y = pGrip(2).Y
Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y


'drawing
If CanSeeStick(i) Then
    picMain.FillStyle = vbFSSolid
    picMain.ForeColor = vbBlack
    picMain.FillColor = vbBlack
    picMain.DrawWidth = 1
    
    
    'before polys get resized
    modStickGame.sLine Stock1X, Stock1Y, CSng(pGrip(1).X), CSng(pGrip(1).Y)
    modStickGame.sLine Stock1X, Stock1Y, Stock2X, Stock2Y 'not for this though
    
    modStickGame.sPoly pGrip, vbBlack
    modStickGame.sPoly pBarrel, vbBlack
    If Not bReloading Then
        modStickGame.sPoly pMag, vbBlack
    End If
    
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
    
    modStickGame.sCircle ForesightX, ForesightY, 20, vbBlack
    
    picMain.FillStyle = vbFSTransparent
End If

End Sub

Private Sub DrawAWM(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, _
    ByRef A_Facing As Single)

'Dim Facing As Single
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const ArmLenDX = ArmLen \ 3, ArmLenDX2 = ArmLen


A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = pi3D5  '1-below
    Else
        A_Facing = A_Facing - Pi 'because of the flip
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = pi2D5 'below is here
    kY = 1
End If

'hand position
Hand1X = sX + ArmLen * Sine(A_Facing) '* 2D3
If Flip Then
    If StickiHasState(i, STICK_PRONE) Then
        Hand1Y = sY - HeadRadius * 1.5 - ArmLen * CoSine(A_Facing) - ArmLenDX
    Else
        Hand1Y = sY - HeadRadius * 1.5 - ArmLen * CoSine(A_Facing)
    End If
Else
    If StickiHasState(i, STICK_PRONE) Then
        Hand1Y = sY + HeadRadius + BodyLen / 6 - ArmLen * CoSine(A_Facing) + ArmLenDX
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 6 - ArmLen * CoSine(A_Facing)
    End If
End If



DrawAWM2 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i ', Reloading, StickCol, _
    (GetTickCount() - Stick(i).LastNade) > (Nade_Arm_Time / GetSticksTimeZone(i))


End Sub
Private Sub DrawAWM2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer)


Dim j As Integer

Dim pMain(1 To 12) As PointAPI
Dim pSights(1 To 4) As PointAPI
Dim sX2 As Single, sY2 As Single, Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single

Const BarrelLen = GunLen

Dim SinFacing As Single, CosFacing As Single
SinFacing = Sine(Facing)
CosFacing = CoSine(Facing)


pMain(1).X = Hand1X
pMain(1).Y = Hand1Y

pMain(2).X = pMain(1).X + GunLen / 3 * Sine(Facing - kY * piD8)
pMain(2).Y = pMain(1).Y - GunLen / 3 * CoSine(Facing - kY * piD8)

pMain(3).X = pMain(2).X + GunLen * k2D3 * SinFacing
pMain(3).Y = pMain(2).Y - GunLen * k2D3 * CosFacing

pMain(4).X = pMain(3).X + GunLen / 6 * Sine(Facing - kY * piD2)
pMain(4).Y = pMain(3).Y - GunLen / 6 * CoSine(Facing - kY * piD2)

pMain(5).X = pMain(4).X - GunLen * k4D3 * SinFacing 'backwards
pMain(5).Y = pMain(4).Y + GunLen * k4D3 * CosFacing

pMain(6).X = pMain(5).X + GunLen / 8 * Sine(Facing + kY * pi5D8) 'backwards
pMain(6).Y = pMain(5).Y - GunLen / 8 * CoSine(Facing + kY * pi5D8)

pMain(7).X = pMain(6).X + GunLen / 3 * Sine(Facing + kY * pi17D16) 'backwards
pMain(7).Y = pMain(6).Y - GunLen / 3 * CoSine(Facing + kY * pi17D16)

pMain(8).X = pMain(7).X - GunLen / 3 * SinFacing 'backwards
pMain(8).Y = pMain(7).Y + GunLen / 3 * CosFacing

pMain(9).X = pMain(8).X + GunLen / 4 * Sine(Facing + kY * piD2)
pMain(9).Y = pMain(8).Y - GunLen / 4 * CoSine(Facing + kY * piD2)

pMain(10).X = pMain(9).X + GunLen / 8 * SinFacing
pMain(10).Y = pMain(9).Y - GunLen / 8 * CosFacing

pMain(11).X = pMain(10).X + GunLen / 8 * Sine(Facing - kY * piD2)
pMain(11).Y = pMain(10).Y - GunLen / 8 * CoSine(Facing - kY * piD2)


'pMain(12).X = pMain(1).X + GunLen / 8 * Sine(Facing + 9 * Pi / 16)
'pMain(12).Y = pMain(1).Y - GunLen / 8 * CoSine(Facing + 9 * Pi / 16)
pMain(12).X = pMain(11).X + GunLen / 2 * Sine(Facing + kY * piD18)
pMain(12).Y = pMain(11).Y - GunLen / 2 * CoSine(Facing + kY * piD18)




'sights
'bottom right
pSights(1).X = pMain(4).X + GunLen / 2 * Sine(Facing - Pi)
pSights(1).Y = pMain(4).Y - GunLen / 2 * CoSine(Facing - Pi)

'top right
pSights(2).X = pSights(1).X + GunLen / 6 * Sine(Facing - kY * piD2) 'GL/x = sight height
pSights(2).Y = pSights(1).Y - GunLen / 6 * CoSine(Facing - kY * piD2)

'top left
pSights(3).X = pSights(2).X - GunLen / 1.6 * SinFacing
pSights(3).Y = pSights(2).Y + GunLen / 1.6 * CosFacing

'bottom left
pSights(4).X = pSights(1).X - GunLen / 2 * SinFacing
pSights(4).Y = pSights(1).Y + GunLen / 2 * CosFacing


If Flip Then
    
    sX2 = sX * 2
    sY2 = sY * 2
    
    'flip image
    For j = 1 To 12
        pMain(j).X = sX2 - pMain(j).X
        pMain(j).Y = sY2 - pMain(j).Y
    Next j
    For j = 1 To 4
        pSights(j).X = sX2 - pSights(j).X
        pSights(j).Y = sY2 - pSights(j).Y
    Next j
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
    
    
    SinFacing = -SinFacing
    CosFacing = -CosFacing
End If


Barrel1X = (pMain(4).X + pMain(3).X) / 2
Barrel1Y = (pMain(4).Y + pMain(3).Y) / 2
Barrel2X = Barrel1X + SinFacing * BarrelLen
Barrel2Y = Barrel1Y - CosFacing * BarrelLen


Hand2X = pMain(1).X
Hand2Y = pMain(1).Y
'end calculation


Stick(i).CasingPoint.X = pMain(1).X
Stick(i).CasingPoint.Y = pMain(1).Y
Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y


'drawing
If CanSeeStick(i) Then
    picMain.DrawStyle = vbFSSolid
    picMain.ForeColor = vbBlack
    
    'barrel
    picMain.DrawWidth = 1
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
    
    
    modStickGame.sPoly_NoOutline pMain, vbBlack
    modStickGame.sPoly_NoOutline pSights, vbBlack
End If

End Sub

Private Sub DrawMP5(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, ByRef A_Facing As Single)

Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const BodyLenD3 = BodyLen / 3, HeadRadius5D4 = HeadRadius * 5 / 4, ArmLenDX = ArmLen / 2

A_Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, STICK_RELOAD)

If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = pi3D4
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = piD4
    kY = 1
End If


'hand position
Hand1X = sX + ArmLenDX

If StickiHasState(i, STICK_CROUCH) Then
    If Flip Then
        Hand1Y = sY - HeadRadius
    Else
        Hand1Y = sY + HeadRadius5D4
    End If
Else
    If Flip Then
        Hand1Y = sY - BodyLenD3
    Else
        Hand1Y = sY + BodyLenD3
    End If
End If

DrawMP52 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading



End Sub
Private Sub DrawMP52(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    Reloading As Boolean)

Dim pMain(0 To 14) As PointAPI, pMag(1 To 4) As PointAPI
Dim j As Integer
Dim sX2 As Single, sY2 As Single
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Const BarrelLen As Single = 30


pMain(0).X = Hand1X
pMain(0).Y = Hand1Y

pMain(1).X = pMain(0).X + GunLen / 4 * Sine(Facing + kY * pi3D4)
pMain(1).Y = pMain(0).Y - GunLen / 4 * CoSine(Facing + kY * pi3D4)

pMain(2).X = pMain(1).X + GunLen / 6 * Sine(Facing)
pMain(2).Y = pMain(1).Y - GunLen / 6 * CoSine(Facing)

pMain(3).X = pMain(2).X + GunLen / 4 * Sine(Facing - kY * piD4)
pMain(3).Y = pMain(2).Y - GunLen / 4 * CoSine(Facing - kY * piD4)

pMain(4).X = pMain(3).X + GunLen / 8 * Sine(Facing)
pMain(4).Y = pMain(3).Y - GunLen / 8 * CoSine(Facing)

pMain(5).X = pMain(4).X + GunLen / 5 * Sine(Facing - kY * piD8)
pMain(5).Y = pMain(4).Y - GunLen / 5 * CoSine(Facing - kY * piD8)

pMain(6).X = pMain(5).X + GunLen / 20 * Sine(Facing - kY * piD2)
pMain(6).Y = pMain(5).Y - GunLen / 20 * CoSine(Facing - kY * piD2)

pMain(7).X = pMain(6).X + GunLen / 2 * Sine(Facing - kY * piD16) 'length of main bottom bit
pMain(7).Y = pMain(6).Y - GunLen / 2 * CoSine(Facing - kY * piD16)

Barrel1X = pMain(7).X
Barrel1Y = pMain(7).Y

Barrel2X = Barrel1X + BarrelLen * Sine(Facing)
Barrel2Y = Barrel1Y - BarrelLen * CoSine(Facing)

pMain(8).X = pMain(7).X + GunLen / 6 * Sine(Facing - kY * piD2) 'top of front sight
pMain(8).Y = pMain(7).Y - GunLen / 6 * CoSine(Facing - kY * piD2)

pMain(9).X = pMain(8).X + GunLen / 8 * Sine(Facing + kY * pi3D4) 'GLDX must be smaller that GLDX from above
pMain(9).Y = pMain(8).Y - GunLen / 8 * CoSine(Facing + kY * pi3D4)

pMain(10).X = pMain(9).X - GunLen / 1.2 * Sine(Facing) 'back bit of straight line
pMain(10).Y = pMain(9).Y + GunLen / 1.2 * CoSine(Facing)

pMain(11).X = pMain(10).X + GunLen / 10 * Sine(Facing + kY * pi3D4)
pMain(11).Y = pMain(10).Y - GunLen / 10 * CoSine(Facing + kY * pi3D4)

pMain(12).X = pMain(11).X - GunLen / 3 * Sine(Facing) 'back of stock
pMain(12).Y = pMain(11).Y + GunLen / 3 * CoSine(Facing)

pMain(13).X = pMain(12).X + GunLen / 3 * Sine(Facing + kY * piD2)
pMain(13).Y = pMain(12).Y - GunLen / 3 * CoSine(Facing + kY * piD2)

pMain(14).X = pMain(13).X + GunLen / 8 * Sine(Facing - kY * piD4)
pMain(14).Y = pMain(13).Y - GunLen / 8 * CoSine(Facing - kY * piD4)


If Reloading = False Then
    pMag(1) = pMain(4)
    
    pMag(2).X = pMag(1).X + GunLen / 10 * Sine(Facing)
    pMag(2).Y = pMag(1).Y - GunLen / 10 * CoSine(Facing)
    
    pMag(4).X = pMag(1).X + GunLen / 3 * Sine(Facing + kY * piD6)
    pMag(4).Y = pMag(1).Y - GunLen / 3 * CoSine(Facing + kY * piD6)
    
    pMag(3).X = pMag(2).X + GunLen / 3 * Sine(Facing + kY * piD5) 'front point
    pMag(3).Y = pMag(2).Y - GunLen / 3 * CoSine(Facing + kY * piD5)
End If


If Flip Then
    sX2 = 2 * sX: sY2 = 2 * sY
    
    'flip image
    For j = 0 To 14
        pMain(j).X = sX2 - pMain(j).X
        pMain(j).Y = sY2 - pMain(j).Y
    Next j
    For j = 1 To 4
        pMag(j).X = sX2 - pMag(j).X
        pMag(j).Y = sY2 - pMag(j).Y
    Next j
    
    Barrel1X = sX2 - Barrel1X
    Barrel1Y = sY2 - Barrel1Y
    Barrel2X = sX2 - Barrel2X
    Barrel2Y = sY2 - Barrel2Y
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
End If


Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y
Stick(i).CasingPoint.X = pMain(4).X
Stick(i).CasingPoint.Y = pMain(4).Y

Hand2X = pMain(5).X
Hand2Y = pMain(5).Y
'end calculation


'drawing
If CanSeeStick(i) Then
    
    picMain.ForeColor = vbBlack
    modStickGame.sPoly pMain, vbBlack
    
    If Reloading = False Then modStickGame.sPoly pMag, vbBlack
    
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
    
End If

picMain.DrawWidth = 1
End Sub

Private Sub DrawMac10(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, A_Facing As Single)

Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const BodyLenD3 = BodyLen / 3, HeadRadius5D4 = HeadRadius * 5 / 4, ArmLenD2 = ArmLen / 2, ArmLenD3 = ArmLen / 3


Reloading = StickiHasState(i, STICK_RELOAD)
A_Facing = FixAngle(Stick(i).Facing)


If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = piD4
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = pi3D4
    kY = 1
End If


'hand position
Hand1X = sX + ArmLenD2

If Flip Then
    Hand1Y = sY - HeadRadius - ArmLenD2 * CoSine(A_Facing)
Else
    Hand1Y = sY + HeadRadius - ArmLenD2 * CoSine(A_Facing)
End If
'If StickiHasState(i, Stick_Crouch) Then
'    If Flip Then
'        Hand1Y = sY - HeadRadius
'    Else
'        Hand1Y = sY + HeadRadius5D4
'    End If
'Else
'    If Flip Then
'        Hand1Y = sY - BodyLenD3
'    Else
'        Hand1Y = sY + BodyLenD3
'    End If
'End If

DrawMac102 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i, Reloading



End Sub
Private Sub DrawMac102(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer, _
    Reloading As Boolean)


Dim pHBar(1 To 4) As PointAPI, pVBar(1 To 4) As PointAPI, pMag(1 To 4) As PointAPI
Dim j As Integer
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Dim sX2 As Single, sY2 As Single

Const BarrelLen As Single = 100
Const VBar_Width As Single = GunLen / 10, _
      VBar_Height As Single = GunLen / 3
Const HBar_Width As Single = GunLen, _
      HBar_Height As Single = GunLen / 8
Const HBar_WidthD3 = HBar_Width / 3
Const Mag_Width = VBar_Width / 3, _
      Mag_Height = VBar_Height * 3 / 2


MakeSquarePoints Hand1X, Hand1Y, VBar_Width, VBar_Height, Facing, pVBar(), kY
MakeSquarePoints Hand1X + HBar_WidthD3 * Sine(Facing - Pi), _
                 Hand1Y - HBar_WidthD3 * CoSine(Facing - Pi), _
                 HBar_Width, HBar_Height, Facing, pHBar(), kY




If Not Reloading Then
    MakeSquarePoints pVBar(3).X, pVBar(3).Y, Mag_Width, Mag_Height, Facing, pMag, kY
End If


Barrel1X = (pHBar(2).X + pHBar(3).X) / 2
Barrel1Y = (pHBar(2).Y + pHBar(3).Y) / 2
Barrel2X = Barrel1X + BarrelLen * Sine(Facing)
Barrel2Y = Barrel1Y - BarrelLen * CoSine(Facing)


If Flip Then
    sX2 = 2 * sX: sY2 = 2 * sY
    
    For j = 1 To 4
        pHBar(j).X = sX2 - pHBar(j).X
        pHBar(j).Y = sY2 - pHBar(j).Y
        
        pVBar(j).X = sX2 - pVBar(j).X
        pVBar(j).Y = sY2 - pVBar(j).Y
        
        pMag(j).X = sX2 - pMag(j).X
        pMag(j).Y = sY2 - pMag(j).Y
    Next j
    
    Barrel1X = sX2 - Barrel1X
    Barrel1Y = sY2 - Barrel1Y
    Barrel2X = sX2 - Barrel2X
    Barrel2Y = sY2 - Barrel2Y
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
End If


Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y
Stick(i).CasingPoint.X = pVBar(1).X
Stick(i).CasingPoint.Y = pVBar(1).Y

Hand2X = pHBar(3).X
Hand2Y = pHBar(3).Y
'end calculation


'drawing
If CanSeeStick(i) Then
    
    picMain.ForeColor = vbBlack
    modStickGame.sPoly pVBar, vbBlack
    modStickGame.sPoly pHBar, vbBlack
    
    If Not Reloading Then modStickGame.sPoly pMag, vbBlack
    
    picMain.DrawWidth = 2
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
    'modStickGame.sCircle Hand2X, Hand2X, 80, vbBlack
End If

picMain.DrawWidth = 1
End Sub

Private Sub DrawSPAS(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single, A_Facing As Single)

Dim Flip As Boolean, Reloading As Boolean
Dim kY As Single
Const BodyLenD3 = BodyLen / 3, HeadRadius5D4 = HeadRadius * 5 / 4, ArmLenD2 = ArmLen / 2


Reloading = StickiHasState(i, STICK_RELOAD)
A_Facing = FixAngle(Stick(i).Facing)


If A_Facing > Pi Then
    Flip = True
    
    If Reloading Then
        A_Facing = piD4
    Else
        A_Facing = A_Facing - Pi
    End If
    
    kY = -1
Else
    If Reloading Then A_Facing = pi3D4
    kY = 1
End If


'hand position
Hand1X = sX + ArmLenD2

If StickiHasState(i, STICK_CROUCH) Then
    If Flip Then
        Hand1Y = sY - HeadRadius
    Else
        Hand1Y = sY + HeadRadius5D4
    End If
Else
    If Flip Then
        Hand1Y = sY - BodyLenD3
    Else
        Hand1Y = sY + BodyLenD3
    End If
End If

DrawSPAS2 A_Facing, Hand1X, Hand1Y, Hand2X, Hand2Y, kY, Flip, sX, sY, i



End Sub
Private Sub DrawSPAS2(Facing As Single, _
    Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, _
    kY As Single, Flip As Boolean, sX As Single, sY As Single, i As Integer)


Dim j As Integer
Dim sX2 As Single, sY2 As Single
Dim BarrelStart(1 To 2) As PointAPI 'top and bottom
Dim BarrelEnd(1 To 2) As PointAPI 'top and bottom
Dim pMain(1 To 4) As PointAPI, pStock(1 To 4) As PointAPI
Dim Handle1X As Single, Handle1Y As Single ', Handle2X As Single, Handle2Y As Single
Dim ForesightX As Single, ForesightY As Single
Dim RearSightX As Single, RearSightY As Single

Const Barrel1Len As Single = 100, Barrel2Len As Single = 80
Const Stock_Width As Single = GunLen / 10, _
      Stock_Height As Single = GunLen / 2, _
      Stock_Angle As Single = piD4
Const Main_Width As Single = GunLen, _
      Main_Height As Single = GunLen / 8
Const HandleLen = GunLen / 6, HandleStartLen = Main_Width * 2 / 3
Const Main_Height_Plus_Foresight_Offset = Main_Height + 10

Dim SineFacing As Single, CosFacing As Single

SineFacing = Sine(Facing)
CosFacing = CoSine(Facing)


MakeSquarePoints Hand1X + Main_Height * Sine(Facing - kY * piD2), _
                 Hand1Y - Main_Height * CoSine(Facing - kY * piD2), _
                 Main_Width, Main_Height, Facing, pMain(), kY

MakeSquarePoints pMain(1).X, pMain(1).Y, Stock_Width, Stock_Height, Facing + kY * Stock_Angle, pStock(), kY



BarrelStart(1).X = pMain(2).X
BarrelStart(1).Y = pMain(2).Y

BarrelStart(2).X = pMain(3).X
BarrelStart(2).Y = pMain(3).Y

BarrelEnd(1).X = BarrelStart(1).X + Barrel1Len * SineFacing
BarrelEnd(1).Y = BarrelStart(1).Y - Barrel1Len * CosFacing

BarrelEnd(2).X = BarrelStart(2).X + Barrel2Len * SineFacing
BarrelEnd(2).Y = BarrelStart(2).Y - Barrel2Len * CosFacing


Handle1X = pMain(4).X + HandleStartLen * SineFacing
Handle1Y = pMain(4).Y - HandleStartLen * CosFacing
'Handle2X = Handle1X + HandleLen * SineFacing
'Handle2Y = Handle1Y - HandleLen * CosFacing

ForesightX = Handle1X + HandleLen * SineFacing + Main_Height_Plus_Foresight_Offset * Sine(Facing - kY * piD2)
ForesightY = Handle1Y - HandleLen * CosFacing - Main_Height_Plus_Foresight_Offset * CoSine(Facing - kY * piD2)

RearSightX = pMain(1).X + HandleLen * Sine(Facing - kY * piD8)
RearSightY = pMain(1).Y - HandleLen * CoSine(Facing - kY * piD8)


If Flip Then
    sX2 = 2 * sX: sY2 = 2 * sY
    
    For j = 1 To 4
        pMain(j).X = sX2 - pMain(j).X
        pMain(j).Y = sY2 - pMain(j).Y
        
        
        pStock(j).X = sX2 - pStock(j).X
        pStock(j).Y = sY2 - pStock(j).Y
    Next j
    
    
    For j = 1 To 2
        BarrelStart(j).X = sX2 - BarrelStart(j).X
        BarrelStart(j).Y = sY2 - BarrelStart(j).Y
        
        BarrelEnd(j).X = sX2 - BarrelEnd(j).X
        BarrelEnd(j).Y = sY2 - BarrelEnd(j).Y
    Next j
    
    
    Handle1X = sX2 - Handle1X
    Handle1Y = sY2 - Handle1Y
'    Handle2X = sX2 - Handle2X
'    Handle2Y = sY2 - Handle2Y
    
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
    
    
    ForesightX = sX2 - ForesightX
    ForesightY = sY2 - ForesightY
    
    RearSightX = sX2 - RearSightX
    RearSightY = sY2 - RearSightY
End If


Stick(i).GunPoint.X = BarrelEnd(1).X
Stick(i).GunPoint.Y = BarrelEnd(1).Y
Stick(i).CasingPoint.X = Handle1X
Stick(i).CasingPoint.Y = Handle1Y


Hand2X = Handle1X
Hand2Y = Handle1Y
'end calculation


'drawing
If CanSeeStick(i) Then
    
    picMain.ForeColor = vbBlack
    modStickGame.sPoly pMain, vbBlack
    modStickGame.sPoly pStock, vbBlack
    
    
    picMain.DrawWidth = 2
    modStickGame.sLine CSng(BarrelStart(1).X), CSng(BarrelStart(1).Y), CSng(BarrelEnd(1).X), CSng(BarrelEnd(1).Y)
    modStickGame.sLine CSng(BarrelStart(2).X), CSng(BarrelStart(2).Y), CSng(BarrelEnd(2).X), CSng(BarrelEnd(2).Y)
    
    
    picMain.FillStyle = vbFSSolid
    picMain.FillColor = vbBlack
    modStickGame.sCircle ForesightX, ForesightY, 20, vbBlack
    modStickGame.sCircle RearSightX, RearSightY, 20, vbBlack
    picMain.FillStyle = vbFSTransparent
    
    'picMain.ForeColor = MGrey
    'modStickGame.sLine Handle1X, Handle1Y, Handle2X, Handle2Y
End If

picMain.DrawWidth = 1
End Sub

Private Sub DrawStaticG3(sX As Single, sY As Single)

Dim pGrip(1 To 6) As PointAPI, pMag(1 To 4) As PointAPI, _
    pBarrel(1 To 4) As PointAPI

Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Dim ForesightX As Single, ForesightY As Single
Dim Stock1X As Single, Stock1Y As Single, Stock2X As Single, Stock2Y As Single
Const BarrelLen = GunLen / 2

Const Grip_Width = GunLen * 2 / 3, Grip_Height = GunLen / 4, _
    Grip_Handle_Height = Grip_Height, Grip_Handle_Width = Grip_Width / 4, _
    Grip_Non_Handle_Width = Grip_Width - Grip_Handle_Width
Const Barrel_Width = GunLen, Barrel_Height = Grip_Height / 2
Const Mag_Width = Grip_Width / 3, Mag_Height = Grip_Height * 3 / 2
'Const Foresight_Offset = 30
Const Stock_Len = BarrelLen, Stock_Height = Grip_Height


Const Facing = piD2


pGrip(1).X = sX + Grip_Height
pGrip(1).Y = sY - Mag_Height

pGrip(2).X = pGrip(1).X + Grip_Width
pGrip(2).Y = pGrip(1).Y

pGrip(3).X = pGrip(2).X
pGrip(3).Y = pGrip(2).Y + Grip_Height

pGrip(4).X = pGrip(3).X - Grip_Non_Handle_Width
pGrip(4).Y = pGrip(3).Y

pGrip(5).X = pGrip(4).X
pGrip(5).Y = pGrip(4).Y + Grip_Handle_Height

pGrip(6).X = pGrip(5).X - Grip_Handle_Width
pGrip(6).Y = pGrip(5).Y


MakeSquarePoints pGrip(2).X, pGrip(2).Y, Barrel_Width, Barrel_Height, Facing, pBarrel, 1


MakeSquarePoints pGrip(3).X - Mag_Width, _
                     pGrip(3).Y, _
                     Mag_Width, Mag_Height, Facing, pMag, 1


Barrel1X = pBarrel(3).X
Barrel1Y = pBarrel(3).Y
Barrel2X = Barrel1X + BarrelLen
Barrel2Y = Barrel1Y

ForesightX = pBarrel(2).X '+ Foresight_Offset * sf_M_PiD2
ForesightY = pBarrel(2).Y '- Foresight_Offset * sf_M_PiD2

Stock1X = pGrip(1).X - Stock_Len
Stock1Y = pGrip(1).Y

Stock2X = Stock1X
Stock2Y = Stock1Y + Stock_Height
'end calculation


'drawing
picMain.FillStyle = vbFSSolid
picMain.ForeColor = vbBlack
picMain.FillColor = vbBlack
picMain.DrawWidth = 1


'before polys get resized
modStickGame.sLine Stock1X, Stock1Y, CSng(pGrip(1).X), CSng(pGrip(1).Y)
modStickGame.sLine Stock1X, Stock1Y, Stock2X, Stock2Y 'not for this though

modStickGame.sPoly pGrip, vbBlack
modStickGame.sPoly pBarrel, vbBlack
modStickGame.sPoly pMag, vbBlack

modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y

modStickGame.sCircle ForesightX, ForesightY, 20, vbBlack

picMain.FillStyle = vbFSTransparent

End Sub

'####################################################################################################
'####################################################################################################
'####################################################################################################

Private Sub MakeSquarePoints(ByVal startX As Long, ByVal startY As Long, _
    sgWidth As Single, sgHeight As Single, Facing As Single, Pts() As PointAPI, kY As Single)


Pts(1).X = startX
Pts(1).Y = startY

Pts(2).X = Pts(1).X + sgWidth * Sine(Facing)
Pts(2).Y = Pts(1).Y - sgWidth * CoSine(Facing)

Pts(3).X = Pts(2).X + sgHeight * Sine(Facing + kY * piD2)
Pts(3).Y = Pts(2).Y - sgHeight * CoSine(Facing + kY * piD2)

Pts(4).X = Pts(3).X + sgWidth * Sine(Facing - kY * Pi)
Pts(4).Y = Pts(3).Y - sgWidth * CoSine(Facing - kY * Pi)

End Sub

Private Function StickIsMoving(ByVal i As Integer) As Boolean

'Dim ST As Integer
'
'ST = Stick(i).State
'
'StickIsMoving = ( _
'    ((ST And stick_Left) = stick_Left) Or _
'    ((ST And stick_Right) = stick_Right))

StickIsMoving = (Stick(i).Speed > 0.5)

End Function

Private Sub AddNade(X As Single, Y As Single, Heading As Single, _
    Speed As Single, iStick As Integer, colour As Long, iType As eNadeTypes, _
    Optional IsRPG As Boolean, Optional bIsMartyrdom As Boolean) ', Optional Sticki As Integer)

ReDim Preserve Nade(NumNades)

With Nade(NumNades)
    '.Decay = GetTickCount() + Nade_Time / GetSticksTimeZone(iStick)
    .Start_Time = GetTickCount()
    
    .Heading = Heading
    .Speed = Speed
    .X = X
    .Y = Y
    .OwnerID = Stick(iStick).ID
    .IsRPG = IsRPG
    .colour = colour
    .iType = iType
    .bIsMartyrdomNade = bIsMartyrdom
End With

If IsRPG Then
#If Hack_Ammo = False Then
    If iStick = 0 Or Stick(iStick).IsBot Then
        If Stick(iStick).WeaponType = RPG Then 'might be a chopper
            Stick(iStick).BulletsFired = Stick(iStick).BulletsFired + 1
        End If
    End If
#End If
    
    If PointHearableOnSticksScreen(X, Y, 0) Then
        modAudio.PlayWeaponSound_Panned RPG, 0
    End If
    
Else
    If PointHearableOnSticksScreen(X, Y, 0) Then
        modAudio.PlayNadeThrow GetRelPan(X)
    End If
End If

NumNades = NumNades + 1

End Sub

Private Sub AddMine(X As Single, Y As Single, OwnerID As Integer, colour As Long, Heading As Single, Speed As Single)
Dim i As Integer

ReDim Preserve Mine(NumMines)

With Mine(NumMines)
    .X = X
    .Y = Y
    .OwnerID = OwnerID
    .colour = colour
    .Heading = Heading
    .Speed = Speed
    
    
    .ID = GenerateMineID()
    
    
    i = FindStick(OwnerID)
    If i > -1 Then
        AddInfoCirc .X, .Y, Stick(i).colour
    End If
End With

NumMines = NumMines + 1

End Sub

Private Sub AddMag(X As Single, Y As Single, Speed As Single, Heading As Single, Optional vMagType As eMagTypes = -1)

If vMagType = -1 Then Exit Sub
If modStickGame.cg_Magazines = False Then Exit Sub

ReDim Preserve Mag(NumMags)

With Mag(NumMags)
    .X = X
    .Y = Y
    .Speed = Speed
    .Heading = Heading
    
    .Decay = GetTickCount() + Mag_Decay
    .iMagType = vMagType
End With

NumMags = NumMags + 1

End Sub

Private Sub AddDeadStick(X As Single, Y As Single, colour As Long, bFacingRight As Boolean, _
    bFlamed As Boolean, bSniper As Boolean, _
    Speed As Single, Heading As Single, bIsMyStick As Boolean)

Dim i As Integer


If modStickGame.cg_DeadSticks Then
    
    If bIsMyStick Then
        For i = 0 To NumDeadSticks - 1
            DeadStick(i).bIsMe = False
        Next i
    End If
    
    
    
    ReDim Preserve DeadStick(NumDeadSticks)
    
    With DeadStick(NumDeadSticks)
        .X = X
        .Y = Y
        
        .Decay = GetTickCount() + DeadStickTime
        
        .Heading = Heading
        
        .bIsMe = bIsMyStick
        
        If Speed > MAX_DeadStick_And_StaticWeap_Speed Then
            .Speed = MAX_DeadStick_And_StaticWeap_Speed
        Else
            .Speed = Speed '20
        End If
        
        .bFacingRight = bFacingRight
        
        .bFlamed = bFlamed
        
        If bFlamed Then
            AddSmokeNadeTrail X + PM_Rnd * ArmLen, Y + BodyLen + HeadRadius, True
            AddSmokeNadeTrail X + PM_Rnd * ArmLen, Y + BodyLen + HeadRadius, True
            .colour = vbBlack
        ElseIf bSniper Then
            .colour = Grass_Col
        Else
            .colour = colour
        End If
        
    End With
    
    NumDeadSticks = NumDeadSticks + 1
End If

End Sub

Private Sub AddDeadChopper(X As Single, Y As Single, colour As Long, iOwner As Integer)

ReDim Preserve DeadChopper(NumDeadChoppers)

With DeadChopper(NumDeadChoppers)
    .X = X
    .Y = Y
    .colour = colour
    .Decay = GetTickCount() + DeadChopperTime
    .Speed = 0
    .iOwner = iOwner
End With

NumDeadChoppers = NumDeadChoppers + 1

End Sub

Private Sub AddBlood(X As Single, Y As Single, Heading As Single)

If modStickGame.cg_Blood Then
    ReDim Preserve Blood(NumBlood)
    
    With Blood(NumBlood)
        .Decay = GetTickCount() + Blood_Time / GetTimeZoneAdjust(X, Y) - 500 * PM_Rnd()
        .Heading = Heading + PM_Rnd * piD8
        .Speed = 70 + PM_Rnd() * 30
        .X = X
        .Y = Y
    End With
    
    NumBlood = NumBlood + 1
End If

End Sub

Private Sub AddSpark(X As Single, Y As Single, Heading As Single, Speed As Single)

If modStickGame.cg_Sparks Then
    'ReDim Preserve Spark(NumSparks)
    
    With Spark(NumSparks)
        .Decay = GetTickCount() + Spark_Time / GetTimeZoneAdjust(X, Y) - 100 * Rnd()
        .Heading = Heading
        .Speed = Speed
        .X = X
        .Y = Y
        ResetTimeLong .LastReduction, Spark_Speed_Reduction_Delay
    End With
    
    NumSparks = NumSparks + 1
End If

If NumSparks >= Max_Sparks Then
    RemoveSpark 0
End If

End Sub

Private Sub AddBullet(X As Single, Y As Single, Speed As Single, Heading As Single, _
    ByVal Damage As Single, iStick As Integer)

Dim sgTmp As Single

Const BulletTrail_Smoke_DelayD2 = BulletTrail_Smoke_Delay / 2, muzzleSmokeOffset = BULLET_SPEED / 2

'ReDim Preserve Bullet(NumBullets)


With Bullet(NumBullets)
    'NEED TO SET ALL VALUES, SINCE THE BULLET AT index MAY HAVE BEEN ALREADY USED
    .LastDiffract = 0
    .LastGravity = 0
    .LastSmoke = 0
    .bHeadingChanged = False
    .bHadCircleBlast = False
    
    
    .Heading = Heading
    '.Facing = Stick(iStick).Facing
    .Speed = Speed
    .X = X
    .Y = Y
    .OwnerIndex = iStick
    
    
    
    .bSniperBullet = WeaponIsSniper(Stick(iStick).WeaponType) 'bSnipe
    .bShotgunBullet = WeaponIsShotgun(Stick(iStick).WeaponType)
    .bDEagleBullet = (Stick(iStick).WeaponType = DEagle)
    
    If Stick(iStick).WeaponType = Chopper Then
        .bChopperBullet = True
        .bTracer = False
    Else
        .bChopperBullet = False
        .bTracer = ((Stick(iStick).BulletsFired Mod 15) = 0) And .bShotgunBullet = False
    End If
    
    
    .LastSmoke = GetTickCount() + BulletTrail_Smoke_DelayD2
    
    
    sgTmp = Rnd() * Speed / 20 + 10
    
'    If Stick(iStick).WeaponType = M82 Then
'        '.Decay = GetTickCount() + Bullet_Decay / GetTimeZoneAdjust(Stick(iStick).X, Stick(iStick).Y) - 100 * Rnd()
'    'Else
'        '.Decay = GetTickCount() + Bullet_Decay * 2 / GetTimeZoneAdjust(Stick(iStick).X, Stick(iStick).Y)
'
'        'y + + Speed * cosine(Heading)
'
'        AddSmokeGroup X, Y, 3, sgTmp, Heading - piD8 - 0.5 * Rnd(), True
'        AddSmokeGroup X, Y, 3, sgTmp, Heading + piD8 + 0.5 * Rnd(), True
'
'    End If
    If .bSniperBullet Then
        AddSmokeGroup X, Y, 3, sgTmp, Heading - piD4 - 0.5 * Rnd(), True
        AddSmokeGroup X, Y, 3, sgTmp, Heading + piD4 + 0.5 * Rnd(), True
        
        AddSmokeGroup X, Y, 3, sgTmp, Heading, True
        AddSmokeGroup X, Y, 3, sgTmp + 10, Heading, True
    End If
    
    
    'AddCirc X, Y, 5000, 1, vbYellow, 1000, True
    
    
    If Stick(iStick).Perk = pStoppingPower Then
        Damage = Damage * StoppingPowerIncrease
    ElseIf Stick(iStick).Perk = pSniper Then
        If Stick(iStick).WeaponType = G3 Then
            Damage = Damage * G3_Sniper_Damage_Factor
        End If
    End If
    
    
    .bSilenced = Stick(iStick).bSilenced
    
    
    If Stick(iStick).bSilenced Then
        .Damage = Damage * Bullet_Silenced_Damage_Factor
    Else
        .Damage = Damage
        
        '.bSilenced = (Stick(iStick).BulletsFired Mod 5)
    End If
    
    
    If modStickGame.cg_Smoke Then
        If Rnd() > 0.6 Then
            If Not .bSniperBullet Then
                AddSmokeGroup X + muzzleSmokeOffset * Sine(.Heading), Y - muzzleSmokeOffset * CoSine(.Heading), 5, sgTmp / 2, Heading
            End If
        End If
    End If
    
    
    If Not WeaponIsShotgun(Stick(iStick).WeaponType) Then
        'If Stick(iStick).WeaponType = Chopper Then
            'AddCasing Stick(iStick).CasingPoint.X, Stick(iStick).CasingPoint.Y, _
                Heading, .bSniperBullet Or .bChopperBullet, iStick
            
            'Casing(NumCasings - 1).Speed = Casing(NumCasings - 1).Speed / Chopper_Casing_Reduction
        'Else
            AddCasing Stick(iStick).CasingPoint.X, Stick(iStick).CasingPoint.Y, _
                Heading, .bSniperBullet Or .bChopperBullet, iStick
        'End If
    End If
    
    
'    With Casing(NumCasings - 1)
'        .Speed = Stick(iStick).Speed
'        .Heading = Stick(iStick).Heading
'    End With
    
End With

NumBullets = NumBullets + 1

#If Hack_Ammo Then
If iStick > 0 Then
#End If
Stick(iStick).BulletsFired = Stick(iStick).BulletsFired + 1
#If Hack_Ammo Then
End If
#End If
'End If

'If kBurstBullets(Stick(iStick).WeaponType) > 0 Then
If Stick(iStick).Burst_Bullets Then
    Stick(iStick).BulletsFired2 = Stick(iStick).BulletsFired2 + 1
End If

End Sub

'Private Sub AddMuzzleFlash(X As Single, Y As Single, Facing As Single)
'ReDim Preserve MFlash(NumMFlashes)
'
'With MFlash(NumMFlashes)
'    .Decay = GetTickCount() + MFlash_Time / GetTimeZoneAdjust
'    .Facing = Facing
'    .X = X
'    .Y = Y
'End With
'
'NumMFlashes = NumMFlashes + 1
'End Sub

Private Sub AddCasing(X As Single, Y As Single, Facing As Single, bSnipe As Boolean, iStick As Integer)

If modStickGame.cg_Casing Then
    'ReDim Preserve Casing(NumCasings)
    
    With Casing(NumCasings)
        .Decay = GetTickCount() + Casing_Time / GetTimeZoneAdjust(X, Y)
        
        .Facing = Facing
        
        '.Speed = IIf(bSnipe, 80, 50)
        '.Heading = IIf(Stick(iStick).Facing < Pi, pi7D4, piD4)
        
        AddVectors IIf(bSnipe, 80, 50) + PM_Rnd() * 8, IIf(Stick(iStick).Facing < Pi, pi7D4, piD4), _
                   Stick(iStick).Speed, Stick(iStick).Heading, .Speed, .Heading
        
        
        .X = X
        .Y = Y
        .bSniperCasing = bSnipe
        
        ResetTimeLong .LastGravity, Gravity_Delay
    End With
    
    NumCasings = NumCasings + 1
    
    If NumCasings >= Max_Casings Then
        RemoveCasing 0
    End If
End If

End Sub

Private Sub AddStaticWeapon(X As Single, Y As Single, vWeapon As eWeaponTypes)
'Const Static_Weapon_Y_Inc = 300

If NumStaticWeapons > Max_Static_Weapons Then
    RemoveStaticWeapon 0
End If


ReDim Preserve StaticWeapon(NumStaticWeapons)

With StaticWeapon(NumStaticWeapons)
    .X = X
    .Y = Y '+ Static_Weapon_Y_Inc
    .bOnSurface = False
    .iWeapon = vWeapon
End With

NumStaticWeapons = NumStaticWeapons + 1

End Sub

Private Sub RemoveStaticWeapon(Index As Integer)

Dim i As Integer

If NumStaticWeapons = 1 Then
    Erase StaticWeapon
    NumStaticWeapons = 0
Else
    For i = Index To NumStaticWeapons - 2
        StaticWeapon(i) = StaticWeapon(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve StaticWeapon(NumStaticWeapons - 2)
    NumStaticWeapons = NumStaticWeapons - 1
End If

End Sub

Private Sub RemoveCasing(Index As Integer)

Dim i As Integer

If NumCasings = 1 Then
    Erase Casing
    NumCasings = 0
Else
    'Remove the bullet
    For i = Index To NumCasings - 2
'        Casing(i).Decay = Casing(i + 1).Decay
'        Casing(i).Heading = Casing(i + 1).Heading
'        Casing(i).Speed = Casing(i + 1).Speed
'        Casing(i).X = Casing(i + 1).X
'        Casing(i).Y = Casing(i + 1).Y
        Casing(i) = Casing(i + 1)
    Next i
    
    'Resize the array
    'ReDim Preserve Casing(NumCasings - 2)
    NumCasings = NumCasings - 1
End If

End Sub

Private Sub DrawBullets()

Dim i As Integer
Dim pX As Single, pY As Single
Const Bullet_Len_F As Single = 4

''Remove any decayed bullets
'i = 0
'Do While i < NumBullets
'    'Is this one decayed?
'    If Bullet(i).Decay < GetTickCount() Then
'        'Kill it!
'        RemoveBullet i, False
'        'Decrement the counter
'        i = i - 1
'    End If
'    'Increment the counter
'    i = i + 1
'Loop

'Step through each bullet and draw it
picMain.DrawWidth = 2
picMain.FillStyle = vbFSTransparent

For i = 0 To NumBullets - 1
    'Draw the bullet
    'modstickgame.sCircle  Bullet(i).x - (BULLET_RADIUS + 0.5), Bullet(i).y - (BULLET_RADIUS + 0.5), _
        Bullet(i).x + BULLET_RADIUS + 0.5, Bullet(i).y + BULLET_RADIUS + 0.5, Me.hdc
    
    If Bullet(i).bSilenced = False Then
        'TO SHOW SILENCED TRAILS, UNCOMMENT ABOVE, AND UNCOMMENT THE SAME IF STATEMENT BELOW
        
        On Error GoTo EH
        pX = CLng(Bullet(i).X + Sine(Bullet(i).Heading) * Bullet(i).Speed / Bullet_Len_F)
        pY = CLng(Bullet(i).Y - CoSine(Bullet(i).Heading) * Bullet(i).Speed / Bullet_Len_F)
        
        'If Bullet(i).bSilenced = False Then
            picMain.ForeColor = vbYellow
            modStickGame.sLine Bullet(i).X, Bullet(i).Y, pX, pY
            
            'modStickGame.sCircle pX, pY, Bullet_Radius, vbBlack
        'End If
        
        If Bullet(i).bSniperBullet Then
            If Bullet(i).LastSmoke + Sniper_Smoke_Delay / GetTimeZoneAdjust(Bullet(i).X, Bullet(i).Y) < GetTickCount() Then
                AddSmoke Bullet(i).X, Bullet(i).Y, 10, Bullet(i).Heading, False, True
                Bullet(i).LastSmoke = GetTickCount()
            End If
        ElseIf modStickGame.cg_ShowBulletTrails Then
            If Bullet(i).LastSmoke + BulletTrail_Smoke_Delay / GetTimeZoneAdjust(Bullet(i).X, Bullet(i).Y) < GetTickCount() Then
                'AddBulletTrail Bullet(i).X, Bullet(i).Y, Bullet(i).Heading + Sgn(Bullet(i).LastTrailDir) * piD2
                AddBulletTrail Bullet(i).X, Bullet(i).Y, Bullet(i).Heading + piD2, Bullet(i).Speed / BULLET_SPEED, Bullet(i).bTracer
                
                Bullet(i).LastSmoke = GetTickCount()
            End If
        End If
        
        
        'PrintStickText "Speed: " & Bullet(i).Speed, Bullet(i).X, Bullet(i).Y, 0
        
        'picMain.fillstyle = vbFSSolid
        'picMain.FillColor = Bullet(i).Colour
        'modstickgame.sCircle  (Bullet(i).X, Bullet(i).Y), Bullet_Radius, Bullet(i).Colour
    End If
    
    
    'If Stick(0).Perk = pBombSquad Then
        'modStickGame.sCircle Bullet(i).X, Bullet(i).Y, 1000, vbBlack
    'End If
    
Next i

EH:
End Sub

Private Sub DrawSmokeBlasts()
Dim i As Integer, j As Integer
Const BlastCol = BoxCol 'SmokeFill
'Const MaxWidth = 20, MaxLen = 300

'old
Const sgMaxSize = 30
Const LineLen = 10
Dim PM_Offset As Single: PM_Offset = PM_Rnd() * Pi / 2

'picMain.DrawWidth = 2
picMain.ForeColor = BlastCol

Do While i < NumSmokeBlasts
    
'    With SmokeBlast(i)
'        If Int(.sWidth) Then
'            picMain.DrawWidth = Int(.sWidth)
'        End If
'
'
'        modStickGame.sLine .X, .Y, .X + .sLength * sine(.Heading), .Y - .sLength * cosine(.Heading), BlastCol
'
'
'
'        If .iDir = 1 Then
'            If .sWidth < MaxWidth Then
'                .sWidth = .sWidth + 0.25 * modStickGame.StickTimeFactor
'            End If
'
'
'            If .sLength < MaxLen Then
'                .sLength = .sLength + 10 * modStickGame.StickTimeFactor
'            Else
'                .iDir = -1
'            End If
'
'        Else
'            '.sLength = .sLength - 30 * modStickGame.StickTimeFactor
'            .sWidth = Round(.sWidth * modStickGame.StickTimeFactor / 2 - 0.01, 2)
'
'        End If
'
'    End With
'
'
'    If SmokeBlast(i).sWidth < 0 Then
'        RemoveSmokeBlast i
'        i = i - 1
'    End If
    
    
    
'    'draw
    '3 lines from .x, .y to .x+ksine(Heading+-a),.y+kcosine(Heading+-a)
    '##################
    With SmokeBlast(i)
        modStickGame.sLine .X + .sLength * Sine(.Heading), .Y - .sLength * CoSine(.Heading), _
                           .X + (.sLength * LineLen) * Sine(.Heading), .Y - (.sLength * LineLen) * CoSine(.Heading)
        
        For j = 0 To 4
            
            modStickGame.sLine .X + .sLength * Sine(.Heading), .Y - .sLength * CoSine(.Heading), _
                           .X + (.sLength * LineLen) * Sine(.Heading + .sOffset + PM_Rnd()), .Y - (.sLength * LineLen) * CoSine(.Heading + .sOffset + PM_Rnd())
            
            modStickGame.sLine .X + .sLength * Sine(.Heading), .Y - .sLength * CoSine(.Heading), _
                           .X + (.sLength * LineLen) * Sine(.Heading - .sOffset + PM_Rnd()), .Y - (.sLength * LineLen) * CoSine(.Heading - .sOffset + PM_Rnd())
            
        Next j
        .sLength = .sLength + 4 * modStickGame.StickTimeFactor * GetTimeZoneAdjust(.X, .Y)
    End With
    '##################
    
    
    If SmokeBlast(i).sLength > sgMaxSize Then
        RemoveSmokeBlast i
        i = i - 1
    End If
    
    
    i = i + 1
Loop

End Sub

Private Sub AddSmokeBlast(X As Single, Y As Single, Heading As Single)

If modStickGame.cg_Smoke Then
    ReDim Preserve SmokeBlast(NumSmokeBlasts)
    
    With SmokeBlast(NumSmokeBlasts)
        .Heading = Heading + PM_Rnd() * piD4
        
        .X = X
        .Y = Y
        
        '.iDir = 1
        
        .sOffset = PM_Rnd() * Pi / 4
    End With
    
    NumSmokeBlasts = NumSmokeBlasts + 1
End If

End Sub

Private Sub RemoveSmokeBlast(Index As Integer)

Dim i As Integer

If NumSmokeBlasts = 1 Then
    Erase SmokeBlast
    NumSmokeBlasts = 0
Else
    'Remove the bullet
    For i = Index To NumSmokeBlasts - 2
'        SmokeBlast(i).Decay = SmokeBlast(i + 1).Decay
'        SmokeBlast(i).Heading = SmokeBlast(i + 1).Heading
'        SmokeBlast(i).Speed = SmokeBlast(i + 1).Speed
'        SmokeBlast(i).X = SmokeBlast(i + 1).X
'        SmokeBlast(i).Y = SmokeBlast(i + 1).Y
        SmokeBlast(i) = SmokeBlast(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve SmokeBlast(NumSmokeBlasts - 2)
    NumSmokeBlasts = NumSmokeBlasts - 1
End If

End Sub

Private Sub DrawCasings()
Dim i As Integer
Const Casing_LenD2 As Single = Casing_Len / 2

i = 0
Do While i < NumCasings
    
    If Casing(i).Decay < GetTickCount() Then
        RemoveCasing i
        i = i - 1
    End If
    
    i = i + 1
Loop


picMain.DrawWidth = 2
For i = 0 To NumCasings - 1
    
    'If Casing(i).bSniperCasing Then
        'picMain.DrawWidth = 2
        
        'modStickGame.sLine Casing(i).X, Casing(i).Y, _
          Casing(i).X + 2 * Casing_Len * Sine(Casing(i).Facing) _
        , Casing(i).Y - 2 * Casing_Len * CoSine(Casing(i).Facing), vbYellow
        
        'picMain.DrawWidth = 1
        
    'Else
    picMain.ForeColor = vbYellow
    modStickGame.sLine Casing(i).X, Casing(i).Y, _
          Casing(i).X + Casing_LenD2 * Sine(Casing(i).Facing) _
        , Casing(i).Y - Casing_LenD2 * CoSine(Casing(i).Facing)
    
    'picMain.ForeColor = vbBlack
    'modStickGame.sLine_FromLast Casing(i).X + Casing_Len * Sine(Casing(i).Facing), _
                                Casing(i).Y - Casing_Len * CoSine(Casing(i).Facing)
    
    'End If
    
    
Next i

'picMain.DrawWidth = 1

End Sub

Private Sub DrawBlood()
Dim i As Integer
Const Blood_Radius = 10

picMain.DrawWidth = 3
picMain.FillStyle = vbFSSolid

For i = 0 To NumBlood - 1
    modStickGame.sCircle Blood(i).X, Blood(i).Y, Blood_Radius, vbRed
    'modstickgame.sCircle  (Blood(i).X + Rnd() * 30, Blood(i).Y + Rnd() * 30), Bullet_Radius / 2, vbRed
Next i

picMain.DrawWidth = 1

End Sub

Private Sub AddBulletExplosion(X As Single, Y As Single)
AddCirc X, Y, 75, 1, vbYellow, 12, True
End Sub

Private Sub RemoveBullet(Index As Integer, bWall As Boolean, Optional bFancy As Boolean = True) ', Optional sgWallSurface As Single) ', Optional ByVal WithSmoke As Boolean = True)

Dim i As Integer
Dim headingMpi As Single

If bFancy Then
    If Bullet(Index).Speed > 50 Then
        If bWall Then
            
            'Bullet(Index).Heading = FixAngle(Bullet(Index).Heading)
            headingMpi = FixAngle(Bullet(Index).Heading - Pi)
            
            AddWallMark Bullet(Index).X, Bullet(Index).Y, WallMark_Bullet_Radius
            AddSmokeBlast Bullet(Index).X, Bullet(Index).Y, headingMpi ', sgWallSurface
            
            If Rnd() > 0.8 Then
                AddNadeTrail Bullet(Index).X, Bullet(Index).Y, headingMpi 'IIf(Bullet(i).Heading > Pi, piD4 + Rnd() * piD2, pi7D4 - Rnd() * piD2)
            End If
            
            If modStickGame.cg_Smoke Then
                If Rnd() > 0.8 Or Bullet(Index).bSniperBullet Then
                    AddSmokeNadeTrail Bullet(Index).X, Bullet(Index).Y
                End If
            End If
            
            If Rnd() > 0.2 Then
                'AddExplosion Bullet(Index).X, Bullet(Index).Y, 150, 0.25, 0, 0
                AddBulletExplosion Bullet(Index).X, Bullet(Index).Y
            End If
            If Rnd() > 0.2 Then
                AddSparks Bullet(Index).X, Bullet(Index).Y, headingMpi
            End If
        ElseIf modStickGame.cg_Smoke Then
            If Rnd() > 0.8 Or Bullet(Index).bSniperBullet Then
                AddSmokeGroup Bullet(Index).X, Bullet(Index).Y, 4, 0, 0, Rnd() > 0.3 '10, Bullet(Index).Heading
            End If
        End If
    End If
End If

'If there's only one bullet left, just erase the array
If NumBullets = 1 Then
    Erase Bullet
    NumBullets = 0
Else
    'Remove the bullet
    For i = Index To NumBullets - 2
        Bullet(i) = Bullet(i + 1)
    Next i
    
    
    'Resize the array
    'ReDim Preserve Bullet(NumBullets - 2)
    NumBullets = NumBullets - 1
End If

End Sub

Private Sub RemoveBlood(Index As Integer)

Dim i As Integer

If NumBlood = 1 Then
    Erase Blood
    NumBlood = 0
Else
    'Remove the bullet
    For i = Index To NumBlood - 2
'        Blood(i).Decay = Blood(i + 1).Decay
'        Blood(i).Heading = Blood(i + 1).Heading
'        Blood(i).Speed = Blood(i + 1).Speed
'        Blood(i).X = Blood(i + 1).X
'        Blood(i).Y = Blood(i + 1).Y
        Blood(i) = Blood(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Blood(NumBlood - 2)
    NumBlood = NumBlood - 1
End If

End Sub

Private Sub RemoveNade(Index As Integer, bWall As Boolean)

Dim i As Integer


If bWall Then
    If Nade(Index).iType = nFrag Then
        AddWallMark Nade(Index).X, Nade(Index).Y, WallMark_Explosion_Radius
    End If
End If


If NumNades = 1 Then
    Erase Nade
    NumNades = 0
Else
    For i = Index To NumNades - 2
'        Nade(i).Decay = Nade(i + 1).Decay
'        Nade(i).Heading = Nade(i + 1).Heading
'        Nade(i).Speed = Nade(i + 1).Speed
'        Nade(i).X = Nade(i + 1).X
'        Nade(i).Y = Nade(i + 1).Y
'        Nade(i).OwnerID = Nade(i + 1).OwnerID
'        Nade(i).IsRPG = Nade(i + 1).IsRPG
'        Nade(i).LastSmoke = Nade(i + 1).LastSmoke
        Nade(i) = Nade(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Nade(NumNades - 2)
    NumNades = NumNades - 1
End If

End Sub

Private Sub RemoveMine(Index As Integer)

Dim i As Integer

If NumMines = 1 Then
    Erase Mine
    NumMines = 0
Else
    For i = Index To NumMines - 2
        Mine(i) = Mine(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Mine(NumMines - 2)
    NumMines = NumMines - 1
End If

End Sub

Private Sub RemoveMag(Index As Integer)

Dim i As Integer

If NumMags = 1 Then
    Erase Mag
    NumMags = 0
Else
    For i = Index To NumMags - 2
        Mag(i) = Mag(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Mag(NumMags - 2)
    NumMags = NumMags - 1
End If

End Sub

Private Sub RemoveSpark(Index As Integer)

Dim i As Integer

If NumSparks = 1 Then
    Erase Spark
    NumSparks = 0
Else
    For i = Index To NumSparks - 2
        Spark(i) = Spark(i + 1)
    Next i
    
    'Resize the array
    'ReDim Preserve Spark(NumSparks - 2)
    NumSparks = NumSparks - 1
End If

End Sub

Private Sub DrawSparks()
Dim i As Integer
Dim pX As Single, pY As Single

picMain.ForeColor = vbYellow
picMain.DrawWidth = 1

For i = 0 To NumSparks - 1
    pX = CSng(Spark(i).X + Sine(Spark(i).Heading) * Spark(i).Speed)
    pY = CSng(Spark(i).Y - CoSine(Spark(i).Heading) * Spark(i).Speed)
    
    modStickGame.sLine Spark(i).X, Spark(i).Y, pX, pY
Next i

End Sub

Private Sub ProcessSparks()
Dim i As Integer

Do While i < NumSparks
    If Spark(i).Decay < GetTickCount() Then
        RemoveSpark i
        i = i - 1
    ElseIf Spark(i).Speed < Spark_Min_Speed Then
        RemoveSpark i
        i = i - 1
    End If
    i = i + 1
Loop

For i = 0 To NumSparks - 1
    If Spark(i).LastReduction + Spark_Speed_Reduction_Delay / GetTimeZoneAdjust(Spark(i).X, Spark(i).Y) < GetTickCount() Then
        Spark(i).Speed = Spark(i).Speed / Spark_Speed_Reduction
        Spark(i).LastReduction = GetTickCount()
    End If
    
    MotionStickObject Spark(i).X, Spark(i).Y, Spark(i).Speed, Spark(i).Heading
Next i

End Sub

Private Sub AddSparks(X As Single, Y As Single, GeneralHeading As Single)
Dim i As Integer

For i = 0 To 4
    AddSpark X, Y, GeneralHeading + PM_Rnd() * Spark_Diffraction, Spark_Speed + Rnd() * 20
Next i


End Sub
Private Sub AddMoreSparks(X As Single, Y As Single, n As Integer)
Dim i As Integer

For i = 0 To n
    AddSpark X, Y, PM_Rnd() * Pi2 + PM_Rnd() * Spark_Diffraction, Spark_Speed + Rnd() * 20
Next i


End Sub

''########################################################################
'shieldwave stuff
Private Sub AddShieldWave(X As Single, Y As Single, Facing As Single)
Const Max_Shield_Size As Single = piD2

ReDim Preserve ShieldWave(NumShieldWaves)

With ShieldWave(NumShieldWaves)
    .Facing = FixAngle(piD4 - Facing)
    .X = X
    .Y = Y
    .Size = Max_Shield_Size
    .colour = RandomRGBColour()
End With

NumShieldWaves = NumShieldWaves + 1


End Sub

Private Sub RemoveShieldWave(Index As Integer)

Dim i As Integer

If NumShieldWaves = 1 Then
    Erase ShieldWave
    NumShieldWaves = 0
Else
    For i = Index To NumShieldWaves - 2
        ShieldWave(i) = ShieldWave(i + 1)
    Next i

    'Resize the array
    ReDim Preserve ShieldWave(NumShieldWaves - 2)
    NumShieldWaves = NumShieldWaves - 1
End If

End Sub

Private Sub DrawShieldWaves()
Dim i As Integer
Const Radius As Single = 500, _
    shield_col As Long = &HFF00FF 'purple

For i = 0 To NumShieldWaves - 1
    With ShieldWave(i)
        modStickGame.sCircleSE .X, .Y, Radius, .colour, .Facing, FixAngle(.Facing + .Size)
    End With
Next i

End Sub

Private Sub ProcessShieldWaves()
Const size_dec As Single = 0.01, size_decX2 = size_dec * 2
Dim i As Integer
Dim amount As Single

While i < NumShieldWaves
    With ShieldWave(i)
        amount = size_dec * GetTimeZoneAdjust(.X, .Y)
        .Size = .Size - amount
        .Facing = FixAngle(.Facing + amount / 2)
    End With
    If ShieldWave(i).Size <= size_decX2 Then
        RemoveShieldWave i
        i = i - 1
    End If
    i = i + 1
Wend

End Sub

Private Sub ProcessShields()
Dim i As Integer
Const shield_inc As Single = 3

#If Not Hack_Shield Then
If Stick(0).Shield > 0 Then
#End If
    If Stick(0).Shield < Max_Shield Then
#If Not Hack_Shield Then
        If Stick(0).LastShieldHitTime + Shield_Recharge_Delay / GetMyTimeZone() < GetTickCount() Then
#End If
            Stick(0).Shield = Round(Stick(0).Shield + shield_inc * GetMyTimeZone(), 1)
            
            If Not Stick(0).ShieldCharging Then
                addShieldExhaustWave 0, 15, vbYellow
                Stick(0).ShieldCharging = True
            End If
            
            If Stick(0).Shield > Max_Shield Then
                Stick(0).Shield = Max_Shield
                Stick(0).ShieldCharging = False
            End If
#If Not Hack_Shield Then
        End If
#End If
    End If
#If Not Hack_Shield Then
End If
#End If


For i = 1 To NumSticksM1
    If Stick(i).IsBot Then
#If Not Hack_AIShield Then
        If Stick(i).Shield > 0 Then
#End If
            If Stick(i).Shield < Max_Shield Then
#If Not Hack_AIShield Then
                If Stick(i).LastShieldHitTime + Shield_Recharge_Delay / GetSticksTimeZone(i) < GetTickCount() Then
#End If
                    Stick(i).Shield = Round(Stick(i).Shield + shield_inc * GetSticksTimeZone(i), 1)
                    If Not Stick(i).ShieldCharging Then
                        addShieldExhaustWave i, 15, vbYellow
                        Stick(i).ShieldCharging = True
                    End If
                    
                    If Stick(i).Shield > Max_Shield Then
                        Stick(i).Shield = Max_Shield
                    End If
#If Not Hack_AIShield Then
                End If
#End If
            End If
#If Not Hack_AIShield Then
        End If
#End If
    End If
Next i

ProcessShieldWaves

End Sub

Private Sub addShieldExhaustWave(iStick As Integer, Size As Single, colour As Long)
Const n As Integer = 5, speedFactor As Single = 200, separationDist = 120
Dim i As Integer
Dim sizeXn As Single, sizeDx As Single: sizeXn = Size * n + 1: sizeDx = Size / 3

For i = 1 To n
    AddCirc Stick(iStick).X + (n - i) * shieldChargeDist * sIn(Stick(iStick).Heading) * Stick(iStick).Speed / speedFactor, _
            Stick(iStick).Y - (n - i) * shieldChargeDist * Cos(Stick(iStick).Heading) * Stick(iStick).Speed / speedFactor + i * separationDist, _
            sizeXn, i * Size, colour, sizeDx, True
    
Next i

End Sub

''########################################################################
'flame stuff
Private Sub AddFlame(X As Single, Y As Single, Heading As Single, Speed As Single, OwnerID As Integer, _
    iStick As Integer) ', bIncBulletsFired As Boolean)

ReDim Preserve Flame(NumFlames)

With Flame(NumFlames)
    .Decay = GetTickCount() + Flame_Time / GetSticksTimeZone(iStick)
    .Heading = Heading
    .Speed = Speed
    .X = X
    .Y = Y
    .OwnerID = OwnerID
End With

NumFlames = NumFlames + 1

If iStick = 0 Or Stick(iStick).IsBot Then
    'If bIncBulletsFired Then
        Stick(iStick).BulletsFired = Stick(iStick).BulletsFired + 1
    'End If
End If

If modStickGame.cg_Smoke Then
    If Rnd() > 0.5 Then
        AddSmokeGroup Stick(iStick).GunPoint.X, Stick(iStick).GunPoint.Y, 4, Rnd() * Speed, Heading
    End If
End If

End Sub

Private Sub RemoveFlame(Index As Integer)

Dim i As Integer

If modStickGame.cg_Smoke Then
    If Rnd() > 0.7 Then
        AddSmokeGroup Flame(Index).X, Flame(Index).Y, 4, Flame(Index).Speed / 3, PM_Rnd(), True
        '                                                                         ^Flame(Index).Heading
    End If
End If

If NumFlames = 1 Then
    Erase Flame
    NumFlames = 0
Else
    For i = Index To NumFlames - 2
        Flame(i) = Flame(i + 1)
    Next i

    'Resize the array
    ReDim Preserve Flame(NumFlames - 2)
    NumFlames = NumFlames - 1
End If

End Sub

Private Sub DrawFlames()
Dim i As Integer

'picMain.FillStyle = vbFSSolid

For i = 0 To NumFlames - 1
    With Flame(i)
        DrawFlame .X, .Y, .Size
        If modStickGame.cg_Smoke Then
            If Rnd() > 0.99 / modStickGame.sv_StickGameSpeed Then
                AddSmokeGroup .X, .Y, 4, .Speed / 3, PM_Rnd(), True
            End If
        End If
    End With
Next i

For i = 0 To NumSticks - 1
    If Stick(i).bOnFire Then
        
        DrawFlame Stick(i).X + ArmLen / 2 * Rnd(), _
            GetStickY(i) + IIf(StickiHasState(i, STICK_PRONE), HeadRadius, BodyLen) * Rnd(), _
            Flame_Burn_Radius
        
    End If
Next i


For i = 0 To NumFires - 1
    With Fire(i)
        DrawFlame .X, .Y, .Size
        If .LastGravity + Fire_Smoke_Delay < GetTickCount() Then
            AddSmokeGroup .X, .Y, 3, 25 * Rnd(), PM_Rnd() * piD8, True, True
            .LastGravity = GetTickCount()
        End If
    End With
Next i


'picMain.FillStyle = vbFSTransparent

End Sub

Private Sub ProcessFlames()
Dim i As Integer, j As Integer
Dim iFlameOwner As Integer ', MinDist As Single

Do While i < NumFlames
    If Flame(i).Decay < GetTickCount() Then
        RemoveFlame i
        i = i - 1
'    ElseIf Flame(i).Speed < 5 Then
'        RemoveFlame i
'        i = i - 1
    End If
    i = i + 1
Loop

i = 0
Do While i < NumFlames
    
    With Flame(i)
        If .Size < Flame_Max_Radius Then
            .Size = .Size + 6 * GetTimeZoneAdjust(.X, .Y)
        End If
        
        MotionStickObject .X, .Y, .Speed, .Heading
        
        ApplyGravityVector .LastGravity, GetTimeZoneAdjust(.X, .Y), .Speed, .Heading, .X, .Y, Gravity_Strength / 14
    End With
    
    
    If ClipFlame(i) = False Then
        iFlameOwner = FindStick(Flame(i).OwnerID)
        If iFlameOwner > -1 Then
            For j = 0 To NumSticksM1
                If j <> iFlameOwner Then
                    If Stick(j).LastFlameTouch + Flame_Impact_Delay / GetSticksTimeZone(j) < GetTickCount() Then
                        If IsAlly(Stick(j).Team, Stick(iFlameOwner).Team) = False Then
                            If StickInGame(j) Then
                                If StickInvul(j) = False Then
                                    If Stick(j).WeaponType <> Chopper Then
                                        If GetDist(Stick(j).X, Stick(j).Y, Flame(i).X, Flame(i).Y) < 500 Then
                                            Stick(j).LastFlameTouch = GetTickCount()
                                            Stick(j).LastFlameTouchOwnerID = Flame(i).OwnerID
                                            Stick(j).bFlameIsFromTag = False
                                            Stick(j).bOnFire = True
                                            
                                            If j = 0 Or Stick(j).IsBot Then
                                                DamageStick Flame_Damage, j, iFlameOwner
                                                
                                                If Stick(j).Health < 1 Then
                                                    Killed j, iFlameOwner, kFlame
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next j
        Else
            RemoveFlame i
            i = i - 1
        End If
    Else
        'flame removed by clipflame()
        i = i - 1
    End If
    
    
    i = i + 1
Loop



'damage sticks who came into contact
For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        'If Stick(i).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
        If StickInvul(i) = False Then
            
            If Stick(i).bOnFire Then
                'damage
                If i = 0 Or Stick(i).IsBot Then
                    If Stick(i).LastFlameDamage + Flame_Burn_Damage_Time / GetSticksTimeZone(i) < GetTickCount() Then
                        
                        iFlameOwner = FindStick(Stick(i).LastFlameTouchOwnerID)
                        DamageStick Flame_Burn_Damage, i, iFlameOwner, False
                        
                        
                        If Stick(i).Health < 1 Then
                            Killed i, iFlameOwner, _
                                IIf(Stick(i).bFlameIsFromTag, eKillTypes.kFlameTag, eKillTypes.kBurn)
                            
                        End If
                        
                        
                        Stick(i).LastFlameDamage = GetTickCount()
                    End If
                End If
                
                
            End If
        End If
    End If
Next i

ProcessFlameTagging
ProcessFires

End Sub

Private Sub ProcessFlameTagging()
Dim i As Integer, j As Integer

For i = 0 To NumSticksM1
    If StickInGame(i) Then
        If Stick(i).bOnFire Then
        
            For j = 0 To NumSticksM1
                If j <> i Then
                    If StickInGame(j) Then
                        
                        If Not IsAlly(Stick(j).Team, Stick(i).Team) Then
                            
                            If Stick(j).bOnFire = False And StickInvul(j) = False Then
                                
                                If CoOrdNearStick(Stick(i).X, Stick(i).Y, j) Then
                                    
                                    Stick(j).bOnFire = True
                                    Stick(j).LastFlameTouch = GetTickCount() 'for me/bots only
                                    Stick(j).bFlameIsFromTag = True
                                    Stick(j).LastFlameTouchOwnerID = Stick(i).ID
                                    
                                    
                                    Exit For
                                    
                                End If 'co-ord endif
                            End If
                            
                        End If 'team endif
                    End If 'stickingame j endif
                End If 'j<>i endif
            Next j
        
        
        End If 'onfire endif
    End If 'stickingame i endif
Next i

End Sub

'###########################
'Fire stuff
Private Sub AddFire(X As Single, Y As Single, iStick As Integer)

ReDim Preserve Fire(NumFires)

With Fire(NumFires)
    .Decay = GetTickCount() + Fire_Time / GetTimeZoneAdjust(X, Y)
    .X = X
    .Y = Y
    .OwnerID = Stick(iStick).ID
    .LastGravity = GetTickCount() 'use this as .lastsmoke
End With

NumFires = NumFires + 1

End Sub

Private Sub RemoveFire(Index As Integer)

Dim i As Integer

If modStickGame.cg_Smoke Then
    If Rnd() > 0.7 Then
        AddSmokeGroup Fire(Index).X, Fire(Index).Y, 4, Fire(Index).Speed / 3, PM_Rnd(), True
        '                                                                         ^Fire(Index).Heading
    End If
End If

If NumFires = 1 Then
    Erase Fire
    NumFires = 0
Else
    For i = Index To NumFires - 2
        Fire(i) = Fire(i + 1)
    Next i

    'Resize the array
    ReDim Preserve Fire(NumFires - 2)
    NumFires = NumFires - 1
End If

End Sub

Private Sub ProcessFires()
Dim i As Integer, j As Integer, iOwner As Integer
Dim iFireOwner As Integer

Do While i < NumFires
    If Fire(i).Decay < GetTickCount() Then
        RemoveFire i
        i = i - 1
    End If
    i = i + 1
Loop


i = 0
Do While i < NumFires
    
    With Fire(i)
        If .Size < Flame_Max_Radius Then
            .Size = .Size + 6 * GetTimeZoneAdjust(.X, .Y)
        End If
    End With
    
    
    iFireOwner = FindStick(Fire(i).OwnerID)
    
    If iFireOwner > -1 Then
        For j = 0 To NumSticksM1
            If j <> iFireOwner Then
                If Stick(j).LastFlameTouch + Flame_Impact_Delay / GetSticksTimeZone(j) < GetTickCount() Then
                    If IsAlly(Stick(j).Team, Stick(iFireOwner).Team) = False Then
                        'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
                        If StickInGame(j) Then
                            If StickInvul(j) = False Then
                                If Stick(j).WeaponType <> Chopper Then
                                    If GetDist(Stick(j).X, Stick(j).Y, Fire(i).X, Fire(i).Y) < 500 Then
                                        Stick(j).LastFlameTouch = GetTickCount()
                                        Stick(j).LastFlameTouchOwnerID = Fire(i).OwnerID
                                        Stick(j).bFlameIsFromTag = False
                                        Stick(j).bOnFire = True
                                        
                                        If j = 0 Or Stick(j).IsBot Then
                                            DamageStick Flame_Damage, j, iFireOwner
                                            
                                            If Stick(j).Health < 1 Then
                                                Killed j, iOwner, kFlame
                                            End If
                                        End If
                                        
                                        If Stick(j).WeaponType = Chopper Then
                                            RemoveFire i
                                            i = i - 1
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    Else
        RemoveFire i
        i = i - 1
    End If
    
    
    i = i + 1
Loop

End Sub


Private Function CoOrdNearStick(X As Single, Y As Single, Sticki As Integer) As Boolean

Const XLimit = ArmLen * 3, YLimit = BodyLen * 2
Dim sY As Single
Dim rcStick As RECT

If Stick(Sticki).WeaponType = Chopper Then
    CoOrdNearStick = False
    
Else
    'If X < (Stick(Sticki).X + XLimit) Then
        'If X > (Stick(Sticki).X - XLimit) Then
    
'    If Abs(X - Stick(Sticki).X) < XLimit Then
'        sY = GetStickY(Sticki)
'        If Y > (sY - HeadRadius) Then
'            CoOrdNearStick = (Y < (sY + YLimit))
'        End If
'    End If
    
    
    With rcStick
        .Left = Stick(Sticki).X - XLimit
        .Right = Stick(Sticki).X + XLimit
        sY = GetStickY(Sticki)
        .Top = sY - HeadRadius
        .Bottom = sY + YLimit
    End With
    
End If

CoOrdNearStick = RectCollision(rcStick, PointToRect(X, Y))


End Function

Private Sub DrawFlame(X As Single, Y As Single, Size As Single)
Const Factor As Single = 2

Dim i As Integer
Dim SizeDF As Single

SizeDF = Size / Factor

For i = 0 To 4
    DrawSingleFlame X + PM_Rnd * SizeDF, Y + PM_Rnd * SizeDF, SizeDF
Next i

End Sub

Private Sub DrawSingleFlame(X As Single, Y As Single, Size As Single)

picMain.FillColor = vbRed
modStickGame.sCircle X, Y, Size, vbRed

picMain.FillColor = vbYellow
modStickGame.sCircle X, Y, Size / 1.4, vbYellow

picMain.FillColor = MOrange
modStickGame.sCircle X, Y, Size / 2.4, MOrange

End Sub

Private Function ClipFlame(i As Integer) As Boolean

Const Lim As Integer = 50, Max_Flame_Size = 500
Dim ClippedX As Boolean, ClippedY As Boolean
Dim XComp As Single, YComp As Single

ClippedX = (Flame(i).X < Lim) Or (Flame(i).X > StickGameWidth - Lim)
ClippedY = (Flame(i).Y < Lim) Or (Flame(i).Y > StickGameHeight - Lim)

If ClippedX Or ClippedY Then
    
    ClipFlame = True
    RemoveFlame i
    
ElseIf FlameInPlatform(i) Then
    
    ClipFlame = True
    RemoveFlame i
    
ElseIf FlameInTBox(i) Then
    
    ClipFlame = True
    RemoveFlame i
    
ElseIf FlameInBox(i) Then
    
    ClipFlame = True
    RemoveFlame i
    
ElseIf Flame(i).Size > Max_Flame_Size Then
    
    ClipFlame = True
    RemoveFlame i
    
End If

End Function

Private Function FlameInBox(i As Integer) As Boolean
Dim j As Integer

For j = 0 To ubdBoxes
    If Box(j).bInUse Then
        If FlameCollision(i, Box(j).Left, Box(j).Top, Box(j).width, Box(j).height) Then
            FlameInBox = True
            Exit For
        End If
    End If
Next j

End Function

Private Function FlameInTBox(i As Integer) As Boolean
Dim j As Integer

For j = 0 To ubdtBoxes
    If FlameCollision(i, tBox(j).Left, tBox(j).Top, tBox(j).width, tBox(j).height) Then
        FlameInTBox = True
        Exit For
    End If
Next j

End Function

Private Function FlameInPlatform(i As Integer) As Boolean
Dim j As Integer

For j = 0 To ubdPlatforms
    If FlameCollision(i, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        FlameInPlatform = True
        Exit For
    End If
Next j

End Function

Private Function FlameCollision(ByVal i As Integer, _
    oLeft As Single, oTop As Single, oWidth As Single, oHeight As Single) As Boolean

Dim SizeD2 As Single

SizeD2 = Flame(i).Size / 2

If Flame(i).X + SizeD2 >= oLeft Then
    If (Flame(i).X + SizeD2 <= (oLeft + oWidth)) Then
        If Flame(i).Y + SizeD2 >= oTop Then
            If Flame(i).Y + SizeD2 <= (oTop + oHeight) Then
                FlameCollision = True
            End If
        End If
    End If
End If

End Function

''########################################################################

'Private Sub RemoveMFlash(Index As Integer)
'
'Dim i as integer
'
'If NumMFlashes = 1 Then
'    Erase MFlash
'    NumMFlashes = 0
'Else
'    For i = Index To NumMFlashes - 2
'        MFlash(i) = MFlash(i + 1)
'    Next i
'
'    'Resize the array
'    ReDim Preserve MFlash(NumMFlashes - 2)
'    NumMFlashes = NumMFlashes - 1
'End If
'
'End Sub

Private Sub RemoveDeadStick(Index As Integer)

Dim i As Integer

If NumDeadSticks = 1 Then
    EraseDeadSticks
Else
    For i = Index To NumDeadSticks - 2
        DeadStick(i) = DeadStick(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve DeadStick(NumDeadSticks - 2)
    NumDeadSticks = NumDeadSticks - 1
End If

End Sub

Private Sub RemoveDeadChopper(Index As Integer)

Dim i As Integer

If NumDeadChoppers = 1 Then
    Erase DeadChopper
    NumDeadChoppers = 0
Else
    For i = Index To NumDeadChoppers - 2
        DeadChopper(i) = DeadChopper(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve DeadChopper(NumDeadChoppers - 2)
    NumDeadChoppers = NumDeadChoppers - 1
End If

End Sub

Private Function StartWinsock() As Boolean

''Init winsock
'If modWinsock.InitWinsock() = WINSOCK_ERROR Then
'    'Handle error..
'    GoTo EH
'End If

'Make a lSocket
lSocket = modWinsock.CreateSocket()
If lSocket = WINSOCK_ERROR Then
    'Handle error
    'modWinsock.TermWinsock
    GoTo EH
End If

'If we're the StickServer, bind to the StickServer port
If StickServer Then
    If modWinsock.BindSocket(lSocket, modPorts.StickPort) = WINSOCK_ERROR Then
        'Handle error
        GoTo EH
    End If
End If

StartWinsock = True

Exit Function
EH:
''AddText  "Error Starting Winsock", TxtError, True
Call EndWinsock
Unload Me
End Function

Private Function RequestMap() As Boolean

Dim JoinTimer As Long
Dim sPacket As String
Dim TempSockAddr As ptSockAddr
Dim CurrentRetry As Integer
Dim Txt As String, sMapName As String, sMapPath As String

Dim LastLine As Long
Const LineDelay = 20


'Make the server's ptsockaddr
If MakeSockAddr(ServerSockAddr, modPorts.StickPort, modStickGame.StickServerIP) = WINSOCK_ERROR Then
    'Handle error
    AddText "Error - IP isn't valid", TxtError, True 'Making Socket", TxtError, True
    Unload Me
    
Else
    CurrentRetry = 1
    
    Do
        If (JoinTimer + StickServer_RETRY_FREQ) < GetTickCount() Then
            'Reset the timer
            JoinTimer = GetTickCount()
            
            'Send the mPacket
            modWinsock.SendPacket lSocket, ServerSockAddr, sMapRequests
            
            If CurrentRetry < 6 Then
                Me.picMain.Cls
                
                Txt = "Requesting Map from Server '" & modStickGame.StickServerIP & ":" & CStr(modPorts.StickPort) & "'..."
                modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY - TextHeight(Txt), vbBlack
                
                Txt = "Waiting For Response... " & CStr(CurrentRetry)
                modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY + TextHeight(Txt), vbBlack
                
                Call BltToForm
                Me.Refresh
                
            End If
            
            CurrentRetry = CurrentRetry + 1
        End If
        
        
        DoEvents
        
        
        'Check for ACKs
        sPacket = modWinsock.ReceivePacket(lSocket, TempSockAddr)
        
        If LenB(sPacket) Then
            
            'Is this an ACK?
            If Left$(sPacket, 1) = sMapNames Then
                
                If IsValidVarPacket(sPacket) Then
                    sMapName = Mid$(sPacket, 2, InStrRev(sPacket, vbSpace) - 2)
                    
                    Txt = "Response Received"
                    modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY + 3 * TextHeight(Txt), vbBlack
                    
                    Txt = "Map Name is " & sMapName
                    modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY + 4 * TextHeight(Txt), vbBlack
                    
                    Call BltToForm
                    
                    'Start playing!
                    sMapPath = modStickGame.GetStickMapPath() & sMapName
                    
                    If FileExists(sMapPath) = False Then
                        AddText "Couldn't Connect to Server - Map '" & sMapName & "' is needed", TxtError, True
                        RequestMap = False
                    ElseIf LoadMapEx(sMapPath) Then
                        RequestMap = True
                        modStickGame.StickMapPath = sMapPath
                    Else
                        AddText "Error loading Map '" & sMapName & "'", TxtError, True
                        RequestMap = False
                    End If
                    Exit Function
                End If
            End If
            
        End If
        
        If LastLine + LineDelay < GetTickCount() Then
            If modVars.Closing Or WindowClosing Then Exit Function
            
            LastLine = GetTickCount()
            picMain.Line (ConnectingkX, ConnectingkY)-(ConnectingkX + LastLine - JoinTimer, ConnectingkY), vbRed
            
            Call BltToForm
            Me.Refresh
            
        End If
        
        
    Loop Until ((CurrentRetry - 1) > StickServer_NUM_RETRIES) Or modVars.Closing Or WindowClosing
    
    RequestMap = False
    
    If modVars.Closing Or WindowClosing Then Exit Function
    
    'We didn't receive an ACK before the timeout
    AddText "Unable to Request Map - No Packets Received", TxtError, True
    
End If

End Function

Private Function ConnectToServer() As Boolean

Dim JoinTimer As Long
'Dim TimeOutTimer As Long
Dim sPacket As String
Dim TempSockAddr As ptSockAddr
Dim CurrentRetry As Integer
Dim Txt As String


Dim LastLine As Long
Const LineDelay = 20



'Send "Join" packets to the server until we receive an "ACK" mPacket
CurrentRetry = 1

Do 'While TimeOutTimer + SERVER_CONNECT_DURATION > GetTickCount()
    
    'Is it time to send a "Join" mPacket?
    If (JoinTimer + StickServer_RETRY_FREQ) < GetTickCount() Then
        'Reset the timer
        JoinTimer = GetTickCount()
        
        'Send the mPacket
        modWinsock.SendPacket lSocket, ServerSockAddr, sJoins
        
        If CurrentRetry < 6 Then
            Me.picMain.Cls
            
            Txt = "Connecting to Server '" & modStickGame.StickServerIP & ":" & CStr(modPorts.StickPort) & "'..."
            modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY - TextHeight(Txt), vbBlack
            
            Txt = "Waiting For Response... " & CStr(CurrentRetry)
            modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY + TextHeight(Txt), vbBlack
            
            Call BltToForm
            Me.Refresh
            
        End If
        
        CurrentRetry = CurrentRetry + 1
        
    End If
    
    
    DoEvents
    
    
    'Check for ACKs
    sPacket = modWinsock.ReceivePacket(lSocket, TempSockAddr)
    
    If LenB(sPacket) Then
        
        'Is this an ACK?
        If Left$(sPacket, 1) = sAccepts Then
            
            'Set our ID
            Stick(0).ID = CInt(Mid$(sPacket, 2))
            AdjustIDArray
            'MyID = Stick(0).ID
            
            Txt = "Response Received"
            modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY + 3 * TextHeight(Txt), vbBlack
            
            Txt = "Setting Up Game..."
            modStickGame.PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY + 4 * TextHeight(Txt), vbBlack
            
            'Start playing!
            ConnectToServer = True
            Exit Function
            
        End If
        
    End If
    
    If LastLine + LineDelay < GetTickCount() Then
        If modVars.Closing Or WindowClosing Then Exit Function
        
        LastLine = GetTickCount()
        picMain.Line (ConnectingkX, ConnectingkY)-(ConnectingkX + LastLine - JoinTimer, ConnectingkY), vbRed
        
        Call BltToForm
        Me.Refresh
        
    End If
    
    
Loop Until ((CurrentRetry - 1) > StickServer_NUM_RETRIES) Or modVars.Closing Or WindowClosing

ConnectToServer = False

If modVars.Closing Or WindowClosing Then Exit Function

'We didn't receive an ACK before the timeout
AddText "Unable to Connect to Server - No Packet Flow", TxtError, True


End Function

Private Sub CheckForKills(sTxt As String)
Dim sName As String
Dim KillType As String
Dim i As Integer, j As Integer

'If InStr(1, sTxt, modMessaging.MsgNameSeparator) = 0 Then
On Error GoTo lEH
i = InStr(1, sTxt, "by", vbTextCompare) + 3
sName = Mid$(sTxt, i)

If LCase$(sName) = LCase$(Trim$(Stick(0).Name)) Then
    j = InStr(1, sTxt, "was", vbTextCompare) + 4
    KillType = Mid$(sTxt, j, i - j - 4)
    
    AddMainMessage "You " & KillType & vbSpace & Left$(sTxt, j - 6), False
    
    If Stick(0).WeaponType <> Chopper Then
        If modStickGame.sv_HPBonus Then
            'health bonus
            Stick(0).Health = 100
        End If
        
'        If Stick(0).Perk = pSpy Then
'            Stick(0).Perk = pNone
'
'            AddMainMessage
        
    End If
    
End If
'End If

lEH:
End Sub

Private Sub CheckForSticksChat(ChatText As String)
Dim i As Integer, j As Integer
Dim sName As String

j = InStr(1, ChatText, modMessaging.MsgNameSeparator)

sName = Left$(ChatText, j - 1)

For i = 0 To NumSticksM1
    If sName = Trim$(Stick(i).Name) Then
        
        Stick(i).LastChatMsg = GetTickCount()
        Stick(i).curChatMsg = Mid$(ChatText, j + 2)
        
        Exit For
    End If
Next i


End Sub

Private Function GetPacket() As Boolean

Dim sPacket As String
Dim TempSockAddr As ptSockAddr
Dim i As Integer, j As Integer

Dim Tmp As String, sTxt As String

'Loop until there were no packets
GetPacket = True

Do
    'Check for packets
    sPacket = modWinsock.ReceivePacket(lSocket, TempSockAddr)
    
'    i = 0
'    j = 0
'    Tmp = vbNullString
'    sTxt = vbNullString
    
    If LenB(sPacket) Then
        
        'Check what type of mPacket this is and take appropriate action
        Select Case Left$(sPacket, 1)
            Case sUpdates
                'A position update mPacket
                ProcessUpdatePacket Mid$(sPacket, 2)
                
            Case sDamageTicks
                
                On Error GoTo EH
                i = CInt(Mid$(sPacket, 2)) 'ID of stick to receive
                j = FindStick(i) 'index...
                
                'AddMainMessage "Recieved Tick - ID: " & i, False
                
                If j = 0 Then
                    ReceiveDamageTick
                    
                ElseIf modStickGame.StickServer Then
                    'tell the stick about his damage
                    
                    If j > -1 Then
                        modWinsock.SendPacket lSocket, Stick(j).SockAddr, sDamageTicks & CStr(i)
                    End If
                End If
                
                
            Case sBoxInfos
                
                If modStickGame.StickServer = False Then
                    On Error Resume Next
                    ReceiveBoxInfo Mid$(sPacket, 2)
                End If
                
            Case sSlowUpdates
                
                ProcessSlowPacket Mid$(sPacket, 2)
                
            Case sChats
                'A chat packet... if we're the server, broadcast
                
                On Error GoTo EH
                Tmp = Right$(sPacket, Len(sPacket) - InStrRev(sPacket, "#", , vbTextCompare))
                sTxt = Mid$(sPacket, 2, InStr(1, sPacket, "#", vbTextCompare) - 2)
                
                If modStickGame.StickServer Then
                    SendChatPacketBroadcast sTxt, CLng(Tmp)
                    '(auto-added to array)
                Else 'Otherwise, add it to the array
                    AddChatText sTxt, CLng(Tmp)
                End If
                
            Case sServerVarss
                
                If modStickGame.StickServer = False Then
                    ProcessServerVarPacket Mid$(sPacket, 2)
                End If
                
            Case sMineRefreshs
                
                If modStickGame.StickServer = False Then
                    ReceiveMineRefresh Mid$(sPacket, 2)
                End If
                
            Case sBarrelRefreshs
                
                If modStickGame.StickServer = False Then
                    ReceiveBarrelRefresh Mid$(sPacket, 2)
                End If
                
            Case sTimeZoneRefreshs
                
                If modStickGame.StickServer = False Then
                    ReceiveTimeZoneRefresh Mid$(sPacket, 2)
                End If
                
            Case sGravityZoneRefreshs
                
                If modStickGame.StickServer = False Then
                    ReceiveGravityZoneRefresh Mid$(sPacket, 2)
                End If
                
            Case sKillAndDeathInfos
                'format: (iKiller)#(iDeadStick)(bToasty)
                
                'On Error GoTo EH
                'i = FindStick(CInt(Mid$(sPacket, 2))) 'index of killer
                'j = FindStick(CInt(Mid$(sPacket, 3))) 'index of deadstick
                
                ProcessKillDeathMessage Mid$(sPacket, 2)
                
            Case sGrassRefreshs
                
                If modStickGame.StickServer = False Then
                    ReceiveGrassRefresh Mid$(sPacket, 2)
                End If
                
'            Case sFireRefreshs
'
'                If modStickGame.StickServer = False Then
'                    ReceiveGrassRefresh Mid$(sPacket, 2)
'                End If
                
                
            Case sHealthPacks
                'sHealthPacks & CStr(HealthPack.X) & "|" & CStr(HealthPack.Y)
                On Error GoTo EH
                Tmp = Mid$(sPacket, 2)
                
                i = InStr(1, Tmp, "|")
                
                If i Then
                    HealthPack.bActive = True
                    HealthPack.X = Mid$(Tmp, 1, i - 1)
                    HealthPack.Y = Mid$(Tmp, i + 1)
                End If
                
            Case sExplodeMines
                
                'If Not modStickGame.StickServer Then
                    On Error GoTo EH
                    
                    j = CInt(Mid$(sPacket, 2)) 'mine ID
                    
                    For i = 0 To NumMines - 1
                        If Mine(i).ID = j Then
                            ExplodeMine i, modStickGame.StickServer
                            RemoveMine i
                            Exit For
                        End If
                    Next i
                    
                'End If
                
            Case sExplodeBarrels
                
                'If Not modStickGame.StickServer Then
                    
                    On Error GoTo EH
                    j = CInt(Mid$(sPacket, 2)) 'barrel ID
                    
                    'AddMainMessage "Explode Barrel - " & j, False
                    
                    For i = 0 To NumBarrels - 1
                        If Barrel(i).ID = j Then
                            'AddMainMessage "Barrel " & j & " Found", False
                            ExplodeBarrel i, modStickGame.StickServer
                            RemoveBarrel i
                            'AddMainMessage "Barrel " & j & " Exploded + Removed", False
                            Exit For
                        End If
                    Next i
                    
                'End If
                
            Case sStaticWeaponUpdates
                
                If Not modStickGame.StickServer Then
                    If Len(sPacket) = 1 Then
                        'no static weapons
                        If NumStaticWeapons Then
                            RemoveStaticWeapons
                        End If
                    Else
                        ProcessStaticWeaponPacket Mid$(sPacket, 2)
                    End If
                End If
                
                
            Case sRoundInfos
                
                If Not StickServer Then
                    ReceivedRoundInfo Mid$(sPacket, 2)
                End If
                
                
            Case sPresences
                
                'packet = ID of stick
                On Error GoTo EH
                i = FindStick(CInt(Mid$(sPacket, 2)))
                
                If i > -1 Then
                    Stick(i).LastPacket = GetTickCount()
                    
                    If modStickGame.StickServer Then
                        SendBroadcast sPacket, Stick(i).ID
                    End If
                    
                End If
                
            Case sWeaponSwapInfos
                
                ProcessWeaponSwapInfo Mid$(sPacket, 2)
                
            Case sJoins
                'A join mPacket.  If we're a StickServer, handle it
                If modStickGame.StickServer Then ProcessJoinPacket TempSockAddr
                
                
            Case sExits
                
                On Error GoTo EH
                j = CInt(Mid$(sPacket, 2)) 'id of quitter
                
                i = FindStick(j)
                
                If i <> -1 Then
                    If Stick(i).ID = 0 Then 'server quit
                        If modStickGame.StickServer = False Then 'make sure it's not a warped packet
                            AddText "Server Quit the Game", TxtError, True
                            GetPacket = False
                            Exit Function
                        End If
                    Else
                        AddChatText Trim$(Stick(i).Name) & " left the game", Stick(i).colour
                        
                        If modStickGame.StickServer Then
                            SendBroadcast sExits & CStr(j), j
                            Pause 5
                        End If
                        
                        RemoveStick i
                    End If
                End If
                
                
            Case sNewMaps
                'mid = map name & vbspace & square
                If Not modStickGame.StickServer Then
                    ProcessNewMap Mid$(sPacket, 2)
                End If
                
            Case sMapRequests
                
                If modStickGame.StickServer Then
                    
                    modWinsock.SendPacket lSocket, TempSockAddr, _
                        sMapNames & GetFileName(modStickGame.StickMapPath) & vbSpace & CStr(MakeSquareNumber())
                    
                End If
                
                
            Case sKicks
                
                If Not modStickGame.StickServer Then
                    AddText "Disconnected - Was Kicked" & IIf(LenB(Mid$(sPacket, 2)) > 0, _
                        " (" & Mid$(sPacket, 2) & ")", vbNullString), TxtError, True
                    
                    bRunning = False
                    GetPacket = False
                    Unload Me
                    Exit Function
                End If
                
        End Select
    End If
Loop While LenB(sPacket)

Exit Function
EH:
'MsgBox "Error - " & Err.Description & vbNewLine & _
    "tmp = " & Tmp & vbNewLine & _
    "stxt = " & sTxt

End Function

Private Sub SendKillDeathMessage(iKiller As Integer, iDeadStick As Integer, bToasty As Boolean)

SendBroadcast sKillAndDeathInfos & CStr(iKiller) & "#" & CStr(iDeadStick) & CStr(Abs(bToasty))

End Sub

Private Sub ProcessKillDeathMessage(sPacket As String)
Dim i As Integer, iKiller As Integer, iDead As Integer
Dim bToasty As Boolean


'FORMAT: (ID_Killer)#(ID_DeadStick)(bToasty)

On Error GoTo EH

i = InStr(1, sPacket, "#")

'index of killer
iKiller = FindStick(CInt( _
    Left$(sPacket, i - 1) _
    ))

'index of deadstick
iDead = FindStick(CInt( _
    Mid$(sPacket, i + 1, Len(sPacket) - i - 1) _
    ))

bToasty = CBool(Right$(sPacket, 1))


If iKiller <> -1 And iDead <> -1 Then
    
    If iKiller > 0 Or modStickGame.StickServer Then
        Stick(iKiller).iKills = Stick(iKiller).iKills + 1
        
        Stick(iKiller).iKillsInARow = Stick(iKiller).iKillsInARow + 1
        
        'trust me
    ElseIf Not modStickGame.StickServer Then
        Stick(iKiller).iKillsInARow = Stick(iKiller).iKillsInARow + 1
        
    End If
    
    
    If iKiller = 0 Then
        
        If Stick(0).WeaponType = FlameThrower Then
            'If Stick(0).LastBullet + 750 > GetTickCount() Then
            FlamesInARow = FlamesInARow + 1
            'End If
        End If
        
        
        Call CheckKillsInARow
    End If
    
    'killer stick above
'######################################################################################################
    'dead stick below
    
    If iDead > 0 Then
        'don't do me twice
        SomeoneDied iDead, iKiller, IIf(bToasty, eKillTypes.kFlame, eKillTypes.kNormal)
    End If
    
    Stick(iDead).iKillsInARow = 0
    Stick(iDead).LastSpawnTime = GetTickCount()
    
    
    If modStickGame.sv_GameType = gCoOp Or modStickGame.sv_GameType = gElimination Then
        Stick(iDead).bAlive = False
    End If
    
    If iDead > 0 Then 'not me
        Stick(iDead).iDeaths = Stick(iDead).iDeaths + 1
        
        'dead stick/chopper added in someonedied()
'        If Stick(iDead).WeaponType = Chopper Then
'            AddDeadChopper Stick(iDead).X, Stick(iDead).Y, Stick(iDead).Colour, iDead
'        Else
'            AddDeadStick Stick(iDead).X, Stick(iDead).Y, Stick(iDead).Colour, (Stick(iDead).Facing < Pi), bToasty, _
'                Stick(iDead).Speed, Stick(iDead).Heading
'        End If
    End If
    
    
    
    'reset some stuff
    SubStickiState iDead, STICK_LEFT
    SubStickiState iDead, STICK_RIGHT
    Stick(iDead).bOnSurface = False
    ResetStickFireAndFlash iDead
    
    
    If modStickGame.StickServer Then 'tell everyone else, and send it back to home stick
        SendBroadcast sKillAndDeathInfos & sPacket
    End If
    
End If


EH:
End Sub

Public Sub ProcessNewMap(sData As String)
Dim sMapName As String, sMapPath As String
Dim i As Integer

i = InStrRev(sData, vbSpace)
On Error GoTo EH
If i Then
    If IsValidVarPacket(sData) Then
        sMapName = Left$(sData, i - 1)
        
        If sMapName <> GetFileName(modStickGame.StickMapPath) Then
            
            sMapPath = modStickGame.GetStickMapPath() & sMapName
            If FileExists(sMapPath) Then
                
                If LoadMapEx(sMapPath) Then
                    AddMainMessage "Server Changed Map - " & sMapName, False
                    
                    modWinsock.SendPacket lSocket, ServerSockAddr, sNewMaps & CStr(1) 'ACK
                    
                    LastUpdatePacket = GetTickCount() + 10000 'prevent lagging out
                Else
                    modWinsock.SendPacket lSocket, ServerSockAddr, sNewMaps & CStr(2) 'ACK^-1
                    
                    AddText "Server Changed Map - Error Loading New Map - " & Err.Description, TxtError, True
                    bRunning = False
                    Unload Me
                End If
                
            Else
                
                AddText "Server Changed Map - New Map Not Found", TxtError, True
                bRunning = False
                Unload Me
                
            End If
        Else
            'ack that we have the map
            'may have received a second sNewMaps packet
            modWinsock.SendPacket lSocket, ServerSockAddr, sNewMaps & CStr(1) 'ACK
        End If
    End If
End If

EH:
End Sub

'########################################################################

Private Sub ReceiveDamageTick()

If modStickGame.cl_DamageTick Then
    modAudio.PlayTickSound
    LastDamageTick = GetTickCount()
End If

End Sub
Private Sub DrawDamageTick()
Const OuterDist = 150, InnerDist = 100, Max_Byte = 255
Dim lCol As Long, sFactor As Single

If modStickGame.cl_DamageTick Then
    sFactor = (LastDamageTick + DamageTickTime - GetTickCount()) / DamageTickTime
    
    If sFactor > 0.5 Then
        
        lCol = RGB(sFactor * Max_Byte, sFactor * Max_Byte, sFactor * Max_Byte)
        
        picMain.DrawWidth = 2
        picMain_Line MouseX - OuterDist, MouseY - OuterDist, _
                     MouseX - InnerDist, MouseY - InnerDist, _
                     lCol
        
        picMain_Line MouseX - OuterDist, MouseY + OuterDist, _
                     MouseX - InnerDist, MouseY + InnerDist, _
                     lCol
        
        picMain_Line MouseX + OuterDist, MouseY - OuterDist, _
                     MouseX + InnerDist, MouseY - InnerDist, _
                     lCol
        
        picMain_Line MouseX + OuterDist, MouseY + OuterDist, _
                     MouseX + InnerDist, MouseY + InnerDist, _
                     lCol
        
        
    End If
End If

End Sub

Private Sub picMain_Line(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, lColour As Long)
picMain.Line (X1, Y1)-(X2, Y2), lColour
End Sub

'########################################################################

Private Sub CheckStickNames()
Static LastCheck As Long
Dim i As Integer, j As Integer

If LastCheck + NameCheckDelay < GetTickCount() Then
    On Error GoTo lEH
    
    For i = 0 To NumSticksM1
        For j = NumSticksM1 To 0 Step -1
            
            If j <> i Then
                If Stick(j).Name = Stick(i).Name Then
                    'Kick Stick j
                    modWinsock.SendPacket lSocket, Stick(j).SockAddr, sKicks & "Same Name"
                    Exit Sub 'so we don't get errors, will be checked again
                End If
            End If
            
        Next j
    Next i
    
    LastCheck = GetTickCount()
End If

lEH:
End Sub


Private Sub ProcessJoinPacket(vSockAddr As ptSockAddr)

Dim i As Long
Dim ID As String
Dim Index As Integer
Dim MaxID As Integer

'If this IP address is already in our Stick array, use pre-assigned ID
For i = 0 To NumSticksM1
    'Is it the same IP and port?
    If (Stick(i).SockAddr.sin_addr = vSockAddr.sin_addr) And _
                (Stick(i).SockAddr.sin_port = vSockAddr.sin_port) Then
        
        ID = CStr(Stick(i).ID)
        
        Exit For
    End If
Next i

'New Stick?
If LenB(ID) = 0 And (vSockAddr.sin_addr <> 0) Then
    'Make a spot
    Index = AddStick() 'ID assigned
'    'Find a new ID
'    MaxID = 0
'    For i = 0 To NumSticksM1
'        'Is this ID greater?
'        If Stick(i).ID > MaxID Then MaxID = Stick(i).ID
'    Next i
'    'Assign the ID
'    Stick(Index).ID = MaxID + 1
    
    'Set the Stick's ptsockaddr
    Stick(Index).SockAddr.sin_addr = vSockAddr.sin_addr
    Stick(Index).SockAddr.sin_family = vSockAddr.sin_family
    Stick(Index).SockAddr.sin_port = vSockAddr.sin_port
    Stick(Index).SockAddr.sin_zero = vSockAddr.sin_zero
    
    'Set the ID String
    ID = CStr(Stick(Index).ID)
End If

'Send the ACK
If (vSockAddr.sin_addr <> 0) Then
    modWinsock.SendPacket lSocket, vSockAddr, sAccepts & ID
    LastGrassRefresh = GetTickCount() - 60000
    '                             one min ^
End If

End Sub

Private Sub ReceiveBoxInfo(sTxt As String)
Dim i As Integer

'format: 10101101
'1 = present
'0 = gone

On Error Resume Next
For i = 0 To ubdBoxes
    Box(i).bInUse = CBool(Mid$(sTxt, i + 1, 1))
Next i

'if lenb(tag) = 0 then showbox

End Sub

Private Sub SendBoxInfo()
Static LastSend As Long
Dim i As Integer
Dim sPacketToSend As String

If LastSend + BoxInfoDelay < GetTickCount() Then
    
    For i = 0 To ubdBoxes
        sPacketToSend = sPacketToSend & Abs(Box(i).bInUse)
    Next i
    
    SendBroadcast sBoxInfos & sPacketToSend
    
    LastSend = GetTickCount()
End If

End Sub

Public Sub SendChatPacket(ChatText As String, colour As Long)

'Is this the StickServer?
If StickServer Then
    'Broadcast the chat mPacket
    SendChatPacketBroadcast ChatText, colour
Else
    'Send it to the StickServer
    modWinsock.SendPacket lSocket, ServerSockAddr, sChats & ChatText & "#" & CStr(colour)
End If

End Sub

Public Sub SendChatPacketBroadcast(ChatText As String, colour As Long)

Dim i As Long

'Send the mPacket to everyone but the local user
For i = 1 To NumSticksM1
    If Stick(i).IsBot = False Then
        modWinsock.SendPacket lSocket, Stick(i).SockAddr, sChats & ChatText & "#" & CStr(colour)
    End If
Next i

'Add text to local user's chat text array
AddChatText ChatText, colour

End Sub

Public Sub SendBroadcast(Text As String, Optional ByVal NtID As Integer = -1)

Dim i As Long

'Send the mPacket to everyone but the local user
For i = 1 To NumSticksM1
    'Is this the local user?
    If Stick(i).ID <> NtID Then 'stick(i).ID <> MyID And
        If Stick(i).IsBot = False Then
            modWinsock.SendPacket lSocket, Stick(i).SockAddr, Text
        End If
    End If
Next i


End Sub

Private Sub AddChatText(ChatText As String, colour As Long)
Dim nChat As Integer, i As Integer

'Add this value to the chat text array
ReDim Preserve Chat(NumChat)


With Chat(NumChat)
    .Decay = GetTickCount() + CHAT_DECAY
    .Text = ChatText
    .colour = colour
    .sTextHeight = TextHeight(ChatText)
    .sTextWidth = TextWidth(ChatText)
    
    
    If InStr(1, ChatText, modMessaging.MsgNameSeparator) Then
        .bChatMessage = True
        Call CheckForSticksChat(ChatText)
        
        For i = 0 To NumChat - 1
            If Chat(i).bChatMessage Then
                nChat = nChat + 1
            End If
        Next i
        
        
        AddAttention Chat_X_Offset + .sTextWidth / 2, _
            IIf(bPlaying, nChat, nChat + Chat_Round_Offset) * .sTextHeight + Chat_Chat_Offset + .sTextHeight / 2, _
            .colour
        
    Else
        Call CheckForKills(ChatText)
    End If
    
End With

NumChat = NumChat + 1


End Sub

Private Sub RemoveChatText(Index As Integer)

Dim i As Long

'Remove the specified chat text
For i = Index To NumChat - 2
'    Chat(i).Decay = Chat(i + 1).Decay
'    Chat(i).Text = Chat(i + 1).Text
'    Chat(i).Colour = Chat(i + 1).Colour
    Chat(i) = Chat(i + 1)
Next i

'Resize the array
If NumChat = 1 Then
    Erase Chat
    NumChat = 0
Else
    ReDim Preserve Chat(NumChat - 2)
    NumChat = NumChat - 1
End If
    
End Sub

Private Sub AddAttention(aX As Single, aY As Single, lCol As Long)

ReDim Preserve Attention(NumAttentions)

With Attention(NumAttentions)
    .Decay = GetTickCount() + Attention_Time
    .X = aX
    .Y = aY
    .lColour = lCol
End With

NumAttentions = NumAttentions + 1

End Sub

Private Sub DrawAttentions()
Dim i As Integer
Dim GTC As Long

GTC = GetTickCount()
picMain.DrawWidth = 2

For i = 0 To NumAttentions - 1
    With Attention(i)
        
        picMain.Circle (.X, .Y), _
            Abs(.Decay - GTC), _
            .lColour
        
    End With
Next i

picMain.DrawWidth = 1


End Sub

Private Sub ProcessAttentions()
Dim i As Integer
Dim GTC As Long

GTC = GetTickCount()

Do While i < NumAttentions
    
    If Attention(i).Decay < GTC Then
        RemoveAttention i
        i = i - 1
    End If
    
    i = i + 1
Loop

End Sub

Private Sub RemoveAttention(Index As Integer)

Dim i As Long

For i = Index To NumAttentions - 2
    Attention(i) = Attention(i + 1)
Next i

If NumAttentions = 1 Then
    Erase Attention
    NumAttentions = 0
Else
    ReDim Preserve Attention(NumAttentions - 2)
    NumAttentions = NumAttentions - 1
End If
    
End Sub

Private Sub AddHead(aX As Single, aY As Single, lCol As Long, Speed As Single, Heading As Single)

ReDim Preserve Head(NumHeads)

With Head(NumHeads)
    .Decay = GetTickCount() + Head_Time
    .X = aX
    .Y = aY
    .lColour = lCol
    .Speed = Speed
    .Heading = Heading
End With

NumHeads = NumHeads + 1

End Sub

Private Sub DrawHeads()
Dim i As Integer

picMain.FillStyle = vbFSSolid
picMain.DrawWidth = 1

For i = 0 To NumHeads - 1
    With Head(i)
        picMain.FillColor = .lColour
        modStickGame.sCircle .X, .Y, HeadRadius, .lColour
    End With
Next i

picMain.FillStyle = vbFSTransparent


End Sub

Private Sub ProcessHeads()
Dim i As Integer
Dim GTC As Long
Dim sTz As Single

GTC = GetTickCount()

Do While i < NumHeads
    
    If Head(i).Decay < GTC Then
        RemoveHead i
        i = i - 1
    ElseIf Head(i).Speed Then
        With Head(i)
            MotionStickObject Head(i).X, Head(i).Y, Head(i).Speed, Head(i).Heading
            
            ClipHead i
            
            sTz = GetTimeZoneAdjust(.X, .Y)
            ApplyGravityVector .LastGravity, sTz, .Speed, .Heading, .X, .Y
            
            
        End With
    End If
    
    i = i + 1
Loop

End Sub

Private Function HeadCollision(ByVal i As Integer, _
    oLeft As Single, oTop As Single, oWidth As Single, oHeight As Single) As Boolean

If Head(i).X >= oLeft - Lim Then
    If Head(i).X <= (oLeft + oWidth + Lim) Then
        If Head(i).Y >= oTop Then
            If Head(i).Y <= (oTop + oHeight) Then
                HeadCollision = True
            End If
        End If
    End If
End If

End Function

Private Sub ClipHead(i As Integer)
Dim j As Integer

For j = 0 To modStickGame.ubdPlatforms
    If HeadCollision(i, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        
        If Head(i).Heading > 0 Then
            Head(i).Y = Platform(j).Top - 1
        Else
            Head(i).Y = Platform(j).Top + Platform(j).height + 1
        End If
        
        ReverseYComp Head(i).Speed, Head(i).Heading
        Head(i).Speed = Head(i).Speed * Head_Bounce_Reduction
        
        If Head(i).Speed < 30 Then Head(i).Speed = 0
        
        Exit Sub
    End If
Next j



If Head(i).X < Lim Then
    
    Head(i).X = Lim
    
    ReverseXComp Head(i).Speed, Head(i).Heading
    Head(i).Speed = Head(i).Speed * Head_Bounce_Reduction
    
ElseIf Head(i).X > StickGameWidth - Lim Then
    
    Head(i).X = StickGameWidth - Lim
    
    ReverseXComp Head(i).Speed, Head(i).Heading
    Head(i).Speed = Head(i).Speed * Head_Bounce_Reduction
    
End If
If Head(i).Y < 1 Then
    Head(i).Y = 1
    ReverseYComp Head(i).Speed, Head(i).Heading
End If

End Sub

Private Sub RemoveHead(Index As Integer)

Dim i As Long

For i = Index To NumHeads - 2
    Head(i) = Head(i + 1)
Next i

If NumHeads = 1 Then
    Erase Head
    NumHeads = 0
Else
    ReDim Preserve Head(NumHeads - 2)
    NumHeads = NumHeads - 1
End If
    
End Sub

Private Sub EndWinsock()

'Kill winsock
modWinsock.DestroySocket lSocket
'modWinsock.TermWinsock

End Sub

Private Function CanMoveControl(Ctrl As Control) As Boolean

If (TypeOf Ctrl Is Timer) = False Then
    If Not TypeOf Ctrl Is Menu Then
        If Ctrl.Visible Then
            If Ctrl.Name = oPlatform(0).Name Then
                CanMoveControl = (Ctrl.Index > 0)
            ElseIf Ctrl.Name = picHandle(0).Name Then
                CanMoveControl = False
            Else
                CanMoveControl = True
            End If
        End If
    End If
End If
         
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bInPosition As Boolean

If modStickGame.bStickEditing Then
    Dim i As Integer

    If Button = vbLeftButton Then
        'Hit test over light-weight (non-windowed) controls
        For i = 0 To (Controls.Count - 1)
            'Check for visible, non-menu controls
            '[Note 1]
            'If any of the sizing handle controls are under the mouse
            'pointer, then they must not be visible or else they would
            'have already intercepted the MouseDown event.
            '[Note 2]
            'This code will fail if you have a control such as the
            'Timer control which has no Visible property. You will
            'either need to make sure your form has no such controls
            'or add code to handle them.
            If CanMoveControl(Controls(i)) Then
                m_DragRect.SetRectToCtrl Controls(i)
                If m_DragRect.PtInRect(X, Y) Then
                    DragBegin Controls(i)
                    Exit Sub
                End If
            End If
        Next i
        
        DragEnd
    End If
Else
    On Error Resume Next
    If bHasFocus Then
        If Button = vbLeftButton Then
            If StickInGame(0) Then
                If Stick(0).BulletsFired < GetMaxRounds(Stick(0).WeaponType) Then
                    If StickiHasState(0, STICK_RELOAD) = False Then
                    
                        If WeaponIsSniper(Stick(0).WeaponType) Then
                            If Stick(0).Perk <> pSniper Or Stick(0).WeaponType = M82 Then
                                
                                bInPosition = (StickiHasState(0, STICK_CROUCH) Or StickiHasState(0, STICK_PRONE)) And Stick(0).bOnSurface
                                
                                If (StickIsMoving(0) Or Stick(0).bOnSurface = False) And Not bInPosition Then
                                    AddMainMessage "You can't shoot while moving", True
                                Else
                                    If bInPosition Then
                                        FireKey = True
                                    Else
                                        AddMainMessage "Go to crouch or prone to fire", True
                                    End If
                                End If
                            Else
                                FireKey = True
                            End If
                        Else
                            FireKey = True
                        End If
                        
                            
                    ElseIf WeaponIsShotgun(Stick(0).WeaponType) Then
                        If StickiHasState(0, STICK_RELOAD) Then
                            SubStickiState 0, STICK_RELOAD
                            
                            'add below to fire immediatly
                            'FireKey = True
                        End If
                    Else
                        FireKey = False
                    End If
                Else
                    FireKey = False
                End If
            Else
                FireKey = False
            End If
            
            
        ElseIf Button = vbRightButton Then
            If StickInGame(0) Then
                If Stick(0).Perk <> pZombie Then
                    'If Stick(0).WeaponType <> Chopper Then
                    AddStickiState 0, STICK_NADE
                    Stick(0).NadeStart = GetTickCount()
                End If
            End If
            
        ElseIf Button = vbMiddleButton Then
            If StickInGame(0) Then
                If Stick(0).Perk <> pZombie Then
                    If modStickGame.cl_MiddleMineDrop Then
                        AddStickiState 0, STICK_MINE
                    End If
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If modStickGame.bStickEditing Then
    If Button = vbLeftButton Then
        If m_DragState = StateDragging Or m_DragState = StateSizing Then
            'Hide drag rectangle
            DrawDragRect
            
'            bCan = True
'
'            If m_CurrCtl.Left < 0 Or (m_CurrCtl.Left * Screen.TwipsPerPixelX) > Me.Width Then
'                m_CurrCtl.Left = Me.Width / (2 * Screen.TwipsPerPixelX)
'                bCan = False
'            End If
'
'            If m_CurrCtl.Top < 0 Or m_CurrCtl.Top > Me.Height Then
'                m_CurrCtl.Top = Me.Height / 2 - m_CurrCtl.Height / 2
'                bCan = False
'            End If
            
            'If bCan Then
                'Move control to new location
                m_DragRect.ScreenToTwips m_CurrCtl
                m_DragRect.SetCtrlToRect m_CurrCtl
            'End If
            
            'Restore sizing handles
            ShowHandles True
            
            'Free mouse movement
            ClipCursor ByVal 0&
            'Release mouse capture
            ReleaseCapture
            'Reset drag state
            m_DragState = StateNothing
            
            
        End If
    End If
Else
    If Button = vbLeftButton Then
        FireKey = False
        FireKeyUpTime = GetTickCount()
    End If
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If modStickGame.bStickEditing Then
    Dim nWidth As Single, nHeight As Single
    Dim pt As PointAPI
    
    
    If m_DragState = StateDragging Then
        'Save dimensions before modifying rectangle
        nWidth = m_DragRect.Right - m_DragRect.Left
        nHeight = m_DragRect.Bottom - m_DragRect.Top
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Update drag rectangle coordinates
        m_DragRect.Left = pt.X - m_DragPoint.X
        m_DragRect.Top = pt.Y - m_DragPoint.Y
        m_DragRect.Right = m_DragRect.Left + nWidth
        m_DragRect.Bottom = m_DragRect.Top + nHeight
        'Draw new rectangle
        DrawDragRect
        
        'bSaved = False
        
    ElseIf m_DragState = StateSizing Then
        'Get current mouse position in screen coordinates
        GetCursorPos pt
        'Hide existing rectangle
        DrawDragRect
        'Action depends on handle being dragged
        Select Case m_DragHandle
            Case 0
                m_DragRect.Left = pt.X
                m_DragRect.Top = pt.Y
            Case 1
                m_DragRect.Top = pt.Y
            Case 2
                m_DragRect.Right = pt.X
                m_DragRect.Top = pt.Y
            Case 3
                m_DragRect.Right = pt.X
            Case 4
                m_DragRect.Right = pt.X
                m_DragRect.Bottom = pt.Y
            Case 5
                m_DragRect.Bottom = pt.Y
            Case 6
                m_DragRect.Left = pt.X
                m_DragRect.Bottom = pt.Y
            Case 7
                m_DragRect.Left = pt.X
        End Select
        'Draw new rectangle
        DrawDragRect
    End If
Else
    MouseX = X
    MouseY = Y
End If


'On error GoTo EH
'With Stick(0)
'    If (.State And Stick_Reload) = 0 Then
'        'If .WeaponType = AK Or .WeaponType = Knife Or _
'            .WeaponType = XM8 Or .WeaponType = M249 Or .WeaponType = DEagle Then
'
'        If .WeaponType <> M82 Then
'            If .WeaponType <> RPG Then
'                If .WeaponType <> W1200 Then
'                    If .WeaponType <> DEagle Then
'                        SetMyStickFacing
'                    End If
'                End If
'            End If
'        End If
'    End If
'End With
'
'EH:
End Sub

Public Sub Form_WheelScroll(bScrollUp As Boolean)
'WARNING - SUBCLASSED

If NumSticks Then
    'If modStickGame.sv_2Weapons Then
    Form_KeyPress vbKey2
'    Else
'        If Stick(0).WeaponType <> Chopper Then
'            If StickiHasState(0, Stick_Reload) = False Then
'                Scroll_WeaponKey = Scroll_WeaponKey + IIf(bScrollUp, -1, 1)
'
'                If Scroll_WeaponKey = -1 Then
'                    Scroll_WeaponKey = Knife
'                ElseIf Scroll_WeaponKey > Knife Then
'                    Scroll_WeaponKey = AK
'                End If
'
'                LastScrollWeaponSwitch = GetTickCount()
'            End If
'        End If
'    End If
End If

End Sub

Private Sub DoMyStickFacing()
Static bFacingClipped As Boolean
Const Sniper_Edge_Limit As Single = 128

If Stick(0).LastBullet + (GetBulletDelay(0) - 60) / GetMyTimeZone() < GetTickCount() Then
    SetMyStickFacing
End If


If StickiHasState(0, STICK_PRONE) Then
    'prevent through-floor shooting
    If Stick(0).bOnSurface Then
        If Stick(0).Facing > Pi Then
            
            'facing left
            If Stick(0).Facing < ProneLeftLimit Then
                'allow looking down if on the left of a platform
                If Stick(0).X > Platform(Stick(0).iCurrentPlatform).Left + Sniper_Edge_Limit Then
                    Stick(0).Facing = ProneLeftLimit
                    bFacingClipped = True
                Else
                    bFacingClipped = False
                End If
            Else
                bFacingClipped = False
            End If
            
        Else
            
            'facing right
            If Stick(0).Facing > ProneRightLimit Then
                'allow looking down if on the left of a platform
                If Stick(0).X < (Platform(Stick(0).iCurrentPlatform).Left + _
                        Platform(Stick(0).iCurrentPlatform).width - Sniper_Edge_Limit) Then
                    
                    Stick(0).Facing = ProneRightLimit
                    bFacingClipped = True
                Else
                    bFacingClipped = False
                End If
            Else
                bFacingClipped = False
            End If
        End If
    Else
        bFacingClipped = False
    End If
    
    
    If bFacingClipped Then
        If Stick(0).bFlashed = False Then
            PrintStickFormText "Can't Aim Lower", MouseX - 500, MouseY + 240, vbRed
            Stick(0).ActualFacing = Stick(0).Facing + 0.01
        End If
    End If
Else
    bFacingClipped = False 'Static
End If


End Sub

Private Sub SetMyStickFacing()
Const HeadRadiusX2 = HeadRadius * 2, M82_Min_Accurate_Dist = 5000, BodyLenX1p3 = BodyLen * 1.3
Const KnifeLockAmount As Single = 0.4, Pi2LessLockAmount = Pi2 - KnifeLockAmount, LeftFacingAngle = Pi2 - 0.0001
Dim X As Single, Y As Single
Dim bDoNormal As Boolean

'Stick(0).Facing = FindAngle(Stick(0).x, Stick(0).y + HeadRadius / 1.5, MouseX, MouseY)

'Stick(0).Facing = FindAngle(Stick(0).x * cg_sZoom - cg_sCamera.x, _
                             Stick(0).y * cg_sZoom - cg_sCamera.y, _
                             MouseX, _
                             MouseY)

#If Hack_AimBot Then
    Dim i As Integer
    'Dim tX As Long, tY As Long
    'Const yOffset = BodyLen * 1.6
    
    i = ClosestTargetI(0, 0)
    
    If i > -1 Then
        'Stick(0).Facing = FindAngle_Actual(Stick(0).X, Stick(0).Y, Stick(i).X, Stick(i).Y)
        'tX = Stick(i).X * cg_sZoom - cg_sCamera.X
        'tY = Stick(i).Y * cg_sZoom - cg_sCamera.Y + yOffset
        
        'If 0 < tX And tX < Screen.width Then
            'If 0 < tY And tY < Screen.height Then
                'SetCursorPos ScaleX(tX, vbTwips, vbPixels), ScaleY(tY, vbTwips, vbPixels)
            'End If
        'End If
        
        AccurateShot Stick(i).X, Stick(i).Y, Stick(i).Speed, Stick(i).Heading, _
                    Stick(0).X, Stick(0).Y, Stick(0).Speed, Stick(0).Heading, _
                    BULLET_SPEED, 0, Stick(0).Facing
        
    End If
#Else
    If StickiHasState(0, STICK_PRONE) Then 'And Stick(0).WeaponType <> AWM Then
        
        Stick(0).Facing = FixAngle(FindAngle_Actual(Stick(0).X * cg_sZoom - cg_sCamera.X, _
                                 (Stick(0).Y + BodyLenX1p3) * cg_sZoom - cg_sCamera.Y, _
                                 MouseX, _
                                 MouseY))
        
        
    ElseIf Stick(0).WeaponType <> Chopper Then
        
        If WeaponIsSniper(Stick(0).WeaponType) Then 'Stick(0).WeaponType = AWM Then
            'If modStickGame.cl_SniperScope Then
                If StickiHasState(0, STICK_RELOAD) = False Then
                    
                    X = Stick(0).GunPoint.X * cg_sZoom - cg_sCamera.X
                    Y = Stick(0).GunPoint.Y * cg_sZoom - cg_sCamera.Y
                    
                    If GetDist(X, Y, MouseX, MouseY) > M82_Min_Accurate_Dist Then
                        
                        Stick(0).Facing = FixAngle(FindAngle_Actual(X, Y, MouseX, MouseY))
                    Else
                        bDoNormal = True
                    End If
                Else
                    bDoNormal = True
                End If
            'Else
                'bDoNormal = True
            'End If
        Else
            bDoNormal = True
        End If
        
        
        If bDoNormal Then
            
            Stick(0).Facing = Round(FixAngle(FindAngle_Actual(Stick(0).X * cg_sZoom - cg_sCamera.X, _
                                         (Stick(0).Y + HeadRadiusX2) * cg_sZoom - cg_sCamera.Y, _
                                         MouseX, _
                                         MouseY)), 2)
            
            
        End If
        
        'Stick(0).Facing = piD2
        
    Else
        'chopper facing
        Stick(0).Facing = FindAngle_Actual((Stick(0).X - CLD6) * cg_sZoom - cg_sCamera.X, _
                                 (Stick(0).Y + CLD4) * cg_sZoom - cg_sCamera.Y, _
                                 MouseX, _
                                 MouseY)
    End If
#End If


If Stick(0).WeaponType = Knife Then
    If Stick(0).Facing < KnifeLockAmount Then '10 degrees
        Stick(0).Facing = 0
    ElseIf Stick(0).Facing > Pi2LessLockAmount Then
        Stick(0).Facing = LeftFacingAngle
    End If
End If

Stick(0).ActualFacing = Stick(0).Facing

End Sub

'Private Function GetStickFacing() As Single 'i As Integer) As Single
'
'If StickiHasState(0, Stick_Prone) Then
'    GetStickFacing = FindAngle(Stick(0).X * cg_sZoom - cg_sCamera.X, _
'                             (Stick(0).Y + BodyLen * 1.3) * cg_sZoom - cg_sCamera.Y, _
'                             MouseX, _
'                             MouseY)
'
'Else
'    GetStickFacing = FindAngle(Stick(0).X * cg_sZoom - cg_sCamera.X, _
'                             Stick(0).Y * cg_sZoom - cg_sCamera.Y, _
'                             MouseX, _
'                             MouseY)
'End If
'
'End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim sTmp As String

If modStickGame.bStickEditing Then
    If map_Changed Then
        If MsgBoxEx("Map has changed, are you sure you want to exit?", "Exiting will lose your map. You should save first", _
            vbYesNo + vbQuestion + vbDefaultButton2, "Exit Map Editor?", , , , , Me.hWnd) = vbYes Then
            
            DragEnd
        Else
            Cancel = True
        End If
    End If
Else
    
    'If modLoadProgram.IsIDE() = False Then
    If modSubClass.bStickSubClassing Then
        modSubClass.SubClassStick Me.hWnd, False
    End If
    
    bRunning = False
    
    If modStickGame.StickServer Then
        SendBroadcast sExits & Stick(0).ID
        
        
        sTmp = eCommands.LobbyCmd & eLobbyCmds.Remove & modStickGame.StickServerIP & "S"
        'remove from lobby
        If Server Then
            DataArrival sTmp
        Else
            SendData sTmp
        End If
    Else
        modWinsock.SendPacket lSocket, ServerSockAddr, sExits & Stick(0).ID
    End If
    
    Call ResetVars
    Call EndWinsock
End If

WindowClosing = True
Call FormLoad(Me, True)

modStickGame.StickFormLoaded = False
End Sub

Public Sub EraseSmoke()
NumSmoke = 0: Erase Smoke
End Sub
Public Sub EraseWallMarks()
NumWallMarks = 0: Erase WallMark
End Sub
Public Sub ReleaseDeadSticks()
Dim i As Integer

For i = 0 To NumDeadSticks - 1
    DeadStick(i).bOnSurface = False
    DeadStick(i).Speed = 0
    DeadStick(i).Heading = Pi
Next i

End Sub
Public Sub EraseDeadSticks()
NumDeadSticks = 0: Erase DeadStick
End Sub

Private Sub ResetVars()
Dim i As Integer

modDXSound.DXSound_Terminate

NumSticks = 0: Erase Stick
NumBullets = 0: Erase Bullet
NumSmoke = 0: Erase Smoke
NumBlood = 0: Erase Blood
NumNades = 0: Erase Nade
NumCasings = 0: Erase Casing
NumMines = 0: Erase Mine
NumMags = 0: Erase Mag
NumDeadChoppers = 0: Erase DeadChopper
NumSparks = 0: Erase Spark
NumFlames = 0: Erase Flame
NumChat = 0: Erase Chat
NumTimeZoneCircs = 0: Erase TimeZoneCircs
NumScreenCircs = 0: Erase ScreenCircs
NumMainMessages = 0: Erase MainMessages
NumStaticWeapons = 0: Erase StaticWeapon
'NumLargeSmokes = 0: Erase LargeSmoke
NumBarrels = 0: Erase Barrel
NumSmokeBlasts = 0: Erase SmokeBlast
NumTimeZones = 0: Erase TimeZone
NumGravityZones = 0: Erase GravityZone
NumCircleBlasts = 0: Erase CircleBlast
NumBulletTrails = 0: Erase BulletTrail
NumNadeTrails = 0: Erase NadeTrail
Erase Attention: NumAttentions = 0
Erase Head: NumHeads = 0
Erase ShieldWave: NumShieldWaves = 0
Erase Fire: NumFires = 0
EraseWallMarks
EraseDeadSticks

'NadesShot = 0


LastServerSettingVar = -1
LastUpdatePacket = 0

strChat = vbNullString
bChatActive = False

UseKey = False
CrouchKey = False

'KillsInARow = 0
FlamesInARow = 0
picToasty.Visible = False
'bHadRadar = False
For i = 0 To CInt(eWeaponTypes.Knife)
    AmmoFired(i) = 0
Next i

ChopperAvail = False
'RadarStartTime = 0

ResetKeys

modStickGame.sv_Hardcore = False
modStickGame.sv_StickGameSpeed = 1
'modStickGame.sv_2Weapons = True

End Sub

Private Function GetObjFacing(bX As Boolean, bIsLeft As Boolean, bIsTop As Boolean) As Single

'GetObjFacing = IIf(bX, IIf(bIsLeft, piD2, pi3D2), IIf(bIsTop, pi, 0))
If bX Then
    If bIsLeft Then
        GetObjFacing = piD2
    Else
        GetObjFacing = pi3D2
    End If
Else
    If bIsTop Then
        GetObjFacing = Pi
    'Else
        'GetObjFacing = 0
    End If
End If

End Function

Private Function ClipBullet(i As Integer) As Boolean

Const Lim As Integer = 50, Edge_In_Lim = Lim - 20
Const Bullet_SpeedX2 = BULLET_SPEED * 2
Const Sniper_Bullet_Diffract_Delay As Long = Bullet_Diffract_Delay / 2
Dim ClippedX As Boolean, ClippedY As Boolean, bSlowDownBullet As Boolean, bIsFastBullet As Boolean, bDeepImpact As Boolean
Dim bPlayRicochetAndMark As Boolean
Dim tX As Single, tY As Single, GenHeading As Single, Adj As Single

'is the bullet on the top, left, bottom or right of a wall?
Dim BulletIsLeft As Boolean, BulletIsTop As Boolean
Dim iPlatform As Integer

If Bullet(i).bSniperBullet Then
    If Bullet(i).Speed < 130 Then
        RemoveBullet i, False, False 'might as well not be fancy, nothing is different when speed<50
        ClipBullet = True
        Exit Function
    End If
ElseIf Bullet(i).Speed < 40 Then
    RemoveBullet i, False, False 'might as well not be fancy, nothing is different when speed<50
    ClipBullet = True
    Exit Function
End If


BulletIsLeft = (Bullet(i).X < Lim)
ClippedX = BulletIsLeft Or (Bullet(i).X > StickGameWidth - Lim)
BulletIsTop = (Bullet(i).Y < Lim)
ClippedY = BulletIsTop Or (Bullet(i).Y > StickGameHeight - Lim)

bIsFastBullet = (Bullet(i).Speed > Bullet_Min_Speed)

Adj = GetTimeZoneAdjust(Bullet(i).X, Bullet(i).Y)


If Bullet(i).Speed > Bullet_SpeedX2 Then
    Bullet(i).Speed = Bullet_SpeedX2
End If

If ClippedX Or ClippedY Then
    
    'force the bullet to be on the edge - all the same
    If ClippedX Then
        If BulletIsLeft Then
            Bullet(i).X = Edge_In_Lim
        Else
            Bullet(i).X = StickGameWidth - Edge_In_Lim
        End If
    Else
        If BulletIsTop Then
            Bullet(i).Y = Edge_In_Lim
        Else
            Bullet(i).Y = StickGameHeight - Edge_In_Lim
        End If
    End If
    
    ClipBullet = True
    RemoveBullet i, True ', GetObjFacing(ClippedX, BulletIsLeft, BulletIsTop)
    
ElseIf BulletInPlatform(i) Then ', BulletIsLeft, BulletIsTop, ClippedX) Then
    
    iPlatform = BulletIniPlatform(i)
    
    If Bullet(i).bHadCircleBlast = False Then
        If Rnd() > 0.8 Or Bullet(i).bSniperBullet Or Bullet(i).bDEagleBullet Then
            GenHeading = FixAngle(Bullet(i).Heading)
            
            If GenHeading < piD2 Or GenHeading > pi3D2 Then
                AddCircleBlast Bullet(i).X, Platform(iPlatform).Top + Platform(iPlatform).height, 1
            Else
                AddCircleBlast Bullet(i).X, Bullet(i).Y, -1
            End If
        End If
        Bullet(i).bHadCircleBlast = True
    End If
    
    
    If bIsFastBullet Then
        bSlowDownBullet = True
    End If
    
ElseIf BulletInTBox(i) Then
    
    'If Bullet(i).bSniperBullet Or Bullet(i).bShotgunBullet Then
        If bIsFastBullet Then
            bSlowDownBullet = True
        End If
    'Else
        'ClipBullet = True
        'RemoveBullet i, True
    'End If
    
ElseIf BulletInBox(i) Then
    
    
    If Bullet(i).LastDiffract + Bullet_Diffract_Delay / Adj < GetTickCount() Then
        
        bDeepImpact = (Stick(Bullet(i).OwnerIndex).Perk = pDeepImpact)
        
        If Bullet(i).bHeadingChanged = False Then
            If Bullet(i).bChopperBullet = False Then
                If Not bDeepImpact Then
                    Bullet(i).Damage = Bullet(i).Damage / 2
                End If
            End If
        End If
        
        
        If Bullet(i).bSniperBullet Or Bullet(i).bDEagleBullet Then
            Bullet(i).Speed = Bullet(i).Speed / IIf(bDeepImpact, 1.05, 1.2)
            
        Else
            '########################################################################################
            'Bullet is knocked off course/diffracted here
            '########################################################################################
            
            Bullet(i).Heading = Bullet(i).Heading + PM_Rnd() * IIf(bDeepImpact, piD8, piD4)
            Bullet(i).Speed = Bullet(i).Speed / IIf(bDeepImpact, 1.1, 2)
            'Bullet(i).Facing = Bullet(i).Heading
            Bullet(i).bHeadingChanged = True
            
            '########################################################################################
        End If
        
        
        Bullet(i).LastDiffract = GetTickCount()
    End If
    
    
ElseIf Bullet(i).Speed < Bullet_Min_Speed Then
    ClipBullet = True
    RemoveBullet i, False
    
End If



If bSlowDownBullet Then 'slow down sniper bullets ONLY
    
    tX = Bullet(i).X
    tY = Bullet(i).Y
    GenHeading = Bullet(i).Heading - Pi
    
    bDeepImpact = (Stick(Bullet(i).OwnerIndex).Perk = pDeepImpact)
    
    If Bullet(i).bSniperBullet Or Bullet(i).bDEagleBullet Then
        
        If Bullet(i).LastDiffract + Sniper_Bullet_Diffract_Delay / Adj < GetTickCount() Then
            bPlayRicochetAndMark = True
            
            With Bullet(i)
                If .LastDiffract = 0 Then
                    If Not bDeepImpact Then
                        .Damage = .Damage / 1.6 'M82_Wall_Damage
                    End If
                End If
                
                .Speed = .Speed / IIf(bDeepImpact, 1.2, 1.5)
                
                .LastDiffract = GetTickCount()
            End With
        End If
        
    ElseIf modStickGame.sv_BulletsThroughWalls Or Bullet(i).bShotgunBullet Or Stick(Bullet(i).OwnerIndex).Perk = pDeepImpact Then
        
        If Bullet(i).LastDiffract + Bullet_Wall_Diffract_Delay / Adj < GetTickCount() Then
            
            With Bullet(i)
                bPlayRicochetAndMark = True
                
                If .bChopperBullet = False Then
                    .Speed = .Speed / (1 + IIf(bDeepImpact, 1, 4) * Rnd())
                Else 'If .LastDiffract = 0 Then
                    .Speed = .Speed / Chopper_Impact_Speed_Dec
                End If
                
                
                If .LastDiffract = 0 Then
                    If Bullet(i).bChopperBullet = False Then
                        .Damage = .Damage / IIf(bDeepImpact, 1.4, 4)
                    End If
                End If
                
                .LastDiffract = GetTickCount()
                
            End With
        End If
    Else
        
        bPlayRicochetAndMark = True
        
        ClipBullet = True
        RemoveBullet i, False
        
    End If
    
    
    
    
    
    If bPlayRicochetAndMark Then
        
        AddSparks tX, tY, GenHeading
        AddWallMark tX, tY, WallMark_Bullet_Radius
        
        If Rnd() > 0.3 Then AddBulletExplosion tX, tY
        
        If Rnd() > 0.7 Then
            If PointHearableOnSticksScreen(tX, tY, 0) Then
                modAudio.PlayRicochet GetRelPan(tX)
            End If
        End If
        
        
    End If
    
End If


End Function

Private Function BulletInBox(i As Integer) As Boolean
Dim j As Integer

For j = 0 To ubdBoxes
    If Box(j).bInUse Then
        If BulletCollision(i, Box(j).Left, Box(j).Top, Box(j).width, Box(j).height) Then
            BulletInBox = True
            Exit For
        End If
    End If
Next j

End Function

Private Function BulletInTBox(i As Integer) As Boolean
Dim j As Integer

For j = 0 To ubdtBoxes
    If BulletCollision(i, tBox(j).Left, tBox(j).Top, tBox(j).width, tBox(j).height) Then
        BulletInTBox = True
        Exit For
    End If
Next j

End Function

Private Function BulletInPlatform(i As Integer) As Boolean ', ByRef bLeft As Boolean, ByRef bTop As Boolean, _
    ByRef bXClip As Boolean) As Boolean

Dim j As Integer
'Const LeftLim = 100

For j = 0 To ubdPlatforms
    If BulletCollision(i, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        BulletInPlatform = True
        
'        bLeft = (Bullet(i).X < (Platform(j).Left + Platform(j).width / 2))
'        bTop = (Bullet(i).Y < (Platform(j).Top + Platform(j).height / 2))
'
'        If Bullet(i).X < Platform(j).Left + LeftLim Then
'            bXClip = True
'        ElseIf Bullet(i).X > (Platform(j).Left + Platform(j).width - LeftLim) Then
'            bXClip = True
'        End If
        
        
        Exit For
    End If
Next j

End Function
Private Function BulletIniPlatform(i As Integer) As Integer
Dim j As Integer

BulletIniPlatform = -1

For j = 0 To ubdPlatforms
    If BulletCollision(i, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        BulletIniPlatform = j
        
'        bLeft = (Bullet(i).X < (Platform(j).Left + Platform(j).width / 2))
'        bTop = (Bullet(i).Y < (Platform(j).Top + Platform(j).height / 2))
'
'        If Bullet(i).X < Platform(j).Left + LeftLim Then
'            bXClip = True
'        ElseIf Bullet(i).X > (Platform(j).Left + Platform(j).width - LeftLim) Then
'            bXClip = True
'        End If
        
        
        Exit For
    End If
Next j

End Function

Private Function BulletCollision(ByVal i As Integer, _
    oLeft As Single, oTop As Single, oWidth As Single, oHeight As Single) As Boolean

If Bullet(i).X >= oLeft Then
    If (Bullet(i).X <= (oLeft + oWidth)) Then
        If Bullet(i).Y >= oTop Then
            If Bullet(i).Y <= (oTop + oHeight) Then
                BulletCollision = True
            End If
        End If
    End If
End If

End Function

Private Function NadeInPlatform(iNade As Integer) As Integer
Dim j As Integer

For j = 0 To ubdPlatforms
    If NadeCollision(iNade, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        NadeInPlatform = j
        Exit Function
    End If
Next j

NadeInPlatform = -1

End Function

Private Function NadeInBox(iNade As Integer) As Boolean
Dim j As Integer

For j = 0 To ubdBoxes
    If Box(j).bInUse Then
        If NadeCollision(iNade, Box(j).Left, Box(j).Top, Box(j).width, Box(j).height) Then
            NadeInBox = True
            'Box(j).Tag = "1"
            
            If Nade(iNade).iType = nFrag Then Box(j).bInUse = False
            Exit For
        End If
    End If
Next j

End Function

Private Function NadeInTBox(iNade As Integer) As Integer
Dim j As Integer

NadeInTBox = -1

For j = 0 To ubdtBoxes
    If NadeCollision(iNade, tBox(j).Left, tBox(j).Top, tBox(j).width, tBox(j).height) Then
        NadeInTBox = j
        Exit For
    End If
Next j

End Function

Private Function NadeInStick(iNade As Integer, iStick As Integer) As Boolean
    
If CoOrdInStick(Nade(iNade).X, Nade(iNade).Y, iStick) Then
    NadeInStick = True
End If

End Function

Private Function NadeCollision(ByVal i As Integer, _
    oLeft As Single, oTop As Single, oWidth As Single, oHeight As Single) As Boolean

If Nade(i).X >= (oLeft - Lim) Then
    If Nade(i).X <= (oLeft + oWidth + Lim) Then
        If Nade(i).Y >= oTop Then
            If Nade(i).Y <= (oTop + oHeight) Then
                NadeCollision = True
            End If
        End If
    End If
End If

End Function

Private Sub ClipStick(i As Integer, bLBoundSpeed As Boolean)

Const Lim As Integer = 50
Const ValIn = 30
Dim ClippedX As Boolean, ClippedY As Boolean
Dim XComp As Single, YComp As Single

ClippedY = (Stick(i).Y < Lim)
ClippedX = (Stick(i).X > StickGameWidth - Lim) Or (Stick(i).X < Lim)

If ClippedX Then 'Or ClippedY Then
    With Stick(i)
        XComp = .Speed * Sine(.Heading)
        YComp = .Speed * CoSine(.Heading)
        
        If Stick(i).X < Lim Then
            XComp = Abs(XComp)
        Else
            XComp = -Abs(XComp)
        End If
        
        SubStickiState i, STICK_LEFT
        SubStickiState i, STICK_RIGHT
        
    End With
End If

If ClippedY Then
    With Stick(i)
        XComp = .Speed * Sine(.Heading)
        YComp = .Speed * CoSine(.Heading)
        
        If Stick(i).Y < Lim Then
            YComp = -Abs(YComp)
            
            If i = 0 Then
                If Stick(i).WeaponType <> Chopper And StickInvul(i) = False Then
                    'Stick(i).Helth = Stick(i).Health - Stick(i).Speed / 15
                    DamageStick Stick(i).Speed / 15, i, i, False, False
                    
                    AddBloodExplosion Stick(i).X, Stick(i).Y
                    
                    If Stick(i).Health < 1 Then
                        Call Killed(i, i, kCeiling)
                    End If
                End If
            End If
            
        'Else
            'YComp = -Abs(YComp)
        End If
        
        SubStickiState i, STICK_LEFT
        SubStickiState i, STICK_RIGHT
        
    End With
End If


If ClippedX Or ClippedY Then
    Stick(i).Speed = Sqr(XComp * XComp + YComp * YComp)
    
    If YComp > 0 Then Stick(i).Heading = Atn(XComp / YComp)
    If YComp < 0 Then Stick(i).Heading = Atn(XComp / YComp) + Pi
End If

If bLBoundSpeed Then
    LBoundSpeed i
End If


With Stick(i)
    If .X < Lim Then
        .X = Lim * 2
    ElseIf .X > (StickGameWidth - Lim) Then
        .X = StickGameWidth - Lim * 2
    End If
    
    If .Y < 0 Then
        .Y = ValIn
    ElseIf .Y > (StickGameHeight - 500) Then
        .Y = StickGameHeight - 500 - ValIn
        .Speed = 0
        
        'touched the bottom, death to him/her
        If StickInvul(i) = False Then
            Call Killed(i, i, kFall)
        End If
        
    End If
End With

End Sub

Private Sub LBoundSpeed(i As Integer)

If Stick(i).bOnSurface Then
    
    If StickiHasState(i, STICK_CROUCH) Or StickiHasState(i, STICK_PRONE) Then
        
        If Stick(i).Speed <= 5 Then
            If StickHasMoveState(i) = False Then
                Stick(i).Speed = 0
            End If
        End If
    ElseIf Stick(i).Speed < IIf(Stick(i).WeaponType = Chopper, 3, 2) Then
        Stick(i).Speed = 0
    End If
    
'ElseIf Stick(i).Speed > Max_Speed Then
    'Stick(i).Speed = Max_Speed
    ' See limitspeed()
    
End If

End Sub

Private Sub ReverseYComp(Speed As Single, Heading As Single)

Dim XComp As Single
Dim YComp As Single
 
'Determine the components of the resultant vector
XComp = Speed * Sine(Heading)
YComp = -Speed * CoSine(Heading)


'Calculate the resultant direction, and adjust for atngent by adding Pi if necessary
If YComp > 0 Then
    Heading = Atn(XComp / YComp)
ElseIf YComp < 0 Then
    Heading = Atn(XComp / YComp) + Pi
End If

End Sub

Private Sub ReverseXComp(Speed As Single, Heading As Single)

Dim XComp As Single
Dim YComp As Single
 
'Determine the components of the resultant vector
XComp = -Speed * Sine(Heading)
YComp = Speed * CoSine(Heading)

'Calculate the resultant direction, and adjust for atngent by adding Pi if necessary
If YComp > 0 Then
    Heading = Atn(XComp / YComp)
ElseIf YComp < 0 Then
    Heading = Atn(XComp / YComp) + Pi
End If

End Sub

Private Function StickHasMoveState(i As Integer) As Boolean

If StickiHasState(i, STICK_LEFT) Then
    StickHasMoveState = True
ElseIf StickiHasState(i, STICK_RIGHT) Then
    StickHasMoveState = True
End If

End Function
'##############################################################################
'Smoke ########################################################################
'##############################################################################
'Private Sub AddLargeSmoke(X As Single, Y As Single, Heading As Single)
''Const MaxSize = 300, MinSize = 100
'Dim i As Integer ', Face As Single
'
'ReDim Preserve LargeSmoke(NumLargeSmokes)
'
'With LargeSmoke(NumLargeSmokes)
'    .CentreX = X
'    .CentreY = Y
'
'    .iDirection = 1
'
'    For i = 1 To 10
'
'        .SingleSmoke(i).DistanceFromMain = 10
'        .SingleSmoke(i).AngleFromMain = Rnd() * Pi2
'        .SingleSmoke(i).AspectDir = 1
'        .SingleSmoke(i).sAspect = 1
'        .SingleSmoke(i).DistanceFromMainInc = 0.5
'
''        .SingleSmoke(i).X = X '+ (i - 2) * Spacing
''        .SingleSmoke(i).Y = Y
''        .SingleSmoke(i).Speed = 2 + Rnd()
''        .SingleSmoke(i).Heading = Heading + piD6 * Sgn(PM_Rnd())
'    Next i
'
''    For i = 3 To 4
''        .SingleSmoke(i).X = X
''        .SingleSmoke(i).Y = Y
''        .SingleSmoke(i).Speed = 2
''        .SingleSmoke(i).Heading = pi3D2 + piD10 * (i - 3) * Sgn(PM_Rnd())
''    Next i
'
'
'
'
''    For i = 1 To 4
''        .X(i) = X + Spacing * (i - 2)
''        .Y(i) = Y
''    Next i
''    .pPoly(1).X = X
''    .pPoly(1).Y = Y - MaxSize
''
''    For i = 2 To 10
''
''        Face = Face + piD10
''
''        .pPoly(i).X = X + Rnd() * MinSize * sine(Face)
''        .pPoly(i).Y = Y + Rnd() * MinSize * cosine(Face)
''    Next i
'
'
'End With
'NumLargeSmokes = NumLargeSmokes + 1
'
'End Sub
'
'Private Sub RemoveLargeSmoke(Index As Integer)
'Dim i As Integer
'
'If NumLargeSmokes = 1 Then
'    Erase LargeSmoke
'    NumLargeSmokes = 0
'Else
'    For i = Index To NumLargeSmokes - 2
'        LargeSmoke(i) = LargeSmoke(i + 1)
'    Next i
'
'    'Resize the array
'    NumLargeSmokes = NumLargeSmokes - 1
'    ReDim Preserve LargeSmoke(NumLargeSmokes - 1)
'End If
'
'End Sub
'
'Private Sub ProcessAndDrawLargeSmokes()
'Dim i As Integer, j As Integer
'Const Size_Inc = 4, Size_Dec = 1, Space_Inc = 4
'Const Max_Size = 2500, Min_Size = 5
'
'picMain.FillStyle = vbFSSolid
'picMain.FillColor = SmokeFill
'
'Do While i < NumLargeSmokes
'
''    For j = 1 To 10
''        LargeSmoke(i).pPoly(j).X = LargeSmoke(i).pPoly(j).X - _
''            Sgn(LargeSmoke(i).X - LargeSmoke(i).pPoly(j).X) * Size_Inc * modStickGame.StickTimeFactor * _
''            LargeSmoke(i).iDirection
''
''
''        LargeSmoke(i).pPoly(j).Y = LargeSmoke(i).pPoly(j).Y - _
''            Sgn(LargeSmoke(i).Y - LargeSmoke(i).pPoly(j).Y) * Size_Inc * modStickGame.StickTimeFactor * _
''            LargeSmoke(i).iDirection
''
''    Next j
'
'    'DrawSmoke LargeSmoke(i).pPoly, MGrey
'    'modStickGame.sCircle LargeSmoke(i).X, LargeSmoke(i).Y, LargeSmoke(i).iSize * 3, MGrey
'
'    For j = 1 To 10
'
'
'        'modStickGame.StickMotion LargeSmoke(i).SingleSmoke(j).X, LargeSmoke(i).SingleSmoke(j).Y, _
'               LargeSmoke(i).SingleSmoke(j).Speed, LargeSmoke(i).SingleSmoke(j).Heading
'
'        'LargeSmoke(i).Y(j) = LargeSmoke(i).Y(j) - _
'            Sgn(LargeSmoke(i).CentreY - LargeSmoke(i).Y(j)) * Space_Inc * modStickGame.StickTimeFactor / 4
'
'        RotateLargeSmokePart i, j
'        DrawLargeSmokePart i, j
'
'    Next j
'
'
'    If LargeSmoke(i).iDirection = 1 Then
'        LargeSmoke(i).iSize = LargeSmoke(i).iSize + Size_Inc * modStickGame.StickTimeFactor '* GetTimeZoneAdjust(LargeSmoke(i).CentreX, LargeSmoke(i).CentreY)
'    Else
'        LargeSmoke(i).iSize = LargeSmoke(i).iSize - Size_Dec * modStickGame.StickTimeFactor '* GetTimeZoneAdjust(LargeSmoke(i).CentreX, LargeSmoke(i).CentreY)
'    End If
'
'
'    If LargeSmoke(i).iDirection = 1 Then
'        If LargeSmoke(i).iSize > Max_Size Then
'            LargeSmoke(i).iDirection = -1
'
'            For j = 1 To 10
'                LargeSmoke(i).SingleSmoke(j).DistanceFromMainInc = -0.5
'            Next j
'
'        End If
'
'    ElseIf LargeSmoke(i).iSize <= Min_Size Then
'        RemoveLargeSmoke i
'        i = i - 1
'
'    ElseIf LargeSmoke(i).iSize > Max_Size Then
'        'limit
'        LargeSmoke(i).iSize = Max_Size
'    End If
'
'
'
'    i = i + 1
'Loop
'
'picMain.FillStyle = vbFSTransparent
'
'End Sub
'
'Private Sub RotateLargeSmokePart(i As Integer, j As Integer)
'Const AngleInc = Pi / 250
'Dim Adj As Single
'
'Adj = GetTimeZoneAdjust(LargeSmoke(i).CentreX, LargeSmoke(i).CentreY)
'
'With LargeSmoke(i).SingleSmoke(j)
'    .DistanceFromMain = .DistanceFromMain + .DistanceFromMainInc * modStickGame.StickTimeFactor * Adj
'
'    .AngleFromMain = FixAngle(.AngleFromMain + AngleInc * modStickGame.StickTimeFactor * Adj)
'
'    .sAspect = .sAspect + 0.001 * .AspectDir * modStickGame.StickTimeFactor * Adj
'
'
'    If .sAspect > 1.2 Then
'        .AspectDir = -1
'    ElseIf .sAspect < 0.8 Then
'        .AspectDir = 1
'    End If
'
'End With
'
'End Sub
'
'Private Sub DrawLargeSmokePart(i As Integer, j As Integer) ', bFull As Boolean)
'Dim tX As Single, tY As Single
'
'With LargeSmoke(i).SingleSmoke(j)
'    tX = LargeSmoke(i).CentreX + .DistanceFromMain * Sine(.AngleFromMain)
'    tY = LargeSmoke(i).CentreY - .DistanceFromMain * CoSine(.AngleFromMain)
'
'
'    modStickGame.sCircleAspect tX, tY, LargeSmoke(i).iSize / 2, SmokeFill, .sAspect
'End With
'
''If bFull Then
''Else
''    modStickGame.sHatchCircle _
''        LargeSmoke(i).SingleSmoke(j).X, _
''        LargeSmoke(i).SingleSmoke(j).Y, _
''        MGrey, LargeSmoke(i).iSize / 25
''End If
'
'End Sub

'Private Sub DrawSmoke(X As Single, Y As Single, lFillCol As Long, iSize As Long)

'Dim pPts() As POINTAPI
'Dim j As Integer
'
'ReDim pPts(LBound(Pts) To UBound(Pts))
'
'For j = LBound(Pts) To UBound(Pts)
'    pPts(j).X = frmStickGame.ScaleX(Pts(j).X, vbTwips, vbPixels)
'    pPts(j).Y = frmStickGame.ScaleY(Pts(j).Y, vbTwips, vbPixels)
'Next j
'
'modGDI.DrawCrossPoly_NoOutline pPts, frmStickGame.picMain.hdc, lFillCol

'pPts = Pts
'modStickGame.sHatchPoly pPts, lFillCol
'Erase pPts
'End Sub

Private Sub AddSmokeNadeTrail(ByVal X As Single, ByVal Y As Single, _
    Optional bLong As Boolean = False, Optional bForce As Boolean = False) ', _
    ByVal Speed As Single, ByVal Heading As Single)

If modStickGame.cg_Smoke Or bForce Then
    AddSmokeGroup X, Y, 4, 3, PM_Rnd * piD4, bLong, True
    AddSmokeGroup X, Y, 3, 2, PM_Rnd * piD4, bLong, True
    'AddSmokeGroup X, Y, 3, 2, pm_rnd * piD4
End If

End Sub

Private Sub AddSmokeGroup(ByVal X As Single, ByVal Y As Single, ByVal HowMany As Integer, _
    ByVal Speed As Single, ByVal Heading As Single, Optional bLong As Boolean = False, _
    Optional bForce As Boolean = False)

Dim i As Integer
Const MaxSpacing = 75
Dim rX As Single, rY As Single

If modStickGame.cg_Smoke Or bForce Then
    For i = 1 To HowMany
        rX = X + (Rnd() - 0.5) * MaxSpacing
        rY = Y + (Rnd() - 0.5) * MaxSpacing
        
        AddSmoke rX, rY, Speed, Heading, bLong, bForce
    Next i
End If

End Sub

Private Sub AddSmoke(X As Single, Y As Single, Speed As Single, Heading As Single, bLongTime As Boolean, _
    Optional bForce As Boolean = False)

If modStickGame.cg_Smoke Or bForce Then
    ReDim Preserve Smoke(NumSmoke)
    
    Smoke(NumSmoke).X = X
    Smoke(NumSmoke).Y = Y
    Smoke(NumSmoke).Direction = 1
    Smoke(NumSmoke).Size = 10 '0.4
    
    Smoke(NumSmoke).Speed = Speed
    Smoke(NumSmoke).Heading = Heading
    
    Smoke(NumSmoke).bLongTime = bLongTime
    
    
    NumSmoke = NumSmoke + 1
End If

End Sub

Private Sub RemoveSmoke(ByVal Index As Integer)

Dim i As Integer

If NumSmoke = 1 Then
    Erase Smoke
    NumSmoke = 0
Else
    For i = Index To NumSmoke - 2
'        Smoke(i).X = Smoke(i + 1).X
'        Smoke(i).Y = Smoke(i + 1).Y
'        Smoke(i).Size = Smoke(i + 1).Size
'        Smoke(i).Direction = Smoke(i + 1).Direction
'
'
'        Smoke(i).Heading = Smoke(i + 1).Heading
'        Smoke(i).Speed = Smoke(i + 1).Speed
        Smoke(i) = Smoke(i + 1)
    Next i
    
    ReDim Preserve Smoke(NumSmoke - 2)
    NumSmoke = NumSmoke - 1
End If

End Sub

Private Sub ProcessCasings()
Dim i As Integer

For i = 0 To NumCasings - 1
    
    If Casing(i).Speed Then
        MotionStickObject Casing(i).X, Casing(i).Y, Casing(i).Speed, Casing(i).Heading
        
        ClipCasing i
        
        ApplyGravityVector Casing(i).LastGravity, _
            GetTimeZoneAdjust(Casing(i).X, Casing(i).Y), Casing(i).Speed, _
            Casing(i).Heading, Casing(i).X, Casing(i).Y
        
        
    End If
    
Next i


End Sub

Private Sub ClipCasing(i As Integer)
Dim j As Integer

For j = 0 To modStickGame.ubdPlatforms
    If CasingCollision(i, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        
        If Casing(i).Heading > 0 Then
            Casing(i).Y = Platform(j).Top - 1
        Else
            Casing(i).Y = Platform(j).Top + Platform(j).height + 1
        End If
        
        ReverseYComp Casing(i).Speed, Casing(i).Heading
        Casing(i).Speed = Casing(i).Speed * Casing_Bounce_Reduction
        
        CheckCasingSpeed i
        
        Exit Sub
    End If
Next j



If Casing(i).X < Lim Then
    
    Casing(i).X = Lim
    
    ReverseXComp Casing(i).Speed, Casing(i).Heading
    Casing(i).Speed = Casing(i).Speed * Casing_Bounce_Reduction
    
ElseIf Casing(i).X > StickGameWidth - Lim Then
    
    Casing(i).X = StickGameWidth - Lim
    
    ReverseXComp Casing(i).Speed, Casing(i).Heading
    Casing(i).Speed = Casing(i).Speed * Casing_Bounce_Reduction
    
End If
If Casing(i).Y < 1 Then
    Casing(i).Y = 1
    ReverseYComp Casing(i).Speed, Casing(i).Heading
ElseIf Casing(i).Y > StickGameHeight Then
    Casing(i).Speed = 0
    Casing(i).Y = Platform(0).Top
End If

End Sub

Private Sub CheckCasingSpeed(i As Integer)

If Casing(i).Speed < 30 Then Casing(i).Speed = 0

End Sub


Private Function CasingCollision(ByVal i As Integer, _
    oLeft As Single, oTop As Single, oWidth As Single, oHeight As Single) As Boolean

If Casing(i).X >= oLeft - Lim Then
    If Casing(i).X <= (oLeft + oWidth + Lim) Then
        If Casing(i).Y >= oTop Then
            If Casing(i).Y <= (oTop + oHeight) Then
                CasingCollision = True
            End If
        End If
    End If
End If

End Function

Private Sub ProcessNades()
Dim i As Integer, j As Integer
Dim bRemoveIt As Boolean, bWall As Boolean
Dim bPlayBounce As Boolean
Dim Adj As Single

Do While i < NumNades
    
    bWall = False
    bRemoveIt = False
    
    '###############################################################
    Adj = GetTimeZoneAdjust(Nade(i).X, Nade(i).Y)
    
    ApplyGravityVector Nade(i).LastGravity, Adj, Nade(i).Speed, Nade(i).Heading, Nade(i).X, Nade(i).Y, _
        IIf(Nade(i).IsRPG, Gravity_Strength / 2, Gravity_Strength)
    
    
    MotionStickObject Nade(i).X, Nade(i).Y, Nade(i).Speed, Nade(i).Heading
    '###############################################################
    
    
    
    ClipNade i, bWall, bRemoveIt, bPlayBounce
    'yes, keep the IF [below] there, since bRemoveIt is passed above byref
    If bRemoveIt = False Then
        For j = 0 To NumSticksM1
            If StickInGame(j) Then
                If Nade(i).OwnerID <> Stick(j).ID Then
                    If NadeInStick(i, j) Then
                        'ReverseXComp Nade(i).Speed, Nade(i).Heading
                        bRemoveIt = True
                        'bWall=false
                        
                        'or explode the nade?
                        Exit For
                    End If
                End If
            End If
        Next j
        
        If bRemoveIt = False Then
            For j = 0 To NumBarrels - 1
                If NadeNearBarrel(i, j) Then
                    
                    
                    'remove the grenade, regardless
                    bRemoveIt = True
                    
                    If Nade(i).iType = nFrag Then
                        Barrel(j).LastTouchID = Nade(i).OwnerID
                        ExplodeBarrel j, True
                        RemoveBarrel j
                    End If
                    
                    Exit For
                End If
            Next j
        
        
            If bRemoveIt = False Then
                If Nade(i).IsRPG = False Then
                    If Nade(i).Start_Time + Nade_Time / Adj < GetTickCount() Then
                        bRemoveIt = True
                        
                        bWall = (NadeInPlatform(i) > -1)
                        If bWall = False Then
                            bWall = (NadeInTBox(i) > -1)
                        End If
                        'main bit for WallMarks above ^^
                        
                    End If
                End If
            End If
            
        End If
    End If
    
    
    '#################################
    
    If bPlayBounce Then
        If Nade(i).Speed > 85 Then
            If PointHearableOnSticksScreen(Nade(i).X, Nade(i).Y, 0) Then
                modAudio.PlayNadeBounce GetRelPan(Nade(i).X)
            End If
        End If
        bPlayBounce = False
    End If
    If bRemoveIt Then
        ExplodeNade i
        
        If Nade(i).iType = nFrag Then
            ExplodeAll Nade(i).X, Nade(i).Y, Nade(i).OwnerID, -1, -1
        End If
        
        RemoveNade i, bWall
        i = i - 1
    End If
    
    i = i + 1
Loop


End Sub

'Private Function Nade_Vulnerable(i As Integer, Adj As Single) As Boolean
'Nade_Vulnerable = (Nade(i).Start_Time + Nade_Bullet_Invul_Time / Adj < GetTickCount())
'End Function

Private Sub ClipNade(i As Integer, ByRef bWall As Boolean, ByRef bRemoveIt As Boolean, ByRef bPlayBounce As Boolean)
Dim j As Integer
Const Nade_Shift As Single = 3 ' Nade_Radius * 2.5

j = NadeInPlatform(i)

If j > -1 Then
    
    If Nade(i).IsRPG Then
        bRemoveIt = True
        bWall = True
    ElseIf Nade(i).iType <> nFrag Then
        bRemoveIt = True
        bWall = True
    Else
        Nade(i).Heading = FixAngle(Nade(i).Heading)
        If Nade(i).Heading > piD2 And Nade(i).Heading < pi3D2 Or j = 0 Then
            '                                                    ^don't allow it to go to the bottom of platform(0)
            
            'was heading down
            Nade(i).Y = Platform(j).Top
        Else
            'was heading up
            Nade(i).Y = Platform(j).Top + Platform(j).height
        End If
        
        Nade(i).Speed = Nade(i).Speed / Nade_Bounce_Reduction
        ReverseYComp Nade(i).Speed, Nade(i).Heading
        
        bPlayBounce = True
    End If
    
    
ElseIf NadeInBox(i) Then
    bRemoveIt = True
    
ElseIf NadeOnEdge(i) Then
    
    If Nade(i).IsRPG Then
        bRemoveIt = True
        bWall = True
    ElseIf Nade(i).iType <> nFrag Then 'Nade(i).iType = nSmoke
        bRemoveIt = True
        bWall = True
    Else
        Nade(i).Speed = Nade(i).Speed / Nade_Bounce_Reduction
        
        If Nade(i).X > (StickGameWidth - Lim - 5) Then
            Nade(i).X = StickGameWidth - Lim - 10
            
            If Nade(i).Speed > 0 Or Nade(i).Heading < Pi Then
                ReverseXComp Nade(i).Speed, Nade(i).Heading
                bPlayBounce = True
            End If
            
        ElseIf Nade(i).X < Lim Then
            Nade(i).X = Lim + 5
            
            If Nade(i).Speed < 0 Or Nade(i).Heading > Pi Then
                ReverseXComp Nade(i).Speed, Nade(i).Heading
                bPlayBounce = True
            End If
        Else
            ReverseXComp Nade(i).Speed, Nade(i).Heading
            bPlayBounce = True
        End If
    End If
    
ElseIf Nade(i).Y < Lim Then
    
    If Nade(i).IsRPG Or Nade(i).iType = nTime Then
        bRemoveIt = True
        bWall = True
    Else
        If Nade(i).Speed > 0 And (Nade(i).Heading < piD2 Or Nade(i).Heading > pi3D2) Then
            ReverseYComp Nade(i).Speed, Nade(i).Heading
        End If
        
        Nade(i).Y = Lim
        bPlayBounce = True
    End If
    
    
Else
    
    j = NadeInTBox(i) 'j = tBox that nade is in
    
    If j > -1 Then
        If Nade(i).IsRPG Then
            bRemoveIt = True
            'bWall=false
        ElseIf Nade(i).iType <> nFrag Then 'Nade(i).iType = nSmoke
            bRemoveIt = True
        Else
            Nade(i).Heading = FixAngle(Nade(i).Heading)
            
            
'            If AnglesRoughlyEqual(Nade(i).Heading, piD2) Then
'                'was heading right
'                Nade(i).X = tBox(j).Left + tBox(j).width
'                ReverseXComp Nade(i).Speed, Nade(i).Heading
'            ElseIf AnglesRoughlyEqual(Nade(i).Heading, pi3D2) Then
'                'was heading left
'                Nade(i).X = tBox(j).Left
'                ReverseXComp Nade(i).Speed, Nade(i).Heading
'            End If
'            If AnglesRoughlyEqual(Nade(i).Heading, Pi) Then
'                'was heading down
'                Nade(i).Y = tBox(j).Top
'                ReverseYComp Nade(i).Speed, Nade(i).Heading
'            ElseIf AnglesRoughlyEqual(Nade(i).Heading, Pi) Then
'                'was heading up
'                Nade(i).Y = tBox(j).Top + tBox(j).height
'                ReverseYComp Nade(i).Speed, Nade(i).Heading
'            End If
            
            
            Do
                Nade(i).X = Nade(i).X - Nade_Shift * Sine(Nade(i).Heading)
                Nade(i).Y = Nade(i).Y + Nade_Shift * CoSine(Nade(i).Heading)
            Loop Until NadeInTBox(i) = -1
            
            
            Nade(i).Speed = Nade(i).Speed / Nade_Bounce_Reduction
            Nade(i).Heading = Nade(i).Heading - Pi
            
            bPlayBounce = True
        End If
    End If
    
End If

End Sub

'                         V byval, so it doesn't lock any arrays
Private Sub ExplodeAll(ByVal X As Single, ByVal Y As Single, ByVal ExplosionOwnerID As Integer, _
    iMineToSkip As Integer, iBarrelToSkip As Integer)

Dim i As Integer

Static bDoing As Boolean
If bDoing Then Exit Sub
bDoing = True


'check barrels

'i=0
Do While i < NumBarrels
    If i <> iBarrelToSkip Then
        If ExplosionNearPoint(X, Y, Barrel(i).X, Barrel(i).Y) Then
            Barrel(i).LastTouchID = ExplosionOwnerID
            
            ExplodeBarrel i, True
            
            If iBarrelToSkip > i Then
                'reduce index by 1
                iBarrelToSkip = iBarrelToSkip - 1
            End If
            
            RemoveBarrel i
            i = i - 1
        End If
    End If
    
    i = i + 1
Loop



'check mines
i = 0
Do While i < NumMines
    If i <> iMineToSkip Then
        If ExplosionNearPoint(X, Y, Mine(i).X, Mine(i).Y) Then
            ExplodeMine i, True
            
            If iMineToSkip > i Then
                'reduce index by 1
                iMineToSkip = iMineToSkip - 1
            End If
            
            RemoveMine i
            i = i - 1
        End If
    End If
    
    i = i + 1
Loop

bDoing = False
End Sub

Private Function ExplosionNearPoint(eX As Single, eY As Single, pX As Single, pY As Single) As Boolean
Const Explosion_Limit = BodyLen * 10

ExplosionNearPoint = (GetDist(eX, eY, pX, pY) <= Explosion_Limit)

End Function

Private Sub ExplodeNade(ByVal i As Integer)
Dim j As Integer

With Nade(i)
    AddSmokeNadeTrail .X, .Y, True, True
    AddMoreSparks .X, .Y, 24
    
    If PointHearableOnSticksScreen(.X, .Y, 0) Then
        modAudio.PlayNadeExplosion GetRelPan(.X)
    Else
        modAudio.PlayBackGroundNade GetRelPan(.X)
    End If
    
End With

Select Case Nade(i).iType
    Case nFrag
        ExplodeFrag i
    Case nFlash
        ExplodeFlash i
    Case nTime
        ExplodeTimeNade i
    Case nGravity
        ExplodeGravityNade i
    Case Else
        ExplodeEMPNade i
End Select

End Sub

Private Sub ExplodeTimeNade(i As Integer)

AddTimeZone Nade(i).X, Nade(i).Y

AddSparks Nade(i).X, Nade(i).Y, Nade(i).Heading - Pi
AddExplosion Nade(i).X, Nade(i).Y, 500

End Sub

Private Sub ExplodeGravityNade(i As Integer)

AddGravityZone Nade(i).X, Nade(i).Y

AddSparks Nade(i).X, Nade(i).Y, Nade(i).Heading - Pi
AddExplosion Nade(i).X, Nade(i).Y, 500

End Sub

Private Sub ExplodeEMPNade(iNade As Integer)
Const EMPNadeRadius As Single = 5000, WaveRadius As Single = 250
Dim i As Integer

With Nade(iNade)
    modStickGame.sCircle .X, .Y, EMPNadeRadius, vbYellow
    For i = 0 To 10
        AddShieldWave .X + PM_Rnd() * WaveRadius, .Y + PM_Rnd() * WaveRadius, Rnd() * Pi2
    Next i
    AddMoreSparks .X, .Y, 20
    AddExplosion .X, .Y, 500
    
    
    'Shields
    For i = 0 To NumSticksM1
        If StickInGame(i) Then
            If GetDist(Stick(i).X, Stick(i).Y, .X, .Y) < EMPNadeRadius Then
                RemoveSticksShield i
            End If
        End If
    Next i
    
    'Mines
    i = 0
    While i < NumMines
        If GetDist(Mine(i).X, Mine(i).Y, .X, .Y) < EMPNadeRadius Then
            ExplodeMine i, True
            RemoveMine i
        End If
        i = i + 1
    Wend
End With

End Sub

Private Function GetRelPan(X As Single) As Long

'Relative to Stick(0)
'GetRelPan = RightPan * CLng(X - Stick(0).X) / StickGameWidth

'Relative to the camera
On Error Resume Next
GetRelPan = RightPan * CLng(X - cg_sCamera.X - StickCentreX) / StickGameWidth
'                           X - (cg_sCamera.X + StickCentreX)
'                                               ^ this is to centre the point, because cg_sCamera.X represents the left 'wall'


End Function

Private Sub ExplodeFrag(i As Integer)
Dim j As Integer

Dim Dist As Single
Dim OwnerIndex As Integer
Dim MaxDist As Single
Dim ExplosionForceDist As Single, SmokeAngle As Single, AngleToStick As Single

Const Nade_Explode_RadiusX2 = Nade_Explode_Radius * 2
Const ChopperLenX1p2 = ChopperLen * 1.2
Const Nade_Multiple_X = NadeMultiple * 12000
Const ShieldWaveDispersion As Single = 600


AddExplosion Nade(i).X, Nade(i).Y, 500
If modStickGame.cg_Smoke Then
    For j = 1 To 10
        AddSmokeGroup Nade(i).X, Nade(i).Y, 5, 100 * Rnd(), Pi2 * Rnd(), , True
    Next j
End If

For j = 1 To 3 + Rnd() * 5
    AddNadeTrail_Simple Nade(i).X, Nade(i).Y
Next j

SmokeAngle = Rnd() * Pi2
AddSmokeGroup Nade(i).X, Nade(i).Y, 10, 150, SmokeAngle, , True
AddSmokeGroup Nade(i).X, Nade(i).Y, 10, 150, SmokeAngle - Pi, , True
'AddSparks Nade(i).X, Nade(i).Y, SmokeAngle
'AddSparks Nade(i).X, Nade(i).Y, SmokeAngle - Pi

OwnerIndex = FindStick(Nade(i).OwnerID)

For j = 0 To NumSticksM1
    'apply damage
    
    If StickInGame(j) Then
        
        Dist = GetDist(Stick(j).X, Stick(j).Y, Nade(i).X, Nade(i).Y)
        
        If Stick(j).WeaponType = Chopper Then
            MaxDist = ChopperLen
            ExplosionForceDist = ChopperLenX1p2
        Else
            MaxDist = Nade_Explode_Radius
            ExplosionForceDist = Nade_Explode_RadiusX2
        End If
        
        AngleToStick = FindAngle(Nade(i).X, Nade(i).Y, Stick(j).X, Stick(j).Y)
        
        If Dist < ExplosionForceDist Then
            If Stick(j).WeaponType <> Chopper Then
                
                If Stick(j).Shield = 0 Then
                    If Nade(i).Y > Stick(j).Y Then
                        If Stick(j).bOnSurface Then
                            If Not StickiHasState(j, STICK_PRONE) Then
                                AddVectors Stick(j).Speed, Stick(j).Heading, _
                                    Nade_Multiple_X / (Dist + 1), AngleToStick, _
                                    Stick(j).Speed, Stick(j).Heading
                            End If
                        End If
                    End If
                End If
                
            End If
        End If
        
        
        If Dist < MaxDist Then
            If Stick(j).Shield Then
                AngleToStick = AngleToStick - Pi
                AddShieldWave Stick(j).X, Stick(j).Y, AngleToStick
                AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
            Else
                Stick(j).bOnSurface = False
                Stick(j).Y = Stick(j).Y - 100
            End If
            
            
            'Exit For
            If j = 0 Or Stick(j).IsBot Then
                
                If OwnerIndex <> -1 Then '                                        V so it can damage me
                    If IsAlly(Stick(j).Team, Stick(OwnerIndex).Team) = False Or OwnerIndex = j Then
                        'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
                        If StickInvul(j) = False Then
                            
                            On Error Resume Next
                            'Stick(j).Helth = Stick(j).Health - 100000 / Dist
                            
                            If Stick(j).WeaponType = Chopper Then
                                If Nade(i).IsRPG Then
                                    
                                    'needs to be big (30000) if it hits the tail
                                    'If Dist > CLD2 Then
                                        'DamageStick Chopper_Damage_Reduction * 100000 / Dist, j 'tail rotor
                                    'Else
                                        'DamageStick Chopper_Damage_Reduction * 5000 / Dist, j 'normal
                                    'End If
                                    
                                    
                                    'fixed damage of 75 (2 RPGs to kill)
                                    DamageStick Chopper_Damage_Reduction * 75, j, OwnerIndex
                                    
                                Else
                                    DamageStick Chopper_Damage_Reduction * 30, j, OwnerIndex 'nade = 30 damage
                                End If
                            Else
                                DamageStick 100000 / Dist, j, OwnerIndex 'bullet
                            End If
                            
                            If Err.Number <> 0 Then 'div zero error
                                Stick(j).Health = 0
                                Err.Clear
                            End If
                            
                            If Stick(j).Health < 1 Then
                                Killed j, OwnerIndex, _
                                    IIf(Nade(i).IsRPG, kRPG, _
                                    IIf(Nade(i).bIsMartyrdomNade, kMartyrdom, kNade))
                                
                            End If
                            
                        End If 'spawn invul endif
                    End If 'ally endif
                End If 'owner index endif
            End If 'myid endif
        ElseIf Stick(j).WeaponType = Chopper Then
            If Dist < 2870 Then
                If j = 0 Or Stick(j).IsBot Then
                    'tail rotor
                    DamageStick Chopper_Damage_Reduction * 250000 / Dist, j, OwnerIndex
                    
                    If Stick(j).Health < 1 Then
                        Call Killed(j, FindStick(Nade(i).OwnerID), IIf(Nade(i).IsRPG, kRPG, kNade))
                    End If
                End If
            End If
        End If 'dist endif
    End If 'stickingame endif
Next j

End Sub

Private Sub ExplodeFlash(i As Integer)
Dim j As Integer
Const SparkLim = 1000, Angle_Redux = 5
Dim Smoke_Speed As Single

AddMoreSparks Nade(i).X + PM_Rnd() * SparkLim, Nade(i).Y + PM_Rnd * SparkLim, 80


'If Nade(i).OwnerID <> Stick(0).ID Then 'you are prepared for your own flash
    If PointVisibleOnSticksScreen(Nade(i).X, Nade(i).Y, 0) Then
        If StickInvul(0) = False And StickInGame(0) Then
            BangFlash i
        End If
    End If
'End If
AddExplosion Nade(i).X, Nade(i).Y, 500


For j = 0 To NumSticks - 1
    'If Stick(j).IsBot Then
    'If Stick(j).WeaponTyp <> Chopper Then
    If Stick(j).ID <> Nade(i).OwnerID Then
        If StickInGame(j) Then
            If StickInvul(j) = False Then
                If PointVisibleOnSticksScreen(Nade(i).X, Nade(i).Y, j) Then
                    Stick(j).LastFlashBang = GetTickCount()
                    Stick(j).bFlashed = True
                End If
            End If
        End If
    End If
Next j

For j = 0 To 3
    Smoke_Speed = 120 + 20 * Rnd()
    AddSmokeGroup Nade(i).X, Nade(i).Y, 3, Smoke_Speed, piD2 + PM_Rnd() / Angle_Redux, True, True
    AddSmokeGroup Nade(i).X, Nade(i).Y, 3, Smoke_Speed, pi3D2 + PM_Rnd() / Angle_Redux, True, True
Next j

AddSmokeGroup Nade(i).X, Nade(i).Y, 3, Smoke_Speed, 0, True, True


End Sub

'Private Sub ExplodeSmoke(i As Integer)
'
'AddSparks Nade(i).X, Nade(i).Y, Nade(i).Heading - Pi
'AddExplosion Nade(i).X, Nade(i).Y, 500, 1, 0, 0
'
'AddLargeSmoke Nade(i).X, Nade(i).Y, Nade(i).Heading
'
'End Sub

Private Function PointHearableOnSticksScreen(X As Single, Y As Single, i As Integer) As Boolean
PointHearableOnSticksScreen = pPointOnSticksScreen(X, Y, i, 15000)
End Function
Private Function PointVisibleOnSticksScreen(X As Single, Y As Single, i As Integer) As Boolean
PointVisibleOnSticksScreen = pPointOnSticksScreen(X, Y, i, 6500)
End Function
Private Function pPointOnSticksScreen(X As Single, Y As Single, i As Integer, XLimit As Single) As Boolean
Const def_YLimit = 5500
Dim YLimit As Single
Dim sX As Single, sY As Single

If i = 0 Then
    If StickInGame(0) = False Then
        sX = modStickGame.cg_sCamera.X + StickCentreX 'me.width/2
        sY = modStickGame.cg_sCamera.Y + StickCentreY 'me.height/2
        YLimit = StickGameHeight
    Else
        sX = Stick(i).X: sY = Stick(i).Y
        YLimit = def_YLimit
    End If
Else
    sX = Stick(i).X: sY = Stick(i).Y
    YLimit = def_YLimit
End If


If Abs(sX - X) < XLimit Then
    If Y > (sY - YLimit) Then
        pPointOnSticksScreen = (Y < (sY + YLimit))
    End If
End If

End Function

Private Function NadeOnEdge(iNade As Integer) As Boolean

If Nade(iNade).X < Lim Then
    NadeOnEdge = True
ElseIf Nade(iNade).X > StickGameWidth - Lim Then
    NadeOnEdge = True
End If

End Function

Private Sub DrawNades()
Dim i As Integer
Dim tY As Single, tX As Single
Dim TimeLeft As Single
Const Line_Len As Single = 1000, Line_LenD2 As Single = Line_Len \ 2, _
    Nade_Warn_Radius = 1000, Nade_Warn_Radius_Vul = Nade_Warn_Radius + 100

'picMain.ForeColor = vbBlack
For i = 0 To NumNades - 1
    
    
    If Nade(i).IsRPG Then
        
        picMain.DrawWidth = 1
        picMain.FillStyle = vbFSTransparent
        
        DrawRocket Nade(i).X, Nade(i).Y, Nade(i).Heading ', Nade(i).Colour
        'picMain.DrawWidth = 2
        
        If Nade(i).LastSmoke + RPG_Smoke_Delay / GetTimeZoneAdjust(Nade(i).X, Nade(i).Y) < GetTickCount() Then
            
            tX = Nade(i).X + GunLen * Sine(Nade(i).Heading - Pi)
            tY = Nade(i).Y - GunLen * CoSine(Nade(i).Heading - Pi)
            
            AddSmokeGroup tX, tY, 3, 0, 0, , True
            'AddSmokeNadeTrail tX, tY
            
'            If modStickGame.cg_RPGFlame Then
            AddExplosion tX, tY, 100
'            End If
            
            Nade(i).LastSmoke = GetTickCount()
            
        End If
        
    Else
        picMain.DrawWidth = 2
        picMain.FillStyle = vbSolid
        
        DrawNade Nade(i).X, Nade(i).Y, Nade(i).colour, Nade(i).iType
        
        
        If modStickGame.sv_Draw_Nade_Time Then
            TimeLeft = (GetTickCount() - Nade(i).Start_Time) * GetTimeZoneAdjust(Nade(i).X, Nade(i).Y)
            
            
            tY = Nade(i).Y - 650
            tX = Nade(i).X - Line_LenD2
            
            picMain.ForeColor = vbBlue
            modStickGame.sLine tX, tY, tX + Line_Len, tY
            picMain.ForeColor = vbRed
            modStickGame.sLine tX, tY, tX + Line_Len * TimeLeft / Nade_Time, tY
            
            '                                       V not /Adj, because it's display, and stuff
            'If Nade(i).LastNadeTrail + NadeTrail_Smoke_Delay < GetTickCount() Then
                'AddNadeTrail i

                'Nade(i).LastNadeTrail = GetTickCount()
            'End If
        End If
        
        'old method
'        tX = Nade(i).X - Nade_Time / 4
'        tY = Nade(i).Y - 650
'
'        modStickGame.sLine tX, tY, tX + Nade_Time / 2, tY, vbBlue
'        modStickGame.sLine tX, tY, tX + Nade_Time / 2 - TimeLeft / 2, tY, vbRed
        
        
    End If
    
Next i

picMain.FillStyle = vbFSTransparent

If Stick(0).Perk = pBombSquad Then
    For i = 0 To NumNades - 1
        
        With Nade(i)
            modStickGame.sCircle .X, .Y, Nade_Warn_Radius, .colour
            
'            If .iType = nFrag Then
'                If Nade_Vulnerable(i, GetTimeZoneAdjust(.X, .Y)) Then
'                    modStickGame.sCircle .X, .Y, Nade_Warn_Radius_Vul, vbRed
'                End If
'            End If
        End With
        
    Next i
End If

End Sub

Private Sub DrawNade(X As Single, Y As Single, Col As Long, iType As eNadeTypes)
Const K = 5, m = 20, Mxk = m * 3

If iType = nFrag Then
    picMain.FillColor = Col
    modStickGame.sCircle X, Y, Nade_Radius, vbBlack
Else
    modStickGame.sBoxFilled X - m, Y - Mxk, X + m, Y + Mxk, Col
    'picMain.DrawWidth = 1
    modStickGame.sBox X - m, Y - Mxk, X + m, Y + Mxk, vbBlack
End If

End Sub

Private Sub DrawRocket(pX As Single, pY As Single, pHeading As Single) ', pCol As Long)

Const RPG_RearLen = 1.5
Dim pt(1 To 7) As PointAPI

pt(1).X = pX
pt(1).Y = pY

pt(2).X = pt(1).X + GunLen / 1.5 * Sine(pHeading - pi8D9)
pt(2).Y = pt(1).Y - GunLen / 1.5 * CoSine(pHeading - pi8D9)

pt(3).X = pt(2).X + GunLen / 2.5 * Sine(pHeading + pi8D9)
pt(3).Y = pt(2).Y - GunLen / 2.5 * CoSine(pHeading + pi8D9)

pt(4).X = pt(3).X + GunLen / 3 * Sine(pHeading - Pi)
pt(4).Y = pt(3).Y - GunLen / 3 * CoSine(pHeading - Pi)


pt(7).X = pt(1).X + GunLen / 1.5 * Sine(pHeading + pi8D9)
pt(7).Y = pt(1).Y - GunLen / 1.5 * CoSine(pHeading + pi8D9)

pt(6).X = pt(7).X + GunLen / 2.5 * Sine(pHeading - pi8D9)
pt(6).Y = pt(7).Y - GunLen / 2.5 * CoSine(pHeading - pi8D9)

pt(5).X = pt(6).X + GunLen / 3 * Sine(pHeading - Pi)
pt(5).Y = pt(6).Y - GunLen / 3 * CoSine(pHeading - Pi)

picMain.ForeColor = vbBlack
picMain.FillStyle = vbFSTransparent
'picMain.DrawWidth = 2
modStickGame.sPoly pt, -1 'pCol


'Dim x(1 To 7) As Single, y(1 To 7) As Single
'x(1) = pX
'y(1) = pY
'
'x(2) = x(1) + GunLen / 1.5 * sine(pHeading - pi8D9)
'y(2) = y(1) - GunLen / 1.5 * cosine(pHeading - pi8D9)
'
'x(3) = x(2) + GunLen / 2.5 * sine(pHeading + pi8D9)
'y(3) = y(2) - GunLen / 2.5 * cosine(pHeading + pi8D9)
'
'x(4) = x(3) + GunLen / 3 * sine(pHeading - pi)
'y(4) = y(3) - GunLen / 3 * cosine(pHeading - pi)
'
'
'x(7) = x(1) + GunLen / 1.5 * sine(pHeading + pi8D9)
'y(7) = y(1) - GunLen / 1.5 * cosine(pHeading + pi8D9)
'
'x(6) = x(7) + GunLen / 2.5 * sine(pHeading - pi8D9)
'y(6) = y(7) - GunLen / 2.5 * cosine(pHeading - pi8D9)
'
'x(5) = x(6) + GunLen / 3 * sine(pHeading - pi)
'y(5) = y(6) - GunLen / 3 * cosine(pHeading - pi)
'
'
'picMain.DrawWidth = 1
'picMain.ForeColor = vbBlack
'
'modStickGame.sLine x(1), y(1), x(2), y(2)
'modStickGame.sLine x(3), y(3), x(2), y(2)
'modStickGame.sLine x(3), y(3), x(4), y(4)
'
'modStickGame.sLine x(5), y(5), x(6), y(6)
'modStickGame.sLine x(7), y(7), x(6), y(6)
'modStickGame.sLine x(7), y(7), x(1), y(1)
'
'modStickGame.sLine x(3), y(3), x(6), y(6)


End Sub

Private Sub ProcessMines()
Dim i As Integer, j As Integer, iOwner As Integer
Dim RemoveIt As Boolean

Do While i < NumMines
    
    RemoveIt = False
    
    iOwner = FindStick(Mine(i).OwnerID)
    
    'PrintStickText "Near Mine " & i + 1 & ": " & MineNearStick(i, 0), Stick(0).X + 1000, Stick(0).Y + i * 250, vbBlack
    
    If MineInBullet(i) Then
        RemoveIt = True
    Else
        For j = 0 To NumSticksM1
            If StickInGame(j) And StickInvul(j) = False Then
                If Mine(i).OwnerID <> Stick(j).ID Then
                    If iOwner <> -1 Then
                        If IsAlly(Stick(j).Team, Stick(iOwner).Team) = False Then
                            If MineNearStick(i, j) Then
                                RemoveIt = True
                            End If
                        End If
                    End If
                End If
            End If
        Next j
    End If
    
    
    If RemoveIt Then
        ExplodeMine i, True
        RemoveMine i
        i = i - 1
    ElseIf Mine(i).bOnSurface = False Then
            
            If Mine(i).X < Lim Then
                If Mine(i).Heading > Pi Then
                    ReverseXComp Mine(i).Speed, Mine(i).Heading
                    Mine(i).Speed = Mine(i).Speed / 2
                End If
                Mine(i).X = Lim
            ElseIf Mine(i).X > (StickGameWidth - Lim) Then
                If Mine(i).Heading < Pi Then
                    ReverseXComp Mine(i).Speed, Mine(i).Heading
                    Mine(i).Speed = Mine(i).Speed / 2
                End If
                Mine(i).X = StickGameWidth - Lim
            End If
            If Mine(i).Y < 1 Then
                If Mine(i).Heading > pi3D2 Or Mine(i).Heading < piD2 Then
                    ReverseYComp Mine(i).Speed, Mine(i).Heading
                    Mine(i).Speed = Mine(i).Speed / 2
                End If
                Mine(i).Y = 1
            ElseIf Mine(i).Y > (StickGameWidth - 1) Then
                If Mine(i).Heading < pi3D2 And Mine(i).Heading > piD2 Then
                    ReverseYComp Mine(i).Speed, Mine(i).Heading
                    Mine(i).Speed = Mine(i).Speed / 2
                End If
                Mine(i).Y = StickGameWidth - 1
            End If
            
            
            ApplyGravityVector Mine(i).LastGravity, GetTimeZoneAdjust(Mine(i).X, Mine(i).Y), _
                Mine(i).Speed, Mine(i).Heading, Mine(i).X, Mine(i).Y
            
'            If Mine(i).LastGravity + Gravity_Delay / GetTimeZoneAdjust(Mine(i).X, Mine(i).Y) < GetTickCount() Then
'                AddVectors Mine(i).Speed, Mine(i).Heading, Gravity_Strength, Gravity_Direction, _
'                    Mine(i).Speed, Mine(i).Heading
'
'                Mine(i).LastGravity = GetTickCount()
'            End If
            
            
            MotionStickObject Mine(i).X, Mine(i).Y, Mine(i).Speed, Mine(i).Heading
            For j = 0 To ubdPlatforms
                MineOnSurface i, j
            Next j
            
            
    End If
    
    i = i + 1
Loop


End Sub

Private Sub MineOnSurface(i As Integer, iPlatform As Integer) 'As Boolean

'Dim rcMine As RECT
'
'With rcMine
'    .Left = Mine(i).X
'    .Right = .Left + 1
'    .Top = Mine(i).Y
'    .Bottom = .Top + 1
'End With

If RectCollision(PointToRect(Mine(i).X, Mine(i).Y), PlatformToRect(Platform(iPlatform))) Then
    With Mine(i)
        .Y = Platform(iPlatform).Top
        .bOnSurface = True
        .Speed = 0
    End With
End If
    
    

'If Mine(i).X > Platform(iPlatform).Left Then
'    If Mine(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
'
'        If Mine(i).Y > Platform(iPlatform).Top Then
'            If Mine(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
'
'                Mine(i).Y = Platform(iPlatform).Top
'
'                'mineOnSurface = True
'                Mine(i).bOnSurface = True
'                Mine(i).Speed = 0
'
'            End If
'        End If
'
'
'    End If
'End If

End Sub

Private Sub ExplodeMine(ByRef i As Integer, ByVal bSendBroadcast As Boolean)
Dim j As Integer, OwnerIndex As Integer
Dim Dist As Single
Dim ExplosionForceDist As Single, AngleToStick As Single
Dim ang As Single
Const Mine_Explode_RadiusX2 = Mine_Explode_Radius * 2
Const ChopperLenX1p2 = ChopperLen * 1.2
Const NadeMultiple_X = NadeMultiple * 12000
Const ShieldWaveDispersion As Single = 600


If bSendBroadcast Then
    If modStickGame.StickServer Then
        SendBroadcast sExplodeMines & CStr(Mine(i).ID)
    Else
        modWinsock.SendPacket lSocket, ServerSockAddr, sExplodeMines & CStr(Mine(i).ID)
    End If
End If


AddExplosion Mine(i).X, Mine(i).Y, 750
AddSmokeNadeTrail Mine(i).X, Mine(i).Y, True, True
For j = 1 To 10
    AddSmokeGroup Mine(i).X, Mine(i).Y, 5, 75 * Rnd(), PM_Rnd() * piD4
    AddSparks Mine(i).X, Mine(i).Y, piD8 * PM_Rnd()
Next j

For j = 1 To 2 + Rnd() * 3
    AddNadeTrail_Simple Mine(i).X, Mine(i).Y
Next j

If PointHearableOnSticksScreen(Mine(i).X, Mine(i).Y, 0) Then
    modAudio.PlayNadeExplosion GetRelPan(Mine(i).X)
Else
    modAudio.PlayBackGroundNade GetRelPan(Mine(i).X)
End If

OwnerIndex = FindStick(Mine(i).OwnerID)

If OwnerIndex > -1 Then AddFire Mine(i).X, Mine(i).Y, OwnerIndex


For j = 0 To NumSticksM1
    
    If StickInGame(j) Then
        
        Dist = GetDist(Stick(j).X, Stick(j).Y, Mine(i).X, Mine(i).Y)
        ang = FixAngle(FindAngle(Mine(i).X, Mine(i).Y, Stick(j).X, Stick(j).Y))
        If ang < piD2 Or ang > pi3D2 Then
            
            If Stick(j).WeaponType = Chopper Then
                ExplosionForceDist = ChopperLenX1p2
            Else
                ExplosionForceDist = Mine_Explode_RadiusX2
            End If
            
            AngleToStick = FindAngle(Mine(i).X, Mine(i).Y, Stick(j).X, Stick(j).Y)
            
            If Dist < ExplosionForceDist Then
                If Stick(j).WeaponType <> Chopper Then
                    If Stick(j).Shield = 0 Then
                        If StickiHasState(j, STICK_PRONE) = False Then
                            AddVectors Stick(j).Speed, Stick(j).Heading, _
                                NadeMultiple_X / (Dist + 1), AngleToStick, _
                                Stick(j).Speed, Stick(j).Heading
                            
                        End If
                    End If
                End If
            End If
            
            
            
            If Dist < Mine_Explode_Radius Then
                If Stick(j).Shield Then
                    AngleToStick = AngleToStick - Pi
                    AddShieldWave Stick(j).X, Stick(j).Y, AngleToStick
                    AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                    AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                    AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                End If
                
                If j = 0 Or Stick(j).IsBot Then
                    
                    If OwnerIndex <> -1 Then
                        If (IsAlly(Stick(j).Team, Stick(OwnerIndex).Team) = False) Or (j = OwnerIndex) Then
                            'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
                            If StickInvul(j) = False Then
                                
                                On Error Resume Next
                                'Stick(j).Helth = Stick(j).Health - 100000 / Dist
                                
                                If Stick(j).Perk = pZombie Then
                                    DamageStick Mine_Damage * Zombie_Mine_Weakness / Dist, j, OwnerIndex
                                Else
                                    DamageStick Mine_Damage / Dist, j, OwnerIndex
                                End If
                                
                                If Err.Number <> 0 Then 'div zero error
                                    Stick(j).Health = 0
                                    Err.Clear
                                End If
                                
                                If Stick(j).Health < 1 Then
                                    Call Killed(j, OwnerIndex, IIf(Mine(i).bOnSurface, kMine, kAirMine))
                                End If
                                
                            End If 'spawn invul endif
                        End If 'ally endif
                    End If 'owner index endif
                End If 'myid endif
                
                
            End If 'dist endif
        End If 'angle endif
    End If 'stickingame endif
Next j

ExplodeAll Mine(i).X, Mine(i).Y, Mine(i).OwnerID, i, -1

End Sub

Private Sub DrawMines()
Dim i As Integer
Dim tY As Single, tX As Single
Dim TimeLeft As Single

picMain.DrawWidth = 2

For i = 0 To NumMines - 1
    DrawMine Mine(i).X, Mine(i).Y, Mine(i).colour
Next i

If Stick(0).Perk = pBombSquad Then
    For i = 0 To NumMines - 1
        modStickGame.sCircleSE Mine(i).X, Mine(i).Y, CSng(Mine_StickLim), Mine(i).colour, -0.0001, -Pi
        '                                                                                   ^neagtive,
        '                                                                         so we get lines to the centre
        
    Next i
End If

End Sub

Private Sub DrawMine(X As Single, Y As Single, colour As Long)
Const kX = 50, kY = 12
Const Mine_RadiusD2 = Mine_Radius / 2

modStickGame.sBoxFilled X - kX, Y - kY, X + kX, Y + kY, colour
modStickGame.sCircle X, Y - Mine_Radius, Mine_Radius, BoxCol

End Sub

Private Function MineNearStick(iMine As Integer, iStick As Integer) As Boolean
Dim ang As Single

If GetDist(Stick(iStick).X, Stick(iStick).Y, Mine(iMine).X, Mine(iMine).Y) < Mine_StickLim Then
    ang = FixAngle(FindAngle(Mine(iMine).X, Mine(iMine).Y, Stick(iStick).X, Stick(iStick).Y))
    MineNearStick = (ang < piD2 Or ang > pi3D2)
End If

'If Mine(iMine).X > (Stick(iStick).X - Mine_StickLimY) Then
'    If Mine(iMine).X < (Stick(iStick).X + Mine_StickLimY) Then
'
'        If Mine(iMine).Y > (Stick(iStick).Y - Mine_StickLimY) Then
'            If Mine(iMine).Y < (Stick(iStick).Y + Mine_StickLimY) Then
'                MineNearStick = True
'            End If
'        End If
'
'    End If
'End If

End Function

Private Function MineInBullet(Minei As Integer) As Boolean
Dim i As Integer

For i = 0 To NumBullets - 1
    If BulletNearMine(Minei, i) Then
        MineInBullet = True
        Exit For
    End If
Next i

End Function

Private Function BulletNearMine(Minei As Integer, Bulleti As Integer) As Boolean
Const MineLim = 150

If Bullet(Bulleti).bHeadingChanged = False Or Bullet(Bulleti).bSniperBullet Then
    If Bullet(Bulleti).X > (Mine(Minei).X - MineLim) Then
        If Bullet(Bulleti).X < (Mine(Minei).X + MineLim) Then
            
            If Bullet(Bulleti).Y > (Mine(Minei).Y - MineLim) Then
                If Bullet(Bulleti).Y < (Mine(Minei).Y + MineLim) Then
                    BulletNearMine = True
                End If
            End If
            
        End If
    End If
End If

End Function

Private Function BulletNearBarrel(iBullet As Integer, iBarrel As Integer) As Boolean
Const ex_BarrelWidth = BarrelWidth * 1.4

'If Bullet(iBullet).bHeadingChanged=False Or Bullet(iBullet).bSniperBullet Then
    If Bullet(iBullet).X > (Barrel(iBarrel).X - ex_BarrelWidth) Then
        If Bullet(iBullet).X < (Barrel(iBarrel).X + ex_BarrelWidth) Then
            
            If Bullet(iBullet).Y > (Barrel(iBarrel).Y) Then
                If Bullet(iBullet).Y < (Barrel(iBarrel).Y + BarrelHeight) Then
                    BulletNearBarrel = True
                End If
            End If
            
        End If
    End If
'End If

End Function

Private Function NadeNearBarrel(iNade As Integer, iBarrel As Integer) As Boolean
Dim rcNade As RECT, rcBarrel As RECT

With rcNade
    .Left = Nade(iNade).X
    .Right = .Left + Nade_Radius
    .Top = Nade(iNade).Y
    .Bottom = .Top + Nade_Radius
End With
With rcBarrel
    .Left = Barrel(iBarrel).X
    .Right = .Left + BarrelWidth
    .Top = Barrel(iBarrel).Y
    .Bottom = .Top + BarrelHeight
End With

NadeNearBarrel = RectCollision(rcNade, rcBarrel)
    
'If Nade(iNade).X > (Barrel(iBarrel).X - BarrelWidth) Then
'    If Nade(iNade).X < (Barrel(iBarrel).X + BarrelWidth) Then
'        If Nade(iNade).Y > (Barrel(iBarrel).Y) Then
'            If Nade(iNade).Y < (Barrel(iBarrel).Y + BarrelHeight) Then
'                NadeNearBarrel = True
'            End If
'        End If
'    End If
'End If

End Function

Private Sub DrawSmoke()
Dim i As Integer

picMain.FillStyle = vbFSSolid
picMain.FillColor = SmokeFill

For i = 0 To NumSmoke - 1
    With Smoke(i)
        modStickGame.sCircle .X, .Y, .Size, SmokeOutline
    End With
Next i

picMain.FillStyle = vbFSTransparent

End Sub

Private Sub ProcessSmoke()
Dim i As Integer
Dim f As Single
Const Speed_Decrease = 0.5

picMain.FillColor = SmokeFill
picMain.FillStyle = vbFSSolid 'vbopaque
picMain.DrawWidth = 1

Do While i < NumSmoke
    
    If Smoke(i).Size <= 0 Then
        RemoveSmoke i
        i = i - 1
    ElseIf Smoke(i).Size > 40 And Smoke(i).Direction = 1 Then
        Smoke(i).Direction = -1
    ElseIf Smoke(i).Size > 100 Then
        RemoveSmoke i 'for when debugging, etc
        i = i - 1
    End If
    
    
    i = i + 1
    
Loop

For i = 0 To NumSmoke - 1
    
    With Smoke(i)
        MotionStickObject .X, .Y, .Speed, .Heading
        'f = GetTimeZoneAdjust(.X, .Y)
        'Motion .X, .Y, .Speed * f, .Heading
        
        '.Speed = .Speed / 1.1 * modStickGame.StickTimeFactor / f
        
        
        f = modStickGame.StickTimeFactor * GetTimeZoneAdjust(.X, .Y) / IIf(.bLongTime, 2.5, 1)
        
        If .Speed > 0 Then .Speed = .Speed - Speed_Decrease * f
        
        
        If .Direction = 1 Then
            .Size = .Size + 2 * f
        Else
            .Size = .Size - 0.5 * f
        End If
    End With
    
Next i

picMain.FillStyle = vbFSTransparent 'transparent

End Sub

Private Sub DrawMuzzleFlashes()
Dim i As Integer

'Do While i < NumMFlashes
'
'    If MFlash(i).Decay < GetTickCount() Then
'        RemoveMFlash i
'        i = i - 1
'    End If
'
'    i = i + 1
'Loop
'
'picMain.ForeColor = vbYellow
'For i = 0 To NumMFlashes - 1
'    DrawMFlash MFlash(i).X, MFlash(i).Y, MFlash(i).Facing
'Next i

picMain.ForeColor = vbYellow
'picMain.DrawWidth = 1.5
For i = 0 To NumSticksM1
    If Stick(i).LastMuzzleFlash + MFlash_Time / GetSticksTimeZone(i) > GetTickCount() Then
        If StickInGame(i) Then
            DrawMFlash Stick(i).GunPoint.X, Stick(i).GunPoint.Y, Stick(i).Facing
        End If
    End If
Next i

End Sub

Private Sub DrawMFlash(X As Single, Y As Single, Facing As Single)

Const SideLen = 10, FrontLen = 110
Const SideFlashLen = 65 ', FrontFlashLen = 5

Dim Pts(1 To 3) As PointAPI
Dim Rd As Single, RD2 As Single

Rd = Rnd()
RD2 = Rnd()

Pts(1).X = X + SideLen * Sine(Facing - piD2) * Rd
Pts(1).Y = Y - SideLen * CoSine(Facing - piD2) * Rd

Pts(2).X = X + FrontLen * Sine(Facing) * Rd
Pts(2).Y = Y - FrontLen * CoSine(Facing) * Rd

'Pts(3).x = x + FrontLen * sine(Facing + piD10)
'Pts(3).y = y - FrontLen * cosine(Facing + piD10)

Pts(3).X = X + SideLen * Sine(Facing + piD2) * Rd
Pts(3).Y = Y - SideLen * CoSine(Facing + piD2) * Rd

modStickGame.sPoly Pts, vbYellow

picMain.ForeColor = vbYellow
picMain.DrawWidth = 2

modStickGame.sLine X, Y, _
    X + SideFlashLen * Sine(Facing - piD3) * RD2, _
    Y - SideFlashLen * CoSine(Facing - piD3) * RD2

modStickGame.sLine X, Y, _
    X + SideFlashLen * Sine(Facing + piD3) * RD2, _
    Y - SideFlashLen * CoSine(Facing + piD3) * RD2


End Sub

Private Sub DrawMagazines()
Dim i As Integer

picMain.DrawWidth = 1
picMain.ForeColor = vbBlack

For i = 0 To NumMags - 1
    DrawMagazine i
Next i

End Sub

Private Sub ProcessMagazines()
Dim i As Integer, j As Integer

Do While i < NumMags
    
    If Mag(i).Decay < GetTickCount() Then
        RemoveMag i
        i = i - 1
    End If
    
    i = i + 1
Loop


For i = 0 To NumMags - 1
    
    If Mag(i).bOnSurface = False Then
        
        ApplyGravityVector Mag(i).LastGravity, GetTimeZoneAdjust(Mag(i).X, Mag(i).Y), _
            Mag(i).Speed, Mag(i).Heading, Mag(i).X, Mag(i).Y
'        If Mag(i).LastGravity + Gravity_Delay / GetTimeZoneAdjust(Mag(i).X, Mag(i).Y) < GetTickCount() Then
'            AddVectors Mag(i).Speed, Mag(i).Heading, Gravity_Strength, Gravity_Direction, _
'                Mag(i).Speed, Mag(i).Heading
'
'            Mag(i).LastGravity = GetTickCount()
'        End If
        
        
        MotionStickObject Mag(i).X, Mag(i).Y, Mag(i).Speed, Mag(i).Heading
        For j = 0 To ubdPlatforms
            MagOnSurface i, j
        Next j
        
        
        ClipMag i
    End If
Next i

End Sub

Private Sub ClipMag(i As Integer)

If Mag(i).X > (StickGameWidth - Lim) Then
    
    If Mag(i).Heading < Pi Then
        ReverseXComp Mag(i).Speed, Mag(i).Heading
    End If
    Mag(i).X = StickGameWidth - Lim
    
ElseIf Mag(i).X < Lim Then
    
    If Mag(i).Heading > Pi Then
        ReverseXComp Mag(i).Speed, Mag(i).Heading
    End If
    Mag(i).X = Lim
    
End If

If Mag(i).Y < Lim Then
    Mag(i).Y = Lim
    If Mag(i).Heading < piD2 Or Mag(i).Heading > pi3D2 Then
        ReverseYComp Mag(i).Speed, Mag(i).Heading
    End If
ElseIf Mag(i).Y > StickGameHeight Then
    Mag(i).bOnSurface = True
    Mag(i).Y = Platform(0).Top - 80
End If


End Sub

Private Sub DrawMagazine(i As Integer)
Dim pt(1 To 4) As PointAPI
Dim j As Integer

If Mag(i).iMagType = mAK Then
    'top left
    pt(1).X = Mag(i).X
    pt(1).Y = Mag(i).Y - 50
    
    'top right
    pt(2).X = Mag(i).X + 50
    pt(2).Y = pt(1).Y
    
    'bottom left
    pt(4).X = Mag(i).X + 10
    pt(4).Y = Mag(i).Y + 70
    
    'bottom right
    pt(3).X = Mag(i).X + 60
    pt(3).Y = Mag(i).Y + 70
    
    picMain.DrawWidth = 2
    modStickGame.sPoly pt, -1
    
ElseIf Mag(i).iMagType = mXM8 Then
    'top left
    pt(1).X = Mag(i).X
    pt(1).Y = Mag(i).Y - 50
    
    'top right
    pt(2).X = Mag(i).X + 50
    pt(2).Y = pt(1).Y
    
    'bottom left
    pt(4).X = Mag(i).X + 10
    pt(4).Y = Mag(i).Y + 70
    
    'bottom right
    pt(3).X = Mag(i).X + 60
    pt(3).Y = Mag(i).Y + 70
    
    picMain.DrawWidth = 1
    modStickGame.sPoly pt, vbBlack
    
ElseIf Mag(i).iMagType = mSniper Then
    'top left
    pt(1).X = Mag(i).X
    pt(1).Y = Mag(i).Y
    
    'top right
    pt(2).X = Mag(i).X + 50
    pt(2).Y = pt(1).Y
    
    'bottom left
    pt(4).X = Mag(i).X + 10
    pt(4).Y = Mag(i).Y + 50
    
    'bottom right
    pt(3).X = Mag(i).X + 75
    pt(3).Y = Mag(i).Y + 50
    
    picMain.DrawWidth = 1
    modStickGame.sPoly pt, -1
    
ElseIf Mag(i).iMagType = mPistol Then
    'top left
    pt(1).X = Mag(i).X
    pt(1).Y = Mag(i).Y - 25
    
    'top right
    pt(2).X = Mag(i).X + 25
    pt(2).Y = pt(1).Y
    
    'bottom left
    pt(4).X = Mag(i).X + 5
    pt(4).Y = Mag(i).Y + 60
    
    'bottom right
    pt(3).X = Mag(i).X + 30
    pt(3).Y = Mag(i).Y + 60
    picMain.DrawWidth = 1
    modStickGame.sPoly pt, vbBlack
    
ElseIf Mag(i).iMagType = mFlameThrower Then
    pt(1).X = Mag(i).X
    pt(1).Y = Mag(i).Y
    
    pt(2).X = pt(1).X + GunLen / 2
    pt(2).Y = pt(1).Y
    
    pt(3).X = pt(2).X
    pt(3).Y = pt(2).Y + GunLen / 3
    
    pt(4).X = pt(3).X - GunLen / 4
    pt(4).Y = pt(3).Y
    
    picMain.DrawWidth = 2
    modStickGame.sPoly pt, vbRed
    
ElseIf Mag(i).iMagType = mAUG Then
    'top left
    pt(1).X = Mag(i).X
    pt(1).Y = Mag(i).Y - 50
    
    'top right
    pt(2).X = Mag(i).X + 25
    pt(2).Y = pt(1).Y
    
    'bottom left
    pt(4).X = Mag(i).X + 10
    pt(4).Y = Mag(i).Y + 70
    
    'bottom right
    pt(3).X = Mag(i).X + 30
    pt(3).Y = Mag(i).Y + 70
    
    picMain.DrawWidth = 1
    modStickGame.sPoly pt, vbBlack
End If


End Sub

Private Sub MagOnSurface(i As Integer, iPlatform As Integer) 'As Boolean

'Dim rcMag As RECT
'
'With rcMag
'    .Left = Mag(i).X
'    .Right = .Left + 1
'    .Top = Mag(i).Y
'    .Bottom = .Top + 1
'End With

If RectCollision(PointToRect(Mag(i).X, Mag(i).Y), PlatformToRect(Platform(iPlatform))) Then
    'position Mag on top of the platform
    Mag(i).Y = Platform(iPlatform).Top - 80

    Mag(i).bOnSurface = True
    Mag(i).Speed = 0
End If


'If Mag(i).X > Platform(iPlatform).Left Then
'    If Mag(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
'
'
'        If Mag(i).Y > Platform(iPlatform).Top - 10 Then
'            If Mag(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
'
'                'position the Mag on top of the platform
'                Mag(i).Y = Platform(iPlatform).Top - 80
'                'If Mag(i).Y > (Platform(iPlatform).Top + 30) Then
'                'End If
'
'                Mag(i).bOnSurface = True
'                Mag(i).Speed = 0
'
'            End If
'        End If
'
'
'    End If
'End If

End Sub

Private Sub ClipStaticWeapon(i As Integer)
'Const Max_SW_Speed = 200

If StaticWeapon(i).X < 10 Then
    StaticWeapon(i).X = 10
    
    ReverseXComp StaticWeapon(i).Speed, StaticWeapon(i).Heading
    StaticWeapon(i).Speed = StaticWeapon(i).Speed / 2
    
ElseIf StaticWeapon(i).X > StickGameWidth Then
    StaticWeapon(i).X = StickGameWidth
    
    ReverseXComp StaticWeapon(i).Speed, StaticWeapon(i).Heading
    StaticWeapon(i).Speed = StaticWeapon(i).Speed / 2
End If


If StaticWeapon(i).Y < 1 Then
    StaticWeapon(i).Y = 1
    StaticWeapon(i).Speed = StaticWeapon(i).Speed / 2
    ReverseYComp StaticWeapon(i).Speed, StaticWeapon(i).Heading
End If


End Sub

Private Sub ProcessStaticWeaponPickup()
Dim i As Integer
Dim bPrompted As Boolean
Dim oldWeap As eWeaponTypes
Dim sText As String

'If modStickGame.sv_2Weapons = False Then Exit Sub

On Error GoTo EH

If NumStaticWeapons < Min_Static_Weapons Then
    If modStickGame.StickServer Then
        Erase StaticWeapon: NumStaticWeapons = 0
        MakeStaticWeapons
    End If
End If

bPrompted = Not StickInGame(0)
'i.e. if not in game, don't say anything

Do While i < NumStaticWeapons
    
    'Debug.Print GetWeaponName(StaticWeapon(i).iWeapon)
    
    If StaticWeapon(i).Y > StickGameHeight Then
        RemoveStaticWeapon i
        i = i - 1
        
    Else
        
        If StickNearStaticWeapon(0, i) Then
            If StickiHasWeapon(0, StaticWeapon(i).iWeapon) = False Then
                If UseKey Then
                    If (Stick(0).LastWeaponSwitch + SwitchWeaponDelay < GetTickCount()) And Stick(0).WeaponType < Knife Then
                        'pickup the weapon and
                        'decide whether we are swapping currentweapon(1) or (2)
                        
                        If Stick(0).CurrentWeapons(1) = Stick(0).WeaponType Then
                            'we are swapping CW(1)
                            Stick(0).CurrentWeapons(1) = StaticWeapon(i).iWeapon
                        Else
                            Stick(0).CurrentWeapons(2) = StaticWeapon(i).iWeapon
                        End If
                        
                        modAudio.PlayWeaponPickUpSound
                        
                        oldWeap = Stick(0).WeaponType
                        
                        SwitchWeapon StaticWeapon(i).iWeapon, False
                        
                        UseKey = False
                        
                        
                        '###############################################################
                        sText = sWeaponSwapInfos & Stick(0).ID & "#" & _
                            CStr(i) & "#" & _
                            CStr(Stick(0).WeaponType) & "#" & _
                            CStr(oldWeap)
                        
                        If modStickGame.StickServer Then
                            SendBroadcast sText
                        Else
                            modWinsock.SendPacket lSocket, ServerSockAddr, sText
                        End If
                        '###############################################################
                        
                        'drop current weapon and remove pickup up one
                        RemoveStaticWeapon i
                        i = i - 1
                        AddStaticWeapon Stick(0).X, Stick(0).Y, oldWeap 'Stick(0).PrevWeapon
                        
                        
                        
                        Stick(0).LastWeaponSwitch = GetTickCount()
                    ElseIf Stick(0).WeaponType >= Knife Then
                        UseKey = False
                    End If
                    
                ElseIf Not bPrompted Then
                    
                    If Stick(0).WeaponType < Knife Then
                        PrintStickText "Press E to pick up " & GetWeaponName(StaticWeapon(i).iWeapon), _
                            Stick(0).X - 1000, Stick(0).Y - 1000, vbBlack
                        
                    End If
                    
                    bPrompted = True
                End If
            Else
                'UseKey = False
            End If
        End If
        
        
'        'check if sticks are near any
'        For 0 = 0 To NumSticksM1
'            If StickInGame(0) Then
'                If Stick(0).WeaponType <> Knife And Stick(0).WeaponType <> Chopper Then
'                    If StickiHasState(0, Stick_Use) And StickiHasState(0, Stick_Reload) = False Then
'                        If StickNearStaticWeapon(0, i) Then
'
'                            If Stick(0).LastWeaponSwitch + SwitchWeaponDelay < GetTickCount() Then
'                                '                                   /GetSticksTimeZone(0)
'
'                                If StickiHasWeapon(0, StaticWeapon(i).iWeapon) = False Then
'
'                                    'pickup the weapon
'                                    If 0 = 0 Then
'                                        'decide whether we are swapping currentweapon(1) or (2)
'                                        If Stick(0).CurrentWeapons(1) = Stick(0).WeaponType Then
'                                            'we are swapping CW(1)
'                                            Stick(0).CurrentWeapons(1) = StaticWeapon(i).iWeapon
'                                        Else
'                                            Stick(0).CurrentWeapons(2) = StaticWeapon(i).iWeapon
'                                        End If
'
'                                        If 0 = 0 Then modAudio.PlayWeaponPickUpSound
'
'                                        SwitchWeapon StaticWeapon(i).iWeapon, False
'
'                                        'On Error Resume Next
'                                        'AmmoFired(StaticWeapon(i).iWeapon) = 0
'                                        'Stick(0).BulletsFired = 0
'                                        'Stick(0).BulletsFired = AmmoFired(StaticWeapon(i).iWeapon)
'
'                                        UseKey = False
'                                    Else
'                                        'Stick(0).WeaponType = StaticWeapon(i).iWeapon
'                                        SetSticksWeapon 0, StaticWeapon(i).iWeapon
'                                    End If
'
'
'
'                                    'drop current weapon
'                                    RemoveStaticWeapon i
'                                    i = i - 1
'                                    AddStaticWeapon Stick(0).X, Stick(0).Y, Stick(0).PrevWeapon
'
'                                    'SubStickState Stick(0).ID, Stick_Use
'                                    Stick(0).LastWeaponSwitch = GetTickCount()
'
'
'                                    Exit For
'
'                                'Else
'
'                                    ''stick has the weapon, sub the state
'                                    'If 0 = 0 Then
'                                        'UseKey = False
'                                    'End If
'
'                                    'SubStickiState 0, Stick_Use
'
'                                    'Exit Do
'                                Else
'                                    If 0 = 0 Then UseKey = False: SubStickiState 0, Stick_Use
'                                End If
'                            Else 'recently swapped
'                                If 0 = 0 Then UseKey = False: SubStickiState 0, Stick_Use
'                                'If 0 = 0 Then
'                                    'UseKey = False
'                                'End If
'                                'SubStickState Stick(0).ID, Stick_Use
'                            End If
'                        'Else 'not near weapon
'                            'don't sub the state, because they may be near the next weapon
'                            'If 0 = 0 Then
'                                'UseKey = False
'                            'End If
'                            'SubStickState Stick(0).ID, Stick_Use
'                        End If
'                    ElseIf StickNearStaticWeapon(0, i) Then
'                        'doesn't have use state...
'
'                        If 0 = 0 Then
'                            If bPrompted = False Then
'                                If StickiHasWeapon(0, StaticWeapon(i).iWeapon) = False Then
'                                    If StickiHasState(0, Stick_Reload) = False Then
'                                        PrintStickText "Press E to pick up " & GetWeaponName(StaticWeapon(i).iWeapon), _
'                                            Stick(0).X - 1000, Stick(0).Y - 1000, vbBlack
'                                    End If
'
'                                    bPrompted = True
'                                End If
'                            End If
'                        End If
'
'                        'prevent from 'using' after reload
'                        If StickiHasState(0, Stick_Reload) Then
'                            SubStickiState 0, Stick_Use
'                            If 0 = 0 Then UseKey = False
'                        End If
'
'                    End If
'                End If
'            End If
'        Next 0
    End If
    
    
    i = i + 1
Loop

EH:
End Sub

Private Sub ProcessStaticWeapons()
Dim i As Integer, j As Integer
Dim iHave(0 To eWeaponTypes.Knife - 1) As Integer
Dim bAdd As Boolean

picMain.DrawWidth = 1
Me.picMain.ForeColor = vbBlack

On Error GoTo EH

For i = 0 To NumStaticWeapons - 1
    
    If StaticWeapon(i).bOnSurface = False Then
        ClipStaticWeapon i
        
'        If StaticWeapon(i).LastGravity + Gravity_Delay / _
'                GetTimeZoneAdjust(StaticWeapon(i).X, StaticWeapon(i).Y) < GetTickCount() Then
'
'
'            AddVectors StaticWeapon(i).Speed, StaticWeapon(i).Heading, Gravity_Strength, Gravity_Direction, _
'                StaticWeapon(i).Speed, StaticWeapon(i).Heading
'
'            StaticWeapon(i).LastGravity = GetTickCount()
'        End If
        ApplyGravityVector StaticWeapon(i).LastGravity, GetTimeZoneAdjust(StaticWeapon(i).X, StaticWeapon(i).Y), _
            StaticWeapon(i).Speed, StaticWeapon(i).Heading, StaticWeapon(i).X, StaticWeapon(i).Y
        
        
        MotionStickObject StaticWeapon(i).X, StaticWeapon(i).Y, StaticWeapon(i).Speed, StaticWeapon(i).Heading
        
        For j = 0 To ubdPlatforms
            StaticWeaponOnSurface i, j
        Next j
        
    End If
    
    
    If modStickGame.sv_AllowedWeapons(StaticWeapon(i).iWeapon) = False Then
        StaticWeapon(i).iWeapon = GetRandomStaticWeapon()
    End If
    
    
    If StaticWeapon(i).iWeapon >= Knife Then
        StaticWeapon(i).iWeapon = GetRandomStaticWeapon()
    End If
    
Next i


'If modStickGame.StickServer Then
    'check we have them all
    For i = 0 To NumStaticWeapons - 1
        iHave(StaticWeapon(i).iWeapon) = iHave(StaticWeapon(i).iWeapon) + 1
    Next i
    
    For i = 0 To eWeaponTypes.Knife - 1
        If iHave(i) = 0 Then
            
            bAdd = True
            
            If modStickGame.sv_AllowedWeapons(i) = False Then
                bAdd = False
            End If
            
            If bAdd Then
                'find a weapon, and make it be this weapon type
                For j = 0 To NumStaticWeapons - 1
                    If iHave(StaticWeapon(j).iWeapon) > 1 Then
                        
                        iHave(StaticWeapon(j).iWeapon) = iHave(StaticWeapon(j).iWeapon) - 1
                        
                        StaticWeapon(j).iWeapon = i
                        
                        iHave(StaticWeapon(j).iWeapon) = iHave(StaticWeapon(j).iWeapon) + 1
                        
                        Exit For
                    End If
                Next j
            End If
            
'        ElseIf iHave(i) > 5 Then
'
'            For 0 = 0 To NumStaticWeapons - 1
'                If 0 <> i Then
'                    If iHave(StaticWeapon(0).iWeapon) < 2 Then
'
'                        iHave(StaticWeapon(0).iWeapon) = iHave(StaticWeapon(0).iWeapon) + 1
'
'                        StaticWeapon(0).iWeapon = i
'
'                        iHave(StaticWeapon(0).iWeapon) = iHave(StaticWeapon(0).iWeapon) - 1
'
'                        Exit For
'                    End If
'                End If
'            Next 0
            
        End If
    Next i
'End If


EH:
End Sub

Private Sub ProcessWeaponSwapInfo(sInfo As String)
Dim iStatic As Integer
Dim newWeap As eWeaponTypes, oldWeap As eWeaponTypes
Dim iID As Integer
Dim sParts() As String

'sInfo = ID#iStaticWeapon#NewWeap#OldWeap
On Error GoTo EH
sParts = Split(sInfo, "#")

iID = CInt(sParts(0))
iStatic = CInt(sParts(1))
newWeap = CInt(sParts(2))
oldWeap = CInt(sParts(3))

If 0 < iStatic And iStatic < NumStaticWeapons Then
    'can't use with block - addstaticweap() removes from staticweapon() if too many
    If StaticWeapon(iStatic).iWeapon = newWeap Then
        
        
        AddStaticWeapon CSng(StaticWeapon(iStatic).X), StaticWeapon(iStatic).Y - BodyLen, oldWeap
        'V. IMPORTANT    ^ CSNG MUST BE THERE, SO THE VALUE IS PASSED BYVAL, SO THE ARRAY ISN'T LOCKED, SO ONE CAN BE ADDED
        
        RemoveStaticWeapon iStatic
        
        
        'i would swap the weapon here, but it'll just go back as an old packet is received,
        'then forward as the newer one arrives, so i've left it
        
    'Else
        'wrong weap, screw it
    End If
    
    'invalid static weapon
End If


EH:
Erase sParts
End Sub

Private Sub DrawStaticWeapons()
Dim i As Integer

For i = 0 To NumStaticWeapons - 1
    DrawStaticWeapon i
Next i

End Sub

Private Sub StaticWeaponOnSurface(i As Integer, iPlatform As Integer) 'As Boolean

'Dim rcWep As RECT
'
'With rcWep
'    .Left = StaticWeapon(i).X
'    .Right = .Left + 1
'    .Top = StaticWeapon(i).Y
'    .Bottom = .Top + 1
'End With

If RectCollision(PointToRect(StaticWeapon(i).X, StaticWeapon(i).Y), PlatformToRect(Platform(iPlatform))) Then
    'position the StaticWeapon on top of the platform
    StaticWeapon(i).Y = Platform(iPlatform).Top '- 80
    
    StaticWeapon(i).bOnSurface = True
    StaticWeapon(i).Speed = 0
    AddSparks StaticWeapon(i).X, StaticWeapon(i).Y, StaticWeapon(i).Heading - Pi
End If

'If StaticWeapon(i).X > Platform(iPlatform).Left Then
'    If StaticWeapon(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
'
'
'        If StaticWeapon(i).Y > Platform(iPlatform).Top - 10 Then
'            If StaticWeapon(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
'
'                'position the StaticWeapon on top of the platform
'                StaticWeapon(i).Y = Platform(iPlatform).Top - 80
'
'                StaticWeapon(i).bOnSurface = True
'                StaticWeapon(i).Speed = 0
'                AddSparks StaticWeapon(i).X, StaticWeapon(i).Y, StaticWeapon(i).Heading - Pi
'
'            End If
'        End If
'
'
'    End If
'End If

End Sub

Private Sub DrawStaticWeapon(i As Integer)

If modStickGame.cg_SimpleStaticWeapons Then
    
    modStickGame.sCircle StaticWeapon(i).X, StaticWeapon(i).Y, 100, vbBlack
    modStickGame.PrintStickText GetWeaponName(StaticWeapon(i).iWeapon), StaticWeapon(i).X, StaticWeapon(i).Y - 500, vbBlack
    
Else
    
    Me.DrawWidth = 1
    
    Select Case StaticWeapon(i).iWeapon
        Case XM8
            DrawStaticXM8 StaticWeapon(i).X, StaticWeapon(i).Y
        Case AK
            DrawStaticAK StaticWeapon(i).X, StaticWeapon(i).Y
        Case DEagle
            DrawStaticDEagle StaticWeapon(i).X, StaticWeapon(i).Y
        Case USP
            DrawStaticUSP StaticWeapon(i).X, StaticWeapon(i).Y
        Case FlameThrower
            DrawStaticFlameThrower StaticWeapon(i).X, StaticWeapon(i).Y
        Case M249
            DrawStaticM249 StaticWeapon(i).X, StaticWeapon(i).Y
        Case M82
            DrawStaticM82 StaticWeapon(i).X, StaticWeapon(i).Y
        Case AWM
            DrawStaticAWM StaticWeapon(i).X, StaticWeapon(i).Y
        Case RPG
            DrawStaticRPG StaticWeapon(i).X, StaticWeapon(i).Y
        Case W1200
            DrawStaticW1200 StaticWeapon(i).X, StaticWeapon(i).Y
        Case AUG
            DrawStaticAUG StaticWeapon(i).X, StaticWeapon(i).Y
        Case MP5
            DrawStaticMP5 StaticWeapon(i).X, StaticWeapon(i).Y
        Case Mac10
            DrawStaticMac10 StaticWeapon(i).X, StaticWeapon(i).Y
        Case SPAS
            DrawStaticSPAS StaticWeapon(i).X, StaticWeapon(i).Y
        Case G3
            DrawStaticG3 StaticWeapon(i).X, StaticWeapon(i).Y
    End Select
    
End If

End Sub

Private Function StickNearStaticWeapon(iStick As Integer, iSWeapon As Integer) As Boolean
Const StaticWeaponLim = 400, BodyLenX2 = BodyLen * 2.4
Dim sY As Single

If Stick(iStick).X > (StaticWeapon(iSWeapon).X - StaticWeaponLim) Then
    If Stick(iStick).X < (StaticWeapon(iSWeapon).X + StaticWeaponLim) Then
        
        sY = GetStickY(iStick)
        
        If sY > (StaticWeapon(iSWeapon).Y - BodyLenX2) Then
            If sY < (StaticWeapon(iSWeapon).Y + StaticWeaponLim) Then
                StickNearStaticWeapon = True
            End If
        End If
        
    End If
End If

End Function

Private Function GetRandomPlatform() As Integer
GetRandomPlatform = Rnd() * modStickGame.ubdPlatforms
End Function
Private Function RandomXOnPlatform(iPlatform As Integer) As Single

'RandomXOnPlatform = Platform(iPlatform).Left + Rnd() * (Platform(iPlatform).width - 400)

RandomXOnPlatform = (Platform(iPlatform).Left + 300) + Rnd() * (Platform(iPlatform).width - IIf(iPlatform = 0, 1000, 400))

End Function
Private Function YOnPlatform(iPlatform As Integer) As Single
YOnPlatform = Platform(iPlatform).Top - 100
End Function

Public Sub MakeStaticWeapons()
Dim i As Single
Dim iPlatform As Integer

For i = 0 To eWeaponTypes.Knife - 1 '- 0.25 Step 0.25
    
    iPlatform = GetRandomPlatform()
    'If iPlatform = 6 Or iPlatform = 7 Or iPlatform = 3 Then
        'reduce amount of weapons in sniper's nest and sniper's nest jump and another
        
        'iPlatform = GetRandomPlatform()
        
        'If iPlatform = 7 Or iPlatform = 6 Then
            'iPlatform = GetRandomPlatform()
        'End If
    'End If
    
    
    AddStaticWeapon RandomXOnPlatform(iPlatform), _
                    YOnPlatform(iPlatform), _
                    CInt(i)
    
Next i


'For iPlatform = 0 To ubdPlatforms \ 2
'    AddStaticWeapon RandomXOnPlatform(iPlatform), _
'                    YOnPlatform(iPlatform), _
'                    GetRandomStaticWeapon()
'Next iPlatform

End Sub

Public Function GetRandomStaticWeapon() As eWeaponTypes
'any up to knife, not including knife
Dim vWep As eWeaponTypes


Do
    vWep = CInt(Rnd() * eWeaponTypes.Knife)
Loop Until vWep <> Knife And modStickGame.sv_AllowedWeapons(vWep)


GetRandomStaticWeapon = vWep
End Function

Public Sub RemoveStaticWeapons()
Erase StaticWeapon
NumStaticWeapons = 0
End Sub

Public Sub SetCurrentWeapons()
If Stick(0).WeaponType <> Stick(0).CurrentWeapons(1) Then
    If Stick(0).WeaponType <> Stick(0).CurrentWeapons(2) Then
        If Stick(0).WeaponType <> Chopper Then
            If Stick(0).WeaponType <> Knife Then
                Stick(0).CurrentWeapons(1) = Stick(0).WeaponType
            Else
                Stick(0).CurrentWeapons(1) = AK
            End If
        Else
            Stick(0).CurrentWeapons(1) = AK
        End If
    End If
End If
End Sub

Private Function GetSticksSecondWeapon(iStick As Integer) As eWeaponTypes
If Stick(iStick).WeaponType = Stick(iStick).CurrentWeapons(1) Then
    GetSticksSecondWeapon = Stick(iStick).CurrentWeapons(2)
Else
    GetSticksSecondWeapon = Stick(iStick).CurrentWeapons(1)
End If
End Function

Private Function StickiHasWeapon(iStick As Integer, vWeapon As eWeaponTypes) As Boolean

If Stick(iStick).CurrentWeapons(1) = vWeapon Then
    StickiHasWeapon = True
ElseIf Stick(iStick).CurrentWeapons(2) = vWeapon Then
    StickiHasWeapon = True
End If

End Function


'STATIC WEAPON DRAWING
'#########################################################################################################
Private Sub DrawStaticW1200(sX As Single, sY As Single)

Dim X(1 To 11) As Single, Y(1 To 11) As Single
Const Facing As Single = piD2
Const SAd2 = SmallAngle / 2

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 2 * Sine(Facing - SmallAngle)
Y(2) = Y(1) - GunLen / 2 * CoSine(Facing - SmallAngle)

X(3) = X(1) + GunLen / 1.5 * Sine(Facing - SmallAngle)
Y(3) = Y(1) - GunLen / 1.5 * CoSine(Facing - SmallAngle)

X(4) = X(1) + GunLen / 1.5 * Sine(Facing - SAd2)
Y(4) = Y(1) - GunLen / 1.5 * CoSine(Facing - SAd2)

X(5) = X(1) + GunLen * Sine(Facing - SAd2)
Y(5) = Y(1) - GunLen * CoSine(Facing - SAd2)

'pump action bit
X(6) = X(1) + GunLen * Sine(Facing - SAd2)
Y(6) = Y(1) - GunLen * CoSine(Facing - SAd2)

X(7) = X(1) + GunLen * 1.5 * Sine(Facing - SmallAngle / 3)
Y(7) = Y(1) - GunLen * 1.5 * CoSine(Facing - SmallAngle / 3)
'end pump action bit

X(8) = X(1) + GunLen * 2 * Sine(Facing - SmallAngle / 3)
Y(8) = Y(1) - GunLen * 2 * CoSine(Facing - SmallAngle / 3)

X(9) = X(1) + GunLen * 2.5 * Sine(Facing - SmallAngle / 3.5)
Y(9) = Y(1) - GunLen * 2.5 * CoSine(Facing - SmallAngle / 3.5)

X(10) = X(9) + GunLen / 6 * Sine(Facing - pi2d3)
Y(10) = Y(9) - GunLen / 6 * CoSine(Facing - pi2d3)

X(11) = X(9) + GunLen / 20 * Sine(Facing - Pi)
Y(11) = Y(9) - GunLen / 20 * CoSine(Facing - Pi)

'end calculation

'picMain.ForeColor = &H555555
picMain.DrawWidth = 2

'handle section
picMain.ForeColor = vbRed
modStickGame.sLine X(1), Y(1), X(3), Y(3)

picMain.DrawWidth = 2
picMain.ForeColor = vbBlack
modStickGame.sLine X(2), Y(2), X(4), Y(4)

picMain.ForeColor = &H555555
modStickGame.sLine X(2), Y(2), X(8), Y(8)
modStickGame.sLine X(3), Y(3), X(9), Y(9)

picMain.ForeColor = vbRed
modStickGame.sLine X(1), Y(1), X(4), Y(4)
modStickGame.sLine X(6), Y(6), X(7), Y(7)

'picMain.ForeColor = &H555555
picMain.DrawWidth = 1
modStickGame.sLine X(10), Y(10), X(11), Y(11)

End Sub

Private Sub DrawStaticSPAS(sX As Single, sY As Single)


Dim BarrelStart(1 To 2) As PointAPI 'top and bottom
Dim BarrelEnd(1 To 2) As PointAPI 'top and bottom
Dim pMain(1 To 4) As PointAPI, pStock(1 To 4) As PointAPI
Dim Handle1X As Single, Handle1Y As Single ', Handle2X As Single, Handle2Y As Single
Dim ForesightX As Single, ForesightY As Single
Dim RearSightX As Single, RearSightY As Single
Const Facing = piD2

Const Barrel1Len As Single = 100, Barrel2Len As Single = 80
Const Stock_Width As Single = GunLen / 10, _
      Stock_Height As Single = GunLen / 2, _
      Stock_Angle As Single = piD4
Const Main_Width As Single = GunLen, _
      Main_Height As Single = GunLen / 8
Const HandleLen = GunLen / 6, HandleStartLen = Main_Width * 2 / 3
Const Main_Height_Plus_Foresight_Offset = Main_Height + 10



MakeSquarePoints sX + Main_Height * Sine(Facing - piD2), _
                 sY - Main_Height * CoSine(Facing - piD2), _
                 Main_Width, Main_Height, Facing, pMain(), 1

MakeSquarePoints pMain(1).X, pMain(1).Y, Stock_Width, Stock_Height, Facing + Stock_Angle, pStock(), 1



BarrelStart(1).X = pMain(2).X
BarrelStart(1).Y = pMain(2).Y

BarrelStart(2).X = pMain(3).X
BarrelStart(2).Y = pMain(3).Y

BarrelEnd(1).X = BarrelStart(1).X + Barrel1Len
BarrelEnd(1).Y = BarrelStart(1).Y

BarrelEnd(2).X = BarrelStart(2).X + Barrel2Len
BarrelEnd(2).Y = BarrelStart(2).Y


Handle1X = pMain(4).X + HandleStartLen
Handle1Y = pMain(4).Y
'Handle2X = Handle1X + HandleLen * Sine(Facing)
'Handle2Y = Handle1Y - HandleLen * CoSine(Facing)

ForesightX = Handle1X + HandleLen + Main_Height_Plus_Foresight_Offset * Sine(Facing - piD2)
ForesightY = Handle1Y - Main_Height_Plus_Foresight_Offset * CoSine(Facing - piD2)

RearSightX = pMain(1).X + HandleLen * Sine(Facing - piD8)
RearSightY = pMain(1).Y - HandleLen * CoSine(Facing - piD8)
'end calculation



'drawing
picMain.ForeColor = vbBlack
picMain.DrawWidth = 2
modStickGame.sPoly pMain, vbBlack
modStickGame.sPoly pStock, vbBlack


'picMain.DrawWidth = 2 '<-- set above
modStickGame.sLine CSng(BarrelStart(1).X), CSng(BarrelStart(1).Y), CSng(BarrelEnd(1).X), CSng(BarrelEnd(1).Y)
modStickGame.sLine CSng(BarrelStart(2).X), CSng(BarrelStart(2).Y), CSng(BarrelEnd(2).X), CSng(BarrelEnd(2).Y)


picMain.FillStyle = vbFSSolid
picMain.FillColor = vbBlack
modStickGame.sCircle ForesightX, ForesightY, 20, vbBlack
modStickGame.sCircle RearSightX, RearSightY, 20, vbBlack
picMain.FillStyle = vbFSTransparent

picMain.DrawWidth = 1
End Sub

Private Sub DrawStaticAK(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 18) As Single, Y(1 To 18) As Single

Const SAd2 = SmallAngle / 2
Const SAd4 = SmallAngle / 4
Const SAd8 = SmallAngle / 8

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 4 * Sine(Facing + 11 * Pi / 18)
Y(2) = Y(1) - GunLen / 4 * CoSine(Facing + 11 * Pi / 18) '90+20deg

X(3) = X(1) + GunLen / 4 * Sine(Facing + piD2)
Y(3) = Y(1) - GunLen / 4 * CoSine(Facing + piD2)

X(4) = X(1) + GunLen / 20 * Sine(Facing)
Y(4) = Y(1) - GunLen / 20 * CoSine(Facing)

X(5) = X(1) + GunLen / 4 * Sine(Facing)
Y(5) = Y(1) - GunLen / 4 * CoSine(Facing)

X(6) = X(1) + GunLen / 3.2 * Sine(Facing - SAd2)
Y(6) = Y(1) - GunLen / 3.2 * CoSine(Facing - SAd2)

X(7) = X(6) + GunLen / 1.5 * Sine(Facing + piD4)
Y(7) = Y(6) - GunLen / 1.5 * CoSine(Facing + piD4)

X(8) = X(7) + GunLen / 4 * Sine(Facing - piD4)
Y(8) = Y(7) - GunLen / 4 * CoSine(Facing - piD4)

X(9) = X(1) + GunLen / 2 * Sine(Facing - SAd2)
Y(9) = Y(1) - GunLen / 2 * CoSine(Facing - SAd2)

X(10) = X(9) + GunLen * Sine(Facing - SAd8)
Y(10) = Y(9) - GunLen * CoSine(Facing - SAd8)

X(11) = X(10) + GunLen / 4 * Sine(Facing - piD2)
Y(11) = Y(10) - GunLen / 4 * CoSine(Facing - piD2)

X(12) = X(11) + GunLen / 4 * Sine(Facing + (piD2 + SmallAngle))
Y(12) = Y(11) - GunLen / 4 * CoSine(Facing + (piD2 + SmallAngle))

X(13) = X(12) + GunLen / 3 * Sine(Facing - Pi)
Y(13) = Y(12) - GunLen / 3 * CoSine(Facing - Pi)

X(14) = X(13) + GunLen / 3 * Sine(Facing - Pi)
Y(14) = Y(13) - GunLen / 3 * CoSine(Facing - Pi)

X(15) = X(14) + GunLen * 0.6 * Sine(Facing + (Pi - SAd4))
Y(15) = Y(14) - GunLen * 0.6 * CoSine(Facing + (Pi - SAd4))

X(16) = X(2) + GunLen / 2 * Sine(Facing - (Pi + SAd4))
Y(16) = Y(2) - GunLen / 2 * CoSine(Facing - (Pi + SAd4))

X(17) = X(16) + GunLen / 4 * Sine(Facing + (piD2 - SAd4))
Y(17) = Y(16) - GunLen / 4 * CoSine(Facing + (Pi / 2 - SAd4))

X(18) = X(1) + GunLen / 8 * Sine(Facing - Pi)
Y(18) = Y(1) - GunLen / 8 * CoSine(Facing - Pi)
'end calculation

'drawing
picMain.DrawWidth = 2
picMain.ForeColor = &H6AD5
'handle
modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(3), Y(3), X(2), Y(2)
modStickGame.sLine X(3), Y(3), X(4), Y(4)

picMain.ForeColor = vbBlack
picMain.DrawWidth = 2
'handle-mag bit
modStickGame.sLine X(5), Y(5), X(4), Y(4)
modStickGame.sLine X(5), Y(5), X(6), Y(6)

'magazine
picMain.DrawWidth = 2
modStickGame.sLine X(7), Y(7), X(6), Y(6)
modStickGame.sLine X(7), Y(7), X(8), Y(8)
modStickGame.sLine X(9), Y(9), X(8), Y(8)

'magazine top bit
modStickGame.sLine X(9), Y(9), X(6), Y(6)

'barrel
picMain.ForeColor = &H6AD5
modStickGame.sLine X(9), Y(9), X(10), Y(10)
picMain.ForeColor = vbBlack
modStickGame.sLine X(11), Y(11), X(10), Y(10) 'iron sight
modStickGame.sLine X(11), Y(11), X(12), Y(12) 'iron sight
picMain.ForeColor = &H6AD5
modStickGame.sLine X(13), Y(13), X(12), Y(12)
modStickGame.sLine X(13), Y(13), X(14), Y(14)
picMain.ForeColor = vbBlack
modStickGame.sLine X(15), Y(15), X(14), Y(14)

'stock
picMain.ForeColor = &H6AD5
modStickGame.sLine X(15), Y(15), X(16), Y(16)
modStickGame.sLine X(17), Y(17), X(16), Y(16)
modStickGame.sLine X(17), Y(17), X(18), Y(18)
picMain.ForeColor = vbBlack
modStickGame.sLine X(18), Y(18), X(1), Y(1)

End Sub

Private Sub DrawStaticXM8(sX As Single, sY As Single)

Const Facing As Single = piD2
Const tSin = 0.92387, tCos = 0.38268  'cosine(Facing - piD8)

Dim pt(1 To 17) As PointAPI
Dim PtGap(1 To 3) As PointAPI
Dim ptMag(1 To 4) As PointAPI

Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
'Dim Grip1X As Single, Grip1Y As Single, Grip2X As Single, Grip2Y As Single
Const XM8_Col = &H101010, XM8_Mag_Col = &H202020
Const Barrel_Len = GunLen / 3, GunLenD6 = GunLen / 6

pt(1).X = sX
pt(1).Y = sY

pt(2).X = pt(1).X + GunLen / 3 * Sine(Facing + pi3D4)
pt(2).Y = pt(1).Y - GunLen / 3 * CoSine(Facing + pi3D4)

pt(3).X = pt(2).X + GunLen / 6
pt(3).Y = pt(2).Y

pt(4).X = pt(1).X + GunLen / 6
pt(4).Y = pt(1).Y

pt(5).X = pt(4).X + GunLen / 6
pt(5).Y = pt(4).Y


pt(6).X = pt(5).X + GunLen / 4 * tSin
pt(6).Y = pt(5).Y - GunLen / 4 * tCos

'#######
ptMag(1) = pt(5)

ptMag(2).X = pt(5).X + GunLen / 3 * Sine(Facing + pi4D9)
ptMag(2).Y = pt(5).Y - GunLen / 3 * CoSine(Facing + pi4D9)

ptMag(3).X = ptMag(2).X + GunLen / 4 * tSin
ptMag(3).Y = ptMag(2).Y - GunLen / 4 * tCos

ptMag(4) = pt(6)
'#######

pt(7).X = pt(6).X + GunLen / 5 * tSin
pt(7).Y = pt(6).Y - GunLen / 5 * tCos

'straight bottom part of barrel
pt(8).X = pt(7).X + GunLen / 1.5
pt(8).Y = pt(7).Y

'wedge
pt(9).X = pt(8).X + GunLen / 2.8 * Sine(Facing - pi3D4)
pt(9).Y = pt(8).Y - GunLen / 2.8 * CoSine(Facing - pi3D4)


pt(10).X = pt(9).X + GunLen / 1.4 * Sine(Facing - Pi)
pt(10).Y = pt(9).Y - GunLen / 1.4 * CoSine(Facing - Pi)

pt(11).X = pt(10).X + GunLen / 6 * Sine(Facing - piD2)
pt(11).Y = pt(10).Y - GunLen / 6 * CoSine(Facing - piD2)

pt(12).X = pt(11).X + GunLen / 3 * Sine(Facing - Pi)
pt(12).Y = pt(11).Y - GunLen / 3 * CoSine(Facing - Pi)

pt(13).X = pt(12).X + GunLen / 6 * Sine(Facing + piD2)
pt(13).Y = pt(12).Y - GunLen / 6 * CoSine(Facing + piD2)

pt(14).X = pt(13).X + GunLen / 15 * Sine(Facing + piD2)
pt(14).Y = pt(13).Y - GunLen / 15 * CoSine(Facing + piD2)

'top buttstock
pt(15).X = pt(14).X + GunLen / 2 * Sine(Facing - (Pi * 1.1))
pt(15).Y = pt(14).Y - GunLen / 2 * CoSine(Facing - (Pi * 1.1))

'bottom buttstock
pt(16).X = pt(15).X + GunLen / 3 * Sine(Facing + piD2)
pt(16).Y = pt(15).Y - GunLen / 3 * CoSine(Facing + piD2)

pt(17).X = pt(16).X + GunLen / 4 * Sine(Facing - piD2)
pt(17).Y = pt(16).Y - GunLen / 4 * CoSine(Facing - piD2)


''start of fancy bits
'Pt(20) = Pt(9) + GunLen / 6 * tSin 'F-piD8
'Pt(20) = Pt(9) - GunLen / 6 * tCos
'
'Pt(21) = Pt(20) + GunLen / 2
'Pt(21) = Pt(20) - GunLen / 2 * CosFacing
'
'Pt(22) = Pt(20) + GunLen / 6 * sine(Facing - piD2)
'Pt(22) = Pt(20) - GunLen / 6 * cosine(Facing - piD2)
'
'Pt(23) = Pt(22) + GunLen / 3
'Pt(23) = Pt(22) - GunLen / 3 * CosFacing


'#############
'Hole in front of scope
PtGap(1).X = pt(14).X + GunLen / 6
PtGap(1).Y = pt(14).Y

PtGap(2).X = PtGap(1).X + GunLen / 8 * Sine(Facing + piD2)
PtGap(2).Y = PtGap(1).Y - GunLen / 8 * CoSine(Facing + piD2)

PtGap(3).X = PtGap(2).X + GunLen / 1.5 * Sine(Facing - piD20)
PtGap(3).Y = PtGap(2).Y - GunLen / 1.5 * CoSine(Facing - piD20)



'#############
'barrel
Barrel1X = pt(8).X + GunLenD6 * Sine(Facing - pi3D4)
Barrel1Y = pt(8).Y - GunLenD6 * CoSine(Facing - pi3D4)

Barrel2X = Barrel1X + Barrel_Len 'GunLen/x = BarrelLen
Barrel2Y = Barrel1Y

'############
'grip
'Grip1X = Pt(7).X + GunLen / 3
'Grip1Y = Pt(7).Y

'Grip2X = Grip1X + GunLen / 3 * Sine(Facing + piD2)
'Grip2Y = Grip1Y - GunLen / 3 * CoSine(Facing + piD2)


picMain.DrawWidth = 1
picMain.ForeColor = XM8_Col
picMain.FillColor = XM8_Col

modStickGame.sPoly pt, XM8_Col
modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y

modStickGame.sPoly PtGap, modStickGame.cg_BGColour
modStickGame.sPoly ptMag, XM8_Mag_Col

'picMain.DrawWidth = 2
'modStickGame.sLine Grip1X, Grip1Y, Grip2X, Grip2Y
'picMain.DrawWidth = 1

End Sub

Private Sub DrawStaticM82(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 32) As Single, Y(1 To 32) As Single

'calc constants
Const GLd10 = GunLen / 10
Const SAd4 = SmallAngle / 4

'Dim SinFacing As Single
'Dim CosFacing As Single
Dim SinFacingLess_kYpiD2 As Single, SinFacingLess_kYpiD4 As Single
Dim CosFacingLess_kYpiD2 As Single, CosFacingLess_kYpiD4 As Single


SinFacingLess_kYpiD2 = Sine(Facing - piD2)
CosFacingLess_kYpiD2 = CoSine(Facing - piD2)
SinFacingLess_kYpiD4 = Sine(Facing - piD4)
CosFacingLess_kYpiD4 = CoSine(Facing - piD4)
'SinFacing = sine(Facing)
'CosFacing = cosine(Facing)

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 4 * Sine(Facing - piD4)
Y(2) = Y(1) - GunLen / 4 * CoSine(Facing - piD4)

X(3) = X(2) + GunLen / 6 '* SinFacing
Y(3) = Y(2) '- GunLen / 6 * CosFacing

X(4) = X(1) + GunLen / 6 '* SinFacing
Y(4) = Y(1) '- GunLen / 6 * CosFacing

X(5) = X(4) + GunLen / 6 '* SinFacing
Y(5) = Y(4) '- GunLen / 6 * CosFacing

X(6) = X(2) + GunLen / 4 '* SinFacing
Y(6) = Y(2) '- GunLen / 4 * CosFacing

X(7) = X(6) + GunLen / 2 '* SinFacing
Y(7) = Y(6) '- GunLen / 2 * CosFacing

X(8) = X(5) + GunLen / 3 '* SinFacing
Y(8) = Y(5) '- GunLen / 3 * CosFacing

X(9) = X(8) + GunLen / 3 '* SinFacing
Y(9) = Y(8) '- GunLen / 3 * CosFacing

X(10) = X(7) + GunLen / 3 '* SinFacing
Y(10) = Y(7) '- GunLen / 3 * CosFacing

X(11) = X(10) + GunLen / 2 '* SinFacing
Y(11) = Y(10) '- GunLen / 2 * CosFacing

X(12) = X(11) + GunLen / 40 * SinFacingLess_kYpiD2
Y(12) = Y(11) - GunLen / 40 * CosFacingLess_kYpiD2

X(13) = X(12) + GunLen * 1.5 '* BarrelLen * SinFacing 'BARREL
Y(13) = Y(12) '- GunLen * 1.5 * CosFacing * BarrelLen

X(14) = X(12) + GunLen / 8 * Sine(Facing - Pi)
Y(14) = Y(12) - GunLen / 8 * CoSine(Facing - Pi)

X(15) = X(14) + GLd10 * SinFacingLess_kYpiD2
Y(15) = Y(14) - GLd10 * CosFacingLess_kYpiD2

X(16) = X(15) + GunLen / 10 * Sine(Facing - Pi) 'iron sight bottom
Y(16) = Y(15) - GunLen / 10 * CoSine(Facing - Pi)

X(17) = X(16) + GunLen / 10 * SinFacingLess_kYpiD2 'iron sight top
Y(17) = Y(16) - GunLen / 10 * CosFacingLess_kYpiD2

X(18) = X(15) + GunLen / 6 * Sine(Facing - Pi)
Y(18) = Y(15) - GunLen / 6 * CoSine(Facing - Pi)

X(19) = X(18) + GunLen / 2 * Sine(Facing - Pi) 'end of straight top bit
Y(19) = Y(18) - GunLen / 2 * CoSine(Facing - Pi)

X(20) = X(1) + GunLen / 4 * SinFacingLess_kYpiD2
Y(20) = Y(1) - GunLen / 4 * CosFacingLess_kYpiD2

'sight stand
'bottom points
X(21) = X(18) + GunLen / 8 * Sine(Facing - Pi) 'forward bottom
Y(21) = Y(18) - GunLen / 8 * CoSine(Facing - Pi)

X(22) = X(21) + GunLen / 4 * Sine(Facing - Pi) 'rearward bottom
Y(22) = Y(21) - GunLen / 4 * CoSine(Facing - Pi)
'top points
X(23) = X(21) + GunLen / 6 * SinFacingLess_kYpiD2 'forward top
Y(23) = Y(21) - GunLen / 6 * CosFacingLess_kYpiD2

X(24) = X(22) + GunLen / 6 * SinFacingLess_kYpiD2 'rearward top
Y(24) = Y(22) - GunLen / 6 * CosFacingLess_kYpiD2
'modstickgame.sLine from 21->23, 22->24

'scope
X(25) = X(24) + GunLen / 4 * Sine(Facing - Pi) 'rear bottom pt
Y(25) = Y(24) - GunLen / 4 * CoSine(Facing - Pi)

X(26) = X(24) + GunLen / 1.5 '* SinFacing 'front bottom pt
Y(26) = Y(24) '- GunLen / 1.5 * CosFacing

X(27) = X(25) + GunLen / 6 * SinFacingLess_kYpiD2 'rear top pt
Y(27) = Y(25) - GunLen / 6 * CosFacingLess_kYpiD2

X(28) = X(26) + GunLen / 8 * SinFacingLess_kYpiD2 'front top pt
Y(28) = Y(26) - GunLen / 8 * CosFacingLess_kYpiD2

'If bProne Then
    'bipod
    X(30) = X(12) + GunLen / 2 '* SinFacing 'GunLen/x = Stand's Connection
    Y(30) = Y(12) '- GunLen / 2 * CosFacing
    
    X(31) = X(30) + GunLen / 2 * Sine(Facing + Pi / 1.8) 'GunLen/x = Height of Stand
    Y(31) = Y(30) - GunLen / 2 * CoSine(Facing + Pi / 1.8)
    
    X(32) = X(31) + GunLen / 4 '* SinFacing 'GunLen/x = separation of stands
    Y(32) = Y(31) '- GunLen / 4 * CosFacing
'End If

'flash thing
X(29) = X(13) - GunLen / 6 '* SinFacing
Y(29) = Y(13) '+ GunLen / 6 * CosFacing
'end calculation

'drawing

'handle
picMain.DrawWidth = 1

'v. d. blue = &H693F3F
picMain.ForeColor = &H3F3F3F
modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(2), Y(2), X(3), Y(3)
modStickGame.sLine X(3), Y(3), X(4), Y(4)
modStickGame.sLine X(4), Y(4), X(5), Y(5)
modStickGame.sLine X(5), Y(5), X(6), Y(6)
modStickGame.sLine X(6), Y(6), X(7), Y(7)
'If Not Reloading Then
modStickGame.sLine X(7), Y(7), X(8), Y(8)
modStickGame.sLine X(8), Y(8), X(9), Y(9)
modStickGame.sLine X(9), Y(9), X(10), Y(10)
'End If
modStickGame.sLine X(10), Y(10), X(11), Y(11)
modStickGame.sLine X(11), Y(11), X(12), Y(12)

picMain.DrawWidth = 1
picMain.ForeColor = vbBlack
modStickGame.sLine X(12), Y(12), X(13), Y(13) 'BARREL

'modStickGame.sLine X(13), Y(13), X(14), Y(14)
'modStickGame.sLine X(14), Y(14), X(15), Y(15)
modStickGame.sLine X(12), Y(12), X(15), Y(15)

picMain.DrawWidth = 1
picMain.ForeColor = &H693F3F
modStickGame.sLine X(15), Y(15), X(16), Y(16)
modStickGame.sLine X(16), Y(16), X(17), Y(17)
modStickGame.sLine X(17), Y(17), X(18), Y(18)
modStickGame.sLine X(18), Y(18), X(19), Y(19)
modStickGame.sLine X(19), Y(19), X(20), Y(20)


'magazine barrier
modStickGame.sLine X(7), Y(7), X(10), Y(10)

'end of stock
modStickGame.sLine X(20), Y(20), X(1), Y(1)

'sight stand
'modstickgame.sLine from 21->23, 22->24
modStickGame.sLine X(21), Y(21), X(23), Y(23)
modStickGame.sLine X(22), Y(22), X(24), Y(24)

'scope
picMain.ForeColor = vbBlack '&H555555
picMain.DrawWidth = 2
modStickGame.sLine X(25), Y(25), X(26), Y(26)
modStickGame.sLine X(26), Y(26), X(28), Y(28)
modStickGame.sLine X(28), Y(28), X(27), Y(27)
modStickGame.sLine X(27), Y(27), X(25), Y(25)

''flash bit
'modstickgame.sLine X(29), Y(29),X(30), Y(30))
'modstickgame.sLine X(30), Y(30),X(31), Y(31))
'modstickgame.sLine X(31), Y(31),X(32), Y(32))
'modstickgame.sLine X(32), Y(32),X(29), Y(29))

'flash bit
Me.picMain.FillStyle = vbFSSolid
Me.picMain.FillColor = vbBlack
modStickGame.sCircle X(29), Y(29), 15, vbBlack 'flash thing on end of barrel
Me.picMain.FillStyle = vbFSTransparent

'bipod
picMain.DrawWidth = 1
modStickGame.sLine X(30), Y(30), X(31), Y(31)
modStickGame.sLine X(30), Y(30), X(32), Y(32)


End Sub

Private Sub DrawStaticAWM(sX As Single, sY As Single)
Const Facing As Single = piD2

Dim pMain(1 To 12) As PointAPI
Dim pSights(1 To 4) As PointAPI
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Const BarrelLen = GunLen


pMain(1).X = sX
pMain(1).Y = sY

pMain(2).X = pMain(1).X + GunLen / 3 * Sine(Facing - piD8)
pMain(2).Y = pMain(1).Y - GunLen / 3 * CoSine(Facing - piD8)

pMain(3).X = pMain(2).X + GunLen * k2D3
pMain(3).Y = pMain(2).Y

pMain(4).X = pMain(3).X + GunLen / 6 * Sine(Facing - piD2)
pMain(4).Y = pMain(3).Y - GunLen / 6 * CoSine(Facing - piD2)

pMain(5).X = pMain(4).X - GunLen * k4D3 'backwards
pMain(5).Y = pMain(4).Y

pMain(6).X = pMain(5).X + GunLen / 8 * Sine(Facing + pi5D8)  'backwards
pMain(6).Y = pMain(5).Y - GunLen / 8 * CoSine(Facing + pi5D8)

pMain(7).X = pMain(6).X + GunLen / 3 * Sine(Facing + pi17D16)  'backwards
pMain(7).Y = pMain(6).Y - GunLen / 3 * CoSine(Facing + pi17D16)

pMain(8).X = pMain(7).X - GunLen / 3 'backwards
pMain(8).Y = pMain(7).Y

pMain(9).X = pMain(8).X + GunLen / 4 * Sine(Facing + piD2)
pMain(9).Y = pMain(8).Y - GunLen / 4 * CoSine(Facing + piD2)

pMain(10).X = pMain(9).X + GunLen / 8
pMain(10).Y = pMain(9).Y

pMain(11).X = pMain(10).X + GunLen / 8 * Sine(Facing - piD2)
pMain(11).Y = pMain(10).Y - GunLen / 8 * CoSine(Facing - piD2)


'pMain(12).X = pMain(1).X + GunLen / 8 * Sine(Facing + 9 * Pi / 16)
'pMain(12).Y = pMain(1).Y - GunLen / 8 * CoSine(Facing + 9 * Pi / 16)
pMain(12).X = pMain(11).X + GunLen / 2 * Sine(Facing + piD18)
pMain(12).Y = pMain(11).Y - GunLen / 2 * CoSine(Facing + piD18)




'sights
'bottom right
pSights(1).X = pMain(4).X + GunLen / 2 * Sine(Facing - Pi)
pSights(1).Y = pMain(4).Y - GunLen / 2 * CoSine(Facing - Pi)

'top right
pSights(2).X = pSights(1).X + GunLen / 6 * Sine(Facing - piD2)  'GL/x = sight height
pSights(2).Y = pSights(1).Y - GunLen / 6 * CoSine(Facing - piD2)

'top left
pSights(3).X = pSights(2).X - GunLen / 1.6
pSights(3).Y = pSights(2).Y

'bottom left
pSights(4).X = pSights(1).X - GunLen / 2
pSights(4).Y = pSights(1).Y




Barrel1X = (pMain(4).X + pMain(3).X) / 2
Barrel1Y = (pMain(4).Y + pMain(3).Y) / 2
Barrel2X = Barrel1X + BarrelLen
Barrel2Y = Barrel1Y


picMain.DrawStyle = vbFSSolid
picMain.ForeColor = vbBlack

'barrel
picMain.DrawWidth = 1
modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y

modStickGame.sPoly_NoOutline pMain, vbBlack
modStickGame.sPoly_NoOutline pSights, vbBlack

End Sub

Private Sub DrawStaticRPG(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 16) As Single, Y(1 To 16) As Single

Const SAd2 = SmallAngle / 2

X(2) = sX
Y(2) = sY

X(1) = X(2) + GunLen / 2 * Sine(Facing - piD2)
Y(1) = Y(2) - GunLen / 2 * CoSine(Facing - piD2)

X(3) = X(1) + GunLen / 1.5 * Sine(Facing)
Y(3) = Y(1) - GunLen / 1.5 * CoSine(Facing)

X(4) = X(3) + GunLen / 2 * Sine(Facing + piD2)
Y(4) = Y(3) - GunLen / 2 * CoSine(Facing + piD2)

X(5) = X(3) + GunLen / 1.5 * Sine(Facing)
Y(5) = Y(3) - GunLen / 1.5 * CoSine(Facing)

X(6) = X(5) + GunLen / 4 * Sine(Facing - piD2)
Y(6) = Y(5) - GunLen / 4 * CoSine(Facing - piD2)

X(7) = X(6) + GunLen * 3 * Sine(Facing - Pi) 'rear top point
Y(7) = Y(6) - GunLen * 3 * CoSine(Facing - Pi)

X(8) = X(1) + GunLen * 1.7 * Sine(Facing - Pi) 'rear bottom point
Y(8) = Y(1) - GunLen * 1.7 * CoSine(Facing - Pi)

'rear funnel
X(9) = X(7) + GunLen / 3 * Sine(Facing - pi3D4) 'rear top point
Y(9) = Y(7) - GunLen / 3 * CoSine(Facing - pi3D4)

X(10) = X(8) + GunLen / 3 * Sine(Facing + pi3D4) 'rear bottom point
Y(10) = Y(8) - GunLen / 3 * CoSine(Facing + pi3D4)

'sights
X(11) = X(6) + GunLen / 1.2 * Sine(Facing - Pi)
Y(11) = Y(6) - GunLen / 1.2 * CoSine(Facing - Pi)

X(12) = X(11) + GunLen / 4 * Sine(Facing - piD2)
Y(12) = Y(11) - GunLen / 4 * CoSine(Facing - piD2)

X(13) = X(12) + GunLen / 4 * Sine(Facing - piD4)
Y(13) = Y(12) - GunLen / 4 * CoSine(Facing - piD4)

X(14) = X(13) + GunLen / 4 * Sine(Facing - piD2)
Y(14) = Y(13) - GunLen / 4 * CoSine(Facing - piD2)

X(15) = X(14) + GunLen / 2 * Sine(Facing + pi3D4)
Y(15) = Y(14) - GunLen / 2 * CoSine(Facing + pi3D4)

X(16) = X(15) + GunLen / 4 * Sine(Facing + piD2)
Y(16) = Y(15) - GunLen / 4 * CoSine(Facing + piD2)
'end calculation

'drawing
picMain.ForeColor = vbBlack
picMain.DrawWidth = 2
'handles
modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(4), Y(4), X(3), Y(3)

picMain.DrawWidth = 1
modStickGame.sLine X(1), Y(1), X(3), Y(3)
modStickGame.sLine X(3), Y(3), X(5), Y(5)
modStickGame.sLine X(6), Y(6), X(7), Y(7)

modStickGame.sLine X(1), Y(1), X(8), Y(8)
modStickGame.sLine X(7), Y(7), X(9), Y(9) 'funnel
modStickGame.sLine X(10), Y(10), X(8), Y(8)
'modstickgame.sLine X(7), Y(7),X(8), Y(8)) 'funnel connection

'sights
modStickGame.sLine X(11), Y(11), X(12), Y(12)
modStickGame.sLine X(13), Y(13), X(12), Y(12)
modStickGame.sLine X(13), Y(13), X(14), Y(14)
modStickGame.sLine X(15), Y(15), X(14), Y(14)
modStickGame.sLine X(15), Y(15), X(16), Y(16)
modStickGame.sLine X(11), Y(11), X(16), Y(16)

DrawRocket X(5) + GunLen / 1.2 * Sine(Facing - piD20), _
            Y(5) - GunLen / 1.2 * CoSine(Facing - piD20), _
            Facing ', Stick(i).Colour


End Sub

Private Sub DrawStaticM249(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 20) As Single, Y(1 To 20) As Single

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 2 * Sine(Facing + pi3D4)
Y(2) = Y(1) - GunLen / 2 * CoSine(Facing + pi3D4)

X(3) = X(2) + GunLen / 4 * Sine(Facing)
Y(3) = Y(2) - GunLen / 4 * CoSine(Facing)

X(4) = X(1) + GunLen / 4 * Sine(Facing)
Y(4) = Y(1) - GunLen / 4 * CoSine(Facing)
'end handle

'gap between handle and handy bit
X(5) = X(4) + GunLen / 4 * Sine(Facing)
Y(5) = Y(4) - GunLen / 4 * CoSine(Facing)

X(6) = X(5) + GunLen / 6 * Sine(Facing + piD2)
Y(6) = Y(5) - GunLen / 6 * CoSine(Facing + piD2)

X(7) = X(6) + GunLen / 2 * Sine(Facing)
Y(7) = Y(6) - GunLen / 2 * CoSine(Facing)

X(8) = X(5) + GunLen / 2 * Sine(Facing)
Y(8) = Y(5) - GunLen / 2 * CoSine(Facing)

'bipod
X(9) = X(2) + GunLen * 1.2 * Sine(Facing + piD10)
Y(9) = Y(2) - GunLen * 1.2 * CoSine(Facing + piD10)

X(10) = X(2) + GunLen * 1.5 * Sine(Facing + piD10)
Y(10) = Y(2) - GunLen * 1.5 * CoSine(Facing + piD10)

'barrel
X(11) = X(8) + GunLen / 1.5 * Sine(Facing)
Y(11) = Y(8) - GunLen / 1.5 * CoSine(Facing)

'sights
X(12) = X(8) + GunLen / 4 * Sine(Facing)
Y(12) = Y(8) - GunLen / 4 * CoSine(Facing)

X(13) = X(12) + GunLen / 4 * Sine(Facing - piD2)
Y(13) = Y(12) - GunLen / 4 * CoSine(Facing - piD2)

'top bit
X(14) = X(8) + GunLen / 10 * Sine(Facing - piD2)
Y(14) = Y(8) - GunLen / 10 * CoSine(Facing - piD2)

'top handle
X(15) = X(14) + GunLen / 4 * Sine(Facing - Pi)
Y(15) = Y(14) - GunLen / 4 * CoSine(Facing - Pi)

X(16) = X(15) + GunLen / 6 * Sine(Facing - piD2)
Y(16) = Y(15) - GunLen / 6 * CoSine(Facing - piD2)

X(17) = X(16) + GunLen / 4 * Sine(Facing - pi3D4)
Y(17) = Y(16) - GunLen / 4 * CoSine(Facing - pi3D4)
'end handle

X(18) = X(15) + GunLen / 4 * Sine(Facing - Pi)
Y(18) = Y(15) - GunLen / 4 * CoSine(Facing - Pi)

X(18) = X(15) + GunLen / 4 * Sine(Facing - Pi)
Y(18) = Y(15) - GunLen / 4 * CoSine(Facing - Pi)

X(19) = X(1) + GunLen / 2 * Sine(Facing - Pi)
Y(19) = Y(1) - GunLen / 2 * CoSine(Facing - Pi)

X(20) = X(19) + GunLen / 4 * Sine(Facing + piD2)
Y(20) = Y(19) - GunLen / 4 * CoSine(Facing + piD2)
'end calculation

picMain.ForeColor = vbBlack
picMain.DrawWidth = 1


modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(2), Y(2), X(3), Y(3)
modStickGame.sLine X(3), Y(3), X(4), Y(4)
modStickGame.sLine X(4), Y(4), X(5), Y(5)
modStickGame.sLine X(5), Y(5), X(6), Y(6)
modStickGame.sLine X(6), Y(6), X(7), Y(7)
modStickGame.sLine X(7), Y(7), X(8), Y(8)
modStickGame.sLine X(8), Y(8), X(9), Y(9)
modStickGame.sLine X(8), Y(8), X(10), Y(10)
modStickGame.sLine X(8), Y(8), X(11), Y(11)
modStickGame.sLine X(12), Y(12), X(13), Y(13)
modStickGame.sLine X(8), Y(8), X(14), Y(14)
modStickGame.sLine X(14), Y(14), X(15), Y(15)
modStickGame.sLine X(16), Y(16), X(15), Y(15)
modStickGame.sLine X(18), Y(18), X(15), Y(15)
modStickGame.sLine X(18), Y(18), X(19), Y(19)
modStickGame.sLine X(20), Y(20), X(19), Y(19)
modStickGame.sLine X(20), Y(20), X(1), Y(1)

picMain.DrawWidth = 2 'handle
modStickGame.sLine X(16), Y(16), X(17), Y(17)

picMain.DrawWidth = 1

End Sub

Private Sub DrawStaticUSP(sX As Single, sY As Single)

Dim Pts(1 To 10) As PointAPI
Const Facing As Single = piD2

Pts(1).X = sX
Pts(1).Y = sY

Pts(2).X = Pts(1).X + GunLen / 3 * Sine(Facing)
Pts(2).Y = Pts(1).Y - GunLen / 3 * CoSine(Facing)

Pts(3).X = Pts(2).X + GunLen / 6 * Sine(Facing - piD3) '60 deg
Pts(3).Y = Pts(2).Y - GunLen / 6 * CoSine(Facing - piD3)

Pts(4).X = Pts(3).X + GunLen / 12 * Sine(Facing - piD2)
Pts(4).Y = Pts(3).Y - GunLen / 12 * CoSine(Facing - piD2)

Pts(5).X = Pts(3).X + GunLen / 10 * Sine(Facing - Pi)
Pts(5).Y = Pts(3).Y - GunLen / 10 * CoSine(Facing - Pi)

Pts(6).X = Pts(3).X + GunLen / 1.6 * Sine(Facing - Pi)
Pts(6).Y = Pts(3).Y - GunLen / 1.6 * CoSine(Facing - Pi)

Pts(6).X = Pts(3).X + GunLen / 1.6 * Sine(Facing - Pi)
Pts(6).Y = Pts(3).Y - GunLen / 1.6 * CoSine(Facing - Pi)

Pts(7).X = Pts(6).X + GunLen / 4 * Sine(Facing + pi8D9)
Pts(7).Y = Pts(6).Y - GunLen / 4 * CoSine(Facing + pi8D9)

Pts(8).X = Pts(1).X + GunLen / 6 * Sine(Facing - Pi)
Pts(8).Y = Pts(1).Y - GunLen / 6 * CoSine(Facing - Pi)

Pts(9).X = Pts(8).X + GunLen / 3 * Sine(Facing + pi13D18)
Pts(9).Y = Pts(8).Y - GunLen / 3 * CoSine(Facing + pi13D18)

Pts(10).X = Pts(9).X + GunLen / 6 * Sine(Facing)
Pts(10).Y = Pts(9).Y - GunLen / 6 * CoSine(Facing)




picMain.ForeColor = vbBlack
picMain.DrawWidth = 2
modStickGame.sPoly Pts, vbBlack


End Sub

Private Sub DrawStaticDEagle(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 10) As Single, Y(1 To 10) As Single
Const HeadRadius2 = HeadRadius * 2 ', DEagle_Bullet_DelayD2 = DEagle_Bullet_Delay / 2


X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 2 * Sine(Facing)
Y(2) = Y(1) - GunLen / 2 * CoSine(Facing)

X(3) = X(2) + GunLen / 6 * Sine(Facing - piD3) '60 deg
Y(3) = Y(2) - GunLen / 6 * CoSine(Facing - piD3)

X(4) = X(3) + GunLen / 12 * Sine(Facing - piD2)
Y(4) = Y(3) - GunLen / 12 * CoSine(Facing - piD2)

X(5) = X(3) + GunLen / 10 * Sine(Facing - Pi)
Y(5) = Y(3) - GunLen / 10 * CoSine(Facing - Pi)

X(6) = X(3) + GunLen / 1.6 * Sine(Facing - Pi)
Y(6) = Y(3) - GunLen / 1.6 * CoSine(Facing - Pi)

X(6) = X(3) + GunLen / 1.6 * Sine(Facing - Pi)
Y(6) = Y(3) - GunLen / 1.6 * CoSine(Facing - Pi)

X(7) = X(6) + GunLen / 4 * Sine(Facing + pi8D9)
Y(7) = Y(6) - GunLen / 4 * CoSine(Facing + pi8D9)

X(8) = X(1) + GunLen / 6 * Sine(Facing - Pi)
Y(8) = Y(1) - GunLen / 6 * CoSine(Facing - Pi)

X(9) = X(8) + GunLen / 3 * Sine(Facing + pi13D18)
Y(9) = Y(8) - GunLen / 3 * CoSine(Facing + pi13D18)

X(10) = X(9) + GunLen / 6 * Sine(Facing)
Y(10) = Y(9) - GunLen / 6 * CoSine(Facing)

'end calculation
picMain.DrawWidth = 2

picMain.ForeColor = MSilver
modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(2), Y(2), X(3), Y(3)
modStickGame.sLine X(3), Y(3), X(4), Y(4)
modStickGame.sLine X(4), Y(4), X(5), Y(5)
modStickGame.sLine X(5), Y(5), X(6), Y(6)
modStickGame.sLine X(6), Y(6), X(7), Y(7)

picMain.ForeColor = vbBlack
modStickGame.sLine X(7), Y(7), X(8), Y(8)
modStickGame.sLine X(8), Y(8), X(9), Y(9)
modStickGame.sLine X(9), Y(9), X(10), Y(10)

modStickGame.sLine X(10), Y(10), X(1), Y(1)

picMain.DrawWidth = 1

End Sub

Private Sub DrawStaticFlameThrower(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim MB(1 To 10) As PointAPI
Dim FB(1 To 4) As PointAPI
'mb = MainBarrel
'fb = FuelBox

Const ArmLenDX = ArmLen / 3
Const BodyLenD2 = BodyLen / 2
Const BodyLenX2 = BodyLen * 2

MB(1).X = sX
MB(1).Y = sY

MB(2).X = MB(1).X + GunLen / 5 * Sine(Facing)
MB(2).Y = MB(1).Y - GunLen / 5 * CoSine(Facing)

MB(3).X = MB(2).X + GunLen / 3 * Sine(Facing - piD4)
MB(3).Y = MB(2).Y - GunLen / 3 * CoSine(Facing - piD4)

MB(4).X = MB(3).X + GunLen * Sine(Facing)
MB(4).Y = MB(3).Y - GunLen * CoSine(Facing)

MB(5).X = MB(4).X + GunLen / 6 * Sine(Facing - piD4)
MB(5).Y = MB(4).Y - GunLen / 6 * CoSine(Facing - piD4)

MB(6).X = MB(5).X + GunLen / 3 * Sine(Facing - piD6)
MB(6).Y = MB(5).Y - GunLen / 3 * CoSine(Facing - piD6)

MB(7).X = MB(6).X + GunLen / 10 * Sine(Facing - piD2)
MB(7).Y = MB(6).Y - GunLen / 10 * CoSine(Facing - piD2)

MB(8).X = MB(7).X + GunLen / 3 * Sine(Facing - Pi)
MB(8).Y = MB(7).Y - GunLen / 3 * CoSine(Facing - Pi)

MB(9).X = MB(8).X + GunLen / 3 * Sine(Facing + pi3D4)
MB(9).Y = MB(8).Y - GunLen / 3 * CoSine(Facing + pi3D4)

MB(10).X = MB(9).X + GunLen * Sine(Facing - Pi)
MB(10).Y = MB(9).Y - GunLen * CoSine(Facing - Pi)


FB(1).X = MB(3).X '+ GunLen / 4 * sine(Facing)
FB(1).Y = MB(3).Y '- GunLen / 4 * sine(Facing)

FB(2).X = MB(3).X + GunLen / 2 * Sine(Facing) 'glDx = boxlen
FB(2).Y = MB(3).Y - GunLen / 2 * CoSine(Facing)

FB(3).X = FB(2).X + GunLen / 3 * Sine(Facing + piD2)  'glDx = boxheight
FB(3).Y = FB(2).Y - GunLen / 3 * CoSine(Facing + piD2)

FB(4).X = FB(3).X + GunLen / 4 * Sine(Facing - Pi)
FB(4).Y = FB(3).Y - GunLen / 4 * CoSine(Facing - Pi)


picMain.ForeColor = vbBlack
picMain.DrawWidth = 2

modStickGame.sPoly MB, -1


modStickGame.sPoly FB, vbRed

End Sub

Private Sub DrawStaticMP5(sX As Single, sY As Single)

Const Facing = piD2
Dim pMain(0 To 14) As PointAPI, pMag(1 To 4) As PointAPI
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Const BarrelLen As Single = 30


pMain(0).X = sX
pMain(0).Y = sY

pMain(1).X = pMain(0).X + GunLen / 4 * Sine(Facing + pi3D4)
pMain(1).Y = pMain(0).Y - GunLen / 4 * CoSine(Facing + pi3D4)

pMain(2).X = pMain(1).X + GunLen / 6
pMain(2).Y = pMain(1).Y

pMain(3).X = pMain(2).X + GunLen / 4 * Sine(Facing - piD4)
pMain(3).Y = pMain(2).Y - GunLen / 4 * CoSine(Facing - piD4)

pMain(4).X = pMain(3).X + GunLen / 8
pMain(4).Y = pMain(3).Y

pMain(5).X = pMain(4).X + GunLen / 5 * Sine(Facing - piD8)
pMain(5).Y = pMain(4).Y - GunLen / 5 * CoSine(Facing - piD8)

pMain(6).X = pMain(5).X + GunLen / 20 * Sine(Facing - piD2)
pMain(6).Y = pMain(5).Y - GunLen / 20 * CoSine(Facing - piD2)

pMain(7).X = pMain(6).X + GunLen / 2 * Sine(Facing - piD16) 'length of main bottom bit
pMain(7).Y = pMain(6).Y - GunLen / 2 * CoSine(Facing - piD16)

Barrel1X = pMain(7).X
Barrel1Y = pMain(7).Y

Barrel2X = Barrel1X + BarrelLen
Barrel2Y = Barrel1Y

pMain(8).X = pMain(7).X + GunLen / 6 * Sine(Facing - piD2) 'top of front sight
pMain(8).Y = pMain(7).Y - GunLen / 6 * CoSine(Facing - piD2)

pMain(9).X = pMain(8).X + GunLen / 8 * Sine(Facing + pi3D4) 'GLDX must be smaller that GLDX from above
pMain(9).Y = pMain(8).Y - GunLen / 8 * CoSine(Facing + pi3D4)

pMain(10).X = pMain(9).X - GunLen / 1.2 'back bit of straight line
pMain(10).Y = pMain(9).Y

pMain(11).X = pMain(10).X + GunLen / 10 * Sine(Facing + pi3D4)
pMain(11).Y = pMain(10).Y - GunLen / 10 * CoSine(Facing + pi3D4)

pMain(12).X = pMain(11).X - GunLen / 3  'back of stock
pMain(12).Y = pMain(11).Y

pMain(13).X = pMain(12).X + GunLen / 3 * Sine(Facing + piD2)
pMain(13).Y = pMain(12).Y - GunLen / 3 * CoSine(Facing + piD2)

pMain(14).X = pMain(13).X + GunLen / 8 * Sine(Facing - piD4)
pMain(14).Y = pMain(13).Y - GunLen / 8 * CoSine(Facing - piD4)



pMag(1) = pMain(4)

pMag(2).X = pMag(1).X + GunLen / 10
pMag(2).Y = pMag(1).Y

pMag(4).X = pMag(1).X + GunLen / 3 * Sine(Facing + piD6)
pMag(4).Y = pMag(1).Y - GunLen / 3 * CoSine(Facing + piD6)

pMag(3).X = pMag(2).X + GunLen / 3 * Sine(Facing + piD5) 'front point
pMag(3).Y = pMag(2).Y - GunLen / 3 * CoSine(Facing + piD5)
'end calculation


'drawing
picMain.ForeColor = vbBlack
picMain.DrawWidth = 1
modStickGame.sPoly pMain, vbBlack
modStickGame.sPoly pMag, vbBlack
modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y

End Sub

Private Sub DrawStaticMac10(sX As Single, ByVal sY As Single)


Dim pHBar(1 To 4) As PointAPI, pVBar(1 To 4) As PointAPI, pMag(1 To 4) As PointAPI
Const Facing = piD2
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single

Const BarrelLen As Single = 100
Const VBar_Width As Single = GunLen / 10, _
      VBar_Height As Single = GunLen / 3
Const HBar_Width As Single = GunLen, _
      HBar_Height As Single = GunLen / 6
Const HBar_WidthD3 = HBar_Width / 3
Const Mag_Width = VBar_Width * 2 / 3, _
      Mag_Height = VBar_Height / 2

sY = sY - Mag_Height

MakeSquarePoints sX, sY, VBar_Width, VBar_Height, Facing, pVBar(), 1
MakeSquarePoints sX + HBar_WidthD3 * Sine(Facing - Pi), _
                 sY - HBar_WidthD3 * CoSine(Facing - Pi), _
                 HBar_Width, HBar_Height, Facing, pHBar(), 1




'If Not Reloading Then
MakeSquarePoints pVBar(4).X, pVBar(4).Y, Mag_Width, Mag_Height, Facing, pMag, 1
'End If


Barrel1X = (pHBar(2).X + pHBar(3).X) / 2
Barrel1Y = (pHBar(2).Y + pHBar(3).Y) / 2
Barrel2X = Barrel1X + BarrelLen
Barrel2Y = Barrel1Y
'end calculation


'drawing
picMain.ForeColor = vbBlack
picMain.DrawWidth = 1
modStickGame.sPoly pVBar, vbBlack
modStickGame.sPoly pHBar, vbBlack

modStickGame.sPoly pMag, vbBlack

picMain.DrawWidth = 2
modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y
'modStickGame.sCircle Hand2X, Hand2X, 80, vbBlack

picMain.DrawWidth = 1
End Sub

Private Sub DrawStaticAUG(sX As Single, sY As Single)

Const Facing As Single = piD2
Const kGreen = 32768 '32768=rgb(0,128,0)

Dim pGrip(1 To 4) As PointAPI
Dim ptBarrel(1 To 4) As PointAPI
Dim ptMain(1 To 5) As PointAPI
Dim ptMag(1 To 4) As PointAPI
Dim ptSights(1 To 4) As PointAPI
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
'Dim Stock1X As Single, Stock1Y As Single, Stock2X As Single, Stock2Y As Single
Const GrayColour As Long = &H666666

'grip
pGrip(1).X = sX
pGrip(1).Y = sY

pGrip(2).X = pGrip(1).X + GunLen / 3 * Sine(Facing + pi3D4)
pGrip(2).Y = pGrip(1).Y - GunLen / 3 * CoSine(Facing + pi3D4)

pGrip(3).X = pGrip(2).X + GunLen / 4
pGrip(3).Y = pGrip(2).Y

pGrip(4).X = pGrip(1).X + GunLen / 4
pGrip(4).Y = pGrip(1).Y
'end grip

'green barrel part
ptBarrel(1).X = pGrip(4).X
ptBarrel(1).Y = pGrip(4).Y

ptBarrel(2).X = ptBarrel(1).X + GunLen * k2D3 'GL/x = Green Len
ptBarrel(2).Y = ptBarrel(1).Y

ptBarrel(3).X = ptBarrel(2).X + GunLen / 5 * Sine(Facing - pi2d3) '100deg
ptBarrel(3).Y = ptBarrel(2).Y - GunLen / 5 * CoSine(Facing - pi2d3)

ptBarrel(4).X = ptBarrel(1).X + GunLen / 4 * Sine(Facing - piD2)
ptBarrel(4).Y = ptBarrel(1).Y - GunLen / 4 * CoSine(Facing - piD2)
'end green barrel

'black barrel
Barrel1X = (ptBarrel(2).X + ptBarrel(3).X) / 2
Barrel1Y = (ptBarrel(2).Y + ptBarrel(3).Y) / 2
Barrel2X = Barrel1X + GunLen / 8
Barrel2Y = Barrel1Y

'main black bit
ptMain(1).X = ptBarrel(4).X
ptMain(1).Y = ptBarrel(4).Y

ptMain(2).X = ptMain(1).X - GunLen 'length that it goes back (to the stock)
ptMain(2).Y = ptMain(1).Y

ptMain(3).X = ptMain(2).X + GunLen / 2 * Sine(Facing + piD2)
ptMain(3).Y = ptMain(2).Y - GunLen / 2 * CoSine(Facing + piD2)

ptMain(4).X = pGrip(1).X - GunLen / 2
ptMain(4).Y = pGrip(1).Y

ptMain(5).X = ptBarrel(1).X
ptMain(5).Y = ptBarrel(1).Y

'magazine

ptMag(1).X = pGrip(1).X - GunLen / 2.3
ptMag(1).Y = pGrip(1).Y

ptMag(2).X = ptMag(1).X - GunLen / 8 'GL/x = Mag Width
ptMag(2).Y = ptMag(1).Y + GunLen / 8

ptMag(3).X = ptMag(2).X + GunLen / 2 * Sine(Facing + piD3)
ptMag(3).Y = ptMag(2).Y - GunLen / 2 * CoSine(Facing + piD3)

ptMag(4).X = ptMag(1).X + GunLen / 2 * Sine(Facing + piD3)
ptMag(4).Y = ptMag(1).Y - GunLen / 2 * CoSine(Facing + piD3)


'sights
'bottom right
ptSights(1).X = pGrip(1).X + GunLen / 3 * Sine(Facing - piD2)
ptSights(1).Y = pGrip(1).Y - GunLen / 3 * CoSine(Facing - piD2)

'top right
ptSights(2).X = ptSights(1).X + GunLen / 6 * Sine(Facing - piD4)
ptSights(2).Y = ptSights(1).Y - GunLen / 6 * CoSine(Facing - piD4)

'top left
ptSights(3).X = ptSights(2).X - GunLen / 2
ptSights(3).Y = ptSights(2).Y

'bottom left
ptSights(4).X = ptSights(1).X - GunLen / 4
ptSights(4).Y = ptSights(1).Y




'#############
'Stock1X = CSng(ptMain(2).X)
'Stock1Y = CSng(ptMain(2).Y)
'Stock2X = CSng(ptMain(3).X)
'Stock2Y = CSng(ptMain(3).Y)
'#############

picMain.DrawWidth = 1
picMain.DrawStyle = vbFSSolid
picMain.ForeColor = vbBlack

'sight stand
modStickGame.sLine CLng(ptSights(1).X), _
                CLng(ptSights(1).Y), _
                CLng(ptSights(1).X + GunLen / 6 * Sine(Facing + piD2)), _
                CLng(ptSights(1).Y - GunLen / 6 * CoSine(Facing + piD2))
modStickGame.sLine CLng(ptSights(4).X), _
                CLng(ptSights(4).Y), _
                CLng(ptSights(4).X + GunLen / 6 * Sine(Facing + piD2)), _
                CLng(ptSights(4).Y - GunLen / 6 * CoSine(Facing + piD2))




modStickGame.sPoly pGrip, vbBlack
modStickGame.sPoly ptSights, vbBlack
modStickGame.sPoly ptMag, vbBlack

picMain.ForeColor = GrayColour
modStickGame.sPoly ptMain, GrayColour
modStickGame.sPoly ptBarrel, GrayColour

picMain.DrawWidth = 2
'barrel
modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y



End Sub
'END STATIC WEAPON DRAWING
'#########################################################################################################

Private Sub DrawDeadSticks()
Dim i As Integer

For i = 0 To NumDeadSticks - 1
    DrawDeadStick DeadStick(i).X, DeadStick(i).Y, DeadStick(i).colour, IIf(DeadStick(i).bFacingRight, -1, 1)
Next i

End Sub

Private Sub ProcessDeadSticks()
Dim i As Integer, j As Integer

picMain.DrawWidth = 2

Do While i < NumDeadSticks
    
    If DeadStick(i).Decay < GetTickCount() Then
        RemoveDeadStick i
        i = i - 1
    ElseIf DeadStick(i).Y > StickGameHeight Then
        RemoveDeadStick i
        i = i - 1
    End If
    
    i = i + 1
Loop


For i = 0 To NumDeadSticks - 1
    If DeadStick(i).bOnSurface = False Then
        
        
        DeadStick(i).Heading = FixAngle(DeadStick(i).Heading)
        
        If DeadStick(i).X < 1 Then
            If DeadStick(i).Heading > Pi Then
                ReverseXComp DeadStick(i).Speed, DeadStick(i).Heading
                DeadStick(i).X = 2
                DeadStick(i).Speed = DeadStick(i).Speed / 2
            End If
        ElseIf DeadStick(i).X > (StickGameWidth - 1) Then
            If DeadStick(i).Heading < Pi Then
                ReverseXComp DeadStick(i).Speed, DeadStick(i).Heading
                DeadStick(i).X = StickGameWidth
                DeadStick(i).Speed = DeadStick(i).Speed / 2
            End If
        ElseIf DeadStick(i).Y < 1 Then
            ReverseYComp DeadStick(i).Speed, DeadStick(i).Heading
            AddBloodExplosion DeadStick(i).X, DeadStick(i).Y
            DeadStick(i).Y = 1
            DeadStick(i).Speed = DeadStick(i).Speed / 2
        End If
        
        
        ApplyGravityVector DeadStick(i).LastGravity, GetTimeZoneAdjust(DeadStick(i).X, DeadStick(i).Y), _
            DeadStick(i).Speed, DeadStick(i).Heading, DeadStick(i).X, DeadStick(i).Y
        
'        If DeadStick(i).LastGravity + Gravity_Delay / GetTimeZoneAdjust(DeadStick(i).X, DeadStick(i).Y) < GetTickCount() Then
'            AddVectors DeadStick(i).Speed, DeadStick(i).Heading, Gravity_Strength, Gravity_Direction, _
'                DeadStick(i).Speed, DeadStick(i).Heading
'
'            DeadStick(i).LastGravity = GetTickCount()
'        End If
        
        MotionStickObject DeadStick(i).X, DeadStick(i).Y, DeadStick(i).Speed, DeadStick(i).Heading
        
        For j = 0 To ubdPlatforms
            DeadStickOnSurface i, j
        Next j
        
    End If
Next i

End Sub

Private Sub DrawDeadStick(X As Single, Y As Single, Col As Long, kY As Single)
Const BodyLenD2 = BodyLen / 2
Const BodyLenPlus = BodyLen * 1.2
Dim YpHR As Single, XmHR As Single

YpHR = Y + HeadRadius
XmHR = X - kY * HeadRadius

picMain.FillStyle = vbFSSolid
picMain.FillColor = Col
modStickGame.sCircle X, Y, HeadRadius, Col

picMain.FillStyle = vbFSSolid
picMain.FillColor = vbRed
modStickGame.sCircleAspect XmHR, YpHR, BodyLenD2, vbRed, 0.2
picMain.FillStyle = vbFSTransparent

picMain.ForeColor = Col
modStickGame.sLine XmHR, Y, X - kY * BodyLen, YpHR
modStickGame.sLine XmHR, Y, X - kY * BodyLenD2, YpHR
modStickGame.sLine X - kY * BodyLen, YpHR, X - kY * BodyLenPlus, YpHR

End Sub

Private Sub DeadStickOnSurface(i As Integer, iPlatform As Integer) 'As Boolean
Const kAmount = HeadRadius * 1.6
Dim j As Integer
'Dim rcStick As RECT
'
'With rcStick
'    .Left = DeadStick(i).X
'    .Right = .Left + 1
'    .Top = DeadStick(i).Y
'    .Bottom = .Top + 1
'End With

If RectCollision(PointToRect(DeadStick(i).X, DeadStick(i).Y), PlatformToRect(Platform(iPlatform))) Then
    'position the DeadStick on top of the platform
    'If DeadStick(i).y > (Platform(iPlatform).Top + 5) Then
        '                       add on a bit
    DeadStick(i).Y = Platform(iPlatform).Top - kAmount
        
    'End If
    
    'DeadStickOnSurface = True
    DeadStick(i).bOnSurface = True
    DeadStick(i).Speed = 0
    
    If DeadStick(i).bFlamed Then
        For j = 0 To 10
            AddSmokeNadeTrail DeadStick(i).X + Rnd() * ArmLen, DeadStick(i).Y + Rnd() * ArmLen, True, True
        Next j
    Else
        For j = 1 To 15
            'splatter
            AddBlood DeadStick(i).X, DeadStick(i).Y, PM_Rnd * piD2
        Next j
    End If
End If

'If DeadStick(i).X > Platform(iPlatform).Left Then
'    If DeadStick(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
'
'        If DeadStick(i).Y > Platform(iPlatform).Top Then
'            If DeadStick(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
'
'                'position the DeadStick on top of the platform
'                'If DeadStick(i).y > (Platform(iPlatform).Top + 5) Then
'                    '                       add on a bit
'                DeadStick(i).Y = Platform(iPlatform).Top - kAmount
'
'                'End If
'
'                'DeadStickOnSurface = True
'                DeadStick(i).bOnSurface = True
'                DeadStick(i).Speed = 0
'
'                If DeadStick(i).bFlamed Then
'                    For J = 0 To 10
'                        AddSmokeNadeTrail DeadStick(i).X + Rnd() * ArmLen, DeadStick(i).Y + Rnd() * ArmLen, True
'                    Next J
'                Else
'                    For J = 1 To 15
'                        'splatter!
'                        AddBlood DeadStick(i).X, DeadStick(i).Y, PM_Rnd * piD2, False
'                    Next J
'                End If
'
'
'            End If
'        End If
'
'
'    End If
'End If

End Sub

Private Sub DrawDeadChoppers()
Dim i As Integer

For i = 0 To NumDeadChoppers - 1
    DrawDeadChopper DeadChopper(i).X, DeadChopper(i).Y, DeadChopper(i).colour
Next i

End Sub

Private Sub ProcessDeadChoppers()
Dim i As Integer, j As Integer
Dim Adj As Single

picMain.DrawWidth = 2

Do While i < NumDeadChoppers
    
    If DeadChopper(i).Decay < GetTickCount() Then
        RemoveDeadChopper i
        i = i - 1
    ElseIf DeadChopper(i).Y > StickGameHeight Then
        RemoveDeadChopper i
        i = i - 1
    End If
    
    i = i + 1
Loop


For i = 0 To NumDeadChoppers - 1
    
    If DeadChopper(i).bOnSurface = False Then
        
        Adj = GetTimeZoneAdjust(DeadChopper(i).X, DeadChopper(i).Y)
        
'        If DeadChopper(i).LastGravity + Gravity_Delay / Adj < GetTickCount() Then
'
'            AddVectors DeadChopper(i).Speed, DeadChopper(i).Heading, Gravity_Strength, Gravity_Direction, _
'                DeadChopper(i).Speed, DeadChopper(i).Heading
'
'            DeadChopper(i).LastGravity = GetTickCount()
'
'            If DeadChopper(i).LastSmoke + DeadChopper_Smoke_Delay / Adj < GetTickCount() Then
'                AddSmokeNadeTrail DeadChopper(i).X - CLD3 * Rnd(), DeadChopper(i).Y, , True
'                AddExplosion DeadChopper(i).X - CLD3 * Rnd(), DeadChopper(i).Y + CLD10, 750, 0.25, DeadChopper(i).Speed / 3, DeadChopper(i).Heading
'            End If
'        End If
        
        ApplyGravityVector DeadChopper(i).LastGravity, Adj, DeadChopper(i).Speed, DeadChopper(i).Heading, DeadChopper(i).X, DeadChopper(i).Y
        
        If DeadChopper(i).LastSmoke + DeadChopper_Smoke_Delay / Adj < GetTickCount() Then
            AddSmokeNadeTrail DeadChopper(i).X - CLD3 * Rnd(), DeadChopper(i).Y, , True
            AddExplosion DeadChopper(i).X - CLD3 * Rnd(), DeadChopper(i).Y + CLD10, 200
        End If
        
        
        MotionStickObject DeadChopper(i).X, DeadChopper(i).Y, DeadChopper(i).Speed, DeadChopper(i).Heading
        
        For j = 0 To ubdPlatforms
            DeadChopperOnSurface i, j
        Next j
        
    End If
Next i

End Sub

Private Sub DrawDeadChopper(X As Single, Y As Single, Col As Long)
Dim pt(1 To 11) As PointAPI, ScreenPt(1 To 3) As PointAPI
Dim t1X As Single, t1Y As Single, t2X As Single, t2Y As Single
Const Facing = piD2

pt(1).X = X
pt(1).Y = Y

pt(2).X = pt(1).X + CLD6 * Sine(Facing + piD6)
pt(2).Y = pt(1).Y - CLD6 * CoSine(Facing + piD6)

pt(3).X = pt(2).X + CLD10 * Sine(Facing + piD3)
pt(3).Y = pt(2).Y - CLD10 * CoSine(Facing + piD3)

pt(4).X = pt(3).X + CLD2 * Sine(Facing - Pi)
pt(4).Y = pt(3).Y - CLD2 * CoSine(Facing - Pi)

pt(5).X = pt(4).X + CLD10 * Sine(Facing - pi3D4)
pt(5).Y = pt(4).Y - CLD10 * CoSine(Facing - pi3D4)

pt(6).X = pt(5).X + CLD3 * Sine(Facing - Pi)
pt(6).Y = pt(5).Y - CLD3 * CoSine(Facing - Pi)

pt(7).X = pt(6).X + CLD6 * Sine(Facing - pi3D4)
pt(7).Y = pt(6).Y - CLD6 * CoSine(Facing - pi3D4)

pt(8).X = pt(7).X + CLD8 * Sine(Facing)
pt(8).Y = pt(7).Y - CLD8 * CoSine(Facing)

pt(9).X = pt(8).X + CLD8 * Sine(Facing + piD4)
pt(9).Y = pt(8).Y - CLD8 * CoSine(Facing + piD4)

pt(10).X = pt(9).X + CLD6 * Sine(Facing)
pt(10).Y = pt(9).Y - CLD6 * CoSine(Facing)

pt(11).X = pt(1).X + CLD8 * Sine(Facing - Pi)
pt(11).Y = pt(1).Y - CLD8 * CoSine(Facing - Pi)


ScreenPt(1).X = pt(1).X + Sine(Facing + piD2) * 50
ScreenPt(1).Y = pt(1).Y - CoSine(Facing + piD2) * 50

ScreenPt(2).X = pt(2).X + Sine(Facing - Pi) * 50
ScreenPt(2).Y = pt(2).Y - CoSine(Facing - Pi) * 50

ScreenPt(3).X = ScreenPt(2).X - CLD6 * Sine(Facing)
ScreenPt(3).Y = ScreenPt(2).Y + CLD6 * CoSine(Facing)


t1X = CSng(pt(3).X)
t1Y = CSng(pt(3).Y)
t2X = CSng(pt(8).X)
t2Y = CSng(pt(8).Y)


picMain.DrawStyle = 5
modStickGame.sPoly pt, MSilver
modStickGame.sPoly ScreenPt, Col
picMain.DrawStyle = 0
picMain.DrawWidth = 2

picMain.ForeColor = vbBlack
modStickGame.sLine t1X, t1Y, t2X, t2Y

End Sub

Private Sub DeadChopperOnSurface(i As Integer, iPlatform As Integer) 'As Boolean
'Const kAmount = CLD6, kAmountDX = kAmount / 1.2
Const kAmount = CLD6 / 1.2

Dim j As Integer
Dim rcChopper As RECT

With rcChopper
    .Left = DeadChopper(i).X
    .Right = .Left + 1
    .Top = DeadChopper(i).Y
    .Bottom = .Top + kAmount
End With
 
If RectCollision(rcChopper, PlatformToRect(Platform(iPlatform))) Then
    'position the DeadChopper on top of the platform
    'If DeadChopper(i).y > (Platform(iPlatform).Top + 5) Then
        '                       add on a bit
    DeadChopper(i).Y = Platform(iPlatform).Top - CLD6
        
    'End If
    
    'DeadChopperOnSurface = True
    DeadChopper(i).bOnSurface = True
    DeadChopper(i).Speed = 0
    
    
    For j = 0 To 5
        AddExplosion DeadChopper(i).X + CLD6 - Rnd() * CLD2, DeadChopper(i).Y + Rnd() * CLD6, 400
    Next j
    
    For j = 0 To IIf(modStickGame.cg_Smoke, 20, 10)
        AddSmokeNadeTrail DeadChopper(i).X + CLD6 - Rnd() * CLD2, DeadChopper(i).Y + Rnd() * CLD6, True, True
    Next j
    
    For j = 0 To 5 + Rnd() * 3
        AddNadeTrail_Simple DeadChopper(i).X - Rnd() * CLD6, DeadChopper(i).Y
    Next j
    
    Call CheckDeadChopperStickCollisions(i)
End If


'If DeadChopper(i).X > Platform(iPlatform).Left Then
'    If DeadChopper(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
'
'        If DeadChopper(i).Y + kAmountDX > Platform(iPlatform).Top Then
'            If DeadChopper(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
'
'                'position the DeadChopper on top of the platform
'                'If DeadChopper(i).y > (Platform(iPlatform).Top + 5) Then
'                    '                       add on a bit
'                DeadChopper(i).Y = Platform(iPlatform).Top - kAmount
'
'                'End If
'
'                'DeadChopperOnSurface = True
'                DeadChopper(i).bOnSurface = True
'                DeadChopper(i).Speed = 0
'
'
'                For J = 0 To 5
'                    AddExplosion DeadChopper(i).X + CLD6 - Rnd() * CLD2, DeadChopper(i).Y + Rnd() * CLD6, 400, 0.25, 0, 0
'                Next J
'
'                For J = 0 To 20
'                    AddSmokeNadeTrail DeadChopper(i).X + CLD6 - Rnd() * CLD2, DeadChopper(i).Y + Rnd() * CLD6, True
'                Next J
'
'                Call CheckDeadChopperStickCollisions(i)
'
'            End If
'        End If
'
'
'    End If
'End If

End Sub

Private Sub CheckDeadChopperStickCollisions(iChopper As Integer)
Dim i As Integer

For i = 0 To NumSticksM1
    If StickInGame(i) Then
        If Stick(i).WeaponType <> Chopper Then
            If Stick(i).X > DeadChopper(iChopper).X - CLD2 Then
                If Stick(i).X < DeadChopper(iChopper).X + CLD3 Then
                    
                    If Stick(i).Y > DeadChopper(iChopper).Y Then
                        If Stick(i).Y < DeadChopper(iChopper).Y + CLD3 Then
                            
                            AddBloodExplosion Stick(i).X, Stick(i).Y
                            
                            If i = 0 Or Stick(i).IsBot Then
                                If StickInvul(i) = False Then
                                    Stick(i).Shield = 0
                                    Call Killed(i, DeadChopper(iChopper).iOwner, kCrushed)
                                End If
                            End If
                            
                        End If
                    End If
                    
                    
                End If
            End If
        End If
    End If
Next i

End Sub

Private Sub AddBloodExplosion(X As Single, Y As Single)
Dim i As Integer
For i = 1 To 30
    'splatter!
    AddBlood X, Y, Rnd() * Pi2
Next i
End Sub

'explosions
Private Sub AddExplosion(X As Single, Y As Single, MaxRadius As Single)
Const expand_Speed = 70

AddCirc X, Y, MaxRadius / 1.7, 1, vbYellow, expand_Speed * 0.6, True
AddCirc X, Y, MaxRadius / 1.3, 1, &HDA6F0, expand_Speed * 0.8, True
AddCirc X, Y, MaxRadius, 1, vbRed, expand_Speed, True

'894704 = orange = &H0DA6F0

End Sub

Private Sub AddCirc(X As Single, Y As Single, MaxRadius As Single, StartRadius As Single, _
    colour As Long, ExpandSpeed As Single, bTimeZoneable As Boolean)


If bTimeZoneable Then
    
    ReDim Preserve TimeZoneCircs(NumTimeZoneCircs)
    
    With TimeZoneCircs(NumTimeZoneCircs)
        .X = X
        .Y = Y
        
        .colour = colour
        .sgDirection = 1
        .ExpandSpeed = ExpandSpeed
        
        .currentRadius = StartRadius
        .MaxRadius = MaxRadius
    End With
    
    NumTimeZoneCircs = NumTimeZoneCircs + 1
    
Else
    
    ReDim Preserve ScreenCircs(NumScreenCircs)
    
    With ScreenCircs(NumScreenCircs)
        .X = X
        .Y = Y
        
        .colour = colour
        .sgDirection = 1
        .ExpandSpeed = ExpandSpeed
        
        .currentRadius = StartRadius
        .MaxRadius = MaxRadius
    End With
    
    NumScreenCircs = NumScreenCircs + 1
    
End If

End Sub

Private Sub RemoveCirc(Index As Integer, bIsTimeZoneCirc As Boolean)

Dim i As Integer

If bIsTimeZoneCirc Then
    If NumTimeZoneCircs = 1 Then
        Erase TimeZoneCircs
        NumTimeZoneCircs = 0
    Else
        For i = Index To NumTimeZoneCircs - 2
            TimeZoneCircs(i) = TimeZoneCircs(i + 1)
        Next i
        
        'Resize the array
        ReDim Preserve TimeZoneCircs(NumTimeZoneCircs - 2)
        NumTimeZoneCircs = NumTimeZoneCircs - 1
    End If
Else
    If NumScreenCircs = 1 Then
        Erase ScreenCircs
        NumScreenCircs = 0
    Else
        For i = Index To NumScreenCircs - 2
            ScreenCircs(i) = ScreenCircs(i + 1)
        Next i
        
        'Resize the array
        ReDim Preserve ScreenCircs(NumScreenCircs - 2)
        NumScreenCircs = NumScreenCircs - 1
    End If
End If


End Sub

Private Sub DrawTimeZoneCircs()
Dim i As Integer

picMain.FillStyle = vbFSSolid

For i = NumTimeZoneCircs - 1 To 0 Step -1
    picMain.FillColor = TimeZoneCircs(i).colour
    modStickGame.sCircle TimeZoneCircs(i).X, TimeZoneCircs(i).Y, TimeZoneCircs(i).currentRadius, TimeZoneCircs(i).colour
Next i

picMain.FillStyle = vbFSTransparent

End Sub
Private Sub DrawScreenCircs()
Dim i As Integer

picMain.FillStyle = vbFSSolid

For i = NumScreenCircs - 1 To 0 Step -1
    picMain.FillColor = ScreenCircs(i).colour
    modStickGame.sCircle ScreenCircs(i).X, ScreenCircs(i).Y, ScreenCircs(i).currentRadius, ScreenCircs(i).colour
Next i

picMain.FillStyle = vbFSTransparent

End Sub

Private Sub ProcessAllCircs()
Dim i As Integer

Do While i < NumScreenCircs
    
    'Circs(i).Prog = Circs(i).Prog + Circs(i).Direction * 100 * _
        modStickGame.StickTimeFactor * GetTimeZoneAdjust(Circs(i).X, Circs(i).Y)
    
    'If Circs(i).Prog > Circs(i).MaxProg Then
        'Circs(i).Direction = -1
        'Circs(i).Prog = Circs(i).MaxProg
    'ElseIf Circs(i).Prog <= 0 Then
        'RemoveCirc i
        'i = i - 1
    'End If
    
    ScreenCircs(i).currentRadius = ScreenCircs(i).currentRadius + ScreenCircs(i).sgDirection * ScreenCircs(i).ExpandSpeed * modStickGame.StickTimeFactor
    
    If ScreenCircs(i).sgDirection = 1 Then
        If ScreenCircs(i).currentRadius > ScreenCircs(i).MaxRadius Then
            ScreenCircs(i).currentRadius = ScreenCircs(i).MaxRadius
            ScreenCircs(i).sgDirection = -1
        End If
        
    ElseIf ScreenCircs(i).currentRadius <= 0 Then
        RemoveCirc i, False
        i = i - 1
    End If
    
    
    i = i + 1
Loop


i = 0
Do While i < NumTimeZoneCircs
    
    TimeZoneCircs(i).currentRadius = TimeZoneCircs(i).currentRadius + TimeZoneCircs(i).sgDirection * TimeZoneCircs(i).ExpandSpeed * modStickGame.StickTimeFactor _
        * GetTimeZoneAdjust(TimeZoneCircs(i).X, TimeZoneCircs(i).Y)
    
    
    If TimeZoneCircs(i).sgDirection = 1 Then
        If TimeZoneCircs(i).currentRadius > TimeZoneCircs(i).MaxRadius Then
            TimeZoneCircs(i).currentRadius = TimeZoneCircs(i).MaxRadius
            TimeZoneCircs(i).sgDirection = -1
        End If
        
    ElseIf TimeZoneCircs(i).currentRadius <= 0 Then
        RemoveCirc i, True
        i = i - 1
    End If
    
    
    i = i + 1
Loop

End Sub

Public Sub HideCursor(bHide As Boolean)

If bHide Then
    Me.MousePointer = vbCustom
    Me.MouseIcon = picBlank.Picture
Else
    Me.MousePointer = vbDefault
End If

End Sub

Private Sub tmrMain_Timer()

Const Cap = "Stick Shooter - "
Dim bSuccess As Boolean

tmrMain.Enabled = False

'Connect winsock
If StartWinsock() Then
    
    
    'If we're not the StickServer, try to connect
    If Not StickServer Then
        Me.Caption = Cap & "Client"
        
        bSuccess = False
        
        If RequestMap() Then
            If ConnectToServer() Then
                bSuccess = True
            End If
        End If
        
        If Not bSuccess Then
            modWinsock.DestroySocket lSocket
            Unload Me
            Exit Sub
        End If
        
    Else
        'lSocket already bound
        Me.Caption = Cap & "Host"
        
        'tell everyone
        SendInfoMessage frmMain.LastName & " Started a Game - Ctrl+G to Join (Port " & modPorts.StickPort & ")"
'        If Server Then
'            DistributeMsg eCommands.Info & frmMain.LastName & " Started a Game - Ctrl+G to Join0", -1
'        Else
'            SendData eCommands.Info & frmMain.LastName & " Started a Game - Ctrl+G to Join0"
'        End If
        Pause 100 'let the above message be sent
    End If
    
    
    If Not StickServer Then
        modWinsock.SendPacket lSocket, ServerSockAddr, sChats & Trim$(Stick(0).Name) & _
            " joined.#" & modVars.TxtForeGround
    End If
    
    'Me.MousePointer = vbCustom
    'Me.MouseIcon = picBlank.Picture
    
    modAudio.PlayNewRoundSound
    
    HideCursor True
    AddMainMessage "Press Tab for Settings", False, vbBlack
    CheckCapsLock 'add mainmessage if it's on
    
    Call MainLoop
    
    EndWinsock
'else
    'error text already added
End If

Unload Me

End Sub

'######################################################
'camera stuff
'######################################################

Public Sub CentreCameraOnPoint(pX As Single, pY As Single)

MoveCameraX pX * cg_sZoom - StickCentreX, False
MoveCameraY pY * cg_sZoom - StickCentreY, False

End Sub

Public Sub MoveCameraX(ByVal nX As Single, Optional bCheck As Boolean = True)
Const CameraLim As Single = 150

#If Clip_X_Camera = False Then
If (StickInGame(0) = False Or bPlaying = False) And bCheck Then
#End If
    If modStickGame.cg_sCamera.X < -CameraLim Then
        If nX > modStickGame.cg_sCamera.X Then
            modStickGame.cg_sCamera.X = nX
        End If
    ElseIf (modStickGame.cg_sCamera.X + Me.width) / cg_sZoom > (StickGameWidth + CameraLim) Then
        If nX < modStickGame.cg_sCamera.X Then
            modStickGame.cg_sCamera.X = nX
        End If
        'modStickGame.cg_sCamera.X = (StickGameHeight + CameraLim) / cg_sZoom - Me.height
    Else
        modStickGame.cg_sCamera.X = nX
    End If
#If Clip_X_Camera = False Then
Else
    modStickGame.cg_sCamera.X = nX
End If
#End If

'If StickInGame(0) = False Or bPlaying = False Then
'    If modStickGame.cg_sCamera.X < -CameraLim Then
'        If nX > modStickGame.cg_sCamera.X Then
'            modStickGame.cg_sCamera.X = nX
'        End If
'    ElseIf (modStickGame.cg_sCamera.X + Me.width) * cg_sZoom > (StickGameWidth + CameraLim) Then
'        modStickGame.cg_sCamera.X = (StickGameWidth + CameraLim) / cg_sZoom - Me.width
'    Else
'        modStickGame.cg_sCamera.X = nX
'    End If
'Else
'    modStickGame.cg_sCamera.X = nX
'End If

End Sub
Public Sub MoveCameraY(ByVal nY As Single, Optional bCheck As Boolean = True)
Const CameraLim = 150

#If Clip_Y_Camera = False Then
If (StickInGame(0) = False Or bPlaying = False) And bCheck Then
#End If
    If modStickGame.cg_sCamera.Y < -CameraLim Then
        If nY > modStickGame.cg_sCamera.Y Then
            modStickGame.cg_sCamera.Y = nY
        End If
    ElseIf (modStickGame.cg_sCamera.Y + Me.height - 700) / cg_sZoom > (StickGameHeight + CameraLim) Then
        If nY < modStickGame.cg_sCamera.Y Then '-700 allows the camera to go down
            modStickGame.cg_sCamera.Y = nY
        End If
        'modStickGame.cg_sCamera.Y = (StickGameHeight + CameraLim) / cg_sZoom - Me.height
    Else
        modStickGame.cg_sCamera.Y = nY
    End If
#If Clip_Y_Camera = False Then
Else
    modStickGame.cg_sCamera.Y = nY
End If
#End If

End Sub

Private Sub ResetKeys()
'LeftKey = False
'RightKey = False
'JumpKey = False
'CrouchKey = False
'ProneKey = False
'ReloadKey = False
FireKey = False
UseKey = False
'MineKey = False

ShowScoresKey = False

SpecUp = False: SpecDown = False: SpecLeft = False: SpecRight = False

WeaponKey = -1
'Scroll_WeaponKey = 0 '-1
'LastScrollWeaponSwitch = 0
End Sub

'#####################################################################################
'Round Stuff
'#####################################################################################

Private Sub SendRoundInfo(Optional bForce As Boolean = False)
Static LastSend As Long


If LastSend + RoundInfoSendDelay < GetTickCount() Or bForce Then
    
    SendBroadcast sRoundInfos & CStr(Abs(bPlaying)) & CStr(RoundWinnerID) '& _
        vbSpace & MakeSquareNumber()
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub ReceivedRoundInfo(sPacket As String)
Dim bWasPlaying As Boolean

'format: (bPlaying)(WinnerID) '(space)(square)

On Error GoTo EH

bWasPlaying = bPlaying
bPlaying = CBool(Left$(sPacket, 1))
HideCursor bPlaying
RoundWinnerID = CInt(Mid$(sPacket, 2))

If Not bPlaying Then
    If bWasPlaying Then
        StopPlay True
    End If
Else
    If Not bWasPlaying Then
        StopPlay False
    End If
End If

EH:
End Sub

Private Sub StopPlay(ByVal bStop As Boolean)
Dim i As Integer

bPlaying = Not bStop
HideCursor bPlaying
ResetKeys

If bStop Then
    
    modAudio.StopWeaponReloadSound Stick(0).WeaponType
    
    If StickServer Then
        SendRoundInfo True
    End If
    
    RoundPausedAtThisTime = GetTickCount()
    For i = 0 To NumBullets - 1
        RemoveBullet 0, False
    Next i
    
    Erase ShieldWave: NumShieldWaves = 0
    
    
    For i = 0 To NumSticksM1
        Stick(i).state = STICK_NONE
        ResetStickFireAndFlash i
    Next i
    
    For i = 0 To UBound(AmmoFired)
        AmmoFired(i) = 0
    Next i
    
    'clip camera
    'modStickGame.cg_sZoom = 1
    'MoveCameraX modStickGame.cg_sCamera.X
    'MoveCameraY modStickGame.cg_sCamera.Y
    'CentreCameraOnPoint CSng(modStickGame.cg_sCamera.X), CSng(modStickGame.cg_sCamera.Y)
    
Else
    
    If StickServer Then
        SendRoundInfo True
    End If
    
    CrouchKey = False
    UseKey = False
    FireKey = False
    
    
    'reset all scores
    For i = 0 To NumSticksM1
        With Stick(i)
            .iKills = 0
            .iDeaths = 0
            .state = STICK_NONE
            .iKillsInARow = 0
            .BulletsFired = 0
            
            '.JumpStartY = StickGameHeight + 100
            'ResetJumpStartY i
            ResetStickFireAndFlash i
            
            .BulletsFired2 = 0
            
            .LastMine = 0
            .LastBullet = 0
            
            ResetStickFireAndFlash i
            
            .bOnSurface = False
            .bTouchedSurface = False
            
            .Speed = 0
            
            If .IsBot = False Then
                If .WeaponType = Chopper Then
                    '.WeaponType = .CurrentWeapons(1)
                    SetSticksWeapon i, .CurrentWeapons(1)
                End If
            End If
            If .Perk = pZombie Then
                .Health = Zombie_Health
                .Shield = 0
            ElseIf .WeaponType = Chopper Then
                'bots only, should be
                .Health = Health_Start
                .Shield = 0
            Else
                .Health = Health_Start
                .Shield = 1 'IIf(modStickGame.sv_SpawnWithShields, Max_Shield, 0)
                'always start a new round with a shield
            End If
            
            
            .LastNade = GetTickCount() - 10000
            
            
            .bLightSaber = False
            
            .bAlive = True
            
            .LastSpawnTime = GetTickCount()
            
            If modStickGame.sv_GameType = gCoOp Then
                
                If .Team = Red Then
                    If .WeaponType = Chopper Then
                        .X = StickGameWidth - 1000
                        .Y = 3000
                    Else
                        RandomizeCoOpBot i
                    End If
                Else 'group blue+neutral -> friendly fire etc
                    MoveStickToCoOpStart i
                End If
                
            ElseIf .IsBot Then
                .X = Rnd() * StickGameWidth
                .Y = Rnd() * StickGameHeight
                
            End If
            
            
            
        End With
    Next i
    
    i = 0
    Do While i < NumChat
        If Chat(i).bChatMessage = False Then
            RemoveChatText i
            i = i - 1
        End If
        i = i + 1
    Loop
    
    
    If modStickGame.sv_GameType <> gCoOp Then
        RandomizeMyStickPos
    'Else
        'RadarStartTime = GetTickCount()
    End If
    
    
    'reset private stuff
    'RadarStartTime = 0
    'bHadRadar = False
    'If modStickGame.sv_GameType <> gElimination Or Stick(0).WeaponType = Chopper Then
        ChopperAvail = False
    'End If
    FlamesInARow = 0
    KnifesInARow = 0
    'NadesShot = 0
    
    'erase stuff
    StickGameSpeedChanged -1, -1
    Erase Nade: NumNades = 0
    Erase Mine: NumMines = 0
    'Erase DeadChopper: NumDeadChoppers = 0
    'Erase WallMark: NumWallMarks = 0
    Erase BulletTrail: NumBulletTrails = 0
    Erase MainMessages: NumMainMessages = 0
    EraseWallMarks
    EraseDeadSticks
    
    'Erase LargeSmoke: NumLargeSmokes = 0
    Erase TimeZone: NumTimeZones = 0
    Erase GravityZone: NumGravityZones = 0
    Erase Casing: NumCasings = 0
    
    For i = 0 To modStickGame.ubdBoxes
        Box(i).bInUse = True
    Next i
    
    ResetKeys
    
    LastScoreCheck = GetTickCount() + 10000 'not too sure what this is doing here...
    
    If Stick(0).Team <> Spec Then cg_sZoom = 1
    
    
    FillTotalMags
    
    Erase Barrel: NumBarrels = 0
    AddBarrels
    SendBarrelRefresh True
    
    
    modAudio.PlayNewRoundSound
End If

End Sub

Public Sub RandomizeCoOpBot(i As Integer)
Const adjusted_Left_Indent As Long = Left_Indent * 2

Stick(i).X = Rnd() * (StickGameWidth + Left_Indent)

If Stick(i).X > StickGameWidth Then
    Stick(i).X = StickGameWidth - adjusted_Left_Indent * Rnd()
ElseIf Stick(i).X < adjusted_Left_Indent Then
    Stick(i).X = adjusted_Left_Indent + Rnd() * adjusted_Left_Indent
End If

Stick(i).Y = 200

End Sub

Public Sub MoveStickToCoOpStart(i As Integer)

Stick(i).X = Rnd() * 9000
Stick(i).Y = StickGameHeight - 1500
Stick(i).iCurrentPlatform = -1
Stick(i).Speed = 0
'ResetJumpStartY i

End Sub


Private Sub ProcessEndRound()
Dim RoundTm As Long
Dim RoundWinneri As Integer
Dim Str As String
Static LastPresenceSend As Long

On Error Resume Next

ShowScores

picMain.Font.Underline = True
picMain.Font.Bold = True

Str = "Round is Over"
PrintStickFormText Str, 7650, 400, vbBlack

picMain.Font.Underline = False
picMain.Font.Bold = False
'########
picMain.Font.Size = 12

'box for side info
'picMain.Line (750, 900)-(4200, 2200), BoxCol, BF
BorderedBox 750, 900, 4200, IIf(modStickGame.StickServer, 2600, 2200), BoxCol

RoundWinneri = FindStick(RoundWinnerID)
If RoundWinneri <> -1 Then
    
    Str = "Round Winner - " & Trim$(Stick(RoundWinneri).Name)
    
    PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 1000, Stick(RoundWinneri).colour
    
    
    If (Stick(RoundWinneri).Team = Neutral Or Stick(RoundWinneri).Team = Spec) = False Then
        
        Str = "Winning Team - " & GetTeamStr(Stick(RoundWinneri).Team)
        
        PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 1400, GetTeamColour(Stick(RoundWinneri).Team)
        
    Else
        Str = "No Winning Team"
        
        PrintStickFormText Str, 2200 - TextWidth(Str) / 2, 1400, vbBlack
        
    End If
    '--------
Else
    PrintStickFormText "No Winner", 1950, 1000, vbBlack
    PrintStickFormText "No Winning Team", 1585, 1400, vbBlack
End If

'picMain.ForeColor = MGrey

RoundTm = Round((RoundPausedAtThisTime + RoundWaitTime - GetTickCount()) / 1000)

Str = "Round will begin in " & CStr(RoundTm) & " second" & IIf(RoundTm > 1, "s", vbNullString)
PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 1800, vbBlack

If modStickGame.StickServer Then
    Str = "Press Space to start now"
    PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 2200, vbBlue
End If


picMain.Font.Size = 8


'decide if new round
If RoundTm <= 0 Then
    StopPlay False
End If


If LastPresenceSend + PresenceSendDelay < GetTickCount() Then
    
    If StickServer Then
        SendBroadcast sPresences & "0"
    Else
        modWinsock.SendPacket lSocket, ServerSockAddr, sPresences & CStr(Stick(0).ID)
    End If
    
    LastPresenceSend = GetTickCount()
End If

End Sub


Private Sub CheckMaxScore()
Dim i As Integer

If modStickGame.sv_GameType = gDeathMatch Then
    If LastScoreCheck + ScoreCheckDelay < GetTickCount() Then
        
        For i = 0 To NumSticksM1
            If (Stick(i).iKills - Stick(i).iDeaths) >= modStickGame.sv_WinScore Then
                RoundWinnerID = Stick(i).ID
                StopPlay True
                Exit For
            End If
        Next i
        
        
        LastScoreCheck = GetTickCount()
    End If
End If

End Sub

Public Sub GameTypeChanged()

AddMainMessage "Game Type - " & GetGameType(), False
HideCursor True

End Sub

'#################################################################################################

'Private Sub AddNadeTrail(iNade As Integer)
'
'ReDim Preserve Nade(iNade).NadeTrail(Nade(iNade).NumNadeTrails)
'
'With Nade(iNade).NadeTrail(Nade(iNade).NumNadeTrails)
'    .X = Nade(iNade).X
'    .Y = Nade(iNade).Y
'
'    .lCreation = GetTickCount()
'End With
'
'Nade(iNade).NumNadeTrails = Nade(iNade).NumNadeTrails + 1
'
'End Sub
'
'Private Sub RemoveNadeTrail(Index As Integer, iNade As Integer)
'
'Dim i As Integer
'
'If Nade(iNade).NumNadeTrails = 1 Then
'    Erase Nade(iNade).NadeTrail
'    Nade(iNade).NumNadeTrails = 0
'Else
'    For i = Index To Nade(iNade).NumNadeTrails - 2
'        Nade(iNade).NadeTrail(i) = Nade(iNade).NadeTrail(i + 1)
'    Next i
'
'    'Resize the array
'    ReDim Preserve Nade(iNade).NadeTrail(Nade(iNade).NumNadeTrails - 2)
'    Nade(iNade).NumNadeTrails = Nade(iNade).NumNadeTrails - 1
'End If
'
'End Sub
'
'Private Sub ProcessNadeTrails()
'Dim iTrail As Integer, iNade As Integer
'
'For iNade = 0 To NumNades - 1
'    iTrail = 0
'    Do While iTrail < Nade(iNade).NumNadeTrails
'        If Nade(iNade).NadeTrail(iTrail).lCreation + NadeTrail_Time / _
'                GetTimeZoneAdjust(Nade(iNade).NadeTrail(iTrail).X, Nade(iNade).NadeTrail(iTrail).Y) < GetTickCount() Then
'
'            RemoveNadeTrail iTrail, iNade
'            iTrail = iTrail - 1
'        End If
'
'        iTrail = iTrail + 1
'    Loop
'Next iNade
'
'
'End Sub
'
'Private Sub DrawNadeTrails()
'Dim i As Integer, j As Integer
'
'picMain.DrawWidth = 2
'picMain.FillStyle = vbFSSolid
'picMain.FillColor = NadeTrail_Colour
'
'For i = 0 To NumNades - 1
'    For j = 0 To Nade(i).NumNadeTrails - 2
'        modStickGame.sLine Nade(i).NadeTrail(j).X, Nade(i).NadeTrail(j).Y, _
'                           Nade(i).NadeTrail(j + 1).X, Nade(i).NadeTrail(j + 1).Y, _
'                           NadeTrail_Colour
'    Next j
'Next i
'
'picMain.FillStyle = vbFSTransparent
'picMain.DrawWidth = 1
'
'End Sub

'#################################################################################################

Private Sub AddGravityZone(X As Single, Y As Single)

ReDim Preserve GravityZone(NumGravityZones)

With GravityZone(NumGravityZones)
    .X = X
    .Y = Y
    
    .sSize = 50
End With

NumGravityZones = NumGravityZones + 1

End Sub

Private Sub RemoveGravityZone(Index As Integer)

Dim i As Integer

If NumGravityZones = 1 Then
    Erase GravityZone
    NumGravityZones = 0
Else
    For i = Index To NumGravityZones - 2
        GravityZone(i) = GravityZone(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve GravityZone(NumGravityZones - 2)
    NumGravityZones = NumGravityZones - 1
End If

End Sub

Private Sub ProcessGravityZones()
Dim i As Integer, j As Integer
Dim GTC As Long
Dim sGZ As Single
Const Growth_Rate As Single = 80, DeGrowth_Rate As Single = 4

'For i = 0 To NumSticksM1
'    Stick(i).sgGravityZone = GetGravityZoneAdjust(Stick(i).X, GetStickY(i), Gravity_Strength)
'Next i


'i = 0
GTC = GetTickCount()

Do While i < NumGravityZones
    
    If GravityZone(i).sSize < 50 Then
        RemoveGravityZone i
        i = i - 1
    ElseIf GravityZone(i).sSize > GravityZone_Radius Then
        GravityZone(i).sSize = GravityZone_Radius
    ElseIf GravityZone(i).bShrinking Then
        GravityZone(i).sSize = GravityZone(i).sSize - DeGrowth_Rate * modStickGame.StickTimeFactor * GetTimeZoneAdjust(GravityZone(i).X, GravityZone(i).Y)
    ElseIf GravityZone(i).sSize < GravityZone_Radius Then
        GravityZone(i).sSize = GravityZone(i).sSize + Growth_Rate * modStickGame.StickTimeFactor * GetTimeZoneAdjust(GravityZone(i).X, GravityZone(i).Y)
    Else
        GravityZone(i).bShrinking = True
    End If
    
    i = i + 1
Loop


End Sub

Private Sub DrawGravityZones()
Dim i As Integer

picMain.DrawWidth = 2


For i = 0 To NumGravityZones - 1
    
    modStickGame.sCircle GravityZone(i).X, GravityZone(i).Y, GravityZone(i).sSize, GravityZone_Colour
    
Next i

picMain.DrawWidth = 1

End Sub

'#################################################################################################

Private Sub AddTimeZone(X As Single, Y As Single, Optional sTimeAdjust As Single = def_TimeAdjust)

ReDim Preserve TimeZone(NumTimeZones)

With TimeZone(NumTimeZones)
    .X = X
    .Y = Y
    
    .sSize = 50
    
    .TimeAdjust = sTimeAdjust
    
    '.Decay = GetTickCount() + TimeZone_Time
End With

NumTimeZones = NumTimeZones + 1

End Sub

Private Sub RemoveTimeZone(Index As Integer)

Dim i As Integer

If NumTimeZones = 1 Then
    Erase TimeZone
    NumTimeZones = 0
Else
    For i = Index To NumTimeZones - 2
        TimeZone(i) = TimeZone(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve TimeZone(NumTimeZones - 2)
    NumTimeZones = NumTimeZones - 1
End If

End Sub

Private Sub ProcessTimeZones()
Dim i As Integer, j As Integer
Dim GTC As Long
Dim sTz As Single
Const Growth_Rate = 80, DeGrowth_Rate = 3

For i = 0 To NumSticksM1
    sTz = GetTimeZoneAdjust(Stick(i).X, GetStickY(i))
    
    If sTz <> Stick(i).sgTimeZone Then
        If i = 0 Then
            If StickInGame(0) Then
                SetSoundFreq sTz
                Stick(i).sgTimeZone = sTz
                'to reduce effect, raise MyTimeZone to a power < 1
                'e.g. SetSoundFreq Sqr(MyTimeZone)
                
            End If
        Else
            Stick(i).sgTimeZone = sTz
        End If
    End If
Next i


i = 0
GTC = GetTickCount()

Do While i < NumTimeZones
    
    If TimeZone(i).sSize < 50 Then
        RemoveTimeZone i
        i = i - 1
    ElseIf TimeZone(i).sSize > TimeZone_Radius Then
        TimeZone(i).sSize = TimeZone_Radius
    ElseIf TimeZone(i).bShrinking Then
        TimeZone(i).sSize = TimeZone(i).sSize - DeGrowth_Rate * modStickGame.StickTimeFactor * modStickGame.sv_StickGameSpeed
    ElseIf TimeZone(i).sSize < TimeZone_Radius Then
        TimeZone(i).sSize = TimeZone(i).sSize + Growth_Rate * modStickGame.StickTimeFactor * modStickGame.sv_StickGameSpeed
    Else
        TimeZone(i).bShrinking = True
    End If
    
    i = i + 1
Loop


End Sub

Private Sub DrawTimeZones()
Dim i As Integer

For i = 0 To NumTimeZones - 1
    
    modStickGame.sCircle TimeZone(i).X, TimeZone(i).Y, TimeZone(i).sSize, TimeZone_Colour
    
Next i

End Sub

Private Sub MotionStickObject(X As Single, Y As Single, Speed As Single, Heading As Single)

StickMotion X, Y, Speed * GetTimeZoneAdjust(X, Y), Heading

End Sub

Private Sub ApplyGravityVector(ByRef LastGravity As Long, _
    ByRef sgTimeZone As Single, ByRef Speed As Single, ByRef Heading As Single, _
    ByRef X As Single, ByRef Y As Single, _
    Optional GStrength As Single = Gravity_Strength)


Dim i As Integer
Dim final_GStrength As Single


If LastGravity + Gravity_Delay / sgTimeZone < GetTickCount() Then
    
    LastGravity = GetTickCount()
    
    For i = 0 To NumGravityZones - 1
        If GetDist(X, Y, GravityZone(i).X, GravityZone(i).Y) < GravityZone(i).sSize Then
            
            AddVectors Speed, Heading, _
                       Gravity_Zone_Strength, Gravity_Zone_Direction, _
                       Speed, Heading
            
        End If
    Next i
    
    
    AddVectors Speed, Heading, _
            GStrength, Gravity_Direction, _
            Speed, Heading
    
End If

End Sub

Private Function GetTimeZoneAdjust(X As Single, Y As Single) As Single
Dim i As Integer
Dim f As Single

'now cumulative
f = modStickGame.sv_StickGameSpeed

For i = 0 To NumTimeZones - 1
    If GetDist(X, Y, TimeZone(i).X, TimeZone(i).Y) < TimeZone(i).sSize Then
        f = f * TimeZone(i).TimeAdjust
    End If
Next i

GetTimeZoneAdjust = f

End Function

Private Sub ResetTimeLong(ByRef longToReset As Long, lTime As Long)

longToReset = GetTickCount() - lTime / 0.001

End Sub

'Private Function GetTimeZoneAdjust(X As Single, Y As Single) As Single
'Dim i As Integer
''Const sngSmall = 10 ^ -5
'
'i = PointInTimeZone(X, Y)
'If i > -1 Then
'    'If TimeZone(i).TimeAdjust Then
'        GetTimeZoneAdjust = TimeZone(i).TimeAdjust
'    'Else
'        'GetTimeZoneAdjust = sngSmall
'    'End If
'Else
'    GetTimeZoneAdjust = modStickGame.sv_StickGameSpeed
'End If
'
'
'End Function
'
'Private Function PointInTimeZone(X As Single, Y As Single) As Integer
'Dim i As Integer
'
'PointInTimeZone = -1
'
'For i = 0 To NumTimeZones - 1
'    If GetDist(X, Y, TimeZone(i).X, TimeZone(i).Y) < TimeZone(i).sSize Then
'        PointInTimeZone = i
'        Exit For
'    End If
'Next i
'
'End Function

Private Function GetSticksTimeZone(iStick As Integer) As Single
GetSticksTimeZone = Stick(iStick).sgTimeZone 'GetTimeZoneAdjust(Stick(iStick).X, Stick(iStick).Y)
End Function

Private Function GetMyTimeZone() As Single
GetMyTimeZone = Stick(0).sgTimeZone
End Function

'#################################################################################################

Private Sub AddBarrels()
Dim i As Integer
Dim iPlatform As Integer
Const nBarrelsToAdd = 8

For i = 1 To nBarrelsToAdd
    iPlatform = GetRandomPlatform()
    'If iPlatform = 7 Or iPlatform = 3 Then 'Or iPlatform = 6
        ''reduce amount of weapons in sniper's nest
        'iPlatform = GetRandomPlatform()
    'End If
    
    AddBarrel RandomXOnPlatform(iPlatform), YOnPlatform(iPlatform) - BarrelHeight + 100
Next i

End Sub

Private Sub AddBarrel(X As Single, Y As Single)

ReDim Preserve Barrel(NumBarrels)

With Barrel(NumBarrels)
    .X = X
    .Y = Y
    .iHealth = 100
    .LastTouchID = -1
    
    .ID = GenerateBarrelID()
End With

NumBarrels = NumBarrels + 1

End Sub

Private Sub RemoveBarrel(Index As Integer)

Dim i As Integer

If NumBarrels = 1 Then
    Erase Barrel
    NumBarrels = 0
Else
    For i = Index To NumBarrels - 2
        Barrel(i) = Barrel(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Barrel(NumBarrels - 2)
    NumBarrels = NumBarrels - 1
End If

End Sub

Private Sub DrawExplosiveBarrels()
Dim i As Integer, j As Integer
Const Barrel_BombSquad_Offset = BodyLen

picMain.DrawWidth = 2
picMain.ForeColor = vbBlack

If Stick(0).Perk = pBombSquad Then
    For i = 0 To NumBarrels - 1
        DrawBarrel i
        
        
        'j = FindStick(Barrel(i).LastTouchID)
        'If j > -1 Then
            'lCol = Stick(j).Colour
        'Else
            'lCol = vbBlack
        'End If
        
        modStickGame.sBox Barrel(i).X - Barrel_BombSquad_Offset, Barrel(i).Y - Barrel_BombSquad_Offset, _
                        Barrel(i).X + BarrelWidth + Barrel_BombSquad_Offset, Barrel(i).Y + BarrelHeight + Barrel_BombSquad_Offset, _
                        vbRed
        
        
    Next i
Else
    For i = 0 To NumBarrels - 1
        DrawBarrel i
    Next i
End If

End Sub

Private Sub DrawBarrel(i As Integer)
Const BarrelColour = vbRed, CrossLen = 75
Dim X As Single, Y As Single

modStickGame.sBoxFilled Barrel(i).X, Barrel(i).Y, _
                        Barrel(i).X + BarrelWidth, Barrel(i).Y + BarrelHeight, _
                        BarrelColour

X = Barrel(i).X + BarrelWidth / 2
Y = Barrel(i).Y + BarrelHeight / 2


modStickGame.sCircle X, Y - 50, HeadRadius, vbBlack

modStickGame.sLine X - CrossLen, Y - CrossLen, X + CrossLen, Y + CrossLen
modStickGame.sLine X - CrossLen, Y + CrossLen, X + CrossLen, Y - CrossLen


'health
modStickGame.sLine Barrel(i).X - 10, Barrel(i).Y - 100, Barrel(i).X + Barrel(i).iHealth * 2, Barrel(i).Y - 100

'modStickGame.PrintStickText "ID: " & Barrel(i).ID, Barrel(i).X + 500, Barrel(i).Y, vbBlack

End Sub

Private Sub ProcessExplosiveBarrels()
Dim i As Integer, j As Integer

Do While i < NumBullets
    
    For j = 0 To NumBarrels - 1
        If BulletNearBarrel(i, j) Then
            
            AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading - Pi
            AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading - piD2
            AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading + piD2
            
            
            If PointHearableOnSticksScreen(Bullet(i).X, Bullet(i).Y, 0) Then
                modAudio.PlayRicochet GetRelPan(Bullet(i).X)
            End If
            
            
            If Bullet(i).bHeadingChanged = False Or Bullet(i).bSniperBullet Then
                
                Barrel(j).LastTouchID = FindStick(Bullet(i).OwnerIndex)
                'Barrel(J).nBulletsHit = Barrel(J).nBulletsHit + IIf(Bullet(i).bSniperBullet, 3, 1)
                Barrel(j).iHealth = Barrel(j).iHealth - Bullet(i).Damage
                
                
                If Barrel(j).iHealth <= 0 Then
                    ExplodeBarrel j, True
                    RemoveBarrel j
                End If
                
            End If
            
            RemoveBullet i, False, True
            i = i - 1
            Exit For
            
        End If
    Next j
    
    i = i + 1
Loop


'nade collision in processnades
End Sub

Private Sub ExplodeBarrel(i As Integer, ByVal bSendBroadcast As Boolean)
Dim j As Integer

Dim Dist As Single
Dim OwnerIndex As Integer
Dim MaxDist As Single
Dim ExplosionForceDist As Single, AngleToStick As Single

Const Barrel_Explode_RadiusX2 As Single = Barrel_Explode_Radius * 2
Const ChopperLenX1p2 As Single = ChopperLen * 1.2
Const BarrelMultiple_X_Inc As Single = BarrelMultiple * 12000
Const Barrel_Explosion_Dist As Single = 250
Const ShieldWaveDispersion As Single = 600

OwnerIndex = FindStick(Barrel(i).LastTouchID)


If bSendBroadcast Then
    If modStickGame.StickServer Then
        SendBroadcast sExplodeBarrels & CStr(Barrel(i).ID)
    Else
        modWinsock.SendPacket lSocket, ServerSockAddr, sExplodeBarrels & CStr(Barrel(i).ID)
    End If
End If


AddExplosion Barrel(i).X, Barrel(i).Y, 750
AddExplosion Barrel(i).X + Barrel_Explosion_Dist, Barrel(i).Y + PM_Rnd() * Barrel_Explosion_Dist, 750
AddExplosion Barrel(i).X - Barrel_Explosion_Dist, Barrel(i).Y + PM_Rnd() * Barrel_Explosion_Dist, 750
For j = 1 To IIf(modStickGame.cg_Smoke, 10, 3)
    AddSmokeGroup Barrel(i).X, Barrel(i).Y, 5, 100 * Rnd(), Pi2 * Rnd(), , True
Next j

For j = 1 To 2 + Rnd() * 3
    AddNadeTrail_Simple Barrel(i).X, Barrel(i).Y
Next j

For j = 1 To 5
    AddFlame Barrel(i).X, Barrel(i).Y, Pi2 * Rnd(), Flame_Speed, Barrel(i).LastTouchID, OwnerIndex ', False
Next j

If PointHearableOnSticksScreen(Barrel(i).X, Barrel(i).Y, 0) Then
    modAudio.PlayNadeExplosion GetRelPan(Barrel(i).X)
Else
    modAudio.PlayBackGroundNade GetRelPan(Barrel(i).X)
End If


If OwnerIndex > -1 Then AddFire Barrel(i).X, Barrel(i).Y + BarrelHeight, OwnerIndex


For j = 0 To NumSticksM1
    'apply damage
    
    If StickInGame(j) Then
        
        Dist = GetDist(Stick(j).X, Stick(j).Y, Barrel(i).X, Barrel(i).Y)
        
        If Stick(j).WeaponType = Chopper Then
            MaxDist = ChopperLen
            ExplosionForceDist = ChopperLenX1p2
        Else
            MaxDist = Barrel_Explode_Radius
            ExplosionForceDist = Barrel_Explode_RadiusX2
        End If
        
        AngleToStick = FindAngle(Barrel(i).X, Barrel(i).Y, Stick(j).X, Stick(j).Y)
        If Dist < ExplosionForceDist Then
            
            If Stick(j).WeaponType <> Chopper Then
                If StickiHasState(j, STICK_PRONE) = False Then
                    If Stick(j).Shield = 0 Then
                        AddVectors Stick(j).Speed, Stick(j).Heading, _
                            BarrelMultiple_X_Inc / (Dist + 1), AngleToStick, _
                            Stick(j).Speed, Stick(j).Heading
                    End If
                End If
            End If
            
        End If


        If Dist < MaxDist Then
            
            If Stick(j).Shield Then
                AngleToStick = AngleToStick - Pi
                AddShieldWave Stick(j).X, Stick(j).Y, AngleToStick
                AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
                AddShieldWave Stick(j).X + PM_Rnd() * ShieldWaveDispersion, Stick(j).Y + PM_Rnd() * ShieldWaveDispersion, AngleToStick
            Else
                Stick(j).bOnSurface = False
                Stick(j).Y = Stick(j).Y - 100
            End If
            
            
            If j = 0 Or Stick(j).IsBot Then
                
                If OwnerIndex <> -1 Then
                    If IsAlly(Stick(j).Team, Stick(OwnerIndex).Team) = False Or OwnerIndex = j Then
                        If StickInvul(j) = False Then
                            
                            On Error Resume Next
                            
                            If Stick(j).WeaponType = Chopper Then
                                DamageStick Chopper_Damage_Reduction * 60, j, OwnerIndex, , False 'barrel = 60 damage
                            Else
                                DamageStick 250000 / Dist, j, OwnerIndex, , False 'bullet
                            End If
                            
                            If Err.Number <> 0 Then 'div zero error
                                Stick(j).Health = 0
                                Err.Clear
                            End If
                            
                            If Stick(j).Health < 1 Then
                                Call Killed(j, OwnerIndex, kBarrel)
                            End If
                            
                        End If 'spawn invul endif
                    End If 'ally endif
                End If 'owner index endif
            End If 'myid endif
        End If 'dist endif
    End If 'stickingame endif
Next j

ExplodeAll Barrel(i).X, Barrel(i).Y, Barrel(i).LastTouchID, -1, i

End Sub

'#################################################################################################
'Explosive Barrels

Private Sub SendBarrelRefresh(Optional bForce As Boolean)
Static LastSend As Long
Const Barrel_Refresh_Delay = 2000
Dim i As Integer
Dim sToSend As String

If LastSend + Barrel_Refresh_Delay < GetTickCount() Or bForce Then
    
    For i = 0 To NumBarrels - 1
        sToSend = sToSend & _
            Barrel(i).X & mPacketSep & _
            Barrel(i).Y & mPacketSep & _
            Barrel(i).iHealth & mPacketSep & _
            Barrel(i).LastTouchID & mPacketSep & _
            Barrel(i).ID & mPacketSep & UpdatePacketSep
        
    Next i
    
    
    SendBroadcast sBarrelRefreshs & sToSend
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub ReceiveBarrelRefresh(sPacket As String)
Dim Parts() As String, SubParts() As String
Dim i As Integer

On Error GoTo EH
Parts = Split(sPacket, UpdatePacketSep)

NumBarrels = UBound(Parts)

If NumBarrels > -1 Then
    ReDim Barrel(NumBarrels - 1)
    
    
    For i = 0 To NumBarrels - 1
        
        SubParts = Split(Parts(i), mPacketSep)
        
        With Barrel(i)
            .X = CSng(SubParts(0))
            .Y = CSng(SubParts(1))
            
            .iHealth = CInt(SubParts(2))
            .LastTouchID = CInt(SubParts(3))
            
            .ID = CInt(SubParts(4))
        End With
        
        Erase SubParts
    Next i
Else
    Erase Barrel
    NumBarrels = 0
End If

Erase Parts
Exit Sub
EH:
Erase Parts
Erase SubParts
End Sub


'Fire
'#################################################################################################
'Private Sub SendFireRefresh()
'Const Fire_Refresh_Delay = 20000
'Dim i As Integer
'ReDim sToSend(0 To modStickGame.ubdPlatforms) As String
'
'If LastFireRefresh + Fire_Refresh_Delay < GetTickCount() Then
'
'    For i = 0 To NumFlames - 1
'        sToSend = sToSend & UpdatePacketSep & Flame(i).mPacketSep
'    Next i
'
'    'sToSend(0) = X1 # X2 # X3 ...
'
'    For i = 0 To modStickGame.ubdPlatforms
'        SendBroadcast sFireRefreshs & CStr(i) & UpdatePacketSep & sToSend(i) & vbSpace & CStr(MakeSquareNumber())
'        '             from here -->    |                                           to here --> |
'        '                                        is the format
'
'
'        'format, for each platform =
'        '(iPlatform)(SEP)(Fire1X)(mPacketSep)(Fire2X)(mPacketSep) ... (FireNX)(mPacketSep)(SPACE)(SQUARE)
'        '             ^ needed in case iPlatform > 9
'
'    Next i
'
'
'
'    LastFireRefresh = GetTickCount()
'End If
'
'End Sub
'
'Private Sub ReceiveFireRefresh(sPacket As String)
'Dim Parts() As String
'Dim i As Integer, nParts As Integer
'Dim iPlatform As Integer
'
'
'If IsValidVarPacket(sPacket) Then
'    On Error GoTo EH
'    i = InStr(1, sPacket, UpdatePacketSep)
'    iPlatform = Left$(sPacket, i - 1)
'
'    Parts = Split(Mid$(sPacket, i + 1, InStrRev(sPacket, vbSpace) - 1), mPacketSep)
'
'    nParts = UBound(Parts)
'
'    'get rid of all the Fire on the platform
'    i = 0
'    Do While i < NumFire
'        If Fire(i).iPlatform = iPlatform Then
'            RemoveFire i
'            i = i - 1
'        End If
'        i = i + 1
'    Loop
'
'
'    If nParts > -1 Then
'
'        For i = 0 To nParts - 1
'            AddSingleFire CSng(Parts(i)), iPlatform
'        Next i
'    'Else
'        'already erased
'    End If
'End If
'
'
'EH:
'Erase Parts
'End Sub

'Grass
'#################################################################################################
Private Sub SendGrassRefresh()
Const Grass_Refresh_Delay = 20000
Dim i As Integer
ReDim sToSend(0 To modStickGame.ubdPlatforms) As String

If LastGrassRefresh + Grass_Refresh_Delay < GetTickCount() Then
    
    For i = 0 To NumGrass - 1
        sToSend(Grass(i).iPlatform) = sToSend(Grass(i).iPlatform) & Grass(i).X & mPacketSep
    Next i
    
    'sToSend(0) = X1 # X2 # X3 ...
    
    For i = 0 To modStickGame.ubdPlatforms
        SendBroadcast sGrassRefreshs & CStr(i) & UpdatePacketSep & sToSend(i) & vbSpace & CStr(MakeSquareNumber())
        '             from here -->    |                                           to here --> |
        '                                        is the format
        
        
        'format, for each platform =
        '(iPlatform)(SEP)(Grass1X)(mPacketSep)(Grass2X)(mPacketSep) ... (GrassNX)(mPacketSep)(SPACE)(SQUARE)
        '             ^ needed in case iPlatform > 9
        
    Next i
    
    
    
    LastGrassRefresh = GetTickCount()
End If

End Sub

Private Sub ReceiveGrassRefresh(sPacket As String)
Dim Parts() As String
Dim i As Integer, nParts As Integer
Dim iPlatform As Integer


If IsValidVarPacket(sPacket) Then
    On Error GoTo EH
    i = InStr(1, sPacket, UpdatePacketSep)
    iPlatform = Left$(sPacket, i - 1)
    
    Parts = Split(Mid$(sPacket, i + 1, InStrRev(sPacket, vbSpace) - 1), mPacketSep)
    
    nParts = UBound(Parts)
    
    'get rid of all the grass on the platform
    i = 0
    Do While i < NumGrass
        If Grass(i).iPlatform = iPlatform Then
            RemoveGrass i
            i = i - 1
        End If
        i = i + 1
    Loop
    
    
    If nParts > -1 Then
        
        For i = 0 To nParts - 1
            AddSingleGrass CSng(Parts(i)), iPlatform
        Next i
    'Else
        'already erased
    End If
End If


EH:
Erase Parts
End Sub

'#################################################################################################

Private Sub ValidateWeapons()
Dim i As Integer

If modStickGame.sv_AllowedWeapons(Stick(0).CurrentWeapons(1)) = False Then
    i = 1
ElseIf modStickGame.sv_AllowedWeapons(Stick(0).CurrentWeapons(2)) = False Then
    i = 2
End If



If i Then
    AddMainMessage GetWeaponName(Stick(0).CurrentWeapons(i)) & " has been banned", False
    
    '##################################################################
    'switch currentweapon()
    
    Stick(0).CurrentWeapons(i) = AK - 1
    Do
        Stick(0).CurrentWeapons(i) = Stick(0).CurrentWeapons(i) + 1
    Loop Until modStickGame.sv_AllowedWeapons(Stick(0).CurrentWeapons(i))
    
    
    '##################################################################
    'don't allow two of the same type
    If Stick(0).CurrentWeapons(1) = Stick(0).CurrentWeapons(2) Then
        i = IIf(i = 1, 2, 1)
        
        Stick(0).CurrentWeapons(i) = AK
        
        Do
            Stick(0).CurrentWeapons(i) = Stick(0).CurrentWeapons(i) + 1
        Loop Until modStickGame.sv_AllowedWeapons(Stick(0).CurrentWeapons(i)) And Stick(0).CurrentWeapons(1) <> Stick(0).CurrentWeapons(2)
        
    End If
    
    
    '##################################################################
    'switch to new allowed weapon, if we need to
    If modStickGame.sv_AllowedWeapons(Stick(0).WeaponType) = False Then
        SwitchWeapon Stick(0).CurrentWeapons(i) '"int i" can be either 1 or 2 here
    End If
ElseIf modStickGame.sv_AllowedWeapons(Stick(0).WeaponType) = False Then
    'i.e. chopper
    
    AddMainMessage GetWeaponName(Stick(0).WeaponType) & " has been banned", False
    SwitchWeapon Stick(0).CurrentWeapons(1)
    
End If


If modStickGame.StickServer Then
    For i = 1 To NumSticksM1
        If Stick(i).IsBot Then
            
            If modStickGame.sv_AllowedWeapons(Stick(i).WeaponType) = False Then
                Stick(i).WeaponType = AK - 1 'eWeaponTypes.RPG * Rnd() - 1
                Do
                    Stick(i).WeaponType = Stick(i).WeaponType + 1
                Loop Until modStickGame.sv_AllowedWeapons(Stick(i).WeaponType)
                
                Make_Weapon_Default_FireMode i
            End If
            
        End If
    Next i
End If


End Sub

'#################################################################################################
'GravityZone refresh

Private Sub SendGravityZoneRefresh()
Static LastSend As Long
Const GravityZone_Refresh_Delay = 2000
Dim i As Integer
Dim sToSend As String

If LastSend + GravityZone_Refresh_Delay < GetTickCount() Then
    
    For i = 0 To NumGravityZones - 1
        
        sToSend = sToSend & _
            GravityZone(i).X & mPacketSep & _
            GravityZone(i).Y & mPacketSep & _
            GravityZone(i).sSize & mPacketSep & _
            Abs(GravityZone(i).bShrinking) & mPacketSep & UpdatePacketSep
        
    Next i
    
    
    SendBroadcast sGravityZoneRefreshs & sToSend
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub ReceiveGravityZoneRefresh(sPacket As String)
Dim Parts() As String, SubParts() As String
Dim i As Integer

On Error GoTo EH
Parts = Split(sPacket, UpdatePacketSep)

NumGravityZones = UBound(Parts)

If NumGravityZones > -1 Then
    ReDim GravityZone(NumGravityZones - 1)
    
    
    For i = 0 To NumGravityZones - 1
        
        SubParts = Split(Parts(i), mPacketSep)
        
        With GravityZone(i)
            .X = CSng(SubParts(0))
            .Y = CSng(SubParts(1))
            
            .sSize = CSng(SubParts(2))
            .bShrinking = CBool(SubParts(3))
            
        End With
        
        Erase SubParts
    Next i
Else
    Erase GravityZone
    NumGravityZones = 0
End If

Erase Parts
Exit Sub
EH:
Erase Parts
Erase SubParts
End Sub

'#################################################################################################
'TimeZone refresh

Private Sub SendTimeZoneRefresh()
Static LastSend As Long
Const TimeZone_Refresh_Delay = 2000
Dim i As Integer
Dim sToSend As String

If LastSend + TimeZone_Refresh_Delay < GetTickCount() Then
    
    For i = 0 To NumTimeZones - 1
        
        sToSend = sToSend & _
            TimeZone(i).X & mPacketSep & _
            TimeZone(i).Y & mPacketSep & _
            TimeZone(i).TimeAdjust & mPacketSep & _
            TimeZone(i).sSize & mPacketSep & _
            Abs(TimeZone(i).bShrinking) & mPacketSep & UpdatePacketSep
        
    Next i
    
    
    SendBroadcast sTimeZoneRefreshs & sToSend
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub ReceiveTimeZoneRefresh(sPacket As String)
Dim Parts() As String, SubParts() As String
Dim i As Integer

On Error GoTo EH
Parts = Split(sPacket, UpdatePacketSep)

NumTimeZones = UBound(Parts)

If NumTimeZones > -1 Then
    ReDim TimeZone(NumTimeZones - 1)
    
    
    For i = 0 To NumTimeZones - 1
        
        SubParts = Split(Parts(i), mPacketSep)
        
        With TimeZone(i)
            .X = CSng(SubParts(0))
            .Y = CSng(SubParts(1))
            
            .TimeAdjust = CSng(SubParts(2))
            
            .sSize = CSng(SubParts(3))
            .bShrinking = CBool(SubParts(4))
            
        End With
        
        Erase SubParts
    Next i
Else
    Erase TimeZone
    NumTimeZones = 0
End If

Erase Parts
Exit Sub
EH:
Erase Parts
Erase SubParts
End Sub

'#################################################################################################
'mine refresh

Private Sub SendMineRefresh()
Static LastSend As Long
Const Mine_Refresh_Delay = 2000
Dim i As Integer
Dim sToSend As String

If LastSend + Mine_Refresh_Delay < GetTickCount() Then
    
    Do While i < NumMines
        
        If FindStick(Mine(i).OwnerID) > -1 Then
            
            sToSend = sToSend & _
                Mine(i).X & mPacketSep & _
                Mine(i).Y & mPacketSep & _
                Mine(i).bOnSurface & mPacketSep & _
                Mine(i).colour & mPacketSep & _
                Mine(i).OwnerID & mPacketSep & _
                Mine(i).ID & mPacketSep & UpdatePacketSep
            
        Else
            RemoveMine i
            i = i - 1
        End If
        
        i = i + 1
    Loop
    
    
    SendBroadcast sMineRefreshs & sToSend
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub ReceiveMineRefresh(sPacket As String)
Dim Parts() As String, SubParts() As String
Dim i As Integer

On Error GoTo EH
Parts = Split(sPacket, UpdatePacketSep)

NumMines = UBound(Parts)

If NumMines > -1 Then
    ReDim Mine(NumMines - 1)
    
    
    For i = 0 To NumMines - 1
        SubParts = Split(Parts(i), mPacketSep)
        
        With Mine(i)
            .X = CSng(SubParts(0))
            .Y = CSng(SubParts(1))
            
            .bOnSurface = CBool(SubParts(2))
            .colour = CLng(SubParts(3))
            
            .OwnerID = CInt(SubParts(4))
            
            .ID = CInt(SubParts(5))
        End With
        
        Erase SubParts
    Next i
Else
    Erase Mine
    NumMines = 0
End If

Erase Parts
Exit Sub
EH:
Erase Parts
Erase SubParts
End Sub


'static weapons
'#################################################################################################
Private Function StaticWeaponToString(vType As ptStaticWeapon) As String

StaticWeaponToString = CStr( _
                    vType.Heading & mPacketSep & _
                    vType.iWeapon & mPacketSep & _
                    vType.Speed & mPacketSep & _
                    vType.X & mPacketSep & _
                    vType.Y & mPacketSep & _
                    Abs(vType.bOnSurface) & mPacketSep _
                    )

'don't send lastgravity or bOnSurface - different on all
'actually, i will

End Function

Private Function StaticWeaponFromString(buf As String) As ptStaticWeapon

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With StaticWeaponFromString
    .LastGravity = 1
    
    .Heading = CSng(Parts(0))
    .iWeapon = CInt(Parts(1))
    .Speed = CSng(Parts(2))
    .X = CSng(Parts(3))
    .Y = CSng(Parts(4))
    .bOnSurface = CBool(Parts(5))
End With

EH:
Erase Parts
End Function

Private Sub ProcessStaticWeaponPacket(sPacket As String)
Dim vSWeapons() As String
Dim i As Integer

On Error GoTo EH

vSWeapons = Split(sPacket, UpdatePacketSep)
NumStaticWeapons = UBound(vSWeapons)

ReDim StaticWeapon(0 To NumStaticWeapons - 1)

For i = 0 To NumStaticWeapons - 1
    StaticWeapon(i) = StaticWeaponFromString(vSWeapons(i + 1))
Next i

EH:
End Sub

Public Sub SendStaticWeaponsPacket()
Dim sToSend As String
Dim i As Integer
Static LastSend As Long


If LastSend + StaticWeaponSendDelay < GetTickCount() Then
    
    For i = 0 To NumStaticWeapons - 1
        sToSend = sToSend & UpdatePacketSep & StaticWeaponToString(StaticWeapon(i))
    Next i
    
    
    SendBroadcast sStaticWeaponUpdates & sToSend
    
    LastSend = GetTickCount()
End If

End Sub

'Server Vars
'#################################################################################################
Private Function ServerVarToString(vType As ptServerVars) As String

ServerVarToString = CStr( _
                    vType.sgGameSpeed & mPacketSep & _
                    Abs(vType.bHardCore) & mPacketSep & _
                    Abs(vType.bHPBonus) & mPacketSep & _
                    vType.iScoreToWin & mPacketSep & _
                    vType.iGameType & mPacketSep & _
                    Abs(vType.bBulletsThroughWalls) & mPacketSep & _
                    vType.iSpawnDelay & mPacketSep & _
                    Abs(vType.bDrawNadeTime) & mPacketSep & _
                    vType.sgDamageFactor & mPacketSep & _
                    vType.sAllowedWeapons & mPacketSep & _
                    Abs(vType.bSpawnWithShield) & mPacketSep & _
                    vType.iSequenceNo & mPacketSep _
                    )

End Function

Private Function ServerVarFromString(buf As String) As ptServerVars

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With ServerVarFromString
    .sgGameSpeed = CSng(Parts(0))
    .bHardCore = CBool(Parts(1))
    .bHPBonus = CBool(Parts(2))
    .iScoreToWin = CInt(Parts(3))
    .iGameType = CInt(Parts(4))
    .bBulletsThroughWalls = CBool(Parts(5))
    .iSpawnDelay = CInt(Parts(6))
    .bDrawNadeTime = CBool(Parts(7))
    .sgDamageFactor = CSng(Parts(8))
    .sAllowedWeapons = Parts(9)
    .bSpawnWithShield = CBool(Parts(10))
    .iSequenceNo = CLng(Parts(11))
End With

EH:
Erase Parts
End Function

Private Sub ProcessServerVarPacket(vPacket As String)
Dim vServerVars As ptServerVars
Dim i As Integer

If IsValidVarPacket(vPacket) Then
    On Error GoTo EH
    vServerVars = ServerVarFromString(Left$(vPacket, InStr(1, vPacket, vbSpace) - 1))
    
    If vServerVars.iSequenceNo >= LastServerSettingVar Then
        LastServerSettingVar = vServerVars.iSequenceNo
        
        
        For i = 0 To eWeaponTypes.Chopper
            modStickGame.sv_AllowedWeapons(i) = CBool(Mid$(vServerVars.sAllowedWeapons, i + 1, 1))
        Next i
        
        
        modStickGame.sv_Draw_Nade_Time = vServerVars.bDrawNadeTime
        
        modStickGame.sv_HPBonus = vServerVars.bHPBonus
        
        modStickGame.sv_WinScore = vServerVars.iScoreToWin
        modStickGame.sv_Spawn_Delay = vServerVars.iSpawnDelay
        
        modStickGame.sv_Damage_Factor = vServerVars.sgDamageFactor
        
        
        modStickGame.sv_SpawnWithShields = vServerVars.bSpawnWithShield
        
        
        'check for a change
        If modStickGame.sv_StickGameSpeed <> vServerVars.sgGameSpeed Then
            If vServerVars.sgGameSpeed <= 1.2 Then
                If vServerVars.sgGameSpeed >= 0.1 Then
                    StickGameSpeedChanged modStickGame.sv_StickGameSpeed, vServerVars.sgGameSpeed
                    
                    modStickGame.sv_StickGameSpeed = vServerVars.sgGameSpeed
                End If
            End If
        End If
        
        
        
        If modStickGame.sv_Hardcore <> vServerVars.bHardCore Then
            AddMainMessage "Hardcore Mode " & IIf(modStickGame.sv_Hardcore, "Off", "On"), False
            modStickGame.sv_Hardcore = vServerVars.bHardCore
        End If
        If modStickGame.sv_BulletsThroughWalls <> vServerVars.bBulletsThroughWalls Then
            AddMainMessage "Bullets can" & IIf(vServerVars.bBulletsThroughWalls, vbNullString, "'t") & " pass through walls", False
            modStickGame.sv_BulletsThroughWalls = vServerVars.bBulletsThroughWalls
        End If
        
    '    If modStickGame.sv_2Weapons <> vServerVars.b2Weapons Then
    '        modStickGame.sv_2Weapons = vServerVars.b2Weapons
    '
    '        If modStickGame.sv_2Weapons Then
    '            AddMainMessage "You can only carry two weapons (1 or 2 to switch)"
    '            SetCurrentWeapons
    '        Else
    '            AddMainMessage "You have access to all weapons"
    '        End If
    '    End If
        
        
        
        
        If modStickGame.sv_GameType <> vServerVars.iGameType Then
            modStickGame.sv_GameType = vServerVars.iGameType
            
            GameTypeChanged
            
            If modStickGame.sv_GameType = gDeathMatch Then
                For i = 0 To NumSticksM1
                    Stick(i).bAlive = True
                Next i
                
            ElseIf modStickGame.sv_GameType = gCoOp Then
                MoveStickToCoOpStart 0
                
            End If
            
        End If
        
        
        'If modStickGame.sv_AllowRockets = False Then
        '    If Stick(0).WeaponType = RPG Then
        '        Stick(0).WeaponType = AK
        '    End If
        'End If
    End If
End If


EH:
End Sub

Private Function IsValidVarPacket(ByVal sPacket As String) As Boolean
Dim iSquare As Single

On Error GoTo EH

iSquare = Mid$(sPacket, InStrRev(sPacket, vbSpace) + 1)

IsValidVarPacket = IsSquare(iSquare)

EH:
End Function

Public Sub SendServerVarPacket(Optional bForce As Boolean = False)
Dim vServerVars As ptServerVars
Dim i As Integer

Static LastSend As Long


If LastSend + ServerVarSendDelay < GetTickCount() Or bForce Then
    
    With vServerVars
        .sgGameSpeed = modStickGame.sv_StickGameSpeed
        .bHardCore = modStickGame.sv_Hardcore
        .bHPBonus = modStickGame.sv_HPBonus
        .iScoreToWin = modStickGame.sv_WinScore
        .iGameType = modStickGame.sv_GameType
        .bBulletsThroughWalls = modStickGame.sv_BulletsThroughWalls
        .iSpawnDelay = modStickGame.sv_Spawn_Delay
        .bDrawNadeTime = modStickGame.sv_Draw_Nade_Time
        .sgDamageFactor = modStickGame.sv_Damage_Factor
        .bSpawnWithShield = modStickGame.sv_SpawnWithShields
        
        For i = 0 To eWeaponTypes.Chopper
            .sAllowedWeapons = .sAllowedWeapons & Abs(modStickGame.sv_AllowedWeapons(i))
        Next i
        
        
        LastServerSettingVar = LastServerSettingVar + 1
        .iSequenceNo = LastServerSettingVar
    End With
    
    
    SendBroadcast sServerVarss & ServerVarToString(vServerVars) & vbSpace & CStr(MakeSquareNumber())
    
    LastSend = GetTickCount()
End If

End Sub

'Position/Status
'#################################################################################################

Private Function mPacketToString() As String

With mPacket
    mPacketToString = _
        CStr(.ActualFacing) & mPacketSep & _
        CStr(.Facing) & mPacketSep & _
        CStr(.Heading) & mPacketSep & _
        CStr(.Health) & mPacketSep & _
        CStr(.Shield) & mPacketSep & _
        CStr(.ID) & mPacketSep & _
        CStr(.PacketID) & mPacketSep & _
        CStr(.Speed) & mPacketSep & _
        CStr(.state) & mPacketSep & _
        CStr(.WeaponType) & mPacketSep & _
        CStr(.X) & mPacketSep & _
        CStr(.Y) & mPacketSep & _
        CStr(.iNadeType) & mPacketSep
End With

End Function

Private Sub mPacketFromString(buf As String) 'As ptPacket

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With mPacket
    .ActualFacing = CSng(Parts(0))
    .Facing = CSng(Parts(1))
    .Heading = CSng(Parts(2))
    .Health = CInt(Parts(3))
    .Shield = CInt(Parts(4))
    .ID = CInt(Parts(5))
    .PacketID = CLng(Parts(6))
    .Speed = CSng(Parts(7))
    .state = CInt(Parts(8))
    .WeaponType = CInt(Parts(9))
    .X = CSng(Parts(10))
    .Y = CSng(Parts(11))
    .iNadeType = CInt(Parts(12))
End With

Erase Parts

EH:
End Sub

Private Sub ProcessUpdatePacket(ByVal sPacket As String)

'Dim Num As Integer
Dim i As Integer, j As Integer ', k As Integer
Dim Sticks() As String

Sticks = Split(sPacket, UpdatePacketSep)


'Loop through each stick's info
For i = 0 To UBound(Sticks)
    
    On Error GoTo EH
    
    'Extract stick info
    'sstick = Left$(sPacket, Len(mPacket) + 1)
    'sPacket = Right$(sPacket, Len(sPacket) - (Len(mPacket) + 1))
    
    
    If LenB(Sticks(i)) Then
        'CopyMemory mPacket, ByVal sstick, Len(sstick)
        
        'copy it into mPacket
        
        mPacketFromString Sticks(i)
        
        'Does this stick already exist?
        If FindStick(mPacket.ID) = -1 Then
            If StickServer Then
                
                'we are server => corrent; ignore him
                GoTo nextBit
                
            Else
                
                'new stick. Make new spot and assign ID
                j = AddStick()
                Stick(j).ID = mPacket.ID
                
            End If
        End If
        
        'Is this the local stick?
        If mPacket.ID <> Stick(0).ID Then
            'Is this a new packet?
            j = FindStick(mPacket.ID) 'error handler invoked above
            If Stick(j).LastPacketID < mPacket.PacketID Then 'And j > 0 Then
                'Replace stick data with new data
                With Stick(j)
                    .LastPacketID = mPacket.PacketID
                    .LastPacket = GetTickCount()
                    
                    '.Colour = mPacket.Colour
                    .ActualFacing = mPacket.ActualFacing
                    
                    .Facing = mPacket.Facing
'                    If (.State And Stick_Fire) = 0 Then
'                        .Facing = .ActualFacing
'                    End If
                    
                    .Heading = mPacket.Heading
                    '.ID = mpacket.ID
                    '.Name = mPacket.Name
                    .Speed = mPacket.Speed
                    .state = mPacket.state
                    .X = mPacket.X
                    .Y = mPacket.Y
                    
                    .Health = mPacket.Health
                    .Shield = mPacket.Shield
                    '.Armour = mPacket.Armour
                    '.bAlive = mPacket.bAlive
                    .WeaponType = mPacket.WeaponType
                    '.PrevWeapon = mPacket.PrevWeapon
                    
                    .iNadeType = mPacket.iNadeType
                    
                    'If Not modStickGame.StickServer Then
                        '.iKills = mPacket.iKills
                        '.iDeaths = mPacket.iDeaths
                        '.iKillsInARow = mPacket.iKillsInARow
                    'End If
                    
                    '.Team = mPacket.Team
                    
                    '.bSilenced = mPacket.bSilenced
                    '.bTyping = mPacket.bTyping
                    '.Perk = mPacket.Perk
                    '.MaskID = mPacket.MaskID
                    '.bFlashed = mPacket.bFlashed
                    '.bOnFire = mPacket.bOnFire
                    
                End With
            End If 'packetid endif
            
'        ElseIf Not modStickGame.StickServer Then
'
'            j = FindStick(mPacket.ID)
'
'            If j <> -1 Then
'                'only let it update our kills+deaths
'                Stick(j).iKills = mPacket.iKills
'                Stick(j).iDeaths = mPacket.iDeaths
'
'                'if our killsinarow < mpacket's, then update
'                If Stick(j).iKillsInARow < mPacket.iKillsInARow Then
'                    Stick(0).iKillsInARow = mPacket.iKillsInARow
'                    CheckKillsInARow
'                End If
'
'            End If
            
            
        End If 'myid/iKills endif
    End If 'lenb endif
    
nextBit:
Next i

EH:

LastUpdatePacket = GetTickCount()
End Sub

Private Sub SendUpdatePacket()

'If it's not time to send a mPacket, exit sub
If PacketTimer + mPacket_SEND_DELAY < GetTickCount() Then
    
    'Reset the mPacket timer
    PacketTimer = GetTickCount()
    
    'Is this a StickServer mPacket, or a client mPacket?
    If StickServer Then
        'StickServer mPacket
        SendServerUpdatePacket
    Else
        'Client mPacket
        SendClientUpdatePacket
    End If
End If

End Sub

Private Sub SendClientUpdatePacket()

Dim sPacket As String

'Populate the mPacket type
Stick(0).LastPacketID = Stick(0).LastPacketID + 1

With mPacket
    '.Colour = Stick(0).Colour
    .ActualFacing = Stick(0).ActualFacing
    .Facing = Stick(0).Facing
    .Heading = Stick(0).Heading
    .ID = Stick(0).ID
'    .Name = Stick(0).Name
    .PacketID = Stick(0).LastPacketID
    .Speed = Stick(0).Speed
    .state = Stick(0).state
    .X = Stick(0).X
    .Y = Stick(0).Y
    
    .Health = Stick(0).Health
    .Shield = Stick(0).Shield
'    .Armour = Stick(0).Armour
'    .bAlive = Stick(0).bAlive
    .WeaponType = Stick(0).WeaponType
    '.PrevWeapon = Stick(0).PrevWeapon
    
'    .iKills = Stick(0).iKills
'    .iDeaths = Stick(0).iDeaths
'    .iKills = Stick(0).iKillsInARow
    
'    .Team = Stick(0).Team
    
    '.bSilenced = Stick(0).bSilenced
    '.bTyping = Stick(0).bTyping
'    .Perk = Stick(0).Perk
'    .MaskID = Stick(0).MaskID
    .iNadeType = Stick(0).iNadeType
    
    '.bFlashed = Stick(0).bFlashed
    '.bOnFire = Stick(0).bOnFire
End With

sPacket = sUpdates & mPacketToString() & UpdatePacketSep

modWinsock.SendPacket lSocket, ServerSockAddr, sPacket

End Sub

Private Sub SendServerUpdatePacket()

Dim i As Long
Dim sPacket As String

'Increment the local Stick's LastPacketID
On Error GoTo EH
Stick(0).LastPacketID = Stick(0).LastPacketID + 1

For i = 0 To NumSticksM1
    
    If Stick(i).IsBot Then
        Stick(i).LastPacketID = Stick(i).LastPacketID + 1
    End If
    
    'Fill the mPacket
    With mPacket
'        .Colour = Stick(i).Colour
        .ActualFacing = Stick(i).ActualFacing
        .Facing = Stick(i).Facing
        .Heading = Stick(i).Heading
        .ID = Stick(i).ID
'        .Name = Stick(i).Name
        '.PacketID = Stick(i).LastPacketID
        .PacketID = Stick(i).LastPacketID 'IIf(Stick(i).IsBot, Stick(0).LastPacketID, Stick(i).LastPacketID)
        .Speed = Stick(i).Speed
        .state = Stick(i).state
        .X = Stick(i).X
        .Y = Stick(i).Y
        
        .Health = Stick(i).Health
        .Shield = Stick(i).Shield
'        .Armour = Stick(i).Armour
'        .bAlive = Stick(i).bAlive
        .WeaponType = Stick(i).WeaponType
        '.PrevWeapon = Stick(i).PrevWeapon
        
'        .iKills = Stick(i).iKills
'        .iDeaths = Stick(i).iDeaths
'        .iKillsInARow = Stick(i).iKillsInARow
        
'        .Team = Stick(i).Team
        
        '.bSilenced = Stick(i).bSilenced
        '.bTyping = Stick(i).bTyping
'        .Perk = Stick(i).Perk
'        .MaskID = Stick(i).MaskID
        .iNadeType = Stick(i).iNadeType
        
        '.bFlashed = Stick(i).bFlashed
        '.bOnFire = Stick(i).bOnFire
    End With
    
    'Append
    sPacket = sPacket & mPacketToString() & UpdatePacketSep
Next i

sPacket = sUpdates & sPacket

'Send it to all non-local Stick
i = 1
'Ensure this isn't the local Stick
Do While i < NumSticks
    If Stick(i).IsBot = False Then
        'Send!
        If modWinsock.SendPacket(lSocket, Stick(i).SockAddr, sPacket) = False Then
            'If there was an error sending this mPacket, remove the Stick
            RemoveStick CInt(i)
            i = i - 1
        End If
    End If
    'Increment the counter
    i = i + 1
Loop

EH:
End Sub

'Extra Info
'#################################################################################################

Private Function SlowPacketToString() As String

With msPacket
    
    SlowPacketToString = _
        Trim$(.Name) & mPacketSep & _
        CStr(.colour) & mPacketSep & _
        CStr(.iKills) & mPacketSep & _
        CStr(.iDeaths) & mPacketSep & _
        CStr(.iKillsInARow) & mPacketSep & _
        CStr(.Team) & mPacketSep & _
        CStr(.bAlive) & mPacketSep & _
        CStr(.Perk) & mPacketSep & _
        CStr(.MaskID) & mPacketSep & _
        CStr(.ID) & mPacketSep & _
        Abs(.bSilenced) & mPacketSep & _
        Abs(.bTyping) & mPacketSep & _
        Abs(.bFlashed) & mPacketSep & _
        Abs(.bOnFire) & mPacketSep & _
        Abs(.bLightSaber) & mPacketSep & _
        Abs(.CurrentWeap1) & mPacketSep & _
        Abs(.CurrentWeap2) & mPacketSep & _
        CStr(.Burst_Bullets) & mPacketSep & _
        CStr(.Burst_Delay) & mPacketSep & _
        CStr(.PacketID) & mPacketSep
    
End With

End Function

Private Sub SlowPacketFromString(buf As String)

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With msPacket
    .Name = Trim$(Parts(0))
    .colour = CLng(Parts(1))
    .iKills = CInt(Parts(2))
    .iDeaths = CInt(Parts(3))
    .iKillsInARow = CInt(Parts(4))
    .Team = CInt(Parts(5))
    .bAlive = CBool(Parts(6))
    .Perk = CInt(Parts(7))
    .MaskID = CInt(Parts(8))
    .ID = CInt(Parts(9))
    .bSilenced = CBool(Parts(10))
    .bTyping = CBool(Parts(11))
    .bFlashed = CBool(Parts(12))
    .bOnFire = CBool(Parts(13))
    .bLightSaber = CBool(Parts(14))
    .CurrentWeap1 = CInt(Parts(15))
    .CurrentWeap2 = CInt(Parts(16))
    .Burst_Bullets = CInt(Parts(17))
    .Burst_Delay = CLng(Parts(18))
    .PacketID = CLng(Parts(19))
End With

Erase Parts
EH:
End Sub

Private Sub ProcessSlowPacket(ByVal sPacket As String)

Dim i As Integer, j As Integer
Dim sStick As String
Dim Sticks() As String

Sticks = Split(sPacket, UpdatePacketSep)


'Loop through each stick's info
For i = 0 To UBound(Sticks)
    On Error GoTo EH
    
    sStick = Sticks(i)
    If LenB(sStick) Then
        
        SlowPacketFromString sStick
        
        j = FindStick(msPacket.ID)
        
        'Does this stick already exist?
        If j = -1 Then
            'balls to him, let ProcessNormalUpdatePacket() handle it
            
            'If Not StickServer Then
                'new stick.  Make new spot and assign ID
                'j = AddStick()
                'Stick(j).ID = msPacket.ID
            'End If
            GoTo nextBit
        End If
        
        'Is this the local stick?
        If msPacket.ID <> Stick(0).ID Then
            
            With Stick(j)
                If .LastSlowPacketID < msPacket.PacketID Then
                    
                    .bAlive = msPacket.bAlive
                    .colour = msPacket.colour
                    .iDeaths = msPacket.iDeaths
                    
                    'if we are the server, don't accept kill updates
                    If Not modStickGame.StickServer Then
                        .iKills = msPacket.iKills
                    End If
                    
                    .iKillsInARow = msPacket.iKillsInARow
                    .MaskID = msPacket.MaskID
                    .Name = msPacket.Name
                    .Perk = msPacket.Perk
                    .Team = msPacket.Team
                    .bSilenced = msPacket.bSilenced
                    .bTyping = msPacket.bTyping
                    .bFlashed = msPacket.bFlashed
                    .bOnFire = msPacket.bOnFire
                    .bLightSaber = msPacket.bLightSaber
                    
                    .CurrentWeapons(1) = msPacket.CurrentWeap1
                    .CurrentWeapons(2) = msPacket.CurrentWeap2
                    
                    .Burst_Bullets = msPacket.Burst_Bullets
                    .Burst_Delay = msPacket.Burst_Delay
                    
                    .LastSlowPacketID = msPacket.PacketID
                End If
            End With
            
        ElseIf Not modStickGame.StickServer Then
            'us + we're not the server
            
            
            If Stick(0).LastSlowPacketID <= msPacket.PacketID Then
                'only let it update our kills
                If Stick(0).iKills < msPacket.iKills Then
                    Stick(0).iKills = msPacket.iKills
                End If
                
                'our deaths are always correct
                
                'if our killsinarow < mspacket's, then update, so long as we haven't just died
                If Stick(0).LastSpawnTime + 1000 < GetTickCount() Then
                    If Stick(0).iKillsInARow < msPacket.iKillsInARow Then
                        
                        Stick(0).iKillsInARow = msPacket.iKillsInARow
                        CheckKillsInARow
                        
                    End If
                End If
            End If
            
            
        End If 'myid/iKills endif
    End If 'lenb endif
    
nextBit:
Next i

EH:

LastUpdatePacket = GetTickCount()
End Sub

Private Sub SendSlowPacket()
Static LastSend As Long

'If it's not time to send a mPacket, exit sub
If LastSend + msPacket_SEND_DELAY < GetTickCount() Then
    
    'Reset the mPacket timer
    LastSend = GetTickCount()
    
    'Is this a StickServer mPacket, or a client mPacket?
    If StickServer Then
        'StickServer mPacket
        SendServerSlowPacket
    Else
        'Client mPacket
        SendClientSlowPacket
    End If
End If

End Sub

Private Sub SendClientSlowPacket()
Dim SendPacket As String

Stick(0).LastSlowPacketID = Stick(0).LastSlowPacketID + 1

With msPacket
    .colour = Stick(0).colour
    .Name = Stick(0).Name
    .bAlive = Stick(0).bAlive
    .iKills = Stick(0).iKills
    .iDeaths = Stick(0).iDeaths
    .iKillsInARow = Stick(0).iKillsInARow
    .Team = Stick(0).Team
    .Perk = Stick(0).Perk
    .MaskID = Stick(0).MaskID
    .ID = Stick(0).ID
    .bSilenced = Stick(0).bSilenced
    .bTyping = Stick(0).bTyping
    .bFlashed = Stick(0).bFlashed
    .bOnFire = Stick(0).bOnFire
    .bLightSaber = Stick(0).bLightSaber
    .CurrentWeap1 = Stick(0).CurrentWeapons(1)
    .CurrentWeap2 = Stick(0).CurrentWeapons(2)
    .Burst_Bullets = Stick(0).Burst_Bullets
    .Burst_Delay = Stick(0).Burst_Delay
    .PacketID = Stick(0).LastSlowPacketID
End With

SendPacket = sSlowUpdates & SlowPacketToString() & UpdatePacketSep

modWinsock.SendPacket lSocket, ServerSockAddr, SendPacket

End Sub

Private Sub SendServerSlowPacket()
Dim i As Long
Dim sPacket As String

On Error GoTo EH
Stick(0).LastSlowPacketID = Stick(0).LastSlowPacketID + 1



For i = 0 To NumSticksM1
    
    If Stick(i).IsBot Then
        Stick(i).LastSlowPacketID = Stick(i).LastSlowPacketID + 1
    End If
    
    
    With msPacket
        .colour = Stick(i).colour
        .Name = Stick(i).Name
        .bAlive = Stick(i).bAlive
        .iKills = Stick(i).iKills
        .iDeaths = Stick(i).iDeaths
        .iKillsInARow = Stick(i).iKillsInARow
        .Team = Stick(i).Team
        .Perk = Stick(i).Perk
        .MaskID = Stick(i).MaskID
        .ID = Stick(i).ID
        .bSilenced = Stick(i).bSilenced
        .bTyping = Stick(i).bTyping
        .bFlashed = Stick(i).bFlashed
        .bOnFire = Stick(i).bOnFire
        .bLightSaber = Stick(i).bLightSaber
        .CurrentWeap1 = Stick(i).CurrentWeapons(1)
        .CurrentWeap2 = Stick(i).CurrentWeapons(2)
        .Burst_Bullets = Stick(i).Burst_Bullets
        .Burst_Delay = Stick(i).Burst_Delay
        .PacketID = Stick(i).LastSlowPacketID
    End With
    
    'Append
    sPacket = sPacket & SlowPacketToString() & UpdatePacketSep
Next i

sPacket = sSlowUpdates & sPacket

'Send it to all non-local Stick
i = 1
Do While i < NumSticks
    'Ensure this isn't the local Stick
    If i > 0 Then
        If Stick(i).IsBot = False Then
            'Send!
            If modWinsock.SendPacket(lSocket, Stick(i).SockAddr, sPacket) = False Then
                'If there was an error sending this mPacket, remove the Stick
                RemoveStick CInt(i)
                i = i - 1
            End If
        End If
    End If
    'Increment the counter
    i = i + 1
Loop

EH:
End Sub


'###############################################################################################
'Map Stuff #####################################################################################
'###############################################################################################
Public Function LoadMapEx(sFileName As String) As Boolean

If LoadMap(sFileName) Then
    InitVarsForMap
    
    modStickGame.StickMapPath = sFileName
    
    LoadMapEx = True
End If

End Function

Public Function LoadMap(sFileName As String) As Boolean

Dim i As Integer, f As Integer ', j As Integer, K As Integer
Dim sFile As String, sPlatforms As String, sBoxes As String, sTBoxes As String, sHealthPack As String
Dim arTmp() As String, arParts() As String

f = FreeFile()

On Error GoTo EH
Open sFileName For Input As #f
    sFile = input(LOF(f), f)
Close #f

arTmp = Split(sFile, "#")
sPlatforms = arTmp(0)
sBoxes = Mid$(arTmp(1), 3)
sTBoxes = Mid$(arTmp(2), 3)
sHealthPack = Mid$(arTmp(3), 3)

Erase arTmp

'##########################################

arTmp = Split(sPlatforms, vbNewLine)
ubdPlatforms = UBound(arTmp)
ReDim Preserve Platform(ubdPlatforms)

With Platform(0)
    '-1000|20000|52000|853.3
    .Left = -10
    .Top = StickGameHeight - 100
    .width = StickGameWidth + 10
    .height = 850
End With

For i = 0 To ubdPlatforms - 1
    
    arParts = Split(arTmp(i), MapSep)
    
    With Platform(i + 1)
        .Left = CSng(arParts(0))
        .Top = CSng(arParts(1))
        .width = CSng(arParts(2))
        .height = CSng(arParts(3))
    End With
    
    Erase arParts
Next i

Erase arTmp

'##########################################

arTmp = Split(sBoxes, vbNewLine)
ubdBoxes = UBound(arTmp) - 1
ReDim Preserve Box(ubdBoxes)
For i = 0 To ubdBoxes
    
    arParts = Split(arTmp(i), MapSep)
    
    With Box(i)
        .Left = CSng(arParts(0))
        .Top = CSng(arParts(1))
        .width = CSng(arParts(2))
        .height = CSng(arParts(3))
        .bInUse = True
    End With
    
    Erase arParts
Next i

Erase arTmp

'##########################################

arTmp = Split(sTBoxes, vbNewLine)
ubdtBoxes = UBound(arTmp) - 1
ReDim Preserve tBox(ubdtBoxes)
For i = 0 To ubdtBoxes
    
    arParts = Split(arTmp(i), MapSep)
    
    With tBox(i)
        .Left = CSng(arParts(0))
        .Top = CSng(arParts(1))
        .width = CSng(arParts(2))
        .height = CSng(arParts(3))
    End With
    
    Erase arParts
Next i

Erase arTmp

'##########################################

i = InStr(1, sHealthPack, MapSep)
HealthPackX = CLng(Left$(sHealthPack, i - 1))
HealthPackY = CLng(Mid$(sHealthPack, i + 1))

If Not modStickGame.bStickEditing Then
    HealthPack.X = HealthPackX
    HealthPack.Y = HealthPackY
End If

'##########################################

LoadMap = True
Exit Function
EH:
LoadMap = False
Close #f
End Function

Public Function SaveMap(sFileName As String) As Boolean
Dim i As Integer, f As Integer

f = FreeFile()

On Error GoTo EH
Open sFileName For Output As #f
    
    For i = 0 To modStickGame.ubdPlatforms
        Print #f, Platform(i).Left & MapSep & _
                  Platform(i).Top & MapSep & _
                  Platform(i).width & MapSep & _
                  Platform(i).height
    Next i
    
    
    Print #f, "#"
    
    
    For i = 0 To modStickGame.ubdBoxes
        Print #f, Box(i).Left & MapSep & _
                  Box(i).Top & MapSep & _
                  Box(i).width & MapSep & _
                  Box(i).height
    Next i
    
    
    Print #f, "#"
    
    
    For i = 0 To modStickGame.ubdtBoxes
        Print #f, tBox(i).Left & MapSep & _
                  tBox(i).Top & MapSep & _
                  tBox(i).width & MapSep & _
                  tBox(i).height
    Next i
    
    
    Print #f, "#"
    
    Print #f, HealthPackX & MapSep & HealthPackY
    
Close #f



SaveMap = True
Exit Function
EH:
SaveMap = False
Close #f
End Function

'###############################################################################################
'###############################################################################################
'###############################################################################################
'###############################################################################################
'Map Editing ###################################################################################
'###############################################################################################
'###############################################################################################
'###############################################################################################
'###############################################################################################


'Because some lightweight controls do not have a MouseDown event,
'when we get a MouseDown event on a form, we do a scan of the
'Controls collection to see if any lightweight controls are under
'the mouse. Note that this code does not work for controls within
'containers. Also, if no control is under the mouse, then we
'remove the sizing handles and clear the current control.

'mousedown

'To handle all mouse message anywhere on the form, we set the mouse
'capture to the form. Mouse movement is processed here

'mousemove

'To handle all mouse message anywhere on the form, we set the mouse
'capture to the form. Mouse up is processed here

'mouseup

'Process MouseDown over handles
Private Sub picHandle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'ReleaseCapture
'MsgBox "You can't resize the boxes", vbExclamation, "Warning"
'End Sub

Dim i As Integer
Dim rc As RECT

If m_CurrCtl Is Nothing Then
    ShowHandles False
ElseIf m_CurrCtl.Name <> shHealthPack.Name Then
    'Handles should only be visible when a control is selected
    'Debug.Assert (Not m_CurrCtl Is Nothing)
    
    'NOTE: m_DragPoint not used for sizing
    'Save control position in screen coordinates
    m_DragRect.SetRectToCtrl m_CurrCtl
    m_DragRect.TwipsToScreen m_CurrCtl
    
    'Track index handle
    m_DragHandle = Index
    
    'Hide sizing handles
    ShowHandles False
    
    'We need to force handles to hide themselves before drawing drag rectangle
    Refresh
    
    'Indicate sizing is under way
    m_DragState = StateSizing
    
    'Show sizing rectangle
    DrawDragRect
    
    'In order to detect mouse movement over any part of the form,
    'we set the mouse capture to the form and will process mouse
    'movement from the applicable form events
    SetCapture hWnd
    
    'Limit cursor movement within form
    GetWindowRect hWnd, rc
    ClipCursor rc
    
    map_Changed = True
End If

End Sub

'Display or hide the sizing handles and arrange them for the current rectangld
Private Sub ShowHandles(Optional bShowHandles As Boolean = True)
Dim i As Integer
Dim xFudge As Long, yFudge As Long
Dim nWidth As Long, nHeight As Long

If bShowHandles And Not m_CurrCtl Is Nothing Then
    With m_DragRect
        'Save some calculations in variables for speed
        nWidth = (picHandle(0).width \ 2)
        nHeight = (picHandle(0).height \ 2)
        xFudge = (0.5 * Screen.TwipsPerPixelX)
        yFudge = (0.5 * Screen.TwipsPerPixelY)
        'Top Left
        picHandle(0).Move (.Left - nWidth) + xFudge, (.Top - nHeight) + yFudge
        'Bottom right
        picHandle(4).Move (.Left + .width) - nWidth - xFudge, .Top + .height - nHeight - yFudge
        'Top center
        picHandle(1).Move .Left + (.width / 2) - nWidth, .Top - nHeight + yFudge
        'Bottom center
        picHandle(5).Move .Left + (.width / 2) - nWidth, .Top + .height - nHeight - yFudge
        'Top right
        picHandle(2).Move .Left + .width - nWidth - xFudge, .Top - nHeight + yFudge
        'Bottom left
        picHandle(6).Move .Left - nWidth + xFudge, .Top + .height - nHeight - yFudge
        'Center right
        picHandle(3).Move .Left + .width - nWidth - xFudge, .Top + (.height / 2) - nHeight
        'Center left
        picHandle(7).Move .Left - nWidth + xFudge, .Top + (.height / 2) - nHeight
    End With
End If

'Show or hide each handle
For i = 0 To 7
    picHandle(i).Visible = bShowHandles
Next i
End Sub

'Draw drag rectangle. The API is used for efficiency and also
'because drag rectangle must be drawn on the screen DC in
'order to appear on top of all controls
Private Sub DrawDragRect()
Dim hPen As Long, hOldPen As Long
Dim hBrush As Long, hOldBrush As Long
Dim hScreenDC As Long, nDrawMode As Long

'Get DC of entire screen in order to
'draw on top of all controls
hScreenDC = GetDC(0)

'Select GDI object
hPen = CreatePen(PS_SOLID, 2, 0)
hOldPen = SelectObject(hScreenDC, hPen)
hBrush = GetStockObject(NULL_BRUSH)
hOldBrush = SelectObject(hScreenDC, hBrush)
nDrawMode = SetROP2(hScreenDC, R2_NOT)

'Draw rectangle
Rectangle hScreenDC, m_DragRect.Left, m_DragRect.Top, _
    m_DragRect.Right, m_DragRect.Bottom

'Restore DC
SetROP2 hScreenDC, nDrawMode
SelectObject hScreenDC, hOldBrush
SelectObject hScreenDC, hOldPen
ReleaseDC 0, hScreenDC

'Delete GDI objects
DeleteObject hPen
End Sub

'=========================== Sample controls ===========================
'To drag a control, simply call the DragBegin function with
'the control to be dragged
'=======================================================================

'Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbLeftButton And m_bDesignMode Then
'        DragBegin Label1
'    End If
'End Sub

'========================== Dragging Code ================================

'Initialization -- Do not call more than once
Private Sub DragInit()
Dim i As Integer, xHandle As Single, yHandle As Single
Const Handle_Size = 3

'Use black Picture box controls for 8 sizing handles
'Calculate size of each handle
xHandle = Handle_Size * Screen.TwipsPerPixelX
yHandle = Handle_Size * Screen.TwipsPerPixelY
'Load array of handles until we have 8
For i = 0 To 7
    If i Then
        Load picHandle(i)
    End If
    picHandle(i).width = xHandle
    picHandle(i).height = yHandle
    'Must be in front of other controls
    picHandle(i).ZOrder
Next i
'Set mousepointers for each sizing handle
picHandle(0).MousePointer = vbSizeNWSE
picHandle(1).MousePointer = vbSizeNS
picHandle(2).MousePointer = vbSizeNESW
picHandle(3).MousePointer = vbSizeWE
picHandle(4).MousePointer = vbSizeNWSE
picHandle(5).MousePointer = vbSizeNS
picHandle(6).MousePointer = vbSizeNESW
picHandle(7).MousePointer = vbSizeWE
'Initialize current control
Set m_CurrCtl = Nothing

'm_bDesignMode = True
End Sub

'Drags the specified control
Private Sub DragBegin(ctl As Control)
Dim rc As RECT

'Hide any visible handles
ShowHandles False
'Save reference to control being dragged
Set m_CurrCtl = ctl
'Store initial mouse position
GetCursorPos m_DragPoint
'Save control position (in screen coordinates)
'Note: control might not have a window handle
m_DragRect.SetRectToCtrl m_CurrCtl
m_DragRect.TwipsToScreen m_CurrCtl
'Make initial mouse position relative to control
m_DragPoint.X = m_DragPoint.X - m_DragRect.Left
m_DragPoint.Y = m_DragPoint.Y - m_DragRect.Top
'Force redraw of form without sizing handles
'before drawing dragging rectangle
Refresh
'Show dragging rectangle
DrawDragRect
'Indicate dragging under way
m_DragState = StateDragging
'In order to detect mouse movement over any part of the form,
'we set the mouse capture to the form and will process mouse
'movement from the applicable form events
ReleaseCapture  'This appears needed before calling SetCapture
SetCapture hWnd
'Limit cursor movement within form
GetWindowRect hWnd, rc
ClipCursor rc

map_Changed = True
    
End Sub

'Clears any current drag mode and hides sizing handles
Public Sub DragEnd()
Set m_CurrCtl = Nothing
ShowHandles False
m_DragState = StateNothing
End Sub

Public Sub setMapChanged(b As Boolean)
map_Changed = b
End Sub
