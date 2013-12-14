VERSION 5.00
Begin VB.Form frmStickGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "Stick Shooter"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
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
      Left            =   6000
      Picture         =   "frmStickGame.frx":0000
      ScaleHeight     =   2040
      ScaleWidth      =   1635
      TabIndex        =   2
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
      TabIndex        =   1
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
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   8
      Left            =   14640
      Top             =   6000
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   7
      Left            =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   6
      Left            =   11400
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   11
      Left            =   13200
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   375
      Index           =   5
      Left            =   13095
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   6
      Left            =   13080
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   10
      Left            =   13920
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   9
      Left            =   15240
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   5
      Left            =   12120
      Top             =   8760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   495
      Index           =   4
      Left            =   13080
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   495
      Index           =   3
      Left            =   9600
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   495
      Index           =   2
      Left            =   5400
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   495
      Index           =   1
      Left            =   8530
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape otBox 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   495
      Index           =   0
      Left            =   7200
      Top             =   8280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1695
      Index           =   8
      Left            =   9600
      Top             =   4200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   7
      Left            =   7440
      Top             =   9120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   6
      Left            =   850
      Top             =   3840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1575
      Index           =   5
      Left            =   7200
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   4
      Left            =   3840
      Top             =   7680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   3
      Left            =   13080
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   2
      Left            =   8760
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   1
      Left            =   6250
      Top             =   5280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oBox 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1095
      Index           =   0
      Left            =   5640
      Top             =   3360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   4
      Left            =   8520
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   3
      Left            =   840
      Top             =   4920
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   6240
      Top             =   6360
      Visible         =   0   'False
      Width           =   10095
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Shape oPlatform 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   855
      Index           =   0
      Left            =   -120
      Top             =   10200
      Visible         =   0   'False
      Width           =   15375
   End
End
Attribute VB_Name = "frmStickGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'kills
Private Enum eKillTypes
    kNormal = 0
    kHead = 1
    kNade = 2
    kRPG = 3
    kKnife = 4
    kMine = 5
    kChoppered = 6
    kFlame = 7
    kBurn = 8
    kSilenced = 9
    kCrushed = 10
    kFlameTag = 11
    kLightSaber = 12
End Enum
Private Enum eMagTypes
    mAK = 0
    mSCAR = 1
    mSniper = 2
    mPistol = 3
    mFlameThrower = 4
    mSA80 = 5
End Enum


Private Type ptBullet
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    'Facing As Single
    Decay As Long
    Owner As Integer
    'Colour As Long
    Damage As Integer 'Single
    LastDiffract As Long
    bSniperBullet As Boolean
    bShotgunBullet As Boolean
    
    bSilenced As Boolean
    LastSmoke As Long
End Type

Private Type ptBlood
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    Decay As Long
    bArmour As Boolean
End Type

Private Type ptNade
    X As Single
    Y As Single
    Decay As Long 'Boom
    Heading As Single
    Speed As Single
    OwnerID As Integer
    IsRPG As Boolean
    LastSmoke As Long
    LastGravity As Long
    Colour As Long
    iType As eNadeTypes
End Type

Private Type ptMine
    X As Single
    Y As Single
    OwnerID As Integer
    Colour As Long
    
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
    Colour As Long
    Decay As Long
    bOnSurface As Boolean
    
    Speed As Single
    Heading As Single
    
    LastGravity As Long
    
    bFacingRight As Single
    
    bFlamed As Boolean
End Type

Private Type ptDeadChopper
    X As Single
    Y As Single
    Colour As Long
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

Private Type ptSmallSmoke
    'X As Single
    'Y As Single
    'Heading As Single
    'Speed As Single
    AngleFromMain As Single
    DistanceFromMain As Single
    
    sAspect As Single
    AspectDir As Integer
    
    DistanceFromMainInc As Single
End Type
Private Type ptLargeSmoke
    CentreX As Single
    CentreY As Single
    
    SingleSmoke(1 To 10) As ptSmallSmoke
    
    iSize As Single
    iDirection As Integer
    
    'pPoly(1 To 10) As POINTAPI
End Type

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
Private Type Circ
    X As Single
    Y As Single
    Prog As Single
    MaxProg As Single
    Radius As Single
    Colour As Long
    Direction As Integer 'explode or implode?
    'Speed As Single
    'Heading As Single
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

Private Circs() As Circ
Private NumCircs As Integer

Private NumBullets As Long
Private Bullet() As ptBullet

Private NumBlood As Long
Private Blood() As ptBlood

Private NumNades As Long
Private Nade() As ptNade

Private NumCasings As Long
Private Casing() As ptCasing

Private NumMines As Long
Private Mine() As ptMine

Private NumDeadSticks As Long
Private DeadStick() As ptDeadStick

Private NumMags As Long
Private Mag() As ptMagazine

Private NumDeadChoppers As Long
Private DeadChopper() As ptDeadChopper

Private NumSparks As Long
Private Spark() As ptSpark

Private NumFlames As Long
Private Flame() As ptFlame

Private NumStaticWeapons As Long
Private StaticWeapon() As ptStaticWeapon

Private NumLargeSmokes As Integer
Private LargeSmoke() As ptLargeSmoke

Private NumWallMarks As Long
Private WallMark() As ptWallMark

Private NumSmokeBlasts As Long
Private SmokeBlast() As ptSmokeBlast

'optimization stuff
Private NumSticksM1 As Integer


'angle stuff
Private Const SmallAngle = pi / 4
'end angle


'stats
Private Const Health_Start = 100

Private Const Accel = 2
Private Const Max_Speed = 112
Private Const JumpMultiple = 75 'move stick up by Accel*JumpMultiple
Private Const NadeMultiple = 500 'force stick away by accel*nademultiple

Private Const Gravity_Strength = 8
Private Const Gravity_Direction = pi
Private Const Gravity_Delay = 60
'Private Const JumpTime = 100

Private Const Blood_Time = 500
Private Const Casing_Time = 3500
Private Const Casing_Len = 25

Private Const DeadStickTime = 25000


Private Const Lim = 50
Private Const Left_Indent = 15000 'for bot coop positioning
'end stats


'Weapon stats
Private AmmoFired(0 To eWeaponTypes.Chopper) As Integer
Private kBulletDelay(0 To eWeaponTypes.Chopper) As Integer
Private kMaxRounds(0 To eWeaponTypes.Chopper) As Integer
Private kReloadTime(0 To eWeaponTypes.Chopper) As Integer
Private kPerkName(0 To eStickPerks.pSpy) As String
Private kWeaponName(0 To eWeaponTypes.Chopper) As String
Private kTeamColour(0 To eTeams.Spec) As Long
Private kGameType(0 To eStickGameTypes.gCoOp) As String
Private kRecoverAmount(0 To eWeaponTypes.Knife) As Single
Private kRecoilTime(0 To eWeaponTypes.Knife) As Long
Private kRecoilForce(0 To eWeaponTypes.Knife) As Boolean
Private kNadeName(0 To eNadeTypes.nSmoke) As String


Private Const Shotgun_Gauge = 12
Private Const Shotgun_Spray_Angle = pi / 17
Private Const Shotgun_Recoil_Time = 400
Private Const Shotgun_SingleRecoil_Angle = SmallAngle
Private Const Shotgun_Recover_Amount = Shotgun_SingleRecoil_Angle / 14 '15.5 '26
Private Const Shotgun_Bullet_Delay = 400
Private Const Shotgun_Bullets = 8
Private Const Shotgun_Reload_Time = 2000
Private Const Shotgun_Bullet_Damage = 100 / Shotgun_Gauge 'was 7 - will kill if all shots hit
Private Const Shotgun_RecoilForce = 10
Private Const Shotgun_Mags = 12

Private Const AK_Spray_Angle = pi / 75
Private Const AK_Recoil_Time = 50
Private Const AK_SingleRecoil_Angle = SmallAngle / 90
Private Const AK_Recover_Amount = AK_SingleRecoil_Angle / AK_Recoil_Time
Private Const AK_Bullet_Delay = 70 '600 rpm
Private Const AK_Bullets = 30
Private Const AK_Reload_Time = 900
Private Const AK_Bullet_Damage = 16 '2010 joules
Private Const AK_Mags = 5

Private Const SA80_Spray_Angle = pi / 250
Private Const SA80_Recoil_Time = 10
Private Const SA80_SingleRecoil_Angle = SmallAngle / 310
Private Const SA80_Recover_Amount = SA80_SingleRecoil_Angle / SA80_Recoil_Time
Private Const SA80_Bullet_Delay = 175
Private Const SA80_Single_Bullet_Delay = 50 '775 rpm
Private Const SA80_Bullets = 30
Private Const SA80_Burst_Bullets = 3
Private Const SA80_Reload_Time = 900
Private Const SA80_Bullet_Damage = 17
'1775 joules - more due to burst fire - (17*3)*2>100
Private Const SA80_Mags = 5

Private Const M82_Recoil_Time = 450
Private Const M82_SingleRecoil_Angle = SmallAngle / 1.5
Private Const M82_Recover_Amount = M82_SingleRecoil_Angle / 15
Private Const M82_Bullet_Delay = 400
Private Const M82_Bullets = 4
Private Const M82_Reload_Time = 3250
Private Const M82_Bullet_Damage = 75 ''''''''''half the health and take off <--
Private Const M82_RecoilForce = 36
Private Const M82_Wall_Damage = 25
Private Const M82_Mags = 3

Private Const SCAR_Spray_Angle = pi / 200
Private Const SCAR_Recoil_Time = 15
Private Const SCAR_SingleRecoil_Angle = SmallAngle / 300
Private Const SCAR_Recover_Amount = SCAR_SingleRecoil_Angle / SCAR_Recoil_Time
Private Const SCAR_Bullet_Delay = 55 '750 rpm
Private Const SCAR_Bullets = 30
Private Const SCAR_Reload_Time = 500
Private Const SCAR_Bullet_Damage = 13 '1775 joules
Private Const SCAR_Mags = 5

Private Const RPG_Recoil_Time = 1500
Private Const RPG_SingleRecoil_Angle = piD4
Private Const RPG_Recover_Amount = RPG_SingleRecoil_Angle / 60
Private Const RPG_Bullet_Delay = 1500
Private Const RPG_Bullets = 1
Private Const RPG_Reload_Time = 2500
Private Const RPG_Smoke_Delay = 15
Private Const RPG_RecoilForce = 12
Private Const RPG_Speed = 250
Private Const RPG_Mags = 4

Private Const M249_Spray_Angle = pi / 60
Private Const M249_Recoil_Time = 50
Private Const M249_SingleRecoil_Angle = SmallAngle / 90
Private Const M249_Recover_Amount = M249_SingleRecoil_Angle / M249_Recoil_Time
Private Const M249_Bullet_Delay = 33 '75
Private Const M249_Bullets = 300
Private Const M249_Reload_Time = 5000
Private Const M249_Bullet_Damage = 12
'Private Const M249_RecoilForce = 1 'Doesn't give game a chance to stop stick moving
Private Const M249_Mags = 3

Private Const DEagle_Recoil_Time = 300
Private Const DEagle_SingleRecoil_Angle = SmallAngle
Private Const DEagle_Recover_Amount = DEagle_SingleRecoil_Angle / 14 '17
Private Const DEagle_Bullet_Delay = 300
Private Const DEagle_Bullets = 7
Private Const DEagle_Reload_Time = 600
Private Const DEagle_Bullet_Damage = 90
Private Const DEagle_RecoilForce = 12
Private Const DEagle_Mags = 8

Private Const Flame_Speed = 80
Private Const Flame_Bullet_Delay = 35
Private Const Flame_Bullets = 50
Private Const Flame_Reload_Time = 2000
Private Const Flame_Time = 900
Private Const Flame_Radius = 20
Private Const Flame_Damage = 1 'burn does proper damage
Private Const Flame_Burn_Time = 5000 'time that a flame'll burn after touch
Private Const Flame_Burn_Damage = 4 'totaldamage = 6*this 'damage to apply to a burn
Private Const Flame_Burn_Damage_Time = 500 'apply above damage every x milliseconds
Private Const Flame_Burn_Radius = 100
Private Const Flame_Inertia_Reduction = 2
Private Const Flamethrower_Mags = 5

Private Const Knife_Delay = 100

Private Const Throwing_Strength = 125

Private Const Nade_Explode_Radius = 2300
Private Const Nade_Radius = 50
Private Const Nade_Time = 2000 'time until BOOM
Private Const Nade_Delay = 5000 'time until can throw next nade
Private Const Nade_Bounce_Reduction = 1.3 'non-elastic
'--------------------------------------------------------------------------
Private Const Mine_Radius = 4
Private Const Mine_Explode_Radius = 3000
Private Const Mine_Delay = 15000
Private Const Mine_Y_Increase As Single = 570 '570.06 'BodyLen + HeadRadius * 3.5
'Private Const Mine_Hold_Time = 1000

Private Const Casing_Bounce_Reduction = 1.5
Private Const MFlash_Time = 25

Private Const Nade_Release_Delay = 1000 'time until state is removed from said stick
Private Const Bullet_Release_Delay = 200 'make sure the shot gets through
Private FireKeyUpTime As Long
Private Const AutoReload_Delay = 400

Private Const SwitchWeaponDelay = 500
Private Const UseKeyReleaseDelay = 300
Private Const StaticWeaponSendDelay = 2000

Private Const Sniper_Smoke_Delay = 25

Private Const Hardcore_Damage_Amp = 2 '1.5
'end weapon stats
'##################################################################
'my stats
Private Const RowKillsForArmour = 3
Private Const Armour_Colour = MSilver 'vbBlack
Private Const Max_Armour = 100

Private Const Radar_Time = 30000
Private Const RowKillsForRadar = 4
Private RadarStartTime As Long
Private bHadRadar As Boolean 'for displaying "Radar Expired"

Public ChopperAvail As Boolean
Private Const RowKillsForChopper = 6

Private Const RowFlameKillsForToasty = 3

'Private KillsInARow As Integer
Private FlamesInARow As Integer

Private KnifesInARow As Integer
Private Const KnivesForSaber = 3
'end my stats
'##################################################################
'other stuff
Private LastWeaponSwitch As Long
Private Const WeaponSwitchDelay = 500

Private Const FlashBang_Time = 10000
'end other stuff
'##################################################################

'chopper stats
Private Const Chopper_Max_Speed = 75
Private Const Chopper_Lift = Accel * 2
Private Const Chopper_RPG_Delay = 2500
Private Const DeadChopperTime = 30000

Private Const DeadChopper_Smoke_Delay = 34

Private Const Chopper_Bullet_Damage = 20
Private Const Chopper_Bullet_Delay = 100
Private Const Chopper_Damage_Reduction = 35 '15 = 4 sniper shots JUST kill it
Private Const ChopperLen = 3500, _
    CLD2 = ChopperLen / 2, _
    CLD3 = ChopperLen / 3, _
    CLD4 = ChopperLen / 4, _
    CLD6 = ChopperLen / 6, _
    CLD8 = ChopperLen / 8, _
    CLD10 = ChopperLen / 10
'end chopper stats

'drawing
Private Const StickSize As Integer = 800
Private Const HeadRadius As Integer = StickSize \ 10
Private Const BodyLen As Integer = HeadRadius * 4
Private Const ArmLen = HeadRadius * 2
Private Const ArmNeckDist As Integer = 250
Private Const LegHeight As Integer = StickSize \ 3
Private Const MaxLegWidth As Integer = 90
Private Const StickHeight As Integer = 1100

Private Const Max_Health As Integer = 100

Private Const Bullet_Radius As Integer = 5
Private Const Bullet_Decay As Integer = 1000
Private Const Bullet_Damage As Integer = 3
Private Const BULLET_SPEED = Max_Speed * 2 '=222
Private Const BULLET_LEN = StickHeight \ 15
Private Const Bullet_Diffract_Delay = 400
Private Const Bullet_Min_Speed = 35

Private Const Mag_Decay = 8000

Private Const GunLen = 213 'BodyLen / 1.5

Private Const Spark_Time = 2000
Private Const Spark_Diffraction = piD3, Spark_Speed = 30
Private Const Spark_Min_Speed = 5, Spark_Speed_Reduction = 1.02
Private Const Spark_Speed_Reduction_Delay = 35
Private Const WallMark_Time = 60000
Private Const WallMark_Bullet_Radius = 30, WallMark_Explosion_Radius = WallMark_Bullet_Radius * 3

Private Const Degrees10 = pi * 1 / 18
Private Const ProneRightLimit = piD2 + Degrees10, ProneLeftLimit = pi3D2 - Degrees10
'end drawing

'Private LastSpawnTime As Long
Private Const Spawn_Invul_Time = 2000 'time they are invulnerable for

Private Const Max_Chat = 24
Private FPS As Integer

Private Const mPacket_LAG_TOL = 1000  'Milliseconds to wait before rendering a stick motionless
Private Const mPacket_LAG_KILL = 7000    'Milliseconds to wait before removing a stick due to lack of info
Private Const StickServer_RETRY_FREQ = 2000   'Milliseconds between attempts to connect to StickServer
Private Const StickServer_NUM_RETRIES = 5
Private Const ServerVarSendDelay = 10000
Private Const BoxInfoDelay = 7000
Private Const LagOut_Delay = mPacket_LAG_TOL * 3 'time to lag out
Private Const NameCheckDelay = 10000
Private LastUpdatePacket As Long 'Are we lagging out?

Private WindowClosing As Boolean

Public MyID As Integer 'Which Stick are we?
Private LastScoreCheck As Long


'########################################################################
'packet stuff

Private Type ptStickPacket
    ID As Integer
    PacketID As Long
    
    ActualFacing As Single
    Facing As Single
    Heading As Single
    Speed As Single
    X As Single
    Y As Single
    State As Integer
    
    WeaponType As eWeaponTypes
    PrevWeapon As eWeaponTypes
    Health As Integer
    
    iNadeType As eNadeTypes
End Type
Private Type ptStickSlowPacket
    Name As String * 15
    ID As Integer
    Colour As Long
    Armour As Integer
    
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
End Type

Private mPacket As ptStickPacket
Private msPacket As ptStickSlowPacket

Private Const mPacket_SEND_DELAY = 65 'Milliseconds between update packets
Private Const msPacket_SEND_DELAY = 300 'Milliseconds between slow update packets
Private Const Client_Nade_Delay = mPacket_SEND_DELAY + 25
'########################################################################

Public bRunning As Boolean  'Is the render loop running?
Private bPlaying As Boolean 'In the middle of a game?
Private PacketTimer As Long 'Time at which last mPacket was sent
Private ServerSockAddr As ptsockaddr   'StickServer's sock addr
Public socket As Long       'Socket with which we'll send/receive essages

Private Type CHATTYPE
    Text As String           'The string to display
    Decay As Long            'When is this one removed?
    Colour As Long
    bChatMessage As Boolean
End Type

Private Chat() As CHATTYPE       'Our chat array
Private NumChat As Long          'How many chat messages are there currently?
Private Const CHAT_DECAY = 15000        'How long before chat messages disappear?

'big message(s)
Private Type ptMainMessage
    Text As String           'The string to display
    Decay As Long            'When is this one removed?
    Colour As Long
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

Private Const SmokeOutline = &HDDDDDD
Private Const SmokeFill = &HE1E1E1
Private Const BoxCol As Long = &HC0C0C0

'##############################################################################

Private Type ptServerVars
    bAllowRockets As Boolean
    bAllowFlameThrowers As Boolean
    bAllowChoppers As Boolean
    
    bShootNades As Boolean
    bHPBonus As Boolean
    
    iGameType As eGameTypes
    sgGameSpeed As Single
    bHardCore As Boolean
    b2Weapons As Boolean
    
    iScoreToWin As Integer
End Type

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
Private Const JuggernautDamageReduction = 3
Private Const StoppingPowerIncrease = 2.4
Private Const SleightOfHandReloadDecrease = 3

Private Const ConditiongSpeedIncrease = 2
Private Const ConditioningMaxSpeedInc = 1.35

Private Const StealthESPDist = StickGameWidth \ 2

'##############################################################################

'Round Stuff
Private Const RoundInfoSendDelay = 2500
Private Const ScoreCheckDelay = 2000
Private Const RoundWaitTime = 10000
Private Const PresenceSendDelay = 1000
Private RoundWinnerID As Integer
Private RoundPausedAtThisTime As Long

'##############################################################################
'resize 'constants'
Private RadarLeft As Single ': RadarLeft = Me.width - RadarWidth - 100
Private PlayingX As Single ': PlayingX = StickCentreX - 600
Private ConnectingkX As Single ': kX = StickCentreX - 900
Private ConnectingkY As Single ': kY = StickCentreY + 650

Private Const RadarWidth = 2000

'##############################################################################


Private MouseX As Single, MouseY As Single
Private StunnedMouseX As Single, StunnedMouseY As Single

Private Const sUpdates As String * 1 = "U"
Private Const sJoins As String * 1 = "J"
Private Const sAccepts As String * 1 = "A"
Private Const sChats As String * 1 = "C"
'Private Const sExits As String * 1 = "E"
Private Const sBoxInfos As String * 1 = "B"
Private Const sServerVarss As String * 1 = "S"
'Private Const sKicks As String * 1 = "K"
Private Const sKillInfos As String * 1 = "I"
Private Const sDeathInfos As String * 1 = "D"
Private Const sHealthPacks As String * 1 = "H"
Private Const sStaticWeaponUpdates As String * 1 = "W"
Private Const sRoundInfos As String * 1 = "R"
Private Const sPresences As String * 1 = "P"
Private Const sSlowUpdates As String * 1 = "L"

Private Const Ally_Colour = vbGreen, Enemy_Colour = vbBlack

'FGIMNOQTVXYZ

Private LeftKey As Boolean, RightKey As Boolean, JumpKey As Boolean, CrouchKey As Boolean, ProneKey As Boolean, _
    ReloadKey As Boolean, FireKey As Boolean, ShowScoresKey As Boolean, MineKey As Boolean

Private SpecUp As Boolean, SpecDown As Boolean, SpecLeft As Boolean, SpecRight As Boolean

Private UseKey As Boolean

'Private Const ControlKey = 17
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private WeaponKey As eWeaponTypes
Private Scroll_WeaponKey As eWeaponTypes
Private LastScrollWeaponSwitch As Long
Private Const Scroll_Delay = 750

'zoom
Private LastZoomPress As Long
Private Const ZoomShowTime = 750
Private Const ZoomInc = 0.05
Private Const MinZoom = 0.9 + ZoomInc
Private Const MaxZoom = 2

Private Sub ProcessAmmoPickups()
Dim i As Integer, bTold As Boolean

GenerateAmmoPickup

If Stick(0).WeaponType < Knife Then
    
    For i = 0 To NumAmmoPickUpsM1
        If AmmoPickup(i).bActive Then
            If CoOrdInStick(AmmoPickup(i).X, AmmoPickup(i).Y, 0) Then
                
                If StickiHasState(0, stick_crouch) Then
                    
                    TotalMags(Stick(0).WeaponType) = GetTotalMags(Stick(0).WeaponType)
                    
                    AmmoPickup(i).LastUsed = GetTickCount()
                    AmmoPickup(i).bActive = False
                    
                ElseIf bTold = False Then
                    
                    PrintStickText "Crouch to Pick Up " & GetMagName(Stick(0).WeaponType), Stick(0).X + 750, Stick(0).Y - 500, vbRed
                    
                    bTold = True
                End If
                
            End If
            
        End If
    Next i
    
End If


End Sub

Private Sub GenerateAmmoPickup()
Dim i As Integer, iPlatform As Integer
Dim GTC As Long

GTC = GetTickCount()

For i = 0 To NumAmmoPickUpsM1
    If AmmoPickup(i).bActive = False Then
        If AmmoPickup(i).LastUsed + AmmoPickUp_Spawn_Delay < GTC Then
            
            iPlatform = Rnd() * modStickGame.nPlatforms
            'between 0 and 7
            
            AmmoPickup(i).Y = Platform(iPlatform).Top - 100
            AmmoPickup(i).X = (Platform(iPlatform).Left + 300) + Rnd() * (Platform(iPlatform).width - 300)
            
            AmmoPickup(i).bActive = True
            AmmoPickup(i).LastUsed = GTC
            
        End If
    End If
Next i

End Sub

Private Sub DrawAmmoPickups()
Dim i As Integer

picMain.FillStyle = vbFSSolid
picMain.FillColor = vbBlack
picMain.DrawWidth = 2
For i = 0 To NumAmmoPickUpsM1
    If AmmoPickup(i).bActive Then
        
        modStickGame.sBox AmmoPickup(i).X, AmmoPickup(i).Y, _
            AmmoPickup(i).X + HealthPack_Radius * 2, AmmoPickup(i).Y + HealthPack_Radius, _
            vbBlack
        
        
        modStickGame.PrintStickText "Ammo Pickup", AmmoPickup(i).X - 300, AmmoPickup(i).Y - 200, vbBlack
        
    End If
Next i

picMain.FillStyle = vbFSTransparent

End Sub

Private Function GetTotalMags(vWeapon As eWeaponTypes) As Byte
'can be slow code here

If vWeapon = AK Then
    GetTotalMags = AK_Mags '5
ElseIf vWeapon = DEagle Then
    GetTotalMags = DEagle_Mags '8
ElseIf vWeapon = FlameThrower Then
    GetTotalMags = Flamethrower_Mags '5
ElseIf vWeapon = M249 Then
    GetTotalMags = M249_Mags '3
ElseIf vWeapon = M82 Then
    GetTotalMags = M82_Mags '6
ElseIf vWeapon = RPG Then
    GetTotalMags = RPG_Mags '4
ElseIf vWeapon = SA80 Then
    GetTotalMags = SA80_Mags '5
ElseIf vWeapon = Shotgun Then
    GetTotalMags = Shotgun_Mags '12
ElseIf vWeapon = SCAR Then
    GetTotalMags = SCAR_Mags '5
End If

End Function

Private Function GetMagName(vWeapon As eWeaponTypes) As String

If vWeapon = RPG Then
    GetMagName = "Rockets"
ElseIf vWeapon = FlameThrower Then
    GetMagName = "Canisters"
ElseIf vWeapon = Shotgun Then
    GetMagName = "Shells"
Else
    GetMagName = "Magazines"
End If

End Function

Private Function PM_Rnd() As Single
PM_Rnd = (Rnd() - Rnd())
End Function

Private Sub ProcessAndDrawWallMarks()
Dim i As Integer

picMain.FillStyle = vbFSSolid
picMain.FillColor = modStickGame.cg_BGColour 'WallMark_Colour
Do While i < NumWallMarks
    
    modStickGame.sCircle WallMark(i).X, WallMark(i).Y, WallMark(i).Radius, modStickGame.cg_BGColour
    
    
    If WallMark(i).Decay < GetTickCount() Then
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
        
        HealthPack.X = 49200
        HealthPack.Y = 4800
        HealthPack.bActive = True
        
        SendBroadcast sHealthPacks & CStr(HealthPack.X) & "|" & CStr(HealthPack.Y)
        
        HealthPack.LastUsed = GetTickCount()
    End If
End If

End Sub

Private Sub DisplayHealthPack()

If HealthPack.bActive Then
    
    If modStickGame.sv_GameType = gCoOp Then
        HealthPack.bActive = False
    Else
        picMain.FillStyle = vbFSSolid
        sBox HealthPack.X, HealthPack.Y, HealthPack.X + HealthPack_Radius * 2, HealthPack.Y + HealthPack_Radius, vbRed
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

If Stick(0).iKillsInARow = RowKillsForArmour Then
    'add armour
    Stick(0).Armour = Max_Armour
    
ElseIf Stick(0).iKillsInARow = RowKillsForRadar Then
    RadarStartTime = GetTickCount()
    bHadRadar = True
    AddMainMessage "Radar Active for 30 Seconds"
    
ElseIf Stick(0).iKillsInARow = RowKillsForChopper Then
    If modStickGame.sv_GameType <> gCoOp Then
        ChopperAvail = modStickGame.sv_AllowChoppers 'True
    End If
    
End If

If FlamesInARow = RowFlameKillsForToasty Then
    
    If Stick(0).WeaponType = FlameThrower Then
        AddMainMessage "TOASTY! (Not 3D)"
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
        If Stick(0).bLightSaber = False Then
            Stick(0).bLightSaber = True
            AddMainMessage "Lightsaber Acquired. Hold Vertically to Block Bullets"
        End If
    End If
    
Else
    KnifesInARow = 0
End If

End Sub

Private Sub BltToForm()

BitBlt Me.hdc, 0, 0, ScaleX(Me.width, vbTwips, vbPixels), ScaleY(Me.height, vbTwips, vbPixels), _
    Me.picMain.hdc, 0, 0, modStickGame.cg_DisplayMode

'BitBlt Me.hdc, 0, 0, ScaleX(StickGameWidth, vbTwips, vbPixels), ScaleY(StickGameHeight, vbTwips, vbPixels), _
    Me.picMain.hdc, 0, 0, modStickGame.cg_DisplayMode

'vbNotSrcCopy
'vbSrcCopy

'RasterOpConstants
End Sub

Public Sub SwitchWeapon(ByVal vWeapon As eWeaponTypes)

If Stick(0).WeaponType <> vWeapon Then
    
    If vWeapon <> -1 Then
        If vWeapon <> Chopper Then
            
            If Stick(0).WeaponType <= Knife Then
                AmmoFired(Stick(0).WeaponType) = Stick(0).BulletsFired
            End If
            
            
            Stick(0).PrevWeapon = Stick(0).WeaponType
            
            On Error GoTo EH
            Stick(0).BulletsFired = AmmoFired(vWeapon)
            
        Else
            'ChopperAvail = False
            modStickGame.cg_LaserSight = False
            Stick(0).Health = Health_Start
        End If
        
        If Stick(0).State And Stick_Reload Then
            frmStickGame.SubStickiState 0, Stick_Reload
            ReloadKey = False
        End If
        
        
        'SWITCH IS HERE #################################################
        Stick(0).WeaponType = vWeapon
        'SWITCH IS HERE #################################################
        
        Scroll_WeaponKey = vWeapon
        WeaponKey = -1
        
        
        Stick(0).LastBullet = GetTickCount() - RPG_Bullet_Delay / modStickGame.sv_StickGameSpeed
        Stick(0).LastMuzzleFlash = 1 'turn it off
        
        If modStickGame.StickTeamFormLoaded Then
            frmStickGameSettings.chkShh.Value = IIf(WeaponSilencable(vWeapon), IIf(Stick(0).bSilenced, 1, 0), 0)
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

'Private Sub PrepareWeaponSelection()
'Dim i As Integer
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
'Dim i As Integer
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

On Error GoTo EH

If bChatActive Then Exit Sub

If KeyCode = vbKeyF1 Then
    ShowScoresKey = True
Else
    If StickInGame(0) And bPlaying Then
        Select Case KeyCode
            Case vbKeySpace, vbKeyW
                JumpKey = Stick(0).OnSurface
                
            Case vbKeyA
                LeftKey = Stick(0).OnSurface
                
            Case vbKeyD
                RightKey = Stick(0).OnSurface
                
            Case vbKeyControl, vbKeyS
                CrouchKey = Stick(0).OnSurface 'And Not StickiHasState(0, Stick_Prone)
                
            Case vbKeyE
                If modStickGame.sv_2Weapons Then
                    For i = 0 To NumStaticWeapons - 1
                        If StickNearStaticWeapon(0, i) Then
                            UseKey = True
                            Exit For
                        End If
                    Next i
                End If
                'UseKey = True
                
            Case vbKeyR
                
                If Stick(0).WeaponType < Knife Then
                    If TotalMags(Stick(0).WeaponType) > 0 Then
                        If StickiHasState(0, Stick_Reload) = False Then
                            ReloadKey = True
                        End If
                    End If
                End If
                
            Case vbKeyK
                
                SwitchWeapon Knife
                
                
            Case Is <= (49 + eWeaponTypes.Chopper) 'eWeaponTypes.Knife)
                
                If modStickGame.sv_2Weapons = False Then
                    'If (Stick(0).State And Stick_Reload) = 0 Then
                        
                        If KeyCode >= 48 Then
                            If Stick(0).WeaponType = Chopper Then
                                WeaponKey = -1
                            Else
                                WeaponKey = KeyCode - 49
                                
                                If WeaponKey = -1 Then
                                    WeaponKey = Chopper
                                    
                                    If ChopperAvail = False Then WeaponKey = -1
                                End If
                            End If
                        End If
                    'End If
                    
                ElseIf KeyCode >= 49 Then
                    If KeyCode <= 50 Then
                        'only allow vbkey1 and vbkey2
                        'to swap weapons
                        If LastWeaponSwitch + WeaponSwitchDelay < GetTickCount() Then
                            If Stick(0).WeaponType = Chopper Then
                                
                                If ChopperAvail Then
                                    WeaponKey = Chopper
                                Else
                                    WeaponKey = -1
                                End If
                                
                            Else
                                If Stick(0).WeaponType = Stick(0).CurrentWeapons(1) Then
                                    WeaponKey = Stick(0).CurrentWeapons(2)
                                Else
                                    WeaponKey = Stick(0).CurrentWeapons(1)
                                End If
                            End If
                            
                            LastWeaponSwitch = GetTickCount()
                        End If
                        
                    End If
                    
                ElseIf KeyCode = 48 Then 'chopper
                    If ChopperAvail Then
                        WeaponKey = Chopper
                    End If
                End If
                
            Case vbKeyAdd
                
                If cg_sZoom < MaxZoom Then
                    cg_sZoom = Round(cg_sZoom + ZoomInc, 2)
                End If
                LastZoomPress = GetTickCount()
            
            Case vbKeySubtract
                
                If cg_sZoom >= MinZoom Then
                    cg_sZoom = Round(cg_sZoom - ZoomInc, 2)
                End If
                LastZoomPress = GetTickCount()
                
            Case vbKeyMultiply
                
                cg_sZoom = 1
                LastZoomPress = GetTickCount()
                
        End Select
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
        End Select
    End If
End If

EH:
End Sub

Private Function GetNadeTypeName() As String
GetNadeTypeName = kNadeName(Stick(0).iNadeType)
End Function

Private Sub MakeNadeNameArray()
Dim i As Integer

For i = 0 To 2
    If i = nFrag Then
        kNadeName(i) = "Frag"
    ElseIf i = nFlash Then
        kNadeName(i) = "Flash Bang"
    Else
        kNadeName(i) = "Smoke"
    End If
Next i

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim bCan As Boolean
Static LastNadeSwitch As Long

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
        
        'vbKeyV
        Case (KeyAscii = 118 Or (KeyAscii = vbKey3 And modStickGame.sv_2Weapons)) And bCan
            'switch nade type
            
            If LastNadeSwitch + 300 < GetTickCount() Then
                If Stick(0).iNadeType = nFrag Then
                    Stick(0).iNadeType = nFlash
                ElseIf Stick(0).iNadeType = nFlash Then
                    Stick(0).iNadeType = nSmoke
                Else
                    Stick(0).iNadeType = nFrag
                End If
                
                AddMainMessage "Grenade Type: " & GetNadeTypeName()
                
                LastNadeSwitch = GetTickCount()
            End If
            
        
        '109 = vbkeyM
        Case KeyAscii = 109 And bCan 'vbKeyZ
            MineKey = True
            
        Case (KeyAscii = 99 Or KeyAscii = vbKeyC Or KeyAscii = 102 Or KeyAscii = vbKeyF) And bCan    'vbKeyC
            If Stick(0).WeaponType <> Chopper Then
                ProneKey = Not ProneKey 'And Stick(0).OnSurface
            End If
            
        Case (KeyAscii = vbKeyQ Or KeyAscii = 113 Or KeyAscii = 17) And bCan
            Stick(0).bSilenced = Not Stick(0).bSilenced
            
            If modStickGame.StickTeamFormLoaded Then
                frmStickGameSettings.chkShh.Value = IIf(Stick(0).bSilenced, 1, 0)
            End If
            
            
        '#########################################
        'Chat handling
        'Escape kills the chat
        Case KeyAscii = vbKeyEscape
            bChatActive = False
            strChat = vbNullString
            
            
            'T TO TALK
        Case ((KeyAscii = 116) Or (KeyAscii = 84)) And (bChatActive = False)
            '116=t
            bChatActive = True
            
            'disable movement
            LeftKey = False: RightKey = False: JumpKey = False
            
            
        Case KeyAscii = vbKeyBack
            If LenB(strChat) Then
                strChat = Left$(strChat, Len(strChat) - 1)
            End If
            
            'Return finishes and sends the chat
        Case KeyAscii = vbKeyReturn
            
            If bChatActive Then
                'Send it!
                
                strChat = Trim$(strChat)
                
                If LenB(strChat) Then
                    SendChatPacket Trim$(Stick(0).Name) & modMessaging.MsgNameSeparator & strChat, Stick(0).Colour
                End If
                
                'Reset
                bChatActive = False
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

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF1 Then
    ShowScoresKey = False
Else
    If StickInGame(0) And bPlaying Then
        Select Case KeyCode
            Case vbKeySpace, vbKeyW
                JumpKey = False
                
            Case vbKeyA
                'If Stick(0).OnSurface Then
                    LeftKey = False
                'End If
                
            Case vbKeyD
                'If Stick(0).OnSurface Then
                    RightKey = False
                'End If
                
            Case vbKeyControl, vbKeyS
                CrouchKey = False
                
            Case vbKeyR
                ReloadKey = False
                
            'Case vbKeyE
                'UseKey = False
                
        End Select
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
    
End If

End Sub

Private Sub SetStickiState(i As Integer, State As eStickStates)

Stick(i).State = State

End Sub

'Private Sub AddStickState(ID As Integer, State As eStickStates)
'Dim i As Integer
'
'i = FindStick(ID)
'
''Find the specified Stick and add to his state
'Stick(i).State = (Stick(i).State Or State)
'
'End Sub
'
'Public Sub SubStickState(ID As Integer, State As eStickStates)
'Dim i As Integer
'
'i = FindStick(ID)
'
''Find the specified Stick and subtract from his state
'Stick(i).State = (Stick(i).State And (Not State))
'
'End Sub

Private Sub AddStickiState(i As Integer, State As eStickStates)

Stick(i).State = (Stick(i).State Or State)

End Sub

Public Sub SubStickiState(i As Integer, State As eStickStates)

Stick(i).State = (Stick(i).State And (Not State))

End Sub

'Private Function StickHasState(ID As Integer, vState As eStickStates) As Boolean
'
'StickHasState = CBool((Stick(FindStick(ID)).State And vState))
'
'End Function

Private Function StickiHasState(Index As Integer, vState As eStickStates) As Boolean

StickiHasState = CBool((Stick(Index).State And vState))

End Function

Public Function FindStick(ID As Integer) As Integer

Dim i As Integer

'Find and return the current array index for this Stick
FindStick = -1
For i = 0 To NumSticksM1
    'Is this the Stick?
    If Stick(i).ID = ID Then
        'This is the one!
        FindStick = i
        Exit Function
    End If
Next i

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

modStickGame.sv_StickGameSpeed = 1
WindowClosing = False
MouseX = 15915
MouseY = 3435
picMain.BackColor = Me.BackColor
picMain.Visible = False


Call FormLoad(Me, , , False, True)

'Display the form
Show 'vbModeless, frmMain
Me.ZOrder vbBringToFront

If modStickGame.cl_Subclass Then
    If Not IsIDE() Then
        modSubClass.SubClassStick Me.hWnd
    Else
        modStickGame.cl_Subclass = False
    End If
End If


'Call PrepareWeaponSelection
Call InitVariables
tmrMain.Enabled = True

End Sub

Private Sub Form_Resize()
picMain.width = Me.width
picMain.height = Me.height

StickCentreX = Me.width \ 2 - 500
StickCentreY = Me.height \ 2 - 500

'sort out constants
RadarLeft = Me.width - RadarWidth - 100
PlayingX = StickCentreX - 600
ConnectingkX = StickCentreX - 900
ConnectingkY = StickCentreY + 650

End Sub

Private Sub MainLoop()
Dim Timer As Long
Dim LastFullSecond As Long
Dim nFrames As Integer
Dim newTick As Long

bRunning = True
bPlaying = True
Timer = GetTickCount() 'prevent elapsed time from being huge

Do While bRunning
    
    newTick = GetTickCount()
    If Timer + Stick_Ms_Required_Delay < newTick Then
        
        modStickGame.StickElapsedTime = newTick - Timer
        StickTimeFactor = modStickGame.sv_StickGameSpeed * StickElapsedTime / Stick_Ms_Delay
        
        
        nFrames = nFrames + 1
        If LastFullSecond + 1000 < newTick Then
            FPS = nFrames
            nFrames = 0
            LastFullSecond = newTick
        End If
        
        Timer = newTick 'GetTickCount()
        
        
        On Error GoTo EH
        
        If GetPacket() = False Then Exit Do
        
        
        SendUpdatePacket
        SendSlowPacket
        
        
        If Not Stick(0).bFlashed Then
            picMain.Cls
        ElseIf ((GetTickCount() - Stick(0).LastFlashBang) * modStickGame.sv_StickGameSpeed / 5000 * PM_Rnd()) > 0.75 Then
            picMain.Cls
        End If
        
        
        If bPlaying Then
            Physics
            ProcessBlood
            DrawBullets
            DrawBlood
            DisplaySticks
            DrawDeadSticks
            DrawLaserSight
            ProcessMagazines '+draw
            ProcessMuzzleFlashes '+draw
            DrawNames
            DrawDeadChoppers
            DrawStaticWeapons 'only the images
            
            DrawPlatforms
            DrawCasings
            DrawBoxes
            DrawMines
            DrawtBoxes
            
            ProcessAndDrawWallMarks
            ProcessSmokeBlasts
            
            DrawAmmoPickups
            ProcessAllCircs '+draw
            DrawNades
            ProcessSmoke '+draw
            ProcessSparks
            
            ProcessAndDrawLargeSmokes
            ProcessStaticWeapons '+ draw "Pick up AK-47" bit
            
            DisplayHealthPack
            ProcessFlames
            
            
            DrawCrosshair
            '---------
            DisplayHUD
            DisplayChat
            
            ProcessKeys
            ProcessAllAI
            
            ProcessNades
            ProcessMines
            ProcessCasings
            
            
            
            DrawRadar
            ShowMainMessages
            ShowChatEntry
            
            ProcessPerk
            ProcessToasty
            ProcessAmmoPickups
        Else
            StaticPhysics
            DrawBullets
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
            
            ProcessAndDrawWallMarks
            ProcessSmokeBlasts
            
            DrawCasings
            ProcessAllCircs '+draw
            ProcessNades
            ProcessCasings
            DrawNades
            ProcessSmoke '+draw
            ProcessSparks
            
            SetMyStickFacing
            'DrawCrosshair
            
            'show scoreboard, etc
            ProcessEndRound
            
            
            DisplayChat
            ShowChatEntry
            
            
            ProcessKeys
        End If
        
        
        'draw on form
        BltToForm
        
        
        If modStickGame.StickServer Then
            CheckStickNames
            SendServerVarPacket
            SendRoundInfo
            
            If bPlaying Then
                SendBoxInfo
                GenerateHealthPack
                SendStaticWeaponsPacket
                CheckMaxScore
                
                If modStickGame.sv_GameType = gElimination Or modStickGame.sv_GameType = gCoOp Then
                    ProcessElimination
                'ElseIf modStickGame.sv_GameType = gCoOp Then
                    'ProcessElimination
                    'processcoop
                'ElseIf modStickGame.sv_GameType = gAssault Then
                End If
                
            End If
        End If
        
        
    End If
    
EH:
    DoEvents
Loop

'On Error GoTo 0
'SavePicture picMain.Image, AppPath() & "\test.bmp"
'
'Stop

End Sub

Private Sub DisplayScoreBoard()
Const ScoreBoardWidth = 2000
Dim sTxt As String
Dim i As Integer
Dim X As Single, Y As Single

X = Me.width - ScoreBoardWidth

If RadarStartTime + Radar_Time > GetTickCount() Then
    Y = 1200
Else
    Y = 10
End If

'On Error Resume Next
BorderedBox X, Y - 200, X + 2000, Y + 195 * CSng(NumSticks), BoxCol

X = X + 100
For i = 0 To NumSticksM1
    sTxt = Trim$(Stick(i).Name) & ": " & CStr(Stick(i).iKills)
    
    PrintStickFormText sTxt, X, i * TextHeight(sTxt) + Y, Stick(i).Colour
Next i

End Sub

Private Sub BangFlash(iNade As Integer)
Dim i As Integer
Const FlashDist = 10000
Const CircDist = 10000

Stick(0).LastFlashBang = GetTickCount()
Stick(0).bFlashed = True

StunnedMouseX = MouseX + 500 * PM_Rnd()
StunnedMouseY = MouseY + 500 * PM_Rnd()

modStickGame.sBoxFilled -FlashDist, -FlashDist, StickGameWidth + FlashDist, StickGameHeight + FlashDist, vbWhite

For i = 0 To 100
    
    AddCirc Stick(0).X + PM_Rnd * CircDist, _
            Stick(0).Y + PM_Rnd * CircDist, _
            2000, 0.2, RandomRGBColour()
    
    
Next i

AddCirc Nade(iNade).X, Nade(iNade).Y, 5000, 0.2, vbYellow

End Sub

Private Sub StaticPhysics()
Dim i As Integer

For i = 0 To NumSticksM1
    If Stick(i).Speed > 0 Then
        Stick(i).Speed = 0
        Stick(i).State = Stick_None
    'Else
        'Stick(i).Speed = Stick(i).Speed * modStickGame.StickTimeFactor / 2
    End If
    
    'Motion Stick(i).X, Stick(i).Y, Stick(i).Speed, Stick(i).Heading
    
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
picToasty.Visible = True
End Sub

Private Sub BorderedBox(x1 As Single, y1 As Single, X2 As Single, Y2 As Single, lColour As Long)
picMain.DrawWidth = 1
picMain.Line (x1, y1)-(X2, Y2), lColour, BF
picMain.Line (x1, y1)-(X2, Y2), vbBlack, B
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
                RoundWinnerID = Stick(0).ID
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

Private Sub ProcessPerk()
Const ESP_Print_Len = 4500
Const ESP_Print_LenDX = ESP_Print_Len * 1.1
Const ESP_Y_Offset = 500

Dim i As Integer
Dim tDist As Single, tAng As Single


If Stick(0).Perk = pStealth Then
    'draw esp map
    
    picMain.DrawWidth = 1
    
    PrintStickText "Stealth Awareness", Stick(0).X - 600, Stick(0).Y + ESP_Y_Offset - ESP_Print_Len - 250, BoxCol
    
    For i = 1 To NumSticksM1
        
        If CanSeeStick(i) Then
            tDist = GetDist(Stick(0).X, Stick(0).Y, Stick(i).X, Stick(i).Y)
            
            If tDist < StealthESPDist Then
                tAng = FindAngle(Stick(0).X, Stick(0).Y, Stick(i).X, Stick(i).Y - 1)
                
                'picMain.Font.Size = 9 - tDist / 10000
                
                PrintStickText Trim$(Stick(i).Name), _
                    Stick(0).X + ESP_Print_Len * Sin(tAng), _
                    Stick(0).Y + ESP_Y_Offset - ESP_Print_Len * Cos(tAng), _
                    Stick(i).Colour
            End If
        End If
    Next i
    
    modStickGame.sCircle Stick(0).X, Stick(0).Y + ESP_Y_Offset, ESP_Print_LenDX, BoxCol
    picMain.Font.Size = 8
End If

End Sub

Public Sub StickGameSpeedChanged(oldSpeed As Single, newSpeed As Single)
Dim i As Integer

Erase Flame: NumFlames = 0
Erase Smoke: NumSmoke = 0

If oldSpeed > -1 Then
    For i = 0 To NumNades - 1
        
        Nade(i).Decay = (Nade(i).Decay - GetTickCount()) * oldSpeed / newSpeed + GetTickCount()
        
        '(Nade(i).Decay - GetTickCount()) * oldSpeed / newSpeed + GetTickCount()
        
        '(n-g)o/n + g
        'o-g(o/n + 1)   ?
        
    Next i
End If

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
                            'If Stick(j).LastSpawnTime + Spawn_Invul_Time / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                            If StickInvul(j) = False Then
                                If CoOrdInStick(Stick(j).X, Stick(j).Y, i) Then
                                    
                                    For K = 1 To 30
                                        'splatter!
                                        AddBlood Stick(j).X, Stick(j).Y, Rnd() * 2 * pi, False
                                    Next K
                                    
                                    If Stick(j).ID = MyID Or Stick(j).IsBot Then
                                        Call Killed(j, i, kChoppered)
                                    End If
                                    
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


If RadarStartTime + Radar_Time > GetTickCount() Then
    
    picMain.DrawWidth = 2
    picMain.Line (RadarLeft, RadarTop)-(RadarLeft + RadarWidth, RadarTop + RadarHeight), vbBlue, B
    
    modStickGame.PrintStickFormText "Time Left: " & _
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
    
    
    
    picMain.FillStyle = vbFSTransparent
    For i = 0 To NumSticksM1
        
        If StickInGame(i) Then
            If Stick(i).Perk <> pRadarJammer Then
                
                pX = RadarLeft + RadarWidth * Stick(i).X / StickGameWidth
                pY = RadarTop + RadarHeight * Stick(i).Y / StickGameHeight
                
                C = Stick(i).Colour 'GetTeamColour(Stick(i).Team)
                picMain.FillColor = C
                picMain.Circle (pX, pY), 60, C
                
                If i = 0 Then
                    'draw an X on me
                    DrawX pX, pY
                End If
            End If
        End If
    Next i
    
    
    
ElseIf bHadRadar Then
    AddMainMessage "Radar Expired"
    bHadRadar = False
End If

End Sub

Private Sub DrawX(ByVal pX As Single, ByVal pY As Single)

Const CrossWidth = 75

picMain.Line (pX - CrossWidth, pY + CrossWidth)-(pX + CrossWidth, pY - CrossWidth), Stick(0).Colour
picMain.Line (pX + CrossWidth, pY + CrossWidth)-(pX - CrossWidth, pY - CrossWidth), Stick(0).Colour

End Sub

Private Sub ShowMainMessages()
'Const WO2 = 3935 'Width \ 2 - 100
'Const HO2 = 3235 'Height \ 2 - 100
Dim i As Integer
Dim Tmp As String
Const TextH = 480

If NumMainMessages Then
    'If Not F1Pressed Then  'And ShowMainMsg Then
    
    picMain.Font.Size = 18
    picMain.ForeColor = Stick(0).Colour 'MGrey
    
    Do While i < NumMainMessages
        If MainMessages(i).Decay < GetTickCount() Then
            RemoveMainMessage i
            i = i - 1
        End If
        
        i = i + 1
    Loop
    
    
    For i = 0 To NumMainMessages - 1
        Tmp = MainMessages(i).Text
        
        PrintStickFormText Tmp, StickCentreX - TextWidth(Tmp) / 2 - 500, i * TextH + StickCentreY + 1000, MainMessages(i).Colour
    Next i
    
    picMain.Font.Size = 8
End If

End Sub

Public Sub AddMainMessage(ChatText As String, Optional lColour As Long = -1)

'Add this value to the chat text array
ReDim Preserve MainMessages(NumMainMessages)
MainMessages(NumMainMessages).Decay = GetTickCount() + MainMessageDecay
MainMessages(NumMainMessages).Text = ChatText
MainMessages(NumMainMessages).Colour = IIf(lColour = -1, Stick(0).Colour, lColour)
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
If StickInvul(i) = False Then
    Stick(i).bOnFire = (Stick(i).LastFlameTouch + Flame_Burn_Time / modStickGame.sv_StickGameSpeed > GetTickCount())
End If
End Sub
Private Sub SetStickFlashed(i As Integer)
If StickInvul(i) = False Then
    Stick(i).bFlashed = (Stick(i).LastFlashBang + FlashBang_Time / modStickGame.sv_StickGameSpeed > GetTickCount())
End If
End Sub

Private Sub ProcessKeys()

If JumpKey Then
    If Stick(0).LastMine + 500 < GetTickCount() Then
        'don't let them jump immediatly - let clients place the mine too
        
        'If Stick(0).StartJumpTime + JumpTime < GetTickCount() Then
            AddStickiState 0, stick_Jump
            'Stick(0).StartJumpTime = GetTickCount()
            
            'JumpKey = False
        'End If
    End If
End If

If LeftKey Then
    If StickiHasState(0, stick_Left) = False Then
        AddStickiState 0, stick_Left
    End If
ElseIf StickiHasState(0, stick_Left) Then
    If Stick(0).OnSurface Then
        SubStickiState 0, stick_Left
    End If
End If
    
If RightKey Then
    If StickiHasState(0, stick_Right) = False Then
        AddStickiState 0, stick_Right
    End If
ElseIf StickiHasState(0, stick_Right) Then
    If Stick(0).OnSurface Then
        SubStickiState 0, stick_Right
    End If
End If


'##########################################
If bChatActive Then
    If Stick(0).bTyping = False Then
        Stick(0).bTyping = True
    End If
    Exit Sub
ElseIf Stick(0).bTyping Then
    Stick(0).bTyping = False
End If
SetStickFlashed 0
SetStickOnFire 0
'##########################################


If StickInGame(0) And bPlaying Then
    
    If CrouchKey And Not JumpKey Then
        If StickiHasState(0, stick_crouch) = False Then
            AddStickiState 0, stick_crouch
        ElseIf StickiHasState(0, Stick_Prone) Then
            SubStickiState 0, Stick_Prone
            ProneKey = False
        End If
    ElseIf StickiHasState(0, stick_crouch) Then
        SubStickiState 0, stick_crouch
    ElseIf ProneKey And Not JumpKey Then
        If Stick(0).WeaponType = Chopper Then
            ProneKey = False
        End If
        If StickiHasState(0, Stick_Prone) = False Then
            AddStickiState 0, Stick_Prone
        End If
    ElseIf StickiHasState(0, Stick_Prone) Then
        SubStickiState 0, Stick_Prone
    End If
    
    
    'THIS IS THE BIT THAT'LL RESET SHOTGUN/SNIPER FACING AFTER A BULLET
    If Stick(0).LastBullet + (GetBulletDelay(0) - 50) / (1.5 * modStickGame.sv_StickGameSpeed) < GetTickCount() Then
        SetMyStickFacing
        
        'With Stick(0)
            
            'If GetDist(Stick(0).GunPoint.X, Stick(0).GunPoint.Y, MouseX, MouseY) > BodyLen Then
                
                '.Facing = FindAngle(Stick(0).GunPoint.X, Stick(0).GunPoint.Y, MouseX, MouseY)
            'End If
            
            
        'End With
    End If
    
    If FireKey Then
        If StickiHasState(0, Stick_Reload) = False Then
            AddStickiState 0, Stick_Fire
        End If
    End If
    
    If ReloadKey Then
        ReloadKey = False
        
        If Stick(0).BulletsFired > 0 Then
            StartReload 0
        End If
    End If
    
    
    If WeaponKey <> -1 Then
        If StickiHasState(0, Stick_Fire) Then
            WeaponKey = -1
        ElseIf Stick(0).WeaponType <> WeaponKey Then
            If modStickGame.sv_AllowRockets = False Then
                If WeaponKey = RPG Then
                    WeaponKey = -1
                    Exit Sub
                End If
            ElseIf modStickGame.sv_AllowFlameThrowers = False Then
                If WeaponKey = FlameThrower Then
                    WeaponKey = -1
                    Exit Sub
                End If
            End If
            
            
            SwitchWeapon WeaponKey
            
            WeaponKey = -1
        End If
    End If
    
    ValidateWeapons
    
    
    If MineKey Then
        AddStickiState 0, Stick_Mine
    'else
        'sub'd in physics
    End If
    
    
    
    If LastScrollWeaponSwitch + Scroll_Delay < GetTickCount() Then
        If LastScrollWeaponSwitch Then
            If Stick(0).WeaponType <> Scroll_WeaponKey Then
                If Scroll_WeaponKey <> -1 Then
                    
                    If StickiHasState(0, Stick_Fire) Then
                        Scroll_WeaponKey = Stick(0).WeaponType
                    ElseIf StickiHasState(0, Stick_Reload) = False Then
                        SwitchWeapon Scroll_WeaponKey
                    Else
                        Scroll_WeaponKey = Stick(0).WeaponType
                    End If
                    
                Else
                    Scroll_WeaponKey = Stick(0).WeaponType
                End If
            End If
        Else
            Scroll_WeaponKey = Stick(0).WeaponType
        End If
        
        'LastScrollWeaponSwitch = GetTickCount()
    End If
    
    If Stick(0).bSilenced Then
        If WeaponSilencable(Stick(0).WeaponType) = False Then
            Stick(0).bSilenced = False
        End If
    End If
    
    
    If UseKey Then
        If StickiHasState(0, Stick_Use) = False Then
            AddStickiState 0, Stick_Use
        End If
    ElseIf StickiHasState(0, Stick_Use) Then
        If Stick(0).LastWeaponSwitch + UseKeyReleaseDelay < GetTickCount() Then
            SubStickiState 0, Stick_Use
        End If
    End If
    
Else
    
    Const CamInc = 100
    
    If SpecUp Then
        MoveCameraY modStickGame.cg_sCamera.Y - CamInc * modStickGame.cl_SpecSpeed
    ElseIf SpecDown Then
        MoveCameraY modStickGame.cg_sCamera.Y + CamInc * modStickGame.cl_SpecSpeed
    End If
    
    
    If SpecLeft Then
        MoveCameraX modStickGame.cg_sCamera.X - CamInc * modStickGame.cl_SpecSpeed
    ElseIf SpecRight Then
        MoveCameraX modStickGame.cg_sCamera.X + CamInc * modStickGame.cl_SpecSpeed
    End If
    
End If

End Sub

Private Sub ValidateWeapons()

If modStickGame.sv_AllowRockets = False Then
    If Stick(0).CurrentWeapons(1) = RPG Then
        Stick(0).CurrentWeapons(1) = AK
        
        If Stick(0).CurrentWeapons(2) = AK Then
            Stick(0).CurrentWeapons(2) = DEagle
        End If
    ElseIf Stick(0).CurrentWeapons(2) = RPG Then
        Stick(0).CurrentWeapons(2) = AK
        
        If Stick(0).CurrentWeapons(1) = AK Then
            Stick(0).CurrentWeapons(1) = DEagle
        End If
    End If
    
    If Stick(0).WeaponType = RPG Then
        Stick(0).WeaponType = AK
    End If
End If

If modStickGame.sv_AllowRockets = False Then
    If Stick(0).CurrentWeapons(1) = FlameThrower Then
        Stick(0).CurrentWeapons(1) = AK
        
        If Stick(0).CurrentWeapons(2) = AK Then
            Stick(0).CurrentWeapons(2) = DEagle
        End If
    ElseIf Stick(0).CurrentWeapons(2) = FlameThrower Then
        Stick(0).CurrentWeapons(2) = AK
        
        If Stick(0).CurrentWeapons(1) = AK Then
            Stick(0).CurrentWeapons(1) = DEagle
        End If
    End If
    
    If Stick(0).WeaponType = FlameThrower Then
        Stick(0).WeaponType = AK
    End If
End If

End Sub

Public Function WeaponSilencable(vWeapon As eWeaponTypes) As Boolean
If vWeapon = AK Then
    WeaponSilencable = True
ElseIf vWeapon = DEagle Then
    WeaponSilencable = True
ElseIf vWeapon = M82 Then
    WeaponSilencable = True
ElseIf vWeapon = SCAR Then
    WeaponSilencable = True
ElseIf vWeapon = SA80 Then
    WeaponSilencable = True
End If
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
Dim iTarget As Integer
Dim Dist As Single

Const AI_Facing_Adjust_Delay = AI_Delay / 4
Const BotHardcoreScanInc = pi / 6

If i <> -1 Then
    
    SetStickFlashed i
    SetStickOnFire i
    
    
    
    If Stick(i).AICurrentTarget > -1 And Stick(i).AICurrentTarget < NumSticksM1 Then
        If Stick(i).AILastFacingAdjust + AI_Facing_Adjust_Delay < GetTickCount() Then
            
            If Stick(i).bFlashed Then
                Stick(i).ActualFacing = Stick(i).ActualFacing + PM_Rnd()
                'PrintStickText "Flashed", Stick(i).X, Stick(i).Y, vbRed
                
                If Rnd() > 0.9 Then
                    If StickiHasState(i, stick_Left) Then
                        SubStickiState i, stick_Left
                        AddStickiState i, stick_Right
                    Else
                        SubStickiState i, stick_Right
                        AddStickiState i, stick_Left
                    End If
                    
                    If StickiHasState(i, Stick_Fire) = False Then
                        AddStickiState i, Stick_Fire
                    End If
                    
                End If
                
            Else
                DoAIFacing i, Stick(i).AICurrentTarget
            End If
            
            Stick(i).AILastFacingAdjust = GetTickCount()
        End If
        
    ElseIf Stick(i).AICurrentTarget > NumSticksM1 Then
        Stick(i).AICurrentTarget = -1
    End If
    
    
    
    If Not Stick(i).bFlashed Then
        If Stick(i).LastAI + AI_Delay < GetTickCount() Then
            
            iTarget = ClosestTargetI(i, Dist)
            
            
            
            If iTarget > -1 Then
                
                Stick(i).AICurrentTarget = iTarget
                pProcessAI i, iTarget, Dist
                
            ElseIf modStickGame.sv_Hardcore Then
                
                'attempt to locate someone
                Stick(i).ActualFacing = Stick(i).ActualFacing + BotHardcoreScanInc
                Stick(i).Facing = Stick(i).ActualFacing
                
            End If
            
            If iTarget = -1 Then
                If Stick(i).State > Stick_None Then
                    SetStickiState i, Stick_None
                    
                    Stick(i).AICurrentTarget = -1
                End If
            End If
            
            
            Stick(i).LastAI = GetTickCount()
            
            
            If modStickGame.sv_AllowFlameThrowers = False Then
                If Stick(i).WeaponType = FlameThrower Then
                    Stick(i).WeaponType = GetRandomStaticWeapon()
                End If
            End If
            If modStickGame.sv_AllowRockets = False Then
                If Stick(i).WeaponType = RPG Then
                    Stick(i).WeaponType = GetRandomStaticWeapon()
                End If
            End If
            
            
        End If
    End If
    
End If

End Sub

Private Sub DoAIFacing(iAi As Integer, iTarget As Integer)
Dim FixedAngle As Single, AngleWantToFace As Single

If Stick(iTarget).WeaponType = Chopper Then
    AngleWantToFace = FindAngle(Stick(iAi).X, Stick(iAi).Y, Stick(iTarget).X, GetStickY(iTarget) - BodyLen - HeadRadius)
Else
    AngleWantToFace = FindAngle(Stick(iAi).X, Stick(iAi).Y, Stick(iTarget).X, GetStickY(iTarget) + HeadRadius)
End If

'Stick(iAi).AIWantToFace = AngleWantToFace

'Adjust facing
FixedAngle = FixAngle(AngleWantToFace - Stick(iAi).ActualFacing)
If FixedAngle <= modStickGame.sv_Bot_Rotation_Rate Or FixedAngle >= modStickGame.sv_Bot_pi2LessRotRate Then
    
    Stick(iAi).ActualFacing = AngleWantToFace
    
    
ElseIf FixedAngle >= pi Then
    'mudtFleet2(i).sngFacing = mudtFleet2(i).sngFacing - FLEET2_ROTATION_RATE * pi / 180
    Stick(iAi).ActualFacing = Stick(iAi).ActualFacing - modStickGame.sv_Bot_Rotation_Rate
    
ElseIf FixedAngle < pi Then
    'mudtFleet2(i).sngFacing = mudtFleet2(i).sngFacing + FLEET2_ROTATION_RATE * pi / 180
    Stick(iAi).ActualFacing = Stick(iAi).ActualFacing + modStickGame.sv_Bot_Rotation_Rate
    
End If


End Sub

Private Sub pProcessAI(iAi As Integer, iTarget As Integer, DistToTarget As Single)
Const MinRange = StickGameWidth / 6
Const LevelGap = 1500
Const ChopperMinYDist = 2000, ChopperMinYDist2 = 800, ChopperMinXDist = 7000

'decision vars
Dim yDist As Single, xDist As Single, bDontMove As Boolean
'act-on vars
Dim bJump As Boolean, IDir As Integer, bCanShoot As Boolean, bCloseToRange As Boolean


yDist = Stick(iAi).Y - Stick(iTarget).Y
xDist = Stick(iAi).X - Stick(iTarget).X



If Stick(iAi).WeaponType = Chopper Then
    
    If modStickGame.sv_AIMove Then
        If modStickGame.sv_GameType <> gCoOp Then
            
            yDist = Stick(iAi).Y - Stick(iTarget).Y
            
            'up + down
            If yDist > ChopperMinYDist2 Then
                'they're above me
                
                If StickiHasState(iAi, stick_Jump) = False Then
                    AddStickiState iAi, stick_Jump
                ElseIf StickiHasState(iAi, stick_crouch) Then
                    SubStickiState iAi, stick_crouch
                End If
                
                
            ElseIf yDist < -ChopperMinYDist Then
                If StickiHasState(iAi, stick_crouch) = False Then
                    AddStickiState iAi, stick_crouch
                ElseIf StickiHasState(iAi, stick_Jump) Then
                    SubStickiState iAi, stick_Jump
                End If
                
                
            Else
                
                'on their level
                If StickiHasState(iAi, stick_Jump) Then
                    SubStickiState iAi, stick_Jump
                ElseIf StickiHasState(iAi, stick_crouch) Then
                    SubStickiState iAi, stick_crouch
                End If
                
                If Stick(iAi).Speed > 0 Then
                    Stick(iAi).Speed = Stick(iAi).Speed / 3
                End If
                
            End If
            
            
            If xDist > ChopperMinXDist Then
                IDir = -1
            ElseIf xDist < -ChopperMinXDist Then
                IDir = 1
            End If
            
            bDontMove = True
            bCanShoot = True
            bJump = False
        Else
            If Stick(iAi).Facing < piD2 Then
                If Stick(iAi).Facing > pi3D2 Then
                    'can't face up
                    Stick(iAi).Facing = IIf(Stick(iAi).Facing < pi, piD2, pi3D2)
                End If
            End If
        End If
    End If
End If




If bDontMove = False Then
    If DistToTarget < MinRange Then
        bCanShoot = True
    Else
        'close to minrange
        bCloseToRange = True
    End If
    
    If yDist > LevelGap Then
        'target above
        
        If Abs(xDist) < IIf(Stick(iTarget).WeaponType = Chopper, 20000, 5000) Then
            bJump = True
            bCanShoot = False
        Else
            bCloseToRange = True
            bJump = False
        End If
        
    ElseIf yDist < -LevelGap Then
        'target 1+ level(s) down
        
        'always go right
        IDir = 1
        
    '    If Stick(iAI).X > Stick(iTarget).X Then
    '        iDir = 1
    '    Else
    '        iDir = -1
    '    End If
    End If
End If


If bCloseToRange Then
    If xDist > 0 Then
        IDir = -1
    Else
        IDir = 1
    End If
    
ElseIf bCanShoot Then
    'shoot+nade
    If Stick(iAi).WeaponType = FlameThrower Or Stick(iAi).WeaponType = RPG Then
        Stick(iAi).Facing = Stick(iAi).ActualFacing
    End If
    
    
    If modStickGame.sv_AIShoot Then
        
        'If AnglesRoughlyEqual(Stick(iAi).Facing, Stick(iAi).AIWantToFace) Then
            If StickiHasState(iAi, Stick_Fire) = False Then
                AddStickiState iAi, Stick_Fire
            End If
        'ElseIf StickiHasState(iAi, Stick_Fire) Then
            'SubStickiState iAi, Stick_Fire
        'End If
        
'        If Stick(iAi).WeaponType = M82 Or Stick(iAi).WeaponType = Shotgun Then
'            'delay between shots
'
'            If Stick(iAi).LastBullet + 300 > GetTickCount() Then
'                If StickiHasState(iAi, Stick_Fire) Then
'                    SubStickiState iAi, Stick_Fire
'                End If
'            End If
'
'        End If
        
        If modStickGame.sv_BotHeliRocket Or (Stick(iAi).WeaponType <> Chopper) Then
            If Stick(iAi).LastNade + Stick(iAi).AINadeDelay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                AddStickiState iAi, Stick_Nade
                
                SetAINadeDelay iAi
                Stick(iAi).AIPickedNade = False
                
            ElseIf Stick(iAi).AIPickedNade = False Then
                
                Stick(iAi).AIPickedNade = True
                Stick(iAi).iNadeType = IIf(Rnd() > 0.75, eNadeTypes.nFlash, eNadeTypes.nFrag)
                
            End If
        End If
    ElseIf StickiHasState(iAi, Stick_Fire) Then
        SubStickiState iAi, Stick_Fire
    End If
    
End If


If IDir Then
    Stick(iAi).Facing = Stick(iAi).ActualFacing
    
    'move to range
    If modStickGame.sv_AIMove Then
        If IDir = -1 Then
            'move left
            If StickiHasState(iAi, stick_Left) = False Then
                AddStickiState iAi, stick_Left
            End If
            
        ElseIf StickiHasState(iAi, stick_Right) = False Then
            AddStickiState iAi, stick_Right
        End If
    End If
Else
    If StickiHasState(iAi, stick_Left) Then
        SubStickiState iAi, stick_Left
    ElseIf StickiHasState(iAi, stick_Right) Then
        SubStickiState iAi, stick_Right
    End If
End If


If bJump Then
    If modStickGame.sv_AIMove Then
        If Stick(iTarget).OnSurface Then
            If Stick(iAi).Speed < 20 Then
                AddStickiState iAi, stick_Jump
            End If
        End If
    End If
'ElseIf StickiHasState(iAI, stick_Jump) Then
    'SubStickiState iAI, stick_Jump
End If


End Sub

Private Function AnglesRoughlyEqual(A1 As Single, A2 As Single) As Boolean

AnglesRoughlyEqual = (Round(A1, 1) = Round(A2, 1))

End Function

Private Sub SetAINadeDelay(i As Integer)

Stick(i).AINadeDelay = Nade_Delay * 5 * Rnd()

End Sub

Private Function ClosestTargetI(iSource As Integer, DistToTarget As Single) As Integer

Dim i As Integer
Dim Dist As Single, TestDist As Single
Dim iCurrent As Integer 'current stick with least dist to iSource
Dim jSpy As Integer
Dim bCan As Boolean
Const AI_Bullet_Wait_Time = 5000 'AI can see you for 1 second after you shoot

Dist = StickGameWidth + 100
iCurrent = -1

For i = 0 To NumSticksM1
    If i <> iSource Then
        If StickInGame(i) Then
            
            
            If Stick(i).Perk = pSpy Then
                If Stick(i).WeaponType <> Chopper Then
                    jSpy = FindStick(Stick(i).MaskID)
                    If jSpy = -1 Then jSpy = 0
                Else
                    jSpy = i
                End If
                
                bCan = Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team) Or (jSpy = iSource)
                
                
                
            ElseIf Stick(i).Perk = pStealth Then
                jSpy = i
                
                If StickiHasState(i, Stick_Prone) Then
                    If Stick(i).Speed = 0 Then
                        If Stick(i).LastBullet + AI_Bullet_Wait_Time < GetTickCount() Then
                            bCan = False
                        ElseIf Stick(i).bSilenced Then
                            If StickiHasState(i, Stick_Fire) = False Then
                                bCan = False
                            Else
                                bCan = Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team)
                            End If
                        Else
                            bCan = Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team)
                        End If
                    Else
                        bCan = Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team)
                    End If
                Else
                    bCan = Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team)
                End If
                
                
            Else
                jSpy = i
                
                bCan = Not IsAlly(Stick(jSpy).Team, Stick(iSource).Team)
                
            End If
            
            If bCan Then
                bCan = StickCanSeeStick(iSource, i)
            End If
            If bCan Then
                bCan = Not StickInSmoke(i)
            End If
            If bCan Then
                If Stick(i).Speed < 20 Then
                    If StickiHasState(i, Stick_Fire) = False Then
                        bCan = Not StickInTBox(i)
                    End If
                End If
            End If
            
            
            If bCan Then
                TestDist = GetDist(Stick(iSource).X, Stick(iSource).Y, Stick(i).X, Stick(i).Y)
                
                
                If TestDist < Dist Then
                    iCurrent = i
                    
                    If Stick(i).bSilenced Then
                        Dist = TestDist * 1.5 'harder to see
                    Else
                        Dist = TestDist
                    End If
                    
                End If
            End If
            
            
        End If
    End If
Next i

ClosestTargetI = iCurrent
DistToTarget = Dist

End Function

Private Function StickInSmoke(iStick As Integer) As Boolean
Dim i As Integer
Const MinDist = 50, Inc = 0.5


If Stick(iStick).WeaponType <> Chopper Then
    
    For i = 0 To NumLargeSmokes - 1
        
        If GetDist(Stick(iStick).X, Stick(iStick).Y, LargeSmoke(i).CentreX, LargeSmoke(i).CentreY) < _
            (MinDist + Inc * LargeSmoke(i).iSize) Then
            
            StickInSmoke = True
            Exit For
            
        End If
    Next i
End If


End Function

Private Function StickInTBox(iStick As Integer) As Boolean
Dim i As Integer
Dim sY As Single

For i = 0 To modStickGame.ntBoxes
    If Stick(iStick).X > tBox(i).Left Then
        If Stick(iStick).X < (tBox(i).Left + tBox(i).width) Then
            
            sY = GetStickY(iStick)
            If Stick(iStick).Perk = pStealth Then
                If StickiHasState(iStick, stick_crouch) Then
                    sY = sY + ArmLen
                End If
            End If
            
            
            If sY > tBox(i).Top Then
                If sY < (tBox(i).Top + tBox(i).height) Then
                    StickInTBox = True
                    Exit For
                End If
            End If
        End If
    End If
Next i

End Function

Private Sub DrawPlatforms()
Dim i As Integer

picMain.FillStyle = vbFSTransparent

For i = 0 To nPlatforms
    modStickGame.sBoxFilled Platform(i).Left, Platform(i).Top, _
        Platform(i).Left + Platform(i).width, Platform(i).Top + Platform(i).height, BoxCol
    
    'PrintStickText "Platform " & CStr(i), Platform(i).Left, Platform(i).Top, vbBlack
Next i

picMain.DrawWidth = 5
modStickGame.sBox 0, 0, StickGameWidth, StickGameHeight, BoxCol
'modStickGame.sLine 0, 0, StickGameWidth, 0, Me.BackColor

picMain.DrawWidth = 1
End Sub

Private Sub DrawBoxes()
Dim i As Integer

picMain.DrawWidth = 2
For i = 0 To nBoxes
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
For i = 0 To ntBoxes
    modStickGame.sBoxFilled tBox(i).Left, tBox(i).Top, _
        tBox(i).Left + tBox(i).width, tBox(i).Top + tBox(i).height, BoxCol
    
    'PrintStickText "tBox " & CStr(i), tBox(i).Left, tBox(i).Top, vbBlack
Next i
End Sub

Private Sub DrawCrosshair()

If StickInGame(0) Then
    If Stick(0).bFlashed = False Then
        
        picMain.DrawWidth = 2
        Me.picMain.ForeColor = vbBlack 'Stick(0).Colour
        
        If modStickGame.cg_LaserSight Then
            'picMain.Line (MouseX, Mousey - 20)-(MouseX, Mousey + 20)
            'picMain.Line (MouseX + 20, Mousey)-(MouseX - 20, Mousey)
            picMain.Circle (MouseX, MouseY), 20, vbBlack
            
        Else
            
            'If Stick(0).bFlashed Then
                'StunnedMouseX = StunnedMouseX + 10 * Rnd()
                'StunnedMouseY = StunnedMouseY + 10 * Rnd()
                
                'MouseX = StunnedMouseX
                'Mousey = StunnedMouseY
            'Else
                'MouseX = MouseX
                'Mousey = MouseY
            'End If
            
            
            Select Case Stick(0).WeaponType
                Case eWeaponTypes.Shotgun, eWeaponTypes.FlameThrower
                    picMain.Circle (MouseX, MouseY + 150), 200
                    
                Case eWeaponTypes.AK, eWeaponTypes.SCAR, eWeaponTypes.M249, eWeaponTypes.Chopper, eWeaponTypes.SA80
                    picMain.Circle (MouseX, MouseY), 90
                    
                    picMain.Line (MouseX, MouseY - 150)-(MouseX, MouseY - 50)
                    picMain.Line (MouseX, MouseY + 150)-(MouseX, MouseY + 50)
                    
                    picMain.Line (MouseX - 150, MouseY)-(MouseX - 50, MouseY)
                    picMain.Line (MouseX + 150, MouseY)-(MouseX + 50, MouseY)
                    
                Case eWeaponTypes.DEagle
                    picMain.Circle (MouseX, MouseY), 90
                    
                    picMain.Line (MouseX, MouseY - 150)-(MouseX, MouseY - 75)
                    picMain.Line (MouseX, MouseY + 150)-(MouseX, MouseY + 75)
                    
                    picMain.Line (MouseX - 150, MouseY)-(MouseX - 75, MouseY)
                    picMain.Line (MouseX + 150, MouseY)-(MouseX + 75, MouseY)
                    
                Case eWeaponTypes.M82
                    
                    picMain.Circle (MouseX, MouseY), 75
                    
                    picMain.DrawWidth = 1
                    picMain.Line (MouseX, MouseY - 150)-(MouseX, MouseY + 150)
                    picMain.Line (MouseX + 150, MouseY)-(MouseX - 150, MouseY)
                    
                Case eWeaponTypes.RPG
                    picMain.Circle (MouseX, MouseY + 50), 100
                    picMain.Line (MouseX + 150, MouseY + 100)-(MouseX - 150, MouseY + 100)
                    picMain.Line (MouseX + 100, MouseY + 150)-(MouseX - 100, MouseY + 150)
                    picMain.Line (MouseX + 50, MouseY + 200)-(MouseX - 50, MouseY + 200)
                    
                    
                Case eWeaponTypes.Knife
                    
                    picMain.Line (MouseX, MouseY - 150)-(MouseX, MouseY + 150)
                    picMain.Line (MouseX + 150, MouseY)-(MouseX - 150, MouseY)
                    
            End Select
            
        End If
        
        
        
        'Me.modstickgame.sCircle (MouseX, Mousey), 150, vbRed
        'picMain.Line (MouseX - 3, Mousey - 100,MouseX - 3, Mousey + 100), Stick(0).Colour
        'picMain.Line (MouseX - 100, Mousey - 3,MouseX + 100, Mousey - 3), Stick(0).Colour
    End If
End If

End Sub

Private Sub DisplayChat()

Dim i As Integer, iChat As Integer, iKill As Integer
Dim FinalTxt As String
Dim nChat As Integer, nKill As Integer

'Check if any chat texts have decayed
i = 0
Do While i < NumChat
    'Is it decay time?
    If Chat(i).Decay < GetTickCount() Then
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

If bPlaying = False Then
    iChat = iChat + 7
    iKill = iKill + 7
End If

For i = 0 To NumChat - 1
    If Chat(i).bChatMessage Then
        'show it in the chat section
        FinalTxt = Chat(i).Text
        PrintStickFormText FinalTxt, 10, iChat * TextHeight(FinalTxt) + 1000, Chat(i).Colour
        iChat = iChat + 1
    Else
        'show it in the kill section
        FinalTxt = Chat(i).Text
        PrintStickFormText FinalTxt, 10, iKill * TextHeight(FinalTxt) + 3750, Chat(i).Colour
        iKill = iKill + 1
    End If
Next i

End Sub


Private Function GetReloadTime(iStick As Integer) As Long

'Select Case Stick(iStick).WeaponType
'    Case eWeaponTypes.AK
'        GetReloadTime = AK_Reload_Time
'    Case eWeaponTypes.M82
'        GetReloadTime = M82_Reload_Time
'    Case eWeaponTypes.DEagle
'        GetReloadTime = DEagle_Reload_Time
'    Case eWeaponTypes.Shotgun
'        GetReloadTime = Shotgun_Reload_Time
'    Case eWeaponTypes.SCAR
'        GetReloadTime = SCAR_Reload_Time
'    Case eWeaponTypes.M249
'        GetReloadTime = M249_Reload_Time
'    Case eWeaponTypes.RPG
'        GetReloadTime = RPG_Reload_Time
'    Case eWeaponTypes.FlameThrower
'        GetReloadTime = Flame_Reload_Time
'    Case eWeaponTypes.Chopper
'        GetReloadTime = 1
'End Select

GetReloadTime = kReloadTime(Stick(iStick).WeaponType)

If Stick(iStick).Perk = pSleightOfHand Then
    GetReloadTime = GetReloadTime / _
        (modStickGame.sv_StickGameSpeed * SleightOfHandReloadDecrease)
Else
    GetReloadTime = GetReloadTime / modStickGame.sv_StickGameSpeed
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
        Case eWeaponTypes.Shotgun
            kReloadTime(i) = Shotgun_Reload_Time
        Case eWeaponTypes.SCAR
            kReloadTime(i) = SCAR_Reload_Time
        Case eWeaponTypes.M249
            kReloadTime(i) = M249_Reload_Time
        Case eWeaponTypes.RPG
            kReloadTime(i) = RPG_Reload_Time
        Case eWeaponTypes.FlameThrower
            kReloadTime(i) = Flame_Reload_Time
        Case eWeaponTypes.SA80
            kReloadTime(i) = SA80_Reload_Time
        Case eWeaponTypes.Chopper
            kReloadTime(i) = 1
    End Select
Next i

End Sub

Private Sub ShowScores()
Const Sp8 As String * 8 = "        "
Const TopY = CentreY - 3000
Const TitleOffset = 290
Dim Txt As String
Dim i As Integer
Dim X As Single, Y As Single

picMain.Font.Size = 11
'picMain.Font.Bold = True

On Error Resume Next
X = StickCentreX - 2000 '#######################THIS NEEDS UPDATING WHEN ADDING NEW COLUMNS###########################

Select Case modStickGame.sv_GameType
    Case eStickGameTypes.gCoOp, eStickGameTypes.gElimination
        BorderedBox X, TopY - 300, X + 8000, Y + 195 * NumSticks + 1100, BoxCol
    Case Else
        BorderedBox X, TopY - 300, X + 7000, Y + 195 * NumSticks + 1100, BoxCol
End Select

'#################################################################################
PrintStickFormText Sp8 & "Name" & Sp8, X, TopY - TitleOffset, vbBlack
'                                 290 = TextHeight(Txt)*1.5

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(Stick(i).Name), 20)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).Colour
Next i
'#################################################################################
X = X + 1500
PrintStickFormText " Kills ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(CStr(Stick(i).iKills)), 6)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).Colour
Next i
'#################################################################################
X = X + 1000
PrintStickFormText "Deaths ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(Trim$(CStr(Stick(i).iDeaths)), 8)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).Colour
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
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).Colour
Next i
'#################################################################################
X = X + 1000
PrintStickFormText "     Perk   ", X, TopY - TitleOffset, vbBlack

For i = 0 To NumSticksM1
    Txt = CentreFill(GetPerkName(Stick(i).Perk), 16)
    PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, Stick(i).Colour
Next i
'#################################################################################
Select Case modStickGame.sv_GameType
    Case gElimination, gCoOp
        X = X + 1500
        PrintStickFormText " Dead ", X, TopY - TitleOffset, vbBlack
        
        For i = 0 To NumSticksM1
            Txt = IIf(StickInGame(i), vbNullString, "  ----")
            PrintStickFormText Txt, X, TopY + TextHeight(Txt) * i, GetTeamColour(Stick(i).Team)
        Next i
End Select



'#################################################################################
'Extra Stuff
'#################################################################################
If modStickGame.sv_GameType = gDeathMatch Then
    PrintStickFormText "Score To Win: " & CStr(modStickGame.sv_WinScore), 6000, 20, vbBlack
End If

PrintStickFormText "Game Type: " & GetGameType(), 9000, 20, vbBlack

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
    ElseIf i = pRadarJammer Then
        kPerkName(i) = "Radar Jammer"
    ElseIf i = pSleightOfHand Then
        kPerkName(i) = "Sleight of Hand"
    ElseIf i = pSpy Then
        kPerkName(i) = "Spy"
    ElseIf i = pStealth Then
        kPerkName(i) = "Stealth"
    ElseIf i = pStoppingPower Then
        kPerkName(i) = "Stopping Power"
    End If
Next i

End Sub

Private Sub DrawNames()
Dim Txt As String
Dim i As Integer, j As Integer
Dim Col As Long
Dim sName As String

'Txt = Trim$(Stick(0).Name) & IIf(StickiHasState(0, Stick_Reload), " (Reloading)", vbNullString)
'PrintStickText Txt, Stick(0).X - TextWidth(Txt) / 2, GetStickY(0) - 250, vbBlack

For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        If CanSeeStick(i) Then
            
            If CBool(i) Then
                If IsAlly(Stick(0).Team, Stick(i).Team) Then
                    Col = Ally_Colour
                Else
                    Col = Enemy_Colour
                End If
            Else
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
                
                
                
                If StickiHasState(i, Stick_Prone) Then
                    
                    If Stick(i).Perk <> pStealth Then
                        Txt = sName & " - " & CStr(Stick(i).Health)  '& _
                            IIf(Stick(i).Armour > 0, ":" & CStr(Stick(i).Armour), vbNullString)
                        
                        PrintStickText Txt, Stick(i).X - TextWidth(Txt) / 2, GetStickY(i) - 250, Col
                    End If
                    
                Else
                    If Not (Stick(i).Perk = pStealth And StickiHasState(i, stick_crouch)) Then
                        Txt = sName & IIf(StickiHasState(i, Stick_Reload), " (Reloading)", vbNullString)
                        PrintStickText Txt, Stick(i).X - TextWidth(Txt) / 2, GetStickY(i) - 250, Col
                        
                        Txt = "Health: " & CStr(Stick(i).Health) '& IIf(Stick(i).Armour > 0, " Armour: " & CStr(Stick(i).Armour), vbNullString)
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
Dim X As Single, Y As Single
Dim mR As Integer, MaxRounds As Integer, Reload_Time As Long
Dim i As Integer
Dim bChopper As Boolean

Dim SemiX As Single, SemiY As Single

'Txt = "Health: " & CStr(Round(Stick(0).Health))
'PrintStickText Txt, Me.width / 2 - TextWidth(Txt) / 2, TextHeight(Txt) + 500, vbblack

If modStickGame.cg_DrawFPS Then
    PrintStickFormText "FPS: " & CStr(FPS), 10, 600, vbBlack
    'PrintStickText "Elapsed: " & CStr(modStickGame.StickElapsedTime), 10, 160, vbBlack
End If


If ShowScoresKey Then
    ShowScores
Else
    DisplayScoreBoard
End If


If LastZoomPress + ZoomShowTime > GetTickCount() Then
    PrintStickFormText "Zoom: " & FormatNumber$(cg_sZoom, 2, vbTrue, vbFalse, vbFalse), _
        StickCentreX, StickCentreY - 4000, vbBlack
End If

If Scroll_WeaponKey <> Stick(0).WeaponType Then
    Txt = "Switching Weapon: " & GetWeaponName(Scroll_WeaponKey) '& Space$(2) & _
        CStr(Format((Scroll_Delay - GetTickCount() + LastScrollWeaponSwitch) / 1000, "0.00"))
    
    PrintStickFormText Txt, StickCentreX - TextWidth(Txt) / 2, StickCentreY - 1000, vbBlack
End If

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
            'If Stick(0).WeaponType = Shotgun Then
                'mR = (MaxRounds - Stick(0).BulletsFired) / Shotgun_Gauge
            'Else
                mR = MaxRounds - Stick(0).BulletsFired
            'End If
            
            If StickiHasState(0, Stick_Reload) Then
                
                Reload_Time = GetReloadTime(0)
                
                TimeLeft = (Reload_Time - GetTickCount() + Stick(0).ReloadStart)
                
                DrawSemiCircle SemiX, SemiY, _
                    vbBlue, vbRed, 1 - TimeLeft / Reload_Time, 600
                
        '        x = StickCentreX - Reload_Time / 2
        '        y = StickCentreY - 650
        '
        '        picMain.Line (x, y)-(x + Reload_Time, y), vbRed
        '        picMain.Line (x, y)-(x + Reload_Time - TimeLeft, y), vbBlue
                
            Else
                
                DrawSemiCircle SemiX, SemiY, _
                    vbBlue, vbRed, mR / MaxRounds, 600
                
            End If
            
            If Stick(0).WeaponType = Shotgun Then mR = mR / Shotgun_Gauge
            
            Txt = "Rounds: " & CStr(mR)
            PrintStickFormText Txt, SemiX - TextWidth(Txt) / 2, SemiY - 300, vbBlack
            
            If StickiHasState(0, Stick_Reload) = False Then
                If Stick(0).WeaponType <> RPG Then
                    If Stick(0).WeaponType = Shotgun Then
                        If mR < (MaxRounds * 0.3 / Shotgun_Gauge) Then
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
            PrintStickFormText Txt, SemiX - TextWidth(Txt) / 2, SemiY - 1500, vbBlack
            
            If TotalMags(Stick(0).WeaponType) = 0 Then
                If Stick(0).WeaponType <> RPG Or Stick(0).BulletsFired = 1 Then
                    Txt = "No " & GetMagName(Stick(0).WeaponType) & " left - Find some ammo"
                    PrintStickFormText Txt, StickCentreX - 1300, CentreY - 900, vbRed
                    'PrintStickFormText Txt, SemiX - TextWidth(Txt) / 2, SemiY - 1750, vbRed
                End If
            End If
            
            
        End If 'weapon type endifs
    End If
    
    
    
    If Stick(0).WeaponType <> Chopper Then
        
        X = SemiX - 700
        Y = SemiY - 950
        
        Me.picMain.FontBold = True
        Me.picMain.DrawWidth = 3
        
        If Stick(0).LastNade + Nade_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            Txt = GetNadeTypeName() & " Ready" 'IIf(Stick(0).iNadeType = nFrag, "Grenade Ready", "Flash-Bang Ready")
            PrintStickFormText Txt, X - TextWidth(Txt) / 2, Y, vbGreen
            'C = MGreen
        Else
            'Txt = "Grenade Not Ready"
            
            TimeLeft = Nade_Delay + (Stick(0).LastNade - GetTickCount()) * modStickGame.sv_StickGameSpeed
            
            Me.picMain.Line (SemiX - 1900, SemiY - 850)-(SemiX - 1900 + TimeLeft / 2, SemiY - 850), vbRed
            
        End If
        
        Y = Y - 250
        If Stick(0).LastMine + Mine_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            Txt = "Mine Ready"
            PrintStickFormText Txt, X - TextWidth(Txt) / 2, Y, vbGreen
            'C = MGreen
        Else
            'Txt = "GreMine Not Ready"
            
            TimeLeft = Mine_Delay + (Stick(0).LastMine - GetTickCount()) * modStickGame.sv_StickGameSpeed
            
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
        
        
        If Stick(0).LastNade + Chopper_RPG_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            Txt = "Rocket Ready"
            PrintStickFormText Txt, X + 550 - TextWidth(Txt) / 2, Y - 100, vbGreen
        Else
            TimeLeft = Chopper_RPG_Delay + (Stick(0).LastNade - GetTickCount()) * modStickGame.sv_StickGameSpeed
            
            Me.picMain.Line (X, Y)-(X + TimeLeft / 2, Y), vbRed
        End If
        
        Me.picMain.FontBold = False
        
        X = Me.ScaleWidth - 750
    Else
        X = Me.ScaleWidth - 2000
        If ChopperAvail Then
            picMain.Font.Bold = True
            PrintStickFormText "Chopper Available - Press 0", 10, 260, vbBlack
            picMain.Font.Bold = False
        End If
    End If
    Txt = "Health: " & CStr(Stick(0).Health)
    PrintStickFormText Txt, X - TextWidth(Txt) / 2, SemiY - 350, vbBlack
    DrawSemiCircle X, SemiY, MGreen, vbBlue, Stick(0).Health / Health_Start, 600
    
    
    Txt = "Armour: " & CStr(Stick(0).Armour)
    PrintStickFormText Txt, X - TextWidth(Txt) / 2, SemiY - 200, vbBlack
    DrawSemiCircle X, SemiY, MSilver, vbBlue, Stick(0).Armour / Max_Armour, 675
    
    If ShowScoresKey = False Then
        Txt = "Kills in a Row: " & CStr(Stick(0).iKillsInARow)
        PrintStickFormText Txt, 10, 10, vbBlack
    End If
End If


If Not modStickGame.StickServer Then
    If (LastUpdatePacket + LagOut_Delay) < GetTickCount() Then
        If LastUpdatePacket Then
            Me.Font.Size = 16
            Me.ForeColor = &HC0C0C0
            
            PrintStickFormText "Connection Interrupted", StickCentreX - 750, StickCentreY - 2000, vbRed '&HC0C0C0
            
            Me.Font.Size = 8
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

Public Function GetWeaponName(vWeapon As eWeaponTypes) As String

GetWeaponName = kWeaponName(vWeapon)

End Function

Private Sub MakeWeaponNameArray()
Dim i As Integer

For i = 0 To eWeaponTypes.Chopper
    If i = SCAR Then
        kWeaponName(i) = "SCAR"
    ElseIf i = AK Then
        kWeaponName(i) = "AK-47"
    ElseIf i = M249 Then
        kWeaponName(i) = "M249 SAW"
    ElseIf i = M82 Then
        kWeaponName(i) = "M82 Sniper"
    ElseIf i = RPG Then
        kWeaponName(i) = "RPG"
    ElseIf i = Shotgun Then
        kWeaponName(i) = "Shotgun"
    ElseIf i = DEagle Then
        kWeaponName(i) = "Desert Eagle"
    ElseIf i = FlameThrower Then
        kWeaponName(i) = "FlameThrower"
    ElseIf i = Knife Then
        kWeaponName(i) = "Sword"
    ElseIf i = SA80 Then
        kWeaponName(i) = "SA80 A2"
    ElseIf i = Chopper Then
        kWeaponName(i) = "Chopper"
    End If
Next i

End Sub

Private Sub DrawSemiCircle(tX As Single, tY As Single, _
    ForeCol As Long, BackCol As Long, _
    sAmountFull As Single, sWidth As Single)

Dim Start As Single
picMain.DrawWidth = 3

If sAmountFull < 1 Then
    picMain.Circle (tX, tY), sWidth, BackCol, 0, pi, 0.75
End If

If sAmountFull Then
    On Error Resume Next
    Start = pi * (1 - sAmountFull)
    If Start > (pi - 0.1) Then 'pi*179 / 180
        Start = pi - 0.1
    End If
    
    picMain.Circle (tX, tY), sWidth, ForeCol, Abs(Start), pi, 0.75
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
Dim aFacing As Single

If modStickGame.cg_LaserSight Then
    If StickInGame(0) Then
        If Stick(0).bFlashed = False Then
            If StickiHasState(0, Stick_Reload) = False Then
                'If StickiHasState(0, Stick_Prone) = False Then
                    If Stick(0).WeaponType <> Knife Then
                        If Stick(0).WeaponType <> Chopper Then
                            'If Stick(0).LastBullet + GetBulletDelay(0) < GetTickCount() Then
                                'aFacing = GetStickFacing(0)
                            'Else
                                'aFacing = Stick(0).Facing
                            'End If
                            
                            'If StickiHasState(0, Stick_Prone) Then
                                'aFacing = Stick(0).Facing - Sin(Stick(0).Facing) / 10
                            'Else
                                aFacing = Stick(0).Facing
                            'End If
                            
                            'aFacing = FindAngle(CSng(Stick(0).GunPoint.x), CSng(Stick(0).GunPoint.y), MouseX, MouseY)
                            
                            'Facing = FindAngle(CSng(Stick(0).GunPoint.x), CSng(Stick(0).GunPoint.y), MouseX, MouseY) 'Stick(0).Facing
                            '.0349 = pi*2/180 = 2 degrees
                            
                            
                            Facing = aFacing - Sin(aFacing) * 0.038 '0.0349
                            
                            OldtX = CSng(Stick(0).GunPoint.X) '+ SF100
                            OldtY = CSng(Stick(0).GunPoint.Y) '+ SF100
                            
                            picMain.DrawWidth = 1
                            
                            'gradient'd line
                            For i = 1 To nLines
                                
                                tX = Stick(0).GunPoint.X + Sin(Facing) * Laser_Len * i / nLines '+ SF100
                                tY = Stick(0).GunPoint.Y - Cos(Facing) * Laser_Len * i / nLines '+ SF100
                                
                                modStickGame.sLine OldtX, OldtY, tX, tY, RGB(255 - i * 10, i * 10, i * 30)
                                
                                OldtX = tX
                                OldtY = tY
                                
                            Next i
                        End If 'chopper endif
                    End If 'knife endif
                'End If 'prone endif
            End If 'reload endif
        End If 'flashed endif
    End If 'ingame endif
End If

End Sub

Private Sub DrawSilencer(X As Single, Y As Single, Facing As Single)
Dim X2 As Single, Y2 As Single
Const SilencerLen = 120

X2 = X + SilencerLen * Sin(Facing)
Y2 = Y - SilencerLen * Cos(Facing)

picMain.DrawWidth = 2
modStickGame.sLine X, Y, X2, Y2, vbBlack

End Sub

Private Sub Physics()

Dim i As Integer, j As Integer, BulletOwnerIndex As Integer
Dim TempMag As Single
Dim TempDir As Single, BaseTempDir As Single
Dim Tmp As Integer
Dim MaxSpeed As Integer
Dim Bullet_Delay As Long
Dim Stick_Moving As Boolean, bLBound As Boolean
Dim XComp As Single, YComp As Single
Dim bThrow As Boolean

Dim tX As Single, tY As Single


i = 1
Do While i < NumSticks
    'Skip the local Stick
    If Stick(i).ID <> MyID Then
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
    End If
    'Increment counter
    i = i + 1
Loop


'check if i lagged out
If Not modStickGame.StickServer Then
    If (LastUpdatePacket + mPacket_LAG_KILL * 2) < GetTickCount() Then
        If LastUpdatePacket Then
            bRunning = False
            AddText "Error - Lagged Out (No Packet Flow)", TxtError, True
            Exit Sub
        End If
    End If
End If


'With Stick(0)
'    modstickgame.sLine .X, .Y + BodyLen,.X + 1000 * Sin(.ActualFacing), .Y - 1000 * Cos(.ActualFacing)), vbRed
'End With

'Loop through each Stick and perform physics

For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        
        If i = 0 Or Stick(i).IsBot Then
            DoReload i
        End If
        
        
        'modStickGame.sCircle CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y), 50, vbRed
        
        Stick_Moving = False
        bLBound = True 'should be true
        
        'Check lag tol
        If Stick(i).ID <> MyID Then
            If Not Stick(i).IsBot Then
                If (Stick(i).LastPacket + mPacket_LAG_TOL) < GetTickCount() Then
                    SetStickiState i, Stick_None
                End If
            End If
        End If
        
        Stick(i).Facing = FixAngle(Stick(i).Facing)
        Stick(i).ActualFacing = FixAngle(Stick(i).ActualFacing)
        
        'Firing
        If (Stick(i).State And Stick_Reload) = 0 Then
            If Stick(i).State And Stick_Fire Then
                
                Bullet_Delay = GetBulletDelay(i)
                
                
                If Stick(i).BulletsFired < GetMaxRounds(Stick(i).WeaponType) Then
                    
                    
                    If Stick(i).LastBullet + Bullet_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                        
                        
                        tX = Stick(i).GunPoint.X
                        tY = Stick(i).GunPoint.Y
                        
                        If Stick(i).WeaponType <> Knife Then
                            
                            If Stick(i).WeaponType <> M82 Then
                                If Stick(i).WeaponType <> RPG Then
                                    If Stick(i).WeaponType <> Chopper Then
                                        If Stick(i).WeaponType <> FlameThrower Then
                                            AddVectors Stick(i).Speed / 4, Stick(i).Heading, BULLET_SPEED, Stick(i).ActualFacing, TempMag, BaseTempDir
                                        Else
                                            'flame only
                                            AddVectors Stick(i).Speed / Flame_Inertia_Reduction, Stick(i).Heading, Flame_Speed, Stick(i).ActualFacing, TempMag, BaseTempDir
                                        End If
                                    End If
                                End If
                            End If
                            
                            'sin because when facing up/down, it is dead on (and sin 0 or sin 180 = 0)
                            If Stick(i).ActualFacing > pi Then
                                Tmp = -1
                                Stick(i).RecoilLeft = True
                            Else
                                Tmp = 1
                                Stick(i).RecoilLeft = False
                            End If
                            
                            
                            
                            If Stick(i).WeaponType = SA80 Then
                                If Stick(i).BulletsFired2 < SA80_Burst_Bullets Then
                                    FireShot i, BaseTempDir, TempMag, Tmp, tX, tY, Stick_Moving, bLBound
                                    
                                    Stick(i).LastBullet = GetTickCount()
                                    
                                ElseIf Stick(i).LastBullet + SA80_Bullet_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                                    Stick(i).BulletsFired2 = 0
                                    
                                End If
                                
                            Else
                                FireShot i, BaseTempDir, TempMag, Tmp, tX, tY, Stick_Moving, bLBound
                                Stick(i).LastBullet = GetTickCount()
                            End If
                            
                            
                            
                        Else 'If Stick(i).WeaponType = Knife Then 'knife else
                            
                            'have knife and are 'firing'
                            
                            For j = 0 To NumSticksM1
                                If j <> i Then
                                    If StickInGame(j) Then
                                        If Not IsAlly(Stick(j).Team, Stick(i).Team) Then
                                            If StickInvul(j) = False Then
                                                If CoOrdNearStick(CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y), j) Then
                                                    AddBloodExplosion CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y)
                                                    
                                                    If Stick(j).ID = MyID Or Stick(j).IsBot Then
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
                        If FireKey = False Then
                            If FireKeyUpTime + Bullet_Release_Delay < GetTickCount() Then
                                SubStickiState i, Stick_Fire
                                Stick(i).BulletsFired2 = 0
                            End If
                        End If
                    End If
                    
                    
                End If 'bullets fired
                
            End If 'state_fire
            
        End If 'state_reload
        
        
        
        If Stick(i).Perk = pConditioning Then
            TempMag = Accel * ConditiongSpeedIncrease
        Else
            TempMag = Accel
        End If
        
        
        If Stick(i).State And stick_Left Then
            'If Stick(i).OnSurface Then
                'Apply acceleration
                AddVectors Stick(i).Speed, Stick(i).Heading, TempMag, CSng(pi3D2), Stick(i).Speed, Stick(i).Heading
            'End If
            
            Stick_Moving = True
            
        ElseIf Stick(i).State And stick_Right Then
            'If Stick(i).OnSurface Then
                'Apply reverse acceleration
                AddVectors Stick(i).Speed, Stick(i).Heading, TempMag, CSng(piD2), Stick(i).Speed, Stick(i).Heading
            'End If
            
            Stick_Moving = True
            
        End If
        
        
        
        If Stick(i).WeaponType <> Chopper Then
            Call DoRecoil(i, Stick_Moving, bLBound)
            
            
            If Stick(i).State And Stick_Nade Then
                
                If Stick(i).LastNade + Nade_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                    
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
                            Stick(i).ActualFacing + IIf(Stick(i).ActualFacing > pi, piD6, -piD6) * Abs(Sin(Stick(i).ActualFacing)), _
                            TempMag, TempDir
                        
                        
                        AddNade Stick(i).X, Stick(i).Y, TempDir, TempMag, i, Stick(i).Colour, Stick(i).iNadeType
                        
                        'If Stick(i).iNadeType = nSmoke Then
                            'Stick(i).LastNade = GetTickCount() + Nade_Time / 2
                        'Else
                        Stick(i).LastNade = GetTickCount()
                        'End If
                        
                    End If
                    
                ElseIf Stick(i).LastNade + Nade_Release_Delay < GetTickCount() Then
                    SubStickiState i, Stick_Nade
                    
                End If
                
            ElseIf Stick(i).State And Stick_Mine Then
                
                If Stick(i).LastMine + Mine_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                    
                    'If Stick(i).OnSurface Then
                    AddMine Stick(i).X, Stick(i).Y + Mine_Y_Increase, Stick(i).ID, Stick(i).Colour, Stick(i).Heading, Stick(i).Speed
                    Stick(i).LastMine = GetTickCount()
                    'End If
                    
                ElseIf Stick(i).LastMine + Nade_Release_Delay < GetTickCount() Then
                    SubStickiState i, Stick_Mine
                    
                    If i = 0 Then
                        MineKey = False
                    End If
                    
                End If
                
            End If
            
            
            If Stick(i).State And stick_Jump Then
                
                Stick_Moving = True
                
                If Stick(i).OnSurface = False Then 'Stick(i).StartJumpTime + JumpTime < GetTickCount() Then
                    If i = 0 Then JumpKey = False
                    
                    SubStickiState i, stick_Jump
                    
                    'Debug.Print "Sub " & Rnd & vbNewLine
                Else
                    'If StickiHasState(i, stick_Left) Then SubStickiState i, stick_Left
                    'If StickiHasState(i, stick_Right) Then SubStickiState i, stick_Right
                    
                    AddVectors Stick(i).Speed, Stick(i).Heading, Accel * JumpMultiple, 0, Stick(i).Speed, Stick(i).Heading
                    'Debug.Print "AV " & Rnd
                End If
                
            End If
            
            
            If Stick(i).State And Stick_Reload Then
                If Stick(i).bHadMag = False Then
                    AddMagForStick i
                    Stick(i).bHadMag = True
                End If
            ElseIf Stick(i).bHadMag Then
                Stick(i).bHadMag = False
            End If
            
            
            ApplyGravity i, bLBound
            
            If StickiHasState(i, stick_crouch) Then
                MaxSpeed = Max_Speed / 4
            ElseIf StickiHasState(i, Stick_Prone) Then
                MaxSpeed = Max_Speed / 8
            Else
                If Stick(i).Perk = pConditioning Then
                    MaxSpeed = Max_Speed * ConditioningMaxSpeedInc
                Else
                    MaxSpeed = Max_Speed
                End If
            End If
            
            
            
            If Stick_Moving = False Then
                'friction
                
                XComp = Stick(i).Speed * Sin(Stick(i).Heading)
                
                If Abs(XComp) > 4 Then
                    XComp = XComp / 1.2
                    
                    YComp = Stick(i).Speed * Cos(Stick(i).Heading)
                    
                    Stick(i).Speed = Sqr(XComp ^ 2 + YComp ^ 2)
                    
                End If
            End If
            
            
        Else
            
            MaxSpeed = Chopper_Max_Speed
            
            If Stick(i).OnSurface = False Then Stick(i).OnSurface = True
            
            If StickiHasState(i, stick_Jump) Then
                'If (Stick(i).StartJumpTime + JumpTime) < GetTickCount() Then
                    'If i = 0 Then JumpKey = False
                    'SubStickiState i, Stick_Jump
                'Else
                    Stick_Moving = True
                    AddVectors Stick(i).Speed, Stick(i).Heading, Chopper_Lift, 0, Stick(i).Speed, Stick(i).Heading
                    If Stick(i).IsBot = False Then SubStickiState i, stick_Jump
                'End If
            ElseIf StickiHasState(i, stick_crouch) Then
                Stick_Moving = True
                AddVectors Stick(i).Speed, Stick(i).Heading, Chopper_Lift * 1.5, pi, Stick(i).Speed, Stick(i).Heading
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
            
            If Stick(i).Y > 12500 Then
                Stick(i).Y = 12500
                Stick(i).Speed = 0
            End If
            
            
            
            
            'chopper rockets
            
            
            If StickiHasState(i, Stick_Nade) Then
                If Stick(i).LastNade + Chopper_RPG_Delay < GetTickCount() Then
                    
                    AddNade CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y), Stick(i).ActualFacing, RPG_Speed, _
                        i, Stick(i).Colour, nFrag, True
                    
                    Stick(i).LastNade = GetTickCount()
                     
                ElseIf Stick(i).LastNade + Nade_Release_Delay < GetTickCount() Then
                    SubStickiState i, Stick_Nade
                End If
            End If
            
            
        End If
        
        
        'health pack
        Call CheckStickHealthPack(i)
        
        
        'Cap Speed
        If Stick(i).Speed > MaxSpeed Then Stick(i).Speed = MaxSpeed
        
        
        'stickmotion
        StickMotion Stick(i).X, Stick(i).Y, Stick(i).Speed, Stick(i).Heading
        
        'Wrap edges + speed
        Call ClipEdges(i, bLBound)
        
    End If 'stickingame endif
    
Next i


Call CheckChopperCollisions


'Loop through each bullet and perform physics
i = 0
Do While i < NumBullets
    'Move it!
    StickMotion Bullet(i).X, Bullet(i).Y, Bullet(i).Speed, Bullet(i).Heading
    
    'Wrap edges
    If ClipBullet(i) = False Then
        BulletOwnerIndex = FindStick(Bullet(i).Owner)
        
        If BulletOwnerIndex <> -1 Then
        'Check for collisions
            For j = 0 To NumSticksM1
                If StickInGame(j) Then
                    If Stick(j).ID <> Bullet(i).Owner Then
                        If Not IsAlly(Stick(j).Team, Stick(BulletOwnerIndex).Team) Then
                            If BulletInStick(j, i) Then 'GetDist(Stick(j).X, Stick(j).Y, Bullet(i).X, Bullet(i).Y) < (Bullet_Radius + BodyLen / 2) Then
                                
                                BulletHitStick i, j, BulletOwnerIndex
                                Exit For
                                
                                
                            End If 'bulletinstick endif
                            
                        End If 'ally endif
                    End If 'bulletID endif
                End If 'stickingame endif
            Next j
        End If 'ownerindex <> -1 endif
    Else
        i = i - 1
    End If 'clip endif
    'Increment counter
    i = i + 1
    
Loop


EH:
End Sub

Private Sub BulletHitStick(i As Integer, j As Integer, BulletOwnerIndex As Integer)
Dim F As Single, BHeading As Single
Dim bHeadShot As Boolean
Dim kType As eKillTypes
Dim bCan As Boolean

bCan = True

'if saber, deflect, else damage
If Stick(j).bLightSaber And Stick(j).WeaponType = Knife Then
    
    F = FixAngle(Stick(j).Facing)
    BHeading = FixAngle(Bullet(i).Heading)
    
    If BHeading > pi Then
        If F > 0 And F < piD8 Or F < pi And F > pi3D8 Then
            'deflect!
            bCan = False
        End If
    Else
        If F > pi And F < pi5D4 Or F > pi7D4 And F < pi2 Then
            'deflect!
            bCan = False
        End If
    End If
    
End If


If (0 = j Or Stick(j).IsBot) And bCan Then 'And modStickGame.StickServer) Then
  '^FindStick(MyID)
  'If this is our Stick...
    
    'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
    If StickInvul(j) = False Then
        bHeadShot = BulletInHead(j, i)
        
        If Bullet(i).bShotgunBullet Then
            AlterShotgunBulletDamage i
        End If
        
        
        If Stick(j).WeaponType <> Chopper Then
            If bHeadShot Then
                DamageStick Bullet(i).Damage * 3, j
            Else
                DamageStick Bullet(i).Damage, j
            End If
        ElseIf Bullet(i).bSniperBullet Then
            'sniper hitting a chopper
            DamageStick Bullet(i).Damage * 10, j
        Else
            'normal bullet hitting chopper
            DamageStick Bullet(i).Damage, j
        End If
        
        If bHeadShot Then
            kType = kHead
        ElseIf BulletOwnerIndex <> -1 Then
            If Stick(BulletOwnerIndex).bSilenced Then
                kType = kSilenced
            Else
                kType = kNormal
            End If
        Else
            kType = kNormal
        End If
        
        If kType = kHead Then
            If Stick(BulletOwnerIndex).WeaponType = Chopper Then
                kType = kNormal
            End If
        End If
        
        If Stick(j).Health < 1 Then
            Call Killed(j, BulletOwnerIndex, kType)
        End If
        
    End If 'spawn invul endif
    
    
    If Stick(j).WeaponType <> Chopper Then
        AddBlood Bullet(i).X, Bullet(i).Y, Bullet(i).Heading, (Stick(j).Armour > 0) 'add blood before removing t'bullet
    Else 'If Rnd() > 0.75 Then
        'AddExplosion Bullet(i).X, Bullet(i).Y, 250, 0.5, 0, 0
        AddCirc Bullet(i).X, Bullet(i).Y, 250, 0.25, vbYellow
        AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading - pi
    End If
    
ElseIf bCan = False Then
    
    AddSparks Bullet(i).X, Bullet(i).Y, Bullet(i).Heading - pi
    
End If 'stick(j) = me endif


If Bullet(i).bSniperBullet = False Or Stick(j).WeaponType = Chopper Then
    RemoveBullet i, False, bCan
End If

End Sub

Private Sub FireShot(i As Integer, BaseTempDir As Single, TempMag As Single, Tmp As Integer, _
    tX As Single, tY As Single, _
    Stick_Moving As Boolean, bLBound As Boolean)

Dim j As Integer
Dim TempDir As Single
Const AccuracyRedux = 2000 'bigger, the more accurate


BaseTempDir = BaseTempDir + PM_Rnd * Stick(i).Speed / AccuracyRedux


If Stick(i).WeaponType = FlameThrower Then
    
    AddFlame CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y), BaseTempDir, TempMag, Stick(i).ID, i
    
ElseIf Stick(i).WeaponType = Chopper Then
    
    AddBullet CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y), BULLET_SPEED * 1.5, _
        Stick(i).ActualFacing, Stick(i).ID, Chopper_Bullet_Damage, i
    
ElseIf Stick(i).WeaponType = Shotgun Then
    
    'adjust since bullets aren't sent for stick(i).pointapi
    BaseTempDir = BaseTempDir - Sin(Stick(i).ActualFacing) / IIf(StickiHasState(i, stick_crouch), 8, 10)
    
    For j = 1 To Shotgun_Gauge
        TempDir = BaseTempDir + Tmp * Rnd() * Shotgun_Spray_Angle
        AddBullet tX, tY, TempMag + Rnd() * 50, TempDir, Stick(i).ID, Shotgun_Bullet_Damage, i
    Next j
    
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * Shotgun_SingleRecoil_Angle
    
    'If i = 0 Then FireKey = False
    
    'recoil
    AddVectors Stick(i).Speed, Stick(i).Heading, Shotgun_RecoilForce, FixAngle(Stick(i).Facing - pi), Stick(i).Speed, Stick(i).Heading
    Stick_Moving = True
    bLBound = False
    
ElseIf Stick(i).WeaponType = AK Then
    
    TempDir = BaseTempDir - Sin(Stick(i).Facing) / 30 + Tmp * Rnd() * AK_Spray_Angle
    
    AddBullet tX, tY, TempMag, TempDir, Stick(i).ID, AK_Bullet_Damage, i
    
    'recoil
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * AK_SingleRecoil_Angle
    
ElseIf Stick(i).WeaponType = M249 Then
    
    
    TempDir = BaseTempDir - Sin(Stick(i).Facing) / 25 + Tmp * Rnd() * M249_Spray_Angle
    
    AddBullet tX, tY, TempMag, TempDir, Stick(i).ID, M249_Bullet_Damage, i
    
    'recoil
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * M249_SingleRecoil_Angle
    
    'AddVectors Stick(i).Speed, Stick(i).Heading, M249_RecoilForce, FixAngle(Stick(i).Facing - pi), Stick(i).Speed, Stick(i).Heading
    'Stick_Moving = True
    'bLBound = False
    
    
ElseIf Stick(i).WeaponType = DEagle Then
    
    If StickiHasState(i, stick_crouch) Then
        TempDir = BaseTempDir - Sin(Stick(i).Facing) / 20
    Else
        TempDir = BaseTempDir - Sin(Stick(i).Facing) / 100
    End If
    
    AddBullet tX, tY, TempMag, TempDir, Stick(i).ID, DEagle_Bullet_Damage, i, , True
    AddSmokeGroup tX + TempMag * Sin(TempDir), tY, 4, Rnd() * TempMag / 10, TempDir
    
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * DEagle_SingleRecoil_Angle
    
    'recoil
    AddVectors Stick(i).Speed, Stick(i).Heading, DEagle_RecoilForce, FixAngle(Stick(i).Facing - pi), Stick(i).Speed, Stick(i).Heading
    Stick_Moving = True
    bLBound = False
    
ElseIf Stick(i).WeaponType = M82 Then
    
    'due to recoil, by the time others find out i'm shooting,
    'i've moved back from where i was, so if it's a foreign stick, adjust the angle more
    
    BaseTempDir = Stick(i).ActualFacing
    
    'If i = 0 Then
        If StickiHasState(i, stick_crouch) Then
            TempDir = BaseTempDir - Sin(Stick(i).Facing) / 25
        Else
            TempDir = BaseTempDir - Sin(Stick(i).Facing) / 50
        End If
'    Else
'        'adjust for them
'        'Shot was going lower, so take off (or add on if facing > pi)
'
'        TempDir = BaseTempDir + IIf(Stick(i).Facing > pi, 1, -1) * Sin(Stick(i).Facing) / 8
'
'    End If
    
    AddBullet tX, tY, BULLET_SPEED, TempDir, Stick(i).ID, M82_Bullet_Damage, i, True
    
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * M82_SingleRecoil_Angle
    
    AddVectors Stick(i).Speed, Stick(i).Heading, M82_RecoilForce, FixAngle(Stick(i).Facing - pi), Stick(i).Speed, Stick(i).Heading
    Stick_Moving = True
    bLBound = False
    
ElseIf Stick(i).WeaponType = SCAR Then
    
    TempDir = BaseTempDir - Sin(Stick(i).Facing) / IIf(StickiHasState(i, stick_crouch), 20, 40) + Tmp * Rnd() * SCAR_Spray_Angle
    
    AddBullet tX, tY, TempMag, TempDir, Stick(i).ID, SCAR_Bullet_Damage, i
    
    'recoil
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * SCAR_SingleRecoil_Angle
    
ElseIf Stick(i).WeaponType = SA80 Then
    
    'TempDir = BaseTempDir + Rnd() * SA80_Spray_Angle - 0.1 * Cos(Stick(i).Facing)
    BaseTempDir = Stick(i).ActualFacing + 0.01 * IIf(Stick(i).Facing > pi, -1, 1) + PM_Rnd * Stick(i).Speed / AccuracyRedux
    
    If StickiHasState(i, stick_crouch) Then
        TempDir = BaseTempDir - Sin(Stick(i).Facing) / 25
    Else
        TempDir = BaseTempDir - Sin(Stick(i).Facing) / 50
    End If
    
    
    AddBullet tX, tY, TempMag, TempDir, Stick(i).ID, SA80_Bullet_Damage, i
    
    'recoil
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * SA80_SingleRecoil_Angle
    
ElseIf Stick(i).WeaponType = RPG Then
    
    AddNade tX, tY, Stick(i).ActualFacing, RPG_Speed, i, Stick(i).Colour, nFrag, True
    
    AddExplosion tX, tY, 500, 0.5, Stick(i).Speed, Stick(i).Heading
    
    
    'rear point
    tX = tX - GunLen * 4 * Sin(Stick(i).Facing)
    tY = tY + GunLen * 4 * Cos(Stick(i).Facing)
    AddExplosion tX, tY, 500, 0.5, Stick(i).Speed, Stick(i).Heading
    
    'recoil
    Stick(i).Facing = Stick(i).ActualFacing - IIf(Stick(i).RecoilLeft, -1, 1) * RPG_SingleRecoil_Angle
    
    For j = 1 To 10
        AddSmokeGroup tX, tY, 4, 20 + 30 * Rnd(), Stick(i).ActualFacing - pi
    Next j
    '+ GunLen * Sin(Stick(i).Facing - pi)
    
    'recoil
    AddVectors Stick(i).Speed, Stick(i).Heading, RPG_RecoilForce, FixAngle(Stick(i).Facing - pi), Stick(i).Speed, Stick(i).Heading
    Stick_Moving = True
    bLBound = False
    
    
End If



End Sub

Private Sub AlterShotgunBulletDamage(i As Integer)
Dim tAdjust As Single
'time left = B(i).Decay - GTC()

tAdjust = modStickGame.StickTimeFactor * (Bullet(i).Decay - GetTickCount()) / Bullet_Decay

'tadjust^3
Bullet(i).Damage = Bullet(i).Damage * tAdjust * tAdjust * tAdjust

End Sub

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
    StickMotion Blood(i).X, Blood(i).Y, Blood(i).Speed, Blood(i).Heading
    i = i + 1
Loop

End Sub

Private Sub DoRecoil(i As Integer, ByRef Stick_Moving As Boolean, ByRef bLBound As Boolean)

If Stick(i).LastBullet + kRecoilTime(Stick(i).WeaponType) / modStickGame.sv_StickGameSpeed > GetTickCount() Then
    
    Stick(i).Facing = Stick(i).Facing + _
        IIf(Stick(i).RecoilLeft, -1, 1) * kRecoverAmount(Stick(i).WeaponType) * modStickGame.sv_StickGameSpeed
    
    
    If kRecoilForce(Stick(i).WeaponType) Then
        
        If Stick(i).WeaponType <> RPG Then
            Stick_Moving = True
            bLBound = False
        ElseIf Stick(i).LastBullet + RPG_Recoil_Time / (3 * modStickGame.sv_StickGameSpeed) > GetTickCount() Then
            Stick_Moving = True
            bLBound = False
        End If
        
    End If
    
End If

'If Stick(i).WeaponType = Shotgun Then
'    If Stick(i).LastBullet + Shotgun_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * Shotgun_Recover_Amount * modStickGame.sv_StickGameSpeed
'
'        'allow recoil to push back
'        Stick_Moving = True
'        bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = AK Then
'    If Stick(i).LastBullet + AK_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * AK_Recover_Amount * modStickGame.sv_StickGameSpeed
'    End If
'ElseIf Stick(i).WeaponType = DEagle Then
'    If Stick(i).LastBullet + DEagle_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * DEagle_Recover_Amount * modStickGame.sv_StickGameSpeed
'
'        'allow recoil to push back
'        Stick_Moving = True
'        bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = M82 Then
'    If Stick(i).LastBullet + M82_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'
'
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * M82_Recover_Amount * modStickGame.sv_StickGameSpeed
'
'        'allow recoil to push back
'        Stick_Moving = True
'        bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = SCAR Then
'    If Stick(i).LastBullet + SCAR_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * SCAR_Recover_Amount * modStickGame.sv_StickGameSpeed
'    End If
'ElseIf Stick(i).WeaponType = M249 Then
'    If Stick(i).LastBullet + M249_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'        Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * M249_Recover_Amount * modStickGame.sv_StickGameSpeed
'
'        ''allow recoil to push back
'        'Stick_Moving = True
'        'bLBound = False
'
'    End If
'ElseIf Stick(i).WeaponType = RPG Then
'    If Stick(i).LastBullet + RPG_Recoil_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
'
'        If i = 0 Then 'prevent wobble
'            Stick(i).Facing = Stick(i).Facing + IIf(Stick(i).RecoilLeft, -1, 1) * RPG_Recover_Amount * modStickGame.sv_StickGameSpeed
'        End If
'
'
'        If Stick(i).LastBullet + RPG_Recoil_Time / (3 * modStickGame.sv_StickGameSpeed) > GetTickCount() Then
'            'allow recoil to push back
'            Stick_Moving = True
'            bLBound = False
'        End If
'
'    End If
'End If
End Sub

Private Sub MakeRecoilEtcAmountArray()
Dim i As Integer

For i = 0 To eWeaponTypes.Knife
    If i = Shotgun Then
        kRecoilTime(i) = Shotgun_Recoil_Time
        kRecoverAmount(i) = Shotgun_Recover_Amount
        kRecoilForce(i) = True
        
    ElseIf i = AK Then
        kRecoilTime(i) = AK_Recoil_Time
        kRecoverAmount(i) = AK_Recover_Amount
        
    ElseIf i = DEagle Then
        kRecoilTime(i) = DEagle_Recoil_Time
        kRecoverAmount(i) = DEagle_Recover_Amount
        kRecoilForce(i) = True
        
    ElseIf i = M82 Then
        kRecoilTime(i) = M82_Recoil_Time
        kRecoverAmount(i) = M82_Recover_Amount
        kRecoilForce(i) = True
        
    ElseIf i = SCAR Then
        kRecoilTime(i) = SCAR_Recoil_Time
        kRecoverAmount(i) = SCAR_Recover_Amount
        
    ElseIf i = M249 Then
        kRecoilTime(i) = M249_Recoil_Time
        kRecoverAmount(i) = M249_Recover_Amount
        
    ElseIf i = RPG Then
        kRecoilTime(i) = RPG_Recoil_Time
        kRecoverAmount(i) = RPG_Recover_Amount
        kRecoilForce(i) = True
        
    ElseIf i = SA80 Then
        kRecoilTime(i) = SA80_Recoil_Time
        kRecoverAmount(i) = SA80_Recover_Amount
        
    End If
Next i

End Sub

Private Function StickInvul(i As Integer) As Boolean
StickInvul = (Stick(i).LastSpawnTime + Spawn_Invul_Time / modStickGame.sv_StickGameSpeed > GetTickCount())
End Function

Private Sub DamageStick(ByVal DamageToDo As Integer, iStick As Integer, Optional bDamageArmour As Boolean = True)

'##################################################
'adjust depending on settings
If modStickGame.sv_Hardcore Then
    DamageToDo = DamageToDo * Hardcore_Damage_Amp
End If
If modStickGame.sv_GameType = gCoOp Then
    If Stick(iStick).IsBot = False Then
        DamageToDo = DamageToDo \ 2
    End If
End If
'##################################################



'##################################################
'perks/weapons, etc
If Stick(iStick).WeaponType = Chopper Then
    DamageToDo = DamageToDo \ Chopper_Damage_Reduction
    bDamageArmour = False
End If
If iStick = 0 Then
    'me, check for low damage perk
    If Stick(iStick).Perk = pJuggernaut Then
        DamageToDo = DamageToDo \ JuggernautDamageReduction
    End If
End If
'##################################################



'##################################################
'apply
If Stick(iStick).Armour > 0 And bDamageArmour Then
    
    If DamageToDo <= 1 Then
        DamageToDo = 2
    End If
    
    Stick(iStick).Armour = Stick(iStick).Armour - DamageToDo \ 2
    
    If Stick(iStick).Armour < 0 Then
        Stick(iStick).Health = Stick(iStick).Health + Stick(iStick).Armour
        Stick(iStick).Armour = 0
    End If
Else
    If DamageToDo = 0 Then
        DamageToDo = 1
    End If
    
    Stick(iStick).Health = Stick(iStick).Health - DamageToDo
End If
'##################################################

End Sub

Private Function GetBulletDelay(i As Integer) As Long

GetBulletDelay = kBulletDelay(Stick(i).WeaponType)

'If Stick(i).WeaponType = Shotgun Then
'    GetBulletDelay = Shotgun_Bullet_Delay
'ElseIf Stick(i).WeaponType = AK Then
'    GetBulletDelay = AK_Bullet_Delay
'ElseIf Stick(i).WeaponType = M82 Then
'    GetBulletDelay = M82_Bullet_Delay
'ElseIf Stick(i).WeaponType = SCAR Then
'    GetBulletDelay = SCAR_Bullet_Delay
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

Private Sub MakeBulletDelayArray()
Dim i As Integer

'kbulletdelay = array(

For i = 0 To eWeaponTypes.Chopper
    
    If i = Shotgun Then
        kBulletDelay(i) = Shotgun_Bullet_Delay
    ElseIf i = AK Then
        kBulletDelay(i) = AK_Bullet_Delay
    ElseIf i = M82 Then
        kBulletDelay(i) = M82_Bullet_Delay
    ElseIf i = SCAR Then
        kBulletDelay(i) = SCAR_Bullet_Delay
    ElseIf i = DEagle Then
        kBulletDelay(i) = DEagle_Bullet_Delay
    ElseIf i = M249 Then
        kBulletDelay(i) = M249_Bullet_Delay
    ElseIf i = RPG Then
        kBulletDelay(i) = RPG_Bullet_Delay  'needed to prevent spam
    ElseIf i = Chopper Then
        kBulletDelay(i) = Chopper_Bullet_Delay
    ElseIf i = FlameThrower Then
        kBulletDelay(i) = Flame_Bullet_Delay
    ElseIf i = SA80 Then
        kBulletDelay(i) = SA80_Single_Bullet_Delay 'SA80_Bullet_Delay
    Else
        kBulletDelay(i) = Knife_Delay
    End If
    
Next i

End Sub

Private Sub DoReload(Optional iStick As Integer)
Dim nBullets As Integer

With Stick(iStick)
    If .WeaponType = Knife Or .WeaponType = Chopper Then '.WeaponType = Shotgun
        If .BulletsFired > 0 Then
            .BulletsFired = 0
        End If
        Exit Sub
    End If
    
    nBullets = GetMaxRounds(Stick(iStick).WeaponType)
    
    If StickiHasState(iStick, Stick_Reload) Then
        
        If .ReloadStart + GetReloadTime(iStick) < GetTickCount() Then
            SubStickiState iStick, Stick_Reload
            
            .BulletsFired = 0
            
            If .WeaponType = RPG Then
                If StickiHasState(iStick, Stick_Fire) Then
                    SubStickiState iStick, Stick_Fire
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
                    If StickiHasState(0, Stick_Reload) Then
                        SubStickiState 0, Stick_Reload
                    ElseIf StickiHasState(0, Stick_Fire) Then
                        SubStickiState 0, Stick_Fire
                    End If
                    
                    Exit Sub
                End If
            End If
            
        End If
        
        
        
        If .LastBullet + AutoReload_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            Call StartReload(iStick)
        End If
        
    'ElseIf StickHasState(.ID, Stick_Reload) Then
        'SubStickState .ID, Stick_Reload
    End If
    
End With

End Sub

Private Sub CheckStickHealthPack(iStick As Integer)
Const HealthPack_RadiusXX = HealthPack_Radius * 4


If HealthPack.bActive Then
    If Stick(iStick).WeaponType <> Chopper Then
        If GetDist(Stick(iStick).X, Stick(iStick).Y, HealthPack.X, HealthPack.Y) < HealthPack_RadiusXX Then
            Stick(iStick).Health = Max_Health
            
            If Stick(iStick).Armour < 25 Then
                Stick(iStick).Armour = 25
            End If
            
            AddCirc HealthPack.X, HealthPack.Y, 500, 2, vbGreen
            HealthPack.bActive = False
            HealthPack.LastUsed = GetTickCount()
        End If
    End If
End If


End Sub
Private Sub StartReload(iStick As Integer)
With Stick(iStick)
    
    AddStickiState iStick, Stick_Reload
    SubStickiState iStick, Stick_Fire
    
    If iStick = 0 Then
        FireKey = False
    End If
    
    '.BulletsFired = 0'done at end
    .BulletsFired2 = 0
    .ReloadStart = GetTickCount()
    
End With
End Sub

Private Sub AddMagForStick(i As Integer)
Dim vMag As eMagTypes

If Stick(i).WeaponType = AK Then
    vMag = mAK
ElseIf Stick(i).WeaponType = SCAR Then
    vMag = mSCAR
ElseIf Stick(i).WeaponType = M82 Or Stick(i).WeaponType = M249 Then
    vMag = mSniper
ElseIf Stick(i).WeaponType = DEagle Then
    vMag = mPistol
ElseIf Stick(i).WeaponType = FlameThrower Then
    vMag = mFlameThrower
ElseIf Stick(i).WeaponType = SA80 Then
    vMag = mSA80
Else
    'vMag = -1
    Exit Sub
End If

AddMag CSng(Stick(i).CasingPoint.X), CSng(Stick(i).CasingPoint.Y), Stick(i).Speed, Stick(i).Heading, vMag

End Sub

Private Function GetMaxRounds(vWeapon As eWeaponTypes) As Integer 'Sticki As Integer) As Integer

GetMaxRounds = kMaxRounds(vWeapon)

'If vWeapon = AK Then
'    GetMaxRounds = AK_Bullets
'ElseIf vWeapon = M82 Then
'    GetMaxRounds = M82_Bullets
'ElseIf vWeapon = Shotgun Then
'    GetMaxRounds = Shotgun_Bullets * Shotgun_Gauge
'ElseIf vWeapon = SCAR Then
'    GetMaxRounds = SCAR_Bullets
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
    ElseIf i = Shotgun Then
        kMaxRounds(i) = Shotgun_Bullets * Shotgun_Gauge
    ElseIf i = SCAR Then
        kMaxRounds(i) = SCAR_Bullets
    ElseIf i = DEagle Then
        kMaxRounds(i) = DEagle_Bullets
    ElseIf i = M249 Then
        kMaxRounds(i) = M249_Bullets
    ElseIf i = RPG Then
        kMaxRounds(i) = RPG_Bullets
    ElseIf i = FlameThrower Then
        kMaxRounds(i) = Flame_Bullets
    ElseIf i = SA80 Then
        kMaxRounds(i) = SA80_Bullets
    Else
        kMaxRounds(i) = 1
    End If
Next i

End Sub

Private Function CoOrdInStick(X As Single, Y As Single, Sticki As Integer) As Boolean

Const AL1p5 = ArmLen * 1.5
Const K = BodyLen + LegHeight * 2
Const CLDx = ChopperLen / 1.2
Dim sY As Single

If Stick(Sticki).WeaponType = Chopper Then
    
    If X > (Stick(Sticki).X - CLDx) Then
        If X < (Stick(Sticki).X + CLD4) Then
            If Y > Stick(Sticki).Y - CLD8 Then
                If Y < Stick(Sticki).Y + CLD6 Then
                    CoOrdInStick = True
                End If
            End If
            
        End If
    End If
    
Else
    
    If X > CLng(Stick(Sticki).X - AL1p5) Then
        If X < CLng(Stick(Sticki).X + AL1p5) Then
            
            sY = GetStickY(Sticki)
            
            If Y > sY Then '(Stick(Sticki).y) Then
                If Y < sY + K Then '(Stick(Sticki).y + BodyLen + LegHeight * 2) Then
                    CoOrdInStick = True
                End If
            End If
            
        End If
    End If
End If

End Function

Private Function BulletInStick(Sticki As Integer, Bulleti As Integer) As Boolean
Dim sY As Single
Dim kArmLen As Single
Const ArmLenExtended = ArmLen * 1.2
Const HeadRadiusX2 = HeadRadius * 2, BodyLenX2 = BodyLen * 2


If Stick(Sticki).WeaponType = Chopper Then
    
    BulletInStick = CoOrdInChopper(Bullet(Bulleti).X, Bullet(Bulleti).Y, Sticki)
    
'    If Bullet(Bulleti).X > (Stick(Sticki).X - CLDx) Then
'        If Bullet(Bulleti).X < (Stick(Sticki).X + CLD4) Then
'            If Bullet(Bulleti).Y > Stick(Sticki).Y - CLD8 Then
'                If Bullet(Bulleti).Y < Stick(Sticki).Y + CLD6 Then
'                    BulletInStick = True
'                End If
'            End If
'
'        End If
'    End If
    
Else
    If StickiHasState(Sticki, Stick_Prone) Then
        kArmLen = ArmLen * 3 'body/leg hit
    Else
        kArmLen = ArmLenExtended
    End If
    
    If Bullet(Bulleti).X > (Stick(Sticki).X - kArmLen) Then
        If Bullet(Bulleti).X < (Stick(Sticki).X + kArmLen) Then
            
            sY = GetStickY(Sticki)
            
            If Bullet(Bulleti).Y > sY Then '(Stick(Sticki).y) Then
                If Bullet(Bulleti).Y < sY + IIf(StickiHasState(Sticki, Stick_Prone), HeadRadiusX2, BodyLenX2) Then '(Stick(Sticki).y + BodyLen * 2) Then
                
                    BulletInStick = True
                    
                End If
            End If
            
        End If
    End If
End If

End Function

Private Function CoOrdInChopper(X As Single, Y As Single, iChopper As Integer) As Boolean
Const CLDx = ChopperLen / 1.2

If X > (Stick(iChopper).X - CLDx) Then
    If X < (Stick(iChopper).X + CLD4) Then
        If Y > Stick(iChopper).Y - CLD8 Then
            If Y < Stick(iChopper).Y + CLD6 Then
                CoOrdInChopper = True
            End If
        End If
        
    End If
End If
End Function

Private Function GetStickY(i As Integer) As Single
Const BodyLenX1p3 = BodyLen * 1.3
Const BodyLenD2 = BodyLen / 2

If StickiHasState(i, Stick_Prone) Then
    GetStickY = Stick(i).Y + BodyLenX1p3
ElseIf StickiHasState(i, stick_crouch) Then
    GetStickY = Stick(i).Y + BodyLenD2
Else
    GetStickY = Stick(i).Y
End If

End Function

Private Function BulletInHead(Sticki As Integer, Bulleti As Integer) As Boolean

Const HeadRadiusX2 = HeadRadius * 2, HeadRadiusXk = HeadRadius * 1.5
Dim sY As Single

If Stick(Sticki).WeaponType <> Chopper Then
    
    If Bullet(Bulleti).X > (Stick(Sticki).X - HeadRadiusX2) Then
        If Bullet(Bulleti).X < (Stick(Sticki).X + HeadRadiusX2) Then
            
            sY = GetStickY(Sticki)
            
            'yes, supposed to by -10 below
            If Bullet(Bulleti).Y > sY - 30 Then '(Stick(Sticki).y - 10) Then
                If Bullet(Bulleti).Y < sY + HeadRadiusXk Then '(Stick(Sticki).y + HeadRadiusX2) Then
                    BulletInHead = True
                End If
            End If
            
        End If
    End If
End If

End Function

Private Function NadeInBullet(Nadei As Integer) As Boolean
Dim i As Integer

For i = 0 To NumBullets - 1
    If Bullet(i).bSilenced = False Then
        If BulletNearNade(Nadei, i) Then
            NadeInBullet = True
            Exit For
        End If
    End If
Next i

End Function

Private Function BulletNearNade(Nadei As Integer, Bulleti As Integer) As Boolean
Const NadeLim = 150

If Bullet(Bulleti).LastDiffract = 0 Or Bullet(Bulleti).bSniperBullet Then
    If Bullet(Bulleti).X > (Nade(Nadei).X - NadeLim) Then
        If Bullet(Bulleti).X < (Nade(Nadei).X + NadeLim) Then
            
            If Bullet(Bulleti).Y > (Nade(Nadei).Y - NadeLim) Then
                If Bullet(Bulleti).Y < (Nade(Nadei).Y + NadeLim) Then
                    BulletNearNade = True
                End If
            End If
            
        End If
    End If
End If

End Function

Private Sub ApplyGravity(i As Integer, Optional bResetSpeed As Boolean = True)
Dim j As Integer
Dim XComp As Single, YComp As Single
Dim StickOnPlatform As Boolean

'fly mode
'If Not Stick(i).OnSurface Then Stick(i).OnSurface = True
'Exit Sub

If StickiHasState(i, stick_Jump) = False Then
    'If Stick(i).StartJumpTime + JumpTime / 2 < GetTickCount() Then
        
        For j = 0 To nPlatforms
            
            If StickOnSurface(i, j) Then
                'Stick(i).Y = Me.height - StickHeight - Lim
                'Stick(i).Speed = Stick(i).Speed / 1.05
                'XComp = Stick(i).Speed * Sin(Stick(i).Heading)
                'Stick(i).Speed = XComp
                
                If bResetSpeed Then
                    With Stick(i)
                        XComp = .Speed * Sin(.Heading)
                        YComp = .Speed * Cos(.Heading)
                        
                        If YComp < -0.01 Then
                            .Speed = 0 'XComp
                        End If
                    End With
                End If
                
                Stick(i).Y = Platform(j).Top - BodyLen - LegHeight + 20
                
                StickOnPlatform = True
                
                Exit For
                
            End If
            
        Next j
        
    'End If
End If

If Not StickOnPlatform Then
    If Stick(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        
        AddVectors Stick(i).Speed, Stick(i).Heading, _
            Gravity_Strength, _
            Gravity_Direction, Stick(i).Speed, Stick(i).Heading
        
        
        Stick(i).LastGravity = GetTickCount()
    End If
    
    If StickiHasState(i, stick_crouch) Then
        SubStickiState i, stick_crouch
    ElseIf StickiHasState(i, Stick_Prone) Then
        SubStickiState i, Stick_Prone
    ElseIf StickiHasState(i, stick_Left) Then
        SubStickiState i, stick_Left
    ElseIf StickiHasState(i, stick_Right) Then
        SubStickiState i, stick_Right
    End If
    
    
End If

Stick(i).OnSurface = StickOnPlatform

End Sub

Private Function StickOnSurface(Sticki As Integer, iPlatform As Integer) As Boolean

Dim StickFootY As Integer ', YComp As Single
Dim Diff As Integer

If Stick(Sticki).X > Platform(iPlatform).Left Then
    
    
    If Stick(Sticki).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
        'PrintStickText "X Col", Platform(iPlatform).Left + 200, Platform(iPlatform).Top - 500, 0
        
        StickFootY = Stick(Sticki).Y + BodyLen + LegHeight
        
        If StickFootY > Platform(iPlatform).Top Then
            If StickFootY < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
                
                
                'position the stick on top of the platform
                If StickFootY > (Platform(iPlatform).Top + 5) Then
                    '                       add on a bit so he's not bouncing up + down V
                    Stick(Sticki).Y = Platform(iPlatform).Top - BodyLen - LegHeight + 20
                    
'                    'prevent bouncing
'                    If StickiHasState(Sticki, Stick_Jump) Then
'                        SubStickState Stick(Sticki).ID, Stick_Jump
'                    End If
                    
                End If
                
                StickOnSurface = True
                
            End If
        End If
        
'        If StickFootY < (Platform(iPlatform).Top) Then
'
'            YComp = Stick(Sticki).Speed * Cos(Stick(Sticki).Heading)
'
'            If StickFootY + 50 > (Platform(iPlatform).Top - 50) Then
'                PrintStickText "Y Col", Platform(iPlatform).Left + 200, Platform(iPlatform).Top - 1000, 0
'                StickOnSurface = True
'            End If
'        End If
        
    End If
End If

End Function

Private Sub Killed(ByVal DeadStickj As Integer, ByVal iKiller As Integer, ByVal KillType As eKillTypes)
Dim ChatText As String, FullText As String
Dim i As Integer
Dim bDeadStickExists As Boolean, bKillerExists As Boolean

On Error Resume Next
bDeadStickExists = Stick(DeadStickj).bTyping Or True
bKillerExists = Stick(iKiller).bTyping Or True

On Error GoTo EH
If DeadStickj <> -1 And bDeadStickExists Then
    
    If Stick(DeadStickj).WeaponType = Chopper Then
        AddDeadChopper Stick(DeadStickj).X, Stick(DeadStickj).Y, Stick(DeadStickj).Colour, DeadStickj
    Else
        AddDeadStick Stick(DeadStickj).X, Stick(DeadStickj).Y, Stick(DeadStickj).Colour, _
            (Stick(DeadStickj).Facing < pi), (KillType = kBurn Or KillType = kFlame)
    End If
    
    Stick(DeadStickj).Speed = 0
    Stick(DeadStickj).Health = Health_Start
    Stick(DeadStickj).BulletsFired = 0
    If StickiHasState(DeadStickj, Stick_Reload) Then
        SubStickiState DeadStickj, Stick_Reload
    End If
    Stick(DeadStickj).X = (StickGameWidth - 1000) * Rnd()
    Stick(DeadStickj).Y = StickGameHeight * Rnd()
    Stick(DeadStickj).Facing = 2 * pi * Rnd()
    Stick(DeadStickj).LastSpawnTime = GetTickCount()
    'SubStickState Stick(DeadStickj).ID, stick_Left
    'SubStickState Stick(DeadStickj).ID, stick_Right
    Stick(DeadStickj).State = Stick_None
    Stick(DeadStickj).OnSurface = False
    
    Stick(DeadStickj).LastFlameTouch = 1
    Stick(DeadStickj).bOnFire = False
    Stick(DeadStickj).LastFlashBang = 1
    Stick(DeadStickj).bFlashed = False
    
    Stick(DeadStickj).iDeaths = Stick(DeadStickj).iDeaths + 1
    Stick(DeadStickj).iKillsInARow = 0
    Stick(DeadStickj).LastFlashBang = 1
    'Stick(DeadStickj).bLightSaber = False
    
    
    If DeadStickj = 0 Then '=FindStick(MyID) Then
        If modStickGame.sv_GameType <> gCoOp Then
            If modStickGame.sv_GameType <> gElimination Then
                AddCirc Stick(DeadStickj).X, Stick(DeadStickj).Y, 1000, 2, vbGreen
            End If
        End If
        
        'ResetAmmoFired
        FillTotalMags
        
        LeftKey = False
        RightKey = False
        JumpKey = False
        CrouchKey = False
        ProneKey = False
        UseKey = False
        
        ChopperAvail = False
        FlamesInARow = 0
        KnifesInARow = 0
        
    End If
    
    
    If modStickGame.sv_GameType = gElimination Or modStickGame.sv_GameType = gCoOp Then
        Stick(DeadStickj).bAlive = False
    End If
    
    
    If iKiller <> -1 And bKillerExists Then
        
        If DeadStickj = iKiller Then
            ChatText = Trim$(Stick(DeadStickj).Name) & " committed suicide"
        ElseIf KillType = kNormal Then
            ChatText = "killed by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kNade Then
            ChatText = "grenaded by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kSilenced Then
            ChatText = "silenced by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kRPG Then
            ChatText = "rocketed by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kHead Then
            ChatText = "headshotted by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kMine Then
            ChatText = "mined by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kFlame Then
            ChatText = "fried by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kBurn Then
            ChatText = "toasted by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kChoppered Then
            ChatText = "diced by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kKnife Then
            ChatText = "knifed by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kCrushed Then
            ChatText = "crushed by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kFlameTag Then
            ChatText = "flame-tagged by " & Trim$(Stick(iKiller).Name)
        ElseIf KillType = kLightSaber Then
            ChatText = "Lightsaber'd by " & Trim$(Stick(iKiller).Name)
        End If
        
        If DeadStickj <> iKiller Then
            FullText = Trim$(Stick(DeadStickj).Name) & " was " & ChatText
            Stick(iKiller).iKills = Stick(iKiller).iKills + 1 'INCREASE HERE
            Stick(iKiller).iKillsInARow = Stick(iKiller).iKillsInARow + 1
        Else
            FullText = ChatText
        End If
        
        
        If iKiller = 0 Then
            If iKiller = DeadStickj Then
                
                If KillType = kMine Then
                    AddMainMessage "Suicide Mine!"
                Else
                    AddMainMessage "You Suck!"
                End If
                
            Else
                Call CheckKillsInARow
            End If
        End If
        
        
        
        If StickServer Then
            SendChatPacketBroadcast FullText, Stick(iKiller).Colour
        Else
            modWinsock.SendPacket socket, ServerSockAddr, sChats & FullText & "#" & CStr(Stick(iKiller).Colour)
            
            'AddChatText ChatText, Stick(iKiller).Colour
            'we'll get it back
        End If
        
        
        If Stick(DeadStickj).ID = MyID Then 'if we're dead, tell the server to add one to the killer's kills
            
            
            'picToasty.Visible = False
            
            For i = 0 To eWeaponTypes.Knife
                AmmoFired(i) = 0
            Next i
            
            If Stick(DeadStickj).WeaponType = Chopper Then
                'it's me
                SwitchWeapon Stick(0).CurrentWeapons(1)
            End If
            
            
            If DeadStickj <> iKiller Then
                AddMainMessage UCase$(Left$(ChatText, 1)) & Mid$(ChatText, 2), Stick(iKiller).Colour
                
                If modStickGame.StickServer = False Then
                    modWinsock.SendPacket socket, ServerSockAddr, sDeathInfos & CStr(Stick(DeadStickj).ID) & _
                        CStr(Abs(KillType = kFlame Or KillType = kBurn))
                    
                    modWinsock.SendPacket socket, ServerSockAddr, sKillInfos & CStr(Stick(iKiller).ID)
                Else
                    SendBroadcast sDeathInfos & CStr(Stick(DeadStickj).ID) & _
                        CStr(Abs(KillType = kFlame Or KillType = kBurn))
                    
                    SendBroadcast sKillInfos & CStr(Stick(iKiller).ID)
                End If
                
            End If
            
        ElseIf Stick(DeadStickj).IsBot Then
            
            If iKiller = 0 Then
                If Stick(iKiller).WeaponType = FlameThrower Then
                    If KillType = kBurn Or KillType = kFlame Then
                        FlamesInARow = FlamesInARow + 1
                    End If
                End If
            Else
                If modStickGame.StickServer = False Then
                    modWinsock.SendPacket socket, ServerSockAddr, sDeathInfos & CStr(Stick(DeadStickj).ID) & _
                        CStr(Abs(KillType = kFlame Or KillType = kBurn))
                    
                    modWinsock.SendPacket socket, ServerSockAddr, sKillInfos & CStr(Stick(iKiller).ID)
                Else
                    SendBroadcast sDeathInfos & CStr(Stick(DeadStickj).ID) & _
                        CStr(Abs(KillType = kFlame Or KillType = kBurn))
                    
                    SendBroadcast sKillInfos & CStr(Stick(iKiller).ID)
                End If
            End If
            
        End If
    End If
End If

EH:
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

Private Sub InitVariables()

Dim Ctrl As Control

For Each Ctrl In Controls
    If Not (TypeOf Ctrl Is Timer) Then
        If Not (TypeOf Ctrl Is Shape) Then
            Ctrl.TabStop = False
        End If
    End If
Next Ctrl


'Add us as a Stick!
AddStick
Stick(0).CurrentWeapons(1) = AK
Stick(0).CurrentWeapons(2) = DEagle
'If we're the host, assign our ID now
If StickServer Then
    Stick(0).ID = 0
    MyID = 0
End If
RandomizeMyStickPos
Stick(0).Facing = 2 * pi * Rnd()
Stick(0).Health = Health_Start
Stick(0).Name = frmMain.LastName
Stick(0).Colour = modVars.TxtForeGround

ChatFlashDelay = GetCursorBlinkTime()

InitPlatforms

MakeBulletDelayArray
MakeMaxRoundsArray
MakeReloadTimeArray
MakePerkNameArray
MakeWeaponNameArray
MakeTeamColourArray
MakeGameTypeArray
MakeRecoilEtcAmountArray
MakeNadeNameArray


If modStickGame.sv_2Weapons Then
    frmStickGame.MakeStaticWeapons
    frmStickGame.SetCurrentWeapons
End If

FillTotalMags

End Sub

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

Private Sub InitPlatforms()
Dim i As Integer

'Platform
Platform(0).Left = -1000: Platform(0).Top = 13572: Platform(0).width = 52000: Platform(0).height = 855
Platform(1).Left = 0: Platform(1).Top = 11400: Platform(1).width = 7575: Platform(1).height = 375
Platform(2).Left = 6240: Platform(2).Top = 8400: Platform(2).width = 25000: Platform(2).height = 375
Platform(3).Left = 840: Platform(3).Top = 6000: Platform(3).width = 4935: Platform(3).height = 375
Platform(4).Left = 13000: Platform(4).Top = 5500: Platform(4).width = 10000: Platform(4).height = 375
Platform(5).Left = 12120: Platform(5).Top = 11400: Platform(5).width = 35000: Platform(5).height = 375
Platform(6).Left = 44500: Platform(6).Top = 5000: Platform(6).width = 5500: Platform(6).height = 375
Platform(7).Left = 42000: Platform(7).Top = 8300: Platform(7).width = 500: Platform(7).height = 375

'tBox
tBox(0).Left = 7200: tBox(0).Top = 10905: tBox(0).width = 375: tBox(0).height = 495
tBox(1).Left = 37000: tBox(1).Top = Platform(5).Top - 495: tBox(1).width = 1215: tBox(1).height = 495
tBox(2).Left = 5400: tBox(2).Top = Platform(3).Top - Platform(3).height: tBox(2).width = 375: tBox(2).height = 495
tBox(3).Left = 9600: tBox(3).Top = 8005: tBox(3).width = 495: tBox(3).height = 495
tBox(4).Left = Platform(4).Left + Platform(4).width - 375: tBox(4).Top = Platform(4).Top - Platform(4).height: tBox(4).width = 375: tBox(4).height = 495
tBox(5).Left = Platform(6).Left: tBox(5).height = 900: tBox(5).Top = Platform(6).Top - tBox(5).height: tBox(5).width = 500
tBox(6).Left = 25000: tBox(6).Top = 8025: tBox(6).width = 1215: tBox(6).height = 375
tBox(7).Left = 0: tBox(7).Top = 11025: tBox(7).width = 1215: tBox(7).height = 375
tBox(8).Left = 15000: tBox(8).Top = Platform(4).Top - Platform(4).height: tBox(8).width = 1200: tBox(8).height = 375

'Box
Box(0).Left = 5587: Box(0).Top = tBox(2).Top - 1095: Box(0).width = 135: Box(0).height = 1095
Box(1).Left = 6290: Box(1).Top = 7305: Box(1).width = 135: Box(1).height = 1095
Box(2).Left = tBox(1).Left + 67: Box(2).Top = tBox(1).Top - 1095: Box(2).width = 135: Box(2).height = 1095
Box(3).Left = tBox(4).Left + 67: Box(3).Top = tBox(4).Top - 1095: Box(3).width = 135: Box(3).height = 1095
Box(4).Left = 3840: Box(4).Top = 10405: Box(4).width = 135: Box(4).height = 1095
Box(5).Left = 7200: Box(5).Top = 8775: Box(5).width = 375: Box(5).height = 2505
Box(6).Left = 28000: Box(6).Top = Platform(2).Top + Platform(2).height - 100: Box(6).width = 135: Box(6).height = Platform(5).Top - Box(6).Top
Box(7).Left = 7440: Box(7).Top = 11775: Box(7).width = 135: Box(7).height = 1897
Box(8).Left = Platform(7).Left: Box(8).Top = Platform(7).Top + Platform(7).height: Box(8).width = 495: Box(8).height = Platform(5).Top - Box(8).Top
Box(9).Left = 15240: Box(9).Top = 8775: Box(9).width = 375: Box(9).height = 2725
Box(10).Left = 13920: Box(10).Top = 5875: Box(10).width = 375: Box(10).height = 2625
'Box(11).Left = tBox(5).Left + 67: Box(11).Top = tBox(5).Top - 1095: Box(11).width = 135: Box(11).height = 1095
For i = 0 To nBoxes
    Box(i).bInUse = True
Next i

End Sub

Public Function RandomRGBColour() As Long

RandomRGBColour = RGB( _
        Int(Rnd() * 256), _
        Int(Rnd() * 256), _
        Int(Rnd() * 256))

End Function

Public Function AddBot(vWeapon As eWeaponTypes, vTeam As eTeams, Col As Long) As Integer
Dim i As Integer

i = AddStick(True)

Stick(i).Colour = Col
Stick(i).Team = vTeam
Stick(i).WeaponType = vWeapon

AddBot = i

End Function

Private Function GenerateBotName() As String
Dim i As Integer, j As Integer

For i = 0 To NumSticksM1
    If Stick(i).IsBot Then
        j = j + 1
    End If
Next i

GenerateBotName = "Bot " & CStr(j)

End Function

Public Function AddStick(Optional ByVal Bot As Boolean = False) As Integer

'Add a Stick onto the array, and return his index
ReDim Preserve Stick(NumSticks)
'ReDim Preserve ScoreList(NumSticks)

'ScoreList(NumSticks).ID = Stick(NumSticks).ID
Stick(NumSticks).LastPacket = GetTickCount()
Stick(NumSticks).LegWidth = 50
Stick(NumSticks).LastBullet = Stick(NumSticks).LastPacket

Stick(NumSticks).Name = vbNullString
Stick(NumSticks).bAlive = True
Stick(NumSticks).LastPacket = GetTickCount() + mPacket_SEND_DELAY * 4

If Bot Then
    Stick(NumSticks).IsBot = True
    
    Stick(NumSticks).Name = GenerateBotName() '"Bot " & NumSticks
    'Stick(NumSticks).Facing = piD2 'right
    'Stick(NumSticks).x = StickGameWidth / 2
    Stick(NumSticks).X = Rnd() * StickGameWidth
    
    Stick(NumSticks).Health = Health_Start
    Stick(NumSticks).Colour = modSpaceGame.RandomRGBColour()
    SetAINadeDelay NumSticks
    
    If NumSticks > 0 Then
        Stick(NumSticks).ID = Stick(NumSticksM1).ID + 1
    End If
    
    
    Stick(NumSticks).AICurrentTarget = -1
End If

AddStick = NumSticks
NumSticks = NumSticks + 1
NumSticksM1 = NumSticks - 1

End Function

Public Sub RemoveStick(Index As Integer)
Dim i As Integer


If Index > 0 Then 'not removing local stick
    If Stick(0).Perk = pSpy Then
        'if we're a spy...
        If Stick(0).MaskID = Stick(Index).ID Then
            'we were masquerading as the removed stick
            AddMainMessage "Target Stick has left the game (Spy)"
            Stick(0).MaskID = Stick(0).ID
        End If
    End If
End If

On Error Resume Next
'Remove this Stick from the array
For i = Index To NumSticks - 2
    Stick(i) = Stick(i + 1)
Next i

'Resize the array
ReDim Preserve Stick(NumSticks - 2)
NumSticks = NumSticksM1
NumSticksM1 = NumSticks - 1

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

If Stick(0).ActualFacing > pi Then 'MouseX < Me.width / 2 Then
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


Private Sub DisplaySticks()

Dim i As Integer
Dim Txt As String
Const Invul_Radius = BodyLen * 2

If StickInGame(0) And bPlaying Then
    If modStickGame.cg_AutoCamera Then
        ProcessCameraMovement
    Else
        MoveCameraX Stick(0).X * cg_sZoom - StickCentreX
        MoveCameraY Stick(0).Y * cg_sZoom - StickCentreY
    End If
'Else
    'MoveCameraX Stick(0).X * cg_sZoom - StickCentreX
    'MoveCameraY Stick(0).Y * cg_sZoom - StickCentreY
End If

picMain.FillStyle = vbFSTransparent 'transparent


For i = 0 To NumSticksM1
    
    If StickInGame(i) Then
        
        Me.picMain.DrawWidth = 2
        
        'on error resume next 'overflow
        DrawStick i
        
        If CanSeeStick(i) Then
            
            
            Me.picMain.DrawWidth = 3
            
            If Stick(i).WeaponType <> Chopper Then
                'If Stick(i).LastSpawnTime + Spawn_Invul_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
                If StickInvul(i) Then
                    modStickGame.sCircle Stick(i).X, Stick(i).Y + 250, Invul_Radius, Stick(i).Colour
                End If
            End If
            
            'PrintStickText "In Smoke: " & StickInSmoke(i), Stick(0).X, Stick(0).Y - 2000, vbBlue
            'PrintStickText "In tBox: " & StickIntBox(i), Stick(0).X, Stick(0).Y - 2000, vbBlue
            
        End If
    End If
    
Next i

End Sub

Private Function CanSeeStick(i As Integer) As Boolean
Const Peripheral_Vision = piD6 '[b]piD3[/b] / 2

Dim Theta As Single
Dim F As Single


If i = 0 Then
    CanSeeStick = True
ElseIf modStickGame.sv_Hardcore = False Then
    CanSeeStick = True
ElseIf StickInGame(0) = False Then
    CanSeeStick = True
ElseIf Stick(i).WeaponType = Chopper Then
    CanSeeStick = True
Else
    
    F = FixAngle(Stick(0).ActualFacing)
    Theta = FixAngle(FindAngle(Stick(0).X, Stick(0).Y, Stick(i).X, GetStickY(i) + 1))
    
    
    'stick(0).ActualFacing -abit < Theta < Stick(0).ActualFacing +abit
    If (F - Peripheral_Vision) < Theta Then
        If Theta < (F + Peripheral_Vision) Then
            CanSeeStick = True
        End If
    End If
    
End If


End Function

Private Function StickCanSeeStick(iSource As Integer, iTarget As Integer) As Boolean
Const Peripheral_Vision = piD6 '[b]piD3[/b] / 2

Dim Theta As Single
Dim F As Single


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
    
    F = FixAngle(Stick(iSource).ActualFacing)
    Theta = FixAngle(FindAngle_Actual(Stick(iSource).X, Stick(iSource).Y, Stick(iTarget).X, GetStickY(iTarget)))
    
    
    'stick(0).ActualFacing -abit < Theta < Stick(0).ActualFacing +abit
    If (F - Peripheral_Vision) < Theta Then
        If Theta < (F + Peripheral_Vision) Then
            StickCanSeeStick = True
        End If
    End If
    
End If


End Function

Private Sub DrawStick(i As Integer)

Dim Crouching As Boolean, Prone As Boolean, bArmoured As Boolean
Dim X As Single, Y As Single, tX As Single, tY As Single 'stick's co-ords
Dim XComp As Single 'for leg width
Const LegWidthK As Single = 8 'leg width speed
Const HRx1p8 = HeadRadius * 1.8, FlashEffect_Radius = HeadRadius / 1.5
Const ArmourTop = HeadRadius * 2 + 50
Const HeadRadiusX2 = HeadRadius * 2
Dim Hand1X As Single, Hand1Y As Single 'hand co-ords
Dim Hand2X As Single, Hand2Y As Single
Dim ShoulderY As Single
Dim GunY As Single
Dim TeamCol As Long
Dim j As Integer 'for spy mask

If Stick(i).Perk = pSpy Then
    j = FindStick(Stick(i).MaskID)
    If j = -1 Then j = 0
Else
    j = i
End If


If Stick(i).WeaponType = Chopper Then
    DrawChopper i
Else
    Crouching = StickiHasState(i, stick_crouch)
    Prone = StickiHasState(i, Stick_Prone)
    Stick(i).Facing = FixAngle(Stick(i).Facing)
    bArmoured = (Stick(i).Armour > 0)
    
    '###################################################
    'Find X and Y
    X = Stick(i).X
    If Crouching Then
        Y = Stick(i).Y + BodyLen / 2
        GunY = Y
    ElseIf Prone Then
        Y = Stick(i).Y + BodyLen * 1.2
        GunY = Y - 50
    Else
        Y = Stick(i).Y
        GunY = Y
    End If
    
    '###################################################
    
    'Draw Head + Body
    If CanSeeStick(i) Then
        picMain.DrawWidth = 2
        'Col = IIf(Stick(i).Armour > 0, Armour_Colour, Stick(i).Colour)
        
        'head
        If Stick(j).Team > Neutral Then
            TeamCol = GetTeamColour(Stick(j).Team)
            Me.picMain.FillStyle = vbSolid
            Me.picMain.FillColor = TeamCol
            modStickGame.sCircle X, Y + HeadRadius, HeadRadius, Stick(j).Colour 'head
            Me.picMain.FillStyle = vbFSTransparent
        Else
            modStickGame.sCircle X, Y + HeadRadius, HeadRadius, Stick(j).Colour 'head
        End If
        
        
        If bArmoured Then
            'modStickGame.sCircleSE X, Y + HeadRadius, HeadRadius, Armour_Colour, -(Stick(i).Facing - piD2), Stick(i).Facing - pi3D2
            modStickGame.sCircleSE X, Y + HeadRadius, HeadRadius, Armour_Colour, -0.01, -pi
        End If
        
        
        'draw legs
        DrawLegs i, X, Y, Crouching, Prone, Stick(j).Colour
        
        If Prone Then
            
            If bArmoured Then
                picMain.DrawWidth = 3
                modStickGame.sLine X + IIf(Stick(i).Facing > pi, HeadRadius, -HeadRadius), Y + HeadRadius, _
                               X + IIf(Stick(i).Facing > pi, BodyLen, -BodyLen), Y + HRx1p8, Armour_Colour
                picMain.DrawWidth = 2
            Else
                modStickGame.sLine X + IIf(Stick(i).Facing > pi, HeadRadius, -HeadRadius), Y + HeadRadius, _
                               X + IIf(Stick(i).Facing > pi, BodyLen, -BodyLen), Y + HRx1p8, Stick(j).Colour
            End If
            
        Else
            modStickGame.sLine X, Y + HeadRadiusX2, X, Y + BodyLen, Stick(j).Colour 'body
            
            If bArmoured Then
                modStickGame.sBoxFilled X - 10, Y + ArmourTop, X + 10, Y + BodyLen + HeadRadius, Armour_Colour
            End If
            
        End If
        
        If Stick(i).bTyping Then
            If StickiHasState(i, Stick_Prone) = False Then
                DrawTypeBubble Stick(i).X, Stick(i).Y - 500
            End If
        End If
    End If
    
    
    '###################################################
    
    
    
    
    'MUST DRAW THE WEAPON TO GET UPDATED GUNPOINT CO-ORDS
    If Stick(i).WeaponType = Shotgun Then
        DrawShotgun i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = AK Then
        DrawAK i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = M82 Then
        DrawM82 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY, Stick(j).Colour
    ElseIf Stick(i).WeaponType = DEagle Then
        DrawDEagle i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = SCAR Then
        DrawSCAR i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = SA80 Then
        DrawSA80 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = RPG Then
        DrawRPG i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = M249 Then
        DrawM249 i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    ElseIf Stick(i).WeaponType = FlameThrower Then
        DrawFlameThrower i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    Else
        DrawKnife i, Hand1X, Hand1Y, Hand2X, Hand2Y, X, GunY
    End If
    
    
    If CanSeeStick(i) Then
        picMain.DrawWidth = 2
        
        'picMain.DrawWidth = 2
        Me.ForeColor = Stick(j).Colour
        ShoulderY = Y + BodyLen / 2
        
        
        If Stick(i).WeaponType <> SA80 Then
            modStickGame.sLine X, ShoulderY, Hand1X, Hand1Y, Stick(j).Colour 'arm1
            modStickGame.sLine X, ShoulderY, Hand2X, Hand2Y, Stick(j).Colour 'arm2
        Else
            modStickGame.sLine X, ShoulderY, Hand1X, Hand1Y, Stick(j).Colour 'arm1
        End If
        
        
        If Stick(i).bFlashed Then
            modStickGame.sCircle X + PM_Rnd() * HeadRadius, Y + HeadRadius * (1 + PM_Rnd()), FlashEffect_Radius, vbYellow
        End If
    End If
    
    'move his legs
    If StickIsMoving(i) Then
        If Stick(i).LegWidth > MaxLegWidth Then
            Stick(i).LegBigger = False
        ElseIf Stick(i).LegWidth < -MaxLegWidth Then
            Stick(i).LegBigger = True
        End If
        
        XComp = Abs(Stick(i).Speed * Sin(Stick(i).Heading))
        
        If Stick(i).LegBigger Then
            Stick(i).LegWidth = Stick(i).LegWidth + XComp * modStickGame.sv_StickGameSpeed / LegWidthK
        Else
            Stick(i).LegWidth = Stick(i).LegWidth - XComp * modStickGame.sv_StickGameSpeed / LegWidthK
        End If
    End If
    
    
    
End If

End Sub

Private Sub DrawChopper(iStick As Integer)
Dim Pt(1 To 11) As POINTAPI, ScreenPt(1 To 3) As POINTAPI
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
Const pi2 = pi * 2
Const GunLen = 200


If StickiHasState(iStick, stick_Left) Then

'    If Stick(iStick).ChopperFacingAmount < piD3 Then
'        Stick(iStick).ChopperFacingAmount = Stick(iStick).ChopperFacingAmount + 0.05
'    Else
'        Stick(iStick).ChopperFacingAmount = piD3
'    End If
    
    Facing = pi5D12 'Stick(iStick).ChopperFacingAmount
    
    
ElseIf StickiHasState(iStick, stick_Right) Then
'    If Stick(iStick).ChopperFacingAmount < pi2d3 Then
'        Stick(iStick).ChopperFacingAmount = Stick(iStick).ChopperFacingAmount - 0.05
'    End If
    
    Facing = pi7D12 'Stick(iStick).ChopperFacingAmount
    
Else
    Facing = piD2
End If


Pt(1).X = Stick(iStick).X
Pt(1).Y = Stick(iStick).Y

Pt(2).X = Pt(1).X + CLD6 * Sin(Facing + piD6)
Pt(2).Y = Pt(1).Y - CLD6 * Cos(Facing + piD6)

Pt(3).X = Pt(2).X + CLD10 * Sin(Facing + piD3)
Pt(3).Y = Pt(2).Y - CLD10 * Cos(Facing + piD3)

Pt(4).X = Pt(3).X + CLD2 * Sin(Facing - pi)
Pt(4).Y = Pt(3).Y - CLD2 * Cos(Facing - pi)

Pt(5).X = Pt(4).X + CLD10 * Sin(Facing - pi3D4)
Pt(5).Y = Pt(4).Y - CLD10 * Cos(Facing - pi3D4)

Pt(6).X = Pt(5).X + CLD3 * Sin(Facing - pi)
Pt(6).Y = Pt(5).Y - CLD3 * Cos(Facing - pi)

Pt(7).X = Pt(6).X + CLD6 * Sin(Facing - pi3D4)
Pt(7).Y = Pt(6).Y - CLD6 * Cos(Facing - pi3D4)

Pt(8).X = Pt(7).X + CLD8 * Sin(Facing)
Pt(8).Y = Pt(7).Y - CLD8 * Cos(Facing)

Pt(9).X = Pt(8).X + CLD8 * Sin(Facing + piD4)
Pt(9).Y = Pt(8).Y - CLD8 * Cos(Facing + piD4)

Pt(10).X = Pt(9).X + CLD6 * Sin(Facing)
Pt(10).Y = Pt(9).Y - CLD6 * Cos(Facing)

Pt(11).X = Pt(1).X + CLD8 * Sin(Facing - pi)
Pt(11).Y = Pt(1).Y - CLD8 * Cos(Facing - pi)



ScreenPt(1).X = Pt(1).X + Sin(Facing + piD2) * 50
ScreenPt(1).Y = Pt(1).Y - Cos(Facing + piD2) * 50

ScreenPt(2).X = Pt(2).X + Sin(Facing - pi) * 50
ScreenPt(2).Y = Pt(2).Y - Cos(Facing - pi) * 50

ScreenPt(3).X = ScreenPt(2).X - CLD6 * Sin(Facing)
ScreenPt(3).Y = ScreenPt(2).Y + CLD6 * Cos(Facing)



WheelPtX = Pt(3).X + CLD6 * Sin(Facing + pi8D9)
WheelPtY = Pt(3).Y - CLD6 * Cos(Facing + pi8D9)

WheelConnectionX = WheelPtX + 250 * Sin(Facing - pi13D18)
WheelConnectionY = WheelPtY - 250 * Cos(Facing - pi13D18)

GunPtX = Pt(4).X + CLD6 * Sin(Facing + piD10)
GunPtY = Pt(4).Y - CLD6 * Cos(Facing + piD10)
GunTipPtX = GunPtX + GunLen * Sin(Stick(iStick).ActualFacing)
GunTipPtY = GunPtY - GunLen * Cos(Stick(iStick).ActualFacing)


RotorX = Pt(1).X + 280 * Sin(Facing - pi3D4)
RotorY = Pt(1).Y - 280 * Cos(Facing - pi3D4)

Rotor1X = RotorX + Stick(iStick).RotorWidth * Sin(Facing)
Rotor1Y = RotorY - Stick(iStick).RotorWidth * Cos(Facing)

Rotor2X = RotorX - Stick(iStick).RotorWidth * Sin(Facing)
Rotor2Y = RotorY + Stick(iStick).RotorWidth * Cos(Facing)



TailRotorX = Pt(6).X + 350 * Sin(Facing - pi5D9)
TailRotorY = Pt(6).Y - 350 * Cos(Facing - pi5D9)

TailRotor1X = TailRotorX + TailRotorLen * Sin(Stick(iStick).TailRotorFacing)
TailRotor1Y = TailRotorY - TailRotorLen * Cos(Stick(iStick).TailRotorFacing)

TailRotor2X = TailRotorX - TailRotorLen * Sin(Stick(iStick).TailRotorFacing)
TailRotor2Y = TailRotorY + TailRotorLen * Cos(Stick(iStick).TailRotorFacing)


Stick(iStick).GunPoint.X = GunTipPtX
Stick(iStick).GunPoint.Y = GunTipPtY

Stick(iStick).CasingPoint.X = GunPtX 'Stick(iStick).GunPoint.X
Stick(iStick).CasingPoint.Y = GunPtY 'Stick(iStick).GunPoint.Y


'MUST BE DONE BEFORE POINTS ARE SCALED INTO PIXELS
modStickGame.sCircle WheelPtX, WheelPtY, 75, vbBlack
modStickGame.sLine WheelPtX, WheelPtY, _
                   WheelConnectionX, _
                   WheelConnectionY, cg_ChopperCol

modStickGame.sLine Rotor1X, Rotor1Y, Rotor2X, Rotor2Y, vbBlack
modStickGame.sLine RotorX, RotorY, CSng(Pt(1).X + 200 * Sin(Facing - pi)), _
                                   CSng(Pt(1).Y - 200 * Cos(Facing - pi)), vbBlack

modStickGame.sLine CSng(Pt(4).X), CSng(Pt(4).Y), GunPtX, GunPtY, cg_ChopperCol

modStickGame.sCircle GunPtX, GunPtY, 50, Stick(iStick).Colour
modStickGame.sLine GunPtX, GunPtY, GunTipPtX, GunTipPtY, Stick(iStick).Colour

modStickGame.sLine WheelConnectionX, WheelConnectionY, GunPtX, GunPtY, cg_ChopperCol

'picMain.ForeColor = picMain.BackColor
If Stick(iStick).Team <> Neutral Then
    picMain.ForeColor = GetTeamColour(Stick(iStick).Team)
Else
    picMain.ForeColor = cg_ChopperCol
End If
modStickGame.sPoly Pt, cg_ChopperCol 'Stick(iStick).Colour

picMain.ForeColor = cg_ChopperCol
modStickGame.sPoly ScreenPt, Stick(iStick).Colour
picMain.DrawStyle = 0
picMain.DrawWidth = 2

modStickGame.sLine TailRotor1X, TailRotor1Y, TailRotor2X, TailRotor2Y, vbBlack


''senser/circle on top
'picMain.fillstyle = vbFSSolid
'picMain.FillColor = MSilver
'modStickGame.sCircle RotorX + 200 * Sin(Facing - piD2), RotorY - 200 * Cos(Facing - piD2), 200, ChopperCol
'picMain.fillstyle = vbFSTransparent

'On Error Resume Next
'rotor adjust
If Stick(iStick).RotorDir Then
    If Stick(iStick).RotorWidth < (RotorInc + 1) * modStickGame.sv_StickGameSpeed Then
        Stick(iStick).RotorDir = Not Stick(iStick).RotorDir
    Else
        Stick(iStick).RotorWidth = Stick(iStick).RotorWidth - RotorInc * modStickGame.sv_StickGameSpeed
    End If
Else
    If Stick(iStick).RotorWidth > CLD2 Then
        Stick(iStick).RotorDir = Not Stick(iStick).RotorDir
    Else
        Stick(iStick).RotorWidth = Stick(iStick).RotorWidth + RotorInc * modStickGame.sv_StickGameSpeed
    End If
End If


Stick(iStick).TailRotorFacing = Stick(iStick).TailRotorFacing + TailRotorInc * modStickGame.sv_StickGameSpeed
If Stick(iStick).TailRotorFacing > pi2 Then
    Stick(iStick).TailRotorFacing = FixAngle(Stick(iStick).TailRotorFacing)
End If


End Sub

Private Sub DrawTypeBubble(X As Single, Y As Single) ', Col As Long)

'modStickGame.sCircle X, Y, 100, Col
modStickGame.PrintStickText "Typing", X - 250, Y - 300, vbRed 'Col

End Sub

Private Sub DrawLegs(i As Integer, X As Single, ByVal Y As Single, _
    Crouching As Boolean, Prone As Boolean, Col As Long)

Dim Knee1X As Single, Knee1Y As Single
Dim Knee2X As Single, Knee2Y As Single
Dim iDirection As Integer, LegSgn As Single
Const HRx1p8 = HeadRadius * 1.8

If Stick(i).Facing > pi Then
    iDirection = -1
Else
    iDirection = 1
End If

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


        modStickGame.sLine X, Y + BodyLen, Knee1X, Knee1Y, Col
        modStickGame.sLine Knee1X, Knee1Y, Knee1X, Y + BodyLen / 2 + LegHeight, Col


        '############################################ 2nd knee
        'make legs slightly wider
        Knee2X = X + iDirection * Abs(Stick(i).LegWidth / 4)
        Knee2Y = Y + BodyLen + LegHeight / 4


        modStickGame.sLine X, Y + BodyLen, Knee2X, Knee2Y, Col
        modStickGame.sLine Knee2X, Knee2Y, Knee2X, Y + BodyLen / 2 + LegHeight, Col

    Else
        
        If Abs(Stick(i).LegWidth) < 44 Then
            Stick(i).LegWidth = 44 * Sgn(Stick(i).LegWidth)
        End If
        
        '############################################ 1st knee
        Knee1X = X + Stick(i).LegWidth / 2
        Knee1Y = Y + BodyLen / 2 + LegHeight / 2
        
        
        modStickGame.sLine X, Y + BodyLen, Knee1X, Knee1Y, Col
        modStickGame.sLine Knee1X, Knee1Y, Knee1X, Y + BodyLen / 2 + LegHeight, Col

        '############################################ 2nd knee
        Knee2X = X - Stick(i).LegWidth / 2
        Knee2Y = Y + BodyLen / 2 + LegHeight / 2
        
        
        modStickGame.sLine X, Y + BodyLen, Knee2X, Knee2Y, Col
        modStickGame.sLine Knee2X, Knee2Y, Knee2X, Y + BodyLen / 2 + LegHeight, Col
    End If
    
    '#########
    'modstickgame.sLine X + Stick(i).LegWidth, Y + BodyLen + LegHeight / 2,X + Stick(i).LegWidth, Y + BodyLen + LegHeight)
    'modstickgame.sLine X + -Stick(i).LegWidth, Y + BodyLen + LegHeight / 2,X - Stick(i).LegWidth, Y + BodyLen + LegHeight)
    
ElseIf Prone Then
    
    LegSgn = IIf(Stick(i).Facing > pi, 1, -1)
    
    Y = Y + HRx1p8
    
    modStickGame.sLine X + LegSgn * BodyLen, Y, _
                       X + LegSgn * (BodyLen + LegHeight / 2), Y, Col
    
Else
    modStickGame.sLine X, Y + BodyLen, X + Stick(i).LegWidth, Y + BodyLen + LegHeight, Col 'leg 1
    modStickGame.sLine X, Y + BodyLen, X - Stick(i).LegWidth, Y + BodyLen + LegHeight, Col 'leg 2
End If


End Sub

Private Sub DrawShotgun(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim X(1 To 11) As Single, Y(1 To 11) As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer

Const SAd2 = SmallAngle / 2

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)

If Facing > pi Then
    Flip = True
    
    If Reloading Then Facing = 5 * pi / 4
    
    Facing = Facing - pi
    kY = 1
Else
    If Reloading Then Facing = pi3D4
    kY = -1
End If

'hand position
Hand1X = Stick(i).X - ArmLen / 2

If StickiHasState(i, stick_crouch) Then
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


X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 2 * Sin(Facing + kY * SmallAngle)
Y(2) = Y(1) - GunLen / 2 * Cos(Facing + kY * SmallAngle)

X(3) = X(1) + GunLen / 1.5 * Sin(Facing + kY * SmallAngle)
Y(3) = Y(1) - GunLen / 1.5 * Cos(Facing + kY * SmallAngle)

X(4) = X(1) + GunLen / 1.5 * Sin(Facing + kY * SAd2)
Y(4) = Y(1) - GunLen / 1.5 * Cos(Facing + kY * SAd2)

X(5) = X(1) + GunLen * Sin(Facing + kY * SAd2)
Y(5) = Y(1) - GunLen * Cos(Facing + kY * SAd2)

'pump action bit
X(6) = X(1) + GunLen * Sin(Facing + kY * SAd2)
Y(6) = Y(1) - GunLen * Cos(Facing + kY * SAd2)

X(7) = X(1) + GunLen * 1.5 * Sin(Facing + kY * SmallAngle / 3)
Y(7) = Y(1) - GunLen * 1.5 * Cos(Facing + kY * SmallAngle / 3)
'end pump action bit

X(8) = X(1) + GunLen * 2 * Sin(Facing + kY * SmallAngle / 3)
Y(8) = Y(1) - GunLen * 2 * Cos(Facing + kY * SmallAngle / 3)

X(9) = X(1) + GunLen * 2.5 * Sin(Facing + kY * SmallAngle / 3.5)
Y(9) = Y(1) - GunLen * 2.5 * Cos(Facing + kY * SmallAngle / 3.5)

X(10) = X(9) + GunLen / 6 * Sin(Facing + kY * pi2d3)
Y(10) = Y(9) - GunLen / 6 * Cos(Facing + kY * pi2d3)

X(11) = X(9) + GunLen / 20 * Sin(Facing + kY * pi)
Y(11) = Y(9) - GunLen / 20 * Cos(Facing + kY * pi)

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
    Me.ForeColor = &H555555
    picMain.DrawWidth = 2
    
    'handle section
    modStickGame.sLine X(1), Y(1), X(3), Y(3), vbRed
    
    picMain.DrawWidth = 2
    modStickGame.sLine X(2), Y(2), X(4), Y(4), vbBlack
    
    modStickGame.sLine X(2), Y(2), X(8), Y(8), &H555555
    modStickGame.sLine X(3), Y(3), X(9), Y(9), &H555555
    
    Me.ForeColor = vbRed
    modStickGame.sLine X(1), Y(1), X(4), Y(4), vbRed
    modStickGame.sLine X(6), Y(6), X(7), Y(7), vbRed
    
    'Me.ForeColor = &H555555
    picMain.DrawWidth = 1
    modStickGame.sLine X(10), Y(10), X(11), Y(11), vbRed
    
    'modstickgame.sLine X(), Y(),X(), Y())
End If

Stick(i).GunPoint.X = X(9)
Stick(i).GunPoint.Y = Y(9)

Stick(i).CasingPoint.X = X(7)
Stick(i).CasingPoint.Y = Y(7)

picMain.DrawWidth = 1

End Sub

Private Sub DrawAK(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim X(1 To 18) As Single, Y(1 To 18) As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer

Dim tX As Single, tY As Single


Const SAd2 = SmallAngle / 2
Const SAd4 = SmallAngle / 4
Const SAd8 = SmallAngle / 8


Facing = FixAngle(Stick(i).Facing)

Reloading = StickiHasState(i, Stick_Reload)

If Facing > pi Then
    Flip = True
    
    If Reloading Then
        Facing = pi3D4
    Else
        Facing = Facing - pi
    End If
    
    kY = -1
Else
    If Reloading Then Facing = piD4
    kY = 1
End If


'hand position
Hand1X = sX + ArmLen / 4

If StickiHasState(i, stick_crouch) Then
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 4
    End If
Else
    If Flip Then
        Hand1Y = sY + HeadRadius + BodyLen / 1.2
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 4
    End If
End If


X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 4 * Sin(Facing + kY * 11 * pi / 18)
Y(2) = Y(1) - GunLen / 4 * Cos(Facing + kY * 11 * pi / 18) '90+20deg


X(3) = X(1) + GunLen / 4 * Sin(Facing + kY * piD2)
Y(3) = Y(1) - GunLen / 4 * Cos(Facing + kY * piD2)

X(4) = X(1) + GunLen / 20 * Sin(Facing)
Y(4) = Y(1) - GunLen / 20 * Cos(Facing)

X(5) = X(1) + GunLen / 4 * Sin(Facing)
Y(5) = Y(1) - GunLen / 4 * Cos(Facing)

X(6) = X(1) + GunLen / 3.2 * Sin(Facing - kY * SAd2)
Y(6) = Y(1) - GunLen / 3.2 * Cos(Facing - kY * SAd2)

X(7) = X(6) + GunLen / 1.5 * Sin(Facing + kY * piD4)
Y(7) = Y(6) - GunLen / 1.5 * Cos(Facing + kY * piD4)

X(8) = X(7) + GunLen / 4 * Sin(Facing - kY * piD4)
Y(8) = Y(7) - GunLen / 4 * Cos(Facing - kY * piD4)

X(9) = X(1) + GunLen / 2 * Sin(Facing - kY * SAd2)
Y(9) = Y(1) - GunLen / 2 * Cos(Facing - kY * SAd2)

X(10) = X(9) + GunLen * Sin(Facing - kY * SAd8)
Y(10) = Y(9) - GunLen * Cos(Facing - kY * SAd8)

X(11) = X(10) + GunLen / 4 * Sin(Facing - kY * piD2)
Y(11) = Y(10) - GunLen / 4 * Cos(Facing - kY * piD2)

X(12) = X(11) + GunLen / 4 * Sin(Facing + kY * (piD2 + SmallAngle))
Y(12) = Y(11) - GunLen / 4 * Cos(Facing + kY * (piD2 + SmallAngle))

X(13) = X(12) + GunLen / 3 * Sin(Facing - kY * pi)
Y(13) = Y(12) - GunLen / 3 * Cos(Facing - kY * pi)

X(14) = X(13) + GunLen / 3 * Sin(Facing - kY * pi)
Y(14) = Y(13) - GunLen / 3 * Cos(Facing - kY * pi)

X(15) = X(14) + GunLen * 0.6 * Sin(Facing + kY * (pi - SAd4))
Y(15) = Y(14) - GunLen * 0.6 * Cos(Facing + kY * (pi - SAd4))

X(16) = X(2) + GunLen / 2 * Sin(Facing - kY * (pi + SAd4))
Y(16) = Y(2) - GunLen / 2 * Cos(Facing - kY * (pi + SAd4))

X(17) = X(16) + GunLen / 4 * Sin(Facing + kY * (piD2 - SAd4))
Y(17) = Y(16) - GunLen / 4 * Cos(Facing + kY * (pi / 2 - SAd4))

X(18) = X(1) + GunLen / 8 * Sin(Facing - kY * pi)
Y(18) = Y(1) - GunLen / 8 * Cos(Facing - kY * pi)


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
    Me.ForeColor = &H6AD5
    'handle
    modStickGame.sLine X(1), Y(1), X(2), Y(2)
    modStickGame.sLine X(3), Y(3), X(2), Y(2)
    modStickGame.sLine X(3), Y(3), X(4), Y(4)
    
    Me.ForeColor = vbBlack
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
    Me.ForeColor = &H6AD5
    modStickGame.sLine X(9), Y(9), X(10), Y(10)
    Me.ForeColor = vbBlack
    modStickGame.sLine X(11), Y(11), X(10), Y(10) 'iron sight
    modStickGame.sLine X(11), Y(11), X(12), Y(12) 'iron sight
    Me.ForeColor = &H6AD5
    modStickGame.sLine X(13), Y(13), X(12), Y(12)
    modStickGame.sLine X(13), Y(13), X(14), Y(14)
    Me.ForeColor = vbBlack
    modStickGame.sLine X(15), Y(15), X(14), Y(14)
    
    'stock
    Me.ForeColor = &H6AD5
    modStickGame.sLine X(15), Y(15), X(16), Y(16)
    modStickGame.sLine X(17), Y(17), X(16), Y(16)
    modStickGame.sLine X(17), Y(17), X(18), Y(18)
    Me.ForeColor = vbBlack
    modStickGame.sLine X(18), Y(18), X(1), Y(1)
    
    
    If Stick(i).bSilenced Then
        DrawSilencer X(10), Y(10), Facing + IIf(Stick(i).Facing > pi, pi, 0)
    End If
End If

Stick(i).GunPoint.X = X(10)
Stick(i).GunPoint.Y = Y(10)
Stick(i).CasingPoint.X = X(6)
Stick(i).CasingPoint.Y = Y(6)

'modstickgame.sLine X(), Y(),X(), Y())

picMain.DrawWidth = 1

End Sub

Private Sub DrawSCAR(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim Pt(1 To 17) As POINTAPI
Dim PtGap(1 To 3) As POINTAPI
Dim PtMag(1 To 4) As POINTAPI

Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single

Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer
Const Scar_Col = &H101010 '&H5B5555

Dim tX As Single, tY As Single

Dim SinFacing As Single, CosFacing As Single
Dim tSin As Single, tCos As Single

Const BodyLenX1p6 = BodyLen * 1.6


Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)

If Facing > pi Then
    Flip = True
    
    If Reloading Then
        Facing = pi3D4 '1-below
    Else
        Facing = Facing - pi
    End If
    
    kY = -1
Else
    If Reloading Then Facing = piD4 'below is here
    kY = 1
End If

SinFacing = Sin(Facing)
CosFacing = Cos(Facing)

'hand position
Hand1X = sX + ArmLen / 2


If Flip Then
    If StickiHasState(i, stick_crouch) Then
        Hand1Y = sY + HeadRadius + BodyLen
    Else
        Hand1Y = sY + HeadRadius + BodyLen / 1.2
    End If
Else
    Hand1Y = sY + HeadRadius + BodyLen / 4
End If


Pt(1).X = Hand1X
Pt(1).Y = Hand1Y

Pt(2).X = Pt(1).X + GunLen / 3 * Sin(Facing + kY * pi3D4)
Pt(2).Y = Pt(1).Y - GunLen / 3 * Cos(Facing + kY * pi3D4)

Pt(3).X = Pt(2).X + GunLen / 6 * SinFacing
Pt(3).Y = Pt(2).Y - GunLen / 6 * CosFacing

Pt(4).X = Pt(1).X + GunLen / 6 * SinFacing
Pt(4).Y = Pt(1).Y - GunLen / 6 * CosFacing

Pt(5).X = Pt(4).X + GunLen / 6 * SinFacing
Pt(5).Y = Pt(4).Y - GunLen / 6 * CosFacing


tSin = Sin(Facing - kY * piD8)
tCos = Cos(Facing - kY * piD8)

'#######
PtMag(1) = Pt(5)

PtMag(2).X = Pt(5).X + GunLen / 3 * Sin(Facing + kY * pi4D9)
PtMag(2).Y = Pt(5).Y - GunLen / 3 * Cos(Facing + kY * pi4D9)

PtMag(3).X = PtMag(2).X + GunLen / 4 * tSin
PtMag(3).Y = PtMag(2).Y - GunLen / 4 * tCos

PtMag(4) = Pt(6)
'#######

Pt(6).X = Pt(5).X + GunLen / 4 * tSin
Pt(6).Y = Pt(5).Y - GunLen / 4 * tCos

Pt(7).X = Pt(8).X + GunLen / 5 * tSin
Pt(7).Y = Pt(8).Y - GunLen / 5 * tCos

'straight bottom part of barrel
Pt(8).X = Pt(9).X + GunLen / 1.5 * SinFacing
Pt(8).Y = Pt(9).Y - GunLen / 1.5 * CosFacing

'wedge
Pt(9).X = Pt(10).X + GunLen / 2.8 * Sin(Facing - kY * pi3D4)
Pt(9).Y = Pt(10).Y - GunLen / 2.8 * Cos(Facing - kY * pi3D4)


Pt(10).X = Pt(11).X + GunLen / 1.4 * Sin(Facing - kY * pi)
Pt(10).Y = Pt(11).Y - GunLen / 1.4 * Cos(Facing - kY * pi)

Pt(11).X = Pt(12).X + GunLen / 6 * Sin(Facing - kY * piD2)
Pt(11).Y = Pt(12).Y - GunLen / 6 * Cos(Facing - kY * piD2)

Pt(12).X = Pt(13).X + GunLen / 3 * Sin(Facing - kY * pi)
Pt(12).Y = Pt(13).Y - GunLen / 3 * Cos(Facing - kY * pi)

Pt(13).X = Pt(14).X + GunLen / 6 * Sin(Facing + kY * piD2)
Pt(13).Y = Pt(14).Y - GunLen / 6 * Cos(Facing + kY * piD2)

Pt(14).X = Pt(15).X + GunLen / 15 * Sin(Facing + kY * piD2)
Pt(14).Y = Pt(15).Y - GunLen / 15 * Cos(Facing + kY * piD2)

'top buttstock
Pt(15).X = Pt(15).X + GunLen / 2 * Sin(Facing - kY * (pi * 1.1))
Pt(15).Y = Pt(15).Y - GunLen / 2 * Cos(Facing - kY * (pi * 1.1))

'bottom buttstock
Pt(16).X = Pt(17).X + GunLen / 3 * Sin(Facing + kY * piD2)
Pt(16).Y = Pt(17).Y - GunLen / 3 * Cos(Facing + kY * piD2)

Pt(17).X = Pt(18).X + GunLen / 4 * Sin(Facing - kY * piD2)
Pt(17).Y = Pt(18).Y - GunLen / 4 * Cos(Facing - kY * piD2)


''start of fancy bits
'Pt(20) = Pt(9) + GunLen / 6 * tSin 'F-piD8
'Pt(20) = Pt(9) - GunLen / 6 * tCos
'
'Pt(21) = Pt(20) + GunLen / 2 * SinFacing
'Pt(21) = Pt(20) - GunLen / 2 * CosFacing
'
'Pt(22) = Pt(20) + GunLen / 6 * Sin(Facing - kY * piD2)
'Pt(22) = Pt(20) - GunLen / 6 * Cos(Facing - kY * piD2)
'
'Pt(23) = Pt(22) + GunLen / 3 * SinFacing
'Pt(23) = Pt(22) - GunLen / 3 * CosFacing


'#############
'Hole in front of scope
PtGap(1).X = Pt(16).X + GunLen / 6 * SinFacing
PtGap(1).Y = Pt(16).Y - GunLen / 6 * CosFacing

PtGap(2).X = PtGap(1).X + GunLen / 8 * Sin(Facing + kY * piD2)
PtGap(2).Y = PtGap(1).Y - GunLen / 8 * Cos(Facing + kY * piD2)

PtGap(3).X = PtGap(2).X + GunLen / 1.5 * Sin(Facing - kY * piD20)
PtGap(3).Y = PtGap(2).Y - GunLen / 1.5 * Cos(Facing - kY * piD20)

'Pt(22) = Pt(21) + GunLen / 5.2 * Sin(Facing - piD4)
'Pt(22) = Pt(21) - GunLen / 5.2 * Cos(Facing - piD4)


'#############
'barrel
Barrel1X = Pt(10).X + GunLen / 6 * Sin(Facing - kY * pi3D4)
Barrel1Y = Pt(10).Y - GunLen / 6 * Cos(Facing - kY * pi3D4)

Barrel2X = Barrel1X + GunLen / 4 * SinFacing 'GunLen/x = BarrelLen
Barrel2Y = Barrel1Y - GunLen / 4 * CosFacing

'Pt(26) = Pt(11) + GunLen / 6 * Sin(Facing + pi3D4)
'Pt(26) = Pt(11) - GunLen / 6 * Cos(Facing + pi3D4)
'
'Pt(25) = Pt(26) + GunLen / 2 * SinFacing
'Pt(25) = Pt(26) - GunLen / 2 * CosFacing


If Flip Then
    'flip image
    For j = 1 To 19
        Pt(j).Y = 2 * sY - Pt(j).Y + BodyLenX1p6
        Pt(j).X = 2 * sX - Pt(j).X
    Next j
    
    For j = 1 To 3
        PtGap(j).Y = 2 * sY - PtGap(j).Y + BodyLenX1p6
        PtGap(j).X = 2 * sX - PtGap(j).X
    Next j
    
    Barrel1X = 2 * sX - Barrel1X
    Barrel2X = 2 * sX - Barrel2X
    Barrel1Y = 2 * sY - Barrel1Y + BodyLenX1p6
    Barrel2Y = 2 * sY - Barrel2Y + BodyLenX1p6
    
    'For j = 1 To 24
        'Pt(j) = Pt(j) - 2 * (Pt(j) - Stick(i).X)
        'Pt(j) = 2 * sX - Pt(j)
    'Next j
    
    Hand1X = 2 * sX - Hand1X
    Hand1Y = Pt(1).Y
End If

Hand2X = Pt(5).X
Hand2Y = Pt(5).Y
'end calculation


Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y

Stick(i).CasingPoint.X = Pt(5).X
Stick(i).CasingPoint.Y = Pt(5).Y


'drawing
If CanSeeStick(i) Then
    picMain.DrawWidth = 1
    'Me.ForeColor = &H2F2F2F
    picMain.ForeColor = Scar_Col
    picMain.DrawWidth = 2
    
    If Stick(i).bSilenced Then
        DrawSilencer Barrel1X, Barrel1Y, Facing + IIf(Stick(i).Facing > pi, pi, 0)
    End If
    
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y, Scar_Col
    
    modStickGame.sPoly Pt, Scar_Col
    modStickGame.sPoly PtGap, modStickGame.cg_BGColour
    
    
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
    StickCol As Long)

Dim Facing As Single
Dim X(1 To 32) As Single, Y(1 To 32) As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean ', bProne As Boolean
Dim kY As Integer
Dim tX As Single, tY As Single

'calc constants
Const GLd10 = GunLen / 10
Const SAd4 = SmallAngle / 4
Dim GTC As Long
Dim BarrelLen As Single '1 = normal

Dim SinFacing As Single
Dim CosFacing As Single
Dim SinFacingLess_kYpiD2 As Single, SinFacingLess_kYpiD4 As Single
Dim CosFacingLess_kYpiD2 As Single, CosFacingLess_kYpiD4 As Single

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)
'bProne = StickiHasState(i, Stick_Prone)

If Facing > pi Then
    Flip = True
    
    If Reloading Then
        Facing = pi5D9
        'Facing = piD2
        'Facing = pi * 0.2
    Else
        Facing = Facing - pi
    End If
    
    kY = -1
Else
    If Reloading Then
        Facing = pi4D9
        'Facing = piD2
        'Facing = pi / 1.2
    End If
    kY = 1
End If

If i = 0 Then
    GTC = GetTickCount()
    If Stick(i).LastBullet + M82_Recoil_Time / modStickGame.sv_StickGameSpeed > GTC Then
        BarrelLen = (GTC - Stick(0).LastBullet) * modStickGame.sv_StickGameSpeed / 850 + 0.55
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


SinFacingLess_kYpiD2 = Sin(Facing - kY * piD2)
CosFacingLess_kYpiD2 = Cos(Facing - kY * piD2)
SinFacingLess_kYpiD4 = Sin(Facing - kY * piD4)
CosFacingLess_kYpiD4 = Cos(Facing - kY * piD4)

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


SinFacing = Sin(Facing)
CosFacing = Cos(Facing)

X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 4 * Sin(Facing - kY * piD4)
Y(2) = Y(1) - GunLen / 4 * Cos(Facing - kY * piD4)

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

X(13) = X(12) + GunLen * 1.5 * SinFacing * BarrelLen 'BARREL
Y(13) = Y(12) - GunLen * 1.5 * CosFacing * BarrelLen

X(14) = X(12) + GunLen / 8 * Sin(Facing - kY * pi)
Y(14) = Y(12) - GunLen / 8 * Cos(Facing - kY * pi)

X(15) = X(14) + GLd10 * SinFacingLess_kYpiD2
Y(15) = Y(14) - GLd10 * CosFacingLess_kYpiD2

X(16) = X(15) + GunLen / 10 * Sin(Facing - kY * pi) 'iron sight bottom
Y(16) = Y(15) - GunLen / 10 * Cos(Facing - kY * pi)

X(17) = X(16) + GunLen / 10 * SinFacingLess_kYpiD2 'iron sight top
Y(17) = Y(16) - GunLen / 10 * CosFacingLess_kYpiD2

X(18) = X(15) + GunLen / 6 * Sin(Facing - kY * pi)
Y(18) = Y(15) - GunLen / 6 * Cos(Facing - kY * pi)

X(19) = X(18) + GunLen / 2 * Sin(Facing - kY * pi) 'end of straight top bit
Y(19) = Y(18) - GunLen / 2 * Cos(Facing - kY * pi)

X(20) = X(1) + GunLen / 4 * SinFacingLess_kYpiD2
Y(20) = Y(1) - GunLen / 4 * CosFacingLess_kYpiD2

'sight stand
'bottom points
X(21) = X(18) + GunLen / 8 * Sin(Facing - kY * pi) 'forward bottom
Y(21) = Y(18) - GunLen / 8 * Cos(Facing - kY * pi)

X(22) = X(21) + GunLen / 4 * Sin(Facing - kY * pi) 'rearward bottom
Y(22) = Y(21) - GunLen / 4 * Cos(Facing - kY * pi)
'top points
X(23) = X(21) + GunLen / 6 * SinFacingLess_kYpiD2 'forward top
Y(23) = Y(21) - GunLen / 6 * CosFacingLess_kYpiD2

X(24) = X(22) + GunLen / 6 * SinFacingLess_kYpiD2 'rearward top
Y(24) = Y(22) - GunLen / 6 * CosFacingLess_kYpiD2
'modstickgame.sLine from 21->23, 22->24

'scope
X(25) = X(24) + GunLen / 4 * Sin(Facing - kY * pi) 'rear bottom pt
Y(25) = Y(24) - GunLen / 4 * Cos(Facing - kY * pi)

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
    
    X(31) = X(30) + GunLen / 2 * Sin(Facing + kY * pi / 1.8) 'GunLen/x = Height of Stand
    Y(31) = Y(30) - GunLen / 2 * Cos(Facing + kY * pi / 1.8)
    
    X(32) = X(31) + GunLen / 4 * SinFacing 'GunLen/x = separation of stands
    Y(32) = Y(31) - GunLen / 4 * CosFacing
'End If

'flash thing
X(29) = X(13) - GunLen / 6 * SinFacing
Y(29) = Y(13) + GunLen / 6 * CosFacing

'X(29) = X(11) + GunLen / 1.6 * Sin(Facing) 'GunLen/x = Start Point of Flashy Bit
'Y(29) = Y(11) - GunLen / 1.6 * Cos(Facing)
'
'X(30) = X(29) + GunLen / 4 * Sin(Facing) 'GunLen/x = Length of Flashy Bit
'Y(30) = Y(29) - GunLen / 4 * Cos(Facing)
'
'X(31) = X(30) + GunLen / 8 * SinFacingLess_kYpiD2 'GunLen/x = Height of Flashy Bit
'Y(31) = Y(30) - GunLen / 8 * CosFacingLess_kYpiD2
'
'X(32) = X(31) + GunLen / 4 * Sin(Facing - kY * pi) 'Must be same as X(30)
'Y(32) = Y(31) - GunLen / 4 * Cos(Facing - kY * pi)


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

Hand2X = X(7) + GunLen / 3 * Sin(Facing + piD2)
Hand2Y = Y(7) - GunLen / 3 * Cos(Facing + piD2)

'end calculation

'drawing

'handle
If CanSeeStick(i) Then
    
    'EXTRA HAND BIT
    modStickGame.sLine Hand2X, Hand2Y, X(10), Y(10), StickCol
    
    picMain.DrawWidth = 1
    
    'v. d. blue = &H693F3F
    Me.ForeColor = &H3F3F3F
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
    Me.ForeColor = vbBlack
    modStickGame.sLine X(12), Y(12), X(13), Y(13) 'BARREL
    
    'modStickGame.sLine X(13), Y(13), X(14), Y(14)
    'modStickGame.sLine X(14), Y(14), X(15), Y(15)
    modStickGame.sLine X(12), Y(12), X(15), Y(15)
    
    picMain.DrawWidth = 1
    Me.ForeColor = &H693F3F
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
    Me.ForeColor = vbBlack '&H555555
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
    
    If Stick(i).bSilenced Then
        DrawSilencer X(13), Y(13), Facing + IIf(Stick(i).Facing > pi, pi, 0)
    End If
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

Dim F As Single

Facing = FixAngle(Stick(i).Facing)

'hand position
Hand1X = sX + ArmLen

If Facing > pi Then
    Flip = True
    Facing = Facing - pi
    
    kY = -1
    
    
    Hand1Y = sY + HeadRadius + BodyLen / 2.8 + 100 * Cos(Facing)
Else
    kY = 1
    
    Hand1Y = sY + HeadRadius + BodyLen / 2.8 + 100 * Cos(Facing)
End If


X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + 400 * Sin(Facing)
Y(2) = Y(1) - 400 * Cos(Facing)

If Stick(i).bLightSaber = False Then
    X(5) = X(1) + 50 * Sin(Facing)
    Y(5) = Y(1) - 50 * Cos(Facing)
    
    X(3) = X(5) + 65 * Sin(Facing + piD2)
    Y(3) = Y(5) - 65 * Cos(Facing + piD2)
    
    X(4) = X(5) + 65 * Sin(Facing - piD2)
    Y(4) = Y(5) - 65 * Cos(Facing - piD2)
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
        'Me.ForeColor = SaberGreen
        modStickGame.sLine X(1), Y(1), X(2), Y(2), SaberGreen
        'modStickGame.sLine X(3), Y(3), X(4), Y(4)
        
        picMain.FillStyle = vbFSSolid
        picMain.FillColor = MSilver
        modStickGame.sCircle X(1), Y(1), 25, MSilver
        picMain.FillStyle = vbFSTransparent
    Else
        picMain.DrawWidth = 1
        'Me.ForeColor = &H3F3F3F
        modStickGame.sLine X(1), Y(1), X(2), Y(2), &H3F3F3F
        modStickGame.sLine X(3), Y(3), X(4), Y(4), &H3F3F3F
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
Dim X(1 To 16) As Single, Y(1 To 16) As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer
Dim tX As Single, tY As Single

Const SAd2 = SmallAngle / 2

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)


If Facing > pi Then
    Flip = True
    
    Facing = Facing - pi
    kY = -1
Else
    kY = 1
End If

'hand position
Hand1X = Stick(i).X + ArmLen / 2

If StickiHasState(i, stick_crouch) Then
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


X(2) = Hand1X
Y(2) = Hand1Y

X(1) = X(2) + GunLen / 2 * Sin(Facing - kY * piD2)
Y(1) = Y(2) - GunLen / 2 * Cos(Facing - kY * piD2)

X(3) = X(1) + GunLen / 1.5 * Sin(Facing)
Y(3) = Y(1) - GunLen / 1.5 * Cos(Facing)

X(4) = X(3) + GunLen / 2 * Sin(Facing + kY * piD2)
Y(4) = Y(3) - GunLen / 2 * Cos(Facing + kY * piD2)

X(5) = X(3) + GunLen / 1.5 * Sin(Facing)
Y(5) = Y(3) - GunLen / 1.5 * Cos(Facing)

X(6) = X(5) + GunLen / 4 * Sin(Facing - kY * piD2)
Y(6) = Y(5) - GunLen / 4 * Cos(Facing - kY * piD2)

X(7) = X(6) + GunLen * 3 * Sin(Facing - kY * pi) 'rear top point
Y(7) = Y(6) - GunLen * 3 * Cos(Facing - kY * pi)

X(8) = X(1) + GunLen * 1.7 * Sin(Facing - kY * pi) 'rear bottom point
Y(8) = Y(1) - GunLen * 1.7 * Cos(Facing - kY * pi)

'rear funnel
X(9) = X(7) + GunLen / 3 * Sin(Facing - kY * pi3D4) 'rear top point
Y(9) = Y(7) - GunLen / 3 * Cos(Facing - kY * pi3D4)

X(10) = X(8) + GunLen / 3 * Sin(Facing + kY * pi3D4) 'rear bottom point
Y(10) = Y(8) - GunLen / 3 * Cos(Facing + kY * pi3D4)

'sights
X(11) = X(6) + GunLen / 1.2 * Sin(Facing - kY * pi)
Y(11) = Y(6) - GunLen / 1.2 * Cos(Facing - kY * pi)

X(12) = X(11) + GunLen / 4 * Sin(Facing - kY * piD2)
Y(12) = Y(11) - GunLen / 4 * Cos(Facing - kY * piD2)

X(13) = X(12) + GunLen / 4 * Sin(Facing - kY * piD4)
Y(13) = Y(12) - GunLen / 4 * Cos(Facing - kY * piD4)

X(14) = X(13) + GunLen / 4 * Sin(Facing - kY * piD2)
Y(14) = Y(13) - GunLen / 4 * Cos(Facing - kY * piD2)

X(15) = X(14) + GunLen / 2 * Sin(Facing + kY * pi3D4)
Y(15) = Y(14) - GunLen / 2 * Cos(Facing + kY * pi3D4)

X(16) = X(15) + GunLen / 4 * Sin(Facing + kY * piD2)
Y(16) = Y(15) - GunLen / 4 * Cos(Facing + kY * piD2)

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
    Me.ForeColor = vbBlack
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
            
            If Stick(i).LastBullet + AutoReload_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                'prevent from drawing rocket straight after firing
                'and before reload state received
                
                If Flip Then
                    tX = X(6)
                    tY = Y(6)
                Else
                    tX = X(5)
                    tY = Y(5)
                End If
                
                DrawRocket tX + GunLen / 1.2 * Sin(Stick(i).Facing - piD20), _
                           tY - GunLen / 1.2 * Cos(Stick(i).Facing - piD20), _
                           Stick(i).Facing ', Stick(i).Colour
                
                
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

End Sub

Private Sub DrawM249(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim X(1 To 20) As Single, Y(1 To 20) As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer
Dim SinFacing As Single, CosFacing As Single

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)

If Facing > pi Then
    Flip = True
    
    If Reloading Then Facing = 5 * pi / 4
    
    Facing = Facing - pi
    kY = -1
Else
    If Reloading Then Facing = pi3D4
    kY = 1
End If

SinFacing = Sin(Facing)
CosFacing = Cos(Facing)


'hand position
Hand1X = Stick(i).X + ArmLen / 2

If Flip Then
    Hand1Y = sY + HeadRadius + BodyLen * 1.5
ElseIf StickiHasState(i, stick_crouch) Then
    Hand1Y = sY + HeadRadius + BodyLen / 6
Else
    Hand1Y = sY + HeadRadius + BodyLen / 4
End If


X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 2 * Sin(Facing + kY * pi3D4)
Y(2) = Y(1) - GunLen / 2 * Cos(Facing + kY * pi3D4)

X(3) = X(2) + GunLen / 4 * SinFacing
Y(3) = Y(2) - GunLen / 4 * CosFacing

X(4) = X(1) + GunLen / 4 * SinFacing
Y(4) = Y(1) - GunLen / 4 * CosFacing
'end handle

'gap between handle and handy bit
X(5) = X(4) + GunLen / 4 * SinFacing
Y(5) = Y(4) - GunLen / 4 * CosFacing

X(6) = X(5) + GunLen / 6 * Sin(Facing + kY * piD2)
Y(6) = Y(5) - GunLen / 6 * Cos(Facing + kY * piD2)

X(7) = X(6) + GunLen / 2 * SinFacing
Y(7) = Y(6) - GunLen / 2 * CosFacing

X(8) = X(5) + GunLen / 2 * SinFacing
Y(8) = Y(5) - GunLen / 2 * CosFacing

'bipod
X(9) = X(2) + GunLen * 1.2 * Sin(Facing + kY * piD10)
Y(9) = Y(2) - GunLen * 1.2 * Cos(Facing + kY * piD10)

X(10) = X(2) + GunLen * 1.5 * Sin(Facing + kY * piD10)
Y(10) = Y(2) - GunLen * 1.5 * Cos(Facing + kY * piD10)

'barrel
X(11) = X(8) + GunLen / 1.5 * SinFacing
Y(11) = Y(8) - GunLen / 1.5 * CosFacing

'sights
X(12) = X(8) + GunLen / 4 * SinFacing
Y(12) = Y(8) - GunLen / 4 * CosFacing

X(13) = X(12) + GunLen / 4 * Sin(Facing - kY * piD2)
Y(13) = Y(12) - GunLen / 4 * Cos(Facing - kY * piD2)

'top bit
X(14) = X(8) + GunLen / 10 * Sin(Facing - kY * piD2)
Y(14) = Y(8) - GunLen / 10 * Cos(Facing - kY * piD2)

'top handle
X(15) = X(14) + GunLen / 4 * Sin(Facing - kY * pi)
Y(15) = Y(14) - GunLen / 4 * Cos(Facing - kY * pi)

X(16) = X(15) + GunLen / 6 * Sin(Facing - kY * piD2)
Y(16) = Y(15) - GunLen / 6 * Cos(Facing - kY * piD2)

X(17) = X(16) + GunLen / 4 * Sin(Facing - kY * pi3D4)
Y(17) = Y(16) - GunLen / 4 * Cos(Facing - kY * pi3D4)
'end handle

X(18) = X(15) + GunLen / 4 * Sin(Facing - kY * pi)
Y(18) = Y(15) - GunLen / 4 * Cos(Facing - kY * pi)

X(18) = X(15) + GunLen / 4 * Sin(Facing - kY * pi)
Y(18) = Y(15) - GunLen / 4 * Cos(Facing - kY * pi)

X(19) = X(1) + GunLen / 2 * Sin(Facing - kY * pi)
Y(19) = Y(1) - GunLen / 2 * Cos(Facing - kY * pi)

X(20) = X(19) + GunLen / 4 * Sin(Facing + kY * piD2)
Y(20) = Y(19) - GunLen / 4 * Cos(Facing + kY * piD2)

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
    Me.ForeColor = vbBlack
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
Dim X(1 To 10) As Single, Y(1 To 10) As Single
Dim Flip As Boolean, Reloading As Boolean ', JustShot As Boolean
Dim kY As Integer, j As Integer
Dim ArmLenDist As Single, GTC As Long
Const HeadRadius2 = HeadRadius * 2 ', DEagle_Bullet_DelayD2 = DEagle_Bullet_Delay / 2

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)
'JustShot = (Stick(i).LastBullet + DEagle_Bullet_DelayD2 > getickcount())



If i = 0 Then
    GTC = GetTickCount()
    
    If Stick(i).LastBullet + DEagle_Recoil_Time / modStickGame.sv_StickGameSpeed > GTC Then
        
        ArmLenDist = (GTC - Stick(0).LastBullet) * modStickGame.sv_StickGameSpeed / 2 + 50
        
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
    If Facing > pi Then
        Flip = True
        
        If Reloading Then Facing = pi5D4
        
        Facing = Facing - pi
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
Hand1X = Stick(i).X + ArmLenDist * Sin(Facing)

If Flip Then
    Hand1Y = sY + ArmLen * 3 - ArmLenDist * Cos(Facing)
Else
    Hand1Y = sY + ArmLen / 1.5 - ArmLenDist * Cos(Facing)
End If
'End If

X(1) = Hand1X
Y(1) = Hand1Y

X(2) = X(1) + GunLen / 2 * Sin(Facing)
Y(2) = Y(1) - GunLen / 2 * Cos(Facing)

X(3) = X(2) + GunLen / 6 * Sin(Facing - kY * piD3) '60 deg
Y(3) = Y(2) - GunLen / 6 * Cos(Facing - kY * piD3)

X(4) = X(3) + GunLen / 12 * Sin(Facing - kY * piD2)
Y(4) = Y(3) - GunLen / 12 * Cos(Facing - kY * piD2)

X(5) = X(3) + GunLen / 10 * Sin(Facing - kY * pi)
Y(5) = Y(3) - GunLen / 10 * Cos(Facing - kY * pi)

X(6) = X(3) + GunLen / 1.6 * Sin(Facing - kY * pi)
Y(6) = Y(3) - GunLen / 1.6 * Cos(Facing - kY * pi)

X(6) = X(3) + GunLen / 1.6 * Sin(Facing - kY * pi)
Y(6) = Y(3) - GunLen / 1.6 * Cos(Facing - kY * pi)

X(7) = X(6) + GunLen / 4 * Sin(Facing + kY * pi8D9)
Y(7) = Y(6) - GunLen / 4 * Cos(Facing + kY * pi8D9)

X(8) = X(1) + GunLen / 6 * Sin(Facing - kY * pi)
Y(8) = Y(1) - GunLen / 6 * Cos(Facing - kY * pi)

X(9) = X(8) + GunLen / 3 * Sin(Facing + kY * pi13D18)
Y(9) = Y(8) - GunLen / 3 * Cos(Facing + kY * pi13D18)

X(10) = X(9) + GunLen / 6 * Sin(Facing)
Y(10) = Y(9) - GunLen / 6 * Cos(Facing)



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
    modStickGame.sLine X(1), Y(1), X(2), Y(2), MSilver
    modStickGame.sLine X(2), Y(2), X(3), Y(3), MSilver
    modStickGame.sLine X(3), Y(3), X(4), Y(4), vbBlack
    modStickGame.sLine X(4), Y(4), X(5), Y(5), vbBlack
    modStickGame.sLine X(5), Y(5), X(6), Y(6), MSilver
    modStickGame.sLine X(6), Y(6), X(7), Y(7), MSilver 'vbYellow
    modStickGame.sLine X(7), Y(7), X(8), Y(8), vbBlack
    modStickGame.sLine X(8), Y(8), X(9), Y(9), vbBlack
    modStickGame.sLine X(9), Y(9), X(10), Y(10), vbBlack
    
    modStickGame.sLine X(10), Y(10), X(1), Y(1), vbBlack
    
    'modstickgame.sLine X(), Y(),X(), Y())
    If Stick(i).bSilenced Then
        DrawSilencer X(2), Y(2), Facing + IIf(Stick(i).Facing > pi, pi, 0)
    End If
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
Dim kY As Integer, j As Integer
Dim MB(1 To 10) As POINTAPI
Dim FB(1 To 4) As POINTAPI
'mb = MainBarrel
'fb = FuelBox

Const ArmLenDX = ArmLen / 3
Const BodyLenD2 = BodyLen / 2
Const BodyLenX2 = BodyLen * 2

Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)

sY = GetStickY(i)

If Facing > pi Then
    Flip = True
    
    If Reloading Then
        Facing = pi3D4
    Else
        Facing = Facing - pi
    End If
    
    kY = -1
    
    If StickiHasState(i, Stick_Prone) Then
        sY = sY + BodyLen
    Else
        sY = sY + BodyLenX2
    End If
    
Else
    If Reloading Then Facing = piD4
    kY = 1
End If

Hand1X = Stick(i).X + ArmLenDX
If StickiHasState(i, Stick_Prone) Then
    If Facing > pi Then
        Hand1Y = sY - BodyLen
    Else
        Hand1Y = sY + BodyLenD2
    End If
Else
    Hand1Y = sY + BodyLen
End If

MB(1).X = Hand1X
MB(1).Y = Hand1Y

MB(2).X = MB(1).X + GunLen / 5 * Sin(Facing)
MB(2).Y = MB(1).Y - GunLen / 5 * Cos(Facing)

MB(3).X = MB(2).X + GunLen / 3 * Sin(Facing - kY * piD4)
MB(3).Y = MB(2).Y - GunLen / 3 * Cos(Facing - kY * piD4)

MB(4).X = MB(3).X + GunLen * Sin(Facing)
MB(4).Y = MB(3).Y - GunLen * Cos(Facing)

MB(5).X = MB(4).X + GunLen / 6 * Sin(Facing - kY * piD4)
MB(5).Y = MB(4).Y - GunLen / 6 * Cos(Facing - kY * piD4)

MB(6).X = MB(5).X + GunLen / 3 * Sin(Facing - kY * piD6)
MB(6).Y = MB(5).Y - GunLen / 3 * Cos(Facing - kY * piD6)

MB(7).X = MB(6).X + GunLen / 10 * Sin(Facing - kY * piD2)
MB(7).Y = MB(6).Y - GunLen / 10 * Cos(Facing - kY * piD2)

MB(8).X = MB(7).X + GunLen / 3 * Sin(Facing - kY * pi)
MB(8).Y = MB(7).Y - GunLen / 3 * Cos(Facing - kY * pi)

MB(9).X = MB(8).X + GunLen / 3 * Sin(Facing + kY * pi3D4)
MB(9).Y = MB(8).Y - GunLen / 3 * Cos(Facing + kY * pi3D4)

MB(10).X = MB(9).X + GunLen * Sin(Facing - kY * pi)
MB(10).Y = MB(9).Y - GunLen * Cos(Facing - kY * pi)

If Not Reloading Then
    FB(1).X = MB(3).X '+ GunLen / 4 * Sin(Facing)
    FB(1).Y = MB(3).Y '- GunLen / 4 * Sin(Facing)
    
    FB(2).X = MB(3).X + GunLen / 2 * Sin(Facing) 'glDx = boxlen
    FB(2).Y = MB(3).Y - GunLen / 2 * Cos(Facing)
    
    FB(3).X = FB(2).X + GunLen / 3 * Sin(Facing + kY * piD2) 'glDx = boxheight
    FB(3).Y = FB(2).Y - GunLen / 3 * Cos(Facing + kY * piD2)
    
    FB(4).X = FB(3).X + GunLen / 4 * Sin(Facing - kY * pi)
    FB(4).Y = FB(3).Y - GunLen / 4 * Cos(Facing - kY * pi)
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

Private Sub DrawSA80(i As Integer, Hand1X As Single, Hand1Y As Single, _
    Hand2X As Single, Hand2Y As Single, ByVal sX As Single, ByVal sY As Single)

Dim Facing As Single
Dim j As Integer
Dim Flip As Boolean, Reloading As Boolean
Dim kY As Integer
Dim sX2 As Single, sY2 As Single
Const kGreen = 32768 '32768=rgb(0,128,0)

Dim pGrip(1 To 4) As POINTAPI
Dim ptBarrel(1 To 4) As POINTAPI
Dim ptMain(1 To 5) As POINTAPI
Dim PtMag(1 To 4) As POINTAPI
Dim ptSights(1 To 4) As POINTAPI
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Dim Stock1X As Single, Stock1Y As Single, Stock2X As Single, Stock2Y As Single

Dim SinFacing As Single, CosFacing As Single


Facing = FixAngle(Stick(i).Facing)
Reloading = StickiHasState(i, Stick_Reload)

If Facing > pi Then
    Flip = True
    
    If Reloading Then
        Facing = piD4 '1-below
    Else
        Facing = Facing - pi
    End If
    
    kY = -1
Else
    If Reloading Then Facing = pi3D4 'below is here
    kY = 1
End If

SinFacing = Sin(Facing)
CosFacing = Cos(Facing)

'hand position
Hand1X = sX + ArmLen * SinFacing '* 2 / 3
If Flip Then
    'If StickiHasState(i, stick_Crouch) Then
        'Hand1Y = sY + HeadRadius + BodyLen
    'Else
        Hand1Y = sY - HeadRadius * 1.5 - ArmLen * CosFacing
    'End If
Else
    Hand1Y = sY + HeadRadius + BodyLen / 6 - ArmLen * CosFacing
End If


'grip
pGrip(1).X = Hand1X
pGrip(1).Y = Hand1Y

pGrip(2).X = pGrip(1).X + GunLen / 3 * Sin(Facing + kY * pi3D4)
pGrip(2).Y = pGrip(1).Y - GunLen / 3 * Cos(Facing + kY * pi3D4)

pGrip(3).X = pGrip(2).X + GunLen / 4 * SinFacing
pGrip(3).Y = pGrip(2).Y - GunLen / 4 * CosFacing

pGrip(4).X = pGrip(1).X + GunLen / 4 * SinFacing
pGrip(4).Y = pGrip(1).Y - GunLen / 4 * CosFacing
'end grip

'green barrel part
ptBarrel(1).X = pGrip(4).X
ptBarrel(1).Y = pGrip(4).Y

ptBarrel(2).X = ptBarrel(1).X + GunLen * 2 / 3 * SinFacing 'GL/x = Green Len
ptBarrel(2).Y = ptBarrel(1).Y - GunLen * 2 / 3 * CosFacing

ptBarrel(3).X = ptBarrel(2).X + GunLen / 5 * Sin(Facing - kY * pi2d3) '100deg
ptBarrel(3).Y = ptBarrel(2).Y - GunLen / 5 * Cos(Facing - kY * pi2d3)

ptBarrel(4).X = ptBarrel(1).X + GunLen / 4 * Sin(Facing - kY * piD2)
ptBarrel(4).Y = ptBarrel(1).Y - GunLen / 4 * Cos(Facing - kY * piD2)
'end green barrel

'black barrel
Barrel1X = (ptBarrel(2).X + ptBarrel(3).X) / 2
Barrel1Y = (ptBarrel(2).Y + ptBarrel(3).Y) / 2
Barrel2X = Barrel1X + GunLen / 8 * SinFacing
Barrel2Y = Barrel1Y - GunLen / 8 * CosFacing

'main black bit
ptMain(1).X = ptBarrel(4).X
ptMain(1).Y = ptBarrel(4).Y

ptMain(2).X = ptMain(1).X - GunLen * SinFacing 'length that it goes back (to the stock)
ptMain(2).Y = ptMain(1).Y + GunLen * CosFacing

ptMain(3).X = ptMain(2).X + GunLen / 2 * Sin(Facing + kY * piD2)
ptMain(3).Y = ptMain(2).Y - GunLen / 2 * Cos(Facing + kY * piD2)

ptMain(4).X = pGrip(1).X - GunLen / 3 * SinFacing
ptMain(4).Y = pGrip(1).Y + GunLen / 3 * CosFacing

ptMain(5).X = ptBarrel(1).X
ptMain(5).Y = ptBarrel(1).Y

'magazine
If Not Reloading Then
    PtMag(1).X = pGrip(1).X - GunLen / 2.5 * SinFacing
    PtMag(1).Y = pGrip(1).Y + GunLen / 2.5 * CosFacing
    
    PtMag(2).X = PtMag(1).X - GunLen / 6 * SinFacing 'GL/x = Mag Width
    PtMag(2).Y = PtMag(1).Y + GunLen / 6 * CosFacing
    
    PtMag(3).X = PtMag(2).X + GunLen / 2 * Sin(Facing + kY * piD3)
    PtMag(3).Y = PtMag(2).Y - GunLen / 2 * Cos(Facing + kY * piD3)
    
    PtMag(4).X = PtMag(1).X + GunLen / 2 * Sin(Facing + kY * piD3)
    PtMag(4).Y = PtMag(1).Y - GunLen / 2 * Cos(Facing + kY * piD3)
End If

'sights
'bottom right
ptSights(1).X = pGrip(1).X + GunLen / 3 * Sin(Facing - kY * piD2)
ptSights(1).Y = pGrip(1).Y - GunLen / 3 * Cos(Facing - kY * piD2)

'top right
ptSights(2).X = ptSights(1).X + GunLen / 6 * Sin(Facing - kY * piD4) 'GL/x = sight height
ptSights(2).Y = ptSights(1).Y - GunLen / 6 * Cos(Facing - kY * piD4)

'top left
ptSights(3).X = ptSights(2).X - GunLen / 2 * SinFacing
ptSights(3).Y = ptSights(2).Y + GunLen / 2 * CosFacing

'bottom left
ptSights(4).X = ptSights(1).X - GunLen / 4 * SinFacing
ptSights(4).Y = ptSights(1).Y + GunLen / 4 * CosFacing




'#############
Stock1X = CSng(ptMain(2).X)
Stock1Y = CSng(ptMain(2).Y)
Stock2X = CSng(ptMain(3).X)
Stock2Y = CSng(ptMain(3).Y)
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
        PtMag(j).X = sX2 - PtMag(j).X
        PtMag(j).Y = sY2 - PtMag(j).Y
    Next j
    For j = 1 To 4
        ptSights(j).X = sX2 - ptSights(j).X
        ptSights(j).Y = sY2 - ptSights(j).Y
    Next j
    Barrel1X = sX2 - Barrel1X: Barrel1Y = sY2 - Barrel1Y
    Barrel2X = sX2 - Barrel2X: Barrel2Y = sY2 - Barrel2Y
    Stock1X = sX2 - Stock1X: Stock1Y = sY2 - Stock1Y
    Stock2X = sX2 - Stock2X: Stock2Y = sY2 - Stock2Y
    
    Hand1X = sX2 - Hand1X
    Hand1Y = sY2 - Hand1Y
End If



Hand2X = (ptBarrel(1).X + ptBarrel(2).X) / 2
Hand2Y = (ptBarrel(1).Y + ptBarrel(2).Y) / 2
'end calculation


Stick(i).GunPoint.X = Barrel2X
Stick(i).GunPoint.Y = Barrel2Y

Stick(i).CasingPoint.X = PtMag(1).X
Stick(i).CasingPoint.Y = PtMag(1).Y



'drawing
If CanSeeStick(i) Then
    picMain.DrawWidth = 1
    picMain.DrawStyle = vbFSSolid
    picMain.ForeColor = vbBlack
    
    'sight stand
    modStickGame.sLine CLng(ptSights(1).X), _
                       CLng(ptSights(1).Y), _
                       CLng(ptSights(1).X + GunLen / 6 * Sin(Facing + piD2)), _
                       CLng(ptSights(1).Y - GunLen / 6 * Cos(Facing + piD2)), vbBlack
    modStickGame.sLine CLng(ptSights(4).X), _
                       CLng(ptSights(4).Y), _
                       CLng(ptSights(4).X + GunLen / 6 * Sin(Facing + piD2)), _
                       CLng(ptSights(4).Y - GunLen / 6 * Cos(Facing + piD2)), vbBlack
    
    
    
    
    modStickGame.sPoly_NoOutline pGrip, kGreen
    modStickGame.sPoly_NoOutline ptBarrel, kGreen
    modStickGame.sPoly ptMain, vbBlack
    modStickGame.sPoly ptSights, vbBlack
    If Not Reloading Then
        modStickGame.sPoly PtMag, vbBlack
    End If
    
    
    picMain.DrawWidth = 2
    'barrel
    modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y, vbBlack
    picMain.DrawWidth = 3
    'rear butt stock
    modStickGame.sLine Stock1X, Stock1Y, Stock2X, Stock2Y, kGreen
    
    
    
    If Stick(i).bSilenced Then
        DrawSilencer Barrel2X, Barrel2Y, Facing + IIf(Stick(i).Facing > pi, pi, 0)
    End If
    
    picMain.DrawWidth = 1
End If



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
    Speed As Single, iStick As Integer, Colour As Long, iType As eNadeTypes, Optional IsRPG As Boolean) ', Optional Sticki As Integer)

ReDim Preserve Nade(NumNades)

With Nade(NumNades)
    .Decay = GetTickCount() + Nade_Time / modStickGame.sv_StickGameSpeed
    .Heading = Heading
    .Speed = Speed
    .X = X
    .Y = Y
    .OwnerID = Stick(iStick).ID
    .IsRPG = IsRPG
    .Colour = Colour
    .iType = iType
End With

If IsRPG Then
    If iStick = 0 Or Stick(iStick).IsBot Then
        If Stick(iStick).WeaponType = RPG Then 'might be a chopper
            Stick(iStick).BulletsFired = Stick(iStick).BulletsFired + 1
        End If
    End If
End If

NumNades = NumNades + 1

End Sub

Private Sub AddMine(X As Single, Y As Single, OwnerID As Integer, Colour As Long, Heading As Single, Speed As Single)

ReDim Preserve Mine(NumMines)

With Mine(NumMines)
    .X = X
    .Y = Y
    .OwnerID = OwnerID
    .Colour = Colour
    .Heading = Heading
    .Speed = Speed
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
    
    .Decay = GetTickCount() + Mag_Decay / modStickGame.sv_StickGameSpeed
    .iMagType = vMagType
End With

NumMags = NumMags + 1

End Sub

Private Sub AddDeadStick(X As Single, Y As Single, Colour As Long, bFacingRight As Boolean, bFlamed As Boolean) ', Heading As Single)

If modStickGame.cg_DeadSticks Then
    ReDim Preserve DeadStick(NumDeadSticks)
    
    With DeadStick(NumDeadSticks)
        .X = X
        .Y = Y
        .Colour = Colour
        .Decay = GetTickCount() + DeadStickTime
        '.Heading = 0
        .Speed = 0 '20
        
        .bFacingRight = bFacingRight
        
        .bFlamed = bFlamed
        
        If bFlamed Then
            AddSmokeTrail X + PM_Rnd * ArmLen, Y + BodyLen + HeadRadius, True
            AddSmokeTrail X + PM_Rnd * ArmLen, Y + BodyLen + HeadRadius, True
        End If
        
    End With
    
    NumDeadSticks = NumDeadSticks + 1
End If

End Sub

Private Sub AddDeadChopper(X As Single, Y As Single, Colour As Long, iOwner As Integer)

ReDim Preserve DeadChopper(NumDeadChoppers)

With DeadChopper(NumDeadChoppers)
    .X = X
    .Y = Y
    .Colour = Colour
    .Decay = GetTickCount() + DeadChopperTime
    .Speed = 0
    .iOwner = iOwner
End With

NumDeadChoppers = NumDeadChoppers + 1

End Sub

Private Sub AddBlood(X As Single, Y As Single, Heading As Single, bArmour As Boolean)

If modStickGame.cg_Blood Then
    ReDim Preserve Blood(NumBlood)
    
    With Blood(NumBlood)
        .Decay = GetTickCount() + Blood_Time / modStickGame.sv_StickGameSpeed - 50 * Rnd()
        .Heading = Heading + PM_Rnd * piD8
        .Speed = 100
        .X = X
        .Y = Y
        .bArmour = bArmour
    End With
    
    NumBlood = NumBlood + 1
End If

End Sub

Private Sub AddSpark(X As Single, Y As Single, Heading As Single, Speed As Single)

If modStickGame.cg_Sparks Then
    ReDim Preserve Spark(NumSparks)
    
    With Spark(NumSparks)
        .Decay = GetTickCount() + Spark_Time / modStickGame.sv_StickGameSpeed - 100 * Rnd()
        .Heading = Heading
        .Speed = Speed
        .X = X
        .Y = Y
    End With
    
    NumSparks = NumSparks + 1
End If

End Sub

Private Sub AddBullet(X As Single, Y As Single, Speed As Single, Heading As Single, _
    OwnerID As Integer, ByVal Damage As Single, iStick As Integer, _
    Optional ByVal bSnipe As Boolean = False, Optional ByVal bAddFlash As Boolean = False)

Dim sgTmp As Single
Dim bChopper As Boolean

ReDim Preserve Bullet(NumBullets)

bChopper = (Stick(iStick).WeaponType = Chopper)

With Bullet(NumBullets)
    
    sgTmp = Rnd() * Speed / 10 + 10
    
    If bSnipe = False Then
        .Decay = GetTickCount() + Bullet_Decay / modStickGame.sv_StickGameSpeed - 100 * Rnd()
    Else
        .Decay = GetTickCount() + Bullet_Decay * 2 / modStickGame.sv_StickGameSpeed
        
        'y + + Speed * Cos(Heading)
        
        AddSmokeGroup X, Y, 3, sgTmp, Heading - piD4 - 0.5 * Rnd(), True
        AddSmokeGroup X, Y, 3, sgTmp, Heading + piD4 + 0.5 * Rnd(), True
        
        AddSmokeGroup X, Y, 3, sgTmp, Heading - piD8 - 0.5 * Rnd(), True
        AddSmokeGroup X, Y, 3, sgTmp, Heading + piD8 + 0.5 * Rnd(), True
        
        AddSmokeGroup X, Y, 3, sgTmp, Heading, True
        AddSmokeGroup X, Y, 3, sgTmp + 10, Heading, True
        
    End If
    
    .Heading = Heading
    '.Facing = Stick(iStick).Facing
    .Speed = Speed
    .X = X
    .Y = Y
    .Owner = OwnerID
    .bSniperBullet = bSnipe
    .bShotgunBullet = (Stick(iStick).WeaponType = Shotgun)
    
    If Stick(iStick).Perk = pStoppingPower Then
        Damage = Damage * StoppingPowerIncrease
    End If
    
    If Stick(iStick).bSilenced Then 'And Not (Rnd() > 0.9) Then
        .Damage = Damage / 2
        .bSilenced = True
    Else
        .Damage = Damage
    End If
    
    '.bDraw = bSnipe Or (Rnd() > 0.9)
    
    
    'AddExplosion X, Y, 300, 0.1, 30, Heading
    If Rnd() > 0.3 Or bSnipe Or bAddFlash Then
        If Stick(iStick).WeaponType <> Shotgun Then
            If Stick(iStick).bSilenced = False Then
                Stick(iStick).LastMuzzleFlash = GetTickCount()
            End If
        Else
            'Stick(iStick).LastMuzzleFlash = GetTickCount() - MFlash_Time / 2
            AddExplosion X, Y, 300, 0.1, 30, Heading
        End If
    End If
    
    If modStickGame.cg_BulletSmoke Then
        If Rnd() > 0.6 Then
            If Not bSnipe Then
                AddSmokeGroup X + .Speed * Sin(.Heading), Y - .Speed * Cos(.Heading), 4, sgTmp / 4, Heading
            End If
        End If
    End If
    
    'If Rnd() > 0.4 Then
    AddCasing CSng(Stick(iStick).CasingPoint.X), CSng(Stick(iStick).CasingPoint.Y), _
            Heading, Stick(iStick).Facing, bSnipe Or bChopper, bChopper
    'End If
    
End With

NumBullets = NumBullets + 1

If iStick = 0 Or Stick(iStick).IsBot Then
    Stick(iStick).BulletsFired = Stick(iStick).BulletsFired + 1
End If

If Stick(iStick).WeaponType = SA80 Then
    Stick(iStick).BulletsFired2 = Stick(iStick).BulletsFired2 + 1
End If

End Sub

'Private Sub AddMuzzleFlash(X As Single, Y As Single, Facing As Single)
'ReDim Preserve MFlash(NumMFlashes)
'
'With MFlash(NumMFlashes)
'    .Decay = GetTickCount() + MFlash_Time / modStickGame.sv_StickGameSpeed
'    .Facing = Facing
'    .X = X
'    .Y = Y
'End With
'
'NumMFlashes = NumMFlashes + 1
'End Sub

Private Sub AddCasing(X As Single, Y As Single, Facing As Single, StickFacing As Single, bSnipe As Boolean, _
    bChopper As Boolean)

If modStickGame.cg_Casing Then
    ReDim Preserve Casing(NumCasings)
    
    With Casing(NumCasings)
        .Decay = GetTickCount() + Casing_Time / modStickGame.sv_StickGameSpeed
        If Not bChopper Then
            .Heading = IIf(StickFacing < pi, pi7D4, piD4) ' 0 'pi
        End If
        
        .Facing = Facing
        If bSnipe Then
            .Speed = 40
        Else
            .Speed = 25 'Rnd() * 10 + 10
        End If
        .X = X
        .Y = Y
        .bSniperCasing = bSnipe
    End With
    
    NumCasings = NumCasings + 1
End If

End Sub

Private Sub AddStaticWeapon(X As Single, Y As Single, vWeapon As eWeaponTypes)

ReDim Preserve StaticWeapon(NumStaticWeapons)

With StaticWeapon(NumStaticWeapons)
    .X = X
    .Y = Y
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
    ReDim Preserve Casing(NumCasings - 2)
    NumCasings = NumCasings - 1
End If

End Sub

Private Sub DrawBullets()

Dim i As Integer
Dim pX As Single, pY As Single

'Remove any decayed bullets
i = 0
Do While i < NumBullets
    'Is this one decayed?
    If Bullet(i).Decay < GetTickCount() Then
        'Kill it!
        RemoveBullet i, False
        'Decrement the counter
        i = i - 1
    End If
    'Increment the counter
    i = i + 1
Loop

'Step through each bullet and draw it
picMain.DrawWidth = 2
picMain.FillStyle = vbFSTransparent

For i = 0 To NumBullets - 1
    'Draw the bullet
    'modstickgame.sCircle  Bullet(i).x - (BULLET_RADIUS + 0.5), Bullet(i).y - (BULLET_RADIUS + 0.5), _
        Bullet(i).x + BULLET_RADIUS + 0.5, Bullet(i).y + BULLET_RADIUS + 0.5, Me.hdc
    
    If Bullet(i).bSilenced = False Then
        
        On Error GoTo EH
        pX = CLng(Bullet(i).X + Sin(Bullet(i).Heading) * Bullet(i).Speed)
        pY = CLng(Bullet(i).Y - Cos(Bullet(i).Heading) * Bullet(i).Speed)
        
        modStickGame.sLine Bullet(i).X, Bullet(i).Y, pX, pY, vbYellow
        
        If Bullet(i).bSniperBullet Then
            If Bullet(i).LastSmoke + Sniper_Smoke_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                AddSmoke Bullet(i).X, Bullet(i).Y, 10, Bullet(i).Heading, False
                Bullet(i).LastSmoke = GetTickCount()
            End If
        End If
        
        'picMain.fillstyle = vbFSSolid
        'picMain.FillColor = Bullet(i).Colour
        'modstickgame.sCircle  (Bullet(i).X, Bullet(i).Y), Bullet_Radius, Bullet(i).Colour
    End If
    
Next i

EH:
picMain.FillStyle = vbFSTransparent
End Sub

Private Sub ProcessSmokeBlasts()
Dim i As Integer, j As Integer
Const BlastCol = BoxCol 'SmokeFill
'Const MaxWidth = 20, MaxLen = 300

'old
Const sgMaxSize = 30
Const LineLen = 10
Dim PM_Offset As Single: PM_Offset = PM_Rnd() * pi / 2

'picMain.DrawWidth = 2

Do While i < NumSmokeBlasts
    
'    With SmokeBlast(i)
'        If Int(.sWidth) Then
'            picMain.DrawWidth = Int(.sWidth)
'        End If
'
'
'        modStickGame.sLine .X, .Y, .X + .sLength * Sin(.Heading), .Y - .sLength * Cos(.Heading), BlastCol
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
    '3 lines from .x, .y to .x+kSin(Heading+-a),.y+kCos(Heading+-a)
    '##################
    With SmokeBlast(i)
        modStickGame.sLine .X + .sLength * Sin(.Heading), .Y - .sLength * Cos(.Heading), _
                           .X + (.sLength * LineLen) * Sin(.Heading), .Y - (.sLength * LineLen) * Cos(.Heading), BlastCol
        
        For j = 0 To 4
            
            modStickGame.sLine .X + .sLength * Sin(.Heading), .Y - .sLength * Cos(.Heading), _
                           .X + (.sLength * LineLen) * Sin(.Heading + .sOffset + PM_Rnd()), .Y - (.sLength * LineLen) * Cos(.Heading + .sOffset + PM_Rnd()), BlastCol
            
            modStickGame.sLine .X + .sLength * Sin(.Heading), .Y - .sLength * Cos(.Heading), _
                           .X + (.sLength * LineLen) * Sin(.Heading - .sOffset + PM_Rnd()), .Y - (.sLength * LineLen) * Cos(.Heading - .sOffset + PM_Rnd()), BlastCol
            
        Next j
        .sLength = .sLength + 4 * modStickGame.StickTimeFactor
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

If modStickGame.cg_BulletSmoke Then
    ReDim Preserve SmokeBlast(NumSmokeBlasts)
    
    With SmokeBlast(NumSmokeBlasts)
        .Heading = Heading + PM_Rnd() * piD4
        
        .X = X
        .Y = Y
        
        '.iDir = 1
        
        .sOffset = PM_Rnd() * pi / 4
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

i = 0
Do While i < NumCasings
    
    If Casing(i).Decay < GetTickCount() Then
        RemoveCasing i
        i = i - 1
    End If
    
    i = i + 1
Loop


picMain.DrawWidth = 1
For i = 0 To NumCasings - 1
    
    If Casing(i).bSniperCasing Then
        'picMain.DrawWidth = 2
        
        modStickGame.sLine Casing(i).X, Casing(i).Y, _
          Casing(i).X + 2 * Casing_Len * Sin(Casing(i).Facing) _
        , Casing(i).Y - 2 * Casing_Len * Cos(Casing(i).Facing), vbYellow
        
        'picMain.DrawWidth = 1
        
    Else
        modStickGame.sLine Casing(i).X, Casing(i).Y, _
          Casing(i).X + Casing_Len * Sin(Casing(i).Facing) _
        , Casing(i).Y - Casing_Len * Cos(Casing(i).Facing), vbYellow
        
        
    End If
    
    
Next i

picMain.DrawWidth = 1

End Sub

Private Sub DrawBlood()
Dim i As Integer
Const Blood_Radius = 10

picMain.DrawWidth = 3
picMain.FillStyle = vbFSSolid

For i = 0 To NumBlood - 1
    modStickGame.sCircle Blood(i).X, Blood(i).Y, Blood_Radius, IIf(Blood(i).bArmour, MSilver, vbRed)
    'modstickgame.sCircle  (Blood(i).X + Rnd() * 30, Blood(i).Y + Rnd() * 30), Bullet_Radius / 2, vbRed
Next i

picMain.DrawWidth = 1

End Sub

Private Sub RemoveBullet(Index As Integer, bWall As Boolean, Optional bFancy As Boolean = True) ', Optional sgWallSurface As Single) ', Optional ByVal WithSmoke As Boolean = True)

Dim i As Integer

If bFancy Then
    If Bullet(Index).Speed > 50 Then
        If bWall Then
            
            AddWallMark Bullet(Index).X, Bullet(Index).Y, WallMark_Bullet_Radius
            
            AddSmokeBlast Bullet(Index).X, Bullet(Index).Y, Bullet(Index).Heading - pi ', sgWallSurface
            
            If modStickGame.cg_BulletSmoke Then
                If Rnd() > 0.8 Or Bullet(Index).bSniperBullet Then
                    AddSmokeTrail Bullet(Index).X, Bullet(Index).Y
                End If
            End If
            
            If Rnd() > 0.3 Then
                'AddExplosion Bullet(Index).X, Bullet(Index).Y, 150, 0.25, 0, 0
                AddCirc Bullet(Index).X, Bullet(Index).Y, 150, 0.25, vbYellow
            End If
            If Rnd() > 0.2 Then
                AddSparks Bullet(Index).X, Bullet(Index).Y, Bullet(Index).Heading - pi
            End If
        ElseIf modStickGame.cg_BulletSmoke Then
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
'        Bullet(i).Decay = Bullet(i + 1).Decay
'        Bullet(i).Heading = Bullet(i + 1).Heading
'        Bullet(i).Speed = Bullet(i + 1).Speed
'        Bullet(i).X = Bullet(i + 1).X
'        Bullet(i).Y = Bullet(i + 1).Y
'        Bullet(i).Owner = Bullet(i + 1).Owner
'        Bullet(i).LastDiffract = Bullet(i + 1).LastDiffract
'        Bullet(i).Damage = Bullet(i + 1).Damage
'        Bullet(i).bSniperBullet = Bullet(i + 1).bSniperBullet
'        'Bullet(i).bDraw = Bullet(i + 1).bDraw
        Bullet(i) = Bullet(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Bullet(NumBullets - 2)
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
    ReDim Preserve Spark(NumSparks - 2)
    NumSparks = NumSparks - 1
End If

End Sub

Private Sub ProcessSparks()
Dim i As Integer
Dim pX As Single, pY As Single

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
    If Spark(i).LastReduction + Spark_Speed_Reduction_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        Spark(i).Speed = Spark(i).Speed / Spark_Speed_Reduction
        Spark(i).LastReduction = GetTickCount()
    End If
    
    StickMotion Spark(i).X, Spark(i).Y, Spark(i).Speed, Spark(i).Heading
    
    pX = CSng(Spark(i).X + Sin(Spark(i).Heading) * Spark(i).Speed)
    pY = CSng(Spark(i).Y - Cos(Spark(i).Heading) * Spark(i).Speed)
    
    modStickGame.sLine Spark(i).X, Spark(i).Y, pX, pY, vbYellow
Next i

End Sub

Private Sub AddSparks(X As Single, Y As Single, GeneralHeading As Single)
Dim i As Integer

For i = 0 To 5
    AddSpark X, Y, GeneralHeading + PM_Rnd * Spark_Diffraction, Spark_Speed + Rnd() * 20
Next i


End Sub

''########################################################################
'flame stuff
Private Sub AddFlame(X As Single, Y As Single, Heading As Single, Speed As Single, OwnerID As Integer, _
    iStick As Integer)

ReDim Preserve Flame(NumFlames)

With Flame(NumFlames)
    .Decay = GetTickCount() + Flame_Time / modStickGame.sv_StickGameSpeed
    .Heading = Heading
    .Speed = Speed
    .X = X
    .Y = Y
    .OwnerID = OwnerID
End With

NumFlames = NumFlames + 1

If iStick = 0 Or Stick(iStick).IsBot Then
    Stick(iStick).BulletsFired = Stick(iStick).BulletsFired + 1
End If

If modStickGame.cg_BulletSmoke Then
    If Rnd() > 0.5 Then
        AddSmokeGroup Stick(iStick).GunPoint.X, Stick(iStick).GunPoint.Y, 4, Rnd() * Speed, Heading
    End If
End If


End Sub

Private Sub RemoveFlame(Index As Integer)

Dim i As Integer

If modStickGame.cg_BulletSmoke Then
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

Private Sub ProcessFlames()
Dim i As Integer, j As Integer
Dim iFlameOwner As Integer ', MinDist As Single
Dim bTouching As Boolean
Const CLDx = ChopperLen / 1.2

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

picMain.FillStyle = vbFSSolid
i = 0
Do While i < NumFlames
    
    Flame(i).Size = Flame(i).Size + 6 * modStickGame.StickTimeFactor
    
    StickMotion Flame(i).X, Flame(i).Y, Flame(i).Speed, Flame(i).Heading
    DrawFlame Flame(i).X, Flame(i).Y, Flame(i).Size
    
    If ClipFlame(i) = False Then
        
        iFlameOwner = FindStick(Flame(i).OwnerID)
        
        If iFlameOwner > -1 Then
            For j = 0 To NumSticksM1
                If Stick(j).ID <> Flame(i).OwnerID Then
                    If IsAlly(Stick(j).Team, Stick(iFlameOwner).Team) = False Then
                        'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
                        If StickInvul(j) = False Then
                            
                            bTouching = False
                            
                            If Stick(j).WeaponType = Chopper Then
                                
                                bTouching = CoOrdInChopper(Flame(i).X, Flame(i).Y, j)
                                
                            Else
                                bTouching = (GetDist(Stick(j).X, Stick(j).Y, Flame(i).X, Flame(i).Y) < 500)
                            End If
                            
                            'MinDist = IIf(Stick(j).WeaponType = Chopper, ChopperLen, FlameRadiusXK)
                            'If GetDist(Stick(j).X, Stick(j).Y, Flame(i).X, Flame(i).Y) < MinDist Then
                            
                            If bTouching Then
                                Stick(j).LastFlameTouch = GetTickCount()
                                Stick(j).LastFlameTouchOwnerID = Flame(i).OwnerID
                                Stick(j).bFlameIsFromTag = False
                                
                                If j = 0 Or Stick(j).IsBot Then
                                    DamageStick Flame_Damage, j
                                    
                                    
                                    If Stick(j).Health < 1 Then
                                        Killed j, FindStick(Flame(i).OwnerID), kFlame
                                    End If
                                    
                                End If
                                
                                If Stick(j).WeaponType = Chopper Then
                                    RemoveFlame i
                                    i = i - 1
                                    Exit For
                                End If
                                
                            End If
                            
                            
                        End If
                    End If
                End If
            Next j
        End If
        
    Else
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
                
                
                DrawFlame Stick(i).X + ArmLen / 2 * Rnd(), _
                    GetStickY(i) + IIf(StickiHasState(i, Stick_Prone), HeadRadius, BodyLen) * Rnd(), _
                    Flame_Burn_Radius
                
                
                'damage
                If i = 0 Or Stick(i).IsBot Then
                    If Stick(i).LastFlameDamage + Flame_Burn_Damage_Time / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                        
                        DamageStick Flame_Burn_Damage, i, False
                        
                        
                        If Stick(i).Health < 1 Then
                            Killed i, FindStick(Stick(i).LastFlameTouchOwnerID), _
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

picMain.FillStyle = vbFSTransparent

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

Private Function CoOrdNearStick(X As Single, Y As Single, Sticki As Integer) As Boolean

Const XLimit = ArmLen * 3, YLimit = BodyLen * 2
Dim sY As Single

If Stick(Sticki).WeaponType = Chopper Then
    CoOrdNearStick = False
    
Else
    If X < (Stick(Sticki).X + XLimit) Then
        If X > (Stick(Sticki).X - XLimit) Then
            
            sY = GetStickY(Sticki)
            If Y > (sY - HeadRadius) Then
                If Y < (sY + YLimit) Then
                    CoOrdNearStick = True
                End If
            End If
            
        End If
    End If
End If

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

Const Lim As Integer = 50
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
    
End If

End Function

Private Function FlameInBox(i As Integer) As Boolean
Dim j As Integer

For j = 0 To nBoxes
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

For j = 0 To ntBoxes
    If FlameCollision(i, tBox(j).Left, tBox(j).Top, tBox(j).width, tBox(j).height) Then
        FlameInTBox = True
        Exit For
    End If
Next j

End Function

Private Function FlameInPlatform(i As Integer) As Boolean
Dim j As Integer

For j = 0 To nPlatforms
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
'Dim i As Integer
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
    Erase DeadStick
    NumDeadSticks = 0
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

'Make a socket
socket = modWinsock.CreateSocket()
If socket = WINSOCK_ERROR Then
    'Handle error
    'modWinsock.TermWinsock
    GoTo EH
End If

'If we're the StickServer, bind to the StickServer port
If StickServer Then
    If modWinsock.BindSocket(socket, PORT_STICKSERVER) = WINSOCK_ERROR Then
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

Private Function ConnectToServer() As Boolean

Dim JoinTimer As Long
'Dim TimeOutTimer As Long
Dim sPacket As String
Dim TempSockAddr As ptsockaddr
Dim CurrentRetry As Integer
Dim Txt As String


Dim LastLine As Long
Const LineDelay = 20

Me.ForeColor = vbBlack
'Me.CurrentX = 7
'Me.CurrentY = 7
'Print "Connecting to Server..."


'Make the server's ptsockaddr
If MakeSockAddr(ServerSockAddr, PORT_STICKSERVER, modStickGame.StickServerIP) = WINSOCK_ERROR Then
    'Handle error
    AddText "Error - IP isn't valid", TxtError, True 'Making Socket", TxtError, True
    Unload Me
    
Else
    
    'Send "Join" packets to the server until we receive an "ACK" mPacket
    CurrentRetry = 1
    
    Do 'While TimeOutTimer + SERVER_CONNECT_DURATION > GetTickCount()
        
        'Is it time to send a "Join" mPacket?
        If (JoinTimer + StickServer_RETRY_FREQ) < GetTickCount() Then
            'Reset the timer
            JoinTimer = GetTickCount()
            
            'Send the mPacket
            modWinsock.SendPacket socket, ServerSockAddr, sJoins
            
            If CurrentRetry < 6 Then
                Me.picMain.Cls
                
                Txt = "Connecting to Server '" & modStickGame.StickServerIP & "'..."
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
        sPacket = modWinsock.ReceivePacket(socket, TempSockAddr)
        
        If LenB(sPacket) Then
            
            'Is this an ACK?
            If Left$(sPacket, 1) = sAccepts Then
                
                'Set our ID
                MyID = CInt(Right$(sPacket, Len(sPacket) - 1))
                Stick(0).ID = MyID
                
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
    
End If

End Function

Private Sub CheckForKills(sTxt As String)
Dim sName As String
Dim KillType As String
Dim i As Integer, j As Integer
Const Space1 = vbSpace

'If InStr(1, sTxt, modMessaging.MsgNameSeparator) = 0 Then
On Error GoTo lEH
i = InStr(1, sTxt, "by", vbTextCompare) + 3
sName = Mid$(sTxt, i)

If LCase$(sName) = LCase$(Trim$(Stick(0).Name)) Then
    j = InStr(1, sTxt, "was", vbTextCompare) + 4
    KillType = Mid$(sTxt, j, i - j - 4)
    
    AddMainMessage "You " & KillType & Space1 & Left$(sTxt, j - 6)
    
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

Private Function GetPacket() As Boolean

Dim sPacket As String
Dim TempSockAddr As ptsockaddr
Dim i As Integer, j As Integer

Dim Tmp As String, sTxt As String
Dim bTmp As Boolean

'Loop until there were no packets
GetPacket = True

Do
    'Check for packets
    sPacket = modWinsock.ReceivePacket(socket, TempSockAddr)
    
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
                
            Case sKillInfos
                
                
                On Error GoTo EH
                i = CInt(Mid$(sPacket, 2)) 'id of killer
                j = FindStick(i) 'index
                
                If j <> -1 Then
                    
                    'Debug.Print Trim$(Stick(j).Name) & " killed"
                    
                    If j > 0 Or modStickGame.StickServer Then
                        Stick(j).iKills = Stick(j).iKills + 1
                    End If
                    
                    If Not modStickGame.StickServer Then
                        Stick(j).iKillsInARow = Stick(j).iKillsInARow + 1
                    End If
                    
                    
                    If modStickGame.StickServer Then 'tell everyone else
                        SendBroadcast sKillInfos & CStr(i)
                    End If
                    
                    If i = MyID Then
                        
                        If Stick(0).WeaponType = FlameThrower Then
                            'If Stick(0).LastBullet + 750 > GetTickCount() Then
                            FlamesInARow = FlamesInARow + 1
                            'End If
                        End If
                        
                        Call CheckKillsInARow
                    End If
                    
                    
                End If
                
            Case sDeathInfos
                'server tells all to add dead body + reset stick's lastspawntime
                
                'spacket = "D41"
                'id = 4
                'toasty = 1
                
                On Error GoTo EH
                i = FindStick(CInt(Mid$(sPacket, 2, Len(sPacket) - 2)))  'index (yes, index) of dead guy
                bTmp = CBool(Right$(sPacket, 1)) 'crusty?
                
                If i <> -1 Then
                    
                    If Stick(i).WeaponType = Chopper Then
                        AddDeadChopper Stick(i).X, Stick(i).Y, Stick(i).Colour, i
                    Else
                        AddDeadStick Stick(i).X, Stick(i).Y, Stick(i).Colour, (Stick(i).Facing < pi), bTmp
                    End If
                    
                    Stick(i).iKillsInARow = 0
                    Stick(i).LastSpawnTime = GetTickCount()
                    
                    If modStickGame.sv_GameType = gCoOp Or modStickGame.sv_GameType = gElimination Then
                        Stick(i).bAlive = False
                    End If
                    
                    If i > 0 Then 'not me
                        Stick(i).iDeaths = Stick(i).iDeaths + 1
                    End If
                    
                    
                    If modStickGame.StickServer Then 'tell everyone else
                        SendBroadcast sDeathInfos & CStr(Stick(i).ID) & CStr(Abs(bTmp)), Stick(i).ID
                    End If
                    
                    
                    
                    'reset some stuff
                    SubStickiState i, stick_Left
                    SubStickiState i, stick_Right
                    Stick(i).OnSurface = False
                    Stick(i).LastFlameTouch = 1
                    
                End If
                
                
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
                        AddChatText Trim$(Stick(i).Name) & " left the game", Stick(i).Colour
                        
                        If modStickGame.StickServer Then
                            SendBroadcast sExits & CStr(j), j
                            Pause 5
                        End If
                        
                        RemoveStick i
                    End If
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
                    modWinsock.SendPacket socket, Stick(j).ptsockaddr, sKicks & "Same Name"
                    Exit Sub 'so we don't get errors, will be checked again
                End If
            End If
            
        Next j
    Next i
    
    LastCheck = GetTickCount()
End If

lEH:
End Sub


Private Sub ProcessJoinPacket(vSockAddr As ptsockaddr)

Dim i As Long
Dim ID As String
Dim Index As Integer
Dim MaxID As Integer

'If this IP address is already in our Stick array, use pre-assigned ID
For i = 0 To NumSticksM1
    'Is it the same IP and port?
    If (Stick(i).ptsockaddr.sin_addr = vSockAddr.sin_addr) And _
                (Stick(i).ptsockaddr.sin_port = vSockAddr.sin_port) Then
        
        ID = CStr(Stick(i).ID)
        
        Exit For
    End If
Next i

'New Stick?
If Len(ID) = 0 And (vSockAddr.sin_addr <> 0) Then
    'Make a spot
    Index = AddStick()
    'Find a new ID
    MaxID = 0
    For i = 0 To NumSticksM1
        'Is this ID greater?
        If Stick(i).ID > MaxID Then MaxID = Stick(i).ID
    Next i
    'Assign the ID
    Stick(Index).ID = MaxID + 1
    
    'Set the Stick's ptsockaddr
    Stick(Index).ptsockaddr.sin_addr = vSockAddr.sin_addr
    Stick(Index).ptsockaddr.sin_family = vSockAddr.sin_family
    Stick(Index).ptsockaddr.sin_port = vSockAddr.sin_port
    Stick(Index).ptsockaddr.sin_zero = vSockAddr.sin_zero
    
    'Set the ID String
    ID = CStr(Stick(Index).ID)
End If

'Send the ACK
If (vSockAddr.sin_addr <> 0) Then
    modWinsock.SendPacket socket, vSockAddr, sAccepts & ID
End If

End Sub

Private Sub ReceiveBoxInfo(sTxt As String)
Dim i As Integer

'format: 10101101
'1 = present
'0 = gone

On Error Resume Next
For i = 0 To nBoxes
    Box(i).bInUse = CBool(Mid$(sTxt, i + 1, 1))
Next i

'if lenb(tag) = 0 then showbox

End Sub

Private Sub SendBoxInfo()
Static LastSend As Long
Dim i As Integer
Dim sPacketToSend As String

If LastSend + BoxInfoDelay < GetTickCount() Then
    
    For i = 0 To nBoxes
        sPacketToSend = sPacketToSend & IIf(Box(i).bInUse, "1", "0")
    Next i
    
    SendBroadcast sBoxInfos & sPacketToSend
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub SendChatPacket(ChatText As String, Colour As Long)

'Is this the StickServer?
If StickServer Then
    'Broadcast the chat mPacket
    SendChatPacketBroadcast ChatText, Colour
Else
    'Send it to the StickServer
    modWinsock.SendPacket socket, ServerSockAddr, sChats & ChatText & "#" & CStr(Colour)
End If

End Sub

Public Sub SendChatPacketBroadcast(ChatText As String, Colour As Long)

Dim i As Long

'Send the mPacket to everyone but the local user
For i = 0 To NumSticksM1
    'Is this the local user?
    If Stick(i).ID <> MyID Then
        If Stick(i).IsBot = False Then
            modWinsock.SendPacket socket, Stick(i).ptsockaddr, sChats & ChatText & "#" & CStr(Colour)
        End If
    End If
Next i

'Add text to local user's chat text array
AddChatText ChatText, Colour

End Sub

Public Sub SendBroadcast(Text As String, Optional ByVal NtID As Integer = -1)

Dim i As Long

'Send the mPacket to everyone but the local user
For i = 1 To NumSticksM1
    'Is this the local user?
    If Stick(i).ID <> NtID Then 'stick(i).ID <> MyID And
        If Stick(i).IsBot = False Then
            modWinsock.SendPacket socket, Stick(i).ptsockaddr, Text
        End If
    End If
Next i


End Sub

Private Sub AddChatText(ChatText As String, Colour As Long)

'Add this value to the chat text array
ReDim Preserve Chat(NumChat)
Chat(NumChat).Decay = GetTickCount() + CHAT_DECAY
Chat(NumChat).Text = ChatText
Chat(NumChat).Colour = Colour

If InStr(1, ChatText, modMessaging.MsgNameSeparator) Then
    Chat(NumChat).bChatMessage = True
Else
    Call CheckForKills(ChatText)
End If

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

Private Sub EndWinsock()

'Kill winsock
modWinsock.DestroySocket socket
'modWinsock.TermWinsock

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Button = vbLeftButton Then
    FireKey = True
ElseIf Button = vbRightButton Then
    'If Stick(0).WeaponType <> Chopper Then
    AddStickiState 0, Stick_Nade
    Stick(0).NadeStart = GetTickCount()
    'End If
ElseIf Button = vbMiddleButton Then
    If modStickGame.cl_MiddleMineDrop Then
        AddStickiState 0, Stick_Mine
    End If
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If Button = vbLeftButton Then
'    If Stick(0).WeaponType <> RPG Then
'        SubStickState MyID, Stick_Fire
'    End If
'Else
'ElseIf Button = vbMiddleButton Then
    'AddStick 'True

'If Button = vbRightButton Then
    'SubStickState MyID, Stick_Nade
'End If

If Button = vbLeftButton Then
    FireKey = False
    FireKeyUpTime = GetTickCount()
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

MouseX = X
MouseY = Y


'On error GoTo EH
'With Stick(0)
'    If (.State And Stick_Reload) = 0 Then
'        'If .WeaponType = AK Or .WeaponType = Knife Or _
'            .WeaponType = SCAR Or .WeaponType = M249 Or .WeaponType = DEagle Then
'
'        If .WeaponType <> M82 Then
'            If .WeaponType <> RPG Then
'                If .WeaponType <> Shotgun Then
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
    If modStickGame.sv_2Weapons Then
        Call Form_KeyDown(vbKey1, 0)
    Else
        If Stick(0).WeaponType <> Chopper Then
            If StickiHasState(0, Stick_Reload) = False Then
                Scroll_WeaponKey = Scroll_WeaponKey + IIf(bScrollUp, -1, 1)
                
                If Scroll_WeaponKey = -1 Then
                    Scroll_WeaponKey = Knife
                ElseIf Scroll_WeaponKey > Knife Then
                    Scroll_WeaponKey = AK
                End If
                
                LastScrollWeaponSwitch = GetTickCount()
            End If
        End If
    End If
End If

End Sub

Private Sub SetMyStickFacing()

'Stick(0).Facing = FindAngle(Stick(0).x, Stick(0).y + HeadRadius / 1.5, MouseX, MouseY)

'Stick(0).Facing = FindAngle(Stick(0).x * cg_sZoom - cg_sCamera.x, _
                             Stick(0).y * cg_sZoom - cg_sCamera.y, _
                             MouseX, _
                             MouseY)


If StickiHasState(0, Stick_Prone) Then
    Stick(0).Facing = FindAngle_Actual(Stick(0).X * cg_sZoom - cg_sCamera.X, _
                             (Stick(0).Y + BodyLen * 1.3) * cg_sZoom - cg_sCamera.Y, _
                             MouseX, _
                             MouseY)
    
    If Stick(0).OnSurface Then
        If Stick(0).Facing > pi Then
            'facing left
            If Stick(0).Facing < ProneLeftLimit Then
                Stick(0).Facing = ProneLeftLimit
            End If
        Else
            'facing right
            If Stick(0).Facing > ProneRightLimit Then
                Stick(0).Facing = ProneRightLimit
            End If
        End If
    End If
    
ElseIf Stick(0).WeaponType <> Chopper Then
    Stick(0).Facing = FindAngle_Actual(Stick(0).X * cg_sZoom - cg_sCamera.X, _
                             (Stick(0).Y + HeadRadius) * cg_sZoom - cg_sCamera.Y, _
                             MouseX, _
                             MouseY)
    
    'Stick(0).Facing = piD2
    
Else
    Stick(0).Facing = FindAngle_Actual((Stick(0).X - CLD6) * cg_sZoom - cg_sCamera.X, _
                             (Stick(0).Y + CLD4) * cg_sZoom - cg_sCamera.Y, _
                             MouseX, _
                             MouseY)
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

'If modLoadProgram.IsIDE() = False Then
If modSubClass.bStickSubClassing Then
    modSubClass.SubClassStick Me.hWnd, False
End If

bRunning = False

If modStickGame.StickServer Then
    SendBroadcast sExits & MyID
    
    
    sTmp = eCommands.LobbyCmd & eLobbyCmds.Remove & modStickGame.StickServerIP & "S"
    If Server Then
        DataArrival sTmp
    Else
        'remove from lobby
        SendData sTmp
    End If
Else
    modWinsock.SendPacket socket, ServerSockAddr, sExits & MyID
End If

Call ResetVars
Call EndWinsock

WindowClosing = True
Call FormLoad(Me, True)

modStickGame.StickFormLoaded = False
End Sub

Private Sub ResetVars()
Dim i As Integer

NumSticks = 0: Erase Stick
NumBullets = 0: Erase Bullet
NumSmoke = 0: Erase Smoke
NumBlood = 0: Erase Blood
NumNades = 0: Erase Nade
NumCasings = 0: Erase Casing
NumMines = 0: Erase Mine
NumDeadSticks = 0: Erase DeadStick
NumMags = 0: Erase Mag
NumDeadChoppers = 0: Erase DeadChopper
NumSparks = 0: Erase Spark
NumFlames = 0: Erase Flame
NumChat = 0: Erase Chat
NumCircs = 0: Erase Circs
NumMainMessages = 0: Erase MainMessages
NumStaticWeapons = 0: Erase StaticWeapon
NumLargeSmokes = 0: Erase LargeSmoke
NumWallMarks = 0: Erase WallMark


MyID = 0

LastUpdatePacket = 0

strChat = vbNullString
bChatActive = False

'KillsInARow = 0
FlamesInARow = 0
picToasty.Visible = False
bHadRadar = False
For i = 0 To CInt(eWeaponTypes.Knife)
    AmmoFired(i) = 0
Next i

ChopperAvail = False
RadarStartTime = 0

ResetKeys

modStickGame.sv_AllowChoppers = True
modStickGame.sv_AllowRockets = True
modStickGame.sv_AllowFlameThrowers = True
modStickGame.sv_Hardcore = False
modStickGame.sv_ShootNades = True
modStickGame.sv_StickGameSpeed = 1
modStickGame.sv_2Weapons = True
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
        GetObjFacing = pi
    'Else
        'GetObjFacing = 0
    End If
End If

End Function

Private Function ClipBullet(i As Integer) As Boolean

Const Lim As Integer = 50
Const Bullet_SpeedX2 = BULLET_SPEED * 2
Const Sniper_Bullet_Diffract_Delay = Bullet_Diffract_Delay * 1.5
Dim ClippedX As Boolean, ClippedY As Boolean, bSlowDownBullet As Boolean, bIsFastBullet As Boolean
Dim XComp As Single, YComp As Single

'is the bullet on the top, left, bottom or right of a wall?
Dim BulletIsLeft As Boolean, BulletIsTop As Boolean

BulletIsLeft = (Bullet(i).X < Lim)
ClippedX = BulletIsLeft Or (Bullet(i).X > StickGameWidth - Lim)
BulletIsTop = (Bullet(i).Y < Lim)
ClippedY = BulletIsTop Or (Bullet(i).Y > StickGameHeight - Lim)

bIsFastBullet = (Bullet(i).Speed > Bullet_Min_Speed)



If Bullet(i).Speed > Bullet_SpeedX2 Then
    Bullet(i).Speed = Bullet_SpeedX2
End If

If ClippedX Or ClippedY Then
    
    ClipBullet = True
    RemoveBullet i, True ', GetObjFacing(ClippedX, BulletIsLeft, BulletIsTop)
    
ElseIf BulletInPlatform(i) Then ', BulletIsLeft, BulletIsTop, ClippedX) Then
    
    If Bullet(i).bSniperBullet Then
        If bIsFastBullet Then
            bSlowDownBullet = True
        End If
    Else
        ClipBullet = True
        RemoveBullet i, True ', GetObjFacing(ClippedX, BulletIsLeft, BulletIsTop)
    End If
    
ElseIf BulletInTBox(i) Then
    
    If Bullet(i).bSniperBullet Then
        If bIsFastBullet Then
            bSlowDownBullet = True
        End If
    Else
        ClipBullet = True
        RemoveBullet i, True
    End If
    
ElseIf BulletInBox(i) Then
    
    
    If Bullet(i).LastDiffract + Bullet_Diffract_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        
        If Bullet(i).bSniperBullet = False Then
            Bullet(i).Heading = Bullet(i).Heading + PM_Rnd * piD6
            Bullet(i).Speed = Bullet(i).Speed / 2 '1.5
            'Bullet(i).Facing = Bullet(i).Heading
        Else
            Bullet(i).Speed = Bullet(i).Speed / 2
        End If
        
        If Bullet(i).LastDiffract = 0 Then
            Bullet(i).Damage = Bullet(i).Damage / 2
        End If
        
        Bullet(i).LastDiffract = GetTickCount()
    End If
    
ElseIf Bullet(i).Speed < Bullet_Min_Speed Then
    ClipBullet = True
    RemoveBullet i, False
    
End If

If bSlowDownBullet Then 'slow down sniper bullets ONLY
    If Bullet(i).LastDiffract + Sniper_Bullet_Diffract_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        
        If Bullet(i).LastDiffract = 0 Then
            Bullet(i).Damage = M82_Wall_Damage 'Bullet(i).Damage / 3
        End If
        
        'If Bullet(i).LastDiffract = 0 Then
            Bullet(i).Speed = Bullet(i).Speed / 1.2
            'Bullet(i).LastDiffract = 1
        'End If
        
        Bullet(i).LastDiffract = GetTickCount()
        
    End If
End If


End Function

Private Function BulletInBox(i As Integer) As Boolean
Dim j As Integer

For j = 0 To nBoxes
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

For j = 0 To ntBoxes
    If BulletCollision(i, tBox(j).Left, tBox(j).Top, tBox(j).width, tBox(j).height) Then
        BulletInTBox = True
        Exit For
    End If
Next j

End Function

Private Function BulletInPlatform(i As Integer) As Boolean ', ByRef bLeft As Boolean, ByRef bTop As Boolean, _
    ByRef bXClip As Boolean) As Boolean

Dim j As Integer
Const LeftLim = 100

For j = 0 To nPlatforms
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

Private Function NadeInPlatform(iNade As Integer) As Boolean
Dim j As Integer

For j = 0 To nPlatforms
    If NadeCollision(iNade, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        NadeInPlatform = True
        Exit For
    End If
Next j

End Function

Private Function NadeInBox(iNade As Integer) As Boolean
Dim j As Integer

For j = 0 To nBoxes
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

For j = 0 To ntBoxes
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

Private Sub ClipEdges(i As Integer, bLBoundSpeed As Boolean)

Const Lim As Integer = 50
Const ValIn = 30
Dim ClippedX As Boolean, ClippedY As Boolean
Dim XComp As Single, YComp As Single

ClippedY = (Stick(i).Y < Lim)
ClippedX = (Stick(i).X > StickGameWidth - Lim) Or (Stick(i).X < Lim)

If ClippedX Then 'Or ClippedY Then
    With Stick(i)
        XComp = .Speed * Sin(.Heading)
        YComp = .Speed * Cos(.Heading)
        
        If Stick(i).X < Lim Then
            XComp = Abs(XComp)
        Else
            XComp = -Abs(XComp)
        End If
        
        SubStickiState i, stick_Left
        SubStickiState i, stick_Right
        
        If i = 0 Then
            LeftKey = False
            RightKey = False
        End If
        
    End With
End If

If ClippedY Then
    With Stick(i)
        XComp = .Speed * Sin(.Heading)
        YComp = .Speed * Cos(.Heading)
        
        If Stick(i).Y < Lim Then
            YComp = -Abs(YComp)
            
            If Stick(i).ID = MyID Then
                If Stick(i).WeaponType <> Chopper Then
                    'Stick(i).Helth = Stick(i).Health - Stick(i).Speed / 15
                    DamageStick Stick(i).Speed / 15, i
                    
                    If Stick(i).Health < 1 Then
                        Call Killed(i, i, kNormal)
                    End If
                End If
            End If
            
        'Else
            'YComp = -Abs(YComp)
        End If
        
        SubStickiState i, stick_Left
        SubStickiState i, stick_Right
        
    End With
End If


If ClippedX Or ClippedY Then
    Stick(i).Speed = Sqr(XComp ^ 2 + YComp ^ 2)
    
    If YComp > 0 Then Stick(i).Heading = Atn(XComp / YComp)
    If YComp < 0 Then Stick(i).Heading = Atn(XComp / YComp) + pi
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
    End If
End With

End Sub

Private Sub LBoundSpeed(i As Integer)

If StickiHasState(i, stick_crouch) Or StickiHasState(i, Stick_Prone) Then
    If Stick(i).Speed <= 5 Then
        If StickHasMoveState(i) = False Then
            Stick(i).Speed = 0
        End If
    End If
Else
    If Stick(i).Speed < 2 Then
        Stick(i).Speed = 0
    End If
End If

End Sub

Private Sub ReverseYComp(Speed As Single, Heading As Single)

Dim XComp As Single
Dim YComp As Single
 
'Determine the components of the resultant vector
XComp = Speed * Sin(Heading)
YComp = Speed * Cos(Heading)

YComp = -YComp

'Calculate the resultant direction, and adjust for arctangent by adding Pi if necessary
If YComp > 0 Then
    Heading = Atn(XComp / YComp)
ElseIf YComp < 0 Then
    Heading = Atn(XComp / YComp) + pi
End If

End Sub

Private Sub ReverseXComp(Speed As Single, Heading As Single)

Dim XComp As Single
Dim YComp As Single
 
'Determine the components of the resultant vector
XComp = Speed * Sin(Heading)
YComp = Speed * Cos(Heading)

XComp = -XComp

'Calculate the resultant direction, and adjust for arctangent by adding Pi if necessary
If YComp > 0 Then
    Heading = Atn(XComp / YComp)
ElseIf YComp < 0 Then
    Heading = Atn(XComp / YComp) + pi
End If

End Sub

Private Function StickHasMoveState(i As Integer) As Boolean

If StickiHasState(i, stick_Left) Then
    StickHasMoveState = True
ElseIf StickiHasState(i, stick_Right) Then
    StickHasMoveState = True
End If

End Function
'##############################################################################
'Smoke ########################################################################
'##############################################################################
Private Sub AddLargeSmoke(X As Single, Y As Single, Heading As Single)
'Const MaxSize = 300, MinSize = 100
Dim i As Integer ', Face As Single

ReDim Preserve LargeSmoke(NumLargeSmokes)

With LargeSmoke(NumLargeSmokes)
    .CentreX = X
    .CentreY = Y
    
    .iDirection = 1
    
    For i = 1 To 10
        
        .SingleSmoke(i).DistanceFromMain = 10
        .SingleSmoke(i).AngleFromMain = Rnd() * pi2
        .SingleSmoke(i).AspectDir = 1
        .SingleSmoke(i).sAspect = 1
        .SingleSmoke(i).DistanceFromMainInc = 0.5
        
'        .SingleSmoke(i).X = X '+ (i - 2) * Spacing
'        .SingleSmoke(i).Y = Y
'        .SingleSmoke(i).Speed = 2 + Rnd()
'        .SingleSmoke(i).Heading = Heading + piD6 * Sgn(PM_Rnd())
    Next i
    
'    For i = 3 To 4
'        .SingleSmoke(i).X = X
'        .SingleSmoke(i).Y = Y
'        .SingleSmoke(i).Speed = 2
'        .SingleSmoke(i).Heading = pi3D2 + piD10 * (i - 3) * Sgn(PM_Rnd())
'    Next i
    
    
    
    
'    For i = 1 To 4
'        .X(i) = X + Spacing * (i - 2)
'        .Y(i) = Y
'    Next i
'    .pPoly(1).X = X
'    .pPoly(1).Y = Y - MaxSize
'
'    For i = 2 To 10
'
'        Face = Face + piD10
'
'        .pPoly(i).X = X + Rnd() * MinSize * Sin(Face)
'        .pPoly(i).Y = Y + Rnd() * MinSize * Cos(Face)
'    Next i
    
    
End With
NumLargeSmokes = NumLargeSmokes + 1

End Sub

Private Sub RemoveLargeSmoke(Index As Integer)
Dim i As Integer

If NumLargeSmokes = 1 Then
    Erase LargeSmoke
    NumLargeSmokes = 0
Else
    For i = Index To NumLargeSmokes - 2
        LargeSmoke(i) = LargeSmoke(i + 1)
    Next i
    
    'Resize the array
    NumLargeSmokes = NumLargeSmokes - 1
    ReDim Preserve LargeSmoke(NumLargeSmokes - 1)
End If

End Sub

Private Sub ProcessAndDrawLargeSmokes()
Dim i As Integer, j As Integer
Const Size_Inc = 4, Size_Dec = 1, Space_Inc = 4
Const Max_Size = 2500, Min_Size = 5

picMain.FillStyle = vbFSSolid
picMain.FillColor = SmokeFill

Do While i < NumLargeSmokes
    
'    For j = 1 To 10
'        LargeSmoke(i).pPoly(j).X = LargeSmoke(i).pPoly(j).X - _
'            Sgn(LargeSmoke(i).X - LargeSmoke(i).pPoly(j).X) * Size_Inc * modStickGame.StickTimeFactor * _
'            LargeSmoke(i).iDirection
'
'
'        LargeSmoke(i).pPoly(j).Y = LargeSmoke(i).pPoly(j).Y - _
'            Sgn(LargeSmoke(i).Y - LargeSmoke(i).pPoly(j).Y) * Size_Inc * modStickGame.StickTimeFactor * _
'            LargeSmoke(i).iDirection
'
'    Next j
    
    'DrawSmoke LargeSmoke(i).pPoly, MGrey
    'modStickGame.sCircle LargeSmoke(i).X, LargeSmoke(i).Y, LargeSmoke(i).iSize * 3, MGrey
    
    For j = 1 To 10
        
        
        'modStickGame.StickMotion LargeSmoke(i).SingleSmoke(j).X, LargeSmoke(i).SingleSmoke(j).Y, _
               LargeSmoke(i).SingleSmoke(j).Speed, LargeSmoke(i).SingleSmoke(j).Heading
        
        'LargeSmoke(i).Y(j) = LargeSmoke(i).Y(j) - _
            Sgn(LargeSmoke(i).CentreY - LargeSmoke(i).Y(j)) * Space_Inc * modStickGame.StickTimeFactor / 4
        
        RotateLargeSmokePart i, j
        DrawLargeSmokePart i, j
        
    Next j
    
    
    If LargeSmoke(i).iDirection = 1 Then
        LargeSmoke(i).iSize = LargeSmoke(i).iSize + Size_Inc * modStickGame.StickTimeFactor
    Else
        LargeSmoke(i).iSize = LargeSmoke(i).iSize - Size_Dec * modStickGame.StickTimeFactor
    End If
    
    
    If LargeSmoke(i).iDirection = 1 Then
        If LargeSmoke(i).iSize > Max_Size Then
            LargeSmoke(i).iDirection = -1
            
            For j = 1 To 10
                LargeSmoke(i).SingleSmoke(j).DistanceFromMainInc = -0.5
            Next j
            
        End If
        
    ElseIf LargeSmoke(i).iSize <= Min_Size Then
        RemoveLargeSmoke i
        i = i - 1
        
    ElseIf LargeSmoke(i).iSize > Max_Size Then
        'limit
        LargeSmoke(i).iSize = Max_Size
    End If
    
    
    
    i = i + 1
Loop

picMain.FillStyle = vbFSTransparent

End Sub

Private Sub RotateLargeSmokePart(i As Integer, j As Integer)
Const AngleInc = pi / 250

With LargeSmoke(i).SingleSmoke(j)
    .DistanceFromMain = .DistanceFromMain + .DistanceFromMainInc * modStickGame.StickTimeFactor
    
    .AngleFromMain = FixAngle(.AngleFromMain + AngleInc * modStickGame.StickTimeFactor)
    
    .sAspect = .sAspect + 0.001 * .AspectDir * modStickGame.StickTimeFactor
    
    
    If .sAspect > 1.2 Then
        .AspectDir = -1
    ElseIf .sAspect < 0.8 Then
        .AspectDir = 1
    End If
    
End With

End Sub

Private Sub DrawLargeSmokePart(i As Integer, j As Integer) ', bFull As Boolean)
Dim tX As Single, tY As Single

With LargeSmoke(i).SingleSmoke(j)
    tX = LargeSmoke(i).CentreX + .DistanceFromMain * Sin(.AngleFromMain)
    tY = LargeSmoke(i).CentreY - .DistanceFromMain * Cos(.AngleFromMain)
    
    
    modStickGame.sCircleAspect tX, tY, LargeSmoke(i).iSize / 2, SmokeFill, .sAspect
End With

'If bFull Then
'Else
'    modStickGame.sHatchCircle _
'        LargeSmoke(i).SingleSmoke(j).X, _
'        LargeSmoke(i).SingleSmoke(j).Y, _
'        MGrey, LargeSmoke(i).iSize / 25
'End If

End Sub

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

Private Sub AddSmokeTrail(ByVal X As Single, ByVal Y As Single, Optional bLong As Boolean = False) ', _
    ByVal Speed As Single, ByVal Heading As Single)


AddSmokeGroup X, Y, 4, 3, PM_Rnd * piD4, bLong
AddSmokeGroup X, Y, 3, 2, PM_Rnd * piD4, bLong
'AddSmokeGroup X, Y, 3, 2, pm_rnd * piD4

End Sub

Private Sub AddSmokeGroup(ByVal X As Single, ByVal Y As Single, ByVal HowMany As Integer, _
    ByVal Speed As Single, ByVal Heading As Single, Optional bLong As Boolean = False)

Dim i As Integer
Const MaxSpacing = 75
Dim rX As Single, rY As Single

For i = 1 To HowMany
    rX = X + (Rnd() - 0.5) * MaxSpacing
    rY = Y + (Rnd() - 0.5) * MaxSpacing
    
    AddSmoke rX, rY, Speed, Heading, bLong
Next i

End Sub

Private Sub AddSmoke(X As Single, Y As Single, Speed As Single, Heading As Single, bLongTime As Boolean)

'If modStickGame.cg_Smoke Then
    ReDim Preserve Smoke(NumSmoke)
    
    Smoke(NumSmoke).X = X
    Smoke(NumSmoke).Y = Y
    Smoke(NumSmoke).Direction = 1
    Smoke(NumSmoke).Size = 10 '0.4
    
    Smoke(NumSmoke).Speed = Speed
    Smoke(NumSmoke).Heading = Heading
    
    Smoke(NumSmoke).bLongTime = bLongTime
    
    
    NumSmoke = NumSmoke + 1
'End If

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
    
    StickMotion Casing(i).X, Casing(i).Y, Casing(i).Speed, Casing(i).Heading
    
    If CasingInPlatform(i) Then
        ReverseYComp Casing(i).Speed, Casing(i).Heading
        Casing(i).Speed = Casing(i).Speed / Casing_Bounce_Reduction
    ElseIf CasingOnEdge(i) Then
        ReverseXComp Casing(i).Speed, Casing(i).Heading
        Casing(i).Speed = Casing(i).Speed / Casing_Bounce_Reduction
    ElseIf Casing(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        AddVectors Casing(i).Speed, Casing(i).Heading, Gravity_Strength, Gravity_Direction, Casing(i).Speed, Casing(i).Heading
        Casing(i).LastGravity = GetTickCount()
    End If
    
Next i


End Sub

Private Function CasingOnEdge(iCasing As Integer) As Boolean

If Casing(iCasing).X < Lim Then
    CasingOnEdge = True
ElseIf Casing(iCasing).X > StickGameWidth - Lim Then
    CasingOnEdge = True
End If

End Function

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

Private Function CasingInPlatform(iCasing As Integer) As Boolean
Dim j As Integer

For j = 0 To nPlatforms
    If CasingCollision(iCasing, Platform(j).Left, Platform(j).Top, Platform(j).width, Platform(j).height) Then
        CasingInPlatform = True
        Exit For
    End If
Next j

End Function

Private Sub ProcessNades()
Dim i As Integer, j As Integer
Dim RemoveIt As Boolean, bWall As Boolean

Do While i < NumNades
    
    bWall = False
    RemoveIt = False
    
    If Nade(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
        If Nade(i).IsRPG Then
            AddVectors Nade(i).Speed, Nade(i).Heading, Gravity_Strength / 2, Gravity_Direction, Nade(i).Speed, Nade(i).Heading
        Else
            AddVectors Nade(i).Speed, Nade(i).Heading, Gravity_Strength, Gravity_Direction, Nade(i).Speed, Nade(i).Heading
        End If
        
        Nade(i).LastGravity = GetTickCount()
    End If
    
    StickMotion Nade(i).X, Nade(i).Y, Nade(i).Speed, Nade(i).Heading
    
    
    
    If modStickGame.sv_ShootNades Then
        If Nade(i).iType = nFrag Then
            If Nade(i).Decay + (600 - Nade_Time) / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                If NadeInBullet(i) Then
                    RemoveIt = True
                    GoTo DoRemove
                End If
            End If
        End If
    End If
    
    
    If NadeInBox(i) Then
        RemoveIt = True
        'bWall =false
        
    ElseIf NadeInPlatform(i) Then
        'RemoveIt = True
        If Nade(i).IsRPG Then
            RemoveIt = True
            bWall = True
        ElseIf Nade(i).iType <> nSmoke Then
            Nade(i).Speed = Nade(i).Speed / Nade_Bounce_Reduction
            ReverseYComp Nade(i).Speed, Nade(i).Heading
        Else
            RemoveIt = True
            bWall = True
        End If
        
    ElseIf NadeOnEdge(i) Then
        
        If Nade(i).IsRPG Then
            RemoveIt = True
            bWall = True
        ElseIf Nade(i).iType <> nSmoke Then
            Nade(i).Speed = Nade(i).Speed / Nade_Bounce_Reduction
            
            If Nade(i).X > (StickGameWidth - Lim - 5) Then
                Nade(i).X = StickGameWidth - Lim - 10
                
                If Nade(i).Speed > 0 Or Nade(i).Heading < pi Then
                    ReverseXComp Nade(i).Speed, Nade(i).Heading
                End If
                
            ElseIf Nade(i).X < Lim Then
                Nade(i).X = Lim + 5
                
                If Nade(i).Speed < 0 Or Nade(i).Heading > pi Then
                    ReverseXComp Nade(i).Speed, Nade(i).Heading
                End If
            Else
                ReverseXComp Nade(i).Speed, Nade(i).Heading
            End If
            
        Else
            RemoveIt = True
            bWall = True
        End If
        
    ElseIf NadePastCeiling(i) Then
        
        If Nade(i).IsRPG Then
            RemoveIt = True
            bWall = True
        Else
            ReverseYComp Nade(i).Speed, Nade(i).Heading
        End If
        
    Else
        
        j = NadeInTBox(i) 'j = tBox that nade is in
        
        If j > -1 Then
            If Nade(i).IsRPG Then
                RemoveIt = True
                'bWall=false
            ElseIf Nade(i).iType <> nSmoke Then
'                If Nade(i).Y > (tBox(j).Top + tBox(j).height / 2) Then
'                    'probably hit the side
'                    ReverseXComp Nade(i).Speed, Nade(i).Heading
'                Else
'                    ReverseYComp Nade(i).Speed, Nade(i).Heading
'                End If
                Nade(i).Heading = Nade(i).Heading - pi
            Else
                RemoveIt = True
                'bWall=False
            End If
        End If
        
    End If
    
DoRemove:
    For j = 0 To NumSticksM1
        If StickInGame(j) Then
            If Nade(i).OwnerID <> Stick(j).ID Then
                If NadeInStick(i, j) Then
                    'ReverseXComp Nade(i).Speed, Nade(i).Heading
                    RemoveIt = True
                    'bWall=false
                    
                    'or explode the nade?
                    Exit For
                End If
            End If
        End If
    Next j
    
    
    
    If Nade(i).Decay < GetTickCount() Then
        RemoveIt = True
        bWall = NadeInPlatform(i)
        If bWall = False Then
            bWall = NadeInTBox(i)
        End If
        'main bit for WallMarks above ^^
    End If
    
    If RemoveIt Then
        ExplodeNade i
        RemoveNade i, bWall
        i = i - 1
    End If
    
    i = i + 1
Loop


End Sub

Private Sub ExplodeNade(ByVal i As Integer)
Dim j As Integer

AddSmokeTrail Nade(i).X, Nade(i).Y, True
For j = 0 To 6
    AddSparks Nade(i).X, Nade(i).Y, CSng(j)
Next j



If Nade(i).iType = nFrag Then
    ExplodeFrag i
ElseIf Nade(i).iType = nFlash Then
    ExplodeFlash i
Else
    ExplodeSmoke i
End If

End Sub

Private Sub ExplodeFrag(i As Integer)
Dim j As Integer

Dim Dist As Single
Dim OwnerIndex As Integer
Dim MaxDist As Single
Dim ExplosionForceDist As Single

Const Nade_Explode_RadiusX2 = Nade_Explode_Radius * 2
Const ChopperLenX1p2 = ChopperLen * 1.2


AddExplosion Nade(i).X, Nade(i).Y, 500, 1, 0, 0
For j = 1 To 10
    AddSmokeGroup Nade(i).X, Nade(i).Y, 5, 100 * Rnd(), 2 * pi * Rnd()
Next j


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
        
        
        If Dist < ExplosionForceDist Then
            If Stick(j).WeaponType <> Chopper Then
                AddVectors Stick(j).Speed, Stick(j).Heading, _
                    NadeMultiple * (Dist + 1), FindAngle(Nade(i).X, Nade(i).Y, Stick(j).X, Stick(j).Y), _
                    Stick(j).Speed, Stick(j).Heading
            End If
        End If
        
        
        If Dist < MaxDist Then
            
            Stick(j).OnSurface = False
            Stick(j).Y = Stick(j).Y - 100
            
            'Exit For
            If Stick(j).ID = MyID Or Stick(j).IsBot Then
                OwnerIndex = FindStick(Nade(i).OwnerID)
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
                                    
                                    
                                    'fixed damage of 51 (2 RPGs to kill)
                                    DamageStick Chopper_Damage_Reduction * 51, j
                                    
                                Else
                                    DamageStick Chopper_Damage_Reduction * 30, j 'nade = 30 damage
                                End If
                            Else
                                DamageStick 100000 / Dist, j 'bullet
                            End If
                            
                            If Err.Number <> 0 Then 'div zero error
                                Stick(j).Health = 0
                                Err.Clear
                            End If
                            
                            If Stick(j).Health < 1 Then
                                Call Killed(j, FindStick(Nade(i).OwnerID), IIf(Nade(i).IsRPG, kRPG, kNade))
                            End If
                            
                        End If 'spawn invul endif
                    End If 'ally endif
                End If 'owner index endif
            End If 'myid endif
        ElseIf Stick(j).WeaponType = Chopper Then
            If Dist < 2870 Then
                If j = 0 Or Stick(j).IsBot Then
                    'tail rotor
                    DamageStick Chopper_Damage_Reduction * 250000 / Dist, j
                    
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

For j = 0 To 20
    AddSparks Nade(i).X + PM_Rnd() * SparkLim, _
              Nade(i).Y + PM_Rnd * SparkLim, CSng(j)
    
Next j


If PointOnSticksScreen(Nade(i).X, Nade(i).Y, 0) And StickInvul(0) = False Then
    BangFlash i
Else
    AddExplosion Nade(i).X, Nade(i).Y, 500, 1, 0, 0
End If

For j = 0 To NumSticks - 1
    'If Stick(j).IsBot Then
    If Stick(j).WeaponType <> Chopper Then
        If StickInvul(j) = False Then
            If PointOnSticksScreen(Nade(i).X, Nade(i).Y, j) Then
                Stick(j).LastFlashBang = GetTickCount()
            End If
        End If
    End If
Next j

For j = 0 To 3
    Smoke_Speed = 120 + 20 * Rnd()
    AddSmokeGroup Nade(i).X, Nade(i).Y, 3, Smoke_Speed, piD2 + PM_Rnd() / Angle_Redux, True
    AddSmokeGroup Nade(i).X, Nade(i).Y, 3, Smoke_Speed, pi3D2 + PM_Rnd() / Angle_Redux, True
Next j

AddSmokeGroup Nade(i).X, Nade(i).Y, 3, Smoke_Speed, 0, True


End Sub

Private Sub ExplodeSmoke(i As Integer)
Dim j As Integer
Dim Smoke_Speed As Single

AddSparks Nade(i).X, Nade(i).Y, Nade(i).Heading - pi
AddExplosion Nade(i).X, Nade(i).Y, 500, 1, 0, 0

AddLargeSmoke Nade(i).X, Nade(i).Y, Nade(i).Heading

End Sub

Private Function PointOnSticksScreen(X As Single, Y As Single, i As Integer) As Boolean
Const XLimit = 6000, YLimit = 5500

If X > (Stick(i).X - XLimit) Then
    If X < (Stick(i).X + XLimit) Then
        If Y > (Stick(i).Y - YLimit) Then
            If Y < (Stick(i).Y + YLimit) Then
                PointOnSticksScreen = True
            End If
        End If
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

Private Function NadePastCeiling(iNade As Integer) As Boolean

If Nade(iNade).Y < Lim Then
    NadePastCeiling = True
End If

End Function

Private Sub DrawNades()
Dim i As Integer
Dim tY As Single, tX As Single
Dim TimeLeft As Single

'Me.ForeColor = vbBlack
For i = 0 To NumNades - 1
    If Nade(i).IsRPG Then
        
        picMain.DrawWidth = 1
        picMain.FillStyle = vbFSTransparent
        
        DrawRocket Nade(i).X, Nade(i).Y, Nade(i).Heading ', Nade(i).Colour
        'picMain.DrawWidth = 2
        
        If Nade(i).LastSmoke + RPG_Smoke_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            
            tX = Nade(i).X + GunLen * Sin(Nade(i).Heading - pi)
            tY = Nade(i).Y - GunLen * Cos(Nade(i).Heading - pi)
            
            AddSmokeGroup tX, tY, 3, 0, 0
            'AddSmokeTrail tX, tY
            
            If modStickGame.cg_RPGFlame Then
                AddExplosion tX, tY, 400, 0.15, 0, 0
            End If
            
            Nade(i).LastSmoke = GetTickCount()
            
        End If
        
    Else
        picMain.DrawWidth = 2
        picMain.FillStyle = vbSolid
        
        DrawNade Nade(i).X, Nade(i).Y, Nade(i).Colour, Nade(i).iType
        
        TimeLeft = (Nade(i).Decay - GetTickCount()) * modStickGame.sv_StickGameSpeed
        
        tX = Nade(i).X - Nade_Time / 4
        tY = Nade(i).Y - 650
        
        modStickGame.sLine tX, tY, tX + Nade_Time / 2, tY, vbBlue
        modStickGame.sLine tX, tY, tX + Nade_Time / 2 - TimeLeft / 2, tY, vbRed
        
        'modStickGame.PrintStickText "On Screen: " & _
            PointOnScreen(Nade(i).X, Nade(i).Y), tX, tY, vbBlack
    End If
    
Next i

picMain.FillStyle = vbFSTransparent

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
Dim Pt(1 To 7) As POINTAPI

Pt(1).X = pX
Pt(1).Y = pY

Pt(2).X = Pt(1).X + GunLen / 1.5 * Sin(pHeading - pi8D9)
Pt(2).Y = Pt(1).Y - GunLen / 1.5 * Cos(pHeading - pi8D9)

Pt(3).X = Pt(2).X + GunLen / 2.5 * Sin(pHeading + pi8D9)
Pt(3).Y = Pt(2).Y - GunLen / 2.5 * Cos(pHeading + pi8D9)

Pt(4).X = Pt(3).X + GunLen / 3 * Sin(pHeading - pi)
Pt(4).Y = Pt(3).Y - GunLen / 3 * Cos(pHeading - pi)


Pt(7).X = Pt(1).X + GunLen / 1.5 * Sin(pHeading + pi8D9)
Pt(7).Y = Pt(1).Y - GunLen / 1.5 * Cos(pHeading + pi8D9)

Pt(6).X = Pt(7).X + GunLen / 2.5 * Sin(pHeading - pi8D9)
Pt(6).Y = Pt(7).Y - GunLen / 2.5 * Cos(pHeading - pi8D9)

Pt(5).X = Pt(6).X + GunLen / 3 * Sin(pHeading - pi)
Pt(5).Y = Pt(6).Y - GunLen / 3 * Cos(pHeading - pi)

picMain.ForeColor = vbBlack
'picMain.fillstyle = vbFSTransparent
'picMain.DrawWidth = 2
modStickGame.sPoly Pt, -1 'pCol


'Dim x(1 To 7) As Single, y(1 To 7) As Single
'x(1) = pX
'y(1) = pY
'
'x(2) = x(1) + GunLen / 1.5 * Sin(pHeading - pi8D9)
'y(2) = y(1) - GunLen / 1.5 * Cos(pHeading - pi8D9)
'
'x(3) = x(2) + GunLen / 2.5 * Sin(pHeading + pi8D9)
'y(3) = y(2) - GunLen / 2.5 * Cos(pHeading + pi8D9)
'
'x(4) = x(3) + GunLen / 3 * Sin(pHeading - pi)
'y(4) = y(3) - GunLen / 3 * Cos(pHeading - pi)
'
'
'x(7) = x(1) + GunLen / 1.5 * Sin(pHeading + pi8D9)
'y(7) = y(1) - GunLen / 1.5 * Cos(pHeading + pi8D9)
'
'x(6) = x(7) + GunLen / 2.5 * Sin(pHeading - pi8D9)
'y(6) = y(7) - GunLen / 2.5 * Cos(pHeading - pi8D9)
'
'x(5) = x(6) + GunLen / 3 * Sin(pHeading - pi)
'y(5) = y(6) - GunLen / 3 * Cos(pHeading - pi)
'
'
'picMain.DrawWidth = 1
'Me.ForeColor = vbBlack
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
    
    If MineInBullet(i) Then
        RemoveIt = True
    Else
        For j = 0 To NumSticksM1
            If StickInGame(j) Then
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
        ExplodeMine i
        RemoveMine i
        i = i - 1
    ElseIf Mine(i).bOnSurface = False Then
            
            If Mine(i).X < 1 Then
                If Mine(i).Heading > pi Then
                    ReverseXComp Mine(i).Speed, Mine(i).Heading
                End If
            ElseIf Mine(i).X > (StickGameWidth - 1) Then
                If Mine(i).Heading < pi Then
                    ReverseXComp Mine(i).Speed, Mine(i).Heading
                End If
            End If
            
            If Mine(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                AddVectors Mine(i).Speed, Mine(i).Heading, Gravity_Strength, Gravity_Direction, _
                    Mine(i).Speed, Mine(i).Heading
                
                Mine(i).LastGravity = GetTickCount()
            End If
            
            
            StickMotion Mine(i).X, Mine(i).Y, Mine(i).Speed, Mine(i).Heading
            For j = 0 To nPlatforms
                MineOnSurface i, j
            Next j
            
            
    End If
    
    i = i + 1
Loop


End Sub

Private Sub MineOnSurface(i As Integer, iPlatform As Integer) 'As Boolean

If Mine(i).X > Platform(iPlatform).Left Then
    If Mine(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
        
        If Mine(i).Y > Platform(iPlatform).Top Then
            If Mine(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
                
                Mine(i).Y = Platform(iPlatform).Top
                
                'mineOnSurface = True
                Mine(i).bOnSurface = True
                Mine(i).Speed = 0
                
            End If
        End If
        
        
    End If
End If

End Sub

Private Sub ExplodeMine(ByVal i As Integer)
Dim j As Integer, OwnerIndex As Integer
Dim Dist As Single

AddExplosion Mine(i).X, Mine(i).Y, 500, 2, 0, 0
AddSmokeTrail Mine(i).X, Mine(i).Y, True
For j = 1 To 10
    AddSmokeGroup Mine(i).X, Mine(i).Y, 5, 75 * Rnd(), 2 * pi * Rnd()
Next j

For j = 0 To NumSticksM1
    
    If StickInGame(j) Then
        Dist = GetDist(Stick(j).X, Stick(j).Y, Mine(i).X, Mine(i).Y)
        
        If Dist < Mine_Explode_Radius Then
            
            If Stick(j).ID = MyID Or Stick(j).IsBot Then
                OwnerIndex = FindStick(Mine(i).OwnerID)
                If OwnerIndex <> -1 Then
                    If (IsAlly(Stick(j).Team, Stick(OwnerIndex).Team) = False) Or (j = OwnerIndex) Then
                        'If Stick(j).LastSpawnTime + Spawn_Invul_Time < GetTickCount() Then
                        If StickInvul(j) = False Then
                            
                            On Error Resume Next
                            'Stick(j).Helth = Stick(j).Health - 100000 / Dist
                            DamageStick 180000 / Dist, j
                            
                            If Err.Number <> 0 Then 'div zero error
                                Stick(j).Health = 0
                                Err.Clear
                            End If
                            
                            If Stick(j).Health < 1 Then
                                Call Killed(j, FindStick(Mine(i).OwnerID), kMine)
                            End If
                            
                        End If 'spawn invul endif
                    End If 'ally endif
                End If 'owner index endif
            End If 'myid endif
            
            
        End If 'dist endif
    End If 'stickingame endif
Next j

End Sub

Private Sub DrawMines()
Dim i As Integer
Dim tY As Single, tX As Single
Dim TimeLeft As Single

picMain.DrawWidth = 2

For i = 0 To NumMines - 1
    DrawMine Mine(i).X, Mine(i).Y, Mine(i).Colour
Next i

End Sub

Private Sub DrawMine(X As Single, Y As Single, Colour As Long)
Const kX = 50, kY = 12
Const Mine_RadiusD2 = Mine_Radius / 2

modStickGame.sBoxFilled X - kX, Y - kY, X + kX, Y + kY, Colour
modStickGame.sCircle X, Y - Mine_Radius, Mine_Radius, BoxCol

End Sub

Private Function MineNearStick(iMine As Integer, iStick As Integer) As Boolean
Const StickLimX = 1500, StickLimY = 2000

If Mine(iMine).X > (Stick(iStick).X - StickLimX) Then
    If Mine(iMine).X < (Stick(iStick).X + StickLimX) Then
        
        If Mine(iMine).Y > (Stick(iStick).Y - StickLimY) Then
            If Mine(iMine).Y < (Stick(iStick).Y + StickLimY) Then
                MineNearStick = True
            End If
        End If
        
    End If
End If

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

If Bullet(Bulleti).LastDiffract = 0 Or Bullet(Bulleti).bSniperBullet Then
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

Private Sub ProcessSmoke()
Dim i As Integer
Dim F As Single

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
        
        StickMotion .X, .Y, .Speed, .Heading
        
        modStickGame.sCircle .X, .Y, .Size, SmokeOutline
        
        
        F = modStickGame.StickTimeFactor / IIf(Smoke(i).bLongTime, 2.5, 1)
        
        If .Direction = 1 Then
            .Size = .Size + 2 * F
        Else
            .Size = .Size - 0.5 * F
        End If
        
    End With
    
Next i

picMain.FillStyle = vbFSTransparent 'transparent

End Sub

Private Sub ProcessMuzzleFlashes()
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
picMain.DrawWidth = 1.5
For i = 0 To NumSticksM1
    If Stick(i).LastMuzzleFlash + MFlash_Time / modStickGame.sv_StickGameSpeed > GetTickCount() Then
        DrawMFlash CSng(Stick(i).GunPoint.X), CSng(Stick(i).GunPoint.Y), Stick(i).Facing
    End If
Next i

End Sub

Private Sub DrawMFlash(X As Single, Y As Single, Facing As Single)

Const SideLen = 5, FrontLen = 110
Const SideFlashLen = 65 ', FrontFlashLen = 5
Dim Pts(1 To 3) As POINTAPI
Dim Rd As Single, RD2 As Single

Rd = Rnd()
RD2 = Rnd()

Pts(1).X = X + SideLen * Sin(Facing - piD2) * Rd
Pts(1).Y = Y - SideLen * Cos(Facing - piD2) * Rd

Pts(2).X = X + FrontLen * Sin(Facing) * Rd
Pts(2).Y = Y - FrontLen * Cos(Facing) * Rd

'Pts(3).x = x + FrontLen * Sin(Facing + piD10)
'Pts(3).y = y - FrontLen * Cos(Facing + piD10)

Pts(3).X = X + SideLen * Sin(Facing + piD2) * Rd
Pts(3).Y = Y - SideLen * Cos(Facing + piD2) * Rd

modStickGame.sPoly Pts, vbYellow


modStickGame.sLine X, Y, _
    X + SideFlashLen * Sin(Facing - piD3) * RD2, _
    Y - SideFlashLen * Cos(Facing - piD3) * RD2, vbYellow

modStickGame.sLine X, Y, _
    X + SideFlashLen * Sin(Facing + piD3) * RD2, _
    Y - SideFlashLen * Cos(Facing + piD3) * RD2, vbYellow


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


picMain.DrawWidth = 1
Me.picMain.ForeColor = vbBlack
For i = 0 To NumMags - 1
    
    If Mag(i).bOnSurface = False Then
        
        If Mag(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            AddVectors Mag(i).Speed, Mag(i).Heading, Gravity_Strength, Gravity_Direction, _
                Mag(i).Speed, Mag(i).Heading
            
            Mag(i).LastGravity = GetTickCount()
        End If
        
        
        StickMotion Mag(i).X, Mag(i).Y, Mag(i).Speed, Mag(i).Heading
        For j = 0 To nPlatforms
            MagOnSurface i, j
        Next j
        
    End If
    
    DrawMagazine i
Next i

End Sub

Private Sub DrawMagazine(i As Integer)
Dim Pt(1 To 4) As POINTAPI
Dim j As Integer

If Mag(i).iMagType = mAK Then
    'top left
    Pt(1).X = Mag(i).X
    Pt(1).Y = Mag(i).Y - 50
    
    'top right
    Pt(2).X = Mag(i).X + 50
    Pt(2).Y = Pt(1).Y
    
    'bottom left
    Pt(4).X = Mag(i).X + 10
    Pt(4).Y = Mag(i).Y + 70
    
    'bottom right
    Pt(3).X = Mag(i).X + 60
    Pt(3).Y = Mag(i).Y + 70
    
    picMain.DrawWidth = 2
    modStickGame.sPoly Pt, -1
    
ElseIf Mag(i).iMagType = mSCAR Then
    'top left
    Pt(1).X = Mag(i).X
    Pt(1).Y = Mag(i).Y - 50
    
    'top right
    Pt(2).X = Mag(i).X + 50
    Pt(2).Y = Pt(1).Y
    
    'bottom left
    Pt(4).X = Mag(i).X + 10
    Pt(4).Y = Mag(i).Y + 70
    
    'bottom right
    Pt(3).X = Mag(i).X + 60
    Pt(3).Y = Mag(i).Y + 70
    
    picMain.DrawWidth = 1
    modStickGame.sPoly Pt, vbBlack
    
ElseIf Mag(i).iMagType = mSniper Then
    'top left
    Pt(1).X = Mag(i).X
    Pt(1).Y = Mag(i).Y
    
    'top right
    Pt(2).X = Mag(i).X + 50
    Pt(2).Y = Pt(1).Y
    
    'bottom left
    Pt(4).X = Mag(i).X + 10
    Pt(4).Y = Mag(i).Y + 50
    
    'bottom right
    Pt(3).X = Mag(i).X + 75
    Pt(3).Y = Mag(i).Y + 50
    
    picMain.DrawWidth = 1
    modStickGame.sPoly Pt, -1
    
ElseIf Mag(i).iMagType = mPistol Then
    'top left
    Pt(1).X = Mag(i).X
    Pt(1).Y = Mag(i).Y - 25
    
    'top right
    Pt(2).X = Mag(i).X + 25
    Pt(2).Y = Pt(1).Y
    
    'bottom left
    Pt(4).X = Mag(i).X + 5
    Pt(4).Y = Mag(i).Y + 60
    
    'bottom right
    Pt(3).X = Mag(i).X + 30
    Pt(3).Y = Mag(i).Y + 60
    picMain.DrawWidth = 1
    modStickGame.sPoly Pt, vbBlack
    
ElseIf Mag(i).iMagType = mFlameThrower Then
    Pt(1).X = Mag(i).X
    Pt(1).Y = Mag(i).Y
    
    Pt(2).X = Pt(1).X + GunLen / 2
    Pt(2).Y = Pt(1).Y
    
    Pt(3).X = Pt(2).X
    Pt(3).Y = Pt(2).Y + GunLen / 3
    
    Pt(4).X = Pt(3).X - GunLen / 4
    Pt(4).Y = Pt(3).Y
    
    picMain.DrawWidth = 2
    modStickGame.sPoly Pt, vbRed
    
ElseIf Mag(i).iMagType = mSA80 Then
    'top left
    Pt(1).X = Mag(i).X
    Pt(1).Y = Mag(i).Y - 50
    
    'top right
    Pt(2).X = Mag(i).X + 25
    Pt(2).Y = Pt(1).Y
    
    'bottom left
    Pt(4).X = Mag(i).X + 10
    Pt(4).Y = Mag(i).Y + 70
    
    'bottom right
    Pt(3).X = Mag(i).X + 30
    Pt(3).Y = Mag(i).Y + 70
    
    picMain.DrawWidth = 1
    modStickGame.sPoly Pt, vbBlack
End If


End Sub

Private Sub MagOnSurface(i As Integer, iPlatform As Integer) 'As Boolean

If Mag(i).X > Platform(iPlatform).Left Then
    If Mag(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
        
        
        If Mag(i).Y > Platform(iPlatform).Top - 10 Then
            If Mag(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
                
                'position the Mag on top of the platform
                Mag(i).Y = Platform(iPlatform).Top - 80
                'If Mag(i).Y > (Platform(iPlatform).Top + 30) Then
                'End If
                
                Mag(i).bOnSurface = True
                Mag(i).Speed = 0
                
            End If
        End If
        
        
    End If
End If

End Sub

Private Sub ProcessStaticWeapons()
Dim i As Integer, j As Integer
Dim bPrompted As Boolean
Dim bHave(0 To eWeaponTypes.Knife - 1) As Boolean

On Error GoTo EH

Do While i < NumStaticWeapons
    If StaticWeapon(i).Y > StickGameHeight Then
        RemoveStaticWeapon i
        i = i - 1
    Else
        
        'check if sticks are near any
        For j = 0 To NumSticksM1
            If StickInGame(j) Then
                If Stick(j).WeaponType < Knife Then
                    If StickiHasState(j, Stick_Use) And StickiHasState(j, Stick_Reload) = False Then
                        If StickNearStaticWeapon(j, i) Then
                            
                            If Stick(j).LastWeaponSwitch + SwitchWeaponDelay < GetTickCount() Then
                                
                                If StickiHasWeapon(j, StaticWeapon(i).iWeapon) = False Then
                                    
                                    'pickup the weapon
                                    If j = 0 Then
                                        'decide whether we are swapping currentweapon(1) or (2)
                                        If Stick(j).CurrentWeapons(1) = Stick(0).WeaponType Then
                                            'we are swapping CW(1)
                                            Stick(j).CurrentWeapons(1) = StaticWeapon(i).iWeapon
                                        Else
                                            Stick(j).CurrentWeapons(2) = StaticWeapon(i).iWeapon
                                        End If
                                        
                                        SwitchWeapon StaticWeapon(i).iWeapon
                                        
                                        On Error Resume Next
                                        AmmoFired(StaticWeapon(i).iWeapon) = 0
                                        Stick(j).BulletsFired = 0
                                        
                                        UseKey = False
                                    Else
                                        Stick(j).WeaponType = StaticWeapon(i).iWeapon
                                    End If
                                    
                                    
                                    'drop current weapon
                                    AddStaticWeapon Stick(j).X, Stick(j).Y, Stick(j).PrevWeapon
                                    
                                    'SubStickState Stick(j).ID, Stick_Use
                                    Stick(j).LastWeaponSwitch = GetTickCount()
                                    
                                    
                                    RemoveStaticWeapon i
                                    
                                    i = i - 1
                                    Exit For
                                    
                                Else
                                    
                                    'stick has the weapon, sub the state
                                    If j = 0 Then
                                        UseKey = False
                                    End If
                                    
                                    SubStickiState j, Stick_Use
                                    
                                End If
                            'Else 'recently swapped
                                'If j = 0 Then
                                    'UseKey = False
                                'End If
                                'SubStickState Stick(j).ID, Stick_Use
                            End If
                        'Else 'not near weapon
                            'If j = 0 Then
                                'UseKey = False
                            'End If
                            'SubStickState Stick(j).ID, Stick_Use
                        End If
                    ElseIf StickNearStaticWeapon(j, i) Then
                        'doesn't have use state...
                        
                        If j = 0 Then
                            If bPrompted = False Then
                                If StickiHasWeapon(j, StaticWeapon(i).iWeapon) = False Then
                                    If StickiHasState(j, Stick_Reload) = False Then
                                        PrintStickText "Press E to pick up " & GetWeaponName(StaticWeapon(i).iWeapon), _
                                            Stick(0).X - 1000, Stick(0).Y - 1000, vbBlack
                                    End If
                                    
                                    bPrompted = True
                                End If
                            End If
                        End If
                        
                        'prevent from 'using' after reload
                        If StickiHasState(j, Stick_Reload) Then
                            SubStickiState j, Stick_Use
                            If j = 0 Then UseKey = False
                        End If
                        
                    End If
                End If
            End If
        Next j
    End If
    
    
    i = i + 1
Loop

'PrintStickText "Use: " & UseKey, Stick(0).X, Stick(0).Y, vbRed

picMain.DrawWidth = 1
Me.picMain.ForeColor = vbBlack

For i = 0 To NumStaticWeapons - 1
    
    If StaticWeapon(i).bOnSurface = False Then
        
        If StaticWeapon(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            AddVectors StaticWeapon(i).Speed, StaticWeapon(i).Heading, Gravity_Strength, Gravity_Direction, _
                StaticWeapon(i).Speed, StaticWeapon(i).Heading
            
            StaticWeapon(i).LastGravity = GetTickCount()
        End If
        
        
        StickMotion StaticWeapon(i).X, StaticWeapon(i).Y, StaticWeapon(i).Speed, StaticWeapon(i).Heading
        
        For j = 0 To nPlatforms
            StaticWeaponOnSurface i, j
        Next j
        
    End If
    
    
    If modStickGame.sv_AllowFlameThrowers = False Then
        If StaticWeapon(i).iWeapon = FlameThrower Then
            StaticWeapon(i).iWeapon = GetRandomStaticWeapon()
        End If
    End If
    If modStickGame.sv_AllowRockets = False Then
        If StaticWeapon(i).iWeapon = RPG Then
            StaticWeapon(i).iWeapon = GetRandomStaticWeapon()
        End If
    End If
    
    
    If StaticWeapon(i).iWeapon = Knife Then
        StaticWeapon(i).iWeapon = GetRandomStaticWeapon()
    End If
    
Next i


If modStickGame.StickServer Then
    'check we have them all
    For i = 0 To NumStaticWeapons - 1
        bHave(StaticWeapon(i).iWeapon) = True
    Next i
    
    For i = 0 To eWeaponTypes.Knife - 1
        If bHave(i) = False Then
            
            bPrompted = True
            
            If i = eWeaponTypes.RPG Then
                If modStickGame.sv_AllowRockets = False Then
                    bPrompted = False
                End If
            ElseIf i = eWeaponTypes.FlameThrower Then
                If modStickGame.sv_AllowFlameThrowers = False Then
                    bPrompted = False
                End If
            End If
            
            If bPrompted Then
                'find a weapon, and make it be this weapon type
                For j = 0 To NumStaticWeapons - 1
                    If Rnd() > 0.5 Then 'StaticWeapon(j).iWeapon < i Then
                        
                        StaticWeapon(j).iWeapon = i
                        
                        
                        Exit For
                    End If
                Next j
            End If
            
        End If
    Next i
End If


EH:
End Sub

Private Sub DrawStaticWeapons()
Dim i As Integer

For i = 0 To NumStaticWeapons - 1
    DrawStaticWeapon i
Next i

End Sub

Private Sub StaticWeaponOnSurface(i As Integer, iPlatform As Integer) 'As Boolean

If StaticWeapon(i).X > Platform(iPlatform).Left Then
    If StaticWeapon(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
        
        
        If StaticWeapon(i).Y > Platform(iPlatform).Top - 10 Then
            If StaticWeapon(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
                
                'position the StaticWeapon on top of the platform
                StaticWeapon(i).Y = Platform(iPlatform).Top - 80
                
                StaticWeapon(i).bOnSurface = True
                StaticWeapon(i).Speed = 0
                AddSparks StaticWeapon(i).X, StaticWeapon(i).Y, StaticWeapon(i).Heading - pi
                
            End If
        End If
        
        
    End If
End If

End Sub

Private Sub DrawStaticWeapon(i As Integer)

If modStickGame.cg_SimpleStaticWeapons Then
    
    modStickGame.sCircle StaticWeapon(i).X, StaticWeapon(i).Y, 100, vbBlack
    modStickGame.PrintStickText GetWeaponName(StaticWeapon(i).iWeapon), StaticWeapon(i).X, StaticWeapon(i).Y - 500, vbBlack
    
Else
    
    Me.DrawWidth = 1
    
    If StaticWeapon(i).iWeapon = SCAR Then
        DrawStaticSCAR StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = AK Then
        DrawStaticAK StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = DEagle Then
        DrawStaticDEagle StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = FlameThrower Then
        DrawStaticFlameThrower StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = M249 Then
        DrawStaticM249 StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = M82 Then
        DrawStaticM82 StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = RPG Then
        DrawStaticRPG StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = Shotgun Then
        DrawStaticShotgun StaticWeapon(i).X, StaticWeapon(i).Y
        
    ElseIf StaticWeapon(i).iWeapon = SA80 Then
        DrawStaticSA80 StaticWeapon(i).X, StaticWeapon(i).Y
        
    End If
    
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

Public Sub MakeStaticWeapons()
Dim i As Single

For i = 0 To eWeaponTypes.Knife - 0.25 Step 0.25
    AddStaticWeapon Rnd() * StickGameWidth, _
                    Rnd() * StickGameHeight, _
                    CInt(i)
    
Next i


For i = 0 To eWeaponTypes.Knife - 1
    AddStaticWeapon Rnd() * StickGameWidth, _
                    Rnd() * StickGameHeight, _
                    GetRandomStaticWeapon()
Next i

End Sub

Public Function GetRandomStaticWeapon() As eWeaponTypes
'any up to knife, not including knife
Dim vWep As eWeaponTypes

vWep = CInt(Rnd() * eWeaponTypes.Knife)

If vWep = Knife Then
    vWep = SA80
End If


If modStickGame.sv_AllowFlameThrowers = False Then
    If vWep = FlameThrower Then
        vWep = AK
    End If
End If
If modStickGame.sv_AllowRockets = False Then
    If vWep = RPG Then
        vWep = AK
    End If
End If


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

Private Function StickiHasWeapon(iStick As Integer, vWeapon As eWeaponTypes) As Boolean

If Stick(iStick).CurrentWeapons(1) = vWeapon Then
    StickiHasWeapon = True
ElseIf Stick(iStick).CurrentWeapons(2) = vWeapon Then
    StickiHasWeapon = True
End If

End Function


'STATIC WEAPON DRAWING
'#########################################################################################################
Private Sub DrawStaticShotgun(sX As Single, sY As Single)

Dim X(1 To 11) As Single, Y(1 To 11) As Single
Const Facing As Single = piD2
Const SAd2 = SmallAngle / 2

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 2 * Sin(Facing - SmallAngle)
Y(2) = Y(1) - GunLen / 2 * Cos(Facing - SmallAngle)

X(3) = X(1) + GunLen / 1.5 * Sin(Facing - SmallAngle)
Y(3) = Y(1) - GunLen / 1.5 * Cos(Facing - SmallAngle)

X(4) = X(1) + GunLen / 1.5 * Sin(Facing - SAd2)
Y(4) = Y(1) - GunLen / 1.5 * Cos(Facing - SAd2)

X(5) = X(1) + GunLen * Sin(Facing - SAd2)
Y(5) = Y(1) - GunLen * Cos(Facing - SAd2)

'pump action bit
X(6) = X(1) + GunLen * Sin(Facing - SAd2)
Y(6) = Y(1) - GunLen * Cos(Facing - SAd2)

X(7) = X(1) + GunLen * 1.5 * Sin(Facing - SmallAngle / 3)
Y(7) = Y(1) - GunLen * 1.5 * Cos(Facing - SmallAngle / 3)
'end pump action bit

X(8) = X(1) + GunLen * 2 * Sin(Facing - SmallAngle / 3)
Y(8) = Y(1) - GunLen * 2 * Cos(Facing - SmallAngle / 3)

X(9) = X(1) + GunLen * 2.5 * Sin(Facing - SmallAngle / 3.5)
Y(9) = Y(1) - GunLen * 2.5 * Cos(Facing - SmallAngle / 3.5)

X(10) = X(9) + GunLen / 6 * Sin(Facing - pi2d3)
Y(10) = Y(9) - GunLen / 6 * Cos(Facing - pi2d3)

X(11) = X(9) + GunLen / 20 * Sin(Facing - pi)
Y(11) = Y(9) - GunLen / 20 * Cos(Facing - pi)

'end calculation

Me.ForeColor = &H555555
picMain.DrawWidth = 2

'handle section
modStickGame.sLine X(1), Y(1), X(3), Y(3), vbRed

picMain.DrawWidth = 2
modStickGame.sLine X(2), Y(2), X(4), Y(4), vbBlack

modStickGame.sLine X(2), Y(2), X(8), Y(8), &H555555
modStickGame.sLine X(3), Y(3), X(9), Y(9), &H555555

Me.ForeColor = vbRed
modStickGame.sLine X(1), Y(1), X(4), Y(4), vbRed
modStickGame.sLine X(6), Y(6), X(7), Y(7), vbRed

'Me.ForeColor = &H555555
picMain.DrawWidth = 1
modStickGame.sLine X(10), Y(10), X(11), Y(11), vbRed

End Sub

Private Sub DrawStaticAK(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 18) As Single, Y(1 To 18) As Single

Const SAd2 = SmallAngle / 2
Const SAd4 = SmallAngle / 4
Const SAd8 = SmallAngle / 8

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 4 * Sin(Facing + 11 * pi / 18)
Y(2) = Y(1) - GunLen / 4 * Cos(Facing + 11 * pi / 18) '90+20deg

X(3) = X(1) + GunLen / 4 * Sin(Facing + piD2)
Y(3) = Y(1) - GunLen / 4 * Cos(Facing + piD2)

X(4) = X(1) + GunLen / 20 * Sin(Facing)
Y(4) = Y(1) - GunLen / 20 * Cos(Facing)

X(5) = X(1) + GunLen / 4 * Sin(Facing)
Y(5) = Y(1) - GunLen / 4 * Cos(Facing)

X(6) = X(1) + GunLen / 3.2 * Sin(Facing - SAd2)
Y(6) = Y(1) - GunLen / 3.2 * Cos(Facing - SAd2)

X(7) = X(6) + GunLen / 1.5 * Sin(Facing + piD4)
Y(7) = Y(6) - GunLen / 1.5 * Cos(Facing + piD4)

X(8) = X(7) + GunLen / 4 * Sin(Facing - piD4)
Y(8) = Y(7) - GunLen / 4 * Cos(Facing - piD4)

X(9) = X(1) + GunLen / 2 * Sin(Facing - SAd2)
Y(9) = Y(1) - GunLen / 2 * Cos(Facing - SAd2)

X(10) = X(9) + GunLen * Sin(Facing - SAd8)
Y(10) = Y(9) - GunLen * Cos(Facing - SAd8)

X(11) = X(10) + GunLen / 4 * Sin(Facing - piD2)
Y(11) = Y(10) - GunLen / 4 * Cos(Facing - piD2)

X(12) = X(11) + GunLen / 4 * Sin(Facing + (piD2 + SmallAngle))
Y(12) = Y(11) - GunLen / 4 * Cos(Facing + (piD2 + SmallAngle))

X(13) = X(12) + GunLen / 3 * Sin(Facing - pi)
Y(13) = Y(12) - GunLen / 3 * Cos(Facing - pi)

X(14) = X(13) + GunLen / 3 * Sin(Facing - pi)
Y(14) = Y(13) - GunLen / 3 * Cos(Facing - pi)

X(15) = X(14) + GunLen * 0.6 * Sin(Facing + (pi - SAd4))
Y(15) = Y(14) - GunLen * 0.6 * Cos(Facing + (pi - SAd4))

X(16) = X(2) + GunLen / 2 * Sin(Facing - (pi + SAd4))
Y(16) = Y(2) - GunLen / 2 * Cos(Facing - (pi + SAd4))

X(17) = X(16) + GunLen / 4 * Sin(Facing + (piD2 - SAd4))
Y(17) = Y(16) - GunLen / 4 * Cos(Facing + (pi / 2 - SAd4))

X(18) = X(1) + GunLen / 8 * Sin(Facing - pi)
Y(18) = Y(1) - GunLen / 8 * Cos(Facing - pi)
'end calculation

'drawing
picMain.DrawWidth = 2
Me.ForeColor = &H6AD5
'handle
modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(3), Y(3), X(2), Y(2)
modStickGame.sLine X(3), Y(3), X(4), Y(4)

Me.ForeColor = vbBlack
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
Me.ForeColor = &H6AD5
modStickGame.sLine X(9), Y(9), X(10), Y(10)
Me.ForeColor = vbBlack
modStickGame.sLine X(11), Y(11), X(10), Y(10) 'iron sight
modStickGame.sLine X(11), Y(11), X(12), Y(12) 'iron sight
Me.ForeColor = &H6AD5
modStickGame.sLine X(13), Y(13), X(12), Y(12)
modStickGame.sLine X(13), Y(13), X(14), Y(14)
Me.ForeColor = vbBlack
modStickGame.sLine X(15), Y(15), X(14), Y(14)

'stock
Me.ForeColor = &H6AD5
modStickGame.sLine X(15), Y(15), X(16), Y(16)
modStickGame.sLine X(17), Y(17), X(16), Y(16)
modStickGame.sLine X(17), Y(17), X(18), Y(18)
Me.ForeColor = vbBlack
modStickGame.sLine X(18), Y(18), X(1), Y(1)

End Sub

Private Sub DrawStaticSCAR(sX As Single, sY As Single)

Const Facing As Single = piD2
Const tSin = 0.92387, tCos = 0.38268  'Cos(Facing - piD8)
Dim X(1 To 24) As Single, Y(1 To 24) As Single

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 3 * Sin(Facing + pi3D4)
Y(2) = Y(1) - GunLen / 3 * Cos(Facing + pi3D4)

X(3) = X(2) + GunLen / 6 '* SinFacing
Y(3) = Y(2) '- GunLen / 6 * CosFacing

X(4) = X(1) + GunLen / 6 '* SinFacing
Y(4) = Y(1) '- GunLen / 6 * CosFacing

X(5) = X(4) + GunLen / 6 '* SinFacing
Y(5) = Y(4) '- GunLen / 6 * CosFacing

X(6) = X(5) + GunLen / 3 * Sin(Facing + pi4D9)
Y(6) = Y(5) - GunLen / 3 * Cos(Facing + pi4D9)

X(7) = X(6) + GunLen / 4 * tSin
Y(7) = Y(6) - GunLen / 4 * tCos

X(8) = X(5) + GunLen / 4 * tSin
Y(8) = Y(5) - GunLen / 4 * tCos

X(9) = X(8) + GunLen / 5 * tSin
Y(9) = Y(8) - GunLen / 5 * tCos

'straight bottom part of barrel
X(10) = X(9) + GunLen / 1.5 '* SinFacing
Y(10) = Y(9) '- GunLen / 1.5 * CosFacing

'wedge
X(11) = X(10) + GunLen / 2.8 * Sin(Facing - pi3D4)
Y(11) = Y(10) - GunLen / 2.8 * Cos(Facing - pi3D4)


X(12) = X(11) + GunLen / 1.4 * Sin(Facing - pi)
Y(12) = Y(11) - GunLen / 1.4 * Cos(Facing - pi)

X(13) = X(12) + GunLen / 6 * Sin(Facing - piD2)
Y(13) = Y(12) - GunLen / 6 * Cos(Facing - piD2)

X(14) = X(13) + GunLen / 3 * Sin(Facing - pi)
Y(14) = Y(13) - GunLen / 3 * Cos(Facing - pi)

X(15) = X(14) + GunLen / 6 * Sin(Facing + piD2)
Y(15) = Y(14) - GunLen / 6 * Cos(Facing + piD2)

X(16) = X(15) + GunLen / 15 * Sin(Facing + piD2)
Y(16) = Y(15) - GunLen / 15 * Cos(Facing + piD2)

'top buttstock
X(17) = X(15) + GunLen / 2 * Sin(Facing - (pi * 1.1))
Y(17) = Y(15) - GunLen / 2 * Cos(Facing - (pi * 1.1))

'bottom buttstock
X(18) = X(17) + GunLen / 2 * Sin(Facing + piD2)
Y(18) = Y(17) - GunLen / 2 * Cos(Facing + piD2)

X(19) = X(18) + GunLen / 4 * Sin(Facing - piD2)
Y(19) = Y(18) - GunLen / 4 * Cos(Facing - piD2)


''start of fancy bits
'X(20) = X(9) + GunLen / 6 * tSin 'F-piD8
'Y(20) = Y(9) - GunLen / 6 * tCos
'
'X(21) = X(20) + GunLen / 2 * SinFacing
'Y(21) = Y(20) - GunLen / 2 * CosFacing
'
'X(22) = X(20) + GunLen / 6 * Sin(Facing - piD2)
'Y(22) = Y(20) - GunLen / 6 * Cos(Facing - piD2)
'
'X(23) = X(22) + GunLen / 3 * SinFacing
'Y(23) = Y(22) - GunLen / 3 * CosFacing


'#############
'Hole in front of scope
X(20) = X(1) + GunLen / 1.6 * Sin(Facing - piD6)
Y(20) = Y(1) - GunLen / 1.6 * Cos(Facing - piD6)

X(21) = X(20) + GunLen / 1.8 * Sin(Facing - 0.07)
Y(21) = Y(20) - GunLen / 1.8 * Cos(Facing - 0.07) 'pi/40

'X(22) = X(21) + GunLen / 5.2 * Sin(Facing - piD4)
'Y(22) = Y(21) - GunLen / 5.2 * Cos(Facing - piD4)

X(22) = X(16) + GunLen / 4 '* SinFacing
Y(22) = Y(16) '- GunLen / 4 * CosFacing

'#############
'barrel
X(23) = X(10) + GunLen / 6 * Sin(Facing - pi3D4)
Y(23) = Y(10) - GunLen / 6 * Cos(Facing - pi3D4)

X(24) = X(23) + GunLen / 4 '* SinFacing 'GunLen/x = BarrelLen
Y(24) = Y(23) '- GunLen / 4 * CosFacing

'X(26) = X(11) + GunLen / 6 * Sin(Facing + pi3D4)
'Y(26) = Y(11) - GunLen / 6 * Cos(Facing + pi3D4)
'
'X(25) = X(26) + GunLen / 2 * SinFacing
'Y(25) = Y(26) - GunLen / 2 * CosFacing

'end calculation


'drawing

picMain.DrawWidth = 1
Me.ForeColor = &H2F2F2F


modStickGame.sLine X(1), Y(1), X(2), Y(2)
modStickGame.sLine X(2), Y(2), X(3), Y(3)
modStickGame.sLine X(3), Y(3), X(4), Y(4)
modStickGame.sLine X(4), Y(4), X(5), Y(5)

'magazine
modStickGame.sLine X(5), Y(5), X(6), Y(6)
modStickGame.sLine X(6), Y(6), X(7), Y(7)
modStickGame.sLine X(7), Y(7), X(8), Y(8)

'mag modstickgame.sLine
modStickGame.sLine X(5), Y(5), X(8), Y(8)

modStickGame.sLine X(8), Y(8), X(9), Y(9)
modStickGame.sLine X(9), Y(9), X(10), Y(10)
modStickGame.sLine X(10), Y(10), X(11), Y(11)
modStickGame.sLine X(11), Y(11), X(12), Y(12)
modStickGame.sLine X(12), Y(12), X(13), Y(13)
modStickGame.sLine X(13), Y(13), X(14), Y(14)
modStickGame.sLine X(14), Y(14), X(15), Y(15)
modStickGame.sLine X(15), Y(15), X(16), Y(16)
modStickGame.sLine X(16), Y(16), X(17), Y(17)
modStickGame.sLine X(17), Y(17), X(18), Y(18)
modStickGame.sLine X(18), Y(18), X(19), Y(19)

'hole bit
modStickGame.sLine X(20), Y(20), X(21), Y(21)
'modstickgame.sLine X(22), Y(22),X(21), Y(21)
modStickGame.sLine X(20), Y(20), X(22), Y(22)

'scope modstickgame.sLine
modStickGame.sLine X(16), Y(16), X(22), Y(22)

'connect stock to handle
modStickGame.sLine X(1), Y(1), X(18), Y(18)

'barrel
picMain.DrawWidth = 1
modStickGame.sLine X(23), Y(23), X(24), Y(24)

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


SinFacingLess_kYpiD2 = Sin(Facing - piD2)
CosFacingLess_kYpiD2 = Cos(Facing - piD2)
SinFacingLess_kYpiD4 = Sin(Facing - piD4)
CosFacingLess_kYpiD4 = Cos(Facing - piD4)
'SinFacing = Sin(Facing)
'CosFacing = Cos(Facing)

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 4 * Sin(Facing - piD4)
Y(2) = Y(1) - GunLen / 4 * Cos(Facing - piD4)

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

X(14) = X(12) + GunLen / 8 * Sin(Facing - pi)
Y(14) = Y(12) - GunLen / 8 * Cos(Facing - pi)

X(15) = X(14) + GLd10 * SinFacingLess_kYpiD2
Y(15) = Y(14) - GLd10 * CosFacingLess_kYpiD2

X(16) = X(15) + GunLen / 10 * Sin(Facing - pi) 'iron sight bottom
Y(16) = Y(15) - GunLen / 10 * Cos(Facing - pi)

X(17) = X(16) + GunLen / 10 * SinFacingLess_kYpiD2 'iron sight top
Y(17) = Y(16) - GunLen / 10 * CosFacingLess_kYpiD2

X(18) = X(15) + GunLen / 6 * Sin(Facing - pi)
Y(18) = Y(15) - GunLen / 6 * Cos(Facing - pi)

X(19) = X(18) + GunLen / 2 * Sin(Facing - pi) 'end of straight top bit
Y(19) = Y(18) - GunLen / 2 * Cos(Facing - pi)

X(20) = X(1) + GunLen / 4 * SinFacingLess_kYpiD2
Y(20) = Y(1) - GunLen / 4 * CosFacingLess_kYpiD2

'sight stand
'bottom points
X(21) = X(18) + GunLen / 8 * Sin(Facing - pi) 'forward bottom
Y(21) = Y(18) - GunLen / 8 * Cos(Facing - pi)

X(22) = X(21) + GunLen / 4 * Sin(Facing - pi) 'rearward bottom
Y(22) = Y(21) - GunLen / 4 * Cos(Facing - pi)
'top points
X(23) = X(21) + GunLen / 6 * SinFacingLess_kYpiD2 'forward top
Y(23) = Y(21) - GunLen / 6 * CosFacingLess_kYpiD2

X(24) = X(22) + GunLen / 6 * SinFacingLess_kYpiD2 'rearward top
Y(24) = Y(22) - GunLen / 6 * CosFacingLess_kYpiD2
'modstickgame.sLine from 21->23, 22->24

'scope
X(25) = X(24) + GunLen / 4 * Sin(Facing - pi) 'rear bottom pt
Y(25) = Y(24) - GunLen / 4 * Cos(Facing - pi)

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
    
    X(31) = X(30) + GunLen / 2 * Sin(Facing + pi / 1.8) 'GunLen/x = Height of Stand
    Y(31) = Y(30) - GunLen / 2 * Cos(Facing + pi / 1.8)
    
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
Me.ForeColor = &H3F3F3F
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
Me.ForeColor = vbBlack
modStickGame.sLine X(12), Y(12), X(13), Y(13) 'BARREL

'modStickGame.sLine X(13), Y(13), X(14), Y(14)
'modStickGame.sLine X(14), Y(14), X(15), Y(15)
modStickGame.sLine X(12), Y(12), X(15), Y(15)

picMain.DrawWidth = 1
Me.ForeColor = &H693F3F
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
Me.ForeColor = vbBlack '&H555555
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

Private Sub DrawStaticRPG(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 16) As Single, Y(1 To 16) As Single

Const SAd2 = SmallAngle / 2

X(2) = sX
Y(2) = sY

X(1) = X(2) + GunLen / 2 * Sin(Facing - piD2)
Y(1) = Y(2) - GunLen / 2 * Cos(Facing - piD2)

X(3) = X(1) + GunLen / 1.5 * Sin(Facing)
Y(3) = Y(1) - GunLen / 1.5 * Cos(Facing)

X(4) = X(3) + GunLen / 2 * Sin(Facing + piD2)
Y(4) = Y(3) - GunLen / 2 * Cos(Facing + piD2)

X(5) = X(3) + GunLen / 1.5 * Sin(Facing)
Y(5) = Y(3) - GunLen / 1.5 * Cos(Facing)

X(6) = X(5) + GunLen / 4 * Sin(Facing - piD2)
Y(6) = Y(5) - GunLen / 4 * Cos(Facing - piD2)

X(7) = X(6) + GunLen * 3 * Sin(Facing - pi) 'rear top point
Y(7) = Y(6) - GunLen * 3 * Cos(Facing - pi)

X(8) = X(1) + GunLen * 1.7 * Sin(Facing - pi) 'rear bottom point
Y(8) = Y(1) - GunLen * 1.7 * Cos(Facing - pi)

'rear funnel
X(9) = X(7) + GunLen / 3 * Sin(Facing - pi3D4) 'rear top point
Y(9) = Y(7) - GunLen / 3 * Cos(Facing - pi3D4)

X(10) = X(8) + GunLen / 3 * Sin(Facing + pi3D4) 'rear bottom point
Y(10) = Y(8) - GunLen / 3 * Cos(Facing + pi3D4)

'sights
X(11) = X(6) + GunLen / 1.2 * Sin(Facing - pi)
Y(11) = Y(6) - GunLen / 1.2 * Cos(Facing - pi)

X(12) = X(11) + GunLen / 4 * Sin(Facing - piD2)
Y(12) = Y(11) - GunLen / 4 * Cos(Facing - piD2)

X(13) = X(12) + GunLen / 4 * Sin(Facing - piD4)
Y(13) = Y(12) - GunLen / 4 * Cos(Facing - piD4)

X(14) = X(13) + GunLen / 4 * Sin(Facing - piD2)
Y(14) = Y(13) - GunLen / 4 * Cos(Facing - piD2)

X(15) = X(14) + GunLen / 2 * Sin(Facing + pi3D4)
Y(15) = Y(14) - GunLen / 2 * Cos(Facing + pi3D4)

X(16) = X(15) + GunLen / 4 * Sin(Facing + piD2)
Y(16) = Y(15) - GunLen / 4 * Cos(Facing + piD2)
'end calculation

'drawing
Me.ForeColor = vbBlack
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

DrawRocket X(5) + GunLen / 1.2 * Sin(Facing - piD20), _
            Y(5) - GunLen / 1.2 * Cos(Facing - piD20), _
            Facing ', Stick(i).Colour


End Sub

Private Sub DrawStaticM249(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 20) As Single, Y(1 To 20) As Single

X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 2 * Sin(Facing + pi3D4)
Y(2) = Y(1) - GunLen / 2 * Cos(Facing + pi3D4)

X(3) = X(2) + GunLen / 4 * Sin(Facing)
Y(3) = Y(2) - GunLen / 4 * Cos(Facing)

X(4) = X(1) + GunLen / 4 * Sin(Facing)
Y(4) = Y(1) - GunLen / 4 * Cos(Facing)
'end handle

'gap between handle and handy bit
X(5) = X(4) + GunLen / 4 * Sin(Facing)
Y(5) = Y(4) - GunLen / 4 * Cos(Facing)

X(6) = X(5) + GunLen / 6 * Sin(Facing + piD2)
Y(6) = Y(5) - GunLen / 6 * Cos(Facing + piD2)

X(7) = X(6) + GunLen / 2 * Sin(Facing)
Y(7) = Y(6) - GunLen / 2 * Cos(Facing)

X(8) = X(5) + GunLen / 2 * Sin(Facing)
Y(8) = Y(5) - GunLen / 2 * Cos(Facing)

'bipod
X(9) = X(2) + GunLen * 1.2 * Sin(Facing + piD10)
Y(9) = Y(2) - GunLen * 1.2 * Cos(Facing + piD10)

X(10) = X(2) + GunLen * 1.5 * Sin(Facing + piD10)
Y(10) = Y(2) - GunLen * 1.5 * Cos(Facing + piD10)

'barrel
X(11) = X(8) + GunLen / 1.5 * Sin(Facing)
Y(11) = Y(8) - GunLen / 1.5 * Cos(Facing)

'sights
X(12) = X(8) + GunLen / 4 * Sin(Facing)
Y(12) = Y(8) - GunLen / 4 * Cos(Facing)

X(13) = X(12) + GunLen / 4 * Sin(Facing - piD2)
Y(13) = Y(12) - GunLen / 4 * Cos(Facing - piD2)

'top bit
X(14) = X(8) + GunLen / 10 * Sin(Facing - piD2)
Y(14) = Y(8) - GunLen / 10 * Cos(Facing - piD2)

'top handle
X(15) = X(14) + GunLen / 4 * Sin(Facing - pi)
Y(15) = Y(14) - GunLen / 4 * Cos(Facing - pi)

X(16) = X(15) + GunLen / 6 * Sin(Facing - piD2)
Y(16) = Y(15) - GunLen / 6 * Cos(Facing - piD2)

X(17) = X(16) + GunLen / 4 * Sin(Facing - pi3D4)
Y(17) = Y(16) - GunLen / 4 * Cos(Facing - pi3D4)
'end handle

X(18) = X(15) + GunLen / 4 * Sin(Facing - pi)
Y(18) = Y(15) - GunLen / 4 * Cos(Facing - pi)

X(18) = X(15) + GunLen / 4 * Sin(Facing - pi)
Y(18) = Y(15) - GunLen / 4 * Cos(Facing - pi)

X(19) = X(1) + GunLen / 2 * Sin(Facing - pi)
Y(19) = Y(1) - GunLen / 2 * Cos(Facing - pi)

X(20) = X(19) + GunLen / 4 * Sin(Facing + piD2)
Y(20) = Y(19) - GunLen / 4 * Cos(Facing + piD2)
'end calculation

Me.ForeColor = vbBlack
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

Private Sub DrawStaticDEagle(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim X(1 To 10) As Single, Y(1 To 10) As Single
Const HeadRadius2 = HeadRadius * 2 ', DEagle_Bullet_DelayD2 = DEagle_Bullet_Delay / 2


X(1) = sX
Y(1) = sY

X(2) = X(1) + GunLen / 2 * Sin(Facing)
Y(2) = Y(1) - GunLen / 2 * Cos(Facing)

X(3) = X(2) + GunLen / 6 * Sin(Facing - piD3) '60 deg
Y(3) = Y(2) - GunLen / 6 * Cos(Facing - piD3)

X(4) = X(3) + GunLen / 12 * Sin(Facing - piD2)
Y(4) = Y(3) - GunLen / 12 * Cos(Facing - piD2)

X(5) = X(3) + GunLen / 10 * Sin(Facing - pi)
Y(5) = Y(3) - GunLen / 10 * Cos(Facing - pi)

X(6) = X(3) + GunLen / 1.6 * Sin(Facing - pi)
Y(6) = Y(3) - GunLen / 1.6 * Cos(Facing - pi)

X(6) = X(3) + GunLen / 1.6 * Sin(Facing - pi)
Y(6) = Y(3) - GunLen / 1.6 * Cos(Facing - pi)

X(7) = X(6) + GunLen / 4 * Sin(Facing + pi8D9)
Y(7) = Y(6) - GunLen / 4 * Cos(Facing + pi8D9)

X(8) = X(1) + GunLen / 6 * Sin(Facing - pi)
Y(8) = Y(1) - GunLen / 6 * Cos(Facing - pi)

X(9) = X(8) + GunLen / 3 * Sin(Facing + pi13D18)
Y(9) = Y(8) - GunLen / 3 * Cos(Facing + pi13D18)

X(10) = X(9) + GunLen / 6 * Sin(Facing)
Y(10) = Y(9) - GunLen / 6 * Cos(Facing)

'end calculation
modStickGame.sLine X(1), Y(1), X(2), Y(2), MSilver
modStickGame.sLine X(2), Y(2), X(3), Y(3), MSilver
modStickGame.sLine X(3), Y(3), X(4), Y(4), vbBlack
modStickGame.sLine X(4), Y(4), X(5), Y(5), vbBlack
modStickGame.sLine X(5), Y(5), X(6), Y(6), MSilver
modStickGame.sLine X(6), Y(6), X(7), Y(7), MSilver 'vbYellow
modStickGame.sLine X(7), Y(7), X(8), Y(8), vbBlack
modStickGame.sLine X(8), Y(8), X(9), Y(9), vbBlack
modStickGame.sLine X(9), Y(9), X(10), Y(10), vbBlack

modStickGame.sLine X(10), Y(10), X(1), Y(1), vbBlack

picMain.DrawWidth = 1

End Sub

Private Sub DrawStaticFlameThrower(sX As Single, sY As Single)

Const Facing As Single = piD2
Dim MB(1 To 10) As POINTAPI
Dim FB(1 To 4) As POINTAPI
'mb = MainBarrel
'fb = FuelBox

Const ArmLenDX = ArmLen / 3
Const BodyLenD2 = BodyLen / 2
Const BodyLenX2 = BodyLen * 2

MB(1).X = sX
MB(1).Y = sY

MB(2).X = MB(1).X + GunLen / 5 * Sin(Facing)
MB(2).Y = MB(1).Y - GunLen / 5 * Cos(Facing)

MB(3).X = MB(2).X + GunLen / 3 * Sin(Facing - piD4)
MB(3).Y = MB(2).Y - GunLen / 3 * Cos(Facing - piD4)

MB(4).X = MB(3).X + GunLen * Sin(Facing)
MB(4).Y = MB(3).Y - GunLen * Cos(Facing)

MB(5).X = MB(4).X + GunLen / 6 * Sin(Facing - piD4)
MB(5).Y = MB(4).Y - GunLen / 6 * Cos(Facing - piD4)

MB(6).X = MB(5).X + GunLen / 3 * Sin(Facing - piD6)
MB(6).Y = MB(5).Y - GunLen / 3 * Cos(Facing - piD6)

MB(7).X = MB(6).X + GunLen / 10 * Sin(Facing - piD2)
MB(7).Y = MB(6).Y - GunLen / 10 * Cos(Facing - piD2)

MB(8).X = MB(7).X + GunLen / 3 * Sin(Facing - pi)
MB(8).Y = MB(7).Y - GunLen / 3 * Cos(Facing - pi)

MB(9).X = MB(8).X + GunLen / 3 * Sin(Facing + pi3D4)
MB(9).Y = MB(8).Y - GunLen / 3 * Cos(Facing + pi3D4)

MB(10).X = MB(9).X + GunLen * Sin(Facing - pi)
MB(10).Y = MB(9).Y - GunLen * Cos(Facing - pi)


FB(1).X = MB(3).X '+ GunLen / 4 * Sin(Facing)
FB(1).Y = MB(3).Y '- GunLen / 4 * Sin(Facing)

FB(2).X = MB(3).X + GunLen / 2 * Sin(Facing) 'glDx = boxlen
FB(2).Y = MB(3).Y - GunLen / 2 * Cos(Facing)

FB(3).X = FB(2).X + GunLen / 3 * Sin(Facing + piD2)  'glDx = boxheight
FB(3).Y = FB(2).Y - GunLen / 3 * Cos(Facing + piD2)

FB(4).X = FB(3).X + GunLen / 4 * Sin(Facing - pi)
FB(4).Y = FB(3).Y - GunLen / 4 * Cos(Facing - pi)


picMain.ForeColor = vbBlack
picMain.DrawWidth = 2

modStickGame.sPoly MB, -1


modStickGame.sPoly FB, vbRed

End Sub

Private Sub DrawStaticSA80(sX As Single, sY As Single)

Const Facing As Single = piD2
Const kGreen = 32768 '32768=rgb(0,128,0)

Dim pGrip(1 To 4) As POINTAPI
Dim ptBarrel(1 To 4) As POINTAPI
Dim ptMain(1 To 5) As POINTAPI
Dim PtMag(1 To 4) As POINTAPI
Dim ptSights(1 To 4) As POINTAPI
Dim Barrel1X As Single, Barrel1Y As Single, Barrel2X As Single, Barrel2Y As Single
Dim Stock1X As Single, Stock1Y As Single, Stock2X As Single, Stock2Y As Single

'grip
pGrip(1).X = sX
pGrip(1).Y = sY

pGrip(2).X = pGrip(1).X + GunLen / 3 * Sin(Facing + pi3D4)
pGrip(2).Y = pGrip(1).Y - GunLen / 3 * Cos(Facing + pi3D4)

pGrip(3).X = pGrip(2).X + GunLen / 4
pGrip(3).Y = pGrip(2).Y

pGrip(4).X = pGrip(1).X + GunLen / 4
pGrip(4).Y = pGrip(1).Y
'end grip

'green barrel part
ptBarrel(1).X = pGrip(4).X
ptBarrel(1).Y = pGrip(4).Y

ptBarrel(2).X = ptBarrel(1).X + GunLen * 2 / 3 'GL/x = Green Len
ptBarrel(2).Y = ptBarrel(1).Y

ptBarrel(3).X = ptBarrel(2).X + GunLen / 5 * Sin(Facing - pi2d3) '100deg
ptBarrel(3).Y = ptBarrel(2).Y - GunLen / 5 * Cos(Facing - pi2d3)

ptBarrel(4).X = ptBarrel(1).X + GunLen / 4 * Sin(Facing - piD2)
ptBarrel(4).Y = ptBarrel(1).Y - GunLen / 4 * Cos(Facing - piD2)
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

ptMain(3).X = ptMain(2).X + GunLen / 2 * Sin(Facing + piD2)
ptMain(3).Y = ptMain(2).Y - GunLen / 2 * Cos(Facing + piD2)

ptMain(4).X = pGrip(1).X - GunLen / 3
ptMain(4).Y = pGrip(1).Y

ptMain(5).X = ptBarrel(1).X
ptMain(5).Y = ptBarrel(1).Y

'magazine

PtMag(1).X = pGrip(1).X - GunLen / 2.5
PtMag(1).Y = pGrip(1).Y

PtMag(2).X = PtMag(1).X - GunLen / 6 'GL/x = Mag Width
PtMag(2).Y = PtMag(1).Y + GunLen / 6

PtMag(3).X = PtMag(2).X + GunLen / 2 * Sin(Facing + piD3)
PtMag(3).Y = PtMag(2).Y - GunLen / 2 * Cos(Facing + piD3)

PtMag(4).X = PtMag(1).X + GunLen / 2 * Sin(Facing + piD3)
PtMag(4).Y = PtMag(1).Y - GunLen / 2 * Cos(Facing + piD3)


'sights
'bottom right
ptSights(1).X = pGrip(1).X + GunLen / 3 * Sin(Facing - piD2)
ptSights(1).Y = pGrip(1).Y - GunLen / 3 * Cos(Facing - piD2)

'top right
ptSights(2).X = ptSights(1).X + GunLen / 6 * Sin(Facing - piD4)
ptSights(2).Y = ptSights(1).Y - GunLen / 6 * Cos(Facing - piD4)

'top left
ptSights(3).X = ptSights(2).X - GunLen / 2
ptSights(3).Y = ptSights(2).Y

'bottom left
ptSights(4).X = ptSights(1).X - GunLen / 4
ptSights(4).Y = ptSights(1).Y




'#############
Stock1X = CSng(ptMain(2).X)
Stock1Y = CSng(ptMain(2).Y)
Stock2X = CSng(ptMain(3).X)
Stock2Y = CSng(ptMain(3).Y)
'#############

picMain.DrawWidth = 1
picMain.DrawStyle = vbFSSolid
picMain.ForeColor = vbBlack

'sight stand
modStickGame.sLine CLng(ptSights(1).X), _
                CLng(ptSights(1).Y), _
                CLng(ptSights(1).X + GunLen / 6 * Sin(Facing + piD2)), _
                CLng(ptSights(1).Y - GunLen / 6 * Cos(Facing + piD2)), vbBlack
modStickGame.sLine CLng(ptSights(4).X), _
                CLng(ptSights(4).Y), _
                CLng(ptSights(4).X + GunLen / 6 * Sin(Facing + piD2)), _
                CLng(ptSights(4).Y - GunLen / 6 * Cos(Facing + piD2)), vbBlack




modStickGame.sPoly_NoOutline pGrip, kGreen
modStickGame.sPoly_NoOutline ptBarrel, kGreen
modStickGame.sPoly ptMain, vbBlack
modStickGame.sPoly ptSights, vbBlack
modStickGame.sPoly PtMag, vbBlack


picMain.DrawWidth = 2
'barrel
modStickGame.sLine Barrel1X, Barrel1Y, Barrel2X, Barrel2Y, vbBlack
picMain.DrawWidth = 3
'rear butt stock
modStickGame.sLine Stock1X, Stock1Y, Stock2X, Stock2Y, kGreen



End Sub
'END STATIC WEAPON DRAWING
'#########################################################################################################

Private Sub DrawDeadSticks()
Dim i As Integer, j As Integer

picMain.DrawWidth = 2

Do While i < NumDeadSticks
    
    If DeadStick(i).Decay < GetTickCount() Then
        RemoveDeadStick i
        i = i - 1
    End If
    
    i = i + 1
Loop


For i = 0 To NumDeadSticks - 1
    
    If DeadStick(i).bOnSurface = False Then
        
        If DeadStick(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            AddVectors DeadStick(i).Speed, DeadStick(i).Heading, Gravity_Strength, Gravity_Direction, _
                DeadStick(i).Speed, DeadStick(i).Heading
            
            DeadStick(i).LastGravity = GetTickCount()
        End If
        
        StickMotion DeadStick(i).X, DeadStick(i).Y, DeadStick(i).Speed, DeadStick(i).Heading
        
        For j = 0 To nPlatforms
            DeadStickOnSurface i, j
        Next j
        
    End If
    
    
    DrawDeadStick DeadStick(i).X, DeadStick(i).Y, IIf(DeadStick(i).bFlamed, vbBlack, DeadStick(i).Colour), IIf(DeadStick(i).bFacingRight, -1, 1)
Next i

End Sub

Private Sub DrawDeadStick(X As Single, Y As Single, Col As Long, kY As Single)
Const BodyLenD2 = BodyLen / 2
Const BodyLenPlus = BodyLen * 1.2
Dim YpHR As Single, XmHR As Single

YpHR = Y + HeadRadius
XmHR = X - kY * HeadRadius

modStickGame.sCircle X, Y, HeadRadius, Col

picMain.FillStyle = vbFSSolid
picMain.FillColor = vbRed
modStickGame.sCircleAspect XmHR, YpHR, BodyLenD2, vbRed, 0.2
picMain.FillStyle = vbFSTransparent

modStickGame.sLine XmHR, Y, X - kY * BodyLen, YpHR, Col
modStickGame.sLine XmHR, Y, X - kY * BodyLenD2, YpHR, Col
modStickGame.sLine X - kY * BodyLen, YpHR, X - kY * BodyLenPlus, YpHR, Col

End Sub

Private Sub DeadStickOnSurface(i As Integer, iPlatform As Integer) 'As Boolean
Const kAmount = HeadRadius * 1.6
Dim j As Integer

If DeadStick(i).X > Platform(iPlatform).Left Then
    If DeadStick(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
        
        If DeadStick(i).Y > Platform(iPlatform).Top Then
            If DeadStick(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
                
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
                        AddSmokeTrail DeadStick(i).X + Rnd() * ArmLen, DeadStick(i).Y + Rnd() * ArmLen, True
                    Next j
                Else
                    For j = 1 To 15
                        'splatter!
                        AddBlood DeadStick(i).X, DeadStick(i).Y, PM_Rnd * piD2, False
                    Next j
                End If
                
                
            End If
        End If
        
        
    End If
End If

End Sub

Private Sub DrawDeadChoppers()
Dim i As Integer, j As Integer

picMain.DrawWidth = 2

Do While i < NumDeadChoppers
    
    If DeadChopper(i).Decay < GetTickCount() Then
        RemoveDeadChopper i
        i = i - 1
    End If
    
    i = i + 1
Loop


For i = 0 To NumDeadChoppers - 1
    
    If DeadChopper(i).bOnSurface = False Then
        
        If DeadChopper(i).LastGravity + Gravity_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
            
            AddVectors DeadChopper(i).Speed, DeadChopper(i).Heading, Gravity_Strength, Gravity_Direction, _
                DeadChopper(i).Speed, DeadChopper(i).Heading
            
            DeadChopper(i).LastGravity = GetTickCount()
            
            If DeadChopper(i).LastSmoke + DeadChopper_Smoke_Delay / modStickGame.sv_StickGameSpeed < GetTickCount() Then
                AddSmokeTrail DeadChopper(i).X - CLD3 * Rnd(), DeadChopper(i).Y
                AddExplosion DeadChopper(i).X - CLD3 * Rnd(), DeadChopper(i).Y + CLD10, 750, 0.25, DeadChopper(i).Speed / 3, DeadChopper(i).Heading
            End If
        End If
        
        StickMotion DeadChopper(i).X, DeadChopper(i).Y, DeadChopper(i).Speed, DeadChopper(i).Heading
        
        For j = 0 To nPlatforms
            DeadChopperOnSurface i, j
        Next j
        
    End If
    
    
    DrawDeadChopper DeadChopper(i).X, DeadChopper(i).Y, DeadChopper(i).Colour
Next i

End Sub

Private Sub DrawDeadChopper(X As Single, Y As Single, Col As Long)
Dim Pt(1 To 11) As POINTAPI, ScreenPt(1 To 3) As POINTAPI
Dim t1X As Single, t1Y As Single, t2X As Single, t2Y As Single
Const Facing = piD2

Pt(1).X = X
Pt(1).Y = Y

Pt(2).X = Pt(1).X + CLD6 * Sin(Facing + piD6)
Pt(2).Y = Pt(1).Y - CLD6 * Cos(Facing + piD6)

Pt(3).X = Pt(2).X + CLD10 * Sin(Facing + piD3)
Pt(3).Y = Pt(2).Y - CLD10 * Cos(Facing + piD3)

Pt(4).X = Pt(3).X + CLD2 * Sin(Facing - pi)
Pt(4).Y = Pt(3).Y - CLD2 * Cos(Facing - pi)

Pt(5).X = Pt(4).X + CLD10 * Sin(Facing - pi3D4)
Pt(5).Y = Pt(4).Y - CLD10 * Cos(Facing - pi3D4)

Pt(6).X = Pt(5).X + CLD3 * Sin(Facing - pi)
Pt(6).Y = Pt(5).Y - CLD3 * Cos(Facing - pi)

Pt(7).X = Pt(6).X + CLD6 * Sin(Facing - pi3D4)
Pt(7).Y = Pt(6).Y - CLD6 * Cos(Facing - pi3D4)

Pt(8).X = Pt(7).X + CLD8 * Sin(Facing)
Pt(8).Y = Pt(7).Y - CLD8 * Cos(Facing)

Pt(9).X = Pt(8).X + CLD8 * Sin(Facing + piD4)
Pt(9).Y = Pt(8).Y - CLD8 * Cos(Facing + piD4)

Pt(10).X = Pt(9).X + CLD6 * Sin(Facing)
Pt(10).Y = Pt(9).Y - CLD6 * Cos(Facing)

Pt(11).X = Pt(1).X + CLD8 * Sin(Facing - pi)
Pt(11).Y = Pt(1).Y - CLD8 * Cos(Facing - pi)


ScreenPt(1).X = Pt(1).X + Sin(Facing + piD2) * 50
ScreenPt(1).Y = Pt(1).Y - Cos(Facing + piD2) * 50

ScreenPt(2).X = Pt(2).X + Sin(Facing - pi) * 50
ScreenPt(2).Y = Pt(2).Y - Cos(Facing - pi) * 50

ScreenPt(3).X = ScreenPt(2).X - CLD6 * Sin(Facing)
ScreenPt(3).Y = ScreenPt(2).Y + CLD6 * Cos(Facing)


t1X = CSng(Pt(3).X)
t1Y = CSng(Pt(3).Y)
t2X = CSng(Pt(8).X)
t2Y = CSng(Pt(8).Y)


picMain.DrawStyle = 5
modStickGame.sPoly Pt, cg_ChopperCol
modStickGame.sPoly ScreenPt, Col
picMain.DrawStyle = 0
picMain.DrawWidth = 2

modStickGame.sLine t1X, t1Y, t2X, t2Y, vbBlack

End Sub

Private Sub DeadChopperOnSurface(i As Integer, iPlatform As Integer) 'As Boolean
Const kAmount = CLD6, kAmountDX = kAmount / 1.2
Dim j As Integer

If DeadChopper(i).X > Platform(iPlatform).Left Then
    If DeadChopper(i).X < (Platform(iPlatform).Left + Platform(iPlatform).width) Then
        
        If DeadChopper(i).Y + kAmountDX > Platform(iPlatform).Top Then
            If DeadChopper(i).Y < (Platform(iPlatform).Top + Platform(iPlatform).height) Then
                
                'position the DeadChopper on top of the platform
                'If DeadChopper(i).y > (Platform(iPlatform).Top + 5) Then
                    '                       add on a bit
                DeadChopper(i).Y = Platform(iPlatform).Top - kAmount
                    
                'End If
                
                'DeadChopperOnSurface = True
                DeadChopper(i).bOnSurface = True
                DeadChopper(i).Speed = 0
                
                
                For j = 0 To 5
                    AddExplosion DeadChopper(i).X + CLD6 - Rnd() * CLD2, DeadChopper(i).Y + Rnd() * CLD6, 400, 0.25, 0, 0
                Next j
                
                For j = 0 To 20
                    AddSmokeTrail DeadChopper(i).X + CLD6 - Rnd() * CLD2, DeadChopper(i).Y + Rnd() * CLD6, True
                Next j
                
                Call CheckDeadChopperStickCollisions(i)
                
            End If
        End If
        
        
    End If
End If

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
                                    Stick(i).Armour = 0
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
    AddBlood X, Y, Rnd() * 2 * pi, False
Next i
End Sub

'explosions
Private Sub AddExplosion(ByVal X As Single, ByVal Y As Single, _
    ByVal TimeLen As Single, ByVal Radius As Single, _
    Speed As Single, Heading As Single)

If modStickGame.cg_Explosions Then
    AddCirc X, Y, TimeLen, Radius, vbYellow ', Speed, Heading
    AddCirc X, Y, TimeLen, Radius, 894704, 100 'Speed, Heading, 100
    AddCirc X, Y, TimeLen, Radius, vbRed, 200 'Speed, Heading, 200
    
    '894704 = orange
End If

End Sub

Private Sub AddCirc(ByVal X As Single, ByVal Y As Single, _
    ByVal MaxProg As Single, ByVal Radius As Single, _
    ByVal Colour As Long, _
    Optional ByVal StartProg As Single = 0) 'Speed As Single, Heading As Single, _

ReDim Preserve Circs(NumCircs)

Circs(NumCircs).X = X
Circs(NumCircs).Y = Y
Circs(NumCircs).MaxProg = MaxProg + StartProg
Circs(NumCircs).Radius = Radius
Circs(NumCircs).Colour = Colour
Circs(NumCircs).Prog = StartProg
Circs(NumCircs).Direction = 1
'Circs(NumCircs).Speed = Speed
'Circs(NumCircs).Heading = Heading

NumCircs = NumCircs + 1

End Sub

Private Sub RemoveCirc(ByVal Index As Integer)

Dim i As Integer

'If there's only one left, just erase the array

If NumCircs = 1 Then
    Erase Circs
    NumCircs = 0
Else
    'Remove the bullet
    For i = Index To NumCircs - 2
'        Circs(i).MaxProg = Circs(i + 1).MaxProg
'        Circs(i).Prog = Circs(i + 1).Prog
'        Circs(i).Radius = Circs(i + 1).Radius
'        Circs(i).X = Circs(i + 1).X
'        Circs(i).Y = Circs(i + 1).Y
'        Circs(i).Direction = Circs(i + 1).Direction
'        Circs(i).Colour = Circs(i + 1).Colour
        Circs(i) = Circs(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve Circs(NumCircs - 2)
    NumCircs = NumCircs - 1
End If

End Sub


Private Sub ProcessAllCircs()
Dim i As Integer
Dim bDraw As Boolean

picMain.FillStyle = vbFSSolid

i = NumCircs - 1
Do While i >= 0 '< NumCircs
    
    Circs(i).Prog = Circs(i).Prog + Circs(i).Direction * 100 * StickTimeFactor
    'Else
    '    Circs(i).Prog = Circs(i).Prog - 100 * modSpaceGame.sv_GameSpeed
    'End If
    
    'Motion Circs(i).x, Circs(i).y, Circs(i).Speed, Circs(i).Heading
    
    
    If Circs(i).Prog > Circs(i).MaxProg Then
        
        Circs(i).Direction = -1
        Circs(i).Prog = Circs(i).MaxProg
        
        bDraw = True
        
    ElseIf Circs(i).Prog <= 0 Then
        
        RemoveCirc i
        i = i - 1
        
        bDraw = False
    Else
        bDraw = True
    End If
    
    
    If bDraw Then
        picMain.FillColor = Circs(i).Colour
        modStickGame.sCircle Circs(i).X, Circs(i).Y, Circs(i).Radius * Circs(i).Prog, Circs(i).Colour
    End If
    
    
    i = i - 1
Loop

picMain.FillStyle = vbFSTransparent

End Sub

Public Sub SetCursor(bHide As Boolean)

If bHide Then
    Me.MousePointer = vbCustom
    Me.MouseIcon = picBlank.Picture
Else
    Me.MousePointer = vbDefault
End If

End Sub

Private Sub tmrMain_Timer()

Const Cap = "Stick Shooter - "

tmrMain.Enabled = False

'Connect winsock
If StartWinsock() Then
    
    
    'If we're not the StickServer, try to connect
    If Not StickServer Then
        If ConnectToServer() = False Then
            modWinsock.DestroySocket socket
            Unload Me
            Exit Sub
        Else
            Me.Caption = Cap & "Client"
        End If
    Else
        'socket already bound
        Me.Caption = Cap & "Host"
        
        'tell everyone
        SendInfoMessage frmMain.LastName & " Started a Game - Ctrl+G to Join"
'        If Server Then
'            DistributeMsg eCommands.Info & frmMain.LastName & " Started a Game - Ctrl+G to Join0", -1
'        Else
'            SendData eCommands.Info & frmMain.LastName & " Started a Game - Ctrl+G to Join0"
'        End If
        Pause 100
    End If
    
    
    If Not StickServer Then
        modWinsock.SendPacket socket, ServerSockAddr, sChats & Trim$(Stick(0).Name) & _
            " joined.#" & modVars.TxtForeGround
    End If
    
    'Me.MousePointer = vbCustom
    'Me.MouseIcon = picBlank.Picture
    SetCursor True
    
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

Public Sub MoveCameraX(ByVal nX As Single)
Const CameraLim = 150

modStickGame.cg_sCamera.X = nX

If StickInGame(0) = False Or bPlaying = False Then
    If modStickGame.cg_sCamera.X < -CameraLim Then
        modStickGame.cg_sCamera.X = -CameraLim
    ElseIf (modStickGame.cg_sCamera.X + Me.width) * cg_sZoom > (StickGameWidth + CameraLim) Then
        modStickGame.cg_sCamera.X = (StickGameWidth + CameraLim) / cg_sZoom - Me.width
    End If
End If

End Sub
Public Sub MoveCameraY(ByVal nY As Single)
Const CameraLim = 150

modStickGame.cg_sCamera.Y = nY

If StickInGame(0) = False Or bPlaying = False Then
    If modStickGame.cg_sCamera.Y < -CameraLim Then
        modStickGame.cg_sCamera.Y = -CameraLim
    ElseIf (modStickGame.cg_sCamera.Y + Me.height) * cg_sZoom > (StickGameHeight + CameraLim) Then
        modStickGame.cg_sCamera.Y = (StickGameHeight + CameraLim) / cg_sZoom - Me.height
    End If
End If

End Sub

Private Sub ResetKeys()
LeftKey = False
RightKey = False
JumpKey = False
CrouchKey = False
ProneKey = False
ReloadKey = False
FireKey = False
UseKey = False
MineKey = False

ShowScoresKey = False

SpecUp = False: SpecDown = False: SpecLeft = False: SpecRight = False

WeaponKey = -1
Scroll_WeaponKey = 0 '-1
LastScrollWeaponSwitch = 0
End Sub

'#####################################################################################
'Round Stuff
'#####################################################################################

Private Sub SendRoundInfo(Optional bForce As Boolean = False)
Static LastSend As Long


If LastSend + RoundInfoSendDelay < GetTickCount() Or bForce Then
    
    SendBroadcast sRoundInfos & CStr(Abs(bPlaying)) & CStr(RoundWinnerID)
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub ReceivedRoundInfo(sPacket As String)
Dim bWasPlaying As Boolean

On Error GoTo EH

bWasPlaying = bPlaying
bPlaying = CBool(Left$(sPacket, 1))
SetCursor bPlaying
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
SetCursor bPlaying
ResetKeys

If bStop Then
    If StickServer Then
        SendRoundInfo True
    End If
    
    RoundPausedAtThisTime = GetTickCount()
    For i = 0 To NumBullets - 1
        RemoveBullet 0, False
    Next i
    
    
    For i = 0 To NumSticksM1
        With Stick(i)
            .State = Stick_None
            '.Speed = 0
            .LastFlashBang = 0
        End With
    Next i
    
    For i = 0 To UBound(AmmoFired)
        AmmoFired(i) = 0
    Next i
    
    'clip camera
    modStickGame.cg_sZoom = 1
    MoveCameraX modStickGame.cg_sCamera.X
    MoveCameraY modStickGame.cg_sCamera.Y
    
Else
    'reset all scores
    For i = 0 To NumSticksM1
        With Stick(i)
            .iKills = 0
            .iDeaths = 0
            .State = Stick_None
            .iKillsInARow = 0
            .BulletsFired = 0
            
            .LastFlashBang = 0
            
            .LastNade = 0
            .LastMine = 0
            .LastBullet = 0
            
            .Speed = 0
            
            .Health = Health_Start
            .Armour = 0
            
            If .IsBot = False Then
                If .WeaponType = Chopper Then
                    .WeaponType = .CurrentWeapons(1)
                End If
            End If
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
                Else
                    MoveStickToCoOpStart i
                End If
            End If
        End With
    Next i
    
    If modStickGame.sv_GameType <> gCoOp Then
        RandomizeMyStickPos
    Else
        RadarStartTime = GetTickCount()
    End If
    
    
    'reset private stuff
    RadarStartTime = 0
    bHadRadar = False
    ChopperAvail = False
    FlamesInARow = 0
    KnifesInARow = 0
    
    
    'erase stuff
    StickGameSpeedChanged -1, -1
    Erase Nade: NumNades = 0
    Erase Mine: NumMines = 0
    Erase DeadChopper: NumDeadChoppers = 0
    Erase DeadStick: NumDeadSticks = 0
    Erase WallMark: NumWallMarks = 0
    NumLargeSmokes = 0: Erase LargeSmoke
    
    For i = 0 To modStickGame.nBoxes
        Box(i).bInUse = True
    Next i
    
    ResetKeys
    
    LastScoreCheck = GetTickCount() + 10000
    
    
    FillTotalMags
    
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
Stick(i).Y = StickGameHeight - 100 * Rnd()

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
BorderedBox 750, 900, 4200, 2200, BoxCol

RoundWinneri = FindStick(RoundWinnerID)
If RoundWinneri <> -1 Then
    
    Str = "Round Winner - " & Trim$(Stick(RoundWinneri).Name)
    
    PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 1000, Stick(RoundWinneri).Colour
    
    
    If (Stick(RoundWinneri).Team = Neutral Or Stick(RoundWinneri).Team = Spec) = False Then
        
        Str = "Winning Team - " & GetTeamStr(Stick(RoundWinneri).Team)
        
        PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 1400, GetTeamColour(Stick(RoundWinneri).Team)
        
    Else
        Str = "No Winning Team"
        
        PrintStickFormText Str, 2200 - TextWidth(Str) / 2, 1400, vbBlack
        
    End If
    '--------
End If

'Me.ForeColor = MGrey

RoundTm = Round((RoundPausedAtThisTime + RoundWaitTime - GetTickCount()) / 1000)

Str = "Round will begin in " & CStr(RoundTm) & " second" & IIf(RoundTm > 1, "s", vbNullString)
PrintStickFormText Str, 2000 - TextWidth(Str) / 2, 1800, vbBlack

picMain.Font.Size = 8


'decide if new round
If RoundTm <= 0 Then
    StopPlay False
End If


If LastPresenceSend + PresenceSendDelay < GetTickCount() Then
    
    If StickServer Then
        SendBroadcast sPresences & "0"
    Else
        modWinsock.SendPacket socket, ServerSockAddr, sPresences & CStr(MyID)
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

AddMainMessage "Game Type - " & GetGameType()

End Sub

'static weapons
'#################################################################################################
Private Function StaticWeaponToString(vType As ptStaticWeapon) As String

StaticWeaponToString = CStr( _
                    vType.Heading & mPacketSep & _
                    vType.iWeapon & mPacketSep & _
                    vType.Speed & mPacketSep & _
                    vType.X & mPacketSep & _
                    vType.Y & mPacketSep _
                    )

'don't send lastgravity or bOnSurface - different on all

End Function

Private Function StaticWeaponFromString(buf As String) As ptStaticWeapon

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With StaticWeaponFromString
    .bOnSurface = True
    .LastGravity = 1
    
    .Heading = CSng(Parts(0))
    .iWeapon = CInt(Parts(1))
    .Speed = CSng(Parts(2))
    .X = CSng(Parts(3))
    .Y = CSng(Parts(4))
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
                    Abs(vType.bAllowRockets) & mPacketSep & _
                    Abs(vType.bAllowFlameThrowers) & mPacketSep & _
                    Abs(vType.bShootNades) & mPacketSep & _
                    vType.sgGameSpeed & mPacketSep & _
                    Abs(vType.bHardCore) & mPacketSep & _
                    Abs(vType.bAllowChoppers) & mPacketSep & _
                    Abs(vType.bHPBonus) & mPacketSep & _
                    Abs(vType.b2Weapons) & mPacketSep & _
                    vType.iScoreToWin & mPacketSep & _
                    vType.iGameType & mPacketSep _
                    )

End Function

Private Function ServerVarFromString(buf As String) As ptServerVars

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With ServerVarFromString
    .bAllowRockets = CBool(Parts(0))
    .bAllowFlameThrowers = CBool(Parts(1))
    .bShootNades = CBool(Parts(2))
    .sgGameSpeed = CSng(Parts(3))
    .bHardCore = CBool(Parts(4))
    .bAllowChoppers = CBool(Parts(5))
    .bHPBonus = CBool(Parts(6))
    .b2Weapons = CBool(Parts(7))
    .iScoreToWin = CInt(Parts(8))
    .iGameType = CInt(Parts(9))
End With

EH:
Erase Parts
End Function

Private Sub ProcessServerVarPacket(vPacket As String)
Dim vServerVars As ptServerVars
Dim i As Integer

If IsValidVarPacket(vPacket) Then
    vServerVars = ServerVarFromString(Left$(vPacket, InStr(1, vPacket, vbSpace) - 1))
    
    modStickGame.sv_AllowRockets = vServerVars.bAllowRockets
    modStickGame.sv_ShootNades = vServerVars.bShootNades
    modStickGame.sv_AllowChoppers = vServerVars.bAllowChoppers
    modStickGame.sv_HPBonus = vServerVars.bHPBonus
    modStickGame.sv_AllowFlameThrowers = vServerVars.bAllowFlameThrowers
    modStickGame.sv_WinScore = vServerVars.iScoreToWin
    
    'check for a change
    If modStickGame.sv_StickGameSpeed <> vServerVars.sgGameSpeed Then
        If vServerVars.sgGameSpeed <= 1.2 Then
            If vServerVars.sgGameSpeed >= 0.1 Then
                StickGameSpeedChanged modStickGame.sv_StickGameSpeed, vServerVars.sgGameSpeed
                
                modStickGame.sv_StickGameSpeed = vServerVars.sgGameSpeed
            End If
        End If
    End If
    
    
    
    If modStickGame.sv_AllowChoppers = False Then
        If Stick(0).WeaponType = Chopper Then
            SwitchWeapon AK
            ChopperAvail = False
        End If
    End If
    If modStickGame.sv_AllowFlameThrowers = False Then
        If Stick(0).WeaponType = FlameThrower Then
            SwitchWeapon AK
            
            If Stick(0).CurrentWeapons(1) = FlameThrower Then
                Stick(0).CurrentWeapons(1) = AK
            ElseIf Stick(0).CurrentWeapons(2) = FlameThrower Then
                Stick(0).CurrentWeapons(2) = AK
            End If
            
        End If
    End If
    If modStickGame.sv_AllowRockets = False Then
        If Stick(0).WeaponType = RPG Then
            SwitchWeapon AK
            
            If Stick(0).CurrentWeapons(1) = RPG Then
                Stick(0).CurrentWeapons(1) = AK
            ElseIf Stick(0).CurrentWeapons(2) = RPG Then
                Stick(0).CurrentWeapons(2) = AK
            End If
            
        End If
    End If
    
    
    
    If modStickGame.sv_Hardcore <> vServerVars.bHardCore Then
        AddMainMessage "Hardcore Mode " & IIf(modStickGame.sv_Hardcore, "Off", "On")
        modStickGame.sv_Hardcore = vServerVars.bHardCore
    End If
    
    
    If modStickGame.sv_2Weapons <> vServerVars.b2Weapons Then
        modStickGame.sv_2Weapons = vServerVars.b2Weapons
        
        If modStickGame.sv_2Weapons Then
            AddMainMessage "You can only carry two weapons (1 or 2 to switch)"
            SetCurrentWeapons
        Else
            AddMainMessage "You have access to all weapons"
        End If
    End If
    
    
    
    
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


End Sub

Private Function IsValidVarPacket(ByVal sPacket As String) As Boolean
Dim iSquare As Single, i As Integer

On Error GoTo EH

i = InStr(1, sPacket, vbSpace)
iSquare = Mid$(sPacket, i + 1)

IsValidVarPacket = IsSquare(iSquare)

EH:
End Function

Public Sub SendServerVarPacket(Optional bForce As Boolean = False)
Dim sToSend As String
Dim vServerVars As ptServerVars

Static LastSend As Long


If LastSend + ServerVarSendDelay < GetTickCount() Or bForce Then
    
    vServerVars.bAllowRockets = modStickGame.sv_AllowRockets
    vServerVars.bAllowFlameThrowers = modStickGame.sv_AllowFlameThrowers
    vServerVars.bShootNades = modStickGame.sv_ShootNades
    vServerVars.sgGameSpeed = modStickGame.sv_StickGameSpeed
    vServerVars.bHardCore = modStickGame.sv_Hardcore
    vServerVars.bAllowChoppers = modStickGame.sv_AllowChoppers
    vServerVars.bHPBonus = modStickGame.sv_HPBonus
    vServerVars.b2Weapons = modStickGame.sv_2Weapons
    vServerVars.iScoreToWin = modStickGame.sv_WinScore
    vServerVars.iGameType = modStickGame.sv_GameType
    
    sToSend = ServerVarToString(vServerVars)
    
    SendBroadcast sServerVarss & sToSend & vbSpace & CStr(MakeSquareNumber())
    
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
        CStr(.ID) & mPacketSep & _
        CStr(.PacketID) & mPacketSep & _
        CStr(.Speed) & mPacketSep & _
        CStr(.State) & mPacketSep & _
        CStr(.WeaponType) & mPacketSep & _
        CStr(.X) & mPacketSep & _
        CStr(.Y) & mPacketSep & _
        CStr(.PrevWeapon) & mPacketSep & _
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
    .ID = CInt(Parts(4))
    .PacketID = CLng(Parts(5))
    .Speed = CSng(Parts(6))
    .State = CInt(Parts(7))
    .WeaponType = CInt(Parts(8))
    .X = CSng(Parts(9))
    .Y = CSng(Parts(10))
    .PrevWeapon = CInt(Parts(11))
    .iNadeType = CInt(Parts(12))
End With

Erase Parts

EH:
End Sub

Private Sub ProcessUpdatePacket(ByVal sPacket As String)

'Dim Num As Integer
Dim i As Integer, j As Integer ', k As Integer
Dim sStick As String
Dim Sticks() As String

Sticks = Split(sPacket, UpdatePacketSep)


'Loop through each stick's info
For i = 0 To UBound(Sticks)
    
    On Error GoTo EH
    
    'Extract stick info
    'sstick = Left$(sPacket, Len(mPacket) + 1)
    'sPacket = Right$(sPacket, Len(sPacket) - (Len(mPacket) + 1))
    
    sStick = Sticks(i)
    
    If LenB(sStick) Then
        'CopyMemory mPacket, ByVal sstick, Len(sstick)
        
        'copy it into mPacket
        
        mPacketFromString sStick
        
        'Does this stick already exist?
        If FindStick(mPacket.ID) = -1 Then
            'If Not StickServer Then
                'new stick.  Make new spot and assign ID
                Stick(AddStick()).ID = mPacket.ID
            'End If
        End If
        
        'Is this the local stick?
        If mPacket.ID <> MyID Then
            'Is this a new packet?
            j = FindStick(mPacket.ID)
            If Stick(j).LastPacketID < mPacket.PacketID Then
                'Replace stick data with new data
                With Stick(j)
                    '.Colour = mPacket.Colour
                    .ActualFacing = mPacket.ActualFacing
                    
                    .Facing = mPacket.Facing
'                    If (.State And Stick_Fire) = 0 Then
'                        .Facing = .ActualFacing
'                    End If
                    
                    .Heading = mPacket.Heading
                    '.ID = mpacket.ID
                    '.Name = mPacket.Name
                    .LastPacketID = mPacket.PacketID
                    .Speed = mPacket.Speed
                    .State = mPacket.State
                    .X = mPacket.X
                    .Y = mPacket.Y
                    
                    .Health = mPacket.Health
                    '.Armour = mPacket.Armour
                    '.bAlive = mPacket.bAlive
                    .WeaponType = mPacket.WeaponType
                    .PrevWeapon = mPacket.PrevWeapon
                    
                    .LastPacket = GetTickCount()
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
    .ID = MyID
'    .Name = Stick(0).Name
    .PacketID = Stick(0).LastPacketID
    .Speed = Stick(0).Speed
    .State = Stick(0).State
    .X = Stick(0).X
    .Y = Stick(0).Y
    
    .Health = Stick(0).Health
'    .Armour = Stick(0).Armour
'    .bAlive = Stick(0).bAlive
    .WeaponType = Stick(0).WeaponType
    .PrevWeapon = Stick(0).PrevWeapon
    
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

modWinsock.SendPacket socket, ServerSockAddr, sPacket

End Sub

Private Sub SendServerUpdatePacket()

Dim i As Long
Dim sPacket As String

'Increment the local Stick's LastPacketID
On Error GoTo EH
Stick(0).LastPacketID = Stick(0).LastPacketID + 1

For i = 0 To NumSticksM1
    'Fill the mPacket
    With mPacket
'        .Colour = Stick(i).Colour
        .ActualFacing = Stick(i).ActualFacing
        .Facing = Stick(i).Facing
        .Heading = Stick(i).Heading
        .ID = Stick(i).ID
'        .Name = Stick(i).Name
        '.PacketID = Stick(i).LastPacketID
        .PacketID = IIf(Stick(i).IsBot, Stick(0).LastPacketID, Stick(i).LastPacketID)
        .Speed = Stick(i).Speed
        .State = Stick(i).State
        .X = Stick(i).X
        .Y = Stick(i).Y
        
        .Health = Stick(i).Health
'        .Armour = Stick(i).Armour
'        .bAlive = Stick(i).bAlive
        .WeaponType = Stick(i).WeaponType
        .PrevWeapon = Stick(i).PrevWeapon
        
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
Do While i < NumSticks
    'Ensure this isn't the local Stick
    If Stick(i).ID <> MyID Then
        If Stick(i).IsBot = False Then
            'Send!
            If modWinsock.SendPacket(socket, Stick(i).ptsockaddr, sPacket) = False Then
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

'Extra Info
'#################################################################################################

Private Function SlowPacketToString() As String

With msPacket
    SlowPacketToString = _
        Trim$(.Name) & mPacketSep & _
        CStr(.Colour) & mPacketSep & _
        CStr(.Armour) & mPacketSep & _
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
        Abs(.bOnFire) & mPacketSep
End With

End Function

Private Sub SlowPacketFromString(buf As String)

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

With msPacket
    .Name = Trim$(Parts(0))
    .Colour = CLng(Parts(1))
    .Armour = CInt(Parts(2))
    .iKills = CInt(Parts(3))
    .iDeaths = CInt(Parts(4))
    .iKillsInARow = CInt(Parts(5))
    .Team = CInt(Parts(6))
    .bAlive = CBool(Parts(7))
    .Perk = CInt(Parts(8))
    .MaskID = CInt(Parts(9))
    .ID = CInt(Parts(10))
    .bSilenced = CBool(Parts(11))
    .bTyping = CBool(Parts(12))
    .bFlashed = CBool(Parts(13))
    .bOnFire = CBool(Parts(14))
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
            'If Not StickServer Then
                'new stick.  Make new spot and assign ID
                j = AddStick()
                Stick(j).ID = msPacket.ID
            'End If
        End If
        
        'Is this the local stick?
        If msPacket.ID <> MyID Then
            
            With Stick(j)
                .Armour = msPacket.Armour
                .bAlive = msPacket.bAlive
                .Colour = msPacket.Colour
                .iDeaths = msPacket.iDeaths
                .iKills = msPacket.iKills
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
            End With
            
        ElseIf Not modStickGame.StickServer Then
            'us + we're not the server
            
            'j = FindStick(msPacket.ID)
            If j = 0 Then
                'should do
                
                'only let it update our kills
                If Stick(0).iKills < msPacket.iKills Then
                    Stick(0).iKills = msPacket.iKills
                End If
                
                'our deaths are always correct
                ''''''Stick(0).iDeaths = msPacket.iDeaths
                
                'if our killsinarow < mspacket's, then update
                'If Stick(0).iKillsInARow < msPacket.iKillsInARow Then
                    
                    'Stick(0).iKillsInARow = msPacket.iKillsInARow
                    'CheckKillsInARow
                    
                'End If
                
            End If
            
            
        End If 'myid/iKills endif
    End If 'lenb endif
    
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

With msPacket
    .Colour = Stick(0).Colour
    .Name = Stick(0).Name
    .Armour = Stick(0).Armour
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
End With

SendPacket = sSlowUpdates & SlowPacketToString() & UpdatePacketSep

modWinsock.SendPacket socket, ServerSockAddr, SendPacket

End Sub

Private Sub SendServerSlowPacket()
Dim i As Long
Dim sPacket As String

On Error GoTo EH

For i = 0 To NumSticksM1
    With msPacket
        .Colour = Stick(i).Colour
        .Name = Stick(i).Name
        .Armour = Stick(i).Armour
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
    End With
    
    'Append
    sPacket = sPacket & SlowPacketToString() & UpdatePacketSep
Next i

sPacket = sSlowUpdates & sPacket

'Send it to all non-local Stick
i = 1
Do While i < NumSticks
    'Ensure this isn't the local Stick
    If Stick(i).ID <> MyID Then
        If Stick(i).IsBot = False Then
            'Send!
            If modWinsock.SendPacket(socket, Stick(i).ptsockaddr, sPacket) = False Then
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

