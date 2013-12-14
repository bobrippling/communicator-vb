VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiplayer Combat"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11580
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11580
   Visible         =   0   'False
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox picBlank 
      Height          =   255
      Left            =   1560
      Picture         =   "frmGame.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   360
   End
   Begin VB.PictureBox picHandle 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblDragInfo 
      BackColor       =   &H00000000&
      Caption         =   "You can only shorten/lengthen the full bars (but resize the transparent one)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Shape ln2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4095
      Left            =   3480
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ob1 
      BackColor       =   &H8000000C&
      BorderColor     =   &H00808080&
      Height          =   5175
      Left            =   6480
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ln1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   2160
      Top             =   6120
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuReset 
      Caption         =   "Reset"
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Tmp As String

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Windows declarations
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
    StateDragging = 1
    StateSizing = 2
End Enum

Private m_CurrCtl As Control
Private m_DragState As ControlState
Private m_DragHandle As Integer
Private m_DragRect As New cRect
Private m_DragPoint As POINTAPI

Private m_bDesignMode As Boolean

Private bSaved As Boolean

Private Const BoxPosDelay = 5000
'----------------------------------------------------------------------------------------------


'########
'barriers
'########
Private Const MaxWidth = 20000  '11670
Private Const MaxHeight = 15000 '8970
'declared as ints to prevent overflow with other ints
Private Const StarBarrier = 6000

Private ClosingWindow As Boolean

Private Type BULLETTYPE
    X As Single          'X coord of this bullet
    Y As Single          'Y coord of this bullet
    Speed As Single      'Speed of this bullet
    Heading As Single    'Direction the bullet is headed
    Decay As Long        'Tick count at which this bullet will decay
    Owner As Integer     'Which player owns this bullet?
    Colour As Long
    Damage As Single
    LastDeflect As Long 'has bullet been through the thing?
End Type

Private NumBullets As Long      'How any bullets are in the array?
Private Bullet() As BULLETTYPE  'A nice little array of bullets

Private Type ptStar
    X As Single          'X Coord of the star
    Y As Single          'Y Coord of the star
    RelSpeed As Single   'Speed of the star relative to the speed of the ship
End Type

Private Const NUM_STARS = 200        'Number of stars in the field
Private Stars(NUM_STARS - 1) As ptStar
'Private Const Star_Speed = 0.8 '0.2
Private Const STAR_RADIUS = 5      'Radius of the stars
Private Const NUM_Star_LAYERS = 5        'Number of speed layers in the field

'Private Const Num_Planets = 3
'Private Const Num_Planet_Layers = 2
'Private Const Planet_Radius = 60
'Private Planets(Num_Planets - 1) As ptStar
'Private Sun As ptStar
'Private Const Sun_Radius = 100

Private Const Max_Chat = 24
Private FPS As Integer
Private Timer As Long
'Public TimeFactor As Single

Private Const Start_Speed = 50
Public ROTATION_RATE As Integer    'Rotation speed of the ship
Private Const SHIP_RADIUS = 150    'Distance from center of triangle to any vertex
Private Const SHIP_Height = SHIP_RADIUS + 75 '150+75
Private Const SHIELD_REGEN = 0.5      'Rate of shield regeneration
Private Const Bullet_Radius = 30     'Radius of the bullets
Private Const BULLET_SPEED = 180
Private Const Bullet_Decay = 1000 '600    'How many milliseconds does each bullet last?
Private Const Bullet_Decay_Extra = 200 'How many milliseconds extra can a bullet last? (rnd * <--)
Private Const Default_Bullet_Damage = 20
'Private Const BULLET_COST = 0       'Energy cost of firString a bullet
Private Const BULLET_LEN = SHIP_Height - 10      'Length of a bullet
Private Const Gun_Len As Integer = 200   'Length of a gun
Private Const GunOffset As Single = 0.03

Private Const Default_Rotation_Rate = 15

'missiles
Private Type MissileType
    X As Single
    Y As Single
    Speed As Single
    Heading As Single
    Facing As Single
    Decay As Long
    Owner As Integer
    Colour As Long
    TargetID As Integer
    Hull As Single
    InRange As Boolean
    
    LastSmoke As Long
End Type

Private Const Missile_Hull_Start = 120
Private MissilesShot As Integer
Private NumMissiles As Long
Private Missiles() As MissileType

Private Const Missile_Delay = 12000
Private Const Missile_SPEED = BULLET_SPEED / 1.1
Private Const Missile_Decay = Bullet_Decay * 4
'Private Const Missile_Damage = Default_Bullet_Damage * 4 '<--- Missile damage is now its Hull
Private Const Missile_LEN = BULLET_LEN / 1.5
Private Const Missile_Radius = Bullet_Radius * 2
Private Const MissileLockDist = 8000 'the min distance for a missile to home

Private Const MissileKeyReleaseDelay = 5600 'wait x seconds before lifting key



'#################
'PACKET STUFF
Private Const mPacket_LAG_TOL = 1000  'Milliseconds to wait before rendering a player motionless due to lack of info
Private Const mPacket_LAG_KILL = 5000    'Milliseconds to wait before removing a player due to lack of info
'                                   ^ must be 11, so that they don't lag out at the end of a round
'Private Const SERVER_CONNECT_DURATION = 10000    'Milliseconds during which we will attempt to connect to server
Private Const SERVER_RETRY_FREQ = 2000 'SERVER_CONNECT_DURATION / 3 - 10 'Milliseconds between attempts to connect to server
Private Const SERVER_NUM_RETRIES = 5

Private Const mPacket_SEND_DELAY = 50   'Milliseconds between update packets
'Private Const Default_mPacket_SEND_DELAY = 100   'Milliseconds between update packets
'Private mPacket_SEND_DELAY As Long
Private Const AntiLagPacketDelay = 1000
'#################

Private LastUpdatePacket As Long 'when did this client last receive an update?
Private ServerSockAddr As ptSockAddr


'validation
'Private Const sPacketLen = 57 '61?
'Private Const BoxPosLen = 53 '35
'Private Const AsteroidLen = 23

'ship specific --------------------------------------------------------------------------------
Private Const Raptor_ACCEL As Single = 9
Private Const Behemoth_ACCEL As Single = 4
Private Const Hornet_Accel As Single = 12 '10
Private Const Mothership_Accel As Single = 0.5
Private Const Wraith_Accel As Single = 7
Private Const Infil_Accel As Single = 7.5
Private Const SDNorm_Accel As Single = 4
Private Const SDGW_Accel As Single = 1

Private Const Raptor_MAX_SPEED = 200
Private Const Behemoth_MAX_SPEED = 90
Private Const Hornet_Max_Speed = BULLET_SPEED + 10
Private Const MotherShip_Max_Speed = 15
Private Const Wraith_Max_Speed = 130
Private Const Infil_Max_Speed = 140
Private Const SDNorm_Max_Speed = 18
Private Const SDGW_Max_Speed = 10

Private Const RaptorBulletDamage = 2.5
Private Const BehemothDmgReduction = 2.2
Private Const HornetDmgIncrease = 1.4
Private Const MothershipDmgReduction = 9
Private Const WraithDmgReduction = 1.7
Private Const InfilDmgIncrease = 0.8 '.1
Private Const SDDmgReduction = 6

Private Const Raptor_Bullet_DELAY = 33
Private Const Behemoth_Bullet_DELAY = 110
Private Const Hornet_Bullet_Delay = 33 '50
Private Const Mothership_Bullet_Delay = 20 '15
Private Const Wraith_Bullet_Delay = 40
Private Const Infil_Bullet_Delay = 300
Private Const SD_Bullet_Delay = 150 '70

'GETSHIPRADIUS AND GETACCEL NEED ADDING TO, TOO

Private Const MotherShipDeflPercent = 0.7
Private Const MotherShipFireTime = 3000
Private Const MotherShipRechargeTime = 6000 + MotherShipFireTime

Private Const Hornet_Bullet_Speed = BULLET_SPEED * 3 '1.5
Private Const Wraith_Bullet_Speed = BULLET_SPEED * 1.3
Private Const Infil_Bullet_Damage_Factor = 7
Private Const MS_Bullet_Damage_Factor = 3

Private Const SD_GravityRadius = 15 'Ship_Height * 6
'end ship specific --------------------------------------------------------------------------------

Private Const BoxColour As Long = &H8000000C
'Private Const Box1Colour As Long = vbGreen
'Private Const Box2Colour As Long = &H8000000C


Public NumPlayers As Long      'How any players in the game?
Public MyID As Integer         'Which player are we?

Private Type ptPacket
    ID As Integer        'Player's unique ID
    PacketID As Long     'This mPacket's ID
    Name As String * 20  'Umm.. what was this one?  Oh yes, player's name
    Facing As Single     'Facing update
    Heading As Single    'Heading update
    Speed As Single      'Speed update
    X As Single          'Location update
    Y As Single          'Location update
    State As Integer     'Current actions update
    Colour As Long
    'ShipType As Byte
    Kills As Integer
    Deaths As Integer
    'IsBot As Integer
    'Team As Byte
    Alive As Boolean
End Type

Private mPacket As ptPacket

Public bRunning As Boolean  'Is the render loop running?
'Private Timer As Long       'Frame timer
Private PacketTimer As Long 'Time at which last mPacket was sent
Public socket As Long       'Socket with which we'll send/receive essages

'Private RightSideBullet As Boolean

Private Type ptChat
    Text As String           'The string to display
    Decay As Long            'When is this one removed?
    Colour As Long
End Type

Private Chat() As ptChat       'Our chat array
Private NumChat As Long          'How many chat messages are there currently?
Private Const CHAT_DECAY = 15000        'How long before chat messages disappear?

'chat
Private Const ChatFlashDelay = 300 'for the _ thing
Private LastFlash As Long
Private bChatCursor As Boolean 'for the _ thing
Private bChatActive As Boolean   'chatting?
Private strChat As String 'current chat string
'end chat

'###########################################
'keys pressed
Private KeyW As Boolean, KeyA As Boolean, keys As Boolean, KeyD As Boolean
Private KeyFire As Boolean, KeySecondary As Boolean, KeyShield As Boolean
'###########################################

'big message(s)
Private Type ptMainMessage
    Text As String           'The string to display
    Decay As Long            'When is this one removed?
    'Colour As Long = mGrey
End Type

Private MainMessages() As ptMainMessage       'Our chat array
Private NumMainMessages As Long          'How many chat messages are there currently?
Private Const MainMessageDecay = 3000        'How long before chat messages disappear?


Public AI_Sample_Rate As Long
'Public Const Default_AI_Sample_Rate = 400 - in modSpaceGame

Private MouseX As Single, MouseY As Single

'Private Type ptScore
'    Name As String * 20
'    Score As Integer
'    ID As Integer
'End Type
'
'Private ScoreList() As ptScore

Private Const ShotBy As String = " was shot by "
Private Const MissiledBy As String = " was missiled by "
Private Const RammedBy As String = " was rammed by "

Private Const sUpdates As String * 1 = "U"
Private Const sJoins As String * 1 = "J"
Private Const sAccepts As String * 1 = "A"
Private Const sChats As String * 1 = "C"
Private Const sQuits As String * 1 = "Q"
Private Const sServerQuits As String * 1 = "W"
Private Const sShipTypes As String * 1 = "T"
Private Const sEndRounds As String * 1 = "E"
Private Const sNewRounds As String * 1 = "N"
Private Const sTeams As String * 1 = "M"
Private Const sBoxPoss As String * 1 = "B" ' B(ob1.x,ob1.y)(ln1.x,ln1.y)
Private Const sPowerUps As String * 1 = "P"
Private Const sGameSpeeds As String * 1 = "G"
Private Const sRemovePlayers As String * 1 = "R"
Private Const sScoreUpdates As String * 1 = "O"
Private Const sAsteroidUpdates As String * 1 = "S"
Private Const sServerVarsUpdates As String * 1 = "V"
Private Const sAntiLagPackets As String * 1 = "L"
Private Const sGameTypes As String * 1 = "Y"
Private Const sHasFlags As String * 1 = "H"
'NOTICE BELOW
'public const sForceTeams as String *1 = "F"
'Public Const sKicks As String * 1 = "K"

'left: DIXZ, abcdefghijklmnopqrstuvwxyz

'##############################################################
'Game Type stuff
Private Const GameTypeSendDelay = 3000
Private Const Flag_Width = SHIP_RADIUS * 2
Private Const Flag_Colour = MGreen
Private Const FlagBaseRadius = Flag_Width * 3
'Private Const BaseStayTime = 5000 'time to stay in base for a win

Private Const FlagBaseX = MaxWidth / 2, FlagBaseY = MaxHeight / 2
Private Const FlagDefaultX = CentreX, FlagDefaultY = 2000

Public FlagOwnerID As Integer
Private BaseEnteredAt As Long
'##############################################################

'round stuff
Private Const ScoreCheckDelay = 2000 'check the scores every x secs
Private Const RoundWaitTime = 10000 '10 seconds between round
'Private MaxScore As Integer ' = 10(numplayers-1)
Private RoundPausedAtThisTime As Long
Public Playing As Boolean 'used in options - can they choose a team?
Private RoundWinnerID As Integer

'score stuff
'Private MyScore As Integer
Private KillsInARow As Integer
Private ShipScores(0 To eShipTypes.SD) As Integer 'holds the score obtained with each ship
Private F1Pressed As Boolean

'MS fire stuff
Public MotherShipAvail As Boolean
Private Const KillsForMS = 2
Private MSStartFire As Long

Public WraithAvail As Boolean
Private Const KillsForWraith = 7

Public InfilAvail As Boolean
Private Const KillsForInfil = 5

Public SDAvail As Boolean
Private Const KillsForSD = 3

'shiptype send delay
Private Const ShipTypeSendDelay = 5000

Private Const ScoreUpdateDelay = 1000

'##########
'respawn gCircle
'Private LastRespawn As Long
'Private Const RespawnCircleRadius = 100 '1000 is added to by below
'Private Const RespawnCircleShowTime = 2000 'show it for 2 seconds


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

Private Circs() As Circ
Private NumCircs As Integer

'ms 2ndry fire
Private Const MS2_Delay = Raptor_Bullet_DELAY * 3
Private Const Wraith2_Delay = MS2_Delay

'font sizes
Private Const BigFontSize = 10
Private Const NormalFontSize = 8
Private Const Thick = 30
Private Const Thin = 1

'teams
Private Const TeamSendDelay = ShipTypeSendDelay + 1
'name checking
Private Const NameCheckDelay = 10000

'time based modeling
'http://gpwiki.org/index.php/VB:Tutorials:Time_Based_Modelling
'http://gpwiki.org/index.php/VB:Tutorials:Time_Based_Modelling_IT
'Public ElapsedTime As Long
'Private Const Max_FPS = 70
'Private LastFullSecond As Long
'Private Const FPSCap = 100

'Public TimeFactor As Single
'Private FPS As Long
'
'Private Sub TbmTimerProc()
'
'Static tbmTimer As Long
'Static FrameTimer As Long
'Static FPSCounter As Long
'
''Determine the time that has elapsed since the last frame was displayed
'ElapsedTime = GetTickCount() - tbmTimer
'
''Reset the general timer
'tbmTimer = GetTickCount()
'
''Check if one second has elapsed
'If GetTickCount() - FrameTimer >= 1000 Then
'    'Set the FPS storage var, and reset the FPS counter/timer
'    FPS = FPSCounter + 1
'    FPSCounter = 0
'    FrameTimer = GetTickCount()
'Else
'    'If a second hasn't elapsed, add to the FPS counter
'    FPSCounter = FPSCounter + 1
'End If
'
'
'End Sub

'end tbm

Private Type ptPowerup
    X As Single
    Y As Single
    Active As Boolean
End Type

Private PowerUp As ptPowerup
Private Const Powerup_Radius = Bullet_Radius * 4
Private Const PowerUpDelay = 15000

Private Const GameSpeedSendDelay = 5000
Private Const VarUpdateSendDelay = 2000

Private Type ptAsteroid
    X As Single '4
    Y As Single '4
    Speed As Single '4
    Heading As Single '4
    Facing As Single '4
    LastPlayerTouchID As Integer '2
End Type

'Private Type ptAsteroidBuff
'    Data As String * 22
'End Type

Private Asteroid As ptAsteroid
Private Const Asteroid_Radius = SHIP_Height * 2
Private Const AsteroidSendDelay = 1000
Private Const AsteroidColour As Long = MBrown
Private Const AsteroidColour2 As Long = &H165195
Private Const AsteroidColour3 As Long = &H165175
Private Const AsteroidSpeed = 12 'slightly slower than MS
Private Const AsteroidMass = 40

Private Const ShieldDmgReduction = 4

'##############################################################################
'Smoke ########################################################################
'##############################################################################

Private Type ptSmoke
    X As Single
    Y As Single
    
    Size As Single
    
    Direction As Integer '1=grow (2x rate), -1=shrink
    
    'Speed As Single
    'Heading As Single
    
End Type

Private Smoke() As ptSmoke
Private NumSmoke As Integer

Private Const SmokeOutline = &H777777
Private Const SmokeFill = &HFDFDFD

'#######################################
'Map/Zoom
'Private Const MapColour = vbBlue

'Private Const MapLeft = 11580 - 1000 '-maplen
'Private Const MapTop = CentreY - 1000  'cg_MapLen / 2
                    '= me.height - 500 - cg_MapLen

'Private MapKeyDown As Boolean

Private Const MaxZoom = 2.49 'take one
Private Const MinZoom = 0.71 'add one
Private Const ZoomInc = 0.05

Private Const ZoomShowTime = 1000
Private LastZoomPress As Long

Private Const CameraInc = 250
'Private Const BaseFontSize = 8
'#######################################

Private Sub BltToForm()

BitBlt Me.hDC, 0, 0, ScaleX(MaxWidth, vbTwips, vbPixels), ScaleY(MaxHeight, vbTwips, vbPixels), _
    Me.picMain.hDC, 0, 0, modSpaceGame.cg_SpaceDisplayMode 'vbSrcCopy

'RasterOpConstants
End Sub

Private Sub DrawMap()
Dim MapTop As Single, MapLeft As Single
Dim i As Integer
Dim C As Long 'yes needed
Dim pX As Single, pY As Single
Dim bCan As Boolean

If modSpaceGame.cg_ShowMap Then
    
    picMain.DrawWidth = 2
    
    MapLeft = 11580 - cg_MapLen
    MapTop = 10 'CentreY - cg_MapLen / 2
    
    picMain.Line (MapLeft, MapTop)-(MapLeft + cg_MapLen, MapTop + cg_MapLen), vbBlue, B
    
    picMain.FillStyle = 0
    picMain.FillColor = AsteroidColour
    picMain.Circle (MapLeft + cg_MapLen * Asteroid.X / MaxWidth, _
           MapTop + cg_MapLen * Asteroid.Y / MaxHeight), _
           75, AsteroidColour
    
    
    If PowerUp.Active Then
        picMain.FillColor = vbBlue
        picMain.Circle (MapLeft + cg_MapLen * PowerUp.X / MaxWidth, _
            MapTop + cg_MapLen * PowerUp.Y / MaxHeight), _
            25, vbBlue
    End If
    
    
    For i = 0 To NumPlayers - 1
        
        If PlayerInGame(i) Then
            
            bCan = True
            If Player(i).ShipType = Infiltrator Then
                If Player(i).State And Player_Secondary Then
                    bCan = False
                End If
            End If
            
            If bCan Then
                pX = MapLeft + cg_MapLen * Player(i).X / MaxWidth
                pY = MapTop + cg_MapLen * Player(i).Y / MaxHeight
                
                C = GetTeamColour(Player(i).Team)
                picMain.FillColor = C
                picMain.Circle (pX, pY), 60, C
                
                'picMain.FillColor = Player(i).Colour
                'picMain.Circle (pX, pX), 70, Player(i).Colour
                
                
                If i = 0 Then
                    'draw an X on me
                    DrawX pX, pY
                End If
                
            End If
            
        End If
        
    Next i
    
    
    picMain.FillStyle = 1
    
End If

End Sub

Private Sub DrawX(ByVal pX As Single, ByVal pY As Single)

Const CrossWidth = 75

'X(0) = pX - CrossWidth '2
'X(1) = pX + CrossWidth '3
'
'Y(0) = pY + CrossWidth '2
'Y(1) = pY - CrossWidth '3

picMain.Line (pX - CrossWidth, pY + CrossWidth)-(pX + CrossWidth, pY - CrossWidth), Player(0).Colour
picMain.Line (pX + CrossWidth, pY + CrossWidth)-(pX - CrossWidth, pY - CrossWidth), Player(0).Colour

End Sub

Private Sub DrawFlag(sX As Single, sY As Single) 'ByVal iPlayer As Integer)
Dim X(1 To 5) As Single
Dim Y(1 To 5) As Single
Dim i As Integer

picMain.ForeColor = Flag_Colour

'5 = centre
X(5) = sX ' Player(iPlayer).X
Y(5) = sY 'Player(iPlayer).Y

X(3) = X(5)
X(1) = X(5)
X(2) = X(5) + Flag_Width
X(4) = X(2)

Y(3) = Y(5) - Flag_Width
Y(1) = Y(3) - Flag_Width
Y(2) = Y(1)
Y(4) = Y(3)

For i = 1 To 3
    gLine X(i), Y(i), X(i + 1), Y(i + 1), Flag_Colour
Next i

gLine X(4), Y(4), X(1), Y(1), Flag_Colour
gLine X(4), Y(4), X(2), Y(2), Flag_Colour
gLine X(3), Y(3), X(1), Y(1), Flag_Colour
gLine X(3), Y(3), X(5), Y(5), Flag_Colour

End Sub

Private Sub DrawFlagBase()

gCircle FlagBaseX, FlagBaseY, FlagBaseRadius, BoxColour
'Me.ForeColor = BoxColour
picMain.Font.Size = NormalFontSize
PrintText "Flag Base", FlagBaseX - 300, FlagBaseY - 20, Player(0).Colour

End Sub

Private Sub DoElimination()
Dim NumAlive As Integer, i As Integer
Static LastCheck As Long

Dim RedPresent As Boolean, BluePresent As Boolean, NeutralPresent As Boolean
Dim RoundEnded As Boolean
Dim BiggestScore As Integer
Dim BestID As Integer

If LastCheck + ScoreCheckDelay < GetTickCount() Then
    
    If NumPlayers <> 1 Then
        
        For i = 0 To NumPlayers - 1
            If PlayerInGame(i) Then
                NumAlive = NumAlive + 1
                
                If Player(i).Team = Blue Then
                    BluePresent = True
                ElseIf Player(i).Team = Red Then
                    RedPresent = True
                Else
                    NeutralPresent = True
                End If
                
            End If
        Next i
        
        
        
        If NumAlive = 1 Then
            For i = 0 To NumPlayers - 1
                If PlayerInGame(i) Then
                    RoundWinnerID = Player(i).ID
                    RoundEnded = True
                    StopPlay True
                    Exit For
                End If
            Next i
            
            
            'For i = 0 To NumPlayers - 1
                'Player(i).Team = Neutral
            'Next i
        ElseIf Not RoundEnded Then
            
            
            'no single player is alive, but a team could be alive
            If RedPresent And BluePresent = False And NeutralPresent = False Then
                'find best red player + end
                
                BiggestScore = -1
                
                For i = 0 To NumPlayers - 1
                    If PlayerInGame(i) Then
                        If Player(i).Team = Red Then
                            If Player(i).Score > BiggestScore Then
                                BiggestScore = Player(i).Score
                                BestID = Player(i).ID
                            End If
                        End If
                    End If
                Next i
                
                RoundWinnerID = BestID
                RoundEnded = True
                StopPlay True
                
            ElseIf BluePresent And RedPresent = False And NeutralPresent = False Then
                'find best blue player + end
                
                BiggestScore = -1
                
                For i = 0 To NumPlayers - 1
                    If PlayerInGame(i) Then
                        If Player(i).Team = Blue Then
                            If Player(i).Score > BiggestScore Then
                                BiggestScore = Player(i).Score
                                BestID = Player(i).ID
                            End If
                        End If
                    End If
                Next i
                
                RoundWinnerID = BestID
                RoundEnded = True
                StopPlay True
                
            End If
            
            
            
        End If
        
    End If
    
    
    LastCheck = GetTickCount()
    
End If

End Sub

Private Sub DoCTF()
Dim i As Integer
Dim FlagOwned As Boolean
Dim TimeLeft As Long
Const Flag_WidthX2 = Flag_Width * 2

If FlagOwnerID <> -1 Then
    i = FindPlayer(FlagOwnerID)
    If i <> -1 Then
        DrawFlag Player(i).X, Player(i).Y
        FlagOwned = True
    Else
        FlagOwnerID = -1
    End If
End If

If FlagOwned = False Then
    DrawFlag FlagDefaultX, FlagDefaultY
    
    'check if a player is picking up t'flag
    For i = 0 To NumPlayers - 1
        With Player(i)
            If GetDist(.X, .Y, FlagDefaultX, FlagDefaultY) < (GetShipRadius(.ShipType) + Flag_WidthX2) Then
                FlagOwnerID = .ID
                FlagOwned = True
                Exit For
            End If
        End With
    Next i
    
End If


DrawFlagBase


If FlagOwned Then
    If GetDist(FlagBaseX, FlagBaseY, Player(i).X, Player(i).Y) < FlagBaseRadius Then
        
        If BaseEnteredAt = 0 Then
            BaseEnteredAt = GetTickCount()
        End If
        
        TimeLeft = BaseEnteredAt / 1000 + modSpaceGame.sv_CTFTime / modSpaceGame.sv_GameSpeed - GetTickCount() / 1000
        '+BaseStayTime
        
        If TimeLeft <= 0 Then
            'they've won
            If modSpaceGame.SpaceServer Then
                RoundWinnerID = Player(i).ID
                BaseEnteredAt = 0
                FlagOwnerID = -1
                StopPlay True
            End If
        Else
            picMain.Font.Size = BigFontSize
            'Me.ForeColor = Player(i).Colour
            PrintFormText "Time Left to Flag Capture: " & CStr(Round(TimeLeft, 1)), CentreX - 800, 2000, Player(i).Colour
            picMain.Font.Size = NormalFontSize
            
            '-5% their speed
            Player(i).Speed = Player(i).Speed - 5 * Player(i).Speed * modSpaceGame.sv_GameSpeed / 10
            
        End If
        
        
    ElseIf BaseEnteredAt Then
        BaseEnteredAt = 0
    End If

End If

If Server Then
    SendPlayerFlagUpdate
End If

End Sub

Public Sub ResetVars()

bRunning = False
modSpaceGame.UseAI = False

Erase modSpaceGame.Player
Erase Bullet
Erase Chat
Erase Missiles
Erase Circs
NumPlayers = 0
NumBullets = 0
NumMissiles = 0
NumChat = 0
NumCircs = 0
MissilesShot = 0

modSpaceGame.sv_GameSpeed = 1
modSpaceGame.sv_BotAI = True
modSpaceGame.sv_GameType = DM

MyID = 0
'MyScore = 0
KillsInARow = 0
Erase ShipScores
F1Pressed = False
'SomeOneShooting = False
'SomeOneThrusting = False
MotherShipAvail = False
WraithAvail = False
SDAvail = False

modSpaceGame.SpaceServer = False
modSpaceGame.SpaceServer = False

MSStartFire = 0

End Sub

Private Sub RndAsteroid()

With Asteroid
    .X = Rnd() * MaxWidth
    .Y = Rnd() * MaxHeight
    .Speed = AsteroidSpeed
    .Heading = Rnd() * Pi2
    .LastPlayerTouchID = -1
End With

End Sub

Private Sub ProcessAsteroid()
Const Lim = 500 '50
Const aSize = 1.05 '1.3
Dim ClipX As Boolean, ClipY As Boolean
Dim XComp As Single, YComp As Single

Dim tX(1 To 4) As Single, tY(1 To 4) As Single

'move it
Motion Asteroid.X, Asteroid.Y, Asteroid.Speed, Asteroid.Heading

Asteroid.Facing = Asteroid.Facing + modSpaceGame.sv_GameSpeed * Asteroid.Speed / 1000

'clip it
With Asteroid
    If .X < Lim Or .X > (MaxWidth - Lim) Then
        ClipX = True
    ElseIf .Y < Lim Or .Y > (MaxHeight - Lim) Then
        ClipY = True
    End If
    
    If ClipX Or ClipY Then
        XComp = .Speed * Sine(.Heading)
        YComp = .Speed * CoSine(.Heading)
    End If
    
    If ClipX Then
        If .X < Lim Then
            .X = Lim
            XComp = Abs(XComp)
        Else
            .X = MaxWidth - Lim
            XComp = -Abs(XComp)
        End If
    ElseIf ClipY Then
        If .Y < Lim Then
            .Y = Lim
            YComp = -Abs(YComp)
        Else
            .Y = MaxHeight - Lim
            YComp = Abs(YComp)
        End If
    End If
    
    If ClipX Or ClipY Then
        'Determine the resultant speed
        .Speed = Sqr(XComp ^ 2 + YComp ^ 2)
        
        'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
        If YComp > 0 Then .Heading = Atn(XComp / YComp)
        If YComp < 0 Then .Heading = Atn(XComp / YComp) + Pi
    End If
    
End With

'draw it
picMain.FillStyle = vbSolid
picMain.FillColor = AsteroidColour
picMain.DrawWidth = 10

tX(1) = Asteroid.X + Asteroid_Radius * Sine(Asteroid.Facing) * aSize / 1.1
tX(2) = Asteroid.X + Asteroid_Radius * Sine(Asteroid.Facing + Pi) * aSize
tX(3) = Asteroid.X + Asteroid_Radius * Sine(Asteroid.Facing + piD2) * aSize
tX(4) = Asteroid.X + Asteroid_Radius * Sine(Asteroid.Facing - piD4) * aSize
tY(1) = Asteroid.Y - Asteroid_Radius * CoSine(Asteroid.Facing) * aSize / 1.1
tY(2) = Asteroid.Y - Asteroid_Radius * CoSine(Asteroid.Facing + Pi) * aSize
tY(3) = Asteroid.Y - Asteroid_Radius * CoSine(Asteroid.Facing + piD2) * aSize
tY(4) = Asteroid.Y - Asteroid_Radius * CoSine(Asteroid.Facing - piD4) * aSize

'gCircleaspect Asteroid.X, Asteroid.Y, Asteroid_Radius, AsteroidColour, , , 0.8
gCircleAspect Asteroid.X, Asteroid.Y, Asteroid_Radius, AsteroidColour, 0.8

gLine tX(1), tY(1), tX(2), tY(2), AsteroidColour2
gLine tX(3), tY(3), tX(2), tY(2), AsteroidColour3
'gline tX(3), tY(3),tX(1), tY(1)), AsteroidColour
gLine tX(4), tY(4), tX(1), tY(1), AsteroidColour2
'gline Asteroid.X, Asteroid.Y,tX(2), tY(2)), AsteroidColour
gLine tX(4), tY(4), tX(2), tY(2), AsteroidColour2
gLine tX(4), tY(4), tX(3), tY(3), AsteroidColour

picMain.FillStyle = 1
picMain.FillColor = vbBlack

If Asteroid.Speed > AsteroidSpeed Then
    'friction-ish
    Asteroid.Speed = Asteroid.Speed - 1
End If

If AsteroidCollision(ln1) Then
    ReverseYComp Asteroid.Heading, Asteroid.Speed
ElseIf AsteroidCollision(ln2) Then
    ReverseXComp Asteroid.Heading, Asteroid.Speed
End If

End Sub

Private Sub AddExplosion(ByVal X As Single, ByVal Y As Single, _
    ByVal TimeLen As Single, ByVal Radius As Single, _
    Speed As Single, Heading As Single)

If modSpaceGame.cg_DrawExplosions Then
    AddCirc X, Y, TimeLen, Radius, vbYellow ', Speed, Heading
    AddCirc X, Y, TimeLen, Radius, MOrange, 100 'Speed, Heading, 100
    AddCirc X, Y, TimeLen, Radius, vbRed, 200 'Speed, Heading, 200
End If

End Sub

Private Sub ProcessAllCircs()
Dim i As Integer

For i = NumCircs - 1 To 0 Step -1
    ProcessCirc i
Next i

End Sub

Private Sub ProcessCirc(ByVal Index As Integer)

On Error GoTo EH

Circs(Index).Prog = Circs(Index).Prog + Circs(Index).Direction * 100 * modSpaceGame.TimeFactor
'Else
'    Circs(Index).Prog = Circs(Index).Prog - 100 * modSpaceGame.sv_GameSpeed
'End If

'Motion Circs(Index).x, Circs(Index).y, Circs(Index).Speed, Circs(Index).Heading


If Circs(Index).Prog > Circs(Index).MaxProg Then
    
    Circs(Index).Direction = -1
    
ElseIf Circs(Index).Prog <= 0 Then
    
    RemoveCirc Index
    
Else
    
    picMain.FillStyle = vbSolid
    picMain.FillColor = Circs(Index).Colour
    
    gCircle Circs(Index).X, Circs(Index).Y, Circs(Index).Radius * Circs(Index).Prog, Circs(Index).Colour
    
    picMain.FillStyle = 1
    picMain.FillColor = vbBlack
    
End If


EH:
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
        Circs(i).MaxProg = Circs(i + 1).MaxProg
        Circs(i).Prog = Circs(i + 1).Prog
        Circs(i).Radius = Circs(i + 1).Radius
        Circs(i).X = Circs(i + 1).X
        Circs(i).Y = Circs(i + 1).Y
        Circs(i).Direction = Circs(i + 1).Direction
        Circs(i).Colour = Circs(i + 1).Colour
    Next i
    
    'Resize the array
    ReDim Preserve Circs(NumCircs - 2)
    NumCircs = NumCircs - 1
End If

End Sub

Private Sub SpawnPowerUp(ByVal X As Single, ByVal Y As Single)
PowerUp.X = X
PowerUp.Y = Y
PowerUp.Active = True
End Sub

Public Sub SendGameSpeed(Optional ByVal Force As Boolean = False)
Static LastSend As Long

If LastSend + GameSpeedSendDelay < GetTickCount() Or Force Then
    
    SendBroadcast sGameSpeeds & CStr(modSpaceGame.sv_GameSpeed)
    
    LastSend = GetTickCount()
End If

End Sub

Public Sub SendGameType(Optional ByVal Force As Boolean = False)
Static LastSend As Long

If LastSend + GameTypeSendDelay < GetTickCount() Or Force Then
    
    SendBroadcast sGameTypes & CStr(modSpaceGame.sv_GameType)
    
    LastSend = GetTickCount()
End If

End Sub

Public Sub SendServerVarsUpdate(Optional ByVal Force As Boolean = False)
Static LastSend As Long

If LastSend + VarUpdateSendDelay < GetTickCount() Or Force Then
    
    'handled in GetPacket()
    
    SendBroadcast sServerVarsUpdates & CStr(IIf(modSpaceGame.sv_BulletsCollide, 1, 0) & _
                                       IIf(modSpaceGame.sv_AddBulletVectorToShip, 1, 0) & _
                                       IIf(modSpaceGame.sv_ClipMissiles, 1, 0) & _
                                       IIf(modSpaceGame.sv_BulletWallBounce, 1, 0) & _
                                       CInt(modSpaceGame.sv_Bullet_Damage) & "#" & _
                                       modSpaceGame.sv_CTFTime) & "@" & CStr(modSpaceGame.sv_ScoreReq)
    
    LastSend = GetTickCount()
End If

End Sub

Public Sub SendPlayerFlagUpdate()
Static LastSend As Long

If LastSend + VarUpdateSendDelay < GetTickCount() Then
    
    SendBroadcast sHasFlags & CStr(FlagOwnerID)
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub CheckPlayerNames()
Static LastCheck As Long
Dim i As Integer, j As Integer

If LastCheck + NameCheckDelay < GetTickCount() Then
    
    On Error GoTo EH
    
    For i = 0 To NumPlayers - 1
        For j = NumPlayers - 1 To 0 Step -1
            
            If Player(j).Name = Player(i).Name And Player(j).ID <> Player(i).ID Then
                'Kick player j
                modWinsock.SendPacket frmGame.socket, Player(j).ptSockAddr, modSpaceGame.sKicks & "Same Name"
                Exit Sub 'so we don't get errors, will be checked again
            End If
            
        Next j
    Next i
    
    LastCheck = GetTickCount()
End If

EH:
End Sub

Private Sub HomeMissile(ByVal i As Integer)

Dim ASpeed As Single ', AHeading As Single
Dim j As Integer
Dim CheckSpeed As Boolean ', AttemptNewLock As Boolean


If Missiles(i).TargetID <> -1 Then
    
    j = FindPlayer(Missiles(i).TargetID)
    
    If PlayerInGame(j) Then
        
        If Player(j).ShipType = Infiltrator Then
            If (Player(j).State And Player_Secondary) = Player_Secondary Then
                Missiles(i).InRange = False
            End If
        End If
        
        If GetDist(Missiles(i).X, Missiles(i).Y, Player(j).X, Player(j).Y) < MissileLockDist Then
            
            
            If Not Missiles(i).InRange Then Missiles(i).InRange = True
            
            
            AccurateShot Player(j).X, Player(j).Y, Player(j).Speed, Player(j).Heading, _
                    Missiles(i).X, Missiles(i).Y, Missiles(i).Speed, Missiles(i).Heading, _
                    Missile_SPEED, ASpeed, Missiles(i).Heading
            
            Missiles(i).Speed = IIf(ASpeed > Missile_SPEED, Missile_SPEED, ASpeed)
            
            Missiles(i).Facing = FindAngle(Missiles(i).X, Missiles(i).Y, Player(j).X, Player(j).Y)
        Else
            'Missiles(i).TargetID = -1
            Missiles(i).InRange = False
        End If
        
    Else
        
        'attempt to aquire lock
        CheckSpeed = True
        'AttemptNewLock = True
        'Missiles(i).Facing = FixAngle(Missiles(i).Facing)
        'Missiles(i).Heading = FixAngle(Missiles(i).Heading)
        
        'If Missiles(i).Facing < pi Then
            'Missiles(i).Heading = 0.25 * modSpaceGame.sv_GameSpeed + Missiles(i).Heading
            'Missiles(i).Facing = 0.25 * modSpaceGame.sv_GameSpeed + Missiles(i).Facing
        'Else
            'Missiles(i).Heading = Missiles(i).Heading - 0.25 * modSpaceGame.sv_GameSpeed
            'Missiles(i).Facing = Missiles(i).Facing - 0.25 * modSpaceGame.sv_GameSpeed
        'End If
        
        
    End If
Else
    Missiles(i).InRange = False 'AttemptNewLock = True
End If

With Missiles(i)
    
    If .InRange = False Then
        'Missiles(i).TargetID = FindPlayer(FindLeastDegreeTarget(FindPlayer(Missiles(i).Owner)))
        .TargetID = FindClosestTarget_ID(.X, .Y, .Owner)
        
        CheckSpeed = True
    End If
    
End With

If CheckSpeed Then
    If Missiles(i).Speed < Missile_SPEED Then
        Missiles(i).Speed = Missiles(i).Speed + 5
    ElseIf Missiles(i).Speed <> Missile_SPEED Then
        Missiles(i).Speed = Missile_SPEED
    End If
End If

End Sub

Private Function FindLeastDegreeTarget(ByVal Playerj As Integer, _
    Optional ByVal IncludeInfil As Boolean = True) As Integer

Dim SmallestAngle As Single, SmallestAngleNo As Integer
Dim CurAngle As Single
Dim i As Integer
Dim bCan As Boolean

SmallestAngle = 100 '> 2*pi
SmallestAngleNo = -1

For i = 0 To NumPlayers - 1
    If Player(i).ID <> Player(Playerj).ID Then
        If PlayerInGame(i) Then
            If IsAlly(Player(Playerj).Team, Player(i).Team) = False Then
                
                bCan = True
                
                If IncludeInfil = False Then
                    If Player(i).ShipType = Infiltrator Then
                        If Player(i).State And Player_Secondary Then bCan = False
                    End If
                End If
                
                If bCan Then
                    CurAngle = FixAngle(FindAngle(Player(Playerj).X, Player(Playerj).Y, Player(i).X, Player(i).Y))
                    
                    If Abs(CurAngle - Player(Playerj).Facing) < Abs(SmallestAngle - Player(Playerj).Facing) Then
                        SmallestAngle = CurAngle
                        SmallestAngleNo = i
                    End If
                End If
                
            End If
        End If
    End If
Next i

FindLeastDegreeTarget = SmallestAngleNo

End Function

Private Sub Do2ndryFire(ByVal i As Integer)

If (Player(i).State And Player_Secondary) = Player_Secondary Then
    
    If (Player(i).State And Player_Shielding) = 0 Then
        
        If Player(i).ShipType <> MotherShip And Player(i).ShipType <> Wraith And Player(i).ShipType <> Infiltrator And _
                Player(i).ShipType <> SD Then
            
            If Player(i).LastSecondary + Missile_Delay / modSpaceGame.sv_GameSpeed <= GetTickCount() Then
                Call FireMissile(i)
            End If
            
        ElseIf Player(i).ShipType = Wraith Then
            'fire forward cannons
            If Player(i).LastSecondary + Wraith2_Delay / modSpaceGame.sv_GameSpeed <= GetTickCount() Then
                Call FireWraith2(i)
            End If
            
        'ElseIf Player(i).ShipType = Infiltrator Then
            
            'If Player(i).Shields >= 100 Then
                'AddPlayerState Player(i).ID, Player_Secondary
            'End If
            
            
        ElseIf Player(i).ShipType <> Infiltrator And Player(i).ShipType <> SD Then 'aka MS
            
            'SHIPTYPE=MS
            
            If Player(i).LastSecondary + MS2_Delay / modSpaceGame.sv_GameSpeed <= GetTickCount() Then
                Call FireMS2(i)
            End If
            
            
        End If
        
    End If 'shielding endif
    
    If Player(i).ShipType <> SD Then
        If Player(i).ShipType <> Infiltrator Then
            If Player(i).LastSecondary + MissileKeyReleaseDelay <= GetTickCount() Then
                SubPlayerState Player(i).ID, Player_Secondary 'release the "virtual key"
            End If
        End If
    End If
    
End If


End Sub

Private Sub FireWraith2(ByVal i As Integer)

Dim X1 As Single, Y1 As Single
Dim ASpeed As Single, AHeading As Single
Dim Target As Integer

Player(i).LastSecondary = GetTickCount()

X1 = Player(i).X '+ SHIP_Height * 1.5 * sine(Player(i).Heading) ' - ProngOffSet)
Y1 = Player(i).Y '- SHIP_Height * 1.5 * cosine(Player(i).Heading) ' - ProngOffSet)
'x2 = Player(i).X + SHIP_Height * 1.5 * sine(Player(i).Heading + ProngOffSet)
'y2 = Player(i).Y - SHIP_Height * 1.5 * cosine(Player(i).Heading + ProngOffSet)

'----------
Target = FindClosestTarget_ID(Player(i).X, Player(i).Y, Player(i).ID)

If Target <> -1 Then
    Target = FindPlayer(Target)
    
    AccurateShot Player(Target).X, Player(Target).Y, Player(Target).Speed, Player(Target).Heading, _
        Player(i).X, Player(i).Y, Player(i).Speed, Player(i).Heading, _
        Wraith_Bullet_Speed, 0!, AHeading '! = single ($=str)
    
Else
    AHeading = Player(i).Heading
End If
'----------


AddBullet X1, Y1, Wraith_Bullet_Speed, AHeading, Player(i).ID, _
    Player(i).Colour, sv_Bullet_Damage * 1.5, i, False

End Sub

Private Sub FireMS2(ByVal i As Integer)

Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single

Player(i).LastSecondary = GetTickCount()

X1 = Player(i).X - SHIP_RADIUS * 2
Y1 = Player(i).Y

AddBullet X1, Y1, BULLET_SPEED / 1.5, Player(i).Facing, Player(i).ID, _
    Player(i).Colour, sv_Bullet_Damage * MS_Bullet_Damage_Factor, i

End Sub

Private Sub FireMissile(ByVal i As Integer)
Dim TempMag As Single, TempDir As Single

AddVectors Player(i).Speed, Player(i).Heading, Missile_SPEED, Player(i).Facing, TempMag, TempDir

AddMissile Player(i).X, Player(i).Y, TempMag, TempDir, Player(i).ID, Player(i).Colour, _
    FindLeastDegreeTarget(i), Player(i).Facing 'FindClosestTarget_ID(MouseX, MouseY, MyID)


'If Player(i).ShipType = Raptor Then
    'recharge missiles twice as quickly
    'Player(i).LastSecondary = GetTickCount() - Missile_Delay * 2 / (3 * modSpaceGame.sv_GameSpeed)
'Else
Player(i).LastSecondary = GetTickCount()
'End If


End Sub

Private Sub AddMissile(X As Single, Y As Single, Speed As Single, _
    Heading As Single, OwnerID As Integer, Col As Long, Targeti As Integer, _
    Facing As Single)

ReDim Preserve Missiles(NumMissiles)
Missiles(NumMissiles).Decay = GetTickCount() + Missile_Decay / modSpaceGame.sv_GameSpeed
Missiles(NumMissiles).Heading = Heading
Missiles(NumMissiles).Speed = Speed
Missiles(NumMissiles).X = X
Missiles(NumMissiles).Y = Y
Missiles(NumMissiles).Owner = OwnerID
Missiles(NumMissiles).Colour = Col
Missiles(NumMissiles).Facing = Facing
Missiles(NumMissiles).Hull = Missile_Hull_Start

If Targeti <> -1 Then
    Missiles(NumMissiles).TargetID = Player(Targeti).ID
Else
    Missiles(NumMissiles).TargetID = -1
End If

picMain.DrawWidth = Thick * 2
gCircle Missiles(NumMissiles).X + BULLET_LEN * Sine(Missiles(NumMissiles).Heading), _
    Missiles(NumMissiles).Y - BULLET_LEN * CoSine(Missiles(NumMissiles).Heading), _
    Bullet_Radius, vbRed

picMain.DrawWidth = Thin

NumMissiles = NumMissiles + 1

End Sub

Private Sub SendShipTypes()

Static LastSend As Long
Dim i As Integer


If LastSend + ShipTypeSendDelay < GetTickCount() Then
    
    For i = 0 To NumPlayers - 1
        SendBroadcast sShipTypes & CStr(Player(i).ShipType) & CStr(Player(i).ID)
    Next i
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub SendTeams()

Static LastSend As Long
Dim i As Integer

If LastSend + TeamSendDelay < GetTickCount() Then
    
    For i = 0 To NumPlayers - 1
        SendBroadcast sTeams & CStr(Player(i).Team) & CStr(Player(i).ID)
    Next i
    
    LastSend = GetTickCount()
End If
'Else
'    If LastSend + TeamSendDelay < GetTickCount() Then
'
'        SendPacket socket, ServerSockAddr, sTeams & CStr(Player(0).Team) & CStr(Player(0).ID)
'
'        LastSend = GetTickCount()
'    End If

End Sub

Private Sub SendScores()

Static LastSend As Long
Dim i As Integer

If LastSend + ScoreUpdateDelay < GetTickCount() Then
    
    For i = 0 To NumPlayers - 1
        CalculateScore i
        SendBroadcast sScoreUpdates & CStr(Player(i).ID) & "|" & CStr(Player(i).Score)
    Next i
    
    LastSend = GetTickCount()
End If

End Sub

Private Sub CalculateScore(Playeri As Integer)

Player(Playeri).Score = Player(Playeri).Kills - Player(Playeri).Deaths

End Sub

Private Sub ReceiveScoreUpdate(ByVal Str As String) 'for a single player
Dim ID As Integer, j As Integer
Dim SC As Integer

On Error GoTo EH
SC = CInt(Mid$(Str, InStr(1, Str, "|", vbTextCompare) + 1))
ID = CInt(Left$(Str, InStr(1, Str, "|", vbTextCompare) - 1))

j = FindPlayer(ID)

Player(j).Score = SC

EH:
End Sub

Private Sub ReceiveShipTypes(ByVal Str As String)
Dim ID As Integer, j As Integer
Dim Tp As eShipTypes

On Error GoTo EH
ID = CInt(Mid$(Str, 2))
Tp = CInt(Left$(Str, 1))

j = FindPlayer(ID)

Player(j).ShipType = Tp

EH:
End Sub

Private Sub ReceiveTeam(ByVal Str As String)
Dim ID As Integer, j As Integer
Dim Tm As eTeams

On Error GoTo EH
ID = CInt(Mid$(Str, 2))
Tm = CInt(Left$(Str, 1))

j = FindPlayer(ID)

Player(j).Team = Tm

EH:
End Sub

Private Sub CheckScores()
Static LastScoreCheck As Long
Dim MaxScore As Integer

Dim i As Integer

If LastScoreCheck + ScoreCheckDelay < GetTickCount() Then
    
    MaxScore = modSpaceGame.sv_ScoreReq 'ScoreToGet()
    
    For i = 0 To NumPlayers - 1
        If Player(i).Score >= MaxScore Then
            RoundWinnerID = Player(i).ID
            StopPlay True
            Exit For
        End If
    Next i
    
    LastScoreCheck = GetTickCount()
End If

End Sub

Private Sub StopPlay(ByVal bStop As Boolean)
Dim i As Integer
'Static LastGameSpeed As Single

Playing = Not bStop

RandomizePlayer

If bStop Then
    
    If modSpaceGame.SpaceServer Then
        SendBroadcast sEndRounds & CStr(RoundWinnerID) 'Player(i).ID
    End If
    
    
    RoundPausedAtThisTime = GetTickCount()
    
    For i = 0 To NumBullets - 1
        RemoveBullet 0, True, 0, 0
    Next i
    
    For i = 0 To NumPlayers - 1
        CalculateScore i
    Next i
    
    'MoveCameraX -MaxWidth / 2
    'MoveCameraY -MaxHeight / 2
    
    
    'LastGameSpeed = modSpaceGame.sv_GameSpeed
    
    'modSpaceGame.sv_GameSpeed = 0.1
    
    'AddExplosion Width / 2 + Width / 4, height / 2, 200, 7
    
Else
    'reset keys
    KeyW = False
    KeyA = False
    keys = False
    KeyD = False
    KeyFire = False
    KeySecondary = False
    KeyShield = False
    
    
    'reset all players' scores
    For i = 0 To NumPlayers - 1
        With Player(i)
            .Kills = 0
            .Deaths = 0
            .State = Player_None
            
            .LastSecondary = 0
            
            '.Team = Neutral
            If .ShipType = MotherShip Or .ShipType = Wraith Or .ShipType = Infiltrator _
                Or .ShipType = SD Then
                
                .ShipType = Raptor
            End If
            
            .Alive = True
            
        End With
    Next i
    
    
'    modSpaceGame.sv_GameSpeed = LastGameSpeed
'
'    If modSpaceGame.GameOptionFormLoaded Then
'        frmGameOptions.sldrSpeed.Value = modSpaceGame.sv_GameSpeed * 10
'    End If
    
    'reset private stuff
    'MyScore = 0
    KillsInARow = 0
    MissilesShot = 0
    
    For i = 0 To UBound(ShipScores)
        ShipScores(i) = 0
    Next i
    
    'remove chat
    Erase Chat
    NumChat = 0
    
    Erase Missiles
    NumMissiles = 0
    
    Erase Circs
    NumCircs = 0
    
    
    'reset MS stuff
    MotherShipAvail = False
    MSStartFire = 0
    WraithAvail = False
    InfilAvail = False
    SDAvail = False
    
    If modSpaceGame.GameOptionFormLoaded Then
        With frmGameOptions.optnShipType(eShipTypes.MotherShip)
            
            If .Value = True Then
                .Value = False
                frmGameOptions.optnShipType(eShipTypes.Raptor).Value = True
            End If
            
            .Enabled = False
        End With
        
        With frmGameOptions.optnShipType(eShipTypes.Wraith)
            
            If .Value = True Then
                .Value = False
                frmGameOptions.optnShipType(eShipTypes.Raptor).Value = True
            End If
            
            .Enabled = False
        End With
        
        
        With frmGameOptions.optnShipType(eShipTypes.Infiltrator)
            
            If .Value = True Then
                .Value = False
                frmGameOptions.optnShipType(eShipTypes.Raptor).Value = True
            End If
            
            .Enabled = False
        End With
        
        With frmGameOptions.optnShipType(eShipTypes.SD)
            
            If .Value = True Then
                .Value = False
                frmGameOptions.optnShipType(eShipTypes.Raptor).Value = True
            End If
            
            .Enabled = False
        End With
        
        'frmGameOptions.optnTeam(eTeams.Neutral).Value = True
    End If
    
    If modSpaceGame.SpaceServer Then Call RndAsteroid
    
    Do
        Player(0).X = Rnd() * MaxWidth
        Player(0).Y = Rnd() * MaxHeight
    Loop Until PlayerIsInAsteroid(0) = False
    
    'show where i am
    'LastRespawn = GetTickCount()
    
    'Call ResetCamera
    
End If

End Sub

Private Sub DrawStars()
Dim i As Integer
Dim X1 As Single, Y1 As Single

picMain.DrawWidth = 1
picMain.FillStyle = vbSolid

picMain.FillColor = vbWhite
For i = 0 To NUM_STARS - 1
    gCircle Stars(i).X, Stars(i).Y, STAR_RADIUS, vbWhite
Next i

'For i = 0 To Num_Planets - 1
'
'    picMain.DrawWidth = 5
'
'    If (i Mod 2) = 0 Then
'        picMain.FillColor = vbBlue
'
'        gCircle (Planets(i).x, Planets(i).y), Planet_Radius, vbBlue
'
'        picMain.FillColor = vbGreen
'
'        X1 = Planets(i).x + Planet_Radius * 0.707  'sine(45)
'        Y1 = Planets(i).y - Planet_Radius * 0.707 'cosine(45)
'
'        'y1 = Planets(i).y '- 0.85 'sine(45)
'        'y2 = Planets(i).y '- 0.525 'cosine(45)
'
'        gline X1, Y1,Planets(i).x, Planets(i).y), vbGreen
'
'        picMain.DrawWidth = 4
'        gCircle (Planets(i).x - 10, Planets(i).y), Planet_Radius / 2, vbGreen
'    Else
'        picMain.FillColor = MOrange
'        gCircle (Planets(i).x, Planets(i).y), Planet_Radius, MOrange
'    End If
'
'Next i

'picMain.DrawWidth = 4
'picMain.FillColor = vbYellow
'With Sun
'    gCircle (.x, .y), Sun_Radius, vbYellow
'End With
'
'picMain.FillStyle = vbTransparent


End Sub

Private Sub ProcessStars()
Dim i As Integer

Const mWidth = MaxWidth + StarBarrier, mHeight = MaxHeight + StarBarrier
Dim SinHeading As Single, CosHeading As Single

If modSpaceGame.cg_Stars3D Then
    SinHeading = Sine(Player(0).Heading)
    CosHeading = CoSine(Player(0).Heading)
    
    For i = 0 To UBound(Stars)
        'Move the stars according to their relative speeds (w.r.t. the inverse of the ship's speed)
        
        
        'Motion Stars(i).X, Stars(i).Y, Stars(i).RelSpeed, piD2
        
        Stars(i).X = Stars(i).X - Player(0).Speed * Stars(i).RelSpeed * SinHeading '* Star_Speed
        Stars(i).Y = Stars(i).Y + Player(0).Speed * Stars(i).RelSpeed * CosHeading '* Star_Speed
        
        'Wrap the stars at the edges
        If Stars(i).X > mWidth Then Stars(i).X = -StarBarrier
        If Stars(i).Y > mHeight Then Stars(i).Y = -StarBarrier
        If Stars(i).X < -StarBarrier Then Stars(i).X = mWidth
        If Stars(i).Y < -StarBarrier Then Stars(i).Y = mHeight
        
    Next i
Else
    
    For i = 0 To UBound(Stars)
        'Move the stars according to their relative speeds (w.r.t. the inverse of the ship's speed)
        
        
        Motion Stars(i).X, Stars(i).Y, Stars(i).RelSpeed, piD2
        
        'Wrap the stars at the edges
        If Stars(i).X > mWidth Then Stars(i).X = -StarBarrier
        If Stars(i).Y > mHeight Then Stars(i).Y = -StarBarrier
        If Stars(i).X < -StarBarrier Then Stars(i).X = mWidth
        If Stars(i).Y < -StarBarrier Then Stars(i).Y = mHeight
        
    Next i
End If


'For i = 0 To Num_Planets - 1
'    Motion Planets(i).x, Planets(i).y, Planets(i).RelSpeed, pi3D4
'
'    'Wrap the stars at the edges of the window
'    If Planets(i).x > MaxWidth Then Planets(i).x = 0
'    If Planets(i).y > MaxHeight Then Planets(i).y = 0
'    If Planets(i).x < 0 Then Planets(i).x = MaxWidth
'    If Planets(i).y < 0 Then Planets(i).y = MaxHeight
'Next i
'
'
'With Sun
'    Motion .x, .y, .RelSpeed, piD4
'
'    'Wrap the stars at the edges of the window
'    If .x > MaxWidth Then .x = 0
'    If .y > MaxHeight Then .y = 0
'    If .x < 0 Then .x = MaxWidth
'    If .y < 0 Then .y = MaxHeight
'End With


End Sub

Private Sub InitStars()
Dim i As Integer

For i = 0 To NUM_STARS - 1
    'Set the star's random X and Y coords
    
    Stars(i).X = Rnd() * (MaxWidth + 2 * StarBarrier) - StarBarrier
    Stars(i).Y = Rnd() * (MaxHeight + 2 * StarBarrier) - StarBarrier
    
    'Set the star's relative speed
    Stars(i).RelSpeed = ((i \ (NUM_STARS \ NUM_Star_LAYERS)) + 1) / NUM_Star_LAYERS
Next i


'For i = 0 To Num_Planets - 1
'    With Planets(i)
'        .x = Rnd() * MaxWidth
'        .y = Rnd() * MaxHeight
'        .RelSpeed = ((i \ (Num_Planets \ Num_Planet_Layers)) + 1) / Num_Planet_Layers
'    End With
'Next i
'
'With Sun
'    .x = Rnd() * MaxWidth
'    .y = Rnd() * MaxHeight
'
'    'i = Round(Rnd() * 2)
'
'    '.RelSpeed = ((i \ (Num_Planets \ Num_Planet_Layers)) + 1) / Num_Planet_Layers
'    .RelSpeed = 1
'End With

End Sub

Private Sub DrawBoxes()
'Const Ax1 As Single = 6480, Ay1 As Single = 960
'Const Ax3 As Single = 6705, Ay3 As Single = 6135

'Const Bx1 As Single = 2160, By1 As Single = 6120
'Const Bx3 As Single = 8055, By3 As Single = 6370

picMain.DrawWidth = 1
picMain.FillColor = 0

gBox ob1.Left, ob1.Top, ob1.Left + ob1.width, ob1.Top + ob1.height, BoxColour
gBoxFilled ln1.Left, ln1.Top, ln1.Left + ln1.width, ln1.Top + ln1.height, BoxColour
gBoxFilled ln2.Left, ln2.Top, ln2.Left + ln2.width, ln2.Top + ln2.height, BoxColour

'draw screen barriers
gLine 0, 0, 0, CSng(MaxHeight), vbRed
gLine 0, 0, CSng(MaxWidth), 0, vbRed
gLine CSng(MaxWidth), 0, CSng(MaxWidth), CSng(MaxHeight), vbRed
gLine 0, CSng(MaxHeight), CSng(MaxWidth), CSng(MaxHeight), vbRed

End Sub

Private Function GetGameType() As String

Select Case modSpaceGame.sv_GameType
    Case eGameTypes.DM
        GetGameType = "DeathMatch"
    Case eGameTypes.CTF
        GetGameType = "Capture the Flag"
    Case eGameTypes.Elimination
        GetGameType = "Elimination"
End Select

End Function

Public Sub ActivateShipType(ByVal SType As eShipTypes)
Dim i As Integer

i = FindPlayer(MyID)

If Player(i).ShipType <> SType Then
    Player(i).ShipType = SType
    
    If modSpaceGame.SpaceServer Then
        SendBroadcast sShipTypes & Player(i).ShipType & MyID
    Else
        modWinsock.SendPacket socket, ServerSockAddr, _
            sShipTypes & Player(i).ShipType & MyID
    End If
    
End If

End Sub

Public Sub ActivateTeam(ByVal vTeam As eTeams)
Dim i As Integer

i = FindPlayer(MyID)

If Player(i).Team <> vTeam Then
    Player(i).Team = vTeam
    
    If modSpaceGame.SpaceServer Then
        SendBroadcast sTeams & Player(i).Team & MyID
    Else
        modWinsock.SendPacket socket, ServerSockAddr, _
            sTeams & Player(i).Team & MyID
    End If
    
End If

End Sub

Private Sub ShowMainMessages()
'Const WO2 = 3935 'Width \ 2 - 100
'Const HO2 = 3235 'Height \ 2 - 100
Dim i As Integer
Dim Tmp As String
Const TextH = 480

If NumMainMessages Then
    'If Not F1Pressed Then  'And ShowMainMsg Then
    
    picMain.Font.Size = 20
    picMain.ForeColor = Player(0).Colour 'MGrey
    
    'Check if any chat texts have decayed
    Do While i < NumMainMessages
        'Is it decay time?
        If MainMessages(i).Decay < GetTickCount() Then
            RemoveMainMessage i
            i = i - 1
        End If
        'Increment the counter
        i = i + 1
    Loop
    
    
    'Display chat text
    For i = 0 To NumMainMessages - 1
        'ShowText Chat(i).Text, 7, 435 - (i * 16), vbBlack, Me.hdc
        Tmp = MainMessages(i).Text
        
        PrintFormText Tmp, CentreX - TextWidth(Tmp) / 2, i * TextH + CentreY + 1000, Player(0).Colour
    Next i
    
    picMain.Font.Size = NormalFontSize
    'End If
End If


'i = FindPlayer(MyID)
If Player(0).ShipType = eShipTypes.MotherShip Then
    
    picMain.Font.Size = BigFontSize
    picMain.ForeColor = MGrey
    
    If MSStartFire + MotherShipRechargeTime / modSpaceGame.sv_GameSpeed < GetTickCount() Then
        PrintFormText "Weapon Ready", 7, 7, MGrey
    ElseIf MSStartFire + MotherShipFireTime / modSpaceGame.sv_GameSpeed < GetTickCount() Then
        
        PrintFormText "Charging Weapon... (" & CStr(Round( _
            (MSStartFire + MotherShipRechargeTime / modSpaceGame.sv_GameSpeed - GetTickCount()) / 1000)) _
            & ")", 7, 7, MGrey
        
        
        If (Player(0).State And PLAYER_FIRE) = PLAYER_FIRE Then
            SubPlayerState Player(0).ID, PLAYER_FIRE
        End If
    ElseIf (Player(0).State And PLAYER_FIRE) <> PLAYER_FIRE Then
        'PrintFormText "Cooling Weapon", 7, 7
        AddPlayerState MyID, PLAYER_FIRE 'let them keep firing until charging
    Else
        PrintFormText "Time Left: " & _
            CStr(Round((MSStartFire + MotherShipFireTime / modSpaceGame.sv_GameSpeed - GetTickCount()) / 1000)), _
            7, 7, MGrey
        
    End If
    
    
    
    picMain.Font.Size = NormalFontSize
    
ElseIf MotherShipAvail Then
    picMain.Font.Size = BigFontSize
    picMain.ForeColor = MGrey
    
    PrintFormText "MotherShip Available!", 7, 7, MGrey
    
    picMain.Font.Size = NormalFontSize
End If

If WraithAvail And Player(0).ShipType <> Wraith Then
    picMain.Font.Size = BigFontSize
    picMain.ForeColor = MGrey
    
    PrintFormText "Wraith Available!", 7, 257, MGrey
    
    picMain.Font.Size = NormalFontSize
End If

If InfilAvail And Player(0).ShipType <> Infiltrator Then
    picMain.Font.Size = BigFontSize
    picMain.ForeColor = MGrey
    
    PrintFormText "Infiltrator Available!", 7, 507, MGrey
    
    picMain.Font.Size = NormalFontSize
End If

If SDAvail And Player(0).ShipType <> SD Then
    picMain.Font.Size = BigFontSize
    picMain.ForeColor = MGrey
    
    PrintFormText "Star Destroyer Available!", 7, 757, MGrey
    
    picMain.Font.Size = NormalFontSize
End If

If Player(0).ShipType = Infiltrator Then
    If (Player(0).State And Player_Secondary) = Player_Secondary Then
        picMain.ForeColor = Player(0).Colour
        PrintFormText "Cloak Engaged", Me.ScaleWidth / 2, 7, MGrey
    End If
ElseIf Player(0).ShipType = SD Then
    If (Player(0).State And Player_Secondary) = Player_Secondary Then
        picMain.ForeColor = Player(0).Colour
        PrintFormText "Gravity Well Engaged", Me.ScaleWidth / 2, 7, MGrey
    End If
End If

If Not modSpaceGame.SpaceServer Then
    If (LastUpdatePacket + mPacket_LAG_TOL) < GetTickCount() Then
        If LastUpdatePacket Then
            picMain.Font.Size = BigFontSize * 2
            picMain.ForeColor = MGrey
            picMain.Font.Bold = True
            
            PrintFormText "Connection Interrupted", 3900, 3000, MGrey
            
            picMain.Font.Bold = False
            picMain.Font.Size = NormalFontSize
        End If
    End If
End If


End Sub

Private Sub ShowChatEntry()

'###########
'show chat
If bChatActive Then
    'CurrentY = i * TextHeight(FinalTxt) + 3000
    picMain.ForeColor = Player(0).Colour
    
    If (LastFlash + ChatFlashDelay) < GetTickCount() Then
        bChatCursor = Not bChatCursor
        LastFlash = GetTickCount()
    End If
    
    PrintFormText Trim$(Player(0).Name) & modMessaging.MsgNameSeparator & strChat & IIf(bChatCursor, "_", vbNullString), _
        7, 2700, Player(0).Colour
    
    
End If

End Sub

Private Sub ShowScores()
Dim Str As String
Dim i As Integer

picMain.ForeColor = vbWhite 'IIf(modSpaceGame.cg_BlackBG, vbWhite, vbBlack)

If F1Pressed Then
    ShowF1Scores
Else
    ShowMainMessages
    
    DrawMap
    
    'Str = "K" & Space$(3) & "D"
    'PrintFormText Str, width - 250 - TextWidth(Str), 7
    
    For i = 0 To NumPlayers - 1
        Str = Trim$(Player(i).Name) & ": " & CStr(Player(i).Score) 'CStr(Player(i).Kills) & Space$(3) & CStr(Player(i).Deaths)
        'CurrentX = Width - 100 - TextWidth(Str)
        'CurrentY = i * TextHeight(Str) + 10
        'Print Str
        'Me.ForeColor = Player(i).Colour
        PrintFormText Str, width - 300 - TextWidth(Str), i * TextHeight(Str) + modSpaceGame.cg_MapLen + 20, Player(i).Colour
    Next i
End If

If modSpaceGame.cg_ShowFPS Then
    'Me.ForeColor = Player(0).Colour
    Str = "FPS: " & CStr(FPS)
    PrintFormText Str, 400 - TextWidth(Str) / 2, 7800 - TextHeight(Str), Player(0).Colour
    'tX = 400
    'tY = 8300
End If

End Sub

Private Sub ShowF1Scores()
Const WO2 = 1600 '5000 'Width \ 2 - 100
Const HO2 = 1700 'Height \ 2 - 100
Const Sp8 As String * 8 = "        "

Dim Txt As String
Dim i As Integer
Dim w As Single, H As Single

picMain.Font.Size = BigFontSize
'Me.ForeColor = Player(0).Colour

On Error Resume Next

Txt = "Score To Win - " & CStr(modSpaceGame.sv_ScoreReq)
PrintFormText Txt, WO2 - TextWidth(Txt) / 2 + 500, HO2 - TextHeight(Txt) * 6 - 200, Player(0).Colour

Txt = "Missiles Shot Down: " & CStr(MissilesShot)
PrintFormText Txt, WO2 - TextWidth(Txt) / 2 + 500, HO2 - TextHeight(Txt) * 5 - 200, Player(0).Colour

Txt = "GameType: " & GetGameType()
PrintFormText Txt, WO2 - TextWidth(Txt) / 2 + 500, HO2 - TextHeight(Txt) * 4 - 200, Player(0).Colour

'draw scores
Txt = "Raptor: " & ShipScores(eShipTypes.Raptor) & _
      "  Behemoth: " & ShipScores(eShipTypes.Behemoth) & _
      "  Hornet: " & ShipScores(eShipTypes.Hornet) & _
      "  Mothership: " & ShipScores(eShipTypes.MotherShip) & _
      "  Wraith: " & ShipScores(eShipTypes.Wraith) & _
      "  Infiltrator: " & ShipScores(eShipTypes.Infiltrator) & _
      "  Star Destroyer: " & ShipScores(eShipTypes.SD)
      
PrintFormText Txt, WO2 + 2000, HO2 - TextHeight(Txt) * 6 - 200, Player(0).Colour
'+ TextWidth(Txt) / 2

w = WO2
Txt = Left$(Sp8, 7) & "Name" & Sp8
PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour

'Me.DrawStyle = vbTransparent
'gline w, H,w + 500, H + 500), MGrey, BF
'Me.DrawStyle = vbSolid

For i = 0 To NumPlayers - 1
    'Me.ForeColor = Player(i).Colour
    Txt = CentreFill(Trim$(Player(i).Name), 20)
    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), Player(i).Colour
Next i


'Me.ForeColor = Player(0).Colour
w = w + 1500
Txt = " Score "
PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour

For i = 0 To NumPlayers - 1
    'Me.ForeColor = Player(i).Colour
    
    'Player(i).Score = Player(i).Kills - Player(i).Deaths
    
    Txt = CentreFill(Trim$(CStr(Player(i).Score)), 10)
    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), Player(i).Colour
Next i

'Me.ForeColor = Player(0).Colour
w = w + 1000
Txt = " Kills "
PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour

For i = 0 To NumPlayers - 1
    'Me.ForeColor = Player(i).Colour
    Txt = CentreFill(Trim$(CStr(Player(i).Kills)), 6)
    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), Player(i).Colour
Next i

'Me.ForeColor = Player(0).Colour
w = w + 1000
Txt = " Deaths "
PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour

For i = 0 To NumPlayers - 1
    'Me.ForeColor = Player(i).Colour
    Txt = CentreFill(Trim$(CStr(Player(i).Deaths)), 8)
    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), Player(i).Colour
Next i

'Me.ForeColor = Player(0).Colour
w = w + 1000
Txt = "  Team  "
PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour

For i = 0 To NumPlayers - 1
    'Me.ForeColor = GetTeamColour(Player(i).Team) 'Player(i).Colour
    Txt = CentreFill(Trim$(GetTeamStr(Player(i).Team)), 10)
    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), GetTeamColour(Player(i).Team)
Next i


w = w + 1000
Txt = "  Ship  "
PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour

For i = 0 To NumPlayers - 1
    'Me.ForeColor = GetTeamColour(Player(i).Team) 'Player(i).Colour
    Txt = GetShipName(Player(i).ShipType)
    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), Player(i).Colour
Next i


'Me.ForeColor = Player(0).Colour
'w = w + 1000
'Txt = " Ping "
'PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3
'
'For i = 0 To NumPlayers - 1
'    Me.ForeColor = Player(i).Colour
'
'    If Player(i).IsBot Or Player(i).ID = MyID Then
'        Ping = 0
'    Else
'        Ping = (GetTickCount() - Player(i).LastPacket) / 1000
'    End If
'
'    Txt = CentreFill(Trim$(CStr(Ping)), 6)
'    PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i)
'Next i

If modSpaceGame.sv_GameType = Elimination Then
    If Playing Then
        'Me.ForeColor = Player(0).Colour
        w = WO2 - 500
        Txt = "Dead? "
        PrintFormText Txt, w, HO2 - TextHeight(Txt) * 3, Player(0).Colour
        
        For i = 0 To NumPlayers - 1
            
            If PlayerInGame(i) = False Then
                'Me.ForeColor = Player(i).Colour
                Txt = "  ----  "
                PrintFormText Txt, w, HO2 - TextHeight(Txt) * (2 - i), Player(i).Colour
            End If
            
        Next i
    End If
End If

picMain.Font.Size = NormalFontSize

If Player(0).State Then
    If PlayerInGame(0) Then
        SetPlayerState MyID, Player_None
    End If
End If

'old method ---------------------------------------------------------------------------------

'Txt = InfoStart & "My Scores" & InfoEnd
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 - TextHeight(Txt)
'
'Txt = CStr(Player(0).Kills) & " Kills"
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2
'
'Txt = CStr(Player(0).Deaths) & " Deaths"
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt)
'
'If Player(0).Deaths > 0 Then
'    Txt = "Kill to Death Ratio: " & CStr(Round(Player(0).Kills / Player(0).Deaths, 2))
'ElseIf Player(0).Kills = 0 Then
'    Txt = "Kill to Death Ratio: 0"
'Else
'    Txt = "Kill to Death Ratio: Infinite!"
'End If
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 2
'
'Txt = GetTeamStr(Player(0).Team) & " Team"
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 3
'
'
'Txt = "Raptor Kills - " & CStr(ShipScores(eShipTypes.Raptor))
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 5
'
'Txt = "Behemoth Kills - " & CStr(ShipScores(eShipTypes.Behemoth))
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 6
'
'Txt = "Hornet Kills - " & CStr(ShipScores(eShipTypes.Hornet))
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 7
'
'Txt = "MotherShip Kills - " & CStr(ShipScores(eShipTypes.MotherShip))
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 8
'
'Txt = "Wraith Kills - " & CStr(ShipScores(eShipTypes.Wraith))
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 9
'
'Txt = "Infiltrator Kills - " & CStr(ShipScores(eShipTypes.Infiltrator))
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 10
'
'Txt = InfoStart & "Enemy Scores" & InfoEnd
'PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + TextHeight(Txt) * 12
'
''show the rest of the players
'For i = 1 To NumPlayers - 1
'
'    On Error Resume Next
'
'    If Player(i).Deaths > 0 Then
'        Txt = "Ratio: " & CStr(Round(Player(i).Kills / Player(i).Deaths, 2))
'    ElseIf Player(0).Kills = 0 Then
'        Txt = "Ratio: 0"
'    Else
'        Txt = "Ratio: Infinite!"
'    End If
'
'    Me.ForeColor = Player(i).Colour
'
'    Txt = Trim$(Player(i).Name) & " - " & CStr(Player(i).Kills) & _
'        " : " & CStr(Player(i).Deaths) & Space$(1) & Txt & " Team: " & GetTeamStr(Player(i).Team)
'
'    PrintFormText Txt, WO2 - TextWidth(Txt) / 2, HO2 + (i + 12) * TextHeight(Txt)
'
'Next i

End Sub

Private Function GetShipName(St As eShipTypes) As String
If St = Wraith Then
    GetShipName = "Wraith"
ElseIf St = Behemoth Then
    GetShipName = "Behemoth"
ElseIf St = Hornet Then
    GetShipName = "Hornet"
ElseIf St = Infiltrator Then
    GetShipName = "Infiltrator"
ElseIf St = MotherShip Then
    GetShipName = "Mothership"
ElseIf St = Raptor Then
    GetShipName = "Raptor"
ElseIf St = SD Then
    GetShipName = "Star Destroyer"
End If
End Function

Private Sub ShowRoundScores()
Const w = 4000
Const H = 2000
Dim RoundTm As Long
Dim RoundWinneri As Integer
Dim Str As String

'Static XProg As Single ', YProg As Single 'for smoke
'Const sY As Single = CentreY + 3000
'Static LastAdd As Long

On Error Resume Next

ShowF1Scores

'########
picMain.Font.Size = BigFontSize
'picMain.ForeColor = MGrey
picMain.Font.Underline = True
picMain.Font.Bold = True

Str = "Round is Over"
PrintFormText Str, 5000 - TextWidth(Str) / 2, 725, MGrey

picMain.Font.Underline = False
picMain.Font.Bold = False
'########

RoundWinneri = FindPlayer(RoundWinnerID)

If RoundWinneri <> -1 Then
    
    'Me.ForeColor = Player(RoundWinneri).Colour
    
    Str = "Round Winner - " & Trim$(Player(RoundWinneri).Name) & Space$(1) & _
        Player(RoundWinneri).Kills & " Kills, " & Player(RoundWinneri).Deaths & " Deaths"
    PrintFormText Str, 9500 - TextWidth(Str) / 2, 1000, Player(RoundWinneri).Colour
    '--------
    If (Player(RoundWinneri).Team = Neutral Or Player(RoundWinneri).Team = Spec) = False Then
        
        'Me.ForeColor = IIf(Player(RoundWinneri).Team = Red, vbRed, _
                        IIf(Player(RoundWinneri).Team = Blue, vbBlue, _
                        Player(RoundWinneri).Colour))
        
        Str = "Winning Team - " & GetTeamStr(Player(RoundWinneri).Team)
        PrintFormText Str, 9500 - TextWidth(Str) / 2, 1400, _
                        IIf(Player(RoundWinneri).Team = Red, vbRed, _
                        IIf(Player(RoundWinneri).Team = Blue, vbBlue, _
                        Player(RoundWinneri).Colour))
        
    End If
    '--------
End If

'Me.ForeColor = MGrey

RoundTm = RoundPausedAtThisTime + RoundWaitTime - GetTickCount()
'Str = "Round will begin in " & Format$(RoundTm / 1000, "0.0") & " seconds"
Str = "Round will begin in " & CStr(Round(RoundTm / 1000)) & " seconds"
PrintFormText Str, 9500 - TextWidth(Str) / 2, 1800, MGrey


'decide if new round
If modSpaceGame.SpaceServer Then
    If RoundTm <= 0 Then
        SendBroadcast sNewRounds
        Pause 10 'allow them to reset their kills + deaths etc
        StopPlay False
    End If
End If

picMain.Font.Size = NormalFontSize

'XProg = XProg + 100 * modSpaceGame.sv_GameSpeed
'If XProg > Me.width Then
'    XProg = 0
'End If
''YProg = YProg + 100
''If YProg > Me.height Then
''    YProg = 0
''End If
'
'If LastAdd + 20 / modSpaceGame.sv_GameSpeed < GetTickCount() Then
'    Call AddSmokeGroup(XProg, sY, 2)
'    Call AddSmokeGroup(Me.width - XProg, sY, 2)
'    LastAdd = GetTickCount()
'End If

End Sub

Private Sub ProcessAI(ByVal ID As Integer, Optional ByVal Detail As Boolean = False) ', _
    Optional ByVal DoThrust As Boolean = False)

Const w = 4000 'Width \ 2 - 100
Const H = 7900 'Height \ 2 - 100
'Const AIFacingAdjust = 5 * pi / 18

Dim ShipIndex As Integer
Dim Sample_Rate As Integer
'Dim wX As Single, wY As Single

On Error GoTo EH

ShipIndex = -1

ShipIndex = FindPlayer(ID)

If ShipIndex <> -1 Then
    
    
    If modSpaceGame.SpaceServer Then
        If Player(ShipIndex).IsBot Then
            Sample_Rate = AI_Sample_Rate
        Else
            Sample_Rate = modSpaceGame.Default_AI_Sample_Rate
        End If
    Else
        Sample_Rate = modSpaceGame.Default_AI_Sample_Rate
    End If
    
    'turn them
    
    'Player(ShipIndex).AIWantToFace = FixAngle(Player(ShipIndex).AIWantToFace)
    'Player(ShipIndex).Facing = FixAngle(Player(ShipIndex).Facing)
    
'    If Round(Player(ShipIndex).Facing, 3) <> Round(Player(ShipIndex).AIWantToFace, 3) Then
'
'        Player(ShipIndex).Facing = Player(ShipIndex).Facing + (Player(ShipIndex).AIWantToFace - Player(ShipIndex).Facing) * AIFacingAdjust
'
'
''        If Player(ShipIndex).AIWantToFace > pi And Player(ShipIndex).Facing < pi Then
''            Player(ShipIndex).Facing = Player(ShipIndex).Facing + AIFacingAdjust * modSpaceGame.sv_GameSpeed
''        Else
''            Player(ShipIndex).Facing = Player(ShipIndex).Facing - AIFacingAdjust * modSpaceGame.sv_GameSpeed
''        End If
'
''        If Abs(Player(ShipIndex).AIWantToFace - Player(ShipIndex).Facing) < Player(ShipIndex).AIWantToFace Then
''            Player(ShipIndex).Facing = Player(ShipIndex).Facing + AIFacingAdjust * modSpaceGame.sv_GameSpeed
''        Else
''            Player(ShipIndex).Facing = Player(ShipIndex).Facing - AIFacingAdjust * modSpaceGame.sv_GameSpeed
''        End If
'
'
'    End If
    
    If PlayerInGame(ShipIndex) Then
        If ((Player(ShipIndex).AITimer + Sample_Rate) < GetTickCount()) Then
            
            Call pDoAI(ShipIndex, ID, Player(ShipIndex).LastAITargetIndex) ', bSeekingPowerUp)
            
            Player(ShipIndex).AITimer = GetTickCount()
            
        End If
    End If
    
    
    'draw where AI wants to face
'    wX = Player(ShipIndex).x + BULLET_LEN * sine(Player(ShipIndex).AIWantToFace)
'    wY = Player(ShipIndex).y - BULLET_LEN * cosine(Player(ShipIndex).AIWantToFace)
'
'    picMain.DrawWidth = Thin * 2
    'gline Player(ShipIndex).x, Player(ShipIndex).y,wX, wY), vbBlue
    
    
    If Detail Then
        If modSpaceGame.UseAI Then
            If Player(ShipIndex).LastAITargetIndex <> -1 Then
                'Me.ForeColor = Player(Player(ShipIndex).LastAITargetIndex).Colour
                
                PrintFormText "AI - Target: " & Trim$(Player(Player(ShipIndex).LastAITargetIndex).Name), _
                    w, H, Player(Player(ShipIndex).LastAITargetIndex).Colour
                'If Trim$(Player(Target).Name) = "Rob" Then Stop
            Else
                'Me.ForeColor = Player(ShipIndex).Colour
'                If bSeekingPowerUp = False Then
                PrintFormText "AI - No Target Found...", w, H, Player(ShipIndex).Colour
'                Else
'                    PrintFormText "AI - Seeking Powerup", w, H
'                End If
            End If
        End If
    End If
    
End If

EH:
End Sub

Private Sub pDoAI(ByRef ShipIndex As Integer, ByRef ID As Integer, ByRef Target As Integer) ', _
    ByRef SeekingPowerUp As Boolean)

'variables
Dim ASpeed As Single, AHeading As Single
Dim sngHeadingForPowerUp As Single
'Dim bTmp As Boolean

'helper flags
Dim bPowerUpClose As Boolean, bHeadingToPowerUp As Boolean, bEnemyClose As Boolean, bLowHull As Boolean, bVLowHull As Boolean
Dim bHasFlag As Boolean, bInBase As Boolean 'CTF only

'act-on flag
Dim bSeekPowerUp As Boolean, bShootTarget As Boolean, bThrust As Integer, bRThrust As Boolean, bShieldsUp As Boolean '1=true, 0=false, -1=revthrust
Dim bFaceBase As Boolean, bGoSlow As Boolean

Const BackOffDist = 5000
Const PowerUpScanDist = 2500
Const TooCloseDist = 1000
Const bsD15 = BULLET_SPEED / 1.5

Target = FindPlayer(FindClosestTarget_ID(Player(ShipIndex).X, Player(ShipIndex).Y, ID))


If PowerUp.Active Then
    'PowerUpHeading = FindAngle(Player(ShipIndex).x, Player(ShipIndex).y, PowerUp.x, PowerUp.y)
    If GetDist(Player(ShipIndex).X, Player(ShipIndex).Y, PowerUp.X, PowerUp.Y) < PowerUpScanDist Then
        bPowerUpClose = True
    End If
    
    'AccurateShot PowerUp.x, PowerUp.y, 0, 0, Player(ShipIndex).x, Player(ShipIndex).y, Player(ShipIndex).Speed, _
        Player(ShipIndex).Heading, Player(ShipIndex).Speed, ASpeed, sngHeadingForPowerUp
    
    sngHeadingForPowerUp = FindAngle(Player(ShipIndex).X, Player(ShipIndex).Y, PowerUp.X, PowerUp.Y)
    
    'ASpeed = 0
    sngHeadingForPowerUp = FixAngle(sngHeadingForPowerUp)
    
    If Round(Player(ShipIndex).Heading, 1) = Round(sngHeadingForPowerUp, 1) Then
        If Player(ShipIndex).Speed > 1 Then
            bHeadingToPowerUp = True
        End If
    End If
    
End If

If Target <> -1 Then
    bEnemyClose = (GetDist(Player(ShipIndex).X, Player(ShipIndex).Y, Player(Target).X, Player(Target).Y) < TooCloseDist)
End If

If Player(ShipIndex).Hull < (Player(ShipIndex).MaxHull * 0.5) Then 'decide whether to retreat
    If Player(ShipIndex).Shields < (Player(ShipIndex).MaxShields * 0.3) Then
        bLowHull = True
        If GetDist(Player(ShipIndex).X, Player(ShipIndex).Y, Player(Target).X, Player(Target).Y) < BackOffDist Then
            'retreat
            bVLowHull = (Player(ShipIndex).Hull < 0.2 * Player(ShipIndex).MaxHull)
        End If
    End If
End If

If modSpaceGame.sv_GameType = CTF Then
    
    With Player(ShipIndex)
        
        If .ID = FlagOwnerID Then
            bHasFlag = True
            
            bInBase = (GetDist(.X, .Y, FlagBaseX, FlagBaseY) < (GetShipRadius(.ShipType) + FlagBaseRadius / 2))
        End If
        
    End With
    
End If

If bHasFlag Then
    If bInBase = False Then
        bFaceBase = True
        
        bSeekPowerUp = False
        bShootTarget = False
        bThrust = 1
        bShieldsUp = True
    Else
        'in base...
        bFaceBase = True
        bGoSlow = True
        
        bSeekPowerUp = False
        bShootTarget = False
        'bThrust = 1 'IIf(Player(ShipIndex).Speed > 40, -1, 0)
        bShieldsUp = True
        
        'cheat slightly
        'Player(ShipIndex).Speed = 0
        bThrust = 0
        
    End If
    
ElseIf bPowerUpClose And bHeadingToPowerUp Then
    bSeekPowerUp = (Player(ShipIndex).Speed = 0)
    bShootTarget = True
    bThrust = 0
    bShieldsUp = False
    
ElseIf bPowerUpClose Then
    bSeekPowerUp = True
    bShootTarget = False
    bThrust = 1
    bShieldsUp = True
    
ElseIf Target <> -1 Then
    
    If bVLowHull Then 'attack!
        
        bThrust = 1 'ram!
        bRThrust = False
        bShieldsUp = False
        bShootTarget = True
        
    ElseIf bLowHull Then
        
        bThrust = -1 'retreat!
        bRThrust = Not bEnemyClose 'do a bit of long range evasion
        bShieldsUp = bEnemyClose 'if enemy is close, stick the shields up
        bShootTarget = Not bShieldsUp
        
    Else 'normal attack
        
        'accelerate to half bullet_speed
        bThrust = IIf(Player(ShipIndex).Speed > bsD15, 0, 1)
        
        bShootTarget = True
        bShieldsUp = False
        bRThrust = False
        
    End If
    
    If bThrust = 0 Then bThrust = IIf(bEnemyClose, -1, 1)
    
    
ElseIf bEnemyClose Then
    
    If bLowHull Then
        bSeekPowerUp = False 'ram them!
        bShootTarget = True
        bShieldsUp = False
        bThrust = -1
    Else
        bSeekPowerUp = False 'ram them!
        bShootTarget = False
        bShieldsUp = True
        bThrust = 1
    End If
    
Else
    'no target found...
    
    bSeekPowerUp = PowerUp.Active
    bShootTarget = False
    bShieldsUp = True
    bThrust = IIf(bSeekPowerUp, 1, 0)
    
End If

If Target <> -1 Then
    'fire a missile
    If Player(ShipIndex).LastSecondary + Missile_Delay / modSpaceGame.sv_GameSpeed <= GetTickCount() Then
        AddPlayerState ID, Player_Secondary
        
        'drop shields for a second to fire
        If (Player(ShipIndex).State And Player_Shielding) = Player_Shielding Then
            SubPlayerState Player(ShipIndex).ID, Player_Shielding
        End If
        
    End If
End If

If Player(ShipIndex).ShipType = SD Then
    If (Player(ShipIndex).State And Player_Secondary) = 0 Then
        AddPlayerState Player(ShipIndex).ID, Player_Secondary
    End If
End If


If bSeekPowerUp Then
    'face the powerup
    Player(ShipIndex).Facing = sngHeadingForPowerUp
    
    bShootTarget = False
End If

If bShootTarget Then
    'face target
    
    AccurateShot Player(Target).X, Player(Target).Y, Player(Target).Speed, Player(Target).Heading, _
                Player(ShipIndex).X, Player(ShipIndex).Y, Player(ShipIndex).Speed, Player(ShipIndex).Heading, _
                BULLET_SPEED, ASpeed, AHeading
    
    Player(ShipIndex).Facing = AHeading
    
    If bThrust = 1 Then
        If Player(ShipIndex).Speed > ASpeed Then
            bThrust = 0
        End If
    End If
    
    If (Player(ShipIndex).State And PLAYER_FIRE) = 0 Then
        AddPlayerState ID, PLAYER_FIRE
    End If
    
    If (Player(ShipIndex).State And Player_Shielding) = Player_Shielding Then
        SubPlayerState Player(ShipIndex).ID, Player_Shielding
    End If
    
ElseIf (Player(ShipIndex).State And PLAYER_FIRE) = PLAYER_FIRE Then
    SubPlayerState ID, PLAYER_FIRE
End If


If bShieldsUp Then
    If (Player(ShipIndex).State And Player_Shielding) = 0 Then
        AddPlayerState Player(ShipIndex).ID, Player_Shielding
    End If
Else
    If (Player(ShipIndex).State And Player_Shielding) = Player_Shielding Then
        SubPlayerState Player(ShipIndex).ID, Player_Shielding
    End If
End If

'ctf shizcakes
If bFaceBase Then
    Player(ShipIndex).Facing = FindAngle(Player(ShipIndex).X, Player(ShipIndex).Y, FlagBaseX, FlagBaseY)
End If


'thrust
If bThrust = 0 Then
    'remove thrust
    If (Player(ShipIndex).State And PLAYER_THRUST) = PLAYER_THRUST Then
        SubPlayerState ID, PLAYER_THRUST
    End If
    
    If (Player(ShipIndex).State And PLAYER_REVTHRUST) = PLAYER_REVTHRUST Then
        SubPlayerState ID, PLAYER_REVTHRUST
    End If
    
ElseIf bThrust = 1 And Not (bGoSlow And Abs(Player(ShipIndex).Speed) > 40) Then
    'add thrust
    If (Player(ShipIndex).State And PLAYER_THRUST) <> PLAYER_THRUST Then
        AddPlayerState ID, PLAYER_THRUST
    End If
    If (Player(ShipIndex).State And PLAYER_REVTHRUST) = PLAYER_REVTHRUST Then
        SubPlayerState ID, PLAYER_REVTHRUST
    End If
Else
    
    'If bGoSlow Then
        'If Player(ShipIndex).Speed < 10 Then
            'bTmp = True
        'End If
    'End If
    
    'If Not bTmp Then
        'reverse thrust
    If (Player(ShipIndex).State And PLAYER_THRUST) = PLAYER_THRUST Then
        SubPlayerState ID, PLAYER_THRUST
    End If
    
    If (Player(ShipIndex).State And PLAYER_REVTHRUST) <> PLAYER_REVTHRUST Then
        AddPlayerState ID, PLAYER_REVTHRUST
    End If
    'End If
    
End If


If bRThrust Then
    If (Player(ShipIndex).State And Player_StrafeRight) <> Player_StrafeRight Then
        AddPlayerState ID, Player_StrafeRight
    End If
ElseIf (Player(ShipIndex).State And Player_StrafeRight) = Player_StrafeRight Then
    SubPlayerState ID, Player_StrafeRight
End If


End Sub

Private Function FindClosestTarget_ID(ByVal fX As Single, ByVal fY As Single, _
    ByVal NtID As Integer) As Integer

Dim Target As Integer, i As Integer
Dim Dist As Single, TmpDist As Single
Dim NtIDIndex As Integer

Target = -1
Dist = MaxWidth + 200

NtIDIndex = FindPlayer(NtID)

'find a target
For i = 0 To NumPlayers - 1
    If Player(i).ID <> NtID Then
        If IsAlly(Player(i).Team, Player(NtIDIndex).Team) = False Then
            If PlayerInGame(i) Then 'Player(i).Team <> Spec Then
                
                If Not (Player(i).ShipType = Infiltrator And (Player(i).State And Player_Secondary) = Player_Secondary) Then
                    
                    TmpDist = GetDist(fX, fY, Player(i).X, Player(i).Y)
                    If TmpDist < Dist Then
                        Target = i
                        Dist = TmpDist
                    End If
                    
                End If
                
            End If
            
        End If
        
    End If
Next i

If Target <> -1 Then
    FindClosestTarget_ID = Player(Target).ID
Else
    FindClosestTarget_ID = -1
End If

End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'modSpaceGame.SpaceEditing, bChatActive, PlayerInGame(0)

Select Case KeyCode
        
    Case vbKeyA
        KeyA = True
        
    Case vbKeyD
        KeyD = True
        
    Case vbKeyS
        keys = True
        
    Case vbKeyW
        KeyW = True
        
        
    Case vbKeyControl
        If Player(0).ShipType <> Infiltrator And Player(0).ShipType <> SD Then
            KeySecondary = True
        Else 'If KeySecondary Then
            If Player(0).State And Player_Secondary Then
                SubPlayerState MyID, Player_Secondary
            Else
                AddPlayerState MyID, Player_Secondary
            End If
        End If
        
        
    Case vbKeyShift
        KeyShield = True
        
'#####################################
        
    Case vbKeyF1
        F1Pressed = True
        
    Case vbKeyAdd
        
        If modSpaceGame.cg_Zoom < MaxZoom Then
            modSpaceGame.cg_Zoom = Round(modSpaceGame.cg_Zoom + ZoomInc, 2)
            
            'Me.Font.Size = BaseFontSize * modSpaceGame.cg_Zoom
            
        End If
        LastZoomPress = GetTickCount()
    
    Case vbKeySubtract
        
        If modSpaceGame.cg_Zoom >= MinZoom Then
            modSpaceGame.cg_Zoom = Round(modSpaceGame.cg_Zoom - ZoomInc, 2)
            
            'Me.Font.Size = BaseFontSize * modSpaceGame.cg_Zoom
            
        End If
        LastZoomPress = GetTickCount()
        
    Case vbKeyMultiply
        
        modSpaceGame.cg_Zoom = 1
        LastZoomPress = GetTickCount()
        
    Case vbKeySpace
        KeyFire = True
        
End Select

End Sub

Private Sub ProcessKeys()

Dim bCan As Boolean
Dim St As eShipTypes


If bChatActive Then Exit Sub


If KeyA Then
    If PlayerInGame(0) = False Then
        MoveCameraX modSpaceGame.cg_Camera.X - CameraInc
        
    ElseIf modSpaceGame.cl_UseMouse Then
        If (Player(0).State And Player_StrafeLeft) = 0 Then
            AddPlayerState MyID, Player_StrafeLeft
        End If
    Else
        If (Player(0).State And PLAYER_LEFT) = 0 Then
            AddPlayerState MyID, PLAYER_LEFT
        End If
    End If
Else
    If modSpaceGame.cl_UseMouse Then
        If (Player(0).State And Player_StrafeLeft) = Player_StrafeLeft Then
            SubPlayerState MyID, Player_StrafeLeft
        End If
    Else
        If (Player(0).State And PLAYER_LEFT) = PLAYER_LEFT Then
            SubPlayerState MyID, PLAYER_LEFT
        End If
    End If
End If
'#####################################
If KeyD Then
    If PlayerInGame(0) = False Then
        MoveCameraX modSpaceGame.cg_Camera.X + CameraInc
    ElseIf modSpaceGame.cl_UseMouse Then
        If (Player(0).State And Player_StrafeRight) = 0 Then
            AddPlayerState MyID, Player_StrafeRight
        End If
    Else
        If (Player(0).State And PLAYER_RIGHT) = 0 Then
            AddPlayerState MyID, PLAYER_RIGHT
        End If
    End If
Else
    If modSpaceGame.cl_UseMouse Then
        If (Player(0).State And Player_StrafeRight) = Player_StrafeRight Then
            SubPlayerState MyID, Player_StrafeRight
        End If
    Else
        If (Player(0).State And PLAYER_RIGHT) = PLAYER_RIGHT Then
            SubPlayerState MyID, PLAYER_RIGHT
        End If
    End If
End If
'#####################################
If KeyW Then
    If PlayerInGame(0) = False Then
        MoveCameraY modSpaceGame.cg_Camera.Y - CameraInc
    ElseIf (Player(0).State And PLAYER_THRUST) = 0 Then
        AddPlayerState MyID, PLAYER_THRUST
    End If
ElseIf Player(0).State And PLAYER_THRUST Then
    SubPlayerState MyID, PLAYER_THRUST
End If
'#####################################
If keys Then 'KeyS
    
    If PlayerInGame(0) = False Then
        MoveCameraY modSpaceGame.cg_Camera.Y + CameraInc
        
    ElseIf (Player(0).State And PLAYER_REVTHRUST) = 0 Then
        AddPlayerState MyID, PLAYER_REVTHRUST
    End If
    
ElseIf Player(0).State And PLAYER_REVTHRUST Then
    SubPlayerState MyID, PLAYER_REVTHRUST
End If


'#####################################
If Not PlayerInGame(0) Then Exit Sub
'#####################################


If KeyFire Then
    'AddPlayerState MyID, Player_Fire
    
    If Player(FindPlayer(MyID)).ShipType = MotherShip Then
        If MSStartFire + MotherShipRechargeTime / modSpaceGame.sv_GameSpeed < GetTickCount() Then
            bCan = True
            MSStartFire = GetTickCount()
        End If
    Else
        bCan = ((Player(0).State And PLAYER_FIRE) = 0)
    End If
    
    If bCan Then AddPlayerState MyID, ePlayerState.PLAYER_FIRE
    
ElseIf Player(0).State And PLAYER_FIRE Then
    If Player(0).ShipType <> MotherShip Then
        SubPlayerState MyID, PLAYER_FIRE
    End If
End If


If KeySecondary Then
    St = Player(0).ShipType
    
    'If St = MotherShip Or St = Wraith Then
    If (Player(0).State And Player_Secondary) = 0 Then
        AddPlayerState MyID, Player_Secondary
    End If
    'End If
    
    'If St <> Infiltrator Then
        'If St <> SD Then
            'KeySecondary = False
        'End If
    'End If
    
ElseIf Player(0).State And Player_Secondary Then
    If Player(0).ShipType <> Infiltrator Then
        If Player(0).ShipType <> SD Then
            If Player(0).ShipType = Wraith Then
                SubPlayerState MyID, Player_Secondary
            Else
                If Player(0).LastSecondary + MissileKeyReleaseDelay < GetTickCount() Then
                    SubPlayerState MyID, Player_Secondary
                ElseIf Player(0).ShipType = MotherShip Then
                    SubPlayerState MyID, Player_Secondary
                End If
            End If
        End If
    End If
    
End If


If KeyShield Then
    St = Player(0).ShipType
    
    If St <> MotherShip And St <> Infiltrator Then
        If (Player(0).State And Player_Shielding) = 0 Then
            AddPlayerState MyID, Player_Shielding
        End If
    End If
ElseIf Player(0).State And Player_Shielding Then
    SubPlayerState MyID, Player_Shielding
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'esc, bkspc,return

If KeyAscii = vbKeyTab Then
    Call Form_MouseUp(vbMiddleButton, 0, 0, 0)
    
Else
    
    'If (LenB(strChat) = 0) And (KeyAscii = 116) Then Exit Sub 'don't add t
    
    Select Case True
        
        '#########################################
        'Chat handling
        'Escape kills the chat
        Case KeyAscii = vbKeyEscape
            bChatActive = False
            strChat = vbNullString
        
        Case (KeyAscii = 116) And (bChatActive = False)
            '116=t
            bChatActive = True
            
            'Backspace removes a character
        Case KeyAscii = vbKeyBack
            If Len(strChat) > 0 Then
                strChat = Left$(strChat, Len(strChat) - 1)
            End If
            
            'Return finishes and sends the chat
        Case KeyAscii = vbKeyReturn
            
            If bChatActive Then
                'Send it!
                If LenB(strChat) Then
                    SendChatPacket Trim$(Player(0).Name) & modMessaging.MsgNameSeparator & strChat, Player(0).Colour
                End If
                
                'Reset
                bChatActive = False
                strChat = vbNullString
            End If
            
        Case Else
            'If chat is on, add keystroke to chat text
            If bChatActive Then
                If KeyAscii > 31 Then
                    strChat = strChat & Chr$(KeyAscii)
                End If
            End If
        '#########################################
    End Select
    
End If



frmMain.SetInactive
End Sub

Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)

'modSpaceGame.SpaceEditing, bChatActive, PlayerInGame(0)

Select Case KeyCode
        
    Case vbKeyF1
        F1Pressed = False
        
    Case vbKeySpace
        KeyFire = False
    '##############################
    'MOVEMENT######################
    '##############################
        
    Case vbKeyA
        KeyA = False
        
    Case vbKeyD
        KeyD = False
        
    Case vbKeyS
        keys = False
        
    Case vbKeyW
        KeyW = False
        
    Case vbKeyControl
        KeySecondary = False '= True 'fire on release
        
    Case vbKeyShift
        KeyShield = False
        
        
End Select

End Sub

'######################################################
'######################################################
Private Sub MoveCameraX(ByVal nX As Single)

'If cg_Zoom <> 1 Then

modSpaceGame.cg_Camera.X = nX

If PlayerInGame(0) = False Then 'Player(0).Team = Spec Then
    If modSpaceGame.cg_Camera.X < -5500 Then
        modSpaceGame.cg_Camera.X = -5500
    ElseIf modSpaceGame.cg_Camera.X > 14250 Then
        modSpaceGame.cg_Camera.X = 14250
    End If
End If

'ElseIf nX > -8500 Then
    'If nX < 0 Then
        'modSpaceGame.cg_Camera.x = nX
    'End If
'End If
'Debug.Print nX

End Sub
Private Sub MoveCameraY(ByVal nY As Single)

'If cg_Zoom <> 1 Then
modSpaceGame.cg_Camera.Y = nY

If PlayerInGame(0) = False Then 'Player(0).Team = Spec Then
    If modSpaceGame.cg_Camera.Y < -4000 Then
        modSpaceGame.cg_Camera.Y = -4000
    ElseIf modSpaceGame.cg_Camera.Y > 11250 Then
        modSpaceGame.cg_Camera.Y = 11250
    End If
End If

'ElseIf nY > -6900 Then
    'If nY < 0 Then
        'modSpaceGame.cg_Camera.y = nY
    'End If
'End If

End Sub

Private Sub ResetCamera()
'force it to re-position
cg_Camera.X = -9999 'CentreX
cg_Camera.Y = -9999 'CentreY
End Sub
'######################################################
'######################################################

Public Sub SetPlayerState(ID As Integer, State As ePlayerState)

'Find the specified player and set his state
Player(FindPlayer(ID)).State = State

End Sub

Private Sub AddPlayerState(ID As Integer, State As ePlayerState)
Dim i As Integer

i = FindPlayer(ID)

'Find the specified player and add to his state
Player(i).State = (Player(i).State Or State)

End Sub

Public Sub SubPlayerState(ID As Integer, State As ePlayerState)
Dim i As Integer

i = FindPlayer(ID)

'Find the specified player and subtract from his state
Player(i).State = (Player(i).State And Not (State))

End Sub

Public Function FindPlayer(ID As Integer) As Integer

Dim i As Integer

'Find and return the current array index for this player
FindPlayer = -1
For i = 0 To NumPlayers - 1
    'Is this the player?
    If Player(i).ID = ID Then
        'This is the one!
        FindPlayer = i
        Exit Function
    End If
Next i

End Function

Private Sub Form_Load()

Dim SetBoxes As Boolean

If modVars.Closing Then
    Unload Me
    Exit Sub
End If

On Error Resume Next
'set box pos
If modSpaceGame.R_ob1.height <> 0 And modSpaceGame.R_ob1.width <> 0 Then
    ob1.Left = modSpaceGame.R_ob1.Left * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ob1.Top = modSpaceGame.R_ob1.Top * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ob1.width = modSpaceGame.R_ob1.width * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ob1.height = modSpaceGame.R_ob1.height * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    SetBoxes = True
End If
If modSpaceGame.R_ln1.height <> 0 And modSpaceGame.R_ln1.width <> 0 Then
    ln1.Left = modSpaceGame.R_ln1.Left * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ln1.Top = modSpaceGame.R_ln1.Top * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ln1.width = modSpaceGame.R_ln1.width * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ln1.height = modSpaceGame.R_ln1.height * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    SetBoxes = True
End If
If modSpaceGame.R_ln2.height <> 0 And modSpaceGame.R_ln2.width <> 0 Then
    ln2.Left = modSpaceGame.R_ln2.Left * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ln2.Top = modSpaceGame.R_ln2.Top * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ln2.width = modSpaceGame.R_ln2.width * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    ln2.height = modSpaceGame.R_ln2.height * IIf(modSpaceGame.SpaceEditing, modSpaceGame.EditZoom, 1)
    SetBoxes = True
End If
ob1.ZOrder vbSendToBack 'for editing

If Not SetBoxes And modSpaceGame.SpaceEditing Then
    Call ResetBoxPos
End If


Me.BackColor = vbBlack 'IIf(modSpaceGame.cg_BlackBG, vbBlack, &H8000000F)
picMain.width = Me.width
picMain.height = Me.height
picMain.BackColor = Me.BackColor
picMain.Visible = False

ClosingWindow = False

If modSpaceGame.SpaceEditing = False Then
    mnuSave.Visible = False
    mnuReset.Visible = False
    
    ob1.BackStyle = 0 'trans
    ob1.Visible = False
    ln2.Visible = False
    ln1.Visible = False
    lblDragInfo.Visible = False
    picHandle(0).Visible = False
    
    'Call MainLoop
    tmrStart.Enabled = True
Else
    picHandle(0).Visible = True
    Call DragInit 'Initialize drag code
    ob1.BackStyle = 1 'opaque
    ob1.BackColor = vbBlack
    
    ob1.ZOrder vbSendToBack
    ln1.ZOrder vbBringToFront
    ln1.ZOrder vbBringToFront
    
    bSaved = True
    
    ob1.Visible = True
    ln1.Visible = True
    ln2.Visible = True
    lblDragInfo.Left = Me.width / 2 - lblDragInfo.width / 2
    lblDragInfo.Visible = True
End If

Me.Left = 200

Call FormLoad(Me, , , False, True)
Show

End Sub

Private Sub MainLoop()

Const Cap As String = "Multiplayer Combat - "
'Dim i As Integer,
Dim nFrames As Integer
Dim cTick As Long, LastFullSecond As Long

'Me.MousePointer = vbCrosshair

'Connect winsock
If StartWinsock() Then
    
    'Init some variables
    InitVariables
    
    'If we're not the server, try to connect
    If Not modSpaceGame.SpaceServer Then
        Me.Caption = Cap & "Client"
        If ConnectToServer() = False Then
            modWinsock.DestroySocket socket
            Unload Me
            Exit Sub
        End If
        
        'reset box positions
        Call ResetBoxPos
        
    Else
        'socket already bound
        Me.Caption = Cap & "Host"
        
        'tell everyone
        SendInfoMessage frmMain.LastName & " Started a Game - Ctrl+G to Join"
        Pause 100
    End If
    
    If Not modSpaceGame.SpaceServer Then
        modWinsock.SendPacket socket, ServerSockAddr, sChats & Trim$(Player(0).Name) & _
            " joined.#" & modVars.TxtForeGround
    End If
    
    'Start the render loop
    bRunning = True
    Timer = GetTickCount()
    LastFullSecond = Timer
    
    SetCursor True
    
    Do While bRunning
        
        cTick = GetTickCount()

        
        'Check if we've waited for the appropriate number of milliseconds
        If Timer + Space_Ms_Required_Delay < cTick Then
            
            nFrames = nFrames + 1
            If LastFullSecond + 1000 < cTick Then
                LastFullSecond = cTick
                FPS = nFrames
                nFrames = 0
            End If
            
            
            modSpaceGame.SpaceElapsedTime = cTick - Timer
            
            modSpaceGame.TimeFactor = modSpaceGame.sv_GameSpeed * modSpaceGame.SpaceElapsedTime / modSpaceGame.Space_Ms_Delay
            
            
            'Call TbmTimerProc
            Timer = cTick 'GetTickCount()  'Reset the timer variable
            
            
            On Error GoTo EH
            
            If GetPacket = False Then Exit Do                   'Check for network mPacket
            
            If modSpaceGame.cg_Cls Then Me.picMain.Cls
            
            If modSpaceGame.cg_StarBG Then 'draw stars first, since they're futhest back
                ProcessStars
                DrawStars
            End If
            
            ProcessSmoke
            
            If Playing Then
                
                
                Physics                     'Perform physics on ships/bullets
                SendUpdatePacket            'Send network mPacket
                DrawBullets                 'Show the bullets
                DrawMissiles
                DisplayPlayers              'Show all the players
                ProcessAsteroid             'process + draw (before boxes)
                DrawBoxes
                DisplayPowerup
                DisplayHUD              'Display the player's shield level etc
                ShowScores                  'including tab scores + mainmessage + fps
                
                Select Case modSpaceGame.sv_GameType
                    Case CTF
                        DoCTF
                    Case Elimination
                        DoElimination
                End Select
                
                CheckCanUseShips
                
                If modSpaceGame.UseAI And PlayerInGame(0) Then
                    ProcessAI MyID, True
                Else
                    ProcessKeys
                End If
                
                If modSpaceGame.SpaceServer Then
                    GeneratePowerUp
                    SendGameSpeed
                    SendBoxPos
                    CheckPlayerNames
                    SendShipTypes 'if host doesn't get shiptype/team message,
                    SendTeams 'you'll get reset back to what you were before
                    SendScores
                    SendAsteroidUpdate
                    SendServerVarsUpdate
                    SendGameType
                    CheckScores
                End If
                
                
            Else
                
                'ProcessAsteroid 'process + draw (before boxes)
                'DrawBoxes
                ShowRoundScores
                SendAntiLagPacket
                
            End If
            
            ProcessAllCircs 'process + draw
            DisplayChat 'Display the chat text
            ShowChatEntry 'what you're typing
            
            If Playing Then
                If modSpaceGame.cl_UseMouse Then
                    DrawCrosshair
                End If
            End If
            
            BltToForm
            
         End If
EH:
            
        'Allow other events to occur
        DoEvents
    Loop
    
'else
    'error text already added
    
End If

End Sub

Private Sub CheckCanUseShips()
Static Last As Long

If Last + ScoreCheckDelay < GetTickCount() Then
    'can we use it now?
    MotherShipAvail = CheckMotherShip()
    WraithAvail = CheckWraith()
    InfilAvail = CheckInfil()
    SDAvail = CheckSD()
    
    If Player(0).Alive = False Then
        If modSpaceGame.sv_GameType <> Elimination Then
            Player(0).Alive = True
        End If
    End If
    
    
    Last = GetTickCount()
End If

End Sub

Private Sub GeneratePowerUp()

Static LastSpawn As Long
Dim i As Integer


If LastSpawn + PowerUpDelay / modSpaceGame.sv_GameSpeed < GetTickCount() Then
    
    PowerUp.X = Rnd() * (MaxWidth - 500)
    PowerUp.Y = Rnd() * (MaxHeight - 1000)
    PowerUp.Active = True
    
    LastSpawn = GetTickCount()
    
    SendBroadcast sPowerUps & PowerUp.X & "|" & PowerUp.Y
    
End If



End Sub

Private Sub DisplayPowerup()

If PowerUp.Active Then
    gCircle PowerUp.X, PowerUp.Y, Powerup_Radius, vbRed
    gCircle PowerUp.X, PowerUp.Y, Powerup_Radius * 0.75, vbGreen
    gCircle PowerUp.X, PowerUp.Y, Powerup_Radius * 0.5, vbBlue
    
    picMain.DrawWidth = Thin
    'Me.ForeColor = MGrey
    PrintText "PowerUp!", PowerUp.X + 200, PowerUp.Y, MGrey
End If

End Sub

Private Sub ResetBoxPos()

ob1.Top = 960 * modSpaceGame.EditZoom
ob1.Left = 6480 * modSpaceGame.EditZoom
ob1.width = 255 * modSpaceGame.EditZoom
ob1.height = 5175 * modSpaceGame.EditZoom
ln1.Top = 6120 * modSpaceGame.EditZoom
ln1.Left = 2160 * modSpaceGame.EditZoom
ln1.width = 5895 * modSpaceGame.EditZoom
ln1.height = 255 * modSpaceGame.EditZoom
ln2.Top = 0 * modSpaceGame.EditZoom '2040
ln2.Left = 3480 * modSpaceGame.EditZoom '2160
ln2.width = 255 * modSpaceGame.EditZoom
ln2.height = 4095 * modSpaceGame.EditZoom

End Sub

'Private Sub PlaySounds()

'If SomeOneShooting Then modSpaceGame.PlayLasers

'If SomeOneThrusting Then modSpaceGame.PlayThrusters

'End Sub

Private Sub DrawCrosshair()
Static Facing As Single
Const Inc = Pi / 90 '5pi/90=10 degrees
Const d120 = pi2d3 '90+30
Const bit = Pi / 18
Dim C As Long

Dim t1X As Single, t1Y As Single
Dim t2X As Single, t2Y As Single
Dim t3x As Single, t3y As Single
Dim t4x As Single, t4y As Single
Dim t5x As Single, t5y As Single
Dim t6x As Single, t6y As Single


If modSpaceGame.UseAI Then Exit Sub
If Not PlayerInGame(0) Then Exit Sub

C = modSpaceGame.cg_SpaceMainCrosshair
picMain.DrawWidth = modSpaceGame.cg_CrossHairWidth 'Thin

If modSpaceGame.cg_PredatorCrossHair Then
    
    t1X = MouseX + 150 * Sine(Facing - bit) ' - pd3)
    t1Y = MouseY - 150 * CoSine(Facing - bit) ' - pd3)
    t4x = MouseX + 150 * Sine(Facing + bit) ' - pd3)
    t4y = MouseY - 150 * CoSine(Facing + bit) ' - pd3)
    
    t2X = MouseX + 150 * Sine(Facing + d120 - bit) '60 degrees, 180/3
    t2Y = MouseY - 150 * CoSine(Facing + d120 - bit)
    t5x = MouseX + 150 * Sine(Facing + d120 + bit) '60 degrees, 180/3
    t5y = MouseY - 150 * CoSine(Facing + d120 + bit)
    
    t3x = MouseX + 150 * Sine(Facing - d120 - bit)
    t3y = MouseY - 150 * CoSine(Facing - d120 - bit)
    t6x = MouseX + 150 * Sine(Facing - d120 + bit)
    t6y = MouseY - 150 * CoSine(Facing - d120 + bit)
    
    
    picMain.Line (t4x, t4y)-(t2X, t2Y), C 'vbRed
    'PrintText "t1", t1x, t1y
    picMain.Line (t5x, t5y)-(t3x, t3y), C
    'PrintText "t2", t2x, t2y
    picMain.Line (t6x, t6y)-(t1X, t1Y), C
    'PrintText "t3", t3x, t3y
    
    
    Facing = Facing + Inc * modSpaceGame.sv_GameSpeed
    
    If Facing > Pi2 Then Facing = FixAngle(Facing)
    
Else
    picMain.Circle (MouseX, MouseY), 100, C
    'Me.picMain.Circle (MouseX, MouseY), 150, vbRed
    picMain.Line (MouseX - 3, MouseY - 100)-(MouseX - 3, MouseY + 100), C
    picMain.Line (MouseX - 100, MouseY - 3)-(MouseX + 100, MouseY - 3), C
End If

If modSpaceGame.cg_DrawLeadCrossHair Then
    
    C = modSpaceGame.cg_SpaceLeadCrosshair
    
    If Player(0).ShipType <> Hornet Then
        t1X = MouseX + Player(0).Speed * Sine(Player(0).Heading) * 10
        t1Y = MouseY - Player(0).Speed * CoSine(Player(0).Heading) * 10
    Else
        t1X = MouseX '+ Player(0).Speed * sine(Player(0).Heading) * 10
        t1Y = MouseY '- Player(0).Speed * cosine(Player(0).Heading) * 10
    End If
    
    picMain.Circle (t1X, t1Y), 100, C 'vbGreen
    
    'Me.gCircle (MouseX, MouseY), 150, vbRed
    picMain.Line (t1X - 3, t1Y - 100)-(t1X - 3, t1Y + 100), C
    picMain.Line (t1X - 100, t1Y - 3)-(t1X + 100, t1Y - 3), C
End If

End Sub

Private Sub DisplayChat()

Dim i As Integer, j As Integer
Dim FinalTxt As String

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

If NumChat > Max_Chat Then
    'remove numchat - max_chat from the beginning
    'j = NumChat - Max_Chat + 1
    'For i = 0 To j
        'RemoveChatText i
    'Next i
    RemoveChatText LBound(Chat)
End If

'Display chat text
For i = 0 To NumChat - 1
    'ShowText Chat(i).Text, 7, 435 - (i * 16), vbBlack, Me.hdc
    FinalTxt = Chat(i).Text
    'Me.ForeColor = Chat(i).Colour
    PrintFormText FinalTxt, 7, i * TextHeight(FinalTxt) + 3000, Chat(i).Colour
Next i

'PrintText FinalTxt, 7, 2300

End Sub

Private Sub DisplayHUD()

'Show the amount of shields remainString
'PrintText Fix(Player(FindPlayer(MyID)).Shields) & " shields", 7, 7

picMain.DrawWidth = Thin

If F1Pressed = False Then
    If LastZoomPress + ZoomShowTime > GetTickCount() Then
        'Me.ForeColor = vbWhite
        
        PrintFormText "Zoom: " & CStr(Format$(modSpaceGame.cg_Zoom, "0.00")), CentreX, CentreY - 2000, vbWhite
    End If
End If


If PlayerInGame(0) = False Then Exit Sub

Dim tX As Single, tY As Single
Dim Shieldr As Single, Hullr As Single, Start As Single
Dim ShieldRatioIsOne As Boolean, HullRatioIsOne As Boolean
Dim i As Integer
Dim Txt As String
Dim Col As Long
Const HUDWidth As Integer = 600

tX = width - HUDWidth - 200
tY = height - 500

i = FindPlayer(MyID)

On Error GoTo EH 'overflow if maxshields = 0
Shieldr = Abs(Player(i).Shields / _
    (IIf(Player(i).Shields > Player(i).MaxShields, 2, 1) * Player(i).MaxShields))

Hullr = Abs(Player(i).Hull / _
    (IIf(Player(i).Hull > Player(i).MaxHull, 2, 1) * Player(i).MaxHull))


ShieldRatioIsOne = (Shieldr = 1)
HullRatioIsOne = (Hullr = 1)


picMain.DrawWidth = 3
'shields
If Not ShieldRatioIsOne Then
    picMain.Circle (tX, tY), HUDWidth, vbBlue, 0, Pi
End If

On Error Resume Next
Start = Pi * (1 - Shieldr)
If Start > (Pi - 0.1) Then 'pi*179 / 180
    Start = Pi - 0.1
End If

picMain.Circle (tX, tY), HUDWidth, vbGreen, Abs(Start), Pi

'------------------------------------------------------------------

'hull
If Not HullRatioIsOne Then
    picMain.Circle (tX + 15, tY), HUDWidth - 100, vbBlue, 0, Pi
End If
On Error Resume Next
picMain.Circle (tX + 15, tY), HUDWidth - 100, vbRed, Pi * Abs((1 - Hullr)), Pi
'                                             was: pi - hullr*pi
picMain.DrawWidth = Thin

'shield txt
If Player(i).ShipType <> Infiltrator Then
    Txt = "Shields: "
    Col = vbWhite
Else
    Txt = "Stealth: "
    Col = vbGreen
End If

Txt = Txt & Round(Player(i).Shields)

'CurrentX = CurrentX - (TextWidth(Txt) / 2)
'CurrentY = CurrentY - 150
'Print Txt
modSpaceGame.PrintFormText Txt, tX - TextWidth(Txt) / 2, tY - 150, Col


'hull txt
'Me.ForeColor = vbWhite
Txt = "Hull: " & Round(Player(i).Hull)
'CurrentX = tX - (TextWidth(Txt) / 2)
'CurrentY = tY - 150 - TextHeight(Txt)
'Print Txt

modSpaceGame.PrintFormText Txt, tX - TextWidth(Txt) / 2, tY - 150 - TextHeight(Txt), vbWhite


'colour indicator
tX = 400
tY = Me.height - 500
picMain.DrawWidth = 10
picMain.Circle (tX, tY), 300, Player(i).Colour, 0, Pi

EH:
End Sub

Private Function IsAlly(ByVal t1 As eTeams, ByVal t2 As eTeams) As Boolean

IsAlly = (Not (t1 = Neutral Or t2 = Neutral)) And (t1 = t2)

'If T1 = Neutral Or T2 = Neutral Then 'two neutrals aren't allies
'    IsAlly = False
'ElseIf T1 = T2 Then 'if they are equal, then they are allies
'    IsAlly = True
'End If

End Function

Private Sub DoShields(ByVal i As Integer)
Const Shield_StartX2 = SHIELD_START * 2

If PlayerInGame(i) Then
    If Player(i).ShipType <> Infiltrator Then
        If Player(i).Shields < Player(i).MaxShields Then
            
            Player(i).Shields = Player(i).Shields + _
                IIf(Player(i).ShipType <> MotherShip, SHIELD_REGEN, SHIELD_REGEN / 4) * modSpaceGame.TimeFactor
                
        ElseIf Round(Player(i).Shields) = Player(i).MaxShields Then
            Player(i).Shields = Player(i).MaxShields
            
        ElseIf (Player(i).Shields + 10) > Player(i).MaxShields Then
            'if lag/fps-lag, prevent shields from going a bit too far over
            Player(i).Shields = Player(i).MaxShields
            
        End If
        
        'If Player(i).Shields >= Player(i).MaxShields Then
            'player(i).
    
    Else
        'infiltrator ------
        If Player(i).Hull < Hull_Start Then
            
            Player(i).Hull = Player(i).Hull + SHIELD_REGEN * 1.5 * modSpaceGame.sv_GameSpeed
            
        ElseIf Round(Player(i).Hull) = Player(i).MaxHull Then
            
            Player(i).Hull = Player(i).MaxHull
            
        End If
        
        If (Player(i).State And Player_Secondary) = Player_Secondary Then
            Player(i).Shields = Player(i).Shields - 0.2 * modSpaceGame.sv_GameSpeed
            
            If Player(i).Shields <= 0 Then
                SubPlayerState Player(i).ID, Player_Secondary
            End If
            
        ElseIf Player(i).Shields < SHIELD_START Then
            Player(i).Shields = Player(i).Shields + 0.1 * modSpaceGame.sv_GameSpeed
            
        ElseIf Player(i).Shields > SHIELD_START Then
            Player(i).Shields = SHIELD_START
            
        End If
        
    End If
    
    If Player(i).Shields > Shield_StartX2 Then
        Player(i).Shields = Shield_StartX2
    End If
    
End If

End Sub

Private Function GetShipRadius(ByVal vShiptype As eShipTypes) As Single

If vShiptype = Raptor Then
    GetShipRadius = SHIP_Height
ElseIf vShiptype = Behemoth Then
    GetShipRadius = SHIP_Height * 1.2
ElseIf vShiptype = Hornet Then
    GetShipRadius = SHIP_Height \ 1.5
ElseIf vShiptype = MotherShip Then
    GetShipRadius = SHIP_Height * 2.5
ElseIf vShiptype = Wraith Then
    GetShipRadius = SHIP_Height * 1.2
ElseIf vShiptype = Infiltrator Then
    GetShipRadius = SHIP_Height * 1.3
Else
    GetShipRadius = SHIP_Height * 2.5
End If

'MinDist = IIf(ShipType = Raptor, SHIP_Height, _
    IIf(ShipType = Behemoth, SHIP_Height * 1.2, _
    IIf(ShipType = Hornet, SHIP_Height \ 1.5, _
    IIf(ShipType = MotherShip, SHIP_Height * 3, _
    IIf(ShipType = Wraith, SHIP_Height * 1.2, _
    SHIP_Height * 1.3)))))

End Function

Private Function GetAccel(vShiptype As eShipTypes, iPlayer As Integer) As Single

If vShiptype = Raptor Then
    GetAccel = Raptor_ACCEL
ElseIf vShiptype = Behemoth Then
    GetAccel = Behemoth_ACCEL
ElseIf vShiptype = Hornet Then
    GetAccel = Hornet_Accel
ElseIf vShiptype = MotherShip Then
    GetAccel = Mothership_Accel
ElseIf vShiptype = Wraith Then
    GetAccel = Wraith_Accel
ElseIf vShiptype = Infiltrator Then
    GetAccel = Infil_Accel
Else
    If (Player(iPlayer).State And Player_Secondary) = 0 Then
        GetAccel = SDNorm_Accel
    Else
        GetAccel = SDGW_Accel
    End If
End If

End Function

Private Function GetShipMass(ByVal vShiptype As eShipTypes) As Single

If vShiptype = Raptor Then
    GetShipMass = 1
ElseIf vShiptype = Behemoth Then
    GetShipMass = 6
ElseIf vShiptype = Hornet Then
    GetShipMass = 0.8
ElseIf vShiptype = MotherShip Then
    GetShipMass = 12
ElseIf vShiptype = Wraith Then
    GetShipMass = 5
ElseIf vShiptype = Infiltrator Then
    GetShipMass = 3
Else 'SD
    GetShipMass = 13
End If

'MinDist = IIf(ShipType = Raptor, SHIP_Height, _
    IIf(ShipType = Behemoth, SHIP_Height * 1.2, _
    IIf(ShipType = Hornet, SHIP_Height \ 1.5, _
    IIf(ShipType = MotherShip, SHIP_Height * 3, _
    IIf(ShipType = Wraith, SHIP_Height * 1.2, _
    SHIP_Height * 1.3)))))

End Function

Private Function PlayerInGame(ByVal iPlayer As Integer) As Boolean

PlayerInGame = Player(iPlayer).Team <> Spec And Player(iPlayer).Alive

End Function

Private Sub LockMissile(ByVal Playeri As Integer)
Dim Targeti As Integer
Dim nX As Single, nY As Single


Player(Playeri).MissileLocki = FindLeastDegreeTarget(Playeri, False)

If modSpaceGame.cg_ShowMissileLock Then
    If Playeri = 0 Then
        'show who we're locked on to
        If Player(Playeri).MissileLocki <> -1 Then
            
            Targeti = FindPlayer(Player(Playeri).MissileLocki)
            
            If Targeti <> -1 Then
                
                If GetDist(Player(Playeri).X, Player(Playeri).Y, Player(Targeti).X, Player(Targeti).Y) < MissileLockDist Then
                    
                    nX = Player(Targeti).X '+ Player(Targeti).Speed * sine(Player(Targeti).Heading)
                    nY = Player(Targeti).Y '- Player(Targeti).Speed * cosine(Player(Targeti).Heading)
                    
                    picMain.DrawWidth = 4
                    
                    gCircle nX, nY, 700, vbRed
                    'gCircle nX, nY, 450, vbGreen
                    'gCircle nX, nY, 400, vbBlue
                    
                    'Me.ForeColor = vbRed
                    PrintText "Lock Acquired", nX + 650, nY - 650, vbRed
                    
                    picMain.DrawWidth = Thin
                End If
                
            End If
            
        End If
    End If
End If

End Sub

Private Sub Physics()

Dim ShipVal As Single
Dim i As Integer
Dim j As Integer
Dim TempMag As Single
Dim TempDir As Single
Dim bX As Single, bY As Single
Dim ShipType As eShipTypes
'Dim HasShield As Boolean
Dim MaxSpeed As Integer
Dim MinDist As Single
'Dim BDamage As Single
Dim K As Integer
Dim Factor As Single
Dim Tmp As String
Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single

'Dim ASpeed As Single, AHeading As Single

'SomeOneShooting = False
'SomeOneThrusting = False


'Add to shields

'On Error GoTo EH
On Error GoTo phyEH

i = 0

If Player(i).ShipType <> MotherShip Then
    If Player(i).ShipType <> SD Then
        If Player(i).Speed < 2.5 Then '* modSpaceGame.sv_GameSpeed) Then '2.5 Then 'let the player stop
            If Player(i).Speed <> 0 Then
                Player(i).Speed = 0
            End If
        End If
    End If
End If


If (Player(i).State And Player_Shielding) = 0 Then
    DoShields 0 'FindPlayer(MyID)
Else
    
    'shields are up
    If Player(i).Shields < 1 Then
        SubPlayerState Player(i).ID, Player_Shielding
        Tmp = "Shields Have No Power"
        'Me.ForeColor = Player(0).Colour
        PrintFormText Tmp, CentreX - TextWidth(Tmp) / 2, 20, Player(0).Colour
    End If
    
    
    If Player(i).ShipType = SD Then
        
        If (Player(i).State And Player_Secondary) = Player_Secondary Then
            SubPlayerState Player(i).ID, Player_Secondary
        End If
        
    End If
    
End If

If modSpaceGame.SpaceServer Then
    
    For j = 0 To NumPlayers - 1 'NumBotIDs - 1
    'If BotID <> -1 Then
        
        'i = FindPlayer(BotIDs(j)) 'BotID)
        If Player(j).IsBot Then
        'If i <> -1 Then
            DoShields j
            
            If modSpaceGame.sv_BotAI Then
                ProcessAI Player(j).ID
            Else
                If Player(j).State <> Player_None Then
                    SetPlayerState Player(j).ID, Player_None
                End If
'                If Player(i).State And Player_Fire) = Player_Fire Then
'                    SubPlayerState BotID, Player_Fire
'                ElseIf (Player(i).State And PLAYER_THRUST) = PLAYER_THRUST Then
'                    SubPlayerState BotID, PLAYER_THRUST
'                End If
            End If
        End If
        
     Next j
     
    'End If
End If


i = 0
Do While i < NumPlayers
    'Skip the local player
    If (Player(i).ID <> MyID) And (Player(i).IsBot = False) Then
        'Time to remove this player?
        If (Player(i).LastPacket + mPacket_LAG_KILL < GetTickCount()) Then
            'Remove!
            SendChatPacketBroadcast Trim$(Player(i).Name) & " lagged out", vbRed
            SendPacket socket, Player(i).ptSockAddr, sKicks & "Lag"
            RemovePlayer i
            i = i - 1
        ElseIf (Player(i).LastPacket + mPacket_LAG_TOL < GetTickCount()) Then 'And Player(i).ID <> MyID Then
            SetPlayerState Player(i).ID, Player_None
        End If
    Else
        'player is either me or a bot
        If modSpaceGame.SpaceServer Or Player(i).ID = MyID Then
            Player(i).LastPacket = GetTickCount()
        End If
    End If
    'Increment counter
    i = i + 1
Loop


For i = 0 To NumPlayers - 1
    If PlayerInGame(i) Then
        
        
        Call LockMissile(i)
        
        
        If Player(i).ShipType = SD Then
            If (Player(i).State And Player_Secondary) = Player_Secondary Then
                
                For j = 0 To NumPlayers - 1
                    If j <> i Then
                        If GetDist(Player(i).X, Player(i).Y, Player(j).X, Player(j).Y) < (SHIP_Height * SD_GravityRadius) Then
                            Player(j).Speed = modSpaceGame.sv_GameSpeed * Player(j).Speed / 2 'yes, this   ^ should be *
                            
                            '-5% their speed
                            'Player(i).Speed = Player(i).Speed - Player(i).Speed * modSpaceGame.sv_GameSpeed / 100
                            
                        End If
                    End If
                Next j
            End If
        End If
    End If
Next i


'Loop through each player and perform physics
For i = 0 To NumPlayers - 1
    
    'If Player(i).IsBot And modSpaceGame.SpaceServer Then
        'ProcessAI Player(i).ID
        'If Not ((Player(i).State And PLAYER_THRUST) = PLAYER_THRUST) Then
            'AddPlayerState Player(i).ID, PLAYER_THRUST
        'End If
    'End If
    If PlayerInGame(i) Then 'Player(i).Team <> Spec Then
        ShipType = Player(i).ShipType
        
    '    'Check lag tol
    '    If ((Player(i).LastPacket + mPacket_LAG_TOL) < GetTickCount()) And (Player(i).ID <> MyID) And (Player(i).IsBot = False) Then
    '        Player(i).State = 0
    '    End If
        
        'Cap Speed up here to avoid error in addvectors (overflow)
        If ShipType = Raptor Then
            MaxSpeed = Raptor_MAX_SPEED
        ElseIf ShipType = Behemoth Then
            MaxSpeed = Behemoth_MAX_SPEED
        ElseIf ShipType = Hornet Then
            MaxSpeed = Hornet_Max_Speed
        ElseIf ShipType = MotherShip Then
            MaxSpeed = MotherShip_Max_Speed
        ElseIf ShipType = Wraith Then
            MaxSpeed = Wraith_Max_Speed
        ElseIf ShipType = Infiltrator Then
            MaxSpeed = Infil_Max_Speed
        Else
            If (Player(i).State And Player_Secondary) = 0 Then
                MaxSpeed = SDNorm_Max_Speed
            Else
                MaxSpeed = SDGW_Max_Speed
            End If
        End If
        If Player(i).Speed > MaxSpeed Then Player(i).Speed = MaxSpeed
        
        
        'Firing
        If (Player(i).State And PLAYER_FIRE) = PLAYER_FIRE Then
            
            If (Player(i).State And Player_Shielding) = 0 Then
                
                ShipVal = IIf(ShipType = Raptor, Raptor_Bullet_DELAY, _
                    IIf(ShipType = Behemoth, Behemoth_Bullet_DELAY, _
                    IIf(ShipType = Hornet, Hornet_Bullet_Delay, _
                    IIf(ShipType = MotherShip, Mothership_Bullet_Delay, _
                    IIf(ShipType = Wraith, Wraith_Bullet_Delay, _
                    IIf(ShipType = Infiltrator, Infil_Bullet_Delay, _
                    SD_Bullet_Delay)))))) / modSpaceGame.sv_GameSpeed
                
                If (Player(i).LastBullet + ShipVal) <= GetTickCount() Then
                
                    'Reset bullet timer
                    Player(i).LastBullet = GetTickCount()
                    Player(i).bRightBullet = Not Player(i).bRightBullet
                    
                    'Fire the bullet!
                    If ShipType <> Hornet Then
                        AddVectors Player(i).Speed, Player(i).Heading, _
                            IIf(ShipType <> Hornet, IIf(ShipType <> Wraith, BULLET_SPEED, Wraith_Bullet_Speed), Hornet_Bullet_Speed) _
                                , Player(i).Facing, TempMag, TempDir
                    Else
                        TempMag = BULLET_SPEED * 1.2
                        TempDir = Player(i).Facing
                    End If
                    
                    Factor = IIf(ShipType = Behemoth, 1.5, 1) 'GetShipRadius(ShipType) / 110
                    
                    If Player(i).ShipType <> SD Then
                        If Player(i).bRightBullet Then
                            bX = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing + pi2d3) * Factor
                            bY = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing + pi2d3) * Factor
                            'If ShipType <> Hornet Then
                            TempDir = TempDir - GunOffset
                        Else
                            bX = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing + 4 * piD3) * Factor
                            bY = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing + 4 * piD3) * Factor
                            'If ShipType <> Hornet Then
                            TempDir = TempDir + GunOffset
                        End If
                        
                        AddBullet bX, bY, TempMag, TempDir, Player(i).ID, Player(i).Colour, _
                            IIf(Player(i).ShipType = Infiltrator, sv_Bullet_Damage * Infil_Bullet_Damage_Factor, _
                            sv_Bullet_Damage), i
                        
                    Else
                        
                        Call GetSDGunTurrets(bX, X1, X2, bY, Y1, Y2, i)
                        
                        If Player(i).bRightBullet Then
                            TempDir = TempDir + GunOffset
                        Else
                            TempDir = TempDir - GunOffset
                            X1 = X2
                            Y1 = Y2
                        End If
                        
                        AddBullet X1, Y1, TempMag, TempDir, Player(i).ID, Player(i).Colour, sv_Bullet_Damage, i
                        
                    End If
                    
                    
                    If ShipType = MotherShip Then
                        
                        AddBullet Player(i).X, Player(i).Y, BULLET_SPEED * 2, Player(i).Facing, Player(i).ID, _
                            Player(i).Colour, sv_Bullet_Damage * 1.2, i
                        
                    ElseIf ShipType = Behemoth Then
                        
                        AddBullet Player(i).X, Player(i).Y, TempMag, TempDir, Player(i).ID, _
                            Player(i).Colour, sv_Bullet_Damage, i
                        
                    ElseIf ShipType = SD Then
                        
                        AddBullet bX, bY, TempMag, TempDir, Player(i).ID, _
                            Player(i).Colour, sv_Bullet_Damage, i
                        
                    End If
                
                End If 'tick endif
                
            End If 'shielding endif
            
        End If 'fire endif
        
        Call Do2ndryFire(i)
        
        If (Player(i).State And PLAYER_THRUST) = PLAYER_THRUST Then
            'Apply acceleration
            
            ShipVal = GetAccel(ShipType, i)
            
            AddVectors Player(i).Speed, Player(i).Heading, ShipVal, Player(i).Facing, Player(i).Speed, Player(i).Heading
            
            'SomeOneThrusting = True
            
        ElseIf (Player(i).State And PLAYER_REVTHRUST) = PLAYER_REVTHRUST Then
            'Apply reverse acceleration
            
            ShipVal = GetAccel(ShipType, i)
            
            AddVectors Player(i).Speed, Player(i).Heading, -ShipVal, Player(i).Facing, Player(i).Speed, Player(i).Heading
            
            'SomeOneThrusting = True
            
        End If
        
        
        'strafing--------------------------------------------------------------
        If (Player(i).State And Player_StrafeRight) = Player_StrafeRight Then
            'Apply acceleration
            
            ShipVal = GetAccel(ShipType, i)
            
            AddVectors Player(i).Speed, Player(i).Heading, ShipVal, Player(i).Facing + piD2, Player(i).Speed, Player(i).Heading
            
            
        ElseIf (Player(i).State And Player_StrafeLeft) = Player_StrafeLeft Then
            'Apply reverse acceleration
            
            ShipVal = GetAccel(ShipType, i)
            
            AddVectors Player(i).Speed, Player(i).Heading, ShipVal, Player(i).Facing - piD2, Player(i).Speed, Player(i).Heading
            
            
        End If
        'end strafing--------------------------------------------------------------
        
        If modSpaceGame.UseAI = False Then
            'Rotation
            If modSpaceGame.cl_UseMouse And i = 0 Then 'FindPlayer(MyID) Then
                'Player(i).Facing = FindAngle(Player(i).X, Player(i).Y, MouseX, MouseY)
                
                Player(i).Facing = FindAngle(Player(i).X * cg_Zoom - cg_Camera.X, _
                                             Player(i).Y * cg_Zoom - cg_Camera.Y, _
                                             MouseX, _
                                             MouseY)
                
            Else
                If (Player(i).State And PLAYER_LEFT) = PLAYER_LEFT Then
                    Player(i).Facing = Player(i).Facing - ROTATION_RATE * Pi / 180
                End If
                If (Player(i).State And PLAYER_RIGHT) = PLAYER_RIGHT Then
                    Player(i).Facing = Player(i).Facing + ROTATION_RATE * Pi / 180
                End If
            End If
        End If
        
         Player(i).Facing = FixAngle(Player(i).Facing)
        
        
        
        If Player(i).Shields < 0 Then
            Player(i).Shields = 0
        End If
        
        Player(i).Facing = FixAngle(Player(i).Facing)
        Player(i).Heading = FixAngle(Player(i).Heading)
        
        'Motion
        Motion Player(i).X, Player(i).Y, Player(i).Speed, Player(i).Heading
        
        'Wrap edges
        Call ClipEdges(i)
        
        
        'check if we've got the powerup
        
        If PowerUp.Active Then
            
            'MinDist = IIf(ShipType = Raptor, SHIP_Height, _
                    IIf(ShipType = Behemoth, SHIP_Height * 1.2, _
                    IIf(ShipType = Hornet, SHIP_Height \ 1.5, _
                    IIf(ShipType = MotherShip, SHIP_Height * 3, _
                    SHIP_Height))))
            
            MinDist = GetShipRadius(ShipType)
            
            If (GetDist(Player(i).X, Player(i).Y, PowerUp.X, PowerUp.Y) < (MinDist + Powerup_Radius)) Then
                
                PowerUp.Active = False
                
                If (Player(i).State And Player_Secondary) = Player_Secondary Then
                    SubPlayerState Player(i).ID, Player_Secondary
                End If
                
                If Player(i).ID = MyID Or (Player(i).IsBot And modSpaceGame.SpaceServer) Then 'IsBotID(Player(i).ID)
                    Player(i).Shields = Player(i).MaxShields * 2
                    Player(i).Hull = Player(i).MaxHull * 2
                    Player(i).LastSecondary = 1
                    
                    
                    If Player(i).ID = MyID Then
                        If Player(i).ShipType = MotherShip Then
                            MSStartFire = 0
                        End If
                    End If
                        
                End If
                
                AddCirc PowerUp.X, PowerUp.Y, 300, 1, vbGreen ', Player(i).Speed, Player(i).Heading
                
            End If
            
            
            
        End If 'powerup endif
        
    End If 'InGame endif
    
    
    
Next i


Call CheckPlayerCollisions '+ asteroid

If modSpaceGame.sv_BulletsCollide Then
    Call CheckBulletCollisions
End If

'Loop through each bullet and perform physics
i = 0
Do While i < NumBullets
    'Move it!
    Motion Bullet(i).X, Bullet(i).Y, Bullet(i).Speed, Bullet(i).Heading
    'Wrap edges
    If ClipBullet(i) = False Then
        
        If modSpaceGame.sv_ClipMissiles Then
            'check missile-bullet collisions
            
            For j = 0 To NumMissiles - 1
                'If Missiles(j).Owner <> Bullet(i).Owner Then
                    If GetDist(Missiles(j).X, Missiles(j).Y, Bullet(i).X, Bullet(i).Y) < (Missile_Radius + Bullet_Radius) Then
                        'kill the bullet, sub hull from missile, goto next bullet
                        
                        Missiles(j).Hull = Missiles(j).Hull - Bullet(i).Damage
                        
                        bX = Missiles(j).Speed
                        bY = Missiles(j).Heading
                        
                        If Missiles(j).Hull < 1 Then
                            RemoveMissile j
                            If Bullet(i).Owner = MyID Then
                                MissilesShot = MissilesShot + 1
                            End If
                        End If
                        
                        RemoveBullet i, True, bX, bY 'need to remove it here, incase we need to know bullet's owner
                        i = i - 1
                        
                        GoTo NextBullet
                        
                    End If 'getdist endif
                'End If'ownerID endif
            Next j
            
        End If
        
        
        'Check for collisions
        For j = 0 To NumPlayers - 1
            
            
            If PlayerInGame(j) Then 'Player(i).Team <> Spec Then
                ShipType = Player(j).ShipType
                
                'Collision?
                
                MinDist = GetShipRadius(ShipType)
                
                'MinDist = IIf(ShipType = Raptor, SHIP_Height, _
                    IIf(ShipType = Behemoth, SHIP_Height * 1.2, _
                    IIf(ShipType = Hornet, SHIP_Height \ 1.5, _
                    IIf(ShipType = MotherShip, SHIP_Height * 3, _
                    SHIP_Height))))
                
    '            If ShipType = Raptor Then
    '                MinDist = SHIP_Height
    '            ElseIf ShipType = Behemoth Then
    '                MinDist = SHIP_Height * 2
    '            Else
    '                MinDist = SHIP_Height \ 1.5
    '            End If
                
                If Player(j).ID <> Bullet(i).Owner Then
                    If GetDist(Player(j).X, Player(j).Y, Bullet(i).X, Bullet(i).Y) < (MinDist + Bullet_Radius) Then
    
                        If IsAlly(Player(j).Team, Player(FindPlayer(Bullet(i).Owner)).Team) = False Then
                            
                            'Collision!
                            Player(j).bDrawShields = True
                            
                            If modSpaceGame.sv_AddBulletVectorToShip Then
                                If (Player(j).State And Player_Shielding) = 0 Then
                                    If Player(j).ShipType <> SD Then
                                        Call BulletShipCol(j, i)
                                    End If
                                End If
                            End If
                            
                            'If this is our player or a bot...
                            'ShipVal = FindPlayer(MyID)
                            
                            'If modSpaceGame.NumBotIDs = 0 Then 'BotID = -1 Then
                                'k = -1
                            'Else
                                'k = FindPlayer(BotID)
                            'End If
                            
                            'If j <> K Then K = -1
                            'if the player's index isn't the bot's index, bot's index = -1
                            
                            K = -1
                            
                            If Player(j).ID = MyID Then
                                K = j
                            ElseIf modSpaceGame.SpaceServer Then
                                If Player(j).IsBot Then
                                    K = j
                                End If
                            End If
                            
                            
                            If K <> -1 Then 'ShipVal = j Or k = j Then 'if it is us or the bot then...
                                
                                'k = j
                                
                                Call DoBulletDamage(K, i) 'apply damage to ship
                                
                                If Player(K).ShipType <> eShipTypes.MotherShip Then 'if we aren't a MS
                                    'Kill the bullet
                                    RemoveBullet i, True, 0, 0 'Player(k).Speed, Player(k).Heading
                                    'Decrement the counter
                                    i = i - 1
                                    'Exit loop - next bullet
                                    Exit For
                                Else
                                    If Rnd() < MotherShipDeflPercent Then
                                        ShipVal = FindClosestTarget_ID(Bullet(i).X, Bullet(i).Y, MyID)
                                        DeflectBullet ShipVal, i, Player(j).ID
                                        Exit For
                                    Else
                                        'Kill the bullet
                                        RemoveBullet i, True, 0, 0 'Player(k).Speed, Player(k).Heading
                                        'Decrement the counter
                                        i = i - 1
                                        'Exit loop - next bullet
                                        Exit For
                                    End If
                                    
                                End If
                                
                            ElseIf Player(j).ShipType = eShipTypes.MotherShip Then
                                If Rnd() < MotherShipDeflPercent Then
                                    ShipVal = FindClosestTarget_ID(Bullet(i).X, Bullet(i).Y, Player(j).ID)
                                    DeflectBullet ShipVal, i, Player(j).ID
                                Else
                                    'Kill the bullet
                                    RemoveBullet i, True, 0, 0 'Player(k).Speed, Player(k).Heading
                                    'Decrement the counter
                                    i = i - 1
                                    'Exit loop - next bullet
                                    Exit For
                                End If
                            Else
                                'Kill the bullet
                                RemoveBullet i, True, 0, 0 'Player(k).Speed, Player(k).Heading
                                'Decrement the counter
                                i = i - 1
                                'Exit loop - next bullet
                                Exit For
                            End If
                            
                        End If 'is ally endif
                        
                    End If 'getdist endif
                End If 'owner = player endif
            End If 'spec endif
            
        Next j
        
        
    End If 'clip bullet endif
    
    
NextBullet:
    
    'Increment counter
    i = i + 1
    
Loop


'missiles --------------------------------------------------------------------------------------------
i = 0
Do While i < NumMissiles
    'Move it!
    
    HomeMissile i
    
    Motion Missiles(i).X, Missiles(i).Y, Missiles(i).Speed, Missiles(i).Heading
    
    If ClipMissile(i) = False Then
        For j = 0 To NumPlayers - 1
            
            If PlayerInGame(j) Then 'Player(i).Team <> Spec Then
                
                ShipType = Player(j).ShipType
                
                'MinDist = IIf(ShipType = Raptor, SHIP_Height, _
                    IIf(ShipType = Behemoth, SHIP_Height * 1.2, _
                    IIf(ShipType = Hornet, SHIP_Height \ 1.5, _
                    IIf(ShipType = MotherShip, SHIP_Height * 3, _
                    IIf(ShipType = Wraith, SHIP_Height * 1.2, _
                    SHIP_Height * 1.3)))))
                
                MinDist = GetShipRadius(ShipType)
                
                If (GetDist(Player(j).X, Player(j).Y, Missiles(i).X, Missiles(i).Y) < (MinDist + Missile_Radius)) _
                    And (Player(j).ID <> Missiles(i).Owner) Then
                    
                    'Collision!
                    Player(j).bDrawShields = True
                    
'                    ShipVal = FindPlayer(MyID)
'                    If BotID = -1 Then
'                        k = -1
'                    Else
'                        k = FindPlayer(BotID)
'                    End If
                    
                    K = -1
                    
                    If Player(j).ID = MyID Then
                        K = j
                    ElseIf modSpaceGame.SpaceServer Then
                        If Player(j).IsBot Then
                            K = j
                        End If
                    End If
                    
                    'If j <> K Then K = -1
                    'if the player's index isn't the bot's index, bot's index = -1
                    
                    'If ShipVal = j Or j = k Then 'if it is us then...
                    If K <> -1 Then
                        
                        If Player(K).Shields > Missiles(i).Hull And Player(j).ShipType <> Infiltrator Then
                            'Decrement shields
                            
                            If (Player(j).State And Player_Shielding) = 0 Then
                                Player(j).Shields = Player(j).Shields - Missiles(i).Hull 'Missile_Damage
                            Else
                                Player(j).Shields = Player(j).Shields - Missiles(i).Hull / ShieldDmgReduction
                            End If
                            
                        Else
                            'Decrement hull
                            If (Player(j).State And Player_Shielding) = 0 Then
                                Player(K).Hull = Player(K).Hull - Missiles(i).Hull + Player(K).Shields
                                Player(K).Shields = 0
                            Else
                                Player(K).Hull = Player(K).Hull - Missiles(i).Hull / ShieldDmgReduction + Player(K).Shields
                                Player(K).Shields = 0
                            End If
                            
                        End If
                        
                        'Check for death
                        If Player(K).Hull < 1 Then Call Killed(i, K, True)
                        
                    End If
                    
                    RemoveMissile i
                End If 'getdist endif
            End If 'spec endif
        Next j
        
    End If
    i = i + 1
Loop


Exit Sub
phyEH:
Resume Next
End Sub

Private Sub DoBulletDamage(ByVal j As Integer, ByVal i As Integer)

Dim HasShield As Boolean
Dim baseDmg As Single, BDamage As Single
Dim iOwner As Integer

HasShield = (Player(j).Shields > 1) And (Player(j).ShipType <> Infiltrator)

baseDmg = Bullet(i).Damage

If Player(j).ShipType = eShipTypes.Raptor Then
    BDamage = baseDmg
ElseIf Player(j).ShipType = eShipTypes.Behemoth Then
    BDamage = baseDmg / BehemothDmgReduction
ElseIf Player(j).ShipType = eShipTypes.Hornet Then
    BDamage = baseDmg * HornetDmgIncrease
ElseIf Player(j).ShipType = Wraith Then
    BDamage = baseDmg / WraithDmgReduction
ElseIf Player(j).ShipType = MotherShip Then
    BDamage = baseDmg / MothershipDmgReduction
ElseIf Player(j).ShipType = Infiltrator Then
    BDamage = baseDmg * InfilDmgIncrease
Else
    BDamage = baseDmg / SDDmgReduction
End If

'if player(findplaybullet(i).Owner - was going to reduce wraith bullet dmg...

If (Player(j).State And Player_Shielding) = Player_Shielding Then
    BDamage = BDamage / ShieldDmgReduction
End If

iOwner = FindPlayer(Bullet(i).Owner)
If iOwner > -1 Then
    If Player(iOwner).ShipType = Raptor Then
        BDamage = BDamage * RaptorBulletDamage
    End If
End If

If HasShield Then
    'Decrement shields
    Player(j).Shields = Player(j).Shields - BDamage
    
    If Player(j).Shields < 0 Then
        Player(j).Hull = Player(j).Hull + Player(j).Shields
        Player(j).Shields = 0
    End If
    
Else
    'Decrement hull
    Player(j).Hull = Player(j).Hull - BDamage + Player(j).Shields
    
    If Player(j).ShipType <> Infiltrator Then
        Player(j).Shields = -1
    End If
End If


'Check for death
If Player(j).Hull < 1 Then Call Killed(i, j)

End Sub

Private Sub DeflectBullet(ByVal ClosestID As Integer, ByVal Bulleti As Integer, ByVal PlayerjID As Integer)
Dim ClosestI As Integer

If ClosestID <> -1 Then
    
    ClosestI = FindPlayer(ClosestID)
    
    Bullet(Bulleti).Heading = FindAngle(Bullet(Bulleti).X, Bullet(Bulleti).Y, _
        Player(ClosestI).X, Player(ClosestI).Y)
    
    Bullet(Bulleti).Damage = sv_Bullet_Damage
    
    Bullet(Bulleti).Decay = GetTickCount() + Bullet_Decay / modSpaceGame.sv_GameSpeed
    
    Bullet(Bulleti).Owner = Player(FindPlayer(PlayerjID)).ID
    
End If

End Sub

Private Sub Killed(ByVal i As Integer, ByVal j As Integer, Optional ByVal bIsMissile As Boolean = False, _
    Optional ByVal bIsCollision As Boolean = False)

Dim ChatText As String
Dim iKiller As Integer

If bIsMissile = False And bIsCollision = False Then
    iKiller = FindPlayer(Bullet(i).Owner)
    ChatText = Trim$(Player(j).Name) & ShotBy & Trim$(Player(iKiller).Name)
ElseIf bIsMissile Then
    iKiller = FindPlayer(Missiles(i).Owner)
    ChatText = Trim$(Player(j).Name) & MissiledBy & Trim$(Player(iKiller).Name)
Else
    iKiller = i 'rammed
    ChatText = Trim$(Player(j).Name) & RammedBy & Trim$(Player(iKiller).Name)
End If


Player(iKiller).Shields = Player(iKiller).MaxShields 'if it's a bot...
Player(iKiller).Hull = Player(iKiller).MaxHull

If modSpaceGame.SpaceServer Then
    SendChatPacketBroadcast ChatText, Player(iKiller).Colour
    
    'getpacket() won't receive the above packet, so won't add to my version of someone's score, so...
    
    'Player(iKiller).Kills = Player(iKiller).Kills + 1
    
    'Player(j).Deaths = Player(j).Deaths + 1
    
    
    
Else
    modWinsock.SendPacket socket, ServerSockAddr, sChats & ChatText & "#" & _
                                    CStr(Player(iKiller).Colour)
    
    
    
    'If Bullet(i).Owner <> 0 Then 'if it's not the server...
        'Player(iKiller).Score = Player(iKiller).Score + 1
    'End If
End If

If Player(j).ID = MyID Then
    KillsInARow = 0
    
    'If modSpaceGame.sv_GameType <> Elimination Then
        'LastRespawn = GetTickCount()
    'End If
    'AddExplosion Player(j).x, Player(j).y, 400, 1, Player(j).Speed, Player(j).Heading
End If

'picMain.DrawWidth = 5
'gCircle (Player(j).x, Player(j).y), SHIP_Height * 2, vbRed
'gCircle (Player(j).x, Player(j).y), SHIP_Height, vbRed

'Reset the player

'If Player(j).ID = MyID Or Player(j).IsBot Then
    Player(j).X = (MaxWidth - 100) * Rnd()
    Player(j).Y = (MaxHeight - 100) * Rnd()
    Player(j).Heading = Pi2 * Rnd()
    'Player(j).Facing = Pi2 * Rnd()
    
    If PlayerIsInAsteroid(j) Then
        Player(j).X = (MaxWidth - 100) * Rnd()
        Player(j).Y = (MaxHeight - 100) * Rnd()
    End If
    
'End If

Player(j).Speed = Start_Speed

If Player(j).IsBot Then
    Player(j).Shields = Player(j).MaxShields 'might be an easy bot
    Player(j).Hull = Player(j).MaxHull
Else
    Player(j).Shields = SHIELD_START
    Player(j).MaxShields = SHIELD_START
    Player(j).Hull = Hull_Start
    Player(j).MaxHull = Hull_Start
End If

Player(j).LastSecondary = 0
Player(j).LastBullet = 0

'player(j).Colour = player(j).Colour
'gCircle (Player(j).x, Player(j).y), SHIP_Height * 2, vbBlue
'gCircle (Player(j).x, Player(j).y), SHIP_Height, vbBlue
'AddCirc Player(j).x, Player(j).y, 300, 1, vbBlue,

End Sub

Private Sub BulletShipCol(iPlayer As Integer, iBullet As Integer)

Dim PlayerShielding As Boolean
Dim Factor As Single

PlayerShielding = ((Player(iPlayer).State And Player_Shielding) = Player_Shielding)

Factor = IIf(PlayerShielding, 5, 10) * GetShipMass(Player(iPlayer).ShipType)


AddVectors Bullet(iBullet).Speed / Factor, Bullet(iBullet).Heading, _
    Player(iPlayer).Speed, Player(iPlayer).Heading, Player(iPlayer).Speed, Player(iPlayer).Heading

'Dim XComp As Single, YComp As Single
'
'
'With Player(iPlayer)
'    XComp = .Speed * sine(.Heading) + Bullet(iBullet).Speed * sine(Bullet(iBullet).Heading)
'    YComp = .Speed * cosine(.Heading) + Bullet(iBullet).Speed * cosine(Bullet(iBullet).Heading)
'
'
'
'End With

End Sub

Private Sub CheckBulletCollisions()
Dim i As Integer, j As Integer
Dim TmpHeading As Single, TmpSpeed As Single

Const Bullet_RadiusX2 = Bullet_Radius * 2

For i = 0 To NumBullets - 1
    For j = (i + 1) To NumBullets - 1
        If Bullet(i).Owner <> Bullet(j).Owner Then
            If GetDist(Bullet(i).X, Bullet(i).Y, Bullet(j).X, Bullet(j).Y) < (Bullet_RadiusX2) Then
                
                TmpHeading = Bullet(i).Heading
                TmpSpeed = Bullet(i).Speed
                
                Bullet(i).Heading = Bullet(j).Heading
                Bullet(i).Speed = Bullet(j).Speed
                
                Bullet(j).Heading = TmpHeading
                Bullet(j).Speed = TmpSpeed
            End If
        End If
    Next j
Next i

End Sub

Private Sub CheckPlayerCollisions()
Dim i As Integer, j As Integer
Dim R1 As Single, R2 As Single
Dim Mass1 As Single, Mass2 As Single

Dim IsInfil As Boolean

'Dim xSide As Integer, ySide As Integer, k As Integer
'Dim XComp As Single, YComp As Single,
Dim Damage As Single, TmpHeading As Single, TmpSpeed As Single

For i = 0 To NumPlayers - 1
    If PlayerInGame(i) Then
        
        For j = (i + 1) To NumPlayers - 1
            
            If PlayerInGame(j) Then
                
                If Player(i).ShipType = Infiltrator Then
                    If (Player(i).State And Player_Secondary) = Player_Secondary Then
                        Exit For
                    End If
                End If
                
                If Player(j).ShipType = Infiltrator Then
                    If (Player(j).State And Player_Secondary) = Player_Secondary Then
                        Exit For
                    End If
                End If
                
                R1 = GetShipRadius(Player(i).ShipType)
                R2 = GetShipRadius(Player(j).ShipType)
                Mass1 = GetShipMass(Player(i).ShipType)
                Mass2 = GetShipMass(Player(j).ShipType)
                
                If GetDist(Player(i).X, Player(i).Y, Player(j).X, Player(j).Y) < (R1 + R2) Then
                    'ships are colliding...
                    
                    Damage = ((Player(i).Speed * Mass1) + (Player(j).Speed * Mass2)) / 13
                    
                    
                    TmpHeading = Player(j).Heading
                    Player(j).Heading = Player(i).Heading
                    Player(i).Heading = TmpHeading
                    
                    TmpSpeed = Player(j).Speed
                    Player(j).Speed = Player(i).Speed * Mass1 / Mass2
                    Player(i).Speed = TmpSpeed * Mass2 / Mass1
                    'swap momentum
                    
                    
                    'apply damage to first ship
                    If Player(i).ID = MyID Or (Player(i).IsBot And modSpaceGame.SpaceServer) Then
                        
                        IsInfil = (Player(i).ShipType = Infiltrator) 'Or (Player(i).ShipType = Infiltrator)
                        
                        If Player(i).Shields > Damage / Mass1 And Not IsInfil Then
                            Player(i).Shields = Player(i).Shields - Damage / Mass1
                        Else
                            Player(i).Hull = Player(i).Hull - Damage / Mass1 + IIf(IsInfil, 0, Player(i).Shields)
                            
                            If IsInfil = False Then
                                Player(i).Shields = 0
                            End If
                        End If
                        
                        
                        'Check for death
                        If Player(i).Hull < 1 Then Call Killed(j, i, , True)
                    End If
                    
                    'apply damage to second ship
                    If Player(j).ID = MyID Or (Player(j).IsBot And modSpaceGame.SpaceServer) Then
                        
                        IsInfil = (Player(i).ShipType = Infiltrator)
                        
                        If Player(j).Shields > Damage / Mass2 And Not IsInfil Then
                            Player(j).Shields = Player(j).Shields - Damage / Mass2
                        Else
                            Player(j).Hull = Player(j).Hull - Damage / Mass2 + IIf(IsInfil, 0, Player(j).Shields)
                            
                            If IsInfil = False Then
                                Player(i).Shields = 0
                            End If
                        End If
                        
                        
                        'Check for death
                        If Player(j).Hull < 1 Then Call Killed(i, j, , True)
                    End If
                    
                    Exit For
                End If
            
            End If 'player j <> spec endif
            
        Next j
        
    End If 'player i <> spec endif
    
Next i

For i = 0 To NumPlayers - 1
    If PlayerInGame(i) Then
        If PlayerIsInAsteroid(i) Then
            TmpHeading = Player(i).Heading
            TmpSpeed = Player(i).Speed
            
            Player(i).Heading = Asteroid.Heading
            Player(i).Speed = Asteroid.Speed
            
            Asteroid.Heading = TmpHeading
            Asteroid.Speed = TmpSpeed
            'swap 'momentum'
            
            
            If Player(i).ID = MyID Or Player(i).IsBot Then
                
                Mass1 = GetShipMass(Player(i).ShipType)
                
                Damage = (Player(i).Speed * Mass1 + Asteroid.Speed * AsteroidMass / 100)
                
                If Player(i).Shields > (Damage / Mass1) Then
                    Player(i).Shields = Player(i).Shields - Damage / Mass1
                Else
                    Player(i).Hull = Player(i).Hull - Damage / Mass1 + Player(i).Shields
                    Player(i).Shields = 0
                    
                    If Asteroid.LastPlayerTouchID <> -1 Then
                        If FindPlayer(Asteroid.LastPlayerTouchID) <> -1 Then
                            If Asteroid.LastPlayerTouchID <> Player(i).ID Then
                                If Player(i).Hull < 1 Then
                                    Call Killed(FindPlayer(Asteroid.LastPlayerTouchID), i, , True)
                                End If
                            End If
                        Else
                            Asteroid.LastPlayerTouchID = -1
                        End If
                    End If
                    
                End If
                
                If Player(i).Hull < 1 Then Player(i).Hull = 1
                
            End If
            
            Asteroid.LastPlayerTouchID = Player(i).ID
            
        End If
    End If
Next i

End Sub

'old method
'With Player(i)
'    XComp = .Speed * sine(.Heading)
'    YComp = .Speed * cosine(.Heading)
'
'    xSide = 0
'    ySide = 0
'
'    If .X > Player(j).X Then  'is on right side of j
'        XComp = Abs(XComp)
'        xSide = 1
'    Else 'If .X > (Player(j).X + Lim) Then 'is on left side of j
'        XComp = -Abs(XComp)
'        xSide = -1
'    End If
'
'    If .Y > Player(j).Y Then 'is below j
'        YComp = Abs(YComp)
'        ySide = 1
'    Else 'If .Y > (Player(j).Y + Lim) Then 'is above j
'        YComp = -Abs(YComp)
'        ySide = -1
'    End If
'
'    k = 0
'
'    'Determine the resultant speed
'    If .Speed > Player(j).Speed Then
'        .Speed = Sqr(XComp ^ 2 + YComp ^ 2) * 1.5
'        k = 1
'    Else
'        .Speed = Sqr(XComp ^ 2 + YComp ^ 2) / 1.5
'        k = -1
'    End If
'
'    'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
'    If YComp > 0 Then .Heading = atn(XComp / YComp)
'    If YComp < 0 Then .Heading = atn(XComp / YComp) + pi
'
'    If .Shields > 1 Then
'        .Shields = .Shields - TempMag
''                ElseIf .Hull > TempMag Then
''                    .Hull = .Hull - TempMag / 2
'    End If
'
'End With
'
'
'With Player(j)
'    XComp = .Speed * sine(.Heading)
'    YComp = .Speed * cosine(.Heading)
'
'
'    If xSide = 1 Then 'is on right side of j
'        XComp = -Abs(XComp)
'    Else 'If xSide = -1 Then 'is on left side of j
'        XComp = Abs(XComp)
'    End If
'
'    If ySide = 1 Then
'        YComp = -Abs(YComp)
'    Else 'If ySide = -1 Then
'        YComp = Abs(YComp)
'    End If
'
'    'Determine the resultant speed
'    If k = 1 Then
'        .Speed = Sqr(XComp ^ 2 + YComp ^ 2) / 1.5
'    Else
'        .Speed = Sqr(XComp ^ 2 + YComp ^ 2) * 1.5
'    End If
'
'    'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
'    If YComp > 0 Then .Heading = atn(XComp / YComp)
'    If YComp < 0 Then .Heading = atn(XComp / YComp) + pi
'
'
'    If .Shields > 1 Then
'        .Shields = .Shields - TempMag
''                ElseIf .Hull > TempMag Then
''                    .Hull = .Hull - TempMag / 2
'    End If
'
'End With

Private Sub InitVariables()
Dim RGBCol As ptRGB

'modSpaceGame.SpaceServer = modSpaceGame.SpaceServer
Playing = True

'modSpaceGame.MaxHeight = DefaultMaxHeight
'modSpaceGame.MaxWidth = DefaultMaxWidth

'Add us as a player!
AddPlayer

'If we're the host, assign our ID now
If modSpaceGame.SpaceServer Then
    Player(0).ID = 0
    MyID = 0
End If

'Get a random spot on the screen

Call RandomizePlayer

'Set name
Player(0).Name = frmMain.txtName.Text

Player(0).Colour = modVars.TxtForeGround

RGBCol = modSpaceGame.RGBDecode(Player(0).Colour)

If RGBCol.Red < 150 And RGBCol.Blue < 150 And RGBCol.Green < 150 Then
    
    'Player(0).Colour = RandomRGBColour()
    Player(0).Colour = RGB(150 + Rnd() * 105, _
                            150 + Rnd() * 105, _
                            150 + Rnd() * 105) '100 + Rnd * (255-150), _

    AddChatText "Colour is Too Dark - Ship Colour Set to Random", vbRed
'Else
    'Player(0).Colour = modVars.TxtForeGround
End If

AI_Sample_Rate = modSpaceGame.Default_AI_Sample_Rate

InitStars

Call RndAsteroid
Asteroid.X = MaxWidth - Asteroid_Radius - 500
Asteroid.Y = Asteroid_Radius + 500

'add script obj for cheating, etc
'frmMain.SC.AddObject "Player(0)", Player(0), True

m_bDesignMode = modSpaceGame.SpaceEditing
modSpaceGame.GameFormLoaded = True
ROTATION_RATE = Default_Rotation_Rate
'mPacket_SEND_DELAY = Default_mPacket_SEND_DELAY

sv_Bullet_Damage = Default_Bullet_Damage
modSpaceGame.sv_CTFTime = 5
cg_Cls = True
FlagOwnerID = -1

ResetCamera

'initialize drawing/texting
'modSpaceGame.SetBkMode Me.hdc, modSpaceGame.TEXT_TRANSPARENT
'frmGame.Print "Shouldn't See Me"

End Sub

Public Sub SetCursor(bHide As Boolean)

If bHide Then
    Me.MousePointer = vbCustom
    Me.MouseIcon = picBlank.Picture
Else
    Me.MousePointer = vbDefault
End If

End Sub

Private Sub RandomizePlayer()
Player(0).X = MaxWidth * Rnd()
Player(0).Y = (MaxHeight - 500) * Rnd()
Player(0).Facing = Pi2 * Rnd()

'Set shields
Player(0).Shields = SHIELD_START
Player(0).MaxShields = SHIELD_START

Player(0).Hull = Hull_Start
Player(0).MaxHull = Hull_Start

'Player(0).Score = 0
End Sub

Public Function AddPlayer() As Integer ' Optional ByVal bIsBot As Boolean = False) As Integer

'Add a player onto the array, and return his index
ReDim Preserve Player(NumPlayers)
'ReDim Preserve ScoreList(NumPlayers)

'ScoreList(NumPlayers).ID = Player(NumPlayers).ID
Player(NumPlayers).LastPacket = GetTickCount()
Player(NumPlayers).Alive = True
'Player(NumPlayers).IsBot = bIsBot

AddPlayer = NumPlayers

NumPlayers = NumPlayers + 1


End Function

Public Sub RemovePlayer(Index As Integer)

Dim i As Integer
On Error Resume Next

If Index = -1 Then Exit Sub

If modSpaceGame.SpaceServer Then
    SendBroadcast sRemovePlayers & CStr(Player(Index).ID)
End If

'Remove this player from the array
For i = Index To NumPlayers - 2
    Player(i).ID = Player(i + 1).ID
    Player(i).LastPacket = Player(i + 1).LastPacket
    Player(i).LastPacketID = Player(i + 1).LastPacketID
    Player(i).Name = Player(i + 1).Name
    Player(i).Facing = Player(i + 1).Facing
    Player(i).Heading = Player(i + 1).Heading
    Player(i).Speed = Player(i + 1).Speed
    Player(i).X = Player(i + 1).X
    Player(i).Y = Player(i + 1).Y
    Player(i).LastBullet = Player(i + 1).LastBullet
    Player(i).Shields = Player(i + 1).Shields
    Player(i).Hull = Player(i + 1).Hull
    Player(i).bDrawShields = Player(i + 1).bDrawShields
    Player(i).bRightBullet = Player(i + 1).bRightBullet
    Player(i).MaxShields = Player(i + 1).MaxShields
    Player(i).MaxHull = Player(i + 1).MaxHull
    Player(i).bDrawShields = Player(i + 1).bDrawShields
    Player(i).bRightBullet = Player(i + 1).bRightBullet
    Player(i).State = Player(i + 1).State
    
    Player(i).ptSockAddr.sin_addr = Player(i + 1).ptSockAddr.sin_addr
    Player(i).ptSockAddr.sin_family = Player(i + 1).ptSockAddr.sin_family
    Player(i).ptSockAddr.sin_port = Player(i + 1).ptSockAddr.sin_port
    Player(i).ptSockAddr.sin_zero = Player(i + 1).ptSockAddr.sin_zero
    
    Player(i).Colour = Player(i + 1).Colour
    Player(i).ShipType = Player(i + 1).ShipType
    
    Player(i).Kills = Player(i + 1).Kills
    Player(i).Deaths = Player(i + 1).Deaths
    Player(i).IsBot = Player(i + 1).IsBot
    
    Player(i).LastSecondary = Player(i + 1).LastSecondary
    
    Player(i).Team = Player(i + 1).Team
    
    Player(i).AITimer = Player(i + 1).AITimer
    Player(i).LastAITargetIndex = Player(i + 1).LastAITargetIndex
    
    Player(i).Score = Player(i + 1).Score
    
    Player(i).Alive = Player(i + 1).Alive
    
    Player(i).LastSmoke = Player(i + 1).LastSmoke
Next i

'Resize the array
ReDim Preserve Player(NumPlayers - 2)
NumPlayers = NumPlayers - 1


End Sub

Private Function PlayerThrusting(Index As Integer) As Boolean

PlayerThrusting = ((Player(Index).State And PLAYER_THRUST) = PLAYER_THRUST Or _
                   (Player(Index).State And PLAYER_REVTHRUST) = PLAYER_REVTHRUST Or _
                   (Player(Index).State And Player_StrafeLeft) = Player_StrafeLeft Or _
                   (Player(Index).State And Player_StrafeRight) = Player_StrafeRight)


End Function

Private Sub DisplayPlayers()

Dim i As Integer
Dim Name As String
Dim w As Single
Dim bCan As Boolean
Dim Tm As Long

'Dim rMag As Single, rDir As Single

'Tm = LastRespawn + RespawnCircleShowTime - GetTickCount()

'If Tm > 0 Then
    'picMain.DrawWidth = 5
    'picMain.FillStyle = 1
    'gCircle Player(i).X, Player(i).Y, RespawnCircleRadius + Tm, MLightBlue 'Player(i).Colour
    'picMain.DrawWidth = 1
'End If

If PlayerInGame(0) Then
    MoveCameraX Player(0).X * cg_Zoom - CentreX
    MoveCameraY Player(0).Y * cg_Zoom - CentreY
End If

'modSpaceGame.cg_Camera.X = Player(i).X - cg_Camera.X - (1 - cg_Zoom) * CentreX
'modSpaceGame.cg_Camera.Y = Player(i).Y - cg_Camera.Y - (1 - cg_Zoom) * CentreY

'Loop through each player and display
For i = 0 To NumPlayers - 1
    
    If PlayerInGame(i) Then 'Player(i).Team <> Spec Then
        
        If modSpaceGame.cg_Smoke Then
            If PlayerThrusting(i) Then
                If Player(i).ShipType <> SD Then
                    If Player(i).ShipType <> MotherShip Then
                        If Player(i).ShipType <> Infiltrator Then
                            If Player(i).LastSmoke + 20 / modSpaceGame.sv_GameSpeed < GetTickCount() Then
                                
                                'AddVectors Player(i).Speed, Player(i).Heading, 100, Player(i).Facing, rMag, rDir
                                
                                AddSmokeGroup Player(i).X, Player(i).Y, 4 ', Player(i).Speed, Player(i).Heading
                                Player(i).LastSmoke = GetTickCount()
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        'Display him
        DrawShip i
        
        
        bCan = True
        
        'Display his name
        If (Player(i).ShipType = Infiltrator) And _
            ((Player(i).State And Player_Secondary) = Player_Secondary) Then
            
            bCan = False
            
        End If
    
        If Player(i).ID = MyID Then bCan = True
        
        
        If bCan Then
            Name = Trim$(Player(i).Name)
            'Me.ForeColor = IIf(modSpaceGame.cg_BlackBG, vbWhite, vbBlack)
            'Me.ForeColor = vbWhite
            
            Select Case Player(i).ShipType
                Case eShipTypes.Raptor, eShipTypes.Hornet
                    w = 210
                Case eShipTypes.Behemoth, eShipTypes.Wraith, eShipTypes.Infiltrator
                    w = 450
                Case eShipTypes.MotherShip, eShipTypes.SD
                    w = 750
            End Select
            
            PrintText Name, Player(i).X + w, Player(i).Y, vbWhite
            
            
            gCircle Player(i).X, Player(i).Y, _
                IIf(Player(i).ShipType <> Behemoth And Player(i).ShipType <> Wraith, 1, 2), _
                    GetTeamColour(Player(i).Team) 'don't show them
            
            If modSpaceGame.SpaceServer Then
                If Player(i).IsBot Then
                    PrintText "Hull: " & CStr(Round(Player(i).Hull)), Player(i).X + w, Player(i).Y + 200, vbWhite
                    PrintText "Shields: " & CStr(Round(Player(i).Shields)), Player(i).X + w, Player(i).Y + 400, vbWhite
                End If
            End If
        End If
        'PrintText "Hull: " & CStr(Round(Player(i).Hull)), Player(i).X + 210, Player(i).Y + 200
        'PrintText "Shield: " & CStr(Round(Player(i).Hull)), Player(i).X + 210, Player(i).Y + 400
        
        'ShowText Name, Player(i).x, Player(i).y, vbBlack, Me.hdc
        '- Me.TextWidth(Name)) \ 2
    End If
Next i

End Sub

Private Function GetTeamColour(ByVal vTeam As eTeams) As Long
Select Case vTeam
    Case eTeams.Neutral, eTeams.Spec
        GetTeamColour = MGrey
    Case eTeams.Red
        GetTeamColour = vbRed
    Case eTeams.Blue
        GetTeamColour = vbBlue
End Select
End Function

Private Sub DrawShip(i As Integer)
Dim Col As Long
Dim bDrawSideAccel As Boolean
Dim bCan As Boolean
Dim A1 As Single, A2 As Single
Dim r As Single

'Col = Player(i).Colour 'IIf(MyID = Player(i).ID, vbBlue, vbBlack)
Col = GetTeamColour(Player(i).Team)

picMain.DrawWidth = IIf(modSpaceGame.cg_DrawThick, Thin * 2, Thin)
picMain.FillStyle = 1 'vbTransparent

'draw side accel
If Player(i).ShipType <> MotherShip And Player(i).ShipType <> SD Then
    
    If Player(i).ID = MyID Then
        bCan = True
    ElseIf Not (Player(i).ShipType = Infiltrator And ((Player(i).State And Player_Secondary) = Player_Secondary)) Then
        bCan = True
    End If
    
    If bCan Then
        If (Player(i).State And ePlayerState.Player_StrafeRight) = ePlayerState.Player_StrafeRight Then
            
            A1 = FixAngle(3.5 * piD4 - Player(i).Facing)
            A2 = FixAngle(4.5 * piD4 - Player(i).Facing)
            r = SHIP_Height
            bDrawSideAccel = True
            
        ElseIf (Player(i).State And ePlayerState.Player_StrafeLeft) = ePlayerState.Player_StrafeLeft Then
            
            A1 = FixAngle(-(0.15 * Pi + Player(i).Facing))
            A2 = FixAngle(0.16 * Pi - Player(i).Facing)
            r = SHIP_Height
            bDrawSideAccel = True
            
            'gCircle (x,y), radius, colour, start, end, aspect/curveness
        End If
        
        If bDrawSideAccel Then
            gCircleSE Player(i).X, Player(i).Y, r, vbRed, A1, A2
            'gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
        End If
        
    End If
End If


If Player(i).ShipType = eShipTypes.Raptor Then
    Call DrawRaptor(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 5
ElseIf Player(i).ShipType = eShipTypes.Behemoth Then
    Call DrawBehemoth(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 7
ElseIf Player(i).ShipType = eShipTypes.Hornet Then
    Call DrawHornet(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 7
ElseIf Player(i).ShipType = MotherShip Then
    Call DrawMothership(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 8
ElseIf Player(i).ShipType = Wraith Then
    Call DrawWraith(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 7
ElseIf Player(i).ShipType = Infiltrator Then
    Call DrawInfiltrator(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 7
Else
    Call DrawSD(i, Player(i).Colour, Col)
    picMain.DrawWidth = Thin * 8
End If


'If (Player(i).State And Player_Secondary) = Player_Secondary) Then
'    If Player(i).ID = MyID Then
'
'    End If
'End If


'gCircle (Player(i).X + 50 * sine(Player(i).Facing), _
    Player(i).Y - 50 * cosine(Player(i).Facing)), _
        1, Col, , , 1.1


End Sub

Private Sub DrawRaptor(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

'Dim x1 As Single
'Dim y1 As Single
'Dim X2 As Single
'Dim Y2 As Single
'Dim x3 As Single
'Dim y3 As Single
Dim Pt(1 To 4) As POINTAPI
Dim G1X As Single
Dim G1Y As Single
Dim G2X As Single
Dim G2Y As Single
Dim ShieldX As Single, ShieldY As Single

Dim A1 As Single, A2 As Single
Dim bDrawAccel As Boolean ', bDrawSideAccel As Boolean
Dim r As Single

On Error GoTo EH
''Calculate the new ship vertices
'x1 = Player(i).X + SHIP_Height * sine(Player(i).Facing)
'y1 = Player(i).Y - SHIP_Height * cosine(Player(i).Facing)
'X2 = Player(i).X + SHIP_RADIUS * sine(Player(i).Facing + Pi2D3)
'Y2 = Player(i).Y - SHIP_RADIUS * cosine(Player(i).Facing + Pi2D3)
'x3 = Player(i).X + SHIP_RADIUS * sine(Player(i).Facing + 4 * piD3)
'y3 = Player(i).Y - SHIP_RADIUS * cosine(Player(i).Facing + 4 * piD3)
Pt(1).X = Player(i).X + SHIP_Height * Sine(Player(i).Facing)
Pt(1).Y = Player(i).Y - SHIP_Height * CoSine(Player(i).Facing)
Pt(2).X = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing + pi2d3)
Pt(2).Y = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing + pi2d3)
Pt(3).X = Player(i).X
Pt(3).Y = Player(i).Y
Pt(4).X = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing + 4 * piD3)
Pt(4).Y = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing + 4 * piD3)


G1X = Pt(2).X + Gun_Len * Sine(Pi - Player(i).Facing + GunOffset * 4.5)
G1Y = Pt(2).Y + Gun_Len * CoSine(Pi - Player(i).Facing + GunOffset * 4.5)
G2X = Pt(4).X + Gun_Len * Sine(Pi - Player(i).Facing - GunOffset * 4.5)
G2Y = Pt(4).Y + Gun_Len * CoSine(Pi - Player(i).Facing - GunOffset * 4.5)


'gLine x1, y1, X2, Y2, Col
'gLine Player(i).X, Player(i).Y, X2, Y2, tCol
'gLine x3, y3, Player(i).X, Player(i).Y, tCol
'gLine x3, y3, x1, y1, Col

'gunz
gLine CSng(Pt(2).X), CSng(Pt(2).Y), G1X, G1Y, Col
gLine CSng(Pt(4).X), CSng(Pt(4).Y), G2X, G2Y, Col

picMain.ForeColor = tCol
modSpaceGame.gPoly Pt, Player(i).Colour

If Player(i).bDrawShields Then
    gCircleAspect Player(i).X + 50 * Sine(Player(i).Facing), _
        Player(i).Y - 50 * CoSine(Player(i).Facing), _
            SHIP_Height + 25, vbGreen, 1.1
    
    Player(i).bDrawShields = False
End If

If (Player(i).State And Player_Shielding) = Player_Shielding Then
    
    
    gCircle Player(i).X + 60 * Sine(Player(i).Facing), _
        Player(i).Y - 60 * CoSine(Player(i).Facing), _
            SHIP_Height + 28, vbGreen
    
End If


If (Player(i).State And ePlayerState.PLAYER_THRUST) = ePlayerState.PLAYER_THRUST Then
    
    A1 = FixAngle(1.25 * Pi - Player(i).Facing)
    A2 = FixAngle(1.75 * Pi - Player(i).Facing)
    r = SHIP_RADIUS \ 1.5
    bDrawAccel = True
    
ElseIf (Player(i).State And ePlayerState.PLAYER_REVTHRUST) = ePlayerState.PLAYER_REVTHRUST Then
    
    A1 = FixAngle(0.35 * Pi - Player(i).Facing)
    A2 = FixAngle(0.66 * Pi - Player(i).Facing)
    r = SHIP_Height + 30
    bDrawAccel = True

    'gCircle (x,y), radius, colour, start, end, aspect/curveness
End If

If bDrawAccel Then
    gCircleSE Player(i).X, Player(i).Y, r, vbRed, A1, A2
    'gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
End If


EH:
End Sub

Private Sub DrawBehemoth(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

'Dim x1 As Single
'Dim y1 As Single
'Dim X2 As Single
'Dim Y2 As Single
'Dim x3 As Single
'Dim y3 As Single
'Dim G1X As Single
'Dim G1Y As Single
'Dim x4 As Single
'Dim y4 As Single
'Dim fx1 As Single, fy1 As Single
'Dim fx2 As Single, fy2 As Single
Dim Pt(1 To 7) As POINTAPI

Dim c1x As Single, c1y As Single
Dim c2x As Single, c2y As Single
Dim c3x As Single, c3y As Single

Dim A1 As Single, A2 As Single
Dim bDrawAccel As Boolean
Dim r As Single

On Error GoTo EH
'x1 = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + piD4)  '45
'y1 = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + piD4) '45
'X2 = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + pi3D4) '135
'Y2 = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + pi3D4) '135
'x3 = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + 5 * piD4) '225
'y3 = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + 5 * piD4) '225
'x4 = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + 7 * piD4) '315
'y4 = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + 7 * piD4) '315
'
'G1X = Player(i).X + SHIP_RADIUS * 2.5 * sine(Player(i).Heading) '315
'G1Y = Player(i).Y - SHIP_RADIUS * 2.5 * cosine(Player(i).Heading) '315
'
'fx1 = Player(i).X + SHIP_Height * 1.2 * sine(Player(i).Heading + pi3D4)  '135
'fy1 = Player(i).Y - SHIP_Height * 1.2 * cosine(Player(i).Heading + pi3D4)  '135
'fx2 = Player(i).X + SHIP_Height * 1.2 * sine(Player(i).Heading + 5 * piD4)  '225
'fy2 = Player(i).Y - SHIP_Height * 1.2 * cosine(Player(i).Heading + 5 * piD4)  '225


'Pt(1).X = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + piD4)  '45
'Pt(1).Y = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + piD4) '45
'Pt(2).X = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + pi3D4) '135
'Pt(2).Y = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + pi3D4) '135
'Pt(3).X = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + 5 * piD4) '225
'Pt(3).Y = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + 5 * piD4) '225
'Pt(4).X = Player(i).X + SHIP_Height / 1.5 * sine(Player(i).Heading + 7 * piD4) '315
'Pt(4).Y = Player(i).Y - SHIP_Height / 1.5 * cosine(Player(i).Heading + 7 * piD4) '315
'Pt(5).X = Player(i).X + SHIP_RADIUS * 2.5 * sine(Player(i).Heading) '315
'Pt(5).Y = Player(i).Y - SHIP_RADIUS * 2.5 * cosine(Player(i).Heading) '315
'Pt(6).X = Player(i).X + SHIP_Height * 1.2 * sine(Player(i).Heading + pi3D4)  '135
'Pt(6).Y = Player(i).Y - SHIP_Height * 1.2 * cosine(Player(i).Heading + pi3D4)  '135
'Pt(7).X = Player(i).X + SHIP_Height * 1.2 * sine(Player(i).Heading + 5 * piD4)  '225
'Pt(7).Y = Player(i).Y - SHIP_Height * 1.2 * cosine(Player(i).Heading + 5 * piD4)  '225


Pt(1).X = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Heading + piD4)  '45
Pt(1).Y = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Heading + piD4) '45
Pt(3).X = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Heading + pi3D4) '135
Pt(3).Y = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Heading + pi3D4) '135
Pt(2).X = Player(i).X + SHIP_Height * 1.2 * Sine(Player(i).Heading + pi3D4)  '135
Pt(2).Y = Player(i).Y - SHIP_Height * 1.2 * CoSine(Player(i).Heading + pi3D4)  '135
Pt(5).X = Player(i).X + SHIP_Height * 1.2 * Sine(Player(i).Heading + 5 * piD4)  '225
Pt(5).Y = Player(i).Y - SHIP_Height * 1.2 * CoSine(Player(i).Heading + 5 * piD4)  '225
Pt(4).X = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Heading + 5 * piD4) '225
Pt(4).Y = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Heading + 5 * piD4) '225
Pt(6).X = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Heading + 7 * piD4) '315
Pt(6).Y = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Heading + 7 * piD4) '315
Pt(7).X = Player(i).X + SHIP_RADIUS * 2.5 * Sine(Player(i).Heading) '315
Pt(7).Y = Player(i).Y - SHIP_RADIUS * 2.5 * CoSine(Player(i).Heading) '315


c1x = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Facing + piD4)
c1y = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Facing + piD4)
c2x = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Facing - piD4)
c2y = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Facing - piD4)
c3x = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing)  '315
c3y = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing)  '315



picMain.FillColor = 0
'gLine x1, y1, X2, Y2, Col
'gLine x3, y3, x4, y4, Col
'gLine G1X, G1Y, x4, y4, Col
'gLine G1X, G1Y, x1, y1, Col
'gLine fx1, fy1, X2, Y2, Col
'gLine fx1, fy1, x3, y3, Col
'gLine fx2, fy2, X2, Y2, Col
'gLine fx2, fy2, x3, y3, Col
picMain.ForeColor = tCol
modSpaceGame.gPoly Pt, Player(i).Colour

gLine Player(i).X, Player(i).Y, c1x, c1y, tCol
gLine Player(i).X, Player(i).Y, c2x, c2y, tCol

gCircle c1x, c1y, 50, tCol
gCircle c2x, c2y, 50, tCol

gLine Player(i).X, Player(i).Y, c3x, c3y, tCol
gLine c2x, c2y, c1x, c1y, tCol

'extras
If Player(i).bDrawShields Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 2, vbGreen  ', , , -1
    Player(i).bDrawShields = False
End If

If (Player(i).State And Player_Shielding) = Player_Shielding Then
    
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 2.2, vbGreen
    
End If

If (Player(i).State And ePlayerState.PLAYER_THRUST) = ePlayerState.PLAYER_THRUST Then
    
    A1 = FixAngle(1.25 * Pi - Player(i).Facing) '1.25 *
    A2 = FixAngle(1.75 * Pi - Player(i).Facing)
    r = SHIP_RADIUS * 1.5
    bDrawAccel = True
    
ElseIf (Player(i).State And ePlayerState.PLAYER_REVTHRUST) = ePlayerState.PLAYER_REVTHRUST Then
    
    A1 = FixAngle(0.35 * Pi - Player(i).Facing)
    A2 = FixAngle(0.66 * Pi - Player(i).Facing)
    r = SHIP_Height * 1.2
    bDrawAccel = True

    'gCircle (x,y), radius, colour, start, end, aspect/curveness
End If

If bDrawAccel Then
    gCircleSE Player(i).X, Player(i).Y, r, vbRed, A1, A2
    'gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
End If

EH:
End Sub

Private Sub DrawHornet(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

'Dim x1 As Single
'Dim y1 As Single
'Dim X2 As Single
'Dim Y2 As Single
Dim x3 As Single
Dim y3 As Single
Dim G1X As Single
Dim G1Y As Single
Dim G2X As Single
Dim G2Y As Single

Dim x4 As Single
Dim y4 As Single
'Dim c1x As Single, c1y As Single
Dim Pt(1 To 3) As POINTAPI

Dim A1 As Single, A2 As Single
Dim bDrawAccel As Boolean
Dim r As Single

On Error GoTo EH
'main triangle
Pt(1).X = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing - piD2) 'left wing
Pt(1).Y = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing - piD2)
Pt(2).X = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing + piD2) 'right wing
Pt(2).Y = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing + piD2)
Pt(3).X = Player(i).X + SHIP_RADIUS * Sine(Player(i).Facing)
Pt(3).Y = Player(i).Y - SHIP_RADIUS * CoSine(Player(i).Facing)


x3 = Pt(1).X + SHIP_RADIUS * Sine(Player(i).Facing) 'left panel front
y3 = Pt(1).Y - SHIP_RADIUS * CoSine(Player(i).Facing)
x4 = Pt(1).X - SHIP_RADIUS * Sine(Player(i).Facing) / 1.5 'left panel back
y4 = Pt(1).Y + SHIP_RADIUS * CoSine(Player(i).Facing) / 1.5

G1X = Pt(2).X + SHIP_RADIUS * Sine(Player(i).Facing) 'right panel front
G1Y = Pt(2).Y - SHIP_RADIUS * CoSine(Player(i).Facing)
G2X = Pt(2).X - SHIP_RADIUS * Sine(Player(i).Facing) / 1.5 'right panel back
G2Y = Pt(2).Y + SHIP_RADIUS * CoSine(Player(i).Facing) / 1.5


gLine x3, y3, x4, y4, Col 'panels
gLine G1X, G1Y, G2X, G2Y, Col

'triangle
'gLine x1, y1, X2, Y2, Col
'gLine x1, y1, c1x, c1y, Col
'gLine X2, Y2, c1x, c1y, Col
picMain.ForeColor = tCol
modSpaceGame.gPoly Pt, Player(i).Colour

picMain.FillStyle = 0
picMain.FillColor = tCol
gCircle Player(i).X, Player(i).Y, SHIP_RADIUS \ 2, tCol
picMain.FillStyle = 1

If Player(i).bDrawShields Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height, vbGreen
    Player(i).bDrawShields = False
End If

If (Player(i).State And Player_Shielding) = Player_Shielding Then
    
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 1.2, vbGreen
    
End If


If (Player(i).State And ePlayerState.PLAYER_THRUST) = ePlayerState.PLAYER_THRUST Then
    
    A1 = FixAngle(1.25 * Pi - Player(i).Facing)
    A2 = FixAngle(1.75 * Pi - Player(i).Facing)
    r = SHIP_Height / 1.5
    bDrawAccel = True
    
ElseIf (Player(i).State And ePlayerState.PLAYER_REVTHRUST) = ePlayerState.PLAYER_REVTHRUST Then
    
    A1 = FixAngle(0.35 * Pi - Player(i).Facing)
    A2 = FixAngle(0.66 * Pi - Player(i).Facing)
    r = SHIP_Height
    bDrawAccel = True

    'gCircle (x,y), radius, colour, start, end, aspect/curveness
End If

If bDrawAccel Then
    gCircleSE Player(i).X, Player(i).Y, r, vbRed, A1, A2
    'gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
End If

EH:
End Sub

Private Sub DrawMothership(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

Dim X1 As Single, Y1 As Single
Dim X2 As Single, Y2 As Single
Dim x3 As Single, y3 As Single
Dim x4 As Single, y4 As Single

Dim nx1 As Single, ny1 As Single
Dim nx2 As Single, ny2 As Single
Dim nx3 As Single, ny3 As Single
Dim nx4 As Single, ny4 As Single

Dim c1x As Single, c1y As Single
Dim c2x As Single, c2y As Single
Dim c3x As Single, c3y As Single
Dim c4x As Single, c4y As Single

Dim G1X As Single, G1Y As Single
Dim G2X As Single, G2Y As Single

Dim mg1x As Single, mg1y As Single

'Dim A1 As Single, A2 As Single, R As Single
'Dim bDrawAccel As Boolean

Const cWidth As Single = SHIP_Height / 5

On Error GoTo EH
X1 = Player(i).X - SHIP_Height * 3
Y1 = Player(i).Y - SHIP_RADIUS
X2 = Player(i).X + SHIP_Height * 3
Y2 = Player(i).Y - SHIP_RADIUS
x3 = Player(i).X + SHIP_Height * 3
y3 = Player(i).Y + SHIP_RADIUS
x4 = Player(i).X - SHIP_Height * 3
y4 = Player(i).Y + SHIP_RADIUS

nx1 = X1 + SHIP_Height * 2
ny1 = Y1 - SHIP_RADIUS
nx2 = X2 - SHIP_Height * 2
ny2 = Y1 - SHIP_RADIUS
nx3 = X2 - SHIP_Height * 2
ny3 = y3 + SHIP_RADIUS
nx4 = X1 + SHIP_Height * 2
ny4 = y3 + SHIP_RADIUS

c1x = Player(i).X + SHIP_RADIUS * 2
c1y = Player(i).Y
c2x = Player(i).X + SHIP_RADIUS * 3.5
c2y = c1y
c3x = Player(i).X - SHIP_RADIUS * 2
c3y = c1y
c4x = Player(i).X - SHIP_RADIUS * 3.5
c4y = c1y

G1X = Player(i).X + SHIP_Height * 2 * Sine(Player(i).Facing + Pi / 8)
G1Y = Player(i).Y - SHIP_Height * 2 * CoSine(Player(i).Facing + Pi / 8)
G2X = Player(i).X + SHIP_Height * 2 * Sine(Player(i).Facing - Pi / 8)
G2Y = Player(i).Y - SHIP_Height * 2 * CoSine(Player(i).Facing - Pi / 8)

mg1x = c3x + Gun_Len * Sine(Player(i).Facing)
mg1y = c3y - Gun_Len * CoSine(Player(i).Facing)

gBox X1, Y1, x3, y3, Col
gLine X1, Y1, nx1, ny1, Col
gLine nx1, ny1, nx2, ny2, Col
gLine X2, Y2, nx2, ny2, Col
gLine nx3, ny3, x3, y3, Col
gLine nx3, ny3, nx4, ny4, Col
gLine x4, y4, nx4, ny4, Col

gCircle c1x, c1y, cWidth, Col
gCircle c2x, c2y, cWidth, Col
gCircle c3x, c3y, cWidth, tCol
gCircle c4x, c4y, cWidth, Col
gLine c3x, c3y, mg1x, mg1y, tCol

'gun
gCircle Player(i).X, Player(i).Y, 100, tCol
gLine G1X, G1Y, G2X, G2Y, tCol
gLine G1X, G1Y, Player(i).X, Player(i).Y, tCol
gLine G2X, G2Y, Player(i).X, Player(i).Y, tCol


If Player(i).bDrawShields Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 3.5, vbGreen
    Player(i).bDrawShields = False
End If


'If (Player(i).State And ePlayerState.PLAYER_THRUST) = ePlayerState.PLAYER_THRUST Then
'
'    A1 = FixAngle(1.25 * pi - Player(i).Facing)
'    A2 = FixAngle(1.75 * pi - Player(i).Facing)
'    R = SHIP_Height
'    bDrawAccel = True
'
'ElseIf (Player(i).State And ePlayerState.PLAYER_REVTHRUST) = ePlayerState.PLAYER_REVTHRUST Then
'
'    A1 = FixAngle(0.35 * pi - Player(i).Facing)
'    A2 = FixAngle(0.66 * pi - Player(i).Facing)
'    R = SHIP_Height * 3
'    bDrawAccel = True
'
'    'gCircle (x,y), radius, colour, start, end, aspect/curveness
'End If
'
'If bDrawAccel Then
'    gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
'    gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
'End If

EH:
End Sub

Private Sub DrawWraith(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

Dim xPts(1 To 15) As Single
Dim yPts(1 To 15) As Single
Dim G1X As Single, G1Y As Single, b1x As Single, b1y As Single
Dim G2X As Single, G2Y As Single, b2x As Single, b2y As Single

Dim A1 As Single, A2 As Single
Dim bDrawAccel As Boolean
Dim r As Single

Const ProngOffSet = 0.175 '10-ish degrees


On Error GoTo EH
xPts(1) = Player(i).X + SHIP_Height * 1.5 * Sine(Player(i).Heading - ProngOffSet)
yPts(1) = Player(i).Y - SHIP_Height * 1.5 * CoSine(Player(i).Heading - ProngOffSet)
xPts(2) = Player(i).X + SHIP_Height * 1.5 * Sine(Player(i).Heading + ProngOffSet)
yPts(2) = Player(i).Y - SHIP_Height * 1.5 * CoSine(Player(i).Heading + ProngOffSet)
xPts(3) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading - piD4) '45 angle
yPts(3) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading - piD4) '45 angle
xPts(4) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading - piD4 - ProngOffSet) '55 angle
yPts(4) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading - piD4 - ProngOffSet) '55 angle
xPts(5) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading - pi3D4) '(90+45) angle
yPts(5) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading - pi3D4) '(90+45) angle
xPts(6) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading - pi3D4 - ProngOffSet) '(90+55) angle
yPts(6) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading - pi3D4 - ProngOffSet) '(90+55) angle
xPts(7) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading + pi3D4 + ProngOffSet) '(90+55) angle
yPts(7) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading + pi3D4 + ProngOffSet) '(90+55) angle
xPts(8) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading + pi3D4) '(90+45) angle
yPts(8) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading + pi3D4) '(90+45) angle
xPts(9) = Player(i).X + SHIP_Height * Sine(Player(i).Heading + piD2)  '+90
yPts(9) = Player(i).Y - SHIP_Height * CoSine(Player(i).Heading + piD2)  '+90
xPts(10) = Player(i).X + SHIP_Height * Sine(Player(i).Heading + piD3)  '+60
yPts(10) = Player(i).Y - SHIP_Height * CoSine(Player(i).Heading + piD3)  '+60
xPts(11) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading + 3 * Pi / 10) '+54
yPts(11) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading + 3 * Pi / 10) '+54
xPts(12) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading + piD4 + ProngOffSet) '55 angle
yPts(12) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading + piD4 + ProngOffSet) '55 angle
xPts(13) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading + piD4) '45 angle
yPts(13) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading + piD4) '45 angle
xPts(14) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading + ProngOffSet)
yPts(14) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading + ProngOffSet)
xPts(15) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Heading - ProngOffSet)
yPts(15) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Heading - ProngOffSet)

G1X = Player(i).X + Gun_Len * 1.2 * Sine(Player(i).Facing + ProngOffSet * 2)
G1Y = Player(i).Y - Gun_Len * 1.2 * CoSine(Player(i).Facing + ProngOffSet * 2)
G2X = Player(i).X + Gun_Len * 1.2 * Sine(Player(i).Facing - ProngOffSet * 2)
G2Y = Player(i).Y - Gun_Len * 1.2 * CoSine(Player(i).Facing - ProngOffSet * 2)

b1x = Player(i).X + 100 * Sine(Player(i).Facing + piD4)
b1y = Player(i).Y - 100 * CoSine(Player(i).Facing + piD4)
b2x = Player(i).X + 100 * Sine(Player(i).Facing - piD4)
b2y = Player(i).Y - 100 * CoSine(Player(i).Facing - piD4)

picMain.FillColor = 0

'main bit
gLine xPts(1), yPts(1), xPts(3), yPts(3), Col
gLine xPts(3), yPts(3), xPts(4), yPts(4), Col
gLine xPts(4), yPts(4), xPts(5), yPts(5), Col
gLine xPts(5), yPts(5), xPts(6), yPts(6), Col
gLine xPts(6), yPts(6), xPts(7), yPts(7), Col
gLine xPts(7), yPts(7), xPts(8), yPts(8), Col
gLine xPts(8), yPts(8), xPts(12), yPts(12), Col
gLine xPts(8), yPts(8), xPts(9), yPts(9), Col
gLine xPts(9), yPts(9), xPts(10), yPts(10), Col
gLine xPts(10), yPts(10), xPts(11), yPts(11), Col
gLine xPts(11), yPts(11), xPts(12), yPts(12), Col
gLine xPts(12), yPts(12), xPts(13), yPts(13), Col
gLine xPts(13), yPts(13), xPts(2), yPts(2), Col
gLine xPts(2), yPts(2), xPts(14), yPts(14), Col
gLine xPts(13), yPts(13), xPts(14), yPts(14), Col
gLine xPts(14), yPts(14), xPts(15), yPts(15), Col
gLine xPts(15), yPts(15), xPts(1), yPts(1), Col
gLine xPts(15), yPts(15), xPts(3), yPts(3), Col

'turret
gCircle Player(i).X, Player(i).Y, 100, tCol
gLine b1x, b1y, G1X, G1Y, tCol
gLine b2x, b2y, G2X, G2Y, tCol


'extras
If Player(i).bDrawShields Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 2, vbGreen  ', , , -1
    Player(i).bDrawShields = False
End If

If (Player(i).State And Player_Shielding) = Player_Shielding Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 2.2, vbGreen
End If

If (Player(i).State And ePlayerState.PLAYER_THRUST) = ePlayerState.PLAYER_THRUST Then
    
    A1 = FixAngle(1.25 * Pi - Player(i).Facing) '1.25 *
    A2 = FixAngle(1.75 * Pi - Player(i).Facing)
    r = SHIP_RADIUS * 1.5
    bDrawAccel = True
    
ElseIf (Player(i).State And ePlayerState.PLAYER_REVTHRUST) = ePlayerState.PLAYER_REVTHRUST Then
    
    A1 = FixAngle(0.35 * Pi - Player(i).Facing)
    A2 = FixAngle(0.66 * Pi - Player(i).Facing)
    r = SHIP_Height * 1.2
    bDrawAccel = True

    'gCircle (x,y), radius, colour, start, end, aspect/curveness
End If

If bDrawAccel Then
    gCircleSE Player(i).X, Player(i).Y, r, vbRed, A1, A2
    'gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
End If

EH:
End Sub

Private Sub DrawInfiltrator(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

Dim xPts(1 To 17) As Single
Dim yPts(1 To 17) As Single
Const sfa = 1.8
'Dim g1x As Integer, g1y As Integer, b1x As Integer, b1y As Integer
'Dim g2x As Integer, g2y As Integer, b2x As Integer, b2y As Integer

Dim A1 As Single, A2 As Single
Dim bDrawAccel As Boolean
Dim r As Single, j As Single

Const D10 = 0.175 '10-ish degrees

If (Player(i).State And Player_Secondary) = Player_Secondary Then
    If Player(i).ID <> MyID Then Exit Sub 'don't show them
End If

On Error GoTo EH
xPts(1) = Player(i).X + SHIP_Height * 2.3 * Sine(Player(i).Facing) / sfa ' - D10 / 1.6) / sfa
yPts(1) = Player(i).Y - SHIP_Height * 2.3 * CoSine(Player(i).Facing) / sfa ' - D10 / 1.6) / sfa
xPts(2) = Player(i).X + SHIP_Height / 1.3 * Sine(Player(i).Facing - piD4) / sfa '45
yPts(2) = Player(i).Y - SHIP_Height / 1.3 * CoSine(Player(i).Facing - piD4) / sfa '45
xPts(3) = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Facing - piD3) / sfa
yPts(3) = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Facing - piD3) / sfa
xPts(4) = Player(i).X + SHIP_Height * 1.2 * Sine(Player(i).Facing - Pi / 3.5) / sfa
yPts(4) = Player(i).Y - SHIP_Height * 1.2 * CoSine(Player(i).Facing - Pi / 3.5) / sfa
xPts(5) = Player(i).X + SHIP_Height * 1.2 * Sine(Player(i).Facing - 2.2 * piD3) / sfa
yPts(5) = Player(i).Y - SHIP_Height * 1.2 * CoSine(Player(i).Facing - 2.2 * piD3) / sfa
xPts(6) = Player(i).X + SHIP_Height * 1.5 * Sine(Player(i).Facing - 3.3 * piD4) / sfa
yPts(6) = Player(i).Y - SHIP_Height * 1.5 * CoSine(Player(i).Facing - 3.3 * piD4) / sfa
xPts(7) = Player(i).X + SHIP_Height / 1.8 * Sine(Player(i).Facing - pi3D4) / sfa
yPts(7) = Player(i).Y - SHIP_Height / 1.8 * CoSine(Player(i).Facing - pi3D4) / sfa
xPts(8) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Facing - 3.7 * piD4) / sfa
yPts(8) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Facing - 3.7 * piD4) / sfa
xPts(9) = Player(i).X + SHIP_Height * Sine(Player(i).Facing - 3.8 * piD4) / sfa
yPts(9) = Player(i).Y - SHIP_Height * CoSine(Player(i).Facing - 3.8 * piD4) / sfa

'xPts(18) = Player(i).X + SHIP_Height * 2.3 * sine(Player(i).Facing) / sfa ' + D10 / 1.6) / sfa
'yPts(18) = Player(i).Y - SHIP_Height * 2.3 * cosine(Player(i).Facing) / sfa ' + D10 / 1.6) / sfa
xPts(17) = Player(i).X + SHIP_Height / 1.3 * Sine(Player(i).Facing + piD4) / sfa '45
yPts(17) = Player(i).Y - SHIP_Height / 1.3 * CoSine(Player(i).Facing + piD4) / sfa '45
xPts(16) = Player(i).X + SHIP_Height / 1.5 * Sine(Player(i).Facing + piD3) / sfa
yPts(16) = Player(i).Y - SHIP_Height / 1.5 * CoSine(Player(i).Facing + piD3) / sfa
xPts(15) = Player(i).X + SHIP_Height * 1.2 * Sine(Player(i).Facing + Pi / 3.5) / sfa
yPts(15) = Player(i).Y - SHIP_Height * 1.2 * CoSine(Player(i).Facing + Pi / 3.5) / sfa
xPts(14) = Player(i).X + SHIP_Height * 1.2 * Sine(Player(i).Facing + 2.2 * piD3) / sfa
yPts(14) = Player(i).Y - SHIP_Height * 1.2 * CoSine(Player(i).Facing + 2.2 * piD3) / sfa
xPts(13) = Player(i).X + SHIP_Height * 1.5 * Sine(Player(i).Facing + 3.3 * piD4) / sfa
yPts(13) = Player(i).Y - SHIP_Height * 1.5 * CoSine(Player(i).Facing + 3.3 * piD4) / sfa
xPts(12) = Player(i).X + SHIP_Height / 1.8 * Sine(Player(i).Facing + pi3D4) / sfa
yPts(12) = Player(i).Y - SHIP_Height / 1.8 * CoSine(Player(i).Facing + pi3D4) / sfa
xPts(11) = Player(i).X + SHIP_Height / 1.2 * Sine(Player(i).Facing + 3.7 * piD4) / sfa
yPts(11) = Player(i).Y - SHIP_Height / 1.2 * CoSine(Player(i).Facing + 3.7 * piD4) / sfa
xPts(10) = Player(i).X + SHIP_Height * Sine(Player(i).Facing + 3.8 * piD4) / sfa
yPts(10) = Player(i).Y - SHIP_Height * CoSine(Player(i).Facing + 3.8 * piD4) / sfa

picMain.FillColor = 0

'main bit

picMain.ForeColor = Col

For j = 1 To 16
    gLine xPts(j), yPts(j), xPts(j + 1), yPts(j + 1), Col
Next j

gLine xPts(17), yPts(17), xPts(1), yPts(1), Col

If Player(i).State And Player_Secondary Then
    picMain.DrawWidth = Thin * 5
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 1.8, MGrey
    picMain.DrawWidth = Thin
End If

'extras
'If Player(i).bDrawShields Then
    'gCircle (Player(i).x, Player(i).y), SHIP_Height * 1.2, vbGreen ', , , -1
    'Player(i).bDrawShields = False
'End If


If (Player(i).State And ePlayerState.PLAYER_THRUST) = ePlayerState.PLAYER_THRUST Then
    
    A1 = FixAngle(1.25 * Pi - Player(i).Facing) '1.25 *
    A2 = FixAngle(1.75 * Pi - Player(i).Facing)
    r = SHIP_RADIUS * 1.2
    bDrawAccel = True
    
ElseIf (Player(i).State And ePlayerState.PLAYER_REVTHRUST) = ePlayerState.PLAYER_REVTHRUST Then
    
    A1 = FixAngle(0.35 * Pi - Player(i).Facing)
    A2 = FixAngle(0.66 * Pi - Player(i).Facing)
    r = SHIP_Height * 1.2
    bDrawAccel = True

    'gCircle (x,y), radius, colour, start, end, aspect/curveness
End If

If bDrawAccel Then
    gCircleSE Player(i).X, Player(i).Y, r, vbRed, A1, A2
    'gCircle (Player(i).X, Player(i).Y), R, vbRed, A1, A2
End If

EH:
End Sub

Private Sub DrawSD(ByVal i As Integer, ByVal Col As Long, ByVal tCol As Long)

Dim xPts(1 To 17) As Single
Dim yPts(1 To 17) As Single

Dim A1 As Single, A2 As Single
Dim bDrawAccel As Boolean
Dim r As Integer, j As Integer

On Error GoTo EH
xPts(1) = Player(i).X + 2.5 * SHIP_Height * Sine(Player(i).Heading - piD10 / 2)
xPts(2) = Player(i).X + 2 * SHIP_Height * Sine(Player(i).Heading - pi3D4)
xPts(3) = Player(i).X + 2 * SHIP_Height * Sine(Player(i).Heading - Pi)
xPts(4) = Player(i).X + 2 * SHIP_Height * Sine(Player(i).Heading + pi3D4)
xPts(5) = Player(i).X + 2.5 * SHIP_Height * Sine(Player(i).Heading + piD10 / 2)
'-
xPts(6) = Player(i).X + 1.2 * SHIP_Height * Sine(Player(i).Heading - pi3D4)
xPts(7) = Player(i).X + 1.2 * SHIP_Height * Sine(Player(i).Heading + pi3D4)
xPts(8) = Player(i).X + 1.4 * SHIP_Height * Sine(Player(i).Heading - 5 * piD6)
xPts(9) = Player(i).X + 1.4 * SHIP_Height * Sine(Player(i).Heading + 5 * piD6)
xPts(10) = Player(i).X + 1.1 * SHIP_Height * Sine(Player(i).Heading - 5 * piD6)
xPts(11) = Player(i).X + 1.1 * SHIP_Height * Sine(Player(i).Heading + 5 * piD6)
'-
Call GetSDGunTurrets(xPts(13), xPts(16), xPts(17), yPts(13), yPts(16), yPts(17), i)
'If modSpaceGame.cl_UseMouse And Player(i).ID = MyID Then
'    xPts(12) = xPts(13) + BULLET_LEN * sine(FindAngle(xPts(13), yPts(13), MouseX, MouseY))
'    xPts(14) = xPts(16) + BULLET_LEN * sine(FindAngle(xPts(16), yPts(16), MouseX, MouseY))
'    xPts(15) = xPts(17) + BULLET_LEN * sine(FindAngle(xPts(17), yPts(17), MouseX, MouseY))
'Else
    xPts(12) = xPts(13) + BULLET_LEN * Sine(Player(i).Facing)
    xPts(14) = xPts(16) + BULLET_LEN * Sine(Player(i).Facing)
    xPts(15) = xPts(17) + BULLET_LEN * Sine(Player(i).Facing)
'End If

'---
yPts(1) = Player(i).Y - 2.5 * SHIP_Height * CoSine(Player(i).Heading - piD10 / 2)
yPts(2) = Player(i).Y - 2 * SHIP_Height * CoSine(Player(i).Heading - pi3D4)
yPts(3) = Player(i).Y - 2 * SHIP_Height * CoSine(Player(i).Heading - Pi)
yPts(4) = Player(i).Y - 2 * SHIP_Height * CoSine(Player(i).Heading + pi3D4)
yPts(5) = Player(i).Y - 2.5 * SHIP_Height * CoSine(Player(i).Heading + piD10 / 2)
'-
yPts(6) = Player(i).Y - 1.2 * SHIP_Height * CoSine(Player(i).Heading - pi3D4)
yPts(7) = Player(i).Y - 1.2 * SHIP_Height * CoSine(Player(i).Heading + pi3D4)
yPts(8) = Player(i).Y - 1.4 * SHIP_Height * CoSine(Player(i).Heading - 5 * piD6)
yPts(9) = Player(i).Y - 1.4 * SHIP_Height * CoSine(Player(i).Heading + 5 * piD6)
yPts(10) = Player(i).Y - 1.1 * SHIP_Height * CoSine(Player(i).Heading - 5 * piD6)
yPts(11) = Player(i).Y - 1.1 * SHIP_Height * CoSine(Player(i).Heading + 5 * piD6)
'-

'If modSpaceGame.cl_UseMouse And Player(i).ID = MyID Then
'    yPts(12) = yPts(13) - BULLET_LEN * cosine(FindAngle(xPts(13), yPts(13), MouseX, MouseY))
'    yPts(14) = yPts(16) - BULLET_LEN * cosine(FindAngle(xPts(16), yPts(16), MouseX, MouseY))
'    yPts(15) = yPts(17) - BULLET_LEN * cosine(FindAngle(xPts(17), yPts(17), MouseX, MouseY))
'Else
    yPts(12) = yPts(13) - BULLET_LEN * CoSine(Player(i).Facing)
    yPts(14) = yPts(16) - BULLET_LEN * CoSine(Player(i).Facing)
    yPts(15) = yPts(17) - BULLET_LEN * CoSine(Player(i).Facing)
'End If


picMain.FillColor = 0

'outline
gLine xPts(1), yPts(1), xPts(5), yPts(5), Col
For j = 1 To 4
    gLine xPts(j), yPts(j), xPts(j + 1), yPts(j + 1), Col
Next j


'bridge
gCircle xPts(10), yPts(10), 50, Col
gCircle xPts(11), yPts(11), 50, Col

gLine xPts(6), yPts(6), xPts(7), yPts(7), Col
gLine xPts(9), yPts(9), xPts(8), yPts(8), Col
'gline xPts(6), yPts(6),xPts(8), yPts(8)), Col
'gline xPts(9), yPts(9),xPts(7), yPts(7)), Col

'turrets
gCircle xPts(13), yPts(13), 40, Col
gCircle xPts(16), yPts(16), 40, Col
gCircle xPts(17), yPts(17), 40, Col
gLine xPts(13), yPts(13), xPts(12), yPts(12), Col
gLine xPts(16), yPts(16), xPts(14), yPts(14), Col
gLine xPts(17), yPts(17), xPts(15), yPts(15), Col

'extras
If (Player(i).State And Player_Secondary) = Player_Secondary Then
    picMain.DrawWidth = Thin * 5
    gCircle Player(i).X, Player(i).Y, SHIP_Height * SD_GravityRadius, MGrey
    picMain.DrawWidth = Thin
End If

'extras
If Player(i).bDrawShields Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 3.2, vbGreen ', , , -1
    Player(i).bDrawShields = False
End If

If (Player(i).State And Player_Shielding) = Player_Shielding Then
    gCircle Player(i).X, Player(i).Y, SHIP_Height * 3, vbGreen
End If

EH:
End Sub

Private Sub GetSDGunTurrets(ByRef X1 As Single, ByRef X2 As Single, ByRef x3 As Single, _
                            ByRef Y1 As Single, ByRef Y2 As Single, ByRef y3 As Single, _
                            ByRef i As Integer)

X1 = Player(i).X + 2 * SHIP_Height * Sine(Player(i).Heading)
X2 = Player(i).X + SHIP_Height * Sine(Player(i).Heading - piD4) / 1.5
x3 = Player(i).X + SHIP_Height * Sine(Player(i).Heading + piD4) / 1.5

Y1 = Player(i).Y - 2 * SHIP_Height * CoSine(Player(i).Heading)
Y2 = Player(i).Y - SHIP_Height * CoSine(Player(i).Heading - piD4) / 1.5
y3 = Player(i).Y - SHIP_Height * CoSine(Player(i).Heading + piD4) / 1.5

End Sub

Private Sub DrawMissiles()

Const WO2 = 1000 'Width \ 2 - 100
'Const h = 8200 '7900 'Height \ 2 - 100
Dim H As Single

H = Me.height - 800

Dim i As Integer
'Dim pX1 As Integer, pY1 As Integer
Dim pX2 As Single, pY2 As Single
Dim T As Long

i = 0
Do While i < NumMissiles
    If (Missiles(i).Decay + Missile_Decay) < GetTickCount() Then
        RemoveMissile i
        i = i - 1
    End If
    i = i + 1
Loop

'Me.ForeColor = vbWhite
For i = 0 To NumMissiles - 1
    
    picMain.FillStyle = 1
    'pX1 = CInt(Missiles(i).X + sine(Missiles(i).Heading) * Missile_LEN)
    'pY1 = CInt(Missiles(i).Y - cosine(Missiles(i).Heading) * Missile_LEN)
    'gline pX1, pY1,Missiles(i).X, Missiles(i).Y), Missiles(i).Colour
    
    picMain.DrawWidth = Thin * 2
    pX2 = CInt(Missiles(i).X + Sine(Missiles(i).Facing - Pi) * Missile_LEN * 2)
    pY2 = CInt(Missiles(i).Y - CoSine(Missiles(i).Facing - Pi) * Missile_LEN * 2)
    gLine pX2, pY2, Missiles(i).X, Missiles(i).Y, vbRed 'Missiles(i).Colour
    
    pX2 = CInt(Missiles(i).X + Sine(Missiles(i).Heading) * Missile_LEN * 1.5)
    pY2 = CInt(Missiles(i).Y - CoSine(Missiles(i).Heading) * Missile_LEN * 1.5)
    gLine pX2, pY2, Missiles(i).X, Missiles(i).Y, vbBlue 'Missiles(i).Colour
    
    picMain.DrawWidth = Thin
    
    picMain.FillColor = Missiles(i).Colour
    picMain.FillStyle = 0
    gCircle Missiles(i).X, Missiles(i).Y, Missile_Radius, Missiles(i).Colour
    picMain.FillStyle = 1
    
    picMain.ForeColor = Missiles(i).Colour
    PrintText "Hull: " & CStr(Missiles(i).Hull), Missiles(i).X + 100, Missiles(i).Y, vbWhite
    
    PrintText "Fuel: " & CStr(Round((Missiles(i).Decay - GetTickCount()) / 1000) + 3), Missiles(i).X + 100, Missiles(i).Y + 200, vbWhite
    
    'PrintText "Lock: " & CStr(Missiles(i).InRange), Missiles(i).X + 100, Missiles(i).Y + 400
    
    
Next i

If modSpaceGame.cg_Smoke Then
    For i = 0 To NumMissiles - 1
        If Missiles(i).LastSmoke + 20 / modSpaceGame.sv_GameSpeed < GetTickCount() Then
            
            'AddVectors Player(i).Speed, Player(i).Heading, 100, Player(i).Facing, rMag, rDir
            
            AddSmokeGroup Missiles(i).X, Missiles(i).Y, 1 ', Player(i).Speed, Player(i).Heading
            Missiles(i).LastSmoke = GetTickCount()
        End If
    Next i
End If


If PlayerInGame(0) Then 'Player(0).Team <> Spec Then
    'show if player(0).canfiremissile
    
    If Player(0).ShipType <> MotherShip Then
        If Player(0).ShipType <> Wraith Then
            If Player(0).ShipType <> Infiltrator Then
                If Player(0).ShipType <> SD Then
                    
                    T = Player(0).LastSecondary + Missile_Delay / modSpaceGame.sv_GameSpeed
                    
                    'Me.ForeColor = MGrey
                    picMain.Font.Size = BigFontSize
                    
                    If T > GetTickCount() Then
                        PrintFormText "Next Missile: " & CStr(Round((T - GetTickCount()) / 1000)), WO2, H, MGrey
                    Else
                        PrintFormText "Missile Ready", WO2, H, MGrey
                    End If
                    
                    picMain.Font.Size = NormalFontSize
                    
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub AddBullet(X As Single, Y As Single, Speed As Single, _
    Heading As Single, OwnerID As Integer, Col As Long, Damage As Single, _
    iPlayer As Integer, Optional ByVal UseFacing_Flash As Boolean = True)

Dim nX As Single, nY As Single

'Create a new bullet, and add the given attributes
ReDim Preserve Bullet(NumBullets)
Bullet(NumBullets).Decay = GetTickCount() + (Bullet_Decay + Rnd() * Bullet_Decay_Extra) / modSpaceGame.sv_GameSpeed
Bullet(NumBullets).Heading = Heading
Bullet(NumBullets).Speed = IIf(Speed >= BULLET_SPEED * 1.2, BULLET_SPEED * 2, Speed)
Bullet(NumBullets).X = X
Bullet(NumBullets).Y = Y
Bullet(NumBullets).Owner = OwnerID
Bullet(NumBullets).Colour = Col
Bullet(NumBullets).Damage = Damage

picMain.DrawWidth = Thin * 2

If UseFacing_Flash Then
    nX = Bullet(NumBullets).X + BULLET_LEN * Sine(Player(iPlayer).Facing)
    nY = Bullet(NumBullets).Y - BULLET_LEN * CoSine(Player(iPlayer).Facing)
Else
    nX = Bullet(NumBullets).X
    nY = Bullet(NumBullets).Y
End If

picMain.FillStyle = 0
gCircle nX, nY, Bullet_Radius * 2, vbYellow
picMain.FillStyle = 1

'If Player(iPlayer).LastGunSmoke + 20 / modSpaceGame.sv_GameSpeed < GetTickCount() Then
If modSpaceGame.cg_GunSmoke Then
    AddSmokeGroup nX, nY, 2
End If
'End If

'- modSpaceGame.sv_GameSpeed * Bullet(NumBullets).Speed * cosine(Bullet(NumBullets).Heading))

picMain.DrawWidth = Thin

NumBullets = NumBullets + 1

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
        RemoveBullet i ', False
        'Decrement the counter
        i = i - 1
    End If
    'Increment the counter
    i = i + 1
Loop

'Step through each bullet and draw it
picMain.ForeColor = vbWhite
For i = 0 To NumBullets - 1
    'Draw the bullet
    'gCircle Bullet(i).x - (BULLET_RADIUS + 0.5), Bullet(i).y - (BULLET_RADIUS + 0.5), _
        Bullet(i).x + BULLET_RADIUS + 0.5, Bullet(i).y + BULLET_RADIUS + 0.5, Me.hdc
    
    picMain.DrawWidth = IIf(modSpaceGame.cg_DrawThick, Thin * 2, Thin) + Bullet(i).Damage / 100
    
    picMain.FillStyle = 1
    pX = CInt(Bullet(i).X - Sine(Bullet(i).Heading) * BULLET_LEN)
    pY = CInt(Bullet(i).Y + CoSine(Bullet(i).Heading) * BULLET_LEN)
    
    'SideX1 = .X + .Speed * sine(D2R(.Facing - 90))
    'SideY1 = .Y + .Speed * cosine(D2R(.Facing - 90))
    
    'SideX2 = .X + .Speed * sine(D2R(.Facing + 90))
    'SideY2 = .Y + .Speed * cosine(D2R(.Facing + 90))
    
    'gline SideX1, SideY1,pX, pY), vbBlack
    'gline SideX2, SideY2,pX, pY), vbBlack
    
    gLine Bullet(i).X, Bullet(i).Y, pX, pY, vbGreen
    
    picMain.FillColor = Bullet(i).Colour
    picMain.FillStyle = 0
    gCircle Bullet(i).X, Bullet(i).Y, Bullet_Radius, Bullet(i).Colour
    picMain.FillStyle = 1
Next i

End Sub

Private Sub RemoveBullet(Index As Integer, Optional ByVal Flsh As Boolean = False, _
    Optional PSpeed As Single, Optional pHeading As Single)

Dim i As Integer

If Flsh Then
    'picMain.DrawWidth = Thin * 10
    'gCircle (Bullet(Index).x, Bullet(Index).y), Bullet_Radius * 2, vbRed
    'picMain.DrawWidth = Thin
    AddExplosion Bullet(Index).X, Bullet(Index).Y, 250, 0.5, PSpeed, pHeading 'Bullet(Index).Speed, Bullet(Index).Heading
'Else
    'gCircle (Bullet(Index).x, Bullet(Index).y), Bullet_Radius * 2, vbYellow
End If

If modSpaceGame.cg_BulletSmoke Then AddSmokeGroup Bullet(Index).X, Bullet(Index).Y, 3

'If there's only one bullet left, just erase the array
If NumBullets = 1 Then
    Erase Bullet
    NumBullets = 0
Else
    'Remove the bullet
    For i = Index To NumBullets - 2
        Bullet(i).Decay = Bullet(i + 1).Decay
        Bullet(i).Heading = Bullet(i + 1).Heading
        Bullet(i).Speed = Bullet(i + 1).Speed
        Bullet(i).X = Bullet(i + 1).X
        Bullet(i).Y = Bullet(i + 1).Y
        Bullet(i).Owner = Bullet(i + 1).Owner
        Bullet(i).Colour = Bullet(i + 1).Colour
        Bullet(i).Damage = Bullet(i + 1).Damage
        Bullet(i).LastDeflect = Bullet(i + 1).LastDeflect
    Next i
    
    'Resize the array
    ReDim Preserve Bullet(NumBullets - 2)
    NumBullets = NumBullets - 1
End If

End Sub

Private Sub RemoveMissile(Index As Integer)

Dim i As Integer
Const Spread = 100

'picMain.DrawWidth = Thick
'gCircle (Missiles(Index).X, Missiles(Index).Y), Missile_Radius * 3, vbRed
'picMain.DrawWidth = Thin
AddExplosion Missiles(Index).X, Missiles(Index).Y, 100, 1, Missiles(Index).Speed, Missiles(Index).Heading

If modSpaceGame.cg_BulletSmoke Then
    AddSmokeGroup Missiles(Index).X, Missiles(Index).Y, 10
    AddSmokeGroup Missiles(Index).X + Spread, Missiles(Index).Y + Spread, 10
    AddSmokeGroup Missiles(Index).X - Spread, Missiles(Index).Y + Spread, 10
    AddSmokeGroup Missiles(Index).X + Spread, Missiles(Index).Y - Spread, 10
    AddSmokeGroup Missiles(Index).X - Spread, Missiles(Index).Y - Spread, 10
End If

'If there's only one bullet left, just erase the array
If NumMissiles = 1 Then
    Erase Missiles
    NumMissiles = 0
Else
    'Remove the bullet
    For i = Index To NumMissiles - 2
        Missiles(i).Decay = Missiles(i + 1).Decay
        Missiles(i).Heading = Missiles(i + 1).Heading
        Missiles(i).Speed = Missiles(i + 1).Speed
        Missiles(i).X = Missiles(i + 1).X
        Missiles(i).Y = Missiles(i + 1).Y
        Missiles(i).Owner = Missiles(i + 1).Owner
        Missiles(i).Colour = Missiles(i + 1).Colour
        Missiles(i).TargetID = Missiles(i + 1).TargetID
        Missiles(i).Hull = Missiles(i + 1).Hull
        Missiles(i).LastSmoke = Missiles(i + 1).LastSmoke
    Next i
    
    'Resize the array
    ReDim Preserve Missiles(NumMissiles - 2)
    NumMissiles = NumMissiles - 1
End If

End Sub

Private Function StartWinsock() As Boolean

AddConsoleText "Initialising Game Winsock", , True, , True

''Init winsock
'If modWinsock.InitWinsock() = WINSOCK_ERROR Then
'    'Handle error..
'    GoTo EH
'End If

modWinsock.DestroySocket socket

AddConsoleText "Creating Socket..."

'Make a socket
socket = modWinsock.CreateSocket()
If socket = WINSOCK_ERROR Then
    'Handle error
    'modWinsock.TermWinsock
    GoTo EH
Else
    AddConsoleText "Created Game Socket: " & CStr(socket)
End If

'If we're the server, bind to the server port
If modSpaceGame.SpaceServer Then
    AddConsoleText "Binding Socket to " & CStr(modPorts.SpacePort) & "..."
    If modWinsock.BindSocket(socket, modPorts.SpacePort) = WINSOCK_ERROR Then
        'Handle error
        GoTo EH
    End If
End If

AddConsoleText "Initialised Game Winsock", , , True

StartWinsock = True

Exit Function
EH:
AddConsoleText "Error Starting Winsock", , , True
AddText "Error Starting Winsock", TxtError, True
Call EndWinsock
Unload Me
End Function

Private Function ConnectToServer() As Boolean

Dim JoinTimer As Long
'Dim TimeOutTimer As Long
Dim sPacket As String
Dim TempSockAddr As ptSockAddr
Dim CurrentRetry As Integer
Dim Txt As String

Const kX = CentreX - 900
Const kY = CentreY + 600
Dim LastLine As Long
Const LineDelay = 20


picMain.ForeColor = Player(0).Colour
'Me.CurrentX = 7
'Me.CurrentY = 7
'Print "Connecting to Server..."


'Make the server's ptsockaddr
If MakeSockAddr(ServerSockAddr, modPorts.SpacePort, modSpaceGame.SpaceServerIP) = WINSOCK_ERROR Then
    'Handle error
    AddText "Error - IP isn't valid", TxtError, True 'Making Socket", TxtError, True
    Unload Me
    
Else
    
    'Send "Join" packets to the server until we receive an "ACK" mPacket
    CurrentRetry = 1
    
    Do 'While TimeOutTimer + SERVER_CONNECT_DURATION > GetTickCount()
        
        'Is it time to send a "Join" mPacket?
        If (JoinTimer + SERVER_RETRY_FREQ) < GetTickCount() Then
            'Reset the timer
            JoinTimer = GetTickCount()
            
            'Send the mPacket
            modWinsock.SendPacket socket, ServerSockAddr, sJoins
            
            If CurrentRetry < 6 Then
                Me.picMain.Cls
                
                Txt = "Connecting to Server '" & modSpaceGame.SpaceServerIP & "'..."
                modSpaceGame.PrintFormText Txt, CentreX - TextWidth(Txt) / 2, CentreY - TextHeight(Txt), Player(0).Colour
                
                Txt = "Waiting For Response... " & CStr(CurrentRetry)
                modSpaceGame.PrintFormText Txt, CentreX - TextWidth(Txt) / 2, CentreY + TextHeight(Txt), Player(0).Colour
                
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
                Player(0).ID = MyID
                
                Txt = "Response Received"
                modSpaceGame.PrintFormText Txt, CentreX - TextWidth(Txt) / 2, CentreY + 3 * TextHeight(Txt), Player(0).Colour
                
                Txt = "Setting Up Game..."
                modSpaceGame.PrintFormText Txt, CentreX - TextWidth(Txt) / 2, CentreY + 4 * TextHeight(Txt), Player(0).Colour
                
                Pause 500 * Rnd()
                                
                'Start playing!
                ConnectToServer = True
                Exit Function
                
            End If
            
        End If
        
        
        If LastLine + LineDelay < GetTickCount() Then
            If modVars.Closing Or ClosingWindow Then Exit Function
            
            LastLine = GetTickCount()
            Me.picMain.Line (kX, kY)-(kX + LastLine - JoinTimer, kY), vbRed
            
            Call BltToForm
            
            Me.Refresh
            
        End If
        
        
    Loop Until ((CurrentRetry - 1) > SERVER_NUM_RETRIES) Or modVars.Closing Or ClosingWindow
    
    
    ConnectToServer = False
    
    If modVars.Closing Or ClosingWindow Then Exit Function
    
    'We didn't receive an ACK before the timeout
    AddText "Unable to Connect to Server - No Packet Flow", TxtError, True
    
End If

End Function

Private Sub ScanForKills(sTxt As String)
Dim j As Integer, i As Integer
Dim l As Integer
Dim KillType As Byte

Dim Name1 As String, Name2 As String, Tmp As String, sKillType As String

Dim DeadPlayeri As Integer, LivePlayeri As Integer 'for ctf

DeadPlayeri = -1
LivePlayeri = -1


If InStr(1, sTxt, modMessaging.MsgNameSeparator, vbTextCompare) Then Exit Sub 'it is a chat message


j = InStr(1, sTxt, ShotBy, vbTextCompare)

If j = 0 Then
    j = InStr(1, sTxt, MissiledBy, vbTextCompare)
    
    If j = 0 Then
        j = InStr(1, sTxt, RammedBy, vbTextCompare)
        l = Len(RammedBy)
        KillType = 3
    Else
        l = Len(MissiledBy)
        KillType = 2
    End If
Else
    l = Len(ShotBy)
    KillType = 1
End If

If j Then
    
    
    Name1 = Mid$(sTxt, j + l) 'killer
    Name2 = Left$(sTxt, j - 1) 'killed
    
    
    For i = 0 To NumPlayers - 1
        If Trim$(Player(i).Name) = Name1 Then
            Player(i).Kills = Player(i).Kills + 1
            
            LivePlayeri = i
            
            If (Player(i).State And Player_Secondary) = Player_Secondary Then
'                If Player(i).ShipType <> SD Then
'                    If Player(i).ShipType <> Infiltrator Then
'                        If Player(i).ShipType <> MotherShip Then
'                            if player(i).ShipType <> Wraith
                
                If Player(i).ShipType = Raptor Or _
                        Player(i).ShipType = Behemoth Or _
                        Player(i).ShipType = Hornet Then
                    
                    SubPlayerState i, Player_Secondary
                End If
                
            End If
            
            If Player(i).ID = MyID Then
                KillsInARow = KillsInARow + 1
                
                Select Case KillType
                    Case 1
                        sKillType = "Shot "
                    Case 2
                        sKillType = "Missiled "
                    Case 3
                        sKillType = "Rammed "
                End Select
                
                AddMainMessage "You " & sKillType & Name2
                
                
                '###########################################
                'check if we should show a grey message
                If KillsInARow > 1 Then

                    If KillsInARow = 2 Then
                        Tmp = "Double Kill!"
                    ElseIf KillsInARow = 3 Then
                        Tmp = "Triple Kill!"
                    ElseIf KillsInARow > 3 Then
                        Tmp = "Uber Kill! (" & CStr(KillsInARow) & ")"
                    End If

                    AddMainMessage Tmp

                End If
                '###########################################
                
                'modSpaceGame.Kills = modSpaceGame.Kills + 1
                
                'LastMainMessage = GetTickCount() + MainMessageDelay
                
                
                'since we have killed, Health Bonus
                If Player(i).ShipType <> MotherShip Then
                    Player(i).Shields = Player(i).MaxShields
                    Player(i).Hull = Player(i).MaxHull
                    Player(i).LastSecondary = 0
                End If
                
                'increase the score with said ship
                On Error Resume Next
                ShipScores(Player(i).ShipType) = ShipScores(Player(i).ShipType) + 1
                
                'can we use it now?
                MotherShipAvail = CheckMotherShip()
                WraithAvail = CheckWraith()
                InfilAvail = CheckInfil()
                SDAvail = CheckSD()
                
            Else
                KillsInARow = 0
            End If
            
        ElseIf Trim$(Player(i).Name) = Name2 Then
            
            'If Player(i).ID <> MyID Then
            Player(i).Deaths = Player(i).Deaths + 1
            'End If
            
            DeadPlayeri = i
            
            
            If Player(i).ID = MyID Then
                'modSpaceGame.Deaths = modSpaceGame.Deaths + 1
                
                Select Case KillType
                    Case 1
                        sKillType = "Shot "
                    Case 2
                        sKillType = "Missiled "
                    Case 3
                        sKillType = "Rammed "
                End Select
                
                AddMainMessage sKillType & "by " & Name1
                
                If modSpaceGame.sv_GameType = Elimination Then
                    AddMainMessage "You will respawn next round"
                    'Player(i).Alive = False
                End If
                
            End If
            
        End If
    Next i
'ElseIf InStr(1, sTxt, "bot removed", vbTextCompare) Then
End If

'If modSpaceGame.SpaceServer Then
    
    If modSpaceGame.sv_GameType = CTF Then
        
        If DeadPlayeri <> -1 Then
            If Player(DeadPlayeri).ID = FlagOwnerID Then
            
                If LivePlayeri <> -1 Then
                    FlagOwnerID = Player(LivePlayeri).ID
                    
                    'SendChatPacketBroadcast "Flag Stolen by: " & Trim$(Player(LivePlayeri).Name), Player(LivePlayeri).Colour
                    
                End If
                
            End If
        End If
        
    ElseIf modSpaceGame.sv_GameType = Elimination Then
        
        If DeadPlayeri <> -1 Then
            'Player(DeadPlayeri).Team = Spec
            Player(DeadPlayeri).Alive = False
            
            SetPlayerState Player(DeadPlayeri).ID, Player_None
            
        End If
        
    End If
    
'End If

If DeadPlayeri <> -1 Then
    CalculateScore DeadPlayeri
End If
If LivePlayeri <> -1 Then
    CalculateScore LivePlayeri
End If


End Sub

Private Function GetPacket() As Boolean

Dim sPacket As String
Dim TempSockAddr As ptSockAddr
Dim i As Integer, j As Integer
Dim Tmp As String, sTxt As String

Dim l As Long

'box pos
Const Sep1 As String = "@"
Const Sep2 As String = "#"


'Loop until there are no packets
GetPacket = True

Do
    'Check for packets
    sPacket = modWinsock.ReceivePacket(socket, TempSockAddr)
    'Was there anything?
    If LenB(sPacket) = 0 Then
        GetPacket = True
    Else
        'Check what type of mPacket this is and take appropriate action
        Select Case Left$(sPacket, 1)
            Case sUpdates
                'A position update mPacket
                ProcessUpdatePacket sPacket
            Case sJoins
                'A join mPacket.  If we're a server, handle it
                If modSpaceGame.SpaceServer Then ProcessJoinPacket TempSockAddr
                
            Case sChats
                'A chat packet... if we're the server, broadcast
                
                On Error GoTo EH
                Tmp = Right$(sPacket, Len(sPacket) - InStrRev(sPacket, "#", , vbTextCompare))
                sTxt = Mid$(sPacket, 2, InStr(1, sPacket, "#", vbTextCompare) - 2)
                
                If modSpaceGame.SpaceServer Then
                    SendChatPacketBroadcast sTxt, CLng(Tmp)
                    '(auto-added to array)
                Else 'Otherwise, add it to the array
                    AddChatText sTxt, CLng(Tmp)
                End If
                
            Case sShipTypes
                
                Call ReceiveShipTypes(Mid$(sPacket, 2))
                
            Case sTeams
                
                Call ReceiveTeam(Mid$(sPacket, 2))
                
            Case sScoreUpdates
                
                Call ReceiveScoreUpdate(Mid$(sPacket, 2))
                
            Case sAsteroidUpdates
                
                If Not modSpaceGame.SpaceServer Then
                    Call ReceiveAsteroid(Mid$(sPacket, 2))
                End If
                
            Case sHasFlags
                
                On Error Resume Next
                FlagOwnerID = CInt(Mid$(sPacket, 2))
                
            Case sGameTypes
                
                On Error Resume Next
                i = CInt(Mid$(sPacket, 2))
                
                If modSpaceGame.sv_GameType <> i Then
                    modSpaceGame.sv_GameType = i
                    
                    AddMainMessage "Game Type - " & GetGameType()
                End If
                
'                Select Case modSpaceGame.sv_GameType
'                    Case eGameTypes.DM
'                        AddMainMessage "Game Type - Deathmatch"
'                    Case eGameTypes.CTF
'                        AddMainMessage "Game Type - CTF"
'                End Select
                
            Case sQuits
                
                On Error Resume Next
                j = CInt(Mid$(sPacket, 2))
                RemovePlayer FindPlayer(j)
                On Error GoTo 0
                
                If modSpaceGame.SpaceServer Then
                    SendBroadcast sPacket, j
                    'tell everyone to remove them, apart from said person
                End If
                
            Case sServerQuits
                AddText "Server Left the Game", TxtError, True
                bRunning = False
                GetPacket = False
                'Unload Me
                Exit Function
                
                
            Case sServerVarsUpdates
                
                If modSpaceGame.SpaceServer = False Then
                    On Error GoTo EH
                    
                    i = InStr(1, sPacket, "#")
                    
                    modSpaceGame.sv_BulletsCollide = CBool(Mid$(sPacket, 2, 1))
                    modSpaceGame.sv_AddBulletVectorToShip = CBool(Mid$(sPacket, 3, 1))
                    modSpaceGame.sv_ClipMissiles = CBool(Mid$(sPacket, 4, 1))
                    modSpaceGame.sv_BulletWallBounce = CBool(Mid$(sPacket, 5, 1))
                    modSpaceGame.sv_Bullet_Damage = CSng(Mid$(sPacket, 6, i - 6))
                    
                    j = InStr(1, sPacket, "@")
                    modSpaceGame.sv_CTFTime = CInt(Mid$(sPacket, i + 1, j - i - 1))
                    
                    modSpaceGame.sv_ScoreReq = CInt(Mid$(sPacket, j + 1))
                    
                End If
                
                
            Case sRemovePlayers
                
                If modSpaceGame.SpaceServer = False Then
                    On Error Resume Next
                    RemovePlayer FindPlayer(CInt(Mid$(sPacket, 2)))
                End If
                
'            Case sShipTypes
'                'sShipTypes & Player(i).ShipType & MyID
'                'i = shiptype
'                'j = playerID
'                On Error GoTo EH
'                i = Mid$(sPacket, 2, 1)
'                j = Mid$(sPacket, 3)
'
'                Player(FindPlayer(j)).ShipType = i
'
'                If modSpaceGame.SpaceServer Then
'                    SendBroadcast sPacket, j
'                End If
            Case sAntiLagPackets
                On Error Resume Next
                i = FindPlayer(CInt(Mid$(sPacket, 2)))
                
                If Player(i).ID <> MyID Then
                    Player(i).LastPacket = GetTickCount()
                End If
                
                'AddConsoleText "Received AntiLag - Index: " & CStr(i) & " LastPacket: " & Player(i).LastPacket & " GTC: " & GetTickCount()
                
                
                
            'Case sAddBots
                'Call AddBot
                
            Case sBoxPoss
                
                On Error Resume Next
                'vars = i, j ,stxt, tmp
                'sTxt =
                
                Call ReceiveBoxPos(Mid$(sPacket, 2))
                
'                i = InStr(1, sTxt, Sep1)
'                j = InStr(i + 1, sTxt, Sep1)
'
'                'On Error GoTo 0
'
'                ob1.Left = CSng(Left$(sTxt, i - 1))
'                ob1.Top = CSng(Mid$(sTxt, i + 1, j - i - 1))
'
'                i = InStr(1, sTxt, Sep2)
'
'                ln1.Left = CSng(Mid$(sTxt, j + 1, i - j - 1))
'                ln1.Top = CSng(Mid$(sTxt, i + 1))
                
                'ob1.l & Sep1 & ob1.t & Sep1 & ln1.l & Sep2 & ln1.t
                
            Case sGameSpeeds
                
                If modSpaceGame.SpaceServer = False Then
                    On Error Resume Next
                    
                    modSpaceGame.sv_GameSpeed = CSng(Mid$(sPacket, 2))
                    
                    If modSpaceGame.GameOptionFormLoaded Then
                        frmGameOptions.sldrSpeed.Value = modSpaceGame.sv_GameSpeed * 10
                    End If
                    
                End If
                
                
            Case sPowerUps
                
                'PowerUp.X & "|" & PowerUp.Y
                
                On Error Resume Next
                
                sTxt = Mid$(sPacket, 2)
                
                SpawnPowerUp CSng(Left$(sTxt, InStr(1, sTxt, "|") - 1)), _
                            CSng(Mid$(sTxt, 1 + InStr(1, sTxt, "|")))
                
            Case sKicks
                If modSpaceGame.SpaceServer = False Then
                    AddText "Disconnected - Was Kicked" & IIf(LenB(Mid$(sPacket, 2)) > 0, _
                        " (" & Mid$(sPacket, 2) & ")", vbNullString), TxtError, True
                    bRunning = False
                    GetPacket = False
                    Unload Me
                    Exit Function
                End If
                
                
            Case sEndRounds
                
                If modSpaceGame.SpaceServer = False Then
                    On Error Resume Next
                    RoundWinnerID = CInt(Mid$(sPacket, 2))
                    
                    For i = 0 To NumBullets - 1
                        RemoveBullet 0, True, 0, 0
                    Next i
                    
                    StopPlay True
                End If
                
            Case sNewRounds
                StopPlay False
                'RandomizePlayer '- done above
                
                
            Case sForceTeams
                ActivateTeam CInt(Mid$(sPacket, 2))
                
'            Case sScores
'                On Error Resume Next
'
'                Call ReceiveScore(Mid$(sPacket, 2))
'
'    '            i = Mid$(sPacket, 2)
'    '
'    '            j = FindPlayer(i)
'    '            Player(j).Score = Player(j).Score + 1
'    '
'    '            If j = MyID Then
'    '                MyScore = Player(j).Score
'    '                Player(j).Shields = Player(j).MaxShields 'add to shields as a bonus
'    '            End If
'    '            On Error GoTo 0
'    '
'    '            If Server Then
'    '                SendBroadcast sScores & CStr(i)
'    '            End If
'
'            Case sTellScores
'                If Server Then
'
'                    i = Mid$(sPacket, 2) 'id of player
'                    j = FindPlayer(i)
'
'
'                    On Error Resume Next
'                    Player(j).Score = Player(j).Score + 1
'                    Me.Caption = "Player(" & j & ")'s score = " & Player(j).Score
'
'                End If
                
        End Select
    End If
Loop Until LenB(sPacket) = 0

'Exit Function
EH:
'MsgBox "Error - " & Err.Description & vbNewLine & _
    "tmp = " & Tmp & vbNewLine & _
    "stxt = " & sTxt
    
End Function

Private Function CheckSD() As Boolean
Dim bCan As Boolean
Dim i As Integer

bCan = True

For i = 0 To eShipTypes.Hornet
    'If i <> eShipTypes.MotherShip Then
    If ShipScores(i) < KillsForSD Then
        bCan = False
    End If
    'End If
Next i

If bCan Then
    If modSpaceGame.GameOptionFormLoaded Then
        frmGameOptions.optnShipType(eShipTypes.SD).Enabled = True
    End If
End If

CheckSD = bCan

End Function

Private Sub SendAntiLagPacket()

Static LastSend As Long
Dim i As Integer

If (LastSend + AntiLagPacketDelay) < GetTickCount() Then
    
    LastSend = GetTickCount()
    
    If modSpaceGame.SpaceServer Then
        
        For i = 0 To NumPlayers - 1
            SendBroadcast sAntiLagPackets & CStr(Player(i).ID)
            
            If Player(i).IsBot Then
                Player(i).LastPacket = LastSend
            End If
            
        Next i
        
    Else
        modWinsock.SendPacket socket, ServerSockAddr, sAntiLagPackets & CStr(MyID)
    End If
    
    
    Player(0).LastPacket = LastSend
End If


End Sub

Private Function CheckMotherShip() As Boolean
Dim bCan As Boolean
Dim i As Integer

bCan = True

For i = 0 To eShipTypes.Hornet
    'If i <> eShipTypes.MotherShip Then
    If ShipScores(i) < KillsForMS Then
        bCan = False
    End If
    'End If
Next i

If bCan Then
    If modSpaceGame.GameOptionFormLoaded Then
        frmGameOptions.optnShipType(eShipTypes.MotherShip).Enabled = True
    End If
End If

CheckMotherShip = bCan

End Function

Private Function CheckWraith() As Boolean
Dim bCan As Boolean

If Player(0).Kills >= KillsForWraith Then
    bCan = True
Else
    bCan = False
End If

If bCan Then
    If modSpaceGame.GameOptionFormLoaded Then
        frmGameOptions.optnShipType(eShipTypes.Wraith).Enabled = True
    End If
End If

CheckWraith = bCan

End Function

Private Function CheckInfil() As Boolean
Dim bCan As Boolean

If Player(0).Kills >= KillsForInfil Then
    bCan = True
Else
    bCan = False
End If

If bCan Then
    If modSpaceGame.GameOptionFormLoaded Then
        frmGameOptions.optnShipType(eShipTypes.Infiltrator).Enabled = True
    End If
End If

CheckInfil = bCan

End Function

Public Function AddBot(ByVal Easy As Boolean, ByVal ShipT As eShipTypes, _
    ByVal Col As Long, ByVal vTeam As eTeams) As Integer

Dim i As Integer, MaxID As Integer, j As Integer
Dim BotNum As Integer
Dim Tmp As String

If modSpaceGame.SpaceServer Then
    'Make a spot
    i = AddPlayer()
    
    'Find a new ID
    'MaxID = 0
    
    For j = 0 To NumPlayers - 1
        'Is this ID greater?
        If Player(j).ID > MaxID Then MaxID = Player(j).ID
    Next j
    
    
    'Assign the ID
    Player(i).ID = MaxID + 1
    
    BotNum = 1
    
    For j = NumPlayers - 1 To 0 Step -1
        If Player(j).IsBot Then
            
            Tmp = Trim$(Player(j).Name)
            
            On Error Resume Next
            BotNum = CInt(Mid$(Tmp, InStr(1, Tmp, "Bot", vbTextCompare) + 3)) + 1
            
            Exit For
        End If
    Next j
    
    Player(i).Name = IIf(Easy, "Easy ", vbNullString) & "Bot" & CStr(BotNum) 'Player(i).ID)
    
    '-------------------
    
    Player(i).X = MaxWidth * Rnd()
    Player(i).Y = (MaxHeight - 500) * Rnd()
    Player(i).Facing = Pi2 * Rnd()
    
    'Set shields
    Player(i).Shields = Round(SHIELD_START / IIf(Easy, 2, 1))  ' --> easy to kill
    Player(i).MaxShields = Player(i).Shields
    
    Player(i).Hull = Round(Hull_Start / IIf(Easy, 2, 1))
    Player(i).MaxHull = Player(i).Hull
    
    Player(i).Colour = Col 'vbYellow 'frmGame.RandomRGBColour()
    
    Player(i).ShipType = ShipT
    Player(i).Team = vTeam
    
    Player(i).IsBot = True
    
    
    'ReDim Preserve BotIDs(NumBotIDs)
    'BotIDs(NumBotIDs) = Player(i).ID
    'NumBotIDs = NumBotIDs + 1
    '-------------------
    
    SendChatPacketBroadcast "Bot Added: " & Trim$(Player(i).Name), Player(i).Colour
End If

AddBot = i

End Function

Private Sub ProcessJoinPacket(ptSockAddr As ptSockAddr)

Dim i As Long
Dim ID As String
Dim Index As Integer
Dim MaxID As Integer

'If this IP address is already in our players array, use pre-assigned ID
For i = 0 To NumPlayers - 1
    'Is it the same IP and port?
    If (Player(i).ptSockAddr.sin_addr = ptSockAddr.sin_addr) And _
                                        (Player(i).ptSockAddr.sin_port = ptSockAddr.sin_port) Then
        
        ID = CInt(Player(i).ID)
        Exit For
    End If
Next i

'New player?
If Len(ID) = 0 And (ptSockAddr.sin_addr <> 0) Then
    'Make a spot
    Index = AddPlayer()
    'Find a new ID
    MaxID = 0
    For i = 0 To NumPlayers - 1
        'Is this ID greater?
        If Player(i).ID > MaxID Then MaxID = Player(i).ID
    Next i
    'Assign the ID
    Player(Index).ID = MaxID + 1
    'Set the player's ptsockaddr
    Player(Index).ptSockAddr.sin_addr = ptSockAddr.sin_addr
    Player(Index).ptSockAddr.sin_family = ptSockAddr.sin_family
    Player(Index).ptSockAddr.sin_port = ptSockAddr.sin_port
    Player(Index).ptSockAddr.sin_zero = ptSockAddr.sin_zero
    'Set the ID String
    ID = CStr(MaxID + 1)
End If

'Send the ACK
If (ptSockAddr.sin_addr <> 0) Then
    modWinsock.SendPacket socket, ptSockAddr, sAccepts & ID
    
'    Pause 5
'
'
'    For i = 0 To NumPlayers - 1
'        modWinsock.SendPacket socket, ptsockaddr, sShipTypes & _
'            Player(i).ShipType & Player(i).ID
'
'        Pause 5
'    Next i
    
End If

End Sub

Private Sub SendChatPacket(ChatText As String, Colour As Long)

'Is this the server?
If modSpaceGame.SpaceServer Then
    'Broadcast the chat mPacket
    SendChatPacketBroadcast ChatText, Colour
Else
    'Send it to the server
    modWinsock.SendPacket socket, ServerSockAddr, sChats & ChatText & "#" & CStr(Colour)
    'server will send it back here
End If

End Sub

Public Sub SendChatPacketBroadcast(ChatText As String, Colour As Long)

Dim i As Long

'Send the mPacket to everyone but the local user
For i = 0 To NumPlayers - 1
    'Is this the local user?
    If Player(i).ID <> MyID Then
        'Send!
        modWinsock.SendPacket socket, Player(i).ptSockAddr, sChats & ChatText & "#" & CStr(Colour)
    End If
Next i

'Add text to local user's chat text array
AddChatText ChatText, Colour

End Sub

Private Sub SendBroadcast(Text As String, Optional ByVal NtID As Integer = -1)

Dim i As Long

'Send the mPacket to everyone but the local user
For i = 0 To NumPlayers - 1
    'Is this the local user?
    If Player(i).ID <> MyID Then
        If Player(i).ID <> NtID Then
            If Player(i).IsBot = False Then
                'Send!
                modWinsock.SendPacket socket, Player(i).ptSockAddr, Text
            End If
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
NumChat = NumChat + 1

Call ScanForKills(ChatText)

End Sub


Private Sub RemoveChatText(Index As Integer)

Dim i As Long

'Remove the specified chat text
For i = Index To NumChat - 2
    Chat(i).Decay = Chat(i + 1).Decay
    Chat(i).Text = Chat(i + 1).Text
    Chat(i).Colour = Chat(i + 1).Colour
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

Public Sub AddMainMessage(ChatText As String)

'Add this value to the chat text array
ReDim Preserve MainMessages(NumMainMessages)
MainMessages(NumMainMessages).Decay = GetTickCount() + MainMessageDecay
MainMessages(NumMainMessages).Text = ChatText
NumMainMessages = NumMainMessages + 1

End Sub

Private Sub RemoveMainMessage(Index As Integer)

Dim i As Long

'Remove the specified chat text
For i = Index To NumMainMessages - 2
    MainMessages(i).Decay = MainMessages(i + 1).Decay
    MainMessages(i).Text = MainMessages(i + 1).Text
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

Private Sub EndWinsock()

'Kill winsock
modWinsock.DestroySocket socket
'modWinsock.TermWinsock

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim bCan As Boolean

On Error Resume Next

If m_bDesignMode = False Then
    
    If PlayerInGame(0) = False Then Exit Sub 'Player(0).Team = Spec Or Not Player(0).Alive Then Exit Sub
    
    If Not modSpaceGame.UseAI Then
        If modSpaceGame.cl_UseMouse And Button = vbLeftButton Then
            
            KeyFire = True
            
'            If Player(0).ShipType = MotherShip Then
'                If MSStartFire + MotherShipRechargeTime / modSpaceGame.sv_GameSpeed < GetTickCount() Then
'                    bCan = True
'                    MSStartFire = GetTickCount()
'                End If
'            Else
'                bCan = True
'            End If
'
'            If bCan Then AddPlayerState MyID, ePlayerState.Player_Fire
            
        ElseIf modSpaceGame.cl_UseMouse And Button = vbRightButton Then
            
            
            If Player(0).ShipType <> Infiltrator And Player(0).ShipType <> SD Then
                'AddPlayerState MyID, Player_Secondary 'fire on keyup
                KeySecondary = True
                
            ElseIf (Player(0).State And Player_Secondary) = Player_Secondary Then
                'we're an infil or SD
                
                'if already stealthed, remove it
                'SubPlayerState MyID, Player_Secondary
                'KeySecondary = False
                
                
                If Player(0).ShipType <> Infiltrator And Player(0).ShipType <> SD Then
                    KeySecondary = True
                Else 'If KeySecondary Then
                    If Player(0).State And Player_Secondary Then
                        SubPlayerState MyID, Player_Secondary
                    Else
                        AddPlayerState MyID, Player_Secondary
                    End If
                End If
                
                
            ElseIf Player(0).Shields > 5 Then
                'we're an infil or SD
                
                'AddPlayerState MyID, Player_Secondary
                KeySecondary = True
                
            End If
            
        End If
    End If
    
Else
    Dim i As Integer

    If Button = vbLeftButton And m_bDesignMode Then
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
            If Controls(i).Visible And Not TypeOf Controls(i) Is Menu Then
                m_DragRect.SetRectToCtrl Controls(i)
                If m_DragRect.PtInRect(X, Y) Then
                    DragBegin Controls(i)
                    Exit Sub
                End If
            End If
        Next i
        'No control is active
        Set m_CurrCtl = Nothing
        'Hide sizing handles
        ShowHandles False
    End If
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim St As eShipTypes
'Dim bCan As Boolean

If m_bDesignMode = False Then
    
    On Error Resume Next
    If (Not PlayerInGame(0)) And Button <> vbMiddleButton Then Exit Sub
    
    If modSpaceGame.cl_UseMouse And Button = vbLeftButton And Not modSpaceGame.UseAI Then
        
        'If Player(FindPlayer(MyID)).ShipType <> eShipTypes.MotherShip Then
        'SubPlayerState MyID, ePlayerState.Player_Fire
        KeyFire = False
        'End If
        'this is to make sure the fire message is transmitted if MotherShip - it is subbed in showmainmessage()
        
    ElseIf Button = vbMiddleButton Then
        If Playing Or modSpaceGame.SpaceServer Then frmGameOptions.Show vbModeless, Me
        
    ElseIf Button = vbRightButton Then
        
        KeySecondary = False
'        St = Player(FindPlayer(MyID)).ShipType
'
'        If St = MotherShip Or St = Wraith Then
'            SubPlayerState MyID, Player_Secondary
'        End If
        
    End If
    
Else
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
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If m_bDesignMode = False Then
    MouseX = X ' - 100
    MouseY = Y ' - 250
Else
    Dim nWidth As Single, nHeight As Single
    Dim Pt As POINTAPI
    
    
    If m_DragState = StateDragging Then
        'Save dimensions before modifying rectangle
        nWidth = m_DragRect.Right - m_DragRect.Left
        nHeight = m_DragRect.Bottom - m_DragRect.Top
        'Get current mouse position in screen coordinates
        GetCursorPos Pt
        'Hide existing rectangle
        DrawDragRect
        'Update drag rectangle coordinates
        m_DragRect.Left = Pt.X - m_DragPoint.X
        m_DragRect.Top = Pt.Y - m_DragPoint.Y
        m_DragRect.Right = m_DragRect.Left + nWidth
        m_DragRect.Bottom = m_DragRect.Top + nHeight
        'Draw new rectangle
        DrawDragRect
        
        bSaved = False
        
    ElseIf m_DragState = StateSizing Then
        'Get current mouse position in screen coordinates
        GetCursorPos Pt
        'Hide existing rectangle
        DrawDragRect
        'Action depends on handle being dragged
        Select Case m_DragHandle
            Case 0
                m_DragRect.Left = Pt.X
                m_DragRect.Top = Pt.Y
            Case 1
                m_DragRect.Top = Pt.Y
            Case 2
                m_DragRect.Right = Pt.X
                m_DragRect.Top = Pt.Y
            Case 3
                m_DragRect.Right = Pt.X
            Case 4
                m_DragRect.Right = Pt.X
                m_DragRect.Bottom = Pt.Y
            Case 5
                m_DragRect.Bottom = Pt.Y
            Case 6
                m_DragRect.Left = Pt.X
                m_DragRect.Bottom = Pt.Y
            Case 7
                m_DragRect.Left = Pt.X
        End Select
        'Draw new rectangle
        DrawDragRect
    End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Ct As String
Dim Ans As VbMsgBoxResult

If FindPlayer(MyID) <> -1 Then
    Ct = Trim$(Player(FindPlayer(MyID)).Name) & IIf(modSpaceGame.SpaceServer, " (Server)", vbNullString) & " left"
    
    If modSpaceGame.SpaceServer Then
        SendChatPacketBroadcast Ct, Player(0).Colour 'FindPlayer(MyID)).Colour
    Else
        On Error Resume Next
        modWinsock.SendPacket socket, ServerSockAddr, sChats & Ct & "#" & Player(0).Colour
    End If
    
    'remove myself
    If modSpaceGame.SpaceServer Then
        Ct = sServerQuits
        SendBroadcast Ct
    Else
        Ct = sQuits & CStr(MyID)
        modWinsock.SendPacket socket, ServerSockAddr, Ct
    End If
End If

modSpaceGame.GameFormLoaded = False

If modSpaceGame.GameOptionFormLoaded Then
    Unload frmGameOptions 'all sub-option forms are show as frmgameoptions' children - will unload
End If

'Disconnect winsock
EndWinsock

ClosingWindow = True

If modSpaceGame.SpaceServer Then
    
    Ct = eCommands.LobbyCmd & eLobbyCmds.Remove & modSpaceGame.SpaceServerIP
    
    If modVars.Server Then
        'modMessaging.DistributeMsg Ct, -1
        DataArrival Ct
    Else
        modMessaging.SendData Ct
    End If
    
End If


Call ResetVars

strChat = vbNullString
bChatActive = False

'remove from game lobby
'SendData eCommands.LobbyCmd & eLobbyCmds.Remove & modSpaceGame.SpaceServerIP
'If modSpaceGame.SpaceServer Then
'    If Server Then
'        modMessaging.LobbyStr = Replace$(modMessaging.LobbyStr, _
'            "#" & frmMain.SckLC.LocalIP & "," & "Space Combat" _
'            , vbNullString, , , vbTextCompare)
'    Else
'        SendData eCommands.LobbyCmd & eLobbyCmds.Remove & "#" & frmMain.SckLC.LocalIP & "," & "Space Combat"
'    End If
'End If

If bSaved = False And m_bDesignMode And Not modVars.Closing Then
    Ans = MsgBoxEx("Save Changes?", "Keep the boxes as you have places them?", vbQuestion + vbYesNo, "Box Positions", , , frmMain.Icon)
    If Ans = vbYes Then mnuSave_Click
End If

Call FormLoad(Me, True)

KeyW = False
KeyA = False
keys = False
KeyD = False
KeyFire = False
KeySecondary = False
KeyShield = False

End Sub

Private Function ClipBullet(i As Integer) As Boolean

Const Lim As Integer = 50, deg As Single = Pi / 6
Dim ClippedX As Boolean, ClippedY As Boolean
'Dim XComp As Single, YComp As Single

ClippedX = (Bullet(i).X < Lim) Or (Bullet(i).X > MaxWidth - Lim)
ClippedY = (Bullet(i).Y < Lim) Or (Bullet(i).Y > MaxHeight - Lim)


If ClippedX Or ClippedY Then
    
    If modSpaceGame.sv_BulletWallBounce Then
        
        If ClippedX Then
            ReverseXComp Bullet(i).Heading, Bullet(i).Speed
        Else
            ReverseYComp Bullet(i).Heading, Bullet(i).Speed
        End If
        
    Else
        ClipBullet = True
        
        RemoveBullet i ', False
    End If
    
'    With Bullet(i)
'        XComp = .Speed * sine(.Heading)
'        YComp = .Speed * cosine(.Heading)
'
'        If ClippedX Then
'            If Bullet(i).X < Lim Then
'                XComp = Abs(XComp)
'            Else
'                XComp = -Abs(XComp)
'            End If
'        Else
'            If Bullet(i).Y < Lim Then
'                YComp = -Abs(YComp)
'            Else
'                YComp = Abs(YComp)
'            End If
'        End If
'
'        'Determine the resultant speed
'        .Speed = Sqr(XComp ^ 2 + YComp ^ 2)
'
'        'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
'        If YComp > 0 Then .Heading = atn(XComp / YComp)
'        If YComp < 0 Then .Heading = atn(XComp / YComp) + Pi
'
'    End With
ElseIf (BulletVCollision(i, ln1) And BulletHCollision(i, ln1)) Then
    
    ReverseYComp Bullet(i).Heading, Bullet(i).Speed
    
ElseIf (BulletVCollision(i, ln2) And BulletHCollision(i, ln2)) Then
    
    ReverseXComp Bullet(i).Heading, Bullet(i).Speed
    
ElseIf BulletVCollision(i, ob1) And BulletHCollision(i, ob1) Then
    'slow it down
    
    With Bullet(i)
        If .LastDeflect + 10 / modSpaceGame.sv_GameSpeed < GetTickCount() Then
            .Heading = .Heading + deg * (Rnd() - Rnd())
            
            If .Speed > 75 Then
                .Speed = .Speed / 2
            End If
            
            .LastDeflect = GetTickCount()
            
        End If
        
    End With
    
    
'old (new)
'    With Bullet(i)
'
'        If .LastDeflect = 0 Then '+ 20 / modSpaceGame.sv_GameSpeed < GetTickCount() Then
'            .Heading = .Heading + deg * (Rnd() - Rnd())
'
'            '.LastDeflect = GetTickCount()
'            .LastDeflect = 1
'
'        End If
'
'        If .Speed > 30 Then
'            .Speed = .Speed / 2
'        End If
'    End With
    
    'Bullet(i).Decay = Bullet(i).Decay + 100
    
ElseIf BulletIsInAsteroid(i) Then
    
    'XComp = Bullet(i).Speed * sine(Bullet(i).Heading) + Asteroid.Speed * sine(Asteroid.Heading)
    'YComp = Bullet(i).Speed * cosine(Bullet(i).Heading) + Asteroid.Speed * cosine(Asteroid.Heading)
    
    'Asteroid.Speed = Sqr(XComp ^ 2 + YComp ^ 2)
    'Asteroid.Heading
    
    AddVectors Bullet(i).Speed / AsteroidMass, Bullet(i).Heading, Asteroid.Speed, Asteroid.Heading, Asteroid.Speed, Asteroid.Heading
    
    Asteroid.LastPlayerTouchID = Bullet(i).Owner
    
    RemoveBullet i, True ', Asteroid.Speed, Asteroid.Heading
    ClipBullet = True
    
End If

End Function

Private Sub ReverseYComp(ByRef Heading As Single, ByRef Speed As Single)

Dim XComp As Single, YComp As Single

XComp = Speed * Sine(Heading)
YComp = Speed * CoSine(Heading)

YComp = -YComp

'Determine the resultant speed
Speed = Sqr(XComp ^ 2 + YComp ^ 2)

'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
If YComp > 0 Then Heading = Atn(XComp / YComp)
If YComp < 0 Then Heading = Atn(XComp / YComp) + Pi

End Sub

Private Sub ReverseXComp(ByRef Heading As Single, ByRef Speed As Single)

Dim XComp As Single, YComp As Single

XComp = Speed * Sine(Heading)
YComp = Speed * CoSine(Heading)

XComp = -XComp

'Determine the resultant speed
Speed = Sqr(XComp ^ 2 + YComp ^ 2)

'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
If YComp > 0 Then Heading = Atn(XComp / YComp)
If YComp < 0 Then Heading = Atn(XComp / YComp) + Pi

End Sub

Private Function ClipMissile(i As Integer) As Boolean

Const Lim As Integer = 50
Dim ClippedX As Boolean, ClippedY As Boolean

ClippedX = (Missiles(i).X < Lim) Or (Missiles(i).X > MaxWidth - Lim)
ClippedY = (Missiles(i).Y < Lim) Or (Missiles(i).Y > MaxHeight - Lim)

If ClippedX Or ClippedY Then
    
    ClipMissile = True
    
    RemoveMissile i
    
ElseIf MissileVCollision(i, ob1) And MissileHCollision(i, ob1) Then
    
    With Missiles(i)
        'If .Speed > 75 Then
        'Motion Missiles(i).x, Missiles(i).y, Missiles(i).Speed / 2, -Missiles(i).Heading
        '.Heading = -.Heading
        .Speed = .Speed / 10
        'End If
    End With
    
ElseIf MissileVCollision(i, ln1) And MissileHCollision(i, ln1) Then
    
    ClipMissile = True
    
    RemoveMissile i
    
ElseIf MissileVCollision(i, ln2) And MissileHCollision(i, ln2) Then
    
    ClipMissile = True
    
    RemoveMissile i
    
ElseIf MissileIsInAsteroid(i) Then
    
    'XComp = Bullet(i).Speed * sine(Bullet(i).Heading) + Asteroid.Speed * sine(Asteroid.Heading)
    'YComp = Bullet(i).Speed * cosine(Bullet(i).Heading) + Asteroid.Speed * cosine(Asteroid.Heading)
    
    'Asteroid.Speed = Sqr(XComp ^ 2 + YComp ^ 2)
    'Asteroid.Heading
    
    AddVectors 10 * Missiles(i).Speed / AsteroidMass, Missiles(i).Heading, Asteroid.Speed, Asteroid.Heading, _
        Asteroid.Speed, Asteroid.Heading
    
    Asteroid.LastPlayerTouchID = Missiles(i).Owner
    
    RemoveMissile i
    ClipMissile = True
    
End If


End Function


Private Function ClipShip(i As Integer) As Boolean

Dim ClippedX As Boolean, ClippedY As Boolean
Dim XComp As Single, YComp As Single

If (ShipVCollision(i, ln1) And ShipHCollision(i, ln1)) Then
    
    With Player(i)
        XComp = .Speed * Sine(.Heading)
        YComp = .Speed * CoSine(.Heading)

'        If .X > (ln1.Left + ln1.Width \ 2) Then 'right side
'            XComp = Abs(XComp)
'        ElseIf .X < (ln1.Left + ln1.Width \ 2) Then
'            XComp = -Abs(XComp)
'        End If
        
        If .Y > (ln1.Top + ln1.height / 2) Then 'bottom side
            YComp = -Abs(YComp) '-Sgn(YComp) * Abs(YComp)
        Else 'If .Y > (ln1.Top + ln1.Height \ 2) Then - never happens
            YComp = Abs(YComp)
        End If
        
        
        'Determine the resultant speed
        .Speed = Sqr(XComp ^ 2 + YComp ^ 2)

        'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
        If YComp > 0 Then .Heading = Atn(XComp / YComp)
        If YComp < 0 Then .Heading = Atn(XComp / YComp) + Pi
        
    End With
    
ElseIf (ShipVCollision(i, ln2) And ShipHCollision(i, ln2)) Then
    
    With Player(i)
        XComp = .Speed * Sine(.Heading)
        YComp = .Speed * CoSine(.Heading)

'        If .X > (ln1.Left + ln1.Width \ 2) Then 'right side
'            XComp = Abs(XComp)
'        ElseIf .X < (ln1.Left + ln1.Width \ 2) Then
'            XComp = -Abs(XComp)
'        End If
        
        If .X > (ln2.Left + ln2.width / 2) Then 'right side
            XComp = Abs(XComp) '-Sgn(YComp) * Abs(YComp)
        Else 'If .Y > (ln1.Top + ln1.Height \ 2) Then - never happens
            XComp = -Abs(XComp)
        End If
        
        'Determine the resultant speed
        .Speed = Sqr(XComp ^ 2 + YComp ^ 2)

        'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
        If YComp > 0 Then .Heading = Atn(XComp / YComp)
        If YComp < 0 Then .Heading = Atn(XComp / YComp) + Pi
        
    End With
    
End If

End Function

Private Sub ClipEdges(i As Integer)

Const Lim As Integer = 50
Const ValIn = 30
Dim ClippedX As Boolean, ClippedY As Boolean
Dim XComp As Single, YComp As Single
Dim BoxX As Boolean, BoxY As Boolean, Col As Boolean
Dim bCan As Boolean

ClippedX = (Player(i).X < Lim) Or (Player(i).X > MaxWidth - Lim)
ClippedY = (Player(i).Y < Lim) Or (Player(i).Y > MaxHeight - Lim)


With Player(i)
    If ClippedX Or ClippedY Then
        XComp = .Speed * Sine(.Heading)
        YComp = .Speed * CoSine(.Heading)
    End If
    
    If ClippedX Then
        If Player(i).X < Lim Then
            XComp = Abs(XComp)
        Else
            XComp = -Abs(XComp)
        End If
    ElseIf ClippedY Then
        If Player(i).Y < Lim Then
            YComp = -Abs(YComp)
        Else
            YComp = Abs(YComp)
        End If
    End If
    
End With


If ClippedX Or ClippedY Then
    'Determine the resultant speed
    Player(i).Speed = Sqr(XComp ^ 2 + YComp ^ 2)
    
    'Calculate the resultant heading, and adjust for atngent by adding Pi if necessary
    If YComp > 0 Then Player(i).Heading = Atn(XComp / YComp)
    If YComp < 0 Then Player(i).Heading = Atn(XComp / YComp) + Pi
End If


With Player(i)
    If .X < 0 Then
        .X = ValIn
    ElseIf .X > MaxWidth Then
        .X = MaxWidth - ValIn
    End If
    
    If .Y < 0 Then
        .Y = ValIn
    ElseIf .Y > MaxHeight Then
        .Y = MaxHeight - ValIn
    End If
End With
'------------

'Call CheckCollisions(BoxX, BoxY, Col, i, ob1)

'If Col Then
'    If LeftSide(ob1, Player(i).x) Then
'        XComp = -Abs(XComp)
'    Else
'        XComp = Abs(XComp)
'    End If
'
'    If TopSide(ob1, Player(i).y) Then
'        YComp = -Abs(YComp)
'    Else
'        YComp = Abs(YComp)
'    End If
'
'End If
'goto top bit

bCan = True

If Player(i).ShipType = Infiltrator Then
    If (Player(i).State And Player_Secondary) = Player_Secondary Then
        bCan = False
    End If
End If

If bCan Then Call ClipShip(i)

End Sub

Private Function AsteroidCollision(ByRef Obj As Shape) As Boolean

Const Lim As Integer = 50
Dim HCol As Boolean, VCol As Boolean

If ((Asteroid.X + Asteroid_Radius) > (Obj.Left - Lim)) And _
    ((Asteroid.X - Asteroid_Radius) < (Obj.Left + Obj.width)) Then
    
    HCol = True
End If

If ((Asteroid.Y + Asteroid_Radius) > (Obj.Top - Lim)) And _
    ((Asteroid.Y - Asteroid_Radius) < (Obj.Top + Obj.height)) Then
    
    VCol = True
End If

AsteroidCollision = HCol And VCol

End Function

Private Function BulletHCollision(ByVal i As Integer, Obj As Shape) As Boolean

If (Bullet(i).X >= Obj.Left) And (Bullet(i).X <= (Obj.Left + Obj.width)) Then
    BulletHCollision = True
End If

End Function

Private Function BulletVCollision(ByVal i As Integer, Obj As Shape) As Boolean

If (Bullet(i).Y >= Obj.Top) And (Bullet(i).Y <= (Obj.Top + Obj.height)) Then
    BulletVCollision = True
End If

End Function

Private Function BulletIsInAsteroid(i As Integer) As Boolean

Const Lim As Integer = 50
Dim HCol As Boolean, VCol As Boolean

If ((Bullet(i).X + Lim) > (Asteroid.X - Asteroid_Radius)) And _
    ((Bullet(i).X - Lim) < (Asteroid.X + Asteroid_Radius)) Then
    
    HCol = True
End If

If ((Bullet(i).Y + Lim) > (Asteroid.Y - Asteroid_Radius)) And _
    ((Bullet(i).Y - Lim) < (Asteroid.Y + Asteroid_Radius)) Then
    
    VCol = True
End If

BulletIsInAsteroid = HCol And VCol

End Function

Private Function PlayerIsInAsteroid(i As Integer) As Boolean

Const Lim As Integer = 50
Dim HCol As Boolean, VCol As Boolean

If ((Player(i).X + Lim) > (Asteroid.X - Asteroid_Radius)) And _
    ((Player(i).X - Lim) < (Asteroid.X + Asteroid_Radius)) Then
    
    HCol = True
End If

If ((Player(i).Y + Lim) > (Asteroid.Y - Asteroid_Radius)) And _
    ((Player(i).Y - Lim) < (Asteroid.Y + Asteroid_Radius)) Then
    
    VCol = True
End If

PlayerIsInAsteroid = HCol And VCol

End Function

Private Function MissileIsInAsteroid(i As Integer) As Boolean

Const Lim As Integer = 50
Dim HCol As Boolean, VCol As Boolean

If ((Missiles(i).X + Lim) > (Asteroid.X - Asteroid_Radius)) And _
    ((Missiles(i).X - Lim) < (Asteroid.X + Asteroid_Radius)) Then
    
    HCol = True
End If

If ((Missiles(i).Y + Lim) > (Asteroid.Y - Asteroid_Radius)) And _
    ((Missiles(i).Y - Lim) < (Asteroid.Y + Asteroid_Radius)) Then
    
    VCol = True
End If

MissileIsInAsteroid = HCol And VCol

End Function

Private Function ShipHCollision(ByVal i As Integer, Obj As Shape) As Boolean
Const Lim As Integer = 100

If ((Player(i).X + Lim) >= Obj.Left) And ((Player(i).X - Lim) <= (Obj.Left + Obj.width)) Then
    ShipHCollision = True
End If

End Function

Private Function ShipVCollision(ByVal i As Integer, Obj As Shape) As Boolean
Const Lim As Integer = 200

If ((Player(i).Y + Lim) > Obj.Top) And ((Player(i).Y - Lim) <= (Obj.Top + Obj.height)) Then
    ShipVCollision = True
End If

End Function

Private Function MissileHCollision(ByVal i As Integer, Obj As Shape) As Boolean

If (Missiles(i).X >= Obj.Left) And (Missiles(i).X <= (Obj.Left + Obj.width)) Then
    MissileHCollision = True
End If

End Function

Private Function MissileVCollision(ByVal i As Integer, Obj As Shape) As Boolean

If (Missiles(i).Y >= Obj.Top) And (Missiles(i).Y <= (Obj.Top + Obj.height)) Then
    MissileVCollision = True
End If

End Function

'CurrentX = Width \ 2
'CurrentY = Player(i).Y + Lim
'Print "y + l"
'Print "S1: " & CStr((Player(i).Y + Lim) > Obj.Top)
'Print "S2: " & CStr((Player(i).Y - Lim) <= (Obj.Top + Obj.Height))
'
'Private Function LeftSide(Obj As Shape, ByVal x As Single) As Boolean
'
'Dim obL As Single, obW As Single
'
'obL = Obj.Left
'obW = Obj.Width
'
'If x <= (obL + 0.5 * obW) Then LeftSide = True
'
'End Function
'
'Private Function TopSide(Obj As Shape, ByVal y As Single) As Boolean
'
'Dim obT As Single, obH As Single
'
'obT = Obj.Top
'obH = Obj.Height
'
'If y <= (obT + 0.5 * obH) Then TopSide = True
'
'End Function

'Private Sub CheckCollisions(ByRef XCollision As Boolean, ByRef YCollision As Boolean, ByRef Collision As Boolean, _
'    i As Integer, Shp As Shape)
'
'XCollision = HCollision(i, Shp)
'YCollision = VCollision(i, Shp)
'
'Collision = XCollision And YCollision
'
'End Sub
'

Private Sub AccurateShot(sngTargetX As Single, sngTargetY As Single, sngTargetSpeed As Single, _
    sngTargetHeading As Single, sngSourceX As Single, sngSourceY As Single, sngSourceSpeed As Single, _
    sngSourceHeading As Single, sngProjectileSpeed As Single, ByRef sngAccurateSpeed As Single, _
    ByRef sngAccurateHeading As Single)

Dim sngDeltaX As Single
Dim sngDeltaY As Single
Dim sngDeltaSpeed As Single
Dim sngDeltaHeading As Single
Dim sngResultX As Single
Dim sngResultY As Single
Dim sngTResult As Single
Dim blnPossible As Boolean

Dim A As Single
Dim B As Single
Dim C As Single
Dim sq As Single
Dim t1 As Single
Dim t2 As Single

'Assume it's possible
blnPossible = True

'Determine the relative location of the target
sngDeltaX = sngTargetX - sngSourceX
sngDeltaY = sngTargetY - sngSourceY

'Subtract the velocity vectors to find the relative velocity
AddVectors sngTargetSpeed, sngTargetHeading, sngSourceSpeed, sngSourceHeading + Pi, sngDeltaSpeed, sngDeltaHeading

'Set up the quadratic equation's variables
A = (sngProjectileSpeed ^ 2 - sngDeltaSpeed ^ 2)
B = -(2 * sngDeltaSpeed * (sngDeltaX * Sine(sngDeltaHeading) - sngDeltaY * CoSine(sngDeltaHeading)))
C = -(sngDeltaX ^ 2 + sngDeltaY ^ 2)

'Ensure there's no problem with the square root, and no divide by zero
sq = (B ^ 2) - (4 * A * C)
If (sq < 0) Or (A = 0) Then
    blnPossible = False
Else
    'We're good to go, get the two results of the quadratic
    t1 = (-B - Sqr(sq)) / (2 * A)
    t2 = (-B + Sqr(sq)) / (2 * A)
    'Is the first Time value the optimal one?
    If t1 > 0 And t1 < t2 Then
        sngTResult = t1
    ElseIf t2 > 0 Then
        sngTResult = t2
    Else
        blnPossible = False
    End If
End If

'Is there a solution?
If blnPossible Then
    'Where will the target be, in sngTResult seconds?
    sngResultX = sngTargetX + sngTargetSpeed * Sine(sngTargetHeading) * sngTResult
    sngResultY = sngTargetY - sngTargetSpeed * CoSine(sngTargetHeading) * sngTResult
    'Return the angle to hit the target
    sngAccurateHeading = FindAngle(sngSourceX, sngSourceY, sngResultX, sngResultY)
    'Return the speed of the bullet (have to add the source's speed vector)
    AddVectors sngSourceSpeed, sngSourceHeading, sngProjectileSpeed, sngAccurateHeading, sngAccurateSpeed, 0
Else
    'It's not possible, just shoot straight at 'em
    AddVectors sngSourceSpeed, sngSourceHeading, sngProjectileSpeed, FindAngle(sngSourceX, sngSourceY, sngTargetX, sngTargetY), sngAccurateSpeed, sngAccurateHeading
End If

End Sub

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------


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

    'Use black Picture box controls for 8 sizing handles
    'Calculate size of each handle
    xHandle = 5 * Screen.TwipsPerPixelX
    yHandle = 5 * Screen.TwipsPerPixelY
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
    
    m_bDesignMode = True
End Sub

'Drags the specified control
Private Sub DragBegin(ctl As Control)
    Dim RC As RECT

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
    GetWindowRect hWnd, RC
    ClipCursor RC
    
End Sub

'Clears any current drag mode and hides sizing handles
Private Sub DragEnd()
    Set m_CurrCtl = Nothing
    ShowHandles False
    m_DragState = StateNothing
End Sub

Private Sub mnuReset_Click()
bSaved = False
Call ResetBoxPos
End Sub

Private Sub mnuSave_Click()


modSpaceGame.R_ob1.Left = ob1.Left / modSpaceGame.EditZoom
modSpaceGame.R_ob1.Top = ob1.Top / modSpaceGame.EditZoom
modSpaceGame.R_ob1.width = ob1.width / modSpaceGame.EditZoom
modSpaceGame.R_ob1.height = ob1.height / modSpaceGame.EditZoom

modSpaceGame.R_ln1.Left = ln1.Left / modSpaceGame.EditZoom
modSpaceGame.R_ln1.Top = ln1.Top / modSpaceGame.EditZoom
modSpaceGame.R_ln1.width = ln1.width / modSpaceGame.EditZoom
modSpaceGame.R_ln1.height = ln1.height / modSpaceGame.EditZoom

modSpaceGame.R_ln2.Left = ln2.Left / modSpaceGame.EditZoom
modSpaceGame.R_ln2.Top = ln2.Top / modSpaceGame.EditZoom
modSpaceGame.R_ln2.width = ln2.width / modSpaceGame.EditZoom
modSpaceGame.R_ln2.height = ln2.height / modSpaceGame.EditZoom


MsgBoxEx "Saved Positions", "Does what is says on the tin...", vbInformation, "Box Positions", , , frmMain.Icon

bSaved = True

End Sub

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
    Dim RC As RECT
    Dim CtrlName As String
    Dim bCan As Boolean
    
    'Handles should only be visible when a control is selected
    Debug.Assert (Not m_CurrCtl Is Nothing)
    
    CtrlName = m_CurrCtl.Name
    
    If CtrlName = ob1.Name Then
        bCan = True
    ElseIf (CtrlName = ln1.Name) And (Index = 7 Or Index = 3) Then
        bCan = True
    ElseIf (CtrlName = ln2.Name) And (Index = 5 Or Index = 1) Then
        bCan = True
'    Else
'        bCan = False
    End If
    
    If bCan Then
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
        GetWindowRect hWnd, RC
        ClipCursor RC
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

Private Sub tmrStart_Timer()
tmrStart.Enabled = False
Call MainLoop
Unload Me
End Sub

'##############################################################################
'Smoke ########################################################################
'##############################################################################

Private Sub AddSmokeGroup(ByVal X As Single, ByVal Y As Single, ByVal HowMany As Integer) ', _
    ByVal Speed As Single, ByVal Heading As Single)

Dim i As Integer
Const MaxSpacing = 75
Dim rX As Single, rY As Single

For i = 1 To HowMany
    rX = X + (Rnd() - 0.5) * MaxSpacing
    rY = Y + (Rnd() - 0.5) * MaxSpacing
    
    AddSmoke rX, rY ', Speed, Heading
Next i

End Sub

Private Sub AddSmoke(X As Single, Y As Single) ', Speed As Single, Heading As Single)

ReDim Preserve Smoke(NumSmoke)

Smoke(NumSmoke).X = X
Smoke(NumSmoke).Y = Y
Smoke(NumSmoke).Direction = 1
Smoke(NumSmoke).Size = 50 '0.4

'Smoke(NumSmoke).Speed = Speed
'Smoke(NumSmoke).Heading = Heading

NumSmoke = NumSmoke + 1

End Sub

Private Sub RemoveSmoke(ByVal Index As Integer)

Dim i As Integer

If NumSmoke = 1 Then
    Erase Smoke
    NumSmoke = 0
Else
    For i = Index To NumSmoke - 2
        Smoke(i).X = Smoke(i + 1).X
        Smoke(i).Y = Smoke(i + 1).Y
        Smoke(i).Size = Smoke(i + 1).Size
        Smoke(i).Direction = Smoke(i + 1).Direction
        
        'Smoke(i).Heading = Smoke(i + 1).Heading
        'Smoke(i).Speed = Smoke(i + 1).Speed
    Next i
    
    ReDim Preserve Smoke(NumSmoke - 2)
    NumSmoke = NumSmoke - 1
End If

End Sub

Private Sub ProcessSmoke()
Dim i As Integer
Dim f As Single

picMain.FillColor = SmokeFill
picMain.FillStyle = 0 'vbopaque
picMain.DrawWidth = 1

Do While i < NumSmoke
    
    If Smoke(i).Size <= 0 Then
        RemoveSmoke i
        i = i - 1
    ElseIf Smoke(i).Size > 40 Then
        Smoke(i).Direction = -1
    End If
    
    
    i = i + 1
    
Loop

For i = 0 To NumSmoke - 1
    
    With Smoke(i)
        
        'Motion .X, .Y, .Speed, .Heading
        
        gCircle .X, .Y, .Size, SmokeOutline
        
        
        If .Direction = 1 Then
            .Size = .Size + 4 * modSpaceGame.TimeFactor
        Else
            .Size = .Size - 4 * modSpaceGame.TimeFactor
        End If
        
    End With
    
Next i

picMain.FillStyle = 1 'transparent

End Sub

'Receiving + Sending Data
'#################################################################################################
'#################################################################################################

'Asteroid
'#################################################################################################


Private Function AsteroidToString() As String

AsteroidToString = CStr(Asteroid.X) & vbNullChar & _
         CStr(Asteroid.Y) & vbNullChar & _
         CStr(Asteroid.Speed) & vbNullChar & _
         CStr(Asteroid.Heading) & vbNullChar & _
         CStr(Asteroid.Facing) & vbNullChar & _
         CStr(Asteroid.LastPlayerTouchID)

End Function

Private Function AsteroidFromString(buf As String) As ptAsteroid

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, vbNullChar)

AsteroidFromString.X = CSng(Parts(0))
AsteroidFromString.Y = CSng(Parts(1))
AsteroidFromString.Speed = CSng(Parts(2))
AsteroidFromString.Heading = CSng(Parts(3))
AsteroidFromString.Facing = CSng(Parts(4))
AsteroidFromString.LastPlayerTouchID = CInt(Parts(5))

Erase Parts

EH:
End Function


Private Sub SendAsteroidUpdate()
Dim buf As String

buf = AsteroidToString()

SendBroadcast sAsteroidUpdates & buf

End Sub

Private Sub ReceiveAsteroid(sTxt As String)

Asteroid = AsteroidFromString(sTxt)

End Sub


'Box Pos
'#################################################################################################

Private Function SquareToString(ByRef Square As ptSquare) As String

SquareToString = CStr(Square.height) & vbNullChar & _
         CStr(Square.Left) & vbNullChar & _
         CStr(Square.Top) & vbNullChar & _
         CStr(Square.width) & vbNullChar

End Function

Private Function SquareFromString(buf As String) As ptSquare

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, vbNullChar)

SquareFromString.height = CSng(Parts(0))
SquareFromString.Left = CSng(Parts(1))
SquareFromString.Top = CSng(Parts(2))
SquareFromString.width = CSng(Parts(3))

Erase Parts

EH:
End Function


Private Sub SendBoxPos()

Static LastSend As Long
'Dim i As Integer
'Const Sep1 As String = "@"
'Const Sep2 As String = "#"

Dim S1 As ptSquare, S2 As ptSquare, S3 As ptSquare
Dim str1 As String, str2 As String, str3 As String

If LastSend + BoxPosDelay < GetTickCount() Then
    
    
    With S1
        .Left = ob1.Left
        .Top = ob1.Top
        .height = ob1.height
        .width = ob1.width
    End With
    
    'str1 = Space$(Len(S1) + 1)
    'CopyMemory ByVal str1, S1, Len(S1) + 1
    
    'On Error GoTo EH
    'LSet str1 = S1
    
    str1 = SquareToString(S1)
    
    '------------------------------------------
    
    With S2
        .Left = ln1.Left
        .Top = ln1.Top
        .height = ln1.height
        .width = ln1.width
    End With
    
    'str2 = Space$(Len(S2) + 1)
    'CopyMemory ByVal str2, S2, Len(S2) + 1
    
    'LSet str2 = S2
    
    str2 = SquareToString(S2)
    
    '------------------------------------------
    
    With S3
        .Left = ln2.Left
        .Top = ln2.Top
        .height = ln2.height
        .width = ln2.width
    End With
    
    'str3 = Space$(Len(S3) + 1)
    'CopyMemory ByVal str3, S3, Len(S3) + 1
    
    'LSet str3 = S3
    
    str3 = SquareToString(S3)
    
    '------------------------------------------
    
    SendBroadcast sBoxPoss & str1 & "#" & str2 & "#" & str3
    
    LastSend = GetTickCount()
End If

EH:
End Sub

Private Sub ReceiveBoxPos(ByVal sTxt As String)

Dim S1 As ptSquare, S2 As ptSquare, S3 As ptSquare
Dim str1 As String, str2 As String, str3 As String
'Dim l As Long
Dim i As Integer
Dim j As Integer

'l = Len(sTxt)

'If l = BoxPosLen Then
    
    j = InStr(1, sTxt, "#")
    i = InStr(j + 1, sTxt, "#")
    
    str1 = Left$(sTxt, j - 1)
    str2 = Mid$(sTxt, j + 1, i - j - 1)
    str3 = Mid$(sTxt, i + 1)
    
'    CopyMemory S1, ByVal str1, Len(str1)
'    CopyMemory S2, ByVal str2, Len(str2)
'    CopyMemory S3, ByVal str3, Len(str3)
    
    S1 = SquareFromString(str1)
    S2 = SquareFromString(str2)
    S3 = SquareFromString(str3)
    
    
    On Error Resume Next
    
    With ob1
        .Left = S1.Left
        .Top = S1.Top
        .height = S1.height
        .width = S1.width
    End With
    
    With ln1
        .Left = S2.Left
        .Top = S2.Top
        .height = S2.height
        .width = S2.width
    End With
    
    With ln2
        .Left = S3.Left
        .Top = S3.Top
        .height = S3.height
        .width = S3.width
    End With
'
'Else
'
'    AddConsoleText "BoxPos Len Error - Len: " & l
'
'End If

End Sub



'Position/Status
'#################################################################################################

Private Function mPacketToString() As String

With mPacket
    mPacketToString = _
        CStr(.Colour) & mPacketSep & _
        CStr(.Deaths) & mPacketSep & _
        CStr(.Facing) & mPacketSep & _
        CStr(.Heading) & mPacketSep & _
        CStr(.ID) & mPacketSep & _
        CStr(.Kills) & mPacketSep & _
        CStr(.Name) & mPacketSep & _
        CStr(.PacketID) & mPacketSep & _
        CStr(.Speed) & mPacketSep & _
        CStr(.State) & mPacketSep & _
        CStr(.X) & mPacketSep & _
        CStr(.Y) & mPacketSep & _
        CStr(.Alive) & mPacketSep
End With

End Function

Private Sub mPacketFromString(buf As String) 'As ptPacket

Dim Parts() As String

On Error GoTo EH
Parts = Split(buf, mPacketSep)

mPacket.Colour = CLng(Parts(0))
mPacket.Deaths = CInt(Parts(1))
mPacket.Facing = CSng(Parts(2))
mPacket.Heading = CSng(Parts(3))
mPacket.ID = CInt(Parts(4))
mPacket.Kills = CInt(Parts(5))
mPacket.Name = Parts(6)
mPacket.PacketID = CLng(Parts(7))
mPacket.Speed = CSng(Parts(8))
mPacket.State = CInt(Parts(9))
mPacket.X = CSng(Parts(10))
mPacket.Y = CSng(Parts(11))
mPacket.Alive = CBool(Parts(12))

Erase Parts


EH:
End Sub

Private Sub ProcessUpdatePacket(ByVal sPacket As String)

'Dim Num As Integer
Dim i As Integer, j As Integer ', k As Integer
Dim sPlayer As String
Dim Players() As String

'How many players' data is inside?
'Num = CInt(Mid$(sPacket, 2, 3))

'Chop the header off the packet
'sPacket = Right$(sPacket, Len(sPacket) - 4)

'If Num = 1 Then 'is from single client
'    l = Len(sPacket)
'    If l <> sPacketLen Then
'        'error!
'        AddConsoleText "Packet Error - Len: " & CStr(l)
'        Exit Sub
'    End If
'End If


sPacket = Mid$(sPacket, 2)

'now we have the pure packet
'like this: Player1Info#Player2Info#...

Players = Split(sPacket, UpdatePacketSep)


'Loop through each player's info
For i = 0 To UBound(Players)
    
    On Error GoTo EH
    
    'Extract player info
    'sPlayer = Left$(sPacket, Len(mPacket) + 1)
    'sPacket = Right$(sPacket, Len(sPacket) - (Len(mPacket) + 1))
    
    sPlayer = Players(i)
    
    If LenB(sPlayer) Then
        'CopyMemory mPacket, ByVal sPlayer, Len(sPlayer)
        
        'copy it into mPacket
        mPacketFromString sPlayer
        
        
        'Does this player already exist?
        If FindPlayer(mPacket.ID) = -1 Then
            'No, this is a new player.  Make new spot and assign ID
            Player(AddPlayer()).ID = mPacket.ID
        End If
        
        'Is this the local player?
        If mPacket.ID <> MyID Then
            'Is this a new packet?
            j = FindPlayer(mPacket.ID)
            If Player(j).LastPacketID < mPacket.PacketID Then
                'Replace player data with new data
                With Player(j)
                    .X = mPacket.X
                    .Y = mPacket.Y
                    .State = mPacket.State
                    .Facing = mPacket.Facing
                    .Heading = mPacket.Heading
                    .Speed = mPacket.Speed
                    .Name = mPacket.Name
                    .LastPacketID = mPacket.PacketID
                    .Colour = mPacket.Colour
                    '.ShipType = mPacket.ShipType
                    '.IsBot = CBool(mPacket.IsBot)
                    
                    If .Kills <= mPacket.Kills Then 'don't reduce
                        .Kills = mPacket.Kills
                    End If
                    
                    If .Deaths <= mPacket.Deaths Then 'don't reduce
                        .Deaths = mPacket.Deaths
                    End If
                    
                    .Alive = mPacket.Alive
                    
                    '.Team = mPacket.Team
                    
                    .LastPacket = GetTickCount()
                    
                End With
            End If 'packetid endif
        End If 'myid endif
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
    
    'Is this a server mPacket, or a client mPacket?
    If modSpaceGame.SpaceServer Then
        'Server mPacket
        SendServerUpdatePacket
    Else
        'Client mPacket
        SendClientUpdatePacket
    End If
End If

End Sub

Private Sub SendClientUpdatePacket()

Dim sPacket As String
Dim j As Integer

'Populate the mPacket type
j = FindPlayer(MyID)
Player(j).LastPacketID = Player(j).LastPacketID + 1

With mPacket
    .ID = MyID
    .State = Player(j).State
    .PacketID = Player(j).LastPacketID
    .Facing = Player(j).Facing
    .Heading = Player(j).Heading
    .Speed = Player(j).Speed
    .X = Player(j).X
    .Y = Player(j).Y
    .Name = Player(j).Name
    .Colour = Player(j).Colour
    '.ShipType = Player(j).ShipType
    .Kills = Player(j).Kills
    .Deaths = Player(j).Deaths
    '.IsBot = cint(player(j).IsBot
    '.IsBot = 0
    '.Team = Player(j).Team
    .Alive = Player(j).Alive
End With

'Copy this info o an update mPacket
'sPacket = Space$(Len(mPacket) + 1)
'CopyMemory ByVal sPacket, mPacket, Len(mPacket) + 1

'sPacket = sUpdates & "001" & mPacketToString()
sPacket = sUpdates & mPacketToString() & UpdatePacketSep

'Send position update to server
modWinsock.SendPacket socket, ServerSockAddr, sPacket

End Sub

Private Sub SendServerUpdatePacket()

Dim i As Long
Dim sPacket As String

i = FindPlayer(MyID)

'Increment the local player's LastPacketID
On Error GoTo EH
Player(i).LastPacketID = Player(i).LastPacketID + 1

For i = 0 To NumPlayers - 1
    'Fill the mPacket
    With mPacket
        .ID = Player(i).ID
        .State = Player(i).State
        .PacketID = IIf(Player(i).IsBot, Player(0).LastPacketID, Player(i).LastPacketID)
        .Facing = Player(i).Facing
        .Heading = Player(i).Heading
        .Speed = Player(i).Speed
        .X = Player(i).X
        .Y = Player(i).Y
        .Name = Player(i).Name
        .Colour = Player(i).Colour
        '.ShipType = Player(i).ShipType
        .Kills = Player(i).Kills
        .Deaths = Player(i).Deaths
        '.IsBot = CInt(Player(i).IsBot)
        '.Team = CByte(Player(i).Team)
        .Alive = Player(i).Alive
    End With
    
    sPacket = sPacket & mPacketToString() & UpdatePacketSep
Next i

'If Len(CStr(NumPlayers)) = 1 Then
'    sPacket = sUpdates & "00" & CStr(NumPlayers) & sPacket
'ElseIf Len(CStr(NumPlayers)) = 2 Then
'    sPacket = sUpdates & "0" & CStr(NumPlayers) & sPacket
'Else
'    sPacket = sUpdates & CStr(NumPlayers) & sPacket
'End If

sPacket = sUpdates & sPacket

'Send it to all non-local players
i = 1
Do While i < NumPlayers
    'Ensure this isn't the local player
    If Player(i).ID <> MyID And Player(i).IsBot = False Then
        'Send!
        If modWinsock.SendPacket(socket, Player(i).ptSockAddr, sPacket) = False Then
            'If there was an error sendString this mPacket, remove the player
            RemovePlayer CInt(i)
            i = i - 1
        End If
    End If
    'Increment the counter
    i = i + 1
Loop

EH:
End Sub




'old
'#################################################################################################

'Private Sub SendAsteroidUpdate()
'Static LastSend As Long
'Dim S As String
''Dim B As ptAsteroidBuff
'
'If LastSend + AsteroidSendDelay < GetTickCount() Then
'
'    S = Space$(Len(Asteroid) + 1)
'    CopyMemory ByVal S, Asteroid, Len(Asteroid) + 1
'
'    SendBroadcast sAsteroidUpdates & S
'
'    'On Error GoTo EH
'    'LSet B = Asteroid
'
'    'SendBroadcast sAsteroidUpdates & B.Data
'
'    'Asteroid.Facing = 0
'    'Asteroid.Heading = 0
'    'Asteroid.LastPlayerTouchID = 0
'    'Asteroid.Speed = 0
'    'Asteroid.X = 0
'    'Asteroid.Y = 0
'    'LSet Asteroid = B
'
'    LastSend = GetTickCount()
'End If
'
'EH:
'End Sub
'
'Private Sub ReceiveAsteroid(ByVal sTxt As String)
'Dim l As Integer
''Dim ABuff As ptAsteroidBuff
'
'l = Len(sTxt)
'
'If l = AsteroidLen Then
'
'    'ABuff.Data = sTxt
'
'    'LSet Asteroid = ABuff
'
'    CopyMemory Asteroid, ByVal sTxt, Len(sTxt)
'Else
'    AddConsoleText "Asteroid Len Error - Len: " & l
'End If
'
'End Sub
