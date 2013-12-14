VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStickOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stick Options"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMap 
      Caption         =   "Server Map"
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3960
      TabIndex        =   103
      Top             =   6960
      Width           =   3375
      Begin VB.PictureBox picMap 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   3135
         TabIndex        =   104
         Top             =   240
         Width           =   3135
         Begin VB.CommandButton cmdLoadMap 
            Caption         =   "Load Map"
            Height          =   375
            Left            =   480
            TabIndex        =   106
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cboMaps 
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   0
            Width           =   3135
         End
      End
   End
   Begin VB.Frame fraCol 
      Caption         =   "Stick Settings"
      ForeColor       =   &H00FF0000&
      Height          =   2775
      Left            =   7440
      TabIndex        =   15
      Top             =   120
      Width           =   4455
      Begin VB.PictureBox picCol 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   2415
         ScaleWidth      =   4215
         TabIndex        =   16
         Top             =   240
         Width           =   4215
         Begin VB.CheckBox chkTick 
            Caption         =   "Damage Tick"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   2160
            Width           =   1455
         End
         Begin VB.CheckBox chkShh 
            Caption         =   "Supressor/Silencer"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CheckBox chkLaserSight 
            Caption         =   "Laser Sight"
            Height          =   255
            Left            =   2400
            TabIndex        =   26
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   24
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CommandButton cmdNameApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   240
            MaxLength       =   20
            TabIndex        =   22
            Text            =   "<Name>"
            Top             =   960
            Width           =   3735
         End
         Begin VB.PictureBox picColour 
            Height          =   855
            Left            =   3840
            ScaleHeight     =   795
            ScaleWidth      =   195
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   18
            Top             =   0
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            SmallChange     =   5
            Max             =   255
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   240
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            SmallChange     =   5
            Max             =   255
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin MSComctlLib.Slider sldrCol 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   450
            _Version        =   393216
            SmallChange     =   5
            Max             =   255
            SelStart        =   100
            TickFrequency   =   10
            Value           =   100
         End
         Begin VB.Label lblColourInfo 
            AutoSize        =   -1  'True
            Caption         =   "R"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   150
         End
      End
   End
   Begin VB.Frame fraGraphics 
      Caption         =   "Graphics"
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   7440
      TabIndex        =   75
      Top             =   3120
      Width           =   4455
      Begin VB.PictureBox picGraphics 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   4215
         TabIndex        =   76
         Top             =   360
         Width           =   4215
         Begin VB.CheckBox chkExSmoke 
            Caption         =   "Explosion Smoke"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   109
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox cmdGraphics 
            Caption         =   "Low"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   0
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   2400
            Width           =   855
         End
         Begin VB.CheckBox cmdGraphics 
            Caption         =   "Medium"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   2400
            Width           =   975
         End
         Begin VB.CheckBox cmdGraphics 
            Caption         =   "High"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   1
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   2400
            Width           =   855
         End
         Begin VB.CheckBox chkTrails 
            Caption         =   "Bullet Trails"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   84
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox chkSparks 
            Caption         =   "Sparks"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   80
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkMagazines 
            Caption         =   "Magazines"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   81
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkDead 
            Caption         =   "Dead Sticks"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   78
            Top             =   240
            Width           =   1695
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00404040&
            Height          =   375
            Index           =   0
            Left            =   3600
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   90
            Top             =   1920
            Width           =   375
         End
         Begin VB.CheckBox chkCasing 
            Caption         =   "Bullet Casings"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   82
            Top             =   720
            Width           =   1695
         End
         Begin VB.CheckBox chkBlood 
            Caption         =   "Blood"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   83
            Top             =   960
            Width           =   1695
         End
         Begin VB.CheckBox chkSmoke 
            Caption         =   "Smoke"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Width           =   1575
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00FF8080&
            Height          =   375
            Index           =   3
            Left            =   2400
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   92
            Top             =   1920
            Width           =   375
         End
         Begin VB.PictureBox picBG 
            BackColor       =   &H00808080&
            Height          =   375
            Index           =   1
            Left            =   3000
            ScaleHeight     =   315
            ScaleWidth      =   315
            TabIndex        =   91
            Top             =   1920
            Width           =   375
         End
         Begin VB.CheckBox chkHolstered 
            Caption         =   "Holstered Weapon"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            TabIndex        =   85
            Top             =   0
            Width           =   1815
         End
         Begin VB.CheckBox chkInvert 
            Caption         =   "Invert Colours"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   87
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CheckBox chkSimple 
            Caption         =   "Draw simple weapons (Two Weapon Mode)"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   0
            TabIndex        =   88
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CheckBox chkWallMarks 
            Caption         =   "Wall Marks"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   79
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkSniperScope 
            Caption         =   "Use Sniper Scope"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   86
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkFPS 
            Caption         =   "Show FPS"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2400
            TabIndex        =   89
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Graphics:"
            Height          =   255
            Left            =   0
            TabIndex        =   94
            Top             =   2460
            Width           =   855
         End
         Begin VB.Label lblColour 
            Caption         =   "Background Colour:"
            Height          =   255
            Left            =   0
            TabIndex        =   93
            Top             =   2040
            Width           =   1455
         End
      End
   End
   Begin VB.Frame fraSound 
      Caption         =   "Sound"
      Height          =   1335
      Left            =   7440
      TabIndex        =   98
      Top             =   6600
      Width           =   4455
      Begin VB.PictureBox picSound 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   4215
         TabIndex        =   99
         Top             =   240
         Width           =   4215
         Begin VB.CheckBox chkEnableSound 
            Caption         =   "Enable Sound"
            Height          =   255
            Left            =   0
            TabIndex        =   100
            Top             =   0
            Width           =   3015
         End
         Begin MSComctlLib.Slider sldrVol 
            Height          =   255
            Left            =   0
            TabIndex        =   102
            Top             =   600
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   100
            SmallChange     =   10
            Min             =   -3000
            Max             =   0
            TickFrequency   =   300
         End
         Begin VB.Label lblVol 
            Caption         =   "Volume:"
            Height          =   255
            Left            =   0
            TabIndex        =   101
            Top             =   360
            Width           =   1695
         End
      End
   End
   Begin VB.Frame fraTeam 
      Caption         =   "Team Settings"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   3960
      TabIndex        =   7
      Top             =   360
      Width           =   3375
      Begin VB.PictureBox picTeam 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1215
         ScaleWidth      =   3135
         TabIndex        =   8
         Top             =   240
         Width           =   3135
         Begin VB.OptionButton optnTeam 
            Caption         =   "Neutral"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Caption         =   "Red"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   10
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optnTeam 
            Caption         =   "Blue"
            Height          =   255
            Index           =   2
            Left            =   1920
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optnTeam 
            Caption         =   "Spectator"
            Height          =   255
            Index           =   3
            Left            =   1920
            TabIndex        =   12
            Top             =   0
            Width           =   1455
         End
         Begin MSComctlLib.Slider sldrSpecSpeed 
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   450
            _Version        =   393216
            Min             =   5
            Max             =   20
            SelStart        =   10
            TickFrequency   =   5
            Value           =   10
         End
         Begin VB.Label lblSpecSpeed 
            Caption         =   "Spectator Speed - WW"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   600
            Width           =   2175
         End
      End
   End
   Begin VB.Frame fraPerks 
      Caption         =   "Perks"
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Left            =   3960
      TabIndex        =   55
      Top             =   2040
      Width           =   3375
      Begin VB.PictureBox picPerks 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   3135
         TabIndex        =   56
         Top             =   240
         Width           =   3135
         Begin VB.OptionButton optnPerk 
            Caption         =   "Zombie - High Health, Meleé only"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   69
            Top             =   3000
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Spy - Take another's name and colour"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   70
            Top             =   3240
            Width           =   3135
         End
         Begin VB.PictureBox picSpy 
            Height          =   375
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   3075
            TabIndex        =   71
            Top             =   3600
            Width           =   3135
            Begin VB.ComboBox cboSpyStick 
               Height          =   315
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   0
               Width           =   1935
            End
            Begin VB.Label lblStick 
               Alignment       =   2  'Center
               Caption         =   "Stick Mask:"
               Height          =   255
               Left            =   0
               TabIndex        =   72
               Top             =   0
               Width           =   975
            End
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Mechanic - Rapid Fire"
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   67
            Top             =   2520
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Steady Aim - Less Recoil"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   66
            Top             =   2280
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "No Perk"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   1815
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Juggernaut - Take less damage"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   58
            Top             =   240
            Width           =   3015
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Sleight of Hand - Reload more quickly"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   59
            Top             =   480
            Width           =   3015
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Stopping Power - Higer bullet damage"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   60
            Top             =   720
            Width           =   3015
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Bomb Squad (Explosive Detection)"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   61
            Top             =   960
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Conditioning - Run more quickly"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   62
            Top             =   1200
            Width           =   2655
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Sniper/Stealth - Name is hidden in prone + sniper rifle training"
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   63
            Top             =   1440
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Martyrdom - Drop a grenade on death"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   65
            Top             =   2040
            Width           =   3135
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Focus - Zoom for all weapons"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   64
            Top             =   1800
            Width           =   3135
         End
         Begin VB.CommandButton cmdPerkApply 
            Caption         =   "Apply"
            Enabled         =   0   'False
            Height          =   375
            Left            =   600
            TabIndex        =   74
            Top             =   4080
            Width           =   1815
         End
         Begin VB.OptionButton optnPerk 
            Caption         =   "Deep Impact - Bullets go deeper"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   68
            Top             =   2760
            Width           =   3135
         End
      End
   End
   Begin VB.Frame fraViewControls 
      Height          =   3375
      Left            =   3480
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
      Begin VB.Label lblControlInfo 
         AutoSize        =   -1  'True
         Caption         =   "Control Info"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.Frame fraControls 
      Caption         =   "Controls"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox picControls 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   3495
         TabIndex        =   1
         Top             =   240
         Width           =   3495
         Begin VB.CheckBox chkToggleCrouch 
            Caption         =   "Crouch is togglable"
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   2175
         End
         Begin VB.CommandButton cmdControls 
            Caption         =   "View Controls"
            Height          =   375
            Left            =   2160
            TabIndex        =   5
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox chkMiddle 
            Caption         =   "Middle Button - Drop Mine"
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox chkAutoCamera 
            Caption         =   "Auto Zoom Camera"
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkScroll 
            Caption         =   "Enable Scroll Wheel"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server Options"
      ForeColor       =   &H00FF0000&
      Height          =   6255
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   3735
      Begin VB.PictureBox picSv 
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   3495
         TabIndex        =   31
         Top             =   240
         Width           =   3495
         Begin VB.CommandButton cmdWeaponList 
            Caption         =   "Allowed Weapon List"
            Height          =   375
            Left            =   0
            TabIndex        =   108
            Top             =   5520
            Width           =   3375
         End
         Begin VB.CheckBox chkSpawnShields 
            Caption         =   "Sticks spawn with shields"
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtBulletDamage 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   42
            Text            =   "100"
            Top             =   2160
            Width           =   855
         End
         Begin VB.CommandButton cmdBot 
            Caption         =   "Bot Settings"
            Height          =   375
            Left            =   1800
            TabIndex        =   54
            Top             =   5040
            Width           =   1575
         End
         Begin VB.CommandButton cmdKick 
            Caption         =   "Kick Player"
            Height          =   375
            Left            =   0
            TabIndex        =   53
            Top             =   5040
            Width           =   1575
         End
         Begin VB.Timer tmrRefresh 
            Interval        =   3000
            Left            =   2880
            Top             =   4800
         End
         Begin VB.CheckBox chkNadeTime 
            Caption         =   "Show time left on grenades"
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   120
            Width           =   3255
         End
         Begin VB.TextBox txtSpawnDelay 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   40
            Text            =   "3"
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chkWalls 
            Caption         =   "Bullets can pass through walls"
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   360
            Width           =   2775
         End
         Begin VB.OptionButton optnGameType 
            Caption         =   "Co-Op"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   45
            Top             =   2640
            Width           =   855
         End
         Begin VB.OptionButton optnGameType 
            Caption         =   "Elimination"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   44
            Top             =   2640
            Width           =   1215
         End
         Begin VB.OptionButton optnGameType 
            Caption         =   "Deathmatch"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox txtWinScore 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   2400
            TabIndex        =   38
            Text            =   "10"
            Top             =   1440
            Width           =   855
         End
         Begin VB.CheckBox chkHPBonus 
            Caption         =   "Health Bonus on kill (Excluding Choppers)"
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   840
            Width           =   3255
         End
         Begin VB.CheckBox chkHC 
            Caption         =   "Hardcore Mode"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   600
            Width           =   2775
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "Bullet Time"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   49
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "Normal"
            Height          =   375
            Index           =   2
            Left            =   2520
            TabIndex        =   51
            Top             =   3600
            Width           =   735
         End
         Begin VB.CommandButton cmdAutoSpeed 
            Caption         =   "Slow"
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   50
            Top             =   3600
            Width           =   735
         End
         Begin VB.CommandButton cmdSetSpeed 
            Caption         =   "Set"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   48
            Top             =   3240
            Width           =   495
         End
         Begin MSComctlLib.Slider sldrSpeed 
            Height          =   255
            Left            =   0
            TabIndex        =   47
            ToolTipText     =   "Speed of the game - 0.1 to 1.2"
            Top             =   3240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   3
            Min             =   1
            Max             =   11
            SelStart        =   10
            Value           =   10
         End
         Begin projMulti.ScrollListBox lstMain 
            Height          =   855
            Left            =   0
            TabIndex        =   52
            Top             =   4080
            Width           =   3375
            _ExtentX        =   9340
            _ExtentY        =   4683
         End
         Begin VB.Label lblBulletDamage 
            Alignment       =   2  'Center
            Caption         =   "Damage Percent"
            Height          =   255
            Left            =   0
            TabIndex        =   41
            Top             =   2160
            Width           =   2295
         End
         Begin VB.Label lblSpawnDelay 
            Alignment       =   2  'Center
            Caption         =   "Spawn Delay (in seconds)"
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label lblWinScore 
            Caption         =   "Score (Kills - Deaths) to Win:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Speed - WW"
            Height          =   255
            Left            =   0
            TabIndex        =   46
            Top             =   3000
            Width           =   1695
         End
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Status: Loaded Window"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   107
      Top             =   7800
      Width           =   3375
   End
End
Attribute VB_Name = "frmStickOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bColClicked As Boolean, bStatusClicked As Boolean, bShowLim As Boolean 'last is a dirty cheap hack because i cbb writing 1 line methods
Private Const MinWinScore = 5, MaxWinScore = 100, Max_Spawn_Delay = 20, Max_Damage = 12 * 100, Min_Damage = 0.1 * 100
Private Const lblWinScoreCap = "Score (Kills - Deaths) to Win:"
Private psCurrentMapName As String

Private Sub chkAutoCamera_Click()
modStickGame.cg_AutoCamera = CBool(chkAutoCamera.Value)
End Sub

Private Sub chkExSmoke_Click()
modStickGame.cg_ExSmoke = CBool(chkExSmoke.Value)
End Sub

Private Sub chkHC_Click()
modStickGame.sv_Hardcore = CBool(chkHC.Value)
UpdateServerVars
End Sub

Private Sub chkHPBonus_Click()
modStickGame.sv_HPBonus = CBool(chkHPBonus.Value)
UpdateServerVars
End Sub

'Private Sub chk2Weapons_Click()
'
'If Me.Visible Then
'    modStickGame.sv_2Weapons = CBool(chk2Weapons.Value)
'    UpdateServerVars
'
'    If modStickGame.sv_2Weapons Then
'        frmStickGame.MakeStaticWeapons
'
'        frmStickGame.SetCurrentWeapons
'    Else
'        frmStickGame.RemoveStaticWeapons
'    End If
'End If
'
'End Sub

Private Sub UpdateServerVars()
If Me.Visible Then frmStickGame.SendServerVarPacket True
End Sub

Private Sub chkMiddle_Click()
modStickGame.cl_MiddleMineDrop = CBool(chkMiddle.Value)
End Sub

Private Sub chkNadeTime_Click()
modStickGame.sv_Draw_Nade_Time = CBool(chkNadeTime.Value)
UpdateServerVars
End Sub

Private Sub chkScroll_Click()

If Me.Visible Then
    If CBool(chkScroll.Value) Then
        If modSubClass.bStickSubClassing = False Then
            modSubClass.SubClassStick frmStickGame.hWnd
        End If
    ElseIf modSubClass.bStickSubClassing Then
        modSubClass.SubClassStick frmStickGame.hWnd, False
    End If
    
    modStickGame.cl_Subclass = CBool(chkScroll.Value)
End If

End Sub

Private Sub chkSpawnShields_Click()
modStickGame.sv_SpawnWithShields = CBool(chkSpawnShields.Value)
End Sub

Private Sub chkTick_Click()
modStickGame.cl_DamageTick = CBool(chkTick.Value)
End Sub

Private Sub chkToggleCrouch_Click()
modStickGame.cl_ToggleCrouch = CBool(chkToggleCrouch.Value)
End Sub

Private Sub chkTrails_Click()
modStickGame.cg_ShowBulletTrails = CBool(chkTrails.Value)
End Sub

Private Sub chkWalls_Click()
modStickGame.sv_BulletsThroughWalls = CBool(chkWalls.Value)
UpdateServerVars
End Sub

Private Sub cmdGraphics_Click(Index As Integer)

cmdGraphics(Index).Value = 0

If Index < 2 Then
    'high and low
    chkSmoke.Value = Index
    chkWallMarks.Value = Index
    chkMagazines.Value = Index
    chkBlood.Value = Index
    chkHolstered.Value = Index
    chkSimple.Value = (1 - Index)
    chkDead.Value = Index
    chkSparks.Value = Index
    chkCasing.Value = Index
    chkTrails.Value = Index
Else
    'medium
    chkSmoke.Value = 0
    chkWallMarks.Value = 1
    chkMagazines.Value = 1
    chkBlood.Value = 1
    chkHolstered.Value = 1
    chkSimple.Value = 1
    chkDead.Value = 1
    chkSparks.Value = 1
    chkCasing.Value = 1
    chkTrails.Value = 0
End If

End Sub

Private Sub cmdWeaponList_Click()
Load frmStickWeapons
frmStickWeapons.Show vbModeless, frmStickGame
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

'Private Sub cmdAddBot_Click()
'Dim i As Integer
'
''cmdAddBot.Enabled = False
'
'For i = 0 To NumSticks - 1
'    If Stick(i).IsBot Then Exit Sub
'Next i
'
'frmStickGame.AddStick True
''cmdRemoveBot.Enabled = True
'
'End Sub



Private Sub optnGameType_Click(Index As Integer)
Dim i As Integer
Dim bEn As Boolean

If modStickGame.sv_GameType <> Index Then
    modStickGame.sv_GameType = Index
    frmStickGame.GameTypeChanged
    UpdateServerVars
    
    If modStickGame.sv_GameType = gDeathMatch Then
        For i = 0 To NumSticks - 1
            Stick(i).bAlive = True
        Next i
    End If
End If


bEn = (modStickGame.sv_GameType = gDeathMatch)
lblWinScore.Enabled = bEn
lblSpawnDelay.Enabled = bEn
txtWinScore.Enabled = bEn
txtSpawnDelay.Enabled = bEn


End Sub

Private Sub cmdAutoSpeed_Click(Index As Integer)

With sldrSpeed
    Select Case Index
        Case 0
            .Value = 1 '0.1
        Case 1
            .Value = 5 '0.5
        Case 2
            .Value = 10 '1
    End Select
End With

cmdSetSpeed.Default = True

SetFocus2 cmdSetSpeed

End Sub

'Private Sub cmdRemoveBot_Click()
'Dim i As Integer
'
''cmdRemoveBot.Enabled = False
'
'For i = 0 To NumSticks - 1
'    If Stick(i).IsBot Then
'        frmStickGame.RemoveStick i
'        'cmdAddBot.Enabled = True
'        Exit Sub
'    End If
'Next i
'
'End Sub

Private Sub cmdBot_Click()
Unload frmStickBot
Load frmStickBot
frmStickBot.Show vbModeless, frmStickGame
End Sub

Private Sub cmdControls_Click()
fraViewControls.Visible = Not fraViewControls.Visible

cmdControls.Caption = IIf(fraViewControls.Visible, "Hide Controls", "View Controls")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call Stick_FormLoad(Me, True)
modStickGame.StickOptionFormLoaded = False
End Sub

Private Sub fraServer_DblClick()

chkNadeTime.Value = 0

txtBulletDamage.Text = "150%"
txtBulletDamage_KeyPress vbKeyReturn

txtSpawnDelay.Text = "3"
txtSpawnDelay_KeyPress vbKeyReturn

txtWinScore.Text = "30"
txtWinScore_KeyPress vbKeyReturn

End Sub

Private Sub fraControls_DblClick()
bColClicked = True
Call CheckDev
End Sub

Private Sub lblStatus_DblClick()
bStatusClicked = True
Call CheckDev
End Sub

Private Sub CheckDev()
If bColClicked Then
    If bStatusClicked Then
        If bDevMode Then
            
            If Stick(0).WeaponType <> Chopper Then
                frmStickGame.ChopperAvail = True
            Else
                frmStickGame.SwitchWeapon Stick(0).CurrentWeapons(1)
            End If
            
        End If
    End If
End If
End Sub

Private Sub sldrSpeed_Scroll()
sldrSpeed_Change
End Sub
Private Sub sldrSpeed_Change()
sldrSpeed_Click

If Me.Visible Then
    cmdSetSpeed.Enabled = True
    cmdSetSpeed.Default = True
    
    SetFocus2 cmdSetSpeed
End If

End Sub

Private Sub sldrSpeed_Click()
Const lblCap As String = "Speed - "
lblSpeed.Caption = lblCap & CStr(sldrSpeed.Value * 10) & "%"
End Sub

Private Sub cmdSetSpeed_Click()
cmdSetSpeed.Enabled = False

If Me.Visible Then
    frmStickGame.StickGameSpeedChanged modStickGame.sv_StickGameSpeed, sldrSpeed.Value / 10
    
    modStickGame.sv_StickGameSpeed = sldrSpeed.Value / 10
    
    frmStickGame.SendServerVarPacket True
End If

End Sub

Private Sub Form_Load()
Dim i As Integer
Const sControls As String = "Controls" & vbNewLine & vbNewLine & _
                         "W, A, S and D to move" & vbNewLine & _
                         "Space to jump" & vbNewLine & _
                         "Hold Ctrl to crouch" & vbNewLine & _
                         "C or F for prone" & vbNewLine & _
                         "R to reload" & vbNewLine & _
                         "4 to toggle firemode" & vbNewLine & _
                         "Q to toggle supressor" & vbNewLine & _
                         "Z to toggle laser sight" & vbNewLine & _
                         "B or 3 to swap grenade type" & vbNewLine & _
                         "M to drop a mine" & vbNewLine & _
                         "K or V for a knife" & vbNewLine & _
                         "T to talk" & vbNewLine & _
                         "F1 to view scores"


'tooltips off
TurnOffToolTip sldrSpeed.hWnd
fraViewControls.Left = 120
fraViewControls.width = 3735
fraViewControls.Visible = False

lblControlInfo.Caption = sControls

bColClicked = False
bStatusClicked = False

On Error Resume Next 'if user presses tab before loading window
chkAutoCamera.Value = IIf(modStickGame.cg_AutoCamera, 1, 0)
chkHC.Value = IIf(modStickGame.sv_Hardcore, 1, 0)
chkScroll.Value = IIf(modSubClass.bStickSubClassing, 1, 0)

chkMiddle.Value = IIf(modStickGame.cl_MiddleMineDrop, 1, 0)
sldrSpeed.Value = modStickGame.sv_StickGameSpeed * 10
chkHPBonus.Value = IIf(modStickGame.sv_HPBonus, 1, 0)
'chk2Weapons.Value = Abs(modStickGame.sv_2Weapons)
optnGameType(modStickGame.sv_GameType).Value = True
txtWinScore.Text = CStr(modStickGame.sv_WinScore)
chkWalls.Value = Abs(modStickGame.sv_BulletsThroughWalls)
chkTick.Value = Abs(modStickGame.cl_DamageTick)
chkExSmoke.Value = Abs(modStickGame.cg_ExSmoke)

chkNadeTime.Value = Abs(modStickGame.sv_Draw_Nade_Time)

txtSpawnDelay.Text = CStr(modStickGame.sv_Spawn_Delay / 1000)
txtBulletDamage.Text = CStr(modStickGame.sv_Damage_Factor * 100) & "%"

chkToggleCrouch.Value = Abs(modStickGame.cl_ToggleCrouch)
chkSpawnShields.Value = Abs(modStickGame.sv_SpawnWithShields)

If Not modStickGame.StickServer Then
    fraServer.Enabled = False
    
    chkHC.Enabled = False
    
    sldrSpeed.Enabled = False
    lblSpeed.Enabled = False
    cmdSetSpeed.Enabled = False
    For i = 0 To cmdAutoSpeed.UBound
        cmdAutoSpeed(i).Enabled = False
    Next i
    
    
    chkHPBonus.Enabled = False
    'chk2Weapons.Enabled = False
    
    txtWinScore.Enabled = False
    lblWinScore.Enabled = False
    txtSpawnDelay.Enabled = False
    lblSpawnDelay.Enabled = False
    txtBulletDamage.Enabled = False
    lblBulletDamage.Enabled = False
    
    chkWalls.Enabled = False
    chkNadeTime.Enabled = False
    
    For i = 0 To optnGameType.UBound
        optnGameType(i).Enabled = False
    Next i
    
    cmdBot.Enabled = False
    chkSpawnShields.Enabled = False
    cmdWeaponList.Enabled = False
    
'Else
'    'cmdAddBot.Enabled = True
'    For i = 0 To NumSticks - 1
'        If Stick(i).IsBot Then
'            'cmdAddBot.Enabled = False
'            Exit For
'        End If
'    Next i
End If

sldrSpeed_Click

Perk_Form_Load
Graphics_Form_Load
Client_Form_Load
Map_Form_Load

'pos
Me.Top = frmStickGame.Top + frmStickGame.height / 2 - Me.height / 2
If Me.Top < 0 Then Me.Top = 0

Me.Left = frmStickGame.Left + frmStickGame.width '/ 2 - Me.Width / 2
If (Me.Left + Me.width) > Screen.width Then
    Me.Left = Screen.width - Me.width - 10
End If
If Me.Left < 0 Then
    Me.Left = Screen.width - Me.width
End If
'end pos

SetStatus "Loaded Window"

Call Stick_FormLoad(Me)

modStickGame.StickOptionFormLoaded = True
End Sub

Private Sub txtWinScore_Change()
Dim sScore As String, iScore As Integer

If Me.Visible Then
    sScore = txtWinScore.Text
    
    If LenB(sScore) Then
        On Error GoTo EH
        iScore = val(sScore)
        
        If iScore >= MinWinScore Then
            If iScore <= MaxWinScore Then
                'modStickGame.sv_WinScore = iScore
                txtWinScore.ForeColor = vbBlue
                'lblInfo.Visible = True
                lblWinScore.Caption = "Press Enter to Set Score"
                
                modDisplay.ShowBalloonTip txtWinScore, "Press Enter!", _
                    "Press enter to set the score to " & CStr(iScore)
                
            Else
                txtWinScore.ForeColor = vbRed
                lblWinScore.Caption = lblWinScoreCap
                
                bShowLim = True
                
                modDisplay.ShowBalloonTip txtWinScore, "You crazy person", _
                    "The score is too large. It must be less than " & CStr(MaxWinScore + 1)
                
            End If
        Else
            txtWinScore.ForeColor = vbRed
            lblWinScore.Caption = lblWinScoreCap
            
            bShowLim = True
            
            modDisplay.ShowBalloonTip txtWinScore, "You crazy person", _
                "The score is too small. It must be greater than " & CStr(MinWinScore - 1)
            
        End If
    Else
        txtWinScore.ForeColor = vbRed
        lblWinScore.Caption = lblWinScoreCap
    End If
End If

txtWinScore_Click

EH:
End Sub

Private Sub txtWinScore_Click()
SetStatus MinWinScore & " <= Score <= " & MaxWinScore, bShowLim
End Sub

Private Sub txtWinScore_KeyPress(KeyAscii As Integer)

If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then
        If KeyAscii = 13 Then
            'enter, accept if we can
            If txtWinScore.ForeColor = vbBlue Then
                modStickGame.sv_WinScore = val(txtWinScore.Text)
                txtWinScore.ForeColor = MGreen
                lblWinScore.Caption = lblWinScoreCap
                
                modDisplay.ShowBalloonTip txtWinScore, "Score Set", _
                    "Score set to " & CStr(modStickGame.sv_WinScore)
                
            End If
        End If
        
        
        KeyAscii = 0
    End If
End If

End Sub

'#####################################################################################

Private Sub txtSpawnDelay_KeyPress(KeyAscii As Integer)

If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then
        If KeyAscii = 13 Then
            'enter, accept if we can
            If txtSpawnDelay.ForeColor = vbBlue Then
                modStickGame.sv_Spawn_Delay = val(txtSpawnDelay.Text) * 1000
                
                txtSpawnDelay.ForeColor = MGreen
                lblSpawnDelay.Caption = "Spawn Delay Set"
                
                modDisplay.ShowBalloonTip txtSpawnDelay, "Spawn Delay Set", _
                    "Delay set to " & CStr(modStickGame.sv_Spawn_Delay / 1000) & " seconds"
                
                UpdateServerVars
            End If
        End If
        
        
        KeyAscii = 0
    End If
End If

End Sub

Private Sub txtSpawnDelay_Change()
Dim sDelay As String, iDelay As Integer
Const Spawn_Delay_Cap = "Spawn Delay (in seconds)"

If Me.Visible Then
    sDelay = txtSpawnDelay.Text
    
    If LenB(sDelay) Then
        On Error GoTo EH
        iDelay = val(sDelay)
        
        If iDelay >= 1 Then
            If iDelay <= Max_Spawn_Delay Then
                
                txtSpawnDelay.ForeColor = vbBlue
                lblSpawnDelay.Caption = "Press Enter to Set Spawn Delay"
                
                
                modDisplay.ShowBalloonTip txtSpawnDelay, "Press Enter!", _
                    "Press enter to set the spawn delay to " & CStr(iDelay) & " seconds"
                
            Else
                txtSpawnDelay.ForeColor = vbRed
                lblSpawnDelay.Caption = Spawn_Delay_Cap
                
                bShowLim = True
                
                modDisplay.ShowBalloonTip txtSpawnDelay, "You crazy person", _
                    "The delay is too large. It must be less than " & CStr(Max_Spawn_Delay + 1) & " seconds", TTI_ERROR
                
            End If
        Else
            txtSpawnDelay.ForeColor = vbRed
            lblSpawnDelay.Caption = Spawn_Delay_Cap
            
            bShowLim = True
            
            modDisplay.ShowBalloonTip txtSpawnDelay, "You crazy person", _
                "The delay is too small. It must be greater than or equal to 1 second", TTI_ERROR
            
        End If
    Else
        txtSpawnDelay.ForeColor = vbRed
        lblSpawnDelay.Caption = Spawn_Delay_Cap
    End If
End If

txtspawndelay_click

EH:
End Sub

Private Sub txtspawndelay_click()
SetStatus "1 <= Spawn Delay <= " & Max_Spawn_Delay, bShowLim
End Sub

'#####################################################################################

Private Sub txtBulletDamage_KeyPress(KeyAscii As Integer)

If KeyAscii <> 8 Then
    If Not IsNumeric(Chr$(KeyAscii)) Then
        If KeyAscii = 13 Then
            'enter, accept if we can
            If txtBulletDamage.ForeColor = vbBlue Then
                modStickGame.sv_Damage_Factor = val(txtBulletDamage.Text) / 100
                
                txtBulletDamage.ForeColor = MGreen
                lblBulletDamage.Caption = "Bullet Damage Set"
                
                modDisplay.ShowBalloonTip txtBulletDamage, "Damage Multiple Set", _
                    "Damage set to " & CStr(modStickGame.sv_Damage_Factor * 100) & "%"
                
                UpdateServerVars
            End If
        ElseIf KeyAscii = 46 Then '46=asc(".")
            Exit Sub
        End If
        
        
        KeyAscii = 0
    End If
End If

End Sub

Private Sub txtBulletDamage_Change()
Dim sDamage As String, sgDamage As Single
Const Damage_Cap = "Damage Multiple"

If Me.Visible Then
    sDamage = txtBulletDamage.Text
    
    If LenB(sDamage) Then
        If Right$(sDamage, 1) <> "%" Then
            txtBulletDamage.Text = txtBulletDamage.Text & "%"
        End If
        
        On Error GoTo EH
        sgDamage = val(sDamage)
        
        If sgDamage >= Min_Damage Then
            If sgDamage <= Max_Damage Then
                
                txtBulletDamage.ForeColor = vbBlue
                lblBulletDamage.Caption = "Press Enter to Set Damage"
                
                modDisplay.ShowBalloonTip txtBulletDamage, "Press Enter!", _
                    "Press enter to set the Damage to " & CStr(sgDamage) & "%"
                
            Else
                txtBulletDamage.ForeColor = vbRed
                lblBulletDamage.Caption = Damage_Cap
                
                bShowLim = True
                
                modDisplay.ShowBalloonTip txtBulletDamage, "You crazy person", _
                    "The damage is too large. It must be less than " & CStr(Max_Damage + 1) & "%", TTI_ERROR
                
            End If
        Else
            txtBulletDamage.ForeColor = vbRed
            lblBulletDamage.Caption = Damage_Cap
            
            bShowLim = True
            
            modDisplay.ShowBalloonTip txtBulletDamage, "You crazy person", _
                "The damage is too small. It must be greater than or equal to " & Min_Damage & "%", TTI_ERROR
            
        End If
    Else
        txtBulletDamage.ForeColor = vbRed
        lblBulletDamage.Caption = Damage_Cap
    End If
    
    txtbulletdamage_click
End If

EH:
End Sub

Private Sub txtbulletdamage_click()
SetStatus Min_Damage & "% <= Damage <= " & Max_Damage & "%", bShowLim
End Sub

'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'PERK SETTINGS

Private Sub cboSpyStick_Change()
If LenB(cboSpyStick.Text) Then
    cmdPerkApply.Enabled = True
Else
    cmdPerkApply.Enabled = False
End If
End Sub

Private Sub cboSpyStick_Click()
cboSpyStick_Change
End Sub

Private Sub cboSpyStick_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdPerkApply_Click()
Dim i As Integer

cmdPerkApply.Enabled = False

If Stick(0).Perk = pSleightOfHand Then
    If modAudio.bDXSoundInited Then
        For i = 0 To eWeaponTypes.Knife - 1
            modDXSound.SetRelativeFrequency CInt(i + eWeaponTypes.Chopper + 1), Stick(0).sgTimeZone
        Next i
    End If
ElseIf Stick(0).Perk = pZombie Then
    frmStickGame.MakeZombie 0, False
End If

For i = optnPerk.LBound To optnPerk.UBound
    If optnPerk(i).Value Then
        Stick(0).Perk = i
        Exit For
    End If
Next i

If Stick(0).Perk = pSpy Then
    For i = 1 To NumSticks - 1
        If Trim$(Stick(i).Name) = cboSpyStick.Text Then
            Stick(0).MaskID = Stick(i).ID
            Exit For
        End If
    Next i
ElseIf Stick(0).Perk = pSleightOfHand Then
    If modAudio.bDXSoundInited Then
        For i = 0 To eWeaponTypes.Knife - 1
            modDXSound.SetRelativeFrequency CInt(i + eWeaponTypes.Chopper + 1), Stick(0).sgTimeZone * modStickGame.SleightOfHandReloadDecrease
        Next i
    End If
ElseIf Stick(0).Perk = pZombie Then
    frmStickGame.MakeZombie 0
End If

End Sub

Private Sub chkLaserSight_Click()

If Stick(0).WeaponType <> Chopper Then
    modStickGame.cg_LaserSight = CBool(chkLaserSight.Value)
Else
    modStickGame.cg_LaserSight = False
    chkLaserSight.Value = 0
    'lblStatus.Caption = "Can't have laser sight on chopper"
End If

End Sub

Private Sub chkShh_Click()

If frmStickGame.WeaponSilencable(Stick(0).WeaponType) Then
    Stick(0).bSilenced = CBool(chkShh.Value)
Else
    SetStatus "Error - Weapon Not Silencable"
    chkShh.Value = 0
End If

End Sub

Private Sub optnPerk_Click(Index As Integer)
Dim i As Integer

If Index = eStickPerks.pSpy Then
    cmdPerkApply.Enabled = False
    cboSpyStick.Enabled = True
    lblStick.Enabled = True
    
    cboSpyStick.Clear
    For i = 1 To NumSticks - 1
        cboSpyStick.AddItem Trim$(Stick(i).Name)
    Next i
    
    'picSpy.Visible = True
Else
    cboSpyStick.Enabled = False
    lblStick.Enabled = False
    cmdPerkApply.Enabled = (Index <> Stick(0).Perk)
    
    'picSpy.Visible = False
End If

cmdPerkApply.Default = cmdPerkApply.Enabled

End Sub

Private Sub optnTeam_Click(Index As Integer)
'Dim bMoveCamera As Boolean
Static bTold As Boolean

If Me.Visible = False Then Exit Sub


If Stick(0).Team = Spec And (Index <> Spec) Then
    frmStickGame.RandomizeMyStickPos
Else
    If Index = Spec Then
        If Not bTold Then
            frmStickGame.AddMainMessage "Use W, A, S and D to spectate", False
            bTold = True
        End If
        
        modStickGame.cg_sZoom = 1
        'modStickGame.cg_sCamera.X = 0
        'modStickGame.cg_sCamera.Y = 0
        
        'bMoveCamera = True
        'can't do it here - not a spectator yet
        
        Stick(0).X = -10: Stick(0).Y = -10
    End If
End If



If Index = eTeams.Spec Then 'Or Stick(0).bAlive = False Then
    frmStickGame.HideCursor False
    
    'sldrSpecSpeed.Enabled = True
    'lblSpecSpeed.Enabled = True
    
    If modStickGame.sv_GameType = gDeathMatch Then
        'prevent attempted respawn
        Stick(0).bAlive = True
    End If
Else
    If frmStickGame.StickInGame(0) = False Then
        modStickGame.cg_sZoom = 1
        Stick(0).bAlive = False
        If Stick(0).Team = Spec Then 'if we were spectating, reset the spawn count timer
            Stick(0).lDeathTime = GetTickCount()
        End If
        frmStickGame.HideCursor True
    End If
    
    'sldrSpecSpeed.Enabled = False
    'lblSpecSpeed.Enabled = False
End If
Stick(0).Team = Index

'If bMoveCamera Then
'    'frmStickGame.MoveCameraX modStickGame.cg_sCamera.X
'    'frmStickGame.MoveCameraY modStickGame.cg_sCamera.Y
'    frmStickGame.CentreCameraOnPoint CSng(modStickGame.cg_sCamera.X), CSng(modStickGame.cg_sCamera.Y)
'End If

End Sub

Private Sub sldrSpecSpeed_Change()
sldrSpecSpeed_Click
End Sub

Private Sub sldrSpecSpeed_Click()
modStickGame.cl_SpecSpeed = sldrSpecSpeed.Value / 10
lblSpecSpeed.Caption = "Spectator Speed - " & CStr(modStickGame.cl_SpecSpeed)
End Sub

Private Sub sldrSpecSpeed_Scroll()
sldrSpecSpeed_Click
End Sub

Private Sub Perk_Form_Load()

picSpy.BorderStyle = 0
LoadPerkFormStats

End Sub

Private Sub LoadPerkFormStats()
Dim i As Integer

optnPerk(Stick(0).Perk).Value = True

TurnOffToolTip sldrSpecSpeed.hWnd

'sldrSpecSpeed.Enabled = (Stick(0).Team = Spec)
'lblSpecSpeed.Enabled = sldrSpecSpeed.Enabled
chkLaserSight.Value = IIf(modStickGame.cg_LaserSight, 1, 0)
chkShh.Value = IIf(Stick(0).bSilenced, 1, 0)
sldrSpecSpeed.Value = modStickGame.cl_SpecSpeed * 10
lblSpecSpeed.Caption = "Spectator Speed - " & CStr(modStickGame.cl_SpecSpeed)
optnTeam(Stick(0).Team).Value = True

If Stick(0).Perk = pSpy Then
    i = frmStickGame.FindStick(Stick(0).MaskID)
    If i > -1 Then
        cboSpyStick.Text = Trim$(Stick(i).Name)
    End If
End If


'If Stick(0).Team <> Spec Then
'    If frmStickGame.StickInGame(0) = False Then
'        For i = 0 To 3
'            optnTeam(i).Enabled = False
'        Next i
'    End If
'End If

End Sub

'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'Map Loading

Private Sub Map_Form_Load()
Dim sDir As String, sCurrentMap As String
Dim i As Integer

sDir = Dir$(modStickGame.GetStickMapPath())

Do While LenB(sDir)
    If Right$(sDir, 3) = Map_Ext Then
        cboMaps.AddItem sDir
    End If
    
    sDir = Dir$()
Loop

sCurrentMap = modStickGame.StickMapPath
sCurrentMap = Mid$(sCurrentMap, InStrRev(sCurrentMap, "\") + 1)
psCurrentMapName = sCurrentMap

For i = 0 To cboMaps.ListCount - 1
    If cboMaps.List(i) = sCurrentMap Then
        cboMaps.ListIndex = i
        Exit For
    End If
Next i

cmdLoadMap.Enabled = False
cboMaps.Enabled = modStickGame.StickServer

End Sub

Private Sub cmdLoadMap_Click()
Const LoadStr = "Loading Map..."
Dim sMapToLoad As String, sMapPath As String
Dim i As Integer

Const lTimeOut As Long = 5000
Dim lStart As Long
Dim TempSockAddr As ptSockAddr
Dim tempStr As String

Dim arReceived() As Integer
Dim bAllReceived As Boolean


cmdLoadMap.Enabled = False
cboMaps.Enabled = False
cmdLoadMap.Caption = LoadStr


sMapToLoad = cboMaps.Text
SetStatus "Loading '" & sMapToLoad & "'..."
Me.Refresh

sMapPath = modStickGame.GetStickMapPath() & sMapToLoad

If frmStickGame.LoadMapEx(sMapPath) = False Then
    SetStatus "Error Loading Map - " & Err.Description
Else
    
    If frmStickGame.bRunning Then
        frmStickGame.ReleaseDeadSticks
    End If
    
    SetStatus "Loaded Map, Informing Sticks..."
    modStickGame.StickMapPath = sMapPath
    
    frmStickGame.SendBroadcast sNewMaps & sMapToLoad & vbSpace & CStr(modStickGame.MakeSquareNumber())
    
    
    ReDim arReceived(0 To NumSticks - 1)
    
    For i = 1 To NumSticks - 1
        If Stick(i).IsBot Then
            arReceived(i) = 1
        End If
    Next i
    
    tempStr = "0"
    lStart = GetTickCount()
    Do
        
        If Left$(tempStr, 1) = sNewMaps Then
            'reply from a client stick
            For i = 1 To NumSticks - 1
                If Stick(i).SockAddr.sin_addr = TempSockAddr.sin_addr Then
                    If Stick(i).SockAddr.sin_port = TempSockAddr.sin_port Then
                        'it's stick #i
                        
                        Select Case Mid$(tempStr, 2)
                            Case "1"
                                'accepted + loaded map
                                arReceived(i) = 1
                            Case Else
                                'couldn't load map/error loading map/map not found
                                arReceived(i) = 2
                        End Select
                                
                        Exit For
                    End If
                End If
            Next i
        ElseIf LenB(tempStr) = 0 Then
            'no replies, re-broadcast
            frmStickGame.SendBroadcast sNewMaps & sMapToLoad & vbSpace & CStr(modStickGame.MakeSquareNumber())
        End If
        
        
        bAllReceived = True
        For i = 1 To NumSticks - 1
            If arReceived(i) = 0 Then
                bAllReceived = False
                Exit For
            End If
        Next i
        
        
        Pause 10
        tempStr = modWinsock.ReceivePacket(frmStickGame.lSocket, TempSockAddr)
        
    Loop Until bAllReceived Or (lStart + lTimeOut < GetTickCount())
    
    
    If Not bAllReceived Then
        'some sticks need removing
        
        i = NumSticks - 1
        Do While i > 0
            If arReceived(i) <> 1 Then
                'no reply received or they couldn't load the map
                '==> remove stick
                
                frmStickGame.RemoveStick i
            End If
            
            i = i - 1 'step backwards so remove a stick doesn't effect arReceived
        Loop
        
        frmStickGame.AddMainMessage "Some Sticks were removed - Errors loading the new map", False
    End If
    
'    i = 1
'    Do While i < NumSticks
'
'        bOK = False
'        lStart = 0
'        lLastSend = 0
'
'        If Stick(i).IsBot = False Then
'            lStart = GetTickCount()
'            SetStatus "Contacting " & Trim$(Stick(i).Name)
'            Do
'                If lLastSend + lSendDelay < GetTickCount() Then
'                    modWinsock.SendPacket frmStickGame.lSocket, Stick(i).SockAddr, sNewMaps & sMapToLoad & vbSpace & CStr(modStickGame.MakeSquareNumber())
'                End If
'
'                Pause 10
'
'                tempStr = modWinsock.ReceivePacket(frmStickGame.lSocket, tempSockAddr)
'
'                If tempStr = sNewMaps Then
'                    bOK = True
'                    Exit Do
'                End If
'            Loop While lStart + lTimeOut > GetTickCount()
'
'
'
'            If Not bOK Then
'                SetStatus "Error Contacting " & Trim$(Stick(i).Name)
'
'                'kick + remove
'                modWinsock.SendPacket frmStickGame.lSocket, Stick(i).SockAddr, sKicks & "No Response to New Map"
'                frmStickGame.RemoveStick i
'                i = i - 1
'            Else
'                SetStatus "Contacted " & Trim$(Stick(i).Name)
'                Stick(i).LastPacket = GetTickCount() + 5000 'prevent lagging out
'            End If
'        End If
'
'
'        i = i + 1
'    Loop
    
    Erase arReceived
    
    frmStickGame.RandomizeMyStickPos
    'Stick(0).JumpStartY = StickGameHeight + 10
    frmStickGame.AddMainMessage "Loaded Map '" & sMapToLoad & "'", False
    Unload Me
End If


End Sub

Private Sub cboMaps_Change()
cmdLoadMap.Enabled = (cboMaps.Text <> psCurrentMapName)
End Sub
Private Sub cboMaps_Click()
cboMaps_Change
End Sub
Private Sub cboMaps_Scroll()
cboMaps_Change
End Sub

'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'GRAPAHICS SETTINGS

Private Sub chkHolstered_Click()
modStickGame.cg_HolsteredWeap = CBool(chkHolstered.Value)
End Sub

Private Sub chkEnableSound_Click()
modAudio.bDXSoundEnabled = CBool(chkEnableSound.Value)

lblVol.Enabled = modAudio.bDXSoundEnabled
sldrVol.Enabled = modAudio.bDXSoundEnabled

If Me.Visible Then
    If modAudio.bDXSoundEnabled Then
        frmStickGame.StickGameSpeedChanged modStickGame.sv_StickGameSpeed, modStickGame.sv_StickGameSpeed
    End If
End If

End Sub

Private Sub sldrVol_Change()

On Error GoTo EH
frmStickGame.SetDXSoundVol sldrVol.Value

EH:
End Sub

'###############################################################

Private Sub chkFPS_Click()
modStickGame.cg_DrawFPS = CBool(chkFPS.Value)
End Sub

Private Sub chkBlood_Click()
modStickGame.cg_Blood = CBool(chkBlood.Value)
End Sub

Private Sub chkCasing_Click()
modStickGame.cg_Casing = CBool(chkCasing.Value)
End Sub

Private Sub chkDead_Click()
modStickGame.cg_DeadSticks = CBool(chkDead.Value)
If modStickGame.cg_DeadSticks = False Then
    frmStickGame.EraseDeadSticks
End If
End Sub

Private Sub chkInvert_Click()

If chkInvert.Value Then
    modStickGame.cg_DisplayMode = modStickGame.cg_DisplayMode_Invert
Else
    modStickGame.cg_DisplayMode = modStickGame.cg_DisplayMode_Normal
End If

End Sub

Private Sub chkMagazines_Click()
modStickGame.cg_Magazines = CBool(chkMagazines.Value)
End Sub

Private Sub chkSimple_Click()
modStickGame.cg_SimpleStaticWeapons = CBool(chkSimple.Value)
End Sub

Private Sub chkSmoke_Click()
modStickGame.cg_Smoke = CBool(chkSmoke.Value)

If Me.Visible Then
    If Not modStickGame.cg_Smoke Then frmStickGame.EraseSmoke
End If

End Sub

Private Sub chkSniperScope_Click()
modStickGame.cl_SniperScope = CBool(chkSniperScope.Value)
End Sub

Private Sub chkSparks_Click()
modStickGame.cg_Sparks = CBool(chkSparks.Value)
End Sub

Private Sub chkWallMarks_Click()
modStickGame.cg_WallMarks = CBool(chkWallMarks.Value)

If Me.Visible Then
    If Not modStickGame.cg_WallMarks Then frmStickGame.EraseWallMarks
End If

End Sub

'##############################################################################

Private Sub cmdNameApply_Click()
Dim NameOp As String, LName As String
Dim PrevName As String

Stick(0).colour = picColour.BackColor

LName = txtName.Text
NameOp = Trim$(Replace$(LName, "@", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, "#", vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modMessaging.MsgEncryptionFlag, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, ":", vbNullString, , , vbTextCompare)) 'for chat in game
NameOp = Trim$(Replace$(NameOp, modSpaceGame.mPacketSep, vbNullString, , , vbTextCompare))
NameOp = Trim$(Replace$(NameOp, modSpaceGame.UpdatePacketSep, vbNullString, , , vbTextCompare))

PrevName = Trim$(Stick(0).Name)
Stick(0).Name = Trim$(NameOp)
txtName.Text = Trim$(Stick(0).Name)

cmdCancel_Click

If Trim$(Stick(0).Name) <> LName Then
    modDisplay.ShowBalloonTip txtName, "Tsk tsk", "Certain characters aren't allowed and have been removed"
Else
    modDisplay.ShowBalloonTip txtName, "Success", "Name and other stuff have been set"
    
    frmStickGame.SendChatPacket PrevName & " renamed to " & Trim$(Stick(0).Name), Stick(0).colour
End If

End Sub

Private Sub cmdCancel_Click()
Dim RGBCol As ptRGB

RGBCol = RGBDecode(Stick(0).colour)

sldrCol(0).Value = RGBCol.Red
sldrCol(1).Value = RGBCol.Green
sldrCol(2).Value = RGBCol.Blue

picColour.BackColor = Stick(0).colour

txtName.Text = Trim$(Stick(0).Name)

cmdNameApply.Enabled = False
cmdCancel.Enabled = False

modDisplay.ShowBalloonTip txtName, "Reset", "Name and other stuff have been reset"

End Sub

Private Sub Graphics_Form_Load()
Dim i As Integer

chkCasing.Value = Abs(modStickGame.cg_Casing)
'chkExplosions.Value = Abs(modStickGame.cg_Explosions)
chkBlood.Value = Abs(modStickGame.cg_Blood)
'chkRPGFlame.Value = Abs(modStickGame.cg_RPGFlame)
chkFPS.Value = Abs(modStickGame.cg_DrawFPS)
chkDead.Value = Abs(modStickGame.cg_DeadSticks)
chkMagazines.Value = Abs(modStickGame.cg_Magazines)
chkSparks.Value = Abs(modStickGame.cg_Sparks)
'chkBlackChopper.Value = Abs(modStickGame.cg_ChopperCol = vbBlack)
chkSmoke.Value = Abs(modStickGame.cg_Smoke)
chkSimple.Value = Abs(modStickGame.cg_SimpleStaticWeapons)
chkWallMarks.Value = Abs(modStickGame.cg_WallMarks)
chkSniperScope.Value = Abs(modStickGame.cl_SniperScope)
chkHolstered.Value = Abs(modStickGame.cg_HolsteredWeap)
chkTrails.Value = Abs(modStickGame.cg_ShowBulletTrails)


chkEnableSound.Enabled = modAudio.bDXSoundInited
chkEnableSound.Value = Abs(modAudio.bDXSoundEnabled)
TurnOffToolTip sldrVol.hWnd
If modAudio.bDXSoundInited Then
    On Error Resume Next
    sldrVol.Value = modDXSound.GetVolume(0)
Else
    lblVol.Enabled = modAudio.bDXSoundEnabled
    sldrVol.Enabled = modAudio.bDXSoundEnabled
End If


If modStickGame.cg_DisplayMode = modStickGame.cg_DisplayMode_Invert Then
    chkInvert.Value = 1
End If

cmdCancel_Click
lblColourInfo.Caption = "R" & vbNewLine & "G" & vbNewLine & "B"
'For i = sldrCol.LBound To sldrCol.UBound
'    TurnOffToolTip sldrCol(i).hWnd
'Next i
picColour.BorderStyle = 0

End Sub

Private Sub picBG_Click(Index As Integer)
Dim lCol As Long

lCol = picBG(Index).BackColor

frmStickGame.picMain.BackColor = lCol
frmStickGame.BackColor = lCol
modStickGame.cg_BGColour = lCol

frmStickGame.BackgroundColourChanged
End Sub

'name + colour stuff
Private Sub sldrCol_Change(Index As Integer)

picColour.BackColor = RGB(sldrCol(0).Value, sldrCol(1).Value, sldrCol(2).Value)

cmdCancel.Enabled = True
cmdNameApply.Enabled = CBool(LenB(txtName.Text))
End Sub

Private Sub sldrCol_Scroll(Index As Integer)
sldrCol_Change Index
End Sub

Private Sub txtName_Change()
cmdNameApply.Enabled = CBool(LenB(txtName.Text))
cmdCancel.Enabled = True
End Sub

'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'#########################################################################################################
'frmClient - Kick + List

Private Sub SetStatus(ByVal T As String, Optional bError As Boolean = False)
Const K As String = "Status: "
lblStatus.Caption = K & T

lblStatus.ForeColor = IIf(bError, vbRed, vbBlue)

lblStatus.Refresh
End Sub

Private Sub cmdKick_Click()
Dim ID As Integer, Sticki As Integer
Dim Txt As String

cmdKick.Enabled = False

Txt = lstMain.Text

ID = CInt(Right$(Txt, Len(Txt) - InStrRev(Txt, Space$(1), , vbTextCompare)))

If LenB(Txt) Then
    On Error GoTo EH
    Sticki = frmStickGame.FindStick(ID)
    
    If Sticki <> -1 Then
        If Stick(Sticki).IsBot Then
            'SetStatus "Error - Can't kick a bot"
            frmStickGame.RemoveBot Sticki
        ElseIf Sticki = 0 Then
            SetStatus "Error - Can't kick self"
        Else
            modWinsock.SendPacket frmStickGame.lSocket, Stick(Sticki).SockAddr, sKicks & "Server Decision"
        End If
    Else
        SetStatus "Error Finding Stick"
    End If
End If

EH:
End Sub

Private Sub lstMain_Click()
cmdKick.Enabled = (LenB(lstMain.Text) And modStickGame.StickServer And (Right$(lstMain.Text, 1) <> "0"))
End Sub

Private Sub Client_Form_Load()

cmdKick.Enabled = False
lstMain.Enabled = modStickGame.StickServer
tmrRefresh_Timer

End Sub

Private Sub tmrRefresh_Timer()
Dim i As Integer, iSelected As Integer

iSelected = lstMain.ListIndex
lstMain.Clear

For i = 0 To modStickGame.NumSticks - 1
    lstMain.AddItem "Name: " & Trim$(modStickGame.Stick(i).Name) & Space$(3) _
        & "ID: " & modStickGame.Stick(i).ID
Next i

On Error Resume Next
lstMain.ListIndex = iSelected

End Sub
