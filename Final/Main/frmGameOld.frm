VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Window"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStart 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Integer
    Y As Integer
End Type

Public bRunning As Boolean

Private Const TankWidth As Integer = 1000
Private Const CannonLen As Integer = TankWidth / 1.2
Private Const MaxFPS As Integer = 5
Private Const Sens As Integer = 4
Private Const Accel As Byte = 1
Private Const MaxSpeed As Byte = 50
Private Const BulletLen As Single = 5
Private Const BulletWidth As Integer = 75
Private Const BulletSpeed As Integer = 60
Private Const BulletLife As Integer = 75

Private Const pi As Single = 3.141593 '3.1415926535
Private Const Rad As Single = pi / 180

Public UpKey As Boolean, DownKey As Boolean, _
    LeftKey As Boolean, RightKey As Boolean, _
    FireKey As Boolean


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyW
        UpKey = True
    Case vbKeyA
        LeftKey = True
    Case vbKeyS
        DownKey = True
    Case vbKeyD
        RightKey = True
    Case 32
        FireKey = True
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyW
        UpKey = False
    Case vbKeyA
        LeftKey = False
    Case vbKeyS
        DownKey = False
    Case vbKeyD
        RightKey = False
    Case 32
        FireKey = False
End Select
End Sub

Private Sub Form_Load()
Call FormLoad(Me)

ReDim OtherTanks(1)

Me.FillStyle = 0 'set the fill style this is defaulted to 1 transparent
Me.DrawStyle = 0 'the style of the border.  You can find out the numbers by looking at the objects properties
Me.DrawWidth = 1 'the sets the width of the border

MyTank.x = Width / 2
MyTank.Y = Height / 2
MyTank.Facing = 0
MyTank.Colour = TxtForeGround

modGame.GameFormLoaded = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
bRunning = False
modGame.GameFormLoaded = False
End Sub

Private Function D2R(ByVal Degrees As Single) As Single

D2R = Degrees * Rad

End Function

Private Function R2D(ByVal Rads As Single) As Single

R2D = Rads / Rad

End Function

Private Sub DrawTank(ByRef Tank As ptTank, Optional ByVal LabelPts As Boolean = False)

Dim mFacing As Single
Dim Pt(1 To 4) As POINTAPI
Dim PtCannon As POINTAPI

Dim Centre As POINTAPI

'On Error Resume Next
With Tank
    
    mFacing = .Facing
    Centre.x = Tank.x
    Centre.Y = Tank.Y
    
    
    Pt(4).x = Centre.x + (TankWidth) * Sin(D2R(mFacing - 45)) / 2
    Pt(4).Y = Centre.Y + (TankWidth) * Cos(D2R(mFacing - 45)) / 2
    
    Pt(3).x = Centre.x + (TankWidth) * Sin(D2R(mFacing + 45)) / 2
    Pt(3).Y = Centre.Y + (TankWidth) * Cos(D2R(mFacing + 45)) / 2
    
    Pt(2).x = Centre.x + (TankWidth) * Sin(D2R(mFacing + 155)) / 2
    Pt(2).Y = Centre.Y + (TankWidth) * Cos(D2R(mFacing + 155)) / 2
    
    Pt(1).x = Centre.x + (TankWidth) * Sin(D2R(mFacing + 205)) / 2
    Pt(1).Y = Centre.Y + (TankWidth) * Cos(D2R(mFacing + 205)) / 2
    
    PtCannon.x = Centre.x + -(CannonLen) * Sin(D2R(mFacing))
    PtCannon.Y = Centre.Y + -(CannonLen) * Cos(D2R(mFacing))
    
End With

Me.FillColor = Tank.Colour 'set the colour to be filled.  I have made it a bit random
Me.ForeColor = Tank.Colour 'this sets the color of the border

'Polygon Me.hdc, Pt(1), 4 'call the polygon function

Line (Pt(1).x, Pt(1).Y)-(Pt(2).x, Pt(2).Y), Tank.Colour
Line (Pt(2).x, Pt(2).Y)-(Pt(3).x, Pt(3).Y), Tank.Colour
Line (Pt(3).x, Pt(3).Y)-(Pt(4).x, Pt(4).Y), Tank.Colour
Line (Pt(4).x, Pt(4).Y)-(Pt(1).x, Pt(1).Y), Tank.Colour
Line (Centre.x, Centre.Y)-(PtCannon.x, PtCannon.Y), Tank.Colour


If LabelPts Then
    CurrentX = Pt(1).x
    CurrentY = Pt(1).Y
    Print "1"
    CurrentX = Pt(2).x
    CurrentY = Pt(2).Y
    Print "2"
    CurrentX = Pt(3).x
    CurrentY = Pt(3).Y
    Print "3"
    CurrentX = Pt(4).x
    CurrentY = Pt(4).Y
    Print "4"
End If

End Sub

Private Sub MainLoop()

bRunning = True
Do While bRunning
    
    If LimitFPS() Then
        Call DoFrame
    End If
    
    DoEvents
Loop

End Sub

Private Sub FixAngle(ByRef Angle As Integer)
Dim bFixed As Boolean

Do Until bFixed
    If Angle > 359 Then
        Angle = Angle - 360
    ElseIf Angle < 0 Then
        Angle = Angle + 360
    Else
        bFixed = True
    End If
Loop

End Sub

Private Sub ApplySpeed(Tank As ptTank)

With Tank
    
    If .Accelerating = 1 Then
        .Speed = .Speed + Accel
    ElseIf .Accelerating = -1 Then
        .Speed = .Speed - Accel
    End If
    
    If Abs(.Speed) > MaxSpeed Then .Speed = Sgn(.Speed) * MaxSpeed
    
    .x = .x + .Speed * Sin(D2R(-.Facing))
    .Y = .Y - .Speed * Cos(D2R(-.Facing))
    
    'Print "XCOMP: " & XComp & "  X: " & .X
    'Print "YCOMP: " & YComp & "  Y: " & .Y
    
End With

End Sub

Private Sub ApplyBulletSpeed(Bullet As ptBullet)

With Bullet
    
    
    .x = .x + .Speed * Sin(D2R(-.Facing))
    .Y = .Y - .Speed * Cos(D2R(-.Facing))
    
    'Print "XCOMP: " & XComp & "  X: " & .X
    'Print "YCOMP: " & YComp & "  Y: " & .Y
    
End With

End Sub

Private Sub DrawBullet(Bullet As ptBullet)
Dim pX As Single, pY As Single
'Dim SideX1 As Single, SideY1 As Single
'Dim SideX2 As Single, SideY2 As Single

If Bullet.Active Then
    With Bullet
        
        Me.FillColor = vbGreen
        Me.FillStyle = 0
        Circle (.x, .Y), BulletWidth, vbGreen
        Me.FillStyle = 1
        
        pX = .x + .Speed * Sin(D2R(.Facing)) * BulletLen
        pY = .Y + .Speed * Cos(D2R(.Facing)) * BulletLen
        
        'SideX1 = .X + .Speed * Sin(D2R(.Facing - 90))
        'SideY1 = .Y + .Speed * Cos(D2R(.Facing - 90))
        
        'SideX2 = .X + .Speed * Sin(D2R(.Facing + 90))
        'SideY2 = .Y + .Speed * Cos(D2R(.Facing + 90))
        
        'Line (SideX1, SideY1)-(pX, pY), vbBlack
        'Line (SideX2, SideY2)-(pX, pY), vbBlack
        Line (.x, .Y)-(pX, pY), vbBlue
        
    End With
End If

End Sub

Private Function LimitFPS() As Boolean
Static LastTick As Long
Dim NewTick As Long

NewTick = GetTickCount()

If NewTick > (LastTick + MaxFPS) Then
    LimitFPS = True
    LastTick = NewTick
End If

End Function

Private Sub ProcessKeys()
If LeftKey Then
    MyTank.Facing = MyTank.Facing + Sens
End If
If RightKey Then
    MyTank.Facing = MyTank.Facing - Sens
End If

If UpKey Then
    MyTank.Accelerating = 1
ElseIf DownKey Then
    MyTank.Accelerating = -1
Else
    MyTank.Accelerating = 0
End If

MyTank.Shooting = FireKey

End Sub

Private Sub CreateBullet(Tank As ptTank)

Dim XC As Single, YC As Single

With Tank.Bullet
    If .Active Then Exit Sub
    
    .LifeLeft = BulletLife
    .x = Tank.x
    .Y = Tank.Y
    .Speed = Tank.Speed + BulletSpeed
    .Facing = Tank.Facing
    .Active = True
End With

End Sub

Private Sub CheckCollisions(Tank As ptTank)

With Tank
    If .x < 0 Then
        .Speed = 0
        .x = 0 'TankWidth
    ElseIf .x > Me.Width Then
        .Speed = 0
        .x = Me.Width ' - TankWidth
    End If
    
    If .Y < 0 Then
        .Speed = 0
        .Y = 0 'TankWidth
    ElseIf .Y > Me.Height - 500 Then
        .Speed = 0
        .Y = Me.Height - 500 ' - TankWidth
    End If
End With

End Sub

Private Sub tmrStart_Timer()
tmrStart.Enabled = False
Call MainLoop
End Sub

Private Sub DoFrame()
Dim i As Integer

Me.Cls

Call DoTank(MyTank)

For i = LBound(OtherTanks) To UBound(OtherTanks)
    Call DoTank(OtherTanks(i))
Next i

Call ProcessKeys

End Sub

Private Sub DoTank(Tank As ptTank)

Call FixAngle(Tank.Facing)

If Tank.Shooting And Not Tank.Bullet.Active Then
    CreateBullet Tank
End If

If Tank.Bullet.Active Then
    ApplyBulletSpeed Tank.Bullet
    Tank.Bullet.LifeLeft = Tank.Bullet.LifeLeft - 1
    If Tank.Bullet.LifeLeft < 0 Then Tank.Bullet.Active = False
    DrawBullet Tank.Bullet
End If

ApplySpeed Tank
CheckCollisions Tank
DrawTank Tank ', True

End Sub

'above = non winsock ------------------------------------------------------------------

