Attribute VB_Name = "modSpaceGame"
Option Explicit

'Private Const EXTRA_PRECISION As Boolean = False

Public Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'constants
Public Const Pi As Single = 3.14159265358979 'given by 4*atn(1)

Public Const Pi2 = Pi * 2 'calculation constants
Public Const piD2 = Pi / 2
Public Const pi2d3 = Pi2 / 3
Public Const pi3D4 = 3 * Pi / 4
Public Const piD4 = Pi / 4
Public Const pi5D4 = Pi * 5 / 4
Public Const pi3D5 = Pi * 3 / 5
Public Const pi2D5 = Pi * 2 / 5
Public Const piD3 = Pi / 3
Public Const piD6 = Pi / 6
Public Const piD10 = Pi / 10
Public Const pi3D2 = 3 * Pi / 2
Public Const piD8 = Pi / 8
Public Const pi3D8 = 3 * Pi / 8
Public Const piD20 = Pi / 20 'chopper spray angle * 2
Public Const piD40 = Pi / 40 '^
Public Const pi8D9 = 8 * Pi / 9
Public Const pi7D4 = 7 * Pi / 4
Public Const pi13D18 = 13 * Pi / 18 '130 deg (deagle)
Public Const pi4D9 = 4 * Pi / 9 '80 deg (sniper reload)
Public Const pi5D9 = 5 * Pi / 9  '100 deg (sniper reload)
Public Const pi7D12 = 7 * Pi / 12 'chopper facing
Public Const pi5D12 = 5 * Pi / 12 '   "      "
Public Const pi5D8 = Pi * 5 / 8
Public Const pi17D16 = Pi * 17 / 16
Public Const piD18 = Pi / 18
Public Const k2D3 = 2 / 3, k4D3 = 4 / 3
Public Const piD16 = Pi / 16
Public Const piD5 = Pi / 5

Public Const Default_AI_Sample_Rate = 150

Public Const SHIELD_START = 50    'Amount of shields at start
Public Const Hull_Start = 100

Public Const UpdatePacketSep = "®"
Public Const mPacketSep = "©"

Public Const sForceTeams As String * 1 = "F"
Public Const sKicks As String * 1 = "K"

Public Const CentreX = 5500 'for text printing etc
Public Const CentreY = 4000

'editing
Public SpaceEditing As Boolean
Public R_ob1 As ptSquare
Public R_ln1 As ptSquare
Public R_ln2 As ptSquare
Public Const EditZoom = 0.58

Public Type ptSquare
    Left As Single
    Top As Single
    width As Single
    height As Single
End Type

'Public Kills As Integer
'Public Deaths As Integer
Public UseAI As Boolean
'Public Sound As Boolean

'client vars
Public cg_StarBG As Boolean
Public cg_Stars3D As Boolean
Public cg_ShowFPS As Boolean
Public cg_DrawThick As Boolean
Public cl_UseMouse As Boolean
'Public cg_BlackBG As Boolean
Public cg_PredatorCrossHair As Boolean
Public cg_DrawLeadCrossHair As Boolean
Public cg_SpaceMainCrosshair As Long, cg_SpaceLeadCrosshair As Long
Public cg_DrawExplosions As Boolean
Public cg_CrossHairWidth As Integer
Public cg_Cls As Boolean
Public cg_Smoke As Boolean
Public cg_GunSmoke As Boolean
Public cg_BulletSmoke As Boolean
Public cg_ShowMap As Boolean
Public cg_ShowMissileLock As Boolean
Public cg_SpaceDisplayMode As Long '=vbNotSrcCopy or vbSrcCopy

Public cg_MapLen As Integer
Private Const DefaultMapLen = 1500

'camera pos
Public cg_Camera As PointAPI
Public cg_Zoom As Single

'server vars
Public sv_GameSpeed As Single
Public sv_BotAI As Boolean
Public sv_BulletsCollide As Boolean
Public sv_AddBulletVectorToShip As Boolean 'bullets push ships
Public sv_ClipMissiles As Boolean 'missiles hit bullets
Public sv_Bullet_Damage As Single
Public sv_BulletWallBounce As Boolean
Public sv_GameType As eGameTypes
Public sv_CTFTime As Integer 'time required for a successfull flag capture. NOTE: IS IN SECONDS
Public sv_ScoreReq As Integer

'frame stuff
Public SpaceElapsedTime As Long
Public Const Space_Ms_Delay = 25
Private Const Space_Required_FPS = 100
Public Const Space_Ms_Required_Delay = 1000 / Space_Required_FPS
Public TimeFactor As Single
'Milliseconds per frame (25 = 40 frames per second)
'fps = 1/(Delay*10^-3)
'Delay = 10^3/FPS

'Public MaxWidth As Long
'Public MaxHeight As Long

Public Enum eGameTypes
    DM = 0
    CTF = 1
    Elimination = 2
End Enum

'Public BotIDs() As Integer
'Public NumBotIDs As Integer

Public SpaceServer As Boolean
Public SpaceServerIP As String

'loaded forms
Public GameFormLoaded As Boolean
Public GameOptionFormLoaded As Boolean
Public GameClientsFormLoaded As Boolean
Public GameBotFormLoaded As Boolean
Public GameCrosshairFormLoaded As Boolean
Public GameClientSettingsFormLoaded As Boolean


'Player state flags
Public Enum ePlayerState
    Player_None = 0       'No motion
    PLAYER_THRUST = 1     'Thrusting forward
    PLAYER_REVTHRUST = 2  'Thrusting reverse
    PLAYER_LEFT = 4       'Rotating left
    PLAYER_RIGHT = 8      'Rotating right
    PLAYER_FIRE = 16      'Player firing
    Player_Secondary = 32
    Player_StrafeLeft = 64
    Player_StrafeRight = 128
    Player_Shielding = 256
End Enum

Public Type PLAYERTYPE
    ID As Integer        'Number to identify player over net
    LastPacket As Long   'When did we last receive a mPacket for this player?
    LastPacketID As Long 'ID value of last mPacket received (so we can ignore packets received out of sequence)
    Name As String * 20  'Player's name
    Facing As Single     'Angle the ship is Facing (ok, ok, it's a triangle, not a ship! Shut it!)
    Heading As Single    'Current direction in which ship is ovString
    Speed As Single      'Current speed with which ship is ovString
    X As Single          'Current X coordinate of ship within form
    Y As Single          'Current Y coordinate of ship within form
    LastBullet As Long   'When was the last time the player fired?
    Shields As Single    'The player's shields
    Hull As Single       'player's hull
    MaxShields As Single 'The player's ax shields
    MaxHull As Single
    bDrawShields As Boolean   'Are the shields to be displayed?
    bRightBullet As Boolean
    state As Integer     'Player state flags
    ptSockAddr As ptSockAddr 'Player's ptsockaddr
    colour As Long
    ShipType As eShipTypes 'As Byte
    
    Kills As Integer
    Deaths As Integer
    IsBot As Boolean
    
    LastSecondary As Long
    
    Team As eTeams
    
    AITimer As Long
    LastAITargetIndex As Integer
    'AIWantToFace As Single
    
    Score As Integer '=k-d
    
    Alive As Boolean
    
    LastSmoke As Long
    
    MissileLocki As Integer
End Type

Public Player() As PLAYERTYPE  'A nice little array of players

Public Enum eShipTypes
    Raptor = 0
    Behemoth = 1
    Hornet = 2
    MotherShip = 3
    Wraith = 4
    Infiltrator = 5
    SD = 6
End Enum
Public Enum eTeams
    Neutral = 0
    Red = 1
    Blue = 2
    Spec = 3
End Enum


'Text drawing API
'Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long


'Line drawing API
'Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, _
'    ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
'
'Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, _
'    ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
'
''Some text drawing API
'Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
'Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, _
'    ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

'Public Sub ShowText(strText As String, ByVal intX As Integer, ByVal intY As Integer, _
'            ByVal lngColour As Long, lngDC As Long, Optional intOpaque As Integer)
'
''This function writes text to the backbuffer
'Call SetBkColor(lngDC, 0)
'If intOpaque = 1 Then Call SetBkMode(lngDC, intOpaque)
'Call SetTextColor(lngDC, lngColour)
'Call TextOut(lngDC, intX, intY, strText, Len(strText))
'
'End Sub

'Public Sub PlayLasers()
'If modSpaceGame.Sound Then
    'modSound.PlaySound 101
'End If
'End Sub

'Public Sub PlayThrusters()
'If modSpaceGame.Sound Then
    'modSound.PlaySound 102
'End If
'End Sub
Public Enum eLobbyCmds
    Add = 1
    Remove = 2
    Refresh = 3
End Enum

Public Type ptLobbyGame
    HostName As String
    IP As String
    bStickGame As Boolean
End Type

Public CurrentGames() As ptLobbyGame


Public sGameModeMessage As String
Public Const ksGameModeMessage As String = "Server is in Game Mode - May Not Reply"

Public Sub Space_FormLoad(Frm As Form, Optional bReverse As Boolean = False)

FormLoad Frm, bReverse, False, False, True

If bReverse Then
    SetFocus2 frmGame
End If
End Sub

Public Function GetGames() As String
Dim St As String
Dim i As Integer, j As Integer

On Error GoTo EH

'check if any duplicates
i = 1
Do While i <= UBound(CurrentGames)
    
    For j = i + 1 To UBound(CurrentGames)
        If CurrentGames(i).IP = CurrentGames(j).IP Then
            RemoveGame i
            i = i - 1
            Exit For
        End If
    Next j
    
    i = i + 1
Loop


For i = 1 To UBound(CurrentGames)
    St = St & CurrentGames(i).HostName & "#" & _
              CurrentGames(i).IP & _
            IIf(CurrentGames(i).bStickGame, "S", vbNullString) & "@"
Next i

If Right$(St, 1) = "@" Then
    'prevent null entry when list is Split()'d
    St = Left$(St, Len(St) - 1)
End If

GetGames = St

Exit Function
EH:
GetGames = vbNullString
End Function

Public Sub ProcessLobbyCmd(ByVal Str As String)

Dim Games() As String
Dim cmd As eLobbyCmds
Dim i As Integer, u As Integer
Dim sTmp As String
Dim bTmp As Boolean, bCan As Boolean

On Error GoTo EH
cmd = Left$(Str, 1)
Str = Mid$(Str, 2)
u = UBound(CurrentGames) + 1

Select Case cmd
    Case eLobbyCmds.Add
        'str = Name#Ip & iif(stickgame, "S")
        
        i = InStr(1, Str, "#", vbTextCompare)
        
        ReDim Preserve CurrentGames(u)
        
        With CurrentGames(u)
            .HostName = Left$(Str, i - 1)
            
            sTmp = Mid$(Str, i + 1)
            
            If Right$(sTmp, 1) = "S" Then
                'stick game
                .IP = Left$(sTmp, Len(sTmp) - 1)
                .bStickGame = True
            Else
                .IP = sTmp
            End If
            
        End With
        
        
    Case eLobbyCmds.Refresh
        'str = Name1#IP1["S"]@Name2#Ip2["S"]...
        
        ReDim CurrentGames(0)
        
        Games = Split(Str, "@", , vbTextCompare)
        
        For i = 0 To UBound(Games)
            ProcessLobbyCmd eLobbyCmds.Add & Games(i)
        Next i
        
        
    Case eLobbyCmds.Remove
        'str = ip to remove
        
        bTmp = (Right$(Str, 1) = "S")
        If bTmp Then
            Str = Left$(Str, Len(Str) - 1)
        End If
        
        For i = 0 To u - 1
            If CurrentGames(i).bStickGame Then
                If bTmp Then 'if we are looking for a stickgame...
                    If CurrentGames(i).IP = Str Then
                        RemoveGame i
                        Exit For
                    End If
                End If
            ElseIf bTmp = False Then
                'it's not a stick game, and we're not looking for a stickgame
                If CurrentGames(i).IP = Str Then
                    RemoveGame i
                    Exit For
                End If
            End If
            
        Next i
        
End Select


EH:
End Sub

Private Sub RemoveGame(ByVal Index As Integer)

Dim i As Integer

'If there's only one left, just erase the array
If UBound(CurrentGames) = 0 Then
    ReDim CurrentGames(0)
Else
    'Remove the bullet
    For i = Index To UBound(CurrentGames) - 1
        CurrentGames(i).HostName = CurrentGames(i + 1).HostName
        CurrentGames(i).IP = CurrentGames(i + 1).IP
        CurrentGames(i).bStickGame = CurrentGames(i + 1).bStickGame
    Next i
    
    'Resize the array
    ReDim Preserve CurrentGames(UBound(CurrentGames) - 1)
End If

End Sub

Public Function CentreFill(ByVal S As String, ByVal iLen As Integer) As String
Dim ln As Integer
Dim nSpaces As Integer
Dim b As String

On Error Resume Next
ln = Len(S)

b = Space$(iLen)

Mid$(b, iLen \ 2 - ln \ 2, Len(S)) = S

CentreFill = b

End Function

'Public Function IsBotID(ID As Integer) As Boolean
''Dim i As Integer
''
''For i = 0 To NumBotIDs - 1
''    If BotIDs(i) = ID Then
''        IsBotID = True
''        Exit Function
''    End If
''Next i
'
'IsBotID = Player(frmGame.FindPlayer(ID)).IsBot
'
'End Function

Public Sub InitVars() 'called from frmMain_Load

'modSpaceGame.cg_BlackBG = True
modSpaceGame.cg_StarBG = True
modSpaceGame.cl_UseMouse = True
modSpaceGame.cg_PredatorCrossHair = True
modSpaceGame.sv_GameSpeed = 1
modSpaceGame.sv_BotAI = True
modSpaceGame.sv_AddBulletVectorToShip = True
modSpaceGame.cg_SpaceMainCrosshair = RGB(0, 255, 0)
modSpaceGame.cg_SpaceLeadCrosshair = RGB(255, 0, 0)
modSpaceGame.sv_ClipMissiles = True
modSpaceGame.cg_DrawThick = True
modSpaceGame.cg_DrawExplosions = True
'modSpaceGame.cg_DrawLeadCrossHair = True
modSpaceGame.cg_CrossHairWidth = 2
modSpaceGame.cg_Cls = True
modSpaceGame.cg_Smoke = True
modSpaceGame.cg_GunSmoke = True
modSpaceGame.cg_BulletSmoke = True
modSpaceGame.cg_ShowMap = True
modSpaceGame.cg_MapLen = modSpaceGame.DefaultMapLen
cg_Zoom = 1
modSpaceGame.sv_ScoreReq = 10

modSpaceGame.cg_SpaceDisplayMode = cg_DisplayMode_Normal

ReDim CurrentGames(0)
'modSpaceGame.sv_BulletsCollide = True
'modSpaceGame.Sound = True

End Sub

'colour
' Works properly for 24 bit colors
Public Function RGBDecode(RGBColour As Long) As ptRGB

With RGBDecode
    .Red = RGBColour And &HFF
    .Green = (RGBColour And &HFF00&) \ 256
    
    'RGBDecode.rgbtBlue = RGBcolor \ 65536
    .Blue = (RGBColour And &HFF0000) \ 65536 '<-- works with alpha blending
End With

End Function

Public Function RandomRGBColour() As Long

RandomRGBColour = RGB( _
        Int(Rnd() * 256), _
        Int(Rnd() * 256), _
        Int(Rnd() * 256))

End Function

Public Sub HostSpaceGame(ByVal IPToDist As String)

Dim Str As String

modSpaceGame.SpaceEditing = False

Str = eCommands.LobbyCmd & eLobbyCmds.Add & frmMain.LastName & "#" & IPToDist

If Server Then
    'modMessaging.DistributeMsg Str, -1
    Call DataArrival(Str) 'this'll distribute it
Else
    SendData Str
End If

modSpaceGame.SpaceServer = True
modSpaceGame.SpaceServerIP = IPToDist

'DoEvents 'refresh screen

On Error Resume Next
Load frmGame

End Sub

Public Sub JoinSpaceGame(ByVal IP As String)

modSpaceGame.SpaceEditing = False

modSpaceGame.SpaceServer = False
modSpaceGame.SpaceServerIP = IP

On Error Resume Next
Load frmGame
End Sub

Public Sub TurnOffToolTip(hWnd As Long)
SendMessageByLong hWnd, TBM_SETTOOLTIPS, 0, 0
End Sub

Public Function GetTeamStr(ByVal vTeam As eTeams) As String
Select Case vTeam
    Case eTeams.Neutral
        GetTeamStr = "Neutral"
    Case eTeams.Red
        GetTeamStr = "Red"
    Case eTeams.Blue
        GetTeamStr = "Blue"
    Case eTeams.Spec
        GetTeamStr = "Spectator"
End Select
End Function

'#################################################################################
'DRAWING##########################################################################
'#################################################################################

Public Sub PrintFormText(ByVal Text As String, X As Single, Y As Single, colour As Long)

'If LenB(Text) = 0 Then Exit Sub

'frmGame.CurrentX = x
'frmGame.CurrentY = y
'
'frmGame.Print Text
Dim lhDC As Long

lhDC = frmGame.picMain.hDC

Call SetBkColor(lhDC, 0)
Call SetTextColor(lhDC, colour)
Call TextOut(lhDC, _
    frmGame.ScaleX(X, vbTwips, vbPixels), _
    frmGame.ScaleY(Y, vbTwips, vbPixels), _
    Text, Len(Text))

End Sub

Public Sub PrintText(ByVal Text As String, X As Single, Y As Single, colour As Long)
Dim Ret As Long

'If LenB(Text) = 0 Then Exit Sub

'                                               zoom
'frmGame.CurrentX = x * cg_Zoom - cg_Camera.x '+ (1 - cg_Zoom) * CentreX
'frmGame.CurrentY = y * cg_Zoom - cg_Camera.y '+ (1 - cg_Zoom) * CentreY

'frmGame.Print Text

Dim lhDC As Long

lhDC = frmGame.picMain.hDC

Call SetBkColor(lhDC, 0)
Call SetTextColor(lhDC, colour)
Call TextOut(lhDC, _
    frmGame.ScaleX(X * cg_Zoom - cg_Camera.X, vbTwips, vbPixels), _
    frmGame.ScaleY(Y * cg_Zoom - cg_Camera.Y, vbTwips, vbPixels), _
    Text, Len(Text))


End Sub

Public Sub gCircle(X As Single, Y As Single, Radius As Single, colour As Long) ', _
    Optional sStart As Single = pi, Optional sEnd As Single = pi, Optional sAspect As Single = 1)

'Circle (X, Y), Radius, Colour, A1, A2

'zoom fix attempt
'frmGame.Circle (X * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
                Y * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY), _
                Radius * cg_Zoom, Colour ', sStart, sEnd, sAspect

frmGame.picMain.Circle (X * cg_Zoom - cg_Camera.X, _
                Y * cg_Zoom - cg_Camera.Y), _
                Radius * cg_Zoom, colour ', sStart, sEnd, sAspect

End Sub

Public Sub gCircleAspect(X As Single, Y As Single, Radius As Single, colour As Long, sAspect As Single)

'zoom fix attempt
'frmGame.Circle (X * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
                Y * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY), _
                Radius * cg_Zoom, Colour, , , sAspect

frmGame.picMain.Circle (X * cg_Zoom - cg_Camera.X, _
                Y * cg_Zoom - cg_Camera.Y), _
                Radius * cg_Zoom, colour, , , sAspect

End Sub

Public Sub gCircleSE(X As Single, Y As Single, Radius As Single, colour As Long, _
    sStart As Single, Send As Single)

'zoom
'frmGame.Circle (X * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
                Y * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY), _
                Radius * cg_Zoom, Colour, sStart, sEnd

frmGame.picMain.Circle (X * cg_Zoom - cg_Camera.X, _
                Y * cg_Zoom - cg_Camera.Y), _
                Radius * cg_Zoom, colour, sStart, Send

End Sub



Public Sub gLine(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, colour As Long)

'zoom
'frmGame.Line (X1 * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
              Y1 * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY) _
            -(X2 * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
              Y2 * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY), _
              Colour

frmGame.picMain.Line (X1 * cg_Zoom - cg_Camera.X, _
              Y1 * cg_Zoom - cg_Camera.Y) _
            -(X2 * cg_Zoom - cg_Camera.X, _
              Y2 * cg_Zoom - cg_Camera.Y), _
              colour

End Sub

Public Sub gBox(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, colour As Long)

'zoom
'frmGame.Line (X1 * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
              Y1 * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY) _
            -(X2 * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
              Y2 * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY), _
              Colour, B

frmGame.picMain.Line (X1 * cg_Zoom - cg_Camera.X, _
              Y1 * cg_Zoom - cg_Camera.Y) _
            -(X2 * cg_Zoom - cg_Camera.X, _
              Y2 * cg_Zoom - cg_Camera.Y), _
              colour, B

End Sub

Public Sub gBoxFilled(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, colour As Long)

'zoom
'frmGame.Line (X1 * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
              Y1 * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY) _
            -(X2 * cg_Zoom - cg_Camera.X + (1 - cg_Zoom) * CentreX, _
              Y2 * cg_Zoom - cg_Camera.Y + (1 - cg_Zoom) * CentreY), _
              Colour, BF

frmGame.picMain.Line (X1 * cg_Zoom - cg_Camera.X, _
              Y1 * cg_Zoom - cg_Camera.Y) _
            -(X2 * cg_Zoom - cg_Camera.X, _
              Y2 * cg_Zoom - cg_Camera.Y), _
              colour, BF

End Sub

Public Sub gPoly(Pts() As PointAPI, lFillCol As Long)
Dim j As Integer

For j = LBound(Pts) To UBound(Pts)
    Pts(j).X = frmGame.ScaleX(Pts(j).X * cg_Zoom - cg_Camera.X, vbTwips, vbPixels)
    Pts(j).Y = frmGame.ScaleY(Pts(j).Y * cg_Zoom - cg_Camera.Y, vbTwips, vbPixels)
Next j

modGDI.DrawPoly Pts, frmGame.picMain.hDC, lFillCol

End Sub

'#################################################################################
'END DRAWING######################################################################
'#################################################################################

'Public Function stickPrintText(ByVal Text As String, X As Single, Y As Single)

'If LenB(Text) = 0 Then Exit Function

'frmStickGame.CurrentX = X
'frmStickGame.CurrentY = Y

'frmStickGame.Print Text

'End Function

Public Function FixAngle(ByVal sngAngle As Single) As Single

Do While sngAngle < 0
    sngAngle = sngAngle + Pi2
Loop
Do While sngAngle > Pi2
    sngAngle = sngAngle - Pi2
Loop

'Return the value
FixAngle = sngAngle


'FixAngle = ((sngAngle * Rad_To_Deg) Mod 360) / Rad_To_Deg


End Function

Public Function FindAngle(intX1 As Single, intY1 As Single, intX2 As Single, intY2 As Single) As Single

Dim sngXComp As Single
Dim sngYComp As Single

'Find the angle between the 2 coords
sngXComp = intX2 - intX1
sngYComp = intY1 - intY2

If Sgn(sngYComp) > 0 Then
    FindAngle = Atn(sngXComp / sngYComp)
ElseIf Sgn(sngYComp) < 0 Then
    FindAngle = Atn(sngXComp / sngYComp) + Pi
End If

End Function

Public Sub AddVectors(sngMag1 As Single, sngDir1 As Single, sngMag2 As Single, sngDir2 As Single, _
Optional ByRef sngMagResult As Single, Optional ByRef sngDirResult As Single)

Dim sngXComp As Single
Dim sngYComp As Single
 
'Determine the components of the resultant vector
sngXComp = (sngMag1 * Sine(sngDir1) + sngMag2 * Sine(sngDir2)) '* frmGame.TimeFactor '* ElapsedTime

sngYComp = (sngMag1 * CoSine(sngDir1) + sngMag2 * CoSine(sngDir2)) '* frmGame.TimeFactor '* ElapsedTime


'Determine the resultant magnitude
sngMagResult = Sqr(sngXComp ^ 2 + sngYComp ^ 2)

'Calculate the resultant direction, and adjust for atngent by adding Pi if necessary
If sngYComp > 0 Then
    sngDirResult = Atn(sngXComp / sngYComp)
ElseIf sngYComp < 0 Then
    sngDirResult = Atn(sngXComp / sngYComp) + Pi
End If

End Sub

Public Sub Motion(ByRef sngX As Single, ByRef sngY As Single, ByVal sngSpeed As Single, ByVal sngHeading As Single)

'Move an object w.r.t. its speed
sngX = sngX + sngSpeed * Sine(sngHeading) * TimeFactor '* sv_GameSpeed * SpaceElapsedTime / Space_Ms_Delay
sngY = sngY - sngSpeed * CoSine(sngHeading) * TimeFactor '* sv_GameSpeed * SpaceElapsedTime / Space_Ms_Delay

End Sub

'Public Sub StickXMotion(ByRef sngX As Single, _
'    ByVal sngSpeed As Single, ByVal byHeading As Byte)
'
'sngX = sngX + IIf(byHeading = 1, sngSpeed, -sngSpeed)
'
'End Sub
'
'Public Sub StickYMotion(ByRef sngY As Single, _
'    ByVal sngSpeed As Single)
'
'sngY = sngY + sngSpeed
'
'End Sub

Public Function GetDist(sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single) As Single
Dim dx As Single, dY As Single

dx = sngX1 - sngX2
dY = sngY1 - sngY2

'Return the distance between the two points (I love you, Mr. Pythagoras)
GetDist = Sqr(dx * dx + dY * dY)

End Function

''ONLY FOR -PID2 <= THETA <= PID2
'Public Function fsine(Theta As Single) As Single
'
'Const B = 4 / pi
'Const C = -4 / (pi * pi)
'
'If Theta < 0 Then
'    Theta = Theta + pi
'End If
'
'fSin = B * Theta + C * Theta * Abs(Theta)
'
'#If EXTRA_PRECISION Then
'    'Const Q = 0.775
'    Const P = 0.225
'
'    fSin = P * (fSin * Abs(fSin) - fSin) + fSin 'Q * y + P * y * abs(y)
'#End If
'
'
'End Function


'Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Declare Function IntersectRect Lib "user32" _
'    (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
'DestRect contains the overlapping rect between src1 and 2


'Public Function Trunc(sngValue As Single, lngDigits As Long) As String
'
'Dim strTemp As String
'Dim lngNumZeros As Long
'Dim i As Long
'
''Truncate to two decimal places
'strTemp = ((CStr(sngValue) * (10 ^ lngDigits)) \ 1) / (10 ^ lngDigits)
'If InStr(1, strTemp, ".") = 0 Then
'    lngNumZeros = lngDigits
'    strTemp = strTemp & "."
'ElseIf LenB(strTemp) - InStr(1, strTemp, ".") < lngDigits Then
'    lngNumZeros = lngDigits - (LenB(strTemp) - InStr(1, strTemp, "."))
'End If
'
'If lngNumZeros > 0 Then
'    For i = 1 To lngNumZeros
'        strTemp = strTemp & "0"
'    Next i
'End If
'
'Trunc = strTemp
'
'End Function
'
'Public Sub LineDraw(ByVal lngX1 As Long, ByVal lngY1 As Long, ByVal lngX2 As Long, ByVal lngY2 As Long, lngDC As Long)
'
'Dim udtPoint As POINTAPI
'
''This routine draws a box of specific colour on the display
'Call MoveToEx(lngDC, lngX1, lngY1, udtPoint)    'Move current pen x,y
'Call LineTo(lngDC, lngX2, lngY2)                'Draw line from current x,y to given x,y
'
'End Sub
'
'Public Sub EllipseDraw(lngX1 As Long, lngY1 As Long, lngX2 As Long, lngY2 As Long, lngDC As Long)
'
''Draw the given ellipse
'Ellipse lngDC, lngX1, lngY1, lngX2, lngY2
'
'End Sub
