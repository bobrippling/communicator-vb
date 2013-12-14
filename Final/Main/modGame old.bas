Attribute VB_Name = "modGame"
Option Explicit

Public Type ptBullet
    x As Integer
    Y As Integer
    Facing As Integer
    LifeLeft As Integer
    Speed As Single
    Active As Boolean 'must be above pttank
End Type

Public Type ptTank
    x As Integer
    Y As Integer
    Facing As Integer
    Colour As Long
    Speed As Single
    Accelerating As Integer '-1 = slowing, 1 = speeding, 0 = k speed
    Shooting As Boolean
    Bullet As ptBullet
End Type

Public Enum PacketType
    UpdatePos = 0
    UpdateBulletPos = 1
    UpdateHealth = 2
End Enum

Private Const GSep As String = "@"
Public GameFormLoaded As Boolean
Public OtherTanks() As ptTank
Public MyTank As ptTank

Public Sub ProcessTankData(ByVal Str As String, ByVal Player As Integer)

Dim U As Boolean, D As Boolean, L As Boolean, R As Boolean, f As Boolean
Dim x As Integer, Y As Integer, j As Integer

If GameFormLoaded = False Then
    Load frmGame
    frmGame.Show , frmMain
End If

'packet = UDLRFX"@"Y

On Error Resume Next
U = CBool(Left$(Str, 1))
D = CBool(Mid$(Str, 2, 1))
L = CBool(Mid$(Str, 3, 1))
R = CBool(Mid$(Str, 4, 1))
f = CBool(Mid$(Str, 5, 1))
j = InStr(1, Str, GSep, vbTextCompare)
x = Mid$(Str, 6, j - 6)
Y = Mid$(Str, j + 1)
On Error GoTo 0

If Exists(OtherTanks(Player)) = False Then
    CreateTank Player
End If

OtherTanks(Player).Accelerating = IIf(U, 1, IIf(D, -1, 0))
OtherTanks(Player).Shooting = f
OtherTanks(Player).x = x
OtherTanks(Player).Y = Y

End Sub

Public Function CreateTankData() As String
Dim U As Boolean, D As Boolean, L As Boolean, R As Boolean, f As Boolean
Dim x As Integer, Y As Integer
Dim Str As String

U = frmGame.UpKey
D = frmGame.DownKey
L = frmGame.LeftKey
R = frmGame.RightKey
f = frmGame.FireKey
x = MyTank.x
Y = MyTank.Y

Str = CStr(CInt(U)) & CStr(CInt(D)) & CStr(CInt(L)) & CStr(CInt(R)) & _
     CStr(CInt(f)) & Fill$(x, 4) & GSep & Fill(Y, 4)

CreateTankData = Str

End Function

Private Function Exists(Tank As ptTank) As Boolean

Exists = True

On Error GoTo EH
If Tank.x > -1 Then
    Exists = True
End If

Exit Function
EH:
Exists = False
End Function

Private Sub CreateTank(ByVal i As Integer)

ReDim Preserve OtherTanks(1 To i)

End Sub
