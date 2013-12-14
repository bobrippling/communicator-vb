VERSION 5.00
Begin VB.UserControl IPTextBox 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   ScaleHeight     =   780
   ScaleWidth      =   4470
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      ForeColor       =   &H002222E6&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "IPTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Change()
Public Event Click()
Public Event tLostFocus()
Public Event tGotFocus()
Public Event KeyPress(KeyAscii As Integer)

Private Const Colour_Valid = &HFF0000
Private Const Colour_Invalid = &H2222E6

Private bIsValid As Boolean

Public Property Get HasValidIP() As Boolean
HasValidIP = bIsValid
End Property
Public Sub ShowIPBalloonTip()

If UserControl.txtIP.Enabled Then
    If bIsValid Then
        modDisplay.ShowBalloonTip txtIP, "Valid IP", _
            "IP accepted! Couldn't think of anything else to put in the balloon"
        
    Else
        modDisplay.ShowBalloonTip txtIP, "Invalid IP", _
            "The IP Address should be in the format xxx.xxx.xxx.xxx, " & _
            "unless it is a server name", TTI_WARNING
        
    End If
End If

End Sub

Public Property Get hWnd() As Long
hWnd = txtIP.hWnd
End Property

Private Sub txtIP_Change()

If IsIP(txtIP.Text) Then
    txtIP.ForeColor = Colour_Valid
    bIsValid = True
Else
    txtIP.ForeColor = Colour_Invalid
    bIsValid = False
End If


RaiseEvent Change

End Sub

Public Property Get Enabled() As Boolean
Enabled = txtIP.Enabled
End Property

Public Property Let Enabled(ByVal bVal As Boolean)
txtIP.Enabled = bVal
End Property

Private Sub txtIP_Click()
RaiseEvent Click
End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtIP_LostFocus()
RaiseEvent tLostFocus
End Sub

Private Sub txtIP_GotFocus()
RaiseEvent tGotFocus
End Sub

Public Property Let Text(Txt As String)
txtIP.Text = Txt
End Property

Public Property Get Text() As String
Text = txtIP.Text
End Property

Public Property Let Selstart(i As Integer)
txtIP.Selstart = i
End Property

Public Property Get Selstart() As Integer
Selstart = txtIP.Selstart
End Property

Public Property Let Sellength(i As Integer)
txtIP.Sellength = i
End Property

Public Property Get Sellength() As Integer
Sellength = txtIP.Sellength
End Property

Private Function IsIP(ByVal IP As String) As Boolean
Dim IPs() As String
Dim i As Integer

IP = LCase$(IP)

If IP = vbNullString Then GoTo NotIP

If IP = "localhost" Then
    IsIP = True
    Exit Function
End If

IPs = Split(IP, ".", , vbTextCompare)

If UBound(IPs) <> 3 Then GoTo NotIP
If LBound(IPs) Then GoTo NotIP

For i = 0 To 3
    If Len(IPs(i)) > 3 Then GoTo NotIP
    If Len(IPs(i)) < 1 Then GoTo NotIP
    If Not IsNumeric(IPs(i)) Then GoTo NotIP
    If IPs(i) > 255 Then GoTo NotIP
    If IPs(i) < 0 Then GoTo NotIP
Next i

IsIP = True

Exit Function
NotIP:
IsIP = False
End Function

Private Sub UserControl_Initialize()

bIsValid = IsIP(txtIP.Text)

'modSubClass.SubclasstxtIP txtIP.hWnd

End Sub

Private Sub UserControl_Terminate()

'modSubClass.SubclasstxtIP txtIP.hWnd, False

End Sub

Private Sub UserControl_Resize()
'On Error Resume Next
txtIP.width = UserControl.width
'txtIP.height = UserControl.height - 100
UserControl.height = 290
End Sub
