VERSION 5.00
Begin VB.Form frmGmail 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   1395
   ClientLeft      =   5985
   ClientTop       =   2715
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   4080
      Top             =   120
   End
   Begin VB.Image imgPic 
      Height          =   525
      Left            =   120
      Picture         =   "frmGmail.frx":0000
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmGmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'###############################################################
'transparency
Private pTransparency As Byte
'###############################################################

Private PosX As Long
Private BottomY As Long
Private Const kHeight = 1615

Private bShowing As Boolean
Private iCount As Integer
Private Const ShowTime = 200

Public Property Let sCaption(sTxt As String)
lblCaption.Caption = sTxt
End Property
Public Property Let sTitle(sTxt As String)
lblTitle.Caption = sTxt
End Property

Private Property Let Transparency(nVal As Byte)

'0 = completely transparent, 255 = completely opaque

modDisplay.SetTransparency Me.hWnd, nVal
pTransparency = nVal

End Property

Private Sub Form_Load()
Dim hWnd As Long, l As Long

hWnd = Me.hWnd
l = GetWindowLong(hWnd, GWL_EXSTYLE)
If (l And WS_EX_LAYERED) = 0 Then
    'add the layerd bit
    SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_LAYERED Or l
End If


Transparency = 1

PosX = Screen.width - Me.width
BottomY = Screen.height - modAlert.TB_Height

Me.Move PosX, BottomY

bShowing = True

End Sub

Private Sub Form_Click()
iCount = ShowTime
frmSystray.BalloonClicked
End Sub

Private Sub imgPic_Click()
Form_Click
End Sub

Private Sub lblText_Click()
Form_Click
End Sub

Private Sub lblCaption_Click()
Form_Click
End Sub

Private Sub lblTitle_Click()
Form_Click
End Sub

Private Sub tmrMain_Timer()
Const Inc = 30

If bShowing Then
    Me.Top = Me.Top - Inc
    AdjustHeight
    
    If Me.height >= kHeight Then
        Me.height = kHeight
        bShowing = False
    End If
    
    If pTransparency < 250 Then
        Transparency = pTransparency + 4
    End If
    
ElseIf iCount < ShowTime Then
    iCount = iCount + 1
    
Else
    
    Me.Top = Me.Top + Inc
    AdjustHeight
    
    If Me.Top = BottomY Then
        tmrMain.Enabled = False
        Unload Me
    ElseIf pTransparency > 5 Then
        Transparency = pTransparency - 5
    End If
    
End If

End Sub

Private Sub AdjustHeight()

Me.height = BottomY - Me.Top

End Sub
