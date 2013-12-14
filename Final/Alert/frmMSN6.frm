VERSION 5.00
Begin VB.Form frmMSN6 
   BorderStyle     =   0  'None
   ClientHeight    =   1785
   ClientLeft      =   6150
   ClientTop       =   4815
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMSN6.frx":0000
   ScaleHeight     =   1785
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   1920
      Top             =   0
   End
   Begin VB.Line lnMain 
      X1              =   0
      X2              =   2760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Image imgClose 
      Height          =   240
      Left            =   2520
      Picture         =   "frmMSN6.frx":10206
      Top             =   45
      Width           =   225
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
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   2295
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   1095
      Left            =   240
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmMSN6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PosX As Long
Private BottomY As Long
Private Const kHeight = 1785

Private Const ShowTime = 200
Private TaskBarHeight As Long

Private bShowing As Boolean
Private iCount As Integer

Public Property Let sCaption(sTxt As String)
lblCaption.Caption = sTxt
End Property
Public Property Let sTitle(sTxt As String)
lblTitle.Caption = sTxt
End Property

Private Sub Form_Load()

PosX = Screen.width - Me.width
BottomY = Screen.height - modAlert.TB_Height

Me.Move PosX, BottomY
Me.height = 1

bShowing = True

End Sub

Private Sub imgClose_Click()
iCount = ShowTime
End Sub

Private Sub lblCaption_Click()
Form_Click
End Sub

Private Sub lblTitle_Click()
Form_Click
End Sub

Private Sub Form_Click()
frmSystray.BalloonClicked
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
    
ElseIf iCount < ShowTime Then
    iCount = iCount + 1
    
Else
    
    Me.Top = Me.Top + Inc
    AdjustHeight
    
    If Me.Top = BottomY Then
        tmrMain.Enabled = False
        Unload Me
    End If
    
End If

End Sub

Private Sub AdjustHeight()

Me.height = BottomY - Me.Top

End Sub
