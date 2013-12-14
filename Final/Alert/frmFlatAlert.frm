VERSION 5.00
Begin VB.Form frmFlatAlert 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   1815
   ClientLeft      =   12495
   ClientTop       =   11520
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMain 
      Interval        =   1
      Left            =   2160
      Top             =   120
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
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
      Width           =   2535
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your text here"
      Height          =   1215
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmFlatAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ShowTime As Long = 200

Private PosX As Long
Private BottomY As Long
Private Const kHeight = 1815

Private bShowing As Boolean
Private iCount As Integer

Public Property Let sCaption(sTxt As String)
lblCaption.Caption = sTxt
End Property
Public Property Let sTitle(sTxt As String)
lblTitle.Caption = sTxt
End Property

Private Sub Form_Click()
iCount = ShowTime
frmSystray.BalloonClicked
End Sub

Private Sub Form_Load()

PosX = Screen.width - Me.width
BottomY = Screen.height - modAlert.TB_Height

Me.Move PosX, BottomY

bShowing = True

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
