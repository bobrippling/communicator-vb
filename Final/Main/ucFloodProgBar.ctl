VERSION 5.00
Begin VB.UserControl ucFloodProgBar 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7890
   ScaleHeight     =   1695
   ScaleWidth      =   7890
   Begin VB.PictureBox picFlood 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "ucFloodProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const Colour_White = &H800000, Colour_Black = &HFFFFFF, _
    Colour_Success = &H3A633D, Colour_Fail = &H3D2785

Public Sub Cls()
picFlood.Cls
End Sub

Public Sub Flood_Show_Percentage(ByVal lPercent As Single, sText As String)

If lPercent > 100 Then
    lPercent = 100
ElseIf lPercent < 0 Then
    lPercent = 0
End If


picFlood.Cls

'calculate the string's X & Y coordinates
'in the PictureBox ... here, left justified and offset slightly
Centre_XY sText

picFlood.ForeColor = Colour_White

'print the percentage string in the text colour
picFlood.Print sText

'print the flood bar to the new progress length in the line colour
picFlood.Line (0, 0)-(lPercent * picFlood.width / 100, picFlood.height), , BF

'without this DoEvents or Refresh, the flood won't update
picFlood.Refresh

End Sub


Public Sub Flood_Show_Result(bSuccess As Boolean, sInfo As String)

picFlood.Cls

'calculate the string's X & Y coordinates
'in the Picture Box ... here, left justified and offset slightly
Centre_XY sInfo


'print the percentage string in the text colour
picFlood.ForeColor = IIf(bSuccess, Colour_Success, Colour_Fail)
picFlood.Print sInfo

'print the flood bar to the new progress length in the line colour
picFlood.Line (0, 0)-(picFlood.ScaleWidth, picFlood.ScaleHeight), , BF

End Sub


Public Sub Flood_Show_Message(sMessage As String)

picFlood.Cls
 
'calculate the string's X & Y coordinates
'in the PictureBox ... here, left justified and offset slightly
Centre_XY sMessage

picFlood.ForeColor = Colour_White

'print the percentage string in the text colour
picFlood.Print sMessage

End Sub

Private Sub Centre_XY(sMsg As String)
picFlood.CurrentX = picFlood.ScaleWidth \ 2 - picFlood.TextWidth(sMsg) \ 2
picFlood.CurrentY = picFlood.ScaleHeight \ 2 - picFlood.TextHeight(sMsg) \ 2
End Sub

Private Sub UserControl_Initialize()

'initialize the control by setting:
'white (the text colour)
'black (the flood panel colour)
'not Xor pen
'solid fill
With picFlood
    .Cls
    .BackColor = Colour_Black
    .ForeColor = Colour_White
    .DrawMode = vbNotXorPen '10
    .FillStyle = vbFSSolid '0
    .AutoRedraw = True
    .Cls
End With

End Sub

Private Sub UserControl_Resize()
With picFlood
    .width = UserControl.width
    .height = UserControl.height
End With
End Sub
