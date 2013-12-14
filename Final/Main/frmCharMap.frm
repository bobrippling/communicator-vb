VERSION 5.00
Begin VB.Form frmCharMap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Character Map"
   ClientHeight    =   3870
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7815
   Icon            =   "frmCharMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   Begin VB.PictureBox picChar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox picBG 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   840
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdComm 
      Caption         =   "Add to Chat"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtDec 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtBin 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtHex 
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.PictureBox picChars 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   120
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   417
      TabIndex        =   5
      Top             =   720
      Width           =   6255
   End
   Begin VB.ComboBox cboFonts 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtCopy 
      Height          =   375
      HideSelection   =   0   'False
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblFont 
      Caption         =   "Font:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblCopy 
      Caption         =   "Characters to copy:"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2160
      TabIndex        =   13
      Top             =   2760
      Width           =   420
   End
   Begin VB.Label lblDec 
      Caption         =   "Dec:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label lblBin 
      Caption         =   "Bin:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label lblHex 
      Caption         =   "Hex:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblCharCode 
      AutoSize        =   -1  'True
      Caption         =   "CharCode"
      Height          =   195
      Left            =   2160
      TabIndex        =   16
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Info"
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Top             =   3465
      Width           =   270
   End
End
Attribute VB_Name = "frmCharMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

Private Declare Function apiShowCursor Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long


Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const SRCCOPY = &HCC0020

Private AsciiList(0 To 249) As String 'list of character descriptions
Private SizeX As Long, SizeY As Long, previousX As Long, previousY As Long
Private bMouseDown As Boolean
Private SelectedSquare As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub CboFonts_Click()

DrawCharacterMap cboFonts.List(cboFonts.ListIndex)

picChar.Font = cboFonts.List(cboFonts.ListIndex)
picChar.FontSize = 18
txtCopy.Font = cboFonts.List(cboFonts.ListIndex)

'Reselect last square
DrawFocusColour previousX, previousY

End Sub

Private Sub cmdComm_Click()
If Status = Connected Then
    frmMain.txtOut.SelText = txtCopy.Text
    
    modDisplay.ShowBalloonTip txtCopy, "Character Map", "Characters have been added to the chat box thing", TTI_INFO
Else
    cmdComm.Enabled = False
End If
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtCopy.Text, vbCFText

SetFocus2 picChars
End Sub

Private Sub cmdSelect_Click()
Dim X1 As Long, Y1 As Long, Char As String, lpRect As RECT, offsetx  As Long, offsety  As Long
Dim iSquare As Long

iSquare = SelectedSquare
Y1 = iSquare \ 32
X1 = iSquare Mod 32

Char = Chr$((Y1 * 32) + (X1 + 1) + 30)
txtCopy.SelText = Char


SetFocus2 picChars
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim sFont As String

SizeX = (picChars.ScaleWidth \ 32) ' + 1 '  32*7=224
SizeY = (picChars.ScaleHeight \ 7) ''' + 1 '  32*7=224

CreateAsciiList

FillListWithFonts cboFonts
cboFonts.ListIndex = 0

sFont = frmMain.rtfFontName
For i = 0 To cboFonts.ListCount - 1
    If cboFonts.List(i) = sFont Then
        cboFonts.ListIndex = i
        Exit For
    End If
Next i


picChar.Visible = False
picBG.Visible = False
cmdCopy.Enabled = False
cmdComm.Enabled = False

picChars_MouseDown 0, 0, 5, 5
picChars_MouseUp 0, 0, 5, 5

cmdComm.Enabled = (Status = Connected)
Call FormLoad(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call FormLoad(Me, True)
End Sub

Private Sub UpdateLabels(X As Long, Y As Long)
Dim key As Integer, K As String
'give keystroke and alt information

key = (Y * 32) + (X + 1) ' + 31

K = "Keystroke: "
Select Case key
    Case 1
        lblCharCode.Caption = K & "Spacebar"
        
    Case 2 To 95 '
        If key = 7 Then 'need && to show in label for ampersand
            lblCharCode.Caption = K & "&&"
        Else
            lblCharCode.Caption = K & Chr$(key + 31)
        End If
    
    Case 96 To 97
        lblCharCode.Caption = K & "Ctrl+" & CStr(key - 95)
        
    Case 98 To 224
        lblCharCode.Caption = K & "Alt+0" & CStr(key + 31)
        
End Select


'hex / bin text
txtHex.Text = Hex(key + 31)
txtBin.Text = Bin(key + 31, 8)
txtDec.Text = key + 31

lblInfo.Caption = "Col: " & X & " Line: " & Y & " Square:" & (Y * 32) + (X + 1) & " Ascii: " & key + 31

lblName.Caption = AsciiList(key - 1)
'Select Case key
'    Case 1 To 98
'        lblName.Caption = asciiList(key - 1)
'    Case 99 To 129
'        lblName.Caption = asciiList(key - 1)
'    Case 130 To 224
'        lblName.Caption = asciiList(key - 1)
'End Select

End Sub

Private Sub CreateAsciiList()
Dim aList() As String
Dim i As Integer

aList = Split(LoadResText(101), ",", , vbTextCompare)

For i = 0 To UBound(aList) - 1
    AsciiList(i) = StrConv(aList(i), vbProperCase)
Next i

End Sub

Private Sub picChars_DblClick()
cmdSelect_Click
End Sub

Private Sub picChars_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyDown
        If SelectedSquare + 32 < 225 Then
            SelectedSquare = SelectedSquare + 32
        End If
    Case vbKeyUp
        If SelectedSquare - 32 > 0 Then
            SelectedSquare = SelectedSquare - 32
        End If
    Case vbKeyRight
        If SelectedSquare + 1 < 225 Then
            SelectedSquare = SelectedSquare + 1
        End If
    Case vbKeyLeft
        If SelectedSquare - 1 > 0 Then
            SelectedSquare = SelectedSquare - 1
        End If
    Case Else
        Exit Sub
End Select

DrawSelected (SelectedSquare - 1)
UpdateLabels (SelectedSquare - 1) Mod 32, (SelectedSquare - 1) \ 32

End Sub

Private Sub DrawSelected(iSquare As Long)

Dim X1 As Long, Y1 As Long, Char As String, lpRect As RECT, offsetx As Long, offsety As Long
Y1 = iSquare \ 32
X1 = iSquare Mod 32

'erase previous ?
picChars.Line (previousX * SizeX + 1, previousY * SizeY + 1)-(previousX * SizeX + (SizeX - 1), previousY * SizeY + (SizeY - 1)), vbWhite, BF
picChars.CurrentX = (previousX * SizeX) + 3
picChars.CurrentY = (previousY * SizeY)

picChars.Print Chr$((previousY * 32) + (previousX + 1) + 31);
previousX = X1
previousY = Y1

Char = Chr$((Y1 * 32) + (X1 + 1) + 31)
picChar.Visible = False: picBG.Visible = False
offsetx = (picChar.ScaleWidth - picChar.TextWidth(Char)) \ 2
offsety = (picChar.ScaleHeight - picChar.TextHeight(Char)) \ 2
picChar.Left = (X1 * SizeX - 5) + 10
picChar.Top = (Y1 * SizeY - 5) + 35
picBG.Left = picChar.Left + 5
picBG.Top = picChar.Top + 5
picChar.CurrentX = offsetx
picChar.CurrentY = offsety
picChar.Picture = LoadPicture()
picChar.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
picChar.Visible = True
picBG.Visible = True

End Sub

Private Sub picChars_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim X1 As Long, Y1 As Long, Ret As Long, lpRect As RECT, offsetx As Long, offsety As Long, Char As String

X1 = X \ SizeX
Y1 = Y \ SizeY

If Button = vbRightButton Then Exit Sub

If Me.Visible Then SetCursor False

'if in square of picture
If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
    'erase previous focus rectangle
    If Not (IsEmpty(previousX) And IsEmpty(previousY)) Then
        lpRect.Left = X1 * SizeX + 1
        lpRect.Top = Y1 * SizeY + 1
        lpRect.Right = X1 * SizeX + (SizeX - 1) + 1 '- 1
        lpRect.Bottom = Y1 * SizeY + (SizeY - 1) + 1
        
        picChars.Line (previousX * SizeX, previousY * SizeY)-(previousX * SizeX + (SizeX), previousY * SizeY + (SizeY)), vbBlack, BF
        picChars.Line (previousX * SizeX + 1, previousY * SizeY + 1)-(previousX * SizeX + (SizeX - 1), previousY * SizeY + (SizeY - 1)), vbWhite, BF
        
        Char = Chr$((previousY * 32) + (previousX + 1) + 31)
        offsetx = (SizeX - picChars.TextWidth(Char)) \ 2
        offsety = (SizeY - picChars.TextHeight(Char)) \ 2
        picChars.CurrentX = (previousX * SizeX) + offsetx
        picChars.CurrentY = (previousY * SizeY) + offsety
        picChars.Print Char;
        
    End If
    
    picChar.Visible = False
    picBG.Visible = False
    picChar.Left = (X1 * SizeX - 5) + 10
    picChar.Top = (Y1 * SizeY - 5) + 35
    picBG.Left = picChar.Left + 5
    picBG.Top = picChar.Top + 5
    picChar.Visible = True
    picBG.Visible = True
    SelectedSquare = (Y1 * 32) + (X1 + 1)

    previousX = X1
    previousY = Y1
End If

'draw focus rectangle

Call UpdateLabels(X1, Y1)

picChar.Visible = True
picBG.Visible = True
bMouseDown = True
End Sub

Private Sub picChars_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim X1 As Long, Y1 As Long, Ret As Long, Char As String, key As Integer
Dim offsetx  As Long, offsety As Long
Static lastx As Long
Static lasty As Long

X1 = X \ SizeX
Y1 = Y \ SizeY


If bMouseDown Then
    If X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6 Then
        
        If lastx = X1 And lasty = Y1 Then Exit Sub
        
        lastx = X1: lasty = Y1
        key = (Y1 * 32) + (X1 + 1)
        
        picChar.Visible = False
        picBG.Visible = False
        picChar.Left = (X1 * SizeX - 5) + 10
        picChar.Top = (Y1 * SizeY - 5) + 35
        picBG.Left = picChar.Left + 5
        picBG.Top = picChar.Top + 5
        
        Char = Chr$((Y1 * 32) + (X1 + 1) + 31)
        
        If picChar.Tag <> Char Then
            previousX = X1
            previousY = Y1
            picChar.Tag = Char
            
            
            offsetx = (picChar.ScaleWidth - picChar.TextWidth(Char)) \ 2
            offsety = (picChar.ScaleHeight - picChar.TextHeight(Char)) \ 2
            picChar.CurrentX = offsetx
            picChar.CurrentY = offsety
            picChar.Picture = LoadPicture()
            picChar.Print Chr$((Y1 * 32) + (X1 + 1) + 31)
            picChar.Visible = True
            picBG.Visible = True
        End If
        
        UpdateLabels X1, Y1
        previousX = X1
        previousY = Y1
    End If
    
End If

End Sub

Private Sub picChars_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Ret As Long, X1 As Long, Y1 As Long, lpRect As RECT

X1 = X \ SizeX
Y1 = Y \ SizeY

DrawFocusColour previousX, previousY

SetCursor

If Not (X1 >= 0 And X1 <= 31 And Y1 >= 0 And Y1 <= 6) Then
    
    If bMouseDown Then
        picChar.Visible = False
        picBG.Visible = False
        'draw red focus rectangle
        DrawFocusColour previousX, previousY
    End If
    
End If

picChar.Visible = False
picBG.Visible = False
bMouseDown = False

End Sub

Private Sub DrawCharacterMap(FName As String)
Dim X As Long, Y As Long, Char As String, lpPT As POINTAPI

Dim offsetx As Long, offsety As Long

picChars.Visible = False
picChars.FontName = FName
picChars.FontSize = 8
Set picChars.Picture = Nothing

For X = 0 To 31
    For Y = 0 To 6
        Char = Chr$((Y * 32) + (X + 1) + 31)
        
        offsetx = (SizeX - picChars.TextWidth(Char)) \ 2
        offsety = (SizeY - picChars.TextHeight(Char)) \ 2
        
        picChars.CurrentX = (X * SizeX) + offsetx
        picChars.CurrentY = (Y * SizeY) + offsety
        picChars.Print Char;
        
    Next Y
Next X

'borders
For X = 0 To 7
    MoveToEx picChars.hDC, 0, X * SizeY, lpPT
    LineTo picChars.hDC, SizeX * 32, X * SizeY
Next X

For X = 0 To 32
    MoveToEx picChars.hDC, X * SizeX, 0, lpPT
    LineTo picChars.hDC, X * SizeX, SizeY * 7 + 1
Next X

picChars.Visible = True
End Sub

Private Sub txtCopy_Change()
cmdCopy.Enabled = LenB(txtCopy.Text)
cmdComm.Enabled = cmdCopy.Enabled And (Status = Connected)
End Sub

Private Sub DrawFocusColour(X As Long, Y As Long)
'make it look like picBox has focus (blue)
Dim lpRect As RECT, offsetx As Long, offsety As Long, Char As String

picChars.Line (X * SizeX + 1, Y * SizeY + 1)-(X * SizeX + (SizeX - 1), _
        Y * SizeY + (SizeY - 1)), vbHighlight, BF

Char = Chr$((Y * 32) + (X + 1) + 31)

offsetx = (SizeX - picChars.TextWidth(Char)) \ 2
offsety = (SizeY - picChars.TextHeight(Char)) \ 2

picChars.CurrentX = (X * SizeX) + offsetx
picChars.CurrentY = (Y * SizeY) + offsety

picChars.ForeColor = vbWhite
picChars.Print Char;
picChars.ForeColor = vbBlack

lpRect.Left = X * SizeX + 1
lpRect.Top = Y * SizeY + 1
lpRect.Right = X * SizeX + (SizeX - 1) + 1
lpRect.Bottom = Y * SizeY + (SizeY - 1) + 1

DrawFocusRect picChars.hDC, lpRect

End Sub

Private Sub SetCursor(Optional bShow As Boolean = True)
Dim Ret As Long
Dim RC As RECT, TopCorner As POINTAPI
Static bClipped As Boolean

If bShow Then
    Do
        Ret = apiShowCursor(True)
    Loop While Ret < 0
    
    If bClipped Then
        ClipCursor ByVal 0&
        bClipped = False
    End If
Else
    Do While Ret >= 0
        Ret = apiShowCursor(False)
    Loop
    
    
    If Not bClipped Then
        GetClientRect picChars.hWnd, RC
        
        TopCorner.X = RC.Left
        TopCorner.Y = RC.Top
        
        ClientToScreen picChars.hWnd, TopCorner
        OffsetRect RC, TopCorner.X, TopCorner.Y
        
        RC.Bottom = RC.Bottom - 5
        RC.Right = RC.Right - 5
        
        ClipCursor RC
        
        bClipped = True
    End If
    
End If

End Sub

Private Function Bin(ByVal Value As Long, Optional digits As Long = -1) As String

Dim Result As String, exponent As Integer
'this is faster than creating the string by appending chars
Result = String$(32, "0")

Do
    If Value And Power2(exponent) Then
        ' we found a bit that is set, clear it
        Mid$(Result, 32 - exponent, 1) = "1"
        Value = Value Xor Power2(exponent)
    End If
    exponent = exponent + 1
Loop While Value
If digits < 0 Then
    ' trim non significant digits, if digits was omitted or negative
    Bin = Mid$(Result, 33 - exponent)
Else
    ' else trim to the requested number of digits
    Bin = Right$(Result, digits)
End If
End Function

' Raise 2 to a power
' the exponent must be in the range [0,31]

'/******************************************************************************
Function Power2(ByVal exponent As Long) As Long
'/******************************************************************************
Static res(0 To 31) As Long
Dim i As Long

' rule out errors
If exponent < 0 Or exponent > 31 Then Err.Raise 5

' initialize the array at the first call
If res(0) = 0 Then
    res(0) = 1
    For i = 1 To 30
        res(i) = res(i - 1) * 2
    Next
    ' this is a special case
    res(31) = &H80000000
End If

' return the result
Power2 = res(exponent)

End Function
