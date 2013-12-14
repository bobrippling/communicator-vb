VERSION 5.00
Begin VB.Form frmThumbnail 
   Caption         =   "Communicator Thumbnail"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5400
End
Attribute VB_Name = "frmThumbnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hThumbNail As Long
Private Const kOpacity = 170

Private Sub Form_Load()
Dim frmMainAspect As Single

modLoadProgram.frmThumbNail_Loaded = True

frmMainAspect = frmMain.width / frmMain.height
Me.width = Me.height * frmMainAspect

Me.Show
InitThumbNail
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
TerminateThumbnail
modLoadProgram.frmThumbNail_Loaded = False
End Sub

Private Sub InitThumbNail()
Dim lErr As Long
Dim sDesc As String

Dim hWndSource As Long, hWndDest As Long
hWndSource = frmMain.hWnd
hWndDest = Me.hWnd


If modDisplay.RegisterThumbnail(hThumbNail, hWndSource, hWndDest) Then
    SetThumbNailProps
Else
    hThumbNail = 0
    
    lErr = Err.LastDllError
    sDesc = Err.Description
    If lErr = 0 Then lErr = Err.Number
    
    AddConsoleText "Error Registering Thumbnail: " & lErr & IIf(LenB(sDesc), " (" & sDesc & ")", vbNullString)
End If

End Sub

Private Sub SetThumbNailProps()
Dim rcSource As RECT, rcDest As RECT
Dim lErr As Long
Dim sDesc As String


If hThumbNail Then
    
    With rcSource
        .Left = 0
        .Top = 0
        .Bottom = ScaleY(frmMain.height, vbTwips, vbPixels)
        .Right = ScaleX(frmMain.width, vbTwips, vbPixels)
    End With
    
    With rcDest
        .Left = 0
        .Top = 0
        .Bottom = ScaleY(Me.height, vbTwips, vbPixels)
        .Right = ScaleX(Me.width, vbTwips, vbPixels)
    End With
    
    
    If modDisplay.SetThumbNailProps(hThumbNail, kOpacity, True, True, rcSource, rcDest) = False Then
        
        lErr = Err.LastDllError
        sDesc = Err.Description
        If lErr = 0 Then lErr = Err.Number
        
        AddConsoleText "Error Setting Thumbnail Properties: " & lErr & IIf(LenB(sDesc), vbSpace & "(" & sDesc & ")", vbNullString)
    End If
End If

End Sub

Public Sub RefreshThumbNailRect(bDestRect As Boolean)
Dim rcRect As RECT
Dim lErr As Long
Dim sDesc As String


If hThumbNail Then
    
    If bDestRect Then
        With rcRect
            .Left = 0
            .Top = 0
            .Bottom = ScaleY(Me.height, vbTwips, vbPixels)
            .Right = ScaleX(Me.width, vbTwips, vbPixels)
        End With
    Else
        With rcRect
            .Left = 0
            .Top = 0
            .Bottom = ScaleY(frmMain.height, vbTwips, vbPixels)
            .Right = ScaleX(frmMain.width, vbTwips, vbPixels)
        End With
    End If
    
    
    If modDisplay.SetThumbNailRect(hThumbNail, rcRect, Not bDestRect) = False Then
        lErr = Err.LastDllError
        sDesc = Err.Description
        If lErr = 0 Then lErr = Err.Number
        
        AddConsoleText "Error Setting Thumbnail Properties: " & lErr & IIf(LenB(sDesc), vbSpace & "(" & sDesc & ")", vbNullString)
    End If
    
End If

End Sub

Private Sub TerminateThumbnail()

If hThumbNail Then
    modDisplay.UnRegisterThumbnail hThumbNail
End If

End Sub

'##################################################################

'Private Sub sldrTransparency_Change()
'SetThumbNailProps
'End Sub
'
'Private Sub sldrTransparency_Scroll()
'sldrTransparency_Change
'End Sub

Private Sub Form_Resize()

'With sldrTransparency
'    .Top = Me.ScaleHeight - .height - 240
'    .width = Me.ScaleWidth - .Left * 2
'End With

RefreshThumbNailRect True
End Sub


