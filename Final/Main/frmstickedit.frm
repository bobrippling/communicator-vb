VERSION 5.00
Begin VB.Form frmStickEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stick Map Editor"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset/New Map"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemoveBox 
      Caption         =   "Remove Barrier"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddBox 
      Caption         =   "Add Barrier"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemovetBox 
      Caption         =   "Remove Box"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddtBox 
      Caption         =   "Add Box"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Map..."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Map..."
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemovePlatform 
      Caption         =   "Remove Platform"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddPlatform 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Add Platform"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmStickEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Current_BackCol = vbBlue, def_BackCol = &HC0C0C0, def_Platform_BackCol = &H808080

Private Sub cmdAddBox_Click()
Dim i As Integer

With frmStickGame
    i = .oBox.UBound + 1
    
    Load .oBox(i)
    With .oBox(i)
        .Left = 375
        .width = 375 * Stick_Edit_Zoom
        .height = 2000 * Stick_Edit_Zoom
        .Top = frmStickGame.height \ 4
        .Visible = True
        
        .BorderColor = Current_BackCol
    End With
    
    .oBox(i - 1).BorderColor = def_BackCol
    
End With

cmdRemoveBox.Enabled = True
End Sub
Private Sub cmdRemoveBox_Click()
Dim i As Integer

With frmStickGame
    .DragEnd
    
    i = .oBox.UBound
    If i Then
        Unload .oBox(i)
        
        If i > 1 Then
            .oBox(i - 1).BorderColor = Current_BackCol
        End If
    End If
    i = i - 1
    
    If i <= 0 Then
        cmdRemoveBox.Enabled = False
    End If
End With

End Sub

'#########################################################################################################

Private Sub cmdAddtBox_Click()
Dim i As Integer

With frmStickGame
    i = .otBox.UBound + 1
    
    Load .otBox(i)
    With .otBox(i)
        .Left = 375
        .width = 375 * Stick_Edit_Zoom
        .height = 750 * Stick_Edit_Zoom
        .Top = frmStickGame.height \ 2
        .Visible = True
        
        .BackColor = Current_BackCol
    End With
    
    .otBox(i - 1).BackColor = def_BackCol
End With

cmdRemovetBox.Enabled = True
End Sub
Private Sub cmdRemovetBox_Click()
Dim i As Integer

With frmStickGame
    .DragEnd
    
    i = .otBox.UBound
    If i Then
        Unload .otBox(i)
        
        If i > 1 Then
            .otBox(i - 1).BackColor = Current_BackCol
        End If
    End If
    i = i - 1
    
    If i <= 0 Then
        cmdRemovetBox.Enabled = False
    End If
End With

End Sub

'#########################################################################################################

Private Sub cmdAddPlatform_Click()
Dim i As Integer

With frmStickGame
    i = .oPlatform.UBound + 1
    
    Load .oPlatform(i)
    With .oPlatform(i)
        .Left = 375
        .width = 4000 * Stick_Edit_Zoom
        .height = 375 * Stick_Edit_Zoom
        .Top = frmStickGame.height \ 2
        .Visible = True
        
        .BackColor = Current_BackCol
    End With
    
    .oPlatform(i - 1).BackColor = def_Platform_BackCol
End With

cmdRemovePlatform.Enabled = True
End Sub
Private Sub cmdRemovePlatform_Click()
Dim i As Integer

With frmStickGame
    .DragEnd
    
    i = .oPlatform.UBound
    If i Then
        Unload .oPlatform(i)
        
        If i > 1 Then
            .oPlatform(i - 1).BackColor = Current_BackCol
        End If
    End If
    i = i - 1
    
    If i <= 0 Then
        cmdRemovePlatform.Enabled = False
    End If
End With

End Sub

'#########################################################################################################

Private Sub cmdLoad_Click()
Dim sFile As String
Dim bError As Boolean

frmMain.CommonDPath sFile, bError, "Load Map", "Stick Maps (*." & Map_Ext & ")|*." & Map_Ext, modStickGame.GetStickMapPath(), True

If bError = False Then
    frmStickGame.DragEnd
    If frmStickGame.LoadMap(sFile) Then
        ShowCurrentMap
        
        cmdRemovePlatform.Enabled = (frmStickGame.oPlatform.UBound > 0)
        cmdRemoveBox.Enabled = (frmStickGame.oBox.UBound > 0)
        cmdRemovetBox.Enabled = (frmStickGame.otBox.UBound > 0)
        frmStickGame.setMapChanged False
    End If
End If

End Sub

Public Sub ShowCurrentMap()
Dim i As Integer


For i = 1 To frmStickGame.oPlatform.UBound
    Unload frmStickGame.oPlatform(i)
Next i
For i = 1 To frmStickGame.otBox.UBound
    Unload frmStickGame.otBox(i)
Next i
For i = 1 To frmStickGame.oBox.UBound
    Unload frmStickGame.oBox(i)
Next i



For i = 0 To ubdPlatforms
    If i Then
        Load frmStickGame.oPlatform(i)
    End If
    
    With frmStickGame.oPlatform(i)
        .Visible = True
        .Left = Platform(i).Left * Stick_Edit_Zoom
        .Top = Platform(i).Top * Stick_Edit_Zoom
        .width = Platform(i).width * Stick_Edit_Zoom
        .height = Platform(i).height * Stick_Edit_Zoom
        
        If (i = ubdPlatforms) And (ubdPlatforms > 0) Then
            .BackColor = Current_BackCol
        Else
            .BackColor = def_Platform_BackCol
        End If
    End With
Next i

For i = 0 To ubdBoxes
    If i Then
        Load frmStickGame.oBox(i)
    End If
    
    With frmStickGame.oBox(i)
        .Visible = True
        .Left = Box(i).Left * Stick_Edit_Zoom
        .Top = Box(i).Top * Stick_Edit_Zoom
        .width = Box(i).width * Stick_Edit_Zoom
        .height = Box(i).height * Stick_Edit_Zoom
        
        If i = ubdBoxes And ubdBoxes > 0 Then
            .BorderColor = Current_BackCol
        Else
            .BackColor = def_BackCol
        End If
    End With
Next i

For i = 0 To ubdtBoxes
    If i Then
        Load frmStickGame.otBox(i)
    End If
    
    With frmStickGame.otBox(i)
        .Visible = True
        .Left = tBox(i).Left * Stick_Edit_Zoom
        .Top = tBox(i).Top * Stick_Edit_Zoom
        .width = tBox(i).width * Stick_Edit_Zoom
        .height = tBox(i).height * Stick_Edit_Zoom
        
        If i = ubdtBoxes And ubdtBoxes > 0 Then
            .BackColor = Current_BackCol
        Else
            .BackColor = def_BackCol
        End If
    End With
Next i

frmStickGame.Show_shpHealthPack_Pos

End Sub

Private Sub cmdReset_Click()
frmStickGame.ResetEditPlatforms
frmStickGame.setMapChanged False
End Sub

'#########################################################################################################

Private Sub cmdSave_Click()
Dim sFile As String, StickMapPath As String
Dim bError As Boolean

StickMapPath = modStickGame.GetStickMapPath()

Retry:
frmMain.CommonDPath sFile, bError, "Save Map", "Stick Maps (*." & Map_Ext & ")|*." & Map_Ext, StickMapPath

If bError = False Then
    
    If LCase$(sFile) = (LCase$(StickMapPath) & "default.map") Then
        MsgBoxEx "You can't overwrite the default map, save it as something else", _
            "The default map can't be overwritten, choose a different filename", vbExclamation, _
            "Map Save Error", , , , , Me.hWnd
        
        GoTo Retry
    Else
        MakeCurrentMap
        If frmStickGame.SaveMap(sFile) Then
            'MsgBoxEx "Map Saved", "Map has been saved to " & sFile, vbInformation, "Saved Map"
            modSpeech.Say "Mad Saved", , , True
            frmStickGame.setMapChanged False
        End If
    End If
End If

End Sub

Private Sub MakeCurrentMap()
Dim i As Integer


ubdPlatforms = frmStickGame.oPlatform.UBound
ubdtBoxes = frmStickGame.otBox.UBound
ubdBoxes = frmStickGame.oBox.UBound

ReDim Platform(0 To ubdPlatforms)
ReDim tBox(0 To ubdtBoxes)
ReDim Box(0 To ubdBoxes)


For i = 0 To ubdPlatforms
    With frmStickGame.oPlatform(i)
        Platform(i).Left = Round(.Left / Stick_Edit_Zoom, 1)
        Platform(i).Top = Round(.Top / Stick_Edit_Zoom, 1)
        Platform(i).width = Round(.width / Stick_Edit_Zoom, 1)
        Platform(i).height = Round(.height / Stick_Edit_Zoom, 1)
    End With
Next i

For i = 0 To ubdtBoxes
    With frmStickGame.otBox(i)
        tBox(i).Left = Round(.Left / Stick_Edit_Zoom, 1)
        tBox(i).Top = Round(.Top / Stick_Edit_Zoom, 1)
        tBox(i).width = Round(.width / Stick_Edit_Zoom, 1)
        tBox(i).height = Round(.height / Stick_Edit_Zoom, 1)
    End With
Next i

For i = 0 To ubdBoxes
    With frmStickGame.oBox(i)
        Box(i).Left = Round(.Left / Stick_Edit_Zoom, 1)
        Box(i).Top = Round(.Top / Stick_Edit_Zoom, 1)
        Box(i).width = Round(.width / Stick_Edit_Zoom, 1)
        Box(i).height = Round(.height / Stick_Edit_Zoom, 1)
    End With
Next i

With frmStickGame
    .HealthPackX = (.shHealthPack.Left + .shHealthPack.width / 2) / Stick_Edit_Zoom
    .HealthPackY = (.shHealthPack.Top + .shHealthPack.height / 2) / Stick_Edit_Zoom
End With

End Sub

'#########################################################################################################

Private Sub Form_Load()
FormLoad Me, , , False

Me.Left = frmStickGame.Left + frmStickGame.width
Me.Top = frmStickGame.Top + frmStickGame.height / 2 - Me.height / 2

If (Me.Left + Me.width) > Screen.width Then
    Me.Left = Screen.width - Me.width
End If

InitEdit


If LenB(modStickGame.StickMapPath) Then
    If frmStickGame.LoadMap(modStickGame.StickMapPath) Then
        ShowCurrentMap
        
        cmdRemovePlatform.Enabled = (frmStickGame.oPlatform.UBound > 0)
        cmdRemoveBox.Enabled = (frmStickGame.oBox.UBound > 0)
        cmdRemovetBox.Enabled = (frmStickGame.otBox.UBound > 0)
    Else
        MsgBox "Error Loading Map - Reseting Platforms", vbExclamation, "Stick Map Editor"
        
        frmStickGame.ResetEditPlatforms
    End If
End If


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FormLoad Me, True
End Sub

Private Sub InitEdit()

ReDim Platform(0)
ReDim Box(0)
ReDim tBox(0)

ubdPlatforms = -1
ubdBoxes = -1
ubdtBoxes = -1

AddPlatform -1000, 16000, 52000, 855
AddBox 5587, 4530, 135, 1095
AddtBox 7200, 10905, 375, 495

End Sub

Private Sub AddPlatform(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)

ubdPlatforms = ubdPlatforms + 1

ReDim Preserve Platform(ubdPlatforms)
With Platform(ubdPlatforms)
    .Left = lLeft
    .Top = lTop
    .width = lWidth
    .height = lHeight
End With


If ubdPlatforms > 0 Then
    Load frmStickGame.oPlatform(ubdPlatforms)
End If

With frmStickGame.oPlatform(ubdPlatforms)
    .Left = lLeft * Stick_Edit_Zoom
    .Top = lTop * Stick_Edit_Zoom
    .height = lHeight * Stick_Edit_Zoom
    .width = lWidth * Stick_Edit_Zoom
    .Visible = True
End With

End Sub
Private Sub AddtBox(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)

ubdtBoxes = ubdtBoxes + 1

ReDim Preserve tBox(ubdtBoxes)
With tBox(ubdtBoxes)
    .Left = lLeft
    .Top = lTop
    .width = lWidth
    .height = lHeight
End With

If ubdtBoxes > 0 Then
    Load frmStickGame.otBox(ubdtBoxes)
End If

With frmStickGame.otBox(ubdtBoxes)
    .Left = lLeft * Stick_Edit_Zoom
    .Top = lTop * Stick_Edit_Zoom
    .height = lHeight * Stick_Edit_Zoom
    .width = lWidth * Stick_Edit_Zoom
    .Visible = True
End With

End Sub
Private Sub AddBox(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)

ubdBoxes = ubdBoxes + 1

ReDim Preserve Box(ubdBoxes)
With Box(ubdBoxes)
    .Left = lLeft
    .Top = lTop
    .width = lWidth
    .height = lHeight
End With


If ubdBoxes > 0 Then
    Load frmStickGame.oBox(ubdBoxes)
End If

With frmStickGame.otBox(ubdBoxes)
    .Left = lLeft * Stick_Edit_Zoom
    .Top = lTop * Stick_Edit_Zoom
    .height = lHeight * Stick_Edit_Zoom
    .width = lWidth * Stick_Edit_Zoom
    .Visible = True
End With

End Sub
