VERSION 5.00
Begin VB.Form frmStickWeapons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Allowed Weapons"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   7350
   Begin VB.CheckBox chkChopper 
      Caption         =   "Allow Chopper"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "<  Allow Weapon"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdBan 
      Caption         =   "Ban Weapon >"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox lstBanned 
      Height          =   2595
      Left            =   4560
      TabIndex        =   5
      Top             =   0
      Width           =   2775
   End
   Begin VB.ListBox lstAllowed 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmStickWeapons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
RefreshList

chkChopper.Value = Abs(modStickGame.sv_AllowedWeapons(eWeaponTypes.Chopper))

cmdBan.Enabled = False
cmdAllow.Enabled = False

Stick_FormLoad Me
Me.Left = frmStickOptions.Left + frmStickOptions.width / 2 - Me.width / 2
Me.Top = frmStickOptions.Top + frmStickOptions.height / 2 - Me.height / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Stick_FormLoad Me, True
If modStickGame.StickOptionFormLoaded Then
    SetFocus2 frmStickOptions
End If
End Sub

'############################################################################################################

Private Sub RefreshList()
Dim i As eWeaponTypes
Dim iAllowed As Integer, iBanned As Integer

iAllowed = lstAllowed.ListIndex
iBanned = lstBanned.ListIndex

lstAllowed.Clear
lstBanned.Clear

For i = 0 To eWeaponTypes.Chopper - 2 'exclude Knife and Chopper
    AddWeaponToList i
Next i

'AddWeaponToList Chopper
'handled separatly to prevent bugzorz

If iAllowed < lstAllowed.ListCount Then
    lstAllowed.ListIndex = iAllowed
Else
    cmdBan.Enabled = False
End If
If iBanned < lstBanned.ListCount Then
    lstBanned.ListIndex = iBanned
Else
    cmdAllow.Enabled = False
End If

End Sub

Private Sub AddWeaponToList(i As eWeaponTypes)

If modStickGame.sv_AllowedWeapons(i) Then
    lstAllowed.AddItem GetWeaponName(i)
Else
    lstBanned.AddItem GetWeaponName(i)
End If

End Sub

Private Sub cmdBan_Click()

If lstAllowed.ListCount <= 2 Then
    lblInfo.Caption = "You must have" & vbNewLine & "at least two weapons"
    cmdBan.Enabled = False
Else
    BanUnBanWeapon False
    lblInfo.Caption = vbNullString
End If

End Sub

Private Sub cmdAllow_Click()

BanUnBanWeapon True
lblInfo.Caption = vbNullString

End Sub

Private Sub chkChopper_Click()
modStickGame.sv_AllowedWeapons(eWeaponTypes.Chopper) = CBool(chkChopper.Value)

frmStickGame.SendServerVarPacket True
End Sub

Private Sub BanUnBanWeapon(bAllow As Boolean)

Dim iWeapon As eWeaponTypes
Dim i As Integer

If bAllow Then
    iWeapon = modStickGame.WeaponNameToInt(lstBanned.Text)
    i = lstBanned.ListIndex
Else
    iWeapon = modStickGame.WeaponNameToInt(lstAllowed.Text)
    i = lstAllowed.ListIndex
End If


If iWeapon > -1 Then
    modStickGame.sv_AllowedWeapons(CInt(iWeapon)) = bAllow
    
    frmStickGame.SendServerVarPacket True
    
    RefreshList
    
    If bAllow Then
        If i < lstBanned.ListCount Then
            lstBanned.ListIndex = i
        End If
    Else
        If i < lstAllowed.ListCount Then
            lstAllowed.ListIndex = i
        End If
    End If
End If

End Sub

Private Sub lstAllowed_Click()
cmdBan.Enabled = LenB(lstAllowed.Text) And modStickGame.StickServer
End Sub

Private Sub lstBanned_Click()
cmdAllow.Enabled = LenB(lstBanned.Text) And modStickGame.StickServer
End Sub
