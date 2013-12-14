Attribute VB_Name = "modAlert"
Option Explicit

Public Enum eAlertTypes
    eMSN = 0
    eGMail = 1
    eFlat = 2
End Enum


'for use by alerts
Public TB_Height As Long

'setting
Public AlertStyle As eAlertTypes
Public bBalloonTips As Boolean

Public Sub Init()
TB_Height = GetTaskbarHeight()
End Sub

Public Sub ShowAlert(ByVal sTitle As String, ByVal sCaption As String)

pShowAlert sTitle, sCaption, AlertStyle

End Sub

Private Sub pShowAlert(ByVal sTitle As String, ByVal sCaption As String, ByVal eType As eAlertTypes)
Dim frmAlert As Form

If eType = eMSN Then
    Set frmAlert = New frmMSN6
ElseIf eType = eGMail Then
    Set frmAlert = New frmGmail
Else
    Set frmAlert = New frmFlatAlert
End If

Load frmAlert

frmAlert.sCaption = sCaption
frmAlert.sTitle = sTitle

SetOnTop frmAlert.hWnd '+show

Set frmAlert = Nothing

modAudio.PlaySysSound sys_MailBeep

End Sub

