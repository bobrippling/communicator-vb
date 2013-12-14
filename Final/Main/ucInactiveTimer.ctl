VERSION 5.00
Begin VB.UserControl ucInactiveTimer 
   BackColor       =   &H80000002&
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   750
   ScaleWidth      =   885
   Begin VB.Timer tmrInactive 
      Interval        =   30000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "ucInactiveTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As LASTINPUTINFO) As Long


Public Event UserInactive()

Private m_InactiveInterval As Long
Private Const def_InactiveInterval As Long = 1000& * 5& * 60&

Private Sub UserControl_InitProperties()
m_InactiveInterval = def_InactiveInterval
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Me.Enabled = PropBag.ReadProperty("Enabled", True)
Me.InactiveInterval = PropBag.ReadProperty("InactiveInterval", def_InactiveInterval)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Enabled", Me.Enabled, "True"
PropBag.WriteProperty "InactiveInterval", Me.InactiveInterval, def_InactiveInterval
End Sub

Private Sub UserControl_Resize()
UserControl.width = ScaleX(32, vbPixels, vbTwips)
UserControl.height = ScaleX(32, vbPixels, vbTwips)
End Sub

Public Property Get InactiveInterval() As Long
InactiveInterval = m_InactiveInterval
End Property
Public Property Let InactiveInterval(ByVal bVal As Long)
m_InactiveInterval = bVal
PropertyChanged "InactiveInterval"
End Property

' Delegate the Enabled property to the timer.
Public Property Get Enabled() As Boolean
Enabled = tmrInactive.Enabled
End Property
Public Property Let Enabled(ByVal bVal As Boolean)
tmrInactive.Enabled = bVal
End Property

Private Function ElapsedIdleTime() As Long
Dim m_lii As LASTINPUTINFO

m_lii.cbSize = Len(m_lii)

If GetLastInputInfo(m_lii) = 0 Then
    'Err.Raise vbObjectError + 1001, "InactiveTimer", _
        "Error getting last input information"
    ElapsedIdleTime = -1
Else
    ElapsedIdleTime = GetTickCount() - m_lii.dwTime
End If

End Function

Private Sub tmrInactive_Timer()
If Not UserControl.Ambient.UserMode Then Exit Sub

If ElapsedIdleTime() > m_InactiveInterval Then
    'Debug.Print "Inactive"
    RaiseEvent UserInactive
End If

End Sub

