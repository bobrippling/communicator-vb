Attribute VB_Name = "modPorts"
Option Explicit

Public Const DefaultMainPort As Integer = 2858
Private pMainPort As Integer

Public Const DefaultStickPort As Integer = 28808
Private pStickPort As Integer

Public Const DefaultSpacePort As Integer = 28807
Private pSpacePort As Integer

Public Const DefaultFTPort As Integer = 28802
Private pFTPort As Integer

Public Const DefaultDPPort As Integer = 28801
Private pDPPort As Integer

Public Const DefaultVoicePort As Integer = 28803
Private pVoicePort As Integer

Public Sub Init()

MainPort = DefaultMainPort
StickPort = DefaultStickPort
SpacePort = DefaultSpacePort
FTPort = DefaultFTPort
DPPort = DefaultDPPort
VoicePort = DefaultVoicePort

End Sub

Private Sub LetPort(ByRef iCurrentPort As Integer, iNewPort As Integer)

If 1 <= iNewPort And iNewPort <= 65535 Then
    iCurrentPort = iNewPort
End If

End Sub

Public Property Get MainPort() As Integer
MainPort = pMainPort
End Property
Public Property Let MainPort(ByVal iMainPort As Integer)
LetPort pMainPort, iMainPort
End Property

Public Property Get StickPort() As Integer
StickPort = pStickPort
End Property
Public Property Let StickPort(ByVal iStickPort As Integer)
LetPort pStickPort, iStickPort
End Property

Public Property Get SpacePort() As Integer
SpacePort = pSpacePort
End Property
Public Property Let SpacePort(ByVal iSpacePort As Integer)
LetPort pSpacePort, iSpacePort
End Property

Public Property Get FTPort() As Integer
FTPort = pFTPort
End Property
Public Property Let FTPort(ByVal iFTPort As Integer)
LetPort pFTPort, iFTPort
End Property

Public Property Get DPPort() As Integer
DPPort = pDPPort
End Property
Public Property Let DPPort(ByVal iDPPort As Integer)
LetPort pDPPort, iDPPort
End Property

Public Property Get VoicePort() As Integer
VoicePort = pVoicePort
End Property
Public Property Let VoicePort(ByVal iVoicePort As Integer)
LetPort pVoicePort, iVoicePort
End Property
