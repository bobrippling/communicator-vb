VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'XML Tags:
'http://msdn.microsoft.com/en-us/library/bb483084.aspx

'<VOICE REQUIRED="NAME=Microsoft Mike">This is the Microsoft Mike voice.</VOICE>
'
'<RATE SPEED="-5">   This is slow speech.         </RATE>
'<PITCH MIDDLE="5">  This is high pitched speech. </PITCH>
'                    This is a pause.             <SILENCE MSEC ="500"/>
'<VOLUME LEVEL="50"> This is quite speech.        </VOLUME>
'<EMPH>              This                         </EMPH> is an emphasis.
'<SILENCE MSEC ="250"/>


Private WithEvents pVoice As SpeechLib.SpVoice
Attribute pVoice.VB_VarHelpID = -1

Public Enum eSpeechStatus
    sIdle = 0
    sSpeaking = 1
End Enum

Private pPitch As Integer

'Private Const SPF_DEFAULT = 0
'Private Const SPF_ASYNC = 1
'Private Const SPF_PURGEBEFORESPEAK = 2
'Private Const SPF_IS_FILENAME = 3
'Private Const SPF_IS_XML = 4
'Private Const SPF_IS_NOT_XML = 5
'Private Const SPF_PERSIST_XML = 6

Public Sub pGetVoices()
Dim arVoices As SpeechLib.ISpeechObjectTokens
Dim i As Integer

'http://www.codeguru.com/vb/gen/vb_misc/samples/article.php/c13893/

Set arVoices = pVoice.GetVoices

ReDim modSpeech.VoiceAr(0 To arVoices.Count - 1)

For i = 0 To arVoices.Count - 1
    Set modSpeech.VoiceAr(i).tVoice = arVoices.Item(i)
    
    modSpeech.VoiceAr(i).sDesc = modSpeech.VoiceAr(i).tVoice.GetDescription()
Next i

Set arVoices = Nothing

End Sub

Public Sub pSetVoice(ByVal iVoice As Integer)
'Dim arVoices As SpeechLib.ISpeechObjectTokens

'Set arVoices = pVoice.GetVoices

'pVoice.Voice = arVoices.Item(iVoice).CreateInstance

On Error Resume Next
Set pVoice.Voice = pVoice.GetVoices.Item(iVoice)

End Sub

Public Sub pStopSpeech()

pVoice.Speak vbNullString, SVSFPurgeBeforeSpeak

End Sub

Public Sub pSay(ByVal Text As String)

With pVoice
    '.WaitUntilDone 500
    '.Speak Text, SVSFlagsAsync
    
    .Speak "<PITCH ABSMIDDLE=""" & pPitch & """>" & _
        Text & _
        "</PITCH>", _
        SVSFlagsAsync Or SVSFIsXML
    
End With

End Sub

Private Sub Class_Initialize()

'AddConsoleText "Speech Initialised, Initialising Voice..."
Set pVoice = New SpVoice

pVoice.Volume = 100

'AddConsoleText "Voice Initialised"

End Sub

Private Sub Class_Terminate()

'AddConsoleText "Terminating Speech..."
Set pVoice = Nothing
'AddConsoleText "Voice Terminated"
'AddConsoleText "Speech Terminated"

End Sub

'##########
'properties ----------------------------------
'##########
Public Property Let Vol(ByVal V As Integer)
pVoice.Volume = V
End Property
Public Property Get Vol() As Integer
Vol = pVoice.Volume
End Property

Public Property Let Speed(ByVal S As Integer)
pVoice.Rate = S
End Property
Public Property Get Speed() As Integer
Speed = pVoice.Rate
End Property

Public Property Let Pitch(P As Integer)
'pPitch = P
'pVoice.Speak "<Pitch AbsMiddle=""" & P & """ />", SVSFIsXML Or SVSFlagsAsync
pPitch = P

'-10<=p<=10
End Property
Public Property Get Pitch() As Integer
Pitch = pPitch
End Property

'get only (language + vendor are included)
Public Property Get Description() As String
Description = pVoice.Voice.GetDescription()
End Property

Public Property Get Attr(ByVal A As String) As String
Attr = pVoice.Voice.GetAttribute(A)
End Property

Public Property Get SupportsAttr(ByVal A As String) As Boolean
SupportsAttr = pVoice.Voice.MatchesAttributes(A)
End Property

Public Property Get SpeechStatus() As eSpeechStatus
If pVoice.WaitUntilDone(1) = True Then 'if it hasn't waited until it's finished speaking
    SpeechStatus = sIdle               'in one ms, then it is still speaking
Else
    SpeechStatus = sSpeaking
End If
End Property
