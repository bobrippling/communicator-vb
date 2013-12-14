Attribute VB_Name = "modSound"
Option Explicit

Private Declare Function pPlaySound Lib "winmm.dll" Alias _
    "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

'Private Const cSndSYNC = &H0, _
            cSndASYNC = &H1, _
            cSndNODEFAULT = &H2, _
            cSndLOOP = &H8, _
            cSndNOSTOP = &H10
            
            '*-------------------------------------*
'* Playsound flags: store in dwFlags   *
'*-------------------------------------*
' lpszName points to a registry entry
' Do not use SND_RESOURSE or SND_FILENAME
'Private Const SND_ALIAS& = &H10000'
' Playsound returns immediately
' Do not use SND_SYNC
'Private Const SND_ASYNC& = &H1
' The name of a wave file.
' Do not use with SND_RESOURCE or SND_ALIAS
'Private Const SND_FILENAME& = &H20000
' Unless used, the default beep will
' play if the specified resource is missing
'Private Const SND_NODEFAULT& = &H2
' Fail the call & do not wait for
' a sound device if it is otherwise unavailable
'Private Const SND_NOWAIT& = &H2000
' Use a resource file as the source.
' Do not use with SND_ALIAS or SND_FILENAME
'Private Const SND_RESOURCE& = &H40004
' Playsound will not return until the
' specified sound has played.  Do not
' use with SND_ASYNC
'Private Const SND_SYNC& = &H0

Private Const SND_ASYNC As Long = &H1
Private Const SND_MEMORY As Long = &H4
Private Const SND_NODEFAULT As Long = &H2

Private Const ResFlags = SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

'Private Declare Function PlaySound& Lib "winmm.dll" Alias _
    "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, _
    ByVal dwFlags As Long)

' hModule is only used if SND_RESOURCE& is set and represents
' an HINSTANCE handle.  This example doesn't support playing
' from a resource file.

' Plays sounds from the registry or a disk file
' Doesn't care if the file is missing
Public Function PlaySound(ByVal ResNo As Integer) As Boolean
  
'PlaySound = (PlaySound(filenamename, 0&, _
    SND_ASYNC Or SND_NODEFAULT Or SND_RESOURCE) <> 0)

Dim Ret As Long

Ret = pPlaySound(StrConv(CStr(LoadResData(ResNo, "WAVE")), vbUnicode), ResFlags)

Pause 1

PlaySound = (Ret = 1)

End Function

Public Function StopSound() As Long
pPlaySound vbNullString, 0
End Function
