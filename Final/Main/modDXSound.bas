Attribute VB_Name = "modDXSound"
Option Explicit


Public Const LeftPan = -10000 'dimensionless. i think
Public Const RightPan = 10000
Public Const CenterPan = 0

Public Const MaxVolume = 0 'decibels
Public Const MinVolume = -10000

Public Const MaxFrequency = 100000 'hertz
Public Const MinFrequency = 100

'DirectX Variables
Private DXMain As DirectX7
Private DS As DirectSound

'User defined type to determine a buffer's capabilities
Private Type BufferCaps
    Volume As Boolean               'Can this buffer's volume be changed?
    Frequency As Boolean            'Can the frequency be altered?
    Pan As Boolean                  'Can we pan the sound from left to right?
    Loop As Boolean                 'Is this sound looping?
    Delete As Boolean               'Should this sound be deleted after playing?
End Type


'User defined type to contain sound data
Public Enum eSoundStates
    s_Empty = 0
    s_Stopped = 1
    s_Paused = 2
    s_Playing = 3
End Enum
Private Type ptSound
    DSBuffer As DirectSoundBuffer 'The buffer that contains the sound
    DSState As eSoundStates         'Describes the current state of the buffer (ie. "Playing", "Stopped")
    DSNotification As Long          'Contains the event reference returned by the DirectX7 object
    DSCaps As BufferCaps            'Describes the buffer's capabilities
    DSSourceName As String          'The name of the source file
    DSFile As Boolean               'Is the source in a seperate file?
    DSResource As Boolean           'Or is it in a resource?
    DSEmpty As Boolean              'Is this SoundArray index empty?
    
    Start_Freq As Long
End Type

Private Sound() As ptSound          'Contains all the data needed for sound manipulation

'Wave Format Setting Contants
Private Const NumChannels = 2              'How many channels will we be playing on?
Private Const SamplesPerSecond = 22050     'How many cycles per second (hertz)?
Private Const BitsPerSample = 16           'What bit-depth will we use?

Public Function DXSound_Init(ByVal hWnd As Long, ByRef sError As String) As Boolean
Dim sTmpError As String

'If we can't initialize properly, trap the error
On Error GoTo DX_Error
sTmpError = "Error Initialising DirectX"
Set DXMain = New DirectX7
sTmpError = "Error Initialising Direct Sound"
Set DS = DXMain.DirectSoundCreate(vbNullString)

'Set the DirectSound object's cooperative level (Priority gives us sole control)
DS.SetCooperativeLevel hWnd, DSSCL_PRIORITY

'Initialize our Sound array to zero
ReDim Sound(0)
Sound(0).DSEmpty = True
Sound(0).DSState = s_Empty


DXSound_Init = True

sTmpError = vbNullString
sError = vbNullString

'Exit sub before the error code
Exit Function

DX_Error:
Set DS = Nothing
Set DXMain = Nothing

Erase Sound

If LenB(sTmpError) Then
    sTmpError = sTmpError & " (" & GetDXError() & ")"
Else
    sTmpError = GetDXError()
    If LenB(sTmpError) = 0 Then
        sTmpError = "Unknown Error...?"
    End If
End If

DXSound_Init = False
End Function
Private Function GetDXError() As String
Dim sTmpError As String
Dim i As Integer

sTmpError = Err.Description
If LenB(sTmpError) Then
    GetDXError = sTmpError
Else
    i = Err.LastDllError
    If i <> 0 Then
        GetDXError = CStr(i)
    Else
        GetDXError = vbNullString
    End If
End If

End Function

Public Function LoadSound(SourceName As String, IsFile As Boolean, IsResource As Boolean, _
    IsFrequency As Boolean, IsPan As Boolean, IsVolume As Boolean, _
    IsLoop As Boolean, FormObject As Form) As Integer

Dim i As Integer
Dim Index As Integer
Dim DSBufferDescription As DSBUFFERDESC
Dim DSFormat As DxVBLib.WAVEFORMATEX
Dim DSPosition(0) As DSBPOSITIONNOTIFY


'Search the sound array for any empty spaces
Index = -1
For i = 0 To UBound(Sound)
    If Sound(i).DSEmpty = True Then 'If there is an empty space, us it
        Index = i
        Exit For
    End If
Next i

If Index = -1 Then                  'If there's no empty space, make a new spot
    ReDim Preserve Sound(UBound(Sound) + 1)
    Index = UBound(Sound)
End If
LoadSound = Index                   'Set the return value of the function


'Load the Sound array with the data given
With Sound(Index)
    .DSEmpty = False                'This Sound(index) is now occupied with data
    .DSFile = IsFile                'Is this sound to be loaded from a file?
    .DSResource = IsResource        'Or is it to be loaded from a resource?
    .DSSourceName = SourceName      'What is the name of the source?
    .DSState = s_Stopped            'Set the current state to "Stopped"
    .DSCaps.Delete = False 'IsDelete       'Is this sound to be deleted after it is played?
    .DSCaps.Frequency = IsFrequency 'Is this sound to have frequency altering capabilities?
    .DSCaps.Loop = IsLoop           'Is this sound to be looped?
    .DSCaps.Pan = IsPan             'Is this sound to have Left and Right panning capabilities?
    .DSCaps.Volume = IsVolume       'Is this sound capable of altered volume settings?
End With

'Set the buffer description according to the data provided
With DSBufferDescription
    'If Sound(Index).DSCaps.Delete Then .lFlags = .lFlags Or DSBCAPS_CTRLPOSITIONNOTIFY
    If Sound(Index).DSCaps.Frequency Then .lFlags = .lFlags Or DSBCAPS_CTRLFREQUENCY
    If Sound(Index).DSCaps.Pan Then .lFlags = .lFlags Or DSBCAPS_CTRLPAN
    If Sound(Index).DSCaps.Volume Then .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
End With

'Set the Wave Format
With DSFormat
    .nFormatTag = WAVE_FORMAT_PCM
    .nChannels = NumChannels
    .lSamplesPerSec = SamplesPerSecond
    .nBitsPerSample = BitsPerSample
    .nBlockAlign = .nBitsPerSample / 8 * .nChannels
    .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
End With

'Load the sound into the buffer
If Sound(Index).DSFile Then
    
    On Error Resume Next
    Set Sound(Index).DSBuffer = DS.CreateSoundBufferFromFile(Sound(Index).DSSourceName, DSBufferDescription, DSFormat)
    
    
ElseIf Sound(Index).DSResource Then
    
    Set Sound(Index).DSBuffer = DS.CreateSoundBufferFromResource(vbNullString, Sound(Index).DSSourceName, DSBufferDescription, DSFormat)
    
End If


''If the sound is to be deleted after it plays, we must create an event for it
'If Sound(Index).DSCaps.Delete Then
'    Sound(Index).DSNotification = DXMain.CreateEvent(FormObject)        'Make the event (has to be created in a Form Object) and get its handle
'
'    DSPosition(0).hEventNotify = Sound(Index).DSNotification        'Place this event handle in an DSBPOSITIONNOTIFY variable
'    DSPosition(0).lOffset = DSBPN_OFFSETSTOP                        'Define the position within the wave file at which you would like the event to be triggered
'
'    Sound(Index).DSBuffer.SetNotificationPositions 1, DSPosition()  'Set the "notification position" by passing the DSBPOSITIONNOTIFY variable
'End If


Sound(Index).Start_Freq = GetFrequency(Index)


Erase DSPosition

End Function

Public Sub RemoveSound(Index As Integer)

''Destroy the event associated with the ending of this sound, if there was one
'If Sound(Index).DSCaps.Delete And Sound(Index).DSNotification <> 0 Then
'    DXMain.DestroyEvent Sound(Index).DSNotification
'End If

'Reset all the variables in the sound array
With Sound(Index)
    Set .DSBuffer = Nothing
    .DSCaps.Delete = False
    .DSCaps.Frequency = False
    .DSCaps.Loop = False
    .DSCaps.Pan = False
    .DSCaps.Volume = False
    .DSEmpty = True
    .DSFile = False
    .DSNotification = 0
    .DSResource = False
    .DSSourceName = vbNullString
    .DSState = s_Empty
    
    .Start_Freq = 0
End With
    
End Sub

Public Sub PlaySound(Index As Integer)

'Check to make sure there is a sound loaded in the specified buffer

If Sound(Index).DSEmpty Then Exit Sub

'If the sound is not "paused" then reset it's position to the beginning
If Sound(Index).DSState <> s_Paused Then Sound(Index).DSBuffer.SetCurrentPosition 0

'Play looped or singly, as appropriate
'If Sound(Index).DSCaps.Loop Then
    'Sound(Index).DSBuffer.Play DSBPLAY_LOOPING
'Else
    Sound(Index).DSBuffer.Play DSBPLAY_DEFAULT
'End If


'Set the state to "playing"
Sound(Index).DSState = s_Playing

End Sub

Public Sub StopSound(Index As Integer)

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Sub

'Stop the buffer and reset to the beginning
Sound(Index).DSBuffer.Stop
Sound(Index).DSBuffer.SetCurrentPosition 0

Sound(Index).DSState = s_Stopped

End Sub

Public Sub PauseSound(Index As Integer)

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Sub

'Stop the buffer
Sound(Index).DSBuffer.Stop
Sound(Index).DSState = s_Paused

End Sub

Public Sub SetFrequency(Index As Integer, Freq As Long)

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then
    'Debug.Print "Couldn't change freq 1 - " & Index
    Exit Sub
End If

'Check to make sure that the buffer has the capability of altering its frequency
If Sound(Index).DSCaps.Frequency = False Then
    'Debug.Print "Couldn't change freq 2 - " & Index
    Exit Sub
End If

'Alter the frequency according to the Freq provided
Sound(Index).DSBuffer.SetFrequency Freq

End Sub

Public Sub SetRelativeFrequency(Index As Integer, Multiple As Single)

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then
    Exit Sub
End If

'Check to make sure that the buffer has the capability of altering its frequency
If Sound(Index).DSCaps.Frequency = False Then
    Exit Sub
End If

'Alter the frequency according to the Freq provided
On Error Resume Next 'in case the buffer isn't set at load/init
Sound(Index).DSBuffer.SetFrequency Sound(Index).Start_Freq * Multiple

End Sub

Public Sub SetVolume(Index As Integer, Vol As Long)

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Sub

'Check to make sure that the buffer has the capability of altering its volume
If Sound(Index).DSCaps.Volume = False Then Exit Sub

'Alter the volume according to the Vol provided
Sound(Index).DSBuffer.SetVolume Vol

End Sub

Public Sub SetPan(Index As Integer, Pan As Long)

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Sub

'Check to make sure that the buffer has the capability of altering its pan
If Sound(Index).DSCaps.Pan = False Then Exit Sub

'Alter the pan according to the Pan provided
Sound(Index).DSBuffer.SetPan Pan

End Sub

Public Function GetFrequency(Index As Integer) As Long

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Function

'Check to make sure that the buffer has the capability of altering its frequency
If Sound(Index).DSCaps.Frequency = False Then Exit Function

'Return the frequency value
GetFrequency = Sound(Index).DSBuffer.GetFrequency()

End Function

Public Function GetVolume(Index As Integer) As Long

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Function

'Check to make sure that the buffer has the capability of altering its volume
If Sound(Index).DSCaps.Volume = False Then Exit Function

'Return the volume value
GetVolume = Sound(Index).DSBuffer.GetVolume()

End Function

Public Function GetPan(Index As Integer) As Long

'Check to make sure there is a sound loaded in the specified buffer
If Sound(Index).DSEmpty Then Exit Function

'Check to make sure that the buffer has the capability of altering its pan
If Sound(Index).DSCaps.Pan = False Then Exit Function

'Return the pan value
GetPan = Sound(Index).DSBuffer.GetPan()

End Function

Public Function GetSoundState(Index As Integer) As eSoundStates
Dim vStat As CONST_DSBSTATUSFLAGS


'Returns the current state of the given sound
GetSoundState = Sound(Index).DSState

'GetSoundState = IIf(Sound(Index).DSBuffer.GetStatus() = DSBSTATUS_PLAYING, eSoundStates.s_Playing, eSoundStates.s_Stopped)


End Function

Public Property Get nSounds() As Integer
On Error GoTo EH
nSounds = UBound(Sound)
Exit Property
EH:
nSounds = -1
End Property

Public Function DXCallback(ByVal EventID As Long) As Integer

Dim i As Integer

'Find the sound that caused this event to be triggered
For i = 0 To UBound(Sound)
    If Sound(i).DSNotification = EventID Then
        Exit For
    End If
Next i

'Return the ID
DXCallback = i

End Function

Public Sub DXSound_Terminate()

Dim i As Integer

If SoundArrayInitited() Then
    On Error GoTo Cont
    
    'Delete all of the sounds created
    For i = 0 To UBound(Sound)
        RemoveSound i
    Next i
    
End If

Cont:

Erase Sound

Set DS = Nothing
Set DXMain = Nothing

End Sub

Private Function SoundArrayInitited() As Boolean
Dim i As Integer

On Local Error GoTo EH

i = UBound(Sound)

SoundArrayInitited = True
Exit Function
EH:
SoundArrayInitited = False
End Function
