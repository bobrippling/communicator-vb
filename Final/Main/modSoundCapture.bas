Attribute VB_Name = "modSoundCap"
Option Explicit

'http://msdn.microsoft.com/en-us/library/ms707311(VS.85).aspx

Private Declare Function api_mciSendString Lib "winmm" Alias "mciSendStringA" ( _
    ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, _
    ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function api_mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" ( _
    ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Const vbSpace = " "
Private plLastError As Long

Private Function mciSendString(sString As String) As Long
plLastError = api_mciSendString(sString, 0&, 0&, 0&)
mciSendString = plLastError
End Function
Private Function mciSetString(sRestOfString As String) As Long
mciSetString = mciSendString("set " & sRestOfString)
End Function
Private Function mciSendString_AndReply(sString As String, sReply As String) As Long
Const kLen As Long = 255
sReply = String$(kLen, 0)
plLastError = api_mciSendString(sString, sReply, kLen, 0&)
mciSendString_AndReply = plLastError
End Function

Public Function GetLastErrorText() As String
Dim sBuffer As String

'create a buffer
sBuffer = Space$(255)

'retrieve the error string
api_mciGetErrorString plLastError, sBuffer, Len(sBuffer)

'strip off the trailing spaces
GetLastErrorText = Trim$(sBuffer)

End Function

'################################################################################

Public Function Set_Stats(sName As String, _
    Optional iChannels As Integer = 2, _
    Optional lSamplePerSec As Long = 44100, _
    Optional iBitsPerSample As Integer = 16) As Boolean

Dim bSuccess As Boolean

If mciSetString(sName & " channels " & CStr(iChannels)) = 0& Then
    If mciSetString(sName & " samplespersec " & CStr(lSamplePerSec)) = 0& Then
        bSuccess = (mciSetString(sName & " bitspersample " & CStr(iBitsPerSample)) = 0&)
    End If
End If

Set_Stats = bSuccess

'Samples Per Second that are supported:
'11025   low quality
'22050   medium quality
'44100 high quality (CD music quality)
'####################################
'Bits per sample is 16 or 8
'####################################
'Channels are 1 (mono) or 2 (stereo)

End Function

Public Function Start_Recording(sName As String) As Boolean
Start_Recording = (mciSendString("record " & sName) = 0&)
End Function
Public Function Stop_Recording(sName As String) As Boolean
Stop_Recording = (mciSendString("stop " & sName) = 0&)
End Function
Public Function Seek_To_Start(sName As String) As Boolean
Seek_To_Start = (mciSendString("seek " & sName & " to start") = 0&)
End Function
'Public Function Get_Record_Pos(sName As String)
'mciSendString_AndReply "status " & sName & " position", ""
'End Function

Public Function Save_Recording(sName As String, sFileName As String) As Boolean
Const Quote As String = """"

Save_Recording = (mciSendString("save " & sName & vbSpace & Quote & sFileName & Quote) = 0&)

End Function
Public Function Save_And_Close_Record(sName As String, sFileName As String) As Boolean

Save_And_Close_Record = Save_Recording(sName, sFileName)

Close_Record sName

FixWaveFile sFileName

End Function

Public Function Open_Record(sName As String) As Boolean
Open_Record = (mciSendString("open new type waveaudio alias " & sName) = 0&)

Set_Stats sName
End Function
Public Function Close_Record(sName As String) As Boolean
Close_Record = (mciSendString("close " & sName) = 0&)
End Function
Public Function Close_All() As Boolean
Close_All = (mciSendString("close all") = 0&)
End Function

Public Function Play_Recording(sName As String) As Boolean
Play_Recording = (mciSendString("play " & sName & " from 0") = 0&)
End Function

'##########################################################################################################

Private Function Get_Record_Length(sName As String, sFormat As String, lResult As Long) As Boolean
Dim sReply As String

mciSendString "set " & sName & " time format " & sFormat 'set format to milliseconds

If mciSendString_AndReply("status " & sName & " length", sReply) = 0& Then
    Get_Record_Length = True
    lResult = CLng(sReply)
Else
    Get_Record_Length = False
End If

End Function
Public Function Get_Record_Length_ms(sName As String, lMs As Long) As Boolean
Get_Record_Length_ms = Get_Record_Length(sName, "ms", lMs)
End Function
Public Function Get_Record_Length_Bytes(sName As String, lBytes As Long) As Boolean
Get_Record_Length_Bytes = Get_Record_Length(sName, "bytes", lBytes)
End Function

'##########################################################################################################

Private Function Get_Record_Status(sName As String, sStatus As String, sResult As String) As Boolean

If mciSendString_AndReply("status " & sName & vbSpace & sStatus, sResult) = 0& Then
    Get_Record_Status = True
Else
    Get_Record_Status = False
End If

End Function
Public Function Get_Record_Channels(sName As String, sChannels As String) As Boolean
Get_Record_Channels = Get_Record_Status(sName, "channels", sChannels)
End Function
Public Function Get_Record_BitsPerSample(sName As String, sBitsPerSample As String) As Boolean
Get_Record_BitsPerSample = Get_Record_Status(sName, "bitspersample", sBitsPerSample)
End Function
Public Function Get_Record_BytesPerSec(sName As String, sBytesPerSec As String) As Boolean
Get_Record_BytesPerSec = Get_Record_Status(sName, "bytespersec", sBytesPerSec)
End Function

'##########################################################################################################

'http://www.rediware.com/programming/vb/vbrecwav/vbrecordwav.htm

Public Sub FixWaveFile(sWav As String, _
    Optional lSamples As Long = 44100, Optional lChannels As Long = 2, Optional lBits As Long = 16)

'this will fix the file so it is playable with WMP
Dim f As Integer
Dim HexCode As String
Dim Hex1 As String
Dim Hex2 As String
Dim Hex3 As String
Dim lByteNum As Long 'byte number (29,30, & 31) in the wave file
Dim bByte As Byte 'will be hex byte to write
Dim lBytes As Long


lBytes = lSamples * (lChannels * lBits) / 8


'get the hexadecimal for the lBytes value
HexCode = Hex$(lBytes) ' lBytes calculated from previous formula
Do While Len(HexCode) < 6 ' make sure the hex code is 6 chars long
    HexCode = "0" & HexCode ' if not, add a zero
Loop


'note: this value had to be written to the file in reverse order!
Hex1 = Right$(HexCode, 2) ' Endian small - reverse order - get last hex byte first
Hex2 = Mid$(HexCode, 3, 2) ' get middle hex byte
Hex3 = Left$(HexCode, 2) ' get first hex byte

'open the file
f = FreeFile() 'get a free file number
Open sWav For Binary Access Write As #f 'binary open file
    
    lByteNum = 29 'first byte to write is 29
    
    bByte = CInt("&H" & Hex1) 'bByte = integer of hex Hex1
    Put #f, lByteNum, bByte 'write bByte value to byte position lByteNum in file
    
    bByte = CInt("&H" & Hex2) 'proceed to write remaining two bytes to consecutive positions
    lByteNum = lByteNum + 1
    Put #f, lByteNum, bByte 'note the Put command for writing bites to binary files
    
    bByte = CInt("&H" & Hex3)
    lByteNum = lByteNum + 1
    Put #f, lByteNum, bByte
    
Close #f

End Sub


