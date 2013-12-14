Attribute VB_Name = "modSpeech"
Option Explicit

Private Const Quote As String = """"

Private cSpeech As clsSpeech

Public Type ptVoice
    tVoice As Speechlib.SpObjectToken
    sDesc As String
End Type

Public VoiceAr() As ptVoice

Public bVoice As Boolean

Public sBalloon As Boolean, sReceived As Boolean, sQuestions As Boolean, sHiBye As Boolean
Public sHi As Boolean, sBye As Boolean, sSayName As Boolean, sGameSpeak As Boolean
Public sOnlyForeground As Boolean, sSayInfo As Boolean, sSent As Boolean

Public bHurgh As Boolean

Private Type ptTagReplacement
    sInStart As String ' don't include the "<"
    sOutStart As String
    endInTag As Boolean
    'ends are implicit - ">"
End Type

Private Type ptReplacement
    sIn As String
    sOut As String
    bWholeWord As Boolean
End Type
Private sReplacements() As ptReplacement
Private sTagReplacements() As ptTagReplacement

Public Sub Say(ByVal sText As String, _
    Optional ByVal iVol As Integer = -1, Optional ByVal ispeed As Integer = 11, _
    Optional ByVal Force As Boolean = False)

'Dim oVol As Integer, ospeed As Integer
Dim can As Boolean

If iVol <> -1 Then 'set the values to our temp ones
    'oVol = Vol
    Vol = iVol
End If
If ispeed <> 11 Then
    'ospeed = speed
    Speed = ispeed
End If

can = bVoice
If Force = False Then
    'can't force it
    If frmMain.mnuFileGameMode.Checked Then
        'if in game mode,
        If modSpeech.sGameSpeak = False Then
            'if can't override,
            can = False
        End If
    End If
    
    If sOnlyForeground Then
        If Not modVars.IsForegroundWindow() Then
            can = False
        End If
    End If
Else
    can = True
End If

If can Then
    If Not (cSpeech Is Nothing) Then
        On Error Resume Next
        
        processReplacements sText
        cSpeech.pSay sText
    End If
End If

End Sub

Public Function String2(n As Integer, S As String) As String
Dim i As Integer
Dim r As String
Dim sLen As Integer

sLen = Len(S) + 1 ' +1 for space

r = Space$(n * sLen)

For i = 0 To n - 1
    Mid$(r, i * sLen + 1) = S & vbSpace
Next i

String2 = Trim$(r)

End Function

Private Sub processReplacements(ByRef sText As String)
Dim i As Integer, iPos As Integer
Dim charLeft As Integer, charRight As Integer

Const powerSymbol As String = "^", MAX_POWER_COUNT As Integer = 12

i = InStr(1, sText, powerSymbol)
While i
    Dim wordBefore As String, wordAfter As String
    Dim n As Integer, j As Integer
    
'    iPos = InStr(i + 1, sText, vbSpace)
'    If iPos Then
'        wordAfter = Mid$(sText, i + 1, iPos - i - 1)
'    Else
'        wordAfter = Mid$(sText, i + 1)
'    End If
    j = i + 1
    While IsNumeric(Mid$(sText, j, 1)) And j <= Len(sText)
        j = j + 1
    Wend
    wordAfter = Mid$(sText, i + 1, j - i - 1)
    
    On Error GoTo Cont
    n = CInt(wordAfter)
    
    If n > 0 And n <= MAX_POWER_COUNT Then
        j = InStrRev(sText, "(", i)
        If j = 0 Then
            j = InStrRev(sText, vbSpace, i)
        End If
        
        If j > 0 Then
            wordBefore = Mid$(sText, j + 1, i - j - 1)
        Else
            wordBefore = Left$(sText, i - 1)
        End If
        
        'sText = Left$(sText, j) & String(n, wordBefore) & Mid$(sText, iPos + n * Len(wordBefore))
        'String() only works with chars
        
        If iPos Then
            sText = Left$(sText, j) & String2(n, wordBefore) & Mid$(sText, iPos)
        Else
            sText = Left$(sText, j) & String2(n, wordBefore)
        End If
    End If
    
    i = InStr(i + 1, sText, powerSymbol)
Wend
    

Cont:
'regex:
's/\<i\>.*\<\/i\>/\<i\>\<emph\>.*\<emph\>\<\/i\>/g
'shell:
'<i>*</i> --> <i><emph>*</emph></i>
i = InStr(sText, "<i>")
While i > 0
    j = InStr(i, sText, "</i>")
    If j > 0 Then
        sText = Left$(sText, i + 2) & "<emph>" & Mid$(sText, i + 3, j - i - 3) & "</emph>" & Mid$(sText, j)
    End If
    i = InStr(i + 1, sText, "<i>")
Wend



For i = LBound(sReplacements) To UBound(sReplacements)
    iPos = InStr(1, sText, sReplacements(i).sIn, vbTextCompare)
    If iPos Then
        If sReplacements(i).bWholeWord Then
            Do
                charLeft = 0
                charRight = 0
                On Error Resume Next
                charLeft = Asc(Mid$(sText, iPos - 1, 1))
                charRight = Asc(Mid$(sText, iPos + Len(sReplacements(i).sIn), 1))
                
                If Not isLetter(charLeft) And Not isLetter(charRight) Then
                    
                    'replace this specific one
                    sText = Left$(sText, iPos - 1) & _
                            sReplacements(i).sOut & _
                            Mid$(sText, iPos + Len(sReplacements(i).sIn))
                    
                    
                End If
                
                iPos = InStr(iPos + 1, sText, sReplacements(i).sIn, vbTextCompare)
            Loop While iPos > 0
        Else
            sText = Replace$(sText, sReplacements(i).sIn, sReplacements(i).sOut, , , vbTextCompare)
        End If
    Else
        j = InStr(1, sReplacements(i).sIn, "*")
        If j > 0 Then
            On Error Resume Next
            If j = 1 Then
                'start of word: [^a-zA-Z]
                j = InStr(1, sText, Mid$(sReplacements(i).sIn, 2), vbTextCompare)
                Do
                    Dim can As Boolean
                    can = False
                    If j > 1 Then
                        can = Not isLetter(Asc(Mid$(sText, j - 1, 1)))
                    Else
                        can = True 'start of string
                    End If
                    If can Then
                        sText = Left$(sText, j - 1) & sReplacements(i).sOut & vbSpace & Mid$(sText, j + Len(sReplacements(i).sIn) - 1)
                    End If
                    j = InStr(j + 1, sText, sReplacements(i).sIn)
                Loop While j > 0
            Else
                'end of word: [^a-zA-Z]
                j = InStr(1, sText, Left$(sReplacements(i).sIn, Len(sReplacements(i).sIn) - 1), vbTextCompare)
                Do
                    If j > 0 And j < Len(sText) Then
                        If isLetter(Asc(Mid$(sText, j + Len(sReplacements(i).sIn) - 1, 1))) = False Then
                            sText = Left$(sText, j - 1) & vbSpace & sReplacements(i).sOut & vbSpace & Mid$(sText, j + Len(sReplacements(i).sIn))
                        End If
                    End If
                    j = InStr(j + 1, sText, sReplacements(i).sIn)
                Loop While j > 0
            End If
        End If
    End If
Next i


For i = LBound(sTagReplacements) To UBound(sTagReplacements)
    iPos = InStr(1, sText, "<" & sTagReplacements(i).sInStart, vbTextCompare)
    If iPos > 0 Then
        Do
            Dim closePos As Integer
            'have "<p" - check for integer
            
            closePos = InStr(iPos + 1, sText, ">")
            If closePos > -1 Then
                Dim Value As Integer
                On Error GoTo cont2
                Value = CInt(Mid$(sText, iPos + 2, closePos - iPos - 2))
                sText = Left$(sText, iPos) & _
                    sTagReplacements(i).sOutStart & _
                    Quote & _
                    CStr(Value) & _
                    Quote & _
                    IIf(sTagReplacements(i).endInTag, " /", "") & ">" & _
                    Mid$(sText, closePos + 1)
                
                If sTagReplacements(i).endInTag = False Then
                    Dim tag As String
                    tag = "</" & sTagReplacements(i).sInStart & ">"
                    closePos = InStr(closePos + 1, sText, tag)
                    If closePos > 0 Then
                        Dim properclose As String
                        Dim itmp As Integer
                        
                        itmp = InStr(1, sTagReplacements(i).sOutStart, vbSpace)
                        If itmp Then
                            properclose = "</" & Left$(sTagReplacements(i).sOutStart, itmp - 1) & ">"
                        Else
                            properclose = "</" & sTagReplacements(i).sOutStart & ">"
                        End If
                        sText = Replace$(sText, tag, properclose)
                    End If
                End If
            End If
cont2:
            iPos = InStr(iPos + 1, sText, "<" & sTagReplacements(i).sInStart, vbTextCompare)
        Loop While iPos > 0
    End If
Next i

End Sub

Private Function isLetter(ascii As Integer) As Boolean
Const littleAChar As Integer = 97, littleZChar As Integer = 122
Const bigAChar As Integer = 65, bigZChar As Integer = 90
'a=97
'z=122
'A=65
'Z=90

If littleAChar <= ascii And ascii <= littleZChar Then
    isLetter = True
ElseIf bigAChar <= ascii And ascii <= bigZChar Then
    isLetter = True
Else
    isLetter = False
End If

End Function

Public Sub StopSpeech()

cSpeech.pStopSpeech

End Sub

Public Sub SetVoice(ByVal iVoice As Integer)

cSpeech.pSetVoice iVoice

End Sub

Public Sub SpeechInit()

modLoadProgram.SetSplashInfo "Initialising Speech..."
'AddConsoleText "Initialising Speech Object..."

Set cSpeech = New clsSpeech
bVoice = True

'AddConsoleText vbNullString

cSpeech.pGetVoices

ReDim sReplacements(0 To 22)
'sIn MUST be lower case
sReplacements(0).sIn = "ke2"
sReplacements(0).sOut = "hhhahahahhahaahhahaahaahahhahhaahahhaahaha"
sReplacements(0).bWholeWord = True

sReplacements(1).bWholeWord = False
sReplacements(1).sIn = "*nq"
sReplacements(1).sOut = "<silence msec =" & Quote & "250" & Quote & "/><pitch middle=" & Quote & "-6" & _
    Quote & "><rate speed=" & Quote & "4" & Quote & ">Not Quite"


sReplacements(2).bWholeWord = False
sReplacements(2).sIn = "byk*"
sReplacements(2).sOut = "But, You Know.</rate></pitch><silence msec =" & Quote & "100" & Quote & " />"

sReplacements(3).sIn = "m$"
sReplacements(3).sOut = "microsoft"

sReplacements(4).sIn = "$am"
sReplacements(4).sOut = "microsoft sam"

sReplacements(5).sIn = "&c"
sReplacements(5).sOut = "etc"

sReplacements(6).sIn = "sec"
sReplacements(6).sOut = "sek"
sReplacements(6).bWholeWord = True


sReplacements(7).sIn = "bsy"
sReplacements(7).sOut = "be seeing you"
sReplacements(7).bWholeWord = True

sReplacements(8).sIn = "sy"
sReplacements(8).sOut = "seeing you"
sReplacements(8).bWholeWord = True

sReplacements(9).sIn = "atm"
sReplacements(9).sOut = "at the moment"
sReplacements(9).bWholeWord = True

sReplacements(10).sIn = "btw"
sReplacements(10).sOut = "by the way"
sReplacements(10).bWholeWord = True

sReplacements(11).sIn = "1337"
sReplacements(11).sOut = "leet"

sReplacements(12).sIn = "hstgp"
sReplacements(12).sOut = "(He stole that guy's pizza)"
sReplacements(12).bWholeWord = True

sReplacements(13).sIn = "brb"
sReplacements(13).sOut = "be right back"
sReplacements(13).bWholeWord = True

sReplacements(14).sIn = "<hl>"
sReplacements(14).sOut = "<pitch middle=""-50""><rate speed=""-5"">"
sReplacements(15).sIn = "</hl>"
sReplacements(15).sOut = "</pitch></rate>"

sReplacements(16).sIn = "<hi>"
sReplacements(16).sOut = "<pitch middle=""500"">"
sReplacements(17).sIn = "</hi>"
sReplacements(17).sOut = "</pitch>"

sReplacements(18).sIn = "inb4"
sReplacements(18).sOut = "in before,"

sReplacements(19).sIn = "schweet"
sReplacements(19).sOut = "<rate speed=" & Quote & "-10" & Quote & "> <pitch middle=" & Quote & "100" & Quote & _
    "> schweet </rate> </pitch>"


sReplacements(20).sIn = "<e>"
sReplacements(20).sOut = "<emph>"
sReplacements(21).sIn = "</e>"
sReplacements(21).sOut = "</emph>"

sReplacements(22).sIn = "<3"
sReplacements(22).sOut = "love"
sReplacements(22).bWholeWord = True

ReDim sTagReplacements(0 To 3)

'the code addes the closing tags
sTagReplacements(0).sInStart = "p"
sTagReplacements(0).sOutStart = "pitch middle="

sTagReplacements(1).sInStart = "q"
sTagReplacements(1).sOutStart = "silence msec="
sTagReplacements(1).endInTag = True

sTagReplacements(2).sInStart = "s"
sTagReplacements(2).sOutStart = "rate speed="

sTagReplacements(3).sInStart = "v"
sTagReplacements(3).sOutStart = "volume level="

End Sub

Public Sub SpeechTerminate()
Dim i As Integer

'AddConsoleText vbNullString ' "Terminating cSpeech Object"
Set cSpeech = Nothing
'AddConsoleText vbNullString

'For i = 0 To UBound(VoiceAr)
    'Set VoiceAr(i).tVoice = Nothing
'Next i

Erase VoiceAr

End Sub

'##########
'properties ----------------------------------
'##########
Public Property Let Vol(ByVal V As Integer)

If V > 100 Then
    V = 100
ElseIf V < 1 Then
    V = 1
End If

cSpeech.Vol = V

End Property
Public Property Get Vol() As Integer
Vol = cSpeech.Vol
End Property

Public Property Let Speed(ByVal S As Integer)

If S > 10 Then
    S = 10
ElseIf S < -10 Then
    S = -10
End If

cSpeech.Speed = S

End Property
Public Property Get Speed() As Integer
Speed = cSpeech.Speed
End Property

Public Property Get pitch() As Integer
pitch = cSpeech.pitch
End Property
Public Property Let pitch(ByVal ipitch As Integer)
cSpeech.pitch = ipitch
End Property

Public Property Get Description() As String
Description = cSpeech.Description
End Property

Public Property Get Attr(ByVal A As String) As String
'age, name, gender

If cSpeech.SupportsAttr(A) Then
    Attr = cSpeech.Attr(A)
Else
    Attr = vbNullString
End If

End Property

Public Property Get nSpeechStatus() As eSpeechStatus

If Not (cSpeech Is Nothing) Then
    On Error Resume Next
    nSpeechStatus = cSpeech.SpeechStatus
Else
    nSpeechStatus = sIdle
End If

End Property

Public Property Get SpeechStatus() As String

If nSpeechStatus = sIdle Then
    SpeechStatus = "Idle"
Else
    SpeechStatus = "Speaking"
End If

End Property
