Attribute VB_Name = "modMessaging"
Option Explicit

Public Const BugReportStr As String = "BUG REPORT!"

Public DevBlockedMessage As String
Public Const Default_DevBlockedMessage = "Error - Dev Command Blocked"


'Public Const JeffTag = "blockednose", GregTag = "lancs"
Public Const BoldTag = "B", ItalicTag = "I", UnderLineTag = "U"


Public Const MaxFont = 20, MinFont = 6

'my socket
Public MySocket As Integer
Public bReceivedWelcomeMessage As Boolean

'Public LobbyStr As String

Public Const MsgEncryptionFlag = "¥"
Public Const MsgFontSep = "%"
Public Const MsgNameSeparator = ": "
Private Const ClientListSep = "#"

'Public Const MessageSeperator1 As String = "¤"
'Public MessageSeperator2 As String
Public Const MessageSeperator As String * 1 = "¤"
Public Const MessageStart = "¢" ', _
             MessageEnd = "Ø"


Public TmpClientList As String

Public Typers() As String
Public TypingStr As String

Public Drawers() As String
Public DrawingStr As String

Public BlockedIPs() As String
Public bAllBlocked As Boolean

'Drawing
Public cx As Integer, cy As Integer
Public colour As Long
Public NewLine As Boolean
Public RubberWidth As Integer
Private Const DrawingParamLen As Integer = 11
'end drawing

'Public mMStealth As Boolean

'##############
Public Type ptHost
    sIP As String
    sName As String
End Type
Public UsedIPs() As ptHost
Public CurIPIndex As Integer
Public LastIP As String
'##############

Private pLastMessageTime As Date
Private pLastSender As String

Public Property Get LastMessageTime() As Date
LastMessageTime = pLastMessageTime
End Property
Public Property Get LastSender() As String
LastSender = pLastSender
End Property

Public Sub AddUsedIP(sIP As String, Optional sName As String = vbNullString)
Dim u As Integer, i As Integer

If LenB(sIP) = 0 Then Exit Sub

u = UBound(UsedIPs)

For i = 0 To u
    If UsedIPs(i).sIP = sIP Then Exit Sub
Next i


If LenB(UsedIPs(u).sIP) > 0 Then
    u = u + 1
    ReDim Preserve UsedIPs(u)
End If

UsedIPs(u).sIP = sIP
UsedIPs(u).sName = sName

End Sub

Public Sub RemoveUsedIP(Index As Integer)
Dim i As Integer, u As Integer

u = UBound(UsedIPs)

If u = 0 Then
    ReDim UsedIPs(0)
Else
    'Remove the bullet
    For i = Index To u - 1
        UsedIPs(i) = UsedIPs(i + 1)
    Next i
    
    'Resize the array
    ReDim Preserve UsedIPs(u - 1)
End If


End Sub

Public Sub DataArrival(ByVal Data As String, Optional ByVal Index As Integer = (-1))

Dim Command As String, Str As String, sFont As String
Dim Tmp As String, i As Integer, j As Integer
Dim bTmp As Boolean

Dim bDistribute As Boolean

'for coloured messaged
Dim colour As Long
Dim Msg As String
'end

'for name + message
Dim sTmp As String

'for dev stuff
Dim Reply As String
Dim sTo As String, sFrom As String, actualCmd As String
'end for dev stuff

If Status <> Connected Then Exit Sub

If frmDev_Loaded Then
    If frmDev.Visible Then frmDev.AddDev Data, Index, True
End If

On Error Resume Next
'Command = Left$(Data, 1)
If Left$(Data, 1) <> "-" Then
    Command = Left$(Data, 1)
    Str = Mid$(Data, 2)
Else
    Command = Left$(Data, 2)
    Str = Mid$(Data, 3)
End If
On Error GoTo 0

If Command <> CStr(eCommands.Draw) Then
    If InStr(1, Str, "-", vbTextCompare) Then
        Tmp = Replace$(Str, "-", "1")
    Else
        Tmp = Str
    End If
    If IsNumeric(Tmp) Then
        If Len(Tmp) > 6 Then 'len("0000000003600000003276800000000010")
            If frmMain.mnuDevShowAll.Checked Then
                AddConsoleText "Dropped Received Data: '" & Command & Str & "'"
            End If
            Exit Sub
        End If
    End If
End If

Tmp = vbNullString
bDistribute = True

Select Case Command
    Case eCommands.Draw
        
        If frmMain.mnuOptionsMessagingDrawingOff.Checked = False Then
            Call DrawData(Str)
            Tmp = Mid$(Data, 2)
            DoEvents
        End If
        
    Case eCommands.Drawing 'like Typing
        
        If frmMain.mnuOptionsMessagingDrawingOff.Checked = False Then
            Call EvalDrawing(Str)
            Tmp = Str
        End If
        
    Case eCommands.FileTransferCmd
        Call ProcessFileTransferCmd(Str)
        Exit Sub
        
    Case eCommands.Message
        'add the new message to our chat buffer
        'If Index = (-1) Then
            'Tmp = Str
        'Else
            'Tmp = "(" & Index & ") " & Str
        'End If
        
        'IIf(Index = (-1), TxtSent, TxtReceived)
        
        '#########################
        'Arial%35176731#Rob: Hello
        '#########################
        
        i = InStr(1, Str, MsgFontSep, vbTextCompare)
        ''in case font beings with a "@"
        'If i <= 1 Then
            'i = InStr(i + 1, Str, MsgFontSep, vbTextCompare)
        'End If
        
        Msg = Mid$(Str, InStr(1, Str, "#", vbTextCompare) + 1)
        sFont = Left$(Str, i - 1)
        
        
        
        '###### DOESN'T MATTER EITHERWAY #########
        'If Index = (-1) Then
            'Colour = TxtForeGround 'TxtSent
        'Else
            colour = Mid$(Str, i + 1, InStr(1, Str, "#", vbTextCompare) - i - 1)
        'End If
        
        If Left$(Msg, 1) = modMessaging.MsgEncryptionFlag Then
            Msg = CryptString(Mid$(Msg, 2))
        End If
        
        
        'get name + text
        i = InStr(1, Msg, MsgNameSeparator, vbTextCompare)
        If i Then
            sFrom = Left$(Msg, i - 1) 'name
            sTmp = Mid$(Msg, i + 2) 'message
        Else 'is a /me 'command'
            i = InStr(1, Msg, Space$(1))
            j = InStr(i + 1, Msg, Space$(1))
            
            sFrom = Mid$(Msg, i + 1, j - i - 1)
            sTmp = Msg 'Mid$(Msg, i + 1)
            
            If sFrom = Mid$(sTmp, i + 1, Len(sFrom)) Then
                sFrom = vbNullString
            End If
            
        End If
        
        
        bTmp = Server
        If (bTmp And Index <> -1) Or (Not bTmp) Or modSpeech.sSent Then 'received from external source
            
            If modVars.bStealth = False Then
                If modSpeech.sReceived Then
                    If modSpeech.sSayName Then
                        sTmp = Msg
                    'Else
                        'sTmp = sTmp
                    End If
                    
                    'hhhahahahhahaahhahaahaahahhahhaahahhaahaha
                    
                    On Error Resume Next 'in case they say nothing (?)
                    modSpeech.Say sTmp
                    
                End If
            End If
        End If
        
        
        '############################################################################
        'balloon/mini window
        sTmp = Msg
        If Len(sTmp) > 28 Then
            Tmp = Left$(sTmp, 25) & "..."
        Else
            Tmp = sTmp
        End If
        
        'chop off infostart+end
        If Left$(Tmp, Len(InfoStart)) = InfoStart Then
            Tmp = Mid$(Tmp, Len(InfoStart) + 1)
        End If
        If Right$(Tmp, Len(InfoEnd)) = InfoEnd Then
            Tmp = Left$(Tmp, Len(Tmp) - Len(InfoEnd))
        End If
        
        SetMiniInfo Tmp
        
        If Not frmMain.Visible Then
            frmSystray.ShowBalloonTip "Message Received - '" & Tmp & "'", , NIIF_INFO
        End If
        
        
        If LenB(sFrom) Then
            pLastMessageTime = Time$()
            pLastSender = sFrom
        'else
            'it's a /me message
        End If
        '############################################################################
        
        
        
        If frmMain.mnuOptionsMessagingDisplayNewLine.Checked Then
            'name on line 1, text on line 2
            
            If frmMain.mnuOptionsTimeStamp.Checked Then sFrom = "[" & FormatDateTime$(Time$, vbLongTime) & "] " & sFrom
            
            AddText sFrom & MsgNameSeparator & vbNewLine & Space$(4) & sTmp, colour, , True, sFont
            
        Else
            If frmMain.mnuOptionsTimeStamp.Checked Then Msg = "[" & FormatDateTime$(Time$, vbLongTime) & "] " & Msg
            
            AddText Msg, colour, , True, sFont
        End If
        
        
        If frmMain.mnuOptionsFlashMsg.Checked And (Not (Index = -1)) Then
            'if index = -1, then we are sending, so don't flash
            FlashWin
        End If
        
        
        Tmp = Str
        
    Case eCommands.Typing
        
        Call EvalTyping(Str)
        
        Tmp = Str
        
    Case eCommands.SetTyping
        
        If Not Server Then
            Call ReceivedTypingList(Str)
        End If
        'Tmp = Str
        Exit Sub
        
    Case eCommands.ClientList
        
        If Not Server Then
            Call ReceivedClientList(Str)
            
            Tmp = Str
        Else
            Exit Sub
        End If
        
'    Case eCommands.GetName
'
'        If Not Server Then
'            If bDevMode Then
'                If frmMain.mnuDevDataCmdsNoReply.Checked Then
'                    Exit Sub
'                End If
'            End If
'            SendData eCommands.ReplyName & frmMain.LastName
'        End If
'
'    Case eCommands.ReplyName
'
'        If Server Then
'            frmMain.lstConnected.AddItem Str
'            TmpClientList = TmpClientList & "," & Str & ":" & Index
'            Exit Sub 'prevent from distributing
'        End If
        
    Case eCommands.cmdOther
        
        ProcessOtherCmd Str, Index
        bDistribute = False
        
    Case eCommands.SetClientVar
        
        If Server Then
            ReceivedClientVar Str, Index
            Exit Sub 'don't resend
        End If
        
    Case eCommands.Shake
        
        If LenB(Str) Then
            Tmp = Replace(Str, "-", vbNullString, , , vbTextCompare)
            If IsNumeric(Tmp) Then Exit Sub
        End If
        
        If frmMain.mnuFileGameMode.Checked = False Then
            If modVars.bStealth = False Then
                If frmMain.mnuOptionsMessagingShake.Checked Then
                    
                    frmMain.tmrShake.Enabled = True
                    FlashWin
                    
                    If modSpeech.sReceived Then
                        modSpeech.Say "Shake Received from " & Tmp
                    End If
                    
                End If
            End If
        End If
        
        AddText "Shake Recieved" & IIf(LenB(Str), " from " & Str, vbNullString), TxtReceived, True
        
        Tmp = Str
        
    Case eCommands.Info
        
        ProcessInfoMessage Str
        
        Tmp = Str
        
    Case eCommands.SetSocket
        
        If Not Server Then
            
            ReceivedSocket Str
            
        End If
        Exit Sub
        
    Case eCommands.matrixMessage
        
        'Dim Sel As Long, Txt As String
        
        'Sel = Val(Left$(Str, InStr(1, Str, "#", vbTextCompare) - 1))
        'Txt = Right$(Str, Len(Str) - InStr(1, Str, "#", vbTextCompare))
        
        If Not frmMain.mnuOptionsMessagingIgnoreMatrix.Checked Then
            If Right$(frmMain.rtfIn.Text, 5) = "-----" Then
                AddText vbNullString 'Newline
            End If
            
            If Index = (-1) Then
                'If frmMain.mnuOptionsMessagingColours.Checked Then
                    colour = Left$(Str, InStr(1, Str, "#", vbTextCompare) - 1)
                'Else
                    'Colour = TxtSent
                'End If
            Else
                'If frmMain.mnuOptionsMessagingColours.Checked Then
                    colour = Left$(Str, InStr(1, Str, "#", vbTextCompare) - 1)
                'Else
                    'Colour = TxtReceived
                'End If
            End If
                
            Msg = Mid$(Str, InStr(1, Str, "#", vbTextCompare) + 1)
            
            If frmMain.rtfIn.SelFontName <> DefaultFontName Then
                frmMain.rtfIn.SelFontName = DefaultFontName
            End If
            
            MidText Msg, colour
        End If
        
        Tmp = Str
        
    Case eCommands.Prvate
        
        If ReceivedPrivateMessage(Str) Then
            bDistribute = False
        Else
            Tmp = Str
        End If
        
        
    Case eCommands.HostCmd
        ProcessHostCmd Str
        
        
'    Case eCommands.Invite
'
'
'        If Len(Str) <= Len("xxx.xxx.xxx.xxx") Then
'            'Call frmMain.CleanUp
'            'Cmds Connecting
'            frmMain.Connect Str
'        End If
'
'        Exit Sub
        
'    Case eCommands.mPing
'
'        i = -1
'        On Error Resume Next
'        Tmp = Left$(Str, 1)
'        On Error GoTo 0
'
'        If LenB(Tmp) > 0 Then
'            If IsNumeric(Tmp) Then
'                i = CInt(Tmp)
'            End If
'        End If
'
'        Select Case i
'            Case ePingCmds.aPing
'                'send pong to index
'                SendData eCommands.mPing & ePingCmds.aPong, Index
'
'            Case ePingCmds.aPong
'                'set client's ping
''                For j = 0 To UBound(Clients)
''                    If Clients(j).iSocket = Index Then
''
''                        clients(j).iPing =
''
''                        Exit For
''                    End If
''                Next i
'
'
'        End Select
        
        
    Case eCommands.PingCmd
        
        ProcessPingCmd Str, Index
        
        bDistribute = False
        
    Case eCommands.LobbyCmd
        
        Call ProcessLobbyCmd(Str) ', Index)
        Tmp = Str
'        Exit Sub

    Case eCommands.DevRecieve
        'display recieved command
'        old way of doing it
'        AddText String$(5, "-") & vbNewLine & "Dev Reply:" & vbNewLine & Str & vbNewLine & String$(5, "-"), TxtReceived, False
'
'        Exit Sub

        Dim tName As String
        
        
        'Name & "#" & DevCmd
        On Error Resume Next
        actualCmd = Right$(Str, Len(Str) - InStr(1, Str, "#", vbTextCompare))
        tName = Left$(Str, InStr(1, Str, "#", vbTextCompare) - 1)
        On Error GoTo 0
        
        If Trim$(LCase$(tName)) = LCase$(Trim$(frmMain.LastName)) Then
            
            AddDevText "Dev Reply: " & actualCmd, True
            Exit Sub
        Else
            
            Tmp = Str
            
        End If
        
        
    Case eCommands.DevSend
        
        If ProcessDevCommand(Str, Index) Then
            bDistribute = False
        Else
            Tmp = Str
        End If
        
        
    Case Else
        
        If Command = "!" And LenB(Str) Then
            AddText "Basic Message Received: " & Str, , True
        Else
            
            If frmMain.mnuDevShowAll.Checked Then '(Not IsNumeric(Data)) Or
                For i = LBound(Clients) To UBound(Clients)
                    If i = Index Then
                        sFrom = Clients(i).sName
                        Exit For
                    End If
                Next i
                
                sTmp = "Unknown Data Recieved: '" & Data & "'" & _
                    IIf(LenB(sFrom), " from " & sFrom, vbNullString)
                
                AddText sTmp, TxtUnknown, True
            End If
            AddConsoleText sTmp
            
            bDistribute = False
        End If
        
End Select



If modVars.bStealth Then
    If Command = eCommands.matrixMessage Then
        
        frmStealth.AddText Mid$(Str, InStr(1, Str, "#", vbTextCompare) + 1), _
            (Command <> eCommands.matrixMessage), (Index <> -1)
        
    End If
End If

If Server Then
    If bDistribute Then
        Call DistributeMsg(Command & Tmp, Index)
    End If
End If

End Sub

Private Sub ProcessOtherCmd(ByVal Str As String, ByVal Index As Integer)
Dim iCmd As eOtherCmds
Dim sCmd As String, sParam As String

On Error GoTo EH
sCmd = Left$(Str, 1)
sParam = Mid$(Str, 2)


iCmd = val(sCmd)

Select Case iCmd
    Case eOtherCmds.SetServerName
        
        If CurIPIndex > -1 Then
            '0000004080000000019200000000000000000000001
            If IsNumeric(Replace$(sParam, "-", vbNullString)) = False Then
                UsedIPs(CurIPIndex).sName = sParam
            End If
        End If
        
        
        
    Case eOtherCmds.ConnectToServerVoicePort
        
        If Not Server Then
            'connect + receive our recording
            frmMain.ucVoiceTransfer.Connect frmMain.SckLC.RemoteHostIP, modPorts.VoicePort
            
            frmMain.SetInfo "Receiving Recording off Server... (Connecting...)", False
            'should receive it once connected
        End If
        
End Select


EH:
End Sub

Public Sub SendInfoMessage(ByVal sTxt As String, _
    Optional bSpeak As Boolean = True, Optional bError As Boolean = False, Optional bBannerInfo As Boolean = False, _
    Optional iSocket As Integer = -1, Optional ntiSocket As Integer = -1)

If Server Then
    If iSocket = -1 Then
        DistributeMsg eCommands.Info & sTxt & Abs(bError) & Abs(bBannerInfo) & Abs(bSpeak), ntiSocket
    Else
        SendData eCommands.Info & sTxt & Abs(bError) & Abs(bBannerInfo) & Abs(bSpeak), iSocket
    End If
Else
    SendData eCommands.Info & sTxt & Abs(bError) & Abs(bBannerInfo) & Abs(bSpeak)
End If


End Sub

Public Sub SendHostCmd(iCmd As eHostCmds, sParam As String, Optional ByVal iSockTo As Integer = -1)

If iSockTo = -1 Then
    DistributeMsg eCommands.HostCmd & iCmd & sParam, -1
Else
    SendData eCommands.HostCmd & iCmd & sParam, iSockTo
End If

End Sub

Private Sub ProcessHostCmd(ByVal sTxt As String)
Dim iCmd As eHostCmds, i As Integer
Dim sParam As String

On Error GoTo EH
iCmd = Left$(sTxt, 1)
sParam = Mid$(sTxt, 2)

If iCmd = eHostCmds.RemoveDP Then
    'socket to remove
    i = CInt(sParam)
    frmMain.ResetImgDP i
    
    
    'client index
    i = FindClient(i)
    If i > -1 Then
        Set Clients(i).IPicture = Nothing
        
        'path of DP
        sParam = modDP.GetClientDPStr(i)
        
        If FileExists(sParam) Then
            On Error GoTo EH
            Kill sParam
        End If
        
    End If
    
End If

EH:
End Sub

Private Sub ProcessInfoMessage(ByVal Str As String)
Dim bError As Boolean, bBanner As Boolean, bSpeak As Boolean
Dim sTxt As String
'format: "txt & Abs(bError) & Abs(bBannerInfo) & Abs(Speak)

On Error Resume Next
bBanner = CBool(Mid$(Str, Len(Str), 1))
bError = CBool(Mid$(Str, Len(Str) - 1, 1))
bSpeak = CBool(Mid$(Str, Len(Str) - 2, 1))
sTxt = Left$(Str, Len(Str) - 3)

If Left$(sTxt, 7) = "Welcome" Then
    bReceivedWelcomeMessage = True
    frmMain.SetInfo "Welcome Message Received, Type Away...", False
End If

If bBanner Then
    frmMain.SetInfo sTxt, bError
    SetMiniInfo sTxt
Else
    AddText sTxt, IIf(bError, TxtError, TxtInfo), True
End If

If sTxt = modMessaging.BugReportStr Then
    modAudio.PlayBugReport
End If

If bSpeak Then
    If modSpeech.sSayInfo Then
        If modVars.bStealth = False Then
            modSpeech.Say sTxt
        End If
    End If
End If

'If Right$(Str, 1) = "1" Then
    'AddText Left$(Str, Len(Str) - 1), TxtError, True
'Else 'if right$(str,1)="0" then
    'AddText Left$(Str, Len(Str) - 1), , True
'End If

End Sub

Private Function ReceivedPrivateMessage(ByVal Str As String) As Boolean
'######## RETURNS TRUE IF MESSAGE WAS FOR US ####################

Dim iSockTo As Integer, iSockFrom As Integer
Dim i As Integer, j As Integer, K As Integer
Dim sFrom As String, sMsg As String
Dim Frm As Form, bFormFound As Boolean

'OLD WhoTo & # & From & @ & Message - ecmds.message format
'NEW WhoToSock & : & WhoFromSock & # & Message        'From & @ & Message


On Error GoTo EH
i = InStr(1, Str, "#", vbTextCompare)
j = InStr(1, Str, "@", vbTextCompare)
'k = InStr(1, Str, "@", vbTextCompare)


iSockTo = Left$(Str, i - 1)
iSockFrom = Mid$(Str, i + 1, j - i - 1)
'sFrom = Mid$(Str, j + 1, k - j - 1)
sMsg = Right$(Str, Len(Str) - j)


If modMessaging.MySocket Then
    If iSockTo = modMessaging.MySocket Then
        
        ReceivedPrivateMessage = True
        'don't distribute
        
        If frmMain.mnuFileGameMode.Checked Then
            
            AddText Trim$(sFrom) & " is trying to private chat with you", TxtError, True
            AddText "Disable Game Mode to converse", TxtError, True
            
        Else
            'find the window
            
            For Each Frm In Forms
                'len("Private Comm Channel - ")
                If Frm.Name = frmPrivateName Then
                    If Frm.SendToSock = iSockFrom Then 'Trim$(Mid$(Frm.Caption, 23)) = sFrom Then
                        bFormFound = True
                        Exit For
                    End If
                End If
                
            Next Frm
            
            If Not bFormFound Then
                
                On Error Resume Next
                'form isn't found, and it's only info
                If Left$(sMsg, Len(InfoStart)) = InfoStart Then Exit Function
                
                'otherwise, not found...
                Set Frm = New frmPrivate
                Load Frm
                
                'Frm.Show vbModeless, frmMain
                Frm.SendToSock = iSockFrom
            End If
            
            
            Frm.AddPvtText sMsg, TxtReceived, , iSockFrom
            
        End If
        
        
    Else
        ReceivedPrivateMessage = False
        'not for us, it will be distributed, if server
    End If
End If

EH:
End Function

Private Function MakeSocketEnding() As String
MakeSocketEnding = CStr(MakeSquareNumber())
End Function

Private Sub ReceivedSocket(ByVal Str As String)
Dim iSock As Integer, iSquare As Single, i As Integer

On Error GoTo EH

i = InStr(1, Str, vbSpace)
iSock = Left$(Str, i - 1)
iSquare = Mid$(Str, i + 1)

If IsSquare(iSquare) Then
    modMessaging.MySocket = iSock
End If

EH:
End Sub

Public Sub SendSetSocketMessage(iSocket As Integer)

SendData CStr(eCommands.SetSocket & iSocket) & _
    vbSpace & MakeSocketEnding(), iSocket

End Sub

Public Sub EvalTyping(ByVal Str As String)

Dim sFrom As String, Tmp As String
Dim i As Integer, j As Integer
Dim bTmp As Boolean

sFrom = Mid$(Str, 2)

Tmp = Replace(sFrom, "-", vbNullString, , , vbTextCompare)
If IsNumeric(Tmp) Then Exit Sub

Tmp = Left$(Str, 1)


If Tmp = "0" Then
    For i = 1 To UBound(Typers)
        If Typers(i) = sFrom Then
            'remove from array
            Typers(i) = vbNullString
            
            For j = i To UBound(Typers) - 1
                Typers(j) = Typers(j + 1)
            Next j
            
            Typers(j) = vbNullString
            
            ReDim Preserve Typers(UBound(Typers) - 1)
            Exit For
        End If
    Next i
    
Else
    For i = 1 To UBound(Typers)
        If Typers(i) = sFrom Then
            bTmp = True 'found = t
            Exit For
        End If
    Next i
    
    If Not bTmp Then 'not found
        j = UBound(Typers)
        
        ReDim Preserve Typers(j + 1)
        
        Typers(j + 1) = sFrom
    End If
End If


UpdateTypingStr
SetTypeCap

End Sub

Private Sub EvalDrawing(ByVal Str As String)

Dim Tmp As String, sFrom As String
Dim i As Integer, j As Integer
Dim bTmp As Boolean

Tmp = Left$(Str, 1)

sFrom = Mid$(Str, 2)

If Tmp = "0" Then
    'DrawingStr = Replace$(DrawingStr, Mid$(Str, 2), vbNullString, , , vbTextCompare)
    
    For i = 1 To UBound(Drawers)
        If Drawers(i) = sFrom Then
            'remove from array
            Drawers(i) = vbNullString
            
            For j = i To UBound(Drawers) - 1
                Drawers(j) = Drawers(j + 1)
            Next j
            
            Drawers(j) = vbNullString
            
            ReDim Preserve Drawers(UBound(Drawers) - 1)
            Exit For
        End If
    Next i
    
Else
'            If Not CBool(InStr(1, DrawingStr, Mid$(Str, 2), vbTextCompare)) Then
'                DrawingStr = "Typing" & MsgNameSeparator & Mid$(Str, 2) & Space$(2) & Mid$(DrawingStr, 8)
'            End If
    
    For i = 1 To UBound(Drawers)
        If Drawers(i) = sFrom Then
            bTmp = True 'found = t
            Exit For
        End If
    Next i
    
    If Not bTmp Then 'not found
        j = UBound(Drawers)
        
        ReDim Preserve Drawers(j + 1)
        
        Drawers(j + 1) = sFrom
    End If
    
End If

UpdateDrawingStr
SetTypeCap

End Sub

Public Function GetTypingList() As String
Dim i As Integer
Dim sTmp As String
Dim sTypers As String, sDrawers As String

For i = 0 To UBound(Typers)
    If LenB(Typers(i)) Then
        sTmp = sTmp & "#" & Typers(i)
    End If
Next i

On Error Resume Next
sTypers = Mid$(sTmp, 2)
sTmp = vbNullString

For i = 0 To UBound(Drawers)
    If LenB(Drawers(i)) Then
        sTmp = sTmp & "#" & Drawers(i)
    End If
Next i

On Error Resume Next
sDrawers = Mid$(sTmp, 2)


GetTypingList = sTypers & "@" & sDrawers

'GetTypingList = Typer1#Typer2@Drawer1#Drawer2

End Function

Private Sub ReceivedTypingList(ByVal Str As String)
Dim i As Integer

Dim pTypers() As String, sTypers As String
Dim pDrawers() As String, sDrawers As String

'Typer1#Typer2#Typer3@Drawer1#Drawer2#Drawer3

i = InStr(1, Str, "@")

If i Then
    On Error Resume Next
    sTypers = Left$(Str, i - 1)
    sDrawers = Mid$(Str, i + 1)
    
    pTypers = Split(sTypers, "#")
    pDrawers = Split(sDrawers, "#")
    
    ReDim Typers(0 To UBound(pTypers) + 1)
    ReDim Drawers(0 To UBound(pDrawers) + 1)
    
    For i = 0 To UBound(pTypers)
        If LenB(pTypers(i)) Then
            Typers(i + 1) = pTypers(i)
        End If
    Next i
    
    
    For i = 0 To UBound(pDrawers)
        If LenB(pDrawers(i)) Then
            Drawers(i + 1) = pDrawers(i)
        End If
    Next i
    
Else
    ReDim Typers(0)
    ReDim Drawers(0)
End If

UpdateTypingStr
UpdateDrawingStr
SetTypeCap

End Sub

Public Sub UpdateTypingStr()
Dim i As Integer

TypingStr = vbNullString

For i = 1 To UBound(Typers)
    If Typers(i) <> frmMain.LastName Then
        If LenB(Typers(i)) Then
            TypingStr = TypingStr & ", " & Typers(i)
        End If
    End If
Next i

On Error Resume Next
TypingStr = "Typing: " & Trim$(Mid$(TypingStr, 2)) 'get rid of 1st msgnamesep
On Error GoTo 0

If Len(TypingStr) <= 8 Then TypingStr = vbNullString

End Sub

Public Sub UpdateDrawingStr()
Dim i As Integer

DrawingStr = vbNullString

For i = 1 To UBound(Drawers)
    If Drawers(i) <> frmMain.LastName Then
        If LenB(Drawers(i)) Then
            DrawingStr = DrawingStr & ", " & Drawers(i)
        End If
    End If
Next i

On Error Resume Next
DrawingStr = "Drawing: " & Trim$(Mid$(DrawingStr, 2)) 'get rid of 1st msgnamesep
On Error GoTo 0

If Len(DrawingStr) <= 9 Then DrawingStr = vbNullString

End Sub

Public Sub SetTypeCap()
If LenB(TypingStr) Then
    frmMain.lblTyping.Caption = TypingStr & vbNewLine & DrawingStr
Else
    frmMain.lblTyping.Caption = DrawingStr
End If
End Sub

'################################################################

Private Sub ProcessPingCmd(sTxt As String, IndexFrom As Integer)
Dim GTC As Long
Dim iClient As Integer, i As Integer

Select Case Left$(sTxt, 1)
    Case ePingCmds.aPing
        If Not Server Then
            SendData eCommands.PingCmd & ePingCmds.aPong
        End If
        
    Case ePingCmds.aPong
        
        If Server Then
            GTC = GetTickCount()
            iClient = FindClient(IndexFrom) '-1
            
'            For i = 0 To UBound(Clients)
'                If Clients(i).iSocket = IndexFrom Then
'                    iClient = i
'                    Exit For
'                End If
'            Next i
            
            If iClient > -1 Then
                With Clients(iClient)
                    .iPing = GTC - .lPingStart
                    
                    If .iPing = 0 Then .iPing = 1
                    
                    .lLastPing = GTC
                End With
            End If
            
            
            
        End If
        
End Select

End Sub

'################################################################

Private Sub ReceivedClientVar(ByVal Str As String, ByVal Index As Integer)
Dim cmd As eClientVarCmds
Dim Clienti As Integer, i As Integer


On Error GoTo EH

cmd = CInt(Left$(Str, 1))
Clienti = FindClient(Index) '-1

'For i = 0 To UBound(Clients)
'    If Clients(i).iSocket = Index Then
'        Clienti = i
'        Exit For
'    End If
'Next i

If Clienti = -1 Then
    ReDim Preserve Clients(UBound(Clients) + 1)
    Clienti = UBound(Clients)
End If

If Clients(Clienti).iSocket <> Index Then
    Clients(Clienti).iSocket = Index
End If


If Clienti <> -1 Then
    If Clients(Clienti).iSocket <> -1 Then
        
        Select Case cmd
            Case eClientVarCmds.SetName
                
                'Dim bJustConnected As Boolean
                'bJustConnected = (LenB(Clients(Clienti).sName) = 0)
                
                Clients(Clienti).sName = Mid$(Str, 2)
                
                If (Not Clients(Clienti).bShownConnection) And Server Then
                    'they've just connected, show a popup
                    
                    frmSystray.ShowBalloonTip Clients(Clienti).sName & " Connected" & vbNewLine & vbNewLine & _
                                              Clients(Clienti).sIP & vbNewLine & _
                                              "(Client " & Clienti & ")", _
                                              "Communicator - New Connection", NIIF_INFO
                    
                    
                    Clients(Clienti).bShownConnection = True
                End If
                
                
            Case eClientVarCmds.SetDrawing
                
                Clients(Clienti).BlockDrawing = CBool(Mid$(Str, 2))
                
                
            Case eClientVarCmds.SetVersion
                
                Str = Mid$(Str, 2)
                
                If Len(Str) <= 10 Then 'Len("xx.xxx.xxx") Then
                    Clients(Clienti).sVersion = Str
                End If
                
            Case eClientVarCmds.SetDPSet
                
                Clients(Clienti).bDPSet = CBool(Mid$(Str, 2))
                
                
            Case eClientVarCmds.SetsStatus
                
                Clients(Clienti).sStatus = Mid$(Str, 2)
                
                
    '        Case eClientVarCmds.SetPing
    '
    '            Clients(Clienti).iPing = GetTickCount() - Clients(Clienti).iPingSent
            
    '        Case eClientVarCmds.SetSocket
    
    '            Clients(Clienti).iSocket = CInt(Mid$(Str, 2))
    
        End Select
    End If
End If

EH:
End Sub

Public Sub SendName()

If Not Server Then
    SendData eCommands.SetClientVar & eClientVarCmds.SetName & frmMain.LastName
End If

End Sub

Private Sub ProcessFileTransferCmd(Str As String)
Dim cmd As eFTCmds, Rest As String
Dim i As Integer, MyClienti As Integer

On Error Resume Next
cmd = Left$(Str, 1)
Rest = Mid$(Str, 2)

Select Case cmd
    Case eFTCmds.FT_SendDPToHost
        
        If FileExists(modDP.My_DP_Path) Then
            
            modDP.bSentMyPicture = False
            
            'For i = 0 To UBound(Clients)
                
                'If Clients(i).sName = frmMain.LastName Then
                    
                    'frmMain.sFileToSend = modDP.My_DP_Path
                    'frmMain.sRemoteFileName = CStr(Clients(i).iSocket) & ".jpg"
                    'frmMain.ucFileTransfer.Connect frmMain.SckLC.RemoteHostIP
                    
                    'Exit For
                'End If
            'Next i
            
        End If
        
        
    Case eFTCmds.FT_ConnectToHost
        
        frmMain.ucFileTransfer.Connect frmMain.SckLC.RemoteHostIP, modPorts.DPPort
        
        
    Case eFTCmds.FT_Close
        
        frmMain.ucFileTransfer.Disconnect
        
End Select

End Sub

'Private Sub ProcessLobbyCmd(ByVal Data As String, ByVal Idx As Integer)
'
'Dim Command As eLobbyCmds
'Dim Str As String
'
'On Error Resume Next
'Command = Left$(Data, 1)
'Str = Mid$(Data, 2)
'On Error GoTo 0
'
'Select Case Command
'    Case eLobbyCmds.Request
'        If Server Then
'            SendData eCommands.LobbyCmd & eLobbyCmds.Reply & LobbyStr, Idx
'        End If
'
'    Case eLobbyCmds.Reply
'
'        Call ParseLobbyReply(Str)
'
'        'frmLobby.lSetStatus "Retrived Game List"
'
''    Case eLobbyCmds.Add
''
''        If Server Then
''            LobbyStr = LobbyStr & Data
''        End If
''
''    Case eLobbyCmds.Remove
''        If Server Then
''            LobbyStr = Replace$(LobbyStr, Data, vbNullString, , , vbTextCompare)
''        End If
'
'End Select
'
'End Sub

'Public Sub ParseLobbyReply(ByVal Data As String)
'Dim Ip As String, Gtype As String
'Dim Dats() As String
'Dim i As Integer
'
'Dats() = Split(Data, "#", , vbTextCompare)
'
'With frmLobby
'    .lstIP.Clear
'    .lstGameType.Clear
'End With
'
'For i = LBound(Dats) To UBound(Dats)
'    If Len(Dats(i)) <> 0 Then
'        On Error Resume Next
'        Ip = Left$(Dats(i), InStr(1, Dats(i), ",", vbTextCompare) - 1)
'        Gtype = Mid$(Dats(i), InStr(1, Dats(i), ",", vbTextCompare) + 1)
'
'        frmLobby.lstIP.AddItem Ip
'        frmLobby.lstGameType.AddItem Gtype
'
'    End If
'Next i
'
'frmLobby.lSetStatus "Retrieved Games"
'
'End Sub

'Public Function GetClientList() As String
'
'On Error Resume Next
'GetClientList = TmpClientList & "," & frmMain.LastName & ":-1"
'
'TmpClientList = vbNullString
'
'End Function

Public Sub SendData(ByVal Str As String, Optional ByVal Client As Integer = (-1))

If modLoadProgram.frmDev_Loaded Then
    frmDev.AddDev Str, Client, False
End If

If Client = (-1) Then
    On Error GoTo SendEH
    frmMain.SckLC.SendData MessageStart & Str & MessageSeperator
Else
    On Error GoTo SendEH
    frmMain.SockAr(Client).SendData MessageStart & Str & MessageSeperator
End If

Exit Sub
SendEH:
'If Err.Number <> 0 Then
'    'had an error, disconnect
'    If frmMain.mnuDevAdvCmdsIgnoreSD.Checked = False And Status = Connected Then
'        'raise it
'        If Client = -1 Then 'i.e. server
'            frmMain.SckLC_Error CustomLagError, Err.Description, 0, "SendData()", vbNullString, 0, False
'        Else
'            frmMain.sockAr_Error Client, CustomLagError, Err.Description, 0, "SendData()", vbNullString, 0, False
'        End If
'
'        Beep
'    End If
'End If
End Sub

Public Sub DistributeMsg(ByVal Msg As String, ByVal Nt As Integer)
Dim n As Integer, Clienti As Integer, i As Integer
Dim cmd As eCommands

'now the client says something, wich arrived at the server...
'the server must now redistibute this message to all other connected
'clients...
On Error Resume Next
cmd = CInt(Left$(Msg, 1))

With frmMain
    On Error Resume Next    'Error Handler
    
    For n = 1 To SocketCounter
        
        If n <> Nt Then  'we don't want to send the msg back to the sender
            
            If .SockAr(n).state = sckConnected Then   'if socket is connected
                
                If cmd = Draw Then
'                    For i = 0 To UBound(Clients)
'                        If Clients(i).iSocket = N Then
'                            If Clients(i).BlockDrawing Then GoTo Nex
'                        End If
'                    Next i
                    
                    i = FindClient(n)
                    
                    If i > -1 Then
                        If Clients(i).BlockDrawing Then GoTo Nex
                    End If
                    
                End If
                
                Call SendData(Msg, n)
                DoEvents
                
            End If
            
        End If
Nex:
        
    Next n
    
End With


End Sub

'Public Sub CreateRandomClients()
'Dim i As Integer
'ReDim Clients(0 To 3)
'
'With Clients(0)
'    .sName = frmMain.LastName
'    .iSocket = -1
'    .sVersion = GetVersion()
'    .sIP = "HostIP"
'End With
'
'Clients(1).sName = "Tim"
'Clients(2).sName = "Greg"
'Clients(3).sName = "Jeff"
'
'For i = 1 To 3
'    Clients(i).sIP = Clients(i).sName & "IP"
'    Clients(i).sVersion = Clients(i).sName & "Version"
'    Clients(i).iSocket = i
'Next i
'
'
'End Sub

Public Sub ReceivedClientList(ByVal lst As String)
Dim Clts() As String, i As Integer
'Dim j As Integer, K As Integer, n As Integer, m As Integer
Dim Parts() As String

'e.g.
'"Test2@-1@1.38.9999@HostIP@,Tim@1@TimVersion@TimIP@,Greg@2@GregVersion@GregIP@,Jeff@3@JeffVersion@JeffIP@"
'                            |                       |                          |

frmMain.EnableCmd 3, False

Clts = Split(lst, ClientListSep, , vbTextCompare)

ReDim Preserve Clients(UBound(Clts))


For i = 0 To UBound(Clts)
    
    On Error Resume Next
    Erase Parts
    Parts = Split(Clts(i), "@")
    
    Clients(i).sName = Parts(0)
    Clients(i).iSocket = Parts(1)
    Clients(i).sVersion = Parts(2)
    Clients(i).sIP = Parts(3)
    Clients(i).iPing = CInt(Parts(4))
    Clients(i).sStatus = Parts(5)
    
'    j = InStr(1, Clts(i), ":", vbTextCompare)
'    K = InStr(1, Clts(i), "#", vbTextCompare)
'    n = InStr(1, Clts(i), "@", vbTextCompare)
'    m = InStr(1, Clts(i), "", vbTextCompare)
'    Clients(i).sName = Left$(Clts(i), j - 1)
'    Clients(i).iSocket = Mid$(Clts(i), j + 1, K - j - 1)
'    Clients(i).sVersion = Mid$(Clts(i), K + 1, n - K - 1)
'    Clients(i).sIP = Mid$(Clts(i), n + 1)
Next i

Erase Parts
Erase Clts

End Sub

Public Function GetClientList() As String
Dim i As Integer, j As Integer
Dim S As String
Dim iSock As Integer
Dim bCan As Boolean
Const Sep = "@"

Do While i <= UBound(Clients)
    bCan = False
    
    iSock = Clients(i).iSocket
    If iSock <> -1 Then
        If frmMain.ControlExists(frmMain.SockAr(iSock)) Then
            bCan = (frmMain.SockAr(iSock).state = sckConnected)
        Else
            bCan = False
        End If
    Else
        bCan = True
        'it's me, server
    End If
    
    If bCan Then
        S = S & ClientListSep & Clients(i).sName & Sep & _
                      Clients(i).iSocket & Sep & _
                      Clients(i).sVersion & Sep & _
                      Clients(i).sIP & Sep & _
                      CStr(Clients(i).iPing) & Sep & _
                      Clients(i).sStatus & Sep
        
    Else
        'don't resize the array - messes up the pictures..?
        For j = i To UBound(Clients) - 1
            Clients(j) = Clients(j + 1)
        Next j

        ReDim Preserve Clients(UBound(Clients) - 1)

        i = i - 1

    End If
    
    i = i + 1
Loop

GetClientList = Mid$(S, 2) 'chop off beginning ','

End Function

Public Sub DrawData(ByVal Data As String)

Dim OldDrawWidth As Long

If LCase$(Data) = "cls" Then
    'frmMain.SaveLastPic
    frmMain.ClearBoard
    'frmMain.pDrawDrawnOn = False
    Exit Sub
End If

On Error Resume Next

With frmMain
    If .pDrawDrawnOn = False Then .pDrawDrawnOn = True
    
    
    OldDrawWidth = .picDraw.DrawWidth
    
    .picDraw.DrawWidth = sParam(Data, 6)
    .picDraw.Line (sParam(Data, 1), sParam(Data, 2))- _
        (sParam(Data, 3), sParam(Data, 4)), sParam(Data, 5)
    
    .picDraw.DrawWidth = OldDrawWidth
End With

End Sub

Public Sub SendLine(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal DrawWidth As Long, _
Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer)

Dim Data As String

If X2 = 0 And Y2 = 0 Then
    Data = eCommands.Draw & sFormatSend(X1) & sFormatSend(Y1) & sFormatSend(X1) & sFormatSend(Y1)
Else
    Data = eCommands.Draw & sFormatSend(X1) & sFormatSend(Y1) & sFormatSend(X2) & sFormatSend(Y2)
End If

'Data = Data & IIf(Grid, sFormatSend(GridColour), sFormatSend(Colour))

Data = Data & sFormatSend(colour) & sFormatSend(DrawWidth)

If Server Then
    DistributeMsg Data, -1
Else
    SendData Data
End If

End Sub

Private Function sFormatSend(ByVal vData As String) As String
'Format data to send.
'Make it exactly PARAM_LEN chars long.
sFormatSend = Format$(vData, String$(DrawingParamLen, "0"))

'If it is (PARAM_LEN + 1) chars long, that means there is a negative sign at the front.
'So format it one character shorter.
If Len(sFormatSend) = DrawingParamLen + 1 Then
    sFormatSend = Format$(vData, String$(DrawingParamLen - 1, "0"))
End If

End Function

Public Function sParam(Data As String, Num As Integer) As String
'This function pulls the (viNum)th parameter from datastream vsData, which is being processed in the ProcessData procedure.
'This parameter is exactly PARAM_LEN characters long.
sParam = Mid$(Data, DrawingParamLen * (Num - 1) + 1, DrawingParamLen)
End Function

'################################################################################################
'translation

'Public Function TranslateToCCB(ByVal sTxt As String) As String
'Dim i As Integer, j As Integer
'Dim sToBeTranslated As String, sNew As String
'
'
'i = InStr(1, sTxt, "<" & JeffTag & ">", vbTextCompare)
'If i Then
'    J = InStr(1, sTxt, "</" & JeffTag & ">", vbTextCompare)
'    If J Then
'
'        sToBeTranslated = Mid$(sTxt, i + Len(JeffTag) + 2, J - i - (Len(JeffTag) + 2))
'
'        sNew = cJeffery(sToBeTranslated)
'
'        Mid$(sTxt, i + Len(JeffTag) + 2, J - i - (Len(JeffTag) + 2)) = sNew
'
'    End If
'End If
'
'
'i = InStr(1, sTxt, "<" & GregTag & ">", vbTextCompare)
'If i Then
'    J = InStr(1, sTxt, "</" & GregTag & ">", vbTextCompare)
'    If J Then
'
'        sToBeTranslated = Mid$(sTxt, i + Len(GregTag) + 2, J - i - (Len(GregTag) + 2))
'
'        sNew = cGregory(sToBeTranslated)
'
'        'Mid$(sTxt, i + Len(GregTag) + 2, j - i - (Len(GregTag) + 2)) = sNew
'
'        sTxt = Left$(sTxt, i + Len(GregTag) + 1) & sNew & Mid$(sTxt, J)
'
'
'    End If
'End If
'
'
'TranslateToCCB = sTxt
'End Function

'Private Function cJeffery(ByVal sTxt As String) As String
'
'cJeffery = Replace$( _
'    Replace$(sTxt, "m", "b", , , vbTextCompare) _
'    , "n", "d", , , vbTextCompare)
'
'End Function
'Private Function cGregory(ByVal sTxt As String) As String
'
'cGregory = _
'    Replace$( _
'    Replace$( _
'    Replace$( _
'    Replace$( _
'    Replace$( _
'    sTxt _
'    , "the", "t'", , , vbTextCompare) _
'    , "good", "grand", , , vbTextCompare) _
'    , "very", "reet", , , vbTextCompare) _
'    , "cake", "reet nice cake", , , vbTextCompare) _
'    , ".", ", eeeeee.", , , vbTextCompare)
'
'End Function
'################################################################################################
