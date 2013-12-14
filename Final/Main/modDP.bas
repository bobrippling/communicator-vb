Attribute VB_Name = "modDP"
Option Explicit

Public My_DP_Path As String
Public bSentMyPicture As Boolean


'for distributing
Public sPictureToSendPath As String, sRemoteFileName As String
Public DP_iClient As Integer, iToAdd As Integer
Public bSetHostBool As Boolean

Public Function GetClientDPStr(iClient As Integer) As String
Dim sPath As String

On Error GoTo EH
If Clients(iClient).iSocket = modMessaging.MySocket Then
    sPath = modDP.DP_Dir_Path & "\Local."
Else
    sPath = modDP.DP_Dir_Path & "\" & CStr(Clients(iClient).iSocket) & "."
End If

If Clients(iClient).bDPIsGIF Then
    sPath = sPath & "gif"
Else
    sPath = sPath & "jpg"
End If

GetClientDPStr = sPath

EH:
End Function

Public Function DP_Path_Exists() As Boolean
DP_Path_Exists = FileExists(My_DP_Path)
End Function

Public Property Get DP_Dir_Path() As String
DP_Dir_Path = frmMain.DP_Path 'AppPath() & "Files"
End Property

Public Sub DelPics()
'ClearFolder DP_Dir_Path & "\"
Dim dr As String, sRoot As String

If bIsIDE Then Exit Sub


sRoot = DP_Dir_Path

If Right$(sRoot, 1) <> "\" Then sRoot = sRoot & "\"

On Error GoTo EH
dr = Dir$(sRoot & "*.*")

Do While LenB(dr)
    If dr <> "." Then
        If dr <> ".." Then
            'If LCase$(dr) <> "local.jpg" Then
                On Error GoTo EH
                Kill sRoot & dr
            'End If
        End If
    End If
    
    dr = Dir$()
Loop


Exit Sub
EH:
End Sub

Public Sub tmrMain_Timer()
Dim i As Integer, j As Integer
Dim Path As String
Dim bCanDistribute As Boolean, bContinue As Boolean

'If frmMain.mnuOptionsDPEnable.Checked = False Then Exit Sub



If (GameFormLoaded = False) And (StickFormLoaded = False) Then
    
    With frmMain.ucFileTransfer
        If LenB(.SaveDir) = 0 Then
            Path = frmMain.DP_Path
            
            If FileExists(Path, vbDirectory) = False Then
                On Error Resume Next
                MkDir Path
            End If
            
            .SaveDir = Path
        End If
    End With
    
    
    
    If Server Then
        
        If frmMain.ucFileTransfer.iCurStatus = tDisconnected Then
            frmMain.ucFileTransfer.Listen modPorts.DPPort
        End If
        
        
        'check we have all the pictures loaded
        For i = 1 To UBound(Clients)
            If Clients(i).bDPSet Then
                If Clients(i).IPicture Is Nothing Then
                    'check for the file
                    'Path = DP_Dir_Path & "\" & CStr(Clients(i).iSocket) & ".jpg"
                    Path = modDP.GetClientDPStr(i)
                    
                    
                    If FileExists(Path) Then
                        Set Clients(i).IPicture = LoadPicture(Path)
                        
                        frmMain.ShowDP i
                    End If
                End If
            End If
        Next i
        
        
        bCanDistribute = True 'assume we should try to send our picture
        
        If frmMain.ucFileTransfer.iCurSockStatus <> sckConnected Then
            
            'find a client whos picture we haven't got, and obtain it
            For i = 1 To UBound(Clients)
                If Clients(i).bDPSet Then
                    If Clients(i).IPicture Is Nothing Then
                        'grab it and exit for (enough transfering for one round)
                        SendData eCommands.FileTransferCmd & eFTCmds.FT_SendDPToHost, Clients(i).iSocket
                        
                        If frmMain.ucFileTransfer.iCurStatus <> tlistening Then
                            frmMain.ucFileTransfer.Listen modPorts.DPPort
                        End If
                        
                        Exit Sub
                    End If
                End If
            Next i
            
            
            'start to send
            sPictureToSendPath = DP_Dir_Path & "\Local.jpg"
            If FileExists(sPictureToSendPath) = False Then
                sPictureToSendPath = DP_Dir_Path & "\Local.gif"
            End If
            
            If FileExists(sPictureToSendPath) Then
                For i = 1 To UBound(Clients)
                    If Clients(i).bSentHostDP = False Then
                        
                        modDP.sRemoteFileName = "-1." & Right$(sPictureToSendPath, 3)
                        modDP.DP_iClient = i
                        modDP.bSetHostBool = True
                        
                        SendData eCommands.FileTransferCmd & eFTCmds.FT_ConnectToHost, Clients(i).iSocket
                        bCanDistribute = False
                        
                        If frmMain.ucFileTransfer.iCurStatus <> tlistening Then
                            frmMain.ucFileTransfer.Listen modPorts.DPPort
                        End If
                        
                        Exit For
                    End If
                Next i
            End If
        Else
            bCanDistribute = False
        End If
        
        If bCanDistribute Then
            'send other client's pictures
            
            bContinue = True
            
            For i = 1 To UBound(Clients)
                
                'find out whos DPs each client doesn't have + send
                For j = 1 To UBound(Clients)
                    If Clients(j).bDPSet Then
                        If j <> i Then
                            If InStr(1, Clients(i).sHasiDPs, CStr(j)) = 0 Then
                                'they don't have DP number j, send it
                                
                                modDP.sPictureToSendPath = modDP.GetClientDPStr(j) 'modDP.DP_Dir_Path & "\" & CStr(Clients(j).iSocket) & ".jpg"
                                
                                modDP.sRemoteFileName = CStr(Clients(j).iSocket) & "." & Right$(modDP.sPictureToSendPath, 3)
                                
                                modDP.DP_iClient = i
                                modDP.iToAdd = Clients(j).iSocket
                                modDP.bSetHostBool = False
                                
                                SendData eCommands.FileTransferCmd & eFTCmds.FT_ConnectToHost, Clients(i).iSocket
                                bContinue = False
                                
                                If frmMain.ucFileTransfer.iCurStatus <> tlistening Then
                                    frmMain.ucFileTransfer.Listen modPorts.DPPort
                                End If
                                
                                Exit For
                            End If
                        End If
                    End If
                Next j
                If bContinue = False Then Exit For
            Next i
        End If
        
        
        
    Else
        
        If Not bSentMyPicture Then
            If SendMyDisplayPicture(frmMain.SckLC.RemoteHostIP) Then
                bSentMyPicture = True
            End If
        End If 'sent mypic endif
            
'        ElseIf frmMain.ucFileTransfer.iCurStatus = tDisconnected Then
'
'            With frmMain.ucFileTransfer
'                If LenB(.SaveDir) = 0 Then
'                    Path = frmMain.DP_Path
'
'                    If FileExists(Path, vbDirectory) = False Then
'                        On Error Resume Next
'                        MkDir Path
'                    End If
'
'                    .SaveDir = Path
'                    frmMain.ucFileTransfer.Listen
'                End If
'            End With
            
'        Else
'
'            'request a picture
'            For i = 0 To UBound(Clients)
'                If Clients(i).IPicture Is Nothing Then
'                    SendData eCommands.SetClientVar & eClientVarCmds.SetRequestedDP & Clients(i).iSocket
'                    Exit For
'                End If
'            Next i
            
            
    End If 'server endif
End If 'game endif

End Sub

Private Function SendMyDisplayPicture(sTo As String) As Boolean
Dim MySock As Integer, i As Integer
Const TimeOut As Long = 1000
Dim Tick As Long

If LenB(My_DP_Path) Then
    If FileExists(My_DP_Path) Then
        
        If modMessaging.MySocket = 0 Then
            MySock = -1
            
            For i = 1 To UBound(Clients)
                If Clients(i).sName = frmMain.LastName Then
                    MySock = Clients(i).iSocket
                    modMessaging.MySocket = MySock
                    Exit For
                End If
            Next i
        Else
            MySock = modMessaging.MySocket
        End If
        
        If MySock <> -1 Then
            frmMain.ucFileTransfer.Connect sTo, modPorts.DPPort
            
            Tick = GetTickCount()
            
            Do
                Pause 10
            Loop While (frmMain.ucFileTransfer.iCurStatus <> tConnected) And _
                       (Tick + TimeOut > GetTickCount()) And _
                       Not modVars.Closing
            
            If modVars.Closing Then Exit Function
            
            
            If frmMain.ucFileTransfer.iCurStatus = tConnected Then
                
                SendMyDisplayPicture = frmMain.ucFileTransfer.SendFile( _
                    My_DP_Path, CStr(MySock) & "." & Right$(My_DP_Path, 3))
                
                
            End If
            
            'they'll disconnect
            'frmMain.ucFileTransfer.Disconnect
            
        End If
        
    End If
End If

End Function

'Public Sub SendDisplayPicture(iSendToClient As Integer, iDPClient As Integer)
'Const TimeOut As Long = 1000
'Dim i As Integer, iDPSock As Integer, iSendToSock As Integer
'Dim Tick As Long
'Dim DP_Path As String
'
'On Error GoTo EH
'iSendToSock = Clients(iSendToClient).iSocket
'iDPSock = Clients(iDPClient).iSocket
'
'
'If iSendToSock And iDPSock Then
'    DP_Path = DP_Dir_Path & CStr(iDPSock) & ".jpg"
'
'    If LenB(DP_Path) Then
'        If FileExists(DP_Path) Then
'
'                frmMain.ucFileTransfer.Connect Clients(iSendToClient).sIP
'
'                Tick = GetTickCount()
'                Do
'                    Pause 10
'                Loop While (frmMain.ucFileTransfer.iCurStatus <> tConnected) And _
'                           (Tick + TimeOut > GetTickCount()) And _
'                           Not modVars.Closing
'
'
'                If modVars.Closing Then Exit Sub
'
'
'                If frmMain.ucFileTransfer.iCurStatus = tConnected Then
'
'                    frmMain.ucFileTransfer.SendFile DP_Path, CStr(iDPSock) & ".jpg"
'
'                End If
'
'                'they'll disconnect
'                'frmMain.ucFileTransfer.Disconnect
'
'            End If
'
'        End If
'    End If
'End If
'
'EH:
'End Sub
