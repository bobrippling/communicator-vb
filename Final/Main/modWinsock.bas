Attribute VB_Name = "modWinsock"
Option Explicit

''winsock control timeout stuff
'Private Declare Function setsockopt Lib "wsock32.dll" ( _
'    ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, _
'    ByVal optlen As Long) As Long
'
'Private Const SO_SNDTIMEO = &H1005
'Private Const SO_RCVTIMEO = &H1006
'Private Const IPPROTO_TCP = 6

'Winsock info structure
Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Long
    iMaxUdpDg As Long
    lpVendorInfo As Long
End Type


'Socket info structure
Public Type ptSockAddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
'Public Type IN_ADDR
'    S_addr As Long
'End Type
'
'Public Type SOCK_ADDR
'    sin_family As Integer
'    sin_port As Integer
'    sin_addr As IN_ADDR
'    sin_zero(0 To 7) As Byte
'End Type


Private m_udtWskData As WSADATA


Private Const PF_INET = 2                'Internet address format
Private Const SOCK_DGRAM = 2             'UDP format
Private Const SOCK_STREAM = 1            'TCP format
Private Const SOCKET_ERROR = -1          'Error return value
Private Const SOCK_VER As Integer = 514  'Version of winsock to use
Private Const FIONBIO = &H8004667E       'Set socket to non-blocking
Private Const MAX_MESSAGE_LENGTH = 1300  'Max possible size of packets
Private Const DEFAULT_PROTOCOL = 0
Public Const WINSOCK_ERROR = SOCKET_ERROR          'A custom error value we'll return when one of our function calls fails

'API calls
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequested As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function socket Lib "wsock32.dll" (ByVal af As Long, ByVal prototype As Long, ByVal Protocol As Long) As Long
Private Declare Function bind Lib "wsock32.dll" (ByVal S As Long, bindto As ptSockAddr, ByVal tolen As Long) As Long
Private Declare Function closesocket Lib "wsock32.dll" (ByVal S As Long) As Long
Private Declare Function SendTo Lib "wsock32.dll" Alias "sendto" (ByVal S As Long, buf As Any, ByVal Length As Long, ByVal flags As Long, addrto As ptSockAddr, ByVal tolen As Long) As Long
Private Declare Function recvfrom Lib "wsock32.dll" (ByVal S As Long, ByRef buf As Any, ByVal Length As Integer, ByVal flags As Integer, addrfrom As ptSockAddr, fromlen As Integer) As Integer
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal Addr As String) As Long
Private Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
Private Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Integer) As Integer

'TCP
Private Const FD_SETSIZE = 64
Private Type FD_SET
    fd_count As Long
    fd_array(0 To FD_SETSIZE - 1) As Long
End Type
Private Type TIME_VAL
    tv_sec As Long
    tv_usec As Long
End Type

Private Const SD_BOTH = 2&

Private Declare Function apiConnect Lib "wsock32" Alias "connect" (ByVal S As Long, Name As ptSockAddr, _
    ByVal namelen As Integer) As Long
Private Declare Function apiSend Lib "wsock32" Alias "send" (ByVal S As Long, Buffer As Any, _
    ByVal Length As Long, ByVal flags As Long) As Long
Private Declare Function apiSelect Lib "wsock32" Alias "select" (ByVal nfds As Long, readfds As FD_SET, _
    writefds As FD_SET, exceptfds As FD_SET, TimeOut As TIME_VAL) As Long
Private Declare Function apiShutdown Lib "wsock32" Alias "shutdown" (ByVal S As Long, _
    ByVal how As Long) As Long
Private Declare Function apiRecv Lib "wsock32" Alias "recv" (ByVal S As Long, ByVal buf As String, _
  ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long

'############################################################################################################
Private Declare Function apiGetIpAddrTable Lib "iphlpapi" Alias "GetIpAddrTable" ( _
    pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

Const MAX_IP = 5

Private Type IPINFO
    dwAddr As Long   'IP address
    dwIndex As Long 'interface index
    dwMask As Long 'subnet mask
    dwBCastAddr As Long 'broadcast address
    dwReasmSize As Long 'assembly size
    unused1 As Integer 'not currently used
    unused2 As Integer 'not currently used
End Type

Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type

Private IPAddrInfo As MIB_IPADDRTABLE
Private psLocalIP As String, psRemoteIP As String


Private Const LocalHost_IP As Long = &H100007F
'Private Type IP_Array
'    mBuffer As MIB_IPADDRTABLE
'    BufferLen As Long
'End Type

'###################################################################################################
Private Const AF_INET As Long = 2

Private Declare Function apiGetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal szHost As String) As Long
Private Declare Function apiGetHostByAddr Lib "wsock32" Alias "gethostbyaddr" (haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function apiStrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Any) As Long

'Private Declare Function apiGetNameInfo Lib "Ws2_32.dll" Alias "getnameinfo" ( _
    sa As Any, ByVal salen As Long, ByRef host As String, ByVal hostlen As Long, _
    ByRef Serv As String, ByVal servlen As Long, ByVal Flags As Long) As Long

Private Type ptHOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

'###################################################################################################

Public Function InitWinsock() As Long

Dim lngRetVal As Long

'Initialize the session
lngRetVal = WSAStartup(SOCK_VER, m_udtWskData)

'Check for errors
If lngRetVal Then
    AddConsoleText "Unable to initialise Winsock session" ', TxtError, True
    InitWinsock = WINSOCK_ERROR
Else
    AddConsoleText "Initialising Winsock...", , True, , True
    With m_udtWskData
        AddConsoleText "Version: " & CStr(.wVersion)
        AddConsoleText "Description: " & Trim0(.szDescription)
        'AddConsoleText "Vendor: " & CStr(.lpVendorInfo)
        'AddConsoleText "Max Sockets: " & CStr(.iMaxSockets)
        AddConsoleText "System Status: " & Trim0(.szSystemStatus)
    End With
    AddConsoleText "Initialised Winsock", , , True
    AddConsoleText vbNullString
    
    InitWinsock = 0
End If


Call InitIPAddr

End Function

Public Function CreateSocket() As Long

Dim lngSocket As Long

'Create a UDP socket
lngSocket = socket(PF_INET, SOCK_DGRAM, DEFAULT_PROTOCOL)

'Check for errors
If lngSocket = SOCKET_ERROR Then
    AddText "Unable to open a socket", TxtError, True
    CreateSocket = WINSOCK_ERROR
    Exit Function
End If

'Set the socket to non-blocking
ioctlsocket lngSocket, FIONBIO, 1

'Assign the socket
CreateSocket = lngSocket

End Function

Public Function MakeSockAddr(ByRef udtSockAddr As ptSockAddr, ByRef intPort As Integer, ByRef strIP As String) As Integer

'Set the address format
udtSockAddr.sin_family = PF_INET

'Convert to network byte order (using htons) and set port
udtSockAddr.sin_port = htons(intPort)

'Specify the IP (and resolve[?])
udtSockAddr.sin_addr = inet_addr(strIP)

'Zero fill
udtSockAddr.sin_zero = String$(8, vbNullChar)

'Check for error in IP
If udtSockAddr.sin_addr = SOCKET_ERROR Then
    'Notify user
    AddText "Unable to resolve IP address", TxtError, True  ' from " & strIP, TxtError, True
    MakeSockAddr = WINSOCK_ERROR
    'Exit!
    Exit Function
End If

End Function

Public Function BindSocket(ByRef lngSocket As Long, ByRef intPort As Integer) As Integer

Dim udtSocketInfo As ptSockAddr
Dim lngRetVal As Long

'Set the address format
udtSocketInfo.sin_family = PF_INET

'Convert to network byte order (using htons) and set port
udtSocketInfo.sin_port = htons(intPort)

'Specify local IP
udtSocketInfo.sin_addr = 0

'Zero fill
udtSocketInfo.sin_zero = String$(8, vbNullChar)

'Bind
lngRetVal = bind(lngSocket, udtSocketInfo, Len(udtSocketInfo))

'have we got error 10048 - addr in use?

'Check for errors
If lngRetVal = SOCKET_ERROR Then
    AddText "Error " & Err.LastDllError & " - Unable to bind socket on port " & CStr(intPort), TxtError, True
    BindSocket = WINSOCK_ERROR
End If

End Function

Public Function SendPacket(ByRef lngSocket As Long, ByRef udtSockAddr As ptSockAddr, _
    ByRef strMessage As String) As Boolean

'Dim udtSocketInfo As ptSockAddr
Dim lngRetVal As Long

'Send mPacket
lngRetVal = SendTo(lngSocket, ByVal strMessage, Len(strMessage), 0, udtSockAddr, Len(udtSockAddr))

'Check for errors
SendPacket = (lngRetVal <> SOCKET_ERROR)

End Function

Public Function ReceivePacket(ByRef lngSocket As Long, ByRef udtSockAddr As ptSockAddr) As String

Dim strMessage As String
Dim lngRetVal As Long

'Space the string
strMessage = Space$(MAX_MESSAGE_LENGTH)

'Check for a message
lngRetVal = recvfrom(lngSocket, ByVal strMessage, Len(strMessage), 0, udtSockAddr, Len(udtSockAddr))

'Check for errors
If lngRetVal <> SOCKET_ERROR Then
    'Return the message
    ReceivePacket = Left$(strMessage, lngRetVal)
'Else
    ''There was no message waiting
    'ReceivePacket = vbNullString
End If

End Function

Public Sub DestroySocket(ByRef lngSocket As Long)

'Close the specified socket
closesocket lngSocket

lngSocket = 0

End Sub

Public Sub TermWinsock()

'End the session and release resources
WSACleanup
'affects scks

End Sub

'Public Sub SetWinsockTimeout(lhSocket As Long, lSendTimeout As Long, lRecvTimeout As Long)
'Dim lRet As Long
'
''4 = byte length of a long
'If lhSocket > 0 Then
'    lRet = setsockopt(lhSocket, IPPROTO_TCP, SO_SNDTIMEO, lSendTimeout, LenB(lSendTimeout))
'    lRet = setsockopt(lhSocket, IPPROTO_TCP, SO_RCVTIMEO, ByVal lRecvTimeout, LenB(lRecvTimeout))
'End If
'
'
'End Sub
'########################################################################################

Public Function IPToHostName(sIP As String) As String

Dim lp As Long
Dim lAddress As Long
Dim sHostName As String
Dim lLength As Integer

'convert string address to long datatype
lAddress = inet_addr(sIP)

If lAddress <> SOCKET_ERROR Then
    
    lp = apiGetHostByAddr(lAddress, 4, AF_INET)
    
    
    If lp <> 0 Then
        CopyMemory lp, ByVal lp, 4
        
        lLength = apiStrLen(lp)
        If lLength > 0 Then
            sHostName = Space$(lLength)
            
            CopyMemory ByVal sHostName, ByVal lp, lLength
            
            IPToHostName = sHostName
        End If
    End If
End If

End Function

Public Function HostNameToIP(ByVal sHostName As String) As String
'converts a host name to an IP address.

Dim ptrHosent As Long
Dim hstHost As ptHOSTENT
Dim ptrIPAddress As Long
Dim sAddress As String  'declare this as Dim sAddress(1) As String if you want 2 ip addresses returned
Dim i As Integer
Dim sTmp As String

'try to get the IP
ptrHosent = apiGetHostByName(sHostName & vbNullChar)

If ptrHosent <> 0 Then
            
    'get the IP address
    CopyMemory hstHost, ByVal ptrHosent, LenB(hstHost)
    CopyMemory ptrIPAddress, ByVal hstHost.hAddrList, 4
      
    'fill buffer
    sAddress = Space$(4)
    'if you want multiple domains returned,
    'fill all items in sAddress array with 4 spaces
    
    CopyMemory ByVal sAddress, ByVal ptrIPAddress, 4 'hstHost.hLength
    
    'change this to
    'CopyMemory ByVal sAddress(0), ByVal ptrIPAddress, hstHost.hLength
    'if you want an array of ip addresses returned
    '(some domains have more than one ip address associated with it)
    
    For i = 1 To 4
        sTmp = sTmp & CStr(Asc(Mid$(sAddress, i, 1))) & Dot
    Next i
    
    'get the IP address
    HostNameToIP = Left$(sTmp, Len(sTmp) - 1) 'AddressToString(lAddress)
    'if you are using multiple addresses, you need IPToText(sAddress(0)) & "," & IPToText(sAddress(1))
    'etc
End If

End Function

'Public Function IPToHostName(sIP As String) As String
'Dim tAddress As ptsockaddr
'Dim sHostName As String
'Dim lR As Long
'
'MakeSockAddr tAddress, 2850, sIP
'sHostName = String$(255, vbNullChar)
'
'
'lR = apiGetNameInfo(tAddress, Len(tAddress), sHostName, Len(sHostName), ByVal vbNullString, 0, 0)
'
'
'IPToHostName = sHostName
'
''11004 = WSANO_DATA  -  http://msdn.microsoft.com/en-us/library/ms740668(VS.85).aspx
'
'End Function


'########################################################################################
'non-socket stuff

Public Property Get LocalIP() As String
LocalIP = psLocalIP
End Property
Public Property Get RemoteIP() As String
RemoteIP = psRemoteIP
End Property

Public Function SetLocalIP(ByVal sIP As String) As Boolean
If IsIP(sIP) Then
    psLocalIP = sIP
    SetLocalIP = True
Else
    SetLocalIP = False
End If
End Function

Public Function GetHTML(ByVal Url As String) As String
Dim Tmp As String

Tmp = modFTP.GetHTML(Url)

GetHTML = Trim$(Tmp)
End Function

Public Function GetIP() As String
Const URL2 As String = "http://www.whatismyip.org/"
Const URL1 As String = "http://www.whatismyip.com/automation/n09230945.asp"
'const URL3 as string = "http://checkip.dyndns.org/" - not just an ip, though

Dim Str As String

Str = GetHTML(URL1)

If IsIP(Str) Then
    GetIP = Trim$(Str)
Else
    Str = GetHTML(URL2)
    If IsIP(Str) Then
        GetIP = Trim$(Str)
    Else
        GetIP = vbNullString
    End If
End If

End Function

Public Function ObtainRemoteIP() As Boolean

If LenB(psRemoteIP) = 0 Then
    psRemoteIP = GetIP()
End If

ObtainRemoteIP = CBool(LenB(psRemoteIP))

If psLocalIP = AddressToString(LocalHost_IP) Then
    ObtainLocalIP 'attempt to re-obtain
End If

End Function
Public Sub ObtainLocalIP()
Dim i As Integer

If GetIPAddrTable(IPAddrInfo) Then
    
    For i = 0 To IPAddrInfo.dEntrys - 1
        If IPAddrInfo.mIPInfo(i).dwAddr > 0 Then 'prevent 0.0.0.0
            If IPAddrInfo.mIPInfo(i).dwAddr <> LocalHost_IP Then 'prevent 127.0.0.1
                psLocalIP = AddressToString(IPAddrInfo.mIPInfo(i).dwAddr)
                Exit For
            End If
        End If
    Next i
    
    If i = IPAddrInfo.dEntrys Then
        psLocalIP = "127.0.0.1"
    End If
    
End If

End Sub

Private Sub InitIPAddr()

ObtainLocalIP

'If modLoadProgram.bQuick = False Then
'    psRemoteIP = GetIP()
'End If

End Sub

Private Function GetIPAddrTable(ByRef IPAddrs As MIB_IPADDRTABLE) As Boolean

Dim buf() As Byte
Dim BUFSIZE As Long, lRet As Long, nEntries As Integer
Dim i As Integer, j As Integer
Dim S As String
Dim Listing As MIB_IPADDRTABLE


apiGetIpAddrTable ByVal 0&, BUFSIZE, 1

If BUFSIZE > 0 Then
    
    ReDim buf(0 To BUFSIZE - 1) As Byte
    
    lRet = apiGetIpAddrTable(buf(0), BUFSIZE, 0)
    
    
    If lRet <> 0 Then
        AddConsoleText "GetIpAddrTable() failed with return value " & CStr(lRet)
        GetIPAddrTable = False
    Else
        
        
        CopyMemory Listing.dEntrys, buf(0), 4
        
        For i = 0 To Listing.dEntrys - 1
            
            CopyMemory Listing.mIPInfo(i), buf(4 + (i * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(i))
            
            '"IP address: " & ConvertAddressToString(Listing.mIPInfo(i).dwAddr)
            '"IP Subnetmask: " & ConvertAddressToString(Listing.mIPInfo(i).dwMask)
            '"BroadCast IP address: " & ConvertAddressToString(Listing.mIPInfo(i).dwBCastAddr)
            
        Next i
        
        IPAddrs = Listing
        
        GetIPAddrTable = True
'        nEntries = buf(1) * 256 + buf(0)
'
'        If nEntries = 0 Then
'            ReDim IPAddrs(0)
'        Else
'
'            ReDim IPAddrs(0 To nEntries - 1) As String
'
'
'            For i = 0 To nEntries - 1
'                S = vbNullString
'
'                For j = 0 To 3
'                    S = S & IIf(j > 0, ".", vbNullString) & buf(4 + i * 24 + j)
'                Next j
'
'                IPAddrs(i) = S
'            Next i
'
'
'        End If
    End If
Else
    GetIPAddrTable = False
End If

End Function

Private Function AddressToString(longAddr As Long) As String
Dim myByte(3) As Byte
Dim i As Integer
Dim sTmp As String

CopyMemory myByte(0), longAddr, 4

For i = 0 To 3
    sTmp = sTmp + CStr(myByte(i)) + "."
Next i

AddressToString = Left$(sTmp, Len(sTmp) - 1)
End Function

'################################################################################
'################################################################################
'################################################################################
'TCP
Public Function TCP_CreateSocket() As Long

Dim lngSocket As Long

'Create a UDP socket
lngSocket = socket(PF_INET, SOCK_STREAM, DEFAULT_PROTOCOL)

'Check for errors
If lngSocket = SOCKET_ERROR Then
    TCP_CreateSocket = WINSOCK_ERROR
Else
    'Set the socket to non-blocking
    ioctlsocket lngSocket, FIONBIO, 1
    
    'Assign the socket
    TCP_CreateSocket = lngSocket
End If

End Function

Public Function TCP_BindSocket(ByRef lngSocket As Long, Optional ByRef intPort As Integer = 0) As Integer

Dim udtSocketInfo As ptSockAddr
Dim lngRetVal As Long


'Set the address format
udtSocketInfo.sin_family = PF_INET

'Convert to network byte order (using htons) and set port
udtSocketInfo.sin_port = htons(intPort)

'Specify local IP
udtSocketInfo.sin_addr = 0 '= use default/current ip

'Zero fill
udtSocketInfo.sin_zero = String$(8, vbNullChar)

'Bind
lngRetVal = bind(lngSocket, udtSocketInfo, Len(udtSocketInfo))

'have we got error 10048 - addr in use?

'Check for errors
If lngRetVal = SOCKET_ERROR Then
    TCP_BindSocket = WINSOCK_ERROR
End If

End Function

Public Function TCP_Connect(lSocket As Long, sIP As String, iPort As Integer) As Long
Dim Addr As ptSockAddr

With Addr
    .sin_family = PF_INET
    .sin_port = htons(iPort)
    .sin_addr = inet_addr(sIP)
    
    If .sin_addr = SOCKET_ERROR Then
        TCP_Connect = WINSOCK_ERROR
        Exit Function
    End If
End With
    

TCP_Connect = apiConnect(lSocket, Addr, Len(Addr))
'returns 0 is blocking+success
'returns -1 if non-blocking

End Function

Public Function TCP_Connected(lSocket As Long) As Boolean
Dim readfds As FD_SET, writefds As FD_SET, exceptfds As FD_SET
Dim TimeOut As TIME_VAL
Dim lR As Long, nfds As Long


nfds = 0
TimeOut.tv_sec = 1
TimeOut.tv_usec = 0 'microseconds

readfds.fd_count = 0

writefds.fd_count = 1 'socket is writable when connected
writefds.fd_array(0) = lSocket

exceptfds.fd_count = 1
exceptfds.fd_array(0) = lSocket


lR = apiSelect(nfds, readfds, writefds, exceptfds, TimeOut)

If lR = SOCKET_ERROR Then
    TCP_Connected = False
Else
    
    If writefds.fd_count > 0 Then
        If writefds.fd_array(0) = lSocket Then
            If exceptfds.fd_count = 0 Then
                TCP_Connected = True
            End If
        End If
    End If
End If


End Function

Public Function TCP_ReceiveData(lSocket As Long, ByRef sData As String, _
    Optional ByVal BufferLen As Long = 255) As Boolean

Dim lR As Long
Dim sBuff As String

sBuff = String$(BufferLen, 0)

lR = apiRecv(lSocket, sBuff, BufferLen, 0)

If lR <> SOCKET_ERROR Then
    TCP_ReceiveData = True
    sData = Left$(sBuff, lR)
Else
    TCP_ReceiveData = False
    sData = vbNullString
End If

End Function

'Public Function TCP_SendData(lSocket As Long, ByVal strData As String) As Boolean
'Dim WSAResult As Long, i As Long, l As Long
'
'l = Len(strData)
'ReDim Buff(l + 1) As Byte
'
'For i = 1 To l
'    Buff(i - 1) = Asc(Mid$(strData, i, 1))
'Next
'Buff(l) = 0 'asc(chr$(0))
'
'WSAResult = apiSend(lSocket, Buff(0), l, 0)
'TCP_SendData = (WSAResult <> SOCKET_ERROR)
'End Function

Public Function TCP_CloseSocket(lSocket As Long) As Boolean
Dim sBuff As String
Const bLen = 255&
Dim start_tick As Long

apiShutdown lSocket, SD_BOTH

start_tick = GetTickCount()
'call recv until 0 returned
Do
    sBuff = String$(bLen, 0)
Loop Until apiRecv(lSocket, sBuff, bLen, 0) = SOCKET_ERROR Or (start_tick + 1000 < GetTickCount())

closesocket lSocket
WSACancelBlockingCall
lSocket = 0

End Function
