Attribute VB_Name = "modNetwork"
Option Explicit

'http://www.developerfusion.com/code/3169/list-network-computers/

'API Declarations & Constants
Private Const NERR_Success = 0&
Private Const NERR_Access_Denied = 5&
Private Const NERR_MoreData = 234&

Private Const SRV_TYPE_SERVER = &H2
Private Const SRV_TYPE_SQLSERVER = &H4
Private Const SRV_TYPE_NT_PDC = &H8
Private Const SRV_TYPE_NT_BDC = &H10
Private Const SRV_TYPE_PRINT = &H200
Private Const SRV_TYPE_NT = &H1000
Private Const SRV_TYPE_RAS = &H400
Public Const SRV_TYPE_ALL = &HFFFF

Private Const SHORT_LEVEL = 10&
Private Const EXTENDED_LEVEL = 3&

Private Const USER_ACC_NOPWD_CHANGE = 577&
Private Const USER_ACC_NOPWD_EXPIRE = 66049
Private Const USER_ACC_DISABLED = 515&
Private Const USER_ACC_LOCKED = 529&

Private Type SERVER_INFO_API
    PlatformId As Long
    ServerName As Long
    Type As Long
    VerMajor As Long
    VerMinor As Long
    Comment As Long
End Type

Private Type WKSTA_INFO_API
    PlatformId As Long
    ComputerName As Long
    LanGroup As Long
    VerMajor As Long
    VerMinor As Long
    LanRoot As Long
End Type

Private Type ServerInfo
    PlatformId As Long
    ServerName As String
    Type As Long
    VerMajor As Long
    VerMinor As Long
    Comment As String
    Platform As String
    ServerType As Integer
    LanGroup As String
    LanRoot As String
End Type

Public Type ListOfServer
    Init As Boolean
    LastErr As Long
    List() As ServerInfo
End Type

Private Type USER_INFO_EXT_API
    Name As Long
    Password As Long
    PasswordAge As Long
    Privilege As Long
    HomeDir As Long
    Comment As Long
    flags As Long
    ScriptPath As Long
    AuthFlags As Long
    FullName As Long
    UserComment As Long
    Parms As Long
    Workstations As Long
    LastLogon As Long
    LastLogoff As Long
    AcctExpires As Long
    MaxStorage As Long
    UnitsPerWeek As Long
    LogonHours As Long
    BadPwCount As Long
    NumLogons As Long
    LogonServer As Long
    CountryCode As Long
    CodePage As Long
    UserID As Long
    PrimaryGroupID As Long
    Profile As Long
    HomeDirDrive As Long
    PasswordExpired As Long
End Type

Private Type UserInfoExt
    Name As String
    Password As String
    PasswordAge As String
    Privilege As Long
    HomeDir As String
    Comment As String
    flags As Long
    NoChangePwd As Boolean
    NoExpirePwd As Boolean
    AccDisabled As Boolean
    AccLocked As Boolean
    ScriptPath As String
    AuthFlags As Long
    FullName As String
    UserComment As String
    Parms As String
    Workstations As String
    LastLogon As Date
    LastLogoff As Date
    AcctExpires As Date
    MaxStorage As Long
    UnitsPerWeek As Long
    LogonHours(0 To 20) As Byte
    BadPwCount As Long
    NumLogons As Long
    LogonServer As String
    CountryCode As Long
    CodePage As Long
    UserID As Long
    PrimaryGroupID As Long
    Profile As String
    HomeDirDrive As String
    PasswordExpired As Boolean
End Type

Public Type ListOfUserExt
    Init As Boolean
    LastErr As Long
    List() As UserInfoExt
End Type

Private Declare Function lstrlenW Lib "kernel32" _
        (ByVal lpString As Long) As Long

Private Declare Function NetApiBufferFree Lib "netapi32" _
        (ByVal lBuffer&) As Long

Private Declare Function NetGetDCName Lib "netapi32" _
        (lpServer As Any, lpDomain As Any, _
         vBuffer As Any) As Long

Private Declare Function NetServerEnum Lib "netapi32" _
        (lpServer As Any, ByVal lLevel As Long, vBuffer As Any, _
         lPreferedMaxLen As Long, lEntriesRead As Long, lTotalEntries As Long, _
         ByVal lServerType As Long, ByVal sDomain$, vResume As Any) As Long

Private Declare Function NetUserEnum Lib "netapi32" _
        (lpServer As Any, ByVal Level As Long, _
         ByVal Filter As Long, lpBuffer As Long, _
         ByVal PrefMaxLen As Long, lpEntriesRead As Long, _
         lpTotalEntries As Long, lpResumeHandle As Long) As Long

'Speed/Bandwidth stuff
Private lBytesRecv     As Long
Private lBytesSent     As Long

Private rNew As Single, sNew As Single
Private rOld As Single, sOld As Single

Private cIP As clsIPHelper

Public Function EnumServer(lServerType As Long, ByRef sError As String) As ListOfServer
    Dim nRet As Long, X As Integer, i As Integer
    Dim lRetCode As Long
    Dim tServerInfo As SERVER_INFO_API
    Dim lServerInfo As Long
    Dim lServerInfoPtr As Long
    Dim ServerInfo As ServerInfo
    Dim lPreferedMaxLen As Long
    Dim lEntriesRead As Long
    Dim lTotalEntries As Long
    Dim sDomain As String
    Dim vResume As Variant
    Dim yServer() As Byte
    Dim SrvList As ListOfServer
    
    yServer = MakeServerName(ByVal "")
    lPreferedMaxLen = 65536
    
    sError = vbNullString
    
    nRet = NERR_MoreData
    Do While (nRet = NERR_MoreData)
        
        'Call NetServerEnum to get a list of Servers
        nRet = NetServerEnum(yServer(0), 101, lServerInfo, _
                             lPreferedMaxLen, lEntriesRead, _
                             lTotalEntries, lServerType, _
                             sDomain, vResume)
        
        
        If (nRet <> NERR_Success And nRet <> NERR_MoreData) Then
            SrvList.Init = False
            SrvList.LastErr = nRet
            sError = NetError(nRet)
            Exit Do
        End If
        
        ' NetServerEnum Index is 1 based
        X = 1
        lServerInfoPtr = lServerInfo
        
        Do While X <= lTotalEntries
            
            CopyMemory tServerInfo, ByVal lServerInfoPtr, Len(tServerInfo)
            
            ServerInfo.Comment = PointerToStringW(tServerInfo.Comment)
            ServerInfo.ServerName = PointerToStringW(tServerInfo.ServerName)
            ServerInfo.Type = tServerInfo.Type
            ServerInfo.PlatformId = tServerInfo.PlatformId
            ServerInfo.VerMajor = tServerInfo.VerMajor
            ServerInfo.VerMinor = tServerInfo.VerMinor
            
            i = i + 1
            ReDim Preserve SrvList.List(1 To i) As ServerInfo
            SrvList.List(i) = ServerInfo
            
            X = X + 1
            lServerInfoPtr = lServerInfoPtr + Len(tServerInfo)
            
        Loop
        
        lRetCode = NetApiBufferFree(lServerInfo)
        SrvList.Init = (X > 1)
        
    Loop
    
    EnumServer = SrvList
    
End Function

Private Function GetPDCName() As String
    Dim lpBuffer As Long, nRet As Long
    Dim yServer() As Byte
    Dim sLocal As String
    
    yServer = MakeServerName(ByVal "")
    
    nRet = NetGetDCName(yServer(0), yServer(0), lpBuffer)
    
    If nRet = 0 Then
        sLocal = PointerToStringW(lpBuffer)
    End If
    
    If lpBuffer Then Call NetApiBufferFree(lpBuffer)
    
    GetPDCName = sLocal
    
End Function

'Function Read User Information - for future development!
Public Function LongEnumUsers(Server As String, ByRef sError As String) As ListOfUserExt
    Dim yServer() As Byte, lRetCode As Long
    Dim nRead As Long, nTotal As Long
    Dim nRet As Long, nResume As Long
    Dim PrefMaxLen As Long
    Dim i As Long, X As Long
    Dim lUserInfo As Long
    Dim lUserInfoPtr As Long
    Dim UserInfo As UserInfoExt
    Dim UserList As ListOfUserExt
    Dim tUserInfo As USER_INFO_EXT_API
    
    yServer = MakeServerName(ByVal Server)
    PrefMaxLen = 65536
    
    nRet = NERR_MoreData
    Do While (nRet = NERR_MoreData)
        
        nRet = NetUserEnum(yServer(0), EXTENDED_LEVEL, 2, _
                           lUserInfo, PrefMaxLen, nRead, _
                           nTotal, nResume)
        
        If (nRet <> NERR_Success And nRet <> NERR_MoreData) Then
            UserList.Init = False
            UserList.LastErr = nRet
            sError = NetError(nRet)
            Exit Do
        End If
        
        lUserInfoPtr = lUserInfo
        
        X = 1
        Do While X <= nRead
            
            CopyMemory tUserInfo, ByVal lUserInfoPtr, Len(tUserInfo)
            
            UserInfo.Name = PointerToStringW(tUserInfo.Name)
            UserInfo.Password = PointerToStringW(tUserInfo.Password)
            UserInfo.PasswordAge = Format(tUserInfo.PasswordAge / 86400, "0.0")
            UserInfo.Privilege = tUserInfo.Privilege
            UserInfo.HomeDir = PointerToStringW(tUserInfo.HomeDir)
            UserInfo.Comment = PointerToStringW(tUserInfo.Comment)
            UserInfo.flags = tUserInfo.flags
            UserInfo.NoChangePwd = CBool((tUserInfo.flags Or USER_ACC_NOPWD_CHANGE) = tUserInfo.flags)
            UserInfo.NoExpirePwd = CBool((tUserInfo.flags Or USER_ACC_NOPWD_EXPIRE) = tUserInfo.flags)
            UserInfo.AccDisabled = CBool((tUserInfo.flags Or USER_ACC_DISABLED) = tUserInfo.flags)
            UserInfo.AccLocked = CBool((tUserInfo.flags Or USER_ACC_LOCKED) = tUserInfo.flags)
            UserInfo.ScriptPath = PointerToStringW(tUserInfo.ScriptPath)
            UserInfo.AuthFlags = tUserInfo.AuthFlags
            UserInfo.FullName = PointerToStringW(tUserInfo.FullName)
            UserInfo.UserComment = PointerToStringW(tUserInfo.UserComment)
            UserInfo.Parms = PointerToStringW(tUserInfo.Parms)
            UserInfo.Workstations = PointerToStringW(tUserInfo.Workstations)
            UserInfo.LastLogon = NetTimeToVbTime(tUserInfo.LastLogon)
            UserInfo.LastLogoff = NetTimeToVbTime(tUserInfo.LastLogoff)
            If tUserInfo.AcctExpires = -1& Then
                UserInfo.AcctExpires = NetTimeToVbTime(0)
            Else
                UserInfo.AcctExpires = NetTimeToVbTime(tUserInfo.AcctExpires)
            End If
            UserInfo.MaxStorage = tUserInfo.MaxStorage
            UserInfo.UnitsPerWeek = tUserInfo.UnitsPerWeek
            CopyMemory UserInfo.LogonHours(0), ByVal tUserInfo.LogonHours, 21
            UserInfo.BadPwCount = tUserInfo.BadPwCount
            UserInfo.NumLogons = tUserInfo.NumLogons
            UserInfo.LogonServer = PointerToStringW(tUserInfo.LogonServer)
            UserInfo.CountryCode = tUserInfo.CountryCode
            UserInfo.CodePage = tUserInfo.CodePage
            UserInfo.UserID = tUserInfo.UserID
            UserInfo.PrimaryGroupID = tUserInfo.PrimaryGroupID
            UserInfo.Profile = PointerToStringW(tUserInfo.Profile)
            UserInfo.HomeDirDrive = PointerToStringW(tUserInfo.HomeDirDrive)
            UserInfo.PasswordExpired = CBool(tUserInfo.PasswordExpired)
            
            i = i + 1
            ReDim Preserve UserList.List(1 To i) As UserInfoExt
            UserList.List(i) = UserInfo
            X = X + 1
            
            lUserInfoPtr = lUserInfoPtr + Len(tUserInfo)
            
        Loop
        
        lRetCode = NetApiBufferFree(lUserInfo)
        UserList.Init = (X > 1)
        
    Loop
    
    LongEnumUsers = UserList
    
End Function

Private Function MakeServerName(ByVal ServerName As String)
    Dim yServer() As Byte
    
    If ServerName <> "" Then
        If InStr(1, ServerName, "\\") = 0 Then
            ServerName = "\\" & ServerName
        End If
    End If
    
    yServer = ServerName & vbNullChar
    MakeServerName = yServer
    
End Function

Private Function NetError(nErr As Long) As String
    Dim Mesage As String
    
    
    Select Case nErr
        Case 5
            Mesage = "Access Denied"
        Case 1722
            Mesage = "Server not Accessible"
        'Case 1326
            'Message = " Sie besitzen nicht die Berechtigungen dafür"
        Case Else
            Mesage = modVars.DllErrorDescription(nErr) '"Error Number " & nErr & " has occured."
    End Select
    
    NetError = Mesage
    
End Function

Private Function NetTimeToVbTime(NetDate As Long) As Double
    Const BaseDate# = 25569   'DateSerial(1970, 1, 1)
    Const SecsPerDay# = 86400
    Dim Tmp As Double
    
    Tmp = BaseDate + (CDbl(NetDate) / SecsPerDay)
    If Tmp <> BaseDate Then
        NetTimeToVbTime = Tmp
    End If
    
End Function

Private Function PointerToStringW(lpStringW As Long) As String
    Dim Buffer() As Byte
    Dim nLen As Long
    
    If lpStringW Then
        nLen = lstrlenW(lpStringW) * 2
        If nLen Then
            ReDim Buffer(0 To (nLen - 1)) As Byte
            CopyMemory Buffer(0), ByVal lpStringW, nLen
            PointerToStringW = Buffer
        End If
    End If
End Function

'########################################################################

Public Sub InitBandwidthStuff(ByVal bInit As Boolean)
If bInit Then
    Set cIP = New clsIPHelper
    GetSpeeds 0, 0
Else
    Set cIP = Nothing
End If
End Sub

Public Sub GetSpeeds(ByRef DownSpeed As Single, ByRef UpSpeed As Single)
Dim rValue As Single, sValue As Single
Static Last As Long

If cIP.GetSpeeds() Then
    
    lBytesRecv = cIP.BytesReceived
    lBytesSent = cIP.BytesSent
    
    rNew = lBytesRecv
    sNew = lBytesSent
    
    If Last = 0 Then Last = GetTickCount()
    
    rValue = (rNew - rOld) '/ (GetTickCount() - Last)
    sValue = (sNew - sOld) '/ (GetTickCount() - Last)
    rOld = rNew
    sOld = sNew
    
    
    
    DownSpeed = rValue / 1000
    UpSpeed = sValue / 1000
End If

End Sub
