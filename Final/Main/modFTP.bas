Attribute VB_Name = "modFTP"
Option Explicit

'################################################
#Const bDebug_Auto_FTP = False
'################################################

Private Const MAX_PATH As Integer = 260
Private Const UNKNOWN_ERROR_STRING = "Unknown Error, Please Retry"

'ftp types
Public Enum eFTP_Methods
    FTP_HTTP = 0
    FTP_Manual = 1
    FTP_Auto = 2
    FTP_Default = FTP_Auto 'can't be HTTP, since the Upload menu hasn't got a HTTP option
End Enum

Public Type ptFTPFile
    sName As String * MAX_PATH
    lFileSize As Long
    
    dDateLastWritten As Date
    dDateLastAccessed As Date
    dDateCreated As Date
End Type

'#############################################################################
'http
'manual download
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
    ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Const BINDF_GETNEWESTVERSION As Long = &H10 'don't use cache
Private Const INET_E_DOWNLOAD_FAILURE = &H800C0008, E_OUTOFMEMORY = &H8007000E


'#############################################################################
Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    ByRef lpdwError As Long, _
    ByVal lpszBuffer As String, _
    ByRef lpdwBufferLength As Long) As Long 'c_bool

'#############################################################################

Public Enum eFTPCustErrs
    cSuccess = -1
    cOther = 0
    cFileNotFoundOnLocal = 1
    cFileNotFoundOnServer = 2
End Enum
    
Public OnlineStatusIs As Boolean 'was online status set to true when uploaded?
Private pbFTP_Doing As Boolean

Public Const FTPControlStr As String = "MicRobSoft OpenUrl"
Private Const FTP_Wait_Time As Integer = 1500
'###################################################################################################
'###################################################################################################
'###################################################################################################
'String Constants

'Private Const DefaultHost As String = "ftp.byethost13.com"
'Private FTP_Host_Name As String'="ftp.byethost13.com"
'Private FTP_User_Name As String '= "b13_1256618"
'Private FTP_Password As String '= "communicator"
Public Type ptFTP_Details
    FTP_Host_Name As String
    FTP_User_Name As String
    FTP_Password As String
    
    FTP_Root As String
    FTP_File_Ext As String 'some servers don't like *.mcc
    'HTTP_Root As String
End Type
Public FTP_Details() As ptFTP_Details
Public FTP_iCustomServer As Integer
Public iCurrent_FTP_Details As Integer

Private pFTP_bStealth As Boolean, pFTP_bDeferToUpdateForm As Boolean


'Public FTP_Root_Location As String


Private Const HTTPRoot As String = "http://microbsoft.byethost13.com/" 'HTTP download method
'Public Const UpdateZip As String = "http://microbsoft.byethost13.com/Version/Communicator.zip" 'used in Manual Download
Public Const UpdateSite As String = "http://microbsoft.110mb.com/" 'website shelling

'Private FTP_IP_Remote_File As String 'IP File


'Private Const Default_FTP_Root_Location As String = "microbsoft.byethost13.com/htdocs"
'###################################################################################################
'###################################################################################################
'###################################################################################################

'Communicator Update Files
Public Const Communicator_File As String = "Communicator.zip"
'Private Const Local_UpdateTxt As String = "Version." & FileExt



'dlls
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
    (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
    ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function apiInternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
    (ByVal hInternetSession As Long, ByVal sServerName As String, _
    ByVal nServerPort As Integer, ByVal sUsername As String, _
    ByVal sPassword As String, ByVal lService As Long, _
    ByVal lFlags As Long, ByVal lContext As Long) As Long


Private Declare Function FTPGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
    (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
    ByVal lpszNewFile As String, ByVal fFailIfExists As Long, _
    ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Boolean


Private Declare Function FTPPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
    (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
    ByVal lpszRemoteFile As String, ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Boolean


'Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
    (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean

Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" ( _
    ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'###################################################################################################
'callbacks
Private Declare Function apiInternetSetStatusCallback Lib "wininet" Alias "InternetSetStatusCallback" ( _
    ByVal hInternet As Long, ByVal lpfnInternetCallback As Long) As Long

'internet callback messages
Private Const INTERNET_INVALID_STATUS_CALLBACK As Long = -1
Private Const INTERNET_STATUS_RESOLVING_NAME As Long = 10
Private Const INTERNET_STATUS_NAME_RESOLVED As Long = 11
Private Const INTERNET_STATUS_CONNECTING_TO_SERVER As Long = 20
Private Const INTERNET_STATUS_CONNECTED_TO_SERVER As Long = 21
Private Const INTERNET_STATUS_SENDING_REQUEST As Long = 30
Private Const INTERNET_STATUS_REQUEST_SENT As Long = 31
Private Const INTERNET_STATUS_RECEIVING_RESPONSE As Long = 40
Private Const INTERNET_STATUS_RESPONSE_RECEIVED As Long = 41
Private Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED As Long = 42
Private Const INTERNET_STATUS_PREFETCH As Long = 43
Private Const INTERNET_STATUS_CLOSING_CONNECTION As Long = 50
Private Const INTERNET_STATUS_CONNECTION_CLOSED As Long = 51
Private Const INTERNET_STATUS_HANDLE_CREATED As Long = 60
Private Const INTERNET_STATUS_HANDLE_CLOSING As Long = 70
Private Const INTERNET_STATUS_DETECTING_PROXY As Long = 80
Private Const INTERNET_STATUS_REQUEST_COMPLETE As Long = 100
Private Const INTERNET_STATUS_REDIRECT As Long = 110
Private Const INTERNET_STATUS_INTERMEDIATE_RESPONSE As Long = 120
Private Const INTERNET_STATUS_USER_INPUT_REQUIRED As Long = 140
Private Const INTERNET_STATUS_STATE_CHANGE As Long = 200

'connected state (mutually exclusive with disconnected)
Private Const INTERNET_STATE_CONNECTED As Long = &H1
'disconnected from network
Private Const INTERNET_STATE_DISCONNECTED As Long = &H2
'disconnected by user request
Private Const INTERNET_STATE_DISCONNECTED_BY_USER As Long = &H10
'no network requests being made (by Wininet)
Private Const INTERNET_STATE_IDLE As Long = &H100
'network requests being made (by Wininet)
Private Const INTERNET_STATE_BUSY As Long = &H200


Private Type ptINTERNET_ASYNC_RESULT
   dwResult As Long
   dwError As Long
End Type

Private Enum FTP_STATES
   FTP_WAIT
   FTP_ENUM
   FTP_DOWNLOAD
   FTP_DOWNLOADING
   FTP_UPLOAD
   FTP_UPLOADING
   FTP_CREATINGDIR
   FTP_CREATEDIR
   FTP_REMOVINGDIR
   FTP_REMOVEDIR
   FTP_DELETINGFILE
   FTP_DELETEFILE
   FTP_RENAMINGFILE
   FTP_RENAMEFILE
   FTP_ENUMFILES
End Enum

'Used in Callback function
Private AUTO_Bytes_Progress As Long, AUTO_Current_FileSize As Long
Private AUTO_State As FTP_STATES

Public bUsingManualMethod As Boolean, bCancelFTP As Boolean

Public Const ERROR_NO_MORE_FILES = 18
'clsFTP uses ^

'Private Const MAXDWORD As Double = (2 ^ 32) - 1

'###################################################################################################
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
'Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const INTERNET_FLAG_NO_CACHE_WRITE As Long = &H4000000

Private Const FTP_TRANSFER_TYPE_BINARY = 2
Private Const INTERNET_DEFAULT_FTP_PORT = 21 'INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000

Private Const scUserAgent = FTPControlStr '"MicRobSoft OpenUrl"
Private Const INTERNET_FLAG_RELOAD = &H80000000, INTERNET_FLAG_PRAGMA_NOCACHE = &H100

'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
    (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, _
    ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
    (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Long 'bool

'Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hinet As Long) As Integer
'end http

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
    (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Long

'directory listing
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
    (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
    lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, _
    ByVal dwContent As Long) As Long

Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

Private Type SYSTEMTIME
    wYear         As Integer
    wMonth        As Integer
    wDayOfWeek    As Integer
    wDay          As Integer
    wHour         As Integer
    wMinute       As Integer
    wSecond       As Integer
    wMilliseconds As Integer
End Type

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Function InternetSetStatusCallback(hInternet As Long, lpfnInternetCallback As Long) As Long
Dim bDoNorm As Boolean

If frmMain_Loaded Then
    If frmMain.mnuDevAdvCmdsNoFTPCallbacks.Checked Then
        InternetSetStatusCallback = INTERNET_INVALID_STATUS_CALLBACK
        bDoNorm = False
    Else
        bDoNorm = True
    End If
Else
    bDoNorm = True
End If

If bDoNorm Then
    InternetSetStatusCallback = apiInternetSetStatusCallback(hInternet, lpfnInternetCallback)
End If

End Function

Public Property Get bFTP_Doing() As Boolean
bFTP_Doing = pbFTP_Doing
End Property

Private Function GetFTPMethod(bGet As Boolean) As eFTP_Methods
Dim i As Integer

With frmMain
    If bGet Then
        For i = .mnuOnlineFTPDLO.LBound To .mnuOnlineFTPDLO.UBound
            If .mnuOnlineFTPDLO(i).Checked Then
                GetFTPMethod = i
                Exit For
            End If
        Next i
    Else
        For i = .mnuOnlineFTPULO.LBound To .mnuOnlineFTPULO.UBound
            If .mnuOnlineFTPULO(i).Checked Then
                GetFTPMethod = i
                Exit For
            End If
        Next i
    End If
End With

End Function

Public Property Get FTP_StealthMode() As Boolean
FTP_StealthMode = pFTP_bStealth
End Property
Public Property Let FTP_StealthMode(b As Boolean)
pFTP_bStealth = b
End Property

Public Property Get FTP_DeferToUpdateForm() As Boolean
FTP_DeferToUpdateForm = pFTP_bDeferToUpdateForm
End Property

Private Function InternetConnect(lNet As Long, ServerName As String, _
    UName As String, Pass As String, bCallBack As Boolean, Optional bForce As Boolean = False) As Long

If modVars.Closing And Not bForce Then
    InternetConnect = 0
Else
    InternetConnect = apiInternetConnect(lNet, ServerName, _
            INTERNET_DEFAULT_FTP_PORT, UName, Pass, _
            INTERNET_SERVICE_FTP, _
            IIf(frmMain.mnuOnlineFTPPassive.Checked, INTERNET_FLAG_PASSIVE, 0), Abs(bCallBack))
End If

End Function

'sites
'http://vbnet.mvps.org/index.html?code/internet/ftpdownload.htm
'http://www.15seconds.com/issue/981203.htm - ftp
'http://www.15seconds.com/issue/990408.htm - tcp comp


'ByVal ServerName As String, ByVal UName As String, ByVal Pass As String,
Private Sub DoFTP( _
    ByVal RemoteFile As String, ByVal LocalFile As String, ByVal bGet As Boolean, _
    ByRef sError As String, ByRef bSuccess As Boolean, ByVal bAllowCancel As Boolean)


'ftp details
Dim ServerName As String, UName As String, Pass As String, sRoot As String, sFileExt As String, sCurrentFileExt As String
GetFTPDetails ServerName, UName, Pass, sRoot, sFileExt

Dim vMethod As eFTP_Methods


If pbFTP_Doing Then
    sError = "Download Already in Progress"
    Exit Sub
Else
    pbFTP_Doing = True
End If
    

vMethod = GetFTPMethod(bGet)


sCurrentFileExt = modVars.GetFileExtension(RemoteFile)

If sCurrentFileExt <> sFileExt And sCurrentFileExt = "mcc" Then 'TRUST ME
    RemoteFile = Left$(RemoteFile, Len(RemoteFile) - Len(sCurrentFileExt) - 1) & _
        IIf(LenB(sFileExt), Dot & sFileExt, vbNullString)
    
End If




If vMethod = FTP_HTTP And (Not pFTP_bStealth) Then
    bSuccess = DoFTP_HTML(RemoteFile, LocalFile, sError)
    
Else
    If Left$(RemoteFile, 1) <> "/" Then RemoteFile = "/" & RemoteFile
    
    bCancelFTP = False
    
    If vMethod = FTP_Manual And (Not pFTP_bStealth) Then 'force Auto Mode for Stealth
        bUsingManualMethod = True
        bSuccess = DoFTP_Manual(ServerName, UName, Pass, RemoteFile, LocalFile, bGet, sError, bAllowCancel)
        
    Else
        bUsingManualMethod = False
        
        bSuccess = DoFTP_Auto(ServerName, UName, Pass, RemoteFile, LocalFile, bGet, sError, bAllowCancel)
        'if bGet=False, i.e. and upload, and ftp_method is HTTP, then upload ends up here
    End If
End If


pbFTP_Doing = False

End Sub

Private Function DoFTP_HTML(ByVal RemoteFile As String, ByVal LocalFile As String, _
    ByRef sError As String) As Boolean

Dim i As Integer
Dim rNew As String
Dim lR As Long


If modVars.bNoInternet Then
    If modLoadProgram.bSafeMode Then
        sError = "Turn off safe mode to access the internet via HTML method"
    Else
        sError = "Internet Not Connected"
    End If
Else
    
'    i = FreeFile()
'
'    rNew = GetHTML(RemoteFile)
'
'    Open LocalFile For Output As #i
'        Print #i, rNew
'    Close #i
'
'
'    DoFTP_HTML = True
    
    i = InStr(1, RemoteFile, "htdocs", vbTextCompare) + 7 'len(htdocs)+1

    rNew = HTTPRoot & Mid$(RemoteFile, i)

    'Ret = URLDownloadToFile(0, RemoteFile, LocalFile, 0, 0)
    lR = URLDownloadToFile(0&, rNew, LocalFile, BINDF_GETNEWESTVERSION, 0&)

    If lR = 0 And FileExists(LocalFile) Then
        DoFTP_HTML = True
    Else
        Select Case lR
            Case 0
                i = Err.LastDllError
                If i = 0 Then
                    sError = "Unknown Error"
                Else
                    sError = modVars.DllErrorDescription(i)
                End If

            Case E_OUTOFMEMORY
                sError = "Download Call is 'Out Of Memory'"
            Case INET_E_DOWNLOAD_FAILURE
                sError = "Download Error - Filename may not be valid"
        End Select
    End If
End If

End Function

Private Function DoFTP_Auto(ByVal ServerName As String, ByVal UName As String, ByVal Pass As String, _
    ByVal RemoteFile As String, ByVal LocalFile As String, ByVal bGet As Boolean, ByRef sError As String, _
    ByVal bAllowCancel As Boolean) As Boolean

Dim lNet As Long
Dim lCon As Long
Dim bSuccess As Boolean, bCallBack As Boolean
Dim Ret As Long, hCallback As Long

'If AddToConsole Then AddConsoleText "Beginning FTP Transfer...", , True, , True

lNet = InternetOpen(FTPControlStr, INTERNET_OPEN_TYPE_DIRECT, _
    vbNullString, vbNullString, INTERNET_FLAG_NO_CACHE_WRITE)


#If bDebug_Auto_FTP Then
    AddConsoleText "Internet Initialised - lNet: " & CStr(lNet)
#End If


If lNet Then
    
    hCallback = InternetSetStatusCallback(lNet, AddressOf FtpCallbackStatus)
    bCallBack = (hCallback <> INTERNET_INVALID_STATUS_CALLBACK)
    If bCallBack Then
        If Not pFTP_bStealth Then
            Load frmFTP
            
            frmFTP.cmdCancel.Enabled = bAllowCancel
        End If
    Else
        AddConsoleText "Error Setting Callback"
        
        If pFTP_bStealth Then
            sError = "Couldn't Set Callback: " & CStr(Err.LastDllError)
            DoFTP_Auto = False
            Exit Function
        End If
    End If
    
    
    #If bDebug_Auto_FTP Then
        AddConsoleText "Callback Set: " & CStr(bCallBack)
    #End If
    
    
    AUTO_State = IIf(bGet, FTP_DOWNLOADING, FTP_UPLOADING)
    
    If Not modVars.bNoInternet Then
        lCon = InternetConnect(lNet, ServerName, UName, Pass, bCallBack)
        
    End If
    '                                      0=default=21     1=ftp
    '                       flags; 0 = normal, &H8000000 = passive ftp
    '                       context (callbacks) therefore 0
    
    #If bDebug_Auto_FTP Then
        AddConsoleText "Connection - lCon: " & CStr(lCon)
    #End If
    
    
    If lCon Then
        
        If bGet Then
            AUTO_Current_FileSize = GetFTPFileSize(lCon, RemoteFile, bCallBack)
        Else
            AUTO_Current_FileSize = GetFileSize_Bytes(LocalFile)
        End If
        
        #If bDebug_Auto_FTP Then
            AddConsoleText "File Size: " & CStr(AUTO_Current_FileSize) & " bytes"
        #End If
        
        If AUTO_Current_FileSize > 0 Or frmMain.mnuDevAdvCmdsNoFTPCallbacks.Checked Then
            If bGet Then
                bSuccess = GetFile(lCon, RemoteFile, LocalFile, sError, bCallBack)
            Else
                bSuccess = PutFile(lCon, RemoteFile, LocalFile, sError, bCallBack)
            End If
            
            
            #If bDebug_Auto_FTP Then
                AddConsoleText "FTP" & IIf(bGet, "Get", "Put") & "File() returned " & CStr(bSuccess) & _
                    IIf(Not bSuccess, " (sError: " & sError & ")", vbNullString)
                
            #End If
        Else
            bSuccess = False
            If AUTO_Current_FileSize = -1 Then
                sError = "File not found (GetFileSize) - Server may not be working"
            Else
                Ret = Err.LastDllError
                sError = "GetFileSize() Error" & IIf(Ret <> 0, " (" & CStr(Ret) & ")", vbNullString)
            End If
        End If
        
        If bCallBack Then RemoveCallback lCon, 0
        Ret = InternetCloseHandle(lCon)
        lCon = 0
    Else
        bSuccess = False
        If modVars.bNoInternet Then
            'If modLoadProgram.bSafeMode Then
                'sError = "Turn off safe mode to download"
            'Else
                sError = "Internet access disabled"
            'End If
        Else
            Ret = Err.LastDllError
            
            If Ret > 0 Then
                sError = FTP_TranslateErrorCode(Ret)
            End If
            
            If LenB(sError) = 0 Then
                sError = "Had trouble connecting to the web server" & IIf(Ret, " (" & CStr(Ret) & ")", vbNullString)
            End If
        End If
    End If
    
    
    
    If bCallBack Then
        'the connection handle inherits the callback, so it must be removed from the connection handle, too
        RemoveCallback lNet, hCallback
        hCallback = 0
        
        If Not pFTP_bStealth Then Unload frmFTP
    End If
    
    AUTO_Bytes_Progress = 0
    AUTO_State = 0
    
    
    Ret = InternetCloseHandle(lNet)
    lNet = 0
Else
    bSuccess = False
    sError = "Could Not Initialise Internet Access"
End If

'If Success = False And AddToConsole Then
'    AddConsoleText "Error - " & Err.LastDllError, , True
'    AddConsoleText "InternetOpen: " & lNet
'    AddConsoleText "InternetConnect: " & lCon
'    AddConsoleText "ServerName: " & ServerName
'    AddConsoleText "RemoteFile: " & RemoteFile
'    AddConsoleText "LocalFile: " & LocalFile
'    AddConsoleText "Valid UserName? " & CBool(LenB(UName))
'    AddConsoleText "Valid Password? " & CBool(LenB(Pass))
'    modConsole.Indent False
'End If
'If AddToConsole Then AddConsoleText "Finished Transfer Procedure", , , True



If bCancelFTP Then
    'was canceled
    DoFTP_Auto = False
    sError = "Canceled"
Else
    DoFTP_Auto = bSuccess
End If

'If AddToConsole Then AddConsoleText "Transfer CleanUp Complete", , , True

End Function

Private Sub RemoveCallback(ByVal lHandle As Long, ByVal hPrevCallBack As Long)
InternetSetStatusCallback lHandle, hPrevCallBack
End Sub

Private Function GetFile(lCon As Long, RemoteFile As String, LocalFile As String, _
    ByRef sError As String, ByVal bCallBack As Boolean) As Boolean

Dim lError As Long

'fFailIfExists - 0=replace localfile, 1=don't replace(therefore fail)
If CBool(FTPGetFile(lCon, RemoteFile, LocalFile, 0, 0, FTP_TRANSFER_TYPE_BINARY, Abs(bCallBack))) Then
    GetFile = True
Else
    GetFile = False
    '##############################################################
    lError = Err.LastDllError
    sError = modVars.DllErrorDescription(lError)
    
    If LenB(sError) = 0 Then sError = "Dll Error: " & CStr(lError)
    '##############################################################
End If


End Function

Private Function PutFile(lCon As Long, RemoteFile As String, LocalFile As String, _
    ByRef sError As String, ByVal bCallBack As Boolean) As Boolean

Dim lError As Long

If CBool(FTPPutFile(lCon, LocalFile, RemoteFile, FTP_TRANSFER_TYPE_BINARY, Abs(bCallBack))) Then
    PutFile = True
Else
    PutFile = False
    
    '##############################################################
    lError = Err.LastDllError
    sError = modVars.DllErrorDescription(lError)
    
    If LenB(sError) = 0 Then sError = "Dll Error: " & CStr(lError)
    '##############################################################
End If

End Function

Private Function GetFTPFileSize(hConnection As Long, sFile As String, bCallBack As Boolean) As Long
Dim Files() As ptFTPFile
Dim sDir As String, sFileName As String, sError As String
Dim i As Integer
Dim lSize As Long

i = InStrRev(sFile, "/")
sDir = Left$(sFile, i)
sFileName = LCase$(Mid$(sFile, i + 1))

modFTP.FTP_StealthMode = True 'so the user doesn't see "Listing files..."

lSize = -1 'file not found

If pGetDirFiles(hConnection, sDir, Files, bCallBack, sError) Then
    
    If modVars.FileArrayDimensioned(Files) Then
        For i = 0 To UBound(Files)
            If LCase$(Trim$(Files(i).sName)) = sFileName Then
                lSize = Files(i).lFileSize
                Exit For
            End If
        Next i
    Else
        lSize = -2 'other error
    End If
    
'Else
'    lSize = -1
End If


GetFTPFileSize = lSize

modFTP.FTP_StealthMode = False

'###############################################################
'old method
'###############################################################
'Dim hFind As Long
'Dim nLastError As Long
'Dim pData As WIN32_FIND_DATA
'
''Dim sDir As String, sFileName As String
''Dim i As Integer
''
''i = InStrRev(sFile, "/")
''sDir = Left$(sFile, i)
''sFileName = Mid$(sFile, i + 1)
''
'''pSetFTPDirectory hConnection, sDir
'
'hFind = FtpFindFirstFile(hConnection, Replace$(sFile, vbSpace, "?"), pData, INTERNET_FLAG_RELOAD, Abs(bCallBack))
'nLastError = Err.LastDllError
'
''pSetFTPDirectory hConnection, "/"
'
'If hFind = 0 Then
'    If (nLastError = ERROR_NO_MORE_FILES) Then
'        GetFTPFileSize = -1  ' File not found
'    Else
'        GetFTPFileSize = -2  ' Other error
'    End If
'Else
'
'    GetFTPFileSize = pData.nFileSizeLow
'    'GetFileSize = (pData.nFileSizeHigh * (MAXDWORD + 1)) + pData.nFileSizeLow
'
'    RemoveCallback hFind, 0
'    InternetCloseHandle hFind
'End If

End Function

Private Function FtpCallbackStatus(ByVal hInternet As Long, _
                                  ByVal dwContext As Long, _
                                  ByVal dwInternetStatus As Long, _
                                  ByVal lpvStatusInfo As Long, _
                                  ByVal dwStatusInfoLength As Long) As Long

Dim sMsg As String
Dim cBuffer As String
Dim dwRead As Long
Dim uStatus As ptINTERNET_ASYNC_RESULT

If bCancelFTP Then
    RemoveCallback hInternet, 0
    InternetCloseHandle hInternet
    
    'causes this:
    '"Error:  12017 The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed."
    Exit Function
End If


If pFTP_bStealth Then
    DoEvents
    Exit Function
End If


cBuffer = Space$(dwStatusInfoLength)
Select Case dwInternetStatus
    
    Case INTERNET_STATUS_RESPONSE_RECEIVED
        
        
        If AUTO_State = FTP_DOWNLOADING Then
            If AUTO_Current_FileSize Then
                CopyMemory dwRead, ByVal lpvStatusInfo, dwStatusInfoLength
                AUTO_Bytes_Progress = AUTO_Bytes_Progress + dwRead
                
                If AUTO_Bytes_Progress < AUTO_Current_FileSize Then
                    frmFTP.cFTP_FileTransferProgress AUTO_Bytes_Progress, AUTO_Current_FileSize, False
                'Else
                    'frmFTP.cFTP_FileTransferProgress AUTO_Current_FileSize, AUTO_Current_FileSize, False
                End If
            End If
        End If
        
    
'    Case INTERNET_STATUS_SENDING_REQUEST
'
'        If AUTO_State = FTP_UPLOADING Then
'            CopyMemory dwRead, ByVal lpvStatusInfo, dwStatusInfoLength
'            AUTO_Bytes_Progress = AUTO_Bytes_Progress + dwRead
'
'            If AUTO_Bytes_Progress < AUTO_Current_FileSize Then
'                frmFTP.cFTP_FileTransferProgress AUTO_Bytes_Progress, AUTO_Current_FileSize, False
'            Else
'                frmFTP.cFTP_FileTransferProgress AUTO_Current_FileSize, AUTO_Current_FileSize, False
'            End If
'        End If
        
        
    Case INTERNET_STATUS_REQUEST_SENT

'        CopyMemory dwRead, ByVal lpvStatusInfo, dwStatusInfoLength
'        sMsg = "Request sent: " & dwRead & " bytes"
'        pub_BytesSent = pub_BytesSent + dwRead

        If AUTO_State = FTP_UPLOADING Then
            If AUTO_Current_FileSize Then
                CopyMemory dwRead, ByVal lpvStatusInfo, dwStatusInfoLength
                AUTO_Bytes_Progress = AUTO_Bytes_Progress + dwRead
                
                If AUTO_Bytes_Progress < AUTO_Current_FileSize Then
                    frmFTP.cFTP_FileTransferProgress AUTO_Bytes_Progress, AUTO_Current_FileSize, False
                'Else
                    'frmFTP.cFTP_FileTransferProgress AUTO_Current_FileSize, AUTO_Current_FileSize, False
                End If
            End If
        End If


'    Case INTERNET_STATUS_RECEIVING_RESPONSE
'    Case INTERNET_STATUS_CTL_RESPONSE_RECEIVED
'    Case INTERNET_STATUS_PREFETCH
        
    Case INTERNET_STATUS_RESOLVING_NAME
        
        CopyMemory ByVal cBuffer, ByVal lpvStatusInfo, dwStatusInfoLength
        sMsg = "Looking up the IP address for " & Trim0(cBuffer)
        
    Case INTERNET_STATUS_NAME_RESOLVED
        
        CopyMemory ByVal cBuffer, ByVal lpvStatusInfo, dwStatusInfoLength
        sMsg = "Name resolved " & Trim0(cBuffer)
        
    Case INTERNET_STATUS_CONNECTING_TO_SERVER
        
        CopyMemory ByVal cBuffer, ByVal lpvStatusInfo, dwStatusInfoLength
        sMsg = "Connecting to server.." & Trim0(cBuffer)
        
    Case INTERNET_STATUS_CONNECTED_TO_SERVER
        
        'sMsg = "Connected to server"
        CopyMemory ByVal cBuffer, ByVal lpvStatusInfo, dwStatusInfoLength
        sMsg = "Connected to " & Trim0(cBuffer)
        
        
        If AUTO_State = FTP_DOWNLOADING Or AUTO_State = FTP_UPLOADING Then
            'frmFTP.cFTP_FileTransferProgress 0, 1, False
            frmFTP.FloodBar.Flood_Show_Result True, "Connected to Server, Starting Transfer..."
            'can't know the file size yet
        ElseIf AUTO_State = FTP_ENUMFILES Then
            frmFTP.FloodBar.Flood_Show_Result True, "Connected to Server, Listing Files..."
            frmFTP.progFTP.Value = 100
            
        ElseIf AUTO_State = FTP_DELETINGFILE Then
            frmFTP.FloodBar.Flood_Show_Result True, "Connected to Server, Deleting File..."
            frmFTP.progFTP.Value = 100
            
        End If
        
        
        'download takes place here, if they were in order
        '(Above is connecting, connected
        'Below is closing, disconnected)
        
        
    Case INTERNET_STATUS_CLOSING_CONNECTION
        sMsg = "Closing connection"
        
    Case INTERNET_STATUS_CONNECTION_CLOSED
        sMsg = "Connection closed"
        
    'Case INTERNET_STATUS_HANDLE_CREATED
        'CopyMemory dwRead, ByVal lpvStatusInfo, dwStatusInfoLength
        'sMsg = "Handle created: " & CStr(dwRead)
        
    'Case INTERNET_STATUS_HANDLE_CLOSING
        'sMsg = "Handle closed"
        
        'If AUTO_State = FTP_DOWNLOADING Then
            'sMsg = "Download complete. " & sMsg
            'AUTO_State = FTP_WAIT
        'ElseIf AUTO_State = FTP_UPLOADING Then
            'sMsg = "Upload complete. " & sMsg
            'AUTO_State = FTP_WAIT
        'End If
        
        
    Case INTERNET_STATUS_DETECTING_PROXY
        sMsg = "Detecting proxy"
        
    Case INTERNET_STATUS_REQUEST_COMPLETE
        sMsg = "Request completed"
        
    Case INTERNET_STATUS_REDIRECT
        CopyMemory ByVal cBuffer, ByVal lpvStatusInfo, dwStatusInfoLength
        sMsg = "HTTP request redirected to " & Trim0(cBuffer)
        
    Case INTERNET_STATUS_INTERMEDIATE_RESPONSE
        sMsg = "Received intermediate status message from the server"
        
    Case INTERNET_STATUS_STATE_CHANGE
        'Moved between a secure and a non-secure site.
        CopyMemory dwRead, ByVal lpvStatusInfo, dwStatusInfoLength
        
        Select Case dwRead
            Case INTERNET_STATE_CONNECTED
                sMsg = "Connected state moved between secure and nonsecure site"
                
            Case INTERNET_STATE_DISCONNECTED
                sMsg = "Disconnected from network"
                
            Case INTERNET_STATE_DISCONNECTED_BY_USER
                sMsg = "Disconnected by user request"
                
            Case INTERNET_STATE_IDLE
                sMsg = "No network requests are being made (by WinlNet)"
                
            Case INTERNET_STATE_BUSY
                sMsg = "Network requests are being made (by WinlNet)"
                
            Case INTERNET_STATUS_USER_INPUT_REQUIRED
                sMsg = "The request requires user input to complete"
        
        End Select
End Select

If LenB(sMsg) Then frmFTP.SetLabelInfo sMsg

DoEvents

End Function

Private Function DoFTP_Manual(ByVal ServerName As String, ByVal UName As String, ByVal Pass As String, _
    ByVal RemoteFile As String, ByVal LocalFile As String, ByVal bGet As Boolean, _
    ByRef sError As String, ByVal bAllowCancel As Boolean) As Boolean

'stealth dealt with in _Load
Load frmFTP
frmFTP.cmdCancel.Enabled = bAllowCancel

DoFTP_Manual = frmFTP.FTP_Transfer(ServerName, UName, Pass, RemoteFile, LocalFile, bGet, sError)

Unload frmFTP

End Function

Public Function DownloadFTPFile(ByVal LocalF As String, ByVal RemoteF As String, _
    ByRef oCaller As Object, ByRef sError As String, ByVal bAllowCancel As Boolean) As Boolean


oCaller.Enabled = False
DoFTP RemoteF, LocalF, True, sError, DownloadFTPFile, bAllowCancel
oCaller.Enabled = True 'function returns ^

End Function

Public Function UploadFTPFile(ByVal LocalF As String, ByVal RemoteF As String, _
    ByRef oCaller As Object, ByRef sError As String, ByVal bAllowCancel As Boolean) As Boolean


oCaller.Enabled = False
DoFTP RemoteF, LocalF, False, sError, UploadFTPFile, bAllowCancel
oCaller.Enabled = True 'function returns ^

End Function

Public Sub GetFileStr(ByRef sFileContents As String, ByRef eType As eFTPCustErrs, ByVal RFile As String, _
    ByRef oCaller As Object, ByRef sError As String, ByVal bAllowCancel As Boolean)

Dim StartT As Long
Dim LFile As String
Dim f As Integer


LFile = Get_FTP_Temp_File()
sFileContents = vbNullString

If FileExists(LFile) Then
    On Error Resume Next
    Kill LFile
    On Error GoTo 0
End If


If DownloadFTPFile(LFile, RFile, oCaller, sError, bAllowCancel) Then
    'AddConsoleText "Waiting for System to Free File..."
    
    StartT = GetTickCount()
    Do
        If IsFileOpen(LFile) = False Then
            Exit Do
        ElseIf (StartT + FTP_Wait_Time) < GetTickCount() Then
            Exit Do
        End If
        
        Pause 10
    Loop
    
    'AddConsoleText "File Free"
    
    
    If FileExists(LFile) Then
        f = FreeFile()
        
CarryOn:
        On Error GoTo EH
        Open LFile For Binary As #f
            sFileContents = input(LOF(f), f)
        Close #f
        
        On Error Resume Next
        Kill LFile
        On Error GoTo 0
        
        eType = cSuccess
    Else
        eType = cFileNotFoundOnLocal
    End If
ElseIf LCase$(sError) = "file not found on server" Then
    eType = cFileNotFoundOnServer
Else
    eType = cOther
    'sError is passed back
End If


Exit Sub
EH:
If InStr(1, Err.Description, "permission", vbTextCompare) Then
    Err.Clear
    Pause 10
    Resume
Else
    eType = cOther
    If LenB(sError) = 0 Then sError = Err.Description
    Close #f
End If
End Sub

Public Function PutFileStr(ByVal sFileContents As String, ByVal RFile As String, _
    ByRef oCaller As Object, ByRef sError As String, _
    ByVal bAllowCancel As Boolean) As Boolean

Dim LFile As String
Dim f As Integer

LFile = Get_FTP_Temp_File()
f = FreeFile()

On Error GoTo EH
Open LFile For Output As #f
    Print #f, sFileContents;
Close #f

PutFileStr = UploadFTPFile(LFile, RFile, oCaller, sError, bAllowCancel)

On Error Resume Next
Kill LFile

Exit Function
EH:
PutFileStr = False
If LenB(sError) = 0 Then sError = Trim$("Error Creating Local File " & Err.Description)
End Function

Private Function Get_FTP_Temp_File() As String
Get_FTP_Temp_File = modVars.Comm_Safe_Path & "Communicator_Tmp." & modVars.FileExt
'modVars.RootDrive & "\
End Function

'###################################################################################################
'###################################################################################################
'###################################################################################################

Private Function GetIPRemoteFilePath() As String
GetIPRemoteFilePath = FTP_Root_Location & "/IP Data/IPs." & FileExt
End Function

Public Sub DownloadIPs(ByRef sIPs As String, ByRef bError As Boolean, _
    ByRef oCaller As Object, ByRef sError As String, ByVal bAllowCancel As Boolean)

Dim v_Error As eFTPCustErrs

GetFileStr sIPs, v_Error, GetIPRemoteFilePath(), oCaller, sError, bAllowCancel
Evaluate_FTP_Error v_Error, sError, bError

'Dim B As Boolean
'Dim StartT As Long
'
'B = DownloadFTPFile(FTP_IPLocal_File, FTP_IP_Remote_File, oCaller, sError)
'
'AddConsoleText "Waiting for System to Free File..."
''Pause FTP_Wait_Time
'StartT = GetTickCount()
'Do
'    If IsFileOpen(FTP_IPLocal_File) = False Then
'        Exit Do
'    ElseIf (StartT + FTP_Wait_Time) < GetTickCount() Then
'        Exit Do
'    End If
'
'    Pause 10
'Loop
'
'AddConsoleText "File Free"
'
'Dim f As Integer
'Dim Str As String
'
'If FileExists(FTP_IPLocal_File) Then
'    f = FreeFile()
'
'    On Error Resume Next
'CarryOn:
'    On Error GoTo EH
'    Open FTP_IPLocal_File For Input As #f
'        Str = Input(LOF(f), f)
'    Close #f
'    On Error GoTo 0
'
'    On Error Resume Next
'    Kill FTP_IPLocal_File
'    On Error GoTo 0
'
'    sIPs = Str
'Else
'    B = False
'End If
'
'Error = Not B
'
'Exit Sub
'EH:
'If InStr(1, Err.Description, "permission", vbTextCompare) Then
'    Err.Clear
'    GoTo CarryOn
'End If
End Sub

Public Function UploadIPs(ByRef oCaller As Object, ByRef sError As String, ByVal sFileContents As String) As Boolean
UploadIPs = PutFileStr(sFileContents, GetIPRemoteFilePath(), oCaller, sError, True)
End Function

Private Sub Evaluate_FTP_Error(ByVal v_Error As eFTPCustErrs, sError As String, bError As Boolean)

If v_Error = cSuccess Then
    bError = False
    
ElseIf v_Error = cOther Then
    If LenB(sError) = 0 Then sError = "Unknown Error"
    bError = True
    
ElseIf v_Error = cFileNotFoundOnLocal Then
    If LenB(sError) = 0 Then sError = "Error In Download"
    bError = True
    
ElseIf v_Error = cFileNotFoundOnServer Then
    If LenB(sError) = 0 Then sError = "Error Contacting Server"
    bError = True
    
End If

End Sub

Public Sub fGetVersion(ByRef oCaller As Object, ByRef sError As String, _
    ByRef bError As Boolean, ByRef sVersion As String)

Dim sFileContents As String
Dim v_Error As eFTPCustErrs
Dim sRemoteFile As String



sRemoteFile = FTP_Root_Location & "/Version/Version." & FileExt

pFTP_bDeferToUpdateForm = True
GetFileStr sFileContents, v_Error, sRemoteFile, oCaller, sError, True
pFTP_bDeferToUpdateForm = False

Evaluate_FTP_Error v_Error, sError, bError


If bError Then
    sVersion = vbNullString
Else
    sVersion = sFileContents
End If

'Dim StartT As Long
'
'Call DownloadFTPFile( _
'        modVars.RootDrive & "\" & Local_UpdateTxt, _
'        FTP_UpdateTxt, oCaller, sError)
'
'AddConsoleText "Waiting for System to Free File..."
''Pause FTP_Wait_Time 'allow the system time to do the file
'StartT = GetTickCount()
'Do
'    If IsFileOpen(modVars.RootDrive & "\" & Local_UpdateTxt) = False Then
'        Exit Do
'    ElseIf (StartT + FTP_Wait_Time) < GetTickCount() Then
'        Exit Do
'    End If
'
'    Pause 10
'Loop
'AddConsoleText "File Free"
'
'
'Dim f As Integer
'Dim Str As String
'
'f = FreeFile()
'
'On Error Resume Next
'CarryOn:
'On Error GoTo EH
'Open (modVars.RootDrive & "\" & Local_UpdateTxt) For Input As #f
'    Str = Input(LOF(f), f)
'Close #f
'On Error GoTo 0
'
'fGetVersion = Str
'
'On Error Resume Next
'Kill (modVars.RootDrive & "\" & Local_UpdateTxt)
'On Error GoTo 0
'
'Exit Function
'EH:
'If InStr(1, Err.Description, "permission", vbTextCompare) Then
'    Err.Clear
'    GoTo CarryOn
'End If
End Sub

Public Property Get FTP_Temp_Dir() As String
FTP_Temp_Dir = modVars.Comm_Safe_Path
End Property
Public Property Get FTP_Comm_Exe_File() As String
FTP_Comm_Exe_File = FTP_Temp_Dir() & Communicator_File
End Property

Public Function DownloadLatest(ByRef oCaller As Object, ByRef sError As String) As Boolean
Dim LFile As String

'LFile = modVars.RootDrive & "\" & Communicator_File
LFile = FTP_Comm_Exe_File()


pFTP_bDeferToUpdateForm = True
If DownloadFTPFile(LFile, FTP_Root_Location & "/Version/" & Communicator_File, _
        oCaller, sError, True) Then
    
    DownloadLatest = FileExists(LFile)
Else
    DownloadLatest = False
End If
pFTP_bDeferToUpdateForm = False


'Dim StartT As Long
'
'DownloadLatest = False
'
'
'Call DownloadFTPFile(modVars.RootDrive & "\" & Communicator_File, _
'                    "/" & FTP_Root_Location & "/Version/" & Communicator_File, oCaller, sError)
'
'AddConsoleText "Waiting for System to Free File..."
''pause ftp_wait_time
'StartT = GetTickCount()
'Do
'    If IsFileOpen(modVars.RootDrive & "\" & Communicator_File) = False Then
'        Exit Do
'    ElseIf (StartT + FTP_Wait_Time) < GetTickCount() Then
'        Exit Do
'    End If
'
'    Pause 10
'Loop
'AddConsoleText "File Free"
'
'If FileExists(RootDrive & "\" & Communicator_File) Then
'    DownloadLatest = True
'End If

End Function

Private Function HTMLCallbackStatus(ByVal hInternet As Long, _
                                  ByVal dwContext As Long, _
                                  ByVal dwInternetStatus As Long, _
                                  ByVal lpvStatusInfo As Long, _
                                  ByVal dwStatusInfoLength As Long) As Long

DoEvents

End Function

Public Function GetHTML(ByVal sUrl As String) As String

Dim sHTML As String
Dim hOpen As Long
Dim hOpenUrl As Long
Dim bDoLoop As Boolean
Dim bRet As Boolean
Dim sReadBuffer As String * 2048
Dim lNumberOfBytesRead As Long
Dim lRet As Long

Dim hCallback As Long
Dim bCallbackSet As Boolean

'AddConsoleText "Beginning HTTP Transfer...", , True

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

If hOpen Then
    
    hCallback = InternetSetStatusCallback(hOpen, AddressOf HTMLCallbackStatus)
    bCallbackSet = (hCallback <> INTERNET_INVALID_STATUS_CALLBACK)
    
    
    If Not modVars.bNoInternet Then
        hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, _
            INTERNET_FLAG_PRAGMA_NOCACHE Or INTERNET_FLAG_RELOAD, Abs(bCallbackSet))
    End If
    
    
    If hOpenUrl Then
        Do
            sReadBuffer = vbNullString
            
            bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
            
            If lNumberOfBytesRead Then
                sHTML = sHTML & Left$(sReadBuffer, lNumberOfBytesRead)
            Else
                Exit Do
            End If
            
        Loop While Not modVars.Closing
        
        If bCallbackSet Then RemoveCallback hOpenUrl, 0
        lRet = InternetCloseHandle(hOpenUrl)
    End If
    
    If bCallbackSet Then RemoveCallback hOpen, hCallback
    lRet = InternetCloseHandle(hOpen)
End If


'AddConsoleText "HTTP Transfer Procedure Complete", , , True

GetHTML = sHTML

End Function

Public Function ListDir(ByVal DirToList As String, ByRef Files() As ptFTPFile, ByRef sError As String) As Boolean

Dim lNet As Long
Dim lCon As Long
Dim lCallback As Long, bCallbackSet As Boolean


If pbFTP_Doing Then
    sError = "FTP already in progress, try again in a few seconds"
    ListDir = False
    Exit Function
Else
    pbFTP_Doing = True
End If


'ftp details
Dim ServerName As String, UName As String, Pass As String
GetFTPDetails ServerName, UName, Pass, vbNullString, vbNullString

Erase Files

'microbsoft.byethost13.com/htdocs/Files
'microbsoft.byethost13.com/htdocs/IP Data/Users

lNet = InternetOpen(FTPControlStr, INTERNET_OPEN_TYPE_DIRECT, _
    vbNullString, vbNullString, INTERNET_FLAG_NO_CACHE_WRITE)

If lNet Then
    lCallback = InternetSetStatusCallback(lNet, AddressOf FtpCallbackStatus)
    bCallbackSet = (lCallback <> INTERNET_INVALID_STATUS_CALLBACK)
    bCancelFTP = False
    If bCallbackSet Then
        AUTO_State = FTP_ENUMFILES
        modFTP.bUsingManualMethod = False
        Load frmFTP
    End If
    
    
    If Not modVars.bNoInternet Then
        lCon = InternetConnect(lNet, ServerName, UName, Pass, bCallbackSet)
    End If
    
    
    If lCon Then
        
        
        ListDir = pGetDirFiles(lCon, DirToList, Files(), bCallbackSet, sError)
        
        
        If bCallbackSet Then RemoveCallback lCon, 0
        InternetCloseHandle lCon
    Else
        ListDir = False
        If bCancelFTP Then
            sError = "Canceled"
        Else
            lCon = Err.LastDllError
            If lCon Then
                sError = FTP_TranslateErrorCode(lCon)
                lCon = 0
            Else
                sError = "Couldn't Connect to Server"
            End If
        End If
    End If
    
    
    If bCallbackSet Then
        RemoveCallback lNet, lCallback
        Unload frmFTP
    End If
    InternetCloseHandle lNet
Else
    ListDir = False
    sError = "Couldn't Initialise Internet Access"
End If

AUTO_State = 0
pbFTP_Doing = False


End Function

Private Function pGetDirFiles(ByVal lCon As Long, DirToList As String, ByRef tFiles() As ptFTPFile, _
    bCallbackSet As Boolean, ByRef sError As String) As Boolean

Dim pData As WIN32_FIND_DATA
Dim lRet As Long, lFind As Long, lError As Long
Dim iCount As Integer
Dim i As Integer


'init the filename buffer
pData.cFileName = String$(MAX_PATH, 0)


pSetFTPDirectory lCon, DirToList

'get the first file in the directory...
lFind = FtpFindFirstFile(lCon, "*", pData, INTERNET_FLAG_RELOAD, Abs(bCallbackSet))

'how'd we do?
If lFind = 0 Then
    
    'get the error from the findfirst call
    lError = Err.LastDllError
    
    'is the directory empty?
    If lError = ERROR_NO_MORE_FILES Then
        pGetDirFiles = True
    ElseIf lError Then
        pGetDirFiles = False
        
        If bCancelFTP Then
            sError = "Canceled"
        Else
            sError = FTP_TranslateErrorCode(lError)
        End If
    Else
        pGetDirFiles = False
        sError = UNKNOWN_ERROR_STRING
    End If
    
Else
    
'    'we got some dir info...
'    'get the name
'    sTemp = Trim0(pData.cFileName) 'Left$(pData.cFileName, InStr(1, pData.cFileName, vbNullChar, vbBinaryCompare) - 1)
'
'    If Extras Then
'        sList = sList & vbNewLine & sTemp & Sep & _
'            pData.nFileSizeLow & " bytes" & Sep & _
'            CStr(FileTimeToDate(pData.ftLastWriteTime))
'    Else
'        sList = sList & vbNewLine & sTemp
'    End If
    
    iCount = 0
    
    
    Do
        
        If Not (Left$(pData.cFileName, 1) = "." Or Left$(pData.cFileName, 2) = "..") Then
            'don't add "." or ".."
            
            ReDim Preserve tFiles(iCount)
            With tFiles(iCount)
                .sName = Trim0(pData.cFileName)
                .lFileSize = pData.nFileSizeLow
                .dDateCreated = FileTimeToDate(pData.ftCreationTime)
                .dDateLastAccessed = FileTimeToDate(pData.ftLastAccessTime)
                .dDateLastWritten = FileTimeToDate(pData.ftLastWriteTime)
            End With
            iCount = iCount + 1
        End If
        
        
        
        'init the filename buffer
        pData.cFileName = String$(MAX_PATH, 0)
        
        
        'how'd we do?
        If InternetFindNextFile(lFind, pData) = 0 Then
            
            'get the error from the findnext call
            lError = Err.LastDllError
            
            'no more items
            If lError = ERROR_NO_MORE_FILES Then
                pGetDirFiles = True
                'no more items...
                Exit Do
            ElseIf lError Then
                pGetDirFiles = False
                If bCancelFTP Then
                    sError = "Canceled"
                Else
                    sError = FTP_TranslateErrorCode(lError)
                End If
            Else
                pGetDirFiles = False
                sError = UNKNOWN_ERROR_STRING
            End If
            
        End If
    Loop
    
    'close the handle for the dir listing
    RemoveCallback lFind, 0
    InternetCloseHandle lFind
End If


pSetFTPDirectory lCon, "/"

End Function

Private Function pSetFTPDirectory(lCon As Long, sDir As String) As Boolean

pSetFTPDirectory = CBool(FtpSetCurrentDirectory(lCon, sDir))

End Function

Private Function FileTimeToDate(File_Time As FILETIME) As Date
Dim System_Time As SYSTEMTIME

' Convert the FILETIME into a SYSTEMTIME.
Call FileTimeToSystemTime(File_Time, System_Time)

' Convert the SYSTEMTIME into a Date.
FileTimeToDate = SystemTimeToDate(System_Time)

End Function

Private Function SystemTimeToDate(System_Time As SYSTEMTIME) As Date
' Convert a SYSTEMTIME into a Date.

On Error Resume Next
With System_Time
    SystemTimeToDate = DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond)
    
    'SystemTimeToDate = CDate( _
        Format$(.wDay) & "/" & _
        Format$(.wMonth) & "/" & _
        Format$(.wYear) & vbSpace & _
        Format$(.wHour) & ":" & _
        Format$(.wMinute, "00") & ":" & _
        Format$(.wSecond, "00"))
    
End With
On Error GoTo 0

End Function

Public Function DelFTPFile(ByVal RFile As String, ByRef sError As String) As Boolean

Dim Success As Boolean, bCallbackSet As Boolean
Dim lNet As Long, lCon As Long, Ret As Long, hCallback As Long

If pbFTP_Doing Then
    sError = "FTP already in progress, try again in a few seconds"
    DelFTPFile = False
    Exit Function
Else
    pbFTP_Doing = True
End If

'ftp details
Dim ServerName As String, UName As String, Pass As String
GetFTPDetails ServerName, UName, Pass, vbNullString, vbNullString


lNet = InternetOpen(FTPControlStr, INTERNET_OPEN_TYPE_DIRECT, _
    vbNullString, vbNullString, INTERNET_FLAG_NO_CACHE_WRITE)

AUTO_State = FTP_DELETINGFILE

If lNet Then
    '########################################################################
    hCallback = InternetSetStatusCallback(lNet, AddressOf FtpCallbackStatus)
    bCallbackSet = (hCallback <> INTERNET_INVALID_STATUS_CALLBACK)
    
    If bCallbackSet Then
        If modFTP.FTP_StealthMode = False Then
            modFTP.bUsingManualMethod = False
            Load frmFTP
            frmFTP.cmdCancel.Enabled = False
        End If
    Else
        AddConsoleText "Error Setting Callback"
    End If
    '########################################################################
    
    
    If Not modVars.bNoInternet Then
        lCon = InternetConnect(lNet, ServerName, UName, Pass, bCallbackSet, True)
    End If
    
    If lCon Then
        Success = CBool(FtpDeleteFile(lCon, RFile)) 'returns a c_bool
        If bCallbackSet Then RemoveCallback lCon, 0
        InternetCloseHandle lCon
    Else
        Success = False
        Ret = Err.LastDllError
        If Ret <> 0 Then
            sError = FTP_TranslateErrorCode(Ret)
        Else
            sError = UNKNOWN_ERROR_STRING
        End If
    End If
    
    If bCallbackSet Then
        RemoveCallback lNet, hCallback
        If modFTP.FTP_StealthMode = False Then Unload frmFTP
    End If
    InternetCloseHandle lNet
Else
    Success = False
    sError = "Couldn't Initialise Internet Access"
End If

AUTO_State = 0

DelFTPFile = Success
pbFTP_Doing = False

End Function

Private Function FTP_TranslateErrorCode(ByVal lErrorCode As Long) As String
Select Case lErrorCode
    Case 0
    Case ERROR_INTERNET_EXTENDED_ERROR: FTP_TranslateErrorCode = FTP_LastResponse() '"An extended error was returned from the server"
    Case 12001: FTP_TranslateErrorCode = "No more handles could be generated at this time"
    Case 12002: FTP_TranslateErrorCode = "The request has timed out"
    Case 12004: FTP_TranslateErrorCode = "An internal error has occurred"
    Case 12005: FTP_TranslateErrorCode = "The URL is invalid"
    Case 12006: FTP_TranslateErrorCode = "The URL scheme could not be recognized, or is not supported"
    Case 12007: FTP_TranslateErrorCode = "The server name could not be resolved"
    Case 12008: FTP_TranslateErrorCode = "The requested protocol could not be located"
    Case 12009: FTP_TranslateErrorCode = "A request to InternetQueryOption or InternetSetOption specified an invalid option value"
    Case 12010: FTP_TranslateErrorCode = "The length of an option supplied to InternetQueryOption or InternetSetOption is incorrect for the type of option specified"
    Case 12011: FTP_TranslateErrorCode = "The request option can not be set, only queried"
    Case 12012: FTP_TranslateErrorCode = "The Win32 Internet support is being shutdown or unloaded"
    
    Case 12013, 12014
        'FTP_TranslateErrorCode = _
        "The request to connect and login to an FTP server could not be completed because the supplied user name is incorrect"
        FTP_TranslateErrorCode = "FTP Server Error - Retry"
        
    'Case 12014
        'FTP_TranslateErrorCode = _
        "The request to connect and login to an FTP server could not be completed because the supplied password is incorrect"
        
        
        
    Case 12015: FTP_TranslateErrorCode = "The request to connect to and login to an FTP server failed"
    Case 12016: FTP_TranslateErrorCode = "The requested operation is invalid"
    Case 12017: FTP_TranslateErrorCode = "The operation was canceled, usually because the handle on which the request was operating was closed before the operation completed"
    Case 12018: FTP_TranslateErrorCode = "The type of handle supplied is incorrect for this operation"
    Case 12019: FTP_TranslateErrorCode = "The requested operation can not be carried out because the handle supplied is not in the correct state"
    Case 12020: FTP_TranslateErrorCode = "The request can not be made via a proxy"
    Case 12021: FTP_TranslateErrorCode = "A required registry value could not be located"
    Case 12022: FTP_TranslateErrorCode = "A required registry value was located but is an incorrect type or has an invalid value"
    Case 12023: FTP_TranslateErrorCode = "Direct network access cannot be made at this time"
    Case 12024: FTP_TranslateErrorCode = "An asynchronous request could not be made because a zero context value was supplied"
    Case 12025: FTP_TranslateErrorCode = "An asynchronous request could not be made because a callback function has not been set"
    Case 12026: FTP_TranslateErrorCode = "The required operation could not be completed because one or more requests are pending"
    Case 12027: FTP_TranslateErrorCode = "The format of the request is invalid"
    Case 12028: FTP_TranslateErrorCode = "The requested item could not be located"
    Case 12029: FTP_TranslateErrorCode = "The attempt to connect to the server failed, don't know why :("
    Case 12030: FTP_TranslateErrorCode = "The connection with the server has been terminated"
    Case 12031: FTP_TranslateErrorCode = "The connection with the server has been reset"
    Case 12036: FTP_TranslateErrorCode = "The request failed because the handle already exists"
    Case 12111: FTP_TranslateErrorCode = "The FTP operation was not completed because the session was aborted"
    Case 12163: FTP_TranslateErrorCode = "Could not connect - The internet is not connected to this PC at this time" 'ERROR_INTERNET_DISCONNECTED
    Case Else
        If lErrorCode = 0 Then
            FTP_TranslateErrorCode = UNKNOWN_ERROR_STRING
        Else
            FTP_TranslateErrorCode = modVars.DllErrorDescription(lErrorCode) '"Could not connect to the web server" ' "Error details not available"
        End If
End Select
End Function

Private Function FTP_LastResponse() As String

Dim lErr As Long, sErr As String, lenBuf As Long
'get the required buffer size
InternetGetLastResponseInfo lErr, sErr, lenBuf

'create a buffer
sErr = String$(lenBuf, 0)

'retrieve the last response info
InternetGetLastResponseInfo lErr, sErr, lenBuf


FTP_LastResponse = sErr
End Function


'##############################################################################
'##############################################################################
'##############################################################################

'Public Sub ApplyFTPRoot(ByVal HostName As String)

'modFTP.FTP_Root_Location = modFTP.Default_FTP_Root_Location

'modFTP.IPPath = FTP_Root_Location & "/IP Data"

'modFTP.FTP_IP_Remote_File = IPPath & "/IPs." & FileExt
'modFTP.FTP_UpdateTxt = "/" & FTP_Root_Location & "/Version/Version." & FileExt

'modFTP.FTP_Host_Name = HostName

'End Sub

Public Property Get FTP_Root_Location() As String
FTP_Root_Location = FTP_Details(iCurrent_FTP_Details).FTP_Root
End Property

Public Property Get FTP_Online_Users_Path() As String
FTP_Online_Users_Path = modFTP.FTP_Root_Location & "/IP Data/Online/"
End Property

Public Sub FTP_Init()
'File Structure:

'/Files
'/IP Data
'    "   /Logs
'    "   /Users
'/Messages
'/Version


FTP_iCustomServer = -1

'##########################################################################################

Add_FTP_Server_Details "microbsoft.110mb.com", "microbsoft", "Communicatore.8", _
    vbNullString, vbNullString, 0 '<-- first one must have an index        ^15-Char Limit

'##########################################################################################
'                                                               V 8-char limit
Add_FTP_Server_Details "ftp.byethost22.com", "b22_2822719", "Comm8721", _
    "/htdocs", "mcc"
'    ^ /home/vol5/byethost22.com/b22_2822719/htdocs


'##########################################################################################

Add_FTP_Server_Details "ftp.byethost13.com", "b13_1256618", "communicator", _
    "/microbsoft.byethost13.com/htdocs", "mcc"
'^down


'THIS CORRESPONDES TO THE BIT IN frmMain.InitVars - menu setup

End Sub
Public Sub Add_FTP_Server_Details(ByVal HostName As String, ByVal UserName As String, _
    ByVal Password As String, ByVal FTPRoot As String, ByVal sFileExt As String, _
    Optional ByVal iIndex As Integer = -1)

If iIndex = -1 Then iIndex = UBound(FTP_Details) + 1

ReDim Preserve FTP_Details(iIndex)

With FTP_Details(iIndex)
    .FTP_Host_Name = HostName
    .FTP_User_Name = UserName
    .FTP_Password = Password
    .FTP_Root = FTPRoot
    .FTP_File_Ext = sFileExt
    '.HTTP_Root = sHTTPRoot
End With

End Sub

Public Sub GetFTPDetails(ByRef ServerName As String, ByRef UserName As String, _
    ByRef Password As String, ByRef FTPRoot As String, ByRef sFileExt As String)

With FTP_Details(iCurrent_FTP_Details)
    ServerName = .FTP_Host_Name
    UserName = .FTP_User_Name
    Password = .FTP_Password
    FTPRoot = .FTP_Root
    sFileExt = .FTP_File_Ext
End With

End Sub
