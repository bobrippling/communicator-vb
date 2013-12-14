Attribute VB_Name = "modCmd"
Option Explicit

'other decs
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'end other decs

' Contants
Private Const STARTF_USESHOWWINDOW     As Long = &H1
Private Const STARTF_USESTDHANDLES     As Long = &H100
Private Const SW_HIDE                  As Integer = 0

' Types
Private Type SECURITY_ATTRIBUTES
    nLength                                As Long
    lpSecurityDescriptor                   As Long
    bInheritHandle                         As Long
End Type
Private Type STARTUPINFO
    cb                                     As Long
    lpReserved                             As String
    lpDesktop                              As String
    lpTitle                                As String
    dwX                                    As Long
    dwY                                    As Long
    dwXSize                                As Long
    dwYSize                                As Long
    dwXCountChars                          As Long
    dwYCountChars                          As Long
    dwFillAttribute                        As Long
    dwFlags                                As Long
    wShowWindow                            As Integer
    cbReserved2                            As Integer
    lpReserved2                            As Long
    hStdInput                              As Long
    hStdOutput                             As Long
    hStdError                              As Long
End Type
Private Type PROCESS_INFORMATION
    hProcess                               As Long
    hThread                                As Long
    dwProcessID                            As Long
    dwThreadId                             As Long
End Type

' Declares
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, _
                                                    phWritePipe As Long, _
                                                    lpPipeAttributes As Any, _
                                                    ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, _
                                                  lpBuffer As Any, _
                                                  ByVal nNumberOfBytesToRead As Long, _
                                                  lpNumberOfBytesRead As Long, _
                                                  lpOverlapped As Any) As Long
                                                  
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                                                              ByVal lpCommandLine As String, _
                                                                              lpProcessAttributes As Any, _
                                                                              lpThreadAttributes As Any, _
                                                                              ByVal bInheritHandles As Long, _
                                                                              ByVal dwCreationFlags As Long, _
                                                                              lpEnvironment As Any, _
                                                                              ByVal lpCurrentDriectory As String, _
                                                                              lpStartupInfo As STARTUPINFO, _
                                                                              lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private pCurrentDir As String

Public Property Get currentDir() As String
currentDir = pCurrentDir
End Property


Private Sub cd(sCommandLine As String, output As String)
Dim oldCD As String, quote As String, path As String
quote = Chr$(34)

On Error GoTo eh:
oldCD = currentDir

path = Trim$(Mid$(sCommandLine, 3))
If LenB(path) = 0 Then path = pCurrentDir


If Left$(path, 1) = quote And Right$(path, 1) = quote Then
    path = Mid$(path, 2, Len(path) - 2)
End If


If path = ".." Then
    path = Left$(path, InStr(1, path, "\"))
End If


If FileExists(path, vbDirectory) = False Then
    
    If InStr(pCurrentDir, ":") And InStr(oldCD, ":") Then
        'both ref. drives, can't combine
        output = "Invalid Path, reverted to '" & oldCD & "'"
        pCurrentDir = oldCD
    Else
        path = GetLocalFileName(oldCD, pCurrentDir)
        
        
        If FileExists(path, vbDirectory) Then
            pCurrentDir = path
            output = "Changed Dir to '" & currentDir & "'"
        Else
            output = "Directory doesn't exist, reverted to '" & oldCD & "'"
            pCurrentDir = oldCD
        End If
        Exit Sub
    End If
Else
    pCurrentDir = path
    output = "Changed Dir to '" & currentDir & "'"
End If


Exit Sub
eh:
output = Err.Description & IIf(LenB(output), " (" & output & ")", vbNullString)
End Sub

'---------------------------------------------------
' Call this sub to execute and capture a console app.
Public Sub ExecAndCapture(ByVal sCommandLine As String, _
                          ByRef output As String, Optional ByVal sStartInFolder As String = vbNullString)

Dim i As Integer
Dim BothCds As String
Dim BothDirs As String
Dim path As String
Dim Tmp As String


If pCurrentDir = vbNullString Or FileExists(currentDir, vbDirectory) = False Then pCurrentDir = AppPath()
If sStartInFolder = vbNullString Then sStartInFolder = pCurrentDir


sCommandLine = Trim$(sCommandLine)


If Left$(sCommandLine, 2) = "cd" Then
    
    cd Mid$(sCommandLine, 3), output
    
    Exit Sub
    
ElseIf LCase$(Left$(sCommandLine, 3)) = "dir" Then 'Left$(sCommandLine, Len("dir")) = "dir" And Len(sCommandLine) = Len("dir") Then
    
    On Error Resume Next
    path = Mid$(sCommandLine, 4)
    
    If path = vbNullString Then
        path = pCurrentDir
        
    ElseIf LCase$(path) = "drives" Then
        output = "Drive Listing: " & vbNewLine & Replace$(DriveList, ",", " , ")
        Exit Sub
    End If
    
    output = DirList(path)
    
    Exit Sub
    'sCommandLine = sCommandLine & " /a:d /o:n " & CurrentDir
    
    
ElseIf LCase$(Left$(sCommandLine, 3)) = "del" Then
    
    path = Right$(sCommandLine, Len(sCommandLine) - 4)
    
    If path = vbNullString Then
        output = "Specify something to delete"
    ElseIf FileExists(path, vbNormal) = False Then
        
        path = GetLocalFileName(pCurrentDir, path)
        
        If FileExists(path) = False Then
            output = "File Not Found"
            Exit Sub
        'else continue
        End If
    End If
    
    On Error Resume Next
    Kill path
    
    'If Err.Number = 53 Then
        'Output = "File Not Found"
    If Err.Number Then
        output = Err.Description
    Else 'no error
        output = path & " was deleted."
    End If
    
    Exit Sub
    
ElseIf LCase$(Left$(sCommandLine, 5)) = "start" Then
    
    output = "Use the shell command to start programs"
    Exit Sub
    
'    On Error Resume Next
'    Path = Mid$(sCommandLine, 7)
'    On Error GoTo 0
'
'    If (Path = vbNullString) Or (Path = Dot) Then
'        Path = CurrentDir
'    ElseIf FileExists(Path) = False Then
'
'        Path = GetLocalFileName(CurrentDir, Path)
'
'        If FileExists(Path) = False Then
'            Output = "File Not Found"
'            Exit Sub
'        'else continue
'        End If
'    End If
'
'    On Error Resume Next
'    Shell "explorer.exe " & Path, vbNormalNoFocus
'    On Error GoTo 0
'
'    If Err.Number Then
'        Output = "Error: " & Err.Description
'    Else 'no error
'        Output = "..." & Right$(Path, 25) & " was shelled/started."
'    End If
'    Exit Sub
End If

'If CurrentDir <> vbNullString Then
    'sStartInFolder = CurrentDir
'End If

Const BUFSIZE         As Long = 1024 * 10
Dim hPipeRead         As Long
Dim hPipeWrite        As Long
Dim sa                As SECURITY_ATTRIBUTES
Dim si                As STARTUPINFO
Dim Pi                As PROCESS_INFORMATION
Dim baOutput(BUFSIZE) As Byte
Dim sOutput           As String
Dim lBytesRead        As Long


    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1    ' get inheritable pipe handles
    End With 'SA
    
    If CreatePipe(hPipeRead, hPipeWrite, sa, 0) = 0 Then
        Error = True
        output = "Error - Couldn't Create Pipe (" & Err.LastDllError & ")"
        Exit Sub
        ' Set an errorlevel? for it can tell if invalid command?
    End If

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE        ' hide the window (0=hide, 1=show)
        .hStdOutput = hPipeWrite
        .hStdError = hPipeWrite
    End With 'SI
    
    'Debug.Print CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, pi)
    
    If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, sStartInFolder, si, Pi) Then
        Call CloseHandle(hPipeWrite)
        Call CloseHandle(Pi.hThread)
        hPipeWrite = 0
        Do
            DoEvents
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
                Exit Do
            End If
            sOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
            
            'added by me
            'cTextBox.SelStart = Len(cTextBox.Text)
            'end added by me
            'cTextBox.SelText = sOutput
            ' Look at above for adding to bottom of textbox
            output = output & sOutput
            
        Loop
        Call CloseHandle(Pi.hProcess)
    Else
        output = "Error Creating Process (" & Err.LastDllError & ")"
        Error = True
    End If
    ' To make sure...
    Call CloseHandle(hPipeRead)
    Call CloseHandle(hPipeWrite)


Exit Sub
ErrH:
'If Err.Number = 52 Then
    'output = "Drive Error (" & Err.Description & ")"
'Else
    output = "Error: " & Err.Description
'End If
Error = True
End Sub

Public Function DirList(ByVal path As String) As String

Dim output As String

Dim dr As String 'dir
Dim done As Boolean
Dim Attrib As Integer
Dim Tmp As String
Dim AttribDone As Boolean

If Right$(path, 1) <> "\" Then path = path & "\"
If Right$(path, 1) <> "*" Then path = path & "*"

dr = Dir$(path, vbArchive + vbDirectory + vbHidden + vbNormal + vbReadOnly + vbSystem + vbVolume)

path = Left$(path, Len(path) - 1) 'remove *

Do While Not done
    'On Error GoTo end_
    
    On Error Resume Next
    Attrib = GetAttr(path & dr)
    On Error GoTo end_
    
    AttribDone = False
    
    If dr <> "." And dr <> ".." And dr <> vbNullString Then
        Do While Not AttribDone
            
            If (Attrib And vbDirectory) = vbDirectory Then  'If InStr(dr, ".") Then 'file
                 Tmp = Tmp & ",Directory"
                 Attrib = Attrib And Not (vbDirectory)
                 
            ElseIf (Attrib And vbHidden) = vbHidden Then
                Tmp = Tmp & ",Hidden"
                Attrib = Attrib And Not (vbHidden)
                
            ElseIf (Attrib And vbArchive) = vbArchive Then
                Tmp = Tmp & ",Archive"
                Attrib = Attrib And Not (vbArchive)
                
            ElseIf (Attrib And vbReadOnly) = vbReadOnly Then
                Tmp = Tmp & ",Read Only"
                Attrib = Attrib And Not (vbReadOnly)
                
            ElseIf (Attrib And vbSystem) = vbSystem Then
                Tmp = Tmp & ",System"
                Attrib = Attrib And Not (vbSystem)
                
            ElseIf (Attrib And vbVolume) = vbVolume Then
                Tmp = Tmp & ",Volume"
                Attrib = Attrib And Not (vbVolume)
                
            Else
                If Tmp <> vbNullString Then
                    Tmp = Mid$(Tmp, 2) 'remove preceding ","
                End If
                AttribDone = True
            End If
            
        Loop
        
    ElseIf dr = vbNullString Then
        done = True
    End If
    
    If Tmp <> vbNullString Then
        output = output & vbNewLine & "'" & dr & "': " & Tmp
    End If
    
    dr = Dir$() 'vbArchive + vbDirectory + vbHidden + vbNormal + vbReadOnly + vbSystem + vbVolume)
    Tmp = vbNullString
    
Loop

end_:

If Left$(output, 2) = vbNewLine Then
    output = Mid$(output, 2)
End If
If Right$(output, 2) = "''" Then
    output = Left$(output, Len(output) - 2)
End If
If Right$(output, 2) = vbNewLine Then
    output = Left$(output, Len(output) - 2)
End If

output = "Listing of '" & IIf(Len(path) > 30, "..." & Right$(path, 27), path) & "'" & vbCrLf & vbCrLf & output

DirList = output

End Function

Public Function DriveList() As String

' Wrapper for calling the GetLogicalDriveStrings API
    
Dim Result As Long          ' Result of our api calls
Dim strDrives As String     ' String to pass to api call
Dim lenStrDrives As Long    ' Length of the above string

' Call GetLogicalDriveStrings with a buffer size of zero to
' find out how large our stringbuffer needs to be
Result = GetLogicalDriveStrings(0, strDrives)

strDrives = String$(Result, 0)
lenStrDrives = Result

' Call again with our new buffer
Result = GetLogicalDriveStrings(lenStrDrives, strDrives)

If Result = 0 Then
    ' There was some error calling the API
    ' Pass back an empty string
    ' NOTE: Implement proper error handling here
    DriveList = vbNullString
Else
    DriveList = Replace$(strDrives, vbNullChar, ",")
    If Right$(DriveList, 2) = ",," Then DriveList = Left$(DriveList, Len(DriveList) - 2)
End If

End Function

Public Function GetLocalFileName(ByVal path As String, ByVal File As String) As String

On Error Resume Next
If Right$(path, 1) <> "\" Then path = path & "\"
If Left$(File, 1) = "\" Then File = Mid$(File, 2)

GetLocalFileName = path & File

End Function
