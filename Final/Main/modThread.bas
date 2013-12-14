Attribute VB_Name = "modThread"
Option Explicit

Private Declare Function CreateThread Lib "kernel32" (ByVal lpSecurityAttributes As Long, _
   ByVal dwStackSize As Long, _
   ByVal lpStartAddress As Long, _
   ByVal lpParameter As Long, _
   ByVal dwCreationFlags As Long, _
   lpThreadId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function AddThread() As Boolean

Dim H As Long, tID As Long

H = CreateThread(0, 0, AddressOf OpenLogin, 0, 0, tID)

AddConsoleText "H: " & H & " tID: " & tID

CloseHandle H

End Function

Private Sub OpenLogin()

Load frmLogin
frmLogin.Show

End Sub

'' Structure to hold IDispatch GUID
'Private Type GUID
'   Data1 As Long
'   Data2 As Integer
'   Data3 As Integer
'   Data4(7) As Byte
'End Type
'
'Private IID_IDispatch As GUID
'
'Private Declare Function CoMarshalInterThreadInterfaceInStream Lib "ole32.dll" _
'   (riid As GUID, ByVal pUnk As IUnknown, ppStm As Long) As Long
'
'Private Declare Function CoGetInterfaceAndReleaseStream Lib "ole32.dll" _
'   (ByVal pStm As Long, riid As GUID, pUnk As IUnknown) As Long
'
'Private Declare Function CoInitialize Lib "ole32.dll" (ByVal pvReserved As Long) As Long
'Private Declare Sub CoUninitialize Lib "ole32.dll" ()
'
'
'Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
'
'
'
'Public Function AddThread() 'ByVal Addr As Long) As Boolean
'Dim c As clsBackGround
'Dim Ret As Long
'
'If IsIDE() = False Then
'    Set c = New clsBackGround
'
'    Ret = StartBackgroundThreadApt(c)
'
'    'added by me
'    'Set c = Nothing
'End If
'
'AddThread = (Ret <> 0)
'
'End Function
'
''--------------------------------------------------------------------------------
'
'' Start the background thread for this object
'' using the apartment model
'' Returns zero on error                                             'added as long
'Private Function StartBackgroundThreadApt(ByVal qobj As clsBackGround) As Long
'Dim threadid As Long
'Dim hnd&, res&
'Dim threadparam As Long
'Dim tobj As Object
'
'Set tobj = qobj
'
'AddConsoleText "Entered Proc", , True
'
''Proper marshalled approach
'InitializeIID
'
'res = CoMarshalInterThreadInterfaceInStream(IID_IDispatch, qobj, threadparam)
'
'AddConsoleText "Res: " & res
'
'If res <> 0 Then
'   StartBackgroundThreadApt = 0
'Else
'
'    hnd = CreateThread(0, 2000, AddressOf BackgroundFuncApt, threadparam, 0, threadid)
'
'    AddConsoleText "hnd: " & hnd
'
'    If hnd = 0 Then
'       'Return with zero (error)
'       StartBackgroundThreadApt = 0
'    Else
'
'        ' We don't need the thread handle
'        CloseHandle hnd
'
'        StartBackgroundThreadApt = threadid
'
'        ' This message box can cause problems
'        ' MsgBox "New thread created " & threadid
'        AddConsoleText "tid: " & threadid
'
'    End If
'End If
'
'End Function
'
'' Initialize the GUID structure
'Private Sub InitializeIID()
'Static Initialized As Boolean
'
'If Initialized = False Then
'    With IID_IDispatch
'       .Data1 = &H20400
'       .Data2 = 0
'       .Data3 = 0
'       .Data4(0) = &HC0
'       .Data4(7) = &H46
'    End With
'    Initialized = True
'End If
'
'End Sub
'
'' An correctly marshalled apartment model callback.
'' This is the correct approch, though slower.
'Private Function BackgroundFuncApt(ByVal param As Long) As Long
'
'Dim qobj As Object
'Dim qobj2 As clsBackGround
'Dim res&
'
''AddConsoleText "Entered funcapt"
'
'' This new thread is a new apartment, we must
'' initialize OLE for this apartment (VB doesn't seem to do it)
'res = CoInitialize(0)
'
'' Proper apartment modeled approach
'res = CoGetInterfaceAndReleaseStream(param, IID_IDispatch, qobj)
'
'Set qobj2 = qobj
'
''AddConsoleText "calling showaform"
'qobj2.ShowAForm
'
'' Alternatively, you can put a wait function here,
'' then call the qobj function when the wait is
'' satisfied
'
'' All calls to CoInitialize must be balanced
'CoUninitialize
'
'End Function
