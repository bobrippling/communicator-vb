Attribute VB_Name = "modDirBrowse"
Option Explicit

''Constantes
'Private Const BIF_RETURNONLYFSDIRS = &H1        'Uniquement des ripertoire
'Private Const BIF_DONTGOBELOWDOMAIN = &H2       'Domaine globale intredit
'Private Const BIF_STATUSTEXT = &H4              'Zone de saisie autorisie
'Private Const BIF_RETURNFSANCESTORS = &H8
'Private Const BIF_EDITBOX = &H10                'Zone de saisie autorisie
'Private Const BIF_VALIDATE = &H20               'insist on valid result (or CANCEL)
'Private Const BIF_BROWSEFORCOMPUTER = &H1000    'Uniquement des PCs.
'Private Const BIF_BROWSEFORPRINTER = &H2000     'Uniquement des imprimantes
'Private Const BIF_BROWSEINCLUDEFILES = &H4000   'Browsing for Everything
'
'Private Const MAX_PATH = 260
'
''Types
'Private Type T_BROWSEINFO
'   hWndOwner      As Long
'   pIDLRoot       As Long
'   pszDisplayName As Long
'   lpszTitle      As Long
'   ulFlags        As Long
'   lpfnCallback   As Long
'   lParam         As Long
'   iImage         As Long
'End Type
'
''Fonctions API Windows
'Private Declare Function SHBrowseForFolder Lib "shell32" _
'                                  (lpbi As T_BROWSEINFO) As Long
'
'Private Declare Function SHGetPathFromIDList Lib "shell32" _
'                                  (ByVal pidList As Long, _
'                                  ByVal lpBuffer As String) As Long
'
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
'                                  (ByVal lpString1 As String, ByVal _
'                                  lpString2 As String) As Long

Private Type BrowseInfo
    lngHwnd        As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'used in frmMain.GetSpecialFolder()

Private Declare Function lstrcat Lib "kernel32" _
   Alias "lstrcatA" (ByVal lpString1 As String, _
   ByVal lpString2 As String) As Long
   
Private Declare Function SHBrowseForFolder Lib "shell32" _
   (lpbi As BrowseInfo) As Long
   
Private Declare Function SHGetPathFromIDList Lib "shell32" _
   (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseFolder(ByVal OwnerHwnd As Long, ByVal sPrompt As String) As String

On Error GoTo ehBrowseForFolder 'Trap for errors

Dim intNull As Integer
Dim lngIDList As Long, lngResult As Long
Dim strPath As String
Dim udtBI As BrowseInfo

'Set API properties (housed in a UDT)
With udtBI
    .lngHwnd = OwnerHwnd
    .lpszTitle = lstrcat(sPrompt, vbNullString)
    .ulFlags = BIF_RETURNONLYFSDIRS
End With

'Display the browse folder...
lngIDList = SHBrowseForFolder(udtBI)

If lngIDList <> 0 Then
    'Create string of nulls so it will fill in with the path
    strPath = String(MAX_PATH, 0)
    
    'Retrieves the path selected, places in the null
     'character filled string
    lngResult = SHGetPathFromIDList(lngIDList, strPath)
    
    'Frees memory
    Call CoTaskMemFree(lngIDList)
    
    'Find the first instance of a null character,
     'so we can get just the path
    intNull = InStr(strPath, vbNullChar)
    
    If intNull Then
        'Set the value
        strPath = Left(strPath, intNull - 1)
    End If
End If

'Return the path name
BrowseFolder = strPath
Exit Function 'Abort

ehBrowseForFolder:

'Return no value
BrowseFolder = vbNullString

End Function




''*************************************************************
''*  BrowseFolder :
''*  Entries :   - HwndOwner     : Handle de la fenjtre appelante
''*              - Titre         : Titre
''*  Sorties :
''*              - string contenant le chemin complet ou Chaine vide
''*                (si annulation)
''*
''*  Affiche une boite de dialogue permettant la silection d'un ripertoire.
''*  Renvoie une chaine vide si l'opirateur annule.
''*************************************************************
'Public Function BrowseFolder(ByVal hWndOwner As Long, _
'    ByRef Title As String) As String
'
'Dim lpIDList As Long
'Dim sBuffer As String
'Dim BrowseInfo As T_BROWSEINFO
'
'
'BrowseFolder = vbNullString
'
'With BrowseInfo
'    .hWndOwner = hWndOwner
'    .lpszTitle = lstrcat(Title, vbNullString)
'    .ulFlags = BIF_RETURNONLYFSDIRS
'End With
'
'lpIDList = SHBrowseForFolder(BrowseInfo)
'
'If (lpIDList) Then
'    sBuffer = Space$(MAX_PATH)
'    SHGetPathFromIDList lpIDList, sBuffer
'    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
'    BrowseFolder = sBuffer
'End If
'
'End Function

