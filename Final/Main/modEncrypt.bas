Attribute VB_Name = "modEncrypt"
Option Explicit

Private Const CryptPass As String = "²¿ÊÐ?ßø??" '"ReetTasty"

Public Function CryptString(ByVal ptSource As String, _
    Optional ByVal ptPassword As String = CryptPass) As String

Dim tdest As String
Dim lteller As Long
Dim lPasswTeller As Long

tdest = Space$(Len(ptSource))

For lteller = 1 To Len(ptSource)
    lPasswTeller = lPasswTeller - 1
    If lPasswTeller < 1 Then lPasswTeller = Len(ptPassword)
    
    Mid$(tdest, lteller, 1) = Chr$(Asc(Mid$(ptSource, lteller, 1)) Xor Asc(Mid$(ptPassword, lPasswTeller, 1)))
Next lteller


CryptString = tdest

End Function
