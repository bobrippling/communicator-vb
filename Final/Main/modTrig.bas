Attribute VB_Name = "modTrig"
Option Explicit

'*****************************************************
'
'This module contains a fast alternative to using the
'standard VB Cos and Sin functions. If you use many
'trig operations in your code this is the way to go.
'
'
'Lucky
'
'theluckyleper@home.com
'http://members.home.net/theluckyleper
'
'*****************************************************

Private SineArray(359) As Single           'Contains the values of Sin from 0-359 degrees
Private CoSineArray(359) As Single         'Contains the values of Cos from 0-359 degrees
Private TanArray(179) As Single         'Contains the values of Cos from 0-359 degrees
Public Const Rad_To_Deg As Single = 180 / Pi

'Public Function Sqrt(n As Single) As Single
'Dim i As Integer
'
'Dim Xn As Single
'
'Xn = 1
'
'For i = 0 To 4
'    Xn = (Xn + n / Xn) / 2
'Next i
'
'Sqrt = Xn
'
'End Function

Public Sub TrigInit()

Dim i As Integer

'This routine fills the values of the arrays so that they can
'Be used by the various functions in this module
For i = 0 To 359
    SineArray(i) = Sin(i / Rad_To_Deg)
    CoSineArray(i) = Cos(i / Rad_To_Deg)
Next i

For i = 0 To 179
    TanArray(i) = Tan(i / Rad_To_Deg)
Next i

End Sub

'Public Function TestTrig() As String
'Dim i As Integer
'Const Accuracy = 1
'
'
''#############################################################################
''test arrays
'For i = 0 To 359
'    If Round(SineArray(i), Accuracy) <> Round(Sin(i / Rad), Accuracy) Then
'        TestTrig = "Sine(" & i & ") is wrong"
'        Exit Function
'    ElseIf Round(CoSineArray(i), Accuracy) <> Round(Cos(i / Rad), Accuracy) Then
'        TestTrig = "Cosine(" & i & ") is wrong"
'        Exit Function
'    End If
'Next i
'
'For i = 0 To 179
'    If Round(TanArray(i), Accuracy) <> Round(Tan(i / Rad), Accuracy) Then
'        TestTrig = "Tan(" & i & ") is wrong"
'        Exit Function
'    End If
'Next i
''#############################################################################
''test functions
'For i = 0 To 359
'    If Round(Sine(i / Rad), Accuracy) <> Round(Sin(i / Rad), Accuracy) Then
'        TestTrig = "Function Sine(" & i & ") is wrong"
'        Exit Function
'    ElseIf Round(CoSine(i / Rad), Accuracy) <> Round(Cos(i / Rad), Accuracy) Then
'        TestTrig = "Function Cosine(" & i & ") is wrong"
'        Exit Function
'    End If
'Next i
'
'For i = 0 To 179
'    If Round(Tangent(i / Rad), Accuracy) <> Round(Tan(i / Rad), Accuracy) Then
'        TestTrig = "Function Tan(" & i & ") is wrong"
'        Exit Function
'    End If
'Next i
''#############################################################################
'
'
'TestTrig = "Yay"
'
'End Function

Public Function Sine(ByVal Angle As Single) As Single

Do While Angle < 0
    Angle = Angle + Pi2
Loop
Sine = SineArray(CLng(Angle * Rad_To_Deg) Mod 360)

''CDegrees + Round
'Angle = CLng(Angle * Rad)
'
''Ensure we're between 0 and 359
'Do While Angle < 0
'    Angle = Angle + 360
'Loop
'Do While Angle > 359
'    Angle = Angle - 360
'Loop
'
''Return the value
'Sine = SineArray(Angle)

End Function

Public Function CoSine(ByVal Angle As Single) As Single

Do While Angle < 0
    Angle = Angle + Pi2
Loop
CoSine = CoSineArray(CLng(Angle * Rad_To_Deg) Mod 360)


''If we're in radians, convert to degrees and round
'Angle = CLng(Angle * Rad)
'
''Ensure we're between 0 and 359
'Do While Angle < 0
'    Angle = Angle + 360
'Loop
'Do While Angle > 359
'    Angle = Angle - 360
'Loop
'
''Return the value
'CoSine = CoSineArray(Angle)

End Function

Public Function Tangent(ByVal Angle As Single) As Single

Do While Angle < 0
    Angle = Angle + Pi2
Loop
Tangent = TanArray(CLng(Angle * Rad_To_Deg) Mod 180)

End Function
