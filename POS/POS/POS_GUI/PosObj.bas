Attribute VB_Name = "PosObj"
Option Explicit
Global sInBox As String
Global sOutBox As String
Global sServerInbox As String
Global sServerOutbox As String
Global sPOSSQLServer As String
Public oZSession As z_ZSession
Public Const POLL_INTERVAL As Long = 3000
Public oPC As z_POSCLIConnection

Global lngExchangeNumber As Long



'Function RoundUp(ByVal x As Long, ByVal Factor As Long)
'Dim tmp As Long
'
'    tmp = x Mod Factor
'    Select Case tmp
'    Case 0
'        RoundUp = x
'    Case Else
'        RoundUp = ((x \ Factor) * Factor) + Factor
'    End Select
'
'End Function
Public Function RoundUp(ByVal Value As Double) As Double
Dim temp As Double
    temp = Int(Value)
    If temp <> Value Then
        temp = temp + 1
    End If
    RoundUp = temp
End Function
Function RoundUp2(ByVal x As Double, ByVal Factor As Long)
Dim tmp As Long

    tmp = x Mod Factor
    Select Case tmp
    Case 0
        RoundUp2 = x
    Case Else
        RoundUp2 = ((x \ Factor) * Factor) + Factor
    End Select
    
End Function
Function RoundDown(ByVal x As Long, ByVal Factor As Long)
Dim tmp As Long
    If Factor = 0 Then
        RoundDown = x
        Exit Function
    End If
    If x > 0 Then
        tmp = x Mod Factor
        Select Case tmp
        Case 0
            RoundDown = x
        Case Else
            RoundDown = ((x \ Factor) * Factor)
        End Select
    ElseIf x >= 0 Then
        RoundDown = 0
    Else
        RoundDown = ((x \ Factor) * Factor)
    End If
End Function

