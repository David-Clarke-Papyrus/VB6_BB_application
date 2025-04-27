Attribute VB_Name = "Maths"
Function RoundUp(ByVal x As Long, ByVal Factor As Long)
Dim tmp As Long

    tmp = x Mod Factor
    Select Case tmp
    Case 0
        RoundUp = x
    Case Else
        RoundUp = ((x \ Factor) * Factor) + Factor
    End Select
    
End Function

Function CurrRound(pIn As Double) As Long
On Error GoTo ERR_CurrRound
Dim pout As Long

    If ((pIn Mod 10 = 0)) And (pIn - Int(pIn) = 0) Then
        pout = pIn
        GoTo EXIT_CurrRound
    End If

    If (pIn < 50) Then
        pout = (Int(pIn \ 5 + 1)) * 5
    Else
        pout = (Int(pIn \ 10 + 1)) * 10
    End If

EXIT_CurrRound:
    CurrRound = pout

    Exit Function

ERR_CurrRound:
    MsgBox Error
    Exit Function

End Function
Function GetMax(lng1 As Long, lng2 As Long) As Long
    If lng1 > lng2 Then
        GetMax = lng1
    Else
        GetMax = lng2
    End If
End Function
Function GetMin(lng1 As Long, lng2 As Long, Optional bExcludeZero As Boolean) As Long
    If bExcludeZero Then
        If lng1 = 0 Then
            GetMin = lng2
            Exit Function
        ElseIf lng2 = 0 Then
            GetMin = lng1
            Exit Function
        End If
    End If
    If lng1 > lng2 Then
        GetMin = lng2
    Else
        GetMin = lng1
    End If
End Function

Function NonNegative_Lng(p1 As Long)
    If p1 < 0 Then
        NonNegative_Lng = 0
    Else
        NonNegative_Lng = p1
    End If
End Function
Function NonNegative_Dbl(p1 As Double)
    If p1 < 0 Then
        NonNegative_Dbl = 0
    Else
        NonNegative_Dbl = p1
    End If
End Function

Function GetMod(p1 As Long)
    If p1 < 0 Then
        GetMod = p1 * -1
    Else
        GetMod = p1
    End If
End Function

Public Function PTAdjustment(RRP As Long, DISC As Double, Within As Long, B1Max As Long, B2Max As Long, B3Max As Long, B1MU As Long, B2MU As Long, B3MU As Long, pRoundTo As Integer) As Long
Dim tmp As Double
Dim tmp2 As Long
Dim tmp3 As Double
Dim tmp4 As Double

    Select Case RRP
    Case Is < (B1Max * 100)
        tmp3 = RRP + (CLng(B1MU) * 100)
    Case Is < (B2Max * 100)
        tmp3 = RRP + (CLng(B2MU) * 100)
    Case Is < (B3Max * 100)
        tmp3 = RRP + (CLng(B3MU) * 100)
    End Select
    
    If DISC <> 0 Then
        tmp3 = tmp3 * ((100 - DISC) / 100)
    End If
    If pRoundTo > 0 Then
        tmp3 = RoundUp(tmp3, pRoundTo)
    End If
    
    tmp4 = RoundUp(tmp3, 100)
    If tmp4 - tmp3 <= Within Then
        tmp3 = tmp4
    End If
    PTAdjustment = CLng(tmp3)
    
    
End Function
Public Function Markup(lngPrice As Long, lngCost As Long) As Double
    If lngCost > 0 Then
        Markup = ((lngPrice - lngCost) / lngCost) * 100
    Else
        Markup = 0
    End If
End Function

