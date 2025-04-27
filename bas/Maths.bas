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
Function RoundDown(ByVal x As Long, ByVal Factor As Long)
Dim tmp As Long

    tmp = x Mod Factor
    Select Case tmp
    Case 0
        RoundDown = x
    Case Else
        RoundDown = ((x \ Factor) * Factor)
    End Select
    
End Function
Function CurrRound(pIn As Double) As Long
On Error GoTo err_CurrRound
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

err_CurrRound:
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
Function GetMin3(lng1 As Long, lng2 As Long, lng3 As Long, Optional bExcludeZero As Boolean) As Long
Dim a() As Long
Dim i As Integer
Dim m As Long
    ReDim a(3)
    
    a(0) = lng1
    a(1) = lng2
    a(2) = lng3
    
    m = 9999
    For i = 0 To 2
        If bExcludeZero Then
            If a(i) <> 0 Then
                If m > a(i) Then
                    m = a(i)
                End If
            End If
        Else
            If m > a(i) And a(i) <> 0 Then
                m = a(i)
            End If
        End If
    Next
    GetMin3 = m
'    If bExcludeZero Then
'        If lng1 = 0 Then
'            GetMin = lng2
'            Exit Function
'        ElseIf lng2 = 0 Then
'            GetMin = lng1
'            Exit Function
'        End If
'    End If
'    If lng1 > lng2 Then
'        GetMin = lng2
'    Else
'        GetMin = lng1
'    End If
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
Function Absolute_Lng(p1 As Long)
    If p1 < 0 Then
        Absolute_Lng = p1 * -1
    Else
        Absolute_Lng = p1
    End If
End Function

Function GetMod(p1 As Long)
    If p1 < 0 Then
        GetMod = p1 * -1
    Else
        GetMod = p1
    End If
End Function

Public Function PTAdjustment(RRP As Long, Disc As Double, Within As Long, B1Max As Long, B2Max As Long, B3Max As Long, B1MU As Long, B2MU As Long, B3MU As Long, pRoundTo As Integer) As Long
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
    
    If Disc <> 0 Then
        tmp3 = tmp3 * ((100 - Disc) / 100)
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
    
    If lngPrice > 0 Then
        Markup = ((lngPrice - lngCost) / lngPrice) * 100
    Else
        Markup = 0
    End If
End Function
Public Function MinimumSP(lngCost As Long) As Long
#If H_CENTRAL = 1 Then
    MinimumSP = 0
#Else
    MinimumSP = lngCost + ((oPC.Configuration.MinMU / 100) * lngCost)
#End If
End Function


Public Function Round17(ByRef v As Double, Optional ByVal lngDecimals As Long = 0) As Double
  ' By Filipe Lage
  ' fclage@garciaresende.com
  ' msn: fclage@clix.pt
  ' Revision C by Donald - 20060201
  Dim xint As Double, yint As Double, xrest As Double
  Static PreviousValue    As Double
  Static PreviousDecimals As Long
  Static PreviousOutput   As Double
  Static m                As Double
  If m = 0 Then m = 1
      ' Initialization - M is never 0 (it is always 10 ^ n)
  
  If PreviousValue = v And PreviousDecimals = lngDecimals Then Round17 = PreviousOutput: Exit Function
      ' Hey... it's the same number and decimals as before...
      ' So, the actual result is the same. No need to recalc it
  
  If v = 0 Then Exit Function
      ' no matter what rounding is made, 0 is always rounded to 0
      
  If PreviousDecimals = lngDecimals Then
      Else
      ' A different number of decimal places, means a new Multiplier
      PreviousDecimals = lngDecimals
      m = 10 ^ lngDecimals
      End If
  
  If m = 1 Then xint = v Else xint = v * CDec(m)
      ' Let's consider the multiplication of the number by the multiplier
      ' Bug fixed: If you just multiplied the value by M, those nasty reals came up
      ' So, we use CDEC(m) to avoid that
                                                              
  Round17 = Fix(xint)
      ' The real integer of the number (unlike INT, FIX reports the actual number)
  
  ' 20060201: fix by Donald
  If Abs(Fix(10 * (xint - Round17))) > 4 Then
    If Round17 >= 0 Then
      Round17 = Round17 + 1
    Else
      Round17 = Round17 - 1
    End If
  End If
      ' First decimal is 5 or bigger ? If so, we'll add +1 or -1 to the result (later to be divided by M)
  
  If m = 1 Then Else Round17 = Round17 / m
      ' Divides by the multiplier. But we only need to devide if M isn't 1
  
  PreviousOutput = Round17
  PreviousValue = v
      ' Let's save this last result in memory... may be handy ;)
End Function
