Attribute VB_Name = "General"
Option Explicit

Public flgGotFocus As Boolean





Function RemoveSpace(sStr As String) As String
    Dim i As Integer
    Dim tmpS As String
    sStr = Trim(sStr)
    i = InStr(sStr, " ")
    If i > 0 Then
        Do While i > 0
            tmpS = tmpS & Left(sStr, i - 1)
            sStr = Right(sStr, Len(sStr) - i)
            i = InStr(sStr, " ")
        Loop
        sStr = tmpS
    End If
    RemoveSpace = sStr
End Function

Function ValidIDNum(pIn) As Boolean
Dim a As Long
Dim b As Long
Dim C As Long

Dim tmpB As String
Dim ichkDigit As Integer

Dim i As Integer
    If IsNull(pIn) Then
        ValidIDNum = False
        GoTo EXIT_Handler
    End If
    If Not IsNumeric(pIn) Then
        ValidIDNum = False
        GoTo EXIT_Handler
    End If
    If Len(pIn) <> 13 Then
        ValidIDNum = False
        GoTo EXIT_Handler
    End If
    a = 0
    For i = 1 To 11 Step 2
        a = a + CInt(Mid(pIn, i, 1))
    Next i
    tmpB = ""
    For i = 2 To 12 Step 2
        tmpB = tmpB & Mid(pIn, i, 1)
    Next i
    C = 2 * (CLng(tmpB))
    tmpB = CStr(C)
    b = 0
    For i = 1 To Len(tmpB)
        b = b + CInt(Mid(tmpB, i, 1))
    Next i
    ichkDigit = 10 - Right(CStr(a + b), 1)
    If ichkDigit = 10 Then ichkDigit = 0
    If ichkDigit = CInt(Right(pIn, 1)) Then
        ValidIDNum = True
    Else
        ValidIDNum = False
    End If
    
EXIT_Handler:
    Exit Function
End Function
Function preparemailaddress2(Addressee, Add1, Add2, Add3, Add4, Add5, Add6, PCode)
On Error GoTo ERR_PrepareMailAddress2
Dim strOut As String
    If IsNull(Add1) Then
        preparemailaddress2 = ""
        GoTo EXIT_PrepareMailAddress2
    End If
    strOut = ""

    If Not IsNull(Addressee) Then
    If Len(Addressee) > 0 Then
        strOut = strOut & Addressee
      End If
    End If

    If Not IsNull(Add1) Then
    If Len(Add1) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & Add1
      End If
    End If
    
    If Not IsNull(Add2) Then
      If Len(Add2) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & Add2
      End If
    End If
    If Not IsNull(Add3) Then
    If Len(Add3) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & Add3
      End If
    End If
    If Not IsNull(Add4) Then
    If Len(Add4) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & Add4
      End If
    End If
    If Not IsNull(Add5) Then
    If Len(Add5) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & Add5
      End If
    End If
    If Not IsNull(Add6) Then
    If Len(Add6) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & Add6
      End If
    End If
    If Not IsNull(PCode) Then
    If Len(PCode) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & PCode
      End If
    End If
    preparemailaddress2 = strOut
    'rs.Close


EXIT_PrepareMailAddress2:
    Exit Function

ERR_PrepareMailAddress2:
    MsgBox "ca't format address"
    Resume EXIT_PrepareMailAddress2
End Function
Public Function GetGenderFromID(pID As String) As String
Dim strFiveToSeven As String

    If CLng(Mid(pID, 7, 4)) < 5000 Then
        GetGenderFromID = "F"
    Else
        GetGenderFromID = "M"
    End If

End Function

Public Function GetDOBFromID(pID As String) As Date
    Dim Year As String
    Dim Month As String
    Dim Day As String
    
    Year = Left(pID, 2)
    Month = Mid(pID, 3, 2)
    Day = Mid(pID, 5, 2)
    GetDOBFromID = CDate(Day & "-" & Month & "-" & Year)
End Function

Public Function HasData(pIn) As Boolean
    
        
    If IsNull(pIn) Then
        HasData = False
    ElseIf Len(pIn) = 0 Then
        HasData = False
    Else
        HasData = (Left(pIn, 1) <> Chr(0))
    End If
End Function

Public Function NumberOnly(KeyAscii As Integer) As Integer
    
    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii <> vbKeyBack Then
            Beep
            NumberOnly = 0
            Exit Function
        End If
        
    End If
    NumberOnly = KeyAscii
End Function

Public Sub SizeGrid(Grid As Control)
    With Grid
        .Columns(0).Width = 3300
        .Columns(1).Width = 2500
    End With
End Sub

Public Sub AutoSelect(CTR As Control)
    With CTR
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Public Sub ParseCurrency(objText As Control)
  Dim strTemp As String
  Dim iOff, iLen, iDec As Long
  If flgGotFocus Then
    flgGotFocus = False
    Exit Sub
  End If
  iLen = Len(objText)
  'check if new entry is a number
  If Val(objText) = 0 Or (Not IsNumeric(objText)) Then
    objText = "0.00"
    GoTo MEX
  End If
  If Val(Right(objText, 1)) = 0 And Right(objText, 1) <> "0" Then
    'not valid
    objText = Left(objText, iLen - 1)
    objText.SelStart = Len(objText)
    objText.SetFocus
    Exit Sub
  End If
  If iLen > 1 Then
    If iLen = 3 Then iDec = 1 Else iDec = 2
    'first remove old point
   iOff = InStr(1, objText, ".")
   If iOff <> 0 Then
    strTemp = Left(objText, iOff - 1) & Right(objText, iLen - iOff)
   Else
    strTemp = objText
   End If
    iLen = Len(strTemp)
    If iLen > 1 Then
      objText = Left(strTemp, iLen - iDec) & "." & Right(strTemp, iDec)
      'objText = strTemp
     objText.SelStart = iLen + 1
    Else
      objText.SelStart = iLen
    End If
  End If
MEX:
    objText = Format(Val(objText), "0.00")
    objText.SelStart = Len(objText)
    objText.SetFocus
End Sub

Public Function IsISBN(sISBN) As Boolean
Dim strISBN As String
Dim i As Integer
Dim X As Integer
Dim iMod As Integer
Dim strChk As String
    IsISBN = True
    If Len(sISBN) = 10 Then
        If Not (UCase(Right(sISBN, 1)) = "X" Or IsNumeric(Right(sISBN, 1))) Then
            IsISBN = False
            GoTo EXIT_Handler
        End If
        If Left(sISBN, 9) <= "0" Then
            IsISBN = False
            GoTo EXIT_Handler
        End If
        strISBN = Left(Right(sISBN, 10), 9)
        X = 0
        For i = 1 To 9
            X = X + (Val(Mid(sISBN, i, 1))) * Abs(i - 11)
        Next
        iMod = X Mod 11
        Select Case iMod
        Case Is > 1
           strChk = Str(11 - iMod)
        Case 1
           strChk = "X"
        Case 0
           strChk = "0"
        End Select
        If Not (UCase(Right(sISBN, 1)) = Trim(strChk)) Then
           IsISBN = False
        End If
     Else
        IsISBN = False
     End If
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
End Function

Public Function NZS(Val As Variant) As Variant
    If Left(Val, 1) = Chr(0) Then
        NZS = ""
        GoTo EXIT_Handler
    End If
    If IsNull(Val) Then
        NZS = ""
    Else
        If IsNull(Val) Then
            NZS = ""
        Else
            NZS = Trim$(Val)
        End If
    End If
EXIT_Handler:
End Function

Public Function NZ(Val As Variant) As Variant
    If IsNull(Val) Then
        NZ = 0
    Else
        If IsNull(Val) Then
            NZ = 0
        Else
            NZ = Val
        End If
    End If

End Function

Public Function Z2N(Val As Variant) As Variant
    If IsNull(Val) Or Val = 0 Then
        Z2N = Null
    Else
        Z2N = Val
    End If
End Function

Public Sub CurrencyInput(oText As TextBox, iKeyCode As Integer)
Dim sCents As String
Dim siCents As Single
Dim iLen As Integer
    sCents = CStr(Val(oText.Text) * 100)
    If Val(sCents) = 0 Then sCents = ""
    iLen = Len(sCents)
    If iKeyCode = vbKeyBack Then
        'handle back key
        If iLen >= 2 Then
            sCents = Left(sCents, iLen - 1)
        Else
            sCents = ""
        End If
    ElseIf iKeyCode = vbKeyDelete Then
        'handle delete key
        If iLen >= 2 Then
            sCents = Mid(sCents, 2)
        Else
            sCents = ""
        End If
    ElseIf iLen >= 7 Then
        'don't allow larger numbers to prevent overflow...
    ElseIf InStr("0123456789", Chr(iKeyCode)) > 0 Then
        'add numbers
        sCents = sCents & Chr(iKeyCode)
    ElseIf iKeyCode >= 96 And iKeyCode <= 105 Then
        sCents = sCents & KeyPadNum(iKeyCode)
    End If
    If Val(sCents) > 0 Then siCents = CStr(Val(sCents) / 100)
    oText.Text = Format(siCents, "0.00")
    oText.SelStart = Len(oText.Text)
End Sub

Private Function KeyPadNum(iKeyCode As Integer) As String
    KeyPadNum = Chr(iKeyCode - 48)
End Function
