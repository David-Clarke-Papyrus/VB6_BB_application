Attribute VB_Name = "General_POS"
Option Explicit
Public flgGotFocus As Boolean

Function RemoveSpace(sStr As String) As String
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.RemoveSpace(sStr)", sStr
End Function

Function ValidIDNum(pIn) As Boolean
    On Error GoTo errHandler
Dim a As Long

Dim b As Long
Dim c As Long

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
        a = a + CInt(MID(pIn, i, 1))
    Next i
    tmpB = ""
    For i = 2 To 12 Step 2
        tmpB = tmpB & MID(pIn, i, 1)
    Next i
    c = 2 * (CLng(tmpB))
    tmpB = CStr(c)
    b = 0
    For i = 1 To Len(tmpB)
        b = b + CInt(MID(tmpB, i, 1))
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.ValidIDNum(pIn)", pIn
End Function
Function preparemailaddress2(Addressee, Add1, Add2, Add3, Add4, Add5, Add6, pCODE)
    On Error GoTo errHandler
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
    If Not IsNull(pCODE) Then
    If Len(pCODE) > 0 Then
        If Len(strOut) > 0 Then strOut = strOut & Chr(13) & Chr(10)
        strOut = strOut & pCODE
      End If
    End If
    preparemailaddress2 = strOut
    'rs.Close


EXIT_PrepareMailAddress2:
    Exit Function

'ERR_PrepareMailAddress2:
'    MsgBox "ca't format address"
'    Resume EXIT_PrepareMailAddress2
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.preparemailaddress2(Addressee,Add1,Add2,Add3,Add4,Add5,Add6,PCode)", _
         Array(Addressee, Add1, Add2, Add3, Add4, Add5, Add6, pCODE)
End Function
Public Function GetGenderFromID(PID As String) As String
    On Error GoTo errHandler
Dim strFiveToSeven As String

    If CLng(MID(PID, 7, 4)) < 5000 Then
        GetGenderFromID = "F"
    Else
        GetGenderFromID = "M"
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.GetGenderFromID(pID)", PID
End Function

Public Function GetDOBFromID(PID As String) As Date
    On Error GoTo errHandler
    Dim Year As String
    Dim Month As String
    Dim Day As String
    
    Year = Left(PID, 2)
    Month = MID(PID, 3, 2)
    Day = MID(PID, 5, 2)
    GetDOBFromID = CDate(Day & "-" & Month & "-" & Year)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.GetDOBFromID(pID)", PID
End Function

Public Function HasData(pIn) As Boolean
    On Error GoTo errHandler
    
        
    If IsNull(pIn) Then
        HasData = False
    ElseIf Len(pIn) = 0 Then
        HasData = False
    Else
        HasData = (Left(pIn, 1) <> Chr(0))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.HasData(pIn)", pIn
End Function

Public Function NumberOnly(KeyAscii As Integer) As Integer
    On Error GoTo errHandler
    
    If Not IsNumeric(Chr(KeyAscii)) Then
        If KeyAscii <> vbKeyBack Then
            Beep
            NumberOnly = 0
            Exit Function
        End If
        
    End If
    NumberOnly = KeyAscii
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.NumberOnly(KeyAscii)", KeyAscii
End Function

Public Sub SizeGrid(Grid As Control)
    On Error GoTo errHandler
    With Grid
        .DBCONNlumns(0).Width = 3300
        .DBCONNlumns(1).Width = 2500
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.SizeGrid(Grid)", Grid
End Sub

Public Sub AutoSelect(CTR As Control)
    On Error Resume Next
    With CTR
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.AutoSelect(CTR)", CTR
End Sub

Public Sub ParseCurrency(objText As Control)
    On Error GoTo errHandler
  Dim strTemp As String
  Dim iOff, iLen, iDec As Long
  If flgGotFocus Then
    flgGotFocus = False
    Exit Sub
  End If
  iLen = Len(objText)
  'check if new entry is a number
  If val(objText) = 0 Or (Not IsNumeric(objText)) Then
    objText = "0.00"
    GoTo MEX
  End If
  If val(Right(objText, 1)) = 0 And Right(objText, 1) <> "0" Then
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
    objText = Format(val(objText), "0.00")
    objText.SelStart = Len(objText)
    objText.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.ParseCurrency(objText)", objText
End Sub

Public Function IsISBN(sISBN) As Boolean
    On Error GoTo errHandler
Dim strISBN As String
Dim i As Integer
Dim x As Integer
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
        x = 0
        For i = 1 To 9
            x = x + (val(MID(sISBN, i, 1))) * Abs(i - 11)
        Next
        iMod = x Mod 11
        Select Case iMod
        Case Is > 1
           strChk = str(11 - iMod)
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.IsISBN(sISBN)", sISBN
End Function

Public Function NZS(val As Variant) As Variant
    On Error GoTo errHandler
    If Left(val, 1) = Chr(0) Then
        NZS = ""
        GoTo EXIT_Handler
    End If
    If IsNull(val) Then
        NZS = ""
    Else
        If IsNull(val) Then
            NZS = ""
        Else
            NZS = Trim$(val)
        End If
    End If
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.NZS(Val)", val
End Function

Public Function NZ(val As Variant) As Variant
    On Error GoTo errHandler
    If IsNull(val) Then
        NZ = 0
    Else
        If IsNull(val) Then
            NZ = 0
        Else
            NZ = val
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.NZ(Val)", val
End Function

Public Function Z2N(val As Variant) As Variant
    On Error GoTo errHandler
    If IsNull(val) Or val = 0 Then
        Z2N = Null
    Else
        Z2N = val
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.Z2N(Val)", val
End Function

Public Sub CurrencyInput(oText As TextBox, iKeyCode As Integer)
    On Error GoTo errHandler
Dim sCents As String
Dim siCents As Single
Dim iLen As Integer
    sCents = CStr(val(oText.Text) * 100)
    If val(sCents) = 0 Then sCents = ""
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
            sCents = MID(sCents, 2)
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
    If val(sCents) > 0 Then siCents = CStr(val(sCents) / 100)
    oText.Text = Format(siCents, "0.00")
    oText.SelStart = Len(oText.Text)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.CurrencyInput(oText,iKeyCode)", Array(oText, iKeyCode)
End Sub

Private Function KeyPadNum(iKeyCode As Integer) As String
    On Error GoTo errHandler
    KeyPadNum = Chr(iKeyCode - 48)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "General_POS.KeyPadNum(iKeyCode)", iKeyCode
End Function

Sub WAIT(pLength As Long)
Dim i
For i = 1 To pLength
Next i
End Sub

Public Function Centre(sText As String, iWidth As Integer) As String
    On Error GoTo errHandler
    If Len(sText) < iWidth Then
        Centre = Space((iWidth - Len(sText)) \ 2) & sText
    Else
        Centre = sText
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.Centre(sText,iWidth)", Array(sText, iWidth)
End Function
Public Sub LoadCombo(oRS As ADODB.Recordset, oCBO As ComboBox)
    On Error GoTo errHandler
    If Not oRS.EOF Then
        oRS.MoveFirst
        oCBO.Clear
    
        With oRS
            Do While Not oRS.EOF
              If Not IsNull(!SM_Shortname) Then
                oCBO.AddItem !SM_Name & " (" & !SM_Shortname & ")"
              Else
                oCBO.AddItem !SM_Name
              End If
                oCBO.ITEMDATA(oCBO.NewIndex) = !SM_ID
                .MoveNext
            Loop
            .MoveFirst
        End With
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Globals_POS.LoadCombo(oRS,oCBO)", Array(oRS, oCBO)
End Sub

'//--------------------------------------------------------------------
'// PURPOSE:
'// Pause or delay a procedure for a specified number of seconds
'//
'// ARGUMENTS:
'// Number of seconds. May use fractions in a decimal format (#.##)
'//
'// COMMENTS:
'// Timer() returns a Single value rounded to the nearest 1/100 of a
'// second like a stopwatch. Also, Timer() has a "bug" - it resets
'// itself at midnight. Therefore we need to adjust for this, using
'// some sort of counter. The simplest way is to concatenate the day
'// in front of it with Day(Date) but then the days get reset when the
'// month changes, and of course we need to adjust when the months are
'// reset by the changing year. Fortunately that's as far as we have
'// to go. To avoid an extremely large number by concatenating one in
'// front of the other, we add the different parts of the Date together
'// and then concatenate with the sum.
'//--------------------------------------------------------------------
Public Sub EventPause(sngSeconds As Single)

    '// A Single will convert to scientific notation when concatenating a
    '//  number resulting in 8-digits or more. This can introduce inaccuracies
    '//  as a result of the number being rounded when converted. Therefore we
    '//  must declare doubles when working with the date counter to avoid
    '//  converting to scientific notation.
    Dim dblTotal As Double, dblDateCounter As Double, sngStart As Single
    Dim dblReset As Double, sngTotalSecs As Single, intTemp As Integer
        '// For our purposes, it's better to concatenate five zeros onto the
        '//  end of our date counter, then ADD any Timer values to it.
        dblDateCounter = ((Year(Date) + Month(Date) + Day(Date)) _
          & 0 & 0 & 0 & 0 & 0)
        '// Initialize start time.
        sngStart = Timer
        '// We also need to adjust for the possible resetting of Timer()
        '//  (such as if the Time happens to be just before midnight) when
        '//  adding the Pause time onto the Start time. The folowing formula
        '//  takes ANY value of the total seconds, whether it's above or below
        '//  the 86400 limit, and converts it to a format compatible to the
        '//  date counter.
        sngTotalSecs = (sngStart + sngSeconds)
        intTemp = (sngTotalSecs \ 86400)   '// Return the integer portion only
        dblReset = (intTemp * 100000) + (sngTotalSecs - (intTemp * 86400))
        '// Now we can initialize our total time.
        dblTotal = dblDateCounter + dblReset
    
    '// Timer loop
    Do
        DoEvents        '// Make sure any other tasks get some attention
    '// For this to work properly, we cannot create a variable with the
    '//  concatenated expression and plug it in unless we reset the variable
    '//  during the loop. Much better to do it like this:
    Loop While (dblDateCounter + Timer) < dblTotal
    
End Sub

