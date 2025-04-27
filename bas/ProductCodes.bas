Attribute VB_Name = "ProductCodes"
Option Explicit


Public Function ISBN13to10(pIn As String) As String
    On Error GoTo errHandler

Dim strISBN As String
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String

     If Len(pIn) = 13 Then      'the EAN code is attached
        If Left(pIn, 3) = "978" Then
            strISBN = Left(Right(pIn, 10), 9)
            x = 0
            For i = 1 To 9
                x = x + (val(Mid(strISBN, i, 1))) * Abs(i - 11)
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
            ISBN13to10 = strISBN & Trim(strChk)
        Else
            ISBN13to10 = ""
        End If
      Else
            If IsISBN10(pIn) Then
                ISBN13to10 = pIn
            Else
                ISBN13to10 = ""
            End If
      End If

EXIT_Handler:
        Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.ClearEAN(pIn)", pIn
End Function

Public Function FormatProductCode(pIn As String, pForEXport As Boolean) As String
 FormatProductCode = pIn
End Function

Public Function IsISBN13(pIn As String, Optional IsBook, Optional IsLocal) As Boolean
Dim bIsISBN13 As Boolean

    If Not IsNumeric(pIn) Or (Not OnlyNumbers(pIn)) Then
        IsISBN13 = False
        Exit Function
    End If
    bIsISBN13 = True
    If Len(pIn) <> 13 Then
        bIsISBN13 = False
      '  MsgBox "Pos 1"
    ElseIf Not IsNumeric(pIn) Then
        bIsISBN13 = False
     '   MsgBox "Pos 2"
    ElseIf Modulo_10(Left(pIn, 12)) <> Right(pIn, 1) Then
        bIsISBN13 = False
     '   MsgBox "Pos 3"
    End If
    
    If Not IsMissing(IsBook) Then
         '       MsgBox "Pos 4"
        If IsBook Then
            If Not (Left(pIn, 3) = "978" Or Left(pIn, 3) = "979") Then
                bIsISBN13 = False
            End If
        Else
            If Left(pIn, 3) = "978" Or Left(pIn, 3) = "979" Then
                bIsISBN13 = False
            End If
        End If
    End If
    
    If Not IsMissing(IsLocal) Then
      '      MsgBox "Pos 5"
        If IsLocal Then
            If Not (Left(pIn, 1) = "2") Then
                bIsISBN13 = False
            End If
        Else
            If Left(pIn, 1) = "2" Then
                bIsISBN13 = False
            End If
        End If
    End If
        
    
    IsISBN13 = bIsISBN13

End Function
Public Function IsHashCode(pIn As String) As Boolean
    IsHashCode = (Left(pIn, 1) = "#")
End Function


Public Function Modulo_10(pIn) As String
    On Error GoTo errHandler
Dim iRes, iSumOdd, iSumEven, i, iLen As Long

   iLen = Len(pIn)
   i = iLen
   iSumOdd = 0
   Do While i > 0
      iSumOdd = iSumOdd + CInt(Mid(pIn, i, 1))
      i = i - 2
   Loop
   iSumOdd = iSumOdd * 3
   i = iLen - 1
   iSumEven = 0
   Do While i > 0
      iSumEven = iSumEven + CInt(Mid(pIn, i, 1))
      i = i - 2
   Loop
   iRes = iSumOdd + iSumEven
   Modulo_10 = (((Int(iRes / 10)) * 10) + 10) - iRes
   If Modulo_10 = "10" Then Modulo_10 = "0"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.Modulo_10(pIn)", pIn
End Function

Public Function IsISBN10(pIn) As Boolean
Dim strCode As String
Dim strISBN As String
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String

    strCode = Trim$(pIn)
    IsISBN10 = True
    
    If Len(strCode) = 10 Then
       If Not (UCase(Right(strCode, 1)) = "X" Or IsNumeric(Right(strCode, 1))) Then
           IsISBN10 = False
           GoTo EXIT_Handler
       End If
       If Left(strCode, 9) <= "0" Then
           IsISBN10 = False
           GoTo EXIT_Handler
       End If
       strISBN = Left(Right(strCode, 10), 9)
       If Not (IsNumeric(strISBN)) Then
           IsISBN10 = False
           GoTo EXIT_Handler
       End If
       x = 0
       For i = 1 To 9
           x = x + (val(Mid(strCode, i, 1))) * Abs(i - 11)
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
       If Not (UCase(Right(strCode, 1)) = Trim(strChk)) Then
          IsISBN10 = False
       End If
    Else
        IsISBN10 = False
    End If
    Exit Function
    
EXIT_Handler:
End Function

Public Function IsPrivateCode(pIn As String) As Boolean
    IsPrivateCode = Len(pIn) > 2 And Len(pIn) < 10 And Left(pIn, 1) <> "#" And Left(pIn, 1) <> "/"
End Function

Public Function StripSerial(pIn As String) As String
    If InStr(1, pIn, "/") > 0 Then
        StripSerial = Left(pIn, InStr(1, pIn, "/") - 1)
    Else
        StripSerial = FNS(pIn)
    End If
End Function

Public Function FormatISBN13(p As String) As String
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

    If Not IsNumeric(p) Then
        FormatISBN13 = p
        Exit Function
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "FormatISBN13"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@IN", adVarChar, adParamInput, 15, p)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@OUT", adVarChar, adParamOutput, 20, f)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FormatISBN13 = cmd.Parameters("@OUT").Value
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.FormatISBN13(p)", p
End Function

Public Function formatSKU(pEAN As String, pCode As String, bForExport As Boolean) As String
    If IsISBN13(pEAN) And Not (Left(pEAN, 2) = "22" Or Left(pEAN, 2) = "23" Or Left(pEAN, 2) = "24" Or Left(pEAN, 2) = "25") Then
        formatSKU = FormatISBN13(pEAN)
    Else
        If bForExport = False Then
            formatSKU = pCode
        End If
    End If
End Function
Public Function ConvertISBN10toEAN(p As String) As String
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f

'    If Not IsNumeric(p) Then
'        FormatISBN13 = p
'        Exit Function
'    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "ConvertISBNtoISBN13"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@ISBN", adVarChar, adParamInput, 15, p)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@OUT", adVarChar, adParamOutput, 20, f)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    f = cmd.Parameters("@OUT").Value
    ConvertISBN10toEAN = IIf(IsNull(f), p, f)
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_StockManager.ConvertISBN10toEAN(p)", p
End Function

