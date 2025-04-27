VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ProdCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim strRaw As String
Dim strCode As String
Dim strEAN As String
Dim strSerial As String
Dim bIsEAN As Boolean
Dim bIsISBN As Boolean
Dim bIsNonISBNCode As Boolean
Event updateCodes(pEAN As String, pCODE As String)

Public Property Get ISBNFormatted() As String

End Property

Public Function Load(pCODE As String, Optional pFBook)

'This object handles all the issues of validation and conversion of codes.
'We often only have one code for a new product and it is useful to be able to derive other codes where we can.
'In particular an EAN is vital if scanning of products is to be widely used.

'If you only have one code and you don't know whether it is EAN or ISBN or INTERNAL or PRIVATE
'then provide pCode only.

'If you have both fields then provide both
'If you know whether it is a book or not, provide pfBook

'If both codes are provided then the LOAD will check to see if they correspond.
'If you only have an EAN then if pfBook is true, the function will generate an ISBN
'otherwise it will return an error

'if you have only a pcode, if pfbook is true and pcode is an ISBN it will generate a provisional EAN (starting "222")
'                          if pfbook is false or pcode is not an ISBN it will generate provisional EAN (starting "221"

Dim i As Integer
Dim strTmp As String
    strCode = ""
    strEAN = ""
    strSerial = ""
    strRaw = UCase(FNS(pCODE))
    
    i = InStr(1, strRaw, "/")
    If i > 0 Then
        strCode = Left(strRaw, i - 1)
        strSerial = Right(strRaw, i - Len(strCode))
    Else
        strCode = strRaw
        strSerial = ""
        If ValidEAN(strCode) Then
            strEAN = strCode
            If Not IsMissing(pFBook) Then
                If pFBook Then
                    SetCodesForBook EAN:=strRaw
                Else
                    SetCodesForProduct EAN:=strRaw
                End If
            Else
                If Left(strEAN, 3) = "978" Then
                    SetCodesForBook EAN:=strEAN
                End If
            End If
        ElseIf IsISBN Then
         '   strCode = strRaw
         '   If Not IsMissing(pFBook) Then
         '       If pFBook Then
                    SetCodesForBook ISBN:=strCode
         '      End If
         '   End If
        ElseIf IsCode Then  'This can only be a hash number cause ISBNs would be picked up earlier
            strCode = strRaw
            If Not IsMissing(pFBook) Then
                If pFBook Then
                    SetCodesForBook Code:=strCode
                Else
                    SetCodesForProduct Code:=strCode
                End If
            Else
                SetCodesForProduct Code:=strCode
            End If
        ElseIf isPrivateCode Then  'not a # number
            strCode = "#" & strRaw
            If Not IsMissing(pFBook) Then
                If pFBook Then
                    SetCodesForBook Code:=strCode
                Else
                    SetCodesForProduct Code:=strCode
                End If
            Else
                SetCodesForProduct Code:=strCode
            End If
        ElseIf strRaw = "#" Then
            strCode = "#"
            strEAN = ""
        End If
    End If
    RaiseEvent updateCodes(Me.EAN, Me.Code)
End Function
Public Function IStoBeGenerated() As Boolean
    IStoBeGenerated = (strCode = "#")
End Function
Public Function IsISBN() As Boolean
Dim strISBN As String
Dim i As Integer
Dim X As Integer
Dim iMod As Integer
Dim strChk As String
    IsISBN = True
    If Len(strCode) = 10 Then
       If Not (UCase(Right(strCode, 1)) = "X" Or IsNumeric(Right(strCode, 1))) Then
           IsISBN = False
           GoTo EXIT_Handler
       End If
       If Left(strCode, 9) <= "0" Then
           IsISBN = False
           GoTo EXIT_Handler
       End If
       strISBN = Left(Right(strCode, 10), 9)
       X = 0
       For i = 1 To 9
           X = X + (Val(Mid(strCode, i, 1))) * Abs(i - 11)
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
       If Not (UCase(Right(strCode, 1)) = Trim(strChk)) Then
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
Public Function FormatCode(Optional pForEXport As Boolean) As String
Dim iGroupLength As Integer
Dim iPublisherLength As Integer
Dim itest As Long
Dim strGroup As String
Dim strPublisher As String
Dim strRem As String
Dim strChk As String
Dim strSeqNum As String
Dim fForExport As Boolean
    If pForEXport = True Then
        fForExport = True
    Else
        fForExport = False
    End If
    If IsNull(strCode) Then
        FormatCode = ""
        GoTo EXIT_FormatCode
    End If
    If Left(strCode, 1) = "#" Then
        If fForExport = False Then
            FormatCode = strCode & strSerial
        Else
            FormatCode = ""
        End If
        GoTo EXIT_FormatCode
    End If
    If Len(strCode) <> 10 Then
        FormatCode = ""
        GoTo EXIT_FormatCode
    End If
    'get the group code
    itest = Val(Left(strCode, 1))
    If itest >= 0 And itest <= 7 Then
      iGroupLength = 1
    Else
      itest = Val(Left(strCode, 2))
      If itest >= 80 And itest <= 94 Then
        iGroupLength = 2
      Else
        itest = Val(Left(strCode, 3))
        If itest >= 950 And itest <= 995 Then
          iGroupLength = 3
        Else
          itest = Val(Left(strCode, 4))
          If itest >= 9960 And itest <= 9989 Then
            iGroupLength = 4
          Else
            itest = Val(Left(strCode, 5))
            If itest >= 99900 And itest <= 99999 Then
              iGroupLength = 5
            Else
              FormatCode = "ERROR"
              Exit Function
            End If
          End If
        End If
      End If
    End If
    strGroup = Left(strCode, iGroupLength)
    strRem = Right(strCode, 10 - iGroupLength)
    strChk = Right(strRem, 1)
    strRem = Left(strRem, Len(strRem) - 1)
    
    'get the publisher code
    itest = Val(Left(strRem, 2))
    If itest >= 0 And itest <= 19 Then
      iPublisherLength = 2
    Else
      itest = Val(Left(strRem, 3))
      If itest >= 200 And itest <= 699 Then
        iPublisherLength = 3
      Else
        itest = Val(Left(strRem, 4))
        If itest >= 7000 And itest <= 8499 Then
          iPublisherLength = 4
        Else
          itest = Val(Left(strRem, 5))
          If itest >= 85000 And itest <= 89999 Then
            iPublisherLength = 5
          Else
            itest = Val(Left(strRem, 6))
            If itest >= 900000 And itest <= 949999 Then
              iPublisherLength = 6
            Else
              itest = Val(Left(strRem, 7))
              If itest >= 9500000 And itest <= 9999999 Then
                iPublisherLength = 7
              Else
                FormatCode = "ERROR"
                Exit Function
              End If
            End If
          End If
        End If
      End If
    End If
    strPublisher = Left(strRem, iPublisherLength)
    strSeqNum = Right(strRem, Len(strRem) - iPublisherLength)
    FormatCode = strGroup & "-" & strPublisher & "-" & strSeqNum & "-" & strChk
EXIT_FormatCode:
End Function
Public Property Get CanHandle() As Boolean
    If IsISBN Or IsEAN Or IsCode Or isPrivateCode Then
        CanHandle = True
    Else
        CanHandle = False
    End If
End Property
Private Function ValidEAN(pIN As String) As Boolean

    bIsEAN = True
    If Len(pIN) <> 13 Then
        bIsEAN = False
    ElseIf Not IsNumeric(pIN) Then
        bIsEAN = False
    ElseIf Modulo_10(Left(pIN, 12)) <> Right(pIN, 1) Then
        bIsEAN = False
    End If
    ValidEAN = bIsEAN
End Function
Public Property Get IsEAN() As Boolean
    IsEAN = bIsEAN
End Property
Public Property Get IsCode() As Boolean
    IsCode = True
    If Len(strCode) = 10 Then
        If Not IsISBN Then
            IsCode = False
        End If
    Else
        If Left(strRaw, 1) <> "#" Or Len(strRaw) = 1 Then
                IsCode = False
        End If
    End If
    
End Property
Private Function ValidISBN(pCODE As String)

End Function
Private Function Modulo_10(pIN) As String
Dim iRes, iSumOdd, iSumEven, i, iLen As Long

   iLen = Len(pIN)
   i = iLen
   iSumOdd = 0
   Do While i > 0
      iSumOdd = iSumOdd + CInt(Mid(pIN, i, 1))
      i = i - 2
   Loop
   iSumOdd = iSumOdd * 3
   i = iLen - 1
   iSumEven = 0
   Do While i > 0
      iSumEven = iSumEven + CInt(Mid(pIN, i, 1))
      i = i - 2
   Loop
   iRes = iSumOdd + iSumEven
   Modulo_10 = (((Int(iRes / 10)) * 10) + 10) - iRes
   If Modulo_10 = "10" Then Modulo_10 = "0"
End Function
Public Function CreateCode()
Dim strTemp As String
    strTemp = Trim$(Str$(GetNextNumber(23)))
    strCode = "#" & strTemp
    strTemp = "210" & strTemp
    strEAN = strTemp & String(12 - Len(strTemp), "0")
    strEAN = strEAN & Modulo_10(strEAN)
End Function
Public Function SetCodesForBook(Optional Code, Optional EAN, Optional ISBN) As Boolean
On Error GoTo ERR_Handler

Dim strTemp
Dim strTemp2
Dim strStrippedCode

    SetCodesForBook = False
    If Not IsMissing(Code) Then
        If Left(Code, 1) = "#" Then
            strCode = Code
        Else
            strCode = "#" & Code
        End If
        strStrippedCode = Right(strCode, Len(strCode) - 1)
        If IsNumeric(strStrippedCode) Then
            strTemp = "221" & "000" & String(6 - Len(strStrippedCode), "0") & strStrippedCode
        Else
            strTemp2 = Trim$(Str$(GetNextNumber(23)))
            strTemp = "221" & "000" & String(6 - Len(strTemp2), "0") & strTemp2
        End If
        strEAN = strTemp & Modulo_10(strTemp)
        SetCodesForBook = True
    End If
    If Not IsMissing(EAN) Then
        strEAN = EAN
        strCode = ClearEAN(CStr(EAN))
        SetCodesForBook = True
    ElseIf Not IsMissing(ISBN) Then
        strCode = ISBN
        strTemp = "978" & Left(ISBN, 9)
        If IsNumeric(strTemp) Then
            strEAN = strTemp & Modulo_10(strTemp)
            SetCodesForBook = True
        End If
    End If
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    Resume EXIT_Handler
    Resume
    
End Function
Public Function SetCodesForProduct(Optional Code, Optional EAN) As Boolean
Dim strTemp
Dim strTemp2
Dim strStrippedCode
        
    SetCodesForProduct = False

    If Not IsMissing(Code) Then
        If Left(Code, 1) = "#" Then
            strCode = Code
        Else
            strCode = "#" & Code
        End If
        strStrippedCode = Right(strCode, Len(strCode) - 1)
      '  If IsNumeric(strStrippedCode) Then
      '      strTemp = "222" & "000" & String(6 - Len(strStrippedCode), "0") & Right(strStrippedCode, 6)
      '  Else
            strTemp2 = Trim$(Str$(GetNextNumber(23)))
            strTemp = "222" & "000" & String(6 - Len(strTemp2), "0") & strTemp2
      '  End If
        strEAN = strTemp & Modulo_10(strTemp)
        SetCodesForProduct = True
    End If
    If Not IsMissing(EAN) Then
        strEAN = EAN
        strCode = ""
        SetCodesForProduct = True
    End If
End Function
Function ClearEAN(pIN As String) As String

'    Accepts a string which contains an EAN code and check digit and produces an
'    unformatted ISBN string.
On Error GoTo ERR_Handler

     Dim strISBN As String
     Dim i As Integer
     Dim X As Integer
     Dim iMod As Integer
     Dim strChk As String
     If Len(pIN) = 13 Then      'the EAN code is attached
         strISBN = Left(Right(pIN, 10), 9)
         X = 0
         For i = 1 To 9
             X = X + (Val(Mid(strISBN, i, 1))) * Abs(i - 11)
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
         ClearEAN = strISBN & Trim(strChk)
      Else
         ClearEAN = pIN
      End If

EXIT_Handler:
        Exit Function

ERR_Handler:
    Select Case Err
'
        Case Else
             MsgBox Error
              Resume EXIT_Handler
    End Select
End Function

Private Function GetNextNumber(pType) As Long
On Error GoTo ERR_GetNextNumber
Dim iNextCode As Long
'Dim wks As Workspace
'Dim db As Database
Dim rs As ADODB.Recordset
Dim cnt As Integer

  
    cnt = 0
    On Error Resume Next
    
'    Do
'        Set rs = New ADODB.Recordset
'        rs.Open "tControl", oPC.CO, adOpenDynamic, adLockPessimistic
'        cnt = cnt + 1
'        If cnt > 50 Then
'            If MsgBox("Cannot allocate a transaction code. Retry?", vbOKCancel, "Table locked") = vbCancel Then
'                Error 9999
'                GoTo ERR_GetNextNumber
'            Else
'                cnt = 0
'            End If
'        End If
'        '        DoEvents
'    Loop While (Err = 3211)
'
'    On Error GoTo ERR_GetNextNumber
'    rs.Find "[ID] = " & pType
'    iNextCode = rs![Value]
'    'rs.Edit
'       rs![Value] = iNextCode + 1
'    rs.Update
    GetNextNumber = 12345
EXIT_GetNextNumber:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Exit Function

ERR_GetNextNumber:
    Select Case Err
    Case Else
        MsgBox Error
        Resume EXIT_GetNextNumber
    End Select
End Function
Public Property Get Code() As String
    Code = strCode
End Property
Public Property Get Serial() As String
    Serial = strSerial
End Property
Public Property Get EAN() As String
    EAN = strEAN
End Property
Public Property Get EANExists() As Boolean
    EANExists = Len(strEAN) = 13
End Property
Public Property Get ISBNExists() As Boolean
    ISBNExists = Len(strCode) = 10
End Property
Public Property Get PrivateCodeExists() As Boolean
    If Len(strRaw) < 10 And Len(strRaw) > 0 Then
        PrivateCodeExists = True
    End If
End Property
Public Property Get SerialExists() As Boolean
    SerialExists = Len(strSerial) > 0
End Property
Public Function isPrivateCode() As Boolean
    isPrivateCode = Len(strRaw) < 9 And Len(strRaw) > 0 And Left(strRaw, 1) <> "#"
End Function

