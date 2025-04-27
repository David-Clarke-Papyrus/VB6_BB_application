VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ProdCode_Short"
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
'Dim bIsISBN As Boolean
'Dim bIsNonISBNCode As Boolean
Event updateCodes(pEAN As String, pCODE As String)



Public Function Load(pCODE As String)
    On Error GoTo errHandler

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
'Dim strTmp As String
    strCode = ""
    strEAN = ""
    strSerial = ""
    strRaw = UCase$(FNS(pCODE))
    
    i = InStr(1, strRaw, "/")
    If i > 0 Then
        strCode = Left$(strRaw, i - 1)
        strSerial = Right$(strRaw, i - Len(strCode))
    Else
        strCode = strRaw
        strSerial = ""
        If ValidEAN(strCode) Then
            strEAN = strCode
            strCode = ClearEAN(strEAN)
        ElseIf IsCode Then  'This can only be a hash number cause ISBNs would be picked up earlier
            strCode = strRaw
        ElseIf isPrivateCode Then  'not a # number
            strCode = "#" & strRaw
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.Load(pCODE)", Array(pCODE)
End Function
Public Function IStoBeGenerated() As Boolean
    On Error GoTo errHandler
    IStoBeGenerated = (strCode = "#")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.IStoBeGenerated"
End Function
Public Function IsISBN() As Boolean
    On Error GoTo errHandler
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String
    IsISBN = True
    
    If Len(strCode) = 10 Then
       If Not (UCase$(Right$(strCode, 1)) = "X" Or IsNumeric(Right$(strCode, 1))) Then
           IsISBN = False
           GoTo EXIT_Handler
       End If
       If Left$(strCode, 9) <= "0" Then
           IsISBN = False
           GoTo EXIT_Handler
       End If
       x = 0
       For i = 1 To 9
           x = x + (val(MID(strCode, i, 1))) * Abs(i - 11)
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
       If Not (UCase$(Right$(strCode, 1)) = Trim(strChk)) Then
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
    ErrorIn "z_ProdCode.IsISBN"
End Function
Public Function IsISBNEx(pCODE As String) As Boolean
    On Error GoTo errHandler
Dim strISBN As String
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String
    IsISBNEx = True
    
    If Len(pCODE) = 10 Then
       If Not (UCase$(Right$(pCODE, 1)) = "X" Or IsNumeric(Right$(pCODE, 1))) Then
           IsISBNEx = False
           GoTo EXIT_Handler
       End If
       If Left$(pCODE, 9) <= "0" Then
           IsISBNEx = False
           GoTo EXIT_Handler
       End If
       strISBN = Left$(Right$(pCODE, 10), 9)
       x = 0
       For i = 1 To 9
           x = x + (val(MID(pCODE, i, 1))) * Abs(i - 11)
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
       If Not (UCase$(Right$(pCODE, 1)) = Trim(strChk)) Then
          IsISBNEx = False
       End If
    Else
       IsISBNEx = False
    End If
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.IsISBNEx(pCODE)", pCODE
End Function

Public Function FormatCode(Optional pForEXport As Boolean) As String
    On Error GoTo errHandler
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
    If Left$(strCode, 1) = "#" Then
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
    itest = val(Left$(strCode, 1))
    If itest >= 0 And itest <= 7 Then
      iGroupLength = 1
    Else
      itest = val(Left$(strCode, 2))
      If itest >= 80 And itest <= 94 Then
        iGroupLength = 2
      Else
        itest = val(Left$(strCode, 3))
        If itest >= 950 And itest <= 995 Then
          iGroupLength = 3
        Else
          itest = val(Left$(strCode, 4))
          If itest >= 9960 And itest <= 9989 Then
            iGroupLength = 4
          Else
            itest = val(Left$(strCode, 5))
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
    strGroup = Left$(strCode, iGroupLength)
    strRem = Right$(strCode, 10 - iGroupLength)
    strChk = Right$(strRem, 1)
    strRem = Left$(strRem, Len(strRem) - 1)
    
    'get the publisher code
    itest = val(Left$(strRem, 2))
    If itest >= 0 And itest <= 19 Then
      iPublisherLength = 2
    Else
      itest = val(Left$(strRem, 3))
      If itest >= 200 And itest <= 699 Then
        iPublisherLength = 3
      Else
        itest = val(Left$(strRem, 4))
        If itest >= 7000 And itest <= 8499 Then
          iPublisherLength = 4
        Else
          itest = val(Left$(strRem, 5))
          If itest >= 85000 And itest <= 89999 Then
            iPublisherLength = 5
          Else
            itest = val(Left$(strRem, 6))
            If itest >= 900000 And itest <= 949999 Then
              iPublisherLength = 6
            Else
              itest = val(Left$(strRem, 7))
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
    strPublisher = Left$(strRem, iPublisherLength)
    strSeqNum = Right$(strRem, Len(strRem) - iPublisherLength)
    FormatCode = strGroup & "-" & strPublisher & "-" & strSeqNum & "-" & strChk
EXIT_FormatCode:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.FormatCode(pForEXport)", pForEXport
End Function
Public Property Get CanHandle() As Boolean
    On Error GoTo errHandler
    If IsISBN Or IsEAN Or IsCode Or isPrivateCode Then
        CanHandle = True
    Else
        CanHandle = False
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.CanHandle"
End Property
Private Function ValidEAN(pIn As String) As Boolean
    On Error GoTo errHandler

    bIsEAN = True
    If Len(pIn) <> 13 Then
        bIsEAN = False
    ElseIf Not IsNumeric(pIn) Then
        bIsEAN = False
    ElseIf Modulo_10(Left$(pIn, 12)) <> Right$(pIn, 1) Then
        bIsEAN = False
    End If
    ValidEAN = bIsEAN
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.ValidEAN(pIn)", pIn
End Function
Public Property Get IsEAN() As Boolean
    IsEAN = bIsEAN
End Property
Public Property Get IsCode() As Boolean
    On Error GoTo errHandler
    IsCode = True
    If Len(strCode) = 10 And IsNumeric(Left$(strCode, 9)) Then
        If Not IsISBN Then
            IsCode = False
        End If
    Else
        If Left$(strRaw, 1) <> "#" Or Len(strRaw) = 1 Then
                IsCode = False
        End If
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.IsCode"
End Property
Private Function ValidISBN(pCODE As String)
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.ValidISBN(pCODE)", pCODE
End Function
Private Function Modulo_10(pIn) As String
    On Error GoTo errHandler
Dim iRes, iSumOdd, iSumEven, i, iLen As Long

   iLen = Len(pIn)
   i = iLen
   iSumOdd = 0
   Do While i > 0
      iSumOdd = iSumOdd + CInt(MID(pIn, i, 1))
      i = i - 2
   Loop
   iSumOdd = iSumOdd * 3
   i = iLen - 1
   iSumEven = 0
   Do While i > 0
      iSumEven = iSumEven + CInt(MID(pIn, i, 1))
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
Private Function CheckMatch(pEAN As String, pCODE As String) As Boolean
    On Error GoTo errHandler
    CheckMatch = True
    If Left$(pEAN, 3) = "978" Or Left$(pEAN, 3) = "979" Then
        If Not (Left$(Right$(pEAN, 10), 9) = Left$(pCODE, 9)) Then
            CheckMatch = False
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.CheckMatch(pEAN,pCODE)", Array(pEAN, pCODE)
End Function
Function ClearEAN(pIn As String) As String
    On Error GoTo errHandler

'    Accepts a string which contains an EAN code and check digit and produces an
'    unformatted ISBN string.

     Dim strISBN As String
     Dim i As Integer
     Dim x As Integer
     Dim iMod As Integer
     Dim strChk As String
     If Len(pIn) = 13 Then      'the EAN code is attached
         strISBN = Left$(Right$(pIn, 10), 9)
         x = 0
         For i = 1 To 9
             x = x + (val(MID(strISBN, i, 1))) * Abs(i - 11)
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
         ClearEAN = strISBN & Trim(strChk)
      Else
         ClearEAN = pIn
      End If

EXIT_Handler:
        Exit Function

'ERR_Handler:
'    Select Case Err
'
'        Case Else
'             MsgBox Error
'              Resume EXIT_Handler
'    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.ClearEAN(pIn)", pIn
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
Public Property Get ISBNExists() As Boolean
    ISBNExists = Len(strCode) = 10 And IsISBN
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
    isPrivateCode = Len(strRaw) > 2 And Left$(strRaw, 1) <> "#"
End Function

