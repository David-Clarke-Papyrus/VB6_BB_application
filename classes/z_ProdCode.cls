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
Dim bCodesOK As Boolean
Dim oSQL As New z_SQL

Event updateCodes(pEAN As String, pCode As String)

Public Property Get ISBNFormatted() As String
    ISBNFormatted = FormatCode(False)
End Property
Public Property Get CodesOK() As Boolean
CodesOK = bCodesOK
End Property
'Public Function LoadOLD(pCODE As String, Optional pFBook)
'    On Error GoTo ErrHandler
'
''This object handles all the issues of validation and conversion of codes.
''We often only have one code for a new product and it is useful to be able to derive other codes where we can.
''In particular an EAN is vital if scanning of products is to be widely used.
'
''If you only have one code and you don't know whether it is EAN or ISBN or INTERNAL or PRIVATE
''then provide pCode only.
'
''If you have both fields then provide both
''If you know whether it is a book or not, provide pfBook
'
''If both codes are provided then the LOAD will check to see if they correspond.
''If you only have an EAN then if pfBook is true, the function will generate an ISBN
''otherwise it will return an or
'
''if you have only a pcode, if pfbook is true and pcode is an ISBN it will generate a provisional EAN (starting "222")
''                          if pfbook is false or pcode is not an ISBN it will generate provisional EAN (starting "221"
'
'Dim i As Integer
'Dim strTmp As String
'    strCode = ""
'    strEAN = ""
'    strSerial = ""
'    strRaw = UCase(FNS(pCODE))
'
'    i = InStr(1, strRaw, "/")
'    If i > 0 Then
'        strCode = Left(strRaw, i - 1)
'        strSerial = Right(strRaw, Len(strRaw) - i)
'        If Not (IsNumeric(strSerial)) Then
'            bIsEAN = False
'            bIsISBN = False
'            bIsNonISBNCode = False
'        End If
'    Else
'        strCode = strRaw
'        strSerial = ""
'        If ValidEAN(strCode) Then
'            strEAN = strCode
'            If Not IsMissing(pFBook) Then
'                If pFBook Then
'                    If Left(strEAN, 3) = "978" Then
'                        strCode = ClearEAN(strCode)
'                    End If
''                Else
''                    SetCodesForProduct EAN:=strRaw
'                End If
''            Else
''                If Left(strEAN, 3) = "978" Then
''                    SetCodesForBook EAN:=strEAN
''                End If
'            End If
'        ElseIf IsISBN Then
'         '   strCode = strRaw
'         '   If Not IsMissing(pFBook) Then
'         '       If pFBook Then
'                    SetCodesForBook ISBN:=strCode
'         '      End If
'         '   End If
'        ElseIf IsCode Then  'This can only be a hash number cause ISBNs would be picked up earlier
'            strCode = strRaw
''            If Not IsMissing(pFBook) Then
''                If pFBook Then
''                    SetCodesForBook Code:=strCode
''                Else
''                    SetCodesForProduct Code:=strCode
''                End If
''            Else
''                SetCodesForProduct Code:=strCode
''            End If
'        ElseIf isPrivateCode Then  'not a # number
'            strCode = "#" & strRaw
''            If Not IsMissing(pFBook) Then
''                If pFBook Then
''                    SetCodesForBook Code:=strCode
''                Else
''                    SetCodesForProduct Code:=strCode
''                End If
''            Else
''                SetCodesForProduct Code:=strCode
''            End If
''        ElseIf strRaw = "#" Then
''            strCode = "#"
''            strEAN = ""
'        End If
'    End If
'   ' RaiseEvent updateCodes(Me.EAN, Me.Code)
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_ProdCode.Load(pCODE,pFBook)", Array(pCODE, pFBook)
'End Function
Public Function LoadNew(ByRef pCode As String, ByRef pEAN As String, ByRef IsNew As Boolean, Optional pFBook As Variant, Optional pMsg As String, Optional pReadonly As Variant) As Boolean
    On Error GoTo errHandler
'Here both codes are supplied or either, but it is assumed that all available codes are supplied;
'so:    if there is only a pCode - a new EAN will be generated (a 221...)
'       if only a pEAN is supplied, then if it is a ISBN13 with a 978 prefix a ISBN10 will be placed in pCode
'       if both are supplied they will be crosschecked and validated before being accepted
'       If pFBook is true EAN will be validated that it is in set( 978,979) and vice versa

Dim i As Integer
Dim strTmp As String
Dim bOK As Boolean



'Prepare fields
    strCode = ""
    strEAN = FNS(pEAN)
    strSerial = ""
    strRaw = UCase(FNS(pCode))
'If IsMissing(pFBook) Then pFBook = False
'Check to see if the pCode specifies a copy number (antiquarian only)
    i = InStr(1, strRaw, "/")
    If i > 0 Then
        strCode = Left(strRaw, i - 1)
        strSerial = Right(strRaw, Len(strRaw) - i)
        If Not (IsNumeric(strSerial)) Then
            LoadNew = False
            Exit Function
        End If
    Else
        strCode = strRaw
        strSerial = ""
    End If
    bOK = False
    If strCode > "" And strEAN > "" Then    'BOTH codes are supplied
        
        If Not IsMissing(pFBook) Then
            If pFBook Then
                If IsISBN13(strEAN, True, False) Then           'Is it a proper book EAN
                    If IsISBN10(strCode) Then
                        If ISBN13toISBN10(strEAN) = strCode Then
                            bOK = True
                        Else
                            pMsg = "ISBN-10 does not match ISBN-13"
                            bOK = False
                        End If
                    Else
                        If IsValidLocalCode(strCode) Then
                            bOK = True
                        Else
                            pMsg = "Invalid local code"
                            bOK = False
                        End If
                    End If
                ElseIf IsISBN13(strEAN, True, True) Then         'Is it a proper EAN but a local number
                    If IsISBN10(strCode) Then       'PROBLEM: we should have a proper book EAN
                        bOK = False
                        pMsg = "ISBN-10 does not match ISBN-13"
                    Else                            'OK we have a local code
                        bOK = True
                    End If
                ElseIf IsISBN13(strEAN, False, True) Then        'Is it a proper EAN but a local number
                    If IsISBN10(strCode) Then       'PROBLEM: we should have a proper book EAN
                        bOK = False
                        pMsg = "ISBN-10 does not match ISBN-13"
                    Else                            'OK we have a local code
                        bOK = True
                    End If
                ElseIf IsISBN13(strEAN, False, False) Then        'Is it not a book  but
                    If IsISBN10(strCode) Then       'PROBLEM: we should have a proper book EAN
                        bOK = False
                        pMsg = "ISBN-10 does not match ISBN-13"
                    Else                            'OK we have a local code
                        bOK = True
                    End If
                Else
                        pMsg = "This does not look like a book EAN/ISBN13 code"
                        bOK = False
                End If
            Else
                If IsISBN13(strEAN, True, False) Then           'Is it a proper book EAN
                        pMsg = "This is a book ISBN-13 but the product is a non-book"
                        bOK = False
                ElseIf IsISBN13(strEAN, False, False) Then         'Is it a proper EAN for a non-book
                    If IsISBN10(strCode) Then
                        pMsg = "ISBN-10 does not match EAN for non-book"
                        bOK = False
                    Else
                        bOK = True
                    End If
                ElseIf IsISBN13(strEAN, False, True) Then         'Is it a proper EAN but local
                    If IsISBN10(strCode) Then
                        pMsg = "ISBN-10 does not match ISBN-13"
                        bOK = False
                    Else
                        bOK = True
                    End If
                End If
            End If
        Else
                If IsISBN13(strEAN) Then
                    If IsISBN10(strCode) Then
           '             If ISBN13toISBN10(strEAn) = strCode Then
                            bOK = True
           '             Else
           '                 pMsg = "ISBN-10 does not match ISBN-13"
           '                 bOK = False
           '            End If
                    Else
                        If strCode = "#" Then
                            strCode = Trim$(str$(GetNextNumber(23)))
                            strCode = "#" & strCode
                            bOK = True
                        Else
                            If Left(strCode, 1) = "#" And IsNew = True Then
                                If IsNumeric(Mid(strCode, 2, 99)) Then
                                    pMsg = "You cannot create hash short codes that are numeric only," & vbCrLf & "these numbers are created automatically by Papyrus when you place only a hash symbol (#) in the code field."
                                    bOK = False
                                Else
                                    If oSQL.IsCodeUsedOnDB(strCode) = False Then
                                        bOK = True
                                    Else
                                        If strEAN = "" Then
                                            pMsg = "This short code is already used."
                                            bOK = False
                                        Else
                                            bOK = True
                                        End If
                                    End If
                                End If
                            Else
                                bOK = True
                            End If
                        End If
                    End If
            Else
                bOK = False
            End If
        End If
    
    ElseIf strEAN > "" Then                         'Only pEAN is supplied
        'Is EAN OK
        If Not IsMissing(pFBook) Then
            If pFBook Then
                If IsISBN13(strEAN, True, False) Then   'Is it a proper book EAN
                    strCode = ISBN13to10(pEAN)        'Will generate ISBN-10 if EAN starts 978
                    bOK = True
                Else
                  '  pMsg = "Invalid ISBN-13 for book"   -we have changed our mind and will accept an number as a book
                    bOK = True
                End If
            Else
                If IsISBN13(strEAN, False) Then           'Is it a proper NON-book EAN
                'should we generate a # number here?
                    bOK = True
                Else
                    pMsg = "Missing local code"   '''''''''''''
                    bOK = False
                End If
            End If
        Else
                If IsISBN13(strEAN) Then           'Is it a proper NON-book EAN
                    If Left(strEAN, 3) = "978" Then
                        If strCode = "" Then
                            strCode = ISBN13to10(strEAN)
                        End If
                    End If
                'should we generate a # number here?
                    bOK = True
                Else
                    pMsg = "Missing local code"   '''''''''''''
                    bOK = False
                End If
        End If
    
    ElseIf strCode > "" Then                         'Only pCode is supplied
        'Is Code OK
        If Not IsMissing(pFBook) Then
            If pFBook Then
                If IsISBN10(strCode) Then   'Is it a proper book ISBN10?
                    bOK = True
                    generateISBN13 strCode, True    'will generate a proper ISBN-13 code with 978 prefix from the ISBN-10
                ElseIf strCode = "#" Then
                    If Not pReadonly Then GenerateBothCodes
                    bOK = True
                ElseIf IsHashCode(strCode) Then
                    If Not pReadonly Then generateISBN13 strCode, False   'will generate a local EAN code
                    bOK = True
                ElseIf Len(strCode) = 10 Then
                    If IsNumeric(Left(strCode, 9)) Then
                         pMsg = "Invalid ISBN-10 for book"
                        bOK = False
                    Else
                        bOK = True
                        generateISBN13 strCode, False
                    End If
                Else
                    generateISBN13 strCode, False
                    bOK = True
                End If
            Else
                If IsISBN10(strCode) Then           'Is it a proper book ISBN?
                    pMsg = "Proper ISBN-10 for a non-book product"   '''''''''''''
                    bOK = False
                ElseIf strCode = "#" Then
                    If Not pReadonly Then GenerateBothCodes
                    bOK = True
                ElseIf IsHashCode(strCode) Then
                    If Not pReadonly Then generateISBN13 strCode, False   'will generate a local EAN code
                    bOK = True
                Else
                    bOK = True
                    If Not pReadonly Then generateISBN13 strCode, False   'will generate a local EAN code
                End If
            End If
        ElseIf Left(strCode, 1) = "#" And IsNew = True Then
                If IsNumeric(Mid(strCode, 2, 99)) Then
                    pMsg = "You cannot create hash short codes that are numeric only," & vbCrLf & "these numbers are created automatically by Papyrus when you place only a hash symbol (#) in the code field."
                    bOK = False
                    GoTo returnresult
                Else
                    If oSQL.IsCodeUsedOnDB(strCode) = False Then
                        bOK = True
                        generateISBN13 strCode, True    'will generate a proper ISBN-13 code with 978 prefix from the ISBN-10
                    Else
                        If strEAN = "" Then
                            pMsg = "This short code is already used."
                            bOK = False
                            GoTo returnresult
                        Else
                            bOK = True
                        End If
                    End If
                End If

        
        
        
'                If IsISBN10(strCode) Then   'Is it a proper book ISBN10?
'                    bOK = True
'                    generateISBN13 strCode, True    'will generate a proper ISBN-13 code with 978 prefix from the ISBN-10
'                ElseIf strCode = "#" Then
'                   ' If Not pReadonly Then GenerateBothCodes
'                    bOK = True
'                Else
'                    If IsHashCode(strCode) Then
'                       ' generateISBN13 strCode, False
'                       bOK = True
'                    Else
'                        pMsg = "Invalid ISBN-10 for book"
'                        bOK = False
'                    End If
'                End If
            End If
        Else                                                    'Neither code is supplied
            LoadNew = False
            Exit Function
    End If
returnresult:
    bCodesOK = bOK
    pCode = strCode
    pEAN = strEAN
    LoadNew = bCodesOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.LoadNew(pCODE,pEAN,pFBook)", Array(pCode, pEAN, pFBook)
End Function

Public Function IStoBeGenerated() As Boolean
    On Error GoTo errHandler
    IStoBeGenerated = (strCode = "#")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.IStoBeGenerated"
End Function
Public Function IsISBNEx(pCode As String) As Boolean
    On Error GoTo errHandler
Dim strISBN As String
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String
    IsISBNEx = True
    
    If Len(pCode) = 10 Then
       If Not (UCase(Right(pCode, 1)) = "X" Or IsNumeric(Right(pCode, 1))) Then
           IsISBNEx = False
           GoTo EXIT_Handler
       End If
       If Left(pCode, 9) <= "0" Then
           IsISBNEx = False
           GoTo EXIT_Handler
       End If
       strISBN = Left(Right(pCode, 10), 9)
       x = 0
       For i = 1 To 9
           x = x + (val(Mid(pCode, i, 1))) * Abs(i - 11)
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
       If Not (UCase(Right(pCode, 1)) = Trim(strChk)) Then
          IsISBNEx = False
       End If
    Else
       IsISBNEx = False
    End If
EXIT_Handler:
    Exit Function
errHandler:
    MsgBox Error
    Exit Function
ErrrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.IsISBNEx(pCODE)", pCode
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
    itest = val(Left(strCode, 1))
    If itest >= 0 And itest <= 7 Then
      iGroupLength = 1
    Else
      itest = val(Left(strCode, 2))
      If itest >= 80 And itest <= 94 Then
        iGroupLength = 2
      Else
        itest = val(Left(strCode, 3))
        If itest >= 950 And itest <= 995 Then
          iGroupLength = 3
        Else
          itest = val(Left(strCode, 4))
          If itest >= 9960 And itest <= 9989 Then
            iGroupLength = 4
          Else
            itest = val(Left(strCode, 5))
            If itest >= 99900 And itest <= 99999 Then
              iGroupLength = 5
            Else
              FormatCode = strCode
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
    itest = val(Left(strRem, 2))
    If itest >= 0 And itest <= 19 Then
      iPublisherLength = 2
    Else
      itest = val(Left(strRem, 3))
      If itest >= 200 And itest <= 699 Then
        iPublisherLength = 3
      Else
        itest = val(Left(strRem, 4))
        If itest >= 7000 And itest <= 8499 Then
          iPublisherLength = 4
        Else
          itest = val(Left(strRem, 5))
          If itest >= 85000 And itest <= 89999 Then
            iPublisherLength = 5
          Else
            itest = val(Left(strRem, 6))
            If itest >= 900000 And itest <= 949999 Then
              iPublisherLength = 6
            Else
              itest = val(Left(strRem, 7))
              If itest >= 9500000 And itest <= 9999999 Then
                iPublisherLength = 7
              Else
                FormatCode = "OR"
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.FormatCode(pForEXport)", pForEXport
End Function
'Public Property Get CanHandle() As Boolean
'    On Error GoTo ErrHandler
'    If IsISBN10 Or IsISBN13 Or IsHashCode Then
'        CanHandle = True
'    Else
'        CanHandle = False
'    End If
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_ProdCode.CanHandle"
'End Property
'Private Function ValidEAN(pIN As String) As Boolean
'    On Error GoTo ErrHandler
'
'    bIsEAN = True
'    If Len(pIN) <> 13 Then
'        bIsEAN = False
'    ElseIf Not IsNumeric(pIN) Then
'        bIsEAN = False
'    ElseIf Modulo_10(Left(pIN, 12)) <> Right(pIN, 1) Then
'        bIsEAN = False
'    End If
'    ValidEAN = bIsEAN
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_ProdCode.ValidEAN(pIn)", pIN
'End Function
'Private Function ValidEANEx(pIN As String, pIsBook As Boolean, pIsLocal As Boolean) As Boolean
'    bIsEAN = True
'    If Len(pIN) <> 13 Then
'        bIsEAN = False
'    ElseIf Not IsNumeric(pIN) Then
'        bIsEAN = False
'    ElseIf Modulo_10(Left(pIN, 12)) <> Right(pIN, 1) Then
'        bIsEAN = False
'    End If
'    If bIsEAN Then
'        If pIsBook And Left(pIN, 3) <> "978" And Left(pIN, 3) <> "979" Then
'            bIsEAN = False
'        End If
'    End If
'    If pIsLocal Then
'        If Not (Left(pIN, 1) = "2") Then
'            bIsEAN = False
'        End If
'    End If
'
'    ValidEANEx = bIsEAN
'
'End Function
'Public Property Get IsEAN() As Boolean
'    bIsEAN = True
'    If Len(strEAn) <> 13 Then
'        bIsEAN = False
'    ElseIf Not IsNumeric(strEAn) Then
'        bIsEAN = False
'    ElseIf Modulo_10(Left(strEAn, 12)) <> Right(strEAn, 1) Then
'        bIsEAN = False
'    End If
'
'    IsEAN = bIsEAN
'End Property
'Public Property Get IsCode() As Boolean
'    On Error GoTo ErrHandler
'    IsCode = True
'    If Len(strCode) = 10 And IsNumeric(Left(strCode, 9)) Then
'        If Not IsISBN Then
'            IsCode = False
'        End If
'    Else
'        If Left(strRaw, 1) <> "#" Or Len(strRaw) = 1 Then
'                IsCode = False
'        Else
'            If Len(strRaw) > 13 Then
'                IsCode = False
'            End If
'        End If
'    End If
'
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_ProdCode.IsCode"
'End Property
'Private Function ValidISBN(pCODE As String)
'    On Error GoTo ErrHandler
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_ProdCode.ValidISBN(pCODE)", pCODE
'End Function
Public Function CreateCode()
    On Error GoTo errHandler
Dim strTemp As String
    strTemp = Trim$(str$(GetNextNumber(23)))
    strCode = "#" & strTemp
    strTemp = "210" & strTemp
    strEAN = strTemp & String(12 - Len(strTemp), "0")
    strEAN = strEAN & Modulo_10(strEAN)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.CreateCode"
End Function
Public Function CreateCodes(pGivenCode As String, pGivenEAN As String, pProductType As String, pResult As Integer)
    On Error GoTo errHandler
Dim strTempCode As String
Dim strTempCodeS As String
Dim strTempEAN As String
Dim strTempEANS As String
Dim strNo As String

    CreateCodes = 99
    pResult = 0
    strTempCode = pGivenCode
    If Len(strTempCode) = 10 And Left(strTempCode, 1) <> "#" Then
        strTempCodeS = Left(strTempCode, Len(strTempCode) - 1)
    Else
        strTempCodeS = ""
    End If
    strTempEAN = pGivenEAN
    If Len(strTempEAN) <> 0 Then
        If IsISBN13(strTempEAN) = False Then  'Must be standard EAN, not locally generated EAN
            pResult = 99
            Exit Function
        End If
    End If
    
    If strTempCode > "" And strTempEAN > "" Then   'Both are present
        If Len(strTempCode) = 10 And Left(strTempCode, 1) <> "#" And IsNumeric(strTempCodeS) Then    'we have an ISBN
            If CheckMatch(strTempEAN, strTempCode) Then
                strCode = strTempCode
                strEAN = strTempEAN
            Else
                pResult = 97
            End If
        Else                            'we have a local code and an EAN
            If pProductType = "B" And (Left(strTempEAN, 3) = "978") And Left(strTempCode, 1) <> "#" Then
                strCode = ISBN13to10(strTempEAN)
            Else
                strCode = strTempCode
            End If
            strEAN = strTempEAN
        End If
    ElseIf strTempEAN > "" Then                    'Only EAN is present
        If pProductType = "B" And Left(strTempEAN, 3) = "978" Or Left(strTempEAN, 3) = "979" Then  'a book
            strCode = ISBN13to10(strTempEAN)
        Else
            strCode = Trim$(str$(GetNextNumber(23)))   '230 is for all non-book products where the EAN is known
            strCode = "#" & strCode
           ' strCode = "23" & strCode
        End If
        strEAN = strTempEAN
    ElseIf strTempCode > "" Then                   'Only Code is present
        If pProductType = "B" Then
            If IsISBNEx(strTempCode) Then
                strCode = strTempCode
                strEAN = Trim$(str$(GetNextNumber(23)))   '240 is for all book products where the ISBN is known
                'strEAN = & strEAN
                strEAN = "224" & String(9 - Len(strEAN), "0") & strEAN
                strEAN = strEAN & Modulo_10(strEAN)
            Else
                strCode = strTempCode
                strEAN = Trim$(str$(GetNextNumber(23)))   '250 is for all book products where no ISBN is known
            '    strEAN = "225" & strEAN
                strEAN = "225" & String(9 - Len(strEAN), "0") & strEAN
                strEAN = strEAN & Modulo_10(strEAN)
            End If
        Else
            If IsISBNEx(strTempCode) Then   'not a book
                pResult = 95
            Else
                strCode = strTempCode
                strEAN = Trim$(str$(GetNextNumber(23)))   '260 is for all non-book products where the EAN is NOT known
             '   strEAN = "226" & strEAN
                strEAN = "226" & String(9 - Len(strEAN), "0") & strEAN
                strEAN = strEAN & Modulo_10(strEAN)
            End If
        End If
    Else                                           'Neither is present
        If pProductType = "B" Then
            strNo = Trim$(str$(GetNextNumber(23)))   '250 is for all book products where no ISBN is known
            strEAN = "250" & String(9 - Len(strNo), "0") & strNo
            strEAN = strEAN & Modulo_10(strEAN)
            strCode = "#" & strNo
        Else
            strNo = Trim$(str$(GetNextNumber(23)))   '260 is for all non-book products where the EAN is NOT known
            strEAN = "260" & String(9 - Len(strNo), "0") & strNo
            strEAN = strEAN & Modulo_10(strEAN)
            strCode = "#" & strNo
        End If
    
    End If
    CreateCodes = 0
    Exit Function
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.CreateCodes(pGivenCode,pGivenEAN,pProductType,pResult)", Array(pGivenCode, _
         pGivenEAN, pProductType, pResult)
End Function
Private Function CheckMatch(pEAN As String, pCode As String) As Boolean
    On Error GoTo errHandler
    CheckMatch = True
    If Left(pEAN, 3) = "978" Or Left(pEAN, 3) = "979" Then
        If Not (Left(Right(pEAN, 10), 9) = Left(pCode, 9)) Then
            CheckMatch = False
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.CheckMatch(pEAN,pCODE)", Array(pEAN, pCode)
End Function
Public Function SetCodesForBook(Optional code, Optional EAN, Optional ISBN) As Boolean
    On Error GoTo errHandler

Dim strTemp
Dim strTemp2
Dim strStrippedCode

    SetCodesForBook = False
    If Not IsMissing(code) Then
        If Left(code, 1) = "#" Then
            strCode = code
        Else
            strCode = "#" & code
        End If
        strStrippedCode = Right(strCode, Len(strCode) - 1)
        If IsNumeric(strStrippedCode) Then
            strTemp = "221" & "000" & String(6 - Len(strStrippedCode), "0") & strStrippedCode
        Else
            strTemp2 = Trim$(str$(GetNextNumber(23)))
            strTemp = "221" & "000" & String(6 - Len(strTemp2), "0") & strTemp2
        End If
        strEAN = strTemp & Modulo_10(strTemp)
        SetCodesForBook = True
    End If
    If Not IsMissing(EAN) Then
        strEAN = EAN
        strCode = ISBN13to10(CStr(EAN))
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
'Err_Handler:
'    MsgBox error
'    Resume EXIT_Handler
'    Resume
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.SetCodesForBook(Code,EAN,ISBN)", Array(code, EAN, ISBN)
End Function
Public Function SetCodesForProduct(Optional code, Optional EAN) As Boolean
    On Error GoTo errHandler
Dim strTemp
Dim strTemp2
Dim strStrippedCode
        
    SetCodesForProduct = False

    If Not IsMissing(code) Then
        If Left(code, 1) = "#" Then
            strCode = code
        Else
            strCode = "#" & code
        End If
        strStrippedCode = Right(strCode, Len(strCode) - 1)
      '  If IsNumeric(strStrippedCode) Then
      '      strTemp = "222" & "000" & String(6 - Len(strStrippedCode), "0") & Right(strStrippedCode, 6)
      '  Else
            strTemp2 = Trim$(str$(GetNextNumber(23)))
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.SetCodesForProduct(Code,EAN)", Array(code, EAN)
End Function

Public Property Get code() As String
    code = strCode
End Property
Public Property Get Serial() As String
    Serial = strSerial
End Property
Public Property Get ISBN13() As String
    ISBN13 = strEAN
End Property
'Public Property Get ISBNExists() As Boolean
'    ISBNExists = (Len(strCode) = 10 And IsISBN) Or ((Len(strCode) = 13) And IsISBN)
'End Property
'Public Property Get PrivateCodeExists() As Boolean
'    If Len(strRaw) < 10 And Len(strRaw) > 0 Then
'        PrivateCodeExists = True
'    End If
'End Property
Public Property Get SerialExists() As Boolean

    SerialExists = False
    If IsNumeric(strSerial) Then
        If CLng(strSerial) > 0 Then
            SerialExists = True
        End If
    End If
        
End Property
Public Function IsPrivateCode() As Boolean
    IsPrivateCode = Len(strCode) > 2 And Len(strCode) < 10 And Left(strCode, 1) <> "#" And Left(strCode, 1) <> "/"
End Function

Public Function GetNextNumber(pType As enCodeType) As Long
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim strDefaultCo As String
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_GetNextProdNo"
    cmd.commandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@Type", adInteger, adParamInput, , pType)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@Number", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    cmd.execute
    GetNextNumber = cmd.Parameters("@Number").Value
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
End Function
'
'Private Function IsISBN10(pIn As String) As Boolean
'    IsISBN10 = (Len(pIn) = 10 And IsISBN)
'End Function
Private Function ISBN13toISBN10(pIn As String) As String
Dim strISBN As String
Dim i As Integer
Dim x As Integer
Dim iMod As Integer
Dim strChk As String
    
    If Left(pIn, 3) = "978" Then
        If Len(pIn) = 13 Then      'the EAN code is attached
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
            ISBN13toISBN10 = strISBN & Trim(strChk)
        End If
    Else
        ISBN13toISBN10 = ""
    End If

EXIT_Handler:
        Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ProdCode.ISBN13toISBN10(pIn)", pIn
End Function

Private Function IsValidLocalCode(pIn As String) As Boolean
    If IsPrivateCode() Or Left(pIn, 1) = "#" Then
        IsValidLocalCode = True
    Else
        IsValidLocalCode = False
    End If
End Function
'Private Function GENERATEISBN10(pIn As String) As Boolean
'    If Left(pIn, 3) = "978" Then
'        strCode = ClearEAN(pIn)
'        GENERATEISBN10 = True
'    Else
'        GENERATEISBN10 = False
'    End If
'End Function
Private Function generateISBN13(pIn As String, pIsBook As Boolean) As Boolean
Dim s As String
Dim s2 As String
    If pIsBook Then
        s = pIn
        s = "978" & Left(s, 9)
        If IsNumeric(s) Then
            strEAN = s & Modulo_10(s)
            generateISBN13 = True
        End If
    Else
        If Left(pIn, 1) = "#" Then
            s = Right(pIn, Len(pIn) - 1)
        End If
        s2 = Trim$(str$(GetNextNumber(23)))
        s = oPC.CodePrefixForInternalBarcodesEAN & String(12 - Len(oPC.CodePrefixForInternalBarcodesEAN) - Len(s2), "0") & s2
        
        strEAN = s & Modulo_10(s)
        generateISBN13 = True
    End If
End Function
Private Function GenerateBothCodes()
Dim s As String

        s = Trim$(str$(GetNextNumber(23)))
       ' strEAn = "230" & "000" & String(6 - Len(s), "0") & s
        strEAN = oPC.CodePrefixForInternalBarcodesEAN & "000" & String(6 - Len(s), "0") & s
        strEAN = strEAN & Modulo_10(strEAN)
        
        strCode = "#" & s
        GenerateBothCodes = True
End Function
