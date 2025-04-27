VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zc_BF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Progress(lngPos As Long, lngMax As Long)
Public Event STATUS(msg As String)

Sub LoadProdRec()
    On Error GoTo ErrHandler
Dim c As Integer
Dim fMoreTags As Integer
Dim strTag As String
Dim strValue As String
Dim i As Integer
Dim od As d_BICCode
Dim strBICDescription As String
    i = 1
    c = bufResult.Temp(1)
    If c <> 26 Then fMoreTags = True
    BFRec.MainAuthor = ""
    BFRec.Availability = ""
    BFRec.Code = ""
    BFRec.Title = ""
    BFRec.SubTitle = ""
    BFRec.BindingCode = ""
    BFRec.Edition = ""
    BFRec.PublisherName = ""
    BFRec.UKPrice = ""
    BFRec.USPrice = ""
    BFRec.LocalPrice = ""
    BFRec.SeriesTitle = ""
    BFRec.PublicationDate = ""
    BFRec.BFClassification = ""
    BFRec.k1 = ""
    BFRec.k6 = ""
    BFRec.Description = ""
    BFRec.MainSupplierName = ""
    BFRec.MaxResults = 0
    BFRec.BindingCode = ""
    BFRec.BFClassification = ""
    BFRec.BICDescription = ""
    BFRec.Note = ""
    i = 4
   ' MsgBox "pos 1"
    Do While fMoreTags = True And i <= 2048               'Handle a tag and text
        strTag = Chr(bufResult.Temp(i))
        i = i + 1
        strTag = strTag & Chr(bufResult.Temp(i))
        i = i + 2                             'get past the single space
        c = bufResult.Temp(i)
  '  MsgBox "pos 2"
        strValue = ""
        Do While c <> 0 And c <> 26
            strValue = strValue & Chr(bufResult.Temp(i))
            i = i + 1
            c = bufResult.Temp(i)
        Loop
        If c = 26 Then fMoreTags = False
        i = i + 1 'Get past the 0
      '  Debug.Print strTag & " : " & strValue
        Select Case strTag
        Case "AU"
            BFRec.MainAuthor = strValue
        Case "AV"
            BFRec.Availability = strValue
        Case "BN"
            BFRec.Code = strValue
        Case "TI"
            BFRec.Title = strValue
        Case "T2"
            BFRec.SubTitle = BFRec.SubTitle & strValue 'Keep t2 and t3 together
        Case "T3"
            BFRec.SubTitle = BFRec.SubTitle & strValue
        Case "BC"
            BFRec.BindingCode = strValue
        Case "DE"
            BFRec.Description = strValue
        Case "EN"
            BFRec.Edition = strValue
        Case "PU"
            BFRec.PublisherName = strValue
        Case "IN"
            BFRec.Note = strValue
        Case "PR"
            BFRec.UKPrice = strValue
        Case "PB"
            BFRec.USPrice = strValue
        Case "R6"
            BFRec.LocalPrice = strValue
        Case "SE"
            BFRec.SeriesTitle = strValue
        Case "K6"
            BFRec.k6 = IIf(Trim(BFRec.k6) > "", Trim(BFRec.k6) & "/" & strValue, strValue)
        Case "PY"
            BFRec.PublicationDate = strValue
        Case "B1"
'            If Not IsNumeric(Left(strValue, 1)) Then
'                BFRec.BFClassification = FNS(BFRec.BFClassification) & IIf(FNS(BFRec.BFClassification) > " ", ",", "") & strValue
'                Set od = oPC.Configuration.BICs.FetchBICByCode(strValue)
'                If Not od Is Nothing Then
'                    strBICDescription = od.Description
'                Else
'                    strBICDescription = ""
'                End If
'                BFRec.BICDescription = FNS(BFRec.BICDescription) & IIf(FNS(BFRec.BICDescription) > "_", ",", "") & strBICDescription
'            End If
        End Select
    Loop
  '  MsgBox "pos 5"
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error & vbCrLf & " IN a_bookfind:LoadProdRec", ",  strTag = " & strTag & "   i = " & i
'    oError.SetError Err, Error, Now(), "a_bookfind:LoadProdRec", ",  strTag = " & strTag & "   i = " & i, ""
'
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.LoadProdRec"
End Sub



Public Function GetRecordByNum(rec_no As Long) As Long
    On Error GoTo ErrHandler

  Dim Text As String
  ' Convert record number into four byte equivalent for G command
  rec_byte1 = 0
  rec_byte2 = 0
  rec_byte3 = 0
  rec_byte4 = 0

  rec_byte1 = rec_no Mod 256
  rec_no_temp = Int(rec_no / 256)

  rec_byte2 = rec_no_temp Mod 256
  rec_no_temp = Int(rec_no / 256 ^ 2)

  rec_byte3 = rec_no_temp Mod 256
  rec_no_temp = Int(rec_no / 256 ^ 3)

  rec_byte4 = rec_no_temp Mod 256
'  LongToBytes Text, 1, rec_no
  ' text contains the G command along with the record number and ^Z as a terminator
  Text = "G" & Chr(rec_byte1) & Chr(rec_byte2) & Chr(rec_byte3) & Chr(rec_byte4) & Chr(26)
  
  'Copy ASCII characters to a byte array
  oPC.BF.SetRequest bufRequest, 1, Text
  
  ' Set the size of the results block to 8K
  Erase bufResult.Temp
  bufResult.Temp(1) = 0
  bufResult.Temp(2) = 64
  
  ' Call database engine
  oPC.BF.RunBookFind

  ' Display the status code for the G command
    GetRecordByNum = bufResult.Temp(3)
    If bufResult.Temp(3) <> 0 Then
        GoTo ErrHandler
    End If
'ERRH:
'MsgBox "Error in GetRecordByNum: " & Error
'    oError.SetError Err, Error, Now, "Error in GetRecordByNum: ", "", ""
'
'    GetRecordByNum = 99
'
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.GetRecordByNum(rec_no)", rec_no
End Function



Public Function Search(SearchText As String, ResultsFound As Variant) As Long
    On Error GoTo ErrHandler
  Dim Text, strTemp As String
  Dim iLen, i As Long
    
 
  Text = "S" & SearchText & Chr(26)
  
  ' Copy ASCII characters to a byte array
  Call oPC.BF.SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 8 bytes
  Erase bufResult.Temp
  bufResult.Temp(1) = 8 'abcd'
  bufResult.Temp(2) = 0
  
  ' Call database engine
  oPC.BF.RunBookFind
  
  ' Display the status code for the S command
  Search = bufResult.Temp(3)
     
  If bufResult.Temp(3) = 0 Then
    ResultsFound = CLng(bufResult.Temp(4)) + CLng(bufResult.Temp(5)) * 256 + CLng(bufResult.Temp(6)) * 256 ^ 2 + CLng(bufResult.Temp(7)) * 256 ^ 3
  End If
'ERRH:
'MsgBox "ERROR: " & Error
'
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Search(SearchText,ResultsFound)", Array(SearchText, ResultsFound)
End Function



Function FetchFromBF(pCode As String) As Long
    On Error GoTo ErrHandler
'Returns 0 if a record is found; 99 if not found
Dim iErr As Long
Dim i, iRecsFound As Long
Dim strTemp As String
Dim strPos As String

    strPos = "pos 1"
    FetchFromBF = 99
   ' strTemp = "FIND" & Chr(9) & "BN" & Chr(9) & pCODE
    If Len(pCode) = 13 Then
        strTemp = "FIND" & Chr(9) & "EA" & Chr(9) & pCode
    Else
        strTemp = "FIND" & Chr(9) & "BN" & Chr(9) & pCode
    End If
    strPos = "pos 2"
    iErr = Search(strTemp, iRecsFound)
    If iErr <> 0 Then
        LogSaveToFile "Bookfind search string: z" & pCode & "z" & strTemp & "z" & strPos & ":" & CStr(iErr)
        GoTo FetchFromBF_EXIT
    End If
    strPos = "pos 4"
    If iRecsFound > 0 Then
        iErr = GetRecordByNum(1)
        Check iErr = 0, EXC_GENERAL, "Bookfind GetRecordByNum"
        FetchFromBF = 0
        LoadProdRec
    strPos = "pos 8"
    End If

FetchFromBF_EXIT:
    Exit Function

ErrHandler:
    ErrPreserve
    Select Case iErr
        Case 333   'Database not initialized
Init:
            iErr = oPC.BF.Initialize("WORD2605AA000SAB")   'initialize the bookfind database
            If iErr <> 0 Then
               Returnval = MsgBox("BookFind CD has not been prepared for use, load the correct CD, wait 30 seconds, and retry or ignore this issue", vbRetryCancel + vbExclamation, "Warning")
               If Returnval = vbRetry Then
                   GoTo Init:
               Else
                   GoTo FetchFromBF_EXIT
               End If
            End If
            iErr = oPC.BF.OpenDB(oPC.BookFindFacet)  'open the bookfind database
            Check iErr = 0, EXC_GENERAL, "Bookfind OPEN Error"
    Case 249 'Invalid search strategy e.g. invalid Code
        iRecsFound = 0
        Resume Next

    End Select

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.FetchFromBF(pCode as string)", Array(pCode), , , "strPos", Array(strPos)

End Function




Function InsertFromBookfind()
    On Error GoTo ErrHandler
Dim strArticle As String
Dim strTitleNet As String

'     This looks for Bookfind data for records on tSS that have no matches on the local database
Dim dteStarted As Date
Dim iresult As Integer
    Dim rs As ADODB.Recordset
    Dim rstmpP As ADODB.Recordset
    Dim oProdCode As New z_ProdCode
    Dim strMsg As String
    Dim StartTime
    Dim x As Long
    Dim iCancelled As Integer
    Dim strShortname As String
    Dim lngProgress As Long
    Dim lngMax As Long
    iCancelled = False
    
        Set rs = New ADODB.Recordset
        Set rstmpP = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM tPRODUCT", oPC.COShort, adOpenDynamic, adLockOptimistic
        rstmpP.CursorLocation = adUseClient
        rstmpP.Open "SELECT * FROM tmpP WHERE PID is NULL", oPC.COShort, adOpenKeyset, adLockOptimistic
    Do While Not rstmpP.EOF
        If Not (IsISBN10(FNS(rstmpP![Code])) Or IsISBN13(FNS(rstmpP![Code]))) Then '(IsISBN13(FNS(rstmpP![Code])) Or IsISBN13(FNS(rstmpP![Code]))) Then
            GoTo ENDOFLOOP
        End If
        If FetchFromBF(FNS(rstmpP![Code])) = 0 Then
            rs.AddNew
                rs![P_MainAuthor] = Left$(FNS(Me.MainAuthor), rs.Fields("P_MainAuthor").DefinedSize)
                rs![P_Code] = Left$(FNS(Me.Code), rs.Fields("P_Code").DefinedSize)
                StripArticle FNS(Left$(FNS(Me.Title), rs.Fields("P_Title").DefinedSize)), strArticle, strTitleNet
                rs!P_Title = strTitleNet
                rs!P_Article = strArticle
                rs![P_SubTitle] = Left$(FNS(Me.SubTitle), rs.Fields("P_SubTitle").DefinedSize)
                rs![P_STATUS] = Left(FNS(Me.Availability), 1) ' rs.Fields("P_SubTitle").DefinedSize)
                rs![P_Bindingcode] = Left$(FNS(Me.BindingCode), rs.Fields("P_Bindingcode").DefinedSize)
                rs![P_Edition] = Left$(FNS(Me.Edition), rs.Fields("P_Edition").DefinedSize)
                rs![P_Publisher] = Left$(FNS(Me.PublisherName), rs.Fields("P_Publisher").DefinedSize)
                rs![P_BFSupplierCode] = Left$(FNS(Me.DistributorCode), rs.Fields("P_BFSupplierCode").DefinedSize)
                strShortname = rs![P_Publisher]


'We are fleshing out the a_Product table, and when there is a skeleton record (Defined as having no Publisher
'name) then we fill in the other fields off Bookfind and then get the matching publisher record and see
'whether it is on the TP table (we match by Publisher name). If not, we add a publisher record with
'ISBNPrefix.
   'Get BF Publisher Name and ISBNPrefix
   'Try to match on papyrus TP_Shortname
   'if not matched then
   '   add new publisher record etc.
   'end if
   'assign new publisher record ID to Title
StartSearch2:

                If IsNumeric(Me.UKPrice) Then
                    rs![P_UKPrice] = CCur(Me.UKPrice)
                Else
                    rs![P_UKPrice] = Null
                End If
                If IsNumeric(Me.LocalPrice) Then
                    If Me.LocalPrice > 0 Then
                        rs![P_RRP] = FNN(Me.LocalPrice * 100)
                    Else
                        rs![P_RRP] = FNN(Me.UKPrice * 1000)
                    End If
                End If
                rs![P_SeriesTitle] = Left$(FNS(Me.SeriesTitle), rs.Fields("P_SeriesTitle").DefinedSize)
        
                rs![P_Pubdate] = Left$(FNS(Me.PublicationDate), rs.Fields("P_Pubdate").DefinedSize)
                
                rs![P_BIC] = Left(Me.BFClassification, rs![P_BIC].DefinedSize)
                rs![P_PRODUCTTYPE] = "B"
                rs![P_PRODUCTTYPE_ID] = 0
                rs![P_SPECIALVAT] = 0
                rs![P_DateLastStockTake] = CDate(0)
                rs![P_QtyLastStockTake] = 0
'                iResult = 0
'                oProdCode.CreateCodes Trim(rs![P_Code]), "", "B", iResult
'                If iResult = 0 Then
'                    rs![P_EAN] = FNS(oProdCode.EAN)
'                End If
                rs.Update
            End If
ENDOFLOOP:
        rstmpP.MoveNext
        DoEvents
    Loop

    rs.Close
    rstmpP.Close
    Set rs = Nothing
    Set rstmpP = Nothing
    Exit Function

ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.UpdateFromBookfind"
End Function

'------------------------------------------------------------------------------------
Public Property Get MainAuthor() As String
    On Error GoTo ErrHandler
    MainAuthor = Trim$(BFRec.MainAuthor)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.MainAuthor"
End Property
Public Property Get Title() As String
    On Error GoTo ErrHandler
    Title = Trim$(BFRec.Title)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Title"
End Property
Public Property Get Code() As String
    On Error GoTo ErrHandler
    Code = Trim$(BFRec.Code)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Code"
End Property
Public Property Get SubTitle() As String
    On Error GoTo ErrHandler
    SubTitle = Trim$(BFRec.SubTitle)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.SubTitle"
End Property
Public Property Get BindingCode() As String
    On Error GoTo ErrHandler
    BindingCode = Trim$(BFRec.BindingCode)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.BindingCode"
End Property
Public Property Get Edition() As String
    On Error GoTo ErrHandler
    Edition = Trim$(BFRec.Edition)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Edition"
End Property
Public Property Get Note() As String
    On Error GoTo ErrHandler
    Note = Trim$(BFRec.Note)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Note"
End Property
Public Property Get Availability() As String
    On Error GoTo ErrHandler
    Availability = Trim$(BFRec.Availability)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Availability"
End Property
Public Property Get PublisherName() As String
    On Error GoTo ErrHandler
    PublisherName = Trim$(BFRec.PublisherName)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.PublisherName"
End Property
Public Property Get UKPrice() As Currency
    On Error GoTo ErrHandler
    If HasData(Trim$(BFRec.UKPrice)) And Len(Trim$(BFRec.UKPrice)) > 0 Then
        UKPrice = CCur(BFRec.UKPrice)
    Else
        UKPrice = 0
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.UKPrice"
End Property
Public Property Get USPrice() As Currency
    On Error GoTo ErrHandler
    If HasData(Trim$(BFRec.USPrice)) And Len(Trim$(BFRec.USPrice)) > 0 Then
        USPrice = CCur(BFRec.USPrice)
    Else
        USPrice = 0
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.USPrice"
End Property
Public Property Get LocalPrice() As Currency
    On Error GoTo ErrHandler
    If HasData(Trim$(BFRec.LocalPrice)) And Len(Trim$(BFRec.LocalPrice)) > 0 Then
        LocalPrice = CCur(Trim(BFRec.LocalPrice))
    Else
        LocalPrice = 0
    End If
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.LocalPrice"
End Property
Public Property Get SeriesTitle() As String
    On Error GoTo ErrHandler
    SeriesTitle = Trim$(BFRec.SeriesTitle)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.SeriesTitle"
End Property
Public Property Get PublicationDate() As String
    On Error GoTo ErrHandler
    PublicationDate = Trim$(BFRec.PublicationDate)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.PublicationDate"
End Property
Public Property Get BFClassification() As String
    On Error GoTo ErrHandler
    BFClassification = Trim$(BFRec.BFClassification)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.BFClassification"
End Property
Public Property Get Description() As String
    On Error GoTo ErrHandler
    Description = Trim$(BFRec.Description)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Description"
End Property
Public Property Get DistributorCode() As String
    On Error GoTo ErrHandler
    DistributorCode = Trim$(BFRec.k6)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.DistributorCode"
End Property
Public Property Get MainSupplierName() As String
    On Error GoTo ErrHandler
    MainSupplierName = Trim$(BFRec.MainSupplierName)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.MainSupplierName"
End Property
Public Property Get SACostPrice() As Currency
    On Error GoTo ErrHandler
    SACostPrice = Trim$(BFRec.SACostPrice)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.SACostPrice"
End Property

Function UpdateFromBookfind(pAuthor As Boolean, _
pTitle As Boolean, _
pSubtitle As Boolean, _
pAvailability As Boolean, _
pBindingcode As Boolean, _
pEdition As Boolean, _
pSUpplierCode As Boolean, _
pPublisherName As Boolean, _
pSeriesTitle As Boolean, _
pPublicationDate As Boolean, _
pUKPrice As Boolean, _
pRRP As Boolean, _
pBIC As Boolean, _
pBookStatus As Boolean, _
Optional pSTAFFID As Long, _
Optional pEAN As String)


    On Error GoTo ErrHandler
Dim strArticle As String
Dim strTitleNet As String

'     This is used when a new CD arrives.
'     It finds all records on tblProduct where there are no details except ISBN and looks on the CD
'     for details, if they are found then the record is updated.
Dim dteStarted As Date
    Dim rs As ADODB.Recordset
Dim strCode As String
    Dim oBF As a_BookFind
    Dim strMsg As String
    Dim StartTime
    Dim x As Long
    Dim iCancelled As Integer
    Dim strShortname As String
    Dim lngProgress As Long
    Dim lngMax As Long
    iCancelled = False
    
        RaiseEvent STATUS("Selecting records to update . . .")
        DoEvents
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
'        If pEAN > "" Then
'            rs.Open "SELECT * FROM vBooksWithISBNs WHERE P_EAN > pEAN", oPC.COShort, adOpenKeyset, adLockOptimistic
'        Else
            rs.Open "SELECT * FROM vBooksWithISBNs ", oPC.COShort, adOpenKeyset, adLockOptimistic
'        End If
'        rs.MoveLast
        lngProgress = 0
        lngMax = rs.RecordCount
        RaiseEvent Progress(lngProgress, lngMax)
        dteStarted = Now()
        RaiseEvent STATUS("Updating from CD . . ." & "Started:" & Format(dteStarted, "dd/mm/yyyy Hh:Nn") & ", Records: " & rs.RecordCount)
        StartTime = Now
        rs.MoveFirst
        Set oBF = New a_BookFind
    Do While Not rs.EOF
        DoEvents
        lngProgress = lngProgress + 1
        RaiseEvent Progress(lngProgress, lngMax)
        If (Not (IsISBN10(FNS(rs![P_Code])) Or IsISBN13(FNS(rs![P_EAN])))) Or Left(rs!P_EAN, 2) = "22" Or Left(rs!P_EAN, 2) = "23" Then
            GoTo ENDOFLOOP
        End If
        strCode = FNS((rs![P_EAN]))
        If strCode = "" Then strCode = CStr(rs![P_Code])
        If oBF.FetchFromBF(strCode) = 0 Then
            If pAuthor Then rs![P_MainAuthor] = Left$(FNS(oBF.MainAuthor), rs.Fields("P_MainAuthor").DefinedSize)
            On Error Resume Next
            rs![P_Code] = Left$(FNS(oBF.Code), rs.Fields("P_Code").DefinedSize)
            rs![P_EAN] = Left$(FNS(oBF.EAN), rs.Fields("P_EAN").DefinedSize)
            On Error GoTo ErrHandler
            If pTitle Then
                StripArticle FNS(Left$(FNS(oBF.Title), rs.Fields("P_Title").DefinedSize)), strArticle, strTitleNet
                rs!P_Title = strTitleNet
                rs!P_Article = strArticle
            End If
            If pSubtitle Then rs![P_SubTitle] = Left$(FNS(oBF.SubTitle), rs.Fields("P_SubTitle").DefinedSize)
            If pAvailability Then rs![P_STATUS] = Left(FNS(oBF.Availability), 1)   ' rs.Fields("P_SubTitle").DefinedSize)
            If pBindingcode Then rs![P_Bindingcode] = Left$(FNS(oBF.BindingCode), rs.Fields("P_Bindingcode").DefinedSize)
            If pEdition Then rs![P_Edition] = Left$(FNS(oBF.Edition), rs.Fields("P_Edition").DefinedSize)
            If pPublisherName Then rs![P_Publisher] = Left$(FNS(oBF.PublisherName), rs.Fields("P_Publisher").DefinedSize)
            If pSUpplierCode Then rs![P_BFSupplierCode] = Left$(FNS(oBF.DistributorCode), rs.Fields("P_BFSupplierCode").DefinedSize)
                strShortname = FNS(rs![P_Publisher])


'We are fleshing out the tProduct table, and when there is a skeleton record (Defined as having no Publisher
'name) then we fill in the other fields off Bookfind and then get the matching publisher record and see
'whether it is on the TP table (we match by Publisher name). If not, we add a publisher record with
'ISBNPrefix.
   'Get BF Publisher Name and ISBNPrefix
   'Try to match on papyrus TP_Shortname
   'if not matched then
   '   add new publisher record etc.
   'end if
   'assign new publisher record ID to Title
StartSearch2:
'       If Not SeekRec("tblTradingPartner", "TP_Shortname", "=", strShortname) Then
'            RetVal = GetMatchingBFPublisher(Trim$(BFRec.k1), ISBNPrefix, strImprintName, strPublisherName, strBFCode) 'returns ISBNPrefix in second param
'            If RetVal = True Then  'found on BookFind''
'
                                            'This publisher is not on database , so add a record to
                                            'tblTradingpartner and tblTradingPartner_Role and tblSupplier_Imprint
'                 RetVal = AddAPublisher(strImprintName, strPublisherName, Trim$(ISBNPrefix), strBFCode, strShortname, 0)
'                 GoTo StartSearch2
'            End If
'        Else
'            rs![P_PublisherID] = gSeekRst![Trading_Partner_ID]
'        End If
                If pUKPrice Then
                    If IsNumeric(oBF.UKPrice) Then
                        rs![P_UKPrice] = CCur(oBF.UKPrice) * oPC.Configuration.DefaultCurrencyDivisor
                    Else
                        rs![P_UKPrice] = Null
                    End If
                    If IsNumeric(oBF.USPrice) Then
                        rs![P_USPrice] = CCur(oBF.USPrice) * oPC.Configuration.DefaultCurrencyDivisor
                    Else
                        rs![P_USPrice] = Null
                    End If
                End If
                If pRRP Then
                    If IsNumeric(oBF.LocalPrice) Then
                        If oBF.LocalPrice > 0 Then
                            rs![P_RRP] = FNN(oBF.LocalPrice * oPC.Configuration.DefaultCurrencyDivisor)
                        Else
                        '    rs![P_RRP] = FNN(oBF.UKPrice / IIf(oPC.Configuration..FindBySysname("GBP") Is Nothing, 10, oPC.Configuration.Currencies.FindBySysname("GBP").Factor) * oPC.Configuration.DefaultCurrency.Divisor)
                        End If
                    End If
                End If
                If pSeriesTitle Then rs![P_SeriesTitle] = Left$(FNS(oBF.SeriesTitle), rs.Fields("P_SeriesTitle").DefinedSize)
                If pPublicationDate Then rs![P_Pubdate] = Left$(FNS(oBF.PublicationDate), rs.Fields("P_Pubdate").DefinedSize)
                If pBIC Then rs![P_BIC] = Left$(FNS(oBF.BFClassification), rs![P_BIC].DefinedSize)
                rs![P_UpdatedFromBookFind] = Date
             '  On Error Resume Next
                rs.Update
              '  On Error GoTo errHandler
            End If
            If pBookStatus Then
                If DateDiff("yyyy", FND(rs!P_LastDateDelivered), Date) > 2 Then 'mark books received last more than 2 years ago and not on Bookfind as Out of Print
                    rs!P_STATUS = "O"
                    rs.Update
                End If
            End If
            
        'exit if operation cancelled
ENDOFLOOP:
      '  rs.CancelUpdate
        rs.MoveNext
        DoEvents
    Loop

    rs.Close
    If pEAN = "" Then
'        oPC.COShort.execute "INSERT INTO tOPERATION (OP_STARTEDAT,OP_ENDEDAT,OP_NOMINALDATE,OP_TYPE,OP_STARTEDBYID,OP_RESULT) VALUES ('" & ReverseDateTime(dteStarted) & "','" & ReverseDateTime(Now) & "','" & ReverseDate(Date) & "'," & BookDataWash & "," & pSTAFFID & ",1)"
        RaiseEvent STATUS("Updating from CD . . . complete.")
        MsgBox "Bookfind wash complete", , "Status"
        RaiseEvent STATUS("")
    End If
    Exit Function

ERRH::
    MsgBox Error
    Exit Function
    Resume
    Exit Function
ErrHandler:
    ErrPreserve
    If Err.Description = "Duplicate CODE" Or Err.Description = "Duplicate EAN" Or Left(Err.Description, 23) = "Cannot insert duplicate" Then
        rs.CancelUpdate
        Resume ENDOFLOOP
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.UpdateFromBookfind(pAuthor,pTitle,pSubtitle,pAvailability,pBindingcode,pEdition," & _
        "pSUpplierCode,pPublisherName,pSeriesTitle,pPublicationDate,pUKPrice,pRRP,pBIC,pBookStatus,pStaffID)", _
         Array(pAuthor, pTitle, pSubtitle, pAvailability, pBindingcode, pEdition, pSUpplierCode, _
         pPublisherName, pSeriesTitle, pPublicationDate, pUKPrice, pRRP, pBIC, pBookStatus, pSTAFFID)
End Function

