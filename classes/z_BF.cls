VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_BF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim iERr As Integer

Sub LoadProdRec()
    On Error GoTo errHandler
Dim c As Integer
Dim fMoreTags As Integer
Dim strTag As String
Dim strValue As String
Dim i As Integer
Dim od As d_BICCode
Dim strBICDescription As String
    i = 1
    c = bufResult.temp(1)
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
        strTag = Chr(bufResult.temp(i))
        i = i + 1
        strTag = strTag & Chr(bufResult.temp(i))
        i = i + 2                             'get past the single space
        c = bufResult.temp(i)
  '  MsgBox "pos 2"
        strValue = ""
        Do While c <> 0 And c <> 26
            strValue = strValue & Chr(bufResult.temp(i))
            i = i + 1
            c = bufResult.temp(i)
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.LoadProdRec"
End Sub



Public Function GetRecordByNum(rec_no As Long) As Long
    On Error GoTo errHandler

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
  Erase bufResult.temp
  bufResult.temp(1) = 0
  bufResult.temp(2) = 64
  
  ' Call database engine
  oPC.BF.RunBookFind

  ' Display the status code for the G command
    GetRecordByNum = bufResult.temp(3)
    If bufResult.temp(3) <> 0 Then
        GoTo errHandler
    End If
'ERRH:
'MsgBox "Error in GetRecordByNum: " & Error
'    oError.SetError Err, Error, Now, "Error in GetRecordByNum: ", "", ""
'
'    GetRecordByNum = 99
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.GetRecordByNum(rec_no)", rec_no
End Function



Public Function Search(SearchText As String, ResultsFound As Variant) As Long
    On Error GoTo errHandler
  Dim Text, strTemp As String
  Dim iLen, i As Long
    
 
  Text = "S" & SearchText & Chr(26)
  
  ' Copy ASCII characters to a byte array
  Call oPC.BF.SetRequest(bufRequest, 1, Text)
  
  ' Set the size of the results block to 8 bytes
  Erase bufResult.temp
  bufResult.temp(1) = 8 'abcd'
  bufResult.temp(2) = 0
  
  ' Call database engine
  oPC.BF.RunBookFind
  
  ' Display the status code for the S command
  Search = bufResult.temp(3)
     
  If bufResult.temp(3) = 0 Then
    ResultsFound = CLng(bufResult.temp(4)) + CLng(bufResult.temp(5)) * 256 + CLng(bufResult.temp(6)) * 256 ^ 2 + CLng(bufResult.temp(7)) * 256 ^ 3
  End If
'ERRH:
'MsgBox "ERROR: " & Error
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Search(SearchText,ResultsFound)", Array(SearchText, ResultsFound)
End Function



Function FetchFromBF(pCODE As String) As Long
    On Error GoTo errHandler
'Returns 0 if a record is found; 99 if not found

Dim i, iRecsFound As Long
Dim strTemp As String
Dim strPos As String

    strPos = "pos 1"
    FetchFromBF = 99
    strTemp = "FIND" & Chr(9) & "BN" & Chr(9) & pCODE
    strPos = "pos 2"
    iERr = Search(strTemp, iRecsFound)
    strPos = "pos 3"
    Check iERr = 0, EXC_GENERAL, "Bookfind search string"
    strPos = "pos 4"
    If iRecsFound > 0 Then
    strPos = "pos 5"
        iERr = GetRecordByNum(1)
    strPos = "pos 6"
        Check iERr = 0, EXC_GENERAL, "Bookfind GetRecordByNum"
        FetchFromBF = 0
    strPos = "pos 7"
        LoadProdRec
    strPos = "pos 8"
    End If

FetchFromBF_EXIT:
    Exit Function

errHandler:
    ErrPreserve
    Select Case iERr
        Case 333   'Database not initialized
Init:
            iERr = oPC.BF.Initialize("WORD2605AA000SAB")   'initialize the bookfind database
            If iERr <> 0 Then
               Returnval = MsgBox("BookFind CD has not been prepared for use, load the correct CD, wait 30 seconds, and retry or ignore this issue", vbRetryCancel + vbExclamation, "Warning")
               If Returnval = vbRetry Then
                   GoTo Init:
               Else
                   GoTo FetchFromBF_EXIT
               End If
            End If
            iERr = oPC.BF.OpenDB(oPC.BookFindFacet)  'open the bookfind database
            Check iERr = 0, EXC_GENERAL, "Bookfind OPEN Error"
    Case 249 'Invalid search strategy e.g. invalid Code
        iRecsFound = 0
        Resume Next

    End Select

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.FetchFromBF(pCode as string)", Array(pCODE), , , "strPos", Array(strPos)

End Function




Function InsertFromBookfind()
    On Error GoTo errHandler
Dim strArticle As String
Dim strTitleNet As String

'     This looks for Bookfind data for records on tSS that have no matches on the local database
Dim dteStarted As Date
Dim iResult As Integer
    Dim rs As ADODB.Recordset
    Dim rsSS As ADODB.Recordset
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
        Set rsSS = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM tPRODUCT", oPC.COShort, adOpenDynamic, adLockOptimistic
        rsSS.CursorLocation = adUseClient
        rsSS.Open "SELECT * FROM tSS WHERE SS_PID is NULL", oPC.COShort, adOpenKeyset, adLockOptimistic
    Do While Not rsSS.EOF
        If Not IsISBN10(CStr(rsSS![SS_ISBN])) Then
            GoTo ENDOFLOOP
        End If
        If FetchFromBF(CStr(rsSS![SS_ISBN])) = 0 Then
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
                rs![P_PRODUCTTYPE_ID] = oPC.UnallocatedPT
                rs![P_SPECIALVAT] = 0
                rs![P_VATRATE] = oPC.VATRate
                rs![P_DateLastStockTake] = CDate(0)
                rs![P_QtyLastStockTake] = 0
                iResult = 0
             '  oProdCode.CreateCodes Trim(rs![P_Code]), "", "B", iResult
             '   If iResult = 0 Then
                    rs![P_EAN] = "978" & Left(CStr(rsSS![SS_ISBN]), 9) & Modulo_10(Left("978" & CStr(rsSS![SS_ISBN]), 12))
             '   End If
                rs.Update
            End If
ENDOFLOOP:
        rsSS.MoveNext
        DoEvents
    Loop

    rs.Close
    rsSS.Close
    Set rs = Nothing
    Set rsSS = Nothing
    Exit Function

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BF.INSERTFromBookfind"
End Function

'------------------------------------------------------------------------------------
Public Property Get MainAuthor() As String
    On Error GoTo errHandler
    MainAuthor = Trim$(BFRec.MainAuthor)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.MainAuthor"
End Property
Public Property Get Title() As String
    On Error GoTo errHandler
    Title = Trim$(BFRec.Title)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Title"
End Property
Public Property Get Code() As String
    On Error GoTo errHandler
    Code = Trim$(BFRec.Code)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Code"
End Property
Public Property Get SubTitle() As String
    On Error GoTo errHandler
    SubTitle = Trim$(BFRec.SubTitle)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.SubTitle"
End Property
Public Property Get BindingCode() As String
    On Error GoTo errHandler
    BindingCode = Trim$(BFRec.BindingCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.BindingCode"
End Property
Public Property Get Edition() As String
    On Error GoTo errHandler
    Edition = Trim$(BFRec.Edition)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Edition"
End Property
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = Trim$(BFRec.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Note"
End Property
Public Property Get Availability() As String
    On Error GoTo errHandler
    Availability = Trim$(BFRec.Availability)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Availability"
End Property
Public Property Get PublisherName() As String
    On Error GoTo errHandler
    PublisherName = Trim$(BFRec.PublisherName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.PublisherName"
End Property
Public Property Get UKPrice() As Currency
    On Error GoTo errHandler
    If HasData(Trim$(BFRec.UKPrice)) And Len(Trim$(BFRec.UKPrice)) > 0 Then
        UKPrice = CCur(BFRec.UKPrice)
    Else
        UKPrice = 0
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.UKPrice"
End Property
Public Property Get USPrice() As Currency
    On Error GoTo errHandler
    If HasData(Trim$(BFRec.USPrice)) And Len(Trim$(BFRec.USPrice)) > 0 Then
        USPrice = CCur(BFRec.USPrice)
    Else
        USPrice = 0
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.USPrice"
End Property
Public Property Get LocalPrice() As Currency
    On Error GoTo errHandler
    If HasData(Trim$(BFRec.LocalPrice)) And Len(Trim$(BFRec.LocalPrice)) > 0 Then
        LocalPrice = CCur(Trim(BFRec.LocalPrice))
    Else
        LocalPrice = 0
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.LocalPrice"
End Property
Public Property Get SeriesTitle() As String
    On Error GoTo errHandler
    SeriesTitle = Trim$(BFRec.SeriesTitle)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.SeriesTitle"
End Property
Public Property Get PublicationDate() As String
    On Error GoTo errHandler
    PublicationDate = Trim$(BFRec.PublicationDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.PublicationDate"
End Property
Public Property Get BFClassification() As String
    On Error GoTo errHandler
    BFClassification = Trim$(BFRec.BFClassification)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.BFClassification"
End Property
Public Property Get Description() As String
    On Error GoTo errHandler
    Description = Trim$(BFRec.Description)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.Description"
End Property
Public Property Get DistributorCode() As String
    On Error GoTo errHandler
    DistributorCode = Trim$(BFRec.k6)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.DistributorCode"
End Property
Public Property Get MainSupplierName() As String
    On Error GoTo errHandler
    MainSupplierName = Trim$(BFRec.MainSupplierName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.MainSupplierName"
End Property
Public Property Get SACostPrice() As Currency
    On Error GoTo errHandler
    SACostPrice = Trim$(BFRec.SACostPrice)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_BookFind.SACostPrice"
End Property

