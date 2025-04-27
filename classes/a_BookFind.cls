VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_BookFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strCurrentFacet As String
Dim bBookfindConnected As Boolean
Event RecordAdded(Success As Boolean)
Dim ior As Integer
Dim oProduct As a_Product
Sub LongToBytes(a As String, iInsert As Long, LongNum As Long)
10        On Error GoTo errHandler
      ' Procedure to transform a Long integer to 4 bytes
      ' LongNum is the Long integer and iInsert is the position where the byte
      ' begins in the request buffer.
       
20        Call ShiftLongToBytes(a, iInsert, LongNum)
       
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.LongToBytes(A,iInsert,LongNum)", Array(a, iInsert, LongNum)
End Sub

'*******************************************************************************************
' Retrieve record
Public Function GetRecordByNum(rec_no As Long) As Long
10        On Error GoTo errHandler

        Dim Text As String
        ' Convert record number into four byte equivalent for G command
20      rec_byte1 = 0
30      rec_byte2 = 0
40      rec_byte3 = 0
50      rec_byte4 = 0

60      rec_byte1 = rec_no Mod 256
70      rec_no_temp = Int(rec_no / 256)

80      rec_byte2 = rec_no_temp Mod 256
90      rec_no_temp = Int(rec_no / 256 ^ 2)

100     rec_byte3 = rec_no_temp Mod 256
110     rec_no_temp = Int(rec_no / 256 ^ 3)

120     rec_byte4 = rec_no_temp Mod 256
      '  LongToBytes Text, 1, rec_no
        ' text contains the G command along with the record number and ^Z as a terminator
130     Text = "G" & Chr(rec_byte1) & Chr(rec_byte2) & Chr(rec_byte3) & Chr(rec_byte4) & Chr(26)
        
        'Copy ASCII characters to a byte array
140     oPC.BF.SetRequest bufRequest, 1, Text
        
        ' Set the size of the results block to 8K
150     Erase bufResult.Temp
160     bufResult.Temp(1) = 0
170     bufResult.Temp(2) = 64
        
        ' Call database engine
180     oPC.BF.RunBookFind

        ' Display the status code for the G command
190       GetRecordByNum = bufResult.Temp(3)
200       If bufResult.Temp(3) <> 0 Then
210           GoTo errHandler
220       End If
      'H:
      'MsgBox "or in GetRecordByNum: " & or
      '    oor.Setor , or, Now, "or in GetRecordByNum: ", "", ""
      '
      '    GetRecordByNum = 99
      '
230       Exit Function
errHandler:
240       If ErrMustStop Then Debug.Assert False: Resume
250       ErrorIn "a_BookFind.GetRecordByNum(rec_no)", rec_no
End Function

'Friend Function IsOnDataBase(Code As String) As Long
 ' Dim rs As ADODB.Recordset
 ' Dim strSQL As String
 ' Dim cmd As New ADODB.Command
 ' IsOnDataBase = 0
 ' strSQL = "SELECT Product_ID FROM Product WHERE P_Code ='" & Code & "'"
 ' Set rs = New ADODB.Recordset
 ' Set cmd.ActiveConnection = CO
 ' cmd.CommandText = strSQL
 ' With rs
 '   .Open cmd, , adOpenDynamic, adLockReadOnly
 '   If Not .EOF Then IsOnDataBase = !Product_ID
 '   .Close
 ' End With
 ' Set rs = Nothing
 ' Set cmd = Nothing

'End Function

'Public Property Get IsOnDB(Code As String) As Long
'  IsOnDB = IsOnDataBase(Code)
'End Property


Function saveToProduct(pCategoryID As Long) As Long
10        On Error GoTo errHandler

      Dim strShort
      Dim strShortname As String
      Dim iLockCOunt As Integer
      Dim iChoice As Integer
      Dim iRndCount As Integer
      Dim x As Integer
      Dim oProd As a_Product
      Dim oProdCode As z_ProdCode
      Dim strCode As String
      Dim strEAN As String

20        Set oProd = New a_Product
30        oProd.BeginEdit
          
40        oProd.SetEAN FNS(BFRec.EAN)
50        oProd.SetCode FNS(BFRec.Code)
         
60        oProd.SetAuthor Trim$(BFRec.MainAuthor)
70        oProd.SetTitle FNS(BFRec.Title)
80        oProd.SetNote FNS(BFRec.Note)
90        oProd.Availability = Left(FNS(BFRec.Availability), 1)
100       oProd.SetSubTitle FNS(BFRec.SubTitle)
110       oProd.SetEdition FNS(BFRec.Edition)
120       oProd.SetBindingCode FNS(BFRec.BindingCode)
130       oProd.SetUKPrice FNDBL(BFRec.UKPrice)
140       oProd.SetUSPrice FNDBL(BFRec.USPrice)
150       oProd.SetSeriesTitle FNS(BFRec.SeriesTitle)
160       oProd.SetPublicationDate FNS(BFRec.PublicationDate)
170       oProd.SetCategoryID FNN(pCategoryID)
180       oProd.SetPublisher FNS(BFRec.PublisherName)
190       oProd.SetVAT oPC.Configuration.VATRate
200       oProd.ApplyEdit
          

210       RaiseEvent RecordAdded(True)
220       saveToProduct = oProd.PID
230       Set oProd = Nothing
EXIT_Handler:
240       Exit Function
errHandler:
250       If ErrMustStop Then Debug.Assert False: Resume
260       ErrorIn "a_BookFind.saveToProduct(pCategoryID)", pCategoryID
End Function
Function FetchFromBF(pCODE As String) As Long
10        On Error GoTo errHandler
      'Returns 0 if a record is found; 99 if not found

      Dim i, iRecsFound As Long
      Dim iPubRecsFound As Integer
      Dim strDistributor As String
      
      Dim strTemp As String
20        FetchFromBF = 99
30        If oPC.BFLoaded = False Then
40            GoTo FetchFromBF_EXIT
50        End If
60        If Len(pCODE) = 13 Then
70            strTemp = "FIND" & Chr(9) & "EA" & Chr(9) & pCODE
80        Else
90            strTemp = "FIND" & Chr(9) & "BN" & Chr(9) & pCODE
100       End If
110       i = search(strTemp, iRecsFound)
120       Check i = 0, EXC_GENERAL, "Bookfind search string"
130       If iRecsFound > 0 Then
140           i = GetRecordByNum(1)
150           Check i = 0, EXC_GENERAL, "Bookfind GetRecordByNum"
160           FetchFromBF = 0
170           LoadProdRec
'              oPC.BF.CloseDB
'970           oPC.BF.OpenDB oPC.BookFindFacet & "    "
'              lngResult = oPC.BF.DefineOutputFormat("PUBLISHER")
'              lngResult = oPC.BF.Search("FIND" & Chr$(9) & "L2" & Chr$(9) & strDistributor, iPubRecsFound)
'              lngResult = GetRecordByNum(1)
'              LoadDistrRec 1
'            '  BFRec.Distributor = ""
180       End If

FetchFromBF_EXIT:
190       Exit Function

errHandler:
200       Select Case i
              Case 333   'Database not initialized
Init:
210               i = oPC.BF.Initialize("WORD2605AA000SAB")   'initialize the bookfind database
220               If i <> 0 Then
230                  Returnval = MsgBox("BookFind CD has not been prepared for use, load the correct CD, wait 30 seconds, and retry or ignore this issue", vbRetryCancel + vbExclamation, "Warning")
240                  If Returnval = vbRetry Then
             GoTo Init:
250                  Else
260                      GoTo FetchFromBF_EXIT
270                  End If
280               End If
290               i = oPC.BF.OpenDB("WEBK")  'open the bookfind database
300               Check i = 0, EXC_GENERAL, "Bookfind OPEN or"
310       Case 249 'Invalid search strategy e.g. invalid Code
320           iRecsFound = 0
330           Resume Next
340       Case Else
350           MsgBox "Unknown Bookfind or"
360       End Select
370       Resume FetchFromBF_EXIT
380       Exit Function
End Function

'Sub GetRecord(Num As Long)
'On Error GoTo _GetRecord
'    ChDrive "C:"
'    ChDir "C:\Bookfind"
'
'    strRecord = ""
'    i = d_Get(Num, strRecord)
'    If i <> 0 Then
'        MsgBox "Bookfind or: or No:" & i & Chr(10) & "Can't get Record No: " & Num
'    End If
'
'EXIT_GetRecord:
'    Exit Sub
'
'_GetRecord:
'    MsgBox "Bookfind or in dGET"
'    Resume EXIT_GetRecord
'
'End Sub
Public Property Get MainAuthor() As String
10        On Error GoTo errHandler
20        MainAuthor = Trim$(BFRec.MainAuthor)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.MainAuthor"
End Property
Public Property Get Title() As String
10        On Error GoTo errHandler
20        Title = Trim$(BFRec.Title)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Title"
End Property
Public Property Get Code() As String
10        On Error GoTo errHandler
20        Code = Trim$(BFRec.Code)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Code"
End Property
Public Property Get EAN() As String
10        On Error GoTo errHandler
20        EAN = Trim$(BFRec.EAN)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.EAN"
End Property
Public Property Get SubTitle() As String
10        On Error GoTo errHandler
20        SubTitle = Trim$(BFRec.SubTitle)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.SubTitle"
End Property
Public Property Get BindingCode() As String
10        On Error GoTo errHandler
20        BindingCode = Trim$(BFRec.BindingCode)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.BindingCode"
End Property
Public Property Get Weight() As String
10        On Error GoTo errHandler
20        Weight = Trim$(BFRec.Weight)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Weight"
End Property

Public Property Get Edition() As String
10        On Error GoTo errHandler
20        Edition = Trim$(BFRec.Edition)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Edition"
End Property
Public Property Get Note() As String
10        On Error GoTo errHandler
20        Note = Trim$(BFRec.Note)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Note"
End Property
Public Property Get Availability() As String
10        On Error GoTo errHandler
20        Availability = Trim$(BFRec.Availability)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Availability"
End Property
Public Property Get PublisherName() As String
10        On Error GoTo errHandler
20        PublisherName = Trim$(BFRec.PublisherName)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.PublisherName"
End Property
Public Property Get UKPrice() As Currency
      Dim tmp As Currency
10        On Error GoTo errHandler
          
20        UKPrice = 0
30        If FNS(BFRec.UKPrice) > "" Then
40            If ConvertToCurr(BFRec.UKPrice, tmp) Then
50                UKPrice = tmp
60            End If
70        End If
80        Exit Property
errHandler:
90        If ErrMustStop Then Debug.Assert False: Resume
100       ErrorIn "a_BookFind.UKPrice"
End Property
Public Property Get USPrice() As Currency
      Dim tmp As Currency
10        On Error GoTo errHandler
20        USPrice = 0
30        If FNS(BFRec.USPrice) > "" Then
40            If ConvertToCurr(BFRec.USPrice, tmp) Then
50                USPrice = tmp
60            End If
70        End If
80        Exit Property
errHandler:
90        If ErrMustStop Then Debug.Assert False: Resume
100       ErrorIn "a_BookFind.USPrice"
End Property
Public Property Get LocalPrice() As Currency
      Dim tmp As Currency
10        On Error GoTo errHandler
20        LocalPrice = 0
30        If FNS(BFRec.LocalPrice) > "" Then
40            If ConvertToCurr(BFRec.LocalPrice, tmp) Then
50                LocalPrice = tmp
60            End If
70        End If
80        Exit Property
errHandler:
90        If ErrMustStop Then Debug.Assert False: Resume
100       ErrorIn "a_BookFind.LocalPrice"
End Property
Public Property Get SeriesTitle() As String
10        On Error GoTo errHandler
20        SeriesTitle = Trim$(BFRec.SeriesTitle)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.SeriesTitle"
End Property
Public Property Get PublicationDate() As String
10        On Error GoTo errHandler
20        PublicationDate = Trim$(BFRec.PublicationDate)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.PublicationDate"
End Property
Public Property Get BFClassification() As String
10        On Error GoTo errHandler
20        BFClassification = Trim$(BFRec.BFClassification)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.BFClassification"
End Property
Public Property Get Description() As String
10        On Error GoTo errHandler
20        Description = Trim$(BFRec.Description)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.Description"
End Property
Public Property Get DistributorCode() As String
10        On Error GoTo errHandler
20        DistributorCode = Trim$(BFRec.k6)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorCode"
End Property
Public Property Get MainSupplierName() As String
10        On Error GoTo errHandler
20        MainSupplierName = Trim$(BFRec.MainSupplierName)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.MainSupplierName"
End Property
Public Property Get SACostPrice() As Currency
10        On Error GoTo errHandler
20        SACostPrice = Trim$(BFRec.SACostPrice)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.SACostPrice"
End Property

Sub LoadProdRec()
10        On Error GoTo errHandler
      Dim c As Integer
      Dim fMoreTags As Integer
      Dim strTag As String
      Dim strValue As String
      Dim i As Integer
      Dim oD As d_BICCode
      Dim strBICDescription As String
20        i = 1
30        c = bufResult.Temp(1)
40        If c <> 26 Then fMoreTags = True
50        BFRec.MainAuthor = ""
60        BFRec.Availability = ""
70        BFRec.Code = ""
80        BFRec.EAN = ""
90        BFRec.Title = ""
100       BFRec.SubTitle = ""
110       BFRec.BindingCode = ""
120       BFRec.Edition = ""
130       BFRec.PublisherName = ""
140       BFRec.UKPrice = ""
150       BFRec.USPrice = ""
160       BFRec.LocalPrice = ""
170       BFRec.SeriesTitle = ""
180       BFRec.PublicationDate = ""
190       BFRec.BFClassification = ""
200       BFRec.k1 = ""
210       BFRec.k6 = ""
220       BFRec.Description = ""
230       BFRec.MainSupplierName = ""
240       BFRec.MaxResults = 0
250       BFRec.BindingCode = ""
260       BFRec.BFClassification = ""
270       BFRec.BICDescription = ""
280       BFRec.Note = ""
290       BFRec.Weight = ""
300       i = 4
         ' MsgBox "pos 1"
310       Do While fMoreTags = True And i <= 2048               'Handle a tag and text
320           strTag = Chr(bufResult.Temp(i))
330           i = i + 1
340           strTag = strTag & Chr(bufResult.Temp(i))
350           i = i + 2                             'get past the single space
360           c = bufResult.Temp(i)
        '  MsgBox "pos 2"
370           strValue = ""
380           Do While c <> 0 And c <> 26
390               strValue = strValue & Chr(bufResult.Temp(i))
400               i = i + 1
410               c = bufResult.Temp(i)
420           Loop
430           If c = 26 Then fMoreTags = False
440           i = i + 1 'Get past the 0
            '  Debug.Print strTag & " : " & strValue
450           Select Case strTag
              Case "AU"
460               BFRec.MainAuthor = strValue
470           Case "AV"
480               BFRec.Availability = strValue
490           Case "BN"
500               BFRec.Code = strValue
510           Case "EA"
520               BFRec.EAN = strValue
530           Case "TI"
540               BFRec.Title = strValue
550           Case "T2"
560               BFRec.SubTitle = FNS(BFRec.SubTitle) & strValue 'Keep t2 and t3 together
570           Case "T3"
580               BFRec.SubTitle = FNS(BFRec.SubTitle) & strValue
590           Case "BC"
600               BFRec.BindingCode = strValue
610           Case "DE"
620               BFRec.Description = strValue
630           Case "EN"
640               BFRec.Edition = strValue
650           Case "PU"
660               BFRec.PublisherName = strValue
670           Case "IN"
680               BFRec.Note = strValue
690           Case "PR"
700               BFRec.UKPrice = strValue
710           Case "PB"
720               BFRec.USPrice = strValue
730           Case "R6"
740               BFRec.LocalPrice = strValue
750           Case "SE"
760               BFRec.SeriesTitle = strValue
770           Case "K6"
780               BFRec.k6 = IIf(Trim(BFRec.k6) > "", Trim(BFRec.k6) & "/" & strValue, strValue)
790           Case "PY"
800               BFRec.PublicationDate = strValue
810           Case "WT"
820               BFRec.Weight = strValue
830           Case "B1"
840               If Not IsNumeric(Left(strValue, 1)) Then
850                   BFRec.BFClassification = FNS(BFRec.BFClassification) & IIf(FNS(BFRec.BFClassification) > " ", ",", "") & strValue
860                   Set oD = oPC.Configuration.BICs.FetchBICByCode(strValue)
870                   If Not oD Is Nothing Then
880                       strBICDescription = oD.Description
890                   Else
900                       strBICDescription = ""
910                   End If
920                   BFRec.BICDescription = FNS(BFRec.BICDescription) & IIf(FNS(BFRec.BICDescription) > "_", ",", "") & strBICDescription
930               End If
940           End Select
950       Loop
EXIT_Handler:
960       Exit Sub
errHandler:
970       If ErrMustStop Then Debug.Assert False: Resume
980       ErrorIn "a_BookFind.LoadProdRec"
End Sub

Sub LoadDistrRec(pIdx As Integer)
10        On Error GoTo errHandler
      Dim c As Integer
      Dim fMoreTags As Integer
      Dim strTag As String
      Dim strValue As String
      Dim i As Integer
      Dim oD As d_BICCode
      Dim strBICDescription As String
20        i = 1
30        c = bufResult.Temp(1)
40        If c <> 26 Then fMoreTags = True
50        DistributorRec.DistributorAddress = ""
60        DistributorRec.DistributorCode = ""
70        DistributorRec.DistributorEMail = ""
80        DistributorRec.DistributorFax = ""
90        DistributorRec.DistributorName = ""
100       DistributorRec.DistributorTel = ""
110       i = 4
120       Do While fMoreTags = True And i <= 2048               'Handle a tag and text
130           strTag = Chr(bufResult.Temp(i))
140           i = i + 1
150           strTag = strTag & Chr(bufResult.Temp(i))
160           i = i + 2                             'get past the single space
170           c = bufResult.Temp(i)
180           strValue = ""
190           Do While c <> 0 And c <> 26
200               strValue = strValue & Chr(bufResult.Temp(i))
210               i = i + 1
220               c = bufResult.Temp(i)
230           Loop
240           If c = 26 Then fMoreTags = False
250           i = i + 1 'Get past the 0
            '  Debug.Print strTag & " : " & strValue
260           Select Case strTag
              Case "AD"
270               DistributorRec.DistributorAddress = strValue
280           Case "L2"
290               DistributorRec.DistributorCode = strValue
300           Case "EM"
310               DistributorRec.DistributorEMail = strValue
320           Case "FX"
330               DistributorRec.DistributorFax = strValue
340           Case "DN"
350               DistributorRec.DistributorName = strValue
360           Case "TL"
370               DistributorRec.DistributorTel = strValue
380           End Select
390       Loop
EXIT_Handler:
      'Err_Handler:
      '    MsgBox error
400       Exit Sub
errHandler:
410       If ErrMustStop Then Debug.Assert False: Resume
420       ErrorIn "a_BookFind.LoadDistrRec(pIdx)", pIdx
End Sub


Public Function SearchByTag(SearchText As String, EFTag As String) As Long
10        On Error GoTo errHandler
      Dim Text As String
20      SearchByTag = 0
        ' text contains the S command along with the EF_ tag and search terms entered and ^Z as a terminator
30      Text = "SFIND" & Chr(9) & EFTag & Chr(9) & SearchText & Chr(26)
        
        ' Copy ASCII characters to a byte array
40      Call oPC.BF.SetRequest(bufRequest, 1, Text)
        
        ' Set the size of the results block to 8 bytes
50      Erase bufResult.Temp
60      bufResult.Temp(1) = 8 'abcd'
70      bufResult.Temp(2) = 0
        
        ' Call database engine
80      oPC.BF.RunBookFind
        
        ' Display the status code for the S command
90      i = bufResult.Temp(3)
           
100     If bufResult.Temp(3) = 0 Then
110       SearchByTag = bufResult.Temp(4) + bufResult.Temp(5) * 256 + bufResult.Temp(6) * 256 ^ 2 + bufResult.Temp(7) * 256 ^ 3
120     End If

130       Exit Function
errHandler:
140       If ErrMustStop Then Debug.Assert False: Resume
150       ErrorIn "a_BookFind.SearchByTag(SearchText,EFTag)", Array(SearchText, EFTag)
End Function

Public Function search(SearchText As String, ResultsFound As Variant) As Long
10        On Error GoTo errHandler
        Dim Text, strTemp As String
        Dim iLen, i As Long
          
       
20      Text = "S" & SearchText & Chr(26)
        
        ' Copy ASCII characters to a byte array
30      Call oPC.BF.SetRequest(bufRequest, 1, Text)
        
        ' Set the size of the results block to 8 bytes
40      Erase bufResult.Temp
50      bufResult.Temp(1) = 8 'abcd'
60      bufResult.Temp(2) = 0
        
        ' Call database engine
70      oPC.BF.RunBookFind
        
        ' Display the status code for the S command
80      search = bufResult.Temp(3)
           
90      If bufResult.Temp(3) = 0 Then
100       ResultsFound = CLng(bufResult.Temp(4)) + CLng(bufResult.Temp(5)) * 256 + CLng(bufResult.Temp(6)) * 256 ^ 2 + CLng(bufResult.Temp(7)) * 256 ^ 3
110     End If
      'H:
      'MsgBox "OR: " & or
      '
120       Exit Function
errHandler:
130       If ErrMustStop Then Debug.Assert False: Resume
140       ErrorIn "a_BookFind.Search(SearchText,ResultsFound)", Array(SearchText, ResultsFound)
End Function


Public Function FetchRecord(RetreiveTags As String, RecordNum As Long) As Long
10        On Error GoTo errHandler
        Dim i, rec_no As Long
        Dim strTemp, Text As String
        
        ' text contains the F command and the first EF_ tag
20      strTemp = RetreiveTags
30      i = InStr(strTemp, ",")
40      Text = "F"
50      Do While i > 0
60        Text = Text & Mid(strTemp, 1, i - 1) & Chr(0)
70        strTemp = Right(strTemp, Len(strTemp) - i)
80        i = InStr(strTemp, ",")
90      Loop
        ' cop off the last 0 char
100     Text = Left(Text, Len(Text) - 1)
        
        ' text contains the F command along with the EF_ tag(s) and ^Z as a terminator
110     Text = Text & Chr(26)
        
        ' Copy ASCII characters to a byte array
120     Call oPC.BF.SetRequest(bufRequest, 1, Text)
        
        ' Set the size of the results block to 4 bytes
130     Erase bufResult.Temp
140     bufResult.Temp(1) = 4
150     bufResult.Temp(2) = 0
        
        ' Call database engine
160     oPC.BF.RunBookFind
        
        ' Display the status code for the F command
170     FetchRecord = bufResult.Temp(3)
180     If bufResult.Temp(3) <> 0 Then
190       Exit Function
200     End If
        
        ' Call record required for display
        'Call rec_required(rec_no)
        
        ' Display data as a result of the database search
      '  DisplayRecord (Text)

210       Exit Function
errHandler:
220       If ErrMustStop Then Debug.Assert False: Resume
230       ErrorIn "a_BookFind.FetchRecord(RetreiveTags,RecordNum)", Array(RetreiveTags, RecordNum)
End Function


Public Function DestroyResultsList() As Long
10        On Error GoTo errHandler
        Dim Text As String
        
        ' text contains the A command, 0 Char and ^Z as a terminator
20      Text = "A" & "0" & Chr(26)
        
        ' Copy ASCII characters to a byte array
30      Call oPC.BF.SetRequest(bufRequest, 1, Text)
        
        ' Set the size of the results block to 1024 bytes
40      Erase bufResult.Temp
50      bufResult.Temp(1) = 0
60      bufResult.Temp(2) = 4
        
        ' Call database engine
70      oPC.BF.RunBookFind
        
        'Return Status Code
80      DestroyResultsList = bufResult.Temp(3)
          
90        Exit Function
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "a_BookFind.DestroyResultsList"
End Function

' Generate request byte array for transfer to DLL
Public Property Get k6() As String
10        On Error GoTo errHandler
20      k6 = Trim$(BFRec.k6)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.k6"
End Property

Public Property Get k1() As String
10        On Error GoTo errHandler
20      k1 = Trim$(BFRec.k1)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.k1"
End Property

Public Property Get MaxResults() As Long
10        On Error GoTo errHandler
20      MaxResults = BFRec.MaxResults
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.MaxResults"
End Property

Public Property Let MaxResults(Value As Long)
10        On Error GoTo errHandler
20      BFRec.MaxResults = Value
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.MaxResults(Value)", Value
End Property

Public Function UpdateBICCodes()
10        On Error GoTo errHandler
      Dim rs As New ADODB.Recordset
      Dim strDescription As String
      Dim oD As d_BICCode
      Dim cmd As ADODB.Command

      Dim OpenResult As Integer
      '-------------------------------
20        OpenResult = oPC.OpenDBSHort
      '-------------------------------
30        Set cmd = New ADODB.Command
40        cmd.CommandText = "UPDATE tPRODUCT SET P_BIC = ''"
50        cmd.CommandType = adCmdText
60        cmd.ActiveConnection = oPC.COShort
70        cmd.CommandTimeout = 0
80        cmd.execute
90        Set cmd = Nothing
100       rs.open "Select P_Code,P_BIC,P_EAN From tProduct WHERE len(P_CODE) = " & CStr(ISBN_LENGTH), oPC.COShort, adOpenKeyset, adLockOptimistic
110       Do While Not rs.eof
      ' lookup each title on bookfind
      '        If rs.Fields("P_Code") = "0195707206" Then
      '            MsgBox "Here"
      '        End If
120           lngResult = FetchFromBF(FNS(rs.fields("P_EAN")))
130           If lngResult <> 0 Then
140               lngResult = FetchFromBF(FNS(rs.fields("P_Code")))
150           End If
      ' update the BIC codes
160           If lngResult = 0 Then
170               rs.fields("P_BIC") = FNS(BFRec.BFClassification)
      ' Save records
180               rs.Update
190           End If
200           rs.MoveNext
210       Loop
220       rs.Close
230       Set rs = Nothing
          
      '    Set cmd = New ADODB.Command
      '    cmd.CommandText = "UpdateMainBIC"
      '    cmd.CommandType = adCmdStoredProc
      '    cmd.ActiveConnection = oPC.COShort
      '    cmd.CommandTimeout = 0
      '    cmd.execute
      '    Set cmd = Nothing
      '---------------------------------------------------
240       If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
250       Exit Function
errHandler:
260       If ErrMustStop Then Debug.Assert False: Resume
270       ErrorIn "a_BookFind.UpdateBICCodes"
End Function

Function InitializeScan(pIndex As String, pTerm As String, iResultFound As Variant)
10        On Error GoTo errHandler
        Dim Text As String
        
          ' text contains the A command, 0 Char and ^Z as a terminator
20        Text = "I" & pIndex & pTerm & Chr(26)
          
          ' Copy ASCII characters to a byte array
30        Call oPC.BF.SetRequest(bufRequest, 1, Text)
          
          ' Set the size of the results block to 1024 bytes
40        Erase bufResult.Temp
50        bufResult.Temp(1) = 0
60        bufResult.Temp(2) = 4
          
          ' Call database engine
70        oPC.BF.RunBookFind
          
          'Return Status Code
80        InitializeScan = bufResult.Temp(3)
90        iResultFound = bufResult.Temp(4) + bufResult.Temp(5) * 256 + bufResult.Temp(6) * 256 ^ 2 + bufResult.Temp(7) * 256 ^ 3
100       Exit Function
errHandler:
110       If ErrMustStop Then Debug.Assert False: Resume
120       ErrorIn "a_BookFind.InitializeScan(pIndex,pTerm,iResultFound)", Array(pIndex, pTerm, _
               iResultFound)
End Function
Function GetNextTerm() As String
10        On Error GoTo errHandler
        Dim Text As String
        Dim x As Long
        Dim strTemp As String
        Dim strResult As String
        ' text contains the A command, 0 Char and ^Z as a terminator
20      Text = "N" & Chr(26)
        
        ' Copy ASCII characters to a byte array
30      Call oPC.BF.SetRequest(bufRequest, 1, Text)
        
        ' Set the size of the results block to 1024 bytes
40      Erase bufResult.Temp
50      bufResult.Temp(1) = 0
60      bufResult.Temp(2) = 4
        
        ' Call database engine
70      oPC.BF.RunBookFind
        
        'Return Status Code
80      strResult = bufResult.Temp(3)
90      If strResult <> 0 Then
100       GetNextTerm = ""
110       Exit Function
120     End If
130     strTemp = ""
140     x = 9
150        Do
160           strTemp = strTemp & Chr(bufResult.Temp(x))
170           x = x + 1
180        Loop Until bufResult.Temp(x) = 26
190       GetNextTerm = strTemp
      'H:
      '    MsgBox error
      '    Exit Function
200       Exit Function
errHandler:
210       If ErrMustStop Then Debug.Assert False: Resume
220       ErrorIn "a_BookFind.GetNextTerm"
End Function
Public Property Get DistributorName() As String
10        On Error GoTo errHandler
20        DistributorName = Trim$(DistributorRec.DistributorName)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorName"
End Property
Public Property Get DistributorAddress() As String
10        On Error GoTo errHandler
20        DistributorAddress = Trim$(DistributorRec.DistributorAddress)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorAddress"
End Property
Public Property Get DistributorTelephone() As String
10        On Error GoTo errHandler
20        DistributorTelephone = Trim$(DistributorRec.DistributorTel)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorTelephone"
End Property
Public Property Get DistributorFax() As String
10        On Error GoTo errHandler
20        DistributorFax = Trim$(DistributorRec.DistributorFax)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorFax"
End Property
Public Property Get DistributorEMail() As String
10        On Error GoTo errHandler
20        DistributorEMail = Trim$(DistributorRec.DistributorEMail)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorEMail"
End Property
Public Property Get DistributorID() As String
10        On Error GoTo errHandler
20        DistributorID = Trim$(DistributorRec.DistributorCode)
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_BookFind.DistributorID"
End Property

Function FetchFromBFEx(pArg As String, pRecsFound As Long, pMax As Long) As Long
10        On Error GoTo errHandler
Dim SQL As String
      Dim i As Long
      Dim Res As Long
      Dim iRecsFound As Long
      Dim iPubRecsFound As Long
      Dim strTemp As String
      Dim lngLimit As Long
      Dim strTmp As String
      Dim strDS As String
      Dim EOCriteria As Boolean
      Dim strArg As String
      Dim tmp As String
      Dim ar() As String
      Dim iWordCount As Integer
      Dim EOFCriteria As Boolean
      Dim strTitleCrit As String
      Dim strPubCrit As String
      Dim strAuthorCrit As String
      Dim iStartTitle As Long
      Dim iStartAuthor As Long
      Dim iStartPub As Long
      Dim rs As ADODB.Recordset
      Dim iPos As Integer
      Dim bISBN As Boolean
      Dim bEAN As Boolean

20        iPos = 1
30        If Len(pArg) = 10 Then
40            If IsNumeric(Left(pArg, 9)) Then
50                bISBN = True
60            End If
70        End If
80        If Len(pArg) = 13 And IsNumeric(pArg) Then
90            bEAN = True
100       End If
110       If Not bISBN And Not bEAN Then
120           If InStr(pArg, "/") = 0 Then
130               pArg = "/" & pArg
140           End If
150           pArg = Replace(pArg, "///", "#")
160           pArg = Replace(pArg, "//", "~")
170           pArg = Replace(pArg, "*", "")
              
180           iStartTitle = InStr(1, pArg, "/")
190           iStartAuthor = InStr(1, pArg, "~")
200           iStartPub = InStr(1, pArg, "#")
          
              'Check for Title search
210           If iStartTitle > 0 Then
220               strTitleCrit = FNS(Mid(pArg, iStartTitle + 1, IIf(GetMin(iStartAuthor, iStartPub, True) > iStartTitle + 1, GetMin(iStartAuthor, iStartPub, True) - iStartTitle - 1, 999)))
230           End If
240           If iStartAuthor > 0 Then
250               strAuthorCrit = FNS(Mid(pArg, iStartAuthor + 1, IIf(GetMin(iStartTitle, iStartPub, True) > iStartAuthor + 1, GetMin(iStartTitle, iStartPub, True) - iStartAuthor - 1, 999)))
260           End If
270           If iStartPub > 0 Then
280               strPubCrit = FNS(Mid(pArg, iStartPub + 1, IIf(GetMin(iStartAuthor, iStartTitle, True) > iStartPub + 1, GetMin(iStartAuthor, iStartTitle, True) - iStartPub - 1, 999)))
290           End If

300           iPos = 2
310           If strTitleCrit > "" Then
320               ar = Split(FNS(strTitleCrit), " ")
330               iWordCount = UBound(ar) + 1
340               tmp = ""
350               For i = 1 To iWordCount
360                   If Len(tmp) > 0 Then tmp = tmp & "*" & "AND" & "^"
370                   tmp = tmp & "TI" & "^" & ar(i - 1)
380               Next
390           End If
400           If strAuthorCrit > "" Then
410               ar = Split(FNS(strAuthorCrit), " ")
420               iWordCount = UBound(ar) + 1
430               For i = 1 To iWordCount
440                   If Len(tmp) > 0 Then tmp = tmp & "*" & "AND" & "^"
450                   tmp = tmp & "AA" & "^" & ar(i - 1)
460               Next
470           End If
480           If strPubCrit > "" Then
490               ar = Split(FNS(strPubCrit), " ")
500               iWordCount = UBound(ar) + 1
510               For i = 1 To iWordCount
520                   If Len(tmp) > 0 Then tmp = tmp & "*" & "AND" & "^"
530                   tmp = tmp & "PU" & "^" & ar(i - 1)
540               Next
550           End If
              ' we use visible characters until this point for debugging purposes
560               iPos = 3
570           tmp = Replace(tmp, "^", Chr(9))
580           tmp = Replace(tmp, "*", Chr(0))
              
590           FetchFromBFEx = 99
600           If oPC.BFLoaded = False Then
610               GoTo FetchFromBF_EXIT
620           End If
630               iPos = 4
                  
                  
                  
                  
640           strTemp = "FIND" & Chr(9) & tmp
650           i = search(strTemp, iRecsFound)
              
660       Else
670           If Len(pArg) = 13 Then
680               strTemp = "FIND" & Chr(9) & "EA" & Chr(9) & pArg
690           Else
700               strTemp = "FIND" & Chr(9) & "BN" & Chr(9) & pArg
710           End If
720           i = search(strTemp, iRecsFound)
730       End If
740       Check i = 0, EXC_GENERAL, "Bookfind or in search function"
750           iPos = 5
760       FetchFromBFEx = 0
770       lngLimit = GetMin(iRecsFound, pMax)
780       For i = 1 To lngLimit
          
790           Res = GetRecordByNum(i)
800               iPos = 51
810           If Res <> 0 Then
820               FetchFromBFEx = 999
830               pRecsFound = i - 1
840               GoTo FetchFromBF_EXIT
850           Else
860                   iPos = 52
870              LoadProdRec
880               BFRec.Title = Replace(BFRec.Title, "'", "''")
890               BFRec.MainAuthor = Replace(BFRec.MainAuthor, "'", "''")
900               BFRec.PublisherName = Replace(BFRec.PublisherName, "'", "''")
               '   On Error Resume Next
910                iPos = 53
920                SQL = "INSERT INTO search" & oPC.WorkstationID & "(Distributor,DistributorCode,P_Status,P_CODE,P_EAN,P_TITLE,P_MainAuthor,P_Publisher,LocalPrice,UKPrice,USPrice) " _
                      & "VALUES ('" & "','" & FNS(BFRec.k6) & "','" & FNS(BFRec.Availability) & "','" & FNS(BFRec.Code) & "','" & FNS(BFRec.EAN) & "','" & FNS(BFRec.Title) & "','" _
                      & FNS(BFRec.MainAuthor) & "','" & FNS(BFRec.PublisherName) & "','" _
                      & CDbl(IIf(FNS(BFRec.LocalPrice) = "", 0, FNS(BFRec.LocalPrice))) * 100 & "','" _
                      & CDbl(IIf(FNS(BFRec.UKPrice) = "", 0, FNS(BFRec.UKPrice))) * 100 & "','" _
                      & CDbl(IIf(FNS(BFRec.USPrice) = "", 0, FNS(BFRec.USPrice))) * 100 & "')"
930               oPC.COShort.execute SQL
940           End If
950       Next
          '''''
960                           iPos = 6
970           oPC.BF.CloseDB
980           oPC.BF.OpenDB oPC.BookFindFacet & "    "
990           lngResult = oPC.BF.DefineOutputFormat("PUBLISHER")
1000          Set rs = New ADODB.Recordset
1010          rs.open "Select * FROM search" & oPC.WorkstationID, oPC.COShort, adOpenKeyset, adLockOptimistic
1020                                  iPos = 7
1030          Do While Not rs.eof
1040                                  iPos = 71
1050              If Trim(rs!DistributorCode) > "" Then
                     ' lngResult = oPC.BF.Search("FIND" & Chr$(9) & "L2" & Chr$(9) & Trim(rs!Distributor), iPubRecsFound)
1060                 lngResult = oPC.BF.search("FIND" & Chr$(9) & "L2" & Chr$(9) & Trim(rs!DistributorCode), iPubRecsFound)
1070                  rs!Distributor = ""
1080                                  iPos = 72
1090                  For i = 1 To iPubRecsFound
1100                                  iPos = 721
1110                      lngResult = GetRecordByNum(i)
1120                                  iPos = 722
1130                      LoadDistrRec 1
1140                                  iPos = 723
1150                      rs!Distributor = Left(IIf(rs!Distributor > "", rs!Distributor & " / " & FNS(DistributorRec.DistributorName), FNS(DistributorRec.DistributorName)), 60) 'Left(rs!Distributor & Left(FNS(DistributorRec.DistributorName), 60), 60)
1160                                  iPos = 724
1170                  Next
1180                  rs.Update
1190                                  iPos = 73
1200              End If
1210              rs.MoveNext
1220          Loop
1230                                  iPos = 8
1240          rs.Close
1250          oPC.BF.CloseDB
1260          oPC.BF.OpenDB
              
1270          lngResult = oPC.BF.DefineOutputFormat("MAIN")
              
1280          iPos = 9
1290      pRecsFound = iRecsFound
          
FetchFromBF_EXIT:
1300      Exit Function

1310      Exit Function
errHandler:
1320      If ErrMustStop Then Debug.Assert False: Resume
1330      ErrorIn "a_BookFind.FetchFromBFEx(pArg,pRecsFound,pMax)", Array(pArg, pRecsFound, pMax)
End Function


