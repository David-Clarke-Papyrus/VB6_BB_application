VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Catalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim wb As Word.Application
Dim doc As Word.Document
Dim ExclConn As ADODB.Connection
Dim rs As ADODB.Recordset
'Dim oCurrMan As z_CurrencyManager
Dim oCurrency As a_Currency

Dim strLastSection As String
Dim strLastAuthor As String
Dim strLastCat As String
Dim strCurrentCat As String
Dim strCurrentSection As String
Dim strFilename As String
Dim iFilenum As Integer
Dim lngCurrID As Long
Dim strCurrFormat As String

Public Event Status(i As Long)
Public Event MaxRecs(i As Long)

Dim lngMax As Long
Dim lngCount As Long
Dim flgEOF As Boolean

Public Sub SetCurrency(PID As Long)
    On Error GoTo ErrHandler

    Set oCurrency = oPC.Configuration.Currencies.FindCurrencyByID(PID)
    strCurrFormat = oCurrency.FormatString
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.SetCurrency(PID)", PID
End Sub
Private Sub DoOneBookOrSection()
    On Error GoTo ErrHandler

    If strCurrentSection <> "" Then
        DoOneSection
    Else
        InsertTitle
        RaiseEvent Status(lngCount)
        rs.MoveNext
        SetCurrentIndicators
    End If
    

EXIT_Handler:
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.DoOneBookOrSection"
End Sub
Private Sub SetCurrentIndicators()
    On Error GoTo ErrHandler
    
    flgEOF = rs.EOF
    If flgEOF Then GoTo EXIT_SetCurrentIndicators
    
    strCurrentCat = FNS(rs![ParentHeadingDescription])
    strCurrentSection = FNS(rs![heading])

    If FNN(rs![CATHEAD_Parent]) = 0 Then
        strCurrentCat = FNS(rs![heading])
        strCurrentSection = ""
    End If

EXIT_SetCurrentIndicators:
    Exit Sub
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.SetCurrentIndicators"
End Sub
Private Sub DoOneCategory()
    On Error GoTo ErrHandler
Dim iSameHeading

'Print category heading
    wb.Selection.Collapse Direction:=wdCollapseEnd
'    Selection.InsertAfter rs![heading]
    wb.Selection.InsertAfter strCurrentCat
    wb.Selection.Style = "Bookhead"
    wb.Selection.InsertParagraphAfter
    wb.Selection.Collapse Direction:=wdCollapseEnd
'Print Book Or Section
    strLastCat = strCurrentCat
    Do Until flgEOF = True Or strCurrentCat <> strLastCat
        DoOneBookOrSection
    Loop

EXIT_Handler:
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.DoOneCategory"
End Sub

Private Sub DoOneSection()
    On Error GoTo ErrHandler

    'Print section heading
    wb.Selection.Collapse Direction:=wdCollapseEnd
    wb.Selection.InsertAfter rs![heading]
    wb.Selection.Style = "Bookhead"
    wb.Selection.InsertParagraphAfter
    
    strLastSection = rs![heading]
    Do Until flgEOF Or strCurrentCat <> strLastCat Or strCurrentSection <> strLastSection
        InsertTitle
        
        rs.MoveNext
        SetCurrentIndicators
    Loop

    'Print end of section line
        wb.Selection.InsertAfter "____________"
        wb.Selection.Style = "SectionEnd"
        wb.Selection.MoveUp Unit:=wdParagraph, Count:=2
        With wb.Selection.ParagraphFormat
            .WidowControl = True
            .KeepWithNext = True
        End With
        wb.Selection.MoveDown Unit:=wdParagraph, Count:=2

        wb.Selection.InsertParagraphAfter

EXIT_Handler:
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.DoOneSection"
End Sub

Public Sub ExportToWORD(ByRef ptemplate As String, ByRef pFilename As String, ByRef pUseHeadings As Boolean, pVisible As Boolean)
    On Error GoTo ErrHandler
Dim strConn As String
Dim i As Long
Dim lResult As Long
Dim oUtil As z_UTIL
Dim strSQL As String


    strSQL = "SELECT  sect, sectauthor, sectauthor2, underscores, brackets, P_MainAuthor, heading, P_CODE, " _
            & " P_Title, P_Article,P_Description, P_Comment, P_Summary, P_PubPlace, PI_Serial, PI_Description, " _
            & " PI_PurchaseDate, PI_Cost, PI_Comment, P_SubTitle, P_Pubdate, PI_Price, price, P_Price1, P_Publisher, " _
            & " P_Edition, P_FlagText, P_CatalogueHEadingID, CATHEAD_Parent, ParentHeadingDescription " _
            & " From TEMPCatalogue " _
            & " ORDER BY sect,sectauthor,sectauthor2,underscores, brackets DESC," _
            & " price DESC , P_Price1 DESC;"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.CO, adOpenForwardOnly
    i = 0
    If rs.EOF Then GoTo EXIT_Handler
    Do While Not rs.EOF
        rs.MoveNext
        i = i + 1
    Loop
    lngMax = i
    RaiseEvent MaxRecs(lngMax)
    rs.MoveFirst
    lngCount = 0
    Set oUtil = Nothing
    RaiseEvent Status(0)
    rs.MoveFirst

    Set wb = CreateObject("Word.application")
    wb.Visible = pVisible
    
    Set doc = wb.Documents.Add(ptemplate)
    wb.ActiveWindow.View = wdNormalView
    wb.ActiveDocument.ShowGrammaticalErrors = False
    wb.ActiveDocument.ShowSpellingErrors = False
    strFilename = pFilename
    doc.Paragraphs(1).Style = "books"
    wb.DisplayAlerts = wdAlertsNone
    wb.Options.Pagination = False
    strLastSection = ""
    strLastCat = ""
    strLastAuthor = ""
    
    flgEOF = False
    If pUseHeadings Then
        SetCurrentIndicators
        Do Until flgEOF
              DoOneCategory
        Loop
    Else
        Do Until rs.EOF
            InsertTitle
            RaiseEvent Status(lngCount)
            rs.MoveNext
        Loop
    End If
    
    doc.SaveAs strFilename

    wb.Quit
    
EXIT_Handler:

    rs.Close
    Set wb = Nothing
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.ExportToWORD(ptemplate,pFilename,pUseHeadings,pVisible)", Array(ptemplate, _
         pFilename, pUseHeadings, pVisible)
End Sub

Private Sub InsertTitle()
    On Error GoTo ErrHandler
Dim tmp, tmp2, strComment1, strComment2, strPubdate
Dim vPrice As Long
Dim strFormat As String
Dim strPubPlace, strEdition, strSummary, strCondition
Dim strSubtitle, strTitle, strAuthor, strFlagtext
Dim lngRoundTo As Long

' If FNS(rs.Fields("P_Title")) = "liberatory vision" Then MsgBox "Here"
    strComment1 = rs.Fields("PI_Comment")
    strComment1 = FNS(strComment1)
    strComment2 = FNS(rs.Fields("P_Comment"))
    strComment2 = FNS(strComment2)
    vPrice = rs.Fields("Price")
    vPrice = FNN(vPrice)
    strPubdate = FNS(rs.Fields("P_Pubdate"))
    strPubdate = FNS(strPubdate)
    strPubPlace = FNS(rs.Fields("P_PubPlace"))
    strPubPlace = FNS(strPubPlace)
    strEdition = rs.Fields("P_Edition")
    strEdition = FNS(strEdition)
    strSummary = rs.Fields("P_Summary")
    strSummary = FNS(strSummary)
    strCondition = FNS(rs.Fields("PI_Description"))
    If strCondition = "" Then strCondition = FNS(rs.Fields("P_Description"))
 '   strCondition = FNS(strCondition)
 If InStr(1, FNS(rs.Fields("P_Title")), "MODERN PALIMPSEST") > 0 Then MsgBox "here"
    strTitle = FNS(rs.Fields("P_Article")) & IIf(Len(FNS(rs.Fields("P_Article"))) > 0, " ", "") & FNS(rs.Fields("P_Title"))
   ' strTitle = FNS(strTitle)
    strSubtitle = rs.Fields("P_SubTitle")
    strSubtitle = FNS(strSubtitle)
    strAuthor = rs.Fields("P_MainAuthor")
    strAuthor = FNS(strAuthor)
    strFlagtext = rs.Fields("P_FlagText")
    strFlagtext = FNS(strFlagtext)
    
    If Len(strFlagtext) > 0 Then
        wb.Selection.Paragraphs(1).Range.Select
        wb.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
        wb.Selection.Font.Bold = True
        wb.Selection.InsertAfter strFlagtext
        wb.Selection.Style = "flags"
        wb.Selection.InsertParagraphAfter
    End If
    wb.Selection.Collapse Direction:=wdCollapseEnd
    If strAuthor = strLastAuthor Then
        wb.Selection.InsertAfter "__________"
    Else
        If strAuthor = "" Then
            wb.Selection.InsertAfter "__________"
            strLastAuthor = ""
        Else
            wb.Selection.InsertAfter strAuthor
            strLastAuthor = strAuthor
        End If
    End If
    wb.Selection.Style = "Books"
    wb.Selection.Font.Bold = True
    wb.Selection.Collapse Direction:=wdCollapseEnd
    wb.Selection.InsertAfter " " & UCase(strTitle) & IIf(Len(strSubtitle) = 0, "", ", " & strSubtitle)
    wb.Selection.Font.Bold = False
    


    wb.Selection.Collapse Direction:=wdCollapseEnd
    
    If Not strCondition = "" Then
        wb.Selection.InsertAfter ", " & strCondition
        wb.Selection.Font.Italic = True
        wb.Selection.Collapse Direction:=wdCollapseEnd
    End If
'    If Not strSummary = "" Then
'        wb.Selection.InsertAfter ", " & strSummary
'        wb.Selection.Font.Italic = True
'        wb.Selection.Collapse Direction:=wdCollapseEnd
'    End If
    If Not strEdition = "" Then
        wb.Selection.InsertAfter ", " & strEdition
        wb.Selection.Font.Italic = False
        wb.Selection.Collapse Direction:=wdCollapseEnd
    End If
    If Not strPubPlace = "" Then
        wb.Selection.InsertAfter ", " & strPubPlace
        wb.Selection.Font.Italic = False
        wb.Selection.Collapse Direction:=wdCollapseEnd
    End If
    If Not strPubdate = "" Then
        wb.Selection.InsertAfter ", " & strPubdate & "."
        wb.Selection.Font.Italic = False
        wb.Selection.Collapse Direction:=wdCollapseEnd
    End If
    If Not vPrice = 0 Then
       ' wb.Selection.InsertAfter Chr$(9) & Format(vPrice, oCurrency.FormatString)
        
   strFormat = oCurrency.FormatString
   strFormat = Replace(strFormat, ".00", "")
        
        vPrice = vPrice * oCurrency.Factor
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(vPrice)
        vPrice = (RoundUp(vPrice, lngRoundTo))
        wb.Selection.InsertAfter Chr$(9) & Format(vPrice / oCurrency.Divisor, strFormat)
        
        
        
'''''
'                    Set oCurr = oPC.Configuration.Currencies.FindByDescription("U.S. DOLLAR")
'                    ConvertedPrice = adjustedPrice * oCurr.Factor
'                    lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(ConvertedPrice)  ' / oCurr.Divisor
'                    ConvertedPrice = (RoundUp(ConvertedPrice, lngRoundTo))
'                    rsOUt.Fields("PriceDollar") = Format((ConvertedPrice) / oCurr.Divisor, oCurr.FormatString)
'''''
        
        wb.Selection.Font.Bold = True
        wb.Selection.Font.Italic = False
        wb.Selection.Collapse Direction:=wdCollapseEnd
    End If
    If (Not strComment1 = "") Or (Not strComment2 = "") Then
            With wb.Selection.ParagraphFormat
                .SpaceAfter = 0
                .WidowControl = True
                .KeepWithNext = True
            End With
    End If
    
    If Not strComment1 = "" Then
        wb.Selection.Collapse Direction:=wdCollapseEnd
        wb.Selection.InsertParagraphAfter
        wb.Selection.Collapse Direction:=wdCollapseEnd
        wb.Selection.InsertAfter CStr(strComment1)
        wb.Selection.Style = "Bookhead2"
        
        wb.Selection.InsertParagraphAfter

    End If

    If Not strComment2 = "" Then
        wb.Selection.Collapse Direction:=wdCollapseEnd
        wb.Selection.InsertParagraphAfter
        wb.Selection.Collapse Direction:=wdCollapseEnd
        wb.Selection.InsertAfter CStr(strComment2) & "."
        wb.Selection.Style = "Bookhead2"
        wb.Selection.InsertParagraphAfter
    End If
    wb.Selection.Collapse Direction:=wdCollapseEnd
    wb.Selection.InsertParagraphAfter
    wb.Selection.Font.Size = 12
    wb.Selection.Collapse Direction:=wdCollapseEnd
    lngCount = lngCount + 1
    If lngCount Mod 10 = 0 Then
        doc.SaveAs strFilename
    End If
    DoEvents

EXIT_Handler:
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.InsertTitle"
End Sub
Private Function TrimAuthor(pIn)
    On Error GoTo ErrHandler

Dim strOut

    If Left(pIn, 1) = "[" Then
        strOut = Right(pIn, Len(pIn) - 1)
        TrimAuthor = Left(strOut, Len(strOut) - 1)
    Else
        TrimAuthor = pIn
    End If

EXIT_Handler:
    Exit Function

    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.TrimAuthor(pIn)", pIn
End Function
Public Sub PrepareData(plngCATID As Long)
    On Error GoTo ErrHandler
Dim strSQL As String
Dim strResult As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.CO
    cmd.CommandText = "DropTableIfExists"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@pName", adVarChar, adParamInput, 30)
    par.Value = "tCataloglist"
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@Result", adVarChar, adParamOutput, 30)
    cmd.Parameters.Append par
    cmd.CommandTimeout = 0
    cmd.Execute
    If cmd.Parameters(0) <> 0 Then
        MsgBox "Cannot delete" & cmd.Parameters(1)
    End If
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.CO
    cmd.CommandText = "DropTableIfExists"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@pName", adVarChar, adParamInput, 30)
    par.Value = "TEMPCatalogue"
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@Result", adVarChar, adParamOutput, 30)
    cmd.Parameters.Append par
    cmd.CommandTimeout = 0
    cmd.Execute
    If cmd.Parameters(0) <> 0 Then
        MsgBox "Cannot delete" & cmd.Parameters(1)
    End If
   
    
    oPC.CO.Execute "CREATE TABLE tCataloglist ([PID] UNIQUEIDENTIFIER,[PIID] INT,[HEAD] INT);"
    oPC.CO.Execute "CREATE TABLE TEMPCatalogue (Sect varChar(250),SectAuthor varChar(250),SectAuthor2 varChar(250), " _
                & "Underscores INT,Brackets INT,P_MainAuthor varChar(250),Heading varChar(250),P_Code TEXT,P_Title TEXT,P_Article TEXT," _
                & "P_Description TEXT,P_Comment TEXT,P_Summary TEXT,P_PubPlace TEXT,PI_Serial INT,PI_Description TEXT," _
                & "PI_PurchaseDATE DateTime,PI_Cost Money,PI_Comment TEXT,P_Subtitle TEXT,P_PUBDATE TEXT," _
                & "PI_Price Money,Price Money,P_Price1 money,P_Publisher TEXT,P_Edition TEXT,P_Flagtext TEXT," _
                & "P_CatalogueHeadingID INT,CATHEAD_Parent INT,ParentHeadingDescription TEXT)"
    
    strSQL = "INSERT INTO tCataloglist ( PID, PIID ) " _
            & " SELECT DISTINCT P_ID AS Expr1, PI_ID " _
            & " FROM tProduct JOIN tPI ON P_ID = PI_P_ID JOIN tCatalPI on CATALPI_PI_ID = PI_ID JOIN tCatalogue on CATALPI_CATAL_ID = CATAL_ID " _
            & " WHERE CATALPI_CATAL_ID= " & plngCATID
    oPC.CO.Execute strSQL
    strSQL = "INSERT INTO tCataloglist ( PID) " _
            & " SELECT DISTINCT P_ID AS Expr1 " _
            & " FROM tProduct JOIN tCatalP on CATALP_P_ID = P_ID JOIN tCatalogue on CATALP_CATAL_ID = CATAL_ID " _
            & " WHERE CATALP_CATAL_ID= " & plngCATID
    oPC.CO.Execute strSQL
    

    oPC.CO.Execute "GENCAT"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM TEMPCATALOGUE", oPC.CO, adOpenForwardOnly, adLockOptimistic
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Catalog.PrepareData(plngCATID)", plngCATID
End Sub


