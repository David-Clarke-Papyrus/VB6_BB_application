VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SearchEngineS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public mbDebug As Boolean

Dim strSQL As String
Dim strcriteria As String
Dim lngretval As Long
Dim colobjs As Collection
Dim strstock As String
Dim intstock As Integer
Dim intWithCopies As Integer
Dim rs As ADODB.Recordset
Dim lngrowcount As Long
Dim inttemptbl As Integer
Dim oSQL As New z_SQL
Dim intsearchcount As Integer
Private colDisplay As Collection
'Public Enum enSearchType
'    enSearchByCatalogue = 1
'    enSearchNormal = 2
'    enSearchAdvanced = 3
'    enSearchBF = 4
'    enSearchBIC = 5
'End Enum
Public Property Get rows() As Long
    On Error GoTo errHandler
    rows = lngrowcount
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SearchEngineB.rows"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.rows"
End Property

Public Function instock(val As Boolean)
    On Error GoTo errHandler
    If val Then
            intstock = 1
    Else
            intstock = 0
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.instock(val)", val
End Function
Public Function WithCopies(val As Integer)
    On Error GoTo errHandler
    Select Case val
        Case 0
            intWithCopies = 0
        Case 1
            intWithCopies = 1
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.WithCopies(val)", val
End Function

Public Function execute(MaxResult As Long)
    On Error GoTo errHandler
    
Dim objDisplay As d_Product
Dim i As Long
Dim strArticle As String
Dim strTitleNet As String

    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.CursorLocation = adUseClient
    lngretval = oSQL.RunGetRecordset("select DISTINCT TOP " & MaxResult & " DISTRIBUTOR,P_Status,P_ID, P_Article,P_Code,P_EAN,P_CodeF, P_Title,P_BIC, P_MainAuthor, P_Publisher,P_Edition, QTYCopiesONHand,QtyOnHand,QtyOnOrder,QtyOnBackOrder,P_PubDate,P_PUBPLACE,QtyTotalSold,LastDateDelivered,LocalPrice,UKPrice,USPrice,OtherstoresOH,Serial,PurchaseDate,SoldDate,CopyPrice,CopiesSold,Obsolete,Img,Length,Width from  Search" & CStr(oPC.WorkstationID), enText, Array(), "", rs)
    rowsreturned
    If lngrowcount > 0 Then
      i = 1
      rs.MoveFirst
      Do While Not rs.eof  ' And i <= MaxResult
          Set objDisplay = New d_Product
          With objDisplay
          .PID = FNS(rs!P_ID)
          .code = FNS(rs!P_Code)
          .EAN = FNS(rs!P_EAN)
          .CodeF = FNS(rs!P_CodeF) & IIf(FNN(rs!OTHERSTORESOH) > 0, "*", "")
          .Status = FNS(rs!P_STATUS)
          strArticle = FNS(rs!P_Article)
          .Length = FNDBL(rs!Length)
          .Title = FNS(rs!P_Title)
          .Width = FNDBL(rs!Width)
          If Not IsNull(rs!Img) Then .Img = rs!Img
            .QtyOnHand = FNN(rs!QtyOnHand)
          .QtyonOrder = FNN(rs!QtyonOrder)
          .QtyOnBackorder = FNN(rs!QtyOnBackorder)
          .QtyTotalSold = FNN(rs!QtyTotalSold)
          .LastDateDelivered = FND(rs!LastDateDelivered)
          .LocalPrice = FNN(rs!LocalPrice)
          .ImageFilename = FNS(rs!P_BIC)
    
          .Obsolete = FNB(rs!Obsolete)
          colDisplay.Add objDisplay '(.BookID)
          Set objDisplay = Nothing
          rs.MoveNext
          i = i + 1
          End With
    Loop
  End If
  rs.Close
  Set rs = Nothing
  Set oSQL = Nothing

  Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.execute(MaxResult)", MaxResult
End Function

Public Property Get getcols() As Collection

    Set getcols = colDisplay
    Set colDisplay = Nothing

End Property

Private Sub rowsreturned()
    On Error GoTo errHandler
    lngrowcount = 0
    
    Do Until rs.eof = True
        lngrowcount = lngrowcount + 1
        rs.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.rowsreturned"
End Sub

Private Sub droptemp()
    On Error GoTo errHandler

    Set oSQL = New z_SQL
    
    lngretval = oSQL.DropTable("search" & oPC.WorkstationID, "")
    Set oSQL = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.droptemp"
End Sub

Public Function SetupSQLwoCriteria2(pCopiesAvailable As Boolean, lngMaxRowsToReturn As Long, pProductType As String, Optional bIncludeObsolete As Boolean = False) 'VARCHAR(40)
    On Error GoTo errHandler
Dim strProductTypeParams As String
Dim i As Integer
Dim strPos As String
Dim bExcludeObsolete As Boolean

Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    strProductTypeParams = ""
    For i = 1 To Len(pProductType)
        strProductTypeParams = strProductTypeParams & IIf(Len(strProductTypeParams) > 0, " OR ", "") & "P_ProductType = '" & Mid(pProductType, i, 1) & "'"
    Next
    If Not bIncludeObsolete Then
        strProductTypeParams = strProductTypeParams & IIf(strProductTypeParams > "", " AND ", "") & " ISNULL(P_OBSOLETE,0) = 0"
    End If
    If strProductTypeParams > "" Then strProductTypeParams = "(" & strProductTypeParams & ")"
    
    If intsearchcount = 0 Then
        droptemp
        oPC.COShort.execute "CREATE TABLE  Search" & CStr(oPC.WorkstationID) & " (Distributor VARCHAR(100),P_ID UNIQUEIDENTIFIER,P_CODE VARCHAR(20),P_EAN VARCHAR(25),P_CodeF VARCHAR(20),P_TITLE VARCHAR(900), " _
            & " P_ARTICLE VARCHAR(10),P_MAINAUTHOR VARCHAR(900),P_PUBLISHER VARCHAR(900),P_Edition VARCHAR(200),QTYONHAND INT,QTYCOPIESONHAND INT,QTYONORDER INT,QTYONBACKORDER INT,P_PubDate VARCHAR(500),P_PUBPLACE VARCHAR(500),P_BIC VARCHAR(51),LASTDATEDELIVERED DATETIME,QTYTOTALSOLD INT,P_SUPPLIERID INT,LocalPrice Int,UKPrice Int,USPrice Int,P_Status Varchar(10),OTHERSTORESOH CHAR(1),Serial INT,CopiesSold INT, CopyPrice INT,PUrchaseDate DateTime,SoldDate DateTime,Obsolete Bit,IMG VARBinary(MAX),Length Real,Width Real)"
            
            strSQL = "INSERT INTO Search" & CStr(oPC.WorkstationID) & " (Distributor,P_ID,P_CODE,P_EAN,P_CODEF,P_TITLE,P_Article,P_MAINAUTHOR,P_PUBLISHER,P_EDITION,P_SUPPLIERID,QTYONHAND, " _
                            & " QTYONORDER,QTYONBACKORDER,P_PubDate,P_PUBPLACE,LASTDATEDELIVERED,QTYTOTALSOLD,P_Status,localPrice,UKPrice,USPrice,OTHERSTORESOH,Obsolete,Img,Length,Width,P_BIC) " _
                            & " SELECT DISTINCT ISNULL(TP_NAme,''),dbo.tProduct.P_ID,Left(P_Code,20),LEFT(P_EAN,25),dbo.codef(P_Code,P_EAN,0), P_Title,  P_ARTICLE,P_MainAuthor, P_Publisher,P_EDITION," _
                            & " P_SupplierID ,P_QTYONHAND,P_QTYONORDER,P_QTYONBACKORDER,P_PubDate,P_PUBPLACE,P_LASTDATEDELIVERED,P_QTYTOTALSOLD,ISNULL(P_Status,'IP'),P_SP,P_UKPrice,P_USPrice" _
                            & " ,OtherStoresOH,P_Obsolete,P_TinyImage,P_Length,P_Width,P_BIC FROM tProduct LEFT JOIN tTP ON P_SUpplierID = TP_ID LEFT JOIN vOtherstoresOH ON P_ID = STP_P_ID  WHERE  " & strProductTypeParams & " AND "
                            
                            
    End If
        strPos = "Pos 11"
    If lngMaxRowsToReturn > 0 Then
        strSQL = Replace(strSQL, "SELECT DISTINCT", "SELECT DISTINCT TOP " & CStr(lngMaxRowsToReturn))
        strPos = "Pos 12"
    End If
        strPos = "Pos 13"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SearchEngineB.SetupSQLwoCriteria(pCopiesAvailable,SearchType,BooksWithCopies," & _
'        "lngMaxRowsToReturn,pProductType,bIncludeObsolete)", Array(pCopiesAvailable, SearchType, _
'         BooksWithCopies, lngMaxRowsToReturn, pProductType, bIncludeObsolete)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.SetupSQLwoCriteria2(pCopiesAvailable,lngMaxRowsToReturn,pProductType," & _
        "bIncludeObsolete)", Array(pCopiesAvailable, lngMaxRowsToReturn, pProductType, bIncludeObsolete)
End Function

Public Sub selectcriteria(T As String, c As String, pRecsFound As Long)
    On Error GoTo errHandler

    Set colDisplay = Nothing
    Set colDisplay = New Collection
    
    Select Case T
        Case "2-Code"
            strcriteria = "P_Code like '" & c & "'"
        Case "5-Author Contains"
            strcriteria = "P_MainAuthor like '%" & c & "%'"
        Case "4-Author Starts With"
            strcriteria = "P_MainAuthor like '" & c & "%'"
        Case "3-Title Contains"
            strcriteria = "P_Title like '%" & c & "%'"
        Case "1-Title Starts With"
            strcriteria = "P_Title like '" & c & "%'"
        Case "7-Publisher"
            strcriteria = "P_Publisher like '" & c & "%'"
        Case "9-BIC"
            strcriteria = "P_BIC like '" & c & "%' OR P_BIC LIKE '," & c & "%'  OR P_BIC LIKE '%," & c & ",%'"
        Case "8-Category"
            Select Case intsearchcount
                Case 0
                    strcriteria = "Category.Category_ID = " & c & ""
                    If intstock = 1 Then
                        strcriteria = strcriteria & " and Product.P_Stockbalance > 0"
                        intstock = 0
                    End If
                Case Else
                    strcriteria = "P_C_ID = " & c & ""
            End Select
        Case "6-Supplier"
            Select Case intsearchcount
                Case 0
                    strcriteria = "[Trading Partner].Trading_Partner_ID = " & c & ""
                    If intstock = 1 Then
                        strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
                        intstock = 0
                    End If
                Case Else
                    strcriteria = "P_LastSupplierUsedID = " & c & ""
            End Select
        Case "Catalogue"
            Select Case intsearchcount
                Case 0
                    If Not IsNumeric(c) Then c = "0"
                    strcriteria = "CATAL_Serial = '" & c & "'"
                    If intstock = 1 Then
                        strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
                        intstock = 0
                    End If
                Case Else
                    strcriteria = "CAT_SerialNum = " & c & ""
            End Select
        End Select
        
        If intstock = 1 Then
            strcriteria = strcriteria & " and P_QtyCopiesOnHand > 0"
        End If
        Set oSQL = New z_SQL

        pRecsFound = oSQL.RunSQL(strSQL & strcriteria, oPC.COShort)
        Set oSQL = Nothing
'H:
'    oor.Setor , or, Now, "z_SearchEngineB:SelectCriteria", "", ""
'    Exit Sub
'    Resume
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.selectcriteria(t,c,pRecsFound)", Array(T, c, pRecsFound)
End Sub

Public Sub prisearch()
    On Error GoTo errHandler

    intsearchcount = 0
    inttemptbl = 0

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.prisearch"
End Sub

Public Sub secsearch()
    On Error GoTo errHandler

    Select Case intsearchcount
        Case 0
            intsearchcount = 1
        Case 1
            intsearchcount = 2
        Case 2
            intsearchcount = 1
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.secsearch"
End Sub


Private Sub Class_Initialize()
    On Error GoTo errHandler
'Debug.Print "Initialize SearchEngine"
'    addcols
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
Debug.Print "Terminate SearchEngine"
    Set colobjs = Nothing
    Set oSQL = Nothing
    Set colDisplay = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.Class_Terminate"
End Sub

Public Sub SimpleSearch(pArg As String, pRecsFound As Long)
    On Error GoTo errHandler
Dim strTmp As String
Const strDS = "TAP"
Dim EOCriteria As Boolean
Dim i As Integer
Dim strArg As String
Dim strRightWildCard As String
Dim strLeftWildCard As String
Dim strPlainarg As String
Dim ar() As String
Dim iWordCount As Integer
Dim EOFCriteria As Boolean
Dim lngRecsReturned As Long
'Dim oProdCode As New z_ProdCode

    pArg = Replace(pArg, "'''", "'")
    pArg = Replace(pArg, "''", "'")
    'THis is removed when we are not looking for free text
    'pArg = Replace(pArg, "'", "''")

    Set colDisplay = Nothing
    Set colDisplay = New Collection
    ar = Split(FNS(pArg), " ")
    If Left(pArg, 1) = "/" Then
        If Mid(pArg, 3, 1) = " " Then 'Use criteria type specified by character 2
        Else
        End If  ' USe first criteria type in default criteria sequence
    ElseIf InStr(pArg, " ") = 0 And Len(pArg) > 0 Then 'Search by Code then ISBN  THere is no multiple word search here, we may have a strange loal code e.g. BGBG
        strcriteria = "P_Code = '" & pArg & "'"
        lngretval = oSQL.RunSQL(strSQL & strcriteria)
        If lngretval = 0 Then
            lngretval = oSQL.RunSQL(strSQL & "P_CODE = dbo.CONVERTEANTOISBN('" & pArg & "')")
        End If
        If lngretval = 0 Then
            strcriteria = "P_EAN = '" & pArg & "'"
            lngretval = oSQL.RunSQL(strSQL & strcriteria)
        End If
        i = i + 1
    ElseIf Left(pArg, 1) = "#" Then 'Search by code
        strcriteria = "P_Code = '" & pArg & "'"
        lngretval = oSQL.RunSQL(strSQL & strcriteria)
        i = i + 1
    End If
    If lngretval = 0 Then
        Set oSQL = New z_SQL
        strcriteria = ""

        If intstock = 1 Then
            strcriteria = strcriteria & IIf(strcriteria > "", " AND ", "") & "  (P_QtyCopiesOnHand > 0 OR P_QTYONHAND > 0)"
        End If
        
        If strcriteria = "" Then
            strSQL = Left(strSQL, Len(strSQL) - 4)
        End If
        
        lngretval = oSQL.RunSQL(strSQL & " AND " & pArg, oPC.COShort)
        lngRecsReturned = lngRecsReturned + lngretval
        i = i + 1
        If i > Len(strDS) Then EOFCriteria = True
            
        pRecsFound = lngRecsReturned
    End If
'H:
'    oor.Setor , or, Now, "z_SearchEngineB:SimpleSearch", "", ""
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.SimpleSearch(pArg,pRecsFound)", Array(pArg, pRecsFound)
End Sub
Private Function NextPart(iStart As Long, arg1 As Long, arg2 As Long, arg3 As Long) As Integer
Dim a() As Long
Dim i As Integer
Dim m As Long
    ReDim a(3)
    
    a(0) = arg1
    a(1) = arg2
    a(2) = arg3
    
    m = 9999
    For i = 0 To 2
        If a(i) > iStart Then
            If a(i) <> 0 Then
                If m > a(i) Then
                    m = a(i)
                End If
            End If
        End If
    Next
    NextPart = m

End Function
Public Sub AdvancedSearch(pRecsFound As Long, pArg As String, Optional pSectionID As Long, Optional pProductTypeID As Long, Optional bMissingCategory As Boolean, Optional bMissingProductType As Boolean)
    On Error GoTo errHandler
Dim strTmp As String
Dim strDS As String
Dim EOCriteria As Boolean
Dim i As Integer
Dim strArg As String
Dim arg1 As String
Dim arg2 As String
Dim arg3 As String
Dim arg4 As String

Dim ar() As String
Dim iWordCount As Integer
Dim EOFCriteria As Boolean
Dim strTitleCrit As String
Dim strPubCrit As String
Dim strDISTCrit As String
Dim strAuthorCrit As String
Dim iStartTitle As Long
Dim iStartAuthor As Long
Dim iStartPub As Long
Dim iStartDist As Long
Dim lngRecsReturned As Long

    
    pArg = Replace(pArg, "'''", "'")
    pArg = Replace(pArg, "''", "'")

    Set colDisplay = Nothing
    Set colDisplay = New Collection
        
        Set oSQL = New z_SQL
        strcriteria = pArg
        i = 1
            If intstock = 1 Then
                    strcriteria = strcriteria & " and P_QtyOnHand > 0"
            End If
            If pSectionID > 0 Then
                i = InStr(1, strSQL, "WHERE") - 1
                strSQL = Left(strSQL, i) & "LEFT JOIN tProductSection ON P_ID = PSEC_P_ID" & Right(strSQL, Len(strSQL) - i + 1)
                If strcriteria = "" Then
                    strcriteria = strcriteria & " PSEC_SEC_ID = " & pSectionID
                Else
                    strcriteria = strcriteria & " AND PSEC_SEC_ID = " & pSectionID
                End If
            Else
                If bMissingCategory Then
                    i = InStr(1, strSQL, "WHERE") - 1
                    strSQL = Left(strSQL, i) & "LEFT JOIN tProductSection ON P_ID = PSEC_P_ID" & Right(strSQL, Len(strSQL) - i + 1)
                    strcriteria = strcriteria & " AND PSEC_SEC_ID  IS NULL"
                End If
            End If
            If pProductTypeID > 0 Then
                If strcriteria = "" Then
                    strcriteria = strcriteria & " P_ProductType_ID = " & pProductTypeID
                Else
                    strcriteria = strcriteria & " AND P_ProductType_ID = " & pProductTypeID
                End If
            Else
                If bMissingProductType Then
                    strcriteria = strcriteria & " AND P_ProductType_ID IS NULL"
                End If
            End If
        If UCase(Left(strcriteria, 4)) = " AND" Then
            strcriteria = Right(strcriteria, Len(strcriteria) - 4)
        End If
        If strcriteria > "" Then
            If mbDebug Then LogSaveToFile strSQL & strcriteria
            lngretval = oSQL.RunSQL(strSQL & strcriteria, oPC.COShort)
            If lngretval = -1 Then
                lngRecsReturned = -1
            Else
                lngRecsReturned = lngRecsReturned + lngretval
            End If
            If mbDebug Then LogSaveToFile "Records found:" & CStr(lngretval)
        End If
        pRecsFound = lngRecsReturned
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.AdvancedSearch(pRecsFound,pArg)", Array(pRecsFound, pArg)
End Sub
Private Function Strip_WHERE(pIn As String) As String
    On Error GoTo errHandler
Dim tmp As String
    tmp = RTrim(pIn)
    If Right(tmp, 5) = "WHERE" Then
        tmp = Left(tmp, Len(tmp) - 5)
    End If
    Strip_WHERE = tmp
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.Strip_WHERE(pIn)", pIn
End Function
'Public Sub BFSearchEx(pArg As String, pRecsFound As Long, MaxResult As Long, Optional pResult As Long)
'    On Error GoTo errHandler
'Dim lngFound As Long
'
''Dim oBF As a_BookFind
'    Set colDisplay = Nothing
'    Set colDisplay = New Collection
'    Set oBF = New a_BookFind
'    lngResult = oBF.FetchFromBFEx(Trim(pArg), pRecsFound, MaxResult)
'    pResult = lngResult
'    Set oBF = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SearchEngineB.BFSearchEx(pArg,pRecsFound,MaxResult,pResult)", Array(pArg, pRecsFound, _
'         MaxResult, pResult)
'End Sub
'Private Sub LoadProductFromBF()
'    On Error GoTo ErrHandler
''    mudtProps.Code = oBF.Code
''    mudtProps.Author = oBF.MainAuthor
''    mudtProps.Title = oBF.Title
''    mudtProps.SubTitle = oBF.SubTitle
''    mudtProps.Edition = oBF.Edition
''    mudtProps.Publisher = oBF.PublisherName
''    mudtProps.Note = oBF.Note
''    mudtProps.SeriesTitle = oBF.SeriesTitle
''    mudtProps.Availability = oBF.Availability
''    mudtProps.UKPrice = oBF.UKPrice
''    mudtProps.USPrice = oBF.USPrice
''    mudtProps.RRP = oBF.LocalPrice
''    mudtProps.PublicationDate = oBF.PublicationDate
''    mudtProps.Description = oBF.Description
''    mudtProps.MainSupplierName = oBF.MainSupplierName
''    mudtProps.BindingCode = oBF.BindingCode
''    mudtProps.BFClassification = oBF.BFClassification
''    MsgBox "Possibly problem assigning properties directly -specially prices"
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SearchEngineB.LoadProductFromBF"
'End Sub
Public Sub SearchBIC(pArg As String, pRecsFound As Long)
    On Error GoTo errHandler
Dim strTmp As String
Const strDS = "TAP"
Dim EOCriteria As Boolean
Dim i As Integer
Dim strArg As String
Dim strRightWildCard As String
Dim strLeftWildCard As String
Dim strPlainarg As String
Dim ar() As String
Dim iWordCount As Integer
Dim EOFCriteria As Boolean
Dim lngRecsReturned As Long


    Set colDisplay = Nothing
    Set colDisplay = New Collection
    Set oSQL = New z_SQL
    strcriteria = "  Patindex('," & pArg & "%',P_BIC) > 0 OR Patindex('" & pArg & "%',P_BIC) > 0"
    If intstock = 1 Then
        strcriteria = strcriteria & " and (P_QtyCopiesOnHand > 0 OR P_QTYONHAND > 0)"
    End If
    If strcriteria > "" Then
        lngretval = oSQL.RunSQL(strSQL & strcriteria, oPC.COShort)
        lngRecsReturned = lngRecsReturned + lngretval
    End If
    pRecsFound = lngRecsReturned
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.SimpleSearch(pArg,pRecsFound)", Array(pArg, pRecsFound)
End Sub

Public Function GoogleSearchString(pArg As String) As String
    On Error GoTo errHandler
Dim strTmp As String
Dim strDS As String
Dim EOCriteria As Boolean
Dim i As Integer
Dim strArg As String
Dim arg1 As String
Dim arg2 As String
Dim arg3 As String
Dim arg4 As String

Dim ar() As String
Dim iWordCount As Integer
Dim EOFCriteria As Boolean
Dim strTitleCrit As String
Dim strPubCrit As String
Dim strDISTCrit As String
Dim strAuthorCrit As String
Dim iStartTitle As Long
Dim iStartAuthor As Long
Dim iStartPub As Long
Dim iStartDist As Long
Dim lngRecsReturned As Long

    If Len(pArg) = 13 Then
        If IsNumeric(pArg) Then
            GoogleSearchString = "ISBN" & pArg
            Exit Function
        End If
    Else
        If Len(pArg) = 10 Then
            If IsNumeric(Left(pArg, 9)) Then
                GoogleSearchString = "ISBN" & pArg
                Exit Function
            End If
        End If
    End If
    pArg = Replace(pArg, "////", "^")
    pArg = Replace(pArg, "///", "#")
    pArg = Replace(pArg, "//", "~")
    pArg = Replace(pArg, "*", "%")
    
    pArg = Replace(pArg, "'''", "'")
    pArg = Replace(pArg, "''", "'")
    pArg = Replace(pArg, "'", "''")
    
    iStartTitle = InStr(1, pArg, "/")
    iStartAuthor = InStr(1, pArg, "~")
    iStartPub = InStr(1, pArg, "#")
    iStartDist = InStr(1, pArg, "^")
    
    'Check for Title search
    If iStartTitle > 0 Then
        strTitleCrit = Mid(pArg, iStartTitle + 1, IIf(NextPart(iStartTitle, iStartAuthor, iStartPub, iStartDist) > iStartTitle + 1, NextPart(iStartTitle, iStartAuthor, iStartPub, iStartDist) - iStartTitle - 1, 999))
    End If
    If iStartAuthor > 0 Then
        strAuthorCrit = Mid(pArg, iStartAuthor + 1, IIf(NextPart(iStartAuthor, iStartTitle, iStartPub, iStartDist) > iStartAuthor + 1, NextPart(iStartAuthor, iStartTitle, iStartPub, iStartDist) - iStartAuthor - 1, 999))
    End If
    If iStartPub > 0 Then
        strPubCrit = Mid(pArg, iStartPub + 1, IIf(NextPart(iStartPub, iStartAuthor, iStartTitle, iStartDist) > iStartPub + 1, NextPart(iStartPub, iStartAuthor, iStartTitle, iStartDist) - iStartPub - 1, 999))
    End If
    If iStartDist > 0 Then
        strDISTCrit = Mid(pArg, iStartDist + 1, IIf(NextPart(iStartDist, iStartAuthor, iStartTitle, iStartPub) > iStartDist + 1, NextPart(iStartDist, iStartAuthor, iStartTitle, iStartPub) - iStartDist - 1, 999))
    End If

    strDS = ""
    If strTitleCrit > "" Then strDS = "T"
    If strAuthorCrit > "" Then strDS = strDS & "A"
    If strPubCrit > "" Then strDS = strDS & "P"
    If strDISTCrit > "" Then strDS = strDS & "D"

    Set colDisplay = Nothing
    Set colDisplay = New Collection
        
        Set oSQL = New z_SQL
        EOCriteria = False
        i = 1
        strcriteria = ""
        Do While Not EOFCriteria
            Select Case UCase(Mid(strDS, i, 1))
                Case "T"
                    ar = Split(FNS(strTitleCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & IIf(strcriteria > "", "+", "") & SpreadArgs(ar(0), "intitle")
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0 AND Patindex('%" & ar(3) & "%',P_TITLE + ISNULL(P_Subtitle,'')) > 0"
                    End Select
                Case "A"
                    ar = Split(FNS(strAuthorCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & IIf(strcriteria > "", "+", "") & SpreadArgs(ar(0), "inauthor")
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(1) & "%',P_MAINAUTHOR) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(1) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(2) & "%',P_MAINAUTHOR) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(1) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(2) & "%',P_MAINAUTHOR) > 0 AND Patindex('" & ar(3) & "%',P_MAINAUTHOR) > 0"
                    End Select
                Case "P"
                    ar = Split(FNS(strPubCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & IIf(strcriteria > "", "+", "") & SpreadArgs(ar(0), "inpublisher")
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex(' " & ar(1) & "%',P_PUBLISHER) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(2) & "%',P_PUBLISHER) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(2) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(3) & "%',P_PUBLISHER) > 0"
                    End Select
                Case "D"
                    ar = Split(FNS(strDISTCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & IIf(strcriteria > "", "+", "") & SpreadArgs(ar(0), "subject")
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',TP_NAme) > 0 AND Patindex(' " & ar(1) & "%',TP_NAme) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',TP_NAme) > 0 AND Patindex('" & ar(1) & "%',TP_NAme) > 0 AND Patindex('" & ar(2) & "%',TP_NAme) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',TP_NAme) > 0 AND Patindex('" & ar(1) & "%',TP_NAme) > 0 AND Patindex('" & ar(2) & "%',TP_NAme) > 0 AND Patindex('" & ar(3) & "%',TP_NAme) > 0"
                    End Select
            End Select
            i = i + 1
            If i > Len(strDS) Then EOFCriteria = True
        Loop
        If UCase(Left(strcriteria, 4)) = "&" Then
            strcriteria = Right(strcriteria, Len(strcriteria) - 1)
        End If
        GoogleSearchString = strcriteria
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SearchEngineB.AdvancedSearch(pRecsFound,pArg)", Array(pRecsFound, pArg)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineB.GoogleSearchString(pArg)", pArg
End Function

Private Function SpreadArgs(p As String, T As String) As String
Dim i As Integer
Dim a() As String
Dim o As String

    i = 0
    o = ""
    a = Split(p, " ")
    Do
        o = o & IIf(o > "", "+", "") & T & ":" & a(i)
        i = i + 1
    Loop Until UBound(a) < i
    SpreadArgs = o
End Function



