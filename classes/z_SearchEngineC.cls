VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SearchEngineC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Public Enum enSearchType
    enSearchByCatalogue = 1
    enSearchNormal = 2
    enSearchAdvanced = 3
    enSearchBF = 4
    enSearchBIC = 5
End Enum
Public Property Get rows() As Long
    rows = lngrowcount
End Property

Public Function instock(val As Integer)
    Select Case val
        Case 0
            intstock = 0
        Case 1
            intstock = 1
    End Select
End Function
Public Function WithCopies(val As Integer)
    Select Case val
        Case 0
            intWithCopies = 0
        Case 1
            intWithCopies = 1
    End Select
End Function

Public Function execute(MaxResult As Long)
    On Error GoTo ERRHANDLR
    
Dim objDisplay As d_Product
Dim i As Long
Dim strArticle As String
Dim strTitleNet As String

    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.CursorLocation = adUseClient
    lngretval = oSQL.RunGetRecordset("select DISTINCT TOP " & MaxResult & "DISTRIBUTOR,P_Status,P_ID, P_Article,P_Code,P_CodeF, P_EAN,P_Title, P_MainAuthor, P_Publisher, QTYCopiesONHand,QtyOnHand,QtyOnOrder,QtyOnBackOrder,QtyTotalSold,LastDateDelivered,LocalPrice,UKPrice,USPrice,SalesList from tempsearch ", enText, Array(), "", rs)
   ' lngrowcount = rs.RecordCount
    rowsreturned
    If lngrowcount > 0 Then
      i = 1
      rs.MoveFirst
      Do While Not rs.EOF  ' And i <= MaxResult
          Set objDisplay = New d_Product
          With objDisplay
          .PID = FNS(rs!P_ID)
          .Code = FNS(rs!P_Code)
          .EAN = FNS(rs!P_EAN)
          .CodeF = FNS(rs!P_CodeF)
          .Status = FNS(rs!P_STATUS)
          .Title = FNS(rs!P_Article) & " " & FNS(rs!P_Title)
          .Author = FNS(rs!P_MainAuthor)
            .QtyOnHand = FNN(rs!QtyOnHand)
          .QtyOnOrder = FNN(rs!QtyOnOrder)
          .QtyOnBackorder = FNN(rs!QtyOnBackorder)
          .QtyTotalSold = FNN(rs!QtyTotalSold)
          .LastDateDelivered = FND(rs!LastDateDelivered)
          .LocalPrice = FNN(rs!LocalPrice)
          .UKPrice = FNN(rs!UKPrice)
          .USPrice = FNN(rs!USPrice)
          .Distributor = FNS(rs!Distributor)
          .Publisher = FNS(rs!P_Publisher)
          .SalesList = FNS(rs!SalesList)
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
                    
ERRHANDLR:
   
 Exit Function
      Resume
End Function

Public Property Get getcols() As Collection

    Set getcols = colDisplay
    Set colDisplay = Nothing

End Property

Private Sub rowsreturned()
    
    On Error GoTo ERRHNDLER
    
    lngrowcount = 0
    
 '   rs.MoveFirst
    Do Until rs.EOF = True
        lngrowcount = lngrowcount + 1
        rs.MoveNext
    Loop
    
    'MsgBox "" & lngrowcount & " rows found", , "Search Complete!"
        
    Exit Sub
    
ERRHNDLER:
  
   MsgBox Error

End Sub

Private Sub droptemp()

    Set oSQL = New z_SQL
    
    lngretval = oSQL.DropTable("tempsearch", "")
    lngretval = oSQL.DropTable("tempsearch", "")
    
    Set oSQL = Nothing

End Sub

Public Function SetupSQLwoCriteria(pAQ As Boolean, SearchType As enSearchType, IncludeAnt As Boolean, lngMaxRowsToReturn As Long, pProductType As String) 'VARCHAR(40)
On Error GoTo ERRH
Dim strParams As String
Dim i As Integer

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strParams = ""
    For i = 1 To Len(pProductType)
        strParams = strParams & IIf(Len(strParams) > 0, " OR ", "") & "P_ProductType = '" & Mid(pProductType, i, 1) & "'"
    Next
    If strParams > "" Then strParams = "(" & strParams & ")"
'    If pProductType = "B" Then
'        strParams = " (P_ProductType = 'B') "
'    Else
'        strParams = " (P_ProductType = 'N' or P_ProductType = 'G' or P_ProductType = 'M') "
'    End If
    If intsearchcount = 0 Then
        droptemp
        oPC.COShort.execute "CREATE TABLE tempSearch (Distributor VARCHAR(60),P_ID UNIQUEIDENTIFIER,P_CODE VARCHAR(16),P_CodeF VARCHAR(20),P_EAN VARCHAR(20),P_TITLE VARCHAR(200),P_ARTICLE VARCHAR(10),P_MAINAUTHOR VARCHAR(200),P_PUBLISHER VARCHAR(200),QTYONHAND INT,QTYCOPIESONHAND INT,QTYONORDER INT,QTYONBACKORDER INT,P_PubDate VARCHAR(500),P_PUBPLACE VARCHAR(500),LASTDATEDELIVERED DATETIME,QTYTOTALSOLD INT,P_SUPPLIERID INT,LocalPrice Int,UKPrice Int,USPrice Int,P_Status Varchar(10),SALESLIST VARCHAR(150))"
        If pAQ Then
            If SearchType = enSearchNormal Or SearchType = enSearchAdvanced Or SearchType = enSearchBIC Then
                If IncludeAnt Then  'only books with copies
                    strSQL = "INSERT INTO tempSearch (Distributor,P_ID,P_CODE,P_CODEF,P_TITLE,P_Article,P_MAINAUTHOR,P_PUBLISHER,P_SUPPLIERID,QTYCOPIESONHAND,QTYONHAND,QTYONORDER,QTYONBACKORDER,LASTDATEDELIVERED,QTYTOTALSOLD,P_Status) SELECT DISTINCT TP_NAme,P_ID,P_Code, dbo.codef(P_Code), P_Title, P_ARTICLE,P_MainAuthor, P_Publisher," _
                            & " P_SupplierID,P_QtyCopiesONHAND,P_QtyONHAND,P_QTYONORDER,P_QTYONBACKORDER,P_LASTDATEDELIVERED,P_QTYTOTALSOLD,P_Status " _
                            & " FROM tProduct LEFT Join tPI ON PI_P_ID = P_ID JOIN tTP ON P_SUpplierID = TP_ID WHERE PI_ID is not null AND " & strParams & " AND "
                Else
                    strSQL = "INSERT INTO tempSearch (Distributor,P_ID,P_CODE,P_CODEF,P_TITLE,P_Article,P_MAINAUTHOR,P_PUBLISHER,P_SUPPLIERID,QTYONHAND,QTYONORDER,QTYONBACKORDER,LASTDATEDELIVERED,QTYTOTALSOLD,P_Status,localPrice,UKPrice,USPrice) SELECT DISTINCT TP_NAme,P_ID, P_Code,dbo.codef(P_Code), P_Title,  P_ARTICLE,P_MainAuthor, P_Publisher," _
                            & " P_SupplierID ,P_QTYONHAND,P_QTYONORDER,P_QTYONBACKORDER,P_LASTDATEDELIVERED,P_QTYTOTALSOLD,P_Status,P_SP,P_UKPrice,P_USPrice" _
                            & " FROM tProduct LEFT JOIN tTP ON P_SUpplierID = TP_ID WHERE  " & strParams & " AND "
                End If
            End If
        Else
            strSQL = "INSERT INTO tempSearch (Distributor,P_ID,P_CODE,P_EAN,P_CODEF,P_TITLE,P_Article,P_MAINAUTHOR,P_PUBLISHER,P_SUPPLIERID,QTYONHAND,QTYONORDER,QTYONBACKORDER, " _
                    & " P_PubDate,P_PUBPLACE,LASTDATEDELIVERED,QTYTOTALSOLD,P_Status,localPrice,UKPrice,USPrice,SALESLIST) " _
                    & " SELECT DISTINCT ISNULL(TP_NAme,''),P_ID,Left(P_Code,20),LEFT(P_EAN,25),dbo.codef(P_Code,P_EAN,0), P_Title,  P_ARTICLE,P_MainAuthor, P_Publisher," _
                    & " P_SupplierID ,P_QTYONHAND,P_QTYONORDER,P_QTYONBACKORDER,P_PubDate,P_PUBPLACE,P_LASTDATEDELIVERED,P_QTYTOTALSOLD,ISNULL(P_Status,'IP'),P_SP,P_UKPrice,P_USPrice,dbo.FormatSOH_List(P_ID)" _
                    & "  FROM tProduct LEFT JOIN tTP ON P_SUpplierID = TP_ID  WHERE  " & strParams & IIf(strParams > "", " AND ", "")
        End If
    Else
    End If
    If lngMaxRowsToReturn > 0 Then
        strSQL = Replace(strSQL, "SELECT DISTINCT", "SELECT DISTINCT TOP " & CStr(lngMaxRowsToReturn))
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ERRH:
    Exit Function
    Resume
End Function
'''Public Function SetupSQLwoCriteria(pAntiquarianUsage As Boolean, pCopiesAvailable As Boolean, SearchType As enSearchType, BooksWithCopies As Boolean, lngMaxRowsToReturn As Long, pProductType As String) 'VARCHAR(40)
'''    On Error GoTo errHandler
'''Dim strProductTypeParams As String
'''Dim i As Integer
'''Dim strPos As String
'''
'''    If SearchType = enSearchBIC Then
'''        strProductTypeParams = "P_ProductType = 'B'"
'''    Else
'''        strProductTypeParams = ""
'''        For i = 1 To Len(pProductType)
'''            strProductTypeParams = strProductTypeParams & IIf(Len(strProductTypeParams) > 0, " OR ", "") & "P_ProductType = '" & Mid(pProductType, i, 1) & "'"
'''        Next
'''        If strProductTypeParams > "" Then strProductTypeParams = "(" & strProductTypeParams & ")"
'''    End If
'''    If intsearchcount = 0 Then
'''        droptemp
'''        oPC.COShort.execute "CREATE TABLE  Search" & CStr(oPC.WorkstationID) & " (Distributor VARCHAR(100),P_ID UNIQUEIDENTIFIER,P_CODE VARCHAR(20),P_EAN VARCHAR(25),P_CodeF VARCHAR(20),P_TITLE VARCHAR(900),P_ARTICLE VARCHAR(10),P_MAINAUTHOR VARCHAR(900),P_PUBLISHER VARCHAR(900),QTYONHAND INT,QTYCOPIESONHAND INT,QTYONORDER INT,QTYONBACKORDER INT,P_PubDate VARCHAR(500),P_PUBPLACE VARCHAR(500),LASTDATEDELIVERED DATETIME,QTYTOTALSOLD INT,P_SUPPLIERID INT,LocalPrice Int,UKPrice Int,USPrice Int,P_Status Varchar(10),OTHERSTORESOH CHAR(1),Serial INT,CopiesSold INT, CopyPrice INT,PUrchaseDate DateTime,SoldDate DateTime)"
'''        If SearchType = enSearchNormal Or SearchType = enSearchAdvanced Or SearchType = enSearchBIC Then
'''                strSQL = "INSERT INTO Search" & CStr(oPC.WorkstationID) & " (Distributor,P_ID,P_CODE,P_EAN,P_CODEF,P_TITLE,P_Article,P_MAINAUTHOR,P_PUBLISHER,P_SUPPLIERID,QTYONHAND,QTYONORDER,QTYONBACKORDER,P_PubDate,P_PUBPLACE,LASTDATEDELIVERED,QTYTOTALSOLD,P_Status,localPrice,UKPrice,USPrice,OTHERSTORESOH) SELECT DISTINCT ISNULL(TP_NAme,''),P_ID,Left(P_Code,20),LEFT(P_EAN,25),dbo.codef(P_Code,P_EAN,0), P_Title,  P_ARTICLE,P_MainAuthor, P_Publisher," _
'''                        & " P_SupplierID ,P_QTYONHAND,P_QTYONORDER,P_QTYONBACKORDER,P_PubDate,P_PUBPLACE,P_LASTDATEDELIVERED,P_QTYTOTALSOLD,ISNULL(P_Status,'IP'),P_SP,P_UKPrice,P_USPrice" _
'''                        & " ,OtherStoresOH FROM tProduct LEFT JOIN tTP ON P_SUpplierID = TP_ID LEFT JOIN vOtherstoresOH ON P_ID = STP_P_ID WHERE  " & strProductTypeParams & " AND "
'''        End If
'''    End If
'''        strPos = "Pos 11"
'''    If lngMaxRowsToReturn > 0 Then
'''        strSQL = Replace(strSQL, "SELECT DISTINCT", "SELECT DISTINCT TOP " & CStr(lngMaxRowsToReturn))
'''        strPos = "Pos 12"
'''    End If
'''        strPos = "Pos 13"
'''    Exit Function
'''errHandler:
'''    If ErrMustStop Then Debug.Assert False: Resume
'''    ErrorIn "z_SearchEngineB.SetupSQLwoCriteria(pAntiquarianUsage,pCopiesAvailable,SearchType," & _
'''        "BooksWithCopies,lngMaxRowsToReturn,pProductType)", Array(pAntiquarianUsage, pCopiesAvailable, SearchType, _
'''         BooksWithCopies, lngMaxRowsToReturn, pProductType), , , "strPos=", Array(strPos)
'''End Function

Public Sub selectcriteria(T As String, c As String, pRecsFound As Long)
On Error GoTo ERRH

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
                    strcriteria = "CATAL_Serial = " & CInt(c) & ""
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

        pRecsFound = oSQL.RunSQL(strSQL & strcriteria)
        Set oSQL = Nothing
    Exit Sub
ERRH:

End Sub

Public Sub prisearch()

    intsearchcount = 0
    inttemptbl = 0

End Sub

Public Sub secsearch()

    Select Case intsearchcount
        Case 0
            intsearchcount = 1
        Case 1
            intsearchcount = 2
        Case 2
            intsearchcount = 1
    End Select

End Sub

Private Sub Class_Initialize()
'Debug.Print "Initialize SearchEngine"
'    addcols
   
End Sub

Private Sub Class_Terminate()
Debug.Print "Terminate SearchEngine"
    Set colobjs = Nothing
    Set oSQL = Nothing
    Set colDisplay = Nothing

End Sub

Public Sub SimpleSearch(pArg As String, pRecsFound As Long)
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

On Error GoTo ERRH

    Set colDisplay = Nothing
    Set colDisplay = New Collection
    ar = Split(FNS(pArg), " ")
    If Left(pArg, 1) = "/" Then
        If Mid(pArg, 3, 1) = " " Then 'Use criteria type specified by character 2
        Else
        End If  ' USe first criteria type in default criteria sequence
    ElseIf IsNumeric(Left(pArg, 9)) Then  'Search by Code then ISBN
        strcriteria = "P_Code = '" & pArg & "'"
        lngretval = oSQL.RunSQL(strSQL & strcriteria)
        strcriteria = "P_EAN = '" & pArg & "'"
        lngretval = oSQL.RunSQL(strSQL & strcriteria)
        i = i + 1
    ElseIf Left(pArg, 1) = "#" Then 'Search by code
        strcriteria = "P_Code = '" & pArg & "'"
        lngretval = oSQL.RunSQL(strSQL & strcriteria)
        i = i + 1
    Else    'Default search sequence
        If Left(pArg, 1) = "*" Then
            strLeftWildCard = "*"
            strPlainarg = Right(pArg, Len(pArg) - 1)
        Else
            strLeftWildCard = ""
            strPlainarg = pArg
        End If
        If Right(strPlainarg, 1) = "*" Then
            strPlainarg = Left(strPlainarg, Len(strPlainarg) - 1)
            strRightWildCard = "*"
        Else
            strRightWildCard = ""
        End If
        ar = Split(strPlainarg, " ")
        iWordCount = UBound(ar) + 1
        Set oSQL = New z_SQL
        EOCriteria = False
        i = 1
        Do While Not EOFCriteria
            Select Case UCase(Mid(strDS, i, 1))
                Case "T"
                    Select Case iWordCount
                    Case 1
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_TITLE) > 0"
                    Case 2
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0"
                    Case 3
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE) > 0"
                    Case 4
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE) > 0 AND Patindex('%" & ar(3) & "%',P_TITLE) > 0"
                    End Select
                Case "A"
                    Select Case iWordCount
                    Case 1
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_MAINAUTHOR) > 0"
                    Case 2
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('%" & ar(1) & "%',P_MAINAUTHOR) > 0"
                    Case 3
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('%" & ar(1) & "%',P_MAINAUTHOR) > 0 AND Patindex('%" & ar(2) & "%',P_MAINAUTHOR) > 0"
                    Case 4
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_MAINAUTHOR) > 0 AND Patindex('%" & ar(1) & "%',P_MAINAUTHOR) > 0 AND Patindex('%" & ar(2) & "%',P_MAINAUTHOR) > 0 AND Patindex('%" & ar(3) & "%',P_MAINAUTHOR) > 0"
                    End Select
                Case "P"
                    Select Case iWordCount
                    Case 1
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_PUBLISHER) > 0"
                    Case 2
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('%" & ar(1) & "%',P_PUBLISHER) > 0"
                    Case 3
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('%" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('%" & ar(2) & "%',P_PUBLISHER) > 0"
                    Case 4
                        strcriteria = "  Patindex('%" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('%" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('%" & ar(2) & "%',P_PUBLISHER) > 0 AND Patindex('%" & ar(3) & "%',P_PUBLISHER) > 0"
                    End Select
            End Select
            If intstock = 1 Then
                strcriteria = strcriteria & " and (P_QtyCopiesOnHand > 0 OR P_QTYONHAND > 0)"
            End If
            If strcriteria > "" Then
                lngretval = oSQL.RunSQL(strSQL & strcriteria)
                lngRecsReturned = lngRecsReturned + lngretval
            End If
            i = i + 1
            If i > Len(strDS) Then EOFCriteria = True
        Loop
        pRecsFound = lngRecsReturned
    End If
    Exit Sub
ERRH:

End Sub

Public Sub AdvancedSearch(pRecsFound As Long, pArg As String)
On Error GoTo ERRH
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
Dim strAuthorCrit As String
Dim iStartTitle As Long
Dim iStartAuthor As Long
Dim iStartPub As Long
Dim lngRecsReturned As Long

    pArg = Replace(pArg, "///", "#")
    pArg = Replace(pArg, "//", "~")
    pArg = Replace(pArg, "*", "%")
    
    iStartTitle = InStr(1, pArg, "/")
    iStartAuthor = InStr(1, pArg, "~")
    iStartPub = InStr(1, pArg, "#")

    'Check for Title search
    If iStartTitle > 0 Then
        strTitleCrit = Mid(pArg, iStartTitle + 1, IIf(GetMin(iStartAuthor, iStartPub, True) > iStartTitle + 1, GetMin(iStartAuthor, iStartPub, True) - iStartTitle - 1, 999))
    End If
    If iStartAuthor > 0 Then
        strAuthorCrit = Mid(pArg, iStartAuthor + 1, IIf(GetMin(iStartTitle, iStartPub, True) > iStartAuthor + 1, GetMin(iStartTitle, iStartPub, True) - iStartAuthor - 1, 999))
    End If
    If iStartPub > 0 Then
        strPubCrit = Mid(pArg, iStartPub + 1, IIf(GetMin(iStartAuthor, iStartTitle, True) > iStartPub + 1, GetMin(iStartAuthor, iStartTitle, True) - iStartPub - 1, 999))
    End If

    strDS = ""
    If strTitleCrit > "" Then strDS = "T"
    If strAuthorCrit > "" Then strDS = strDS & "A"
    If strPubCrit > "" Then strDS = strDS & "P"

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
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0"
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_TITLE) > 0 AND Patindex('%" & ar(1) & "%',P_TITLE) > 0 AND Patindex('%" & ar(2) & "%',P_TITLE) > 0 AND Patindex('%" & ar(3) & "%',P_TITLE) > 0"
                    End Select
                Case "A"
                    ar = Split(FNS(strAuthorCrit), "+")
                    iWordCount = UBound(ar) + 1
                    Select Case iWordCount
                    Case 1
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_MAINAUTHOR) > 0"
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
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0"
                    Case 2
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex(' " & ar(1) & "%',P_PUBLISHER) > 0"
                    Case 3
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(2) & "%',P_PUBLISHER) > 0"
                    Case 4
                        strcriteria = strcriteria & " AND  Patindex('" & ar(0) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(1) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(2) & "%',P_PUBLISHER) > 0 AND Patindex('" & ar(3) & "%',P_PUBLISHER) > 0"
                    End Select
            End Select
            i = i + 1
            If i > Len(strDS) Then EOFCriteria = True
        Loop
        If intstock = 1 Then
            strcriteria = strcriteria & " and P_QtyOnHand > 0"
        End If
        If Left(strcriteria, 4) = " AND" Then
            strcriteria = Right(strcriteria, Len(strcriteria) - 4)
        End If
        If strcriteria > "" Then
            lngretval = oSQL.RunSQL(strSQL & strcriteria)
            lngRecsReturned = lngRecsReturned + lngretval
        End If
       ' lngretval = oSQL.RunSQL(strSQL & strcriteria)
        pRecsFound = lngretval
    Exit Sub
ERRH:
    MsgBox Error
    Exit Sub
    Resume

End Sub
Private Function Strip_WHERE(pIn As String) As String
Dim tmp As String
    tmp = RTrim(pIn)
    If Right(tmp, 5) = "WHERE" Then
        tmp = Left(tmp, Len(tmp) - 5)
    End If
    Strip_WHERE = tmp
End Function
Public Sub BFSearchEx(pArg As String, pRecsFound As Long, MaxResult As Long, Optional pResult As Long)
Dim lngFound As Long

Dim oBF As a_BookFind
    Set colDisplay = Nothing
    Set colDisplay = New Collection
    Set oBF = New a_BookFind
    lngResult = oBF.FetchFromBFEx(Trim(pArg), pRecsFound, MaxResult)
    pResult = lngResult
    Set oBF = Nothing
End Sub


Public Sub SearchBIC(pArg As String, pRecsFound As Long)
    On Error GoTo ErrHandler
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
        lngretval = oSQL.RunSQL(strSQL & strcriteria)
        lngRecsReturned = lngRecsReturned + lngretval
    End If
    pRecsFound = lngRecsReturned
    Exit Sub
    
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SearchEngineC.SimpleSearch(pArg,pRecsFound)", Array(pArg, pRecsFound)
End Sub

Private Sub LoadProductFromBF()
'    mudtProps.Code = oBF.Code
'    mudtProps.Author = oBF.MainAuthor
'    mudtProps.Title = oBF.Title
'    mudtProps.SubTitle = oBF.SubTitle
'    mudtProps.Edition = oBF.Edition
'    mudtProps.Publisher = oBF.PublisherName
'    mudtProps.Note = oBF.Note
'    mudtProps.SeriesTitle = oBF.SeriesTitle
'    mudtProps.Availability = oBF.Availability
'    mudtProps.UKPrice = oBF.UKPrice
'    mudtProps.USPrice = oBF.USPrice
'    mudtProps.RRP = oBF.LocalPrice
'    mudtProps.PublicationDate = oBF.PublicationDate
'    mudtProps.Description = oBF.Description
'    mudtProps.MainSupplierName = oBF.MainSupplierName
'    mudtProps.BindingCode = oBF.BindingCode
'    mudtProps.BFClassification = oBF.BFClassification
'    MsgBox "Possibly problem assigning properties directly -specially prices"
End Sub

