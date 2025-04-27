VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum enumDestination
    enABE = 1
    enBIBLIO = 2
End Enum
Dim rs As ADODB.Recordset

Public Function Export_CAT_ABE(pFilename As String, pstrFilter As String, pPriceFactor As Double)
    On Error GoTo errHandler
Dim strSQL As String
Dim strSQL2 As String
Dim rsOUt As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter
Dim lngResult As Long

Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT P_ID, PI_ID, P_Title, HEAD1, HEAD2, P_MainAuthor, PI_PurchaseDate," _
        & " P_Subtitle, PI_SoldDate, P_Description,  P_Edition, P_PubPlace," _
        & "P_PubDate, PI_Price, PI_Serial, P_Code, CATAL_Serial AS CatalogueNum " _
        & "INTO EXPB From vGetExportData1 JOIN tCATALOGUE ON arg_CATALID = CATAL_ID "
    strSQL2 = " ORDER BY P_Title"
'1. Delete the existing temp file
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "DropTableIfExists"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@pName", adVarChar, adParamInput, 30)
    par.Value = "EXP"
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@Result", adVarChar, adParamOutput, 30)
    cmd.Parameters.Append par
    
    cmd.execute
    If cmd.Parameters(0) <> 0 Then
        MsgBox "Cannot delete" & cmd.Parameters(1)
    End If
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "DropTableIfExists"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@pName", adVarChar, adParamInput, 30)
    par.Value = "EXPB"
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@Result", adVarChar, adParamOutput, 30)
    cmd.Parameters.Append par
    
    cmd.execute
    If cmd.Parameters(0) <> 0 Then
        MsgBox "Cannot delete" & cmd.Parameters(1)
    End If
 '2. Create a empty table with the necessary field types
                'title TEXT
                'subtitle TEXT
                'Author TEXT
                'description TEXT
                'PUBPLACE TEXT
                'PUBDATE TEXT
                'HeadDescription TEXT
                'pricerand TEXT
                'priceDollar TEXT
                'PricePound TEXT
                'Comm TEXT
                'Booknum TEXT
                'CatalogueNum TEXT
                'DateSold TEXT
    oPC.COShort.execute "CREATE TABLE EXP (P_ID UNIQUEIDENTIFIER,PI_ID INT, title varchar(800),subtitle VarChar(3000),Author VarChar(500),description TEXT,PUBPLACE Varchar(500),PUBDATE Varchar(500),HeadDescription Varchar(300),  pricerand Varchar(20),priceDollar Varchar(20), PricePound Varchar(20),Comm TEXT,Booknum Varchar(15),CatalogueNum Varchar(10),DateSold DateTime);"
 '3. Retrieve raw data from tables and join fields in appropriate ways
    Set rs = New ADODB.Recordset
    oPC.COShort.execute strSQL & pstrFilter & strSQL2
    strSQL = "SELECT EXPB.P_ID, tPI.PI_Comment,tPI.PI_DESCRIPTION,tProduct.P_Comment,tProduct.P_Summary,EXPB.PI_ID, EXPB.P_Title, EXPB.HEAD1, EXPB.HEAD2, EXPB.P_MainAuthor, EXPB.PI_PurchaseDate," _
        & " EXPB.P_Subtitle, EXPB.PI_SoldDate, EXPB.P_Description,  EXPB.P_Edition, EXPB.P_PubPlace," _
        & "EXPB.P_PubDate, EXPB.PI_Price, EXPB.PI_Serial, EXPB.P_Code,EXPB.CatalogueNum " _
        & "From EXPB JOIN tPI ON EXPB.PI_ID = tPI.PI_ID JOIN tProduct on EXPB.P_ID = tProduct.P_ID WHERE ISNULL(tPI.PI_SoldDate,'1900-01-01') <= {d '1980-01-01'}  ORDER BY EXPB.P_Title"
 '   rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    lngResult = ExportToABE(rs, pFilename, pPriceFactor)
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox "EXPORT_CAT_ABE:  " & Error
    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.Export_CAT_ABE(pFilename,pstrFilter,pPriceFactor)", Array(pFilename, pstrFilter, _
         pPriceFactor)
End Function
Public Function Export_CAT_BIBLIO(pFilename As String, pstrFilter As String, pPriceFactor As Double)
    On Error GoTo errHandler
Dim strSQL As String
Dim strSQL2 As String
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT P_ID, PI_ID, P_Title, HEAD1, HEAD2, P_MainAuthor, PI_PurchaseDate," _
        & " P_Subtitle, PI_SoldDate, P_Description,  P_Edition, P_PubPlace," _
        & "P_PubDate, PI_Price, PI_Serial, P_Code, CATAL_Serial AS CatalogueNum " _
        & "INTO EXPB From vGetExportData1 JOIN tCATALOGUE ON arg_CATALID = CATAL_ID "
    strSQL2 = " ORDER BY P_TITLE"

 '1. Delete the existing temp file
    oPC.COShort.execute "DROP TABLE EXP"
    oPC.COShort.execute "DROP TABLE EXPB"
 '2. Create a empty table with the necessary field types
     'title  TEXT
     'Author  TEXT
     'description  TEXT
     'HeadDescription  TEXT
     'pricerand  TEXT
     'priceDollar  TEXT
     'PricePound  TEXT
     'Comm  TEXT
     'Booknum  TEXT
     'CatalogueNum  TEXT
    oPC.COShort.execute "CREATE TABLE EXP (P_ID UNIQUEIDENTIFIER,PI_ID INT, title varchar(800),subtitle VarChar(3000),Author VarChar(500),description TEXT,PUBPLACE Varchar(500),PUBDATE Varchar(500),HeadDescription Varchar(300),  pricerand Varchar(20),priceDollar Varchar(20), PricePound Varchar(20),Comm TEXT,Booknum Varchar(15),CatalogueNum Varchar(10),DateSold DateTime);"
'    oPC.COShort.execute "CREATE TABLE EXP2 (title TEXT, Author TEXT, description TEXT, HeadDescription TEXT, pricerand TEXT, priceDollar TEXT, PricePound TEXT, Comm TEXT, Booknum TEXT, CatalogueNum TEXT);"
 '3. Retrieve raw data from tables and join fields in appropriate ways
   ' Set rs = New ADODB.Recordset
   ' rs.Open strSQL & " " & pstrFilter & strSQL2, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    
    Set rs = New ADODB.Recordset
    oPC.COShort.execute strSQL & pstrFilter & strSQL2
    strSQL = "SELECT EXPB.P_ID, tPI.PI_Comment,tPI.PI_Condition,tProduct.P_Comment,tProduct.P_Summary,EXPB.PI_ID, EXPB.P_Title, EXPB.HEAD1, EXPB.HEAD2, EXPB.P_MainAuthor, EXPB.PI_PurchaseDate," _
        & " EXPB.P_Subtitle, EXPB.PI_SoldDate, EXPB.P_Description,  EXPB.P_Edition, EXPB.P_PubPlace," _
        & "EXPB.P_PubDate, EXPB.PI_Price, EXPB.PI_Serial, EXPB.P_Code,EXPB.CatalogueNum " _
        & "From EXPB JOIN tPI ON EXPB.PI_ID = tPI.PI_ID JOIN tProduct on EXPB.P_ID = tProduct.P_ID WHERE ISNULL(tPI.PI_SoldDate,'1900-01-01') <= {d '1980-01-01'}  ORDER BY EXPB.P_Title"
 '   rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    
    
    
    
    lngResult = ExportToBIBLIO(rs, pFilename, pPriceFactor)
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
H:
    MsgBox "EXPORT_CAT_ABE:  " & Error
    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.Export_CAT_BIBLIO(pFilename,pstrFilter,pPriceFactor)", Array(pFilename, _
         pstrFilter, pPriceFactor)
End Function
Public Function ExportToABE(rs As ADODB.Recordset, pFilename As String, pPriceFactor As Double) As Long
    On Error GoTo errHandler
Dim adjustedPrice As Long
Dim ConvertedPrice As Long
Dim oCurr As a_Currency
Dim lngRoundTo As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Dim rsOUt As ADODB.Recordset
 '   Set cmd = New ADODB.Command
  '  Set cmd.ActiveConnection = oPC.COShort
    Dim v1, v2, v3, v4, v5, v6, t1, t2, a1, d1, h1, n1, c1, ds
    Dim oTextFile As z_TextFile
    Dim strExport As String
    Dim lngLastCopyID, lngRecordsExported As Long
'"PI_Comment,P_Summary, PI_Condition, P_Summary"
    Set rsOUt = New ADODB.Recordset
    rsOUt.open "SELECT * FROM EXP", oPC.COShort, adOpenDynamic, adLockOptimistic
    lngLastCopyID = 0
    lngRecordsExported = 0
    Do While Not rs.eof
        If rs.fields("PI_ID") <> lngLastCopyID Then 'Only append one of each copy
            rsOUt.AddNew
                v1 = FNS(rs.fields("P_Description"))
                v2 = FNS(rs.fields("PI_Description"))
                v3 = FNS(rs.fields("P_Edition"))
                v4 = FNS(rs.fields("P_PubPlace"))
                v5 = FNS(rs.fields("P_Pubdate"))
                t1 = FNS(rs.fields("P_Title"))
                t2 = FNS(rs.fields("P_Subtitle"))
                rsOUt.fields("Title") = t1
                rsOUt.fields("Subtitle") = t2
                rsOUt.fields("Author") = FNS(rs.fields("P_MainAuthor"))
                d1 = FormatAntDescription(v1, v2, v3, "", "")
                rsOUt.fields("PUBPLACE") = v4
                rsOUt.fields("PUBDATE") = v5
                rsOUt.fields("Description") = d1
                rsOUt.fields("HeadDescription") = FNS(IIf(IsNull(rs.fields("Head2")), rs.fields("Head1"), rs.fields("Head2")))
                If Not IsNull(rs.fields("PI_Price")) Then
                
                    adjustedPrice = CurrRound(FNN(rs.fields("PI_Price")) * pPriceFactor / 100)
                    Set oCurr = oPC.Configuration.LocalCurrency
                    If Not oCurr Is Nothing Then
                        ConvertedPrice = adjustedPrice * oCurr.Factor
                        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(ConvertedPrice)  ' / oCurr.Divisor
                        ConvertedPrice = (RoundUp(ConvertedPrice, lngRoundTo))
                        rsOUt.fields("PriceRand") = Format(ConvertedPrice / oCurr.Divisor, oCurr.FormatString)
                    End If
                    Set oCurr = oPC.Configuration.Currencies.FindBySysname("USD")
                    If Not oCurr Is Nothing Then
                        ConvertedPrice = adjustedPrice * oCurr.Factor
                        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(ConvertedPrice)  ' / oCurr.Divisor
                        ConvertedPrice = (RoundUp(ConvertedPrice, lngRoundTo))
                        rsOUt.fields("PriceDollar") = Format((ConvertedPrice) / oCurr.Divisor, oCurr.FormatString)
                    End If
                    Set oCurr = oPC.Configuration.Currencies.FindBySysname("GBP")
                    If Not oCurr Is Nothing Then
                        ConvertedPrice = adjustedPrice * oCurr.Factor
                        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(ConvertedPrice)
                        ConvertedPrice = (RoundUp(ConvertedPrice, lngRoundTo))
                        rsOUt.fields("Pricepound") = Format(ConvertedPrice / oCurr.Divisor, oCurr.FormatString)
                    End If
                    
                End If
                v6 = FNS(rs.fields("PI_Comment"))
                rsOUt.fields("Comm") = stripCRLF(v6)
                rsOUt.fields("Booknum") = FNS(Trim(rs.fields("P_Code")) & "/" & rs.fields("PI_Serial"))
                rsOUt.fields("CatalogueNum") = FNS(rs.fields("CatalogueNum"))
                ds = FND(rs.fields("PI_SoldDate"))
                rsOUt.fields("DateSold") = Format(ds, "dd/mm/yyyy")
                
            rsOUt.Update
            lngRecordsExported = lngRecordsExported + 1
        End If
        lngLastCopyID = rs.fields("PI_ID")
        rs.MoveNext
    Loop
    
'4. Export EXP2 as delimited file
    Set oTextFile = New z_TextFile
    If Len(pFilename) = 0 Then pFilename = "ExportDefault.txt"
    oTextFile.OpenTextFile (pFilename)
    If Not rsOUt.eof Then rsOUt.MoveFirst
    strExport = ""
    Do While Not rsOUt.eof
        strExport = rsOUt.fields("Title") & "~" & rsOUt.fields("subTitle") & "~" & rsOUt.fields("Author") & "~" & rsOUt.fields("description") & "~" & rsOUt.fields("PUBPLACE") & "~" & rsOUt.fields("PUBDATE") & "~" & rsOUt.fields("headdescription") & "~" & rsOUt.fields("PriceRand") & "~" & rsOUt.fields("PriceDollar") & "~" & rsOUt.fields("pricePound") & "~" & rsOUt.fields("Comm") & "~" & rsOUt.fields("Cataloguenum") & "~" & rsOUt.fields("Booknum") & "~" & rsOUt.fields("DateSold")
        oTextFile.WriteToTextFile (strExport)
        oTextFile.WriteToTextFile ("")
        rsOUt.MoveNext
    Loop
    oTextFile.CloseTextFile
    rsOUt.Close
    Set rsOUt = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:

    MsgBox CStr(lngRecordsExported) & " records exported.", vbOKOnly, "Finished"
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.ExportToABE(rs,pFilename,pPriceFactor)", Array(rs, pFilename, pPriceFactor)
End Function
Public Function ExportToBIBLIO(rs As ADODB.Recordset, pFilename As String, pPriceFactor As Double) As Long
    On Error GoTo errHandler
Dim adjustedPrice As Long
Dim ConvertedPrice As Long
Dim rsOUt As ADODB.Recordset
Dim v1, v2, v3, v4, v5, v6, t1, t2, a1, d1, h1, n1, c1
Dim oTextFile As z_TextFile
Dim strExport As String
Dim lngLastCopyID, lngRecordsExported As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rsOUt = New ADODB.Recordset
    rsOUt.open "SELECT * FROM EXP", oPC.COShort, adOpenDynamic, adLockOptimistic
    lngLastCopyID = 0
    lngRecordsExported = 0
    Do While Not rs.eof
        If rs.fields("PI_ID") <> lngLastCopyID Then 'Only append one of each copy
            rsOUt.AddNew
                v1 = FNS(rs.fields("P_Description"))
                v2 = FNS(rs.fields("PI_Condition"))
                v3 = FNS(rs.fields("P_Edition"))
                v4 = FNS(rs.fields("P_PubPlace"))
                v5 = FNS(rs.fields("P_Pubdate"))
                rsOUt.fields("Title") = FNS(rs.fields("P_Title") & IIf(Len(rs.fields("P_SubTitle")) > 0, ", " & rs.fields("P_SubTitle"), ""))
                rsOUt.fields("Author") = FNS(rs.fields("P_MainAuthor"))
                d1 = FormatAntDescription(v1, v2, v3, v4, v5)
                If Not IsNull(d1) Then rsOUt.fields("Description") = d1
                rsOUt.fields("HeadDescription") = FNS(IIf(IsNull(rs.fields("Head2")), rs.fields("Head1"), rs.fields("Head2")))
                If Not IsNull(rs.fields("PI_Price")) Then
                    adjustedPrice = CurrRound(FNN(rs.fields("PI_Price")) * pPriceFactor / 100)
  '                  ConvertedPrice = adjustedPrice
                    rsOUt.fields("PriceDollar") = (adjustedPrice * oPC.Configuration.Currencies.FindBySysname("USD").Factor) / oPC.Configuration.Currencies.FindBySysname("USD").Divisor
                End If
                v6 = FNS(rs.fields("PI_Comment"))
                If Not IsNull(v6) Then rsOUt.fields("Comm") = stripCRLF(v6)
                rsOUt.fields("Booknum") = FNS(Trim(rs.fields("P_Code")) & "/" & rs.fields("PI_Serial"))
                rsOUt.fields("CatalogueNum") = FNS(rs.fields("CatalogueNum"))
            rsOUt.Update
            lngRecordsExported = lngRecordsExported + 1
        End If
        lngLastCopyID = rs.fields("PI_ID")
        rs.MoveNext
    Loop
    
'4. Export EXP2 as delimited file
    Set oTextFile = New z_TextFile
    If Len(pFilename) = 0 Then pFilename = "ExportDefault.txt"
    oTextFile.OpenTextFile (pFilename)
    If rs.eof And rs.BOF Then
        oTextFile.WriteToTextFile ("<no records>")
        oTextFile.CloseTextFile
        GoTo EXIT_Handler
    End If
    rsOUt.MoveFirst
    strExport = ""
    Do While Not rsOUt.eof
        strExport = rsOUt.fields("Author") & "~" & rsOUt.fields("Title") & "~" & rsOUt.fields("description") & "~" & rsOUt.fields("priceDollar")
        oTextFile.WriteToTextFile (strExport)
        rsOUt.MoveNext
    Loop
    oTextFile.CloseTextFile
    rs.Close
    rsOUt.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Set rsOUt = Nothing
EXIT_Handler:
    MsgBox CStr(lngRecordsExported) & " records exported.", vbOKOnly, "Finished"

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.ExportToBIBLIO(rs,pFilename,pPriceFactor)", Array(rs, pFilename, pPriceFactor)
End Function
Private Function FormatAntDescription(f1 As Variant, f2 As Variant, f3 As Variant, f4 As Variant, f5 As Variant) As Variant
    On Error GoTo errHandler
Dim strOut As String
    If FNS(CStr(f1)) > "" Then
        strOut = f1
    End If
    If FNS(CStr(f2)) > "" Then
        If Len(strOut) > 0 Then
            strOut = strOut & ". " & f2
        Else
            strOut = f2
        End If
    End If
    If FNS(CStr(f3)) > "" Then
        If Len(strOut) > 0 Then
            strOut = strOut & ". " & f3
        Else
            strOut = f3
        End If
    End If
    If FNS(CStr(f4)) > "" Then
        If Len(strOut) > 0 Then
            strOut = strOut & ". " & f4
        Else
            strOut = f4
        End If
    End If
    If FNS(CStr(f5)) > "" Then
        If Len(strOut) > 0 Then
            strOut = strOut & " " & f5 & "."
        Else
            strOut = f5 & "."
        End If
    End If
    FormatAntDescription = strOut
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.FormatAntDescription(f1,f2,f3,f4,f5)", Array(f1, f2, f3, f4, f5)
End Function
Public Function ExportWantsListToABE(pFilename As String, pSince As Date) As Long
    On Error GoTo errHandler
    Dim cmd As ADODB.Command
    Dim Param As ADODB.Parameter
    Dim rs As ADODB.Recordset
    Dim rsOUt As ADODB.Recordset
    Dim oBatch As New z_SQL
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = oPC.COShort
    Dim v1, v2, v3, v4, v5, v6, t1, t2, a1, d1, h1, n1, c1, ds
    Dim oTextFile As z_TextFile
    Dim strExport As String
    Dim lngLastCopyID, lngRecordsExported As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
 '1. Delete the existing temp file
    oPC.COShort.execute "DROP TABLE EXP2"
 '2. Create a empty table with the necessary field types
    oPC.COShort.execute "CREATE TABLE [EXP2] ( [title] TEXT,[Author] TEXT);"
 '3. Retrieve raw data from tables and join fields in appropriate ways
    Set rs = New ADODB.Recordset
    rs.open "Select P_Title, P_MainAuthor FROM tPRODUCT WHERE P_DateRecordAdded > '" & Format(pSince, "yyyy-mm-dd") & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
    
'4. Export EXP2 as delimited file
    Set oTextFile = New z_TextFile
    If Len(pFilename) = 0 Then pFilename = "ExportDefault.txt"
    oTextFile.OpenTextFile (pFilename)
'    Set rs = New ADODB.Recordset
'    rs.Open "EXP2", oPC.COShort, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then rs.MoveFirst
    strExport = ""
    lngRecordsExported = 0
    Do While Not rs.eof
        strExport = rs.fields("P_Title") & "~" & rs.fields("P_MainAuthor")
        oTextFile.WriteToTextFile (strExport)
        oTextFile.WriteToTextFile ("")
        lngRecordsExported = lngRecordsExported + 1
        rs.MoveNext
    Loop
    oTextFile.CloseTextFile
    rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    MsgBox CStr(lngRecordsExported) & " records exported.", vbOKOnly, "Finished"

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.ExportWantsListToABE(pFilename,pSince)", Array(pFilename, pSince)
End Function
Public Function ExportWantsListToBib(pFilename As String) As Long
    On Error GoTo errHandler
    Dim cmd As ADODB.Command
    Dim Param As ADODB.Parameter
    Dim rs As ADODB.Recordset
    Dim rsOUt As ADODB.Recordset
    Dim oBatch As New z_SQL
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = oPC.COShort
    Dim v1, v2, v3, v4, v5, v6, t1, t2, a1, d1, h1, n1, c1, ds
    Dim oTextFile As z_TextFile
    Dim strExport As String
    Dim lngLastCopyID, lngRecordsExported As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
'1. Delete the existing temp file
    cmd.CommandText = "DROP TABLE EXP2"
    cmd.CommandType = adCmdText
    cmd.execute
 '2. Create a empty table with the necessary field types
    cmd.CommandText = "CREATE TABLE [EXP2] ( [Author] TEXT,[title] MEMO);"
    cmd.execute
 '3. Retrieve raw data from tables and join fields in appropriate ways
 '   Set rs = New ADODB.Recordset
 '   retval = oBatch.RunProc("EX_Bookset_2", Array(pBooksetID), "")
    rs.open "Select * FROM vGetExportData1 ORDER BY CATALPI_ID DESC", oPC, adOpenKeyset, adLockOptimistic
'4. Export EXP2 as delimited file
    Set oTextFile = New z_TextFile
    If Len(pFilename) = 0 Then pFilename = "ExportDefault.txt"
    oTextFile.OpenTextFile (pFilename)
    Set rsOUt = New ADODB.Recordset
    rsOUt.open "EXP2", oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rsOUt.eof Then rsOUt.MoveFirst
    strExport = ""
    lngRecordsExported = 0
    Do While Not rsOUt.eof
        strExport = rsOUt.fields("Author") & "~" & rsOUt.fields("Title")
        oTextFile.WriteToTextFile (strExport)
        oTextFile.WriteToTextFile ("")
        lngRecordsExported = lngRecordsExported + 1
        rsOUt.MoveNext
    Loop
    oTextFile.CloseTextFile
    rsOUt.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    MsgBox CStr(lngRecordsExported) & " records exported.", vbOKOnly, "Finished"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.ExportWantsListToBib(pFilename)", pFilename
End Function

Public Function ExportToID(pFilename As String)
    On Error GoTo errHandler
    Dim cmd As ADODB.Command
    Dim Param As ADODB.Parameter
    Dim rs As ADODB.Recordset
    Dim rsOUt As ADODB.Recordset
    Dim oBatch As New z_SQL
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = oPC.COShort
    Dim v1, v2, v3, v4, v5, v6, t1, t2, a1, d1, h1, n1, c1, ds
    Dim oTextFile As z_TextFile
    Dim strExport As String
    Dim lngLastCopyID, lngRecordsExported As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
 '1. Delete the existing temp file
    oBatch.DropTable "EXP", ""
 '2. Create a empty table with the necessary field types
    oBatch.RunProc "GetBooksForIDBooks", Array(), ""
'4. Export EXP as delimited file
    Set oTextFile = New z_TextFile
    If Len(pFilename) = 0 Then pFilename = "ExportIDBooks.txt"
    oTextFile.OpenTextFile (pFilename)
    Set rsOUt = New ADODB.Recordset
    rsOUt.open "EXP", oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rsOUt.eof Then rsOUt.MoveFirst
    strExport = ""
    lngRecordsExported = 0
    Do While Not rsOUt.eof
        With rsOUt
        strExport = .fields("Code") & "~" & .fields("Author") & "~" & .fields("Title") & "~" & .fields("qoh") & "~" & .fields("Edition") & "~" & .fields("Pubdate") & "~" & .fields("Bindingcode") & "~" & .fields("Price") & "~" & .fields("Publisher")
        End With
        oTextFile.WriteToTextFile (strExport)
        oTextFile.WriteToTextFile ("")
        lngRecordsExported = lngRecordsExported + 1
        rsOUt.MoveNext
    Loop
    oTextFile.CloseTextFile
    rsOUt.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    MsgBox CStr(lngRecordsExported) & " records exported.", vbOKOnly, "Finished"

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Export.ExportToID(pFilename)", pFilename
End Function
