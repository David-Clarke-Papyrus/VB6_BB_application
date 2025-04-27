VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim FTP1 As New FTPClass
Dim strLog As String
Dim res As Boolean
Dim arStores() As String
Public Function ConnectToFTP()
    On Error GoTo ErrHandler
    
    Set fINET = New wininet
 '   If Not fINET.IsNetConnectOnline Then
 '   On Error Resume Next
     '   lngResult = fINET.StartDUN(0, oPC.DUN, True)
       ' Check lngResult = 0, ERR_DUNALREADYOPEN, "Cannot open connection,perhaps it is already open"
    On Error GoTo ErrHandler
    
    SetupCentralFTP

    Exit Function
ErrHandler:
    ErrPreserve
    On Error Resume Next
    fINET.Hangup
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.ManageLoyaltyImport"
End Function
Function InStoreArray(pName As String) As Long
Dim i As Long

    For i = 1 To UBound(arStores)
        If arStores(1, i) = pName Then
            InStoreArray = i
            Exit For
        End If
    Next
End Function
Public Sub FetchLoyaltyFiles(pMsg As String)
On Error GoTo ErrHandler
Dim oFSO As New FileSystemObject
Dim Fold
Dim res As Boolean
Dim strFolder As String
Dim f As FTPFileClass
Dim strFrom As String
Dim iPosInArray As Long
Dim iLevel As Long
Dim strOut As String
Dim iTotal As Long
'
    iLevel = 1
    ReDim Preserve arStores(2, oPC.Configuration.Stores.Count)
    For Each f In FTP1.files
        If oFSO.FileExists(oPC.SharedFolderRoot & "\Data\Loyalty\UP\" & f.FileName) Then
            oFSO.DeleteFile oPC.SharedFolderRoot & "\Data\Loyalty\UP\" & f.FileName, True
        End If
        strFrom = ExtractStoreCodeFromFilename(f.FileName)
        iPosInArray = InStoreArray(strFrom)
        If iPosInArray > 0 Then
            arStores(2, iPosInArray) = arStores(2, iPosInArray) + 1
        Else
            arStores(1, iLevel) = strFrom
            arStores(2, iLevel) = 1
            iLevel = iLevel + 1
        End If
        If Not oPC.Configuration.Stores.FindStoreByCode(strFrom) Is Nothing Then
            res = FTP1.GetFile(f.FileName, oPC.SharedFolderRoot & "\Data\Loyalty\UP\" & f.FileName, True)    ', EXC_GENERAL, "Fail in fetching file"
            If res = True Then
                res = FTP1.DeleteFile(f.FileName)
            End If
        End If
    Next
    strOut = ""
    iTotal = 0
    For iPosInArray = 1 To iLevel - 1
        strOut = strOut & arStores(1, iPosInArray) & ": " & arStores(2, iPosInArray) & Chr(13)
        iTotal = iTotal + CLng(arStores(2, iPosInArray))
    Next
    
    strOut = strOut & "Total: " & CStr(iTotal)
    pMsg = strOut
    Exit Sub

ErrHandler:
    ErrPreserve
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.FetchLoyaltyFiles"
End Sub


Public Sub SetupCentralFTP()
    On Error GoTo ErrHandler
Dim res As Boolean
    res = FTP1.OpenFTP(oPC.CentralFTPAddress, oPC.CentralFTPUsername, oPC.CentralFTPPassword)
  '  If oPC.CentralFTPFolder > "" Then
        Check FTP1.SetCurrentFolder("/LOYALTY/UP"), EXC_GENERAL, "Error setting FTP folder"
  '  End If
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Export.SetupCentralFTP"
End Sub

Public Sub CloseFTP()
    FTP1.CloseFTP
End Sub
Private Function ExtractStoreCodeFromFilename(pFilename As String) As String
Dim i As Integer
Dim strCode As String
Dim c As String

    strCode = ""
    For i = 1 To Len(pFilename)
        c = UCase(Mid(pFilename, i, 1))
        If IsAlphaCaps(c) Then
            strCode = strCode & c
        Else
            Exit For
        End If
    Next
    ExtractStoreCodeFromFilename = strCode
End Function
Public Function UpdateDBFromFiles(pLog As String) As Long
    On Error GoTo ErrHandler
Dim oFSO As New FileSystemObject
Dim f, fc, fol
Dim ZIP
Dim strFrom As String
Dim strType As String
Dim iRecCount As Long

    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\Data\Loyalty")
    Set fc = fol.files
    
    For Each f In fc
        f.Delete
    Next
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\Data\Loyalty\UP")
    Set fc = fol.files
    
    For Each f In fc
        If InStr(1, f.Name, "_SA") > 0 Or UCase(Right(f.Name, 4)) = ".CUS" Then
            f.Copy oPC.SharedFolderRoot & "\Data\Loyalty\", True
        ElseIf UCase(Right(f.Name, 4)) = ".ZIP" Then
            Set ZIP = CreateObject("FathZIP.FathZIPCtrl.1")
            ZIP.OpenZip (f.Path)
            ZIP.BasePath = oPC.SharedFolderRoot & "\Data\Loyalty\"
            ZIP.ExtractFile ("*.*")
            ZIP.Close
            Set ZIP = Nothing
        End If
    Next
    
'
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\Data\Loyalty")
    Set fc = fol.files
    For Each f In fc
        strType = Right(f.Name, 3)
        If UCase(strType) = "CUS" Then   'First we handle the new and updated customers
   '         strFrom = Left(F.Name, InStr(1, F.Name, "_") - 1)
            strFrom = ExtractStoreCodeFromFilename(f.Name)
            If Not oPC.Configuration.Stores.FindStoreByCode(strFrom) Is Nothing Then
                If UCase(oPC.Configuration.Stores.FindStoreByCode(strFrom).SystemName) = "PAPYRUS" Then
                    iRecCount = UpdateCustomerFromFile_Papyrus(f.Path, strFrom)
                ElseIf UCase(oPC.Configuration.Stores.FindStoreByCode(strFrom).SystemName) = "WORDSTOCK" Then
                    iRecCount = UpdateCustomerFromFile_Wordstock(f.Path, strFrom)
                End If
            End If
            pLog = pLog & f.Name & "  :  " & iRecCount & vbCrLf
        End If
    Next
'
    For Each f In fc
        strType = Right(f.Name, 3)
        If UCase(strType) = "ZIP" Then   'Then the sales for loyalty customers
            strFrom = ExtractStoreCodeFromFilename(f.Name)
          ' If UCase(oPC.Configuration.Stores.FindStoreByCode(strFrom).SystemName) = "PAPYRUS" Then
                If UpdateCustomerSalesFromFile_Papyrus(f.Path, strFrom) <> 0 Then
                    MsgBox "Trying to insert duplicate rows"
                End If
          '  ElseIf UCase(oPC.Configuration.Stores.FindStoreByCode(strFrom).SystemName) = "WORDSTOCK" Then
          '      UpdateCustomerSalesFromFile_Wordstock f.Path, strFrom
          '  End If
            pLog = pLog & f.Name & vbCrLf
        End If
    Next
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "Exec sp_UpdateSales " & oPC.Configuration.DefaultCurrencyDivisor
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.UpdateDBFromFiles"
End Function
Private Function UpdateCustomerFromFile_Wordstock(pFilePath As String, pStoreCode As String) As Long
    On Error GoTo ErrHandler
Dim oTF As New z_TextFile
Dim strLine As String
Dim ar() As String
Dim cmd As ADODB.Command
Dim oLC As New z_LCManager
Dim lngTPID As Long
Dim arAdd() As String
Dim i, j As Integer
Dim fso As New FileSystemObject
Dim strStoreCode As String
Dim strACNo As String
Dim iRecCount As Long
    iRecCount = 0
    strStoreCode = pStoreCode
    oTF.OpenTextFileToRead pFilePath
    Do While Not oTF.IsEOF
        ReDim arAdd(6)
        ReDim ar(25)
        strLine = oTF.ReadLinefromTextFile
        If Trim(strLine) = "" Then GoTo Tail
     '   If oTF.IsEOF Then Exit Do
        ar = Split(strLine, "|")

        i = 0
        j = 0
'        For i = 0 To 4  'skip blank input addres lines
'            If Trim(ar(i + 5)) > "" Then
'                arAdd(j) = Trim(ar(i + 5))
'                j = j + 1
'            End If
'        Next
        ar(0) = Trim(ar(0))
        If IsNumeric(ar(0)) Then
            strACNo = CStr(CLng(ar(0)))
        Else
            strACNo = ""
        End If
      '  If ar(1) = "SONNENBERG" Then MsgBox "here"
        oLC.InsertLoyaltyCustomer lngTPID, strACNo, Trim(ar(1)), Trim(ar(2)), Trim(ar(3)), Trim(ar(4)), _
        CreateAddressee(Trim(ar(1)), Trim(ar(3)), Trim(ar(2)), Trim(ar(4))), Trim(ar(5)), Trim(ar(6)), "", "", Trim(ar(7)), Trim(ar(8)), _
        Trim(ar(9)), Trim(ar(10)), Trim(ar(11)), Trim(ar(12)), Trim(ar(13)), Trim(ar(14)), _
        Mid(Trim(ar(16)), 2, 1) = "Y", Mid(Trim(ar(16)), 1, 1) = "Y", Mid(Trim(ar(16)), 3, 1) = "Y", strStoreCode, Trim(ar(23))
        iRecCount = iRecCount + 1
Tail:
    Loop
    oTF.CloseTextFile
    'Delete file
    UpdateCustomerFromFile_Wordstock = iRecCount
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.UpdateCustomerFromFile(pFilePath)", pFilePath
End Function
Private Function UpdateCustomerFromFile_Papyrus(pFilePath As String, pStoreCode As String)
    On Error GoTo ErrHandler
Dim oTF As New z_TextFile
Dim strLine As String
Dim ar() As String
Dim cmd As ADODB.Command
Dim oLC As New z_LCManager
Dim lngTPID As Long
Dim arAdd() As String
Dim i, j As Integer
Dim fso As New FileSystemObject
Dim strStoreCode As String
Dim iRecCount As Long
    iRecCount = 0
    oTF.OpenTextFileToRead pFilePath
    strStoreCode = pStoreCode
    Do While Not oTF.IsEOF
        ReDim arAdd(6)
        ReDim ar(26)
        strLine = oTF.ReadLinefromTextFile
        If strLine = "" Then Exit Do
        If InStr(1, strLine, vbTab) > 0 Then
            ar = Split(strLine, vbTab)
        Else
            ar = Split(strLine, "|")
        End If

        i = 0
        j = 0
'        For i = 0 To 3  'skip blank input addres2 lines
'            If Trim(ar(i + 4)) > "" Then
'                arAdd(j) = Trim(ar(i + 4))
'                j = j + 1
'            End If
'        Next
        oLC.InsertLoyaltyCustomer lngTPID, Trim(ar(17)), Trim(ar(0)), Trim(ar(2)), Trim(ar(1)), "", _
        CreateAddressee(Trim(ar(1)), Trim(ar(0)), Trim(ar(2)), ""), Trim(ar(4)), Trim(ar(5)), "", "", Trim(ar(6)), Trim(ar(7)), _
       Trim(ar(22)), Trim(ar(8)), Trim(ar(10)), Trim(ar(11)), Trim(ar(3)), Trim(ar(13)), _
        Mid(Trim(ar(20)), 1, 1) = "1", Mid(Trim(ar(18)), 1, 1) = "1", Mid(Trim(ar(19)), 1, 1) = "1", strStoreCode, ""
        iRecCount = iRecCount + 1
    Loop
    oTF.CloseTextFile
    'Delete file
    UpdateCustomerFromFile_Papyrus = iRecCount
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.UpdateCustomerFromFile(pFilePath)", pFilePath
End Function

Private Function UpdateCustomerSalesFromFile_Papyrus(pFilePath As String, pBranchCode As String) As Long
    On Error GoTo ErrHandler
Dim oTF As New z_TextFile
Dim strLine As String
Dim ar() As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

Dim oLC As New z_LCManager
Dim lngTPID As Long
Dim arAdd() As String
Dim ofs As New FileSystemObject
Dim strNewFIlename As String
Dim ZIP
Dim strStoreCode As String
Dim dteFileDate As Date
Dim strSQL As String
Dim lngTimeOut As Long
Dim lngRecs As Long
Dim oBF As zc_BF

'We expect a ZIP file, unzip into folder
        If UCase(Right(pFilePath, 4)) = ".ZIP" Then
            Set ZIP = CreateObject("FathZIP.FathZIPCtrl.1")
            ZIP.OpenZip (pFilePath)
            ZIP.BasePath = oPC.SharedFolderRoot & "\Data\Loyalty\"
            ZIP.ExtractFile ("*.*")
            ZIP.Close
            Set ZIP = Nothing
        End If
        
        dteFileDate = ExtractDateFromFilename(pFilePath)
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        lngTimeOut = oPC.COShort.CommandTimeout
        oPC.COShort.CommandTimeout = 0
        
  '  oPC.COShort.BeginTrans
        
        strSQL = "DELETE FROM dbo.tmpEXCHANGE"
        oPC.COShort.execute strSQL, lngRecs
        strSQL = "DELETE FROM dbo.tmpCSL"
        oPC.COShort.execute strSQL, lngRecs
        strSQL = "DELETE FROM dbo.tmpPAYMENT"
        oPC.COShort.execute strSQL, lngRecs
        
        strSQL = "BULK INSERT  " & oPC.CatalogName & ".dbo.tmpEXCHANGE From '" & oPC.SharedFolderRoot & "\Data\Loyalty\Exch.txt'" _
        & " WITH (FIELDTERMINATOR = '\t', ROWTERMINATOR = '\n',KEEPNULLS)"
        oPC.COShort.execute strSQL
        
        On Error Resume Next
        oPC.COShort.execute "ALTER TABLE tmpCSL DROP COLUMN PID"
        On Error GoTo ErrHandler
        strSQL = "BULK INSERT " & oPC.CatalogName & ".dbo.tmpCSL From '" & oPC.SharedFolderRoot & "\Data\Loyalty\CSL.txt'" _
        & " WITH (FIELDTERMINATOR = '\t', ROWTERMINATOR = '\n')"
        oPC.COShort.execute strSQL
        oPC.COShort.execute "ALTER TABLE tmpCSL ADD PID UNIQUEIDENTIFIER NULL"
        
        strSQL = "BULK INSERT " & oPC.CatalogName & ".dbo.tmpPAYMENT From '" & oPC.SharedFolderRoot & "\Data\Loyalty\Pay.txt'" _
        & " WITH (FIELDTERMINATOR = '\t', ROWTERMINATOR = '\n')"
        oPC.COShort.execute strSQL
        
        oPC.COShort.execute "DROP Table tmpP"
        strSQL = "SELECT DISTINCT CAST(CODE AS VARCHAR(20)) as CODE,CAST(EAN AS VARCHAR(15)) as EAN,CAST(NULL as UNIQUEIDENTIFIER) as PID  INTO tmpP FROM tmpCSL "
        oPC.COShort.execute strSQL
        
        oPC.COShort.execute "EXEC dbo.sp_UpdateCSL_1"
        
        Set oBF = New zc_BF
        oBF.InsertFromBookfind
        Set oBF = Nothing
        
        oPC.COShort.execute "EXEC dbo.sp_UpdateCSL_1"
        
        oPC.COShort.execute "EXEC dbo.InsertRogueTitles"
        
        oPC.COShort.execute "EXEC dbo.sp_UpdateCSL_1"
        oPC.COShort.execute "EXEC dbo.sp_UpdateCSL_2"

        Set cmd = New ADODB.Command
        cmd.CommandText = "InsertSalesToTables"
        cmd.CommandType = adCmdStoredProc
    
        Set par = cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 5)
        par.Value = pBranchCode
        cmd.Parameters.Append par
        Set par = Nothing
        
        Set par = cmd.CreateParameter("@Res", adInteger, adParamInputOutput)
        par.Value = 0
        cmd.Parameters.Append par
    
        cmd.ActiveConnection = oPC.COShort
        cmd.execute
    
        lngResult = cmd.Parameters(1)
        Set cmd = Nothing
        
        On Error GoTo ErrHandler
        
        UpdateCustomerSalesFromFile_Papyrus = lngResult
        oPC.COShort.CommandTimeout = lngTimeOut
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.UpdateCustomerSalesFromFile_Papyrus(pFilePath)", pFilePath, EA_DFTRBK
End Function
Private Sub UpdateCustomerSalesFromFile_Wordstock(pFilePath As String, pBranchCode As String)
    On Error GoTo ErrHandler
Dim oTF As New z_TextFile
Dim strLine As String
Dim ar() As String
Dim cmd As ADODB.Command
Dim oLC As New z_LCManager
Dim lngTPID As Long
Dim arAdd() As String
Dim i, j As Integer
Dim strACNo As String
Dim strCode As String
Dim lngTrackingNo As Long
Dim lngRegisterNumber As Long
Dim lngPrice As Long
Dim lngSP As Long
Dim lngQty As Long
Dim dteDate As Date
Dim lngSuccess As Long
Dim lngResult As Long
Dim oProduct As a_Product
Dim ofs As New FileSystemObject
Dim strNewFIlename As String

    'replaces CHR(10) with vbCRLF so file can be read line by line
    oTF.OpenTextFileToRead pFilePath
    strLine = oTF.ReadLinefromTextFile
    
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Do While Not oTF.IsEOF
        ReDim arAdd(6)
        ReDim ar(20)
        strLine = oTF.ReadLinefromTextFile
        If oTF.IsEOF Then Exit Do
        ar = Split(strLine, "|")
        If Trim(ar(4)) <> "CR" And Trim(ar(4)) <> "GC" And Trim(ar(4)) <> "ENROLL" Then
             strACNo = CStr(CLng(Trim(ar(0))))
             lngTrackingNo = CLng(Trim(ar(3)))
             
            '' lngRegisterNumber = CLng(Trim(ar(4)))
             strCode = Trim(ar(4))
             dteDate = CDate(Trim(ar(5)))
             lngSP = CDbl(Trim(ar(7))) * oPC.Configuration.DefaultCurrencyDivisor
             lngPrice = CDbl(Trim(ar(8))) * oPC.Configuration.DefaultCurrencyDivisor
             lngQty = CLng(Trim(ar(9)))
             oLC.InsertLoyaltyCustomerSale strACNo, strCode, lngSP, lngPrice, lngQty, dteDate, lngRegisterNumber, lngTrackingNo, lngSuccess
             If lngSuccess <> 0 Then   'UnSuccessful
                 Set oProduct = New a_Product
                 lngResult = oProduct.Load(0, 0, strCode) 'loads the record from BF if not on database
                 If lngResult = 0 Then
                     oLC.InsertLoyaltyCustomerSale strACNo, strCode, lngSP, lngPrice, lngQty, dteDate, lngRegisterNumber, lngTrackingNo, lngSuccess
                 End If
                 If lngSuccess <> 0 Or lngResult <> 0 Then
                 'add to failures file for inspection
                '     oPC.CO.execute "CREATE TABLE tRETRY (ACNO VARCHAR(10),CDE VARCHAR(15),SP INTEGER,PRICE INTEGER,QTY INTEGER,DTE DATETIME,REGNO INTEGER,TRNO INTEGER)"
                     oPC.COShort.execute "INSERT INTO tRETRY (ACNO,CDE,SP,PRICE,QTY,DTE,REGNO,TRNO) VALUES ('" & strACNo & "','" & strCode & "'," & lngSP & "," & lngPrice & "," & lngQty & ",'" & ReverseDateTime(dteDate) & "'," & lngRegisterNumber & "," & lngTrackingNo & ")"
                 End If
             End If
        End If
    Loop
    oTF.CloseTextFile
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.UpdateCustomerSalesFromFile_Wordstock(pFilePath)", pFilePath
End Sub

Public Function ConfirmReceipt()
Dim oFSO As New FileSystemObject
Dim f, fc, fol
Dim strConfirm As String
Dim oTF As New z_TextFile
Dim dte As Date
Dim strFilename As String

    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\Data\Loyalty\UP")
    Set fc = fol.files
    strConfirm = ""
    For Each f In fc
        If UCase(Right(f.Name, 4)) = ".ZIP" Or UCase(Right(f.Name, 4)) = ".CUS" Or UCase(Right(f.Name, 4)) = ".SAL" Then
            strConfirm = strConfirm & IIf(Len(strConfirm) > 0, ",", "") & f.Name
        End If
    Next
    dte = Date
    strFilename = "LCR" & ReverseDateCompact(dte) & ".txt"
    oTF.OpenTextFile oPC.SharedFolderRoot & "\Data\Loyalty\DOWN\" & strFilename
    oTF.WriteToTextFile strConfirm
    oTF.CloseTextFile
    If Not FTP1.FolderExists(oPC.CentralFTPFolder & "/LOYALTY/DOWN*") Then
        Check FTP1.CreatFolder(oPC.CentralFTPFolder & "/LOYALTY/DOWN/"), EXC_GENERAL, "Error creating folder:ConfirmReceipt"
    End If
    Check FTP1.SetCurrentFolder(oPC.CentralFTPFolder & "/LOYALTY/DOWN/"), EXC_GENERAL, "Error setting FTP folder"
    If oFSO.FileExists(oPC.SharedFolderRoot & "\Data\Loyalty\DOWN\" & strFilename) Then
        res = FTP1.PutFile(oPC.SharedFolderRoot & "\Data\Loyalty\DOWN\" & strFilename, strFilename, True)  ', EXC_GENERAL, "Fail in putting file"
        If res = False Then
            strLog = "Error putting receiptfile in FTP site"
        End If
    End If
End Function

Public Sub Hangup()
    On Error Resume Next
    fINET.Hangup

End Sub
Public Function AppendEditedCustomers(pTextFilename As String, pFrom As Date, Optional pTo As Date, Optional pAll As Boolean) As Long
    On Error GoTo ErrHandler
Dim strCommand As String
Dim oFSO As New FileSystemObject
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim lngCount As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.ActiveConnection = oPC.COShort
    
    If pAll = False Then
        If CDate(pTo) = 0 Then pTo = pFrom
        strSQL = "SELECT * FROM PBKSC.dbo.vLCustomers " _
        & " WHERE (TP_DateLastModified >= '" & ReverseDate(pFrom) & "' AND TP_DateLastModified <='" & ReverseDate(DateAdd("d", 1, pTo)) & "')" _
        & " OR (TP_DateRecordAdded >= '" & ReverseDate(pFrom) & "' AND TP_DateRecordAdded <='" & ReverseDate(DateAdd("d", 1, pTo)) & "')"
      '  MsgBox strSQL
    Else
        strSQL = "SELECT * FROM PBKSC.dbo.vLCustomers "
    End If
    
    rs.Open strSQL, , adOpenForwardOnly
    lngCount = rs.RecordCount
    MsgBox "Count = " & lngCount
    rs.Close
    Set rs = Nothing
    
    strCommand = "bcp """ & strSQL & """ queryout """ & "\PBKS\data\loyalty\Edited\" & pTextFilename & """ -eBCPError.txt -c -q -r\r -Usa -P" & oPC.Password & " -S " & oPC.ServerName
'        MsgBox strCommand
'    oPC.COShort.execute "Exec Master..xp_Cmdshell " & strCommand
    ShellandWaitold strCommand, vbHide, False
    
    AppendEditedCustomers = lngCount

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Loyalty.AppendNewCustomers"
    Exit Function
ErrHandler:
    ErrorIn "z_Import.AppendEditedCustomers(pTextFilename,pFrom,pTo,pAll)", Array(pTextFilename, _
         pFrom, pTo, pAll)
End Function

Private Function ExtractDateFromFilename(pFN As String, Optional pStart As Integer) As Date
Dim strDateTime As String
Dim iStart As Integer
Dim iEnd As Integer

    If pStart > 0 Then
        iEnd = pStart
    Else
        iEnd = InStrRev(pFN, ".") '- 12
    End If
    iStart = iEnd - 12
    strDateTime = Mid(pFN, iStart, 4) & "/" & Mid(pFN, iStart + 4, 2) & "/" & Mid(pFN, iStart + 6, 2)
    If IsDate(strDateTime) Then
        ExtractDateFromFilename = CDate(strDateTime)
    Else
        ExtractDateFromFilename = CDate(0)
    End If
End Function
Public Sub SendLoyaltyEditingChangesToBranches()
    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "ExportLoyaltyEditingChanges"
    cmd.CommandType = adCmdStoredProc
    
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@NominalDate", adDate, adParamInput, , pNominalDate)
'    cmd.Parameters.Append prm
'
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@OperatorID", adInteger, adParamInput, , pOperatorID)
'    cmd.Parameters.Append prm
'
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@TypeID", adInteger, adParamInput, , pTypeID)
'    cmd.Parameters.Append prm
'
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@OPID", adInteger, adParamOutput)
'    cmd.Parameters.Append prm

    
    cmd.execute lngRecordsAffected
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.SendLoyaltyEditingChangesToBranches"
End Sub

Public Sub ImportCSVAcnoList(pFilename As String)
    On Error GoTo ErrHandler
Dim strSQL As String
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

        oPC.COShort.execute "PrepareTmpAcno"
        
        strSQL = "BULK INSERT  " & oPC.CatalogName & ".dbo.tTempAcno From '" & pFilename _
        & "' WITH (FIELDTERMINATOR = '\t', ROWTERMINATOR = '\n',KEEPNULLS)"
        oPC.COShort.execute strSQL
        oPC.COShort.execute "TRUNCATE TABLE tS"
        strSQL = "INSERT INTO tS SELECT TP_ID FROM tTEMPAcno a Join tTP b ON a.ACNO = b.TP_ACNO "
        oPC.COShort.execute strSQL
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Import.ImportCSVAcnoList(pFilename)", pFilename
End Sub


