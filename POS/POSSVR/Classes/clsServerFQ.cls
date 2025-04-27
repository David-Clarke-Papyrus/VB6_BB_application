VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_PollingServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event FileProcessed(FileType As Integer) '0 = Other, 1 = Sale, 2 = Update
Dim bDBOK As Boolean
Dim bFoldersOK As Boolean
Dim sOutBoxRoot As String
Dim sRoot As String
Dim sInBox As String
Dim sInBoxSharedName As String



Public Function CheckFolders()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim oFShare As New Z_ManageFolders
Dim sSharedName As String
    
    CheckFolders = False

    sRoot = oPC.SharedFolderRoot
    sInBox = sRoot & "\POSSVR_IN"
    sOutBoxRoot = sRoot & "\POSSVR_OUT"
    
    sSharedName = "POSSVR_IN_S"
    
    If Not fs.FolderExists(sInBox) Then
        If Not oFShare.CreateFolder(sInBox, True, sSharedName) Then Exit Function
    Else
        oFShare.SetShared oFShare.GetCompName, sSharedName, sInBox
    End If
    If Not fs.FolderExists(sOutBoxRoot) Then
        If Not oFShare.CreateFolder(sOutBoxRoot, False) Then Exit Function
    End If
    sInBoxSharedName = "\\" & oFShare.GetCompName & "\" & sSharedName
    CheckFolders = True

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.CheckFolders"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CheckFolders"
End Function
Public Property Get SharedInboxName() As String
    On Error GoTo errHandler
    SharedInboxName = sInBoxSharedName
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.SharedInboxName"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.SharedInboxName"
End Property

Public Property Get FoldersOK() As Boolean
    On Error GoTo errHandler
    FoldersOK = bFoldersOK
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.FoldersOK"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.FoldersOK"
End Property
Public Property Get DBConnectionOK() As Boolean
    On Error GoTo errHandler
    DBConnectionOK = bDBOK
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.DBConnectionOK"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.DBConnectionOK"
End Property
Public Property Get ClientListCount() As Integer
    On Error GoTo errHandler
    If ClientList(UBound(ClientList)).Name <> "" Then
        ClientListCount = UBound(ClientList)
    Else
        ClientListCount = -1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.ClientListCount"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientListCount"
End Property

Public Function ClientName(Index As Integer) As String
    On Error GoTo errHandler
    If Index <= UBound(ClientList) Then
        ClientName = ClientList(Index).Name
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.ClientName(Index)", Index
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientName(Index)", Index
End Function
Public Function ClientPathName(Index As Integer) As String
    On Error GoTo errHandler
    If Index <= UBound(ClientList) Then
        ClientPathName = ClientList(Index).pathname
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.ClientPathName(Index)", Index
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientPathName(Index)", Index
End Function

Public Function LoadClientList() As Integer
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim i As Integer
    
    sSQL = "SELECT * FROM tPOSClient"
    ReDim ClientList(0)
    rs.Open sSQL, oPC.CO, adOpenDynamic, adLockReadOnly
    Do While Not rs.EOF
        If ClientList(UBound(ClientList)).Name <> "" Then
            ReDim Preserve ClientList(UBound(ClientList) + 1)
        End If
        With ClientList(UBound(ClientList))
            If Not IsNull(rs!TillCode) Then .Name = rs!TillCode
            If Not IsNull(rs!pathname) Then .pathname = rs!pathname
            If Not IsNull(rs!OutBox) Then .OutBox = rs!OutBox
        End With
        i = i + 1
        rs.MoveNext
    Loop
    LoadClientList = i
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.LoadClientList"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.LoadClientList"
End Function

Public Function PutMasterDataInOutbox() As Boolean
    On Error GoTo errHandler
Dim rs1 As ADODB.Recordset
Dim sSQL As String
Dim sTable As String
Dim i As Integer
    Set rs1 = New ADODB.Recordset
    For i = 1 To 4
        oPC.CO.BeginTrans
        If i = 1 Then
          ' check if Product records are new / changed
          sTable = "tProdUpdates"
        ElseIf i = 2 Then
          ' check if any StaffMember records are new / changed
          sTable = "tStaffMemberUpdate"
        ElseIf i = 3 Then
          ' check if any Customer records are new / changed
          sTable = "tTPUpdate_Cust"
        ElseIf i = 4 Then
          ' check if any Customer records are new / changed
          sTable = "tCOUpdate"
        End If
        sSQL = "SELECT * FROM " & sTable
        rs1.CursorLocation = adUseClient
        rs1.CursorType = adOpenStatic
        rs1.LockType = adLockBatchOptimistic
        rs1.Open sSQL, oPC.CO
        rs1.ActiveConnection = Nothing         'DISCONNECT RECORDSET
      
        If Not rs1.EOF Then
            Check PrepareMasterDataSet(rs1, sTable), EXC_MASTERDATASET, "Problem preparing master data set"
            sSQL = "DELETE " & sTable
            oPC.CO.CommandTimeout = 0
            oPC.CO.Execute sSQL
            PutMasterDataInOutbox = True
            RaiseEvent FileProcessed(2)
        End If
        If rs1.State = adStateOpen Then rs1.Close
        oPC.CO.CommitTrans
    Next i
MEX:
    If rs1.State = adStateOpen Then rs1.Close
    Set rs1 = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.TransmitMasterData"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.PutMasterDataInOutbox"
End Function

Public Function PrepareMasterDataSet(rs As ADODB.Recordset, sTable As String) As Boolean
    On Error GoTo errHandler
Dim i As Integer
Dim oRS As ADODB.Recordset
Dim fs As FileSystemObject
Dim sFile As String

    Err.Clear
    Set fs = New FileSystemObject
    Set oRS = rs
    oRS.MoveFirst
    
    If sTable = "tProdUpdates" Then
      sFile = oRS!PRU_GUID & ".upd"
    ElseIf sTable = "tStaffMemberUpdate" Then
      sFile = oRS!SMU_GUID & ".usd"
    ElseIf sTable = "tTPUpdate_Cust" Then
      sFile = oRS!CU_GUID & ".ucd"
    ElseIf sTable = "tCOUpdate" Then
      sFile = CStr(oRS!COU_COLID) & ".uod"
    End If
    oRS.Save sOutBoxRoot & "\" & sFile, adPersistADTG
    For i = 0 To UBound(ClientList)
        With ClientList(i)
            If Not fs.FolderExists(sOutBoxRoot & "\" & .OutBox) Then
                fs.CreateFolder sOutBoxRoot & "\" & .OutBox
            End If
            fs.CopyFile sOutBoxRoot & "\" & sFile, sOutBoxRoot & "\" & .OutBox & "\"
        End With
    Next i
    fs.DeleteFile sOutBoxRoot & "\" & sFile
    PrepareMasterDataSet = Err = 0
    Err.Clear

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.PrepareUpdate(rs,sTable)", Array(rs, sTable)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.PrepareMasterDataSet(rs,sTable)", Array(rs, sTable)
End Function
Public Sub PollForNewClient()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim F As TextStream
Dim sFile As String
Dim sTillCode As String
Dim sPath As String
Dim xFiles() As String
Dim i As Integer, k As Integer

    ReDim xFiles(0)
    sFile = Dir(sInBox & "\*.*") '"\*_Reg.dat")
    Do While sFile <> ""
        If xFiles(UBound(xFiles)) <> "" Then
            ReDim Preserve xFiles(UBound(xFiles) + 1)
        End If
        xFiles(UBound(xFiles)) = sInBox & "\" & sFile
        sFile = Dir
    Loop
    For k = 0 To UBound(xFiles)
        If UCase(Right(xFiles(k), 7)) = "REG.DAT" Then
            'This is a registration request from a new POS client
            Set F = fs.OpenTextFile(xFiles(k))
            sTillCode = F.ReadLine
            sPath = F.ReadLine
            F.Close
            
            i = AddClient(sTillCode, sPath)
            If i >= 0 Then
                fs.DeleteFile (xFiles(k))
                'Create and send confirmation file
                Set F = fs.CreateTextFile(sOutBoxRoot & "\RegOK.dat")
                F.WriteLine sInBoxSharedName
                F.Close
                fs.MoveFile sOutBoxRoot & "\RegOK.dat", sPath & "\RegOK.dat"
                If Err Then Err.Clear
            End If
            
        ElseIf UCase(Right(xFiles(k), 4)) = ".REQ" Then
            'This is a request to send an empty disconnected Recordset file
            Dim ClientPath As String
            Dim TillCode  As String
            Dim sTmp As String
            
            Set F = fs.OpenTextFile(xFiles(k))
            sTmp = F.ReadLine
            TillCode = F.ReadLine
            ClientPath = F.ReadLine
            F.Close
            If sTmp = "Send RS Data" Then
                If SendRSFile(TillCode, ClientPath) Then
                    fs.DeleteFile xFiles(k), True
                    RaiseEvent FileProcessed(0)
                End If
            End If
        End If
    Next k
        
    Set fs = Nothing
    Set F = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.PollForNewClient"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.PollForNewClient"
End Sub

Private Function ClientExists(sTillCode As String, sPath As String) As String
    On Error GoTo errHandler
Dim i As Integer
    For i = 0 To UBound(ClientList)
        If UCase(sPath) = UCase(ClientList(i).pathname) Then
            ClientExists = "PA"
            Exit For
        ElseIf UCase(sTillCode) = UCase(ClientList(i).Name) Then
            ClientExists = "TC"
            Exit For
        End If
    Next i
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.ClientExists(sTillCode,sPath)", Array(sTillCode, sPath)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClientExists(sTillCode,sPath)", Array(sTillCode, sPath)
End Function

Friend Function AddClient(Name As String, pathname As String) As Integer
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim oFS As New Z_ManageFolders
Dim sRes As String
Dim sOldPath As String

    sRes = ClientExists(Name, pathname)
    If sRes = "TC" Then
        sSQL = "SELECT * FROM tPOSClient WHERE tPOSClient.TillCode ='" & UCase(Name) & "'"
    ElseIf sRes = "PA" Then
        sSQL = "SELECT * FROM tPOSClient WHERE tPOSClient.PathName ='" & UCase(pathname) & "'"
    Else
        sSQL = "SELECT * FROM tPOSClient WHERE tPOSClient.TillCode =''"
    End If
'    ReDim ClientList(0)
    oPC.CO.BeginTrans
    rs.Open sSQL, oPC.CO, adOpenDynamic, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
    End If
    rs!TillCode = UCase(Name)
    If Not IsNull(rs!OutBox) Then sOldPath = rs!OutBox
    rs!pathname = UCase(pathname)
    rs!OutBox = "\" & rs!TillCode & "_OutBox"
    rs!CreateDate = Now
    rs.Update
    
    'delete old path if nessesary
    If sOldPath <> "" And sOldPath <> rs!OutBox Then DeletePath (sOutBoxRoot & sOldPath)
    
    'create new OutBox for new client
    If Not oFS.CreateFolder(sOutBoxRoot & rs!OutBox, False) Then GoTo errHandler
    oPC.CO.CommitTrans
    
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    LoadClientList
'EH:
'
'
'    oPC.CO.RollbackTrans
'    AddClient = -99
'    GoTo MEX
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.AddClient(Name,pathname)", Array(Name, pathname)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.AddClient(Name,pathname)", Array(Name, pathname)
End Function

Private Sub DeletePath(sPath As String)
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim sFile As String

    'check if folder is empty
    sFile = Dir(sPath & "\*.*")
    If sFile <> "" Then
'        if msgbox("The folder:" & vblf & sPath & vblf & "is not empty!" & vblf & _
'                  "Delete it anyway?", vbyesno, "
    Else
        fs.DeleteFolder sPath
    End If
    
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.DeletePath(sPath)", sPath
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.DeletePath(sPath)", sPath
End Sub

Public Sub DeleteClient(ClientName As String)
    On Error GoTo errHandler
Dim i As Integer
Dim sSQL As String
Dim msg As String
    
    For i = 0 To UBound(ClientList)
        If ClientList(i).Name = ClientName Then
            sSQL = "DELETE FROM tPOSClient WHERE tPOSClient.TillCode ='" & ClientName & "'"
            oPC.CO.Execute sSQL
            'reload client list
            i = LoadClientList
            If i = 0 Then
                msg = "Client list is empty!"
            ElseIf i = -99 Then
                msg = "Failed to load client list!"
            End If
        End If
    Next i
'EH:
'    If msg = "" Then
'        msg = "Failed to delete client: " & ClientName & vbLf & _
'              "from list."
'    End If
'    MsgBox msg, vbOKOnly + vbCritical, "Delete Client"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.DeleteClient(ClientName)", ClientName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.DeleteClient(ClientName)", ClientName
End Sub
Public Property Get DBName() As String
    On Error GoTo errHandler
    DBName = "DATABASE NAME"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.DBName"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.DBName"
End Property

Public Sub FetchClientSalesandUpdateServer()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sFile As String
Dim xFiles() As String
Dim fs As New FileSystemObject
Dim i As Integer, k As Integer

    ReDim xFiles(0)
    sFile = Dir(sInBox & "\*.*") 'Look in inbox for files from client - Only sales data is expected
    Do While sFile <> ""
        If xFiles(UBound(xFiles)) <> "" Then
            ReDim Preserve xFiles(UBound(xFiles) + 1)
        End If
        xFiles(UBound(xFiles)) = sInBox & "\" & sFile    'store filenames in array
        sFile = Dir
    Loop
    For k = 0 To UBound(xFiles)
        If UCase(Right(xFiles(k), 4)) = ".POS" Then
            rs.Open xFiles(k)
            If SaveSaleData(rs) Then                    'Update sales data
                rs.Close
                fs.DeleteFile (xFiles(k)), True
                RaiseEvent FileProcessed(1)
            Else
                fs.MoveFile xFiles(k), xFiles(k) & ".bak"
            End If
        End If
    Next k
        
    Set fs = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.FetchClientSalesandUpdate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.FetchClientSalesandUpdateServer"
End Sub

'~~~~~~~~~ SaveSaleData ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   This reconnects the disconnected Recordset which was send by POS.
'   Then we loop through all the hyrachies and run a batch update on each record.
'
'   Any data checking and manipulation prior to saving could be added to this
'   function, e.g. checking and creating valid product codes, checking of sale
'   prices etc.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function SaveSaleData(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim rsOP As ADODB.Recordset
Dim rsEXCH As ADODB.Recordset
Dim rsCSL As ADODB.Recordset
Dim rsPAY As ADODB.Recordset
Dim msg As String
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim TRID As Long

    oPC.CO.BeginTrans
    'Update or add to tZSession
    'Update or add to tCS
    'Update or add to tOPSession
    'Add to tExchange
    Select Case rs.Fields("typ")
    Case "O"
      '  strSQL = "UPDATE tOPSession SET OPS_ENDTIME = '" & ReverseDateTime(rs!OPS_ENDTIME) & "' WHERE OPS_ID = '" & rs!OPS_ID & "'"
        strSQL = "UPDATE tCS SET CS_ENDSession = '" & ReverseDateTime(rs!OPS_ENDTIME) & "' WHERE CS_GUID = '" & rs!OPS_ID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.CO
        cmd.CommandText = strSQL
        cmd.CommandType = adCmdText
        cmd.Execute
    Case "Z"
        strSQL = "UPDATE tZSession SET Z_ENDDATE = '" & rs!Z_ENDDATE & "' WHERE Z_ID = '" & rs!Z_ID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.CO
        cmd.CommandText = strSQL
        cmd.CommandType = adCmdText
        cmd.Execute
    Case "E"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.CO
        cmd.CommandText = "sp_UPDATEPOS_1"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        Set prm = cmd.CreateParameter("@ZID", adGUID, adParamInput, , rs!Z_ID)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@STARTDATE", adDate, adParamInput, , rs!Z_STARTDATE)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@ENDDATE", adDate, adParamInput, , rs!Z_ENDDATE)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@SUPERVISORID", adInteger, adParamInput, , rs!Z_SUPERVISORID)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@NOMINALDATE", adDate, adParamInput, , rs!Z_NOMINALDATE)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 20, rs!Z_TILLPOINT)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set rsOP = rs.Fields("OPSession").Value
        
        Set prm = cmd.CreateParameter("@OPSID", adGUID, adParamInput, , rsOP!OPS_ID)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@OPSSTARTTIME", adDate, adParamInput, , rsOP!OPS_STARTTIME)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@OPSENDTIME", adDate, adParamInput, , rsOP!OPS_ENDTIME)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@OPERATORID", adInteger, adParamInput, , rsOP!OPS_OPERATORID)
        cmd.Parameters.Append prm
        Set prm = Nothing
               
        Set rsEXCH = rsOP.Fields("POSEXCHANGE").Value
        
        Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_ID))
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXZSESSID", adGUID, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_ZSESSION_ID))
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXOPSESSID", adGUID, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_OPSESSION_ID))
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@EXCHDATE", adDate, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_SALEDATE))
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXCHTPID", adInteger, adParamInput, , oPC.CustomerTPID)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@TRID", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        cmd.Execute
        
        Check cmd.Parameters("@pErrCode") = 0, EXC_GENERAL, "Error in Stored procedure sp_UPDATEPOS_1 Error code is " & cmd.Parameters("@pErrCode") & " in position: " & cmd.Parameters("@ERRPOSITION")
        TRID = IIf(IsNull(cmd.Parameters("@TRID")), 0, cmd.Parameters("@TRID"))
        
    'Add  to tCSL
        Set rsCSL = rsEXCH.Fields("SALESLINES").Value
        Do While Not rsCSL.EOF
            Set cmd = Nothing
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = oPC.CO
            cmd.CommandType = adCmdStoredProc
            cmd.NamedParameters = True
            cmd.CommandText = "sp_CSL"
            Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , rsCSL.Fields("CSL_P_ID"))
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@Qty", adInteger, adParamInput, , rsCSL.Fields("CSL_QTY"))
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@Price", adInteger, adParamInput, , rsCSL.Fields("CSL_PRICE"))
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@DiscountRate", adNumeric, adParamInput, , rsCSL.Fields("CSL_DISCOUNTRate"))
            prm.NumericScale = 4
            prm.Precision = 8
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@Discount", adInteger, adParamInput, , rsCSL.Fields("CSL_Discount"))
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@VRate", adInteger, adParamInput, , rsCSL.Fields("CSL_VATRATE"))
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@STID", adInteger, adParamInput, , oPC.DefaultStoreID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , oPC.CustomerTPID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , rsEXCH!EXCH_ID)
            cmd.Parameters.Append prm
            Set prm = Nothing
'
'            Set prm = cmd.CreateParameter("@ZID", adGUID, adParamInput, , rs!Z_ID)
'            cmd.Parameters.Append prm
'            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@ID", adInteger, adParamOutput) 'dummy in this context
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            cmd.Execute
            Check cmd.Parameters("@pErrCode") = 0, EXC_GENERAL, "Error in Stored procedure sp_CSL Error code is " & cmd.Parameters("@pErrCode") & " in position: " & cmd.Parameters("@ERRPOSITION")
            rsCSL.MoveNext
        Loop
    
    'Add to tPayment
        Set rsPAY = rsEXCH.Fields("PAYMENTS").Value
        Do While Not rsPAY.EOF
            Set cmd = Nothing
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = oPC.CO
            cmd.CommandType = adCmdStoredProc
            cmd.NamedParameters = True
            cmd.CommandText = "sp_UPDATEPOS_2"
            Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , rsEXCH!EXCH_ID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@PAYTYPE", adChar, adParamInput, 1, rsPAY!PAY_PAYMENTTYPE)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@PAYAMT", adInteger, adParamInput, , rsPAY!PAY_AMT)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            
            Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
            cmd.Execute
            Check cmd.Parameters("@pErrCode") = 0, EXC_GENERAL, "Error in Stored procedure sp_UPDATEPOS_3 Error code is " & cmd.Parameters("@pErrCode") & " in position: " & cmd.Parameters("@ERRPOSITION")
            rsPAY.MoveNext
        Loop
    
        Set cmd = Nothing
    End Select
    
    oPC.CO.CommitTrans
    SaveSaleData = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.SaveSaleData(rs)", rs, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.CO
End Function

Public Sub TransmitUpdatesToClients()
    On Error GoTo errHandler
Dim sFile As String
Dim xFile() As String
Dim i As Integer, k As Integer
Dim fs As New FileSystemObject

    ReDim xFile(0)
    For i = 0 To UBound(ClientList)
        With ClientList(i)
            sFile = Dir(sOutBoxRoot & "\" & .OutBox & "\*.U*D")
            Do While sFile <> ""
                If xFile(UBound(xFile)) <> "" Then
                    ReDim Preserve xFile(UBound(xFile) + 1)
                End If
                xFile(UBound(xFile)) = sFile
                sFile = Dir
            Loop
            For k = 0 To UBound(xFile)
                If xFile(k) <> "" Then
                    fs.CopyFile sOutBoxRoot & "\" & .OutBox & "\" & xFile(k), .pathname & "\" & xFile(k)
                    If Not Err Then
                        fs.DeleteFile sOutBoxRoot & "\" & .OutBox & "\" & xFile(k), True
                    Else
                        Err.Clear
                    End If

                End If
            Next k
        End With
    Next i
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.SendUpdates"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.TransmitUpdatesToClients"
End Sub
Public Sub SendNewDRStoAll()
    On Error GoTo errHandler
Dim sFile As String
Dim rs As ADODB.Recordset
Dim sName As String
Dim i As Integer
Dim fs As New FileSystemObject
    
    Set rs = CreateDisconnRecSet()
    sName = "DRS.rs"
    If fs.FileExists(sOutBoxRoot & "\" & sName) Then
        fs.DeleteFile sOutBoxRoot & "\" & sName
    End If
    rs.Save sOutBoxRoot & "\" & sName, adPersistADTG
    
    For i = 0 To UBound(ClientList)
        With ClientList(i)
            fs.CopyFile sOutBoxRoot & "\" & sName, ClientList(i).pathname & "\" & sName
'            sFile = Dir(sOutBoxRoot & "\" & .OutBox & "\*.U*D")
'            Do While sFile <> ""
'                If xFile(UBound(xFile)) <> "" Then
'                    ReDim Preserve xFile(UBound(xFile) + 1)
'                End If
'                xFile(UBound(xFile)) = sFile
'                sFile = Dir
'
'                fs.CopyFile sOutBoxRoot & "\" & .OutBox & "\" & xFile(k), .pathname & "\" & xFile(k)
        End With
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.SendNewDRStoAll"
End Sub

Private Function SendRSFile(TillCode As String, ClientPath As String) As Boolean
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim rs As ADODB.Recordset
Dim sName As String

    If fs.FolderExists(ClientPath) Then
        sName = TillCode & ".rs"
        Set rs = CreateDisconnRecSet(False)
        If Not rs Is Nothing Then
            rs.Save sOutBoxRoot & "\" & sName, adPersistADTG
            fs.CopyFile sOutBoxRoot & "\" & sName, ClientPath & "\" & sName
            If Not Err Then
                fs.DeleteFile sOutBoxRoot & "\" & sName
                SendRSFile = True
            Else
                Err.Clear
            End If
        End If
        
    End If
    
    Set fs = Nothing
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.SendRSFile(TillCode,ClientPath)", Array(TillCode, ClientPath)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.SendRSFile(TillCode,ClientPath)", Array(TillCode, ClientPath)
End Function

Public Function CreateDisconnRecSet() As ADODB.Recordset
    On Error GoTo errHandler
Dim oShapeDB As New PapyConn
Dim rs As New ADODB.Recordset
Dim sSQL As String
Dim msg As String
    If oShapeDB.Connect("MSDataShape") <> 0 Then
        msg = "Failed to create database connection!"
        GoTo errHandler
    End If
    sSQL = "Shape {SELECT * FROM tZSession WHERE Z_ID = NULL} " & _
            "        APPEND ({SELECT * FROM tCS} RELATE Z_ID to CS_ZSessionID) as chpOpSessions, " & _
            "        ((SHAPE  {SELECT * FROM tExchange } " & _
            "                           APPEND ({SELECT * FROM tCSL} as rsCSL " & _
            "                              RELATE EXCH_Exchange_GUID TO CSL_Exchange_GUID) AS chpPayment , " & _
            "                               ({SELECT * FROM tPayment} as rsPay " & _
            "                              RELATE EXCH_Exchange_GUID TO PAY_Exchange_GUID)) AS chpExchange " & _
            "             RELATE Z_ID TO EXCH_ZSESSIONID) "
    rs.ActiveConnection = oShapeDB.CO
    rs.CursorLocation = adUseClient
    rs.Open sSQL, , adOpenStatic, adLockBatchOptimistic
    Set rs.ActiveConnection = Nothing
    Set CreateDisconnRecSet = rs
    
    
MEX:
    oShapeDB.Disconnect
    Set oShapeDB = Nothing
    Set rs = Nothing
'EH:
'    If msg = "" Then msg = Err.Description
'    MsgBox msg
'    Err.Clear
'    GoTo MEX
'Resume
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "clsServerFQ.CreateDisconnRecSet()"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CreateDisconnRecSet"
End Function




