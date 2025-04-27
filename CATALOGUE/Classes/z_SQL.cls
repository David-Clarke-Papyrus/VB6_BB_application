VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SQL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim f
Dim fMode
Dim strMMsg As String

Public Event ProgressB(lngPos As Long, lngMax As Long, msg As String)

Public Event Status(msg As String)
Public Sub SetMainMsg(pMsg As String)
    On Error GoTo ErrHandler
    strMMsg = pMsg
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.SetMainMsg(pMsg)", pMsg
End Sub
Function TableExists(pTableName As String)
    On Error GoTo ErrHandler
Dim oBatch As z_SQL
Dim rs As ADODB.Recordset
    Set oBatch = New z_SQL
    Set rs = New ADODB.Recordset
    oBatch.RunGetRecordset pTableName, enTable, Array(), "", rs
    TableExists = True
EXIT_Handler:
    Set oBatch = Nothing
    Set rs = Nothing
'ErrHandler:
'    Select Case
'        Case 510
'            TableExists = False
'        Case Else
'            MsgBox error
'    End Select
'    GoTo EXIT_Handler
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.TableExists(pTableName)", pTableName
End Function


Function DropTable(pName As String, Optional pMsg As String) As Long
    On Error GoTo ErrHandler
Dim cmd As ADODB.Command
Dim strSQL As String
Dim strResult As String
Dim par As ADODB.Parameter
Dim OpenResult As Integer

    Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "TableExists"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@pName", adVarChar, adParamInput, 50, pName)
    cmd.Parameters.Append cmd.CreateParameter("@Result", adInteger, adParamOutput)
    cmd.Execute
    
    If cmd.Parameters(1).Value = 1 Then
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
    
        cmd.CommandText = "DROP TABLE " & pName & ";"
        cmd.CommandType = adCmdText
        If pMsg > "" Then RaiseEvent Status(strMMsg & ": " & pMsg)
        cmd.Execute
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.DropTable(pName,pMsg)", Array(pName, pMsg)
End Function
Function RunProc(pName As String, pParams As Variant, pMsg As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim retval
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pName
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    RaiseEvent Status(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            retval = cmd.Execute(, pParams)
        Else
            cmd.Execute
        End If
    Else
        cmd.Execute
    End If
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.RunProc(pName,pParams,pMsg,pConn)", Array(pName, pParams, pMsg, pConn)
End Function
Function RunSQL(pSQL As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pSQL
    cmd.CommandType = adCmdText
    
    cmd.Execute lngRecordsAffected
    Set cmd = Nothing
    RunSQL = lngRecordsAffected
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
ErrHandler:
    ErrPreserve
    Select Case Err
        Case -2147217900, -2147227904
        RunSQL = -1
        Clear
        Exit Function
    End Select
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.RunSQL(pSQL,pConn)", Array(pSQL, pConn), , cmd
End Function
Function RunGetRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset) As Long
    On Error GoTo ErrHandler

Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
Dim OpenResult As Integer
''-------------------------------
    OpenResult = oPC.OpenDBSHort
''-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = pName
    cmd.CommandType = pType
    cmd.CommandTimeout = 360
    RaiseEvent Status(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            Set rs = cmd.Execute(lngRecordsAffected, pParams)
        Else
            Set rs = cmd.Execute(lngRecordsAffected)
        End If
    Else
        Set rs = cmd.Execute(lngRecordsAffected)
    End If
    Set cmd = Nothing
    RunGetRecordset = lngRecordsAffected
''---------------------------------------------------
'        If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
EXIT_Handler:
    Exit Function
ErrHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    'If err = -2147467259 Then 'Connection to database lost
    
    ErrorIn "z_SQL.RunGetRecordset(pName,pType,pParams,pMsg,rs)", Array(pName, pType, pParams, pMsg, rs), EA_DFTRBKCLS, oPC.COShort
End Function

Function GetDynamicRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        rs.Open pName, oPC.COShort, adOpenDynamic, adLockOptimistic
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetDynamicRecordset(pName,pType,pParams,pMsg,rs)", Array(pName, pType, _
         pParams, pMsg, rs)
End Function

Function StartOperation(pNominalDate As Date, pOperatorID As Long, pTypeID As OpTypes) As Long
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
    cmd.CommandText = "StartOperation"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@NominalDate", adDate, adParamInput, , pNominalDate)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@OperatorID", adInteger, adParamInput, , pOperatorID)
    cmd.Parameters.Append prm

    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TypeID", adInteger, adParamInput, , pTypeID)
    cmd.Parameters.Append prm

    Set prm = Nothing
    Set prm = cmd.CreateParameter("@OPID", adInteger, adParamOutput)
    cmd.Parameters.Append prm

    
    cmd.Execute lngRecordsAffected
    StartOperation = CLng(cmd.Parameters(3).Value)
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
ErrHandler:
    ErrPreserve
    Select Case Err
        Case -2147217900, -2147227904
        StartOperation = -1
        Clear
        Exit Function
    End Select
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.StartOperation(pNominalDate,pOperatorID,pTypeID)", Array(pNominalDate, pOperatorID, pTypeID)
End Function
Public Function ExportToHO()
    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim lngRes As Long
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "ExportTRsToPastel_Auto"
    cmd.CommandType = adCmdStoredProc
    
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@Res", adInteger, adParamInput, , lngRes)
'    cmd.Parameters.Append prm
    cmd.Execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.ExportToHO"
End Function


Public Function CompleteOperation(pOPID As Long, Result As Boolean)
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.Execute "UPDATE tOPERATION SET OP_RESULT = " & IIf(Result = True, 1, 0) & ",OP_EndedAt = '" & ReverseDateTime(Now()) & "' WHERE OP_ID = " & pOPID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Public Function CreateConsolidatedGRN(lngRecordsAffected As Long)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "ProduceConsolidatedGRN"
    cmd.CommandType = adCmdStoredProc
    cmd.Execute lngRecordsAffected
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Public Function GetTrackingActions(PXMLArgs As String, CXMLArgs As String, lngTopNumber As Long) As ADODB.Recordset   ', Optional pPOLID As Long, Optional pCOLID As Long

    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "TrackingActions"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@NominalDate", adInteger, adParamInput, , lngTopNumber)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@POLS", adVarChar, adParamInput, 10000, PXMLArgs)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@COLS", adVarChar, adParamInput, 10000, CXMLArgs)
    cmd.Parameters.Append prm
    
'    If pPOLID > 0 Then
'        Set prm = Nothing
'        Set prm = cmd.CreateParameter("@POLID", adInteger, adParamInput, , pPOLID)
'        cmd.Parameters.Append prm
'    End If
'    If pCOLID > 0 Then
'        Set prm = Nothing
'        Set prm = cmd.CreateParameter("@POLID", adInteger, adParamInput)
'        cmd.Parameters.Append prm
'        Set prm = Nothing
'        Set prm = cmd.CreateParameter("@COLID", adInteger, adParamInput, , pCOLID)
'        cmd.Parameters.Append prm
'    End If
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    Set GetTrackingActions = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetTrackingActions(lngTopNumber)", lngTopNumber
End Function
Public Function GetTrackingActionsDetails(PAID As Long, pSupplierMessage As String, pCustomerMessage As String, pSignature As String) As ADODB.Recordset
    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "TrackingActionsDetails"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@PAID", adInteger, adParamInput, , PAID)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@SupplierMessage", adVarChar, adParamOutput, 150, pSupplierMessage)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@CustomerMessage", adVarChar, adParamOutput, 300, pCustomerMessage)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Signature", adVarChar, adParamOutput, 100, pSignature)
    cmd.Parameters.Append prm
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    Set rs.ActiveConnection = Nothing
    pSupplierMessage = FNS(cmd.Parameters(1))
    pCustomerMessage = FNS(cmd.Parameters(2))
    pSignature = FNS(cmd.Parameters(3))
    Set cmd = Nothing
    Set GetTrackingActionsDetails = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetTrackingActionsDetails"
End Function

Public Function GetTrackingActionsDetailsbyPOL(POLID As Long) As ADODB.Recordset
    On Error GoTo ErrHandler
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "TrackingActionsDetailsByPOL"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@POLID", adInteger, adParamInput, , POLID)
    cmd.Parameters.Append prm
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    Set GetTrackingActionsDetailsbyPOL = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetTrackingActionsDetailsbyPOL"
End Function

Function SaveMailLabel(pDescription As String, pLeft As Long, pRowHeight As Long, pColumnSpacing As Long, pTopMargin As Long, pPrintWidth As Long) As Long
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
    cmd.CommandText = "SetMailLabel"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Description", adVarChar, adParamInput, 50, pDescription)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Left", adInteger, adParamInput, , pLeft)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@RowHeight", adInteger, adParamInput, , pRowHeight)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@ColumnSpacing", adInteger, adParamInput, , pColumnSpacing)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@TopMargin", adInteger, adParamInput, , pTopMargin)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@PrintWidth", adInteger, adParamInput, , pPrintWidth)
    cmd.Parameters.Append prm

    
    cmd.Execute lngRecordsAffected
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.SaveMailLabel(pDescription,pLeft,pRowHeight,pColumnSpacing)", Array(pDescription, _
         pLeft, pRowHeight, pColumnSpacing)
End Function

Public Sub MarkForWebExport(PID As String)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "MarkForWebExport"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, PID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub

End Sub

Public Sub ReSendSalesToCentral(pDate As Date)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "ReSendSalesToCentral"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Dte", adVarChar, adParamInput, 50, Format(pDate, "YYYY-MM-DD"))
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub

End Sub
Public Sub ImportCustomerListFromCSVFILE(X As XArrayDB, pFilename As String)
Dim oTF As New z_TextFile
Dim i As Long
Dim s As String

    i = 0
    X.ReDim 0, 1, 0, 1
    oTF.OpenTextFileToRead pFilename
    Do While Not oTF.IsEOF
        X.ReDim 0, 1, 0, 1
        s = oTF.ReadLinefromTextFile
        X(i, 1) = s
        i = i + 1
    Loop
    


End Sub
Public Sub PrepareSalesSpreadsheetData(DateFrom As Date, DateTo As Date)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateSalesReportLines"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DteFrom", adVarChar, adParamInput, 50, Format(DateFrom, "YYYY-MM-DD"))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@DteTo", adVarChar, adParamInput, 50, Format(DateTo, "YYYY-MM-DD"))
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub

End Sub
Public Sub RemoveSBConversation(Handl As String)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "_ClearQueue_ByHandle"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@H", adVarChar, adParamInput, 50, Handl)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.RemoveSBConversation(Handl)", Handl
End Sub

Public Function QtyDocumentsUsingCurrency(pCurrID As Long) As Long
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "QtyDocumentsUsingCurrency"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@CURRID", adInteger, adParamInput, , pCurrID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@QTYD_DOCUMENTS", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    QtyDocumentsUsingCurrency = cmd.Parameters("@QTYD_DOCUMENTS")
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.IsCurrencyInUse(pCurrID)", pCurrID
End Function

'Public Sub PrepareStatements(pOnlyActiveAccounts As Boolean)
'    On Error GoTo ErrHandler
'Dim rs As ADODB.Recordset
'Dim strSQL As String
'Dim oC As a_Customer
'Dim oFSO As New FileSystemObject
'Dim fol
'Dim fil
'Dim f
'Dim iProgress As Long
'Dim iMax As Long
'Dim OpenResult As Integer
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    Me.RunProc "AgeInvoices_All", "", "Prepare statements"
'    ClearStatementsFolder
'    If pOnlyActiveAccounts Then
'        strSQL = "SELECT TP_ID FROM tTP WHERE TP_ROLE = 3 AND (TP_BALANCE <> 0 or TP_BALANCE_120PLUS <> 0 or TP_BALANCE_30 <> 0 or TP_BALANCE_60 <> 0 or TP_BALANCE_90 <> 0) ORDER BY TP_Name,TP_ACNO "
'    Else
'        strSQL = "SELECT TP_ID FROM tTP WHERE TP_ROLE = 3  ORDER BY TP_Name,TP_ACNO "
'    End If
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open strSQL, oPC.COShort, adOpenForwardOnly, adLockReadOnly
'    iMax = rs.RecordCount
'    iProgress = 1
'    Do While Not rs.EOF
'        Set oC = Nothing
'        Set oC = New a_Customer
'        oC.Load FNN(rs.Fields(0))
'        oC.ExportStatement
'        Set oC = Nothing
'        rs.MoveNext
'        iProgress = iProgress + 1
'        RaiseEvent ProgressB(iProgress, iMax, "Preparing statement files . . .")
'        DoEvents
'
'    Loop
'    rs.Close
'    Set rs = Nothing
'    RaiseEvent Status("")
'    DoEvents
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.PrepareStatements"
'End Sub
'
'Public Sub ClearStatementsFolder()
'Dim oFSO As New FileSystemObject
'Dim fil
'Dim f
'    RaiseEvent Status("Clearing folder of old files . . .")
'    DoEvents
'    If Not oFSO.FolderExists(oPC.SharedFolderRoot & "\Statements") Then
'        oFSO.CreateFolder (oPC.SharedFolderRoot & "\Statements")
'    End If
'    Set fil = oFSO.GetFolder(oPC.SharedFolderRoot & "\Statements").Files
'    For Each f In fil
'        f.Delete True
'    Next
'    RaiseEvent Status("")
'
'End Sub
'
'Public Function GetAccountingPeriods() As String
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim rs As ADODB.Recordset
'Dim str As String
'
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "GetAccountingPeriods"
'    cmd.CommandType = adCmdStoredProc
'
''    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput)
''    cmd.Parameters.Append par
''    par.Value = lngTPID
'
'    cmd.ActiveConnection = oPC.CO
'    Set rs = cmd.Execute
'
'    Set cmd = Nothing
'    str = ""
'    Do While Not rs.EOF
'        str = str & FNS(rs.Fields(0)) & vbTab & FNS(rs.Fields(1)) & vbCrLf
'        rs.MoveNext
'    Loop
'    GetAccountingPeriods = str
'    Exit Function
'
'End Function
'
'Public Function Request_SOH_ALLBRANCHES(EAN As String, REQID As Long) As ADODB.Recordset
'    On Error GoTo ErrHandler
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'Dim rs As ADODB.Recordset
'Dim prm As ADODB.Parameter
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandTimeout = 0
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "_Request_SOH_ALLBRANCHES"
'    cmd.CommandType = adCmdStoredProc
'
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 20, EAN)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@REQID", adInteger, adParamOutput, , REQID)
'    cmd.Parameters.Append prm
'
'
'  '  Set rs = New ADODB.Recordset
'  '  rs.CursorLocation = adUseClient
'  '  rs.Open cmd, , adOpenStatic
'  '  Set rs.ActiveConnection = Nothing
'    cmd.Execute
'    REQID = FNN(cmd.Parameters(1))
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
''errHandler:
''    If ErrMustStop Then Debug.Assert False: Resume
''    ErrorIn "z_SQL.GetSOHAllStores"
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SQL.Request_SOH_ALLBRANCHES(EAN,REQID)", Array(EAN, REQID)
'End Function
'
'Public Function GetSOHAllStores(REQID As Long) As ADODB.Recordset
'    On Error GoTo ErrHandler
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'Dim rs As ADODB.Recordset
'Dim prm As ADODB.Parameter
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandTimeout = 0
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "Fetch_SOH_Response"
'    cmd.CommandType = adCmdStoredProc
'
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@REQID", adInteger, adParamInput, , REQID)
'    cmd.Parameters.Append prm
'
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open cmd, , adOpenStatic
'    Set rs.ActiveConnection = Nothing
'    Set cmd = Nothing
'    Set GetSOHAllStores = rs
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
''errHandler:
''    ErrorIn "z_SQL.GetSOHAllStores(EAN,REQID)", Array(EAN, REQID)
'    Exit Function
'ErrHandler:
'    ErrorIn "z_SQL.GetSOHAllStores(REQID)", REQID
'End Function
'
'Public Function LoadAlert(MsgType As String, MsgText As String, Optional AcNo As String, Optional ProductID As String)
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim bRestart As Boolean
'
'    bRestart = False
'restart:
''-------------------------------
'    OpenResult = oPC.OpenDBSHort(bRestart)
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "AlertLoad"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@MsgText", adVarChar, adParamInput, 500, MsgText)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@MsgType", adVarChar, adParamInput, 3, MsgType)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@Acno", adVarChar, adParamInput, 15, AcNo)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.Execute
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
''errHandler:
''    ErrorIn "z_SQL.SendAlert(MsgType,MsgText,DestTPID,ProductID)", Array(MsgType, MsgText, DestTPID, _
''         ProductID)
'    Exit Function
'ErrHandler:
'    ErrorIn "z_SQL.LoadAlert(MsgType,MsgText,Acno,ProductID)", Array(MsgType, MsgText, AcNo, _
'         ProductID)
'End Function
'
'
'
'Public Sub ImportStock(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "ImportBulk_FromExtFile"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@Type", adVarChar, adParamInput, 20, "Stock")
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@DatafilePath", adVarChar, adParamInput, 500, pDatafilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@FormatFilePath", adVarChar, adParamInput, 500, oPC.GetProperty("StockInputFormatFilePath"))
'    cmd.Parameters.Append par
'    pErrorFilePath = oPC.SharedFolderRoot & "\Logs\BulkInsertErrors" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
'    Set par = cmd.CreateParameter("@ErrorFilePath", adVarChar, adParamInput, 500, pErrorFilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@ResultsMessage", adVarChar, adParamOutput, 500)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.Execute
'
'    pMsg = FNS(cmd.Parameters("@ResultsMessage"))
'    Set cmd = Nothing
'
' '---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Sub
'ErrHandler:
'    ErrorIn "z_SQL.ImportStock"
'End Sub
'
'Public Sub ImportSupplier(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "ImportBulk_FromExtFile"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@Type", adVarChar, adParamInput, 20, "Supplier")
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@DatafilePath", adVarChar, adParamInput, 500, pDatafilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@FormatFilePath", adVarChar, adParamInput, 500, oPC.GetProperty("SupplierInputFormatFilePath"))
'    cmd.Parameters.Append par
'    pErrorFilePath = oPC.SharedFolderRoot & "\Logs\BulkInsertErrors" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
'    Set par = cmd.CreateParameter("@ErrorFilePath", adVarChar, adParamInput, 500, pErrorFilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@ResultsMessage", adVarChar, adParamOutput, 500)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.Execute
'
'    pMsg = FNS(cmd.Parameters("@ResultsMessage"))
'    Set cmd = Nothing
'
' '---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Sub
'ErrHandler:
'    ErrorIn "z_SQL.ImportSupplier"
'End Sub
'
'Public Sub ImportCustomer(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "ImportBulk_FromExtFile"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@Type", adVarChar, adParamInput, 20, "Customer")
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@DatafilePath", adVarChar, adParamInput, 500, pDatafilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@FormatFilePath", adVarChar, adParamInput, 500, oPC.GetProperty("CustomerInputFormatFilePath"))
'    cmd.Parameters.Append par
'    pErrorFilePath = oPC.SharedFolderRoot & "\Logs\BulkInsertErrors" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
'    Set par = cmd.CreateParameter("@ErrorFilePath", adVarChar, adParamInput, 500, pErrorFilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@ResultsMessage", adVarChar, adParamOutput, 500)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.Execute
'
'    pMsg = FNS(cmd.Parameters("@ResultsMessage"))
'    Set cmd = Nothing
'
' '---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Sub
'ErrHandler:
'    ErrorIn "z_SQL.ImportCustomer"
'End Sub
'
'Public Sub ImportStockCategory(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo ErrHandler
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "ImportBulk_FromExtFile"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@Type", adVarChar, adParamInput, 20, "StockCategory")
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@DatafilePath", adVarChar, adParamInput, 500, pDatafilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@FormatFilePath", adVarChar, adParamInput, 500, oPC.GetProperty("StockCategoryInputFormatFilePath"))
'    cmd.Parameters.Append par
'    pErrorFilePath = oPC.SharedFolderRoot & "\Logs\BulkInsertErrors" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
'    Set par = cmd.CreateParameter("@ErrorFilePath", adVarChar, adParamInput, 500, pErrorFilePath)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@ResultsMessage", adVarChar, adParamOutput, 500)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.Execute
'
'    pMsg = FNS(cmd.Parameters("@ResultsMessage"))
'    Set cmd = Nothing
'
' '---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Sub
'ErrHandler:
'    ErrorIn "z_SQL.ImportStockCategory"
'End Sub
