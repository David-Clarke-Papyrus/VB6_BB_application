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

Public Event STATUS(msg As String)
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
Dim strresult As String
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
    cmd.execute
    
    If cmd.Parameters(1).Value = 1 Then
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
    
        cmd.CommandText = "DROP TABLE " & pName & ";"
        cmd.CommandType = adCmdText
        If pMsg > "" Then RaiseEvent STATUS(strMMsg & ": " & pMsg)
        cmd.execute
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
    RaiseEvent STATUS(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            retval = cmd.execute(, pParams)
        Else
            cmd.execute
        End If
    Else
        cmd.execute
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
    
    cmd.execute lngRecordsAffected
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
    RaiseEvent STATUS(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            Set rs = cmd.execute(lngRecordsAffected, pParams)
        Else
            Set rs = cmd.execute(lngRecordsAffected)
        End If
    Else
        Set rs = cmd.execute(lngRecordsAffected)
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

Function GetDynamicRecordset(pName As String, pType As enumCommandType, pMsg As String, rs As ADODB.Recordset)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If rs Is Nothing Then
        Set rs = New ADODB.Recordset
    Else
        If rs.State <> 0 Then rs.Close
    End If
        rs.Open pName, oPC.COShort, adOpenDynamic, adLockOptimistic
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetDynamicRecordset(pName,pType,pParams,pMsg,rs)", Array(pName, pType, _
         pMsg, rs)
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

    
    cmd.execute lngRecordsAffected
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
    cmd.execute
    
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
    oPC.COShort.execute "UPDATE tOPERATION SET OP_RESULT = " & IIf(Result = True, 1, 0) & ",OP_EndedAt = '" & ReverseDateTime(Now()) & "' WHERE OP_ID = " & pOPID
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
    cmd.execute lngRecordsAffected
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

    
    cmd.execute lngRecordsAffected
    
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
    cmd.execute
    
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
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Sub

End Sub
Public Sub ImportCustomerListFromCSVFILE(x As XArrayDB, pFilename As String)
Dim oTF As New z_TextFile
Dim i As Long
Dim s As String

    i = 0
    x.ReDim 0, 1, 0, 1
    oTF.OpenTextFileToRead pFilename
    Do While Not oTF.IsEOF
        x.ReDim 0, 1, 0, 1
        s = oTF.ReadLinefromTextFile
        x(i, 1) = s
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
    cmd.execute
    
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
    cmd.execute
    
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
    cmd.execute
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

Public Function InsertAlert(MsgType As String, MsgText As String, Optional TPACNO As String, Optional ProductID As String)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim bRestart As Boolean

    bRestart = False
restart:
'-------------------------------
    OpenResult = oPC.OpenDBSHort(bRestart)
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertAlert"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@MsgType", adVarChar, adParamInput, 3, MsgType)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@MsgText", adVarChar, adParamInput, 500, MsgText)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TPACNO", adVarChar, adParamInput, 25, TPACNO)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
ErrHandler:
    ErrPreserve
    If Error = "Connection Failure" Then
        bRestart = True
        Resume restart
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.InsertAlert(MsgType,MsgText,DestTPID,ProductID)", Array(MsgType, MsgText, TPACNO, _
         ProductID)
End Function

Public Function SendInvocation(MsgType As String, BranchCode As String, Selection As String, StartDate As String, EndDate As String)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim bRestart As Boolean
Dim Strguid As String
Dim oSM As z_StockManager
    bRestart = False
restart:
'-------------------------------
    OpenResult = oPC.OpenDBSHort(bRestart)
'-------------------------------

    If MsgType = "CustomerStats" Then
        Set cmd = New ADODB.Command
        cmd.CommandText = "_SendInvocation_CustomerStats"
        cmd.CommandType = adCmdStoredProc
        cmd.CommandTimeout = 0
        Set par = cmd.CreateParameter("@BranchSelection", adVarChar, adParamInput, 3000, Selection)
        cmd.Parameters.Append par
        
        cmd.ActiveConnection = oPC.COShort
        cmd.execute
        Set cmd = Nothing
    Else
        If MsgType = "SalesSet" Then
            Set oSM = New z_StockManager
                Set cmd = New ADODB.Command
                cmd.CommandText = "_SendInvocation_SalesSet"
                cmd.CommandType = adCmdStoredProc
                
                Set par = cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 20, BranchCode)
                cmd.Parameters.Append par
                
                Set par = cmd.CreateParameter("@From", adVarChar, adParamInput, 50, StartDate)
                cmd.Parameters.Append par
                Set par = cmd.CreateParameter("@To", adVarChar, adParamInput, 50, EndDate)
                cmd.Parameters.Append par
                cmd.ActiveConnection = oPC.COShort
                cmd.execute
                Set cmd = Nothing
        Else
            If MsgType = "CustomerSet" Then
                Set cmd = New ADODB.Command
                cmd.CommandText = "_SendInvocation_CustomerSet"
                cmd.CommandType = adCmdStoredProc
                
                Set par = cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 20, BranchCode)
                cmd.Parameters.Append par
                Set par = cmd.CreateParameter("@Selection", adVarChar, adParamInput, 6000, Selection)
                cmd.Parameters.Append par
                
                cmd.ActiveConnection = oPC.COShort
                cmd.execute
                Set cmd = Nothing
            Else
                If MsgType = "SOHBulkUpdate" Then
                    Set cmd = New ADODB.Command
                    cmd.CommandText = "_SendInvocation_SOHALL"
                    cmd.CommandType = adCmdStoredProc
                    
                    Set par = cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 20, BranchCode)
                    cmd.Parameters.Append par
                    Set par = cmd.CreateParameter("@Selection", adVarChar, adParamInput, 6000, Selection)
                    cmd.Parameters.Append par
                    
                    cmd.ActiveConnection = oPC.COShort
                    cmd.execute
                    Set cmd = Nothing
                Else
                    If MsgType = "CASHUP_RESEND" Then
                        Set cmd = New ADODB.Command
                        cmd.CommandText = "_SendInvocation_CASHUPS"
                        cmd.CommandType = adCmdStoredProc
                        
                        Set par = cmd.CreateParameter("@BranchCode", adVarChar, adParamInput, 20, BranchCode)
                        cmd.Parameters.Append par
                        Set par = cmd.CreateParameter("@Selection", adVarChar, adParamInput, 6000, Selection)
                        cmd.Parameters.Append par
                        
                        cmd.ActiveConnection = oPC.COShort
                        cmd.execute
                        Set cmd = Nothing
                    Else
                        If MsgType = "COLS_SEND" Then
                            Set cmd = New ADODB.Command
                            cmd.CommandText = "_SendInvocation_COLS"
                            cmd.CommandType = adCmdStoredProc
                            
                            Set par = cmd.CreateParameter("@Selection", adVarChar, adParamInput, 6000, Selection)
                            cmd.Parameters.Append par
                            
                            cmd.ActiveConnection = oPC.COShort
                            cmd.execute
                            Set cmd = Nothing
                        End If
                    End If
                End If
            End If
        End If
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

'ErrHandler:
'    ErrorIn "z_SQL.SendInvocation(MsgType,BranchSelection)", Array(MsgType, BranchSelection)
    Exit Function
ErrHandler:
    ErrorIn "z_SQL.SendInvocation(MsgType,BranchCode,Selection)", Array(MsgType, BranchCode, _
         Selection)
End Function

Public Function MatchLoyaltyCustomers_MissingAtBranch(BRCode As String) As ADODB.Recordset
Dim strRemoteServerName As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

    strRemoteServerName = oPC.Configuration.Stores.FindStoreByCode(BRCode).VPNAddress & "\PBKSINSTANCE2"
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort(False)
'-------------------------------
    Set rs = oPC.COShort.execute("TRUNCATE TABLE tMissingLoyaltyCustomers")
   ' oPC.COShort.CommandTimeout
    oPC.COShort.execute ("INSERT INTO PBKSC.dbo.tMissingLoyaltyCustomers (MLC_Acno,MLC_Name,MLC_DLM,MLC_Store)" _
                    & "Select loc.TP_ACNO acno,dbo.FullNameReversedF(loc.TP_INITIALS ,loc.TP_TITLE ,loc.TP_NAME ) name,loc.TP_DateLastModified dlm, s.STORE_Code Store From vLoyaltyCustomers loc " _
                    & " INNER JOIN tStore s ON loc.TP_LoyaltyHomeStoreID = s.STORE_ID " _
                    & " LEFT JOIN OPENROWSET('SQLOLEDB','Driver={SQL SERVER};SERVER=" & strRemoteServerName & ";UID=sa;PWD=car', " _
                    & " 'SELECT TP_ACNO " _
                    & " FROM PBKS.dbo.vLoyaltyCustomers') as a ON loc.TP_ACNO = a.TP_ACNO WHERE a.TP_ACNO IS NULL ORDER BY store,name")
    Set rs = oPC.COShort.execute("SELECT * FROM tMissingLoyaltyCustomers")
    Set MatchLoyaltyCustomers_MissingAtBranch = rs
       
End Function
Public Function MatchLoyaltyCustomers_MissingAtCentral(BRCode As String) As ADODB.Recordset
Dim strRemoteServerName As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

    strRemoteServerName = oPC.Configuration.Stores.FindStoreByCode(BRCode).VPNAddress & "\PBKSINSTANCE2"
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort(False)
'-------------------------------
    Set rs = oPC.COShort.execute("TRUNCATE TABLE tMissingLoyaltyCustomers")

    oPC.COShort.execute ("INSERT INTO PBKSC.dbo.tMissingLoyaltyCustomers (MLC_Acno,MLC_Name,MLC_DLM,MLC_Store)" _
                    & "Select a.TP_ACNO acno,dbo.FullNameReversedF(a.TP_INITIALS ,a.TP_TITLE ,a.TP_NAME ) name,a.TP_DateLastModified dlm, ISNULL(s.STORE_Code,'Unknown') Store From vLoyaltyCustomers loc " _
                    & " LEFT JOIN tStore s ON loc.TP_LoyaltyHomeStoreID = s.STORE_ID " _
                    & " RIGHT JOIN OPENROWSET('SQLOLEDB','Driver={SQL SERVER};SERVER=" & strRemoteServerName & ";UID=sa;PWD=car', " _
                    & " 'SELECT TP_ACNO,TP_INITIALS ,TP_TITLE ,TP_NAME,TP_DateLastModified" _
                    & " FROM PBKS.dbo.vLoyaltyCustomers') as a ON loc.TP_ACNO = a.TP_ACNO WHERE loc.TP_ACNO IS NULL ORDER BY store,name")
    Set rs = oPC.COShort.execute("SELECT * FROM tMissingLoyaltyCustomers")
    Set MatchLoyaltyCustomers_MissingAtCentral = rs
       
End Function

Public Function MatchLoyaltyCustomers_InterestGroupsMismatch(BRCode As String) As ADODB.Recordset
Dim strRemoteServerName As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

    strRemoteServerName = oPC.Configuration.Stores.FindStoreByCode(BRCode).VPNAddress & "\PBKSINSTANCE2"
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort(False)
'-------------------------------
    Set rs = oPC.COShort.execute("TRUNCATE TABLE tMismatchedLoyaltyCustomers")

    oPC.COShort.execute ("INSERT INTO PBKSC.dbo.tMismatchedLoyaltyCustomers (MLC_Acno,MLC_IGSale,MLC_IGLaunch,MLC_IGPromotion,MLC_IGLiteraryLunch,MLC_Store)" _
                    & "Select loc.TP_ACNO,SA,LA,PR,LL, s.STORE_Code Store From vMismatchedLoyaltyCustomers ml " _
                    & " INNER JOIN tStore s ON ml.TP_LoyaltyHomeStoreID = s.STORE_ID " _
                    & " LEFT JOIN OPENROWSET('SQLOLEDB','Driver={SQL SERVER};SERVER=" & strRemoteServerName & ";UID=sa;PWD=car', " _
                    & " 'SELECT TP_ACNO " _
                    & " FROM PBKS.dbo.vMismatchedLoyaltyCustomers') as mF ON ml.TP_ACNO = mF.TP_ACNO WHERE a.TP_ACNO IS NULL ORDER BY store,name")
    Set rs = oPC.COShort.execute("SELECT * FROM tMismatchedLoyaltyCustomers")
    Set MatchLoyaltyCustomers_InterestGroupsMismatch = rs
       
End Function

Public Sub SendLoyaltyMissingSetToBranch(Optional pBranchCode As String = "")
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute ("INSERT INTO _tCQ(CQ_INT,CQ_TYPE,CQ_DESTINATIONBRANCH) SELECT TOP 300 TP_ID,'TP','" & pBranchCode & "' FROM tTP tp JOIN tMissingLoyaltyCustomers mlc ON tp.TP_ACNO = mlc.MLC_ACNO")
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub

Public Sub SendLoyaltySelectedBranch(Optional pBranchCode As String = "")
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute ("INSERT INTO _tCQ(CQ_INT,CQ_TYPE,CQ_DESTINATIONBRANCH) SELECT  TPID,'TP','" & pBranchCode & "' FROM tS")
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub
'

Public Sub CountingLocalLoyaltyCustomerRecords(pTPCnt As Long, pADDCnt As Long)
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
    cmd.CommandText = "CountingLocalLoyaltyCustomerRecords"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPRecCount", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@AddRecCount", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.execute
    
    pTPCnt = FNS(cmd.Parameters(0))
    pADDCnt = FNS(cmd.Parameters(1))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub
Public Function MatchExchangeRecs(BRCode As String) As ADODB.Recordset
Dim strRemoteServerName As String
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim s As String
 
    If BRCode = "DC" Then
        strRemoteServerName = "PBKS-SVR"
    Else
        strRemoteServerName = oPC.Configuration.Stores.FindStoreByCode(BRCode).VPNAddress & "\PBKSINSTANCE2"
    End If
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort(False)
'-------------------------------
    Set rs = oPC.COShort.execute("TRUNCATE TABLE tMissingExchanges")
    oPC.COShort.CommandTimeout = 0
    
    s = "INSERT INTO PBKSC.dbo.tMissingExchanges (ME_EXCHID, ME_EXCH_DATE,ME_EXCH_NUMBER,ME_EXCH_BRANCHCODE,ME_EXCH_TYPE) " _
                    & " SELECT EXCH_ID,EXCH_SALEDATE,EXCH_NUMBER,' " & BRCode & "',EXCH_TYPE FROM OPENROWSET('SQLOLEDB','Driver={SQL SERVER};SERVER=" & strRemoteServerName & ";UID=sa;PWD=car', " _
                    & " ' SELECT EXCH_ID,EXCH_SALEDATE,EXCH_NUMBER,EXCH_TYPE,EXCH_VOIDED FROM PBKS.dbo.vEXCHANGES') as br LEFT JOIN tEXCHANGE loc ON loc.EXCHANGEID = br.EXCH_ID " _
                    & " WHERE loc.EXCHANGEID IS NULL AND CAST(ISNULL(br.EXCH_VOIDED, 0) AS VARCHAR(5)) = '0' ORDER BY br.EXCH_SALEDATE DESC"
                    MsgBox "Ready to execute . . ." & vbCrLf & s
                    
     oPC.COShort.execute (s)
                   
    Set rs = oPC.COShort.execute("SELECT * FROM tMissingExchanges")
    Set MatchExchangeRecs = rs
 
End Function

Public Function ExportCustomerList(pTextFilename As String)
    On Error GoTo ErrHandler
Dim strCommand As String
Dim oFSO As New FileSystemObject
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim lngCount As Long
    
    strSQL = "SELECT * FROM PBKSC.dbo.vResultDetails "
    
    strCommand = "bcp """ & strSQL & """ queryout """ & pTextFilename & """ -eBCPError.txt -t\t -c -q -r\n -Usa -P" & oPC.Password & " -S " & oPC.servername
    ShellandWaitold strCommand, vbHide, False
    
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.ExportCustomerList(pTextFilename)", pTextFilename
End Function

Public Function LoadBlindCashups(rs As ADODB.Recordset, pBranchCode As String, pDiscrepancy As String, pTIllpoint, pStartDate As Date, pEndDate As Date, pStatus As String) As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetCashups"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@BRANCHCODE", adVarChar, adParamInput, 50, pBranchCode)
    cmd.Parameters.Append par
    Set par = Nothing
    
    If pDiscrepancy > "" Then
        Set par = cmd.CreateParameter("@DISCREPANCY", adNumeric, adParamInput, , CDbl(pDiscrepancy))
    Else
        Set par = cmd.CreateParameter("@DISCREPANCY", adNumeric, adParamInput)
    End If
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 50, pTIllpoint)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@SINCE", adDate, adParamInput, , pStartDate)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@UNTIL", adDate, adParamInput, , pEndDate)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@STATUS", adVarChar, adParamInput, 50, pStatus)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    rs.Open cmd
    
    Set cmd = Nothing

End Function

Public Function LoadCOLS(rs As ADODB.Recordset, pBranchCode As String, pStartDate As Date, pEndDate As Date) As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetCOLS"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@BRANCHCODE", adVarChar, adParamInput, 50, pBranchCode)
    cmd.Parameters.Append par
    Set par = Nothing
    
    
    Set par = cmd.CreateParameter("@SINCE", adDate, adParamInput, , pStartDate)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@UNTIL", adDate, adParamInput, , pEndDate)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    rs.Open cmd
    
    Set cmd = Nothing

End Function

Public Sub ImportFromFile(PSuppID As Long, pSupplierAcno As String, pTemplateFile As String, pDatafilePath As String, pMsg As String, pErrorFilePath As String)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "ImportSupplierInvoices"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@SupplierID", adInteger, adParamInput, , PSuppID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@SupplierAcno", adVarChar, adParamInput, 50, pSupplierAcno)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@DatafilePath", adVarChar, adParamInput, 500, pDatafilePath)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@FormatFilePath", adVarChar, adParamInput, 500, pTemplateFile)
    cmd.Parameters.Append par
    pErrorFilePath = oPC.SharedFolderRoot & "\Logs\BulkInsertErrors" & Format(Now(), "ddmmyy_HHNNSS") & ".TXT"
    Set par = cmd.CreateParameter("@ErrorFilePath", adVarChar, adParamInput, 500, pErrorFilePath)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ResultsMessage", adVarChar, adParamOutput, 500)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ServerName", adVarChar, adParamInput, 500, oPC.servername)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    pMsg = FNS(cmd.Parameters("@ResultsMessage"))
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Sub
ErrHandler:
    ErrorIn "z_SQL.ImportFromFile"
End Sub

Public Sub TransferDeliveryDataToStores(pMsg As String)
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "TransferDeliveryDataToStores"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@ResultsMessage", adVarChar, adParamOutput, 500)
    cmd.Parameters.Append par

    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    pMsg = FNS(cmd.Parameters("@ResultsMessage"))
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
'errHandler:
'    ErrorIn "z_SQL.TransferDeliveryDataToStores"
    Exit Sub
ErrHandler:
    ErrorIn "z_SQL.TransferDeliveryDataToStores(pMsg)", pMsg
End Sub

Public Sub SetupBudgets()
    On Error GoTo ErrHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "SetupBudgets"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
ErrHandler:
    ErrorIn "z_SQL.SetupBudgets"
End Sub
