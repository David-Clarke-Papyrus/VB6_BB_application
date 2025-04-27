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
Attribute VB_Exposed = False
Option Explicit
Dim f
Dim fMode
Dim strMMsg As String
Public Event Status(msg As String)
Event MaxImportRows(lngCount As Long)
Event LineCountChange(lngCurrentCount As Long)

Dim enImportType As enumImportType



Public Sub SetMainMsg(pMsg As String)
    On Error GoTo errHandler
    strMMsg = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.SetMainMsg(pMsg)", pMsg
End Sub

Public Sub Write_tCommand(ID As String, commandType As String, xml As String)
Dim cmd As ADODB.Command
Dim strSQL As String
Dim strResult As String
Dim OpenResult As Integer

    Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "Insert into tCommand (CMD_ID,CMD_CommandType,CMD_XML) Values ('" & FNS(ID) & "','" & commandType & "','" & xml & "')"
    cmd.commandType = adCmdText
    cmd.execute
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.Write_tCommand(ID,commandType,xml)", Array(ID, commandType, xml)

End Sub

Function TableExists(pTableName As String)
    On Error GoTo errHandler
Dim oBatch As z_SQL
Dim rs As ADODB.Recordset
    Set oBatch = New z_SQL
    Set rs = New ADODB.Recordset
    oBatch.RunGetRecordset pTableName, enTable, Array(), "", rs
    TableExists = True
EXIT_Handler:
    Set oBatch = Nothing
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.TableExists(pTableName)", pTableName
End Function
Function ReportExists(pReportName As String) As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Me.RunGetRecordset "SELECT RPT_NAME FROM tAdHocReports WHERE RPT_NAME = '" & CLARG(pReportName) & "'", enText, Array(), "", rs
    ReportExists = (rs.eof = False)
    rs.Close
    
EXIT_Handler:
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.ReportExists(pReportName)", pReportName
End Function

Function DropTable(pName As String, Optional pMsg As String) As Long
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@pName", adVarChar, adParamInput, 50, pName)
    cmd.Parameters.Append cmd.CreateParameter("@Result", adInteger, adParamOutput)
    cmd.execute
    
    If cmd.Parameters(1).Value = 1 Then
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
    
        cmd.CommandText = "DROP TABLE " & pName & ";"
        cmd.commandType = adCmdText
        If pMsg > "" Then RaiseEvent Status(strMMsg & ": " & pMsg)
        cmd.execute
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.DropTable(pName,pMsg)", Array(pName, pMsg)
End Function
Function RunProc(pName As String, pParams As Variant, pMsg As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    RaiseEvent Status(strMMsg & ": " & pMsg)
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.RunProc(pName,pParams,pMsg,pConn)", Array(pName, pParams, pMsg, pConn)
End Function
Function RunSQL(pSQL As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler
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
    cmd.commandType = adCmdText
    
    cmd.execute lngRecordsAffected
    Set cmd = Nothing
    RunSQL = lngRecordsAffected
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    Select Case Err
        Case -2147217900, -2147227904
        RunSQL = -1
        Err.Clear
        Exit Function
    End Select
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.RunSQL(pSQL,pConn)", Array(pSQL, pConn), , cmd
End Function
Function RunGetRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset) As Long
    On Error GoTo errHandler

Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
Dim OpenResult As Integer
''-------------------------------
    OpenResult = oPC.OpenDBSHort
''-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = pName
    cmd.commandType = pType
    cmd.CommandTimeout = 360
    RaiseEvent Status(strMMsg & ": " & pMsg)
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


   '     rs.Open pName, oPC.COShort, adOpenDynamic, adLockOptimistic



EXIT_Handler:
    RunGetRecordset = 0
    Exit Function
errHandler:
    ErrPreserve
    RunGetRecordset = 1
    If ErrMustStop Then Debug.Assert False: Resume
    'If err = -2147467259 Then 'Connection to database lost
    
    ErrorIn "z_SQL.RunGetRecordset(pName,pType,pParams,pMsg,rs)", Array(pName, pType, pParams, pMsg, rs), EA_DFTRBKCLS, oPC.COShort
End Function
Public Function IsCodeUsedOnDB(Cde As String)
#If H_REPORTS Then
#Else
    Dim oProd As New a_Product
    Dim Res As Integer
    
    Res = oProd.Load("", 0, Cde)
    IsCodeUsedOnDB = (Res = 0)
#End If
End Function
Function GetDynamicRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        rs.Open pName, oPC.COShort, adOpenDynamic, adLockOptimistic
       ' Set rs.ActiveConnection = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetDynamicRecordset(pName,pType,pParams,pMsg,rs)", Array(pName, pType, _
         pParams, pMsg, rs)
End Function
Function GetDynamicRecordset_Improved(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
        cmd.CommandText = pName
        cmd.ActiveConnection = oPC.COShort
        rs.CursorLocation = adUseClient
        rs.Open cmd, , adOpenStatic
        Set rs.ActiveConnection = Nothing
        Set cmd = Nothing
        Set GetDynamicRecordset_Improved = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetDynamicRecordset_Improved(pName,pType,pParams,pMsg,rs)", Array(pName, pType, _
         pParams, pMsg, rs)
End Function

Function StartOperation(pNominalDate As Date, pOperatorID As Long, pTypeID As OpTypes) As Long
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
    ErrPreserve
    Select Case Err
        Case -2147217900, -2147227904
        StartOperation = -1
        Err.Clear
        Exit Function
    End Select
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.StartOperation(pNominalDate,pOperatorID,pTypeID)", Array(pNominalDate, pOperatorID, pTypeID)
End Function
Public Function ExportToHO()
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@Res", adInteger, adParamInput, , lngRes)
'    cmd.Parameters.Append prm
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
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
    cmd.commandType = adCmdStoredProc
    cmd.execute lngRecordsAffected
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Public Function GetTrackingActions(PXMLArgs As String, CXMLArgs As String, lngTopNumber As Long) As ADODB.Recordset   ', Optional pPOLID As Long, Optional pCOLID As Long

    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetTrackingActions(lngTopNumber)", lngTopNumber
End Function
Public Function GetTrackingActionsDetails(PAID As Long, pSupplierMessage As String, pCustomerMessage As String, pSignature As String) As ADODB.Recordset
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetTrackingActionsDetails"
End Function

Public Function GetTrackingActionsDetailsbyPOL(POLID As Long) As ADODB.Recordset
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.GetTrackingActionsDetailsbyPOL"
End Function

Function SaveMailLabel(pDescription As String, pLeft As Long, pRowHeight As Long, pColumnSpacing As Long, pTopMargin As Long, pPrintWidth As Long) As Long
    On Error GoTo errHandler
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
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
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
    cmd.commandType = adCmdStoredProc
    
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
    cmd.commandType = adCmdStoredProc
    
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
'Public Sub ImportIBTFromFile(pFilePath As String, pBadRecords As Long)
'    On Error GoTo errHandler
'Dim lngLineCount As Long
'Dim strCode As String
'Dim strQty As String
'Dim lngQty As Long
'Dim oBatch As z_SQL
'Dim strCodeOut As String
'Dim strEANOut As String
'Dim strLine As String
'Dim fs As New FileSystemObject
'Dim arTF() As String
'Dim arLine() As String
''Dim oProd As a_Product
'Dim res As Long
'Dim txtStream As Scripting.TextStream
'Dim oS As New z_StockManager
'Dim DestCode As String
'Dim lngTFRID As Long
'
'    Set oBatch = New z_SQL
'    Set txtStream = fs.OpenTextFile(pFilePath)
'
'    lngLineCount = 0
'    Do While Not txtStream.AtEndOfStream
'        txtStream.ReadLine
'        lngLineCount = lngLineCount + 1
'    Loop
'    RaiseEvent MaxImportRows(lngLineCount)
'    lngLineCount = 0
'    txtStream.Close
'    Set txtStream = Nothing
'    Set txtStream = fs.OpenTextFile(pFilePath)
'    If Not txtStream.AtEndOfStream Then
'        DestCode = txtStream.ReadLine
'    End If
'    lngLineCount = 0
'    Do While Not txtStream.AtEndOfStream
'        strLine = txtStream.ReadLine
'        If strLine = "" Then GoTo skip
'        If Right(strLine, 1) = "," Then strLine = Left(strLine, Len(strLine) - 1)
'        arLine = Split(strLine, ",")
'        strCode = Trim(arLine(0))
'        If UBound(arLine) = 0 Then
'            lngQty = 1
'        Else
'            lngQty = CLng(Trim(arLine(1)))
'        End If
''        If oPC.BFLoaded Then
''            Set oProd = New a_Product
''            If Not oProd.Exists(strCode) Then
''                Res = oProd.Load("", 0, strCode)
''            End If
''            Set oProd = Nothing
''        End If
'
'
'        ReDim arTF(lngLineCount, 2)
'        arTF(lngLineCount, 0) = arLine(0)
'        If UBound(arLine) > 1 Then
'            arTF(lngLineCount, 2) = arLine(1)
'        End If
'        lngLineCount = lngLineCount + 1
'        RaiseEvent LineCountChange(lngLineCount)
'
'skip:
'    Loop
'
'
'Dim xMLDoc As ujXML
'Dim XMLArgs As String
'    Set xMLDoc = New ujXML
'    With xMLDoc
'        .docProgID = "MSXML2.DOMDocument"
'        .docInit "doc_IBT"
'            .chCreate "MessageType"
'                .elText = "IBT"
'            .elCreateSibling "MessageCreationDate"
'                .elText = Format(Now(), "yyyymmddHHNN")
'            .elCreateSibling "WORKSTATION"
'                .elText = oPC.WorkstationName
'            .elCreateSibling "StoreCode"
'                .elText = DestCode
'            .elCreateSibling "INOUT"
'                .elText = "OUT"
'            .elCreateSibling "DetailLine", True
'            For i = 0 To UBound(arTF, 1)
'                If arTF(i, 0) > "" Then
'                    .chCreate "ITEM"
'                    .chCreate "EAN"
'                        .elText = arTF(i, 0)
'                    .elCreateSibling "Qty", True
'                        .elText = "1"
'                    .navUP
'                    .navUP
'                End If
'            Next i
'         XMLArgs = .docXML
'    End With
'
'
'    lngTFRID = oS.CreateNewTransferFromFile(XMLArgs, "OUT", "Transfer from scanned file")
'    If lngTFRID = 0 Then
'        MsgBox "There is a problem generating a transfer"
'    End If
'
'    Set oBatch = Nothing
'EXIT_Handler:
'    Exit Sub
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_SQL.ImportIBTFromFile(pFilePath,pBadRecords)", Array(pFilePath, pBadRecords)
'End Sub
'Public Sub PrepareSalesSpreadsheetData(DateFrom As Date, DateTo As Date)
'Dim OpenResult As Integer
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'Dim f As String
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "CreateSalesReportLines"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@DteFrom", adVarChar, adParamInput, 50, Format(DateFrom, "YYYY-MM-DD"))
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@DteTo", adVarChar, adParamInput, 50, Format(DateTo, "YYYY-MM-DD"))
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.execute
'
'    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Sub
'
'End Sub
Public Sub RemoveSBConversation(Handl As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "_ClearQueue_ByHandle"
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.RemoveSBConversation(Handl)", Handl
End Sub

Public Function QtyDocumentsUsingCurrency(pCurrID As Long) As Long
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim f As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "QtyDocumentsUsingCurrency"
    cmd.commandType = adCmdStoredProc
    
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.IsCurrencyInUse(pCurrID)", pCurrID
End Function

'Public Sub PrepareStatements(pOnlyActiveAccounts As Boolean)
'    On Error GoTo errHandler
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.PrepareStatements"
'End Sub


Public Function GetAccountingPeriods() As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
Dim str As String

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetAccountingPeriods"
    cmd.commandType = adCmdStoredProc
    
'    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput)
'    cmd.Parameters.Append par
'    par.Value = lngTPID
    
    cmd.ActiveConnection = oPC.CO
    Set rs = cmd.execute
    
    Set cmd = Nothing
    str = ""
    Do While Not rs.eof
        str = str & FNS(rs.Fields(0)) & vbTab & FNS(rs.Fields(1)) & vbCrLf
        rs.MoveNext
    Loop
    GetAccountingPeriods = str
    Exit Function

End Function

Public Function Request_SOH_ALLBRANCHES(EAN As String, REQID As Long) As ADODB.Recordset
    On Error GoTo errHandler
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
    cmd.CommandText = "_Request_SOH_ALLBRANCHES"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 20, EAN)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@REQID", adInteger, adParamOutput, , REQID)
    cmd.Parameters.Append prm
    
    cmd.execute
    REQID = FNN(cmd.Parameters(1))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.Request_SOH_ALLBRANCHES(EAN,REQID)", Array(EAN, REQID)
End Function

Public Function GetSOHAllStores(REQID As Long) As ADODB.Recordset
    On Error GoTo errHandler
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
    cmd.CommandText = "Fetch_SOH_Response"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@REQID", adInteger, adParamInput, , REQID)
    cmd.Parameters.Append prm
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic
    Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    Set GetSOHAllStores = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    ErrorIn "z_SQL.GetSOHAllStores(EAN,REQID)", Array(EAN, REQID)
    Exit Function
errHandler:
    ErrorIn "z_SQL.GetSOHAllStores(REQID)", REQID
End Function

Public Function LoadAlert(MsgType As String, MsgText As String, Optional AcNo As String, Optional ProductID As String)
    On Error GoTo errHandler
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
    cmd.CommandText = "AlertLoad"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@MsgText", adVarChar, adParamInput, 500, MsgText)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@MsgType", adVarChar, adParamInput, 3, MsgType)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Acno", adVarChar, adParamInput, 15, AcNo)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    ErrorIn "z_SQL.SendAlert(MsgType,MsgText,DestTPID,ProductID)", Array(MsgType, MsgText, DestTPID, _
'         ProductID)
    Exit Function
errHandler:
    ErrorIn "z_SQL.LoadAlert(MsgType,MsgText,Acno,ProductID)", Array(MsgType, MsgText, AcNo, _
         ProductID)
End Function

Public Function FindCORefMatch(pRef As String, pDocnoSet As String)
    On Error GoTo errHandler
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
    cmd.CommandText = "FindCORefMatch"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Ref", adVarChar, adParamInput, 500, pRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@DocnoSet", adVarChar, adParamOutput, 1000, pDocnoSet)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    pDocnoSet = FNS(cmd.Parameters("@DocNoSet"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.FindCORefMatch(pRef,pDocnoSet)", Array(pRef, pDocnoSet)
End Function
Public Function FindSuppInvMatch(pRef As String, pDocnoSet As String, pCurrentTRID As Long)
    On Error GoTo errHandler
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
    cmd.CommandText = "FindSuppInvMatch"
    cmd.commandType = adCmdStoredProc
    
'    Set par = cmd.CreateParameter("@Ref", adVarChar, adParamInput, 500, pRef)
'    cmd.Parameters.Append par
'
'    Set par = cmd.CreateParameter("@CurrentTRIDForExclusion", adInteger, adParamInput, , pCurrentTRID)
'    cmd.Parameters.Append par
'
'    Set par = cmd.CreateParameter("@DocnoSet", adVarChar, adParamOutput, 1000, pDocnoSet)
'    cmd.Parameters.Append par
'
'
    Set par = cmd.CreateParameter("@Ref", adVarChar, adParamInput, 500, pRef)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@DocnoSet", adVarChar, adParamOutput, 1000, pDocnoSet)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@CurrentTRIDForExclusion", adInteger, adParamInput, , pCurrentTRID)
    cmd.Parameters.Append par
    
    
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    pDocnoSet = FNS(cmd.Parameters("@DocNoSet"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.FindSuppInvMatch(pRef,pDocnoSet)", Array(pRef, pDocnoSet)
End Function

'Public Sub ImportStock(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo errHandler
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
'    Set par = cmd.CreateParameter("@ServerName", adVarChar, adParamInput, 500, oPC.Servername)
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
'errHandler:
'    ErrorIn "z_SQL.ImportStock"
'End Sub
'Public Sub ValidateFile(pDatafilePath As String, enImportType As enumImportType)
'Dim oTF As New z_TextFile
'Dim strLine As String
'Dim El() As String
'Dim lngColCount As Long
'
'    oTF.OpenTextFileToRead (pDatafilePath)
'    strLine = oTF.ReadLinefromTextFile
'    Do While Not oTF.IsEOF
'        El = Split(strLine, vbTab)
'        lngColCount = UBound(El)
'        strLine = oTF.ReadLinefromTextFile
'    Loop
'    oTF.CloseTextFile
'
'    Select Case enImportType
'    Case enSupplierImport
'
'
'    End Select
'End Sub
'Public Sub ImportSupplier(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo errHandler
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
'    Set par = cmd.CreateParameter("@ServerName", adVarChar, adParamInput, 500, oPC.Servername)
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
'errHandler:
'    ErrorIn "z_SQL.ImportSupplier"
'End Sub
'
'Public Sub ImportCustomer(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo errHandler
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
'    Set par = cmd.CreateParameter("@ServerName", adVarChar, adParamInput, 500, oPC.Servername)
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
'errHandler:
'    ErrorIn "z_SQL.ImportCustomer"
'End Sub
'
'Public Sub ImportStockCategory(pDatafilePath As String, pMsg As String, pErrorFilePath As String)
'    On Error GoTo errHandler
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
'    Set par = cmd.CreateParameter("@ServerName", adVarChar, adParamInput, 500, oPC.Servername)
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
'errHandler:
'    ErrorIn "z_SQL.ImportStockCategory"
'End Sub

'''''
Public Sub ImportFromFile(pType As String, pTemplateFile As String, pDatafilePath As String, pMsg As String, pErrorFilePath As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "ImportBulk_FromExtFile"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set par = cmd.CreateParameter("@Type", adVarChar, adParamInput, 20, pType) ' "StockCategory")
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
errHandler:
    ErrorIn "z_SQL.ImportFromFile"
End Sub
'''''



Public Sub LoadBrowsedProductsToTempTable(x As XArrayDB)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    oPC.COShort.execute ("TRUNCATE TABLE tmpBrowsedProducts")
    For i = 1 To x.UpperBound(1)
        oPC.COShort.execute ("INSERT INTO tmpBrowsedProducts (PID) VALUES ('" & FNS(x(i, 8)) & "')")
    Next
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.LoadBrowsedProductsToTempTable(X)", x
End Sub

'Public Sub FetchPotentialReturns(pTPID As Long, pFrom As Date, pTo As Date, pMsg As String, rs As adodb.Recordset)
Public Sub FetchPotentialReturns(pTPID As Long, pSpecifiedMonth As Date, pMsg As String, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

'    Set cmd = New adodb.Command
'    cmd.CommandText = "GetReturnsData"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 0
'
'    If pTPID > 0 Then
'        Set par = cmd.CreateParameter("@pTPID", adInteger, adParamInput, , pTPID)
'    Else
'        Set par = cmd.CreateParameter("@pTPID", adInteger, adParamInput)
'    End If
'    cmd.Parameters.Append par
'
'    Set par = cmd.CreateParameter("@From", adDate, adParamInput, , pFrom)
'    cmd.Parameters.Append par
'
'    Set par = cmd.CreateParameter("@To", adDate, adParamInput, , pTo)
'    cmd.Parameters.Append par
'
'
'    cmd.ActiveConnection = oPC.COShort
'    rs.Open cmd
'    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.CommandText = "GetReturnsData2"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    If pTPID > 0 Then
        Set par = cmd.CreateParameter("@pTPID", adInteger, adParamInput, , pTPID)
    Else
        Set par = cmd.CreateParameter("@pTPID", adInteger, adParamInput)
    End If
    cmd.Parameters.Append par
    If pSpecifiedMonth > CDate("2005-01-01") Then
        Set par = cmd.CreateParameter("@SpecifiedMonth", adDate, adParamInput, , pSpecifiedMonth)
        cmd.Parameters.Append par
    End If
    
    
    
    cmd.ActiveConnection = oPC.COShort
    rs.Open cmd
    Set cmd = Nothing
    
 '---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Sub
errHandler:
    ErrorIn "z_SQL.ImportFromFile"
End Sub

Public Function RerunTransactions(pTRID As Long, pUntil As Date, ResetToZero As Boolean) As Long
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "ReRunTransactions"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@ResetCostToZero", adBoolean, adParamInput, , ResetToZero)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@SinceStockTakeID", adInteger, adParamInput, , pTRID)
    cmd.Parameters.Append par
    
  '  Set par = cmd.CreateParameter("@UNTIL", adDate, adParamInput, , pUntil)
  '  cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    RerunTransactions = cmd.Parameters(2)
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Function
errHandler:
    ErrorIn "z_SQL.RerunTransactions"
End Function
Public Function RerunTransactionsPerPID(pCode As String, pRes As Long) As Long
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "ReRunTransactionsperPID"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 20, pCode)
    cmd.Parameters.Append par
    
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    RerunTransactionsPerPID = cmd.Parameters(1)
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Function
errHandler:
    ErrorIn "z_SQL.RerunTransactionsPerPID"
End Function


Public Sub DeletePrinter(PrinterID As Long)
 Dim OpenResult As Integer
Dim cmd As ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "DELETE FROM tPRINTERS WHERE PRINT_ID = " & CStr(PrinterID)
    cmd.ActiveConnection = oPC.COShort
    cmd.commandType = adCmdText
    cmd.CommandTimeout = 0
    cmd.execute
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
End Sub

Public Function GetOpenDaySessions() As String
Dim s As String
Dim rs As ADODB.Recordset
    
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.ActiveConnection = oPC.COShort
    
    rs.Open "SELECT * FROM tZSession WHERE ISNULL(Z_ENDDATE,'1950-01-01') < Z_STARTDATE"
    s = ""
    Do While Not rs.eof
        s = IIf(s > "", s & vbCrLf, "") & FNS(rs.Fields("Z_TILLPOINT")) & ":" & FND(rs.Fields("Z_STARTDATE"))
        rs.MoveNext
    Loop
    GetOpenDaySessions = s
    
End Function

Public Function SaveReport(ReportName As String, ReportMetaData As String)

    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim s As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "SaveReport"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@ReportName", adVarChar, adParamInput, 200, ReportName)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@Metadata", adVarChar, adParamInput, 100000, ReportMetaData)
    cmd.Parameters.Append par
    Set par = Nothing
   
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Function
errHandler:
    ErrorIn "z_SQL.SaveReport"
End Function
Public Function DeleteReport(ReportName As String)
    On Error GoTo errHandler

Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim s As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "DeleteReport"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@ReportName", adVarChar, adParamInput, 200, ReportName)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
'errHandler:
'    ErrorIn "z_SQL.DeleteReport"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.DeleteReport(ReportName)", ReportName
End Function
Public Function LoadReport(ReportName As String, ReportMetaData As String)  ', ReportLayout() As Byte)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------


Dim rs As ADODB.Recordset
Dim mstream As ADODB.Stream
Set rs = New ADODB.Recordset
    rs.Open "Select * from tAdhocReports WHERE RPT_NAME = '" & ReportName & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
    
    ReportMetaData = CStr(rs.Fields("RPT_METADATA").Value)
    
    rs.Close

   
    Exit Function
errHandler:
    ErrorIn "z_SQL.LoadReport"
End Function

Public Function TestSQL(SQL As String, SQLsmg As String) As Boolean
Dim OpenResult As Integer
Dim strCommandFilePath As String
Dim oTF As New z_TextFileSimple
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
Dim strPath As String
Dim fs As New FileSystemObject

    TestSQL = False
    strCommandFilePath = "c:\PBKS\TESTSQL.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "SET NOEXEC ON;"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile SQL
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "SET NOEXEC OFF;"
        oTF.WriteToTextFile "GO"
        
        oTF.CloseTextFile
        Set oTF = Nothing
        TestSQL = ExecuteScript(strCommandFilePath, SQLsmg)
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If

End Function
Private Function ExecuteScript(strCommandFilePath, SQLsmg As String) As Boolean
Dim strCommand As String
Dim Res As Boolean
Dim fs As New FileSystemObject
Dim f As File
Dim oF As z_TextFileSimple

    ExecuteScript = False
    strCommand = "SQLCMD -Usa -P" & oPC.Password & " -S" & oPC.servername & " -d" & oPC.DatabaseName & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand, , 10000)
    End If
    If fs.FileExists(Replace(strCommandFilePath, ".SQL", ".ERR")) Then
        Set f = fs.GetFile(Replace(strCommandFilePath, ".SQL", ".ERR"))
        If f.Size = 0 Then
            ExecuteScript = True
        Else
            Set oF = New z_TextFileSimple
            oF.OpenTextFileToRead f.Path
            SQLsmg = oF.ReadWholeFile
            oF.CloseTextFile
            ExecuteScript = False
        End If
    End If
    
    
End Function


Private Sub PrepareScript()
'txtPublicKeyFile
End Sub

Public Function GetPerformanceStats() As ADODB.Recordset

Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetPerformanceStats"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
'    Set par = cmd.CreateParameter("@ReportName", adVarChar, adParamInput, 200, ReportName)
'    cmd.Parameters.Append par
'    Set par = Nothing
    Set rs = New ADODB.Recordset
    cmd.ActiveConnection = oPC.COShort
    rs.CursorLocation = adUseClient
    rs.Open cmd
    
    Set cmd = Nothing
    Set GetPerformanceStats = rs
End Function

Public Sub CategoryCheck(rs As ADODB.Recordset, lngCatChkID As Long)

Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetCategoryCheck"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@CATCHKID", adInteger, adParamInput, , lngCatChkID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    rs.Open cmd
    
    Set cmd = Nothing
End Sub

Public Function CalculateJournalSummary(Optional pFrom As Date, Optional pTo As Date, Optional ZID As String) As Long
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CalculateJournalSummary"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@From", adDate, adParamInput, , pFrom)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@From", adDate, adParamInput, , pTo)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@ZID", adVarChar, adParamInput, 100, ZID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Function
errHandler:
    ErrorIn "z_SQL.RerunTransactions"
End Function

Public Sub ReturnRecordset(rs As ADODB.Recordset, lngTRID As Long, pSpecifiedMonth As Date)
    On Error GoTo errHandler

Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetReturnLines"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , lngTRID)
    cmd.Parameters.Append par
    Set par = Nothing
    If pSpecifiedMonth > CDate("2005-01-01") Then
        Set par = cmd.CreateParameter("@SpecifiedMonth", adDate, adParamInput, , pSpecifiedMonth)
    Else
        Set par = cmd.CreateParameter("@SpecifiedMonth", adDate, adParamInput)
    End If
        cmd.Parameters.Append par
        Set par = Nothing
    
    
    cmd.ActiveConnection = oPC.COShort
    rs.Open cmd
    
    Set cmd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.ReturnRecordset(rs,lngTRID,pSpecifiedMonth)", Array(rs, lngTRID, pSpecifiedMonth)
End Sub

Public Function SaveCashup( _
StageNumber As Integer, _
CU_ID As String, _
SavingOperator As Long, _
IssuingOperator As Long, _
ExplainingOperator As Long, _
FloatVal As Double, _
ChequesVal As Double, _
CCardsVal As Double, _
DCardsVal As Double, _
VouchersVal As Double, _
DepositsVal As Double, _
CashR200 As Long, _
CashR100 As Long, _
CashR50 As Long, _
CashR20 As Long, _
CashR10 As Long, _
CashC500 As Long, _
CashC200 As Long, _
CashC100 As Long, _
CashC50 As Long, _
CashC20 As Long, _
CashC10 As Long, _
CashC5 As Long, FloatBreakdown As String, sExplanation As String, DiscCash As Double, DiscCheques As Double, DiscCards As Double, DiscDeposits As Double, DiscVouchers As Double, DiscFloat As Double, Banked As Double, _
PettyCashNett As Double, Retained As Double, Returned As Double, pTotalSales As Double, pCOGS As Double, pGiftVouchers As Double, pDepositsReceived, pDepositsRedeemed, pDepositsRefunder, pOtherVouchers As Double, pWages As Double, pSickLeave As Double, pLeavePay As Double) As Long
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "SaveCashup"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    
    Set par = Nothing
    Set par = cmd.CreateParameter("@CU_XID", adGUID, adParamInput, 40, CU_ID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_WordstationID", adInteger, adParamInput, , oPC.WorkstationID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CapturedByStaffID", adInteger, adParamInput, , SavingOperator)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_IssuedByStaffID", adInteger, adParamInput, , IssuingOperator)
    cmd.Parameters.Append par
    Set par = Nothing
   
    Set par = cmd.CreateParameter("@CU_ExplainedByStaffID", adInteger, adParamInput, , ExplainingOperator)
    cmd.Parameters.Append par
    Set par = Nothing
   
    Set par = cmd.CreateParameter("@CU_Float", adNumeric, adParamInput, , FloatVal)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Cheques", adNumeric, adParamInput, , ChequesVal)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CreditCards", adNumeric, adParamInput, , CCardsVal)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_DebitCards", adNumeric, adParamInput, , DCardsVal)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_DirectDeposits", adNumeric, adParamInput, , DepositsVal)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Vouchers", adNumeric, adParamInput, , VouchersVal)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashN200", adInteger, adParamInput, , CashR200)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashN100", adInteger, adParamInput, , CashR100)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashN50", adInteger, adParamInput, , CashR50)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashN20", adInteger, adParamInput, , CashR20)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashN10", adInteger, adParamInput, , CashR10)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC500", adInteger, adParamInput, , CashC500)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC200", adInteger, adParamInput, , CashC200)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC100", adInteger, adParamInput, , CashC100)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC50", adInteger, adParamInput, , CashC50)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC20", adInteger, adParamInput, , CashC20)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC10", adInteger, adParamInput, , CashC10)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_CashC5", adInteger, adParamInput, , CashC5)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_FloatBreakdownAtEnd", adVarChar, adParamInput, 200, FloatBreakdown)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Explanation", adVarChar, adParamInput, 1000, sExplanation)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_DiscrepancyCash", adNumeric, adParamInput, , DiscCash)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_DiscrepancyCheques", adNumeric, adParamInput, , DiscCheques)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_DiscrepancyCards", adNumeric, adParamInput, , DiscCards)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
     
    Set par = cmd.CreateParameter("@CU_DiscrepancyDeposits", adNumeric, adParamInput, , DiscDeposits)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_DiscrepancyVouchersRedeemed", adNumeric, adParamInput, , DiscVouchers)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
 
   
    Set par = cmd.CreateParameter("@CU_DiscrepancyFloat", adNumeric, adParamInput, , DiscFloat)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Banked", adNumeric, adParamInput, , Banked)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_PettyCashNett", adNumeric, adParamInput, , PettyCashNett)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Retained", adNumeric, adParamInput, , Retained)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Returned", adNumeric, adParamInput, , Returned)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_TotalSales", adNumeric, adParamInput, , pTotalSales)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_COGS", adNumeric, adParamInput, , pCOGS)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_GiftVouchers", adNumeric, adParamInput, , pGiftVouchers)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
'
'1600      Set par = cmd.CreateParameter("@CU_DepositsReceived", adNumeric, adParamInput, , pGiftVouchers)
'1610      par.NumericScale = 2
'1620      par.Precision = 12
'1630      cmd.Parameters.Append par
'1640      Set par = Nothing
'
'1650      Set par = cmd.CreateParameter("@CU_DepositsRedeemed", adNumeric, adParamInput, , pGiftVouchers)
'1660      par.NumericScale = 2
'1670      par.Precision = 12
'1680      cmd.Parameters.Append par
'1690      Set par = Nothing
'
'1700      Set par = cmd.CreateParameter("@CU_DepositsRefunded", adNumeric, adParamInput, , pGiftVouchers)
'1710      par.NumericScale = 2
'1720      par.Precision = 12
'1730      cmd.Parameters.Append par
'1740      Set par = Nothing
'
    Set par = cmd.CreateParameter("@CU_OtherVouchers", adNumeric, adParamInput, , pOtherVouchers)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_Wages", adNumeric, adParamInput, , pWages)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_SickLeave", adNumeric, adParamInput, , pSickLeave)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@CU_LeavePay", adNumeric, adParamInput, , pLeavePay)
    par.NumericScale = 2
    par.Precision = 12
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.SaveCashup(CU_ID,SavingOperator,IssuingOperator,FloatVal,ChequesVal,CCardsVal," & _
        "DCardsVal,VouchersVal,DepositsVal,CashR200,CashR100,CashR50,CashR20,CashR10,CashC500,CashC200," & _
        "CashC100,CashC50,CashC20,CashC10,CashC5)", Array(CU_ID, SavingOperator, IssuingOperator, FloatVal, _
         ChequesVal, CCardsVal, DCardsVal, VouchersVal, DepositsVal, CashR200, CashR100, CashR50, CashR20, _
         CashR10, CashC500, CashC200, CashC100, CashC50, CashC20, CashC10, CashC5, Banked, PettyCashNett), , , "line number", Array(Erl())
End Function

Public Function LoadBlindCashup(rs As ADODB.Recordset, ZID As String) As ADODB.Recordset
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetBlindCashup"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@XID", adVarChar, adParamInput, 40, Trim(ZID))
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    rs.Open cmd
    
    Set cmd = Nothing

End Function

Public Function GetDateOfMostRecentFullySignedOffDailySession(pSince As Date) As Date
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetDateOfMostRecentFullySignedOffDailySession"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@Since", adDate, adParamInput, , pSince)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Dte", adDate, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    GetDateOfMostRecentFullySignedOffDailySession = FND(cmd.Parameters(1))
    
    Set cmd = Nothing

End Function
Public Function GetDateOfEarliestUnSignedSession() As Date
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetDateOfEarliestUnSignedSession"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@Dte", adDate, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    GetDateOfEarliestUnSignedSession = FND(cmd.Parameters(0))
    
    Set cmd = Nothing

End Function
Public Function GetDateOfEarliestUnSignedSessionSince(pSince As Date) As Date
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetDateOfEarliestUnSignedSessionSince"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@Since", adDate, adParamInput, , ReverseDate(pSince))
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Dte", adDate, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    GetDateOfEarliestUnSignedSessionSince = FND(cmd.Parameters(1))
    
    Set cmd = Nothing


End Function
Public Sub ExportDataToFile( _
    pView As String, _
    pFilePath As String, _
    pSeparator As String, _
    pServer As String, _
    pErr As Long, _
    pErrMsg As String _
)
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "ExportStockToTabDelimited"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@SERVERname", adVarChar, adParamInput, 500, pServer)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@TableName", adVarChar, adParamInput, 500, pView)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@FilePath", adVarChar, adParamInput, 500, pFilePath)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@Separator", adVarChar, adParamInput, 1, pSeparator)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@Result", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@ErrDesc", adVarChar, adParamOutput, 500, "")
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    pErr = FNN(cmd.Parameters(4))
    pErrMsg = FNS(cmd.Parameters(5))
    
    Set cmd = Nothing

End Sub
Public Function CreatePayment(TPID As Long, Amt As Double, lngStatementLineID As Long, Ref As String, dte As Date)
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
    cmd.CommandText = "CreatePayment"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , TPID)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Amt", adDouble, adParamInput, , Amt)
    cmd.Parameters.Append prm

    Set prm = Nothing
    Set prm = cmd.CreateParameter("@lngStatementLineID", adInteger, adParamInput, , lngStatementLineID)
    cmd.Parameters.Append prm

    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Ref", adVarChar, adParamInput, 100, Ref)
    cmd.Parameters.Append prm

    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Dte", adDate, adParamInput, , dte)
    cmd.Parameters.Append prm

    
    cmd.execute lngRecordsAffected
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
End Function




Public Sub FindCustomerByOurRefNo(Ref As String, TPID As Long, Name As String)
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
    cmd.CommandText = "FindCustomerByOurRefNo"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@OurRefWithCustomer", adVarChar, adParamInput, 50, Ref)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Name", adVarChar, adParamOutput, 200, "")
    cmd.Parameters.Append prm
    
    
    cmd.execute lngRecordsAffected
    
    TPID = FNN(cmd.Parameters(1))
    Name = FNS(cmd.Parameters(2))
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub

Public Sub UpdateRemittanceOrderStatus(REMID As Long, RemInvID As Long, Remittance As Double, RemittanceNote As String, Status As String)
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
    cmd.CommandText = "UpdateRemittanceOrderStatus"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@REMID", adInteger, adParamInput, , REMID)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@REMInvID", adInteger, adParamInput, , RemInvID)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@REMAmt", adNumeric, adParamInput, , Remittance)
    prm.NumericScale = 4
    prm.Precision = 20
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@REMNote", adVarChar, adParamInput, 200, RemittanceNote)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Status", adVarChar, adParamInput, 10, Status)
    cmd.Parameters.Append prm
    
    cmd.execute lngRecordsAffected
    
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub

Public Function GetPaymentOrder() As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPaymentOrder"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
    
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetPaymentOrder = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetPaymentOrder"
End Function

Public Sub CreateCreditorsPayments()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'
'    cmd.CommandText = "CreateCreditorsPayments"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 3000
'
'
'    rs.CursorLocation = adUseClient
'    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
'    Set cmd = Nothing
'
'    Set CreateCreditorsPayments = rs

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.CreateCreditorsPayments"
End Sub

Public Sub CreateSupplierClaim(TPID As Long, TRID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim lngRecordsAffected As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "CreateSupplierClaim"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
'
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , TPID)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TRID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.execute lngRecordsAffected
    TRID = FNN(cmd.Parameters(1))
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.CreateSupplierClaim"

End Sub
Public Sub LoadSCDetails(SCID As Long, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
 '   rs.ActiveConnection = oPC.COShort
    rs.Open "SELECT * FROM SCLs Where SCL_ID = " & CStr(SCID) & " Order by GRNDate ", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQL.LoadSCDetails(rs)", rs
End Sub


Public Sub CloseSuppliersClaim(lngTRID As Long, lngTPID As Long)
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim lngRecordsAffected As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "IssueSuppliersClaim"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
'
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , lngTRID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append prm
    
    
    cmd.execute lngRecordsAffected
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Sub

End Sub
Public Function SwitchTriggers(EnableOrDisable As String)
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
Dim lngRecordsAffected As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "SwitchTriggers"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
'
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@DB_Name", adVarChar, adParamInput, 30, Null)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Action", adVarChar, adParamInput, 30, EnableOrDisable)
    cmd.Parameters.Append prm
    
    
    cmd.execute lngRecordsAffected
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    

End Function
Public Function MatchCashbook()
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
    cmd.CommandText = "MatchCashbook"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@RowsAffected", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.execute lngRecordsAffected
    
    lngRecordsAffected = FND(cmd.Parameters(0))
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function

Public Function CalculateStatementBalance(pAccount As String, dteSince As Date, bStartOfDay) As Double
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
    cmd.CommandText = "GetStatementBalance"
    cmd.commandType = adCmdStoredProc
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Account", adVarChar, adParamInput, 50, pAccount)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Since", adDate, adParamInput, , dteSince)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@StartOfDay", adBoolean, adParamInput, , bStartOfDay)
    cmd.Parameters.Append prm
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Balance", adDouble, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.execute lngRecordsAffected
    
    CalculateStatementBalance = FNDBL(cmd.Parameters(3))
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function


Public Function IssueSuppliersClaim(lngTRID As Long, lngTPID As Long) As String
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim lngRecordsAffected As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "IssueSuppliersClaim"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
'
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , lngTRID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , lngTPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Result", adVarChar, adParamOutput, 100)
    cmd.Parameters.Append prm
    
    cmd.execute lngRecordsAffected
    IssueSuppliersClaim = FNS(cmd.Parameters(2))
   
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function

End Function
Public Function FindSUpplierFromDistributorCode(TPACNO As String, TPNAME As String) As Long
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim lngRecordsAffected As Long

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort

    cmd.CommandText = "FindSupplierByCode"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
'
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Code", adVarChar, adParamInput, 100, TPACNO)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@ID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Name", adVarChar, adParamOutput, 100)
    cmd.Parameters.Append prm
    
    cmd.execute lngRecordsAffected
    FindSUpplierFromDistributorCode = FNN(cmd.Parameters(1))
    TPNAME = FNS(cmd.Parameters(2))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function

End Function

Public Function GetPossibleDuplicateCOLByRef(pTPID As Long, pRef As String) As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPossibleDuplicateCOLByRef"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@Ref", adVarChar, adParamInput, 100, pRef)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetPossibleDuplicateCOLByRef = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetPaymentOrder"
End Function
Public Function GetDashboardData( _
        QtyCO_Unissued As Long, _
        QtyCO_UnfulfilledOrders As Long, _
        QtyCO_UnfulfilledOrderLines As Long, _
        QtyCO_UnfulfilledOrderItems As Long, _
        QtyPO_UnIssued As Long, _
        QtyPO_UnfulfilledOrders As Long, _
        QtyPO_UnfulfilledOrderLines As Long, _
        QtyPO_UnfulfilledOrderItems As Long, _
        QtyOrderRequests As Long, _
        QtyDEL_Unissued As Long, _
        QtyCOUnActioned As Long, _
        MostRecentCustomerOrderDate, _
        MostRecentPurchaseOrder, _
        MostRecentDelivery) As Long



    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim lngRecordsAffected As Integer
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetDashboardData"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
    
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@QtyCO_Unissued", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyCO_UnfulfilledOrders", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyCO_UnfulfilledOrderLines", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyCO_UnfulfilledOrderItems", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyPO_UnIssued", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyPO_UnfulfilledOrders", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyPO_UnfulfilledOrderLines", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyPO_UnfulfilledOrderItems", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyDEL_Unissued", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyOrderRequests", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@QtyCOUnActioned", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
        
    Set prm = cmd.CreateParameter("@MostRecentCustomerOrderDate", adDate, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@MostRecentPurchaseOrder", adDate, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@MostRecentDelivery", adDate, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing

        
        
        
    cmd.execute lngRecordsAffected
    QtyCO_Unissued = FNN(cmd.Parameters(0))
    QtyCO_UnfulfilledOrders = FNN(cmd.Parameters(1))
    QtyCO_UnfulfilledOrderLines = FNN(cmd.Parameters(2))
    QtyCO_UnfulfilledOrderItems = FNN(cmd.Parameters(3))
    QtyPO_UnIssued = FNN(cmd.Parameters(4))
    QtyPO_UnfulfilledOrders = FNN(cmd.Parameters(5))
    QtyPO_UnfulfilledOrderLines = FNN(cmd.Parameters(6))
    QtyPO_UnfulfilledOrderItems = FNN(cmd.Parameters(7))
    QtyDEL_Unissued = FNN(cmd.Parameters(8))
    QtyOrderRequests = FNN(cmd.Parameters(9))
    QtyCOUnActioned = FNN(cmd.Parameters(10))
    MostRecentCustomerOrderDate = FND(cmd.Parameters(11))
    MostRecentPurchaseOrder = FND(cmd.Parameters(12))
    MostRecentDelivery = FND(cmd.Parameters(13))
    Set cmd = Nothing
    
    GetDashboardData = 0
    Exit Function
errHandler:
    GetDashboardData = 99
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetDashboardDataCOLByRef"
End Function
