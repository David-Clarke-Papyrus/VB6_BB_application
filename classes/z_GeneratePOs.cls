VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_GenerateTRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Progress(pCount As Long)
Event PBMax(pMax As Long)

Public Property Get POS_StationName() As String
    On Error GoTo errHandler
'See if there is a POS database on this machine  i.e. PBKSFD
Dim DBLocalConn As ADODB.Connection
Dim strLocalSQLServerName As String
Dim rs As ADODB.Recordset

        strLocalSQLServerName = GetIniKeyValue(oPC.LocalFolder & "PBKS.INI", "NETWORK", "POSSQLSERVER", "")
        
        Set DBLocalConn = New ADODB.Connection
        DBLocalConn.Provider = "sqloledb"
        DBLocalConn.open "Data provider=sqloledb;Data Source=" & strLocalSQLServerName & ";Initial Catalog=PBKSFD;User Id=sa" & ";Password=" & "" & "; Network Library=dbmssocn;"
        DBLocalConn.CommandTimeout = 0
        Set rs = New ADODB.Recordset
        rs.open "SELECT StationName FROM tAPPSettings", DBLocalConn, adOpenForwardOnly, adLockOptimistic
        If Not rs.eof Then
            POS_StationName = rs.fields("StationName")
        End If
        rs.Close
        Set rs = Nothing
        DBLocalConn.Close
        Set DBLocalConn = Nothing
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.POS_StationName"
End Property
Public Function GeneratePOs(pSTAFFID As Long, pType As String)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim rsCOL As ADODB.Recordset
Dim oPO As a_PO
Dim oPOL As a_POL
'Dim oProdCode As New z_ProdCode
Dim curTPID As Long
Dim i As Long
Dim lngExistingPOID As Long
Dim OpenResult As Integer
Dim tmpTPID As Long


    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM tCREATEPOS_TEMP LEFT JOIN tDEAL ON DLID = DL_ID JOIN tPRODUCT on P_ID = PID ORDER BY TPID"
    rs.open strSQL, oPC.COShort, adOpenStatic
    If rs.RecordCount > 0 Then RaiseEvent PBMax(rs.RecordCount)
    i = 0
    Do While Not rs.eof
        curTPID = rs!TPID
        'Create PO
        Set oPO = Nothing
        Set oPO = New a_PO
        If oPC.Configuration.AggregatePOs Then
            tmpTPID = CLng(rs!TPID)
            lngExistingPOID = FindOpenPO(tmpTPID)
            If lngExistingPOID > 0 Then
                oPO.Load lngExistingPOID, False
            End If
        End If
        oPO.BeginEdit
        oPO.OrderType = "RC"
        If lngExistingPOID = 0 Then
            oPO.CaptureDate = Now()
            tmpTPID = CLng(rs!TPID)
            oPO.SetTP tmpTPID
            oPO.StaffID = pSTAFFID
            oPO.SetCaptureCurrency oPO.Supplier.DefaultCurrency
        End If
        oPO.ContainsCO = False
        Do While rs!TPID = curTPID
            'Add line
            i = i + 1
            RaiseEvent Progress(i)
            Set oPOL = oPO.POLines.Add
            oPOL.BeginEdit
            oPOL.Discount = FNDBL(rs!DL_Rate)
            oPOL.DealID = FNN(rs!DLID)
            oPOL.PID = FNS(rs!PID)
            oPOL.COLID = FNN(rs!COLID)
            oPO.ContainsCO = (pType = "CUST")
            oPOL.Ref = FNS(rs!Ref)
            oPOL.VATRate = FNN(rs!P_VATRATE)
            oPOL.ProductCode = FNS(rs!code)
            oPOL.ProductCodeF = FormatProductCode(FNS(rs!code), False)
            oPOL.LocalPrice = FNN(rs!Price)
            oPOL.QtyFirm = FNN(rs!QtyFirm)
            oPOL.QtySS = FNN(rs!QtySS)
            'This really does not make any difference all values are already discerned - leaving in for now - no harm done.
            If oPO.CaptureCurrency.SYSNAME = "GBP" Then
                oPOL.ForeignPrice = FNN(rs!ForeignPrice)
            ElseIf oPO.CaptureCurrency.SYSNAME = "USD" Then
                oPOL.ForeignPrice = FNN(rs!ForeignPrice)
            ElseIf oPO.CaptureCurrency.SYSNAME = "EUR" Then
                oPOL.ForeignPrice = FNN(rs!ForeignPrice)
            End If
            oPOL.ApplyEdit
            rs.MoveNext
            If rs.eof Then Exit Do
        Loop
        oPO.SetStatus stInProcess
        oPO.Post
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.GeneratePOs(pStaffID,pType)", Array(pSTAFFID, pType)
End Function
Private Function FindOpenPO(pTPID As Long) As Long
    On Error GoTo errHandler
Dim lng As Long
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
      
    cmd.CommandText = "sp_GetOpenPOID"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@POID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    cmd.execute
    FindOpenPO = IIf(IsNull(cmd.Parameters("@POID")), 0, cmd.Parameters("@POID"))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.FindOpenPO(pTPID)", pTPID
End Function
Public Property Get OPenCSID() As Long
    On Error GoTo errHandler
Dim lng As Long
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
    Set cmd = New ADODB.Command
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
      
    cmd.CommandText = "sp_GetOpenCSID"
    cmd.CommandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@CSID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    cmd.execute
    OPenCSID = IIf(IsNull(cmd.Parameters(0)), 0, cmd.Parameters(0))
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.OPenCSID"
End Property
Public Function GenerateCSandAPPR(pCSID As Long, pAPPID As Long, pAPPRID As Long, pINVID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_ApproReturn_POS"
    cmd.CommandType = adCmdStoredProc

    Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pCSID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@APPID", adInteger, adParamInput, , pAPPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TRSTATUS", adChar, adParamInput, 1, 2)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@IID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@ARID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@POSITION", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    cmd.execute
    
    pINVID = cmd.Parameters("@IID")
    pAPPRID = cmd.Parameters("@ARID")
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.GenerateCSandAPPR(pCSID,pAPPID)", Array(pCSID, pAPPID)
End Function
Public Function GenerateCS(pCSID As Long)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open "SELECT * FROM tCREATECS_TEMP", oPC.COShort, adOpenKeyset, adLockOptimistic
    Do While Not rs.eof
            Set cmd = Nothing
            Set cmd = New ADODB.Command
            cmd.CommandTimeout = 0
            Set cmd.ActiveConnection = oPC.COShort
            cmd.CommandText = "sp_CSL"
            cmd.CommandType = adCmdStoredProc
        
        '1
            Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , pCSID)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '2
            Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , rs.fields("PID"))
            cmd.Parameters.Append prm
            Set prm = Nothing
        '3
            Set prm = cmd.CreateParameter("@Qty", adInteger, adParamInput, , rs.fields("QTY"))
            cmd.Parameters.Append prm
            Set prm = Nothing
        '4
            Set prm = cmd.CreateParameter("@Price", adInteger, adParamInput, , rs.fields("PRICE"))
            cmd.Parameters.Append prm
            Set prm = Nothing
        '5
            Set prm = cmd.CreateParameter("@PriceAlteration", adInteger, adParamInput, , 0)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '6
            Set prm = cmd.CreateParameter("@DiscountRate", adNumeric, adParamInput, , rs.fields("DISCOUNT"))
            prm.NumericScale = 5
            prm.Precision = 4
            cmd.Parameters.Append prm
            Set prm = Nothing
        '7
            Set prm = cmd.CreateParameter("@Discount", adInteger, adParamInput, , 0)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '8
            Set prm = cmd.CreateParameter("@VRate", adSmallInt, adParamInput, , oPC.Configuration.VATRate)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '9
            Set prm = cmd.CreateParameter("@STID", adInteger, adParamInput)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '10
            Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , 0)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '11
            Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '12
            Set prm = cmd.CreateParameter("@COUNTERFOIL", adVarChar, adParamInput, 30, "")
            cmd.Parameters.Append prm
            Set prm = Nothing
        '13
            Set prm = cmd.CreateParameter("@COUNTERFOIL", adVarChar, adParamInput, 20, "")
            cmd.Parameters.Append prm
            Set prm = Nothing
        '14
            Set prm = cmd.CreateParameter("@ID", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '15
            Set prm = cmd.CreateParameter("@pCode", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
        '16
            Set prm = cmd.CreateParameter("@pPosition", adInteger, adParamOutput)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            cmd.execute
            If cmd.Parameters("ID") <> 0 Then
  '              pCSLID = cmd.Parameters("ID")
            End If
            Set cmd = Nothing

    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.GenerateCS(pCSID)", pCSID
End Function
Public Function GenerateINVGDN(pCust As a_Customer, pCOMPID As Long, Optional pSTAFFID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "CreateInvoiceGDNFromAppro"
    cmd.CommandType = adCmdStoredProc

    Set prm = cmd.CreateParameter("@CUSTID", adInteger, adParamInput, , pCust.ID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , pCOMPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@StaffID", adInteger, adParamInput, , pSTAFFID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Rows", adVarChar, adParamInput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    cmd.execute
    
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


'
'
'Dim strSQL As String
'Dim rs As ADODB.Recordset
'Dim oINV As a_Invoice
'Dim oINVLine As a_InvoiceLine
'Dim oProdCode As New z_ProdCode
'Dim curTPID As Long
'Dim i As Long
'Dim oSM As New z_StockManager
'Dim OpenResult As Integer
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    strSQL = "SELECT * FROM tTempToBeInvoiced JOIN tPRODUCT on P_ID = PID WHERE Workstation_ID = '" & oPC.WorkstationID & "' ORDER BY P_Title"
'    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockReadOnly
'    If rs.RecordCount > 0 Then RaiseEvent PBMax(rs.RecordCount)
'    Do While Not rs.EOF
'        Set oINV = Nothing
'        Set oINV = New a_Invoice
'        oINV.BeginEdit
'        oINV.SetBillToAddress pCust.DelToAddress
'        oINV.StaffID = pStaffID
'        oINV.SetCustomer pCust.ID
'        oINV.CaptureDate = Now()
'        If oPC.IssueBookclubReturnDocs = True Then
'            oINV.SetStatus stCOMPLETE
'        End If
'        oINV.setDelToAddress pCust.DelToAddress
'        oINV.ShowVAT = True
'        i = 0
'        Do While Not rs.EOF
'            'Add line
'            i = i + 1
'            RaiseEvent Progress(i)
'            Set oINVLine = oINV.InvoiceLines.Add
'            oINVLine.SetParentInvoice oINV
'            oINVLine.BeginEdit
'            oINVLine.DiscountPercent = FNDBL(rs!Discount)
'            oINVLine.PID = FNS(rs!PID)
'            oINVLine.VATRate = FNDBL(rs!P_VATRATE)
'            oINVLine.APPLID = FNN(rs!APPLID)
'            oINVLine.APPLQTY = FNN(rs!QTYTAKEN)
'        '    oProdcode.Load FNS(rs!Code)
'            oINVLine.Price = FNN(rs!Price)
'            oINVLine.Qty = FNN(rs!QTYTAKEN)
'            oINVLine.ApplyEdit
'            rs.MoveNext
'            If rs.EOF Then Exit Do
'        Loop
'        oINV.Post 'stInProcess
'    Loop
'    rs.Close
'    Set rs = Nothing
'
'    oSM.AUTOGenerateReturnFromInvoice oINV.InvoiceID
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_GenerateTRs.GenerateINV(pCust,pCOMPID)", Array(pCust, pCOMPID)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_GenerateTRs.GenerateINVGDN(pCust,pCOMPID,pStaffID)", Array(pCust, pCOMPID, pStaffID), _
'         EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.GenerateINVGDN(pCust,pCOMPID,pStaffID)", Array(pCust, pCOMPID, pSTAFFID)
End Function

Public Function GenerateAPPR(pCust As a_Customer, pAddr As a_Address, pCOMPID As Long)
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim oAPPR As a_APPR
Dim oAPPRLine As a_APPRL
Dim oProdCode As New z_ProdCode
Dim curTPID As Long
Dim i As Long
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    strSQL = "SELECT * FROM tCREATEAPPR_TEMP JOIN tAPPL ON APPL_ID = APPLID JOIN tPRODUCT on P_ID = APPL_P_ID ORDER BY P_Title"
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockReadOnly
    If rs.RecordCount > 0 Then RaiseEvent PBMax(rs.RecordCount)
    Do While Not rs.eof
        Set oAPPR = Nothing
        Set oAPPR = New a_APPR
        oAPPR.BeginEdit
        oAPPR.SetCustomer pCust.ID
        oAPPR.COMPID = pCOMPID
        oAPPR.IssDate = Now()
        oAPPR.SetStatus stInProcess
        i = 0
        Do While Not rs.eof
            'Add line
            If rs!Qty > 0 Then
                i = i + 1
    '            RaiseEvent Progress(i)
                Set oAPPRLine = oAPPR.APPRLines.Add
                oAPPRLine.SetParentAPPR oAPPR
                oAPPRLine.BeginEdit
                oAPPRLine.Qty = rs!Qty
                oAPPRLine.APPLID = rs!APPLID
                oAPPRLine.ApplyEdit
            End If
            rs.MoveNext
            If rs.eof Then Exit Do
        Loop
        oAPPR.ApplyEdit
        oAPPR.Post
    Loop
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_GenerateTRs.GenerateAPPR(pCust,pAddr,pCOMPID)", Array(pCust, pAddr, pCOMPID)
End Function

