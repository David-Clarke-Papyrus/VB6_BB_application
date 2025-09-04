VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_POSCLIConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mCurrentSecurityCode As String
Public DBLocalConn As ADODB.Connection
Public DBMainConn As ADODB.Connection
Public DBPropertyConn As ADODB.Connection
Dim rsProperty As New ADODB.Recordset
Dim rsRoundingRule As New ADODB.Recordset
Dim rsMultibuys As New ADODB.Recordset
Dim bUseA4Printer As Boolean
Dim f As Long

Dim oGD As z_GetData
Dim sDBPath As String
Dim strSharedFolderRoot As String
Dim strDDB As String
Dim bPrinterHasCutter As Boolean
Dim strServer As String
Dim strTillCode As String
Dim strZSessionID As Variant
Dim strCurrencyFormat As String
Dim lngCurrencyDivisor As Long
Dim strLocalSQLServerName As String
Dim strPassword As String
Dim strPBKSSERVERMACHINE As String
Dim strBookfindRoot As String
Dim strLocalRootFolder As String
Dim strPCNAme As String
Dim strServername As String
Dim lngLoyaltyRate As Long
Dim strPOSCompanyName As String
Dim strPOSBranchName As String
Dim strPOSBranchAddress As String
Dim strPOSReceiptMessage As String
Dim strPOSEMAILAddress As String
Dim strCashDrawerPort As String
Dim CNLimit As Long
Dim bSendsCR As Boolean
Dim strCOMPORTSettings As String
Dim strCOMMPORTNumber As String
Dim bPOSACTIVE As Boolean
Dim strStationName As String
Dim strSlipPrinterName As String
Dim strTillDrawerName As String
Dim strCOMPANYNAME As String
Dim strBRANCHNAME As String
Dim strBRANCHADDRESS As String
Dim strRECEIPTMESSAGE As String
Dim strEMAILADDRESS As String
Dim strDriveDrawer As String
Dim strPrintSlips As String
Dim iInvoiceCopyCount As Integer
Dim iDepositCopyCount As Integer
Dim iCreditNoteCopyCount As Integer
Dim iReturnCopyCount As Integer
Dim iApproCopyCount As Integer
Dim iOrderCopyCount As Integer
Dim iPettyCashCount As Integer
Dim strVouchersAccepted As String
Dim strDiscountsGiven As String
Dim strPettyCashSet As String
Dim strCurrencyNotes As String
Dim lngDefaultDeposit As Long
Dim strLocalConnectionString As String
Dim strDatabasename As String
Dim bUseTestDatabase As Boolean
Dim mLoyaltyCode As String
Dim mBookclubCode As String
Dim bEnableXEND As Boolean

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Property Let UseTestDatabase(val As Boolean)
    bUseTestDatabase = val
End Property
Public Property Get UseA4Printer() As Boolean
    UseA4Printer = bUseA4Printer
End Property
Public Property Get EnableXEND() As Boolean
    EnableXEND = GetProperty("EnableXEND")
End Property

Public Property Let DatabaseName(val As String)
    strDatabasename = val
End Property
Public Property Get DatabaseName() As String
    DatabaseName = strDatabasename
End Property
Public Property Get POSMessageTimeout() As String
    POSMessageTimeout = GetProperty("POSMessageTimeout")
End Property
Public Property Get ServerIPAddress() As String
    ServerIPAddress = GetProperty("ServerIPAddress")
End Property
Public Property Get PrintSlips() As Boolean
    PrintSlips = (GetProperty("PrintSlipsTF") = "TRUE")
End Property
Public Property Get UseCashDrawer() As Boolean
    UseCashDrawer = (GetProperty("UseCashDrawerTF") = "TRUE")
End Property
Public Property Get UsageContext() As String
    UsageContext = (GetProperty("UsageContext"))
End Property

Public Property Get DriveDrawer() As String
    DriveDrawer = UCase(GetProperty("DriveDrawerTF"))
End Property
Public Property Get StationName() As String
    StationName = GetProperty("StationName")
End Property
Public Property Get DefaultDeposit() As Long
    DefaultDeposit = CInt(GetProperty("DefaultDeposit"))
End Property
'Public Property Get DefaultDeposit_Dec() As Double
''MsgBox "Default deposit:" & CStr(DefaultDeposit)
'    DefaultDeposit_Dec = CDbl(DefaultDeposit) / oPC.CurrencyDivisor
'End Property
'Public Property Get DefaultDeposit_DecF() As String
'    DefaultDeposit_DecF = Format(DefaultDeposit_Dec, "###,##0.00")
'MsgBox "Default DefaultDeposit_DecF:" & CStr(DefaultDeposit_DecF)
    
'End Property

Public Property Get CashDrawerPort() As String
    CashDrawerPort = GetProperty("CashDrawerPort")
End Property
Public Property Get COMPORTSettings() As String
    COMPORTSettings = GetProperty("COMPortSettings")
End Property
Public Property Get PrinterHasCutter() As Boolean
    PrinterHasCutter = (GetProperty("PrinterHasCutterTF") = "TRUE")
End Property
Public Property Get POSActive() As Boolean
    POSActive = (GetProperty("POSActiveTF") = "TRUE")
End Property
Public Function VoucherSet() As String
    VoucherSet = GetProperty("VouchersAccepted")
End Function
Public Function DiscountSet() As String
    DiscountSet = GetProperty("DiscountsGiven")
End Function
Public Function PettyCashSet() As String
    PettyCashSet = GetProperty("PettyCashAccounts")
End Function

Public Property Get InvoiceCopyCount() As Integer
    InvoiceCopyCount = CInt(GetProperty("InvoiceCopyCount"))
End Property
Public Property Get DepositCopyCount() As Integer
    DepositCopyCount = CInt(GetProperty("DepositCopyCount"))
End Property
Public Property Get DepositRefundCopyCount() As Integer
    DepositRefundCopyCount = CInt(GetProperty("DepositRefundCopyCount"))
End Property
Public Property Get CreditNoteCopyCount() As Integer
    CreditNoteCopyCount = CInt(GetProperty("CreditNoteCopyCount"))
End Property
Public Property Get ReturnCopyCount() As Integer
    ReturnCopyCount = CInt(GetProperty("ReturnCopyCount"))
End Property
Public Property Get ApproCopyCount() As Integer
    ApproCopyCount = CInt(GetProperty("ApproCopyCount"))
End Property
Public Property Get ReceiptCopyCount() As Integer
    ReceiptCopyCount = CInt(GetProperty("ReceiptCopyCount"))
End Property
Public Property Get OrderCopyCount() As Integer
    OrderCopyCount = CInt(GetProperty("OrderCopyCount"))
End Property
Public Property Get PettyCashCopyCount() As Integer
    PettyCashCopyCount = CInt(GetProperty("PettyCashCopyCount"))
End Property
Public Property Get LoyaltyVoucherCopyCount() As Integer
    LoyaltyVoucherCopyCount = CInt(GetProperty("LoyaltyVoucherCopyCount"))
End Property
Public Property Get AccountSaleCopyCount() As Integer
    AccountSaleCopyCount = CInt(GetProperty("AccountSaleCopyCount"))
End Property



Public Property Get CreditNoteLimit() As Long
    CreditNoteLimit = CLng(GetProperty("CreditNoteLimit"))
End Property
Public Property Get POSCompanyName() As String
    POSCompanyName = GetProperty("CompanyName")
End Property
Public Property Get POSBranchName() As String
    POSBranchName = GetProperty("BranchName")
End Property
Public Property Get POSBranchAddress() As String
    POSBranchAddress = GetProperty("BranchAddress")
End Property
Public Property Get POSReceiptMessage() As String
    POSReceiptMessage = GetProperty("ReceiptMessage")
End Property
Public Property Get POSEMAILAddress() As String
    POSEMAILAddress = GetProperty("EmailAddress")
End Property
Public Property Get Printername() As String
    Printername = GetProperty("SlipPrinterName")
End Property
Public Property Get TillDrawerName() As String
    TillDrawerName = GetProperty("TillDrawerName")
End Property
Public Property Get LoyaltyRate() As Long
    LoyaltyRate = CInt(GetProperty("LoyaltyRate"))
End Property
Public Property Get CurrencyFormat() As String
    CurrencyFormat = GetProperty("CurrencyFormat")
End Property
Public Property Get CurrencyDivisor() As Long
    CurrencyDivisor = CLng(GetProperty("CurrencyDivisor"))
End Property
Public Sub SetupZSession(pSupervisorID As Long, pName As String)
    On Error GoTo errHandler
    If CurrentZSessionID > "" Then
        Set oZSession = New z_ZSession
        oZSession.Load CurrentZSessionID
    End If
    If oZSession Is Nothing Then
        Set oZSession = Nothing
        Set oZSession = New z_ZSession
        oZSession.Start_Z_Session pSupervisorID, pName
    ElseIf oZSession.EndDate > CDate("2005-01-01") Then
        Set oZSession = Nothing
        Set oZSession = New z_ZSession
        oZSession.Start_Z_Session pSupervisorID, pName
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.SetupZSession(pSupervisorID,pName)", Array(pSupervisorID, pName)
End Sub
Public Function OpenLocalDatabase() As Integer
    On Error GoTo errHandler
    If DBLocalConn Is Nothing Then
        Set DBLocalConn = New ADODB.Connection
    End If
    
    If DBLocalConn.State <> adStateClosed Then DBLocalConn.Close
    DBLocalConn.Open strLocalConnectionString
    DBLocalConn.CommandTimeout = 360
    OpenLocalDatabase = 0
    
    Exit Function
errHandler:
    ErrPreserve
    If oPC.DatabaseName = "PBKSFD_TEST" Then
        MsgBox "You are trying to connect to a test database that has not been prepared. You will be connected to the LIVE database.", vbExclamation + vbOKOnly, "Can't do this"
        oPC.DatabaseName = "PBKSFD"
        strLocalConnectionString = "Provider=SQLOLEDB.1;Data provider=sqloledb;Data Source=" & Me.LocalSQLServerName & ";Initial Catalog=" & oPC.DatabaseName & ";User Id=sa;Password=" & strPassword & ";"""
        SaveSetting "POS", "StartDatabaseName", "DBNAME", "PBKSFD"

        Resume
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.OpenLocalDatabase", , , , "strLocalConnectionString", Array(strLocalConnectionString)
End Function
Public Sub CloseLocalDatabase()
    On Error GoTo errHandler
    DBLocalConn.Close
    Set DBLocalConn = Nothing
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 91 Then 'Already closed
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.CloseLocalDatabase"
End Sub

Public Function dbConnect()
    On Error GoTo errHandler
    strLocalConnectionString = "Provider=SQLOLEDB.1;Data provider=sqloledb;Data Source=" & Me.LocalSQLServerName & ";Initial Catalog=" & oPC.DatabaseName & ";User Id=sa;Password=" & strPassword & ";"""
 '''''   strLocalConnectionString = "Provider=SQLOLEDB.1;Data provider=sqloledb;Data Source=5.136.116.79\PBKSINSTANCE2;Initial Catalog=PBKSFD;User Id=sa;Password=car;"""
    OpenLocalDatabase
    If DBPropertyConn Is Nothing Then
        Set DBPropertyConn = New ADODB.Connection
        DBPropertyConn.Open "Provider=SQLOLEDB.1;Data provider=sqloledb;Data Source=" & Me.LocalSQLServerName & ";Initial Catalog=" & oPC.DatabaseName & ";User Id=sa;Password=" & strPassword & ";"""
       '''' DBPropertyConn.Open strLocalConnectionString
        DBPropertyConn.CommandTimeout = 0
    End If
    ConfigureFromDB  'looka for CurrentZSessionID among other things
EXIT_Handler:
    Exit Function
errHandler:
    bIsConnectedtoServer = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.dbConnect", , , , "strLocalConnectionString", Array(strLocalConnectionString)
End Function
Public Function dbConnectMain()
Dim strConnection As String
    On Error GoTo errHandler
    If bTryToCOnnectToMainServer Then
        If DBMainConn Is Nothing Then
            Set DBMainConn = New ADODB.Connection
            DBMainConn.Provider = "sqloledb"
            strConnection = "Data provider=sqloledb;Data Source=" & Me.MainSQLServerName & ";Initial Catalog=" & mDBName & ";User Id=sa;Password=" & strPassword & ";"
          ' strConnection = "Data provider=sqloledb;Data Source=" & "CENTRALBACKUP\PBKSINSTANCE2" & ";Initial Catalog=" & mDBName & ";User Id=sa;Password=" & strPassword & ";"
          '  DBMainConn.Close
        End If
        If DBMainConn.State = 0 Then
            DBMainConn.Open strConnection
            DBMainConn.CommandTimeout = 0
        End If
         bIsConnectedtoServer = True
    End If
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    bIsConnectedtoServer = False
    bTryToCOnnectToMainServer = False
    ErrorIn "z_POSCLIConnection.dbConnectMain", , , , "strConnection", Array(strConnection)
    Exit Function
    If ErrMustStop Then Debug.Assert False: Resume
   ' ErrSaveToFile
    ErrorIn "z_POSCLIConnection.dbConnectMain", , , , strConnection, Array(strConnection)
End Function
Public Function dbMainDisConnect() As Integer
    On Error GoTo errHandler
    If Not DBMainConn Is Nothing Then
        If DBMainConn.State = adStateOpen Then DBMainConn.Close
        Set DBMainConn = Nothing
        bIsConnectedtoServer = False
    End If
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.dbMainDisConnect"
End Function
Public Property Get IsConnectedToServer() As Boolean
    IsConnectedToServer = bIsConnectedtoServer
End Property
Public Property Get ServerConnectionRef() As ADODB.Connection
    Set ServerConnectionRef = DBLocalConn
End Property

Public Function dbCloseLocalConnect() As Integer
    On Error GoTo errHandler
  '  MsgBox "Close Local Connection"
    If Not DBLocalConn Is Nothing Then
        If DBLocalConn.State = adStateOpen Then DBLocalConn.Close
        Set DBLocalConn = Nothing
    End If
    Set oZSession = Nothing
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.dbCloseLocalConnect"
End Function

Friend Property Get DBPath() As String
    On Error GoTo errHandler
    DBPath = sDBPath
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.DBPath"
End Property

Public Property Get SharedFolderRoot() As String
    On Error GoTo errHandler
    SharedFolderRoot = strSharedFolderRoot
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.SharedFolderRoot"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler

  '  InitializeSettings
    
    Set oPC = Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.Class_Initialize"
End Sub
Public Sub InitializeSettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String

    
    strPCNAme = Trim(Me.NameOfPC)
        
    strLocalRootFolder = "C:\PBKS"
    If bUseTestDatabase Then
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSERVERMACHINE", strPCNAme)
    Else
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCNAme)
    End If
    If IsNetConnectionAlive Then
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strSharedFolderRoot = "C:\PBKS"
    End If

    If bUseTestDatabase Then
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSQLSERVER", strPCNAme & "\PBKSInstance2")
    Else
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", strPCNAme & "\PBKSInstance2")
    End If
    strLocalSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "POSSQLSERVER", strPCNAme & "\PBKSInstance2")
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    
    SynchronizeTOD strPBKSSERVERMACHINE

   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.InitializeSettings"
End Sub
Public Property Get LocalRootFolder() As String
    LocalRootFolder = strLocalRootFolder
End Property
Public Property Get LocalSQLServerName() As String
    LocalSQLServerName = strLocalSQLServerName
End Property
Public Property Get MainSQLServerName() As String
    MainSQLServerName = strServername
End Property
Public Property Get Password() As String
    Password = strPassword
End Property

Public Property Get PBKSSERVERMACHINE() As String
    PBKSSERVERMACHINE = strPBKSSERVERMACHINE
End Property
Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Property

Public Property Let CurrentSecurityCode(val As String)
    mCurrentSecurityCode = val
End Property
Public Property Get CurrentSecurityCode() As String
    CurrentSecurityCode = mCurrentSecurityCode
End Property

Public Property Get GD() As z_GetData
    Set GD = oGD
End Property
Public Property Get ZSession() As z_ZSession
    Set ZSession = oZSession
End Property

'Public Property Let TillCode(val As String)
'    sTillCode = val
'End Property
'Public Property Get TillCode() As String
'    TillCode = Trim(sTillCode)
'End Property
Public Property Get CurrentZSessionID() As String
    CurrentZSessionID = strZSessionID
End Property
Public Property Get ExchangeNumber() As String
    ExchangeNumber = lngExchangeNumber
End Property



Public Function ExchangeNumberIncrement(pCurrent As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    
    oPC.OpenLocalDatabase
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = DBLocalConn
    cmd.CommandText = "ExchangeCountIncrement"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@Current", adInteger, adParamOutput)
    cmd.Parameters.Append par
    cmd.Execute
    pCurrent = cmd.Parameters(0).Value
    lngExchangeNumber = pCurrent
    Set par = Nothing
    Set cmd = Nothing
    
    oPC.CloseLocalDatabase
    
End Function
Private Function ConfigureFromDB() As Boolean
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim sSQL As String
    
    oPC.OpenLocalDatabase
    
    sSQL = "SELECT * FROM tAppSettings"
    rs.Open sSQL, oPC.DBLocalConn, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        strZSessionID = FNS(rs!OPENZSessionID)
        lngExchangeNumber = FNN(rs!EX_Number)
        f = FNN(rs!f)
        mLoyaltyCode = FNS(rs!Loyalty_Code)
        mBookclubCode = FNS(rs!BC_Code)
    End If
    rs.Close
    Set rs = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.ConfigureFromDB"
End Function
Public Function GetProperty(pKey As String) As String
    On Error GoTo errHandler
    rsProperty.MoveFirst
    rsProperty.Find "PropertyKey = '" & pKey & "'"
    GetProperty = Trim(CStr(rsProperty.Fields(1)))
    Exit Function
errHandler:
    ErrorIn "z_POSCLIConnection.GetProperty(pKey)", pKey
End Function
Public Function GetLoyaltyCode() As String
    GetLoyaltyCode = FNS(mLoyaltyCode)
End Function
Public Function GetBookClubCode() As String
    GetBookClubCode = mBookclubCode
End Function
Public Function GetRoundingRule(pVal As Long) As Long
Dim T As Long
    On Error GoTo errHandler
'Dim T As Long
'    rsRoundingRule.Filter = ""
'    rsRoundingRule.MoveFirst
'
'    rsRoundingRule.Filter = "RR_LowerBound < " & pVal & " And RR_UpperBound >= " & pVal
'    If rsRoundingRule.EOF Then
'        GetRoundingRule = 0
'    Else
'        GetRoundingRule = CLng(rsRoundingRule.Fields(2))
'    End If
    
    
    rsRoundingRule.Filter = ""
    If rsRoundingRule Is Nothing Then loadRoundingRules
    If rsRoundingRule.RecordCount = 0 Then
        loadRoundingRules
    End If
    If pVal < 0 Then
        T = pVal * -1
    Else
        T = pVal
    End If
    GetRoundingRule = pVal
    rsRoundingRule.MoveFirst
    rsRoundingRule.Filter = "RR_LowerBound < " & T & " And RR_UpperBound >= " & T
    If rsRoundingRule.EOF Then
        GetRoundingRule = 0
    Else
        GetRoundingRule = CLng(rsRoundingRule.Fields(2))
    End If
'=======
'    If rsRoundingRule Is Nothing Then loadRoundingRules
'    If rsRoundingRule.RecordCount = 0 Then
'        loadRoundingRules
'    End If
'    If pVal < 0 Then
'        T = pVal * -1
'        GetRoundingRule = pVal
'        rsRoundingRule.MoveFirst
'        rsRoundingRule.Filter = "RR_LowerBound < " & T & " And RR_UpperBound >= " & T
'        If rsRoundingRule.EOF Then
'            GetRoundingRule = 0
'        Else
'            GetRoundingRule = CLng(rsRoundingRule.Fields(2))
'        End If
'>>>>>>> .merge-right.r375
'=======
'    If rsRoundingRule Is Nothing Then loadRoundingRules
'    If rsRoundingRule.RecordCount = 0 Then
'        loadRoundingRules
'    End If
'    If pVal < 0 Then
'        T = pVal * -1
'        GetRoundingRule = pVal
'        rsRoundingRule.MoveFirst
'        rsRoundingRule.Filter = "RR_LowerBound < " & T & " And RR_UpperBound >= " & T
'        If rsRoundingRule.EOF Then
'            GetRoundingRule = 0
'        Else
'            GetRoundingRule = CLng(rsRoundingRule.Fields(2))
'        End If
'>>>>>>> .merge-right.r375
'    Else
'        GetRoundingRule = pVal
'        rsRoundingRule.MoveFirst
'        rsRoundingRule.Filter = "RR_LowerBound < " & pVal & " And RR_UpperBound >= " & pVal
'        If rsRoundingRule.EOF Then
'            GetRoundingRule = 0
'        Else
'            GetRoundingRule = CLng(rsRoundingRule.Fields(2))
'        End If
'   End If
    
    
    
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.GetRoundingRule(pVal)", pVal
End Function
Public Function LoadProperties() As Boolean
    On Error GoTo errHandler
Dim sSQL As String
'   If (32767 And 32767) = 32767 Then
'  ' If (16384 & 32767) Then
'    MsgBox "true"
'Else
'    MsgBox "false"
'End If
    
    
    
    
    oPC.OpenLocalDatabase
    
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.Open sSQL, oPC.DBPropertyConn, adOpenKeyset, adLockOptimistic
    Set rsProperty.ActiveConnection = Nothing
    
    oPC.CloseLocalDatabase
    bUseA4Printer = UCase(oPC.GetProperty("UseA4Printer")) = "TRUE"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.LoadProperties"
End Function
Public Function loadRoundingRules() As Boolean
Dim sSQL As String
    
    oPC.OpenLocalDatabase
    
    sSQL = "SELECT RR_LowerBound,RR_UpperBound,RR_RoundTo FROM tRoundingRule"
    Set rsRoundingRule = New ADODB.Recordset
    rsRoundingRule.CursorLocation = adUseClient
    rsRoundingRule.Open sSQL, oPC.DBPropertyConn, adOpenKeyset, adLockOptimistic
    Set rsRoundingRule.ActiveConnection = Nothing
    oPC.CloseLocalDatabase
   
End Function

Public Function loadMultibuys() As Boolean
Dim sSQL As String
    
    oPC.OpenLocalDatabase
    
    sSQL = "SELECT * FROM tMultibuys"
    Set rsMultibuys = New ADODB.Recordset
    rsMultibuys.CursorLocation = adUseClient
    rsMultibuys.Open sSQL, oPC.DBPropertyConn, adOpenKeyset, adLockOptimistic
    Set rsMultibuys.ActiveConnection = Nothing
    oPC.CloseLocalDatabase
   
End Function
Public Function GetMultibuy(pCODE As String, pPrice As Long, pQty As Long) As Boolean
    On Error GoTo errHandler
Dim T As Long
    
    rsMultibuys.Filter = ""
    If rsMultibuys Is Nothing Then loadMultibuys
    If rsMultibuys.RecordCount = 0 Then
        loadMultibuys
    End If
    GetMultibuy = False
    rsMultibuys.MoveFirst
    rsMultibuys.Filter = "MB_SYSTEMCODE  = '" & pCODE & "'"
    If Not rsMultibuys.EOF Then
        GetMultibuy = True
        pPrice = FNN(rsMultibuys.Fields("MB_PRICE"))
        pQty = FNN(rsMultibuys.Fields("MB_QtyGroup"))
    End If
    rsMultibuys.Filter = ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.GetMultibuy(pCode,pPrice,pQty)", Array(pCODE, pPrice, pQty)
End Function

Public Property Get MultibuyRS() As ADODB.Recordset
    Set MultibuyRS = rsMultibuys
End Property

Public Property Get PCNAme() As String
    PCNAme = strPCNAme
End Property

Public Property Get CurrencyNotes()
Dim ar() As Integer
Dim ars() As String
Dim i As Integer
    strCurrencyNotes = GetProperty("CurrencyNotes")
    ars = Split(strCurrencyNotes, ",")
    ReDim ar(UBound(ars))
    
    For i = 1 To UBound(ars)
        ar(i) = CInt(ars(i))
    Next
    CurrencyNotes = ar
End Property

Private Sub Class_Terminate()
On Error Resume Next
    rsProperty.Close
    Set rsProperty = Nothing
End Sub

Public Property Get SecurityNumber() As Long
    SecurityNumber = f
End Property

