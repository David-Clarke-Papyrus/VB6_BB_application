VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PapyConn"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mSecurity As String
Private dteTimeStarted As Date
Private strPapyConn As String
Private strServername As String
Private strPassword As String
Private strUsername As String
Dim strPOSDBName As String
Private strSharedFolderRoot As String
'Private strBookfindRoot As String
'Dim strBookfindFacet As String
Dim strBackupFolder As String
Dim strDDB As String
Private oConfig As a_Configuration
Private strApplicationDatabase As String
Private bBFLoaded As Boolean
Private bEnableBookCLubReturn As Boolean
Private bShowBudget As Boolean
Dim bCanGenerateApprosFromOrderFulfilment As Boolean
'Private oBookfindmanager As z_BookfindManager
Private flgDbObjConnected As Boolean
'Private strDatabaseFullPath As String
Private cnPapy As ADODB.Connection
Private cnPapyShort As ADODB.Connection
Private cnScheduler As ADODB.Connection
'Private oBookFind  As z_BookfindManager
Private mCurrentSecurityCode As String
Event RecordAdded(Success As Boolean)
Event Status(pMsg As String)
Dim bMultiStore As Boolean
Dim strLocalRootFolder As String
Dim strRemovableBackupShareName As String
Dim strPBKSSERVERMACHINE As String
Dim strMarkCustomerOrderLinesOnPOLines As String
Dim strDeliveryStyle As String

Dim strBackupMachine As String
Dim strCDType As String
Dim strLabelPrinter As String
Dim bCompressBackup As Boolean
Dim strBackupMedium As String
Dim mlngWSID As Long
Dim strPCName As String
Dim bEDIEnabled As Boolean
Dim bShowAllAppros As Boolean
Dim strDefaultAreaCode As String
Dim iTimerInterval As Integer
Dim bSendsCR As Boolean
Dim strCOMPORTSettings As String
Dim strCOMMPORTNumber As String
Dim strSpreadsheetType As String
Dim bPOSActive As Boolean
Dim bAllowInvoiceDateOverride As Boolean
Dim bAllowPODateOverride As Boolean
'Dim bBOOKFINDISBN13ENABLED As Boolean
Dim bAllowGeneralStock As Boolean
Dim bSupportsMultibuys As Boolean
Dim bSupplierBasedCurrencyConversion As Boolean
Dim bSupportsCatalogue As Boolean
Dim bShowWordstockSales As Boolean
Dim strIssueBookclubReturnDocs As String
Dim bTransferIsExVAT As Boolean
Dim bAllowsSSInvoicing As Boolean
Dim dblSupplierInvoiceTolerance As Double
Dim iRoundPriceTo As Integer
Dim iKeepTemporaryfilesFor_n_Days As Integer
Dim strVouchersAccepted As String
Dim strVouchersReportTogether As String
Dim strInvoiceTotalSequence As String
Dim iAllowAntiquarionSearch As Integer
Dim strMainConnectionString As String
Dim bSetSupplierIDFROMPO As Boolean
Dim bSetSupplierIDFROMDEL As Boolean
Dim rsProperty As New ADODB.Recordset
Dim rsBudgets As New ADODB.Recordset
Dim rsLinesClipboard As New ADODB.Recordset
Dim bSetPricesInGRN As Boolean
'Dim strQuote As String
Dim bUseQuotation As Boolean
Dim bHideLocalSKUonINV As Boolean
Dim bAllowsInvoicePicking As Boolean
Dim bRunsAccountsTF As Boolean
Dim bSupportsBookClubsTF As Boolean
Dim bSupportsLoyaltyCustomersTF As Boolean
Dim bEmailPO As Boolean
Dim bEmailCO As Boolean
Dim bEMailINV As Boolean
Dim bEMailQuote As Boolean
Dim bEmailPOShowHTML As Boolean
Dim bEmailSalesOrderShowHTML As Boolean
Dim bEmailShowHTML As Boolean
Dim bEmailQuoteShowHTML As Boolean
Dim bBlindCashup As Boolean
Dim bUsesOutlookForPOEmail As Boolean
Dim bUsesOutlookForINVEmail As Boolean
Dim bUsesOutlookForCOEmail As Boolean
Dim bUsesOutlookForQuoteEmail As Boolean
Dim bAllowsZeroDiscountPOs As Boolean
Dim bUpdatePriceOnForeignDelivery As Boolean
Dim bAllowZeropricedPOLines As Boolean
Dim bAllowZeropricedCOLines As Boolean
Dim bExportToPastelEnabled As Boolean
Dim bSupportsUNISA As Boolean
Dim bUniqueProducts As Boolean
Dim strContraAccount_Inv As String
Dim strContraAccount_SInv As String
Dim lngTranslationOffset As Long
Dim bISSUE_PO_ON_THIS_WS As Boolean
Dim bIsFrontDeskWorkstation As Boolean
Dim strInternetDialup As String
Dim strConnectionName As String
Dim mCL As String
Dim strEmailFrom As String
Dim strEmailAddressForTesting As String
Dim strPOSubject As String
Dim strInvoiceSubject As String
Dim strSenderName As String
Dim bTestMode As Boolean
Dim bPrintPackingSlip As Boolean
Dim strDimension As String
Dim strEmailForDBWarnings As String
Dim strCodePrefixForInternalBarcodesEAN As String

Dim bIncludeSupplierFeatures As Boolean
Dim bEnforceSupplierOnProduct As Boolean
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long


Public Property Get CodePrefixForInternalBarcodesEAN() As String
    CodePrefixForInternalBarcodesEAN = strCodePrefixForInternalBarcodesEAN
End Property
Public Property Get EnforceSupplierOnProduct() As Boolean
    EnforceSupplierOnProduct = bEnforceSupplierOnProduct
End Property
Public Property Get IncludeSupplierFeatures() As Boolean
    IncludeSupplierFeatures = bIncludeSupplierFeatures
End Property
Public Property Get EmailForDBWarnings() As String
    EmailForDBWarnings = strEmailForDBWarnings
End Property
Public Property Get DeliveryStyle() As String
    DeliveryStyle = strDeliveryStyle
End Property
Public Property Get EmailPOShowHTML() As Boolean
    EmailPOShowHTML = bEmailPOShowHTML
End Property
Public Property Get CanGenerateApprosFromOrderFulfilment() As Boolean
    CanGenerateApprosFromOrderFulfilment = bCanGenerateApprosFromOrderFulfilment
End Property

Public Property Get EmailINVShowHTML() As Boolean
    EmailINVShowHTML = bEmailShowHTML
End Property
Public Property Get EmailSalesOrderShowHTML() As Boolean
    EmailSalesOrderShowHTML = bEmailSalesOrderShowHTML
End Property
Public Property Get EmailQuoteShowHTML() As Boolean
    EmailQuoteShowHTML = bEmailQuoteShowHTML
End Property

Public Property Get TranslationOffset() As Long
   TranslationOffset = lngTranslationOffset
End Property
Public Property Get EmailAddressForTesting() As String
    EmailAddressForTesting = strEmailAddressForTesting
End Property
Public Property Get EmailFrom() As String
    EmailFrom = strEmailFrom
End Property
Public Property Get InvoiceSubject() As String
    InvoiceSubject = strInvoiceSubject
End Property
Public Property Get POSubject() As String
    POSubject = strPOSubject
End Property
Public Property Get SenderName() As String
    SenderName = strSenderName
End Property

'Public Property Get Username() As String
'    Username = strUsername
'End Property
Public Property Get TestMode() As Boolean
    TestMode = bTestMode
End Property
Public Property Get SupplierBasedCurrencyConversion() As Boolean
    SupplierBasedCurrencyConversion = bSupplierBasedCurrencyConversion
End Property

Public Property Get ShowBudget() As Boolean
    ShowBudget = bShowBudget
End Property

Public Property Get ShowWordstockSales() As Boolean
    ShowWordstockSales = bShowWordstockSales
End Property
Public Property Get PrintPackingSlip() As Boolean
    PrintPackingSlip = bPrintPackingSlip
End Property
Public Property Get ConnectionString() As String
    ConnectionString = strMainConnectionString
End Property
Public Property Get HideLocalSKUonINV() As Boolean
    HideLocalSKUonINV = bHideLocalSKUonINV
End Property
Public Property Get ContraAccount_SInv() As String
    ContraAccount_SInv = Trim(strContraAccount_SInv)
End Property
Public Property Get fSecurity() As Long
    If IsNumeric(FNS(mSecurity)) Then
        fSecurity = CLng(mSecurity)
    Else
        fSecurity = 0
    End If
End Property

Public Property Get ContraAccount_Inv() As String
    ContraAccount_Inv = Trim(strContraAccount_Inv)
End Property
Public Property Get NewQuotation() As String
    NewQuotation = Trim(GetQuotation)
End Property

Public Property Get RoundPriceTo() As Integer
    RoundPriceTo = iRoundPriceTo
End Property
Public Property Get KeepTemporaryfilesFor_n_Days() As Integer
    KeepTemporaryfilesFor_n_Days = iKeepTemporaryfilesFor_n_Days
End Property

Public Property Get SupplierInvoiceTolerance() As Double
    SupplierInvoiceTolerance = dblSupplierInvoiceTolerance
End Property

Public Property Get AllowsZeroDiscountPOs() As Boolean
    AllowsZeroDiscountPOs = bAllowsZeroDiscountPOs
End Property
Public Property Get UpdatePriceOnForeignDelivery() As Boolean
    UpdatePriceOnForeignDelivery = bUpdatePriceOnForeignDelivery
End Property

Public Property Get TransferIsExVAT() As Boolean
    TransferIsExVAT = UCase(bTransferIsExVAT)
End Property
Public Property Get AllowsSSInvoicing() As Boolean
    AllowsSSInvoicing = bAllowsSSInvoicing
End Property
Public Property Get AllowsInvoicePicking() As Boolean
    AllowsInvoicePicking = bAllowsInvoicePicking
End Property
Public Property Get RunsAccountsTF() As Boolean
    RunsAccountsTF = bRunsAccountsTF
End Property
Public Property Get SupportsBookClubsTF() As Boolean
    SupportsBookClubsTF = bSupportsBookClubsTF
End Property
Public Property Get SupportsLoyaltyCustomersTF() As Boolean
    SupportsLoyaltyCustomersTF = bSupportsLoyaltyCustomersTF
End Property
Public Property Get VouchersAccepted() As String
    VouchersAccepted = UCase(strVouchersAccepted)
End Property

Public Property Get VouchersReportTogether() As String
    VouchersReportTogether = UCase(strVouchersReportTogether)
End Property

Public Property Get CDType() As String
    CDType = Trim(strCDType)
End Property
Public Property Get LabelPrinter() As String
    LabelPrinter = Trim(strLabelPrinter)
End Property
Public Property Get COMPORTSettings() As String
    COMPORTSettings = strCOMPORTSettings
End Property
Public Property Get COMMPORTNumber() As String
    COMMPORTNumber = strCOMMPORTNumber
End Property

Public Property Get TimerInterval() As Integer
    TimerInterval = iTimerInterval
End Property
Public Property Get ScannerSendsCR() As Boolean
    ScannerSendsCR = bSendsCR
End Property
Public Property Get ExportToPastelEnabled() As Boolean
    ExportToPastelEnabled = bExportToPastelEnabled
End Property
Public Property Get ISSUE_PO_ON_THIS_WS() As Boolean
    ISSUE_PO_ON_THIS_WS = bISSUE_PO_ON_THIS_WS
End Property

Public Property Get IsFrontDeskWorkstation() As Boolean
    IsFrontDeskWorkstation = bIsFrontDeskWorkstation
End Property

Public Property Get POSActive() As Boolean
    POSActive = bPOSActive
End Property
Public Property Get IssueBookclubReturnDocs() As Boolean
    IssueBookclubReturnDocs = (strIssueBookclubReturnDocs = "TRUE")
End Property
Public Property Get EnableBookCLubReturn() As Boolean
    EnableBookCLubReturn = bEnableBookCLubReturn
End Property
Public Property Get ShowAllAppros() As Boolean
    ShowAllAppros = bShowAllAppros
End Property

'Public Property Get BF() As z_BookfindManager
'    Set BF = oBookfindmanager
'End Property
Public Property Get Configuration() As a_Configuration
    On Error GoTo errHandler
    Set Configuration = oConfig
    Exit Property
errHandler:
    ErrorIn "PapyConn.Configuration"
End Property
Private Sub objBookfind_RecordAdded(Success As Boolean)
    RaiseEvent RecordAdded(Success)
End Sub
Public Property Get DimensionUnits() As String
    DimensionUnits = strDimension
End Property
Public Property Get SpreadsheetType() As String
    SpreadsheetType = strSpreadsheetType
End Property
Public Property Get UsesExcel() As Boolean
    UsesExcel = (SpreadsheetType = "EXCEL")
End Property
Public Property Get UsesOpenOffice() As Boolean
    UsesOpenOffice = (SpreadsheetType = "OPENOFFICE")
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
    dteTimeStarted = Now()
    errSysHandlerSet
    Set oPC = Me
    bBFLoaded = False
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Class_Initialize"
End Sub
Public Property Get CO() As ADODB.Connection
    Set CO = cnPapy
End Property
Public Property Get COShort() As ADODB.Connection
    On Error GoTo errHandler
    Set COShort = cnPapyShort
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.COShort"
End Property

'Public Sub DisconnectFromBookfind()
'    On Error GoTo errHandler
'    If Not oBookFind Is Nothing Then
'        oBookFind.CloseDB
'        Set oBookFind = Nothing
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "PapyConn.DisconnectFromBookfind"
'End Sub
Private Sub NotifyConnection()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_NotifyConnection"
    cmd.commandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@COMPNAME", adVarChar, adParamInput, 30, Me.NameOfPC)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@WSID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.execute
    
    
    mlngWSID = cmd.Parameters("@WSID").Value
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.NotifyCOnnection"
End Sub
Public Property Get WorkstationID() As Long
    WorkstationID = mlngWSID
End Property
Public Property Get WorkstationName() As String
    WorkstationName = strPCName
End Property

'Public Function Connect() As Long
'    On Error GoTo ErrHandler
'Dim iConnect As Integer
'Dim i As Integer
'Dim strPos As String
'Dim fs As New FileSystemObject
'
'    RaiseEvent STATUS("Loading configuration data . . .")
'    Me.OpenDB
'    NotifyConnection
'    If Me.NameOfPC = PrintServerMachine Then
'        LoadPrintersToDB
'    End If
'    Set oConfig = New a_Configuration
'    ReloadConfiguration
'    i = 1
'    Do While oPC.Configuration.LookupSequence(i) > ""
'        Select Case oPC.Configuration.LookupSequence(i)
'        Case "BF"
'                If Not fs.FileExists(oPC.BookFindRoot & "\Endev32.dll") Then
'                    RaiseEvent STATUS("Bookfind not installed in folder expected by Papyrus")
'                    Exit Function
'                End If
'                Set oBookfindmanager = New z_BookfindManager
'                RaiseEvent STATUS("Linking to Bookfind . . .")
'                If Not oBookfindmanager.PrepareBookfind() Then
'                    iConnect = 2 'Connected to DB but not to Bookfind
'                End If
'                bBFLoaded = oBookfindmanager.BookfindConnected
'        Case "WH"
'        End Select
'        i = i + 1
'    Loop
'    RaiseEvent STATUS("")
'
'EXIT_Handler:
'    Connect = iConnect
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "PapyConn.Connect"
'End Function
Public Sub ReloadConfiguration()
Dim OpenResult As Integer

    '-------------------------------
        OpenResult = oPC.OpenDBSHort
    '-------------------------------

    Set oConfig = Nothing
    Set oConfig = New a_Configuration
    oConfig.Load
    LoadProperties
    RefreshBudgets
End Sub
Public Sub ReloadBudget()
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    '-------------------------------
        OpenResult = oPC.OpenDBSHort(True)
    '-------------------------------
  '  rsBudgets.Close
    Set rsBudgets = Nothing
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "Budget_Update_InDayend"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0

    Set prm = Nothing
    Set prm = cmd.CreateParameter("@DATE", adDate, adParamInput, , Date)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@DateBack", adInteger, adParamInput, , -1)
    cmd.Parameters.Append prm
    Set prm = Nothing
    cmd.execute
    
    RefreshBudgets
    
End Sub

Public Sub RefreshBudgets()
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    '-------------------------------
        OpenResult = oPC.OpenDBSHort(True)
    '-------------------------------
    Set rsBudgets = Nothing
    Set rsBudgets = New ADODB.Recordset
    rsBudgets.CursorLocation = adUseClient
    rsBudgets.Open "SELECT * FROM vBudget_Pivot1", oPC.COShort, adOpenKeyset, adLockOptimistic
    Set rsBudgets.ActiveConnection = Nothing

End Sub

Public Property Get rsBudgetData() As ADODB.Recordset
    Set rsBudgetData = rsBudgets
End Property
Public Sub OpenScheduler()
Dim cs As String

   cs = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & Me.servername & ";Initial Catalog=PapyrusScheduler;User Id=" & strUsername & ";Password=" & strPassword & ";Connect Timeout=50"
   Set cnScheduler = New ADODB.Connection
   cnScheduler.Open cs
   cnScheduler.CommandTimeout = 120
End Sub
Public Sub CloseScheduler()
On Error Resume Next
    cnScheduler.Close
   Set cnScheduler = Nothing
End Sub
Public Property Get SchedulerConnection() As ADODB.Connection
    Set SchedulerConnection = cnScheduler
End Property
Public Function OpenDBSHort(Optional restart As Boolean) As Integer
    On Error GoTo errHandler
Dim iCnt As Integer
Dim tmprs As ADODB.Recordset

    iCnt = 0
    OpenDBSHort = 0
    If mCL = "" Then
        mCL = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "DATABASENAME", "")
        If mCL = "" Then mCL = "PBKS"
    End If
                                                                                        'strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=" & Me.Servername & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPWD & "; Network Library=dbmssocn;Connect Timeout=45"
    If cnPapyShort Is Nothing Then
        Set cnPapyShort = New ADODB.Connection
    End If

    'Can't do the following if we have one procedure using OpenDBShort calling another also using OpenDBShort
  '  If cnPapyShort.State <> adStateClosed Then
  '      cnPapyShort.Close
  '  End If
TryAgain:
    If cnPapyShort.State = adStateClosed Or restart Then
        If Me.servername = "PBKS-SVR" Then
            strUsername = "sa"
        Else
            strUsername = "sa"
        End If
        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & Me.servername & ";Initial Catalog=" & mCL & ";User Id=" & strUsername & ";Password=" & strPassword & ";Connect Timeout=150"
      '        MsgBox ("strMainConnectionString = " & strMainConnectionString)

        If cnPapyShort.State = 1 Then cnPapyShort.Close
        cnPapyShort.Open strMainConnectionString
        cnPapyShort.CommandTimeout = 150

    Else
        Set tmprs = cnPapyShort.execute("SELECT CF_VATRATE FROM tConfiguration")
        tmprs.Close
        OpenDBSHort = 99
    End If


EXIT_Handler:
    Exit Function
errHandler:
    If Err = -2147467259 Then
        If cnPapyShort.State <> adStateClosed Then
            cnPapyShort.Close
            cnPapyShort.errors.Clear
            Err.Clear
            GoTo TryAgain
        End If
      '  Err.Clear
        If MsgBox("The application cannot connect to the database on the server. " & vbCrLf & "This difficulty may be caused by:  " & vbCrLf _
                & "1. The server computer not running, or " & vbCrLf _
                & "2. by the network router being off or " & vbCrLf _
                & "3. by a network fault or " & vbCrLf _
                & "4. by a firewall being altered on this computer or the server. " & vbCrLf _
                & "Check these and correct the situation before continuing, then click YES else click NO to close application. " & vbCrLf _
                & "Try again?", vbCritical + vbYesNo, "Connection difficulty") = vbYes Then
                DoEvents
                Err.Clear
            Resume
        Else
          '  Err.Clear
            OpenDBSHort = 1
            Err.Raise EXC_NOSERVER, "OpenDBShort", "Can't connect - stopping application."

            Exit Function
        End If
    End If

    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.OpenDBSHort", , , , strMainConnectionString, Array(strMainConnectionString)
End Function
Public Sub DisconnectDBShort()
    On Error GoTo errHandler
    cnPapyShort.Close
    Set cnPapyShort = Nothing
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 91 Then  'Already closed
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.DisconnectDBShort"
End Sub
Public Property Get DatabaseName() As String
    DatabaseName = mCL
End Property
Public Property Let DatabaseName(val As String)
    mCL = val
End Property
Public Function OpenDB() As Integer
    On Error GoTo errHandler
 
    If mCL = "" Then
        mCL = "PBKS13"
    End If
    OpenDB = 0
    If Me.servername = "PBKS-SVR" Then
        strUsername = "sa"
   '     strPassword = "jadarow"
    Else
        strUsername = "sa"
    End If
    If cnPapy Is Nothing Then
        Set cnPapy = New ADODB.Connection
        cnPapy.Provider = "sqloledb"
        cnPapy.CommandTimeout = 60
        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & Me.servername & ";Initial Catalog=" & mCL & ";User Id=" & strUsername & ";Password=" & strPassword & "; Connect Timeout=85"

        cnPapy.Open strMainConnectionString
        cnPapy.CommandTimeout = 0
    Else
    End If
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.OpenDB"
End Function

Public Property Get Connected() As Boolean
    Connected = flgDbObjConnected
End Property
Public Property Get BFLoaded() As Boolean
    BFLoaded = FNB(bBFLoaded)
End Property
Public Property Let UserName(val As String)
    strUsername = val
End Property
Public Property Get UserName() As String
    UserName = strUsername
End Property
Public Sub InitializeSettings(Optional bUseTest As Boolean)
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String

    Set oPC = Me
    strPCName = Trim(Me.NameOfPC)
  '  MsgBox App.Path
  '  MsgBox fs.GetDriveName(App.Path)
 '   strLocalRootFolder = fs.GetDriveName(App.Path) & "\PBKS"
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        
      '  MsgBox "strLocalRootFolder: " & strLocalRootFolder
        If bUseTest Then
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSERVERMACHINE", strPCName)
        Else
    '    MsgBox "POS Aa"
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
   '     MsgBox strPBKSSERVERMACHINE
        End If
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        If bUseTest Then
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSERVERMACHINE", strPCName)
        Else
   '     MsgBox "POS Ba"
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
    '    MsgBox strServername
        End If
        strSharedFolderRoot = "C:\PBKS"
    End If
    If bUseTest Then
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSQLSERVER", "")
    Else
  '      MsgBox "POS Ca"
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
  '      MsgBox strServername
   '     MsgBox "URL Problem"
        strServername = GetIniKeyValue("C:\PBKS\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
    End If
    
    
  '  MsgBox ("strsharedFolderRoot = " & strSharedFolderRoot)
    
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
   ' strPassword = GetIniKeyValue("C:\PBKS\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    '  MsgBox ("strPassword = " & strPassword)
  
    strPOSDBName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "POSSQLServer", "")
    strBackupMachine = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "BACKUPMACHINE", strPCName)
  '  bIsFrontDeskWorkstation = True '(Dir(strLocalRootFolder & "\DATA\PBKSFD*") > "")
    bISSUE_PO_ON_THIS_WS = UCase(GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "OPTIONS", "ISSUE_PO_ON_THIS_WS", "TRUE")) = "TRUE"
   ' MsgBox "Pos 1"
    LoadProperties
    If oPC.ShowBudget Then RefreshBudgets

Exit Sub
errHandler:
'MsgBox "error in InitializeSettings " & Error
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.InitializeSettings", , , , "settings", Array(strPCName, strLocalRootFolder, strServername, strPBKSSERVERMACHINE, strRemovableBackupShareName, bCompressBackup)
End Sub
Public Property Get SetSupplierIDFROMPO() As Boolean
    SetSupplierIDFROMPO = bSetSupplierIDFROMPO
End Property
Public Property Get SetSupplierIDFROMDEL() As Boolean
    SetSupplierIDFROMDEL = bSetSupplierIDFROMDEL
End Property
Public Property Get SetPricesInGRN() As Boolean
    SetPricesInGRN = bSetPricesInGRN
End Property
Public Property Get InternetDialup() As Boolean
    InternetDialup = (strInternetDialup = "YES")
End Property
Public Property Get Connectionname() As String
    Connectionname = strConnectionName
End Property

Public Property Get InvoiceTotalSequence() As String
    InvoiceTotalSequence = strInvoiceTotalSequence
End Property






Public Property Get servername() As String
    servername = strServername
End Property
Public Property Get Password() As String
    Password = strPassword
End Property

Public Property Get POSServername() As String
    POSServername = strPOSDBName
End Property
Public Property Get IsServerMachine() As Boolean
    IsServerMachine = (UCase(strPCName) = UCase(strPBKSSERVERMACHINE))
End Property
Public Property Get AllowInvoiceDateOverride() As Boolean
    AllowInvoiceDateOverride = bAllowInvoiceDateOverride
End Property
Public Property Get AllowPODateOverride() As Boolean
    AllowPODateOverride = bAllowPODateOverride
End Property
'Public Property Get BOOKFINDISBN13ENABLED() As Boolean
'    BOOKFINDISBN13ENABLED = bBOOKFINDISBN13ENABLED
'End Property
Public Property Get SupportsCatalogue() As Boolean
    SupportsCatalogue = bSupportsCatalogue
End Property
Public Property Get AllowGeneralStock() As Boolean
    AllowGeneralStock = bAllowGeneralStock
End Property
Public Property Get SupportsMultibuys() As Boolean
    SupportsMultibuys = bSupportsMultibuys
End Property
Public Property Get BlindCashup() As Boolean
    BlindCashup = bBlindCashup
End Property
Public Property Get AllowZeropricedPOLines() As Boolean
    AllowZeropricedPOLines = bAllowZeropricedPOLines
End Property
Public Property Get AllowZeropricedCOLines() As Boolean
    AllowZeropricedCOLines = bAllowZeropricedCOLines
End Property
Public Property Get AllowAntiquarionSearch() As Integer
    AllowAntiquarionSearch = iAllowAntiquarionSearch
End Property
Public Property Get UsesOutlookForPOEmail() As Boolean
    UsesOutlookForPOEmail = bUsesOutlookForPOEmail
End Property
Public Property Get UsesOutlookForINVEmail() As Boolean
    UsesOutlookForINVEmail = bUsesOutlookForINVEmail
End Property
Public Property Get UsesOutlookForCOEmail() As Boolean
    UsesOutlookForCOEmail = bUsesOutlookForCOEmail
End Property
Public Property Get UsesOutlookForQuoteEmail() As Boolean
    UsesOutlookForQuoteEmail = bUsesOutlookForQuoteEmail
End Property
Public Property Get EmailPO() As Boolean
    EmailPO = bEmailPO
End Property
Public Property Get EmailINV() As Boolean
    EmailINV = bEMailINV
End Property
Public Property Get EmailCO() As Boolean
    EmailCO = bEmailCO
End Property
Public Property Get EmailQuote() As Boolean
    EmailQuote = bEMailQuote
End Property

Public Property Get UniqueProducts() As Boolean
    UniqueProducts = bUniqueProducts
End Property




Public Property Get SupportsUNISA() As Boolean
    SupportsUNISA = bSupportsUNISA
End Property
Public Property Get NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
 '   MsgBox "NameofPC: " & NameOfPC
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.NameOfPC"
End Property
'
Public Property Get TimeStarted() As Date
    TimeStarted = Format(dteTimeStarted, "dd/mm/yyyy hh:mm")
End Property
Public Sub CloseDB()
    If cnPapy Is Nothing Then Exit Sub
    If cnPapy.State = 1 Then   'it is open
        cnPapy.Close
        Set cnPapy = Nothing
    End If
End Sub
Public Sub Disconnect()
    On Error GoTo errHandler
  '  DisconnectFromBookfind
    Set oConfig = Nothing
    If Not cnPapy Is Nothing Then
    If cnPapy.State = 1 Then   'it is open
        cnPapy.Close
        Set cnPapy = Nothing
    End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Disconnect"
End Sub

'Public Property Get BookFindFacet() As String
'    BookFindFacet = strBookfindFacet
'End Property
'
'Public Property Get BookFindRoot() As String
'    BookFindRoot = strBookfindRoot
'End Property
'Public Property Get PrintServerMachine() As String
'    PrintServerMachine = strPrintServerMachine
'End Property

Public Property Get RemovableBackupShareName() As String
    If InStr(1, strRemovableBackupShareName, "\") > 0 Then
        RemovableBackupShareName = strRemovableBackupShareName & IIf(Right(strRemovableBackupShareName, 1) = "\", "", "\")
   Else
        RemovableBackupShareName = "\\" & strBackupMachine & "\" & strRemovableBackupShareName & "\"
    End If
End Property
Public Property Get RemoveableBackupMedium() As String
    RemoveableBackupMedium = UCase(strBackupMedium)
End Property
Public Property Get CompressBackup() As Boolean
    CompressBackup = bCompressBackup
End Property
Public Property Get EDIEnabled() As Boolean
    EDIEnabled = bEDIEnabled
End Property
Public Property Let SharedFolderRoot(x As String)
    strSharedFolderRoot = x
End Property
Public Property Get SharedFolderRoot() As String
    SharedFolderRoot = strSharedFolderRoot  '  "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
End Property
Public Property Get DistributorDB() As String
    DistributorDB = strDDB
End Property
Public Property Get BackupFolder() As String
    BackupFolder = strSharedFolderRoot & "\BU\"
End Property
Public Property Get LocalFolder() As String
    LocalFolder = strLocalRootFolder & "\"
End Property
Public Property Get MarkCustomerOrderLinesOnPOLines() As String
    MarkCustomerOrderLinesOnPOLines = strMarkCustomerOrderLinesOnPOLines
End Property
Public Property Let CurrentSecurityCode(val As String)
    mCurrentSecurityCode = val
End Property
Public Property Get CurrentSecurityCode() As String
    CurrentSecurityCode = mCurrentSecurityCode
End Property
Public Property Get DefaultAreaCode() As String
    DefaultAreaCode = strDefaultAreaCode
End Property

Property Get MaxBrowseRecs() As Integer
    MaxBrowseRecs = MAXBROWSE
End Property
Private Sub LoadPrintersToDB()
    On Error GoTo errHandler
Dim pdf As XpdfPrint.XpdfPrint
Set pdf = New XpdfPrint.XpdfPrint
Dim nPrinters As Long
Dim i As Integer
Dim j As Integer
Dim PrinterName As String
Dim strPort As String

Dim p As Printer
        oPC.COShort.execute "UPDATE tPRINTERS SET PRINT_ACTIVE = 0"
'    For Each p In Printers
'        LoadPrinterFromString ParseDeviceName(p.DeviceName), p.Port, p.DeviceName = Printer.DeviceName
'    Next
  '  oPC.COShort.execute "DELETE FROM tPRINTERS WHERE ISNULL(PRINT_ACTIVE,0) <> 1"
  'Commented because this causes printers to be deleted when they are network printers and the workstation is turned off - ptoblem


    nPrinters = pdf.getNumPrinters
    For i = 0 To nPrinters - 1
        PrinterName = pdf.getPrinterName(i)
        strPort = ""
        For j = 0 To Printers.Count - 1
            If Printers(j).DeviceName = PrinterName Then
                strPort = Printers(j).Port
            End If
        Next j
        LoadPrinterFromString ParseDeviceName(PrinterName), strPort, PrinterName = Printer.DeviceName
    Next i



    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.LoadPrintersToDB"
End Sub
Public Sub LoadPrinterFromString(pPrinter As String, pPort As String, pDefault As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_LoadPrinter"
    cmd.commandType = adCmdStoredProc

    Set prm = cmd.CreateParameter("@Printer", adVarChar, adParamInput, 100, pPrinter)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@Port", adVarChar, adParamInput, 20, pPort)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@Default", adTinyInt, adParamInput, , IIf(pDefault, 1, 0))
    cmd.Parameters.Append prm

    cmd.execute
    Set cmd = Nothing
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadPrinterFromString(pPrinter)", pPrinter
End Sub

Public Sub WRiteToErrors(pString As String)
Dim oTF As New z_TextFile
    oTF.OpenTextFileToAppend oPC.SharedFolderRoot & "\DebugLog.TXT"
    oTF.WriteToTextFile "_____________"
    oTF.WriteToTextFile "FROM " & App.EXEName & " at " & Format(Now, "dd/mm/yyyy HH:NN AMPM")
    oTF.WriteToTextFile pString
    oTF.WriteToTextFile "============="
    oTF.CloseTextFile
End Sub

Private Function GetQuotation() As String



Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer

    If bUseQuotation = True Then
       ' If oPC.InternetDialup Then
            Set cmd = New ADODB.Command
    '-------------------------------
        OpenResult = oPC.OpenDBSHort
    '-------------------------------
            cmd.ActiveConnection = oPC.COShort
            cmd.CommandText = "sp_GET_A_QUOTE"
            cmd.commandType = adCmdStoredProc
            
            Set prm = cmd.CreateParameter("@QUOTATION", adVarChar, adParamOutput, 500)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            cmd.execute
    '-------------------------------------------------
        If OpenResult = 0 Then oPC.DisconnectDBShort
    '---------------------------------------------------
            If IsNull(cmd.Parameters(0)) = True Then
               GetQuotation = ""
            Else
               GetQuotation = Replace(Trim$(cmd.Parameters(0)), "-", " - ")
            End If
            Set cmd = Nothing
        'Else
        '    OpenBrowser "http://feeds.feedburner.com/brainyquote/QUOTENA"
        'End If
    Else
        GetQuotation = ""
    End If
    
    
End Function

Public Function LoadInitialData(Optional pQuiet As Boolean, Optional msg As String) As Boolean
    On Error GoTo errHandler
Dim i As Integer
Dim fs As New FileSystemObject

    LoadInitialData = True
    Set oPC = Me
'-------------------------------
    OpenDBSHort
'-------------------------------
    NotifyConnection
    If UCase(Me.NameOfPC) = UCase(oPC.GetProperty("PrintServerMachine")) Then
        LoadPrintersToDB
    End If
    Set oConfig = New a_Configuration
    
    ReloadConfiguration
'---------------------------------------------------
    DisconnectDBShort
'---------------------------------------------------
    i = 1

    Do While oPC.Configuration.LookupSequence(i) > ""
        Select Case oPC.Configuration.LookupSequence(i)
        Case "BF"
'                    If Me.GetProperty("NielsenUserID") = "" Or Me.GetProperty("NielsenPassword") = "" Then
'160                       msg = "You do not have the credentials for access to Bookfind set." & vbCrLf _
'                              & "You must contact your support person to correct this."
'170                       Exit Function
'180                   End If
'190                   Set oBookfindmanager = New z_BookfindManager
'200                   RaiseEvent STATUS("Linking to Bookfind . . .")
'210                   If Not oBookfindmanager.PrepareBookfind() Then
'220                       If Not pQuiet Then
'230                           If MsgBox("Connection to Bookfind failed" & vbCrLf & "Probably the disk is not in the drive or possibly it is incorrectly installed." & vbCrLf & "Choose OK to to continue without Bookfind and CANCEL to exit." & vbCrLf & "Status = " & oPC.Configuration.BookfindStatus, vbCritical + vbOKCancel, "Bookfind problem") = vbCancel Then
'240                               LoadInitialData = False 'to force closing of program
'250                               Set oPC = Nothing
'260                               Set oConfig = Nothing
'270                               GoTo EXIT_Handler
'280                           End If
'290                       Else
'300                       bBFLoaded = False
'310                       End If
'320                   End If
                bBFLoaded = True
        Case "WH"
        End Select
        i = i + 1
    Loop
    
    If oPC.Configuration.DefaultPT = 0 Then MsgBox "WARNING, no product type has been set as the default, you connot add product records until this is set. " & vbCrLf & "Use the Product types option under Master files on the menu to select a product type as the default.", vbExclamation, "WARNING"
    
EXIT_Handler:
    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.LoadInitialData(pQuiet,msg)", Array(pQuiet, msg)
End Function

Public Function GetProperty(pKey As String) As String
    On Error GoTo errHandler
    rsProperty.MoveFirst
    rsProperty.Find "PropertyKey = '" & pKey & "'"
    On Error Resume Next
    GetProperty = FNS(CStr(rsProperty.Fields(1)))
    If Err Then
        GetProperty = ""
        Err.Clear
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.GetProperty(pKey)", pKey
End Function
Public Function LoadProperties() As Boolean
    On Error GoTo errHandler
Dim sSQL As String
Dim fs As New FileSystemObject
Dim sec1 As String
Dim sec2 As String
Dim s1 As String
Dim s2 As String
Dim s3 As String
Dim sar() As String
 Dim OpenResult As Integer
    '   MsgBox "Load properties 1"
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
   '    MsgBox "Load properties 2"
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.Open sSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    Set rsProperty.ActiveConnection = Nothing
'-------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    strRemovableBackupShareName = GetProperty("BACKUP")  '  GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FOLDERS", "BACKUP", "REM")
    bCompressBackup = GetProperty("BACKUPCOMPRESSION") = "TRUE"  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "BACKUPCOMPRESSION ", "TRUE") = "TRUE"
    strBackupMedium = GetProperty("BACKUPMEDIUM")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "BACKUPMEDIUM ", "DISK")
    strCDType = GetProperty("CDTYPE")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "CDTYPE ", "RO")
    strLabelPrinter = GetProperty("LABELPRINTER")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "LABELPRINTER ", "OKI")
'     strBookfindRoot = GetProperty("BOOKFINDROOT")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "BOOKDATA", "BOOKFINDROOT", "c:\Bookfind")
  '  strBookfindFacet = GetProperty("BOOKFINDFACET")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "BOOKDATA", "BOOKFINDFACET", "WEBK")
    MAXBROWSE = CLng(GetProperty("MAXBROWSE"))  '   = CInt(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "MAXBROWSE", "1000"))
    bEDIEnabled = GetProperty("EDIENABLED") = "TRUE" '   = IIf(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PRINTING", "EDIENABLED", "FALSE") = "TRUE", True, False)
    strDefaultAreaCode = GetProperty("DEFAULTAREACODE")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "LOCAL", "DEFAULTAREACODE", "")
    strInternetDialup = GetProperty("INTERNETDIALUP")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP ", "TRUE")
    strConnectionName = GetProperty("CONNECTIONNAME")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME ", "")
    iTimerInterval = CInt(GetProperty("TIMERINTERVAL"))  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "TIMERINTERVAL", "3000")
    bSendsCR = GetProperty("SendsCR") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "SENDSCR", False)
    strCOMPORTSettings = GetProperty("COMPORTSETTINGS")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "COMPORTSETTINGS", "9600,n,8,1")
    strCOMMPORTNumber = GetProperty("COMPORTNumber")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "COMPORTNumber", "1")
    strSpreadsheetType = GetProperty("Spreadsheet_type")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "COMPORTNumber", "1")
    bPOSActive = GetProperty("POSACTIVE") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "POS", "POSACTIVE", "FALSE") = "TRUE"
    bShowBudget = GetProperty("ShowBudget") = "TRUE"
      s1 = GetProperty("SUPPLERINVOICETOLERANCE")
      If IsNumeric(s1) Then
          dblSupplierInvoiceTolerance = CDbl(s1)
      Else
          dblSupplierInvoiceTolerance = 0
      End If
    iRoundPriceTo = FNN(GetProperty("ROUNDPRICETO"))  '   = CDbl(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "ROUNDPRICETO", "0"))
    strVouchersReportTogether = GetProperty("VOUCHEPORTTOGETHER")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "VOUCHEPORTTOGETHER", "")
    bEnableBookCLubReturn = GetProperty("ENABLEBOOKCLUBRETURN") = "TRUE" '   = IIf(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "ENABLEBOOKCLUBRETURN", "FALSE") = "TRUE", True, False)
   ' strPrintServerMachine = GetProperty("PRINTSERVERMACHINE")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "NETWORK", "PRINTSERVERMACHINE", strPBKSSERVERMACHINE)
    strIssueBookclubReturnDocs = GetProperty("ISSUEBOOKCLUBRETURNDOCS")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "POS", "ISSUEBOOKCLUBRETURNDOCS", "FALSE")
    iAllowAntiquarionSearch = CInt(GetProperty("ALLOWANTIQUARIANSEARCH"))  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "ALLOWANTIQUARIANSEARCH", "1")
    bUseQuotation = GetProperty("SHOWQUOTES") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SHOWQUOTES", "TRUE") = "TRUE"
    bShowAllAppros = GetProperty("SHOWALLAPPROS") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SHOWALLAPPROS", "TRUE") = "TRUE"
    bExportToPastelEnabled = GetProperty("EXPORTTOPASTELENABLED") = "TRUE"   '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PASTEL", "EXPORTTOPASTELENABLED", "FALSE") = "TRUE"
    strContraAccount_Inv = GetProperty("CONTRA_ACCOUNT_INV")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PASTEL", "CONTRA_ACCOUNT_INV", "")
    strContraAccount_SInv = GetProperty("CONTRA_ACCOUNT_SINV")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PASTEL", "CONTRA_ACCOUNT_SINV", "")
    strInvoiceTotalSequence = GetProperty("INVTOTALSEQ")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "INVTOTALSEQ", "EVR")
    bHideLocalSKUonINV = GetProperty("HIDELOCALSKUONINV") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "HIDELOCALSKUONINV", "FALSE") = "TRUE"
    bSetSupplierIDFROMPO = GetProperty("SetSupplierIDFROMPO") = "TRUE"  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SetSupplierIDFROMPO", "FALSE") = "TRUE"
    bSetSupplierIDFROMDEL = GetProperty("SetSupplierIDFROMDEL") = "TRUE"  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SetSupplierIDFROMPO", "FALSE") = "TRUE"
    bAllowInvoiceDateOverride = GetProperty("AllowInvoiceDateOverride") = "TRUE"
    bAllowPODateOverride = GetProperty("AllowPODateOverride") = "TRUE"
'   bBOOKFINDISBN13ENABLED = GetProperty("BOOKFINDISBN13ENABLED") = "TRUE"
    bSupplierBasedCurrencyConversion = GetProperty("SupplierBasedCurrencyConversion") = "TRUE"
    bMultiStore = GetProperty("MultiStore") = "TRUE"
    bShowWordstockSales = GetProperty("ShowWordstockSales") = "TRUE"
    bAllowGeneralStock = GetProperty("AllowGeneralStock") = "TRUE"
    bBlindCashup = GetProperty("BlindCashup") = "TRUE"
    bIsFrontDeskWorkstation = False '(Dir(strLocalRootFolder & "\DATA\PBKSFD*") > "")
    strEmailForDBWarnings = GetProperty("EmailAddressForDatabaseIntegrityWarnings")
    bIncludeSupplierFeatures = GetProperty("IncludeSupplierOptions") = "TRUE"
    sec1 = GetProperty("SecondaryEDIAddress")
    sec2 = IIf(GetProperty("UNISASupport") = "", "999", GetProperty("UNISASupport"))
    If sec1 > "" Then
        sar = Split(sec1, ".")
        s1 = Right(sar(1), 2)
        s2 = Right(sar(2), 2)
        s3 = Right(sar(3), 2)
    End If
    
    bSupportsUNISA = False
    If s1 > "" And s2 > "" And s3 > "" Then
        If sec2 = Chr(s1) & Chr(s2) & Chr(s3) Then
            bSupportsUNISA = True
        End If
    End If
    
    bEnforceSupplierOnProduct = GetProperty("EnforceSupplierOnProduct") = "TRUE"
    bSupportsCatalogue = GetProperty("SupportsCatalogue") = "TRUE"
    iKeepTemporaryfilesFor_n_Days = FNN(GetProperty("KeepTemporaryfilesFor_n_Days"))
    bAllowZeropricedPOLines = GetProperty("AllowZeropricedPOLines") = "TRUE"
    bAllowZeropricedCOLines = GetProperty("AllowZeropricedCOLines") = "TRUE"
    bSupportsLoyaltyCustomersTF = GetProperty("SupportsLoyaltyCustomers") = "TRUE"
    bSupportsBookClubsTF = GetProperty("SupportsBookClubs") = "TRUE"
    bTransferIsExVAT = GetProperty("TransferIsEXVAT") = "TRUE"
    bAllowsSSInvoicing = GetProperty("AllowsSSInvoicing") = "TRUE"
    bSetPricesInGRN = GetProperty("SetPricesInGRN") = "TRUE"
    bAllowsInvoicePicking = GetProperty("AllowsInvoicePicking") = "TRUE"
    bRunsAccountsTF = GetProperty("RunsAccountsTF") = "TRUE"
    bUsesOutlookForPOEmail = GetProperty("UsesOutlookForPOEmail") = "TRUE"
    bUsesOutlookForCOEmail = GetProperty("UsesOutlookForCOEmail") = "TRUE"
    bUsesOutlookForINVEmail = GetProperty("UsesOutlookForINVEmail") = "TRUE"
    bUsesOutlookForQuoteEmail = GetProperty("UsesOutlookForQuoteEmail") = "TRUE"
    bAllowsZeroDiscountPOs = GetProperty("AllowsZeroDiscountPOs") = "TRUE"
    bUpdatePriceOnForeignDelivery = GetProperty("UpdatePriceOnForeignDelivery") = "TRUE"
    bAllowGeneralStock = GetProperty("AllowGeneralStock") = "TRUE"
    bSupportsMultibuys = GetProperty("SUPPORTSMULTIBUY") = "TRUE"
    strMarkCustomerOrderLinesOnPOLines = GetProperty("MarkCustomerOrderLinesOnPOLines")
    bEmailPO = GetProperty("EmailPO") = "TRUE"
    bEmailCO = GetProperty("EmailSalesOrder") = "TRUE"
    bEMailINV = GetProperty("EMailINV") = "TRUE"
    bEMailQuote = GetProperty("EMailQuote") = "TRUE"
    bEmailPOShowHTML = GetProperty("Email_PO_ShowHTML") = "TRUE"
    bEmailSalesOrderShowHTML = GetProperty("Email_SalesOrder_ShowHTML") = "TRUE"
    bEmailShowHTML = GetProperty("Email_INV_ShowHTML") = "TRUE"
    bEmailQuoteShowHTML = GetProperty("Email_Quote_ShowHTML") = "TRUE"
    strDeliveryStyle = GetProperty("DeliveryStyle")
    bCanGenerateApprosFromOrderFulfilment = GetProperty("CanGenerateApprosFromOrderFulfilment") = "TRUE"
    bUniqueProducts = GetProperty("UniqueProducts") = "TRUE"
   ' bAllowGDNs = GetProperty("AllowsGDNs") = "TRUE"
    
    strEmailFrom = GetProperty("EmailFrom")
    strEmailAddressForTesting = GetProperty("EmailAddressForTesting")
    strPOSubject = GetProperty("Subject")
    strInvoiceSubject = GetProperty("InvoiceSubject")
    strSenderName = GetProperty("SenderName")
    strDimension = GetProperty("DimensionMeasurementUnits")
      strCodePrefixForInternalBarcodesEAN = GetProperty("CodePrefixForInternalBarcodesEAN")

    mSecurity = GetProperty("SecuritySetting")
    bTestMode = GetProperty("TestMode") = "TRUE"
    bPrintPackingSlip = GetProperty("PrintPackingSlip") = "TRUE"
    lngTranslationOffset = CLng(IIf(GetProperty("Translation_offset") = "", 0, GetProperty("Translation_offset")))
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        Set rsLinesClipboard = New ADODB.Recordset
        rsLinesClipboard.Open oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        PrepareLinesClipboard
    End If
    Exit Function
errHandler:
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.LoadProperties"
End Function

Private Sub InsertProperty(pPropertyName As String, pVal As String)
    oPC.CO.execute "INSERT INTO tPROPERTY (PropertyKey,PropertyValue) VALUES ('" & pPropertyName & "','" & pVal & "')"
End Sub

Public Property Get LinesClipboard() As ADODB.Recordset
    Set LinesClipboard = rsLinesClipboard
End Property
Public Property Get QtyLinesinClipboard() As Long
    If rsLinesClipboard Is Nothing Then
        QtyLinesinClipboard = 0
        Exit Property
    End If
    If rsLinesClipboard.State = 0 Then
        QtyLinesinClipboard = 0
    Else
        QtyLinesinClipboard = rsLinesClipboard.RecordCount
    End If
End Property
Public Sub PrepareLinesClipboard()
Dim f As ADODB.Field

    If rsLinesClipboard.State <> 0 Then rsLinesClipboard.Close
    Set rsLinesClipboard = Nothing
    Set rsLinesClipboard = New ADODB.Recordset
    rsLinesClipboard.CursorLocation = adUseClient
    With rsLinesClipboard
        .Fields.Append "GUID", adVarChar, 40
        .Fields.Append "PID", adVarChar, 40
        .Fields.Append "REF", adVarChar, 30
        .Fields.Append "QTY", adInteger
        .Fields.Append "QTYFIRM", adInteger
        .Fields.Append "QTYSS", adInteger
        .Fields.Append "PRICE", adInteger
        .Fields.Append "DISCOUNTRATE", adNumeric
        Set f = .Fields("DISCOUNTRATE")
        f.NumericScale = 2
        f.Precision = 15
        .Fields.Append "CODEF", adVarChar, 20
        .Fields.Append "EANF", adVarChar, 20
        .Fields.Append "EAN", adVarChar, 20
        .Fields.Append "TITLE", adVarChar, 1000
        .Fields.Append "VATRATE", adNumeric
        .Fields.Append "EXTRACHARGEPID", adVarChar, 40
        .Fields.Append "EXTRACHARGEVALUE", adInteger
        Set f = .Fields("VATRATE")
        f.NumericScale = 2
        f.Precision = 15
        .Fields.Append "SELECTED", adBoolean
        .Fields.Append "ETA", adDate
        .Fields.Append "FCPRICE", adInteger
        .Fields.Append "FCFACTOR", adNumeric
        Set f = .Fields("FCFACTOR")
        f.NumericScale = 2
        f.Precision = 15
        .Fields.Append "FCID", adInteger
        
   End With
    
End Sub

Public Sub SwapConnectionToDatabase()

    If Me.DatabaseName = "PBKS" Then
        Me.DatabaseName = "PBKS_TEST"
    Else
        Me.DatabaseName = "PBKS"
    End If
    oPC.Disconnect
    LoadProperties
    LoadInitialData
    
End Sub

Public Property Get IsMultiStore() As Boolean
    IsMultiStore = bMultiStore
End Property

Public Function OpenSuppliedConnection(c As ADODB.Connection)
Dim CC As ADODB.Connection

    Set CC = New ADODB.Connection
    CC.Open strMainConnectionString
    CC.CommandTimeout = 0
    Set c = CC

End Function
Public Function CloseSUppliedConnection(c As ADODB.Connection)
    c.Close
End Function

