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
Private flgDbObjConnected As Boolean
'Private cnPapy As ADODB.Connection
Private cnPapyShort As ADODB.Connection
Dim rsProperty As New ADODB.Recordset
Dim strSpreadsheetType As String
Dim mCASHSALESTPID As Long
Dim mDefaultStoreID As Long
Dim strPBKSSERVERMACHINE As String
Dim strSharedFolderRoot As String
Dim strDBName As String
Dim strMainConnectionString As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Property Get LocalMachineName() As String
    LocalMachineName = sLocalMachineName
End Property
Public Sub DBName(val As String)
    strDBName = val
End Sub
Public Property Get DatabaseName() As String
    DatabaseName = strDBName
End Property
Public Function NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Function
Public Property Get COShort() As ADODB.Connection
    On Error GoTo errHandler
    Set COShort = cnPapyShort
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.COShort"
End Property
Public Function Connect(Optional strProvider As String) As Long
    On Error GoTo errHandler
Dim iConnect As Integer
    Set cnPapyShort = Nothing
    
    iConnect = OpenDBSHort
    DisconnectDBShort
    
    Set oPC = Me
    
EXIT_Handler:
    Connect = iConnect
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Connect(strProvider)", strProvider
End Function

Public Sub DisconnectDBShort()
On Error Resume Next
    If Not cnPapyShort Is Nothing Then
        If Not cnPapyShort.State = adStateClosed Then
            cnPapyShort.Close
        End If
        Set cnPapyShort = Nothing
    End If
    Exit Sub

End Sub
Public Sub Initialize()
10        On Error GoTo errHandler
      Dim rs As ADODB.Recordset
      Dim fs As New FileSystemObject
      Dim iRetryCount As Integer
      Dim cmd1 As New ADODB.Command

20   ' Err.Raise 123
30        iRetryCount = 0
          
40        If Not fs.FolderExists(oPC.SharedFolderRoot & "\Exchanges") Then
50            fs.CreateFolder oPC.SharedFolderRoot & "\Exchanges"
60        End If

70        Set rs = New ADODB.Recordset

      'using a command so we can set commandtimeout to 0. Occasionally getting errors on starting app with a time out here
80        cmd1.ActiveConnection = cnPapyShort
90        cmd1.CommandTimeout = 0
100       cmd1.CommandType = adCmdText
110       cmd1.CommandText = "SELECT CF_CSCUSTOMERID,CF_DEFAULTSTOREID FROM tCONFIGURATION"
          
120       rs.Open cmd1, , adOpenKeyset, adLockOptimistic
              
              
              
          '    rs.Open "SELECT CF_CSCUSTOMERID,CF_DEFAULTSTOREID FROM tCONFIGURATION", cnPapyShort, adOpenKeyset, adLockOptimistic
130           mCASHSALESTPID = FNN(rs!CF_CSCUSTOMERID)
140           mDefaultStoreID = rs!CF_DEFAULTSTOREID
150           rs.Close
160           Set rs = Nothing
170         LoadProperties
180       Exit Sub
errHandler:
190       ErrPreserve
200       If Err = -2147467259 And iRetryCount < 5 Then
210           iRetryCount = iRetryCount + 1
220           MsgWaitObj (2000)
230           Resume
240       End If
250       ErrorIn "PapyConn.Initialize", , , , "line number", Array(Erl())
End Sub
Public Property Get ConnectionString() As String
        ConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Data Source=" & Me.servername & ";Initial Catalog=" & strDBName & ";User Id=sa;Password=" & Me.Password & "; Connect Timeout=180"
End Property

Public Function OpenDBSHort() As Integer
    On Error GoTo errHandler
    Dim retry As Integer
    retry = 0
TOP:
    OpenDBSHort = 0
    If strDBName = "" Then
        strDBName = "PBKS"
    End If
    
    If cnPapyShort Is Nothing Then
        Set cnPapyShort = New ADODB.Connection
    End If
    If cnPapyShort.Errors.Count > 0 Or cnPapyShort.State = adStateClosed Then
        If Not cnPapyShort.State = adStateClosed Then
            cnPapyShort.Close
        End If
        Set cnPapyShort = Nothing
        Set cnPapyShort = New ADODB.Connection
        strMainConnectionString = "Provider=SQLOLEDB.1;User ID=sa;Data Source=" & Me.servername & ";Initial Catalog=" & strDBName & ";User Id=sa;Password=" & Me.Password & "; Connect Timeout=180"
        cnPapyShort.Open strMainConnectionString
    Else
        OpenDBSHort = 99
    End If

EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    If retry < 300 Then
        retry = retry + 1
        MsgWaitObj 100
        Err.Clear
        Set cnPapyShort = Nothing
        Resume TOP
    Else
        OpenDBSHort = 2
        Resume EXIT_Handler
    End If
        
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.OpenDBSHort", , , , "strMainConnectionString,retry", Array(strMainConnectionString, retry)
End Function

Public Property Get CustomerTPID() As Long
    CustomerTPID = mCASHSALESTPID
End Property
Public Property Get DefaultStoreID() As Long
    DefaultStoreID = mDefaultStoreID
End Property
Private Sub InitializeSettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String

    sLocalMachineName = NameOfPC

    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & NameOfPC & "\PBKS_S"
        strMainSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strMainSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
        strSharedFolderRoot = "C:\PBKS"
    End If
    
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.InitializeSettings"
End Sub
Public Property Get SharedFolderRoot() As String
    SharedFolderRoot = strSharedFolderRoot
End Property

Public Property Get servername() As String
    servername = strMainSQLServerName
End Property
Public Property Let servername(val As String)
    strMainSQLServerName = val
End Property
Public Property Get Password() As String
    Password = strPassword
End Property
Public Property Get LocalRoot() As String
    LocalRoot = strLocalRootFolder
End Property
Private Sub Class_Initialize()
    InitializeSettings

End Sub
Public Property Get TransactionCount() As Long
Dim rs As ADODB.Recordset

        Set rs = cnPapyShort.Execute("SELECT @@TRANCOUNT")
        TransactionCount = CLng(rs.Fields(0))
        
End Property

Public Sub BeginTransaction()
Dim rs As ADODB.Recordset

        cnPapyShort.Execute "BEGIN TRANSACTION"
        cnPapyShort.Execute "set implicit_transactions off"
        
End Sub
Public Sub CommitTransaction()
Dim rs As ADODB.Recordset

        cnPapyShort.Execute "COMMIT TRANSACTION"
        
End Sub

Public Sub RollbackTransaction()
Dim rs As ADODB.Recordset

        cnPapyShort.Execute "ROLLBACK TRANSACTION"
        
End Sub

Public Function LoadProperties() As Boolean
10        On Error GoTo errHandler
      Dim sSQL As String
      Dim fs As New FileSystemObject
      Dim sec1 As String
      Dim sec2 As String
      Dim s1 As String
      Dim s2 As String
      Dim s3 As String
      Dim sar() As String
       Dim OpenResult As Integer
      '-------------------------------
20        OpenResult = oPC.OpenDBSHort
      '-------------------------------
         
30        sSQL = "SELECT * FROM tProperty"
40        Set rsProperty = New ADODB.Recordset
50        rsProperty.CursorLocation = adUseClient
60        rsProperty.Open sSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
70        Set rsProperty.ActiveConnection = Nothing
      '-------------------------------------------------
80        If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
'90        strRemovableBackupShareName = GetProperty("BACKUP")  '  GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FOLDERS", "BACKUP", "REM")
'100       bCompressBackup = GetProperty("BACKUPCOMPRESSION") = "TRUE"  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "BACKUPCOMPRESSION ", "TRUE") = "TRUE"
'110       strBackupMedium = GetProperty("BACKUPMEDIUM")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "BACKUPMEDIUM ", "DISK")
'120       strCDType = GetProperty("CDTYPE")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "CDTYPE ", "RO")
'130       strLabelPrinter = GetProperty("LABELPRINTER")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "LABELPRINTER ", "OKI")
'140       strBookfindRoot = GetProperty("BOOKFINDROOT")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "BOOKDATA", "BOOKFINDROOT", "c:\Bookfind")
'150       strBookfindFacet = GetProperty("BOOKFINDFACET")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "BOOKDATA", "BOOKFINDFACET", "WEBK")
'160       MAXBROWSE = CInt(GetProperty("MAXBROWSE"))  '   = CInt(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "MAXBROWSE", "1000"))
'170       bEDIEnabled = GetProperty("EDIENABLED") = "TRUE" '   = IIf(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PRINTING", "EDIENABLED", "FALSE") = "TRUE", True, False)
'180       strDefaultAreaCode = GetProperty("DEFAULTAREACODE")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "LOCAL", "DEFAULTAREACODE", "")
'190       strInternetDialup = GetProperty("INTERNETDIALUP")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP ", "TRUE")
'200       strConnectionName = GetProperty("CONNECTIONNAME")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME ", "")
'210       iTimerInterval = CInt(GetProperty("TIMERINTERVAL"))  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "TIMERINTERVAL", "3000")
'220       bSendsCR = GetProperty("SendsCR") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "SENDSCR", False)
'230       strCOMPORTSettings = GetProperty("COMPORTSETTINGS")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "COMPORTSETTINGS", "9600,n,8,1")
'240       strCOMMPORTNumber = GetProperty("COMPORTNumber")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "COMPORTNumber", "1")
250       strSpreadsheetType = GetProperty("Spreadsheet_type")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FRONTDESK", "COMPORTNumber", "1")
'260       bPOSActive = GetProperty("POSACTIVE") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "POS", "POSACTIVE", "FALSE") = "TRUE"
'270       bShowBudget = GetProperty("ShowBudget") = "TRUE"
'280       dblSupplierInvoiceTolerance = CDbl(GetProperty("SUPPLERINVOICETOLERANCE"))  '   = CDbl(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SUPPLERINVOICETOLERANCE", ".005"))
'290       iRoundPriceTo = FNN(GetProperty("ROUNDPRICETO"))  '   = CDbl(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "ROUNDPRICETO", "0"))
'300       strVouchersReportTogether = GetProperty("VOUCHEPORTTOGETHER")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "VOUCHEPORTTOGETHER", "")
'310       bEnableBookCLubReturn = GetProperty("ENABLEBOOKCLUBRETURN") = "TRUE" '   = IIf(GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "ENABLEBOOKCLUBRETURN", "FALSE") = "TRUE", True, False)
'         ' strPrintServerMachine = GetProperty("PRINTSERVERMACHINE")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "NETWORK", "PRINTSERVERMACHINE", strPBKSSERVERMACHINE)
'320       strIssueBookclubReturnDocs = GetProperty("ISSUEBOOKCLUBRETURNDOCS")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "POS", "ISSUEBOOKCLUBRETURNDOCS", "FALSE")
'330       iAllowAntiquarionSearch = CInt(GetProperty("ALLOWANTIQUARIANSEARCH"))  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "ALLOWANTIQUARIANSEARCH", "1")
'340       bUseQuotation = GetProperty("SHOWQUOTES") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SHOWQUOTES", "TRUE") = "TRUE"
'350       bShowAllAppros = GetProperty("SHOWALLAPPROS") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SHOWALLAPPROS", "TRUE") = "TRUE"
'360       bExportToPastelEnabled = GetProperty("EXPORTTOPASTELENABLED") = "TRUE"   '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PASTEL", "EXPORTTOPASTELENABLED", "FALSE") = "TRUE"
'370       strContraAccount_Inv = GetProperty("CONTRA_ACCOUNT_INV")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PASTEL", "CONTRA_ACCOUNT_INV", "")
'380       strContraAccount_SInv = GetProperty("CONTRA_ACCOUNT_SINV")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "PASTEL", "CONTRA_ACCOUNT_SINV", "")
'390       strInvoiceTotalSequence = GetProperty("INVTOTALSEQ")  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "INVTOTALSEQ", "EVR")
'400       bHideLocalSKUonINV = GetProperty("HIDELOCALSKUONINV") = "TRUE" '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "HIDELOCALSKUONINV", "FALSE") = "TRUE"
'410       bSetSupplierIDFROMPO = GetProperty("SetSupplierIDFROMPO") = "TRUE"  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SetSupplierIDFROMPO", "FALSE") = "TRUE"
'420       bSetSupplierIDFROMDEL = GetProperty("SetSupplierIDFROMDEL") = "TRUE"  '   = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "SetSupplierIDFROMPO", "FALSE") = "TRUE"
'430       bAllowInvoiceDateOverride = GetProperty("AllowInvoiceDateOverride") = "TRUE"
'440       bAllowPODateOverride = GetProperty("AllowPODateOverride") = "TRUE"
'450       bBOOKFINDISBN13ENABLED = GetProperty("BOOKFINDISBN13ENABLED") = "TRUE"
'460       bSupplierBasedCurrencyConversion = GetProperty("SupplierBasedCurrencyConversion") = "TRUE"
'470       bMultiStore = GetProperty("MultiStore") = "TRUE"
'480       bShowWordstockSales = GetProperty("ShowWordstockSales") = "TRUE"
'490       bAllowGeneralStock = GetProperty("AllowGeneralStock") = "TRUE"
'500       bBlindCashup = GetProperty("BlindCashup") = "TRUE"
'510       bIsFrontDeskWorkstation = False '(Dir(strLocalRootFolder & "\DATA\PBKSFD*") > "")
'520       strEmailForDBWarnings = GetProperty("EmailAddressForDatabaseIntegrityWarnings")
'          bIncludeSupplierFeatures = GetProperty("IncludeSupplierOptions") = "TRUE"
'530       sec1 = GetProperty("SecondaryEDIAddress")
'540       sec2 = IIf(GetProperty("UNISASupport") = "", "999", GetProperty("UNISASupport"))
'550       If sec1 > "" Then
'560           sar = Split(sec1, ".")
'570           s1 = Right(sar(1), 2)
'580           s2 = Right(sar(2), 2)
'590           s3 = Right(sar(3), 2)
'600       End If
'
'610       bSupportsUNISA = False
'620       If s1 > "" And s2 > "" And s3 > "" Then
'630           If sec2 = Chr(s1) & Chr(s2) & Chr(s3) Then
'640               bSupportsUNISA = True
'650           End If
'660       End If
'
'670       bSupportsCatalogue = GetProperty("SupportsCatalogue") = "TRUE"
'680       iKeepTemporaryfilesFor_n_Days = FNN(GetProperty("KeepTemporaryfilesFor_n_Days"))
'690       bAllowZeropricedPOLines = GetProperty("AllowZeropricedPOLines") = "TRUE"
'700       bAllowZeropricedCOLines = GetProperty("AllowZeropricedCOLines") = "TRUE"
'710       bSupportsLoyaltyCustomersTF = GetProperty("SupportsLoyaltyCustomers") = "TRUE"
'720       bSupportsBookClubsTF = GetProperty("SupportsBookClubs") = "TRUE"
'730       bTransferIsExVAT = GetProperty("TransferIsEXVAT") = "TRUE"
'740       bAllowsSSInvoicing = GetProperty("AllowsSSInvoicing") = "TRUE"
'750       bSetPricesInGRN = GetProperty("SetPricesInGRN") = "TRUE"
'760       bAllowsInvoicePicking = GetProperty("AllowsInvoicePicking") = "TRUE"
'770       bRunsAccountsTF = GetProperty("RunsAccountsTF") = "TRUE"
'780       bUsesOutlookForPOEmail = GetProperty("UsesOutlookForPOEmail") = "TRUE"
'790       bUsesOutlookForCOEmail = GetProperty("UsesOutlookForCOEmail") = "TRUE"
'800       bUsesOutlookForINVEmail = GetProperty("UsesOutlookForINVEmail") = "TRUE"
'810       bUsesOutlookForQuoteEmail = GetProperty("UsesOutlookForQuoteEmail") = "TRUE"
'820       bAllowsZeroDiscountPOs = GetProperty("AllowsZeroDiscountPOs") = "TRUE"
'830       bUpdatePriceOnForeignDelivery = GetProperty("UpdatePriceOnForeignDelivery") = "TRUE"
'840       bAllowGeneralStock = GetProperty("AllowGeneralStock") = "TRUE"
'850       bSupportsMultibuys = GetProperty("SUPPORTSMULTIBUY") = "TRUE"
'860       strMarkCustomerOrderLinesOnPOLines = GetProperty("MarkCustomerOrderLinesOnPOLines")
'870       bEmailPO = GetProperty("EmailPO") = "TRUE"
'880       bEmailCO = GetProperty("EmailSalesOrder") = "TRUE"
'890       bEMailINV = GetProperty("EMailINV") = "TRUE"
'900       bEMailQuote = GetProperty("EMailQuote") = "TRUE"
'910       bEmailPOShowHTML = GetProperty("Email_PO_ShowHTML") = "TRUE"
'920       bEmailSalesOrderShowHTML = GetProperty("Email_SalesOrder_ShowHTML") = "TRUE"
'930       bEmailShowHTML = GetProperty("Email_INV_ShowHTML") = "TRUE"
'940       bEmailQuoteShowHTML = GetProperty("Email_Quote_ShowHTML") = "TRUE"
'950       strDeliveryStyle = GetProperty("DeliveryStyle")
'955       bCanGenerateApprosFromOrderFulfilment = GetProperty("CanGenerateApprosFromOrderFulfilment") = "TRUE"
'960       bUniqueProducts = GetProperty("UniqueProducts") = "TRUE"
'965      ' bAllowGDNs = GetProperty("AllowsGDNs") = "TRUE"
'
'970       strEmailFrom = GetProperty("EmailFrom")
'980       strEmailAddressForTesting = GetProperty("EmailAddressForTesting")
'990       strPOSubject = GetProperty("Subject")
'1000      strInvoiceSubject = GetProperty("InvoiceSubject")
'1010      strSenderName = GetProperty("SenderName")
'1020      strDimension = GetProperty("DimensionMeasurementUnits")
'
'1030      bTestMode = GetProperty("TestMode") = "TRUE"
'1040      bPrintPackingSlip = GetProperty("PrintPackingSlip") = "TRUE"
'1050      lngTranslationOffset = CLng(IIf(GetProperty("Translation_offset") = "", 0, GetProperty("Translation_offset")))
'1060      If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
'1070          Set rsLinesClipboard = New ADODB.Recordset
'1080          rsLinesClipboard.Open oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
'1090      Else
'1100          PrepareLinesClipboard
'1110      End If
1120      Exit Function
errHandler:
          
1130      If ErrMustStop Then Debug.Assert False: Resume
1140      ErrorIn "clsExchange.LoadProperties"
End Function

Public Property Get SpreadsheetType() As String
    SpreadsheetType = strSpreadsheetType
End Property
Public Property Get UsesExcel() As Boolean
    UsesExcel = (SpreadsheetType = "EXCEL")
End Property
Public Property Get UsesOpenOffice() As Boolean
    UsesOpenOffice = (SpreadsheetType = "OPENOFFICE")
End Property

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

