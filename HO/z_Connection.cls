VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private flgDbObjConnected As Boolean
Dim strSMTPServer As String
Dim strSMTP_Username As String
Dim strSMTP_Password As String
Dim strEmailFrom As String
Dim strSenderName As String
Private cnPapyShort As ADODB.Connection
Dim storeType As String
Dim mCASHSALESTPID As Long
Dim mDefaultStoreID As Long
Dim strPBKSSERVERMACHINE As String
Dim strSharedFolderRoot As String
Dim strDBName As String
Dim strMainConnectionString As String
Dim mCL As String
Dim mVATAccount As String
Dim mSupplierControlAccount As String
Dim mInventoryControlAccount As String
Dim mPeriodic_or_Perpetual As String
Dim mCostOfSalesAccount As String
Dim sStoreType As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Property Get PapyrusConnectionstring() As String
    PapyrusConnectionstring = strMainConnectionString
End Property
Public Property Get PastelConnectionstring() As String
    PastelConnectionstring = sPastelConnectionstring
End Property
Public Property Get SupplierControlAccount() As String
    SupplierControlAccount = mSupplierControlAccount
End Property
Public Property Get CostOfSalesAccount() As String
    CostOfSalesAccount = mCostOfSalesAccount
End Property
Public Property Get InventoryControlAccount() As String
    InventoryControlAccount = mInventoryControlAccount
End Property

Public Property Get Periodic_or_Perpetual() As String
    Periodic_or_Perpetual = mPeriodic_or_Perpetual
End Property
Public Property Get VATAccount() As String
    VATAccount = mVATAccount
End Property
Public Property Get LocalMachineName() As String
    LocalMachineName = sLocalMachineName
End Property
Public Property Get DBName() As String
    DBName = strDBName
End Property

Public Sub SetDBName(val As String)
    strDBName = val
End Sub

Public Function NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Function
Public Property Get COShort() As ADODB.Connection
    On Error GoTo ErrHandler
    Set COShort = cnPapyShort
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.COShort"
End Property
Public Function Connect(Optional strProvider As String) As Long
    On Error GoTo ErrHandler
Dim iConnect As Integer
    Set cnPapyShort = Nothing
    
    iConnect = OpenDBSHort
   ' DisconnectDBShort
    
    Set oPC = Me
    
EXIT_Handler:
    Connect = iConnect
    Exit Function
ErrHandler:
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
Dim rs As ADODB.Recordset
Dim fs As New FileSystemObject

'        If Not fs.FolderExists(oPC.SharedFolderRoot & "\Exchanges") Then
'            fs.CreateFolder oPC.SharedFolderRoot & "\Exchanges"
'        End If


End Sub
Public Function OpenDBSHort(Optional reStart As Boolean) As Integer
    On Error GoTo ErrHandler
Dim iCnt As Integer

    iCnt = 0
    OpenDBSHort = 0
    If mCL = "" Then
        If sStoreType = "INDEPENDENT" Then
            mCL = "PBKS"
        Else
            mCL = "PBKSHO"
        End If
    End If
    
    If cnPapyShort Is Nothing Then
        Set cnPapyShort = New ADODB.Connection
    End If
    'Can't do the following if we have one procedure using OpenDBShort calling another also using OpenDBShort
  '  If cnPapyShort.State <> adStateClosed Then
  '      cnPapyShort.Close
  '  End If
TryAgain:
    If cnPapyShort.State = adStateClosed Or reStart Then
        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strMainSQLServerName & ";Initial Catalog=" & mCL & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=36"
        MsgBox strMainConnectionString
        cnPapyShort.Open strMainConnectionString
        cnPapyShort.CommandTimeout = 240
        
    Else
        OpenDBSHort = 99
    End If
    

EXIT_Handler:
    Exit Function
ErrHandler:
    ErrPreserve
    If Err = -2147467259 And iCnt < 2 Then
        If cnPapyShort.State <> adStateClosed Then cnPapyShort.Close
        cnPapyShort.oRS.Clear
        iCnt = iCnt + 1
        GoTo TryAgain
    End If
        
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.OpenDBSHort", , , , strMainConnectionString, Array(strMainConnectionString)
End Function

Public Property Get CustomerTPID() As Long
    CustomerTPID = mCASHSALESTPID
End Property
Public Property Get DefaultStoreID() As Long
    DefaultStoreID = mDefaultStoreID
End Property
Private Sub InitializeSettings()
    On Error GoTo ErrHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String
Dim strINIFILE As String

    sLocalMachineName = NameOfPC
#If HTYPE = 1 Then
    strINIFILE = "\PBKSHO.INI"
#Else
    strINIFILE = "\PBKSWS.INI"
#End If
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & NameOfPC & "\PBKS_S"
        strMainSQLServerName = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "MAINSQLSERVER", "")
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "PBKSSERVERMACHINE", "")
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strMainSQLServerName = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "MAINSQLSERVER", "")
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "PBKSSERVERMACHINE", "")
        strSharedFolderRoot = "C:\PBKS"
    End If
    
    strPassword = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "PASSWORD", "")
    mVATAccount = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "VATACCOUNT", "")
    mPeriodic_or_Perpetual = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "Periodic_or_Perpetual", "")
    mSupplierControlAccount = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "SupplierControlAccount", "")
    mCostOfSalesAccount = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "COSTOFSALESACCOUNT", "")
    'mPeriodic_or_Perpetual
    mInventoryControlAccount = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "InventoryControlAccount", "")
    sPastelServer = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "PASTEL_SERVER", "")
    sPastelDSN = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "PASTEL_DSN", "")
    sPastelConnectionstring = "DSN=" & sPastelDSN
    sStoreType = GetIniKeyValue(strLocalRootFolder & strINIFILE, "NETWORK", "StoreType", "INDEPENDENT")
    
    Exit Sub
ErrHandler:
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

Public Sub GetSettings()
  Dim rsConfig As New ADODB.Recordset
    If sStoreType = "HO" Then
        rsConfig.Open "SELECT * FROM tConfiguration", Me.COShort
        strSMTPServer = FNS(rsConfig.Fields("CF_SMTP_Server"))
        strSMTP_Username = FNS(rsConfig.Fields("CF_SMTP_Username"))
        strSMTP_Password = FNS(rsConfig.Fields("CF_SMTP_Password"))
        strEmailFrom = FNS(rsConfig.Fields("CF_EmailFrom"))
        strSenderName = FNS(rsConfig.Fields("CF_EmailSenderName"))
    End If
End Sub

Public Property Get EMAIL_SMTP_Server() As String
    EMAIL_SMTP_Server = strSMTPServer
End Property
Public Property Get EMAIL_SMTP_Username() As String
    EMAIL_SMTP_Username = strSMTP_Username
End Property
Public Property Get EMAIL_SMTP_Password() As String
    EMAIL_SMTP_Password = strSMTP_Password
End Property
Public Property Get EMAIL_EmailFrom() As String
    EMAIL_EmailFrom = strEmailFrom
End Property
Public Property Get EMAIL_SenderName() As String
    EMAIL_SenderName = strSenderName
End Property
