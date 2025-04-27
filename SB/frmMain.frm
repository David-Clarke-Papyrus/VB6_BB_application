VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Prepare Service Broker client"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDestPort 
      Alignment       =   2  'Center
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5970
      TabIndex        =   45
      Text            =   "4025"
      Top             =   4890
      Width           =   630
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6495
      TabIndex        =   43
      Text            =   "4025"
      Top             =   1320
      Width           =   630
   End
   Begin VB.CheckBox chkCOLSReporting 
      Caption         =   "COLS reporting"
      Height          =   300
      Left            =   4020
      TabIndex        =   42
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CheckBox chkCASHUPReporting 
      Caption         =   "Cashup reporting"
      Height          =   300
      Left            =   120
      TabIndex        =   41
      Top             =   4050
      Width           =   3105
   End
   Begin VB.CheckBox chkIBT 
      Caption         =   "IBT service"
      Height          =   300
      Left            =   6975
      TabIndex        =   40
      Top             =   3825
      Width           =   2175
   End
   Begin VB.CheckBox chkInvocation 
      Caption         =   "Invocation service"
      Height          =   300
      Left            =   6960
      TabIndex        =   39
      Top             =   3150
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackupCertificate 
      Caption         =   "Backup certificate"
      Height          =   495
      Left            =   90
      TabIndex        =   38
      Top             =   8295
      Width           =   1635
   End
   Begin VB.CheckBox chkAlert 
      Caption         =   "Alert service"
      Height          =   300
      Left            =   4020
      TabIndex        =   37
      Top             =   3780
      Width           =   2175
   End
   Begin VB.CommandButton cmdPreparePublicCert_Accounting 
      Caption         =   "Step 5. (opt) Prepare public certificate for HO"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5355
      TabIndex        =   35
      Top             =   7755
      Width           =   4230
   End
   Begin VB.CommandButton Command1 
      Height          =   480
      Left            =   4785
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7665
      Width           =   510
   End
   Begin VB.TextBox txtPublicKeyFileAccounting 
      Height          =   345
      Left            =   75
      TabIndex        =   33
      Text            =   "SERVER\PBKSINSTANCE2"
      Top             =   7815
      Width           =   4695
   End
   Begin VB.TextBox txtAccountingAddress 
      Height          =   345
      Left            =   2745
      TabIndex        =   31
      Text            =   "0.0.0.0"
      Top             =   4905
      Width           =   2505
   End
   Begin VB.CheckBox chkSOH 
      Caption         =   "SOH source (share soh figures)"
      Height          =   300
      Left            =   4005
      TabIndex        =   30
      Top             =   3150
      Width           =   3420
   End
   Begin VB.CheckBox chkPastel 
      Caption         =   "Pastel source (send messages to Central)"
      Height          =   300
      Left            =   4020
      TabIndex        =   29
      Top             =   3465
      Width           =   3420
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Step 1. The database must have service broker enabled - check here before continuing"
      Height          =   405
      Left            =   240
      TabIndex        =   28
      Top             =   1800
      Width           =   7185
   End
   Begin VB.CommandButton cmdCreateRouteHUB 
      Caption         =   "Prepare route for HUB"
      Height          =   495
      Left            =   3390
      TabIndex        =   27
      Top             =   9660
      Width           =   2445
   End
   Begin VB.TextBox txtPublicKeyFileHUB 
      Height          =   345
      Left            =   90
      TabIndex        =   25
      Text            =   "SERVER\PBKSINSTANCE2"
      Top             =   7065
      Width           =   4695
   End
   Begin VB.CommandButton Command3 
      Height          =   480
      Left            =   4800
      Picture         =   "frmMain.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6915
      Width           =   510
   End
   Begin VB.CommandButton cmdPreparePublicCert_HUB 
      Caption         =   "Step 4. (opt) Prepare public certificate for HUB"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5370
      TabIndex        =   23
      Top             =   7020
      Width           =   4230
   End
   Begin VB.CommandButton cmdCreateRouteCentral 
      Caption         =   "Prepare route for CENTRAL"
      Height          =   495
      Left            =   195
      TabIndex        =   22
      Top             =   9495
      Width           =   2445
   End
   Begin VB.TextBox txtHUBAddress 
      Height          =   345
      Left            =   75
      TabIndex        =   20
      Text            =   "0.0.0.0"
      Top             =   5610
      Width           =   2565
   End
   Begin VB.TextBox txtCENTRALAddress 
      Height          =   345
      Left            =   90
      TabIndex        =   18
      Text            =   "0.0.0.0"
      Top             =   4905
      Width           =   2505
   End
   Begin VB.CommandButton cmdPreparePublicCert_CENTRAL 
      Caption         =   "Step 3. (opt) Prepare public certificate for CENTRAL"
      Enabled         =   0   'False
      Height          =   390
      Left            =   5340
      TabIndex        =   16
      Top             =   6300
      Width           =   4260
   End
   Begin VB.CommandButton cmdGetPublicKeyFile 
      Height          =   480
      Left            =   4800
      Picture         =   "frmMain.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6195
      Width           =   510
   End
   Begin VB.TextBox txtPublicKeyFileCENTRAL 
      Height          =   345
      Left            =   75
      TabIndex        =   13
      Text            =   "SERVER\PBKSINSTANCE2"
      Top             =   6345
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   8385
      Top             =   2310
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   11
      Text            =   "PBKS"
      Top             =   1305
      Width           =   630
   End
   Begin VB.TextBox txtInstancename 
      Height          =   345
      Left            =   2430
      TabIndex        =   9
      Text            =   "SERVER\PBKSINSTANCE2"
      Top             =   810
      Width           =   2985
   End
   Begin VB.CommandButton cmdPrepare 
      Caption         =   "Step 2. Prepare msg,contracts,queues and services and routes"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3015
      TabIndex        =   8
      Top             =   5490
      Width           =   5670
   End
   Begin VB.CheckBox chkHub 
      Caption         =   "Hub source (share HUB data)"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   3750
      Width           =   3105
   End
   Begin VB.CheckBox chkSales 
      Caption         =   "Sales source (sends messages to CENTRAL)"
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   3465
      Width           =   3690
   End
   Begin VB.CheckBox chkLoyalty 
      Caption         =   "Loyalty source (sends messages to CENTRAL)"
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   3165
      Width           =   3900
   End
   Begin VB.TextBox txtStoreCode 
      Height          =   345
      Left            =   2430
      TabIndex        =   3
      Text            =   "WWLB"
      Top             =   2355
      Width           =   1710
   End
   Begin VB.TextBox txtDBName 
      Height          =   345
      Left            =   2415
      TabIndex        =   0
      Text            =   "PBKS"
      Top             =   1335
      Width           =   1305
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   240
      Left            =   5565
      TabIndex        =   46
      Top             =   4950
      Width           =   360
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   240
      Left            =   6090
      TabIndex        =   44
      Top             =   1380
      Width           =   360
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Public key from HO endpoint file"
      Height          =   315
      Left            =   195
      TabIndex        =   36
      Top             =   7605
      Width           =   2925
   End
   Begin VB.Label Label11 
      Caption         =   "Accounting TCP/IP address"
      Height          =   270
      Left            =   2745
      TabIndex        =   32
      Top             =   4680
      Width           =   2745
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Public key from HUB endpoint file"
      Height          =   315
      Left            =   210
      TabIndex        =   26
      Top             =   6855
      Width           =   2925
   End
   Begin VB.Label Label9 
      Caption         =   "CENTRAL TCP/IP address"
      Height          =   270
      Left            =   90
      TabIndex        =   21
      Top             =   4680
      Width           =   2745
   End
   Begin VB.Label Label8 
      Caption         =   "HUB TCP/IP address"
      Height          =   270
      Left            =   75
      TabIndex        =   19
      Top             =   5355
      Width           =   2745
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Is this computer a . . ."
      Height          =   300
      Left            =   90
      TabIndex        =   17
      Top             =   2895
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Public key from CENTRAL endpoint file"
      Height          =   315
      Left            =   90
      TabIndex        =   14
      Top             =   6105
      Width           =   2925
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   240
      Left            =   4035
      TabIndex        =   12
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server name"
      Height          =   360
      Left            =   1110
      TabIndex        =   10
      Top             =   855
      Width           =   1260
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Store code (4 chars max)"
      Height          =   360
      Left            =   285
      TabIndex        =   4
      Top             =   2400
      Width           =   2070
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This application must be run on the computer hosting the SQL Server instance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   690
      TabIndex        =   2
      Top             =   165
      Width           =   4950
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Database name"
      Height          =   360
      Left            =   1110
      TabIndex        =   1
      Top             =   1380
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMainConnectionString As String
Dim strPBKSSERVERMACHINE As String
Dim cnn As ADODB.Connection
Dim oTF As z_TextFileSimple
Dim strFilename As String
Dim strCommandFilePath As String
Dim fs As New FileSystemObject
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long


Private Sub cmdCheck_Click()
    If CheckBrokerEnabled Then
        Me.cmdPrepare.Enabled = True
        Me.cmdPreparePublicCert_CENTRAL.Enabled = True
        Me.cmdPreparePublicCert_HUB.Enabled = True
        Me.cmdPreparePublicCert_Accounting.Enabled = True
    End If
End Sub

Private Sub cmdGetPublicKeyFile_Click()
    cd1.DialogTitle = "Public key file"
    cd1.InitDir = "c:\PBKS\BU"
    cd1.DefaultExt = "cer"
    cd1.FileName = txtPublicKeyFileCENTRAL
    cd1.ShowOpen
    strFilename = cd1.FileName
    txtPublicKeyFileCENTRAL = strFilename
End Sub

Private Sub cmdPrepare_Click()
    If txtStoreCode = "" Then
        MsgBox "Enter a store code"
        Exit Sub
    End If
    If MsgBox("You must have run the UPDATES.SQL script before doing this as is places the activation stored procedures in place, If they are not in place some of the commands generated by this application will fail. (The creation of queues for instance)", vbQuestion + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    End If
    SaveSettings
    
    If OpenDB() = 0 Then
    
        PrepareScript
        If fs.FileExists(strCommandFilePath) Then
            ExecuteScript
        Else
            MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
        End If
        CloseDB
    Else
        MsgBox "Cannot open database. Script has not run"
    End If
    MsgBox "Remember: if you are installing for the first time to set the EXCH_SentToCentralAt value to something like '2008-01-01'" & vbCrLf _
       & "update tExchange SET EXCH_SentTOCentralAt = '2008-01-01' WHERE EXCH_SENTTOCENTRALAT IS NULL"
End Sub
Private Sub cmdBackupCertificate_Click()
    If OpenDB() = 0 Then
    
        PrepareBackupCertScript
        If fs.FileExists(strCommandFilePath) Then
            ExecuteScript
        Else
            MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
        End If
        CloseDB
    Else
        MsgBox "Cannot open database. Script has not run"
    End If

End Sub

Private Sub PrepareBackupCertScript()
Dim strPath As String
Dim fs As New FileSystemObject

    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    strPath = "C:\PBKS\BU\"
    
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        If fs.FileExists(strPath & UCase(txtStoreCode) & "_CERT.cer") Then
            fs.DeleteFile strPath & UCase(txtStoreCode) & "_CERT.cer"
        End If
        oTF.WriteToTextFile "BACKUP CERTIFICATE [" & UCase(txtStoreCode) & "_CERT] TO FILE = '" & strPath & UCase(txtStoreCode) & "_CERT.cer'"
        oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
    End If
End Sub
Private Sub PrepareScript()
Dim strPath As String
Dim fs As New FileSystemObject
Dim strEndpointname As String
Dim strCertificatename As String

    strEndpointname = txtStoreCode & "_ENDPOINT"
    strCertificatename = txtStoreCode & "_CERT"
    strPath = "C:\PBKS\BU\" & strCertificatename & ".CER"

    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE MASTER KEY ENCRYPTION BY PASSWORD = '9Crank0" + txtStoreCode & "'"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP CERTIFICATE [" & UCase(strCertificatename) & "]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [" & UCase(strCertificatename) & "] WITH SUBJECT  = '" & UCase(strCertificatename) & "', START_DATE = '01/01/2005', EXPIRY_DATE = '01/01/2100'"
        oTF.WriteToTextFile "GO"
        
        strPath = "C:\PBKS\BU\"
        
        oTF.WriteToTextFile "CREATE ENDPOINT [" & UCase(strEndpointname) & "] STATE=STARTED AS TCP (LISTENER_PORT = " & txtPort & ") FOR SERVICE_BROKER ( AUTHENTICATION = CERTIFICATE [" & UCase(strCertificatename) & "])"
        oTF.WriteToTextFile "GO"
        
        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
        oTF.WriteToTextFile "GO"

        oTF.WriteToTextFile "UPDATE tCONFIGURATION SET CF_INSTALLATIONCODE = '" & UCase(Trim(txtStoreCode)) & "'"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GO"
       
        
        oTF.WriteToTextFile "CREATE MESSAGE TYPE [EndOfStream] AUTHORIZATION [dbo] VALIDATION = NONE"
        oTF.WriteToTextFile "GO"
        If Me.chkAlert = 1 Then
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [ALERT_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [ALERTREAD_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [ALERT_CONTRACT] AUTHORIZATION [dbo] ([ALERT_MSG] SENT BY INITIATOR,[ALERTREAD_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[ALERT_CONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_ALERT_CONSUMER] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [ALERT_CONSUMER_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[ALERT_CONSUMER_Q] ([ALERT_CONTRACT])"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [ALERTLOAD_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [ALERTLOAD_CONTRACT] AUTHORIZATION [dbo] ([ALERTLOAD_MSG] SENT BY INITIATOR,[ALERTREAD_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[ALERTLOAD_SOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_ALERTLOAD_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [ALERTLOAD_SOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[ALERTLOAD_SOURCE_Q] ([ALERTLOAD_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::ALERT_CONSUMER_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        
        End If
        
        If Me.chkSales = 1 Then
            oTF.WriteToTextFile "DROP SERVICE [SALESSOURCE_" & UCase(txtStoreCode) & "_SERVICE]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP QUEUE [dbo].[SALESSOURCE_Q] "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP CONTRACT [SALES_CONTRACT]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP MESSAGE TYPE [EndOfStream]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP MESSAGE TYPE [SALES_MSG]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SALES_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [EndOfStream] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [SALES_CONTRACT] AUTHORIZATION [dbo] ([SALES_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[SALESSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_SALES_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [SALESSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[SALESSOURCE_Q] ([SALES_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::SALESSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        End If
        'POS reporting
        If chkCASHUPReporting = 1 Then
            oTF.WriteToTextFile "DROP SERVICE [CASHUPSOURCE_" & UCase(txtStoreCode) & "_SERVICE]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP QUEUE [dbo].[CASHUPSOURCE_Q] "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP CONTRACT [CASHUP_CONTRACT]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP MESSAGE TYPE [CASHUP_MSG]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP MESSAGE TYPE [CASHUP_TEXT_MSG]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUP_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUP_TEXT_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [EndOfStream] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [CASHUP_CONTRACT] AUTHORIZATION [dbo] ([CASHUP_MSG] SENT BY INITIATOR,[CASHUP_TEXT_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[CASHUPSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_CASHUP_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [CASHUPSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[CASHUPSOURCE_Q] ([CASHUP_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::CASHUPSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
        End If
'        If chkCOLSReporting = 1 Then
'            oTF.WriteToTextFile "DROP SERVICE [COLSSOURCE_" & UCase(txtStoreCode) & "_SERVICE]"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "DROP QUEUE [dbo].[CASHUPSOURCE_Q] "
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "DROP CONTRACT [CASHUP_CONTRACT]"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "DROP MESSAGE TYPE [CASHUP_MSG]"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "DROP MESSAGE TYPE [CASHUP_TEXT_MSG]"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUP_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUP_TEXT_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "CREATE MESSAGE TYPE [EndOfStream] AUTHORIZATION [dbo] VALIDATION = NONE"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "CREATE CONTRACT [CASHUP_CONTRACT] AUTHORIZATION [dbo] ([CASHUP_MSG] SENT BY INITIATOR,[CASHUP_TEXT_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "CREATE QUEUE [dbo].[CASHUPSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_CASHUP_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "CREATE SERVICE [CASHUPSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[CASHUPSOURCE_Q] ([CASHUP_CONTRACT])"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "GRANT SEND ON SERVICE::CASHUPSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
'        End If
        If Me.chkSOH = 1 Then
        'Data provision service to Central
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOH_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOHALL_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [SOH_CONTRACT] AUTHORIZATION [dbo] ([SOHALL_MSG] SENT BY INITIATOR,[SOH_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[SOHSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_SOH_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [SOHSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[SOHSOURCE_Q] ([SOH_CONTRACT])"
            oTF.WriteToTextFile "GO"
        'Data service in use by branch
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOH_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOH_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [SOH_DIALOG_CONTRACT] AUTHORIZATION [dbo] ([SOH_RQ_MSG] SENT BY INITIATOR,[SOH_RS_MSG] SENT BY TARGET,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[SOH_RQ_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_SOH_RQ_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [SOH_RQ_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[SOH_RQ_Q] ([SOH_DIALOG_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::SOH_RQ_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::SOHSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        End If
        If Me.chkLoyalty = 1 Then
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [LOYALTY_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [LOYALTY_CONTRACT] AUTHORIZATION [dbo] ([LOYALTY_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[LOYALTYSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_LOYALTY_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [LOYALTYSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[LOYALTYSOURCE_Q] ([LOYALTY_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [MASTERLOYALTY_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [MASTERLOYALTY_CONTRACT] AUTHORIZATION [dbo] ([EndOfStream] SENT BY INITIATOR,[MASTERLOYALTY_MSG] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[MASTERLOYALTYCONSUMER_Q] WITH STATUS = OFF , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_MASTERLOYALTY] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) ON [PRIMARY]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [MLC_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[MASTERLOYALTYCONSUMER_Q] ([MASTERLOYALTY_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::LOYALTYSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::MLC_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        End If
        If Me.chkHub = 1 Then
        'Data provision service to HUB
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [HUB_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [HUB_CONTRACT] AUTHORIZATION [dbo] ([HUB_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[HUBSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_HUB_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [HUBSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[HUBSOURCE_Q] ([HUB_CONTRACT])"
            oTF.WriteToTextFile "GO"
            
        'Data service in use by branch
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [HUB_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [HUB_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [HUB_DIALOG_CONTRACT] AUTHORIZATION [dbo] ([HUB_RQ_MSG] SENT BY INITIATOR,[HUB_RS_MSG] SENT BY TARGET,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[HUB_RQ_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_HUB_RQ_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [HUB_RQ_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[HUB_RQ_Q] ([HUB_DIALOG_CONTRACT])"
            oTF.WriteToTextFile "GO"
        
        End If
        If Me.chkPastel = 1 Then
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [PASTELDRCRJOURNALS_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [PASTEL_CONTRACT] AUTHORIZATION [dbo] ([PASTELDRCRJOURNALS_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[PASTELSOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_PASTEL_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [PASTELSOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[PASTELSOURCE_Q] ([PASTEL_CONTRACT])"
            oTF.WriteToTextFile "GO"
        End If
        
    'I B T and I B T R----------------------------------------
        If Me.chkIBT = 1 Then

            oTF.WriteToTextFile "CREATE MESSAGE TYPE [IBT_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [IBT_CONTRACT] AUTHORIZATION [dbo] ([IBT_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[IBT_SOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_IBT_Response] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [IBT_SOURCE_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[IBT_SOURCE_Q] ([IBT_CONTRACT])"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [IBTR_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [IBTR_CONTRACT] AUTHORIZATION [dbo] ([IBTR_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[IBTR_CONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_IBTR_CONSUMER] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [IBTR_CONSUMER_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[IBTR_CONSUMER_Q] ([IBTR_CONTRACT])"
            oTF.WriteToTextFile "GO"
     
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::IBTR_CONSUMER_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::IBT_SOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
     '--------------------------------------------------------
        End If
        
        
        If Me.chkInvocation = 1 Then
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CUSTOMERSTATS_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CUSTOMERSTATS_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CUSTOMERSET_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CUSTOMERSET_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SALESSET_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SALESSET_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOHALL_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOHALL_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUPS_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUPS_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [COLS_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [COLS_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [BUDGETDATA_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [GENERAL_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
           
            
'--xxxxxxxxxxxxxxClearing
            oTF.WriteToTextFile "DROP SERVICE [INVOCATION_RQ_" & UCase(txtStoreCode) & "_SERVICE]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP QUEUE [dbo].[INVOCATION_RQ_Q] "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP CONTRACT [INVOCATION_CONTRACT]"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE CONTRACT [INVOCATION_CONTRACT] AUTHORIZATION [dbo] ([SOHALL_RQ_MSG] SENT BY INITIATOR," _
                                                                                  & "[SOHALL_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[CUSTOMERSTATS_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[CUSTOMERSTATS_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[CUSTOMERSET_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[CUSTOMERSET_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[CASHUPS_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[CASHUPS_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[COLS_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[COLS_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[SALESSET_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[SALESSET_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[BUDGETDATA_MSG] SENT BY INITIATOR, " _
                                                                                  & "[GENERAL_MSG] SENT BY INITIATOR, " _
                                                                                  & "[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[INVOCATION_RQ_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_INVOCATION_RQ_CONSUMER] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [INVOCATION_RQ_" & UCase(txtStoreCode) & "_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[INVOCATION_RQ_Q] ([INVOCATION_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::INVOCATION_RQ_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        
        End If

'--------------------Prepare Routes
        If txtHUBAddress > "" Then
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_HUB]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_HUB] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'HUBCONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtHUBAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
        End If
        If txtCENTRALAddress > "" Then
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_SALES]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_LOYALTY]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_MASTERLOYALTY]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_SOH]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_CASHUP]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_SOH_RS]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_ALERT_SOURCE]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_ALERTLOAD_CONSUMER]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CENTRAL_INVOCATION_RS]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_SALES] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'SALESCONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_LOYALTY] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'LOYALTYCONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_MASTERLOYALTY]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'MASTERLOYALTYSOURCE_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_SOH]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'SOHCONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_CASHUP]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'CASHUPCONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_SOH_RS]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'SOH_RS_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_ALERT_SOURCE]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'ALERT_SOURCE_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_ALERTLOAD_CONSUMER]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'ALERTLOAD_CONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_INVOCATION_RS]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'INVOCATION_RS_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_IBT_CONSUMER] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'IBT_CONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CENTRAL_IBTR_SOURCE] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'IBTR_CONSUMER_SOURCE_SERVICE' ,  ADDRESS  = N'TCP://" & txtCENTRALAddress & ":" & txtDestPort & "'"
            
            oTF.WriteToTextFile "GO"
        If txtAccountingAddress > "" Then
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_ACCOUNTING_PASTEL]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_ACCOUNTING_PASTEL]   AUTHORIZATION [dbo]   WITH  SERVICE_NAME  = N'PASTELCONSUMER_SERVICE' ,  ADDRESS  = N'TCP://" & txtAccountingAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
        End If
        
        End If

        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
'txtPublicKeyFile
End Sub
Private Sub ExecuteScript()
Dim strCommand As String
Dim res As Boolean
    
    strCommand = "SQLCMD -Usa -P" & Me.txtPassword & " -S" & Me.txtInstancename & " -d" & txtDBName & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
End Sub
Sub GetInitialValues()
  '  strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)

End Sub

'--------------------------------------------------------------------------

'
'
'Private Sub cmdPreparePublicCert_Click()
'    strCommandFilePath = "\\" & NameOfPC & "\PBKS_S\PrepareServiceBrokerScriptPub.SQL"
'        Set oTF = New z_TextFileSimple
'        oTF.OpenTextFile strCommandFilePath
'        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
'        oTF.WriteToTextFile "GO"
'        oTF.WriteToTextFile "CREATE LOGIN [CENTRAL_LOGIN] WITH PASSWORD=N'9Crank0';"
'        oTF.WriteToTextFile "GO"
'        oTF.WriteToTextFile "CREATE USER [CENTRAL_USER];"
'        oTF.WriteToTextFile "GO"
'        oTF.WriteToTextFile "CREATE CERTIFICATE [CENTRAL_CERT] AUTHORIZATION [CENTRAL_DATASOURCE] FROM FILE = '" & Me.txtPublicKeyFile & "';"
'        oTF.WriteToTextFile "GO"
'        oTF.WriteToTextFile "GRANT CONNECT ON ENDPOINT::WWLB_FEEDDATASOURCE_ENDPOINT TO [CENTRAL_DATASOURCE];"
'        oTF.WriteToTextFile "GO"
'        oTF.WriteToTextFile "GRANT SEND ON SERVICE::C_LOYALTYFEED_DATASOURCE_SERVICE to [PUBLIC]"
'        oTF.WriteToTextFile "GO"
'        oTF.CloseTextFile
'        Set oTF = Nothing
'    Set oTF = Nothing
'End Sub
'


Public Property Get NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
errHandler:
MsgBox Error
End Property


Public Function OpenDB() As Integer
    On Error GoTo errHandler
    OpenDB = 0
    If cnn Is Nothing Then
        Set cnn = New ADODB.Connection
    End If
    If cnn.Errors.Count > 0 Or cnn.State = adStateClosed Then
        On Error Resume Next
        cnn.Close
        On Error GoTo errHandler
        strMainConnectionString = "Provider=SQLOLEDB;Data Source=" & Me.txtInstancename & ";Initial Catalog=" & txtDBName & ";User Id=sa;Password=" & txtPassword & ";Connect Timeout=45"
        cnn.Open strMainConnectionString
        cnn.CommandTimeout = 360
    Else
        If cnn.State <> adStateOpen Then
            OpenDB = 99
        End If
    End If
EXIT_HANDLER:
    Exit Function
    
errHandler:
    MsgBox Error
End Function
Public Sub CloseDB()
    If cnn Is Nothing Then Exit Sub
    If cnn.State = 1 Then   'it is open
        cnn.Close
        Set cnn = Nothing
    End If
End Sub
Private Function CheckBrokerEnabled() As Boolean
Dim rs As New ADODB.Recordset
Dim bEnabled As Boolean

    If OpenDB() = 0 Then
StartCheck:
        rs.Open "SELECT is_broker_enabled FROM master.sys.databases where name = '" & txtDBName & "'", cnn, adOpenKeyset
        If rs.State <> 0 Then
            If rs.EOF <> True Then
                bEnabled = rs.Fields(0)
            Else
                bEnabled = False
            End If
        Else
            bEnabled = False
        End If
        rs.Close
        If Not bEnabled Then
            If MsgBox("Service Broker is not enabled on " & txtDBName & "." & vbCrLf & "Do you want to enable it?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                Exit Function
            Else
                MsgBox "Ensure all applications connected to the database are closed before continuing."
                cnn.CommandTimeout = 30
                On Error Resume Next
                cnn.Execute "ALTER DATABASE  " & txtDBName & " SET ENABLE_BROKER"
                If Err <> 0 Then
                    MsgBox "The following error occurred: " & Error
                End If
            End If
            GoTo StartCheck
        End If
    Else
        MsgBox "Cannot open database. Script has not run"
    End If
    CheckBrokerEnabled = bEnabled
End Function

Private Sub cmdPreparePublicCert_Accounting_Click()
Dim str As String
    str = Replace(UCase(Replace(txtInstancename, "\", "_")), "-", "_")
    If OpenDB() = 0 Then
        strCommandFilePath = "\\" & NameOfPC & "\PBKS_S\PrepareServiceBrokerScriptPubAccounting.SQL"
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP LOGIN HO"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP CERTIFICATE HO_CERT"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [HO_CERT] FROM FILE = '" & Me.txtPublicKeyFileAccounting & "';"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE LOGIN [HO] FROM CERTIFICATE [HO_CERT]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GRANT CONNECT ON ENDPOINT::" & UCase(Me.txtStoreCode) & "_ENDPOINT TO [HO];"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GRANT SEND ON SERVICE::HOSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
        oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
        Set oTF = Nothing
        If fs.FileExists(strCommandFilePath) Then
            ExecuteScript
        Else
            MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
        End If
        CloseDB
    Else
        MsgBox "Cannot open database. Script has not run"
    End If
    SaveSetting "SB", "SBC", "PublicKeyFileAccounting", Me.txtPublicKeyFileAccounting
End Sub

Private Sub cmdPreparePublicCert_CENTRAL_Click()
Dim str As String

    If txtStoreCode = "" Then
        MsgBox "Enter a store code"
        Exit Sub
    End If
    str = Replace(UCase(Replace(txtInstancename, "\", "_")), "-", "_")
    If OpenDB() = 0 Then
        strCommandFilePath = "\\" & NameOfPC & "\PBKS_S\PrepareServiceBrokerScriptPubCENTRAL.SQL"
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP LOGIN CENTRAL"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP CERTIFICATE CENTRAL_CERT"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [CENTRAL_CERT] FROM FILE = '" & Me.txtPublicKeyFileCENTRAL & "';"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE LOGIN [CENTRAL] FROM CERTIFICATE [CENTRAL_CERT]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GRANT CONNECT ON ENDPOINT::" & UCase(Me.txtStoreCode) & "_ENDPOINT TO [CENTRAL];"
        oTF.WriteToTextFile "GO"
        
        oTF.CloseTextFile
        Set oTF = Nothing
        If fs.FileExists(strCommandFilePath) Then
            ExecuteScript
        Else
            MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
        End If
        CloseDB
    Else
        MsgBox "Cannot open database. Script has not run"
    End If

    SaveSetting "SB", "SBC", "PublicKeyFileCENTRAL", Me.txtPublicKeyFileCENTRAL

End Sub

Private Sub cmdPreparePublicCert_HUB_Click()
Dim str As String

    str = Replace(UCase(Replace(txtInstancename, "\", "_")), "-", "_")
    If OpenDB() = 0 Then
        strCommandFilePath = "\\" & NameOfPC & "\PBKS_S\PrepareServiceBrokerScriptPubHUB.SQL"
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP LOGIN HUB"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP CERTIFICATE HUB_CERT"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [HUB_CERT] FROM FILE = '" & Me.txtPublicKeyFileHUB & "';"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE LOGIN [HUB] FROM CERTIFICATE [HUB_CERT]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GRANT CONNECT ON ENDPOINT::" & UCase(Me.txtStoreCode) & "_ENDPOINT TO [HUB];"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GRANT SEND ON SERVICE::HUBSOURCE_" & UCase(txtStoreCode) & "_SERVICE to [PUBLIC]"
        oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
        Set oTF = Nothing
        If fs.FileExists(strCommandFilePath) Then
            ExecuteScript
        Else
            MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
        End If
        CloseDB
    Else
        MsgBox "Cannot open database. Script has not run"
    End If
    SaveSetting "SB", "SBC", "PublicKeyFileHUB", Me.txtPublicKeyFileHUB

End Sub

Private Sub Command1_Click()
    cd1.DialogTitle = "Public key file"
    cd1.InitDir = "c:\PBKS\BU"
    cd1.DefaultExt = "cer"
    cd1.FileName = txtPublicKeyFileAccounting
    cd1.ShowOpen
    strFilename = cd1.FileName
    txtPublicKeyFileAccounting = strFilename
End Sub

Private Sub Command3_Click()
    cd1.DialogTitle = "Public key file"
    cd1.InitDir = "c:\PBKS\BU"
    cd1.DefaultExt = "cer"
    cd1.FileName = txtPublicKeyFileHUB
    cd1.ShowOpen
    strFilename = cd1.FileName
    txtPublicKeyFileHUB = strFilename

End Sub

Private Sub Form_Load()
    Set oTF = New z_TextFileSimple
    
    txtInstancename = GetSetting("SB", "SBC", "INSTANCENAME", "")
    txtDBName = GetSetting("SB", "SBC", "DBNAME", "")
    txtStoreCode = GetSetting("SB", "SBC", "STORECODE", "")
    
    chkLoyalty = IIf(GetSetting("SB", "SBC", "bLOYALTY", "") = "1", 1, 0)
    chkSOH = IIf(GetSetting("SB", "SBC", "bSOH", "") = "1", 1, 0)
    chkSales = IIf(GetSetting("SB", "SBC", "bSales", "") = "1", 1, 0)
    chkHub = IIf(GetSetting("SB", "SBC", "bHUB", "") = "1", 1, 0)
    chkPastel = IIf(GetSetting("SB", "SBC", "bPastel", "") = "1", 1, 0)
    chkAlert = IIf(GetSetting("SB", "SBC", "bAlert", "") = "1", 1, 0)
    chkInvocation = IIf(GetSetting("SB", "SBC", "bInvocation", "") = "1", 1, 0)
    chkCASHUPReporting = IIf(GetSetting("SB", "SBC", "bCashupReporting", "") = "1", 1, 0)
    
    txtCENTRALAddress = GetSetting("SB", "SBC", "CENTRALAddress", "")
    txtHUBAddress = GetSetting("SB", "SBC", "HUBAddress", "")
    txtAccountingAddress = GetSetting("SB", "SBC", "AccountingAddress", "")
        
    txtPublicKeyFileAccounting = GetSetting("SB", "SBC", "PublicKeyFileAccounting", "")
    txtPublicKeyFileHUB = GetSetting("SB", "SBC", "PublicKeyFileHUB", "")
    txtPublicKeyFileCENTRAL = GetSetting("SB", "SBC", "PublicKeyFileCENTRAL", "")
        
End Sub
Private Sub SaveSettings()
    SaveSetting "SB", "SBC", "INSTANCENAME", Me.txtInstancename
    SaveSetting "SB", "SBC", "DBNAME", Me.txtDBName
    SaveSetting "SB", "SBC", "STORECODE", Me.txtStoreCode
    
    SaveSetting "SB", "SBC", "bLOYALTY", IIf(Me.chkLoyalty = 1, "1", "0")
    SaveSetting "SB", "SBC", "bSOH", IIf(Me.chkSOH = 1, "1", "0")
    SaveSetting "SB", "SBC", "bSales", IIf(Me.chkSales = 1, "1", "0")
    SaveSetting "SB", "SBC", "bHUB", IIf(Me.chkHub = 1, "1", "0")
    SaveSetting "SB", "SBC", "bPastel", IIf(Me.chkPastel = 1, "1", "0")
    SaveSetting "SB", "SBC", "bAlert", IIf(Me.chkAlert = 1, "1", "0")
    SaveSetting "SB", "SBC", "bInvocation", IIf(Me.chkInvocation = 1, "1", "0")
    SaveSetting "SB", "SBC", "bCashupReporting", IIf(Me.chkCASHUPReporting = 1, "1", "0")
    
    SaveSetting "SB", "SBC", "CENTRALAddress", Me.txtCENTRALAddress
    SaveSetting "SB", "SBC", "HUBAddress", Me.txtHUBAddress
    SaveSetting "SB", "SBC", "AccountingAddress", Me.txtAccountingAddress

    SaveSetting "SB", "SBC", "PublicKeyFileCENTRAL", Me.txtPublicKeyFileCENTRAL
    SaveSetting "SB", "SBC", "PublicKeyFileHUB", Me.txtPublicKeyFileHUB
    SaveSetting "SB", "SBC", "PublicKeyFileAccounting", Me.txtPublicKeyFileAccounting

End Sub


