VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Prepare Service Broker on CENTRAL or HUB"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDestPort 
      Alignment       =   2  'Center
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5055
      TabIndex        =   30
      Text            =   "4025"
      Top             =   4365
      Width           =   630
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6390
      TabIndex        =   28
      Text            =   "4025"
      Top             =   1800
      Width           =   630
   End
   Begin VB.TextBox txtNotes 
      Height          =   1545
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   27
      Text            =   "frmMain.frx":0000
      Top             =   6930
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Backup Certificate"
      Height          =   495
      Left            =   60
      TabIndex        =   26
      Top             =   6945
      Width           =   2025
   End
   Begin VB.TextBox txtEndpointName 
      Height          =   345
      Left            =   2430
      TabIndex        =   23
      Text            =   "SERVER_PBKSINSTANCE2"
      Top             =   1785
      Width           =   2985
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1050
      Left            =   345
      TabIndex        =   20
      Top             =   2910
      Width           =   2385
      Begin VB.OptionButton optHUB 
         Caption         =   "HUB"
         Height          =   300
         Left            =   225
         TabIndex        =   22
         Top             =   615
         Width           =   1695
      End
      Begin VB.OptionButton optCentral 
         Caption         =   "CENTRAL"
         Height          =   300
         Left            =   225
         TabIndex        =   21
         Top             =   210
         Value           =   -1  'True
         Width           =   2085
      End
   End
   Begin VB.CommandButton cmdPrepareCENTRAL 
      Caption         =   "Step 2a. Prepare CENTRAL Services"
      Enabled         =   0   'False
      Height          =   495
      Left            =   390
      TabIndex        =   19
      Top             =   4995
      Width           =   2445
   End
   Begin VB.TextBox txtClientName 
      Height          =   345
      Left            =   345
      TabIndex        =   17
      Text            =   "HUB"
      Top             =   4365
      Width           =   1305
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Step 1. The database must have service broker enabled - check here before continuing"
      Height          =   405
      Left            =   225
      TabIndex        =   16
      Top             =   2295
      Width           =   7185
   End
   Begin VB.CommandButton cmdCreateRouteHUB 
      Caption         =   "Prepare route for HUB"
      Height          =   495
      Left            =   3390
      TabIndex        =   15
      Top             =   9660
      Width           =   2445
   End
   Begin VB.CommandButton cmdCreateRouteCentral 
      Caption         =   "Prepare route for CENTRAL"
      Height          =   495
      Left            =   195
      TabIndex        =   14
      Top             =   9495
      Width           =   2445
   End
   Begin VB.TextBox txtAddress 
      Height          =   345
      Left            =   1860
      TabIndex        =   12
      Text            =   "0.0.0.0"
      Top             =   4365
      Width           =   2505
   End
   Begin VB.CommandButton cmdPreparePublicCert 
      Caption         =   "Step 3. Prepare public certificate for client"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2400
      TabIndex        =   11
      Top             =   6330
      Width           =   3000
   End
   Begin VB.CommandButton cmdGetPublicKeyFile 
      Height          =   480
      Left            =   5265
      Picture         =   "frmMain.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5865
      Width           =   510
   End
   Begin VB.TextBox txtPublicKeyFile 
      Height          =   345
      Left            =   510
      TabIndex        =   8
      Text            =   "SERVER\PBKSINSTANCE2"
      Top             =   5925
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   -30
      Top             =   5565
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
      TabIndex        =   6
      Text            =   "PBKS"
      Top             =   1305
      Width           =   630
   End
   Begin VB.TextBox txtInstancename 
      Height          =   345
      Left            =   2430
      TabIndex        =   4
      Text            =   "SERVER\PBKSINSTANCE2"
      Top             =   810
      Width           =   4575
   End
   Begin VB.CommandButton cmdPrepareHUB 
      Caption         =   "Step 2b. Prepare HUB Services"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2925
      TabIndex        =   3
      Top             =   4980
      Width           =   2025
   End
   Begin VB.TextBox txtDBName 
      Height          =   345
      Left            =   2415
      TabIndex        =   0
      Text            =   "HUB"
      Top             =   1335
      Width           =   1305
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   240
      Left            =   4650
      TabIndex        =   31
      Top             =   4425
      Width           =   360
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      Height          =   240
      Left            =   5985
      TabIndex        =   29
      Top             =   1860
      Width           =   360
   End
   Begin VB.Label Label8 
      Caption         =   "If address > ''then a route is prepared for that client"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   330
      TabIndex        =   25
      Top             =   4710
      Width           =   4380
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Endpoint name"
      Height          =   360
      Left            =   1110
      TabIndex        =   24
      Top             =   1830
      Width           =   1260
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENT name"
      Height          =   270
      Left            =   330
      TabIndex        =   18
      Top             =   4140
      Width           =   1260
   End
   Begin VB.Label Label9 
      Caption         =   "CLIENT TCP/IP address"
      Height          =   270
      Left            =   1860
      TabIndex        =   13
      Top             =   4140
      Width           =   2745
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Public key from client endpoint"
      Height          =   315
      Left            =   495
      TabIndex        =   9
      Top             =   5685
      Width           =   2925
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   240
      Left            =   4035
      TabIndex        =   7
      Top             =   1365
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SQL Server name"
      Height          =   360
      Left            =   1110
      TabIndex        =   5
      Top             =   855
      Width           =   1260
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
        Frame1.Enabled = True
    End If
End Sub

Private Sub cmdGetPublicKeyFile_Click()
    cd1.DialogTitle = "Public key file"
    cd1.InitDir = "c:\PBKS\BU"
    cd1.DefaultExt = "cer"
    cd1.ShowOpen
    strFilename = cd1.FileName
    txtPublicKeyFile = strFilename
    cmdPreparePublicCert.Enabled = (Len(strFilename) > 0)
End Sub

Private Sub cmdPrepareCENTRAL_Click()
    If MsgBox("You must have run the UPDATES.SQL script before doing this as is places the activation stored procedures in place, If they are not in place some of the commands generated by this application will fail. (The creation of queues for instance)", vbQuestion + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    End If
    SaveSettings
    
    If OpenDB() = 0 Then
    
        PrepareScript_CENTRAL
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

Private Sub cmdPrepareHUB_Click()

    If MsgBox("You must have run the UPDATES.SQL script before doing this as is places the activation stored procedures in place, If they are not in place some of the commands generated by this application will fail. (The creation of queues for instance)", vbQuestion + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    End If
    SaveSettings
    
    If OpenDB() = 0 Then
    
        PrepareScript_HUB
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

Private Sub PrepareScript_HUB()
Dim strPath As String
Dim fs As New FileSystemObject
Dim strEndpointname As String
Dim strCertificatename As String

    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
    strEndpointname = txtEndpointName & "_ENDPOINT"
    strCertificatename = txtEndpointName & "_CERT"
    strPath = "C:\PBKS\BU\" & strCertificatename & ".CER"
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE MASTER KEY ENCRYPTION BY PASSWORD = '9Crank0HUB'"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [" & strCertificatename & "] WITH SUBJECT  = '" & strCertificatename & "', START_DATE = '01/01/2005', EXPIRY_DATE = '01/01/2100'"
        oTF.WriteToTextFile "GO"
        
        If fs.FileExists(strPath) Then
            fs.DeleteFile strPath
        End If
        oTF.WriteToTextFile "CREATE ENDPOINT [" & strEndpointname & "] STATE=STARTED AS TCP (LISTENER_PORT = " & txtPort & ") FOR SERVICE_BROKER ( AUTHENTICATION = CERTIFICATE [" & strCertificatename & "])"
        oTF.WriteToTextFile "GO"
        
        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
        oTF.WriteToTextFile "GO"

        oTF.WriteToTextFile "CREATE MESSAGE TYPE [HUB_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE MESSAGE TYPE [EndOfStream] AUTHORIZATION [dbo] VALIDATION = NONE"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CONTRACT [HUB_CONTRACT] AUTHORIZATION [dbo] ([HUB_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE QUEUE [dbo].[HUBCONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_HUBCONSUMER] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE SERVICE [HUBCONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[HUBCONSUMER_Q] ([HUB_CONTRACT])"
        oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::HUBCONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        If txtAddress > "" Then
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'HUBSOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
        End If

        oTF.CloseTextFile
        Set oTF = Nothing
        txtNotes.Text = "End-point created: '" & strEndpointname & "'"
        txtNotes.Text = "Certificate created: '" & strCertificatename & "'"
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
'txtPublicKeyFile
End Sub
Private Sub PrepareScript_CENTRAL()
Dim strPath As String
Dim fs As New FileSystemObject
Dim strEndpointname As String
Dim strCertificatename As String

    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
    strEndpointname = txtEndpointName & "_ENDPOINT"
    strCertificatename = txtEndpointName & "_CERT"
    strPath = "C:\PBKS\BU\" & strCertificatename & ".CER"
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE MASTER KEY ENCRYPTION BY PASSWORD = '9Crank0CENTRAL'"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [" & strCertificatename & "] WITH SUBJECT  = '" & strCertificatename & "', START_DATE = '01/01/2005', EXPIRY_DATE = '01/01/2100'"
        oTF.WriteToTextFile "GO"
        'The following allows T-SQL to call .DLLs
        oTF.WriteToTextFile "sp_configure 'show advanced options', 1;"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "RECONFIGURE;"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "sp_configure 'Ole Automation Procedures', 1;"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "RECONFIGURE;"
        oTF.WriteToTextFile "GO"

        oTF.WriteToTextFile "CREATE ENDPOINT [" & strEndpointname & "] STATE=STARTED AS TCP (LISTENER_PORT = " & txtPort & ") FOR SERVICE_BROKER ( AUTHENTICATION = CERTIFICATE [" & strCertificatename & "])"
        oTF.WriteToTextFile "GO"

        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
        oTF.WriteToTextFile "GO"

            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SALES_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [EndOfStream] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [SALES_CONTRACT] AUTHORIZATION [dbo] ([SALES_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[SALESCONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_SALES] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [SALESCONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[SALESCONSUMER_Q] ([SALES_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::SALESCONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [LOYALTY_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [LOYALTY_CONTRACT] AUTHORIZATION [dbo] ([LOYALTY_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[LOYALTYCONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_LOYALTY] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [LOYALTYCONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[LOYALTYCONSUMER_Q] ([LOYALTY_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::LOYALTYCONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOH_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOH_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOH_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [SOH_CONTRACT] AUTHORIZATION [dbo] ([SOH_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [SOH_DIALOG_CONTRACT] AUTHORIZATION [dbo] ([SOH_RS_MSG] SENT BY TARGET,[SOH_RQ_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[SOHCONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_SOH] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[SOH_RS_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_SOH_RS] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [SOHCONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[SOHCONSUMER_Q] ([SOH_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [SOH_RS_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[SOH_RS_Q] ([SOH_DIALOG_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::SOHCONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
'Cashup stuff --------------------------------------------------------------
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUP_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUP_TEXT_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [CASHUP_CONTRACT] AUTHORIZATION [dbo] ([CASHUP_MSG] SENT BY INITIATOR,[CASHUP_TEXT_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[CASHUPCONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_Cashup_Consumer] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [CASHUPCONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[CASHUPCONSUMER_Q] ([CASHUP_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::CASHUPCONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"

'-----------
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [ALERT_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [ALERTREAD_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [ALERT_CONTRACT] AUTHORIZATION [dbo] ([ALERT_MSG] SENT BY INITIATOR,[ALERTREAD_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[ALERTRESPONSE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_ALERTSOURCE_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [ALERT_SOURCE_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[ALERTRESPONSE_Q] ([ALERT_CONTRACT])"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [ALERTLOAD_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [ALERTLOAD_CONTRACT] AUTHORIZATION [dbo] ([ALERTLOAD_MSG] SENT BY INITIATOR,[ALERTREAD_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[ALERTLOAD_CONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_ALERTLOAD] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [ALERTLOAD_CONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[ALERTLOAD_CONSUMER_Q] ([ALERTLOAD_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::ALERT_SOURCE_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::ALERTLOAD_CONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
    
    'I B T and I B T R----------------------------------------
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [IBT_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [IBT_CONTRACT] AUTHORIZATION [dbo] ([IBT_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[IBT_CONSUMER_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_IBT_CONSUMER] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [IBT_CONSUMER_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[IBT_CONSUMER_Q] ([IBT_CONTRACT])"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [IBTR_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE CONTRACT [IBTR_CONTRACT] AUTHORIZATION [dbo] ([IBTR_MSG] SENT BY INITIATOR,[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[IBTR_SOURCE_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_IBTR_RESPONSE] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [IBTR_SOURCE_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[IBTR_SOURCE_Q] ([IBTR_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::IBT_CONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::IBTR_SOURCE_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
     '--------------------------------------------------------
            
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
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUPS_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [CASHUPS_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [COLS_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [COLS_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOHALL_RQ_MSG] AUTHORIZATION [dbo] VALIDATION = NONE"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [SOHALL_RS_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"

            oTF.WriteToTextFile "CREATE MESSAGE TYPE [BUDGETDATA_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE MESSAGE TYPE [GENERAL_MSG] AUTHORIZATION [dbo] VALIDATION = WELL_FORMED_XML"
            oTF.WriteToTextFile "GO"
'--xxxxxxxxxxxxxxClearing
            oTF.WriteToTextFile "DROP SERVICE [INVOCATION_RS_SERVICE]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP QUEUE [dbo].[INVOCATION_RS_Q] "
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
                                                                                  & "[SALESSET_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[SALESSET_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[COLS_RQ_MSG] SENT BY INITIATOR, " _
                                                                                  & "[COLS_RS_MSG] SENT BY TARGET, " _
                                                                                  & "[BUDGETDATA_MSG] SENT BY INITIATOR, " _
                                                                                  & "[GENERAL_MSG] SENT BY INITIATOR, " _
                                                                                  & "[EndOfStream] SENT BY INITIATOR)"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE QUEUE [dbo].[INVOCATION_RS_Q] WITH STATUS = ON , RETENTION = OFF , ACTIVATION (  STATUS = ON , PROCEDURE_NAME = [dbo].[_actp_INVOCATION_RS_CONSUMER] , MAX_QUEUE_READERS = 1 , EXECUTE AS N'dbo'  ) "
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE SERVICE [INVOCATION_RS_SERVICE]  AUTHORIZATION [dbo]  ON QUEUE [dbo].[INVOCATION_RS_Q] ([INVOCATION_CONTRACT])"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::INVOCATION_RS_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::SOH_RS_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            
'-------------------
        If txtAddress > "" Then
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_MLC_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_MLC_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'MLC_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_LOYALTY_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_LOYALTY_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'LOYALTYSOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_SALES_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_SALES_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'SALESSOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_SOH_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_SOH_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'SOHSOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_SOH_RQ_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_SOH_RQ_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'SOH_RQ_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_ALERT_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_ALERT_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'ALERT_CONSUMER_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [[ROUTE_TO_ALERTLOAD_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_ALERTLOAD_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'ALERTLOAD_SOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_INVOCATION_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_INVOCATION_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'INVOCATION_RQ_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            'IBT and IBTR
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_IBT_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_IBT_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'IBT_SOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_IBTR_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_IBTR_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'IBTR_CONSUMER_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
        
            oTF.WriteToTextFile "DROP ROUTE [ROUTE_TO_CASHUPSOURCE_" & UCase(txtClientName) & "]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "CREATE ROUTE [ROUTE_TO_CASHUPSOURCE_" & UCase(txtClientName) & "] AUTHORIZATION [dbo] WITH  SERVICE_NAME  = N'CASHUPSOURCE_" & UCase(txtClientName) & "_SERVICE' ,  ADDRESS  = N'TCP://" & txtAddress & ":" & txtDestPort & "'"
            oTF.WriteToTextFile "GO"
        End If

        oTF.CloseTextFile
        Set oTF = Nothing
        txtNotes.Text = "End-point created: '" & strEndpointname & "'"
        txtNotes.Text = "Certificate created: '" & strCertificatename & "'"
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
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

Private Sub SaveSettings()
    SaveSetting "SB", "SB", "INSTANCENAME", Me.txtInstancename
    SaveSetting "SB", "SB", "DBNAME", Me.txtDBName
    SaveSetting "SB", "SB", "CLIENTNAME", Me.txtClientName
    SaveSetting "SB", "SB", "CLIENTADDRESS", Me.txtAddress
    SaveSetting "SB", "SB", "CENTRALorHUB", IIf(optCentral = True, "CENTRAL", "HUB")
    SaveSetting "SB", "SB", "ENDPOINTNAME", txtEndpointName

End Sub
'--------------------------------------------------------------------------


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
'
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
    If cnn.State = adStateClosed Then
      '  cnn.Close
        strMainConnectionString = "Provider=SQLOLEDB;Data Source=" & Me.txtInstancename & ";Initial Catalog=" & txtDBName & ";User Id=sa;Password=" & txtPassword & ";Connect Timeout=45"
        MsgBox strMainConnectionString
        cnn.Open strMainConnectionString
        cnn.CommandTimeout = 360
    Else
        If cnn.State <> adStateOpen Then
            OpenDB = 99
        End If
    End If
EXIT_HANDLER:
    Exit Function
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.OpenDB"
End Function
Public Sub CloseDB()
    If cnn Is Nothing Then Exit Sub
    If cnn.State = 1 Then   'it is open
        cnn.Close
        Set cnn = Nothing
    End If
End Sub
Private Function CheckBrokerEnabled() As Boolean
    On Error GoTo errHandler
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
    Exit Function
errHandler:
MsgBox "Pos 13"
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CheckBrokerEnabled"
End Function

Private Sub Command1_Click()
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
    
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        strPath = "C:\PBKS\BU\" & txtEndpointName & "_CERT.CER"
        If fs.FileExists(strPath) Then
            fs.DeleteFile strPath
        End If
        oTF.WriteToTextFile "BACKUP CERTIFICATE [" & txtEndpointName & "_CERT] TO FILE = '" & strPath & "'"
        oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
    End If
End Sub

Private Sub cmdPreparePublicCert_Click()
    SaveSettings
    If OpenDB() = 0 Then
        strCommandFilePath = "\\" & NameOfPC & "\PBKS_S\PrepareServiceBrokerScriptPub.SQL"
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        oTF.WriteToTextFile "USE [Master]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP LOGIN " & UCase(txtClientName)
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "DROP CERTIFICATE " & UCase(txtClientName) & "_CERT"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE CERTIFICATE [" & UCase(txtClientName) & "_CERT] FROM FILE = '" & Me.txtPublicKeyFile & "';"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "CREATE LOGIN [" & UCase(txtClientName) & "] FROM CERTIFICATE [" & UCase(txtClientName) & "_CERT]"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "GRANT CONNECT ON ENDPOINT::" & txtEndpointName & "_ENDPOINT TO [" & UCase(txtClientName) & "];"
        oTF.WriteToTextFile "GO"
        
        oTF.WriteToTextFile "USE [" & Me.txtDBName & "]"
        oTF.WriteToTextFile "GO"
        If optHUB Then
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::HUBCONSUMER_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
        Else
            oTF.WriteToTextFile "GRANT SEND ON SERVICE::MASTERLOYALTYSOURCE_SERVICE to [PUBLIC]"
            oTF.WriteToTextFile "GO"
            
        End If
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

End Sub

Private Sub endpoint_Click()

End Sub

Private Sub Form_Load()


        Set oTF = New z_TextFileSimple
    txtEndpointName = GetSetting("SB", "SB", "ENDPOINTNAME", "")
    Me.txtInstancename = GetSetting("SB", "SB", "INSTANCENAME", "")
    Me.txtDBName = GetSetting("SB", "SB", "DBNAME", "")
    txtClientName = GetSetting("SB", "SB", "CLIENTNAME", "")
    txtAddress = GetSetting("SB", "SB", "CLIENTADDRESS", "")
    optCentral = IIf(GetSetting("SB", "SB", "CENTRALorHUB", "") = "CENTRAL", 1, 0)
    optHUB = IIf(GetSetting("SB", "SB", "CENTRALorHUB", "") = "HUB", 1, 0)

       
End Sub

Private Sub optCentral_Click()
        Me.cmdPrepareCENTRAL.Enabled = True
        Me.cmdPrepareHUB.Enabled = False
End Sub

Private Sub optHUB_Click()
        Me.cmdPrepareCENTRAL.Enabled = False
        Me.cmdPrepareHUB.Enabled = True
End Sub

