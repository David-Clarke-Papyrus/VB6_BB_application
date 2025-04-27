VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{A9CD2883-061D-11D4-B62B-00004C937F50}#1.0#0"; "CoCash19.ocx"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr19.ocx"
Object = "{CCB90150-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSPOSPrinter.ocx"
Object = "{CCB90040-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSCashDrawer.ocx"
Begin VB.Form frmPOSMain 
   BackColor       =   &H00E1E1E1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DiscountSet"
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12210
   FillColor       =   &H00E1E1E1&
   Icon            =   "frmPOSMain7.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   814
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E1E1E1&
      Height          =   615
      Left            =   30
      TabIndex        =   20
      Top             =   2985
      Width           =   12015
      Begin VB.Label txtQtyTotal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   5625
         TabIndex        =   22
         Top             =   135
         Width           =   1425
      End
      Begin VB.Label txtExtTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   9570
         TabIndex        =   21
         Top             =   135
         Width           =   2340
      End
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2970
      Left            =   615
      OleObjectBlob   =   "frmPOSMain7.frx":038A
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   12000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   9705
      TabIndex        =   16
      Top             =   7080
      Width           =   1935
      Begin VB.Label lblSaleOnHold 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "** Sale on hold **"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   165
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblProg 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   60
         TabIndex        =   18
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label lblUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Updating"
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   60
         TabIndex        =   17
         Top             =   555
         Visible         =   0   'False
         Width           =   1860
      End
   End
   Begin VB.Timer ConnectionTimer 
      Interval        =   10000
      Left            =   5010
      Top             =   5100
   End
   Begin VB.TextBox txtDiscounts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2880
      Left            =   5850
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmPOSMain7.frx":5415
      Top             =   4245
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtVouchers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2820
      Left            =   5775
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmPOSMain7.frx":541B
      Top             =   4305
      Visible         =   0   'False
      Width           =   2595
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3780
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin TrueOleDBGrid60.TDBGrid G4 
      Height          =   2070
      Left            =   30
      OleObjectBlob   =   "frmPOSMain7.frx":5421
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   11985
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   900
      Left            =   75
      TabIndex        =   1
      Top             =   6225
      Width           =   5700
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Height          =   1380
      Left            =   30
      OleObjectBlob   =   "frmPOSMain7.frx":AE30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3705
      Width           =   4485
   End
   Begin TrueOleDBGrid60.TDBGrid G5 
      Height          =   2010
      Left            =   1950
      OleObjectBlob   =   "frmPOSMain7.frx":E4A3
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   7440
   End
   Begin TrueOleDBGrid60.TDBGrid G3 
      Height          =   2400
      Left            =   1500
      OleObjectBlob   =   "frmPOSMain7.frx":1270A
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   8745
   End
   Begin OposCashDrawer_1_8_LibCtl.OPOSCashDrawer OPOSCashDrawerPosiflex 
      Left            =   8880
      OleObjectBlob   =   "frmPOSMain7.frx":17C61
      Top             =   4815
   End
   Begin OposPOSPrinter_1_8_LibCtl.OPOSPOSPrinter OPOSPOSPrinterPosiflex 
      Left            =   8940
      OleObjectBlob   =   "frmPOSMain7.frx":17C85
      Top             =   4110
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   4695
      Top             =   4035
      _Version        =   65536
      _ExtentX        =   900
      _ExtentY        =   556
      _StockProps     =   0
   End
   Begin COCASHLib.OPOSCashDrawer OPOSCashDrawer1 
      Left            =   4770
      Top             =   4470
      _Version        =   65536
      _ExtentX        =   1244
      _ExtentY        =   767
      _StockProps     =   0
   End
   Begin VB.Label SB 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   315
      TabIndex        =   23
      Top             =   7185
      Width           =   9285
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   7380
      Width           =   360
   End
   Begin VB.Label lblState 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   8670
      TabIndex        =   10
      Top             =   6675
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Label lblChange 
      BackColor       =   &H00E1E1E1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   5820
      TabIndex        =   11
      Top             =   6240
      Width           =   5745
   End
   Begin VB.Label txtPaymentTotal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   30
      TabIndex        =   9
      Top             =   5070
      Width           =   5505
   End
   Begin VB.Label txtVatValue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   9900
      TabIndex        =   8
      Top             =   3075
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblCustomername 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   975
      Left            =   4530
      TabIndex        =   7
      Top             =   4290
      Width           =   6060
   End
   Begin VB.Label lblLoyaltyValue 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   9480
      TabIndex        =   6
      Top             =   6150
      Width           =   1965
   End
   Begin VB.Label lblReplacement 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5190
      Visible         =   0   'False
      Width           =   4920
   End
   Begin VB.Label lblInput 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   405
      Left            =   285
      TabIndex        =   2
      Top             =   5850
      Width           =   5400
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSwaptoTest 
         Caption         =   "Swap to TEST/LIVE database"
      End
      Begin VB.Menu mnuNewTestFromLive 
         Caption         =   "Create new test database from live database"
      End
      Begin VB.Menu mnuSep_01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRounding 
         Caption         =   "Rounding rules"
      End
      Begin VB.Menu mnuEAN 
         Caption         =   "Load EAN values"
      End
      Begin VB.Menu mnuSep_02 
         Caption         =   "-"
      End
      Begin VB.Menu mnusavecol 
         Caption         =   "Save column widths"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmPOSMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bUpdating As Boolean
Dim QI As MSMQQueueInfo
Dim QPOS As MSMQQueue
Dim QPOSACK As MSMQQueue
Dim QSVR As MSMQQueue
Dim POSmsg As MSMQMessage
Dim POSAckMsg As MSMQMessage
Dim SVRMsg As MSMQMessage
Private WithEvents POSEvent As MSMQEvent
Attribute POSEvent.VB_VarHelpID = -1
Private WithEvents SVREvent As MSMQEvent
Attribute SVREvent.VB_VarHelpID = -1
Private WithEvents POSACKEvent As MSMQEvent
Attribute POSACKEvent.VB_VarHelpID = -1
Dim arApproReturnLines() As ReturnRec
Dim arInvoiceLines() As InvoiceRec
Dim strMsg As String
Dim strOrderedTitle As String
Dim strCreditLimitExceededMessage As String
Public WithEvents qTimer As XTimer
Attribute qTimer.VB_VarHelpID = -1
Dim iColWidth As Integer
Dim bSaleOnHold As Boolean

Public Enum enOperatorType
    eOperator = 3
    eSupervisor = 4
End Enum

Private Enum enModeType
    emode_Appro = 4
    eMode_ApproReturn = 1
    emode_Sale = 2
    eMode_AcceptDeposit = 3
    eMode_ReturnDeposit = 5
    emode_PayAccount = 6
    emode_CreditNote = 7
End Enum
Dim enMode As enModeType

Private Enum enPaymentMode
    ePaymentMode_Cash = 1
    ePaymentMode_Cheque = 2
    ePaymentMode_CreditCard = 3
    ePaymentMode_Voucher = 4
    ePaymentMode_RedeemedDeposit = 5
    ePaymentMode_CreditVoucher = 6
    ePaymentMode_Account = 7
    ePaymentMode_DIrectDeposit = 8
End Enum


Private Enum enumDocumentType
    eTypReceipt = 1
    eTypVoucher = 2
    eTypCashRefund = 3
    etypCreditVoucher = 4
    eTypDeposit = 5
    eTypDepositRefund = 6
    eTypAppro = 7
    eTypPettyCash = 8
    eTypPettyCashCredit = 9
    eTypApproReturn = 10
    eTypOrder = 11
    eTypChangeVoucher = 12
    eTypPaymentReceipt = 13
    eTypCreditNote = 14
End Enum
Private Enum enumConnectionStatus
    eOnline = 0
    eConnectedOnly = 1
    eOffline = 2
    eError = 3
End Enum

Dim Stack() As Integer

Dim frmOpRep As frmPOSOPREP
'Dim frmStatus As frmPOSStatus
Dim frmPC As frmPettyCash
Dim frmPCC As frmPettyCashCredit
Dim frmH As frmHelp
Dim frmExchange As frmExchange
Dim frmDisc As frmDiscretionaryDiscount
Dim frmCustID As frmIDCustomer



Dim bItemExchange As Boolean
Dim bShiftDown As Boolean
Dim bValid As Boolean
Dim bEnvironmentOK As Boolean
Dim bLogonOK As Boolean
Dim bUnloading As Boolean
Dim bSaleActive As Boolean
Dim bCustomerVisible As Boolean
Dim bCloseXsession As Boolean
Dim bCloseZsession As Boolean
Dim bIssueCreditNote As Boolean
Dim lngDeposit As Long
Dim iCOLForDeposit As Long
Dim lngOPID As Long
Dim lngSalesItemCount As Long
Dim iToVoid As Long
Dim itmp As Integer
Dim lngCustomerID As Long
Dim iCurrentSaleLine As Integer
Dim iCurrentPaymentLine As Integer
Dim lngStaffID As Long
Dim lngRecordUpdateCount As Long
Dim lngPayable As Long
Dim lngVAT As Long
Dim mlngTotalDepositValue As Long

Dim enNewState As eState

Dim arLineNumber() As String
Dim arDiscounts() As String

Dim strPrefix As String
Dim strSuffix As String
Dim strValidVoucherTypes As String
Dim strValidDiscountTypes As String
Dim strDepositTitle As String
Dim strCustomer As String
Dim strDocCode As String
Dim strArg As String
Dim strArg2 As String
Dim strName As String
Dim strDepositMode As String

Dim strRaw As String

Dim ESC As String
Dim strCustomername As String

'For stored sale
Dim iCurrentSaleLine_store1 As Integer
Dim iCurrentPaymentLine_Store1 As Integer
Dim bIssueCreditNote_Store1 As Boolean
Dim oExchangeCopy1 As a_Exchange
Dim X1_Store1 As New XArrayDB
Dim X2_Store1 As New XArrayDB
Dim X3_Store1 As New XArrayDB
Dim X4_Store1 As New XArrayDB
Dim X5_Store1 As New XArrayDB

''''

Dim WithEvents oExchange As a_Exchange
Attribute oExchange.VB_VarHelpID = -1
Dim oPAYMENTLine As a_Payment
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2
Dim cCOLS As c_COLS
Dim cApps As c_APPs
Dim WithEvents oSALELine As a_Sale
Attribute oSALELine.VB_VarHelpID = -1
Dim ADOConn As ADODB.Connection

Dim enRequestState As eState
Dim enPresentState As eState
Dim enPreviousState As eState

Dim X1 As New XArrayDB
Dim X2 As New XArrayDB
Dim X3 As New XArrayDB
Dim X4 As New XArrayDB
Dim X5 As New XArrayDB

Private Type ITEMDATA
    TType As String
    Name As String
    Ext As String
    At As String
    Disc As String
    DiscDesc As String
    Alteration As Boolean
    Counterfoil As String
End Type

Public Function GetEnvironmentstatus() As Boolean
    On Error GoTo errHandler
    GetEnvironmentstatus = bEnvironmentOK And bLogonOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetEnvironmentstatus"
End Function
Private Sub ShowTransactions(bShow As Boolean)
    On Error GoTo errHandler
    If bShow Then
        G4.Visible = True
        G1.Visible = False
        Frame2.Visible = False
    Else
        G4.Visible = False
        G1.Visible = True
        Frame2.Visible = True
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ShowTransactions(bShow)", bShow
End Sub


Private Sub mnuEAN_Click()
    On Error GoTo errHandler
    oPC.OpenLocalDatabase
    oPC.DBLocalConn.Execute "EXEC LoadEAN"
    oPC.CloseLocalDatabase
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuEAN_Click"
End Sub



Private Sub mnuNewTestFromLive_Click()
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject
Dim strFilefolder As String
Dim strFileName As String
Dim tmp As String

    strFileName = oPC.LocalRootFolder & "\BU\PBKSFD_TEST.BAK"
    
    Set oBU = New z_PBKSBackup
    Screen.MousePointer = vbHourglass
    DoEvents
    oBU.BackupToBriefcase strFileName, True
            DoEvents
    
    Screen.MousePointer = vbDefault
    MsgBox "New test database has been created. You are still connected to the " & IIf(oPC.DatabaseName = "PBKSFD_TEST", "TEST", "LIVE") & " database", vbOKOnly, "Status"

End Sub

Private Sub mnuRounding_Click()
Dim F As New frmRoundingRules
    F.Show
End Sub

Private Sub mnuSwaptoTest_Click()

    If MsgBox("You want to open this application connected to the " & IIf(oPC.DatabaseName = "PBKSFD", "TEST", "LIVE") & " database?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    Else
        SaveSetting "POS", "StartDatabaseName", "DBNAME", IIf(oPC.DatabaseName = "PBKSFD", "PBKSFD_TEST", "PBKSFD")
    End If
    MsgBox "The application will close. Reopening it will have it connected to the " & IIf(oPC.DatabaseName = "PBKSFD", "TEST", "LIVE") & " database."
    Unload Me
End Sub

Private Sub oExchange_CreditLimitExceeded(Excess As String)
    If Excess > "" Then
        strCreditLimitExceededMessage = "Credit Limit exceeded by : " & Excess
    Else
        strCreditLimitExceededMessage = ""
    End If
    DisplayCustomerDetails
End Sub

Private Sub oSaleLine_ProvisionalPrice()
    txtExtTotal.ForeColor = RGB(41, 133, 46)
End Sub




Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
        If (X1(Bookmark, 3) < 1) And (X1(Bookmark, 3) <> "") Then
            RowStyle.BackColor = 65135
        Else
            RowStyle.BackColor = &HFFFFFF
        End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblPrompt_Click()
Dim strMsg As String
    strMsg = "Additional options are: " & vbCrLf & vbCrLf _
    & "REPRINT - lets you review exchanges and reprint if necessary." & vbCrLf _
    & "OD - open the drawer." & vbCrLf & vbCrLf _
    & "Please note that the 'n' in Dn, DPn, Vn etc refers to an exchange number, " & vbCrLf _
    & "         so to void an exchange number 312 for example, use 'V312'"
    MsgBox strMsg, vbInformation + vbOKOnly, "Hints"
End Sub

Private Sub mnusavecol_Click()
    On Error GoTo errHandler
    SaveLayout Me.G4, Me.Name & "B"
    SaveLayout Me.G1, Me.Name & "A"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnusavecol_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub oPS_ConnectionStatus(iStatus As Integer)
    On Error GoTo errHandler
    Select Case iStatus
    Case eOnline
      '  lblOnlineStatus = "Online"
      '  lblOnlineStatus.ForeColor = &H8000000D
    Case eConnectedOnly
     '   lblOnlineStatus = "Server off"
     '   lblOnlineStatus.ForeColor = vbRed
    Case eOffline
    '    lblOnlineStatus = "No network"
     '   lblOnlineStatus.ForeColor = vbRed
    Case eError
        MsgBox "The transmission of data to the server has been interrupted or the updating of local data has failed." & vbCrLf & "Finish this transaction, then restart this application. " & vbCrLf & "PLEASE INFORM PAPYRUS SUPPORT"
     '   lblOnlineStatus = "Error transmitting"
     '   lblOnlineStatus.ForeColor = vbRed
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_ConnectionStatus(iStatus)", iStatus, EA_NORERAISE
    HandleError
End Sub

Private Sub UpdatingLocalDatabase(bOn As Boolean, lngCnt As Long)
    On Error GoTo errHandler
Static strMsg As String
    If bUnloading Then Exit Sub
    If bOn Then
        strMsg = SB.Caption
        lngRecordUpdateCount = lngCnt
        lblUpdate.Caption = "updating (" & CStr(lngCnt) & ")"
        lblUpdate.Visible = True
        Me.Refresh
     '   SB.caption = "Updating local database . . . (" & CStr(lngCnt) & " records)"
    Else
       ' SB.caption = strMsg
        lblUpdate.Visible = False
        Me.Refresh
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.UpdatingLocalDatabase(bOn,lngCnt)", Array(bOn, lngCnt)
End Sub
Private Sub Counter(msg As String, lngCnt As Long)
    On Error GoTo errHandler
Static strMsg As String
    If bUnloading Then Exit Sub
    lblUpdate.Caption = "updating " & msg & "(" & CStr(lngCnt) & ")"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Counter(lngCnt)", lngCnt
End Sub

Private Sub SetPresentState(val As eState)
    On Error GoTo errHandler
    If val = eEND Then
        Unload Me
        Exit Sub
    End If
    enPresentState = val
    Me.lblState.Caption = InterpretState
    PrepareForm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetPresentState(val)", val
End Sub
Private Function InterpretState() As String
    On Error GoTo errHandler
    Select Case enPresentState
    Case 0
        InterpretState = "eStart"
    Case 1
        InterpretState = "eSale"
    Case 2
        InterpretState = "eTitle"
    Case 3
        InterpretState = "eQty"
    Case 4
        InterpretState = "eDiscount"
    Case 5
        InterpretState = "ePrice"
    Case 6
        InterpretState = "elogin"
    Case 7
        InterpretState = "ePaymentAmt"
    Case 8
        InterpretState = "eConfirmation"
    Case 9
        InterpretState = "eSearchCustomer"
    Case 20
        InterpretState = "eXTerminate"
    Case 21
        InterpretState = "eZTerminate"
    Case 22
        InterpretState = "eRebuildIndexes"
    Case 23
        InterpretState = "eHelp"
    Case 24
        InterpretState = "ecancelsale"
    Case 25
        InterpretState = "eCashRefund"
    Case 26
        InterpretState = "ePriceCashRefund"
    Case 27
        InterpretState = "eQtyCashRefund"
    Case 28
        InterpretState = "eDiscountCashRefund"
    Case 29
        InterpretState = "eConfirmationCashrefund"
    Case 30
        InterpretState = "eVoid"
    Case 31
        InterpretState = "eReviewExchanges"
    Case 32
        InterpretState = "eShowExchange"
    Case 33
        InterpretState = "eOPenDrawer"
    Case 34
        InterpretState = "eStatus"
    Case 35
        InterpretState = "enull"
    Case 36
        InterpretState = "ePrevious"
    Case 37
        InterpretState = "eDelete"
    Case 38
        InterpretState = "eDeletePayment"
    Case 39
        InterpretState = "eShowvoucherType"
    Case 40
        InterpretState = "eOperatorsReport"
    Case 41
        InterpretState = "eCreditNote"
    Case 42
        InterpretState = "ePriceCreditNote"
    Case 43
        InterpretState = "eDiscountCreditNote"
    Case 44
        InterpretState = "eQtyCreditNote"
    Case 45
        InterpretState = "eRefundDeposit"
    Case 46
        InterpretState = "eConfirmationRefundDeposit"
    Case 47
        InterpretState = "eSearchCustomerfordepositRefund"
    Case 48
        InterpretState = "eRefundType_Cash"
    Case 49
        InterpretState = "eRefundType_Creditcard"
    Case 50
        InterpretState = "eSearchCustomerforAppro"
    Case 51
        InterpretState = "eAppro"
    Case 52
        InterpretState = "ePriceAppro"
    Case 53
        InterpretState = "eDiscountAppro"
    Case 54
        InterpretState = "eQtyAppro"
    Case 55
        InterpretState = "eConfirmationAppro"
    Case 56
        InterpretState = "eApproReturn"
    Case 57
        InterpretState = "eSearchCustomerforApproReturn"
    Case 58
        InterpretState = "ePettyCash"
    Case 59
        InterpretState = "ePettyCashAmt"
    Case 60
        InterpretState = "ePettyCashConfirmation"
    Case 61
        InterpretState = "ePettyCashReason"
    Case 62
        InterpretState = "ePettyCashCredit"
    Case 63
        InterpretState = "ePettyCashCreditConfirmation"
    Case 64
        InterpretState = "ePettyCashCreditAmt"
    Case 65
        InterpretState = "eSearchCustomerfordeposit"
    Case 66
        InterpretState = "eDiscountDeposit"
    Case 67
        InterpretState = "eSelectDepositLineRef"
    Case 68
        InterpretState = "eSelectDepositLine"
    Case 69
        InterpretState = "eSelectDepositLineForRefund"
    Case 70
        InterpretState = "ePriceDeposit"
    Case 71
        InterpretState = "eQtyDeposit"
    Case 72
        InterpretState = "eInvoice"
    Case 73
        InterpretState = "eInvoiceno"
    Case 74
        InterpretState = "eInvoiceMode"
    Case 75
        InterpretState = "eConfirmationInvoiceCollection"
    Case 76
        InterpretState = "eConfirmationDepositRefund"
    Case 77
        InterpretState = "eConfirmationDeposit"
    Case 78
        InterpretState = "eConfirmationCreditNote"
    Case 89
        InterpretState = "eCollect"
    Case 90
        InterpretState = "ePaymentType_Cash"
    Case 91
        InterpretState = "ePaymentType_Cheque"
    Case 92
        InterpretState = "ePaymentType_CreditCard"
    Case 93
        InterpretState = "ePaymentType_CreditVoucher"
    Case 94
        InterpretState = "ePaymentType_CreditVoucherRef"
    Case 95
        InterpretState = "ePaymentType_voucher"
    Case 96
        InterpretState = "ePaymentType_ChequeRef"
    Case 97
        InterpretState = "ePaymentType_CreditVoucherRef"
    Case 98
        InterpretState = "ePaymentType_voucherRef"
    Case 99
        InterpretState = "ePaymentType_RedeemDeposit"
    Case 100
        InterpretState = "eRefundType_CreditVoucher"
    End Select

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.InterpretState"
End Function

Private Sub oExchange_ContainsLines(pYesNo As Boolean)
    On Error GoTo errHandler
    If bUnloading Then Exit Sub
    bSaleActive = pYesNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_ContainsLines(pYesNo)", pYesNo, EA_NORERAISE
    HandleError
End Sub
Private Sub SetTitleBar(pShowExchangeNumber As Boolean)
    On Error GoTo errHandler
    Caption = "Papyrus Point-of-Sale       " & oPC.StationName & "      Supervisor: " & oPC.ZSession.SupervisorName & "/" & oPC.ZSession.OpSession.Name & IIf(pShowExchangeNumber = True, "              #" & oExchange.ExchangeNumber, "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetTitleBar(pShowExchangeNumber)", pShowExchangeNumber
End Sub

Sub POSACKEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim lngResult As Integer
Dim oSM As New z_SM
Dim strEXCHID As String

    Set QPOSACK = Queue
    Set POSAckMsg = QPOSACK.Receive
    If Not (POSAckMsg Is Nothing) Then
        If lblProg.Caption > "" Then
            lblProg.Caption = Left(lblProg.Caption, Len(lblProg.Caption) - 1)
            strEXCHID = FNS(POSAckMsg.Body)
            oSM.UpdateExchAck strEXCHID
        End If
    End If
    On Error Resume Next
    QPOSACK.EnableNotification POSACKEvent
    On Error GoTo errHandler
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.POSACKEvent_Arrived(Queue,Cursor)", Array(Queue, Cursor), EA_NORERAISE
    HandleError
End Sub

Private Sub qTimer_Tick()
On Error Resume Next
    QSVR.EnableNotification SVREvent
    QPOSACK.EnableNotification POSACKEvent
End Sub

Sub SVREvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim lngResult As Integer
    
    Set QSVR = Queue
    Set SVRMsg = QSVR.Receive
    If Not (SVRMsg Is Nothing) Then
        Screen.MousePointer = vbHourglass
        Me.Refresh
        If Left(SVRMsg.Label, 5) = "Clear" Then
            UpdateClientFromServerFiles , SVRMsg.Label
        Else
            UpdateClientFromServerFiles SVRMsg.Body, ""
        End If
        DoEvents
        Screen.MousePointer = vbDefault
    End If
    On Error Resume Next
    QSVR.EnableNotification SVREvent
    On Error GoTo errHandler
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SVREvent_Arrived(Queue,Cursor)", Array(Queue, Cursor), EA_NORERAISE
    HandleError
End Sub

Private Sub SetupQueues()
    On Error GoTo errHandler
   
   'Set up receiving queue (always local queue)
    Set QI = New MSMQQueueInfo
    QI.PathName = oPC.NameOfPC & "\Private$\qposack"
    On Error Resume Next
    QI.Create , True
    On Error GoTo errHandler
    Err.Clear
    Set QPOSACK = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set POSACKEvent = New MSMQEvent
    On Error Resume Next
    QPOSACK.EnableNotification POSACKEvent
    On Error GoTo errHandler
    'Set up our SVR queue for receiving notifications about DB changes
    Set QI = Nothing
    Set QI = New MSMQQueueInfo
    QI.PathName = oPC.NameOfPC & "\Private$\qsvr"
    On Error Resume Next
    QI.Create , True
    Err.Clear
    On Error GoTo errHandler
    Set QSVR = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set SVREvent = New MSMQEvent
    On Error Resume Next
    QSVR.EnableNotification Event:=SVREvent
    On Error GoTo errHandler

Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetupQueues"
End Sub
Public Sub OpenQSVR()
    If QSVR Is Nothing Then
        Set QSVR = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
        Set SVREvent = New MSMQEvent
        SetQSVRTrigger
    End If
    If Not QSVR.IsOpen2 Then
        SetQSVRTrigger
    End If
End Sub
Public Sub CloseQSVR()
    If Not QSVR Is Nothing Then
        If QSVR.IsOpen2 Then
            QSVR.Close
            Set QSVR = Nothing
          '  RaiseEvent QSVRTriggerStatus(False)
        End If
    Else
           ' RaiseEvent QSVRTriggerStatus(False)
    End If
End Sub

Private Sub SetQSVRTrigger()
    On Error Resume Next
    QSVR.EnableNotification SVREvent
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim Result As Integer
Dim bLoggedOnAlready As Boolean
Dim strPos As String
Dim strDBName As String
Dim bGetFloat As Boolean
Dim dblFLoat As Double
Dim sFloatBreakdown As String

Dim i As Integer
        'SaveSetting "POS", "StartDatabaseName", "DBNAME", IIf(oPC.DatabaseName = "PBKSFD", "PBKSFD_TEST", "PBKSFD")
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "A", CStr(i), G1.Columns(i - 1).Width)
    Next
    For i = 1 To G4.Columns.Count
        G4.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "B", CStr(i), G4.Columns(i - 1).Width)
    Next

    strPos = "1"
        Frame2.TOP = 202
        Frame2.Left = 2
    
        G1.TOP = 4
        G1.Left = 2
    
    bUpdating = False
    bEnvironmentOK = True
    ESC = Chr(27)
    iToVoid = 0
    'Try to load local DB connection
    If oPC Is Nothing Then
        Set oPC = New z_POSCLIConnection
        If UBound(arCommandLine) >= 1 Then
            oPC.DatabaseName = arCommandLine(0)
        Else
            strDBName = GetSetting("POS", "StartDatabaseName", "DBNAME", "PBKSFD")
            If strDBName = "" Then
                oPC.DatabaseName = "PBKSFD"
                oPC.UseTestDatabase = False
            Else
                oPC.UseTestDatabase = False
                oPC.DatabaseName = strDBName
            End If
        End If
    strPos = "2"
'        If UBound(arCommandLine) > 1 Then
'            oPC.UseTestDatabase = True
'        Else
'            oPC.UseTestDatabase = False
'        End If
        mnuSwaptoTest.Caption = IIf(oPC.DatabaseName = "PBKSFD", "Swap to TEST database", "Swap to LIVE database")
        oPC.InitializeSettings
        oPC.dbConnect
        oPC.LoadProperties
        oPC.loadRoundingRules
        oPC.loadMultibuys
        If oPC.ServerIPAddress <> oPC.NameOfPC Then
            SynchronizeTOD oPC.ServerIPAddress
        End If
    End If
    strPos = "3"
    bLoggedOnAlready = False
    bLogonOK = True
    oPC.SetupZSession lngStaffID, strName
'    MsgBox "oPC.ZSession.SupervisorID=" & oPC.ZSession.SupervisorID & " PC.GetProperty(CaptureFloat): " & oPC.GetProperty("CaptureFloat")
    If oPC.ZSession.SupervisorID = 0 Then
        If oPC.GetProperty("CaptureFloat") = "TRUE" Then
            bGetFloat = True
        Else
            bGetFloat = False
        End If
        LogonOperator bGetFloat, dblFLoat, sFloatBreakdown
        If bLogonOK = False Then
            bCloseZsession = True
            GoTo EXITHANDLER
        End If
        oPC.ZSession.SupervisorID = lngStaffID
        oPC.ZSession.SupervisorName = strName
        bLoggedOnAlready = True
    End If
    strPos = "4"
    If oPC.ZSession.LoadOpenXSession = False Then
        oPC.ZSession.OpSession.Start_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
        If oPC.ZSession.OpSession.SupervisorID = 0 Then
            If bLoggedOnAlready = False Then
                If oPC.GetProperty("CaptureFloat") = "TRUE" Then
                    bGetFloat = True
                Else
                    bGetFloat = False
                End If
                LogonOperator bGetFloat, dblFLoat, sFloatBreakdown
                If bLogonOK = False Then
                    bCloseXsession = True
                    bCloseZsession = True
                    GoTo EXITHANDLER
                End If
            End If
            oPC.OpenLocalDatabase
            oPC.ZSession.OpSession.SetOperatorID lngStaffID, dblFLoat, sFloatBreakdown
           ' oPC.ZSession.OpSession.OperatorID = lngStaffID
            oPC.ZSession.OpSession.Name = strName
            oPC.CloseLocalDatabase
        End If
    End If
    strPos = "5"
    SetForCOLSVisible False
    If oPC.UseCashDrawer = "TRUE" Then
        If oPC.DriveDrawer = True Then  'There is a COM connected Cash Drawer
            MSComm1.Settings = oPC.COMPORTSettings
            MSComm1.CommPort = oPC.CashDrawerPort
            If MSComm1.PortOpen = False Then
                MSComm1.PortOpen = True
            End If
        Else                            'There is a cash drawer connected to the Printer
            If oPC.GetProperty("ReceiptPrinterType") = "" Or oPC.GetProperty("ReceiptPrinterType") = "Epson" Then
                OPOSCashDrawer1.DeviceEnabled = True
            Else
                OPOSCashDrawerPosiflex.DeviceEnabled = True
            End If
        End If
    End If
    strPos = "6"
    If oPC.PrintSlips = "TRUE" Then
        strPos = "6.1"
        If oPC.UseA4Printer <> "TRUE" Then
        strPos = "6.2"
           SetupPrinter
        strPos = "6.3"
            If oPC.UseCashDrawer = True Then
        strPos = "6.4"
                SetupCashDrawer
            End If
        Else
        strPos = "6.3"
            Printer.FontName = "COURIER"
            Printer.FontSize = 12
            iColWidth = 50
        End If
    End If
    strPos = "7"
    If Not bEnvironmentOK Then
        GoTo EXITHANDLER
    End If
    LoadVoucherTypes
    LoadDiscountTypes
    txtInput.BackColor = RGB(230, 250, 210)
    
    G1.Array = X1
    G4.Height = 380
    If oPC.DatabaseName <> "PBKSFD_TEST" Then
        ReSendExchanges
    End If
    strPos = "8"
    X4.Clear
    X4.ReDim 1, 0, 1, 13
    LoadExchanges
    SetupQueues
    
    Set qTimer = New XTimer
    qTimer.Interval = 10000
    qTimer.Enabled = True
    
    If oPC.DatabaseName = "PBKSFD_TEST" Then
        Me.BackColor = vbRed
    End If
    
EXITHANDLER:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Load", , EA_NORERAISE, , "strpos", Array(strPos)
    HandleError
End Sub


Private Sub SetupPrinter()
Dim lngResult As Long

        MsgBox "in SetupPrinter"

    If oPC.PrintSlips = True Then
            If oPC.GetProperty("ReceiptPrinterType") = "" Or oPC.GetProperty("ReceiptPrinterType") = "Epson" Then
        MsgBox "Trying to set up Epson Printer name is " & oPC.Printername
                With OPOSPOSPrinter1
                    lngResult = .Open(oPC.Printername)
                    
                    If lngResult = 0 Then
                        lngResult = .ClaimDevice(5)
                        If lngResult = OPOS_SUCCESS Then
                            .ClaimDevice 1000
                            .DeviceEnabled = True
                            .MapMode = PTR_MM_METRIC
                            .RecLetterQuality = True
                            .RecLineChars = 40
                        Else
                            MsgBox "The till printer (" & oPC.Printername & ") is not online. This application will close."
                            bEnvironmentOK = False
                            Exit Sub
                        End If
                    Else
                        MsgBox "The till printer is not online. This application will close."
                        bEnvironmentOK = False
                        Exit Sub
                    End If
                End With
        ElseIf oPC.GetProperty("ReceiptPrinterType") = "Posiflex" Then
        MsgBox "Trying to set up POSIFLEX Printer name is " & oPC.Printername
                With OPOSPOSPrinterPosiflex
                    lngResult = .Open(oPC.Printername)
                    
                    If lngResult = 0 Then
                    MsgBox "Before claim"
                        lngResult = .ClaimDevice(5)
                        MsgBox "after claim: result " & CStr(lngResult)
                        If lngResult = OPOS_SUCCESS Then
                            .ClaimDevice 1000
                            .DeviceEnabled = True
                            .MapMode = PTR_MM_METRIC
                            .RecLetterQuality = True
                            .RecLineChars = 40
                        Else
                            MsgBox "The till printer (" & oPC.Printername & ") is not online. This application will close."
                            bEnvironmentOK = False
                            Exit Sub
                        End If
                    Else
                        MsgBox "The till printer is not online. This application will close."
                        bEnvironmentOK = False
                        Exit Sub
                    End If
                End With
        End If
    End If
    Me.lblState.Visible = False

End Sub
Private Sub SetupCashDrawer()
Dim lngResult As Long

            
    If oPC.DriveDrawer = False Then
        If oPC.GetProperty("ReceiptPrinterType") = "" Or oPC.GetProperty("ReceiptPrinterType") = "Epson" Then
            With OPOSCashDrawer1
                lngResult = .Open(oPC.TillDrawerName)
                If lngResult = 0 Then
                    .DeviceEnabled = True
                    lngResult = .ClaimDevice(1000)
                Else
                    MsgBox "The till drawer is not available. This application will close."
                    bEnvironmentOK = False
                    Exit Sub
                End If
            End With
        ElseIf oPC.GetProperty("ReceiptPrinterType") = "Posiflex" Then
            With OPOSCashDrawerPosiflex
                lngResult = .Open(oPC.TillDrawerName)
                If lngResult = 0 Then
                    .DeviceEnabled = True
                    lngResult = .ClaimDevice(1000)
                Else
                    MsgBox "The till drawer is not available. This application will close."
                    bEnvironmentOK = False
                    Exit Sub
                End If
            End With
        End If
    End If
    
    Me.lblState.Visible = False

End Sub

Private Sub LoadVoucherTypes()
    On Error GoTo errHandler
Dim ar() As String
Dim i As Integer
    ar = Split(oPC.VoucherSet, ";")
    
    strValidVoucherTypes = ""
    For i = 0 To UBound(ar)
        strValidVoucherTypes = strValidVoucherTypes & Left(ar(i), 1)
    Next
    
    txtVouchers = Replace(oPC.VoucherSet, ";", vbCrLf)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadVoucherTypes"
End Sub
Private Sub LoadDiscountTypes()
    On Error GoTo errHandler
Dim i As Integer
    arDiscounts = Split(oPC.DiscountSet, ";")
    
    strValidDiscountTypes = ""
    For i = 0 To UBound(arDiscounts)
        strValidDiscountTypes = strValidDiscountTypes & Left(arDiscounts(i), 1)
    Next
    
    txtDiscounts = Replace(oPC.DiscountSet, ";", vbCrLf)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadDiscountTypes"
End Sub

Private Function GetDiscount(pCODE As String, pDescription) As Integer
    On Error GoTo errHandler
Dim i As Integer
Dim str As String
Dim k As Integer
Dim iDisc As Integer
Dim str2 As String
    iDisc = 0

    For i = 0 To UBound(arDiscounts)
        str = arDiscounts(i)
        If pCODE = Left(str, 1) Then
            If InStr(str, "(") > 0 Then
                k = InStr(1, str, "(")
                iDisc = CInt(MID(Left(str, InStr(1, str, "%") - 1), k + 1))
                str2 = Left(str, k - 1)
                pDescription = Right(str2, Len(str2) - 2)
            Else
                pDescription = Right(str, Len(str) - 2)
                iDisc = 0
            End If
            Exit For
        End If
    Next
    GetDiscount = iDisc
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetDiscount(pCODE,pDescription)", Array(pCODE, pDescription)
End Function
Private Function LogonOperator(Optional bAskForFloat As Boolean, Optional dblFLoat As Double, Optional sFloatBreakdown) As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean
Dim Res As Boolean
Dim F As frmGetFloat

    If bAskForFloat Then
        Set F = New frmGetFloat
        F.Show vbModal
        If F.IsCancelled Then
            bLogonOK = False
            Unload F
            Exit Function
        Else
            dblFLoat = F.FloatValue
            sFloatBreakdown = F.GetFloatBreakdown
            Unload F
        End If
    End If
    
    Res = False
    Do Until Res = True
        If Not SecurityControl(eOperator, lngStaffID, strName, bCancelled, "Enter your signature.", "Your signature is invalid", True) Then
            If bCancelled Then Res = True
            bLogonOK = False
        Else
            Res = True
            bLogonOK = True
        End If
    Loop
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LogonOperator"
End Function

Private Function SwapOperator() As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean

    If oPC.ZSession.OpSession.InSession Then
        oPC.ZSession.OpSession.Close_OP_Session
    End If
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid", True) Then
        oPC.ZSession.OpSession.Start_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
    Else
        SetPresentState elogin
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SwapOperator"
End Function


Public Sub StartSale()
    On Error GoTo errHandler
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SetExchangeType eSaleType
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    bIssueCreditNote = False
    SetTitleBar False
    enPresentState = eStart
    enMode = emode_Sale
    PrepareForm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StartSale"
End Sub



Private Sub Stat(msg As String)
    On Error GoTo errHandler
    SB.Caption = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Stat(msg)", msg
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
 '   If bEnvironmentOK = False Then
 '   End If
    If Not bCloseXsession And Not bCloseZsession And Not bLogonOK = False And bEnvironmentOK = True Then
        If MsgBox("You want to close this application? Confirm", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
    If bEnvironmentOK = True Then
        bUnloading = True
        ConnectionTimer.Enabled = False
        CloseApplication Cancel
        If Cancel Then
            bUnloading = False
            ConnectionTimer.Enabled = False
            Exit Sub
        End If
    End If
    
    oPC.dbCloseLocalConnect
    Set oPC = Nothing
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
    End If
    Set QI = Nothing
    Set QPOS = Nothing
    Set QPOSACK = Nothing
    Set QSVR = Nothing
    Set POSmsg = Nothing
    Set POSAckMsg = Nothing
    Set SVRMsg = Nothing
    If Not qTimer Is Nothing Then
        qTimer.Enabled = False
        Set qTimer = Nothing
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub CloseApplication(bCancel As Integer)
    On Error GoTo errHandler
    If bSaleActive Then
        If MsgBox("There is still a transaction in process!" & vbLf & _
                  "Do you want to close this application anyway?", _
                  vbYesNo, "Transaction In Process!") = vbNo Then
            bCancel = True
            Exit Sub
        Else
            RejectSale
        End If
    End If
    Screen.MousePointer = vbHourglass
    Me.SB.Caption = "Wait. The local data is being transmitted to the server."
    
    If Not oExchange Is Nothing Then
        If oExchange.IsEditing Then oExchange.CancelEdit
    End If
    
    If bCloseXsession Then
        If Not oPC.ZSession.OpSession Is Nothing Then
            oPC.ZSession.OpSession.Close_OP_Session
        End If
    End If
    
    If bCloseZsession Then
        If Not oPC.ZSession Is Nothing Then
            oPC.ZSession.Close_Z_Session
        End If
    End If
    
    If Not OPOSPOSPrinter1 Is Nothing Then
    With OPOSPOSPrinter1
        .DeviceEnabled = False
        .ReleaseDevice
        .Close
    End With
    End If
    
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CloseApplication(bCancel)", bCancel
End Sub


Private Sub mnuClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ShowExchange()
    On Error GoTo errHandler
Dim lngRow As Long
Dim lngTmp As Long
Dim oTmpExchange As a_Exchange
Dim strPos As String

    Set frmExchange = New frmExchange
strPos = "01"
    If IsNumeric(strSuffix) And Len(strSuffix) < 10 Then
        lngRow = CLng(strSuffix)
strPos = "02"
        If lngRow <= X4(1, 1) And lngRow > 0 Then 'X4(X4.UpperBound(1) - 1, 1) And lngRow > 0 Then
strPos = "03"
            lngTmp = X4.Find(1, 1, lngRow, , , XTYPE_LONG)
strPos = "04"
            If lngTmp > 0 Then
                frmExchange.component X4(lngTmp, 10)
strPos = "05"
                frmExchange.Show vbModal
                If frmExchange.MustPrint = True Then
                    Set oTmpExchange = oExchange
strPos = "06"
                    Set oExchange = New a_Exchange
strPos = "07"
                    oExchange.Load (X4(lngTmp, 10)), True
strPos = "08"
                    If oExchange.TransactionTypeEnum = eOrderRequestType Then
                        PrintORDERSlip oPC.GetProperty("DepositCopyCount"), True
                    ElseIf oExchange.TransactionTypeEnum = ePettyCashType Then
                        PrintPettyCashVoucher oPC.GetProperty("PettyCashCopyCount")
                    ElseIf oExchange.TransactionTypeEnum = ePettyCashCreditType Then
                        PrintPettyCashVoucher oPC.GetProperty("PettyCashCopyCount")
                    Else
                        PrintSalesSlip oPC.GetProperty("InvoiceCopyCount"), True
                    End If
                    Set oExchange = Nothing
                    Set oExchange = oTmpExchange
                    Set oTmpExchange = Nothing
                End If
                Unload frmExchange
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ShowExchange", , , , "strSuffix", Array(strSuffix, X4(X4.UpperBound(1) - 1, 1), strPos)
End Sub

Private Sub PrepareForm()
    On Error GoTo errHandler
    txtVouchers.Visible = False
    txtDiscounts.Visible = False
'    txtPettyCash.Visible = False
    Select Case enPresentState
        Case ePettyCashCredit
            setInputBox "", "", "", False
            lblInput.Caption = "Pett cash return amount"
            Stat " .. to reverse"
        Case ePettyCashCreditAmt
            setInputBox "", "", "", False
            lblInput.Caption = "Petty cash amount returned"
            Stat " .. to reverse"
        Case ePettyCashCreditConfirmation
            setInputBox "OK", "*", "", True
            lblInput.Caption = "Confirm petty cash return"
            Stat " .. to reverse"
        Case ePettyCash
            setInputBox "", "", "", False
            lblInput.Caption = "Select petty cash account"
            Stat " .. to reverse"
        Case ePettyCashAmt
            setInputBox "", "", "", False
            lblInput.Caption = "Petty cash amount"
            Stat " .. to reverse"
        Case ePettyCashReason
            setInputBox "", "", "", False
            lblInput.Caption = "Reason"
            Stat " .. to reverse"
        Case ePettyCashConfirmation
            setInputBox "OK", "*", "", True
            lblInput.Caption = "Confirm petty cash withdrawal"
            Stat " .. to reverse"
        Case eSelectDepositLine, eSelectDepositLineForRefund
            setInputBox "", "", "", True
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            lblInput.Caption = "Select order line number from list "
            Stat "'.. to reverse"
        Case eRefundDeposit
            setInputBox "", "", "", True
            lblInput.ForeColor = vbBlue
            txtInput.ForeColor = vbBlue
            lblInput.Caption = "Select refund type "
            Stat "(X) Cancel,(CV)Credit voucher,(C)Cash"
        Case eInvoiceno
            setInputBox "", "", "", False
            lblInput.Caption = "Select line number of invoice to pay "
            Stat " .. to reverse"
        Case eCollect, eInvoiceMode
            G5.Visible = False
            setInputBox "", "", "", False
            If oExchange.transactionType = "RDEP" Or oExchange.TotalPayable < 0 Then
                lblInput.Caption = "Select refund type "
                Stat " .. to reverse,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account"
            ElseIf oExchange.transactionType = "AR" Then
                lblInput.Caption = "Select payment type "
                Stat " (X) to Cancel,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account,(DDP) Direct deposit"
            ElseIf oExchange.transactionType = "PA" Then
                lblInput.Caption = "Select payment type "
                Stat " (X) to Cancel,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(DDP) Direct deposit"
            Else
                lblInput.Caption = "Select payment type "
                Stat " .. to reverse,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account,(DDP) Direct deposit"
            End If
            SetForCOLSVisible False
        Case eShowvoucherType
            lblInput.Caption = "Select voucher type "
            txtVouchers = Replace(oPC.VoucherSet, ";", vbCrLf)
            txtVouchers.Visible = True
            Stat "  .. to reverse"
        Case ecancelsale
            setInputBox "", "", "", True
        Case eAppro
          '  ClearTextFields
            setInputBox "", "", "", False
            lblInput.ForeColor = vbBlue
            txtInput.ForeColor = vbBlue
            If bSaleActive Then
                lblInput.Caption = "Scan code or action."
                Stat "Scan or (F)Finalize,(X)Cancel appro,(Dn)Del prod,(DPn)Del paymt"
            Else
                lblInput.Caption = "Start Appro "
                Stat "Start Appro by entering product code,  (X) to cancel appro"
            End If
        Case eApproReturn
            ClearTextFields
            setInputBox "", "", "", False
            lblInput.ForeColor = vbBlue
            txtInput.ForeColor = vbBlue
            If bSaleActive Then
                lblInput.Caption = "Product code."
                Stat "Scan or (F)Finalize,(X)Cancel appro return,(Dn)Del prod,(DPn)Del paymt"
            Else
                lblInput.Caption = "Start Appro return"
                Stat "Start Appro return by entering product code of books returned,   (X) to cancel appro return"
            End If
        Case eConfirmation
            Stat "'.. to reverse"
            Select Case oExchange.transactionType
            Case "RDEP"
                lblInput.Caption = "Confirm deposit refund"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "S"
                lblInput.Caption = "Confirm sale"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "APP"
                lblInput.Caption = "Confirm appro"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "AR"
                lblInput.Caption = "Confirm appro return payment"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "DEP"
                lblInput.Caption = "Confirm deposit payment"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "OR"
                lblInput.Caption = "Confirm deposit payment"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "PA"
                lblInput.Caption = "Confirm account payment"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            Case "CN"
                lblInput.Caption = "Confirm issue credit note"
                setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
            End Select
        Case eDiscount                       ', eDiscountCashRefund, eDiscountCreditNote, eDiscountAppro
            Stat "   .. to reverse"
            setInputBox "", "", "", True
            lblInput.Caption = "Select discount type "
            txtDiscounts = Replace(oPC.DiscountSet, ";", vbCrLf)
            txtDiscounts.Visible = True
            
        Case elogin
            lblInput.Caption = "Staff code."
        Case ePaymentType_Account
            lblInput.Caption = "Charge to account."
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
            Stat "'.. to reverse"
        
        Case ePaymentType_Cash
            lblInput.Caption = "Cash received."
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
            Stat "'.. to reverse"
        Case ePaymentType_Cheque
            lblInput.Caption = "Cheque value."
            Stat "'.. to reverse"
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case ePaymentType_ChequeRef
            setInputBox "", "", "", True
            lblInput.Caption = "Cheque reference."
            Stat "'.. to reverse"
        Case ePaymentType_CreditCard
            lblInput.Caption = "Credit card charge value."
            Stat "'.. to reverse"
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case ePaymentType_DirectDeposit
            lblInput.Caption = "Amount deposited."
            Stat "'.. to reverse"
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case ePaymentType_CreditVoucher
            lblInput.Caption = "Credit voucher value."
            Stat "'.. to reverse"
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case ePaymentType_CreditVoucherRef
            setInputBox "", "", "", True
            lblInput.Caption = "Credit voucher reference."
            Stat "'.. to reverse"
        Case ePaymentType_CreditVoucherRef
            setInputBox "", "", "", True
            lblInput.Caption = "Credit card reference."
            Stat "'.. to reverse"
        Case ePaymentType_RedeemDeposit
            If Len(Trim(txtInput)) > 2 Then
                If IsNumeric(Right(Trim(txtInput), Len(Trim(txtInput)) - 2)) Then
                    iCOLForDeposit = CInt(Right(Trim(txtInput), Len(Trim(txtInput)) - 2))
                    If X3.UpperBound(1) >= iCOLForDeposit And X3.LowerBound(1) <= iCOLForDeposit Then
                        setInputBox CStr(X3(iCOLForDeposit, 12)), "", "", True
                    End If
                End If
            End If
            lblInput.Caption = "Select deposit if available."
            Stat "Select line number, (X) closes deposit list"
            
        Case ePaymentType_voucher
            lblInput.Caption = "Credit voucher value."
            Stat "'.. to reverse"
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case ePaymentType_voucherRef
            setInputBox "", "", "", True
           ' LoadVoucherTypes
            lblInput.Caption = "Voucher code and serial."
            Stat "'.. to reverse"
            txtVouchers.Visible = True
        Case ePrice
            lblInput.Caption = "Price"
            Stat "Hold shift key down and press Enter for discount, '..' to reverse"
            setInputBox oSALELine.Price, "", "", True
            SetTitleBar True
        Case eSale
            setInputBox "", "", "", True
            ShowTransactions False
            lblInput.ForeColor = &H714942
            txtInput.ForeColor = &H714942
            lblInput.Caption = "Scan code or action."
            If bCustomerVisible = True Then
                If G3.Visible = True Then   '  The customers orders are being displayed
                    If oExchange.BalanceOwing < 0 Then
                        Stat "Scan or (X)Cancel trans.,(C)Cash refund,(CC)Reverse credit card,(CV)Credit voucher,(Dn)Del sale,(DPn)Del paymt, (SS) Hold sale "
                    Else
                        Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account,(Dn)Del prod,(DPn)Del paymt,(DDP) Direct deposit, (SS) Hold sale"
                    End If
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account,(Dn)Del sale,(DPn)Del paymt.,(DDP) Direct deposit, (SS) Hold sale"
                End If
            Else
                If oExchange.BalanceOwing < 0 Then
                    Stat "Scan or (X)Cancel trans.,(C)Cash refund,(CC)Reverse credit card,(CV)Issue credit voucher,(Dn)Del sale,(DPn)Del paymt.,(FC) Find customer"
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,Q)Cheque,(AC)On account,(DDP) Direct deposit,(Dn)Del sale,(DPn)Del paymt.,(FC) Find customer"
                End If
            End If
            SetForCOLSVisible False
            DisplayTotals
            AutoSelect txtInput
        Case eStart
            If oExchange.IsEditing Then oExchange.CancelEdit
            Set oExchange = Nothing
            Set oExchange = New a_Exchange
            oExchange.BeginEdit
            ClearSaleLines
            ClearPayments
            setInputBox "", "", "", True
            ShowTransactions False
            lblInput.ForeColor = &H714942
            txtInput.ForeColor = &H714942
            lblInput.Caption = "Start"
           ' Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(CN)Credit note,(Q)Cheque,(AC)Account,(Dn)Del sale,(DPn)Del paymt.,(FC) Find customer"
            Stat "Start by scan or (A)Appro,(AR)Appro retn,(Vn)Void,(OR)Place order,(RDEP)Refund dep, (PA) Pay a/c, (PC)Pet.cash,(PCR)Pet.cash retn, (CNA) Cr. note (a/c cust)"
            ClearCustomer
            ClearTextFields
            SetForCOLSVisible False
            AutoSelect txtInput
        Case eQty                                      ', eQtyCashRefund, eQtyCreditNote, eQtyDeposit, eQtyAppro
            lblInput.Caption = "Qty "
            Stat "'.. to reverse"
            setInputBox oSALELine.Qty, "", "", True
        Case eReviewExchanges
            lblInput.Caption = "Reviewing exchanges"
            Stat "Line number to print, DD to end review."
        Case eSearchCustomer                        ', eSearchCustomerfordeposit, eSearchCustomerfordepositRefund, eSearchCustomerforAppro, eSearchCustomerforApproReturn
            lblInput.Caption = "Search for . . . "
            strArg = Right(Trim(txtInput), Len(Trim(txtInput)) - 1)
            strArg2 = "Name"
            Stat ""
        Case eVoid
            lblInput.Caption = "Voiding #" & iToVoid & " and replacing"
            If bSaleActive Then
                lblInput.Caption = "Scan code or action."
                If bCustomerVisible = True Then
                    Stat "Scan or (A)Appro,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(RDn)Redeem deposit,(Dn)Del prod,(DPn)Del paymt"
                Else
                    Stat "Scan or (A)Appro,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(Dn)Del prod,(DPn)Del paymt,(FC)Find customer"
                End If
            Else
                lblInput.Caption = "Start cash refund "
                Stat "Start replacement by entering product code,   .. to reverse"
            End If
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrepareForm"
End Sub



Private Sub RemovePaymentLine(Optional iRow As Integer, Optional pCurrent As Boolean)
    On Error GoTo errHandler
    If pCurrent Then
        iRow = oExchange.PaymentLines.Count
    End If
    If iRow = 0 Then Exit Sub
    oExchange.PaymentLines.Remove (iRow)
    oExchange.PaymentLines.ApplyEdit
    oExchange.PaymentLines.BeginEdit
    oExchange.CalculateTotals
    txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF
    X2.DeleteRows (iRow)
    G2.ReBind
    iCurrentPaymentLine = iCurrentPaymentLine - 1
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemovePaymentLine(iRow,pCurrent)", Array(iRow, pCurrent)
End Sub

Private Sub RemoveSaleLine(Optional iRow As Integer, Optional pCurrent As Boolean)
    On Error GoTo errHandler
    If pCurrent Then
        iRow = iCurrentSaleLine
    End If
    If iRow = 0 Then Exit Sub
    oExchange.SaleLines.Remove (iRow)
    oExchange.SaleLines.ApplyEdit
    oExchange.SaleLines.BeginEdit
    oExchange.CalculateTotals
   ' X1.DeleteRows (X1.UpperBound(1) - iRow + 1)
    'G1.ReOpen
    'G1.Refresh
    RefreshSaleDisplay
    iCurrentSaleLine = iCurrentSaleLine - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemoveSaleLine(iRow,pCurrent)", Array(iRow, pCurrent)
End Sub

Private Function Action_SaveSale() As eState
   ' oExchange.ApplyEdit
    If oExchange.SaleLines.Count < 1 And oExchange.PaymentLines.Count < 1 Then Exit Function
    Set oExchangeCopy1 = oExchange
    iCurrentSaleLine_store1 = iCurrentSaleLine
    iCurrentPaymentLine_Store1 = iCurrentPaymentLine
    bIssueCreditNote_Store1 = bIssueCreditNote
    CopyArray X1, X1_Store1
    CopyArray X2, X2_Store1
    CopyArray X3, X3_Store1
   ' CopyArray X4, X4_Store1
    CopyArray X5, X5_Store1
    bSaleOnHold = True
    lblSaleOnHold.Visible = bSaleOnHold
    
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    PrepareForNewSale
    Action_SaveSale = eSale
End Function
Private Sub CopyArray(xFrom As XArrayDB, xTo As XArrayDB)
Dim i As Integer
Dim j As Integer
On Error GoTo errHandler
  '  If xFrom Is Empty Then Exit Sub
    xTo.ReDim xFrom.LowerBound(1), xFrom.UpperBound(1), xFrom.LowerBound(2), xFrom.UpperBound(2)
    For i = xFrom.LowerBound(1) To xFrom.UpperBound(1)
        For j = xFrom.LowerBound(2) To xFrom.UpperBound(2)
            xTo(i, j) = xFrom(i, j)
        Next j
    Next i
    Exit Sub
errHandler:
    If Err = 9 Then
        Err.Clear
        Exit Sub
    End If
End Sub



Private Function Action_RetrieveSale() As eState
    If oExchangeCopy1 Is Nothing Then
        Exit Function
    End If
    If oExchange.IsEditing Then oExchange.CancelEdit
    Set oExchange = Nothing
    Set oExchange = oExchangeCopy1
    Set oExchangeCopy1 = Nothing
    iCurrentSaleLine = iCurrentSaleLine_store1
    iCurrentPaymentLine = iCurrentPaymentLine_Store1
    bIssueCreditNote = bIssueCreditNote_Store1
    CopyArray X1_Store1, X1
    CopyArray X2_Store1, X2
    CopyArray X3_Store1, X3
  '  CopyArray X4_Store1, X4
    CopyArray X5_Store1, X5
    bSaleOnHold = False
    RefreshSaleDisplay
    Action_RetrieveSale = eSale
End Function
Private Function Action_CancelSale() As eState
    If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        RejectSale
    Else
        Action_CancelSale = enPresentState
    End If
End Function

Private Sub RejectSale()
    On Error GoTo errHandler
    oExchange.CancelEdit
    Set oExchange = Nothing
    PrepareForNewSale
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RejectSale"
End Sub

Private Sub AcceptSale()
    On Error GoTo errHandler
Dim lngRow As Long
Dim lngLowerBound As Long
Dim oPayment As a_Payment
Dim bCustomerOK As Boolean
Dim strMsg As String
Dim strPos As String
Dim i As Integer

'Save and send exchange
strPos = "00"
'MsgBox strPos
    bItemExchange = False
    If oExchange.NeedsCustomerInfo = True Then
strPos = "01"
'MsgBox strPos
        bCustomerOK = False
        Do Until bCustomerOK = True
strPos = "02"
'MsgBox strPos
            If GetCustomer() Then
                lblCustomername.Caption = DisplayCustomerDetails
            End If
            Set frmCustID = New frmIDCustomer
            If oExchange.Note > "" Then
                frmCustID.component oExchange.Note
            End If
strPos = "03"
'MsgBox strPos
            frmCustID.Show vbModal
            If oExchange.Customer.Name = "" Then
                oExchange.Note = frmCustID.CustomerName
                strMsg = "Confirm customer details:" & vbCrLf & "Name: " & frmCustID.CustomerName & vbCrLf & ""
            Else
                oExchange.Note = vbNullString
                strMsg = "Confirm customer details:" & vbCrLf & "Name: " & oExchange.Customer.Name & vbCrLf & "A/c No.;" & oExchange.Customer.AcNo
            End If
strPos = "04"
            If oExchange.transactionType = "S" Then
                oExchange.Note = FNS(oExchange.Note) & "(" & frmCustID.Counterfoil & ")"
            Else
                oSALELine.Counterfoil = frmCustID.Counterfoil
            End If
strPos = "05"
'MsgBox strPos
            If MsgBox(strMsg, vbInformation + vbYesNo) = vbNo Then
                ClearCustomer
                bCustomerOK = False
            Else
                bCustomerOK = True
            End If
        Loop
    End If
strPos = "06"
'MsgBox strPos
        
    If oExchange.CustomerToBeCredited Then 'THis is to determine in the case of an exchange (not a RDEP) if money is to go out
      'Replaced 8/8/6 so that credit card refunds and cash refunds are both described as rfunds  If oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_Cash) Then
        If oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_Cash) Or oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditCard) Then
            oExchange.SetExchangeType ereturntype
        ElseIf oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditVoucher) Then
            oExchange.SetExchangeType eCreditVoucherType
        ElseIf oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditCard) Then
            oExchange.SetExchangeType eSaleType
        End If
    End If
strPos = "07"
'MsgBox strPos
    oExchange.OperatorID = lngOPID
    oExchange.StaffName = strName
    
    If iToVoid > 0 Then
        oExchange.ToVoid = iToVoid
        lngLowerBound = X4.LowerBound(1)
        lngRow = X4.Find(lngLowerBound, 1, iToVoid)
        Do While lngRow >= lngLowerBound
            X4(lngRow, 12) = oExchange.ExchangeNumber
            If lngRow < X4.UpperBound(1) Then
                lngRow = X4.Find(lngRow + 1, 1, iToVoid)
            End If
        Loop
        G4.Refresh
    End If

'Check to see if a sale has an account component - if so the exchange type must be "A"
    For i = 1 To oExchange.PaymentLines.Count
        If oExchange.PaymentLines(i).PaymentType = "AC" And oExchange.transactionType <> "AR" Then  'A payment has been placed on account
            oExchange.SetExchangeType eAccountSaleType
        End If
    Next i
    oExchange.ApplyEdit
strPos = "11"
'MsgBox strPos
    oPC.OpenLocalDatabase
    oPC.DBLocalConn.BeginTrans
strPos = "12"
'MsgBox strPos



    oExchange.PostExchange
strPos = "13"
'MsgBox strPos
    oPC.DBLocalConn.CommitTrans
    oPC.CloseLocalDatabase
    
'Adds exchange to Xarraydb structure for display
    AddExchange
strPos = "14"
'MsgBox strPos
    SendPOSExchange oExchange.ExchangeID, oExchange.ZID    'oExchange.OPSID,
strPos = "15"
'MsgBox strPos

'Print Till Slip
    If oPC.PrintSlips = True Then
        Select Case oExchange.transactionType
        Case "S", "AR", "A", "CN"
            If (oExchange.Customer.CustomerType = "L1" Or oExchange.Customer.CustomerType = "L2" Or oExchange.Customer.CustomerType = "L3") And (oExchange.LoyaltyValue > 0) Then
                PrintLoyaltyVoucher
            End If
            If oExchange.transactionType = "A" Then
                PrintSalesSlip 3
            Else
                PrintSalesSlip oPC.InvoiceCopyCount
            End If
        Case "R"
            PrintSalesSlip oPC.ReturnCopyCount
        Case "PC", "PCC"
            PrintPettyCashVoucher oPC.PettyCashCopyCount
        Case "C"
            PrintSalesSlip oPC.CreditNoteCopyCount
        Case "DEP"
            PrintDepositSlip oPC.DepositCopyCount
        Case "RDEP"
            PrintDepositRefundSlip oPC.DepositCopyCount
        Case "APP"
            PrintAPPROSlip oPC.ApproCopyCount
        Case "OR"
            PrintORDERSlip oPC.OrderCopyCount
        Case "PA"
            PrintReceiptSlip 3 'oPC.ApproCopyCount
        End Select
    
    'If there is a CV being paid out as change - we must print it
        If bIssueCreditNote Then
            PrintCNasChange oExchange.ChangeVoucherValueF, 1, False
            bIssueCreditNote = False
        End If
    End If
strPos = "16"
'Start new exchange
    Set oExchange = Nothing
    PrepareForNewSale
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AcceptSale", , , , "EXCH:IsEditing,strPOS", Array(oExchange.IsEditing, strPos)
End Sub
Private Sub PrepareForNewSale()
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
    oExchange.SetExchangeType eSaleType
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 0, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
    txtInput.BackColor = RGB(230, 250, 210)
    
    lblCustomername.Caption = vbNullString
    lblReplacement.Visible = False
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    iToVoid = 0
    bSaleActive = False
    bCustomerVisible = False
    SetPresentState eStart
    enMode = emode_Sale
    SetTitleBar True
    SetForCOLSVisible False

End Sub
Private Sub ClearTextFields()
    On Error GoTo errHandler
    txtExtTotal = ""
    txtQtyTotal = ""
    txtVatValue = ""
    txtPaymentTotal = ""
    lblUpdate.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearTextFields"
End Sub
Private Sub ClearPayments()
    On Error GoTo errHandler
    oExchange.PaymentLines.Delete
    oExchange.PaymentLines.ApplyEdit
    oExchange.PaymentLines.BeginEdit
    iCurrentPaymentLine = 0
    X2.Clear
    X2.ReDim 1, iCurrentPaymentLine, 1, 3
    G2.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearPayments"
End Sub
Private Sub ClearSaleLines()
    On Error GoTo errHandler
                
    oExchange.SaleLines.Delete
    oExchange.SaleLines.ApplyEdit
    oExchange.SaleLines.BeginEdit
    iCurrentSaleLine = 0
    X1.Clear
    X1.ReDim 1, 1, 0, 8
    G1.ReBind
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearSaleLines"
End Sub
Private Function LoadProductFromCode(pIn As String) As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oGD As New z_GetData
Dim strPos As String

Dim oLU As z_Lookup
Dim strPID As String
Dim strEAN As String
Dim strCode As String
Dim strTitle As String
Dim strAuthor As String
Dim strDiscountRule As String
Dim lngPrice As Long
Dim lngVatrate As Long
Dim lngDiscount As Long
Dim lngLoyaltyDiscount As Long
Dim bIdentifyCustomer As Boolean
Dim bNoDiscountAllowable As Boolean
Dim rs1 As ADODB.Recordset
Dim bDeleteFromHere As Boolean
Dim i As Integer
Dim bSuccess As Boolean

    Set oLU = New z_Lookup
    strEAN = pIn ' Trim$(txtInput)
    strCode = pIn  'Trim$(txtInput)
    Set rs = oLU.GetProduct(strEAN, strCode, oExchange.Customer.CustomerType, bSuccess)
    If bSuccess = False Then
        MsgBox "Cannot connect to database. Cancel this transaction and inform your supervisor." & vbCrLf & "You can try closing this application and starting it again to clear the error.", vbCritical + vbOKOnly, "Warning"
        Set rs = Nothing
        Exit Function
    End If
    If rs Is Nothing Then
        LoadProductFromCode = False
        Set oLU = Nothing
        Exit Function
    ElseIf rs.State = 0 Then
        LoadProductFromCode = False
        Set rs = Nothing
        Set oLU = Nothing
        Exit Function
    ElseIf rs.RecordCount = 0 Then
        LoadProductFromCode = False
        rs.Close
        Set rs = Nothing
        Set oLU = Nothing
        Exit Function
    End If

'Create new sale line for Exchange
    Set oSALELine = oExchange.SaleLines.Add
    iCurrentSaleLine = iCurrentSaleLine + 1
    X1.ReDim 1, iCurrentSaleLine, 0, 8
    
'Load rules into Sales line
    oSALELine.LoadRules rs
    Set rs = Nothing
    Set oLU = Nothing
    
    oSALELine.FindRule oExchange.NominalValue, bIdentifyCustomer

    oExchange.IdentifyCustomer = bIdentifyCustomer
    If oExchange.IdentifyCustomer = True And oExchange.Note = "" Then
        Set frmCustID = New frmIDCustomer
        frmCustID.component oExchange.Note
        frmCustID.Show vbModal
        oExchange.Note = frmCustID.CustomerName
        oSALELine.Counterfoil = frmCustID.Counterfoil
        Unload frmCustID
    End If
     
     
     oSALELine.ApplyEdit
     oSALELine.BeginEdit
     LoadProductFromCode = True
     
     oPC.CloseLocalDatabase
     
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadProductFromCode"
End Function
Private Sub oExchange_Recalculate()
    On Error GoTo errHandler
    If bUnloading Then Exit Sub
    RefreshSaleDisplay
    DisplayTotals
    lblCustomername.Caption = DisplayCustomerDetails

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_Recalculate", , EA_NORERAISE
    HandleError
End Sub
Private Sub oExchange_GetCustomer()
    If (Not oExchange.Customer.ID > 0) And (oExchange.Note = "") Then
        GetCustomer
    End If
End Sub

Private Sub RefreshRules()
Dim oLU As z_Lookup
Dim oSL As a_Sale
Dim rs As ADODB.Recordset
    Set oLU = New z_Lookup
    For Each oSL In oExchange.SaleLines
       ' Call oLU.GetProduct(strEAN, strCode, strPID, strTitle, strAuthor, lngVatrate, lngPrice, lngDiscount, lngLoyaltyDiscount, bIdentifyCustomer, bNoDiscountAllowable, strDiscountRule)
        Set rs = oLU.GetProduct(oSL.Code, oSL.Code, oExchange.Customer.CustomerType)
        If rs.RecordCount = 0 Then
            rs.Close
            Set rs = Nothing
            Set oLU = Nothing
            Exit Sub
        End If
    'Load rules into Sales line
        oSL.LoadRules rs
        Set rs = Nothing
        oSL.FindRule oExchange.NominalValue, False
    Next
    Set oLU = Nothing
End Sub
Private Function GetCustomer() As Boolean
    On Error GoTo errHandler
Dim frm As New frmBrowseCustomers2
Dim cnt As Integer
Dim s As String

s = "pos 1"
    GetCustomer = False
    frm.Show vbModal
    If frm.IsCancelled Then Exit Function
s = "pos 2"
Start:
    If frm.CustomerName > "" Then
s = "pos 3"
        bCustomerVisible = True
        strCustomername = frm.CustomerName
        G3.Caption = frm.CustomerName
s = "pos 4"
        lngCustomerID = frm.CustomerID
        oExchange.SetCustomer lngCustomerID
        oExchange.Note = frm.CustomerName
s = "pos 5"
        GetCustomer = True
        RefreshRules
s = "pos 6"
        LookForAlert frm.Accnum, frm.CustomerName
s = "pos 7"
    Else
s = "pos 8"
        GetCustomerFromMasterDB_control frm.Accnum, cnt
s = "pos 9"
        ClearCustomer
s = "pos 10"
        If cnt > 0 Then GoTo Start
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetCustomer", , , , "s", Array(s)
End Function
Private Sub LookForAlert(TPACNO As String, CustomerName As String)
Dim frm As frmAlert
Dim oLU As New z_Lookup
Dim rs As ADODB.Recordset

    Set rs = oLU.GetAlertFromMasterDB(TPACNO)
    If Not rs.EOF Then
        Set frm = New frmAlert
        frm.component rs
'        frm.lblCustomer.Caption = CustomerName & "  (" & TPACNO & ")"
'        frm.lblMsg.Caption = FNS(rs.Fields("AL_MSGTEXT"))
        frm.Show vbModal
    End If
End Sub
Private Sub GetCustomerFromMasterDB_control(AcNo As String, cnt As Integer)
    On Error GoTo errHandler
Dim oLU As New z_Lookup
    oLU.GetCustomerFromMasterDB AcNo, cnt
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetCustomerFromMasterDB_control(AcNo,cnt)", Array(AcNo, cnt)
End Sub

Private Function ClearCustomer()
    On Error GoTo errHandler
        oExchange.Customer.SetName ""
        oExchange.Customer.SetAcNO ""
        lngCustomerID = 0
        strCustomername = ""
        bCustomerVisible = False
        G3.Caption = ""
        G3.Visible = False
        Me.lblCustomername.Caption = ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearCustomer"
End Function
Private Function DisplayCustomerDetails() As String
    On Error GoTo errHandler
Dim strDetails As String
    Select Case UCase(oExchange.Customer.CustomerType)
    Case "L1"   'Loyalty club 1 member
            strDetails = oExchange.Customer.NameAndCode(99) & " " & "Loyalty value: " & oExchange.LoyaltyValueF
    Case "BC"
        strDetails = oExchange.Customer.NameAndCode(99) & " " & "(Book club)" & ":" & oExchange.DiscountRateF
    Case ""
        strDetails = oExchange.Customer.NameAndCode(99)
    End Select
    
    DisplayCustomerDetails = strDetails & vbCrLf & strCreditLimitExceededMessage
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayCustomerDetails"
End Function



'Private Sub FetchInvs()
'    On Error GoTo errHandler
'    Set cINVs = New c_Invs
'    cINVs.Load
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.FetchInvs"
'End Sub
'Private Sub LoadINVs()
'    On Error GoTo errHandler
'Dim i As Long
'
'    G5.Visible = True
'    X5.Clear
'    X5.ReDim 1, cINVs.Count, 1, 8
'    For i = 1 To cINVs.Count
'        With cINVs(i)
'            X5.Value(i, 1) = i
'            X5.Value(i, 2) = .docCode
'            X5.Value(i, 3) = .CustomerName
'            X5.Value(i, 4) = .PayableF
'            X5.Value(i, 5) = .INVID
'            X5.Value(i, 6) = .Payable
'            X5.Value(i, 7) = .TPID
'            X5.Value(i, 8) = .VAT
'
'        End With
'    Next
'  '  X5.QuickSort 1, X5.UpperBound(1), 10, XORDER_DESCEND, XTYPE_STRING   'sorted in query - we need the ordinal poistion (column 1) to be in sequence
'    G5.Array = X5
'    Me.G5.ReBind
'    SetForINVSVisible True
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LoadINVs"
'End Sub
Private Sub FetchCOLS()
    On Error GoTo errHandler
    Set cCOLS = New c_COLS
    cCOLS.Load lngCustomerID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.FetchCOLS"
End Sub
Private Sub LoadCOLS()
    On Error GoTo errHandler
Dim i As Long

    G3.Visible = True
    X3.Clear
    X3.ReDim 1, cCOLS.Count, 1, 14
    For i = 1 To cCOLS.Count
        With cCOLS(i)
            X3.Value(i, 1) = i
            X3.Value(i, 2) = .COLDateF
            X3.Value(i, 3) = .Code
            X3.Value(i, 5) = .Description
            X3.Value(i, 4) = .Qty & "(" & .QtyDispatched & ")"
            X3.Value(i, 6) = .DepositF
            X3.Value(i, 7) = .DepositStatus
            X3.Value(i, 8) = .PriceF
            X3.Value(i, 9) = .DiscountRateF
            X3.Value(i, 10) = .COLDateForSORT
            X3.Value(i, 11) = .COLID
            X3.Value(i, 12) = .Deposit
            X3.Value(i, 13) = .PID
            X3.Value(i, 14) = .Qty
        End With
    Next
  '  X3.QuickSort 1, X3.UpperBound(1), 10, XORDER_DESCEND, XTYPE_STRING   'sorted in query - we need the ordinal position (column 1) to be in sequence
    G3.Array = X3
    Me.G3.ReBind
    SetForCOLSVisible True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadCOLS"
End Sub
Private Sub SetForCOLSVisible(pYes As Boolean)
    On Error GoTo errHandler
    If pYes Then
        G3.Visible = True
        G3.ZOrder 0
    Else
        G3.Visible = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetForCOLSVisible(pYes)", pYes
End Sub
Private Sub RefreshSaleDisplay()
    On Error GoTo errHandler
Dim i As Integer

    X1.Clear
    X1.ReDim 1, oExchange.SaleLines.Count, 0, 8
    For i = 1 To oExchange.SaleLines.Count
        LoadSaleRow i, oExchange.SaleLines.Count
    Next
    For i = 1 To oExchange.PaymentLines.Count
        LoadPaymentRow i
    Next
    lblSaleOnHold.Visible = bSaleOnHold
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefreshSaleDisplay"
End Sub



Private Sub LoadSaleRow(Index As Integer, MaxLines As Long)
    On Error GoTo errHandler
Dim i As Long
Dim strPos As String
    strPos = "Pos 1"
    G1.Visible = True
        strPos = "Pos 2"
        
    X1.Value(MaxLines - Index + 1, 0) = Index
        strPos = "Pos 2b"
    X1.Value(MaxLines - Index + 1, 1) = oExchange.SaleLines(Index).CodeF
        strPos = "Pos 3"
    X1.Value(MaxLines - Index + 1, 2) = IIf(enPresentState = eSelectDepositLine, "(DEP)", "") & oExchange.SaleLines(Index).title & " (" & oExchange.SaleLines(Index).MainAuthor & ")"
    X1.Value(MaxLines - Index + 1, 3) = oExchange.SaleLines(Index).Qty
            strPos = "Pos 4"
    X1.Value(MaxLines - Index + 1, 4) = oExchange.SaleLines(Index).PriceF & IIf(oExchange.SaleLines(Index).IsSpecialPrice, "**", "")
    X1.Value(MaxLines - Index + 1, 5) = oExchange.SaleLines(Index).DiscountRateF
        strPos = "Pos 5"
    X1.Value(MaxLines - Index + 1, 6) = oExchange.SaleLines(Index).PLessDiscExtF
        strPos = "Pos 6"
    If oExchange.transactionType <> "INV" Then
        strPos = "Pos 7"
        X1.Value(MaxLines - Index + 1, 7) = oExchange.SaleLines(Index).PLessDiscExtVATF & "(" & oExchange.SaleLines(Index).VATRateF & ")"
    End If
        strPos = "Pos 8"
    G1.ReBind
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadSaleRow(Index,MaxLines)", Array(Index, MaxLines), , , "strPOS", Array(strPos)
End Sub




Private Sub DisplayTotals()
    On Error GoTo errHandler
    
    txtExtTotal = oExchange.TotalPayableF
    'MsgBox oPC.CurrencyDivisor
    txtQtyTotal = oExchange.TotalQty
    txtVatValue = oExchange.TotalVATF
      '  txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF & IIf(oExchange.ChangeGiven > 0, " (To customer: " & oExchange.ChangeGivenF & ")", "")
      If oExchange.transactionType = "S" Then
        If oExchange.ChangeGiven < 0 Then
            txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF & "(Still owing " & oExchange.ChangeGivenF & ")"
        Else
            txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF & "(Change: " & oExchange.ChangeGivenF & ")"
        End If
    Else
        txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF & IIf(oExchange.ChangeGiven > 0, " (To customer: " & oExchange.ChangeGivenF & ")", "")
    End If
    lblUpdate.Caption = "Pre-disc: " & oExchange.NominalValueF
    lblUpdate.Visible = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayTotals"
End Sub
Private Sub LoadPaymentRow(iIndex As Integer)
    On Error GoTo errHandler
Dim i As Long
    If iIndex > X2.UpperBound(1) Then Exit Sub
    G2.Visible = True
   ' X2.ReDim 1, iIndex, 1, 3
    X2.Value(iIndex, 3) = oExchange.PaymentLines(iIndex).ReferenceComplete
    X2.Value(iIndex, 2) = oExchange.PaymentLines(iIndex).AmtF
    X2.Value(iIndex, 1) = oExchange.PaymentLines(iIndex).PaymentTypeF
    G2.Array = X2
    G2.ReBind
    G2.Refresh
    DoEvents
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadPaymentRow(iIndex)", iIndex
End Sub




Private Sub Connect()
    On Error GoTo errHandler
    Set oSQLServer = New SQLDMO.SQLServer
    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
    With oSQLServer
        .LoginSecure = False
        .AutoReConnect = False
        .Connect oPC.LocalSQLServerName, "sa", ""
    End With
    
    Set oDatabase = oSQLServer.Databases("PBKSFD")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Connect"
End Sub

Private Sub RebuildIndexes()
    On Error GoTo errHandler
Dim oTable As SQLDMO.Table
    For Each oTable In oDatabase.Tables
        If Not oTable.SystemObject Then oTable.RebuildIndexes
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RebuildIndexes"
End Sub
Private Function Disconnect()
    On Error GoTo errHandler
    oSQLServer.Disconnect
    Set ADOConn = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Disconnect"
End Function

Private Sub PrintSalesSlip(pCopyCount As Integer, Optional bReprint As Boolean, Optional toA4 As Boolean = False)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
    Dim strDiscountDescription As String
    Dim lValue As Long
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
    Dim sDiscDesc As String
    Dim sCounterfoil As String
    Dim bPriceAlteration As Boolean
    Dim iLineCount As Integer
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass
            BcData = "4902720005074"
            
            If oExchange.SaleLines.Count > 0 Then
            ReDim idBuf(1 To oExchange.SaleLines.Count)
                For i = 1 To oExchange.SaleLines.Count
                    If Not oExchange.SaleLines(i).IsDeleted Then
                        idBuf(i).TType = IIf(oExchange.SaleLines(i).Qty < 0, "R ", "S ")
                        idBuf(i).Name = oExchange.SaleLines(i).title
                        idBuf(i).Disc = oExchange.SaleLines(i).DiscountRateF
                        idBuf(i).Ext = oExchange.SaleLines(i).PLessDiscExtF
                        idBuf(i).At = oExchange.SaleLines(i).QtyF & " @ " & oExchange.SaleLines(i).PriceF
                        idBuf(i).Alteration = oExchange.SaleLines(i).PriceAlteration
                        idBuf(i).Counterfoil = oExchange.SaleLines(i).Counterfoil
                        If oExchange.SaleLines(i).DiscountRule = "" Then
                            idBuf(i).DiscDesc = oExchange.SaleLines(i).DiscountDescription
                        Else
                            idBuf(i).DiscDesc = oExchange.SaleLines(i).DiscountRule
                        End If
                    End If
                Next i
            End If
            For c = 1 To pCopyCount
                If oPC.UseA4Printer Then
                        PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint
                        iLineCount = 0
                        If oExchange.SaleLines.Count > 0 Then
                            For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                                If iLineCount > 24 Then Printer.NewPage
                                sAt = idBuf(i).At
                                sBuf = idBuf(i).Name
                                sExt = idBuf(i).Ext
                                sType = idBuf(i).TType
                                sDisc = idBuf(i).Disc
                                sDiscDesc = idBuf(i).DiscDesc
                                bPriceAlteration = idBuf(i).Alteration
                                sCounterfoil = idBuf(i).Counterfoil
                                
                                sValue = MakePrintStringDetail(50, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                                '.PrintNormal PTR_S_RECEIPT, sValue + vbLf
                                Printer.Print sValue
                                strDiscountDescription = ""
                                If sDiscDesc > "" Then
                                    strDiscountDescription = Left(sDiscDesc, 20) & " " & sDisc
                                Else
                                    strDiscountDescription = "Disc: " & sDisc
                                End If
                                
                               ' .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "") & vbLf
                                Printer.Print oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "")
                                If sCounterfoil > "" Then
                                    '.PrintNormal PTR_S_RECEIPT, "Ref: " & sCounterfoil & vbLf
                                    Printer.Print "Ref: " & sCounterfoil
                                End If
                                iLineCount = iLineCount + 1
                            Next
                        End If
                        Printer.Print ""
                        PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
                        PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                        Printer.EndDoc
               Else
                    With OPOSPOSPrinter1
                        PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                        
                        If oExchange.SaleLines.Count > 0 Then
                            For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                                If .ResultCode <> OPOS_SUCCESS Then Exit For
                                sAt = idBuf(i).At
                                sBuf = idBuf(i).Name
                                sExt = idBuf(i).Ext
                                sType = idBuf(i).TType
                                sDisc = idBuf(i).Disc
                                sDiscDesc = idBuf(i).DiscDesc
                                bPriceAlteration = idBuf(i).Alteration
                                sCounterfoil = idBuf(i).Counterfoil
                                
                                sValue = MakePrintStringDetail(.RecLineChars, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                                strDiscountDescription = ""
                                If sDiscDesc > "" Then
                                    strDiscountDescription = Left(sDiscDesc, 20) & " " & sDisc
                                Else
                                    strDiscountDescription = "Disc: " & sDisc
                                End If
                                
                                .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "") & vbLf
                                If sCounterfoil > "" Then
                                    .PrintNormal PTR_S_RECEIPT, "Ref: " & sCounterfoil & vbLf
                                End If
                            Next
                        End If
                        .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                            
                        PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
                        PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                        
                        .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                        .CutPaper 90
            
                        .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
                
                        'Back to the synchronous mode
                        .AsyncMode = False
                        
                    End With
                End If
            Next
        Printer.EndDoc

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintSalesSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
End Sub
Private Sub PrintReceiptSlip(pCopyCount As Integer, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
    Dim strDiscountDescription As String
    Dim lValue As Long
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
    Dim sDiscDesc As String
    Dim sCounterfoil As String
    Dim bPriceAlteration As Boolean
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
        ''''
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                
                sValue = MakePrintStringDetail(iColWidth, "Payment received ", "", "", oExchange.PaymentLines(1).AmtF, 0, False)
                Printer.Print ""
                Printer.Print sValue
                    
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                
                Printer.Print ""
                Printer.EndDoc
        ''''
        Else
            With OPOSPOSPrinter1
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                
                sValue = MakePrintStringDetail(.RecLineChars, "Payment received ", "", "", oExchange.PaymentLines(1).AmtF, 0, False)
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                    
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf      'create gap
                    
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
                
            End With
        End If
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintReceiptSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
End Sub

Private Sub PrintAPPROSlip(pCopyCount As Integer, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
    Dim strDiscountDescription As String
    Dim lValue As Long
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
    Dim sDiscDesc As String
    Dim bPriceAlteration As Boolean
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    ReDim idBuf(1 To oExchange.SaleLines.Count)
    For i = 1 To oExchange.SaleLines.Count
        If Not oExchange.SaleLines(i).IsDeleted Then
            idBuf(i).TType = IIf(oExchange.SaleLines(i).Qty < 0, "R ", "S ")
            idBuf(i).Name = oExchange.SaleLines(i).title
            idBuf(i).Disc = oExchange.SaleLines(i).DiscountRateF
            idBuf(i).Ext = oExchange.SaleLines(i).PLessDiscExtF
            idBuf(i).At = oExchange.SaleLines(i).QtyF & " @ " & oExchange.SaleLines(i).PriceF
            idBuf(i).Alteration = oExchange.SaleLines(i).PriceAlteration
            idBuf(i).DiscDesc = oExchange.SaleLines(i).DiscountRule
        End If
    Next i
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                    sAt = idBuf(i).At
                    sBuf = idBuf(i).Name
                    sExt = idBuf(i).Ext
                    sType = idBuf(i).TType
                    sDisc = idBuf(i).Disc
                    sDiscDesc = idBuf(i).DiscDesc
                    bPriceAlteration = idBuf(i).Alteration
                    sValue = MakePrintStringDetail(iColWidth, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    Printer.Print sValue
                    strDiscountDescription = ""
                    If sDiscDesc > "" Then
                        strDiscountDescription = Left(sDiscDesc, 20) & " " & sDisc
                    Else
                        strDiscountDescription = "Disc: " & sDisc
                    End If
                    Printer.Print oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "")
                Next
                Printer.Print ""
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                
                Printer.Print ""
                Printer.EndDoc
       Else
            With OPOSPOSPrinter1
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                    If .ResultCode <> OPOS_SUCCESS Then Exit For
                    sAt = idBuf(i).At
                    sBuf = idBuf(i).Name
                    sExt = idBuf(i).Ext
                    sType = idBuf(i).TType
                    sDisc = idBuf(i).Disc
                    sDiscDesc = idBuf(i).DiscDesc
                    bPriceAlteration = idBuf(i).Alteration
                    sValue = MakePrintStringDetail(.RecLineChars, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                    strDiscountDescription = ""
                    If sDiscDesc > "" Then
                        strDiscountDescription = Left(sDiscDesc, 20) & " " & sDisc
                    Else
                        strDiscountDescription = "Disc: " & sDisc
                    End If
                    .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "") & vbLf
                Next
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                    
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
                
            End With
        End If
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintAPPROSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
End Sub
Private Sub PrintORDERSlip(pCopyCount As Integer, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
    Dim strDiscountDescription As String
    Dim lValue As Long
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
    Dim sDiscDesc As String
    Dim bPriceAlteration As Boolean
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                Printer.Print ""
                
                
                sBuf = "Order:" & strOrderedTitle
                Printer.Print sBuf
                Printer.Print ""
                
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint          'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                Printer.Print ""
                Printer.EndDoc
        Else
            With OPOSPOSPrinter1
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                
                
                sBuf = "Order:" & strOrderedTitle
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sBuf + vbLf
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|1200uF"     'create gap
                
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint          'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
            End With
        End If
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Printer.EndDoc
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintAPPROSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
End Sub

Private Sub PrintCNasChange(pAmtF As String, pCopyCount As Integer, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
Dim lValue As Long
Dim idBuf() As ITEMDATA
Dim fDate As String
Dim BcData  As String
Dim sBuf As String
Dim sExt As String
Dim sType As String
Dim sDisc As String
Dim sAt As String
Dim sValue As String
Dim bPriceAlteration As Boolean
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader ConvertToType("CV"), OPOSPOSPrinter1, bReprint      'Print header
                Printer.Print ""
                    
                With OPOSPOSPrinter1
                    sBuf = "Change Voucher"
                    sExt = pAmtF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    Printer.Print sValue
                End With
                PrintFooter c, ConvertToType("C"), OPOSPOSPrinter1          'print footer
                Printer.Print ""
                Printer.EndDoc
       Else
            With OPOSPOSPrinter1
                PrintHeader ConvertToType("CV"), OPOSPOSPrinter1, bReprint      'Print header
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                    
                With OPOSPOSPrinter1
                    sBuf = "Change Voucher"
                    sExt = pAmtF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                End With
                PrintFooter c, ConvertToType("C"), OPOSPOSPrinter1          'print footer
                .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                .AsyncMode = False
            End With
        End If
    Next

    MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintCNasChange(pAmtF,pCopyCount,bReprint)", Array(pAmtF, pCopyCount, _
         bReprint)
End Sub

Private Sub PrintLoyaltyVoucher()
    On Error GoTo errHandler
    Dim lValue As Long
    Dim i As Integer
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String

' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    If oPC.UseA4Printer Then
            PrintHeader eTypVoucher, OPOSPOSPrinter1           'Print header
            Printer.Print ""
            Printer.Print "Credit value: " & oExchange.LoyaltyValueF
            Printer.Print ""
            PrintFooter 1, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
            Printer.Print ""
            Printer.EndDoc
    Else
        With OPOSPOSPrinter1
            PrintHeader eTypVoucher, OPOSPOSPrinter1           'Print header
            .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
           
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "Credit value: " & oExchange.LoyaltyValueF + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                
            PrintFooter 1, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
            
        End With
    End If
' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintLoyaltyVoucher"
End Sub
Private Sub PrintPettyCashVoucher(pCopyCount As Integer, Optional bReprint As Boolean)
    On Error GoTo errHandler
    Dim lValue As Long
    Dim i As Integer
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
Dim c As Integer
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader eTypPettyCash, OPOSPOSPrinter1, bReprint          'Print header
                Printer.Print ""
                If oExchange.PaymentLines(1).PaymentType = "W" Then
                    Printer.Print "Petty Cash: " & oExchange.PaymentLines(1).AmtF
                Else
                    Printer.Print "Petty Cash Refund: " & oExchange.PaymentLines(1).AmtF
                End If
                Printer.Print ""
                Printer.Print oExchange.Note
                Printer.EndDoc
        Else
            With OPOSPOSPrinter1
                PrintHeader eTypPettyCash, OPOSPOSPrinter1, bReprint          'Print header
                .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
                If oExchange.PaymentLines(1).PaymentType = "W" Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash: " & oExchange.PaymentLines(1).AmtF + vbLf
                Else
                    .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash Refund: " & oExchange.PaymentLines(1).AmtF + vbLf
                End If
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                .PrintNormal PTR_S_RECEIPT, oExchange.Note
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|5500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
            End With
        End If
    Next
' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintPettyCashVoucher(pCopyCount)", pCopyCount
End Sub

Private Function MakePrintStringDetail(ByVal lRecLineChars As Long, sType As String, sBuf As String, sAt As String, sExt As String, sDisc As String, PriceAlteration As Boolean) As String
    On Error GoTo errHandler
Dim sValue As String
Dim strNotChangeable As String
Dim iAvailable As Integer
    sAt = " " & sAt
    sExt = " " & sExt
    iAvailable = lRecLineChars - Len(sAt) - Len(sType) - Len(sExt)
    If PriceAlteration = True Then
        sAt = ESC + "|uC" & sAt
    End If
    sBuf = Left(sBuf, iAvailable)
    sBuf = sBuf & Space(iAvailable - Len(sBuf))
    If oPC.UseA4Printer Then
        MakePrintStringDetail = sType & sBuf & sAt & sExt
    Else
        MakePrintStringDetail = sType & sBuf & sAt & ESC + "|N" & sExt
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.MakePrintStringDetail(lRecLineChars,sType,sBuf,sAt,sExt,sDisc," & _
        "PriceAlteration)", Array(lRecLineChars, sType, sBuf, sAt, sExt, sDisc, PriceAlteration)
End Function
Private Sub PrintDepositRefundSlip(pCopyCount As Integer)
    On Error GoTo errHandler
    Dim lValue As Long
    Dim i As Integer
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String

' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For i = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader eTypDepositRefund, OPOSPOSPrinter1           'Print header
                Printer.Print ""      'create gap
               
                Printer.Print "REFUNDED.: " & oExchange.PaymentLines(1).AmtF_nonNegative
                Printer.Print strDepositTitle
                Printer.Print ""      'create gap
                PrintFooter i, eTypDepositRefund, OPOSPOSPrinter1          'print footer
                Printer.Print ""      'create gap
                Printer.EndDoc
        Else
            With OPOSPOSPrinter1
                PrintHeader eTypDepositRefund, OPOSPOSPrinter1           'Print header
                .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
               
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "REFUNDED.: " & oExchange.PaymentLines(1).AmtF_nonNegative + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & strDepositTitle
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
              '  .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(i)
                PrintFooter i, eTypDepositRefund, OPOSPOSPrinter1          'print footer
                
                .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
                
            End With
        End If
    Next i

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Exit Sub


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintDepositRefundSlip(pCopyCount)", pCopyCount
End Sub
Private Sub PrintDepositSlip(pCopyCount As Integer)
    On Error GoTo errHandler
Dim lValue As Long
Dim i As Integer
Dim j As Integer
Dim idBuf() As ITEMDATA
Dim fDate As String
Dim BcData  As String
Dim sBuf As String
Dim sExt As String
Dim sType As String
Dim sDisc As String
Dim sAt As String
Dim sValue As String
Dim bPriceAlteration As Boolean

Dim strPos As String

' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    ReDim idBuf(1 To oExchange.SaleLines.Count)
    For j = 1 To oExchange.SaleLines.Count
        If Not oExchange.SaleLines(j).IsDeleted Then
            idBuf(j).TType = IIf(oExchange.SaleLines(j).Qty < 0, "R ", "S ")
            idBuf(j).Name = oExchange.SaleLines(j).title
            idBuf(j).Disc = oExchange.SaleLines(j).DiscountRateF
            idBuf(j).Ext = oExchange.SaleLines(j).PLessDiscExtF
            idBuf(j).At = oExchange.SaleLines(j).QtyF & " @ " & oExchange.SaleLines(j).PriceF
            idBuf(j).Alteration = oExchange.SaleLines(j).PriceAlteration
            idBuf(j).DiscDesc = oExchange.SaleLines(j).DiscountRule
        End If
    Next j
    For i = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader eTypDeposit, OPOSPOSPrinter1             'Print header
                Printer.Print ""      'create gap
                For j = LBound(idBuf) To UBound(idBuf)          'Print each line
                    sAt = idBuf(j).At
                    sBuf = idBuf(j).Name
                    sExt = idBuf(j).Ext
                    sType = idBuf(j).TType
                    sDisc = idBuf(j).Disc
                    bPriceAlteration = idBuf(j).Alteration
                    sValue = MakePrintStringDetail(iColWidth, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    Printer.Print sValue
                    Printer.Print oExchange.SaleLines(1).CodeF & ":DEPOSIT PAID"
                Next j
                Printer.Print "Deposit paid: " & oExchange.TotalPayableF
                Printer.Print "Change: " & oExchange.ChangeGivenF
                Printer.Print ""      'create gap
                Printer.Print "Copy number: " & CStr(i)
                PrintFooter i, eTypDeposit, OPOSPOSPrinter1          'print footer
                Printer.Print ""     'create gap
                Printer.EndDoc
        Else
            With OPOSPOSPrinter1
                PrintHeader eTypDeposit, OPOSPOSPrinter1             'Print header
                .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
                For j = LBound(idBuf) To UBound(idBuf)          'Print each line
                    If .ResultCode <> OPOS_SUCCESS Then Exit For
                    sAt = idBuf(j).At
                    sBuf = idBuf(j).Name
                    sExt = idBuf(j).Ext
                    sType = idBuf(j).TType
                    sDisc = idBuf(j).Disc
                    bPriceAlteration = idBuf(j).Alteration
                    sValue = MakePrintStringDetail(.RecLineChars, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                    .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(1).CodeF & ":DEPOSIT PAID" & vbLf
                Next j
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Deposit paid: " & oExchange.TotalPayableF + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Change: " & oExchange.ChangeGivenF + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(i)
                PrintFooter i, eTypDeposit, OPOSPOSPrinter1          'print footer
                .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
                'Back to the synchronous mode
                .AsyncMode = False
            End With
        End If
    Next i

    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintDepositSlip(pCopyCount)", pCopyCount
End Sub
Private Function MakePrintString(ByVal lRecLineChars As Long, sBuf As String, sPrice As String) As String
    On Error GoTo errHandler
Dim sValue As String
    If lRecLineChars < (Len(sBuf) + Len(sPrice)) Then
        sValue = sBuf + sPrice
    Else
        sValue = sBuf + Space(lRecLineChars - (Len(sBuf) + Len(sPrice))) + sPrice
    End If

    MakePrintString = sValue
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.MakePrintString(lRecLineChars,sBuf,sPrice)", Array(lRecLineChars, sBuf, _
         sPrice)
End Function

Private Sub AddExchange()
    On Error GoTo errHandler
Dim oSALE As a_Sale
    Select Case oExchange.TransactionTypeEnum
    Case eSaleType, ereturntype, eCreditVoucherType, eAccountCreditNoteType, eAccountSaleType
        For Each oSALE In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = oSALE.CodeF  ' & " (" & oSale.QtyF & ") " & oSale.TitleF(30) & " " & oSale.PLessDiscF
            X4.Value(lngSalesItemCount, 6) = oSALE.QtyF
            X4.Value(lngSalesItemCount, 7) = oSALE.TitleF(30) & IIf(oExchange.ToVoid > 0, " (Voids:" & oExchange.ToVoid & ")", "")
            X4.Value(lngSalesItemCount, 8) = oSALE.PriceF
            X4.Value(lngSalesItemCount, 9) = oSALE.PLessDiscExtF
            
            'Add 4 columns
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 11) = oSALE.PID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next
    Case eApproType
        For Each oSALE In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = oSALE.CodeF  ' & " (" & oSale.QtyF & ") " & oSale.TitleF(30) & " " & oSale.PLessDiscF
            X4.Value(lngSalesItemCount, 6) = oSALE.QtyF
            X4.Value(lngSalesItemCount, 7) = oSALE.TitleF(30) & IIf(oExchange.ToVoid > 0, " (Voids:" & oExchange.ToVoid & ")", "")
            X4.Value(lngSalesItemCount, 8) = oSALE.PriceF
            X4.Value(lngSalesItemCount, 9) = oSALE.PLessDiscExtF
            
            'Add 4 columns
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 11) = oSALE.PID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next
    Case eDepositType, eOrderRequestType
        For Each oSALE In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "DEPOSIT"
            X4.Value(lngSalesItemCount, 8) = oSALE.PriceF
            X4.Value(lngSalesItemCount, 9) = oSALE.PLessDiscExtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next
    Case eApproReturnType
        For Each oSALE In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = oSALE.Qty
            X4.Value(lngSalesItemCount, 7) = oSALE.TitleF(30)
            X4.Value(lngSalesItemCount, 8) = oSALE.PriceF
            X4.Value(lngSalesItemCount, 9) = oSALE.PLessDiscExtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next

    Case eReturnDepositType
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "RETURN DEPOSIT"
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
    Case ePettyCashType, eAccountPaymentType
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            If oExchange.TransactionTypeEnum = eAccountPaymentType Then
                X4.Value(lngSalesItemCount, 7) = "ACOUNT PAYMENT" & ":" & oExchange.Note
            Else
                X4.Value(lngSalesItemCount, 7) = "PETTY CASH" & ":" & oExchange.Note
            End If
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
    Case ePettyCashCreditType
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "PETTY CASH CREDIT" & ":" & oExchange.Note
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
    Case eVoidAction
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "VOIDING" & ":" & oExchange.Note
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = ""
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
    Case eOpenDrawerType
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "OPEN DRAWER" & ":" & oExchange.Note
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = ""
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
    End Select
    X4.QuickSort 1, X4.UpperBound(1), 1, XORDER_DESCEND, XTYPE_NUMBER
    G4.Array = X4
    G4.ReBind
    G4.Bookmark = 1 'lngSalesItemCount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AddExchange"
End Sub
Private Sub G4_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    
    If X4(Bookmark, 12) <> 0 Then
        RowStyle.BackColor = RGB(192, 192, 192)
    ElseIf X4(Bookmark, 13) > 0 Then
        RowStyle.BackColor = RGB(176, 222, 173)
    Else
        RowStyle.BackColor = &HFFFFFF
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.G4_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub LoadExchanges()
    On Error GoTo errHandler
Dim ZID As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim i As Integer

'    For i = 1 To G4.Columns.Count
'        G4.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G4.Columns(i - 1).Width)
'    Next
    
    ZID = oPC.ZSession.Current_Z_Session_ID
    
    oPC.OpenLocalDatabase

    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = oPC.DBLocalConn
    cmd.CommandText = "q_ExchangeDetails"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@ZSESSID", adGUID, adParamInput, , ZID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TITLELENGTH", adInteger, adParamInput, , 50)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@CurrencyDivisor", adInteger, adParamInput, , 100)
    cmd.Parameters.Append prm
    Set prm = Nothing
   
    lngSalesItemCount = 0
    Set rs = cmd.Execute
    Do While Not rs.EOF
        lngSalesItemCount = lngSalesItemCount + 1
        X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = FNN(rs.Fields("EXCH_NUMBER"))
            X4.Value(lngSalesItemCount, 2) = Format(rs.Fields("EXCH_SaleDate"), "HH:NN")
            X4.Value(lngSalesItemCount, 3) = FNS(rs.Fields("SM_SHORTNAME"))
            X4.Value(lngSalesItemCount, 4) = FNS(rs.Fields("EXCH_TYPE"))
            X4.Value(lngSalesItemCount, 5) = FNS(rs.Fields("Code"))
            X4.Value(lngSalesItemCount, 6) = FNN(rs.Fields("CSL_Qty"))
            X4.Value(lngSalesItemCount, 7) = FNS(rs.Fields("TITLE")) & IIf(FNN(rs.Fields("EXCH_Voids")) > 0, " (Voids:" & FNN(rs.Fields("EXCH_Voids")) & ")", "")
            X4.Value(lngSalesItemCount, 8) = IIf(FNS(rs.Fields("EXCH_TYPE")) = "D", "", Format(rs.Fields("PRICE"), "Currency"))
            X4.Value(lngSalesItemCount, 9) = Format(rs.Fields("DiscountedValueIncVAT"), "Currency")
            X4.Value(lngSalesItemCount, 10) = FNS(rs.Fields("EXCH_ID"))
            X4.Value(lngSalesItemCount, 11) = FNS(rs.Fields("P_ID"))
            X4.Value(lngSalesItemCount, 12) = FNN(rs.Fields("EXCH_Voided"))
            X4.Value(lngSalesItemCount, 13) = FNN(rs.Fields("EXCH_Voids"))
        rs.MoveNext
    Loop
    X4.QuickSort 1, X4.UpperBound(1), 1, XORDER_DESCEND, XTYPE_NUMBER
    G4.Array = X4
    G4.ReBind
    G4.Bookmark = 1
    oPC.CloseLocalDatabase
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadExchanges"
End Sub

Private Sub PrintTotals(eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim sBuf As String
Dim sExt As String
Dim sValue As String
Dim oPmt As a_Payment

    Select Case eDocumentType
    Case eTypReceipt, eTypApproReturn
        If oPC.UseA4Printer Then
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                If oExchange.DiscountRate > 0 Then
                    sBuf = "Discount - " & oExchange.DiscountRateF
                    sExt = oExchange.TotalDiscountF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
                    sBuf = "Discounted subtotal"
                    sExt = oExchange.TotalLessDiscExtF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
                End If
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString(iColWidth, sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                Printer.Print sValue
                    sBuf = "Total received"
                    sExt = oExchange.TotalPaymentF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
                    For Each oPmt In oExchange.PaymentLines
                        If oPmt.Amt <> 0 Then
                            If oPmt.PaymentType = "V" Then
                                sBuf = oPmt.PaymentTypeF & " " & oPmt.ReferenceComplete
                            Else
                                sBuf = oPmt.PaymentTypeF
                            End If
                            sExt = oPmt.AmtF
                            sValue = MakePrintString(iColWidth, sBuf, sExt)
                            Printer.Print sValue
                        End If
                    Next
                    sBuf = "Change given"
                    sExt = oExchange.ChangeGivenF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
        Else
            If oPC.UseA4Printer Then
                    sBuf = "subtotal"
                    sExt = oExchange.TotalPayableF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
                    If oExchange.DiscountRate > 0 Then
                        sBuf = "Discount - " & oExchange.DiscountRateF
                        sExt = oExchange.TotalDiscountF
                        sValue = MakePrintString(iColWidth, sBuf, sExt)
                        Printer.Print sValue
                        sBuf = "Discounted subtotal"
                        sExt = oExchange.TotalLessDiscExtF
                        sValue = MakePrintString(iColWidth, sBuf, sExt)
                        Printer.Print sValue
                    End If
                    sBuf = "includes V.A.T."
                    sExt = oExchange.TotalVATF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
                    sBuf = "Total"
                    sExt = oExchange.TotalLessDiscExtF
                    sValue = MakePrintString((iColWidth \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                    Printer.Print sValue
                    sBuf = "Total received"
                    sExt = oExchange.TotalPaymentF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
                    For Each oPmt In oExchange.PaymentLines
                        If oPmt.Amt <> 0 Then
                            If oPmt.PaymentType = "V" Then
                                sBuf = oPmt.PaymentTypeF & " " & oPmt.ReferenceComplete
                            Else
                                sBuf = oPmt.PaymentTypeF
                            End If
                            sExt = oPmt.AmtF
                            sValue = MakePrintString(iColWidth, sBuf, sExt)
                            Printer.Print sValue
                        End If
                    Next
                    sBuf = "Change given"
                    sExt = oExchange.ChangeGivenF
                    sValue = MakePrintString(iColWidth, sBuf, sExt)
                    Printer.Print sValue
            Else
                With pPrinter
                    sBuf = "subtotal"
                    sExt = oExchange.TotalPayableF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    .PrintNormal PTR_S_RECEIPT, ESC + "|bC" + sValue + vbLf
                    If oExchange.DiscountRate > 0 Then
                        sBuf = "Discount - " & oExchange.DiscountRateF
                        sExt = oExchange.TotalDiscountF
                        sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                        .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sValue + vbLf
                        sBuf = "Discounted subtotal"
                        sExt = oExchange.TotalLessDiscExtF
                        sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                        .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sValue + vbLf
                    End If
                    sBuf = "includes V.A.T."
                    sExt = oExchange.TotalVATF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sValue + vbLf
                    sBuf = "Total"
                    sExt = oExchange.TotalLessDiscExtF
                    sValue = MakePrintString((.RecLineChars \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + sValue + vbLf
                    sBuf = "Total received"
                    sExt = oExchange.TotalPaymentF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sValue + vbLf
                    For Each oPmt In oExchange.PaymentLines
                        If oPmt.Amt <> 0 Then
                            If oPmt.PaymentType = "V" Then
                                sBuf = oPmt.PaymentTypeF & " " & oPmt.ReferenceComplete
                            Else
                                sBuf = oPmt.PaymentTypeF
                            End If
                            sExt = oPmt.AmtF
                            sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                            .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                        End If
                    Next
                    sBuf = "Change given"
                    sExt = oExchange.ChangeGivenF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                End With
            End If
        End If
    Case eTypCashRefund
        If oPC.UseA4Printer Then
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((iColWidth \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                Printer.Print sValue
                
                If oExchange.PaymentLines(1).PaymentType = "A" Then
                    sBuf = "Refund to credit card"
                Else
                    sBuf = "Cash refund"
                End If
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
        Else
            With pPrinter
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|bC" + sValue + vbLf
                
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|uC" + sValue + vbLf
                
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((.RecLineChars \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + sValue + vbLf
                
                If oExchange.PaymentLines(1).PaymentType = "A" Then
                    sBuf = "Refund to credit card"
                Else
                    sBuf = "Cash refund"
                End If
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
            End With
        End If
    Case eTypAppro
        If oPC.UseA4Printer Then
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((iColWidth \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                Printer.Print sValue
        Else
            With pPrinter
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|bC" + sValue + vbLf
                
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|uC" + sValue + vbLf
                
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((.RecLineChars \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + sValue + vbLf
                
                
            End With
        End If
    Case etypCreditVoucher
        If oPC.UseA4Printer Then
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((iColWidth \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                Printer.Print sValue
                sBuf = "Credit voucher"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
        Else
            With pPrinter
                sBuf = "subtotal"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|bC" + sValue + vbLf
                
                sBuf = "includes V.A.T."
                sExt = oExchange.TotalVATF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|uC" + sValue + vbLf
                
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((.RecLineChars \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + sValue + vbLf
                
                
                sBuf = "Credit voucher"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
            End With
        End If
    Case eTypOrder
        If oPC.UseA4Printer Then
                sBuf = "Deposit received"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                sBuf = "Total"
                sExt = oExchange.TotalLessDiscExtF
                sValue = MakePrintString((iColWidth \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                Printer.Print sValue
                sBuf = "Total"
                sExt = oExchange.TotalPaymentF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
                For Each oPmt In oExchange.PaymentLines
                    If oPmt.Amt <> 0 Then
                        If oPmt.PaymentType = "V" Then
                            sBuf = oPmt.PaymentTypeF & " " & oPmt.ReferenceComplete
                        Else
                            sBuf = oPmt.PaymentTypeF
                        End If
                        sExt = oPmt.AmtF
                        sValue = MakePrintString(iColWidth, sBuf, sExt)
                        Printer.Print sValue
                    End If
                Next
                sBuf = "Change given"
                sExt = oExchange.ChangeGivenF
                sValue = MakePrintString(iColWidth, sBuf, sExt)
                Printer.Print sValue
        Else
            With pPrinter
                sBuf = "Deposit received"
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|bC" + sValue + vbLf
                sBuf = "Total"
               '''TESTING 24/11/2009 sExt = oExchange.TotalLessDiscExtF
                sExt = oExchange.TotalPayableF
                
                sValue = MakePrintString((.RecLineChars \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
                .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + sValue + vbLf
                sBuf = "Total"
                '''TESTING 24/11/2009 sExt = oExchange.TotalPaymentF
                sExt = oExchange.TotalPayableF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sValue + vbLf
                For Each oPmt In oExchange.PaymentLines
                    If oPmt.Amt <> 0 Then
                        If oPmt.PaymentType = "V" Then
                            sBuf = oPmt.PaymentTypeF & " " & oPmt.ReferenceComplete
                        Else
                            sBuf = oPmt.PaymentTypeF
                        End If
                        sExt = oPmt.AmtF
                        sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                        .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                    End If
                Next
                sBuf = "Change given"
                sExt = oExchange.ChangeGivenF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
            End With
        End If
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintTotals(eDocumentType,pPrinter)", Array(eDocumentType, pPrinter)
End Sub

Private Sub PrintHeader(eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim fDate As String
Dim ar() As String
Dim i As Integer

    If oPC.UseA4Printer Then
        Printer.Print ""
        Printer.Print ""
    End If
    
    Select Case eDocumentType
    Case eTypReceipt, eTypCreditNote
        If oPC.UseA4Printer Then
                Printer.Print oPC.POSCompanyName
                If eDocumentType = eTypReceipt Then
                        Printer.Print "TAX INVOICE"
                        Printer.Print ""
                Else
                        Printer.Print "TAX CREDIT NOTE"
                        Printer.Print ""
                End If
                        Printer.Print oPC.POSBranchName
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                        Printer.Print ar(i)
                Next i
                Printer.Print ""
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                        Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
                        Printer.Print fDate
                        Printer.Print ""
                If oExchange.Customer.Name > "" Then
                        Printer.Print oExchange.Customer.NameAndCodeandType(50)
                ElseIf oExchange.Note > "" Then
                        Printer.Print Left(oExchange.Note, (50))
                End If
                If bReprint = True Then
                        Printer.Print "REPRINT"
                End If
        Else
            With pPrinter
                .AsyncMode = True
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
                .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If eDocumentType = eTypReceipt Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX INVOICE" + vbLf
                Else
                    .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX CREDIT NOTE" + vbLf
                End If
                .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
                If bReprint = True Then
                      .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                      .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                End If
            End With
        End If
    Case eTypApproReturn
        If oPC.UseA4Printer Then
                Printer.Print oPC.POSCompanyName
                Printer.Print ""
                Printer.Print "TAX INVOICE (from Appro)"
                Printer.Print ""
                Printer.Print oPC.POSBranchName
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                    Printer.Print ar(i)
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                Printer.Print ""
                Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
                Printer.Print fDate + vbLf
                Printer.Print ""
                
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, (iColWidth))
                End If
    
                If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
                End If
        Else
            With pPrinter
                .AsyncMode = True
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
                .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX INVOICE (from Appro)" + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
    
                If bReprint = True Then
                      .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                      .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                End If
            End With
        End If
    Case eAppro
        If oPC.UseA4Printer Then
                Printer.Print oPC.POSCompanyName
                Printer.Print ""
                Printer.Print "APPRO OUT"
                Printer.Print ""
                Printer.Print oPC.POSBranchName
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                    Printer.Print ar(i)
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                Printer.Print ""
                Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
                Printer.Print fDate
                Printer.Print ""
                
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, (iColWidth))
                End If
    
                If bReprint = True Then
                      Printer.Print "REPRINT"
                      Printer.Print ""
                End If
        Else
            With pPrinter
                .AsyncMode = True
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
                .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "APPRO OUT" + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
    
                If bReprint = True Then
                      .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                      .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                End If
            End With
        End If
     Case eTypPaymentReceipt
        If oPC.UseA4Printer Then
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "PAYMENT ACCEPTED"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
              Printer.Print fDate
              Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, iColWidth)
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "PAYMENT ACCEPTED" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, .RecLineChars) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypCashRefund
        If oPC.UseA4Printer Then
              Printer.Print ""
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "CASH REFUND"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
              Printer.Print fDate + vbLf
              Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, iColWidth)
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CASH REFUND" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, .RecLineChars) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypChangeVoucher
        If oPC.UseA4Printer Then
            Printer.Print ""
            Printer.Print oPC.POSCompanyName
            Printer.Print ""
            Printer.Print "CHANGE VOUCHER"
            Printer.Print ""
            Printer.Print oPC.POSBranchName
            ar = Split(oPC.POSBranchAddress, ",")
            For i = 0 To UBound(ar)
              Printer.Print ar(i)
            Next i
            fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
            Printer.Print ""
            Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
            Printer.Print fDate
            Printer.Print ""
            If oExchange.Customer.Name > "" Then
                Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
            ElseIf oExchange.Note > "" Then
                Printer.Print Left(oExchange.Note, (iColWidth))
            End If
            If bReprint = True Then
                Printer.Print "REPRINT"
                Printer.Print ""
            End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CHANGE VOUCHER" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case etypCreditVoucher
        If oPC.UseA4Printer Then
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "CREDIT VOUCHER"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName
              Printer.Print fDate
              Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, (iColWidth))
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CREDIT VOUCHER" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypDepositRefund
        If oPC.UseA4Printer Then
              Printer.Print ""
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "DEPOSIT REFUND"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
              fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              Printer.Print fDate
              Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, (iColWidth))
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "DEPOSIT REFUND" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypDeposit
        If oPC.UseA4Printer Then
              Printer.Print ""
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "DEPOSIT PAID"
              Printer.Print
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
              Printer.Print fDate
              Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                     Printer.Print Left(oExchange.Note, (iColWidth))
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "DEPOSIT PAID" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars)
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars))
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT"
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypVoucher
        If oPC.UseA4Printer Then
            Printer.Print ""
            Printer.Print oPC.POSCompanyName
            Printer.Print ""
            Printer.Print "LOYALTY CLUB CREDIT VOUCHER"
            Printer.Print ""
            Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                    Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
            Printer.Print ""
            Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
            Printer.Print fDate
            Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, (iColWidth))
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "LOYALTY CLUB CREDIT VOUCHER" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypAppro
        If oPC.UseA4Printer Then
              Printer.Print ""
              Printer.Print oPC.POSCompanyName + vbLf
              Printer.Print ""
              Printer.Print "APPRO"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
              Printer.Print fDate
              Printer.Print ""
                If oExchange.Customer.Name > "" Then
                    Printer.Print oExchange.Customer.NameAndCodeandType(iColWidth)
                ElseIf oExchange.Note > "" Then
                    Printer.Print Left(oExchange.Note, (iColWidth))
                End If
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "APPRO" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
        End With
        End If
    Case eTypPettyCash
        If oPC.UseA4Printer Then
              Printer.Print ""
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "PETTY CASH"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
              Printer.Print fDate
              Printer.Print ""
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "PETTY CASH" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    Case eTypOrder
        If oPC.UseA4Printer Then
              Printer.Print ""
              Printer.Print oPC.POSCompanyName
              Printer.Print ""
              Printer.Print "CUSTOMER ORDER"
              Printer.Print ""
              Printer.Print oPC.POSBranchName
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
                Printer.Print ar(i)
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              Printer.Print ""
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName
              Printer.Print fDate
              Printer.Print ""
              If bReprint = True Then
                    Printer.Print "REPRINT"
                    Printer.Print ""
              End If
        Else
            With pPrinter
              .AsyncMode = True
              .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
              .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CUSTOMER ORDER" + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & strName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
              End If
            End With
        End If
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintHeader(eDocumentType,pPrinter,bReprint)", Array(eDocumentType, pPrinter, _
         bReprint)
End Sub
Private Sub PrintFooter(pCopyNumber As Integer, eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter)
        On Error GoTo errHandler
Dim ar() As String
Dim i As Integer
Dim sValue As String
    Select Case eDocumentType
    Case eTypReceipt, eTypCashRefund, etypCreditVoucher, eTypDeposit, eTypCreditNote, eTypDepositRefund, eTypPettyCash, eTypAppro, eTypApproReturn, eTypPaymentReceipt, eTypOrder
        If oPC.UseA4Printer = "TRUE" Then
                            Printer.Print ""
                If oExchange.DiscountRate > 0 Then
                    sValue = MakePrintString(50, "List" & oExchange.TotalLessDiscExtF & " Sell" & oExchange.TotalLessDiscExtF & "Your savings" & oExchange.TotalDiscountF, "")
                    Printer.Print sValue
                End If
                ar = Split(oPC.POSReceiptMessage, ",")
                For i = 0 To UBound(ar)
 '                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                    Printer.Print ar(i)
                Next i
               ' .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSEMAILAddress + vbLf
                Printer.Print oPC.POSEMAILAddress
                If pCopyNumber > 1 Then
                  '  .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "Copy number: " & CStr(pCopyNumber) + vbLf
                    Printer.Print "Copy number: " & CStr(pCopyNumber)
                End If
        Else
            With pPrinter
                .AsyncMode = True
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
                .PrintNormal PTR_S_RECEIPT, ESC + "|700uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
                If oExchange.Customer.ID > 0 Then
                End If
                If oExchange.DiscountRate > 0 Then
                    sValue = MakePrintString(.RecLineChars, "List" & oExchange.TotalLessDiscExtF & " Sell" & oExchange.TotalLessDiscExtF & "Your savings" & oExchange.TotalDiscountF, "")
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                End If
                ar = Split(oPC.POSReceiptMessage, ",")
                For i = 0 To UBound(ar)
                  .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSEMAILAddress + vbLf
                If pCopyNumber > 1 Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "Copy number: " & CStr(pCopyNumber) + vbLf
                End If
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
            End With
        End If
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintFooter(pCopyNumber,eDocumentType,pPrinter)", Array(pCopyNumber, _
         eDocumentType, pPrinter)
End Sub
Private Function ConvertToType(val As String) As Integer
    On Error GoTo errHandler
    Select Case UCase(val)
    Case "S", "A"
        ConvertToType = eTypReceipt
    Case "R"
        ConvertToType = eTypCashRefund
    Case "C"
        ConvertToType = etypCreditVoucher
    Case "D"
        ConvertToType = eTypDeposit
    Case "APP"
        ConvertToType = eTypAppro
    Case "PC"
        ConvertToType = eTypPettyCash
    Case "PCC"
        ConvertToType = eTypPettyCashCredit
    Case "DEP"
        ConvertToType = eTypDeposit
    Case "AR"
        ConvertToType = eTypApproReturn
    Case "OR"
        ConvertToType = eTypOrder
    Case "CV"
        ConvertToType = eTypChangeVoucher
    Case "PA"
        ConvertToType = eTypPaymentReceipt
    Case "CN"
        ConvertToType = eTypCreditNote
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ConvertToType(val)", val
End Function

Private Function CollectPettyCashArray() As String()
    On Error GoTo errHandler
Dim i As Integer
Dim j As Integer
Dim arPC() As String

    ReDim Preserve arPC(0)
    For i = X4.LowerBound(1) To X4.UpperBound(1)
        If X4(i, 4) = "PC" And X4(i, 12) <> 1 Then ' it is a petty cash exchange
            j = j + 1
            ReDim Preserve arPC(j)
            arPC(j) = X4(i, 10) & "|" & X4(i, 1) & "|" & X4(i, 2) & "|" & X4(i, 9) & "|" & X4(i, 7)
        End If
    Next
    CollectPettyCashArray = arPC
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CollectPettyCashArray())"
End Function
'Private Sub Command1_Click()
'    On Error GoTo errHandler
'    fixed
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Command1_Click", , EA_NORERAISE
'    HandleError
'End Sub
'Public Sub fixed()
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    rs.Open "SELECT * FROM vFix", oPC.DBLocalConn, adOpenStatic, adLockOptimistic
'    Do While Not rs.EOF
'        If MsgBox("Number " & rs.Fields("EXCH_NUMBER"), vbQuestion + vbYesNo, "Check") = vbNo Then
'            Exit Do
'        End If
'        SendPOSExchange FNS(rs.Fields("EXCH_ID")), FNS(rs.Fields("OPS_ID")), FNS(rs.Fields("OPS_Z_ID"))
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set rs = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.fixed"
'End Sub

Public Sub SendPOSExchange(pEXCHID As String, pZID As String)  'pOPSID As String,
    On Error GoTo errHandler
Dim msg As String
Dim sFileName As String
Dim i As Integer
Dim strExchangeMsg As String
Dim strCSLPart As String
Dim strPayPart As String
Dim strExchangeType As String

    If Not rsZSession Is Nothing Then
        If rsZSession.State <> 0 Then
            rsZSession.Close
        End If
    End If
'================================================================================================================================
    strExchangeType = "E"
    For i = 1 To oExchange.PaymentLines.Count
        If oExchange.PaymentLines(i).PaymentType = "AC" Then  'A purchase has been placed on account (even if partially)
            strExchangeType = "A"
        End If
    Next i
    If oExchange.transactionType = "PA" Then  'It is a payment made to an account
        strExchangeType = "P"
    End If
    If oExchange.transactionType = "CN" Then  'It is a credit made to an account
        strExchangeType = "CN"
    End If
    strExchangeMsg = strExchangeType & vbTab & pZID & vbTab & oPC.StationName & vbTab & ReverseDateTime(oPC.ZSession.FromDate) & vbTab & ReverseDateTime(oPC.ZSession.EndDate) & vbTab _
    & ReverseDateTime(oPC.ZSession.NominalDate) & vbTab & oPC.ZSession.OpSession.OPSID & vbTab & ReverseDateTime(oPC.ZSession.OpSession.DateStarted) & vbTab _
    & ReverseDateTime(oPC.ZSession.OpSession.DateEnded) & vbTab & oPC.ZSession.OpSession.OperatorID & vbTab & oPC.ZSession.OpSession.SupervisorID & vbTab _
    & oExchange.ExchangeID & vbTab & ReverseDateTime(oExchange.ExchangeDate) & vbTab & oExchange.OperatorID & vbTab & oExchange.ExchangeNumber & vbTab _
    & oExchange.TotalPayable & vbTab & oExchange.TotalDiscount & vbTab & oExchange.TotalVAT & vbTab _
    & oExchange.ChangeGiven & vbTab & oExchange.LoyaltyValue & vbTab & oExchange.transactionType & vbTab _
    & oExchange.Note & vbTab & oExchange.ToVoid & vbTab & oExchange.Customer.CustomerID & "|"
'MsgBox "RDT = " & ReverseDateTime(oExchange.ExchangeDate)
    For i = 1 To oExchange.SaleLines.Count
        strCSLPart = oExchange.SaleLines(i).PID & vbTab & oExchange.SaleLines(i).COLID & vbTab & oExchange.SaleLines(i).Qty & vbTab _
        & oExchange.SaleLines(i).Price & vbTab & oExchange.SaleLines(i).PriceAlteration & vbTab & oExchange.SaleLines(i).PDiscExt & vbTab _
        & oExchange.SaleLines(i).DiscountRate & vbTab & oExchange.SaleLines(i).VATRate & vbTab & oExchange.SaleLines(i).Counterfoil & vbTab _
        & oExchange.SaleLines(i).DiscountDescription
        strExchangeMsg = strExchangeMsg & strCSLPart & IIf(i = oExchange.SaleLines.Count, "", "~")
    Next i
    strExchangeMsg = strExchangeMsg & "|"
    
    For i = 1 To oExchange.PaymentLines.Count
        strPayPart = oExchange.PaymentLines(i).PaymentType & vbTab & oExchange.PaymentLines(i).Amt & vbTab _
        & oExchange.PaymentLines(i).ReferenceComplete & vbTab & oExchange.PaymentLines(i).Note & vbTab & oExchange.PaymentLines(i).COLID
        strExchangeMsg = strExchangeMsg & strPayPart & IIf(i = oExchange.PaymentLines.Count, "", "~")
    Next i
                    'Place  POS MESSAGE in queue for server to fetch
    DispatchMessageEx strExchangeMsg
    
    lblProg.Caption = lblProg.Caption & "X"
    If pEXCHID > "" Then
        SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
        oPC.OpenLocalDatabase
        oPC.DBLocalConn.Execute SQL
        oPC.CloseLocalDatabase
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SendPOSExchange(pEXCHID,pZID)", Array(pEXCHID, pZID)
End Sub

Private Sub ReSendExchanges()
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strEXCHNUM As String
Dim sr As New Scripting.FileSystemObject
Dim strFileName As String
Dim iLoc As Integer
Dim strStart As String
Dim strEnd As String
Dim iStart As Long
Dim iEnd As Long
Dim bErr As Boolean
Dim i As Long

    strFileName = oPC.LocalRootFolder & "\RESEND.TXT"
    If sr.FileExists(strFileName) Then
        oTF.OpenTextFileToRead strFileName
        If Not oTF.IsEOF Then
            If MsgBox("There are exchanges to re-send. Continue?", vbYesNo + vbQuestion, "Warning") <> vbYes Then
                oTF.CloseTextFile
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Do While Not oTF.IsEOF
        strEXCHNUM = oTF.ReadLinefromTextFile
        If IsNumeric(strEXCHNUM) Then
            SendPOSExchange_ByExchangeNumber strEXCHNUM
        Else
            iLoc = InStr(1, strEXCHNUM, "-")
            If iLoc > 0 Then
                strStart = Left(strEXCHNUM, iLoc - 1)
                strEnd = Right(strEXCHNUM, Len(strEXCHNUM) - iLoc)
                If IsNumeric(strStart) Then
                    iStart = CLng(strStart)
                Else
                    bErr = True
                End If
                If IsNumeric(strEnd) Then
                    iEnd = CLng(strEnd)
                Else
                    bErr = True
                End If
                If Not bErr Then
                    For i = iStart To iEnd
                        SendPOSExchange_ByExchangeNumber CStr(i)
                    Next i
                End If
            End If
        End If
    Loop
        
    oTF.CloseTextFile
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ReSendExchanges"
End Sub
Public Sub SendPOSExchange_ByExchangeNumber(pEXCHNumber As String)
    On Error GoTo errHandler
Dim msg As String
Dim sFileName As String
Dim oShapeDB As New z_POSCLIConnectionShape
Dim sSQL As String
Dim tmprs As ADODB.Stream
Dim strPos As String
Dim rs As ADODB.Recordset
Dim pZID As String
Dim pOPSID As String
Dim pEXCHID As String
Dim strExchangeMsg As String
Dim strCSLPart As String
Dim strPaymentPart As String
Dim strPayPart As String
Dim strTyp As String

    oPC.OpenLocalDatabase
    Set rs = New ADODB.Recordset
    rs.Open "SELECT     EXCH_ID, EXCH_ZSessionID, EXCH_OPSESSIONID, EXCH_Voided, EXCH_VOIDS FROM dbo.tExchange WHERE EXCH_NUMBER = " & pEXCHNumber, oPC.DBLocalConn, adOpenStatic
    If Not rs.EOF Then
        pZID = rs.Fields("EXCH_ZSessionID")
        pOPSID = rs.Fields("EXCH_OPSESSIONID")
        pEXCHID = rs.Fields("EXCH_ID")
    Else
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
    oPC.CloseLocalDatabase
    Check (oShapeDB.dbConnecttoShape = 0), ERR_GENERAL, "Failed to create database connection!"
    If Not rsZSession Is Nothing Then
        If rsZSession.State <> 0 Then
            rsZSession.Close
        End If
    End If

    sSQL = "SHAPE {SELECT tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession " _
        & " APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession " _
        & " APPEND (( SHAPE {SELECT EXCH_TYPE as TYP,EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE, " _
        & " EXCH_SALEVALUE,EXCH_DISCOUNTVALUE,EXCH_VATVALUE,EXCH_CHANGEGIVEN,EXCH_LOYALTYVALUE,EXCH_TYPE,EXCH_OPERATORID,EXCH_SUPERVISORID,EXCH_NUMBER,EXCH_VOIDS,EXCH_NOTE " _
        & " FROM tEXCHANGE WHERE EXCH_ID = '" & pEXCHID & "'}  AS POSExchange " _
        & " APPEND ({SELECT * FROM tCSL}  AS rsSALESLINES " _
        & " RELATE EXCH_ID TO CSL_EXCH_ID) AS SALESLINES,({SELECT * FROM tPayment}  AS rsPAYMENTS " _
        & " RELATE EXCH_ID TO PAY_EXCH_ID) AS PAYMENTS) AS POSExchange " _
        & " RELATE OPS_ID TO EXCH_OPSESSIONID) AS POSExchange) AS OPSession RELATE Z_ID TO OPS_Z_ID) AS OPSession"
    Set rsZSession = Nothing
    
    Set rsZSession = New ADODB.Recordset
    
    rsZSession.CursorLocation = adUseClient
    rsZSession.Properties("Release Shape On Disconnect") = True
    
    rsZSession.Open sSQL, oShapeDB.DBConn, adOpenStatic
    DisconnectAll rsZSession
    
''''''''''''''''''''''''''''''
Dim rsOP As ADODB.Recordset
Dim rsEx As ADODB.Recordset
Dim rsCSL As ADODB.Recordset
Dim rsPAY As ADODB.Recordset
    Set rsOP = rsZSession.Fields("OPSESSION").Value
    Set rsEx = rsOP.Fields("POSExchange").Value
    Set rsCSL = rsEx.Fields("SALESLINES").Value
    Set rsPAY = rsEx.Fields("PAYMENTS").Value
    

    strTyp = "E"
    Do While Not rsPAY.EOF
        If rsPAY.Fields("PAY_PaymentType") = "AC" Then
            strTyp = "A"
        End If
        rsPAY.MoveNext
    Loop
    If rsPAY.RecordCount > 0 Then rsPAY.MoveFirst
    If rsEx!Typ = "PA" Then
        strTyp = "P"
    ElseIf rsEx!Typ = "CN" Then
        strTyp = "CN"
    End If
    strExchangeMsg = strTyp & vbTab & rsZSession!Z_ID & vbTab & rsZSession!Z_TILLPOINT & vbTab & ReverseDateTime(rsZSession!Z_STARTDATE) & vbTab _
    & ReverseDateTime(FND(rsZSession!Z_ENDDATE)) & vbTab _
    & ReverseDateTime(rsZSession!Z_NOMINALDATE) & vbTab & rsZSession.Fields("OPSession")!OPS_ID & vbTab & ReverseDateTime(CDate(rsZSession.Fields("OPSession")!OPS_STARTTIME)) & vbTab _
    & ReverseDateTime(rsOP!OPS_endtime) & vbTab & rsOP!OPS_OPERATORID & vbTab & rsOP!OPS_OPERATORID & vbTab _
    & rsEx!EXCH_ID & vbTab & ReverseDateTime(rsEx!EXCH_SaleDate) & vbTab _
    & rsEx!EXCH_OperatorID & vbTab & rsEx!EXCH_Number & vbTab _
    & rsEx!EXCH_SaleValue & vbTab & rsEx!EXCH_DiscountValue & vbTab & rsEx!EXCH_VATValue & vbTab _
    & rsEx!EXCH_ChangeGiven & vbTab & rsEx!EXCH_LoyaltyValue & vbTab & rsEx!EXCH_TYPE & vbTab _
    & rsEx!EXCH_Note & vbTab & rsEx!EXCH_VOIDS & vbTab & rsEx!EXCH_TP_ID & "|"

    Do While rsCSL.EOF = False
        strCSLPart = rsCSL!CSL_P_ID & vbTab & rsCSL!CSL_COLID & vbTab & rsCSL!CSL_Qty & vbTab _
        & rsCSL!CSL_Price & vbTab & rsCSL!CSL_PriceAlteration & vbTab & rsCSL!CSL_Discount & vbTab _
        & rsCSL!CSL_DiscountRate & vbTab & rsCSL!CSL_VATRATE & vbTab & rsCSL!CSL_Counterfoil & vbTab _
        & rsCSL!CSL_DiscountDescription
        rsCSL.MoveNext
        strExchangeMsg = strExchangeMsg & strCSLPart & IIf(rsCSL.EOF, "", "~")
    Loop
    strExchangeMsg = strExchangeMsg & "|"

    Do While rsPAY.EOF = False
        strPayPart = rsPAY!PAY_PaymentType & vbTab & rsPAY!PAY_Amt & vbTab _
        & rsPAY!PAY_Ref & vbTab & rsPAY!PAY_Note & vbTab & rsPAY!PAY_COLID
        rsPAY.MoveNext
        strExchangeMsg = strExchangeMsg & strPayPart & IIf(rsPAY.EOF, "", "~")
    Loop
                    'Place  POS MESSAGE in queue for server to fetch
    DispatchMessageEx strExchangeMsg
    
    oShapeDB.dbCloseConnectShape
''''''''''''''''''''''''''''''
    oPC.OpenLocalDatabase
    lblProg.Caption = lblProg.Caption & "X"
    If pEXCHID > "" Then
        SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
        oPC.DBLocalConn.Execute SQL
    End If
    oPC.CloseLocalDatabase
    'oShapeDB.dbCloseConnectShape
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SendPOSExchange(pEXCHID,pOPSID,pZID)", Array(pEXCHID, pOPSID, pZID)
End Sub

Private Function DisconnectAll(rs As ADODB.Recordset)
    Dim i As Long
With rs
   Set .ActiveConnection = Nothing
   For i = 0 To rs.Fields.Count - 1
      If (rs.Fields(i).Type = adChapter) Then
         DisconnectAll rs.Fields(i).Value
      End If
   Next i
End With
End Function
Private Sub DispatchMessageEx(pMsgXML As String)
    On Error GoTo errHandler
Dim oStream As ADODB.Stream
'Dim oTR As MSMQTransaction

    Set QI = New MSMQQueueInfo
   ' QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\qpos"
    QI.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\qpos"
  '  QI.FormatName = "DIRECT=OS:DAVID-LT\Private$\qpos"
    Set QPOS = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    
    QI.FormatName = "DIRECT=OS:" & oPC.NameOfPC & "\Private$\qposack"
    
    Set POSmsg = New MSMQMessage
    POSmsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
    POSmsg.Journal = MQMSG_DEADLETTER
    POSmsg.MaxTimeToReachQueue = oPC.POSMessageTimeout  '9 days
    
    POSmsg.Label = "POSN," & Format(Now, "dd/mm/yyyy HH:NN")

    POSmsg.Body = pMsgXML

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''This must be removed when no longer necessary
 '   POSmsg.Journal = MQMSG_JOURNAL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set POSmsg.ResponseQueueInfo = QI
    POSmsg.Send QPOS  ', MQ_SINGLE_MESSAGE ', oTR

    Exit Sub
errHandler:
    ErrorIn "frmPOSMain.DispatchMessageEx(pMsgXML)", pMsgXML
End Sub

Private Sub DispatchMessage(rs As ADODB.Recordset)
Dim oStream As ADODB.Stream
'Dim oTR As MSMQTransaction

    On Error GoTo errHandler
    Set QI = New MSMQQueueInfo
    'QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
    QI.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
    Set QPOS = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    
    QI.FormatName = "DIRECT=OS:" & oPC.NameOfPC & "\Private$\qposack"
    
    Set POSmsg = New MSMQMessage
    POSmsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
    POSmsg.Journal = MQMSG_DEADLETTER
    POSmsg.MaxTimeToReachQueue = oPC.POSMessageTimeout
    
    POSmsg.Label = "POSN," & Format(Now, "dd/mm/yyyy HH:NN")
    Set oStream = New ADODB.Stream
    rs.Save oStream, adPersistXML
    POSmsg.Body = oStream.ReadText

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''This must be removed when no longer necessary
    POSmsg.Journal = MQMSG_JOURNAL
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Set POSmsg.ResponseQueueInfo = QI
    
    POSmsg.Send QPOS  ', MQ_SINGLE_MESSAGE ', oTR

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DispatchMessage(rs)", rs
End Sub
Private Sub PreparePaymentLine(ePaymentMode As enPaymentMode)
    On Error GoTo errHandler
    Set oPAYMENTLine = oExchange.PaymentLines.Add
    oPAYMENTLine.ApplyEdit
    oPAYMENTLine.BeginEdit
    iCurrentPaymentLine = iCurrentPaymentLine + 1
    X2.ReDim 1, iCurrentPaymentLine, 1, 3
    oPAYMENTLine.SetType ConvertPaymentStateToCode(ePaymentMode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PreparePaymentLine(ePaymentMode)"
End Sub
Private Function ConvertPaymentStateToCode(ePaymentMode As enPaymentMode) As String
    On Error GoTo errHandler
    Select Case ePaymentMode
    Case ePaymentMode_Cheque
        ConvertPaymentStateToCode = "Q"
    Case ePaymentMode_CreditCard
        ConvertPaymentStateToCode = "A"
    Case ePaymentMode_Voucher
        ConvertPaymentStateToCode = "V"
    Case ePaymentMode_RedeemedDeposit
        ConvertPaymentStateToCode = "RD"
    Case ePaymentMode_CreditVoucher
        ConvertPaymentStateToCode = "CV"
    Case ePaymentMode_Cash
        ConvertPaymentStateToCode = "C"
    Case ePaymentMode_Account
        ConvertPaymentStateToCode = "AC"
    Case ePaymentMode_DIrectDeposit
        ConvertPaymentStateToCode = "DDP"
        
    End Select
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ConvertPaymentStateToCode(ePaymentMode)"
End Function

'=====================================================================================================


Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    bShiftDown = (Shift = 1)
    If KeyCode = 13 Then
        lblChange.Caption = ""
        enNewState = GetNewState(txtInput)  ', itmp, strArg, strArg2
        
        If bValid Then SetPresentState enNewState
        If enNewState = eStart Then txtInput.BackColor = RGB(230, 250, 210)

     End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE, , "input value", Array(txtInput)
    HandleError
End Sub
'=====================================================================================================

Private Function GetNewState(txtIn As String) As eState
    On Error GoTo errHandler

    
    strRaw = UCase(Trim$(txtIn))
    GetNewState = eNull
    If LenB(strRaw) = 0 Then
        bValid = False
        Exit Function
    Else
        bValid = True
    End If
    If SeparateInput(strRaw, strPrefix, strSuffix) = False Then 'We cannot process this input
        GetNewState = enPresentState
        Exit Function
    End If
    txtInput.BackColor = vbWhite
    Select Case enPresentState
    Case eAppro
        Select Case strPrefix
        Case "X"
            GetNewState = Action_CancelSale
        Case "F"
            GetNewState = eConfirmation
        Case "D"
            GetNewState = Action_DeleteSaleLine(strSuffix)
        Case Else
          '  bValid = False
            GetNewState = Action_Appro
        End Select
    Case eCollect
            If enMode = eMode_ApproReturn And strPrefix = "X" Then
                    GetNewState = Action_CancelSale
            
            Else
                Select Case strPrefix
                Case "C"
                    GetNewState = ePaymentType_Cash
                Case "Q"
                    GetNewState = ePaymentType_Cheque
                Case "V"
                    GetNewState = ePaymentType_voucher
                Case "CC"
                    GetNewState = ePaymentType_CreditCard
                Case "CN"
                    GetNewState = ePaymentType_CreditVoucher
                Case "DDP"
                    GetNewState = ePaymentType_DirectDeposit
                Case "AC"
                    If enMode = emode_PayAccount Then
                        bValid = False
                    Else
                        GetNewState = ePaymentType_Account
                    End If
                Case "X"
                    GetNewState = Action_CancelSale
                Case Else
                    bValid = False
                End Select
            End If
       ' End If
    Case eConfirmation
        GetNewState = Action_Confirmation()
    Case eDiscount
        GetNewState = Action_Discount
    Case ePaymentType_Cash
        GetNewState = Action_PaymentType_Cash()
    Case ePaymentType_Cheque
        GetNewState = Action_PaymentType_Cheque()
    Case ePaymentType_ChequeRef
        GetNewState = Action_PaymentType_ChequeRef()
    Case ePaymentType_voucher
        GetNewState = Action_PaymentType_Voucher()
    Case ePaymentType_voucherRef
        GetNewState = Action_PaymentType_VoucherRef()
    Case ePaymentType_CreditCard
        GetNewState = Action_PaymentType_Creditcard()
    Case ePaymentType_CreditVoucherRef
        GetNewState = Action_PaymentType_CreditcardRef()
    Case ePaymentType_CreditVoucher
        GetNewState = Action_PaymentType_CreditVoucher()
    Case ePaymentType_CreditVoucherRef
        GetNewState = Action_PaymentType_CreditVoucherRef()
    Case ePaymentType_RedeemDeposit
        GetNewState = Action_PaymentType_RedeemDeposit()
    Case ePaymentType_DirectDeposit
        GetNewState = Action_PaymentType_DirectDeposit()
    Case ePaymentType_Account
        GetNewState = Action_PaymentType_Account()
    Case ePettyCash
        GetNewState = Action_PettyCash()
    Case ePettyCashCredit
        GetNewState = Action_PettyCashCredit()
    Case ePrice
        If strPrefix = ".." Then
            RemoveSaleLine , True
            If bSaleActive Then
                If enMode = emode_Sale Then
                    GetNewState = eSale
                ElseIf enMode = eMode_ApproReturn Then
                    GetNewState = eApproReturn
                ElseIf enMode = emode_Appro Then
                    GetNewState = eAppro
                End If
            Else
                GetNewState = eStart
            End If
        Else
            If Not IsNumeric(strRaw) Then
                GetNewState = ePrice
            Else
                GetNewState = Action_Price()
            End If
        End If
    Case eQty
        GetNewState = Action_Qty
    Case eSelectDepositLine
        If strPrefix = ".." Then
            GetNewState = eStart
        Else
            GetNewState = Action_SelectDepositLine()
        End If
    Case eSelectDepositLineForRefund
        If strPrefix = ".." Then
            GetNewState = eStart
        Else
            GetNewState = Action_SelectDepositLineForRefund()
        End If
    Case eStart
        Select Case strPrefix
        Case "REPRINT"
            GetNewState = Action_Reprint
        Case "QU"
            GetNewState = Action_LoadFromQuotation
        Case "XEND"
            GetNewState = Action_eXTerminate
        Case "ZEND"
            GetNewState = Action_eZTerminate
        Case "A"
            If Not SecurityControl(enSECURITY_ISSUEPOSAPPRO, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permissionto issue an appro.") Then
                GetNewState = eStart
            Else
                enMode = emode_Appro
                GetNewState = Action_SearchCustomer("A")
            End If
        Case "RDEP"
            If SecurityControl(enSECURITY_REFUNDDEPOSITONPOS, lngOPID, strName, , "Enter your signature.", "You do not have permission to refund deposits.") Then
                enMode = eMode_ReturnDeposit
                GetNewState = Action_SearchCustomer("R")
            Else
                GetNewState = eStart
            End If
        Case "AR"
            enMode = eMode_ApproReturn
            GetNewState = Action_SearchCustomer("I")
        Case "PA"
            If SecurityControl(enSECURITY_ACCEPTACPAYMENT, lngOPID, strName, , "Enter your signature.", "You do not have permission to accept payments.") Then
                    enMode = emode_PayAccount
                    GetNewState = Action_SearchCustomer("PA")
            Else
                GetNewState = eStart
            End If
        Case "DD"
            GetNewState = Action_ReviewExchanges
        Case "YY"
            GetNewState = Action_ReviewDeadLetterQueue
        Case "PC"
            GetNewState = Action_PettyCash
        Case "PCR"
            GetNewState = Action_PettyCashCredit
        Case "OP"
            GetNewState = Action_operatorsReport
        Case "OD"
            GetNewState = Action_OpenDrawer
        Case "V"
            GetNewState = Action_Void
        Case "OR"
            GetNewState = Action_OrderRequest
        Case "CNA"
            If SecurityControl(enSECURITY_ISSUEPOSCREDITNOTE, lngOPID, strName, , "Enter your signature.", "You do not have permission to issue credit notes.") Then
                enMode = emode_CreditNote
                GetNewState = Action_SearchCustomer("CN")
            Else
                GetNewState = eStart
            End If
        Case "SS"
            If bSaleOnHold Then
                MsgBox "There is already a sale saved, you should either cancel this sale or complete it before using RS to retrieve the saved sale and cancel or complete it before continuing", vbInformation + vbOKOnly, "Can't do this"
                bValid = False
            Else
                GetNewState = Action_SaveSale
            End If
        Case "RS"
            If bSaleOnHold = False Then
                MsgBox "There is no sale saved.", vbInformation + vbOKOnly, "Can't do this"
                bValid = False
            Else
                GetNewState = Action_RetrieveSale
            End If
        Case Else
            GetNewState = Action_Sale
        End Select
        
    Case eSale
        Select Case strPrefix
        Case ".."
            GetNewState = enPresentState  'DOnt go back from here
        Case "D"
            GetNewState = Action_DeleteSaleLine(strSuffix)
        Case "DP"
            GetNewState = Action_DeletePayment(strSuffix)
        Case "FC"
            GetNewState = Action_SearchCustomer("S")
        Case "SS"
            If bSaleOnHold Then
                MsgBox "There is already a sale saved, you should either cancel this sale or complete it before using RS to retrieve the saved sale and cancel or complete it before continuing", vbInformation + vbOKOnly, "Can't do this"
                bValid = False
            Else
                GetNewState = Action_SaveSale
            End If
        Case "RS"
            If bSaleOnHold = False Then
                MsgBox "There is no sale saved.", vbInformation + vbOKOnly, "Can't do this"
                bValid = False
            Else
                GetNewState = Action_RetrieveSale
            End If
        Case "X"
            GetNewState = Action_CancelSale
        Case Else
            If oExchange.transactionType = "RDEP" Or oExchange.TotalPayable < 0 = True Then    'This is a refund credit note
                    Select Case strPrefix
                    Case "C"
                        If Not SecurityControl(enSECURITY_ISSUEPOSREFUND, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission to issue a cash refund.") Then
                            GetNewState = eSale
                        Else
                            RefundPayment ePaymentMode_Cash
                            GetNewState = eConfirmation
                        End If
                    Case "CV"
                        RefundPayment ePaymentMode_CreditVoucher
                        GetNewState = eConfirmation
                    Case "CC"
                        RefundPayment ePaymentMode_CreditCard
                        GetNewState = eConfirmation
                    Case Else
                        GetNewState = Action_Sale
                    End Select
            Else
                Select Case strPrefix
                Case "C"
                    GetNewState = ePaymentType_Cash
                Case "CC"
                    GetNewState = ePaymentType_CreditCard
                Case "Q"
                    GetNewState = ePaymentType_Cheque
                Case "V"
                    GetNewState = ePaymentType_voucher
                Case "CC"
                    GetNewState = ePaymentType_CreditCard
                Case "DDP"
                    If Not SecurityControl(enSECURITY_ACCEPTDIRECTDEPOSIT, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission to pay with a direct deposit.") Then
                        GetNewState = eSale
                    Else
                        GetNewState = ePaymentType_DirectDeposit
                    End If
                    'GetNewState = eConfirmation
                    
                    
                    ''''''Removed because CV is just another voucher type - 'other'
            '    Case "CV"
            '        GetNewState = ePaymentType_CreditVoucher
            
            
            
                Case "AC"
                    GetNewState = ePaymentType_Account
                Case "DDP"
                    GetNewState = ePaymentType_DirectDeposit
'Commented 5/3/06 Deposits are redeemed after selecting a customer - with the order on the screen
'                Case "RD"
'                    If Valid_RedeemDeposit_Arg(strSuffix) Then
'                        GetNewState = ePaymentType_RedeemDeposit
'                    Else
'                        GetNewState = enPresentState
'                    End If
                Case Else
                  '========  GetNewState = Action_Sale
                GetNewState = Action_Sale

                End Select
            End If
        End Select
    Case eSearchCustomerfordeposit
        If strPrefix = ".." Then
            GetNewState = eStart
        ElseIf ValidRowset(strPrefix) Then
            GetNewState = eCollect
        End If
    Case eSearchCustomerfordepositRefund
        If strPrefix = ".." Then
            GetNewState = eStart
        ElseIf ValidRowset(strPrefix) Then
            GetNewState = eRefundDeposit
        End If
    Case eSearchCustomerforAppro
        If strPrefix = ".." Then
            GetNewState = eStart
        ElseIf ValidRowset(strPrefix) Then
            GetNewState = eCollect
        End If
    Case eRefundDeposit
        If strPrefix = "X" Then
            GetNewState = eStart
        Else
            Select Case strPrefix
            Case "C"
                RefundPayment ePaymentMode_Cash
                GetNewState = eConfirmation
            Case "V"
                GetNewState = ePaymentType_voucher
            Case "CV"
                RefundPayment ePaymentMode_CreditVoucher
                GetNewState = eConfirmation
            Case "CC"
                RefundPayment ePaymentMode_CreditCard
                GetNewState = eConfirmation
            Case Else
                    bValid = False
            End Select
        End If
    Case eReviewExchanges
        GetNewState = Action_ReviewExchanges
    Case Else
        bValid = False
    End Select
            
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetNewState(txtIn)", txtIn
End Function
Private Function DetermineReturnToState() As eState
    Select Case enMode
    Case emode_Appro
        DetermineReturnToState = eAppro
    Case eMode_AcceptDeposit
        DetermineReturnToState = eStart
    Case eMode_ApproReturn
        DetermineReturnToState = eCollect
    Case eMode_ReturnDeposit
        If enPresentState = eConfirmation Then
            RemovePaymentLine , True
        End If
        DetermineReturnToState = eRefundDeposit
    Case emode_Sale
        If enPresentState = eConfirmation Or enPresentState = ePaymentType_CreditVoucherRef Or enPresentState = ePaymentType_voucherRef Then
            RemovePaymentLine , True
        End If
        DetermineReturnToState = eSale
    End Select
End Function
Private Function SeparateInput(pRaw As String, pPrefix As String, pSuffix As String) As Boolean
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
Dim c As String
Dim bAlpha As Boolean

    
    pPrefix = ""
    pSuffix = ""
    SeparateInput = True
    iMax = Len(pRaw)
    If InStr(1, pRaw, ",") > 0 Then  'there are commas in the string meaning a multiple selection
        For i = 1 To iMax
            c = MID(pRaw, i, 1)
            If Not (IsNumeric(c) Or c = ",") Then
                SeparateInput = False
                Exit Function
            End If
        Next
        pSuffix = pRaw
    Else
        bAlpha = True
        If iMax > 9 Then
            SeparateInput = True
            pPrefix = pRaw
            pSuffix = pRaw
            Exit Function
        End If
        If Left(pRaw, 1) = "#" Then
            SeparateInput = True
            pPrefix = pRaw
            pSuffix = pRaw
            Exit Function
        End If
        If (enPresentState <> eStart And enPresentState <> eSale) Or UCase(Left(pRaw, 1)) = "D" Or UCase(Left(pRaw, 1)) = "V" Then
            For i = 1 To iMax
                c = MID(pRaw, i, 1)
                If IsNumeric(c) Then
                    bAlpha = False
                    pSuffix = pSuffix & c
                Else
                    If bAlpha = False Then
                        SeparateInput = False
                        Exit For
                    Else
                        pPrefix = pPrefix & c
                    End If
                End If
            Next i
        Else
            SeparateInput = True
            pPrefix = pRaw
            pSuffix = pRaw
        
        End If
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SeparateInput(pRaw,pPrefix,pSuffix)", Array(pRaw, pPrefix, pSuffix)
End Function
Private Function Action_eXTerminate() As eState


    If MsgBox("Confirm cash up?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Action_eXTerminate = eStart
    Else
    
        'Check for messages in deadletter queue, If they exist, they should be sent again
        If oExchange.SaleLines.Count > 0 Then
            oExchange.CancelEdit
        End If
        bCloseXsession = True
        Action_eXTerminate = eEND
        oPC.OpenLocalDatabase
        oPC.ZSession.OpSession.Close_OP_Session
        oPC.CloseLocalDatabase
    End If
End Function
Private Function Action_eZTerminate() As eState
    If MsgBox("Close day ?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Action_eZTerminate = eStart
    Else
        If oExchange.SaleLines.Count > 0 Then
            oExchange.CancelEdit
        End If
        bCloseZsession = True
        Action_eZTerminate = eEND
        oPC.ZSession.Close_Z_Session
    End If
End Function

Private Sub DisplayProduct()
    On Error GoTo errHandler
    LoadSaleRow iCurrentSaleLine, oExchange.SaleLines.Count
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayProduct"
End Sub
Private Sub DisplayPayment()
    On Error GoTo errHandler
    LoadPaymentRow iCurrentPaymentLine
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayPayment"
End Sub
Private Sub SetTip(pMsg As String)
    On Error GoTo errHandler
    lblInput.Caption = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetTip(pMsg)", pMsg
End Sub
Private Sub setInputBox(pText As String, pPasswordChar As String, pChange As String, bAutoSelect As Boolean)
    On Error GoTo errHandler
    txtInput = pText
    txtInput.PasswordChar = pPasswordChar
    If bAutoSelect Then
        AutoSelect txtInput
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.setInputBox(pText,pPasswordChar,pChange,bAutoSelect)", Array(pText, _
         pPasswordChar, pChange, bAutoSelect)
End Sub
Private Sub OpenDrawer()
Dim ret As Long
Dim sControl As String
Dim ar() As String
Dim i As Integer

    On Error GoTo errHandler
   ' MsgBox "Open Drawer"
    If oPC.DriveDrawer = True Then
        sControl = oPC.GetProperty("CashDrawerKick")
        If sControl > "" Then
            ar = Split(sControl, ",")
            If UBound(ar) = 0 Then
                sControl = Chr(ar(0))
            Else
                sControl = ""
                For i = 0 To UBound(ar)
                    sControl = sControl & Chr(ar(i))
                Next
            End If
            MSComm1.Output = sControl
        Else
            MSComm1.Output = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10)
        End If
    Else
        OPOSCashDrawer1.OpenDrawer
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.OpenDrawer"
End Sub

Private Function CheckAllStatus(pRefund As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim iRow As Long
Dim bValid As Boolean
Dim lngTmp As Long
Dim strDepositStatus As String

    bValid = True
    For i = 0 To UBound(arLineNumber)
        iRow = CLng(arLineNumber(i))
        If Not X3 Is Nothing Then
            If iRow >= X3.LowerBound(1) And iRow <= X3.UpperBound(1) Then
                lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                lngDeposit = X3(lngTmp, 12)
                strDepositStatus = X3(lngTmp, 7)
                If lngDeposit <= 0 Or strDepositStatus <> IIf(pRefund = True, "P", "O") Then
                    bValid = False
                End If
            Else
                bValid = False
            End If
        End If
    Next
    CheckAllStatus = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CheckAllStatus(pRefund)", pRefund
End Function
Private Function Valid_RedeemDeposit_Arg(pArg As String) As Boolean
    On Error GoTo errHandler
    Valid_RedeemDeposit_Arg = False
    If IsNumeric(pArg) Then
        Valid_RedeemDeposit_Arg = True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Valid_RedeemDeposit_Arg(pArg)", pArg, EA_NORERAISE
    HandleError
End Function
Private Function ValidRowNumberSet(pString As String)
    On Error GoTo errHandler
Dim i As Integer
Dim bValid As Boolean

    arLineNumber = Split(pString, ",")
    bValid = True
    If LenB(pString) = 0 Then bValid = False
    For i = 0 To UBound(arLineNumber)
        If Not IsNumeric(arLineNumber(i)) Then
            bValid = False
            Exit For
        End If
    Next i
    ValidRowNumberSet = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ValidRowNumberSet(pString)", pString
End Function
Private Function Validate_eSelectDepositLine(pIn As String, pRefund As Boolean) As Boolean
Dim strMsg As String

    On Error GoTo errHandler
    Validate_eSelectDepositLine = True
    If Not ValidRowNumberSet(pIn) Then
        MsgBox "Invalid row selection.", vbOKOnly, "Can't do this"
        Validate_eSelectDepositLine = False
        Exit Function
    End If
    If Not CheckAllStatus(pRefund) Then
        If pRefund Then
            strMsg = "Invalid row selection,  check status is 'P' " & vbCrLf _
                & "'P' means deposit has been paid." & vbCrLf _
                & "'E' means that the deposit has been redeemed already." & vbCrLf _
                & "'X' means that the deposit has been refunded."
                MsgBox strMsg, vbOKOnly, "Can't do this"
        End If
        Validate_eSelectDepositLine = False
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Validate_eSelectDepositLine(pIn,pRefund)", Array(pIn, pRefund), EA_NORERAISE
    HandleError
End Function

Private Function Action_SelectApproLine() As eState
'    On Error GoTo errHandler
'Dim i As Integer
'Dim iRow As Integer
'Dim lngTmp As Long
'
'    If Validate_eSelectApproLine(strSuffix, False) = False Then
'        Action_SelectApproLine = eSelectApproLine
'        Exit Function
'    End If
'    mlngTotalDepositValue = 0
'    For i = 0 To UBound(arLineNumber)
'        iRow = CLng(arLineNumber(i))
'        lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
'        Set oSALELine = oExchange.SaleLines.Add
'        oSALELine.ApplyEdit
'        oSALELine.BeginEdit
'        iCurrentSaleLine = iCurrentSaleLine + 1
'        X1.ReDim 1, iCurrentSaleLine, 1, 7
'        oSALELine.PID = X3(lngTmp, 13)
'        oSALELine.Price = FNN(X3(lngTmp, 12))
'        oSALELine.Title = X3(lngTmp, 5)
'        oSALELine.Code = X3(lngTmp, 3)
'        oSALELine.SetQty 1
'        oSALELine.IsDepositItem = True
'        oSALELine.CalculateLine
'        oExchange.CalculateTotals
'        mlngTotalDepositValue = mlngTotalDepositValue + FNN(X3(lngTmp, 12))
'        oSALELine.COLID = FNN(X3(lngTmp, 11))
'        oSALELine.ApplyEdit
'        oSALELine.BeginEdit
'        DisplayProduct
'    Next
'    Action_SelectApproLine = eCollect
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_SelectApproLine", , EA_NORERAISE
'    HandleError
End Function
Private Function Action_SelectDepositLine() As eState
    On Error GoTo errHandler
Dim i As Integer
Dim iRow As Integer
Dim lngTmp As Long

    If Validate_eSelectDepositLine(strSuffix, False) = False Then
        Action_SelectDepositLine = eSelectDepositLine
        Exit Function
    End If
    mlngTotalDepositValue = 0
    For i = 0 To UBound(arLineNumber)
        iRow = CLng(arLineNumber(i))
        lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
        Set oSALELine = oExchange.SaleLines.Add
        oSALELine.ApplyEdit
        oSALELine.BeginEdit
        iCurrentSaleLine = iCurrentSaleLine + 1
        X1.ReDim 1, iCurrentSaleLine, 0, 8
        oSALELine.PID = X3(lngTmp, 13)
        oSALELine.Price = FNN(X3(lngTmp, 12))
        oSALELine.title = X3(lngTmp, 5)
        oSALELine.Code = X3(lngTmp, 3)
        oSALELine.SetQty 1
        oSALELine.IsDepositItem = True
        oSALELine.CalculateLine
        oExchange.CalculateTotals
        mlngTotalDepositValue = mlngTotalDepositValue + FNN(X3(lngTmp, 12))
        oSALELine.COLID = FNN(X3(lngTmp, 11))
        oSALELine.ApplyEdit
        oSALELine.BeginEdit
        DisplayProduct
    Next
    Action_SelectDepositLine = eCollect
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SelectDepositLine", , EA_NORERAISE
    HandleError
End Function
Private Function Action_SelectDepositLineForRefund() As eState
    On Error GoTo errHandler
Dim i As Integer
Dim iRow As Integer
Dim lngTmp As Long
Dim lngTotalDeposit As Long
    If Validate_eSelectDepositLine(strSuffix, True) = False Then
        Action_SelectDepositLineForRefund = eSelectDepositLineForRefund
        Exit Function
    End If
    lngTotalDeposit = 0
    For i = 0 To UBound(arLineNumber)
        iRow = CLng(arLineNumber(i))
        lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
        Set oSALELine = oExchange.SaleLines.Add
        oSALELine.ApplyEdit
        oSALELine.BeginEdit
        iCurrentSaleLine = iCurrentSaleLine + 1
        X1.ReDim 1, iCurrentSaleLine, 0, 8
        oSALELine.PID = X3(lngTmp, 13)
        oSALELine.Price = FNN(X3(lngTmp, 12))
        oSALELine.Qty = (FNN(X3(lngTmp, 14))) * -1
        oSALELine.title = X3(lngTmp, 5)
        oSALELine.Code = X3(lngTmp, 3)
        lngTotalDeposit = lngTotalDeposit + FNN(X3(lngTmp, 12))
        oSALELine.COLID = FNN(X3(lngTmp, 11))
        DisplayProduct
    Next
    SetForCOLSVisible False
    Action_SelectDepositLineForRefund = eRefundDeposit 'Action_Refund("R")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SelectDepositLineForRefund", , EA_NORERAISE
    HandleError
End Function

Private Function Action_SearchCustomer(pType As String) As eState
Dim lngInvValue As Long

    On Error GoTo errHandler
    If GetCustomer Then
        If pType = "R" Then
            FetchCOLS
            If cCOLS.Count > 0 Then
                LoadCOLS
                G3.Caption = DisplayCustomerDetails
                oExchange.SetExchangeType eReturnDepositType
                Action_SearchCustomer = eSelectDepositLineForRefund
            Else
                MsgBox "There are no orders for this customer", vbInformation, "Can't find orders"
                Action_SearchCustomer = eStart
            End If
        ElseIf pType = "I" Then   'Appro return
                G3.Caption = DisplayCustomerDetails
                oExchange.SetExchangeType eApproReturnType
                If CreateApproReturnAndInvoice(lngCustomerID, lngInvValue) = True Then
                    If lngInvValue > 0 Then
                        Action_SearchCustomer = eCollect
                    Else
                        Action_SearchCustomer = eConfirmation
                    End If
                Else
                    Action_SearchCustomer = eStart
                End If
            
        ElseIf pType = "A" Then
            Action_SearchCustomer = eAppro
        ElseIf pType = "S" Then
            FetchCOLS
            If cCOLS.Count > 0 Then
                LoadCOLS
                G3.Caption = DisplayCustomerDetails
                Action_SearchCustomer = ePaymentType_RedeemDeposit
            Else
                MsgBox "There are no orders for this customer", vbInformation, "Can't find orders"
                Action_SearchCustomer = eSale
            End If
        ElseIf pType = "AC" Then 'Pay on Account
            G3.Caption = DisplayCustomerDetails
            'oExchange.SetExchangeType eAccountPaymentType
            Action_SearchCustomer = ePaymentType_Account 'next state
        ElseIf pType = "DDP" Then 'Pay on Account
            G3.Caption = DisplayCustomerDetails
            'oExchange.SetExchangeType eAccountPaymentType
            Action_SearchCustomer = ePaymentType_Account 'next state
        ElseIf pType = "PA" Then 'Pay Account balance
            G3.Caption = DisplayCustomerDetails
            oExchange.SetExchangeType eAccountPaymentType
            
            GetPaymentReference
            
            Action_SearchCustomer = eCollect
        ElseIf pType = "CN" Then 'Find Invoices
            G3.Caption = DisplayCustomerDetails
            oExchange.SetExchangeType eAccountCreditNoteType
            If GetInvoicesPerCustomer Then
                Action_SearchCustomer = eConfirmation
            Else
                Action_SearchCustomer = eStart
            End If
            
        Else
            Action_SearchCustomer = eSelectDepositLine
            oExchange.SetExchangeType eDepositType
        End If
        lblCustomername.Caption = DisplayCustomerDetails
        oExchange.CalculateTotals
        RefreshSaleDisplay
        txtInput = ""
     Else
         Action_SearchCustomer = enPresentState
     End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SearchCustomerfordeposit(pType)", pType, EA_NORERAISE
    HandleError
End Function
Private Function GetPaymentReference() As Boolean
Dim F As New frmPaymentReference
    F.Show vbModal
    oExchange.Note = F.PaymentReference & ": " & oExchange.Note
    
End Function
Private Function GetInvoicesPerCustomer() As Boolean
    oPC.dbConnectMain
    GetInvoicesPerCustomer = CreditAccount(lngCustomerID, 100)
    oPC.dbMainDisConnect
End Function
Private Function CreateApproReturnAndInvoice(pTPID As Long, pInvValue As Long) As Boolean
    On Error GoTo errHandler
Dim frm As New frmPOSAPPRet
Dim rs As Boolean
Dim lngAPPID As Long
Dim strAppCOde As String
Dim dteAppDate As Date
Dim Res
Dim i As Integer

    CreateApproReturnAndInvoice = False
    frm.component pTPID
    frm.Show vbModal
    If frm.Cancelled Then
        MsgBox "Appro return cancelled.", vbOKOnly, "Status"
        CreateApproReturnAndInvoice = False
        Exit Function
    End If
    
    ReDim arApproReturnLines(0)
    Res = frm.ApproReturnData(lngAPPID, strAppCOde, dteAppDate, pInvValue, arApproReturnLines)
    Unload frm
    If Res = False Then
        Exit Function
    End If
    If UBound(arApproReturnLines) > 0 Then
    For i = 1 To UBound(arApproReturnLines)
        CreateApproReturnAndInvoice = True
        txtInput = arApproReturnLines(i).Code
        LoadProductFromCode arApproReturnLines(i).Code
        oSALELine.IsDepositItem = False
        'Here we might allow override of quoted customer's price
        If FNN(arApproReturnLines(i).Price) > oSALELine.Price Then
            If MsgBox("The item price is less than the price used on the quote (" & oSALELine.PriceF & " and " & Format(FNN(arApproReturnLines(i).Price) / oPC.CurrencyDivisor, oPC.CurrencyFormat) & ")" & vbCrLf & "Do you want to give the lesser price?", vbQuestion + vbYesNo, "Changed price") = vbNo Then
                oSALELine.Price = FNN(arApproReturnLines(i).Price)
            End If
        End If
        oSALELine.Qty = arApproReturnLines(i).APPLQtySold
'        oSALELine.DiscountRate = 0  'Discount will have to be recalculated based on value of stock taken '  FNN(arApproReturnLines(i).DiscountRate) '   oExchange.Customer.DefaultDiscount  '
'        oSALELine.Title = arApproReturnLines(i).Title
'        oSALELine.Code = arApproReturnLines(i).Code
'        oSALELine.VATRATE = arApproReturnLines(i).VATRATE
'        oSALELine.PID = arApproReturnLines(i).PID
        oSALELine.COLID = arApproReturnLines(i).APPLID
        'oSALELine.CalculateLine
        'oExchange.CalculateTotals
        oSALELine.CalculateLine
      '  oExchange.TotalPayable = pInvValue
        DisplayProduct
        DoEvents
    Next i
    End If
    oExchange.CalculateTotals
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CreateApproReturnAndInvoice(pTPID,pInvValue)", Array(pTPID, pInvValue)
End Function
Private Function ValidRowset(strIn As String) As Boolean
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ValidRowset(strIn)", strIn
End Function
Private Sub Action_eStart()
    On Error GoTo errHandler
    lblReplacement.Caption = ""
    lblReplacement.Visible = False
    txtInput = ""
    txtInput.BackColor = RGB(230, 250, 210)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_eStart", , EA_NORERAISE
    HandleError
End Sub
Private Function CreditAccount(pTPID As Long, pInvValue As Long) As Boolean
Dim frm As New frmCreditAccountPOS
Dim rs As Boolean
Dim lngInvoiceID As Long
Dim strInvoiceCOde As String
Dim dteInvoiceDate As Date
Dim Res
Dim i As Integer

    CreditAccount = False
    frm.component pTPID, oExchange.Customer.NameAndCode(45)
    frm.Show vbModal
    If frm.Cancelled Then
        MsgBox "Credit note entry cancelled.", vbOKOnly, "Warning"
        CreditAccount = False
        Exit Function
    End If
    ReDim arInvoiceLines(0)
'   ' ReDim arApproReturnLines(1)
    Res = frm.InvoiceReturnData(lngInvoiceID, strInvoiceCOde, dteInvoiceDate, pInvValue, arInvoiceLines)
    Unload frm
    If Res = False Then
        Exit Function
    End If
    If UBound(arInvoiceLines) > 0 Then
    For i = 1 To UBound(arInvoiceLines)
        If FNN(arInvoiceLines(i).Qty) > 0 Then
            CreditAccount = True
            Set oSALELine = oExchange.SaleLines.Add
            oSALELine.ApplyEdit
            oSALELine.BeginEdit
            iCurrentSaleLine = iCurrentSaleLine + 1
            X1.ReDim 1, iCurrentSaleLine, 0, 8
            oSALELine.IsDepositItem = True '''Does not napply discount again
            oSALELine.Price = FNN(arInvoiceLines(i).Price)
            oSALELine.Qty = FNN(arInvoiceLines(i).Qty) * -1
            oExchange.TotalPayable = pInvValue
            oSALELine.DiscountRate = FNN(arInvoiceLines(i).DiscountRate) '   oExchange.Customer.DefaultDiscount  '
            oSALELine.title = arInvoiceLines(i).title
            oSALELine.Code = arInvoiceLines(i).Code
            oSALELine.VATRate = arInvoiceLines(i).VATRate
            oSALELine.PID = arInvoiceLines(i).PID
            oSALELine.COLID = arInvoiceLines(i).ILID
            oSALELine.CalculateLine
            DisplayProduct
        End If
    Next i
    End If
    
    
    
End Function

Private Function RefundPayment(ePaymentMode As enPaymentMode)

    On Error GoTo errHandler
    PreparePaymentLine ePaymentMode
    If oExchange.transactionType = "RDEP" Then
    'We can only refund one deposit at a time and the price field of the salesline holds the deposit value
        oPAYMENTLine.Amt = oExchange.SaleLines(1).Price * -1 'oExchange.BalanceOwing
        oExchange.SaleLines(1).Qty = 0
    Else
        oPAYMENTLine.Amt = oExchange.BalanceOwing
    End If
    
    
    DisplayPayment
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefundPayment", , EA_NORERAISE
    HandleError
End Function

Private Function Action_PaymentType_Cash() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_Cash = DetermineReturnToState
    Else
        PreparePaymentLine ePaymentMode_Cash
        If oPAYMENTLine.SetAmt(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_Cash = eConfirmation
            Else
                Action_PaymentType_Cash = eSale
            End If
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_Cash = ePaymentType_Cash
        End If
        DisplayPayment
    End If
    If strMsg = "NOTOK" Then
        If MsgBox("The payment made seems excessive: " & oPAYMENTLine.AmtF & vbCrLf & "Click 'Cancel' to re-enter", vbInformation + vbOKCancel, "Warning") = vbCancel Then
            RemovePaymentLine , True
            Action_PaymentType_Cash = eSale
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Cash", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_Cheque() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        RemovePaymentLine , True
        Action_PaymentType_Cheque = DetermineReturnToState
    Else
        PreparePaymentLine ePaymentMode_Cheque
        If oPAYMENTLine.SetAmt(strSuffix) Then
            oExchange.CalculateTotals
            Action_PaymentType_Cheque = ePaymentType_ChequeRef
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_Cheque = ePaymentType_Cheque
        End If
        DisplayPayment
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Cheque", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_ChequeRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        RemovePaymentLine , True
        Action_PaymentType_ChequeRef = DetermineReturnToState
    Else
        If oPAYMENTLine.SetReference(strRaw) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_ChequeRef = eConfirmation
            Else
                Action_PaymentType_ChequeRef = eSale
            End If
            DisplayPayment
        Else
            SetTip "Invalid Reference."
            Action_PaymentType_ChequeRef = ePaymentType_ChequeRef
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_ChequeRef", , EA_NORERAISE
    HandleError
End Function

Private Function Action_PaymentType_Creditcard() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        RemovePaymentLine , True
        Action_PaymentType_Creditcard = DetermineReturnToState
    Else
        PreparePaymentLine ePaymentMode_CreditCard
        If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
            oExchange.CalculateTotals
            Action_PaymentType_Creditcard = ePaymentType_CreditVoucherRef
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_Creditcard = ePaymentType_CreditCard
        End If
        DisplayPayment
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Creditcard", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_CreditcardRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditcardRef = DetermineReturnToState
    Else
        If oPAYMENTLine.SetReference(Trim(txtInput)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_CreditcardRef = eConfirmation
            Else
                Action_PaymentType_CreditcardRef = eSale
            End If
            DisplayPayment
        Else
            SetTip "Invalid Reference."
            Action_PaymentType_CreditcardRef = ePaymentType_CreditVoucherRef
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditcardRef", , EA_NORERAISE
    HandleError
End Function

Private Function Action_PaymentType_Voucher() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        RemovePaymentLine , True
        Action_PaymentType_Voucher = DetermineReturnToState
    Else
        PreparePaymentLine ePaymentMode_Voucher
        If oPAYMENTLine.SetAmt(Trim(strSuffix)) Then
            oExchange.CalculateTotals
            Action_PaymentType_Voucher = ePaymentType_voucherRef
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_Voucher = ePaymentType_voucher
        End If
        DisplayPayment
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Voucher", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_VoucherRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_VoucherRef = DetermineReturnToState
    Else
        If InStr(1, strValidVoucherTypes, strPrefix) > 0 And Len(strRaw) > 1 And strPrefix > "" Then 'valid voucher type
            If oPAYMENTLine.SetReference(Trim(strRaw)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete(, strMsg) Then
                    Action_PaymentType_VoucherRef = eConfirmation
                Else
                    Action_PaymentType_VoucherRef = eSale
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
                Action_PaymentType_VoucherRef = ePaymentType_voucherRef
            End If
        Else
            Action_PaymentType_VoucherRef = ePaymentType_voucherRef
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_VoucherRef", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_CreditVoucher() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditVoucher = DetermineReturnToState
    Else
        PreparePaymentLine ePaymentMode_CreditVoucher
        If oPAYMENTLine.SetAmt(Trim(strSuffix)) Then
            oExchange.CalculateTotals
            Action_PaymentType_CreditVoucher = ePaymentType_CreditVoucherRef
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_CreditVoucher = ePaymentType_CreditVoucher
        End If
        DisplayPayment
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditVoucher", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_CreditVoucherRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditVoucherRef = DetermineReturnToState
    Else
        If oPAYMENTLine.SetReference(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_CreditVoucherRef = eConfirmation
            Else
                Action_PaymentType_CreditVoucherRef = eSale
            End If
            DisplayPayment
        Else
            SetTip "Invalid Reference."
            Action_PaymentType_CreditVoucherRef = ePaymentType_CreditVoucherRef
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditVoucherRef", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_Account() As eState
Dim lngDeposit As Long
Dim iRow As Long
Dim lngTmp As Long

    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_Account = DetermineReturnToState
    Else
        If Not IsNumeric(strSuffix) Then
            Action_PaymentType_Account = eSale
            Exit Function
        End If
        
        If Not bCustomerVisible Then
            If Not GetCustomer Then
                Action_PaymentType_Account = eSale
                Exit Function
            End If
        End If
        
        PreparePaymentLine ePaymentMode_Account
        If oPAYMENTLine.SetAmt(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_Account = eConfirmation
            Else
                Action_PaymentType_Account = eSale
            End If
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_Account = ePaymentType_Account
        End If
        DisplayPayment
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Account", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_DirectDeposit() As eState
Dim lngDeposit As Long
Dim iRow As Long
Dim lngTmp As Long


    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_DirectDeposit = DetermineReturnToState
    Else
        If Not IsNumeric(strSuffix) Then
            Action_PaymentType_DirectDeposit = eSale
            Exit Function
        End If
        If Not SecurityControl(enSECURITY_ACCEPTDIRECTDEPOSIT, lngOPID, strName, , "Enter your security key.", "Your signature does not give you permission for theis operation") Then
            Action_PaymentType_DirectDeposit = eSale
            Exit Function
        End If
        PreparePaymentLine ePaymentMode_DIrectDeposit
        If oPAYMENTLine.SetAmt(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_DirectDeposit = eConfirmation
            Else
                Action_PaymentType_DirectDeposit = eSale
            End If
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_DirectDeposit = ePaymentMode_DIrectDeposit
        End If
        DisplayPayment
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_DirectDeposit", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_RedeemDeposit() As eState
Dim lngDeposit As Long
Dim iRow As Long
Dim lngTmp As Long

    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_RedeemDeposit = DetermineReturnToState
    Else
        lngDeposit = 0
        If Not IsNumeric(strSuffix) Then
            Action_PaymentType_RedeemDeposit = eSale
            Exit Function
        End If
        iRow = CLng(strSuffix)
        If Not (iRow <= X3.UpperBound(1) And iRow >= X3.LowerBound(1)) Then
            Action_PaymentType_RedeemDeposit = ePaymentType_RedeemDeposit
            MsgBox "Invalid row number", , "Can't accept"
           
           bValid = False
            Exit Function
        End If
        lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
        lngDeposit = FNN(X3(lngTmp, 12))
        If lngDeposit <= 0 Or FNS(X3(lngTmp, 7)) <> "P" Then   'Or lngDeposit > 100000
            MsgBox "Deposit is invalid amount or not paid", , "Can't accept"
            Action_PaymentType_RedeemDeposit = ePaymentType_RedeemDeposit
            Exit Function
        End If
        PreparePaymentLine ePaymentMode_RedeemedDeposit
        oPAYMENTLine.Amt = lngDeposit
        If IsNumeric(X3(lngTmp, 11)) Then
            oPAYMENTLine.COLID = CLng(X3(lngTmp, 11))
        End If
        oExchange.CalculateTotals
        If oExchange.PaymentsComplete(, strMsg) Then
            Action_PaymentType_RedeemDeposit = eConfirmation
        Else
            Action_PaymentType_RedeemDeposit = eSale
        End If
        DisplayPayment
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_RedeemDeposit", , EA_NORERAISE
    HandleError
End Function

Private Function Action_Refund(pType As String) As eState
Dim frm As frmRefunds
    Set frm = New frmRefunds
    If pType = "CN" Then
        frm.component "Refund issued by Credit Note", "YES, Issue credit note", oExchange.BalanceOwingF
    ElseIf pType = "C" Then
        frm.component "Refund issued by CASH", "YES, Issue CASH refund", oExchange.BalanceOwingF
    End If
    frm.Show vbModal
    If frm.Cancelled Then
        Action_Refund = eSale
    Else
        Action_Refund = eConfirmation
    End If
    Unload frm

End Function
Private Function Action_Confirmation()
Dim frm As frmChangeToGive
Dim strMsg As String
Dim oPmt As a_Payment
Const errMsg = "You do not have the authority to accept payment. Talk to your supervisor."
    On Error GoTo errHandler
    
    If strPrefix = ".." Then
        Action_Confirmation = DetermineReturnToState
    Else
        If IsRole(enSECURITY_ISOPERATOR, strPrefix & strSuffix, strName, lngOPID) = True Then
            oExchange.SalesPersonID = lngOPID
            
            
            If oExchange.IssueCreditNoteForChange(strMsg) = True Then
                Set frm = New frmChangeToGive
                frm.component strMsg
                frm.Show vbModal
                bIssueCreditNote = frm.IssueChangeAsCreditNote
                If bIssueCreditNote Then
                    Set oPmt = oExchange.PaymentLines.Add
                    oPmt.BeginEdit
                    oPmt.Amt = oExchange.ChangeGiven * -1
                    oPmt.SetType "CNR"
                    oPmt.ApplyEdit
                End If
                Unload frm
            End If
            If oExchange.CashTransaction Then
                OpenDrawer
            End If
            
            If oExchange.transactionType = "S" Or oExchange.transactionType = "AR" Or oExchange.transactionType = "OR" Or oExchange.transactionType = "" Then
                If oExchange.ChangeGiven > 0 Then
                    lblChange.ForeColor = COLOUR_CHANGE
                    lblChange.Caption = "CHANGE: " & oExchange.ChangeGivenF
                Else
                    Set oPmt = oExchange.CCPayment
                    If Not oPmt Is Nothing Then
                        lblChange.ForeColor = COLOUR_CREDITCARD
                        lblChange.Caption = "CARD: " & oPmt.AmtF
                    End If
                    
                End If
            End If
            
            AcceptSale
        ElseIf UCase(txtInput) = "XX" Then
            If MsgBox("Confirm cancel?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                RejectSale
                Action_Confirmation = eSale
            End If
        Else
            MsgBox errMsg, vbInformation + vbOKOnly, "Security"
            Action_Confirmation = enPresentState
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Confirmation()", , EA_NORERAISE
    HandleError
End Function
Private Function Action_operatorsReport()
    If SecurityControl(eSupervisor, lngStaffID, strName, , "Enter security code to view operators' report") Then
        Set frmOpRep = New frmPOSOPREP
        setInputBox "", "", "", True
        frmOpRep.Show vbModal
    End If
End Function
Private Function Action_PettyCash() As eState
    On Error GoTo errHandler
    Set frmPC = New frmPettyCash
    frmPC.Show vbModal
    If frmPC.Cancelled Then
        'clear fields
        Unload frmPC
        setInputBox "", "", "", True
    Else
        If SecurityControl(eOperator, lngOPID, strName, , "Enter your security key.", "Your key is invalid", True) Then
       ' If IsRole(enSECURITY_ISOPERATOR, strPrefix, strName, lngOPID) = True Then
            oExchange.SetExchangeType ePettyCashType
            oExchange.Note = frmPC.Reason
            Set oPAYMENTLine = oExchange.PaymentLines.Add
            oPAYMENTLine.ApplyEdit
            oPAYMENTLine.BeginEdit
            oPAYMENTLine.SetAmt CStr(frmPC.Amount)
            oPAYMENTLine.SetType "W"
            AcceptSale
            OpenDrawer
            Unload frmPC
            setInputBox "", "", "", True
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PettyCash", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PettyCashCredit() As eState
    On Error GoTo errHandler
    Set frmPCC = New frmPettyCashCredit
    frmPCC.component CollectPettyCashArray
    frmPCC.Show vbModal
    If frmPCC.Cancelled Then
        'clear fields
        Unload frmPCC
        setInputBox "", "", "", True
    Else
        If SecurityControl(eOperator, lngOPID, strName, , "Enter your security key.", "Your key is invalid", True) Then
      '  If IsRole(enSECURITY_ISOPERATOR, strPrefix, strName, lngOPID) = True Then
            oExchange.SetExchangeType ePettyCashCreditType
            oExchange.Note = frmPCC.Reason
            Set oPAYMENTLine = oExchange.PaymentLines.Add
            oPAYMENTLine.ApplyEdit
            oPAYMENTLine.BeginEdit
            oPAYMENTLine.SetAmt CStr(frmPCC.Amount)
            oPAYMENTLine.SetType "R"
            AcceptSale
            OpenDrawer
            Unload frmPCC
            setInputBox "", "", "", True
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PettyCashCredit", , EA_NORERAISE
    HandleError
End Function

Private Function Action_ReviewExchanges() As eState
    On Error GoTo errHandler
    setInputBox "", "", "", True
    If enPresentState = eReviewExchanges Then
        If strPrefix = "DD" Then
            Action_ReviewExchanges = eStart
            ShowTransactions False
        Else
            ShowExchange
            setInputBox "", "", "", True
        End If
    Else
        Action_ReviewExchanges = eReviewExchanges
        ShowTransactions True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_ReviewExchanges", , EA_NORERAISE
    HandleError
End Function
'Private Function Action_PayAccount() As eState
'    On Error GoTo errHandler
'    Set frmPCC = New frmPettyCashCredit
'    frmPCC.component CollectPettyCashArray
'    frmPCC.Show vbModal
'    If frmPCC.Cancelled Then
'        'clear fields
'        Unload frmPCC
'        setInputBox "", "", "", True
'    Else
'      '  If SecurityControl(2, lngOPID, strName, , "Enter your security key.", "Your key is invalid") Then
'        If IsRole(enSECURITY_ISOPERATOR, strPrefix, strName, lngOPID) = True Then
'            oExchange.SetExchangeType ePettyCashCreditType
'            oExchange.Note = frmPCC.Reason
'            Set oPAYMENTLine = oExchange.PaymentLines.Add
'            oPAYMENTLine.ApplyEdit
'            oPAYMENTLine.BeginEdit
'            oPAYMENTLine.SetAmt CStr(frmPCC.Amount)
'            oPAYMENTLine.SetType "R"
'            AcceptSale
'            Unload frmPCC
'            setInputBox "", "", "", True
'        End If
'    End If
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PettyCashCredit", , EA_NORERAISE
'    HandleError
'End Function

Private Function Action_ReviewDeadLetterQueue() As eState
    On Error GoTo errHandler
Dim frm As New frmDeadLetterQueue

    frm.Show vbModal
    Unload frm
    Action_ReviewDeadLetterQueue = eStart
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_ReviewDeadLetterQueue", , EA_NORERAISE
    HandleError
End Function
Private Function Action_Reprint() As eState
Dim frm As New frmReprint
Dim oTmpExchange As a_Exchange



    oPC.dbConnectMain

    If strPrefix = ".." Then
        Action_Reprint = eStart
    Else
        frm.Show vbModal
        If frm.Cancelled Then
            Action_Reprint = eStart
        Else
     '       frm.Show True
                Set oTmpExchange = oExchange
                Set oExchange = New a_Exchange
                oExchange.LoadFromMainDB frm.EXCHID, True
                
                    If oExchange.TransactionTypeEnum = eOrderRequestType Then
                        PrintORDERSlip 1, True
                    ElseIf oExchange.TransactionTypeEnum = ePettyCashType Then
                        PrintPettyCashVoucher 1
                    ElseIf oExchange.TransactionTypeEnum = ePettyCashCreditType Then
                        PrintPettyCashVoucher 1
                    Else
                        PrintSalesSlip 1, True
                    End If

                
                PrintSalesSlip 1, True
                Set oExchange = Nothing
                Set oExchange = oTmpExchange
                Set oTmpExchange = Nothing
            Action_Reprint = eStart
        End If
        Unload frm
    End If
    oPC.dbMainDisConnect
    
End Function

Private Function Action_LoadFromQuotation() As eState
Dim frm As New frmBrowseQuotations
Dim oTmpExchange As a_Exchange
Dim i As Integer
Dim rs As ADODB.Recordset
Dim oSL As a_Sale
    oPC.dbConnectMain

    If strPrefix = ".." Then
        Action_LoadFromQuotation = eStart
    Else
        frm.Show vbModal
        If frm.Cancelled Then
            Action_LoadFromQuotation = eSale
        Else
            
            Set rs = frm.SelectedLines
            If rs Is Nothing Then
                Exit Function
            End If
            rs.MoveFirst
            Do While Not rs.EOF
                LoadProductFromCode FNS(rs.Fields("P_EAN"))
                oExchange.SaleLines(iCurrentSaleLine).SetQty (FNN(rs.Fields("QUL_QTY"))), bItemExchange
                oExchange.SaleLines(iCurrentSaleLine).SetPrice (FNN(rs.Fields("QUL_PRICE"))), bItemExchange
                oExchange.SaleLines(iCurrentSaleLine).SetDiscountRate (FNDBL(rs.Fields("QUL_DISCOUNTPERCENT"))), "Quotation"
                oExchange.CalculateTotals
                oExchange.SetExchangeType eSaleType
                DisplayProduct
              '  oSL.ApplyEdit
              '  oSL.BeginEdit
              '  oSL.CalculateLine
                
                rs.MoveNext
            Loop
            DisplayTotals
'            For i = 1 To oExchange.SaleLines.Count
'                LoadSaleRow i, oExchange.SaleLines.Count
'            Next
            Action_LoadFromQuotation = eSale
        End If
        
        Unload frm
    End If
    oPC.dbMainDisConnect
    
End Function




Private Function Action_OrderRequest() As eState
Dim frm As New frmORREQ
Dim strRequestDetails As String
Dim lngDepositValue As Long

    strOrderedTitle = ""
    If strPrefix = ".." Then
        Action_OrderRequest = eStart
    Else
        frm.Show vbModal
        If frm.Cancelled Then
            Action_OrderRequest = eStart
        Else
            strRequestDetails = frm.Customer & "~~" & frm.Item
            lngDepositValue = frm.Deposit
            oExchange.Note = strRequestDetails & vbCrLf & "Deposit:" & Format((CDbl(lngDepositValue) / oPC.CurrencyDivisor), "###,##0.00")
            oExchange.SetExchangeType eOrderRequestType
        
            Set oSALELine = oExchange.SaleLines.Add
            oSALELine.ApplyEdit
            oSALELine.BeginEdit
            iCurrentSaleLine = iCurrentSaleLine + 1
            X1.ReDim 1, iCurrentSaleLine, 0, 8
            oSALELine.PID = ""
            oSALELine.Price = lngDepositValue
            oSALELine.title = frm.Item
            oSALELine.Code = ""
            oSALELine.SetQty 1
            oSALELine.IsDepositItem = True
            oSALELine.CalculateLine
            oExchange.CalculateTotals
        
            DisplayProduct
            DisplayTotals
            If oExchange.BalanceOwing = 0 Then
                Action_OrderRequest = eConfirmation
            Else
                Action_OrderRequest = eCollect
            End If
        End If
        strOrderedTitle = frm.Item & vbCrLf & "For: " & frm.Customer
        Unload frm
    End If
End Function
Private Function Action_Sale() As eState
''
Dim frm As New frmQuickProductFind
Dim strCode As String
Dim lngQtyQuickFound As Long
''
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_Sale = eStart
    ElseIf LoadProductFromCode(FNS(txtInput)) Then
        oExchange.CalculateTotals
        oExchange.SetExchangeType eSaleType
        DisplayProduct
        DisplayTotals
        Action_Sale = ePrice
    Else
    ''
        
        strCode = Replace(FNS(txtInput), """", "")
        
        lngQtyQuickFound = frm.component(strCode)
        If lngQtyQuickFound = 0 Then
            MsgBox "Nothing found", vbInformation, "Status"
                    Action_Sale = eSale
        Else
            frm.Show vbModal
            If frm.Cancelled = False Then
                If frm.EAN > "" Then
                    txtInput = frm.EAN
                    If LoadProductFromCode(FNS(txtInput)) Then
                        oExchange.CalculateTotals
                        oExchange.SetExchangeType eSaleType
                        DisplayProduct
                        DisplayTotals
                        Action_Sale = ePrice
                    Else
                        bValid = False
                        Action_Sale = eStart
                        MsgBox "Not on database or invalid action", vbInformation, "Status"
                    End If
                End If
            Else
                Action_Sale = eSale
            End If
            Unload frm
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Sale", , EA_NORERAISE
    HandleError
End Function
'Private Sub FindProduct()
'Dim frm As New frmQuickProductFind
'Dim strCode As String
'    strCode = txtCode
'
'    lngQtyQuickFound = frm.Component(FNS(txtInput))
'    If lngQtyQuickFound = 0 Then
'        MsgBox "Nothing found", vbInformation, "Status"
'    ElseIf lngQtyQuickFound = 1 Then
'        txtCode = strCode
'    Else
'        frm.Show vbModal
'    End If
'    If frm.Cancelled = False Then
'        If frm.EAN > "" Then txtCode = frm.EAN
'    End If
'    Unload frm
'End Sub

Private Function Action_CreditNote() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_CreditNote = eStart
    ElseIf Action_SearchCustomer("CN") Then
        oExchange.CalculateTotals
        oExchange.SetExchangeType eSaleType
        DisplayProduct
        DisplayTotals
        Action_CreditNote = ePrice
    Else
        bValid = False
        Action_CreditNote = eStart
        MsgBox "Not on database or invalid action", vbInformation, "Status"
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_CreditNote", , EA_NORERAISE
    HandleError
End Function

Private Function Action_Appro() As eState
    On Error GoTo errHandler
    If LoadProductFromCode(FNS(txtInput)) Then
        oExchange.CalculateTotals
        oExchange.SetExchangeType eApproType
        DisplayProduct
        DisplayTotals
        Action_Appro = ePrice
    Else
        MsgBox "Not on database or invalid action", vbInformation, "Status"
        Action_Appro = eAppro
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Appro", , EA_NORERAISE
    HandleError
End Function

Private Function Action_Price() As eState
    On Error GoTo errHandler
 Dim strErrPos As String
 
 strErrPos = "pos 1"
    If strPrefix = ".." Then
        Action_Price = DetermineReturnToState
        RemoveSaleLine iCurrentSaleLine
        DisplayTotals
        If oExchange.LoyaltyValue > 0 Then
            lblCustomername.Caption = DisplayCustomerDetails
        End If
        Exit Function
    End If
 strErrPos = "pos 2"
    If Not oSALELine Is Nothing And strSuffix > "" Then
        If oSALELine.Price <> CLng(Trim(strSuffix)) Then
            If Not SecurityControl(enSECURITY_POSPRICECHANGE, lngOPID, strName, , "Enter your signature for price change.", "Your signature does not give you permission to change a price") Then
                Action_Price = ePrice
                Exit Function
            End If
        End If
    End If
 strErrPos = "pos 3"
    If bShiftDown And oSALELine.IsDiscountAllowed Then 'And oExchange.transactionType <> "APP" Then
 strErrPos = "pos 4"
        If oSALELine.SetPrice(Trim(txtInput)) Then
            oExchange.CalculateTotals
            DisplayProduct
            Action_Price = eDiscount
        Else
            SetTip "Invalid price."
        End If
    Else
 strErrPos = "pos 5"
        If oSALELine.SetPrice(Trim(strSuffix)) Then
            oExchange.CalculateTotals
            DisplayProduct
            Action_Price = eQty
            lblCustomername.Caption = DisplayCustomerDetails
        Else
            SetTip "Invalid price."
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Price", , EA_NORERAISE, , "strErrpos", Array(strErrPos)
    HandleError
End Function

Private Function Action_Qty() As eState
    On Error GoTo errHandler
    bItemExchange = False
    If strPrefix = ".." Then
        Action_Qty = ePrice 'DetermineReturnToState
    Else
        If oExchange.transactionType = "S" Then
            If Left(strPrefix, 1) = "-" Then
                bItemExchange = True
                
                txtInput = strPrefix & strSuffix   'Right(strSuffix, Len(strSuffix) - 1)
               ' MsgBox "change"
            End If
        End If

     '   If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(strSuffix), bItemExchange) Then
        If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(strSuffix), bItemExchange) Then
               ' MsgBox "change2"
            oExchange.CalculateTotals
            DisplayProduct
            If oExchange.TransactionTypeEnum = eSaleType Then
                Action_Qty = eSale
            ElseIf oExchange.TransactionTypeEnum = eApproType Then
                Action_Qty = eAppro
            End If
            oSALELine.ApplyEdit
            oSALELine.BeginEdit
        Else
            SetTip "Invalid quantity."
            bValid = False
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Qty", , EA_NORERAISE
    HandleError
End Function
Private Function Action_Discount() As eState
    On Error GoTo errHandler
Dim strDiscountCode As String
Dim dblDiscountRate As Double
Dim strDiscountDescription As String
Dim bCancelled As Boolean
    If strPrefix = ".." Then
        Action_Discount = DetermineReturnToState
    Else
        If Not SecurityControl(enSECURITY_POSDISCOUNT, lngOPID, strName, , "Enter your signature.", "Your signature does not give you discount permissions") Then
            Action_Discount = DetermineReturnToState
        Else
    
            strDiscountCode = UCase(Left(txtInput, 1))
            If InStr(1, strValidDiscountTypes, strDiscountCode) > 0 Then 'valid discount type
                If strDiscountCode = "X" Then
                    ConnectionTimer.Enabled = False
                 '   If SecurityControl(3, lngStaffID, strName, , "Enter security code to allow discretionary discount") Then
                    If SecurityControl(enSECURITY_ISSUPERVISOR, lngStaffID, strName, bCancelled, "Enter security code to allow discretionary discount", "You are not entitled to offer discount.") = True Then
                        Set frmDisc = New frmDiscretionaryDiscount
                        frmDisc.Show vbModal
                        dblDiscountRate = frmDisc.DiscountRate
                        Unload frmDisc
                    End If
                    ConnectionTimer.Enabled = True
                Else
                    dblDiscountRate = GetDiscount(strDiscountCode, strDiscountDescription)
                End If
                oSALELine.SetDiscountRateDbl dblDiscountRate, strDiscountDescription
                oExchange.CalculateTotals
                DisplayProduct
                Action_Discount = eQty
            End If
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Discount", , EA_NORERAISE
    HandleError
End Function
Private Function Action_OpenDrawer() As eState
Dim frm As frmOD
    On Error GoTo errHandler
    If SecurityControl(eOperator, lngStaffID, strName, , "Enter security code to open drawer", , True) Then
        Set frm = New frmOD
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Exit Function
        End If
            
        OpenDrawer
        lngOPID = lngStaffID
        oExchange.SalesPersonID = lngStaffID
        oExchange.Note = frm.Reason
        oExchange.SetExchangeType eOpenDrawerType
        AcceptSale
        Unload frm
        
    End If
    Action_OpenDrawer = eStart
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_OpenDrawer", , EA_NORERAISE
    HandleError
End Function
Private Function Action_DeletePayment(iRow As String) As eState
Dim lngRow As Integer
    If IsNumeric(iRow) Then
        lngRow = CInt(iRow)
        If lngRow <= oExchange.PaymentLines.Count And lngRow > 0 Then
            RemovePaymentLine lngRow
        End If
    Else
        MsgBox "Invalid line to delete"
    End If
    If oExchange.SaleLines.Count = 0 And oExchange.PaymentLines.Count = 0 Then
        Action_DeletePayment = eStart
    Else
        Action_DeletePayment = eSale
    End If
End Function
Private Function Action_DeleteSaleLine(iRow As String) As eState
Dim lngRow As Integer
    If IsNumeric(iRow) Then
        lngRow = CInt(iRow)
        If lngRow <= oExchange.SaleLines.Count And lngRow > 0 Then
            RemoveSaleLine lngRow
        End If
    Else
        MsgBox "Invalid line to delete"
    End If
    If oExchange.SaleLines.Count = 0 And oExchange.PaymentLines.Count = 0 Then
        Action_DeleteSaleLine = eStart
    Else
        Action_DeleteSaleLine = eSale
    End If
End Function

Private Function Action_Void() As eState
Dim Res As Boolean
Dim bCancelled As Boolean

    If Not SecurityControl(enSECURITY_VOIDEXCHANGE, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission for this operation", True) Then
        Action_Void = eStart
        Exit Function
    End If

    If IsNumeric(strSuffix) Then
        iToVoid = CLng(strSuffix)
        If iToVoid >= CLng(X4(X4.UpperBound(1), 1)) And iToVoid < oExchange.ExchangeNumber Then
            If (X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 12) = 0) Then
                If Not oExchange.CanCancel(iToVoid) Then
              '  If X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 7) = "DEPOSIT" Then
                    MsgBox "This exhange cannot be voided.", vbInformation, "Can't do this"
                Else
                    Res = False
                    bCancelled = False
                    Do Until Res = True Or bCancelled = True
                        If Not SecurityControl(eSupervisor, lngStaffID, strName, bCancelled, "Enter your signature.", "Your signature is invalid for this action", True) Then
                            Res = False
                        Else
                            Res = True
                        End If
                    Loop
                    
                    
'                        If IsRole(enSECURITY_ISSUPERVISOR, strPrefix, strName, lngOPID) = True Then
                    '        If bCancelled
                    If Res = True Then
                        lngOPID = lngStaffID
                        oExchange.SalesPersonID = lngStaffID
                        oExchange.Note = oExchange.Note & "#" & CStr(iToVoid)
                        oExchange.SetExchangeType eVoidAction
                        AcceptSale
                    End If
                End If
            Else
                MsgBox "This transaction has been replaced already by exchange number " & CStr(X4(X4.Find(1, 1, iToVoid), 11)), vbInformation, "Can't do this"
                bValid = False
            End If
        Else
            MsgBox "This transaction number is out of range ", vbInformation, "Can't do this"
                bValid = False
        End If
    End If
    Action_Void = eStart
    
End Function

'====================================================================================================
'====================================================================================================
'====================================================================================================
'====================================================================================================
'====================================================================================================


Private Sub UpdateClientFromServerFiles(Optional rs As ADODB.Recordset, Optional msg As String)
    On Error GoTo errHandler
Dim ar() As String

    ar = Split(SVRMsg.Label, ",")
    If Not rs Is Nothing Then
        UpdatingLocalDatabase True, rs.RecordCount
    Else
        UpdatingLocalDatabase True, 0
    End If
    Select Case ar(0)
        Case "PROD"
                'Load product updates
                SaveProductUpdate rs
        Case "STAF"
                'Load Staff Member updates
                SaveStaffUpdate rs
        Case "CUST"
                'Load Customer updates
                SaveCustomerUpdate rs
        Case "ORDR"
                'Load Customer order updates
                SaveCustomerOrderUpdate rs
        Case "MARK"
                'Load marketing updates
                SaveMarketingUpdate rs
        Case "APPL"
                'Load appro updates
                SaveApproUpdate rs
        Case "APPRL"
                'Load appro updates
                SaveApproRUpdate rs
        Case "ClearCustomers"
                oPC.OpenLocalDatabase
                oPC.DBLocalConn.Execute "Delete FROM tCustomer"
                oPC.CloseLocalDatabase
        Case "ClearProducts"
                oPC.OpenLocalDatabase
                oPC.DBLocalConn.Execute "Delete FROM tProduct"
                oPC.CloseLocalDatabase
        Case "ClearAppros"
                oPC.OpenLocalDatabase
                oPC.DBLocalConn.Execute "Delete FROM tAPPL"
                oPC.CloseLocalDatabase
        Case "ClearCustomerOrders"
                oPC.OpenLocalDatabase
                oPC.DBLocalConn.Execute "Delete FROM tCOL"
                oPC.CloseLocalDatabase
        Case "ClearMarketingRules"
                oPC.OpenLocalDatabase
                oPC.DBLocalConn.Execute "Delete FROM tMarketing"
                oPC.CloseLocalDatabase
        Case "ClearStaffMembers"
                oPC.OpenLocalDatabase
                oPC.DBLocalConn.Execute "Delete FROM tStaffMembers"
                oPC.CloseLocalDatabase
    End Select
    UpdatingLocalDatabase False, 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.UpdateClientFromServerFiles(rs)", rs
End Sub

Private Function SaveProductUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim lngCnt As Long
    
    oPC.OpenLocalDatabase
    
    bUpdating = True
    rs.MoveFirst
    lngCnt = 1
    Do While Not rs.EOF
  '  If rs!PRU_EAN = "9771995429008" Then MsgBox "HERE"
        lngCnt = lngCnt + 1
        If lngCnt Mod 100 = 0 Then
            Counter "Pr", lngCnt
        End If
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.ActiveConnection = oPC.DBLocalConn
        
        If FNS(rs!PRU_LOG_TYPE) = "DEL" Then
            cmd.CommandText = "sp_DeleteProductOnFD"
            Set par = cmd.CreateParameter("@PID", adGUID, , , rs!PRU_P_ID)
            cmd.Parameters.Append par
            Set par = Nothing
            cmd.Execute
            
        Else
            cmd.CommandText = "dbo.sp_InsertProductUpdateToFD"
                
            Set par = cmd.CreateParameter("@PID", adGUID, , , rs!PRU_P_ID)
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInput, 50, FNS(rs!PRU_Code))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 50, FNS(rs!PRU_EAN))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@PUBLISHER", adVarChar, adParamInput, 500, FNS(rs!PRU_Publisher))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@SERIESTITLE", adVarChar, adParamInput, 225, FNS(rs!PRU_SeriesTitle))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@AUTHOR", adVarChar, adParamInput, 225, FNS(rs!PRU_MainAuthor))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@TITLE", adVarChar, adParamInput, 225, FNS(rs!PRU_Title))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@SP", adInteger, adParamInput, , FNN(rs!PRU_SP))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@SSP", adInteger, adParamInput, , FNN(rs!PRU_SSP))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@VATRATE", adNumeric, adParamInput, 10, FNDBL(rs!PRU_VATRATE))
            par.Precision = 8
            par.NumericScale = 2
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@PTID", adInteger, adParamInput, , FNN(rs!PRU_PTID))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@SECID", adInteger, adParamInput, , FNN(rs!PRU_SECID))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@NDA", adBoolean, adParamInput, , FNB(rs!PRU_NDA))
            cmd.Parameters.Append par
            Set par = Nothing
            Set par = cmd.CreateParameter("@MultibuyCode", adVarChar, adParamInput, 15, FNS(rs!PRU_MultibuyCode))
            cmd.Parameters.Append par
            Set par = Nothing
            cmd.Execute
            
        End If

        Set cmd = Nothing
        rs.MoveNext
    Loop

    bUpdating = False
    SaveProductUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function

Private Function SaveStaffUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
Dim strPos As String

    bUpdating = True
    strPos = "1"
    oPC.OpenLocalDatabase
strPos = "2"
    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tStaffMembers WHERE tStaffMembers.SM_ID =" & rs!SMU_ID
strPos = "3"
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
strPos = "4"
        If FNS(rs!SMU_NAME) = "X" Then
            If Not NewRS.EOF Then NewRS.Delete
        Else
            If NewRS.EOF Then
                NewRS.AddNew
            End If
strPos = "5"
            If Not IsNull(rs!SMU_ID) Then NewRS!SM_ID = rs!SMU_ID
            If Not IsNull(rs!SMU_NAME) Then NewRS!SM_Name = Trim$(rs!SMU_NAME)
            If Not IsNull(rs!SMU_Role) Then NewRS!SM_Role = rs!SMU_Role 'DO NOT TRIM this field
            If Not IsNull(rs!SMU_Telephone) Then NewRS!SM_Telephone = Trim$(rs!SMU_Telephone)
            If Not IsNull(rs!SMU_Mobile) Then NewRS!SM_Mobile = Trim$(rs!SMU_Mobile)
            If Not IsNull(rs!SMU_Password) Then NewRS!SM_Password = Trim$(rs!SMU_Password)
            If Not IsNull(rs!SMU_Shortname) Then NewRS!SM_Shortname = rs!SMU_Shortname
            sName = Trim$(rs!SMU_NAME)
            i = 1
strPos = "6"
DoUpdate:
            NewRS.Update
strPos = "7"
            If Err = -2147217887 Then
strPos = "8"
              NewRS!SM_Code = Left(NewRS!SM_Code, 3) & CStr(i)
strPos = "9"
              i = i + 1
              Err.Clear
              GoTo DoUpdate
            ElseIf Err <> 0 Then
strPos = "10"
                NewRS.Close
strPos = "11"
                bUpdating = False
                GoTo errHandler
            End If
strPos = "12"
            NewRS.Close
        End If
        rs.MoveNext
    Loop
strPos = "13"
    SaveStaffUpdate = True
    bUpdating = False
MEX:
strPos = "14"
    If rs.State = adStateOpen Then rs.Close
strPos = "15"
    Set rs = Nothing
    Set NewRS = Nothing
    oPC.CloseLocalDatabase
strPos = "16"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function

Private Function SaveCustomerUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
Dim lngCnt As Long

    bUpdating = True
    oPC.OpenLocalDatabase
    
    rs.MoveFirst
    Do While Not rs.EOF
        lngCnt = lngCnt + 1
        If lngCnt Mod 100 = 0 Then
            Counter "Cu", lngCnt
        End If
        sSQL = "SELECT * FROM tCustomer WHERE tCustomer.Customer_ID =" & rs!CU_ID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL
        
        If UCase(FNS(rs!CU_TYPE)) = "X" Then
            If Not NewRS.EOF Then
                NewRS.Delete
            End If
        Else
            If NewRS.EOF Then
                NewRS.AddNew
            End If
            NewRS!Customer_ID = FNN(rs!CU_ID)

            NewRS!C_Name = Left(FNS(rs!CU_Name), 50)
            If FNN(rs!CU_DefaultDiscount) <= 100 Then
                NewRS!C_DefaultDiscount = FNN(rs!CU_DefaultDiscount)
            Else
                NewRS!C_DefaultDiscount = 0
            End If
            NewRS!C_Acno = FNS(rs!CU_Acno)
            NewRS!C_Initials = Left(FNS(rs!CU_Initials), 8)
            NewRS!C_Title = FNS(rs!CU_Title)
            NewRS!C_Phone = FNS(rs!CU_Phone)
            NewRS!C_VATABLE = FNN(rs!CU_VATABLE)
            NewRS!C_BALANCE = FNN(rs!CU_BALANCE)
            NewRS!C_Type = FNS(rs!CU_TYPE)
            NewRS!C_BALANCES = FNS(rs!CU_BALANCES)
            NewRS!C_TERMS = FNN(rs!CU_TERMS)
            NewRS!C_CREDITLIMIT = FNN(rs!CU_CREDITLIMIT)
        End If
        NewRS.Update
        NewRS.Close
        rs.MoveNext
    Loop
    SaveCustomerUpdate = True
        bUpdating = False
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    'ErrorIn "frmPOSMain.SaveCustomerUpdate(rs)", rs, EA_NORERAISE
    GoTo MEX
End Function

Private Function SaveCustomerOrderUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True
    
    oPC.OpenLocalDatabase
    
    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tCOL WHERE COL_COLID = " & rs!COU_COLID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        
        If NewRS.EOF Then
            NewRS.AddNew
        End If
        NewRS!COL_COLID = FNN(rs!COU_COLID)
        NewRS!COL_TPID = FNN(rs!COU_TPID)
        NewRS!COL_TRID = FNN(rs!COU_TRID)
        NewRS!COL_Date = FND(rs!COU_Date)

        NewRS!COL_CODE = FNS(rs!COU_CODE)
        NewRS!COL_PID = FNS(rs!COU_PID)
        NewRS!COL_Qty = FNN(rs!COU_QTY)
        NewRS!COL_QTYDISPATCHED = FND(rs!COU_QTYDISPATCHED)

        NewRS!COL_PRICE = FNN(rs!COU_PRICE)
        NewRS!COL_DISCOUNTRATE = FNN(rs!COU_DISCOUNTRATE)
        NewRS!COL_DEPOSIT = FNN(rs!COU_DEPOSIT)
        NewRS!COL_DEPOSITSTATUS = FNS(rs!COU_DEPOSITSTATUS)

        NewRS!COL_DELETE = FNS(rs!COU_DOCSTATUS)
DoUpdate:
        NewRS.Update
        NewRS.Close
        rs.MoveNext
    Loop
    
    SaveCustomerOrderUpdate = True
        bUpdating = False
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function
Private Function SaveApproUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True
    
    oPC.OpenLocalDatabase

    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tAPPL WHERE APPL_APPLID = " & rs!APPL_APPLID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        
        If FNS(rs!APPL_CODE) = "X" Then
            NewRS.Delete
            NewRS.Close
        Else
            If NewRS.EOF Then
                NewRS.AddNew
            End If
            NewRS!APPL_APPLID = FNN(rs!APPL_APPLID)
            NewRS!APPL_TPID = FNN(rs!APPL_TPID)
            NewRS!APPL_TRID = FNN(rs!APPL_TRID)
            NewRS!APPL_Date = FND(rs!APPL_Date)
    
            NewRS!APPL_CODE = FNS(rs!APPL_CODE)
            NewRS!APPL_PID = FNS(rs!APPL_PID)
            NewRS!APPL_Qtyout = FNN(rs!APPL_Qtyout)
            NewRS!APPL_QtyBack = FND(rs!APPL_QtyBack)
    
            NewRS!APPL_PRICE = FNN(rs!APPL_PRICE)
            NewRS!APPL_DISCOUNTRATE = FNN(rs!APPL_DISCOUNTRATE)
            NewRS!APPL_EXCHANGEID = rs!APPL_EXCHANGEID
    
            NewRS!APPL_DELETE = FNS(rs!APPL_DELETE)
DoUpdate:
            NewRS.Update
            NewRS.Close
        End If
        rs.MoveNext
    Loop
    
    SaveApproUpdate = True
    bUpdating = False
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function
Private Function SaveApproRUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True
    
    oPC.OpenLocalDatabase
    
    rs.MoveFirst
    Do While Not rs.EOF
        oPC.DBLocalConn.Execute "UPDATE tAPPL SET APPL_QTYBACK = ISNULL(APPL_QTYBACK,0) + " & FNN(rs.Fields(1)) & " WHERE APPL_APPLID = " & FNN(rs.Fields(0))
        oPC.DBLocalConn.Execute "DELETE FROM tAPPL WHERE APPL_QTYOUT = APPL_QTYBACK"
        rs.MoveNext
    Loop
    
    SaveApproRUpdate = True
    bUpdating = False
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function

Private Function SaveMarketingUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True
    
    oPC.OpenLocalDatabase
    
    rs.MoveFirst
    Do While Not rs.EOF
'        sSQL = "SELECT * FROM tMarketing WHERE M_PT_ID = " & rs!MC_PT_ID _
'            & " AND M_SECTION_ID = " & rs!MC_SECTION_ID & " AND M_CUSTTYPE = '" & rs!MC_CUSTTYPE & "'"
        NewRS.LockType = adLockOptimistic
        sSQL = "SELECT * FROM tMarketing WHERE M_ID = " & rs!MC_ID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        If rs!MC_TYPE = "DEL" Then   'must delete record
            If Not NewRS.EOF Then
                NewRS.Delete
                NewRS.Update
            End If
        Else
            If NewRS.EOF Then
                NewRS.AddNew
            End If
            NewRS!M_PT_ID = FNN(rs!MC_PT_ID)
            NewRS!M_SECTION_ID = FNN(rs!MC_SECTION_ID)
            NewRS!M_CUSTTYPE = FNS(rs!MC_CUSTTYPE)
            NewRS!M_DISCOUNT = FND(rs!MC_DISCOUNT)
            NewRS!M_MINVALUE = FNN(rs!MC_MINVALUE)
            NewRS!M_DESCRIPTION = FNS(rs!MC_DESCRIPTION)
            NewRS!M_NODISCOUNTALLOWABLE = FNS(rs!MC_NODISCOUNTALLOWABLE)
            NewRS!M_IDENTIFYCUSTOMER = FNN(rs!MC_IDENTIFYCUSTOMER)
            NewRS!M_ACTIVE = FNN(rs!MC_ACTIVE)
            NewRS!M_ID = FNN(rs!MC_ID)
DoUpdate:
            NewRS.Update
            NewRS.Close
        End If
        rs.MoveNext
    Loop
    SaveMarketingUpdate = True
    bUpdating = False
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrSaveToFile
    GoTo MEX
End Function




