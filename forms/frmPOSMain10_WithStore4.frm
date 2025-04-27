VERSION 5.00
Object = "{CCB90150-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSPOSPrinter.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{CCB90040-B81E-11D2-AB74-0040054C3719}#1.0#0"; "OPOSCashDrawer.ocx"
Begin VB.Form frmPOSMain 
   BackColor       =   &H00E1E1E1&
   Caption         =   "DiscountSet"
   ClientHeight    =   9765
   ClientLeft      =   165
   ClientTop       =   150
   ClientWidth     =   15240
   ControlBox      =   0   'False
   FillColor       =   &H00E1E1E1&
   Icon            =   "frmPOSMain10_WithStore4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCloseDrawerMessage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1140
      Left            =   60
      TabIndex        =   27
      Text            =   "Close drawer"
      Top             =   4545
      Visible         =   0   'False
      Width           =   12120
   End
   Begin VB.TextBox lblCHange 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1320
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   2370
      Width           =   5505
   End
   Begin VB.Frame frTotals 
      BackColor       =   &H00E1E1E1&
      Height          =   1185
      Left            =   5820
      TabIndex        =   19
      Top             =   5970
      Width           =   3810
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00714942&
         Height          =   405
         Left            =   165
         TabIndex        =   25
         Top             =   645
         Width           =   750
      End
      Begin VB.Label txtExtTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   1035
         TabIndex        =   24
         Top             =   675
         Width           =   2625
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00714942&
         Height          =   405
         Left            =   165
         TabIndex        =   23
         Top             =   210
         Width           =   750
      End
      Begin VB.Label txtQtyTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2235
         TabIndex        =   20
         Top             =   240
         Width           =   1425
      End
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2925
      Left            =   105
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":038A
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   345
      Width           =   8640
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   9690
      TabIndex        =   15
      Top             =   7080
      Width           =   1935
      Begin VB.Label lblSaleOnHold 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "** Sale on hold **"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   165
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblProg 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label lblUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Updating"
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   75
         TabIndex        =   16
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
      TabIndex        =   12
      Text            =   "frmPOSMain10_WithStore4.frx":5415
      Top             =   4245
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtVouchers 
      Enabled         =   0   'False
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmPOSMain10_WithStore4.frx":541B
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
      Left            =   1620
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":5421
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1260
      Visible         =   0   'False
      Width           =   12345
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
      Height          =   1050
      Left            =   45
      TabIndex        =   1
      Top             =   6060
      Width           =   5655
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Height          =   1245
      Left            =   60
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":B41C
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3855
      Width           =   4485
   End
   Begin TrueOleDBGrid60.TDBGrid G5 
      Height          =   2010
      Left            =   1950
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":EA8F
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   7440
   End
   Begin TrueOleDBGrid60.TDBGrid G3 
      Height          =   2400
      Left            =   2160
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":12CF6
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   11385
   End
   Begin OposCashDrawer_CCOCtl.OPOSCashDrawer OPOSCashDrawer 
      Left            =   13140
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":18845
      Top             =   4035
   End
   Begin OposPOSPrinter_CCOCtl.OPOSPOSPrinter OPOSPOSPrinter 
      Left            =   14235
      OleObjectBlob   =   "frmPOSMain10_WithStore4.frx":18869
      Top             =   8415
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
      TabIndex        =   21
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
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   0
      TabIndex        =   18
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
   Begin VB.Label txtPaymentTotal 
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
      ForeColor       =   &H00B07B64&
      Height          =   435
      Left            =   15
      TabIndex        =   9
      Top             =   5145
      Width           =   7125
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
      Left            =   4800
      TabIndex        =   7
      Top             =   4245
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
      Left            =   270
      TabIndex        =   2
      Top             =   5655
      Width           =   5400
   End
End
Attribute VB_Name = "frmPOSMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      Private Declare Function DrawText _
       Lib "user32.dll" Alias "DrawTextA" ( _
       ByVal hdc As Long, _
       ByVal lpStr As String, _
       ByVal nCount As Long, _
       ByRef lpRect As RECT, _
       ByVal wFormat As Long) As Long
      Private Const DT_CALCRECT As Long = &H400
      Private Type RECT
          Left As Long
          TOP As Long
          Right As Long
          Bottom As Long
      End Type

Private Enum PrinterType
    en_Epson = 1
    en_DigiPos = 2
    en_DDigipos = 3
End Enum
Dim Res As Boolean
Dim enPrinterType As PrinterType
Dim bCanSenseDrawer As Boolean
Dim itest As Integer
Dim frmLoading As Boolean
Dim oTmpExchange As a_Exchange
Dim bIgnorestatus As Boolean

Const M_VOID As Long = 1
Const M_OPENDRAWER As Long = 2
Const M_DISCOUNT As Long = 4
Const M_CREDITNOTE As Long = 8
Const M_ISSUEAPPRO As Long = 16
Const M_REFUNDDEPOSIT As Long = 32
Const M_ACCEPTACPAYMENT As Long = 64
Const M_ISSUEPOSCREDITNOTE As Long = 128
Const M_ISSUEPOSREFUND As Long = 256
Const M_ACCEPTDIRECTDEPOSIT As Long = 512
Const M_PETTYCASH As Long = 1024
Const M_POSPRICECHANGE As Long = 2048
Const M_POSDISCOUNT As Long = 4096
Const M_CLOSEAPPLICATION As Long = 8192
Const M_DELETELINE As Long = 16384
Const M_SENSEDRAWEROFF As Long = 32767
Dim lngSecurityFlags As Long
Dim bBarcodeNotPrice As Boolean
Dim bNoMoreSaleLines As Boolean
Dim iChangeGivenLines As Long
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
Dim OPOSPrinter As Object
'Dim OPOSCashDrawer As opo
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
    eTypeCancelledSale = 15
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
Dim lngSupervisorID As Long
Dim lngSalesItemCount As Long
Dim iToVoid As Long
Dim strEXCHtoVoidGUID As String
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
Dim enPresentState_Store1 As eState
Dim cCOLS_Store1 As c_COLS
Dim cApps_Store1 As c_APPs

''''

Dim WithEvents oExchange As a_Exchange
Attribute oExchange.VB_VarHelpID = -1
Dim oPAYMENTLine As a_Payment
'Dim oDatabase As SQLDMO.Database2
'Dim oSQLServer As SQLDMO.SQLServer2
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
Dim bCollectRepcode As Boolean
Dim bDrawerFlag As Boolean

Private Sub DisplayPayment()
    On Error GoTo errHandler
    LoadPaymentRow iCurrentPaymentLine
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayPayment"
End Sub

Private Sub RefreshAllSaleRows()
    On Error GoTo errHandler
Dim i As Integer
  '  X1.ReDim 1, oExchange.SaleLines.Count, 0, 8
  '  For i = 1 To oExchange.SaleLines.Count
     '   LoadTopSaleRow i, oExchange.SaleLines.Count, False, True
     
     
     '3/12/2024 changed this
     '    For i = X1.UpperBound(1) To X1.UpperBound(1) - 15 Step -1
    ' to the code below
    For i = X1.UpperBound(1) To X1.LowerBound(1) Step -1
        If i < 1 Then Exit For
        UpdateSpecifiedSalesRow i - 1, oExchange.SaleLines(i)
    Next i
        RenumberSalesRows
    For i = 1 To oExchange.PaymentLines.Count
        LoadPaymentRow i
    Next
    lblSaleOnHold.Visible = bSaleOnHold
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefreshAllSaleRows"
End Sub
Private Sub RenumberSalesRows()
Dim i
    For i = 1 To X1.UpperBound(1)
        X1.Value(i, 0) = X1.UpperBound(1) - i + 1
    Next i
    G1.ReBind
End Sub
Private Sub LoadTopSaleRow(Index As Integer, MaxLines As Long, bUpdateCurrent As Boolean, bUpdateAll As Boolean)
    On Error GoTo errHandler
Dim i As Long
Dim strPos As String
    
    i = Index
    G1.Visible = True
    If bUpdateCurrent = False And bUpdateAll = False Then
        X1.InsertRows (1)
        i = 1
    End If

If bUpdateCurrent Then i = 1
    X1.Value(i, 0) = X1.UpperBound(1) ' - 1
    X1.Value(i, 1) = oSALELine.CodeF
    X1.Value(i, 2) = IIf(enPresentState = eSelectDepositLine, "(DEP)", "") & oSALELine.title & " (" & oSALELine.MainAuthor & ")"
    X1.Value(i, 3) = oSALELine.Qty
    X1.Value(i, 4) = oSALELine.PriceF & IIf(oSALELine.IsSpecialPrice, "**", "")
    X1.Value(i, 5) = oSALELine.DiscountRateF
    X1.Value(i, 6) = oSALELine.PLessDiscExtF
    If oExchange.transactionType <> "INV" Then
        X1.Value(i, 7) = oSALELine.PLessDiscExtVATF & "(" & oSALELine.VATRateF & ")"
    End If
    Me.G1.ReBind
    DisplayTotals
  '  oExchange.CalculateTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadTopSaleRow(Index,MaxLines)", Array(Index, MaxLines)
End Sub

Private Sub UpdateSpecifiedSalesRow(RowNumber As Integer, oSL As a_Sale)
Dim i As Integer
    oSL.CalculateLine
    i = X1.UpperBound(1) - RowNumber
    X1.Value(i, 0) = X1.UpperBound(1) - 1
    X1.Value(i, 1) = oSL.CodeF
    X1.Value(i, 2) = IIf(enPresentState = eSelectDepositLine, "(DEP)", "") & oSL.title & " (" & oSL.MainAuthor & ")"
    X1.Value(i, 3) = oSL.Qty
    X1.Value(i, 4) = oSL.PriceF & IIf(oSL.IsSpecialPrice, "**", "")
    X1.Value(i, 5) = oSL.DiscountRateF
    X1.Value(i, 6) = oSL.PLessDiscExtF
    If oExchange.transactionType <> "INV" Then
        X1.Value(i, 7) = oSL.PLessDiscExtVATF & "(" & oSL.VATRateF & ")"
    End If
    Me.G1.ReBind
    DisplayTotals

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
            txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF & " (Still owing " & oExchange.ChangeGivenNonNegativeF & ")"
        Else
            txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF & " (Change: " & oExchange.ChangeGivenF & ")"
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
    If oExchange.PaymentLines.Count = 0 Then Exit Sub
    If iIndex > X2.UpperBound(1) Then Exit Sub
    G2.Visible = True
   ' X2.ReDim 1, iIndex, 1, 3
    X2.Value(iIndex, 3) = oExchange.PaymentLines(iIndex).ReferenceComplete
    X2.Value(iIndex, 2) = oExchange.PaymentLines(iIndex).AmtF
    X2.Value(iIndex, 1) = oExchange.PaymentLines(iIndex).PaymentTypeF
    G2.Array = X2
    G2.ReBind
    G2.Refresh
  '  DoEvents
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LoadPaymentRow(iIndex)", iIndex
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadPaymentRow(iIndex)", iIndex
End Sub




Private Function Action_StoreSale() As eState
    On Error GoTo errHandler
   ' oExchange.ApplyEdit
    If oExchange.SaleLines.Count < 1 And oExchange.PaymentLines.Count < 1 Then Exit Function
    Set oExchangeCopy1 = oExchange
    iCurrentSaleLine_store1 = iCurrentSaleLine
    iCurrentPaymentLine_Store1 = iCurrentPaymentLine
    bIssueCreditNote_Store1 = bIssueCreditNote
    If X1.Count(1) > 0 Then
        CopyArray X1, X1_Store1
    End If
    If X2.Count(1) > 0 Then
        CopyArray X2, X2_Store1
    End If
    If X3.Count(1) > 0 Then
        CopyArray X3, X3_Store1
    End If
   ' CopyArray X4, X4_Store1
    If X5.Count(1) > 0 Then
        CopyArray X5, X5_Store1
    End If
    bSaleOnHold = True
    lblSaleOnHold.Visible = bSaleOnHold
    enPresentState_Store1 = enPresentState
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    PrepareForNewSale
    Action_StoreSale = eSale
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_StoreSale"
End Function
Private Sub CopyArray(xFrom As XArrayDB, xTo As XArrayDB)
    On Error GoTo errHandler
Dim i As Integer
Dim j As Integer
  '  If xFrom Is Empty Then Exit Sub
    xTo.ReDim xFrom.LowerBound(1), xFrom.UpperBound(1), xFrom.LowerBound(2), xFrom.UpperBound(2)
    For i = xFrom.LowerBound(1) To xFrom.UpperBound(1)
        For j = xFrom.LowerBound(2) To xFrom.UpperBound(2)
            xTo(i, j) = xFrom(i, j)
        Next j
    Next i
'errHandler:
'    If Err = 9 Then
'        Err.Clear
'        Exit Sub
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CopyArray(xFrom,xTo)", Array(xFrom, xTo)
End Sub



Private Function Action_RetrieveSale() As eState
    On Error GoTo errHandler
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
    If X1_Store1.Count(1) > 0 Then
        CopyArray X1_Store1, X1
    End If
    If X2_Store1.Count(1) > 0 Then
        CopyArray X2_Store1, X2
    End If
    If X3_Store1.Count(1) > 0 Then
        CopyArray X3_Store1, X3
    End If
  '  CopyArray X4_Store1, X4
    If X5_Store1.Count(1) > 0 Then
        CopyArray X5_Store1, X5
    End If
    bSaleOnHold = False
    RefreshAllSaleRows
    Action_RetrieveSale = eSale
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_RetrieveSale"
End Function




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
        frTotals.Visible = False
    Else
        G4.Visible = False
        G1.Visible = True
        frTotals.Visible = True
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ShowTransactions(bShow)", bShow
End Sub



Private Sub Form_Resize()
Dim lngDiff As Long
    
    If frmLoading Then Exit Sub
    G4.Left = 90
    G4.TOP = 90
    G1.Width = NonNegative_Lng(Me.Width - 500)
    G4.Width = NonNegative_Lng(Me.Width - 500)

    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - 4830)
    G4.Height = NonNegative_Lng(Me.Height - 4830)
'
    lngDiff = NonNegative_Lng(Me.Height - 2900)

    txtInput.TOP = lngDiff + 95
    lblInput.TOP = NonNegative_Lng(Me.Height - 3200)
    frTotals.TOP = lngDiff
    frTotals.Left = NonNegative_Lng(txtInput.Left + txtInput.Width + 200)
    
    G2.TOP = NonNegative_Lng(Me.Height - 4700)
    G2.Height = 1200
    txtPaymentTotal.TOP = G2.TOP + 1300
    txtPaymentTotal.Left = G2.Left
    
    SB.TOP = NonNegative_Lng(Me.Height - 1600)
    SB.Width = NonNegative_Lng(Me.Width - 3000)
    Frame1.TOP = NonNegative_Lng(SB.TOP - 150)
    Frame1.Left = NonNegative_Lng(Me.Width - 2120)
    lblPrompt.TOP = lngDiff
    txtDiscounts.TOP = NonNegative_Lng(lngDiff - txtDiscounts.Height)
    txtVouchers.TOP = NonNegative_Lng(lngDiff - txtVouchers.Height)
    lblLoyaltyValue.TOP = lngDiff
    lblCustomername.TOP = NonNegative_Lng(Me.Height - 4700)
    lblCustomername.Width = NonNegative_Lng(Me.Width - 4700)
    lblCHange.TOP = NonNegative_Lng((Me.Height / 2) - (lblCHange.Height / 2) - 800)
    lblCHange.Left = NonNegative_Lng(Me.Width / 2 - lblCHange.Width / 2)
    lblPrompt.TOP = lngDiff + 1400
    txtPaymentTotal.TOP = G2.TOP + G2.Height + 0
    txtPaymentTotal.Left = 120
    txtCloseDrawerMessage.Width = G1.Width
    txtCloseDrawerMessage.TOP = NonNegative_Lng((Me.Height / 2) - (txtCloseDrawerMessage.Height / 2))
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
    On Error GoTo errHandler
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject
Dim strFilefolder As String
Dim strFilename As String
Dim tmp As String

    strFilename = oPC.LocalRootFolder & "\BU\PBKSFD_TEST.BAK"
    
    Set oBU = New z_PBKSBackup
    Screen.MousePointer = vbHourglass
    DoEvents
    oBU.BackupToBriefcase strFilename, True
            DoEvents
    
    Screen.MousePointer = vbDefault
    MsgBox "New test database has been created. You are still connected to the " & IIf(oPC.DatabaseName = "PBKSFD_TEST", "TEST", "LIVE") & " database", vbOKOnly, "Status"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuNewTestFromLive_Click"
End Sub

Private Sub mnuRounding_Click()
    On Error GoTo errHandler
Dim f As New frmRoundingRules
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuRounding_Click"
End Sub

'Private Sub mnuSetProperty_Click()
'Dim f As New frmPrinterProperties
'
'    f.Show vbModal
'
'End Sub

Private Sub mnuSwaptoTest_Click()
    On Error GoTo errHandler

    If MsgBox("You want to open this application connected to the " & IIf(oPC.DatabaseName = "PBKSFD", "TEST", "LIVE") & " database?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    Else
        SaveSetting "POS", "StartDatabaseName", "DBNAME", IIf(oPC.DatabaseName = "PBKSFD", "PBKSFD_TEST", "PBKSFD")
    End If
    MsgBox "The application will close. Reopening it will have it connected to the " & IIf(oPC.DatabaseName = "PBKSFD", "TEST", "LIVE") & " database."
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuSwaptoTest_Click"
End Sub

Private Sub oExchange_CreditLimitExceeded(Excess As String)
    On Error GoTo errHandler
    If Excess > "" Then
        strCreditLimitExceededMessage = "Credit Limit exceeded by : " & Excess
    Else
        strCreditLimitExceededMessage = ""
    End If
    DisplayCustomerDetails
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_CreditLimitExceeded(Excess)", Excess
End Sub



Private Sub oSaleLine_ProvisionalPrice()
    On Error GoTo errHandler
    txtExtTotal.ForeColor = RGB(41, 133, 46)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oSaleLine_ProvisionalPrice"
End Sub




Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
        If Bookmark = 0 Then Exit Sub
        If X1.Count(1) <= 0 Then Exit Sub
        If (X1(Bookmark, 3) < 1) And (X1(Bookmark, 3) <> "") Then
            RowStyle.BackColor = 65135
        Else
            RowStyle.BackColor = &HFFFFFF
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle)
End Sub

Private Sub Label2_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Label2_Click"
End Sub

Private Sub lblPrompt_Click()
    On Error GoTo errHandler
Dim strMsg As String
    strMsg = "Additional options are: " & vbCrLf & vbCrLf _
    & "REPRINT - lets you review exchanges and reprint if necessary." & vbCrLf _
    & "OD - open the drawer." & vbCrLf & vbCrLf _
    & "Please note that the 'n' in Dn, DPn, Vn etc refers to an exchange number, " & vbCrLf _
    & "         so to void an exchange number 312 for example, use 'V312'"
    MsgBox strMsg, vbInformation + vbOKOnly, "Hints"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.lblPrompt_Click"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.oPS_ConnectionStatus(iStatus)", iStatus, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_ConnectionStatus(iStatus)", iStatus
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.UpdatingLocalDatabase(bOn,lngCnt)", Array(bOn, lngCnt)
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Counter(lngCnt)", lngCnt
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Counter(msg,lngCnt)", Array(msg, lngCnt)
End Sub

Private Sub SetPresentState(val As eState)
    On Error GoTo errHandler
    If val = eEND Then
        Unload Me
        Exit Sub
    End If
    If enPresentState <> val Then
        enPresentState = val
        Me.lblState.Caption = InterpretState
    End If
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

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.InterpretState"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.InterpretState"
End Function

Private Sub oExchange_ContainsLines(pYesNo As Boolean)
    On Error GoTo errHandler
    If bUnloading Then Exit Sub
    bSaleActive = pYesNo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.oExchange_ContainsLines(pYesNo)", pYesNo, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_ContainsLines(pYesNo)", pYesNo
End Sub
Private Sub SetTitleBar(pShowExchangeNumber As Boolean)
    On Error GoTo errHandler
    Caption = "Papyrus Point-of-Sale       " & oPC.StationName & "      Supervisor: " & oPC.ZSession.SupervisorName & "/" & oPC.ZSession.OpSession.Name & IIf(pShowExchangeNumber = True, "              #" & oExchange.ExchangeNumber, "")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SetTitleBar(pShowExchangeNumber)", pShowExchangeNumber
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
    QPOSACK.EnableNotification POSACKEvent
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.POSACKEvent_Arrived(Queue,Cursor)", Array(Queue, Cursor), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.POSACKEvent_Arrived(Queue,Cursor)", Array(Queue, Cursor)
End Sub

Private Sub qTimer_Tick()
'    On Error GoTo ErrHandler
On Error Resume Next
  '  If UCase(oPC.MainSQLServerName) = "DAVID-PC\PBKSINSTANCE2" Then Exit Sub
  QSVR.EnableNotification SVREvent
    QPOSACK.EnableNotification POSACKEvent
    Exit Sub
End Sub

Sub SVREvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
10        On Error GoTo errHandler
      Dim rs As New ADODB.Recordset
      Dim lngResult As Integer
          
20        Set QSVR = Queue
30        Set SVRMsg = QSVR.Receive
40        If Not (SVRMsg Is Nothing) Then
50            Screen.MousePointer = vbHourglass
60            Me.Refresh
70            If Left(SVRMsg.Label, 5) = "Clear" Then
80                UpdateClientFromServerFiles , SVRMsg.Label
90            Else
                  '  MsgBox "SVRMsg.Body = " & SVRMsg.Body
100               UpdateClientFromServerFiles SVRMsg.Body, ""
110           End If
120           DoEvents
130           Screen.MousePointer = vbDefault
140       End If
        On Error Resume Next
150     QSVR.EnableNotification SVREvent
160       Exit Sub
errHandler:
170       If ErrMustStop Then Debug.Assert False: Resume
180       ErrorIn "frmPOSMain.SVREvent_Arrived(Queue,Cursor)", Array(Queue, Cursor), , , "Line number", Erl()
End Sub

Private Sub SetupQueues()
On Error Resume Next
   'Set up receiving queue (always local queue)
    Set QI = New MSMQQueueInfo
    QI.PathName = oPC.NameOfPC & "\Private$\qposack"
    QI.Create , True
    Err.Clear
    Set QPOSACK = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set POSACKEvent = New MSMQEvent
    QPOSACK.EnableNotification POSACKEvent
    'Set up our SVR queue for receiving notifications about DB changes
    Set QI = Nothing
    Set QI = New MSMQQueueInfo
    QI.PathName = oPC.NameOfPC & "\Private$\qsvr"
    QI.Create , True
    Err.Clear
    Set QSVR = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set SVREvent = New MSMQEvent
  'XXX  QSVR.EnableNotification Event:=SVREvent

End Sub
Public Sub OpenQSVR()
    On Error GoTo errHandler
    If QSVR Is Nothing Then
        Set QSVR = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
        Set SVREvent = New MSMQEvent
        SetQSVRTrigger
    End If
    If Not QSVR.IsOpen2 Then
        SetQSVRTrigger
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.OpenQSVR"
End Sub
Public Sub CloseQSVR()
    On Error GoTo errHandler
    If Not QSVR Is Nothing Then
        If QSVR.IsOpen2 Then
            QSVR.Close
            Set QSVR = Nothing
          '  RaiseEvent QSVRTriggerStatus(False)
        End If
    Else
           ' RaiseEvent QSVRTriggerStatus(False)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CloseQSVR"
End Sub

Private Sub SetQSVRTrigger()
    On Error GoTo errHandler
  'XXX  QSVR.EnableNotification SVREvent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetQSVRTrigger"
End Sub

Private Sub Form_Load()
10        On Error GoTo errHandler
20        On Error GoTo errHandler
      Dim Result As Integer
      Dim bLoggedOnAlready As Boolean
      Dim strPos As String
      Dim strDBName As String
      Dim bGetFloat As Boolean
      Dim dblFLoat As Double
      Dim sFloatBreakdown As String
      Dim rst As ADODB.Recordset
      Dim strSQL As String
      Dim i As Integer

30      bTryToCOnnectToMainServer = True
40      lblCHange.Visible = False
50        strPos = "1"
60            frTotals.TOP = 202
70            frTotals.Left = 2
          
80            G1.TOP = 4
90            G1.Left = 2
          
100       bUpdating = False
110       bEnvironmentOK = True
120       ESC = Chr(27)
130       iToVoid = 0
          'Try to load local DB connection
140       If oPC Is Nothing Then
150           Set oPC = New z_POSCLIConnection
160           If UBound(arCommandLine) >= 1 Then
170               oPC.DatabaseName = arCommandLine(0)
180           Else
190               strDBName = GetSetting("POS", "StartDatabaseName", "DBNAME", "PBKSFD")
200               If strDBName = "" Then
210                   oPC.DatabaseName = "PBKSFD"
220                   oPC.UseTestDatabase = False
230               Else
240                   oPC.UseTestDatabase = False
250                   oPC.DatabaseName = strDBName
260               End If
270           End If

280           If UBound(arCommandLine) >= 2 Then
290               oPC.DatabaseName = arCommandLine(1)
300           Else
310               strDBName = GetSetting("POS", "StartDatabaseName", "DBNAME", "PBKSFD")
320               If strDBName = "" Then
330                   oPC.DatabaseName = "PBKSFD"
340                   oPC.UseTestDatabase = False
350               Else
360                   oPC.UseTestDatabase = False
370                   oPC.DatabaseName = strDBName
380               End If
390           End If
400       strPos = "2"
       '       mnuSwaptoTest.Caption = IIf(oPC.DatabaseName = "PBKSFD", "Swap to TEST database", "Swap to LIVE database")
410           oPC.InitializeSettings
420           oPC.dbConnect
430           oPC.LoadProperties
440           oPC.loadRoundingRules
450           oPC.loadMultibuys
460           If oPC.ServerIPAddress <> oPC.NameOfPC Then
470               SynchronizeTOD oPC.ServerIPAddress
480           End If
490       End If
500     bCanSenseDrawer = True ' = Not CheckThisPoint(M_SENSEDRAWEROFF)

510         If oPC.GetLoyaltyCode = "" Then
520             oPC.dbConnectMain
530             Set rst = New ADODB.Recordset
540             rst.Open "SELECT CF_BC_ID,CF_LOYALTYCLUBTYPE FROM tCONFIGURATION", oPC.DBMainConn
550             strSQL = "UPDATE tAPPSETTINGS SET BC_Code = '" & FNS(rst.Fields(0)) & "',Loyalty_Code = '" & FNS(rst.Fields(1)) & "'"
560            oPC.OpenLocalDatabase
570            oPC.DBLocalConn.Execute strSQL
580             rst.Close
590         End If
600       bCollectRepcode = oPC.GetProperty("CollectRepCodeTF")
          
          
          
610       bLoggedOnAlready = False
620       bLogonOK = True
630       oPC.SetupZSession lngStaffID, strName
640       If oPC.ZSession.SupervisorID = 0 Then
650           If oPC.GetProperty("CaptureFloat") = "TRUE" Then
660               bGetFloat = True
670           Else
680               bGetFloat = False
690           End If
700           LogonOperator bGetFloat, dblFLoat, sFloatBreakdown
710           If bLogonOK = False Then
720               bCloseZsession = True
730               GoTo EXITHANDLER
740           End If
750           oPC.ZSession.SupervisorID = lngStaffID
760           oPC.ZSession.SupervisorName = strName
770           bLoggedOnAlready = True
780       End If
790       strPos = "4"
800       If oPC.ZSession.LoadOpenXSession = False Then
810           oPC.ZSession.OpSession.Start_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
820           If oPC.ZSession.OpSession.SupervisorID = 0 Then
830               If bLoggedOnAlready = False Then
840                   If oPC.GetProperty("CaptureFloat") = "TRUE" Then
850                       bGetFloat = True
860                   Else
870                       bGetFloat = False
880                   End If
890                   LogonOperator bGetFloat, dblFLoat, sFloatBreakdown
900                   If bLogonOK = False Then
910                       bCloseXsession = True
920                       bCloseZsession = True
930                       GoTo EXITHANDLER
940                   End If
950               End If
960               oPC.OpenLocalDatabase
970               oPC.ZSession.OpSession.SetOperatorID lngStaffID, dblFLoat, sFloatBreakdown
                 ' oPC.ZSession.OpSession.OperatorID = lngStaffID
980               oPC.ZSession.OpSession.Name = strName
990               oPC.CloseLocalDatabase
1000          End If
1010      End If
1020      strPos = "5"
1030      SetForCOLSVisible False
1040      If oPC.UseCashDrawer Then
1050          If oPC.DriveDrawer = True Then  'There is a COM connected Cash Drawer
1060              MSComm1.Settings = oPC.COMPORTSettings
1070              MSComm1.CommPort = oPC.CashDrawerPort
1080              If MSComm1.PortOpen = False Then
1090                  MSComm1.PortOpen = True
1100              End If
1110          Else                            'There is a cash drawer connected to the Printer
'                  If UCase(oPC.GetProperty("ReceiptPrinterType")) = "" Or UCase(oPC.GetProperty("ReceiptPrinterType")) = "EPSON" Then
'950                   OPOSCashDrawer.DeviceEnabled = True
'960               Else
'970                   OPOSCashDrawerDigipos.DeviceEnabled = True
'980               End If
1120          End If
1130      End If
1140      strPos = "6"
1150      If oPC.PrintSlips Then
1160          If oPC.UseA4Printer = False Then
1170             SetupPrinter
1180          Else
1190              Printer.FontName = "COURIER"
1200              Printer.FontSize = 12
1210              iColWidth = 40
1220          End If
1230      End If
1240      If Not bEnvironmentOK Then
1250          GoTo EXITHANDLER
1260      End If
1270      LoadVoucherTypes
1280      LoadDiscountTypes
1290      txtInput.BackColor = RGB(230, 250, 210)
          
1300      G1.Array = X1
1310      G4.Height = 380
1320      If oPC.DatabaseName <> "PBKSFD_TEST" Then
1330          ReSendExchanges
1340      End If
1350      strPos = "8"
1360      X4.Clear
1370      X4.ReDim 1, 0, 1, 13
1380      LoadExchanges
1390      SetupQueues
          
1400      Set qTimer = New XTimer
1410      qTimer.Interval = 10000
1420      qTimer.Enabled = True
          
1430      If oPC.DatabaseName = "PBKSFD_TEST" Then
1440          Me.BackColor = vbRed
1450      End If
1460      Me.ScaleMode = vbTwips
1470      SetGridLayout G1, Me.Name & "A"
1480      SetGridLayout G4, Me.Name & "B"
1490      SetFormSize Me
     
EXITHANDLER:
1500      Exit Sub
'errHandler:
'1430      If ErrMustStop Then Debug.Assert False: Resume
'1440      ErrorIn "frmPOSMain.Form_Load", , EA_NORERAISE, , "strpos", Array(strPos)
'1450      HandleError
1510      Exit Sub
errHandler:
1520      If ErrMustStop Then Debug.Assert False: Resume
1530      ErrorIn "frmPOSMain.Form_Load"
End Sub


Private Sub SetupPrinter()
    On Error GoTo errHandler
Dim lngResult As Long

    If UCase(oPC.GetProperty("ReceiptPrinterType")) = "" Or UCase(oPC.GetProperty("ReceiptPrinterType")) = "EPSON" Then
        enPrinterType = en_Epson
    ElseIf UCase(oPC.GetProperty("ReceiptPrinterType")) = "DIGIPOS" Then
        enPrinterType = en_DigiPos
    ElseIf UCase(oPC.GetProperty("ReceiptPrinterType")) = "Digipos" Then
        enPrinterType = en_DDigipos
    End If
    If oPC.PrintSlips = True Then
        If enPrinterType = en_Epson Then
            Set OPOSPrinter = Me.OPOSPOSPrinter
            With OPOSPOSPrinter
                lngResult = .Open(oPC.Printername)
                If lngResult = 0 Then
                    lngResult = .ClaimDevice(50)
                    If lngResult = OPOS_SUCCESS Then
                   '     .ClaimDevice 1000
                        .DeviceEnabled = True
                        .MapMode = PTR_MM_METRIC
                        .RecLetterQuality = True
                        .RecLineChars = 40
                            
                        If oPC.UseCashDrawer Then
                            If oPC.DriveDrawer = False Then
                                With OPOSCashDrawer
                                    lngResult = .Open(oPC.TillDrawerName)
                                    If lngResult = 0 Then
                                        lngResult = .ClaimDevice(1000)
                                        If lngResult = 0 Then
                                            .DeviceEnabled = True
                                        Else
                                            MsgBox "The till drawer is not available. This application will close(1)."
                                            bEnvironmentOK = False
                                        End If
                                        If .CapStatus = True Then
                                           If .DrawerOpened Then
                                               txtCloseDrawerMessage.Visible = True
                                               txtCloseDrawerMessage.ZOrder 0
                                               .WaitForDrawerClose 5000, 1000, 100, 1000
                                               txtCloseDrawerMessage.Visible = False
                                           End If
                                        End If
                                    Else
                                        MsgBox "The till drawer is not available. This application will close(2)."
                                        bEnvironmentOK = False
                                        Exit Sub
                                    End If
                                End With
                            End If
                        End If
                    Else
                        MsgBox "The till printer (" & oPC.Printername & ") cannot be claimed by the application." & vbCrLf & "Result is " & CStr(lngResult) & ". This application will close."
                        bEnvironmentOK = False
                        Exit Sub
                    End If
                Else
                    MsgBox "The till printer is not online. This application will close."
                    bEnvironmentOK = False
                    Exit Sub
                End If
            End With
        Else
''''''''''''''            If oPC.UseCashDrawer Then
''''''''''''''                lngResult = OPOSCashDrawerDigipos.Open(oPC.TillDrawerName)
''''''''''''''                If lngResult = 0 Then
''''''''''''''                    lngResult = OPOSCashDrawerDigipos.ClaimDevice(1000)
''''''''''''''                    If lngResult = 0 Then
''''''''''''''                        OPOSCashDrawerDigipos.DeviceEnabled = True
''''''''''''''                    Else
''''''''''''''                        MsgBox "The till drawer is not available. This application will close(3)."
''''''''''''''                        bEnvironmentOK = False
''''''''''''''                    End If
''''''''''''''                    If OPOSCashDrawerDigipos.CapStatus = True Then
''''''''''''''                       If OPOSCashDrawerDigipos.DrawerOpened Then
''''''''''''''                           txtCloseDrawerMessage.Visible = True
''''''''''''''                           txtCloseDrawerMessage.ZOrder 0
''''''''''''''                           OPOSCashDrawerDigipos.WaitForDrawerClose 5000, 1000, 100, 1000
''''''''''''''                           txtCloseDrawerMessage.Visible = False
''''''''''''''                       End If
''''''''''''''                    End If
''''''''''''''                Else
''''''''''''''                    MsgBox "The till drawer is not available. This application will close(4)."
''''''''''''''                    bEnvironmentOK = False
''''''''''''''                    Exit Sub
''''''''''''''                End If
''''''''''''''            End If
''''''''''''''            Set OPOSPrinter = OPOSPOSPrinterDigipos
''''''''''''''            With OPOSPOSPrinterDigipos
''''''''''''''                bIgnorestatus = True
''''''''''''''                lngResult = OPOSPOSPrinterDigipos.Open(oPC.Printername)
''''''''''''''                If lngResult = 0 Then
''''''''''''''                    lngResult = OPOSPOSPrinterDigipos.ClaimDevice(50)
''''''''''''''                    If lngResult = OPOS_SUCCESS Then
''''''''''''''                        OPOSPOSPrinterDigipos.DeviceEnabled = True
''''''''''''''                        OPOSPOSPrinterDigipos.MapMode = PTR_MM_METRIC
''''''''''''''                        OPOSPOSPrinterDigipos.RecLetterQuality = True
''''''''''''''                        OPOSPOSPrinterDigipos.RecLineChars = 40
''''''''''''''                    Else
''''''''''''''                        MsgBox "The till printer (" & oPC.Printername & ") cannot be claimed by the application." & vbCrLf & "Result is " & CStr(lngResult) & ". This application will close."
''''''''''''''                        bEnvironmentOK = False
''''''''''''''                        Exit Sub
''''''''''''''                    End If
''''''''''''''                Else
''''''''''''''                    MsgBox "The till printer is not online. This application will close."
''''''''''''''                    bEnvironmentOK = False
''''''''''''''                    Exit Sub
''''''''''''''                End If
''''''''''''''                bIgnorestatus = False
''''''''''''''           End With
        End If
   End If
   Me.lblState.Visible = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetupPrinter", , , , "line", Array(Erl())
End Sub
'Private Sub SetupCashDrawer()
'    On Error GoTo errHandler
'Dim lngResult As Long
'    If UCase(oPC.GetProperty("ReceiptPrinterType")) = "" Or UCase(oPC.GetProperty("ReceiptPrinterType")) = "EPSON" Then
'    Else
'    End If
'
'    Me.lblState.Visible = False
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SetupCashDrawer"
'End Sub
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LoadVoucherTypes"
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LoadDiscountTypes"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.GetDiscount(pCODE,pDescription)", Array(pCODE, pDescription)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetDiscount(pCODE,pDescription)", Array(pCODE, pDescription)
End Function
Private Function LogonOperator(Optional bAskForFloat As Boolean, Optional dblFLoat As Double, Optional sFloatBreakdown) As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean
Dim Res As Boolean
Dim f As frmGetFloat

    If bAskForFloat Then
        Set f = New frmGetFloat
        f.Show vbModal
        If f.IsCancelled Then
            bLogonOK = False
            Unload f
            Exit Function
        Else
            dblFLoat = f.FloatValue
            sFloatBreakdown = f.GetFloatBreakdown
            Unload f
        End If
    End If
    
    Res = False
    Do Until Res = True
        If Not SecurityControl(enSECURITY_ISOPERATOR, lngStaffID, strName, bCancelled, "Enter your signature.", "Your signature is invalid", True) Then
            If bCancelled Then Res = True
            bLogonOK = False
        Else
            Res = True
            bLogonOK = True
        End If
    Loop
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LogonOperator"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LogonOperator(bAskForFloat,dblFLoat,sFloatBreakdown)", Array(bAskForFloat, _
         dblFLoat, sFloatBreakdown)
End Function

Private Function SwapOperator() As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean

    If oPC.ZSession.OpSession.InSession Then
        oPC.ZSession.OpSession.Close_OP_Session False
    End If
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid", True) Then
        oPC.ZSession.OpSession.Start_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
    Else
        SetPresentState elogin
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SwapOperator"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SwapOperator"
End Function


Public Sub StartSale()
    On Error GoTo errHandler
    frmLoading = True
    bNoMoreSaleLines = False
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SetExchangeType eSaleType
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    lngOPID = 0
    bIssueCreditNote = False
    SetTitleBar False
    enPresentState = eStart
    enMode = emode_Sale
    PrepareForm
    frmLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StartSale"
End Sub



Private Sub Stat(msg As String)
    On Error GoTo errHandler
    SB.Caption = msg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Stat(msg)", msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Stat(msg)", msg
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim bCancelled As Boolean
Dim frm As frmOD
    On Error GoTo errHandler
    Me.ScaleMode = vbTwips
    SaveLayout Me.G1, Me.Name, Me.Height, Me.Width
    SaveLayout Me.G4, Me.Name & "B"

    If Not bForceClose Then
        If Not bCloseXsession And Not bCloseZsession And Not bLogonOK = False And bEnvironmentOK = True Then
            If bSaleOnHold Then
                If MsgBox("There is a transaction parked. Do you still want to close this application? Confirm", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
                    Cancel = True
                    Exit Sub
                End If
                oExchangeCopy1.CancelEdit
                Set oExchangeCopy1 = Nothing
            Else
                If MsgBox("You want to close this application? Confirm", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
                        Cancel = True
                        Exit Sub
                End If
            End If
        End If
    End If
    If bEnvironmentOK = True Then
        bUnloading = True
        ConnectionTimer.Enabled = False
        CloseApplication Cancel, bForceClose
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
    ErrorIn "frmPOSMain.Form_Unload(Cancel)", Cancel
End Sub

Private Sub CloseApplication(bCancel As Integer, Optional bForceClose As Boolean)
    On Error GoTo errHandler
Dim frm As frmOD
Dim bCancelled As Boolean
Dim LastNumber As Long
    lngSupervisorID = 0
    bCancel = False
    If bSaleActive And (Not bForceClose) Then
        If MsgBox("There is still a transaction in process!" & vbLf & _
                  "Do you want to close this application anyway?", _
                  vbYesNo, "Transaction In Process!") = vbNo Then
            bCancel = True
            Exit Sub
        Else
            If CheckThisPoint(M_CLOSEAPPLICATION) Then
                If SecurityControl(enSECURITY_ISSUPERVISOR, lngSupervisorID, strName, bCancelled, "Enter security code", "You are not entitled to Close the application.") = False Then
                        bCancel = True
                        Exit Sub
                End If
                Set frm = New frmOD
                frm.component "Closing application - provide reason"
                frm.Show vbModal
                If frm.Cancelled Then
                    Unload frm
                    bCancel = True
                    Exit Sub
                End If
                
                LastNumber = oExchange.ExchangeNumber
                AcceptSale True
                oExchange.SetExchangeType eVoidAction
                oExchange.ToVoid = CLng(LastNumber)
                oExchange.Note = "Closing application. " & frm.Reason
                oExchange.Note = oExchange.Note & "#" & CStr(LastNumber)
                oExchange.SupervisorID = lngSupervisorID
                AcceptSale True
                AddExchange
            Else
                RejectSale
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    Me.SB.Caption = "Wait. The local data is being transmitted to the server."
    
    If Not oExchange Is Nothing Then
        If oExchange.IsEditing Then oExchange.CancelEdit
    End If
    
    If bCloseXsession Then
        If Not oPC.ZSession.OpSession Is Nothing Then
            oPC.ZSession.OpSession.Close_OP_Session True
        End If
    End If
    
    If bCloseZsession Then
        If Not oPC.ZSession Is Nothing Then
            oPC.ZSession.Close_Z_Session
        End If
    End If
    
    If enPrinterType = en_Epson Then
        If Not Me.OPOSPOSPrinter Is Nothing Then
            With OPOSPOSPrinter
                .DeviceEnabled = False
                .ReleaseDevice
                .Close
            End With
        End If
    Else
''''''''''        If Not Me.OPOSPOSPrinterDigipos Is Nothing Then
''''''''''            With OPOSPOSPrinterDigipos
''''''''''                .DeviceEnabled = False
''''''''''                .ReleaseDevice
''''''''''                .Close
''''''''''            End With
''''''''''        End If
    End If
    
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CloseApplication(bCancel,bForceClose)", Array(bCancel, bForceClose)
End Sub


Private Sub mnuClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuClose_Click"
End Sub

Private Sub ShowExchange()
    On Error GoTo errHandler
Dim lngRow As Long
Dim lngTmp As Long
Dim oTmpExchange As a_Exchange
Dim strPos As String

    Set frmExchange = New frmExchange
    If IsNumeric(strSuffix) And Len(strSuffix) < 10 Then
        lngRow = CLng(strSuffix)
        If lngRow <= X4(1, 1) And lngRow > 0 Then 'X4(X4.UpperBound(1) - 1, 1) And lngRow > 0 Then
            lngTmp = X4.Find(1, 1, lngRow, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
            If lngTmp > 0 Then
                frmExchange.component X4(lngTmp, 10)
                frmExchange.Show vbModal
                If frmExchange.MustPrint = True And oPC.PrintSlips = True Then
                    Set oTmpExchange = oExchange
                    Set oExchange = New a_Exchange
                    oExchange.Load (X4(lngTmp, 10)), True
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
    ErrorIn "frmPOSMain.ShowExchange"
End Sub
Private Function IsVoucherUsedInSale() As Boolean
    On Error GoTo errHandler
Dim b As Boolean
Dim i As Integer
    b = False
    For i = 1 To oExchange.SaleLines.Count
        If InStr(1, UCase(oExchange.SaleLines(i).title), "VOUCHER") > 0 Then
            b = True
        End If
    Next
    For i = 1 To oExchange.PaymentLines.Count
        If oExchange.PaymentLines(i).PaymentType = "V" Or _
            oExchange.PaymentLines(i).PaymentType = "CV" Or _
            oExchange.PaymentLines(i).PaymentType = "CNR" Then
            b = True
        End If
    Next
    IsVoucherUsedInSale = b
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.IsVoucherUsedInSale"
End Function
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
            Stat "(X) Cancel,(CV)Credit voucher,(C)Cash,(CC)Card"
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
            txtVouchers.ZOrder 0
            Stat "  .. to reverse"
        Case ecancelsale
            setInputBox "", "", "", True
        Case eCollectRep
                setInputBox "***", "*", "", True
                lblInput.Caption = "Sales representative code."
                Stat "Use 'X' for no sales representative"
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
                lblInput.Caption = "Confirm sale" & IIf(strName > "" And bCollectRepcode = True, " for " & strName, "")
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
                Stat ""
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
            txtDiscounts.ZOrder 0
            
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
        Case ePaymentType_CreditCardRef
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
            txtVouchers.ZOrder 0
            
        Case ePrice
            lblInput.Caption = "Price"
            Stat "Hold shift key down and press Enter for discount"    ', '..' to reverse
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
                        Stat "Scan or (X)Cancel trans.,(C)Cash refund,(CC)Reverse credit card,(CV)Credit voucher,(Dn)Del sale,(DPn)Del payment " ', (PS) Park sale
                    Else
                        Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account,(Dn)Del prod,(DPn)Del payment,(DDP) Direct deposit"   ', (PS) Park sale
                    End If
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(AC)On account,(Dn)Del sale,(DPn)Del payment,(DDP) Direct deposit"   ', (PS) Park sale
                End If
            Else
                If oExchange.BalanceOwing < 0 Then
                    Stat "Scan or (X)Cancel trans.,(C)Cash refund,(CC)Reverse credit card,(CV)Issue credit voucher,(Dn)Del sale,(DPn)Del payment,(FC) Find customer"  ', (PS) Park sale
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,Q)Cheque,(AC)On account,(DDP) Direct deposit,(Dn)Del sale,(DPn)Del payment,(FC) Find customer"   ', (PS) Park sale
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrepareForm"
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
    X1.DeleteRows (X1.UpperBound(1) - iRow + 1)
    oExchange.CalculateTotals
    G1.ReOpen
    G1.Refresh
    RefreshAllSaleRows
    iCurrentSaleLine = iCurrentSaleLine - 1
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.RemoveSaleLine(iRow,pCurrent)", Array(iRow, pCurrent)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemoveSaleLine(iRow,pCurrent)", Array(iRow, pCurrent)
End Sub

Private Function Action_CancelSale() As eState
    On Error GoTo errHandler
Dim bCancelled As Boolean
Dim frm As frmOD
Dim LastNumber As Long
Dim oPL As a_Payment
Dim i As Integer
    lngSupervisorID = 0
    If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        If CheckThisPoint(M_CLOSEAPPLICATION) Then
            If SecurityControl(enSECURITY_ISSUPERVISOR, lngSupervisorID, strName, bCancelled, "Enter security code", "You are not entitled to cancel an exchange.") = False Then
                    Action_CancelSale = enPresentState
                    Exit Function
            End If
            Set frm = New frmOD
            frm.component "Cancelling exchange - provide reason"
            frm.Show vbModal
            If frm.Cancelled Then
                Unload frm
                bCancelled = True
                Action_CancelSale = enPresentState
                Exit Function
            End If
            
            LastNumber = oPC.ExchangeNumber    'oExchange.ExchangeNumber  most recent number - beware problems of parked sales being retrieved before a void
            If oExchange.PaymentLines.Count > 0 Then
                For i = oExchange.PaymentLines.Count To 1
                    oExchange.PaymentLines.Remove (i)
            Next i
            End If
            oExchange.SetExchangeType eVoidAction
            oExchange.ToVoid = CLng(LastNumber)
            oExchange.Note = "Cancelling exchange. " & frm.Reason
            oExchange.Note = oExchange.Note '& " #" & CStr(LastNumber)
            oExchange.OperatorID = 0
            oExchange.SupervisorID = lngSupervisorID
            AcceptSale True
            AddExchange
        Else
            RejectSale
        End If
    Else
        Action_CancelSale = enPresentState
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_CancelSale"
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
Private Sub CancelSale_Ex()
Dim EXCHID As String

    oExchange.ApplyEdit
    EXCHID = oExchange.ExchangeID
    Set oExchange = Nothing
    PrepareForNewSale
'    oExchange.ToVoid
'    oExchange.SupervisorID = lngStaffID
'    oExchange.Note = oExchange.Note & "#" & CStr(iToVoid)
'    oExchange.SetExchangeType eVoidAction
End Sub
Private Function AcceptSale(IsBeingCancelled As Boolean) As Boolean
    On Error GoTo errHandler
10        On Error GoTo errHandler
      Dim lngRow As Long
      Dim lngLowerBound As Long
      Dim oPayment As a_Payment
      Dim bCustomerOK As Boolean
      Dim strMsg As String
      Dim strPos As String
      Dim i As Integer
      Dim bSuccessfulReadOfMainDB As Boolean
      Dim reslt As String
      
          
20        bItemExchange = False
30        If oExchange.NeedsCustomerInfo = True And IsBeingCancelled = False And oExchange.transactionType <> "OR" Then
40            bCustomerOK = False
50            Do Until bCustomerOK = True
60                If GetCustomer() Then
70                    lblCustomername.Caption = DisplayCustomerDetails
80                End If
90                Set frmCustID = New frmIDCustomer
100               If oExchange.Note > "" Then
110                   frmCustID.component oExchange.Note
120               End If
130               frmCustID.Show vbModal
140               If oExchange.Customer.Name = "" Then
150                   oExchange.Note = frmCustID.CustomerName
160                   strMsg = "Confirm customer details:" & vbCrLf & "Name: " & frmCustID.CustomerName & vbCrLf & ""
170               Else
180                   oExchange.Note = vbNullString
190                   strMsg = "Confirm customer details:" & vbCrLf & "Name: " & oExchange.Customer.Name & vbCrLf & "A/c No.;" & oExchange.Customer.AcNo
200               End If
210               If oExchange.transactionType = "S" Then
220                   oExchange.Note = FNS(oExchange.Note) & "(" & frmCustID.Counterfoil & ")"
230               Else
240                   oSALELine.Counterfoil = frmCustID.Counterfoil
250               End If
260               If MsgBox(strMsg, vbInformation + vbYesNo) = vbNo Then
270                   ClearCustomer
280                   bCustomerOK = False
290               Else
300                   bCustomerOK = True
310               End If
320           Loop
330       End If
340   strPos = "06"
350       If oExchange.CustomerToBeCredited And IsBeingCancelled = False Then 'This is to determine in the case of an exchange (not a RDEP) if money is to go out
            'Replaced 8/8/6 so that credit card refunds and cash refunds are both described as rfunds  If oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_Cash) Then
360           If oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_Cash) Or oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditCard) Then
370               oExchange.SetExchangeType ereturntype
380           ElseIf oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditVoucher) Then
390               oExchange.SetExchangeType eCreditVoucherType
400           ElseIf oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditCard) Then
410               oExchange.SetExchangeType eSaleType
420           End If
430       End If
440   strPos = "07"
450       oExchange.OperatorID = lngOPID
460       oExchange.StaffName = strName
470       If iToVoid > 0 And IsBeingCancelled = False Then
480           oExchange.ToVoid = iToVoid
490           lngLowerBound = X4.LowerBound(1)
500           lngRow = X4.Find(lngLowerBound, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
510           strEXCHtoVoidGUID = X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 10)
              'mark all rows voided as such in case operator tries to void again
           '   MsgBox "Before AcceptSale lngRow = " & CStr(lngRow) & "    " & "lngLowerBound = " & CStr(lngLowerBound)
520           Do While lngRow >= lngLowerBound
              '  MsgBox "In AcceptSale lngRow = " & CStr(lngRow) & "    " & "lngLowerBound = " & CStr(lngLowerBound)
530               X4(lngRow, 12) = oExchange.ExchangeNumber
                  If lngRow = X4.UpperBound(1) Then
                    lngRow = 0
                    Exit Do
                  End If
540               If lngRow < X4.UpperBound(1) Then
550                   lngRow = X4.Find(lngRow + 1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
                  ElseIf lngRow = X4.UpperBound(1) Then
                      lngRow = X4.Find(lngRow, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG)
560               End If
570           Loop
580           G4.Refresh
590       End If
      'Check to see if a sale has an account component - if so the exchange type must be "A"
600       For i = 1 To oExchange.PaymentLines.Count
610           If oExchange.PaymentLines(i).PaymentType = "AC" And oExchange.transactionType <> "AR" Then  'A payment is being placed on account
620               oExchange.SetExchangeType eAccountSaleType
630           End If
640       Next i
650       reslt = oExchange.ApplyEdit
            If reslt = "calc" Then
                AcceptSale = False
                Exit Function
            End If
            
660       oPC.OpenLocalDatabase
670       oPC.DBLocalConn.BeginTrans
680       oExchange.PostExchange
690       oPC.DBLocalConn.CommitTrans
700       oPC.CloseLocalDatabase
      'Adds exchange to Xarraydb structure for display
710       AddExchange
720       SendPOSExchange oExchange.ExchangeID, oExchange.ZID    'oExchange.OPSID,

            If oExchange.CashTransaction Then
                OpenDrawer
            End If

      'Print Till Slip
730       If oPC.PrintSlips = True And IsBeingCancelled = False Then
740           Select Case oExchange.transactionType
              Case "S", "AR", "A", "CN"
750               If ((oExchange.Customer.CustomerType = oPC.GetLoyaltyCode) And oPC.GetLoyaltyCode > "") And (oExchange.LoyaltyValue > 0) Then
                      MsgBox "PLEASE CALL PAPYRUS SUPPORT:" & vbCrLf & " Ready to print loyalty voucher. Transaction type = " & oExchange.transactionType & vbCrLf & "Customer type = " & oExchange.Customer.CustomerType & vbCrLf & "Click OK to continue."
                      MsgBox "PLEASE CALL PAPYRUS SUPPORT:" & vbCrLf & " Ready to print loyalty voucher. Transaction type = " & oExchange.transactionType & vbCrLf & "Customer type = " & oExchange.Customer.CustomerType & vbCrLf & "Click OK to continue."
                      MsgBox "PLEASE CALL PAPYRUS SUPPORT:" & vbCrLf & " Ready to print loyalty voucher. Transaction type = " & oExchange.transactionType & vbCrLf & "Customer type = " & oExchange.Customer.CustomerType & vbCrLf & "Click OK to continue."
760                   PrintLoyaltyVoucher
770               End If
780               If oExchange.transactionType = "A" Then
790                   PrintSalesSlip oPC.AccountSaleCopyCount
800               Else
810                   PrintSalesSlip oPC.InvoiceCopyCount
820               End If
830           Case "R"
840               PrintSalesSlip oPC.ReturnCopyCount
850           Case "PC", "PCC"
860               PrintPettyCashVoucher oPC.PettyCashCopyCount
870           Case "C"
880               PrintSalesSlip oPC.CreditNoteCopyCount
890           Case "DEP"
900               PrintDepositSlip oPC.DepositCopyCount
910           Case "RDEP"
920               PrintDepositRefundSlip oPC.DepositCopyCount
930           Case "APP"
940               PrintAPPROSlip oPC.ApproCopyCount
950           Case "OR"
960               PrintORDERSlip oPC.OrderCopyCount
970           Case "PA"
980               PrintReceiptSlip oPC.ReceiptCopyCount
990           Case "V"   'Print a copy of the voided transaction
1000              If strEXCHtoVoidGUID > "" Then
1010                  Set oTmpExchange = oExchange
1020                  Set oExchange = New a_Exchange
1030                  oExchange.LoadFromMainDB strEXCHtoVoidGUID, True, bSuccessfulReadOfMainDB
1040                  If bSuccessfulReadOfMainDB Then
1050                      If oExchange.TransactionTypeEnum = eOrderRequestType Then
1060                          PrintORDERSlip 1, True, True
1070                      ElseIf oExchange.TransactionTypeEnum = ePettyCashType Then
1080                          PrintPettyCashVoucher 1, , True
1090                      ElseIf oExchange.TransactionTypeEnum = ePettyCashCreditType Then
1100                          PrintPettyCashVoucher 1, , True
1110                      Else
1120                          PrintSalesSlip 1, False, , True, oTmpExchange.Note
1130                      End If
1140                  Else
1150                      MsgBox "Cannot connect to main database at this time and so cannor reprint the voided exchange. Do a reprint later", vbInformation, "Can't do this"
1160                  End If
1170                  Set oExchange = Nothing
1180                  Set oExchange = oTmpExchange
1190                  Set oTmpExchange = Nothing
1200              End If
1210          End Select
      
          'If there is a CV being paid out as change - we must print it
1220          If bIssueCreditNote And IsBeingCancelled = False Then
                  'print an extra copy of the exchange
1230              PrintSalesSlip 1
1240              PrintCNasChange oExchange.ChangeVoucherValueF, oPC.CreditNoteCopyCount, False
1250              bIssueCreditNote = False
1260          End If
1270      End If
          If oPC.PrintSlips = True Then
             If IsBeingCancelled Then
                 PrintSalesSlip 1, False, , True
             End If
          End If
1280      Set oExchange = Nothing
1290      PrepareForNewSale
1300      strEXCHtoVoidGUID = ""
          AcceptSale = True
1310      Exit Function
'errHandler:
'1320      If ErrMustStop Then Debug.Assert False: Resume
'1330      ErrorIn "frmPOSMain.AcceptSale", , , , "EXCH:IsEditing,strPOS,line", Array(oExchange.IsEditing, strPos, Erl())
    Exit Function
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AcceptSale"
End Function
Private Sub PrepareForNewSale()
10        On Error GoTo errHandler
20        On Error GoTo errHandler
30        Set oExchange = New a_Exchange
40        oExchange.BeginEdit
50        oExchange.SupervisorID = oPC.ZSession.OpSession.OperatorID
          
60        oExchange.SetExchangeType eSaleType
70        ClearTextFields
80        X1.Clear
90        X1.ReDim 1, 0, 0, 8
100       G1.ReBind
110       X2.Clear
120       X2.ReDim 1, 1, 1, 3
130       G2.ReBind
140       txtInput.BackColor = RGB(230, 250, 210)
          
150       lblCustomername.Caption = vbNullString
160       lblReplacement.Visible = False
170       iCurrentSaleLine = 0
180       iCurrentPaymentLine = 0
190       iToVoid = 0
200       lngOPID = 0
210       lngSupervisorID = 0
220       bSaleActive = False
230       bCustomerVisible = False
240       SetPresentState eStart
250       enMode = emode_Sale
260       SetTitleBar True
270       SetForCOLSVisible False
280       strName = ""
          
290       Exit Sub
'errHandler:
'260       If ErrMustStop Then Debug.Assert False: Resume
'270       ErrorIn "frmPOSMain.PrepareForNewSale"
300       Exit Sub
errHandler:
310       If ErrMustStop Then Debug.Assert False: Resume
320       ErrorIn "frmPOSMain.PrepareForNewSale"
End Sub
Private Sub ClearTextFields()
    On Error GoTo errHandler
    txtExtTotal = ""
    txtQtyTotal = ""
    txtVatValue = ""
    txtPaymentTotal = ""
    lblUpdate.Visible = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ClearTextFields"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ClearPayments"
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
    X1.ReDim 1, 0, 0, 8
    G1.ReBind
    DisplayTotals
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ClearSaleLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearSaleLines"
End Sub
Private Function LoadProductFromCode(pIn As String) As Boolean
    On Error GoTo errHandler
10        On Error GoTo errHandler
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
   '   Dim lngLoyaltyDiscount As Long
      Dim bIdentifyCustomer As Boolean
      Dim bNoDiscountAllowable As Boolean
      Dim rs1 As ADODB.Recordset
      Dim bDeleteFromHere As Boolean
      Dim i As Integer
      Dim bSuccess As Boolean

20        Set oLU = New z_Lookup
30        strEAN = pIn ' Trim$(txtInput)
40        strCode = pIn  'Trim$(txtInput)
50        Set rs = oLU.GetProduct(strEAN, strCode, oExchange.Customer.CustomerTypeRaw, bSuccess)
'60        If bSuccess = False Then
'70            MsgBox "Cannot connect to database. Cancel this transaction and inform your supervisor." & vbCrLf & "You can try closing this application and starting it again to clear the error.", vbCritical + vbOKOnly, "Warning"
'80            Set rs = Nothing
'90            Exit Function
'100       End If
110       If rs Is Nothing Then
120           LoadProductFromCode = False
130           Set oLU = Nothing
140           Exit Function
150       ElseIf rs.State = 0 Then
160           LoadProductFromCode = False
170           Set rs = Nothing
180           Set oLU = Nothing
190           Exit Function
200       ElseIf rs.RecordCount = 0 Then
210           LoadProductFromCode = False
220           rs.Close
230           Set rs = Nothing
240           Set oLU = Nothing
250           Exit Function
260       End If

      'Create new sale line for Exchange
      itest = itest = 1
270       Set oSALELine = oExchange.SaleLines.Add
280       iCurrentSaleLine = iCurrentSaleLine + 1
290   '    X1.ReDim 1, iCurrentSaleLine, 0, 8
          
      'Load rules into Sales line
300       oSALELine.LoadRules rs, enPresentState = eAppro     'TEMPRORARY
310       Set rs = Nothing
320       Set oLU = Nothing
          
330       oSALELine.FindRule oExchange.NominalValue, bIdentifyCustomer

340       oExchange.IdentifyCustomer = bIdentifyCustomer
350       If oExchange.IdentifyCustomer = True And oExchange.Note = "" Then
360           Set frmCustID = New frmIDCustomer
370           frmCustID.component oExchange.Note
380           frmCustID.Show vbModal
390           oExchange.Note = frmCustID.CustomerName
400           oSALELine.Counterfoil = frmCustID.Counterfoil
410           Unload frmCustID
420       End If
           
           
430        oSALELine.ApplyEdit
440        oSALELine.BeginEdit
450        LoadProductFromCode = True
           
460        oPC.CloseLocalDatabase
           
470       Exit Function
'errHandler:
'480       If ErrMustStop Then Debug.Assert False: Resume
'490       ErrorIn "frmPOSMain.LoadProductFromCode"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadProductFromCode(pIn)", pIn
End Function
Private Sub oExchange_Recalculate(Redisplay As Boolean)
    On Error GoTo errHandler
    If bUnloading Then Exit Sub
    'If Redisplay Then
        RefreshAllSaleRows
    'End If
    G1.ReBind
    G1.Refresh
    DisplayTotals
    lblCustomername.Caption = DisplayCustomerDetails

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_Recalculate"
End Sub
Private Sub oExchange_GetCustomer()
    On Error GoTo errHandler
    If (Not oExchange.Customer.ID > 0) And (oExchange.Note = "") Then
        GetCustomer
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_GetCustomer"
End Sub

Private Sub RefreshRules()
    On Error GoTo errHandler
Dim oLU As z_Lookup
Dim oSL As a_Sale
Dim rs As ADODB.Recordset
    Set oLU = New z_Lookup
    For Each oSL In oExchange.SaleLines
       ' Call oLU.GetProduct(strEAN, strCode, strPID, strTitle, strAuthor, lngVatrate, lngPrice, lngDiscount, lngLoyaltyDiscount, bIdentifyCustomer, bNoDiscountAllowable, strDiscountRule)
        Set rs = oLU.GetProduct(oSL.Code, oSL.Code, oExchange.Customer.CustomerTypeRaw)
        If rs.RecordCount = 0 Then
            rs.Close
            Set rs = Nothing
            Set oLU = Nothing
            Exit Sub
        End If
    'Load rules into Sales line
        oSL.LoadRules rs, oExchange.TransactionTypeEnum = eApproType   'Temporary
        Set rs = Nothing
        oSL.FindRule oExchange.NominalValue, False
    Next
    Set oLU = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefreshRules"
End Sub
Private Function GetCustomer() As Boolean
    On Error GoTo errHandler
Dim frm As New frmBrowseCustomers2
Dim cnt As Integer
Dim s As String

    GetCustomer = False
    frm.Show vbModal
    If frm.IsCancelled Then Exit Function
    If frm.CustomerName > "" Then
        bCustomerVisible = True
        strCustomername = frm.CustomerName
        G3.Caption = frm.CustomerName
        lngCustomerID = frm.CustomerID
        oExchange.SetCustomer lngCustomerID
        oExchange.Note = frm.CustomerName
        GetCustomer = True
        RefreshRules
        oExchange.CalculateTotals
        LookForAlert frm.Accnum, frm.CustomerName
    Else
        GetCustomerFromMasterDB_control frm.Accnum, cnt
        ClearCustomer
        If cnt > 0 Then
            MsgWaitObj 2000
            MsgBox "Attempted to fetch missing customer record from master database. Please try again.", vbInformation + vbOKOnly, "Customer record missing on local database"
            Exit Function
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.GetCustomer", , , , "s", Array(s)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetCustomer"
End Function
Private Sub LookForAlert(TPACNO As String, CustomerName As String)
    On Error GoTo errHandler
Dim frm As frmAlert
Dim oLU As New z_Lookup
Dim rs As ADODB.Recordset

    
    Set rs = oLU.GetAlertFromMasterDB(TPACNO)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            Set frm = New frmAlert
            frm.component rs
    '        frm.lblCustomer.Caption = CustomerName & "  (" & TPACNO & ")"
    '        frm.lblMsg.Caption = FNS(rs.Fields("AL_MSGTEXT"))
            frm.Show vbModal
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LookForAlert(TPACNO,CustomerName)", Array(TPACNO, CustomerName)
End Sub
Private Sub GetCustomerFromMasterDB_control(AcNo As String, cnt As Integer)
    On Error GoTo errHandler
Dim oLU As New z_Lookup
    oLU.GetCustomerFromMasterDB AcNo, cnt
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.GetCustomerFromMasterDB_control(AcNo,cnt)", Array(AcNo, cnt)
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ClearCustomer"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearCustomer"
End Function
Private Function DisplayCustomerDetails() As String
    On Error GoTo errHandler
Dim strDetails As String
Dim ar() As String
Dim i As Long
Dim bFound As Boolean

    ' LogSaveToFile "Customer type = " & Replace(UCase(oExchange.Customer.CustomerType), "*", "")
    ' LogSaveToFile "GetLoyaltyCode = " & oPC.GetLoyaltyCode
    strDetails = ""
    bFound = False
    ar() = Split(oExchange.Customer.CustomerType, ",")
    For i = 0 To UBound(ar)
        If oPC.GetLoyaltyCode = ar(i) Then
            strDetails = oExchange.Customer.NameAndCode(99) & " " & "Loyalty value: " & oExchange.LoyaltyValueF
            bFound = True
            Exit For
        End If
    Next
    If Not bFound Then
        For i = 0 To UBound(ar)
            If oPC.GetBookClubCode = ar(i) Then
            strDetails = oExchange.Customer.NameAndCode(99) & " " & "(Book club)" & ":" & oExchange.DiscountRateF
                bFound = True
                Exit For
            End If
        Next
    End If
    If Not bFound Then
            strDetails = oExchange.Customer.NameAndCode(99)
    End If
'    Select Case oExchange.Customer.CustomerType
'   ' Case "L1"   'Loyalty club 1 member
'    Case oPC.GetLoyaltyCode
'            strDetails = oExchange.Customer.NameAndCode(99) & " " & "Loyalty value: " & oExchange.LoyaltyValueF
'    Case oPC.GetBookClubCode
'        strDetails = oExchange.Customer.NameAndCode(99) & " " & "(Book club)" & ":" & oExchange.DiscountRateF
'    Case ""
'        strDetails = oExchange.Customer.NameAndCode(99)
'    End Select
    
    DisplayCustomerDetails = strDetails & vbCrLf & strCreditLimitExceededMessage
    
    If strDetails = "" Then MsgBox "Customer details missing", vbInformation + vbOKOnly, "Warning"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.FetchCOLS"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.FetchCOLS"
End Sub
Private Function LoadCOLS() As Boolean
    On Error GoTo errHandler
Dim i As Long

    G3.Visible = True
    X3.Clear
    X3.ReDim 0, 0, 1, 14
    For i = 1 To cCOLS.Count
        With cCOLS(i)
            
            If IsCOLIDAlreadyInPayments(cCOLS(i).COLID) = False Then
            X3.ReDim 1, cCOLS.Count, 1, 14
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
            ' if we want EAN to display this is where would have to put it

            End If
        End With
    Next
  '  X3.QuickSort 1, X3.UpperBound(1), 10, XORDER_DESCEND, XTYPE_STRING   'sorted in query - we need the ordinal position (column 1) to be in sequence
 
        G3.Array = X3
        Me.G3.ReBind
    If X3.UpperBound(1) > 0 Then
        G3.Bookmark = 1
        SetForCOLSVisible True
        LoadCOLS = True
    Else
        SetForCOLSVisible False
      '  MsgBox "There are no unused deposits"
        LoadCOLS = False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LoadCOLS"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadCOLS"
End Function

Private Function IsCOLIDAlreadyInPayments(COLID As Long)
    For Each oPAYMENTLine In oExchange.PaymentLines
        If oPAYMENTLine.COLID = COLID Then
            IsCOLIDAlreadyInPayments = True
            Exit Function
        End If
    Next
     IsCOLIDAlreadyInPayments = False
End Function
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




Private Sub Connect()
    On Error GoTo errHandler
'    Set oSQLServer = New SQLDMO.SQLServer
'    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
'    With oSQLServer
'        .LoginSecure = False
'        .AutoReConnect = False
'        .Connect oPC.LocalSQLServerName, "sa", ""
'    End With
'
'    Set oDatabase = oSQLServer.Databases("PBKSFD")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Connect"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Connect"
End Sub

Private Sub RebuildIndexes()
    On Error GoTo errHandler
'Dim oTable As SQLDMO.Table
'    For Each oTable In oDatabase.Tables
'        If Not oTable.SystemObject Then oTable.RebuildIndexes
'    Next
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.RebuildIndexes"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RebuildIndexes"
End Sub
Private Function Disconnect()
    On Error GoTo errHandler
 '   oSQLServer.Disconnect
    Set ADOConn = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Disconnect"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Disconnect"
End Function

Private Sub PrintSalesSlip(pCopyCount As Integer, Optional bReprint As Boolean, Optional toA4 As Boolean = False, Optional bBeingVoided As Boolean, Optional strNote As String)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
Dim strErrPos As String
    Dim strDiscountDescription As String
    Dim lValue As Long
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim SType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
    Dim sDiscDesc As String
    Dim sCounterfoil As String
    Dim bPriceAlteration As Boolean
    Dim iLineCount As Integer
' When outputting to a printer,a mouse cursor becomes like a hourglass.

    If IsVoucherUsedInSale And oPC.GetProperty("ExtraSlipForVoucherExchanges") = "TRUE" Then
        pCopyCount = pCopyCount + 1
    End If
    MousePointer = vbHourglass
            BcData = "4902720005074"
strErrPos = "1"
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
                        PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint
                        iLineCount = 0
                        If oExchange.SaleLines.Count > 0 Then
                            For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                                If iLineCount > 24 Then Printer.NewPage
                                sAt = idBuf(i).At
                                sBuf = idBuf(i).Name
                                sExt = idBuf(i).Ext
                                SType = idBuf(i).TType
                                sDisc = idBuf(i).Disc
                                sDiscDesc = idBuf(i).DiscDesc
                                bPriceAlteration = idBuf(i).Alteration
                                sCounterfoil = idBuf(i).Counterfoil
                                
                                sValue = MakePrintStringDetail(50, SType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
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
                        PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter           'print totals
                        PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                        Printer.EndDoc
               Else
strErrPos = "2"
                    With OPOSPrinter
                        PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint, bBeingVoided, strNote    'Print header
strErrPos = "3"
                        
                        If oExchange.SaleLines.Count > 0 Then
                            For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                                If .ResultCode <> OPOS_SUCCESS Then Exit For
                                sAt = idBuf(i).At
                                sBuf = idBuf(i).Name
                                sExt = idBuf(i).Ext
                                SType = idBuf(i).TType
                                sDisc = idBuf(i).Disc
                                sDiscDesc = idBuf(i).DiscDesc
                                bPriceAlteration = idBuf(i).Alteration
                                sCounterfoil = idBuf(i).Counterfoil
                                
                                sValue = MakePrintStringDetail(.RecLineChars, SType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                                strDiscountDescription = ""
                                If sDiscDesc > "" Then
                                    strDiscountDescription = Left(sDiscDesc, 20) & " " & sDisc
                                Else
                                    strDiscountDescription = "Disc: " & sDisc
                                End If
                                .PrintNormal PTR_S_RECEIPT, ESC + "|N"
                                .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "") & vbLf
                                If sCounterfoil > "" Then
                                    .PrintNormal PTR_S_RECEIPT, "Ref: " & sCounterfoil & vbLf
                                End If
                            Next
                        End If
                        .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf     'create gap
strErrPos = "4"
                            
                        PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter           'print totals
                        PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                        
                     '   .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
                        .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf    'create gap
                        .CutPaper 90
            
                        .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
                
strErrPos = "5"
                        'Back to the synchronous mode
                        .AsyncMode = False
 '                   MsgBox "Just finished printing claimed printer"
                    End With
                End If
            Next
        Printer.EndDoc
strErrPos = "6"

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintSalesSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint), , , "strErrpos", Array(strErrPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintSalesSlip(pCopyCount,bReprint,toA4)", Array(pCopyCount, bReprint, toA4)
End Sub
Private Sub PrintReceiptSlip(pCopyCount As Integer, Optional bReprint As Boolean)
    On Error GoTo errHandler
Dim i As Integer
Dim c As Integer
Dim strErrPos As String

    Dim strDiscountDescription As String
    Dim lValue As Long
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim SType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
    Dim sDiscDesc As String
    Dim sCounterfoil As String
    Dim bPriceAlteration As Boolean
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
strErrPos = "1"
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
        ''''
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint      'Print header
                
                sValue = MakePrintStringDetail(iColWidth, "Payment received ", "", "", oExchange.PaymentLines(1).AmtF, 0, False)
                Printer.Print ""
                Printer.Print sValue
                    
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                
                Printer.Print ""
                Printer.EndDoc
        ''''
        Else
strErrPos = "2"

            With OPOSPrinter
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint      'Print header
strErrPos = "3"
                
                sValue = MakePrintStringDetail(.RecLineChars, "Payment received ", "", "", oExchange.PaymentLines(1).AmtF, 0, False)
                .PrintNormal PTR_S_RECEIPT, vbCrLf    'create gap
strErrPos = "4"
                    
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf      'create gap
                    
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf    'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
                
            End With
        End If
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintReceiptSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint), , , "strErrPos", Array(strErrPos)
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
    Dim SType As String
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
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint      'Print header
                For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                    sAt = idBuf(i).At
                    sBuf = idBuf(i).Name
                    sExt = idBuf(i).Ext
                    SType = idBuf(i).TType
                    sDisc = idBuf(i).Disc
                    sDiscDesc = idBuf(i).DiscDesc
                    bPriceAlteration = idBuf(i).Alteration
                    sValue = MakePrintStringDetail(iColWidth, SType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
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
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                
                Printer.Print ""
                Printer.EndDoc
       Else
            With OPOSPrinter
                PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint      'Print header
                For i = LBound(idBuf) To UBound(idBuf)          'Print each line
                    If .ResultCode <> OPOS_SUCCESS Then Exit For
                    sAt = idBuf(i).At
                    sBuf = idBuf(i).Name
                    sExt = idBuf(i).Ext
                    SType = idBuf(i).TType
                    sDisc = idBuf(i).Disc
                    sDiscDesc = idBuf(i).DiscDesc
                    bPriceAlteration = idBuf(i).Alteration
                    sValue = MakePrintStringDetail(.RecLineChars, SType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                    strDiscountDescription = ""
                    If sDiscDesc > "" Then
                        strDiscountDescription = Left(sDiscDesc, 20) & " " & sDisc
                    Else
                        strDiscountDescription = "Disc: " & sDisc
                    End If
                    .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(i).CodeF + " " + IIf(sDisc > "", strDiscountDescription, "") & vbLf
                Next
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf     'create gap
                    
                PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter           'print totals
                PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
                
            End With
        End If
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintAPPROSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintAPPROSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
End Sub
Private Sub PrintORDERSlip(pCopyCount As Integer, Optional bReprint As Boolean, Optional bBeingVoided As Boolean)
10        On Error GoTo errHandler
      Dim i As Integer
      Dim c As Integer
          Dim strDiscountDescription As String
          Dim lValue As Long
          Dim idBuf() As ITEMDATA
          Dim fDate As String
          Dim BcData  As String
          Dim sBuf As String
          Dim sExt As String
          Dim SType As String
          Dim sDisc As String
          Dim sAt As String
          Dim sValue As String
          Dim sDiscDesc As String
          Dim bPriceAlteration As Boolean
          Dim ar() As String
          Dim k As Integer
          Dim ar2() As String
          Dim ar3() As String
          
      ' When outputting to a printer,a mouse cursor becomes like a hourglass.
20        MousePointer = vbHourglass

30        BcData = "4902720005074"
          
40        For c = 1 To pCopyCount
50            If oPC.UseA4Printer Then
60                    PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint, bBeingVoided     'Print header
70                    Printer.Print ""
                      
                      
80                    ar = Split(strOrderedTitle, "~")
                      
90                    Printer.Print "Customer:"
100                   sBuf = ar(1) & " " & ar(2) & " " & ar(3)
110                   Printer.Print sBuf
                      
120                   If ar(4) > "" Then
130                       sBuf = "Phone: " & ar(4)
140                       Printer.Print sBuf
150                   End If
160                   If ar(5) > "" Then
170                       sBuf = "Email: " & ar(5)
180                       Printer.Print sBuf
190                   End If
                      
200                   sBuf = ar(6)
210                   Printer.Print sBuf
                      
220                   If ar(0) > "" Then
230                       sBuf = "Account number: " & ar(0)
240                       Printer.Print sBuf
250                   End If
                      
260                   Printer.Print ""
                      
270                   ar2 = Split(ar(8), "|")
280                   For k = 1 To UBound(ar2)
290                       ar3 = Split(ar2(k), "^^")
300                       sBuf = ar3(0) & " " & Left(ar3(1), 25)
310                       Printer.Print sBuf
320                   Next
                      
330                   If ar(7) > "" Then
340                       sBuf = "Notes: " & ar(7)
350                       Printer.Print sBuf
360                   End If
                      
370                   Printer.Print ""
                      
380                   PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint          'print totals
390                   PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
400                   Printer.Print ""
410                   Printer.EndDoc
420           Else
430               With OPOSPrinter
440                   PrintHeader ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint, bBeingVoided     'Print header
                      
          '           .PrintNormal PTR_S_RECEIPT, vbCrLf     'create gap
                      
450                   ar = Split(strOrderedTitle, "~")
                      
460                   .PrintNormal PTR_S_RECEIPT, vbCrLf + "Customer:" + vbLf
470                   sBuf = ar(1) & " " & ar(2) & " " & ar(3)
480                   .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
                      
490                   If ar(4) > "" Then
500                       sBuf = "Phone: " & ar(4)
510                       .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
520                   End If
530                   If ar(5) > "" Then
540                       sBuf = "Email: " & ar(5)
550                       .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
560                   End If
                      
570                   sBuf = ar(6)
580                   .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
                      
590                   If ar(0) > "" Then
600                       sBuf = "Account number: " & ar(0)
610                       .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
620                   End If
                      
630                   .PrintNormal PTR_S_RECEIPT, vbCrLf  'create gap
                      
640                   ar2 = Split(ar(8), "|")
650                   For k = 1 To UBound(ar2)
660                       ar3 = Split(ar2(k), "^^")
670                       sBuf = ar3(0) & " " & Left(ar3(1), 25)
680                       .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
690                   Next
                      
700                   If ar(7) > "" Then
710                       sBuf = "Notes: " & ar(7)
720                       .PrintNormal PTR_S_RECEIPT, sBuf + vbLf
730                   End If
                      
740                   .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf   'create gap
                      
750                   PrintTotals ConvertToType(oExchange.transactionType), OPOSPrinter, bReprint          'print totals
760                   PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
                      
770                   .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf     'create gap
780                   .CutPaper 90
790                   .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
              
                      'Back to the synchronous mode
800                   .AsyncMode = False
810               End With
820           End If
830       Next

      ' When a cursor is back to its default shape, it means the process ends.
840       MousePointer = vbDefault
850       Printer.EndDoc
      'errHandler:
      '    If ErrMustStop Then Debug.Assert False: Resume
      '    ErrorIn "frmPOSMain.PrintAPPROSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
860       Exit Sub
errHandler:
870       If ErrMustStop Then Debug.Assert False: Resume
880       ErrorIn "frmPOSMain.PrintORDERSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint), , , "Error line", Array(Erl)
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
Dim SType As String
Dim sDisc As String
Dim sAt As String
Dim sValue As String
Dim bPriceAlteration As Boolean
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader ConvertToType("CV"), OPOSPrinter, bReprint      'Print header
                Printer.Print ""
                    
                With OPOSPrinter
                    sBuf = "Change Voucher"
                    sExt = pAmtF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    Printer.Print sValue
                End With
                PrintFooter c, ConvertToType("C"), OPOSPrinter          'print footer
                Printer.Print ""
                Printer.EndDoc
       Else
            With OPOSPrinter
                PrintHeader ConvertToType("CV"), OPOSPrinter, bReprint      'Print header
                
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf     'create gap
                    
                With OPOSPrinter
                    sBuf = "Change Voucher"
                    sExt = pAmtF
                    sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                End With
                PrintFooter c, ConvertToType("C"), OPOSPrinter          'print footer
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                .AsyncMode = False
            End With
        End If
    Next

    MousePointer = vbDefault

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintCNasChange(pAmtF,pCopyCount,bReprint)", Array(pAmtF, pCopyCount, _
'         bReprint)
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
    Dim SType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String

' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    If oPC.UseA4Printer Then
            PrintHeader eTypVoucher, OPOSPrinter           'Print header
            Printer.Print ""
            Printer.Print "Credit value: " & oExchange.LoyaltyValueF
            Printer.Print ""
            PrintFooter 1, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
            Printer.Print ""
            Printer.EndDoc
    Else
        With OPOSPrinter
            PrintHeader eTypVoucher, OPOSPrinter           'Print header
            .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf   'create gap
           
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "Credit value: " & oExchange.LoyaltyValueF + vbLf
            .PrintNormal PTR_S_RECEIPT, vbCrLf     'create gap
                
            PrintFooter 1, ConvertToType(oExchange.transactionType), OPOSPrinter          'print footer
            
            .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf    'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
            
        End With
    End If
' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintLoyaltyVoucher"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintLoyaltyVoucher"
End Sub
Private Sub PrintPettyCashVoucher(pCopyCount As Integer, Optional bReprint As Boolean, Optional bBeingVoided As Boolean)
    On Error GoTo errHandler
    Dim lValue As Long
    Dim i As Integer
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim SType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String
Dim c As Integer
' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For c = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader eTypPettyCash, OPOSPrinter, bReprint, bBeingVoided         'Print header
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
            With OPOSPrinter
                PrintHeader eTypPettyCash, OPOSPrinter, bReprint, bBeingVoided         'Print header
                .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf + vbCrLf     'create gap
                If oExchange.PaymentLines(1).PaymentType = "W" Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash: " & oExchange.PaymentLines(1).AmtF + vbLf
                Else
                    .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash Refund: " & oExchange.PaymentLines(1).AmtF + vbLf
                End If
                .PrintNormal PTR_S_RECEIPT, vbCrLf     'create gap
                .PrintNormal PTR_S_RECEIPT, oExchange.Note
                
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf    'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
        
                'Back to the synchronous mode
                .AsyncMode = False
            End With
        End If
    Next
' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintPettyCashVoucher(pCopyCount)", pCopyCount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintPettyCashVoucher(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
End Sub

Private Function MakePrintStringDetail(ByVal lRecLineChars As Long, SType As String, sBuf As String, sAt As String, sExt As String, sDisc As String, PriceAlteration As Boolean) As String
    On Error GoTo errHandler
Dim sValue As String
Dim strNotChangeable As String
Dim iAvailable As Integer
    sAt = " " & sAt
    sExt = " " & sExt
    iAvailable = lRecLineChars - Len(sAt) - Len(SType) - Len(sExt)
    If PriceAlteration = True Then
        sAt = ESC + "|uC" & sAt
    End If
    sBuf = Left(sBuf, iAvailable)
    sBuf = sBuf & Space(iAvailable - Len(sBuf))
    If oPC.UseA4Printer Then
        MakePrintStringDetail = SType & sBuf & sAt & sExt
    Else
        MakePrintStringDetail = SType & sBuf & sAt & ESC + "|N" & sExt
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.MakePrintStringDetail(lRecLineChars,sType,sBuf,sAt,sExt,sDisc," & _
'        "PriceAlteration)", Array(lRecLineChars, SType, sBuf, sAt, sExt, sDisc, PriceAlteration)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.MakePrintStringDetail(lRecLineChars,SType,sBuf,sAt,sExt,sDisc," & _
        "PriceAlteration)", Array(lRecLineChars, SType, sBuf, sAt, sExt, sDisc, PriceAlteration)
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
    Dim SType As String
    Dim sDisc As String
    Dim sAt As String
    Dim sValue As String

' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    For i = 1 To pCopyCount
        If oPC.UseA4Printer Then
                PrintHeader eTypDepositRefund, OPOSPrinter           'Print header
                Printer.Print ""      'create gap
               
                Printer.Print "REFUNDED.: " & oExchange.PaymentLines(1).AmtF_nonNegative
                Printer.Print strDepositTitle
                Printer.Print ""      'create gap
                PrintFooter i, eTypDepositRefund, OPOSPrinter          'print footer
                Printer.Print ""      'create gap
                Printer.EndDoc
        Else
            With OPOSPrinter
                PrintHeader eTypDepositRefund, OPOSPrinter           'Print header
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf     'create gap
               
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "REFUNDED.: " & oExchange.PaymentLines(1).AmtF_nonNegative + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf & strDepositTitle
                .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf       'create gap
              '  .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(i)
                PrintFooter i, eTypDepositRefund, OPOSPrinter          'print footer
                
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf    'create gap
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


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintDepositRefundSlip(pCopyCount)", pCopyCount
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
Dim SType As String
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
                PrintHeader eTypDeposit, OPOSPrinter             'Print header
                Printer.Print ""      'create gap
                For j = LBound(idBuf) To UBound(idBuf)          'Print each line
                    sAt = idBuf(j).At
                    sBuf = idBuf(j).Name
                    sExt = idBuf(j).Ext
                    SType = idBuf(j).TType
                    sDisc = idBuf(j).Disc
                    bPriceAlteration = idBuf(j).Alteration
                    sValue = MakePrintStringDetail(iColWidth, SType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    Printer.Print sValue
                    Printer.Print oExchange.SaleLines(1).CodeF & ":DEPOSIT PAID"
                Next j
                Printer.Print "Deposit paid: " & oExchange.TotalPayableF
                Printer.Print "Change: " & oExchange.ChangeGivenF
                Printer.Print ""      'create gap
                Printer.Print "Copy number: " & CStr(i)
                PrintFooter i, eTypDeposit, OPOSPrinter          'print footer
                Printer.Print ""     'create gap
                Printer.EndDoc
        Else
            With OPOSPrinter
                PrintHeader eTypDeposit, OPOSPrinter             'Print header
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf      'create gap
                For j = LBound(idBuf) To UBound(idBuf)          'Print each line
                    If .ResultCode <> OPOS_SUCCESS Then Exit For
                    sAt = idBuf(j).At
                    sBuf = idBuf(j).Name
                    sExt = idBuf(j).Ext
                    SType = idBuf(j).TType
                    sDisc = idBuf(j).Disc
                    bPriceAlteration = idBuf(j).Alteration
                    sValue = MakePrintStringDetail(.RecLineChars, SType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                    .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                    .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(1).CodeF & ":DEPOSIT PAID" & vbLf
                Next j
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Deposit paid: " & oExchange.TotalPayableF + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Change: " & oExchange.ChangeGivenF + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf     'create gap
                .PrintNormal PTR_S_RECEIPT, vbCrLf & "Copy number: " & CStr(i)
                PrintFooter i, eTypDeposit, OPOSPrinter          'print footer
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf     'create gap
                .CutPaper 90
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
                'Back to the synchronous mode
                .AsyncMode = False
            End With
        End If
    Next i

    MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintDepositSlip(pCopyCount)", pCopyCount
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.MakePrintString(lRecLineChars,sBuf,sPrice)", Array(lRecLineChars, sBuf, _
'         sPrice)
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
            If oExchange.PaymentLines.Count > 0 Then
                X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            Else
                X4.Value(lngSalesItemCount, 9) = ""
            End If
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
                X4.Value(lngSalesItemCount, 7) = "Payment" & ": " & (oExchange.Customer.Name) & " " & (oExchange.Customer.Initials) & " " & oExchange.Note
            Else
                X4.Value(lngSalesItemCount, 7) = "Petty cash" & ":" & oExchange.Note
            End If
            X4.Value(lngSalesItemCount, 8) = ""
            If oExchange.PaymentLines.Count > 0 Then
                X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            Else
                X4.Value(lngSalesItemCount, 9) = ""
            End If
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
            If oExchange.PaymentLines.Count > 0 Then
                X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            Else
                X4.Value(lngSalesItemCount, 9) = ""
            End If
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.AddExchange"
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.G4_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.G4_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle)
End Sub

Private Sub LoadExchanges()
10        On Error GoTo errHandler
      Dim ZID As String
      Dim rs As ADODB.Recordset
      Dim cmd As ADODB.Command
      Dim prm As ADODB.Parameter
      Dim i As Integer

          
20        ZID = oPC.ZSession.Current_Z_Session_ID
          
30        oPC.OpenLocalDatabase

40        Set cmd = New ADODB.Command
50        cmd.CommandTimeout = 0
60        Set cmd.ActiveConnection = oPC.DBLocalConn
70        cmd.CommandText = "q_ExchangeDetails"
80        cmd.CommandType = adCmdStoredProc
          
90        Set prm = cmd.CreateParameter("@ZSESSID", adGUID, adParamInput, , ZID)
100       cmd.Parameters.Append prm
110       Set prm = Nothing
120       Set prm = cmd.CreateParameter("@TITLELENGTH", adInteger, adParamInput, , 50)
130       cmd.Parameters.Append prm
140       Set prm = Nothing
150       Set prm = cmd.CreateParameter("@CurrencyDivisor", adInteger, adParamInput, , 100)
160       cmd.Parameters.Append prm
170       Set prm = Nothing
         
180       lngSalesItemCount = 0
190       Set rs = cmd.Execute
200       Do While Not rs.EOF
210           lngSalesItemCount = lngSalesItemCount + 1
220           X4.InsertRows (lngSalesItemCount)
230               X4.Value(lngSalesItemCount, 1) = FNN(rs.Fields("EXCH_NUMBER"))
240               X4.Value(lngSalesItemCount, 2) = Format(rs.Fields("EXCH_SaleDate"), "HH:NN")
250               X4.Value(lngSalesItemCount, 3) = FNS(rs.Fields("SM_SHORTNAME"))
260               X4.Value(lngSalesItemCount, 4) = FNS(rs.Fields("EXCH_TYPE"))
270               X4.Value(lngSalesItemCount, 5) = FNS(rs.Fields("Code"))
280               X4.Value(lngSalesItemCount, 6) = FNN(rs.Fields("CSL_Qty"))
290               X4.Value(lngSalesItemCount, 7) = FNS(rs.Fields("TITLE")) & " (" & FNS(rs.Fields("Cust")) & ") " & IIf(FNN(rs.Fields("EXCH_Voids")) > 0, " (Voids:" & FNN(rs.Fields("EXCH_Voids")) & ")", "")
300               X4.Value(lngSalesItemCount, 8) = IIf(FNS(rs.Fields("EXCH_TYPE")) = "D", "", Format(rs.Fields("PRICE"), "Currency"))
310               X4.Value(lngSalesItemCount, 9) = Format(rs.Fields("DiscountedValueIncVAT"), "Currency")
320               X4.Value(lngSalesItemCount, 10) = FNS(rs.Fields("EXCH_ID"))
330               X4.Value(lngSalesItemCount, 11) = FNS(rs.Fields("P_ID"))
340               X4.Value(lngSalesItemCount, 12) = FNN(rs.Fields("EXCH_Voided"))
350               X4.Value(lngSalesItemCount, 13) = FNN(rs.Fields("EXCH_Voids"))
360           rs.MoveNext
370       Loop
380       X4.QuickSort 1, X4.UpperBound(1), 1, XORDER_DESCEND, XTYPE_NUMBER
390       G4.Array = X4
400       G4.ReBind
410       G4.Bookmark = 1
420       oPC.CloseLocalDatabase
          
430       Exit Sub
errHandler:
440       If ErrMustStop Then Debug.Assert False: Resume
450       ErrorIn "frmPOSMain.LoadExchanges"
End Sub

Private Sub PrintTotals(eDocumentType As enumDocumentType, pPrinter As Object, Optional bReprint As Boolean)
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
                            If oPmt.Amt < 0 Then   'Change vouchers should not show negative - confusing for customers
                                sExt = oPmt.AmtF_nonNegative
                            Else
                                sExt = oPmt.AmtF
                            End If
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintTotals(eDocumentType,pPrinter)", Array(eDocumentType, pPrinter)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintTotals(eDocumentType,pPrinter,bReprint)", Array(eDocumentType, pPrinter, _
         bReprint)
End Sub

Private Sub PrintHeader(eDocumentType As enumDocumentType, pPrinter As Object, Optional bReprint As Boolean, Optional bBeingVoided As Boolean, Optional strNote As String)
    On Error GoTo errHandler
Dim fDate As String
Dim ar() As String
Dim i As Integer
              Dim tmp As String
              Dim T As String
              Dim st As Integer

    If oPC.UseA4Printer Then
        Printer.Print ""
        Printer.Print ""
    End If
    
    Select Case eDocumentType
    Case eTypReceipt, eTypCreditNote, eTypeCancelledSale
        If oPC.UseA4Printer Then
                Printer.Print oPC.POSCompanyName
                If eDocumentType = eTypReceipt Then
                        If bBeingVoided Then
                            Printer.Print "THIS VOID t/a #" & oTmpExchange.ExchangeNumber
                            Printer.Print "VOIDING t/a #" & oExchange.ExchangeNumber
                        Else
                            Printer.Print "TAX INVOICE"
                        End If
                        Printer.Print ""
                Else
                        If bBeingVoided Then
                            Printer.Print "THIS VOID t/a #" & oTmpExchange.ExchangeNumber
                            Printer.Print "VOIDING t/a #" & oExchange.ExchangeNumber
                        Else
                            Printer.Print "TAX CREDIT NOTE"
                        End If
                        Printer.Print ""
                End If
                        Printer.Print oPC.POSBranchName
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                        Printer.Print ar(i)
                Next i
                Printer.Print ""
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                        Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
               ' .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If eDocumentType = eTypReceipt And Not bBeingVoided Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX INVOICE" + vbLf
                Else
                    If eDocumentType = eTypeCancelledSale Or bBeingVoided Then
                        .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CANCELLATION" + vbLf
                    Else
                        .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX CREDIT NOTE" + vbLf
                    End If
                End If
                If bBeingVoided Then
                    If eDocumentType = eTypeCancelledSale Then
                        .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "Cancelling t/a #" & oExchange.ExchangeNumber & vbLf
                    Else
                        .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "THIS VOID t/a #" & oTmpExchange.ExchangeNumber & vbLf
                        .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "VOIDING t/a #" & oExchange.ExchangeNumber & vbLf
                    End If
                End If
                
                
                
                ''''
                .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf   'ESC "|150uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf
              '  .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf
              '  .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                End If
                If oExchange.Note > "" And bBeingVoided Then
                    st = 1
                    tmp = IIf(strNote > "", strNote, oExchange.Note)
                    For i = 1 To 5
                        st = 1 + (.RecLineChars * (i - 1))
                        T = MID(tmp, st, .RecLineChars)
                        If T > "" Then .PrintNormal PTR_S_RECEIPT, T + vbLf
                    Next
                End If
                If bReprint = True Then
                      .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                      .PrintNormal PTR_S_RECEIPT, vbCrLf + vbCrLf
                     ' .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
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
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
               ' .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX INVOICE (from Appro)" + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                .PrintNormal PTR_S_RECEIPT, vbCrLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
    
                If bReprint = True Then
                      .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                      .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
                Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "APPRO OUT" + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
                ar = Split(oPC.POSBranchAddress, ",")
                For i = 0 To UBound(ar)
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
                Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
                .PrintNormal PTR_S_RECEIPT, vbCrLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
    
                If bReprint = True Then
                      .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                      .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "PAYMENT ACCEPTED" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, .RecLineChars) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CASH REFUND" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, .RecLineChars) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
            Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CHANGE VOUCHER" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CREDIT VOUCHER" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "DEPOSIT REFUND" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "DEPOSIT PAID" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars)
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars))
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT"
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
            Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "LOYALTY CLUB CREDIT VOUCHER" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "APPRO" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
                If oExchange.Customer.Name > "" Then
                    .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
                ElseIf oExchange.Note > "" Then
                    .PrintNormal PTR_S_RECEIPT, Left(oExchange.Note, (.RecLineChars)) + vbLf
                End If
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              Printer.Print "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "PETTY CASH" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CUSTOMER ORDER" + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
              ar = Split(oPC.POSBranchAddress, ",")
              For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
              Next i
                fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
              .PrintNormal PTR_S_RECEIPT, vbCrLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "  " & oExchange.StaffName & vbLf
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
              .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              If bReprint = True Then
                    .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                    .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
              End If
            End With
        End If
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PrintHeader(eDocumentType,pPrinter,bReprint)", Array(eDocumentType, pPrinter, _
'         bReprint)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintHeader(eDocumentType,pPrinter,bReprint)", Array(eDocumentType, pPrinter, _
         bReprint)
End Sub
Private Sub PrintFooter(pCopyNumber As Integer, eDocumentType As enumDocumentType, pPrinter As Object)
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
               ' .PrintNormal PTR_S_RECEIPT, ESC + "|700uF"
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf
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
                '.PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
                .PrintNormal PTR_S_RECEIPT, vbCrLf & vbCrLf & vbCrLf
            End With
        End If
    End Select



Dim strLine As String
Dim Src As z_TextFile
Dim strFilename As String
Dim fs As New FileSystemObject
Dim bBold As Boolean
Dim bUnderline As Boolean
Dim strControl As String

    bBold = False
    bUnderline = False
    If eDocumentType = etypCreditVoucher Then
        strFilename = "C:\PBKS\T&C_Voucher.txt"
    Else
        strFilename = "C:\PBKS\T&C.txt"
    End If
    If fs.FileExists(strFilename) Then
        Set Src = New z_TextFile
        Src.OpenTextFileToRead strFilename
        Do While Not Src.IsEOF
            strLine = Src.ReadLinefromTextFile
            If strLine = "" Then
                bBold = False
                bUnderline = False
            End If
            If Left(strLine, 3) = "/B1" Then
                bBold = True
                strLine = Right(strLine, Len(strLine) - 3)
            End If
            If Left(strLine, 3) = "/B0" Then
                bBold = False
                strLine = Right(strLine, Len(strLine) - 3)
            End If
            strControl = ""
            If bBold Then
                strControl = IIf(bBold, ESC + "E1", ESC + "E0")
            End If

            
            With pPrinter
                .AsyncMode = True
                .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
                
                .PrintNormal PTR_S_RECEIPT, strControl + strLine + vbLf     'ESC + "|N" + ESC + "M1" +
            End With
        Loop
        Src.CloseTextFile
    
    
    End If



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
    Case "V"
        ConvertToType = eTypeCancelledSale
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ConvertToType(val)", val
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.CollectPettyCashArray())"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CollectPettyCashArray()"
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
10        On Error GoTo errHandler
      Dim msg As String
      Dim sFileName As String
      Dim i As Integer
      Dim strExchangeMsg As String
      Dim strCSLPart As String
      Dim strPayPart As String
      Dim strExchangeType As String

20        If Not rsZSession Is Nothing Then
30            If rsZSession.State <> 0 Then
40                rsZSession.Close
50            End If
60        End If
      '================================================================================================================================
70        strExchangeType = "E"
80        For i = 1 To oExchange.PaymentLines.Count
90            If oExchange.PaymentLines(i).PaymentType = "AC" Then  'A purchase has been placed on account (even if partially)
100               strExchangeType = "A"
110           End If
120       Next i
130       If oExchange.transactionType = "PA" Then  'It is a payment made to an account
140           strExchangeType = "P"
150       End If
160       If oExchange.transactionType = "CN" Then  'It is a credit made to an account
170           strExchangeType = "CN"
180       End If
                '  & ReverseDateTime(oPC.ZSession.OpSession.DateEnded) & vbTab & oPC.ZSession.OpSession.OperatorID & vbTab & oPC.ZSession.OpSession.SupervisorID & vbTab _

    #If SALESREPTransmit = 1 Then
190           strExchangeMsg = strExchangeType & vbTab & pZID & vbTab & oPC.StationName & vbTab & ReverseDateTime(oPC.ZSession.FromDate) & vbTab & ReverseDateTime(oPC.ZSession.EndDate) & vbTab _
              & ReverseDateTime(oPC.ZSession.NominalDate) & vbTab & oPC.ZSession.OpSession.OPSID & vbTab & ReverseDateTime(oPC.ZSession.OpSession.DateStarted) & vbTab _
              & ReverseDateTime(oPC.ZSession.OpSession.DateEnded) & vbTab & oPC.ZSession.OpSession.OperatorID & vbTab & oExchange.SupervisorID & vbTab _
              & oExchange.ExchangeID & vbTab & ReverseDateTime(oExchange.ExchangeDate) & vbTab & oExchange.OperatorID & vbTab & oExchange.ExchangeNumber & vbTab _
              & oExchange.TotalPayable & vbTab & oExchange.TotalDiscount & vbTab & oExchange.TotalVAT & vbTab _
              & oExchange.ChangeGiven & vbTab & oExchange.LoyaltyValue & vbTab & oExchange.transactionType & vbTab _
              & oExchange.Note & vbTab & oExchange.ToVoid & vbTab & oExchange.Customer.CustomerID & vbTab & oExchange.SalesRepID
    #Else
200           strExchangeMsg = strExchangeType & vbTab & pZID & vbTab & oPC.StationName & vbTab & ReverseDateTime(oPC.ZSession.FromDate) & vbTab & ReverseDateTime(oPC.ZSession.EndDate) & vbTab _
              & ReverseDateTime(oPC.ZSession.NominalDate) & vbTab & oPC.ZSession.OpSession.OPSID & vbTab & ReverseDateTime(oPC.ZSession.OpSession.DateStarted) & vbTab _
              & ReverseDateTime(oPC.ZSession.OpSession.DateEnded) & vbTab & oPC.ZSession.OpSession.OperatorID & vbTab & oExchange.SupervisorID & vbTab _
              & oExchange.ExchangeID & vbTab & ReverseDateTime(oExchange.ExchangeDate) & vbTab & oExchange.OperatorID & vbTab & oExchange.ExchangeNumber & vbTab _
              & oExchange.TotalPayable & vbTab & oExchange.TotalDiscount & vbTab & oExchange.TotalVAT & vbTab _
              & oExchange.ChangeGiven & vbTab & oExchange.LoyaltyValue & vbTab & oExchange.transactionType & vbTab _
              & oExchange.Note & vbTab & oExchange.ToVoid & vbTab & oExchange.Customer.CustomerID
    #End If
          'Clean of any bar characters
210       strExchangeMsg = Replace(strExchangeMsg, "|", "/") & "|"
220       For i = 1 To oExchange.SaleLines.Count
230           strCSLPart = oExchange.SaleLines(i).PID & vbTab & oExchange.SaleLines(i).COLID & vbTab & oExchange.SaleLines(i).Qty & vbTab _
              & oExchange.SaleLines(i).Price & vbTab & oExchange.SaleLines(i).PriceAlteration & vbTab & oExchange.SaleLines(i).PDiscExt & vbTab _
              & oExchange.SaleLines(i).DiscountRate & vbTab & oExchange.SaleLines(i).VATRate & vbTab & oExchange.SaleLines(i).Counterfoil & vbTab _
              & oExchange.SaleLines(i).DiscountDescription & vbTab & oExchange.SaleLines(i).ActionSignatureID
240           strExchangeMsg = strExchangeMsg & Replace(strCSLPart, "|", "/") & IIf(i = oExchange.SaleLines.Count, "", "~")
250       Next i
260       strExchangeMsg = strExchangeMsg & "|"
          
270       For i = 1 To oExchange.PaymentLines.Count
280           strPayPart = oExchange.PaymentLines(i).PaymentType & vbTab & oExchange.PaymentLines(i).Amt & vbTab _
              & oExchange.PaymentLines(i).ReferenceComplete & vbTab & oExchange.PaymentLines(i).Note & vbTab & oExchange.PaymentLines(i).COLID
290           strExchangeMsg = strExchangeMsg & Replace(strPayPart, "|", "/") & IIf(i = oExchange.PaymentLines.Count, "", "~")
300       Next i
                          'Place  POS MESSAGE in queue for server to fetch
310       DispatchMessageEx strExchangeMsg
          
320       lblProg.Caption = lblProg.Caption & "X"
330       If pEXCHID > "" Then
340           SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
350           oPC.OpenLocalDatabase
360           oPC.DBLocalConn.Execute SQL
370           oPC.CloseLocalDatabase
380       End If
      'errHandler:
      '    If ErrMustStop Then Debug.Assert False: Resume
      '    ErrorIn "frmPOSMain.SendPOSExchange(pEXCHID,pZID)", Array(pEXCHID, pZID)
390       Exit Sub
errHandler:
400       If ErrMustStop Then Debug.Assert False: Resume
410       ErrorIn "frmPOSMain.SendPOSExchange(pEXCHID,pZID)", Array(pEXCHID, pZID)
End Sub

Private Sub ReSendExchanges()
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strEXCHNUM As String
Dim sr As New Scripting.FileSystemObject
Dim strFilename As String
Dim iLoc As Integer
Dim strStart As String
Dim strEnd As String
Dim istart As Long
Dim iEnd As Long
Dim bErr As Boolean
Dim i As Long

    strFilename = oPC.LocalRootFolder & "\RESEND.TXT"
    If sr.FileExists(strFilename) Then
        oTF.OpenTextFileToRead strFilename
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
                    istart = CLng(strStart)
                Else
                    bErr = True
                End If
                If IsNumeric(strEnd) Then
                    iEnd = CLng(strEnd)
                Else
                    bErr = True
                End If
                If Not bErr Then
                    For i = istart To iEnd
                        SendPOSExchange_ByExchangeNumber CStr(i)
                    Next i
                End If
            End If
        End If
    Loop
        
    oTF.CloseTextFile
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ReSendExchanges"
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
#If SALESREPTransmit Then
    sSQL = "SHAPE {SELECT tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession " _
        & " APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession " _
        & " APPEND (( SHAPE {SELECT EXCH_TYPE as TYP,EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE, " _
        & " EXCH_SALEVALUE,EXCH_DISCOUNTVALUE,EXCH_VATVALUE,EXCH_CHANGEGIVEN,EXCH_LOYALTYVALUE,EXCH_TYPE,EXCH_OPERATORID,EXCH_SUPERVISORID,EXCH_NUMBER,EXCH_VOIDS,EXCH_SalesRepID,EXCH_NOTE " _
        & " FROM tEXCHANGE WHERE EXCH_ID = '" & pEXCHID & "'}  AS POSExchange " _
        & " APPEND ({SELECT * FROM tCSL}  AS rsSALESLINES " _
        & " RELATE EXCH_ID TO CSL_EXCH_ID) AS SALESLINES,({SELECT * FROM tPayment}  AS rsPAYMENTS " _
        & " RELATE EXCH_ID TO PAY_EXCH_ID) AS PAYMENTS) AS POSExchange " _
        & " RELATE OPS_ID TO EXCH_OPSESSIONID) AS POSExchange) AS OPSession RELATE Z_ID TO OPS_Z_ID) AS OPSession"
#Else
    sSQL = "SHAPE {SELECT tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession " _
        & " APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession " _
        & " APPEND (( SHAPE {SELECT EXCH_TYPE as TYP,EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE, " _
        & " EXCH_SALEVALUE,EXCH_DISCOUNTVALUE,EXCH_VATVALUE,EXCH_CHANGEGIVEN,EXCH_LOYALTYVALUE,EXCH_TYPE,EXCH_OPERATORID,EXCH_SUPERVISORID,EXCH_NUMBER,EXCH_VOIDS,EXCH_NOTE " _
        & " FROM tEXCHANGE WHERE EXCH_ID = '" & pEXCHID & "'}  AS POSExchange " _
        & " APPEND ({SELECT * FROM tCSL}  AS rsSALESLINES " _
        & " RELATE EXCH_ID TO CSL_EXCH_ID) AS SALESLINES,({SELECT * FROM tPayment}  AS rsPAYMENTS " _
        & " RELATE EXCH_ID TO PAY_EXCH_ID) AS PAYMENTS) AS POSExchange " _
        & " RELATE OPS_ID TO EXCH_OPSESSIONID) AS POSExchange) AS OPSession RELATE Z_ID TO OPS_Z_ID) AS OPSession"

#End If
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
#If SALESREPTransmit Then
    strExchangeMsg = strTyp & vbTab & rsZSession!Z_ID & vbTab & rsZSession!Z_TILLPOINT & vbTab & ReverseDateTime(rsZSession!Z_STARTDATE) & vbTab _
    & ReverseDateTime(FND(rsZSession!Z_ENDDATE)) & vbTab _
    & ReverseDateTime(rsZSession!Z_NOMINALDATE) & vbTab & rsZSession.Fields("OPSession")!OPS_ID & vbTab & ReverseDateTime(CDate(rsZSession.Fields("OPSession")!OPS_STARTTIME)) & vbTab _
    & ReverseDateTime(rsOP!OPS_endtime) & vbTab & rsOP!OPS_OPERATORID & vbTab & rsOP!OPS_OPERATORID & vbTab _
    & rsEx!EXCH_ID & vbTab & ReverseDateTime(rsEx!EXCH_SaleDate) & vbTab _
    & rsEx!EXCH_OperatorID & vbTab & rsEx!EXCH_Number & vbTab _
    & rsEx!EXCH_SaleValue & vbTab & rsEx!EXCH_DiscountValue & vbTab & rsEx!EXCH_VATValue & vbTab _
    & rsEx!EXCH_ChangeGiven & vbTab & rsEx!EXCH_LoyaltyValue & vbTab & rsEx!EXCH_TYPE & vbTab _
    & rsEx!EXCH_Note & vbTab & rsEx!EXCH_VOIDS & vbTab & rsEx!EXCH_TP_ID & vbTab & rsEx!EXCH_SalesRepID & "|"
#Else
    strExchangeMsg = strTyp & vbTab & rsZSession!Z_ID & vbTab & rsZSession!Z_TILLPOINT & vbTab & ReverseDateTime(rsZSession!Z_STARTDATE) & vbTab _
    & ReverseDateTime(FND(rsZSession!Z_ENDDATE)) & vbTab _
    & ReverseDateTime(rsZSession!Z_NOMINALDATE) & vbTab & rsZSession.Fields("OPSession")!OPS_ID & vbTab & ReverseDateTime(CDate(rsZSession.Fields("OPSession")!OPS_STARTTIME)) & vbTab _
    & ReverseDateTime(rsOP!OPS_endtime) & vbTab & rsOP!OPS_OPERATORID & vbTab & rsOP!OPS_OPERATORID & vbTab _
    & rsEx!EXCH_ID & vbTab & ReverseDateTime(rsEx!EXCH_SaleDate) & vbTab _
    & rsEx!EXCH_OperatorID & vbTab & rsEx!EXCH_Number & vbTab _
    & rsEx!EXCH_SaleValue & vbTab & rsEx!EXCH_DiscountValue & vbTab & rsEx!EXCH_VATValue & vbTab _
    & rsEx!EXCH_ChangeGiven & vbTab & rsEx!EXCH_LoyaltyValue & vbTab & rsEx!EXCH_TYPE & vbTab _
    & rsEx!EXCH_Note & vbTab & rsEx!EXCH_VOIDS & vbTab & rsEx!EXCH_TP_ID & vbTab & "|"
#End If
    Do While rsCSL.EOF = False
        strCSLPart = rsCSL!CSL_P_ID & vbTab & rsCSL!CSL_COLID & vbTab & rsCSL!CSL_Qty & vbTab _
        & rsCSL!CSL_Price & vbTab & rsCSL!CSL_PriceAlteration & vbTab & rsCSL!CSL_Discount & vbTab _
        & rsCSL!CSL_DiscountRate & vbTab & rsCSL!CSL_VATRATE & vbTab & rsCSL!CSL_Counterfoil & vbTab _
        & rsCSL!CSL_DiscountDescription & vbTab & rsCSL!CSL_ActionSignature
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SendPOSExchange(pEXCHID,pOPSID,pZID)", Array(pEXCHID, pOPSID, pZID)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SendPOSExchange_ByExchangeNumber(pEXCHNumber)", pEXCHNumber
End Sub

Private Function DisconnectAll(rs As ADODB.Recordset)
    On Error GoTo errHandler
    Dim i As Long
With rs
   Set .ActiveConnection = Nothing
   For i = 0 To rs.Fields.Count - 1
      If (rs.Fields(i).Type = adChapter) Then
         DisconnectAll rs.Fields(i).Value
      End If
   Next i
End With
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisconnectAll(rs)", rs
End Function
Private Sub DispatchMessageEx(pMsgXML As String)
10        On Error GoTo errHandler
      Dim oStream As ADODB.Stream
      'Dim oTR As MSMQTransaction

20        Set QI = New MSMQQueueInfo
30      '''  QI.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\qpos"
'''''''changed 4/12/2017 to accomodate IPaddressing to server (Change due to Metacom
    If IsNumeric(Replace(oPC.ServerIPAddress, ".", "")) Then
        QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
    '    MsgBox "QI.FormatName=" & QI.FormatName
    Else
        QI.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
    '    MsgBox "QI.FormatName=" & QI.FormatName
    End If
''''''''

40        Set QPOS = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
          
50        QI.FormatName = "DIRECT=OS:" & oPC.NameOfPC & "\Private$\qposack"
          
60        Set POSmsg = New MSMQMessage
70        POSmsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
80        POSmsg.Journal = MQMSG_DEADLETTER
90        POSmsg.MaxTimeToReachQueue = oPC.POSMessageTimeout  '9 days
          
100       POSmsg.Label = "POSN," & Format(Now, "dd/mm/yyyy HH:NN")

110       POSmsg.Body = pMsgXML

      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '''''''''''''This must be removed when no longer necessary
       '   POSmsg.Journal = MQMSG_JOURNAL
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
120       Set POSmsg.ResponseQueueInfo = QI
130       POSmsg.Send QPOS  ', MQ_SINGLE_MESSAGE ', oTR

      'errHandler:
      '    ErrorIn "frmPOSMain.DispatchMessageEx(pMsgXML)", pMsgXML
140       Exit Sub
errHandler:
150       If ErrMustStop Then Debug.Assert False: Resume
160       ErrorIn "frmPOSMain.DispatchMessageEx(pMsgXML)"
            HandleError
End Sub

Private Sub DispatchMessage(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim oStream As ADODB.Stream
'Dim oTR As MSMQTransaction

    Set QI = New MSMQQueueInfo
    If IsNumeric(Replace(oPC.ServerIPAddress, ".", "")) Then
        QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
       ' MsgBox "QI.FormatName=" & QI.FormatName
    Else
        QI.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
      '  MsgBox "QI.FormatName=" & QI.FormatName
    End If
    'QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
   ' QI.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
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

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.DispatchMessage(rs)", rs
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.PreparePaymentLine(ePaymentMode)"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PreparePaymentLine(ePaymentMode)", ePaymentMode
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ConvertPaymentStateToCode(ePaymentMode)"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ConvertPaymentStateToCode(ePaymentMode)", ePaymentMode
End Function

'=====================================================================================================




Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHandler
    
    
    If oPC.UseCashDrawer Then
        If enPrinterType = en_Epson Then
            If OPOSCashDrawer.DrawerOpened = True Then
                MsgBox "Please close drawer before continuing."
                Exit Sub
            End If
        Else
            If bDrawerFlag = True Then
                MsgBox "Please close drawer before continuing."
                Exit Sub
            End If
        End If
    End If
    bShiftDown = (Shift = 1)
    If KeyCode = 13 Then
        itest = itest + 1
        
        
'Come back to here is barcode scanned instead of price confirmed - auto scanning
ReEvaluate:
        lblCHange.Visible = False
        enNewState = GetNewState(txtInput)  ', itmp, strArg, strArg2
        If bForceClose Then Exit Sub 'and error has occurred
        If bBarcodeNotPrice Then
            enPresentState = enNewState
            bBarcodeNotPrice = False
            GoTo ReEvaluate
        End If
        If bValid Then SetPresentState enNewState
        If enNewState = eStart Then txtInput.BackColor = RGB(230, 250, 210)

     End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift)
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
            GetNewState = Action_Appro
        End Select
    Case eCollectRep
        If UCase(strPrefix) = "X" Then
            GetNewState = eConfirmation
        Else
            GetNewState = Action_GetSalesRep
        End If
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
                Case ".."
                    GetNewState = eSale
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
    Case ePaymentType_CreditCardRef
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
''''        If strPrefix = ".." Then
''''            RemoveSaleLine , True
''''            If bSaleActive Then
''''                If enMode = emode_Sale Then
''''                    GetNewState = eSale
''''                ElseIf enMode = eMode_ApproReturn Then
''''                    GetNewState = eApproReturn
''''                ElseIf enMode = emode_Appro Then
''''                    GetNewState = eAppro
''''                End If
''''            Else
''''                GetNewState = eStart
''''            End If
''''        Else
            If Not IsNumeric(strRaw) Then
                GetNewState = ePrice
            Else
                GetNewState = Action_Price()
            End If
''''        End If
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
        itest = 0

        Select Case strPrefix
        Case "END"
            GetNewState = Action_eCloseApp
        Case "REPRINT"
            GetNewState = Action_Reprint
        Case "QU"
            GetNewState = Action_LoadFromQuotation
        Case "XEND"
            If oPC.EnableXEND Then
                GetNewState = Action_eXTerminate
            Else
                GetNewState = eStart
            End If

        Case "ZEND"
            GetNewState = Action_eZTerminate
        Case "A"
            If CheckThisPoint(M_ISSUEAPPRO) Then
                If Not SecurityControl(enSECURITY_ISSUEPOSAPPRO, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission to issue an appro.") Then
                    GetNewState = eStart
                Else
                    enMode = emode_Appro
                    GetNewState = Action_SearchCustomer("A")
                End If
            Else
                enMode = emode_Appro
                GetNewState = Action_SearchCustomer("A")
            End If
        Case "RDEP"
            If CheckThisPoint(M_REFUNDDEPOSIT) Then
                If SecurityControl(enSECURITY_REFUNDDEPOSITONPOS, lngOPID, strName, , "Enter your signature.", "You do not have permission to refund deposits.") Then
                    enMode = eMode_ReturnDeposit
                    GetNewState = Action_SearchCustomer("R")
                Else
                    GetNewState = eStart
                End If
            Else
                enMode = eMode_ReturnDeposit
                GetNewState = Action_SearchCustomer("R")
            End If
        Case "AR"
            enMode = eMode_ApproReturn
            GetNewState = Action_SearchCustomer("I")
        Case "PA"
            If CheckThisPoint(M_ACCEPTACPAYMENT) Then
                If SecurityControl(enSECURITY_ACCEPTACPAYMENT, lngOPID, strName, , "Enter your signature.", "You do not have permission to accept payments.") Then
                        enMode = emode_PayAccount
                        GetNewState = Action_SearchCustomer("PA")
                Else
                    GetNewState = eStart
                End If
            Else
                enMode = emode_PayAccount
                GetNewState = Action_SearchCustomer("PA")
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
            If CheckThisPoint(M_ISSUEPOSCREDITNOTE) Then
                If SecurityControl(enSECURITY_ISSUEPOSCREDITNOTE, lngOPID, strName, , "Enter your signature.", "You do not have permission to issue credit notes.") Then
                    enMode = emode_CreditNote
                    GetNewState = Action_SearchCustomer("CN")
                Else
                    GetNewState = eStart
                End If
            Else
                    enMode = emode_CreditNote
                    GetNewState = Action_SearchCustomer("CN")
            End If
'        Case "PS"
'            If bSaleOnHold Then
'                MsgBox "There is already a sale saved, you should either cancel this sale or complete it before using RS to retrieve the saved sale and cancel or complete it before continuing", vbInformation + vbOKOnly, "Can't do this"
'                bValid = False
'            Else
'                GetNewState = Action_StoreSale
'            End If
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
'        Case "PS"
'            If bSaleOnHold Then
'                MsgBox "There is already a sale saved, you should either cancel this sale or complete it before using RS to retrieve the saved sale and cancel or complete it before continuing", vbInformation + vbOKOnly, "Can't do this"
'                bValid = False
'            Else
'                GetNewState = Action_StoreSale
'            End If
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
                        If CheckThisPoint(M_ISSUEPOSREFUND) Then
                            If Not SecurityControl(enSECURITY_ISSUEPOSREFUND, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission to issue a cash refund.") Then
                                GetNewState = eSale
                            Else
                                RefundPayment ePaymentMode_Cash
                                GetNewState = eConfirmation
                            End If
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
               '     If Not SecurityControl(enSECURITY_ACCEPTDIRECTDEPOSIT, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission to pay with a direct deposit.") Then
               '         GetNewState = eSale
               '     Else
                        GetNewState = ePaymentType_DirectDeposit
               '     End If
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
            
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.GetNewState(txtIn)", txtIn
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetNewState(txtIn)", txtIn
    HandleError
End Function
Private Function DetermineReturnToState() As eState
    On Error GoTo errHandler
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
        If oExchange.transactionType <> "OR" Then
            If enPresentState = eConfirmation Or enPresentState = ePaymentType_CreditVoucherRef Or enPresentState = ePaymentType_voucherRef Or enPresentState = ePaymentType_CreditCardRef Then
                RemovePaymentLine , True
            End If
            DetermineReturnToState = eSale
        Else
                RemovePaymentLine , True
            DetermineReturnToState = eCollect
        End If
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DetermineReturnToState"
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
            'change here 23/10/2024
        If strPrefix = "" Then strPrefix = strSuffix

    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SeparateInput(pRaw,pPrefix,pSuffix)", Array(pRaw, pPrefix, pSuffix)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SeparateInput(pRaw,pPrefix,pSuffix)", Array(pRaw, pPrefix, pSuffix)
End Function
Private Function Action_eXTerminate() As eState
    On Error GoTo errHandler


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
        oPC.ZSession.OpSession.Close_OP_Session True
        OpenDrawer
        oPC.CloseLocalDatabase
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_eXTerminate"
End Function
Private Function Action_eZTerminate() As eState
    On Error GoTo errHandler
    If MsgBox("Close day ?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Action_eZTerminate = eStart
    Else
        If oExchange.SaleLines.Count > 0 Then
            oExchange.CancelEdit
        End If
        bCloseZsession = True
        Action_eZTerminate = eEND
        oPC.ZSession.Close_Z_Session
        OpenDrawer
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_eZTerminate"
End Function
Private Function Action_eCloseApp() As eState
    On Error GoTo errHandler
'    If oExchange.SaleLines.Count > 0 Then
'        oExchange.CancelEdit
'    End If
    Action_eCloseApp = eEND
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_eCloseApp"
End Function

Private Sub SetTip(pMsg As String)
    On Error GoTo errHandler
    lblInput.Caption = pMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SetTip(pMsg)", pMsg
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.setInputBox(pText,pPasswordChar,pChange,bAutoSelect)", Array(pText, _
'         pPasswordChar, pChange, bAutoSelect)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.setInputBox(pText,pPasswordChar,pChange,bAutoSelect)", Array(pText, _
         pPasswordChar, pChange, bAutoSelect)
End Sub
Private Sub OpenDrawer()
    On Error GoTo errHandler
Dim ret As Long
Dim sControl As String
Dim ar() As String
Dim i As Integer

    If oPC.UseCashDrawer = False Then
        Exit Sub
    End If
    
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
            If enPrinterType = en_Epson Then
                OPOSCashDrawer.OpenDrawer
                If bCanSenseDrawer = True Then
                    If OPOSCashDrawer.CapStatus = True Then
                        txtCloseDrawerMessage.Visible = True
                        txtCloseDrawerMessage.ZOrder 0
                        Me.Refresh
                    End If
                End If
            Else
'''''''''                OPOSCashDrawerDigipos.OpenDrawer
'''''''''                If bCanSenseDrawer = True Then
'''''''''                    If OPOSCashDrawerDigipos.CapStatus = True Then
'''''''''                        txtCloseDrawerMessage.Visible = True
'''''''''                        txtCloseDrawerMessage.ZOrder 0
'''''''''                        Me.Refresh
'''''''''                    End If
'''''''''                End If
            End If
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.OpenDrawer"
End Sub
Private Sub OPOSCashDrawerDigipos_StatusUpdateEvent(ByVal Data As Long)
    If bCanSenseDrawer = False Then Exit Sub
    If Data = False Then
        bDrawerFlag = False
        txtCloseDrawerMessage.Visible = False
    End If
End Sub
Private Sub OPOSCashDrawer_StatusUpdateEvent(ByVal Data As Long)
    If bCanSenseDrawer = False Then Exit Sub
    If Data = False Then
        bDrawerFlag = False
        txtCloseDrawerMessage.Visible = False
    End If

End Sub
'Private Sub OPOSCashDrawerDigipos_StatusUpdateEvent(ByVal Data As Long)
''    If bCanSenseDrawer = False Then Exit Sub
''    If bIgnorestatus Then Exit Sub   '-THis is to handle a strange problem Digipos printer statuses are coming thru here
''    If OPOSCashDrawerDigipos.DrawerOpened = False Then
''        bDrawerFlag = False
''        txtCloseDrawerMessage.Visible = False
''    Else
''        bDrawerFlag = True
''        txtCloseDrawerMessage.Visible = True
''    End If
''
'End Sub

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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.CheckAllStatus(pRefund)", pRefund
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Valid_RedeemDeposit_Arg(pArg)", pArg, EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Valid_RedeemDeposit_Arg(pArg)", pArg
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ValidRowNumberSet(pString)", pString
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ValidRowNumberSet(pString)", pString
End Function
Private Function Validate_eSelectDepositLine(pIn As String, pRefund As Boolean) As Boolean
    On Error GoTo errHandler
Dim strMsg As String

    Validate_eSelectDepositLine = True
    If Not ValidRowNumberSet(pIn) Then
        MsgBox "Invalid row selection.", vbOKOnly, "Can't do this"
        Validate_eSelectDepositLine = False
        Exit Function
    End If
    If Not CheckAllStatus(pRefund) Then
        If pRefund Then
            strMsg = "Invalid row selection,  check status is 'P' and deposit is greater than 0.00 " & vbCrLf _
                & "'P' means deposit has been paid." & vbCrLf _
                & "'E' means that the deposit has been redeemed already." & vbCrLf _
                & "'X' means that the deposit has been refunded."
                MsgBox strMsg, vbOKOnly, "Can't do this"
        End If
        Validate_eSelectDepositLine = False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Validate_eSelectDepositLine(pIn,pRefund)", Array(pIn, pRefund), EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Validate_eSelectDepositLine(pIn,pRefund)", Array(pIn, pRefund)
End Function

Private Function Action_SelectApproLine() As eState
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SelectApproLine"
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
        LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
    Next
    Action_SelectDepositLine = eCollect
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_SelectDepositLine", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SelectDepositLine"
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
        LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
    Next
    SetForCOLSVisible False
    Action_SelectDepositLineForRefund = eRefundDeposit 'Action_Refund("R")

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SelectDepositLineForRefund"
End Function

Private Function Action_SearchCustomer(pType As String) As eState
    On Error GoTo errHandler
Dim lngInvValue As Long

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
            If cCOLS.Count > 0 And LoadCOLS = True Then
                G3.Caption = DisplayCustomerDetails
                Action_SearchCustomer = ePaymentType_RedeemDeposit
            Else
                MsgBox "There are no unredeemed deposits for this customer", vbInformation, "Can't find active orders"
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
        'RefreshAllSaleRows
        txtInput = ""
     Else
         Action_SearchCustomer = enPresentState
     End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_SearchCustomerfordeposit(pType)", pType, EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_SearchCustomer(pType)", pType
End Function
Private Function GetPaymentReference() As Boolean
    On Error GoTo errHandler
Dim f As New frmPaymentReference
    f.Show vbModal
    oExchange.Note = f.PaymentReference & ": " & oExchange.Note
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetPaymentReference"
End Function
Private Function GetInvoicesPerCustomer() As Boolean
    On Error GoTo errHandler
    oPC.dbConnectMain
    GetInvoicesPerCustomer = CreditAccount(lngCustomerID, 100)
    oPC.dbMainDisConnect
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetInvoicesPerCustomer"
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
        LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, False, False
        DoEvents
    Next i
    End If
    oExchange.CalculateTotals
    
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.CreateApproReturnAndInvoice(pTPID,pInvValue)", Array(pTPID, pInvValue)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CreateApproReturnAndInvoice(pTPID,pInvValue)", Array(pTPID, pInvValue)
End Function
Private Function ValidRowset(strIn As String) As Boolean
    On Error GoTo errHandler

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.ValidRowset(strIn)", strIn
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_eStart", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_eStart"
End Sub
Private Function CreditAccount(pTPID As Long, pInvValue As Long) As Boolean
    On Error GoTo errHandler
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
            LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
        End If
    Next i
    End If
    
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CreditAccount(pTPID,pInvValue)", Array(pTPID, pInvValue)
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.RefundPayment", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefundPayment(ePaymentMode)", ePaymentMode
End Function

Private Function Action_PaymentType_Cash() As eState
    On Error GoTo errHandler
    Dim lngTmp As Long
    Dim bOK As Boolean
    If strPrefix = ".." Then
        If oExchange.transactionType = "OR" Then
            Action_PaymentType_Cash = eCollect
        Else
            Action_PaymentType_Cash = DetermineReturnToState
        End If
    Else
'        bOK = SetField_strAsCurrencyToLong(lngTmp, Trim(strRaw), 99, "Amt", False, oPC.CurrencyDivisor)
'        If bOK = False Then
'        Else
'            If lngTmp = 0 Then
'            End If
'        End If
        PreparePaymentLine ePaymentMode_Cash
        If oPAYMENTLine.SetAmt(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                If bCollectRepcode Then
                    Action_PaymentType_Cash = eCollectRep
                Else
                    Action_PaymentType_Cash = eConfirmation
                End If
            Else
                Action_PaymentType_Cash = eSale
            End If
        Else
            SetTip "Invalid payment amount."
            RemovePaymentLine , True
            Action_PaymentType_Cash = eSale
        End If
        DisplayPayment
    End If
    If strMsg = "NOTOK" Then
        If MsgBox("The payment made seems excessive: " & oPAYMENTLine.AmtF & vbCrLf & "Click 'Cancel' to re-enter", vbInformation + vbOKCancel, "Warning") = vbCancel Then
            RemovePaymentLine , True
            Action_PaymentType_Cash = eSale
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_Cash", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Cash"
End Function
Private Function Action_PaymentType_Cheque() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_Cheque = eCollect
        Else
            Action_PaymentType_Cheque = DetermineReturnToState
        End If
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_Cheque", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Cheque"
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
                If bCollectRepcode Then
                    Action_PaymentType_ChequeRef = eCollectRep
                Else
                    Action_PaymentType_ChequeRef = eConfirmation
                End If
            Else
                Action_PaymentType_ChequeRef = eSale
            End If
            DisplayPayment
        Else
            SetTip "Invalid Reference."
            Action_PaymentType_ChequeRef = ePaymentType_ChequeRef
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_ChequeRef", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_ChequeRef"
End Function

Private Function Action_PaymentType_Creditcard() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_Creditcard = eCollect
        Else
            Action_PaymentType_Creditcard = DetermineReturnToState
        End If
    Else
        PreparePaymentLine ePaymentMode_CreditCard
        If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
            oExchange.CalculateTotals
            Action_PaymentType_Creditcard = ePaymentType_CreditCardRef
        Else
            RemovePaymentLine , True
            SetTip "Invalid payment amount."
            Action_PaymentType_Creditcard = ePaymentType_CreditCard
        End If
        DisplayPayment
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Creditcard"
End Function
Private Function Action_PaymentType_CreditcardRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditcardRef = DetermineReturnToState
    Else
        If strPrefix = strSuffix And IsNumeric(strSuffix) Then
            If oPAYMENTLine.SetReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete(, strMsg) Then
                    If bCollectRepcode Then
                        Action_PaymentType_CreditcardRef = eCollectRep
                    Else
                        Action_PaymentType_CreditcardRef = eConfirmation
                    End If
                Else
                    Action_PaymentType_CreditcardRef = eSale
                End If
                DisplayPayment
            Else
                MsgBox ("Invalid Reference.")
                'SetTip "Invalid Reference."
                Action_PaymentType_CreditcardRef = ePaymentType_CreditCardRef
            End If
        Else
            MsgBox ("Reference can only contain numbers")
           ' SetTip "Reference can only contain numbers"
             Action_PaymentType_CreditcardRef = ePaymentType_CreditCardRef
       End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditcardRef"
End Function

Private Function Action_PaymentType_Voucher() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_Voucher = eCollect
        Else
            Action_PaymentType_Voucher = DetermineReturnToState
        End If
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_Voucher", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Voucher"
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
                    If bCollectRepcode Then
                        Action_PaymentType_VoucherRef = eCollectRep
                    Else
                        Action_PaymentType_VoucherRef = eConfirmation
                    End If
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_VoucherRef", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_VoucherRef"
End Function
Private Function Action_PaymentType_CreditVoucher() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_CreditVoucher = eCollect
        Else
            Action_PaymentType_CreditVoucher = DetermineReturnToState
        End If
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_CreditVoucher", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditVoucher"
End Function
Private Function Action_PaymentType_CreditVoucherRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditVoucherRef = DetermineReturnToState
    Else
        If oPAYMENTLine.SetReference(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                If bCollectRepcode Then
                    Action_PaymentType_CreditVoucherRef = eCollectRep
                Else
                    Action_PaymentType_CreditVoucherRef = eConfirmation
                End If
            Else
                Action_PaymentType_CreditVoucherRef = eSale
            End If
            DisplayPayment
        Else
            SetTip "Invalid Reference."
            Action_PaymentType_CreditVoucherRef = ePaymentType_CreditVoucherRef
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_CreditVoucherRef", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditVoucherRef"
End Function
Private Function Action_PaymentType_Account() As eState
    On Error GoTo errHandler
Dim lngDeposit As Long
Dim iRow As Long
Dim lngTmp As Long

    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_Account = eCollect
        Else
            Action_PaymentType_Account = DetermineReturnToState
        End If
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
                If bCollectRepcode Then
                    Action_PaymentType_Account = eCollectRep
                Else
                    Action_PaymentType_Account = eConfirmation
                End If
            Else
                Action_PaymentType_Account = eSale
            End If
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_Account = ePaymentType_Account
        End If
        DisplayPayment
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_Account", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_Account"
End Function
Private Function Action_PaymentType_DirectDeposit() As eState
    On Error GoTo errHandler
Dim lngDeposit As Long
Dim iRow As Long
Dim lngTmp As Long


    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_DirectDeposit = eCollect
        Else
            Action_PaymentType_DirectDeposit = DetermineReturnToState
        End If
    Else
        If Not IsNumeric(strSuffix) Then
            Action_PaymentType_DirectDeposit = eSale
            Exit Function
        End If
        If CheckThisPoint(M_ACCEPTDIRECTDEPOSIT) Then
            If Not SecurityControl(enSECURITY_ACCEPTDIRECTDEPOSIT, lngOPID, strName, , "Enter your signature.", "Your signature does not give you permission for theis operation") Then
                Action_PaymentType_DirectDeposit = eSale
                Exit Function
            End If
        Else
        End If
        PreparePaymentLine ePaymentMode_DIrectDeposit
        If oPAYMENTLine.SetAmt(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                If bCollectRepcode Then
                    Action_PaymentType_DirectDeposit = eCollectRep
                Else
                    Action_PaymentType_DirectDeposit = eConfirmation
                End If
            Else
                Action_PaymentType_DirectDeposit = eSale
            End If
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_DirectDeposit = ePaymentMode_DIrectDeposit
        End If
        DisplayPayment
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_DirectDeposit", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_DirectDeposit"
End Function
Private Function Action_PaymentType_RedeemDeposit() As eState
    On Error GoTo errHandler
Dim lngDeposit As Long
Dim iRow As Long
Dim lngTmp As Long

    If strPrefix = ".." Then
        If oExchange.transactionType <> "OR" Then
            Action_PaymentType_RedeemDeposit = eCollect
        Else
            Action_PaymentType_RedeemDeposit = DetermineReturnToState
        End If
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
        If lngDeposit < 0 Or FNS(X3(lngTmp, 7)) <> "P" Then   'Or lngDeposit > 100000
            MsgBox "Deposit is invalid amount or not paid or already redeemed", , "Can't accept"
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
            If bCollectRepcode Then
                Action_PaymentType_RedeemDeposit = eCollectRep
            Else
                Action_PaymentType_RedeemDeposit = eConfirmation
            End If
        Else
            Action_PaymentType_RedeemDeposit = eSale
        End If
        DisplayPayment
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PaymentType_RedeemDeposit", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_RedeemDeposit"
End Function

Private Function Action_Refund(pType As String) As eState
    On Error GoTo errHandler
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

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Refund(pType)", pType
End Function
Private Function Action_Confirmation()
    On Error GoTo errHandler
Dim frm As frmChangeToGive
Dim strMsg As String
Dim oPmt As a_Payment
Const errMsg = "You do not have the authority to accept payment. Talk to your supervisor."
    
    If strPrefix = ".." Then
        Action_Confirmation = DetermineReturnToState
    Else
        If IsRole(enSECURITY_ISOPERATOR, strPrefix & strSuffix, strName, lngOPID) = True Then
            oExchange.OperatorID = lngOPID
            
            
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
            
            If oExchange.transactionType = "S" Or oExchange.transactionType = "AR" Or oExchange.transactionType = "OR" Or oExchange.transactionType = "" Then
                lblCHange.Text = ""
                iChangeGivenLines = 0
                If oExchange.ChangeGiven > 0 Then
                    lblCHange.ForeColor = COLOUR_CHANGE
                    lblCHange.Text = lblCHange.Text & IIf(Len(lblCHange.Text) > 0, vbCrLf, "") & "CHANGE: " & oExchange.ChangeGivenF
                    iChangeGivenLines = iChangeGivenLines + 1
                Else
                    For Each oPmt In oExchange.PaymentLines
                        If oPmt.PaymentTypeF = "Credit card" Then
                            lblCHange.Text = lblCHange.Text & IIf(Len(lblCHange.Text) > 0, vbCrLf, "") & "CARD: " & oPmt.AmtF
                            lblCHange.ForeColor = COLOUR_CREDITCARD
                            lblCHange.BackColor = COLOR_PALEYELLOW
                           
                            iChangeGivenLines = iChangeGivenLines + 1
                        End If
                    Next
                End If
            End If
            If lblCHange.Text > "" Then lblCHange.Visible = True
            lblCHange.ZOrder 0
            AdjustHeight lblCHange, iChangeGivenLines
           ' TextTrans lblChange
            Res = AcceptSale(False)
            If Res = False Then
                MsgBox "Problem with calculation, possibly values are too large.", vbInformation + vbOKOnly, "Can't continue"
                Action_Confirmation = enPresentState
            End If
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

'errHandler:
   ' If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_Confirmation()"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Confirmation"
End Function
Private Function Action_operatorsReport()
    On Error GoTo errHandler
    If SecurityControl(eSupervisor, lngSupervisorID, strName, , "Enter security code to view operators' report") Then
        Set frmOpRep = New frmPOSOPREP
        setInputBox "", "", "", True
        frmOpRep.Show vbModal
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_operatorsReport"
End Function
Private Function Action_PettyCash() As eState
    On Error GoTo errHandler
        If CheckThisPoint(M_PETTYCASH) Then
            If Not SecurityControl(enSECURITY_ISSUEPETTYCASH, lngOPID, strName, , "Enter your signature.", "You do not have permission to issue petty cash.", True) Then
                Exit Function
            End If
        End If
    
    
    Set frmPC = New frmPettyCash
    frmPC.Show vbModal
    If frmPC.Cancelled Then
        'clear fields
        Unload frmPC
        setInputBox "", "", "", True
        Exit Function
    Else
        oExchange.SetExchangeType ePettyCashType
        oExchange.Note = frmPC.Reason
        Set oPAYMENTLine = oExchange.PaymentLines.Add
        oPAYMENTLine.ApplyEdit
        oPAYMENTLine.BeginEdit
        oPAYMENTLine.SetAmt CStr(frmPC.Amount)
        oPAYMENTLine.SetType "W"
        AcceptSale False
        OpenDrawer
        Unload frmPC
        setInputBox "", "", "", True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PettyCash", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PettyCash"
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
        If SecurityControl(enSECURITY_ISSUEPETTYCASH, lngOPID, strName, , "Enter your security key.", "Your key is invalid", True) Then
      '  If IsRole(enSECURITY_ISOPERATOR, strPrefix, strName, lngOPID) = True Then
            oExchange.SetExchangeType ePettyCashCreditType
            oExchange.Note = frmPCC.Reason
            Set oPAYMENTLine = oExchange.PaymentLines.Add
            oPAYMENTLine.ApplyEdit
            oPAYMENTLine.BeginEdit
            oPAYMENTLine.SetAmt CStr(frmPCC.Amount)
            oPAYMENTLine.SetType "R"
            AcceptSale False
            OpenDrawer
            Unload frmPCC
            setInputBox "", "", "", True
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_PettyCashCredit", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PettyCashCredit"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_ReviewExchanges", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_ReviewExchanges"
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_ReviewDeadLetterQueue", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_ReviewDeadLetterQueue"
End Function
Private Function Action_Reprint() As eState
    On Error GoTo errHandler
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
                    strEXCHtoVoidGUID = frm.EXCHID
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
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Reprint"
End Function


Private Function Action_GetSalesRep() As eState
Dim lngID As Long

    If bCollectRepcode Then
        lngID = 0
        strName = ""
        If IsRep(enSECURITY_ISREP, strPrefix, strName, lngID) Then
            oExchange.SalesRepID = lngID
            Action_GetSalesRep = eConfirmation
        Else
            MsgBox "This is not a valid rep code, please enter again.", vbInformation + vbOKOnly
            Action_GetSalesRep = eCollectRep
        End If
    Else
        Action_GetSalesRep = eConfirmation
    End If
End Function

Private Function Action_LoadFromQuotation() As eState
    On Error GoTo errHandler
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
                LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
              '  oSL.ApplyEdit
              '  oSL.BeginEdit
              '  oSL.CalculateLine
                
                rs.MoveNext
            Loop
            DisplayTotals
'            For i = 1 To oExchange.SaleLines.Count
'                LoadTopSaleRow i, oExchange.SaleLines.Count
'            Next
            Action_LoadFromQuotation = eSale
        End If
        
        Unload frm
    End If
    oPC.dbMainDisConnect
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_LoadFromQuotation"
End Function




Private Function Action_OrderRequest() As eState
    On Error GoTo errHandler
Dim frm As New frmORREQ
Dim xmlRequestDetails As String
Dim lngDepositValue As Long

    strOrderedTitle = ""
    If strPrefix = ".." Then
        Action_OrderRequest = eStart
    Else
        frm.Show vbModal
        If frm.Cancelled Then
            Action_OrderRequest = eStart
        Else
            xmlRequestDetails = frm.GetDetailsXML
            strOrderedTitle = frm.GetDetailsForSlip
           ' strRequestDetails = frm.Customer & "~~" & frm.Item
            lngDepositValue = frm.Deposit
            oExchange.Note = xmlRequestDetails '& "~~" & frm.txtName & "~~" & "Deposit:" & Format((CDbl(lngDepositValue) / oPC.CurrencyDivisor), "###,##0.00")
            oExchange.SetExchangeType eOrderRequestType
        
            Set oSALELine = oExchange.SaleLines.Add
            oSALELine.ApplyEdit
            oSALELine.BeginEdit
            X1.ReDim 1, 0, 0, 8
            oSALELine.PID = ""
            oSALELine.Price = lngDepositValue
            oSALELine.title = "Deposit taken"
            oSALELine.Code = ""
            oSALELine.SetQty 1
            oSALELine.IsDepositItem = True
            oSALELine.CalculateLine
            oExchange.CalculateTotals
        
            LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, False, False
            DisplayTotals
            If oExchange.BalanceOwing = 0 Then
                Action_OrderRequest = eConfirmation
            Else
                Action_OrderRequest = eCollect
            End If
        End If
        'MsgBox "Next line to be fixed"
'        strOrderedTitle = "UNKNOWN"  'frm.Item & vbCrLf & "For: " & frm.Customer
        Unload frm
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_OrderRequest"
End Function
Private Function Action_Sale() As eState
10        On Error GoTo errHandler
      ''
      Dim frm As New frmQuickProductFind
      Dim strCode As String
      Dim lngQtyQuickFound As Long
      Dim arIP() As String
      Dim iLineNumber As Integer
      Dim lngEditedPrice As Long

      ''
20        If strPrefix = ".." Then
30            Action_Sale = eStart
40        ElseIf InStr(2, txtInput, "/") > 0 Then 'We are going to edit a price
50            If CheckThisPoint(M_POSPRICECHANGE) Then
60                If Not SecurityControl(enSECURITY_POSPRICECHANGE, lngOPID, strName, , "Enter your signature for price change.", "Your signature does not give you permission to change a price") Then
70                    Action_Sale = enPresentState
80                    Exit Function
90                End If
100           End If
          
110           arIP() = Split(txtInput, "/")
120           If IsNumeric(arIP(0)) Then
130               If CDbl(arIP(0)) < 32000 Then
140                   iLineNumber = CInt(arIP(0))
150               Else
160                   iLineNumber = 0
170               End If
180               If CDbl(arIP(1)) < 1000000 Then
190                   lngEditedPrice = CLng(arIP(1))
200               Else
210                   lngEditedPrice = 0
220               End If
230           End If
240           If iLineNumber > 0 And iLineNumber <= oExchange.SaleLines.Count And lngEditedPrice > 0 Then
250               oExchange.SaleLines(iLineNumber).BeginEdit
260               oExchange.SaleLines(iLineNumber).SetPrice arIP(1)
                  If lngOPID > 0 Then
                    oExchange.SaleLines(iLineNumber).ActionSignatureID = lngOPID
                  End If
270               oExchange.SaleLines(iLineNumber).ApplyEdit
    
                  UpdateSpecifiedSalesRow iLineNumber, oExchange.SaleLines(iLineNumber)
280               oExchange.CalculateTotals True
290               txtInput = ""

300           Else
310               MsgBox "Invalid entry for price alteration.", vbInformation, "Can't edit price"
320           End If
            '  MsgBox "Line to alter = " & CStr(iLineNumber) & vbCrLf & "New price: " & CStr(lngEditedPrice)
330           Action_Sale = eSale
340       ElseIf LoadProductFromCode(FNS(txtInput)) Then
350           oExchange.CalculateTotals
360           oExchange.SetExchangeType eSaleType
370           LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, False, False
380           DisplayTotals
390           Action_Sale = ePrice
400       Else
410           strCode = Replace(FNS(txtInput), """", "")

420           lngQtyQuickFound = frm.component(strCode)
430           frm.Show vbModal
440           If lngQtyQuickFound = 0 Then
450               MsgBox "Nothing found", vbInformation, "Status"
                    If oExchange.SaleLines.Count = 0 Then
                        Action_Sale = eStart
                    Else
460                     Action_Sale = eSale
                    End If
470           Else
               '   frm.Show vbModal
480               If frm.Cancelled = False Then
490                   If frm.EAN > "" Then
500                       txtInput = frm.EAN
510                       If LoadProductFromCode(FNS(txtInput)) Then
520                           oExchange.CalculateTotals
530                           oExchange.SetExchangeType eSaleType
540                           LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, False, False
550                           DisplayTotals
560                           Action_Sale = ePrice
570                       Else
580                           bValid = False
590                           Action_Sale = eStart
600                           MsgBox "Not on database or invalid action", vbInformation, "Status"
610                       End If
620                   End If
630               Else
                        If oExchange.SaleLines.Count = 0 Then
                            Action_Sale = eStart
                        Else
640                         Action_Sale = eSale
                        End If
650               End If
660               Unload frm
670           End If
680       End If
690       Exit Function
errHandler:
700       If ErrMustStop Then Debug.Assert False: Resume
710       ErrorIn "frmPOSMain.Action_Sale", , , , "Error line", Array(Erl())
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
        LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
        DisplayTotals
        Action_CreditNote = ePrice
    Else
        bValid = False
        Action_CreditNote = eStart
        MsgBox "Not on database or invalid action", vbInformation, "Status"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_CreditNote", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_CreditNote"
End Function

Private Function Action_Appro() As eState
    On Error GoTo errHandler
    If LoadProductFromCode(FNS(txtInput)) Then
        oExchange.CalculateTotals
        oExchange.SetExchangeType eApproType
        LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, False, False
        DisplayTotals
        Action_Appro = ePrice
    Else
        MsgBox "Not on database or invalid action", vbInformation, "Status"
        Action_Appro = eAppro
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_Appro", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Appro"
End Function
Private Function Action_Price() As eState
    On Error GoTo errHandler
    On Error GoTo errHandler
 
If strPrefix = ".." Then
        Action_Price = DetermineReturnToState
        RemoveSaleLine iCurrentSaleLine
        DisplayTotals
        If oExchange.LoyaltyValue > 0 Then
            lblCustomername.Caption = DisplayCustomerDetails
        End If
        Exit Function
    End If
    
    
    'THis next part is if the operator scans a second product rather than going through the process of confiming price and qty
    If Not IsISBN13(strSuffix) Then
        'Set Default price
'        If oSALELine.SetPrice(Trim(strSuffix)) Then
'            oExchange.CalculateTotals
'            LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count
'            Action_Price = eQty
'            lblCustomername.Caption = DisplayCustomerDetails
'        Else
'            SetTip "Invalid price."
'              Action_Price = ePrice
'              Exit Function
'        End If
    Else
        oSALELine.LogSalesline
        bBarcodeNotPrice = True
        Action_Price = eSale
        Exit Function
    End If
    
    If Not oSALELine Is Nothing And strSuffix > "" Then
      If CDbl(strSuffix) < 10000000 And CDbl(strSuffix) > 0 Then 'Maximum price value
        If oSALELine.Price <> CLng(Trim(strSuffix)) Then
            If CheckThisPoint(M_POSPRICECHANGE) Then
                If Not SecurityControl(enSECURITY_POSPRICECHANGE, lngOPID, strName, , "Enter your signature for price change.", "Your signature does not give you permission to change a price") Then
                    Action_Price = ePrice
                    Exit Function
                Else
                    oSALELine.ActionSignatureID = lngOPID
                End If
            Else
              SetTip "Invalid price."
            End If
        End If
      Else
              SetTip "Invalid price."
              Action_Price = ePrice
              Exit Function
      End If
    End If
    If bShiftDown And oSALELine.IsDiscountAllowed Then  'And oExchange.transactionType <> "APP" Then
        If oSALELine.SetPrice(Trim(txtInput)) Then
            oExchange.CalculateTotals
            LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
            Action_Price = eDiscount
        Else
            SetTip "Invalid price."
              Action_Price = ePrice
              Exit Function
        End If
    Else
        If oSALELine.SetPrice(Trim(strSuffix)) Then
            oExchange.CalculateTotals
            LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
            Action_Price = eQty
            lblCustomername.Caption = DisplayCustomerDetails
        Else
            SetTip "Invalid price."
              Action_Price = ePrice
              Exit Function
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Price"
End Function

Private Function Action_Qty() As eState
    On Error GoTo errHandler
Dim zLU As z_Lookup

    bItemExchange = False
    If strPrefix = ".." Then
        Action_Qty = ePrice 'DetermineReturnToState
    Else
        If oExchange.transactionType = "S" Then
            If Left(strPrefix, 1) = "-" Then
                If CheckThisPoint(M_ISSUEPOSREFUND) Then
                    If Not SecurityControl(enSECURITY_ISSUEPOSREFUND, lngOPID, strName, , "Enter your signature", "You do not have permission to issue a refund", True) Then
                        Action_Qty = eQty
                        Exit Function
                    End If
                End If
            
                bItemExchange = True
                txtInput = strPrefix & strSuffix   'Right(strSuffix, Len(strSuffix) - 1)
            End If
        End If

     '   If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(strSuffix), bItemExchange) Then
        If strSuffix = "" Or IsNumeric(strSuffix) = False Or strSuffix = "0" Then
                SetTip "Invalid quantity."
                bValid = False
                Exit Function
        End If
        If CDbl(strSuffix) < 1000000 Then  'Maximum qty value then
            If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(strSuffix), bItemExchange) Then
                   ' MsgBox "change2"
                Res = oExchange.CalculateTotals
                If Res = False Then
                    Action_Qty = eQty
                    SetTip "Calculation too large."
                    bValid = False
                    Exit Function
                End If
                LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
               
                If oExchange.TransactionTypeEnum = eSaleType Then
                    Action_Qty = eSale
                ElseIf oExchange.TransactionTypeEnum = eApproType Then
                    Action_Qty = eAppro
                Else
                    MsgBox "Error situation please report to Papyrus Support: oExchange.TransactionTypeEnum = " & CStr(oExchange.TransactionTypeEnum)
                End If
                oSALELine.ApplyEdit
                oSALELine.BeginEdit
                'Log line to database for security purposes
                Set zLU = New z_Lookup
                oSALELine.LogSalesline
            Else
                SetTip "Invalid quantity."
                bValid = False
            End If
        Else
                SetTip "Invalid quantity."
                bValid = False
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Qty"
End Function
Private Function Action_Discount() As eState
    On Error GoTo errHandler
Dim strDiscountCode As String
Dim dblDiscountRate As Double
Dim strDiscountDescription As String
Dim bCancelled As Boolean
    If strPrefix = ".." Then
        Action_Discount = DetermineReturnToState
        Exit Function
    Else
        If CheckThisPoint(M_POSDISCOUNT) Then
            If Not SecurityControl(enSECURITY_POSDISCOUNT, lngOPID, strName, , "Enter your signature.", "Your signature does not give you discount permissions") Then
                Action_Discount = DetermineReturnToState
                Exit Function
            Else
                oSALELine.ActionSignatureID = lngOPID
            End If
        End If
    End If
    strDiscountCode = UCase(Left(txtInput, 1))
    If InStr(1, strValidDiscountTypes, strDiscountCode) > 0 Then 'valid discount type
        If strDiscountCode = "X" Then
            ConnectionTimer.Enabled = False
            lngSupervisorID = 0
            If SecurityControl(enSECURITY_ISSUPERVISOR, lngSupervisorID, strName, bCancelled, "Enter security code to allow discretionary discount", "You are not entitled to offer discount.") = True Then
                Set frmDisc = New frmDiscretionaryDiscount
                frmDisc.Show vbModal
                dblDiscountRate = frmDisc.DiscountRate
                Unload frmDisc
                oExchange.SupervisorID = lngSupervisorID
            End If
            ConnectionTimer.Enabled = True
        Else
            dblDiscountRate = GetDiscount(strDiscountCode, strDiscountDescription)
        End If
        oSALELine.SetDiscountRateDbl dblDiscountRate, strDiscountDescription
        oExchange.CalculateTotals
        LoadTopSaleRow iCurrentSaleLine, oExchange.SaleLines.Count, True, False
        Action_Discount = eQty
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_Discount", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Discount"
End Function
Private Function Action_OpenDrawer() As eState
    On Error GoTo errHandler
Dim frm As frmOD
        lngSupervisorID = 0
        If CheckThisPoint(M_OPENDRAWER) Then
            If Not SecurityControl(enSECURITY_OPENDRAWER, lngSupervisorID, strName, , "Enter your signature", "You do not have permission to perform Open Drawer action", True) Then
                Exit Function
            End If
        End If
        Set frm = New frmOD
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Exit Function
        End If
            
        OpenDrawer
        lngOPID = lngStaffID
        oExchange.SupervisorID = lngSupervisorID
        oExchange.Note = frm.Reason
        oExchange.SetExchangeType eOpenDrawerType
        AcceptSale False
        Unload frm
        setInputBox "", "", "", True
        Action_OpenDrawer = eStart
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.Action_OpenDrawer", , EA_NORERAISE
'    HandleError
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_OpenDrawer"
End Function
Private Function Action_DeletePayment(iRow As String) As eState
    On Error GoTo errHandler
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_DeletePayment(iRow)", iRow
End Function
Private Function Action_DeleteSaleLine(iRow As String) As eState
    On Error GoTo errHandler
Dim lngRow As Integer
Dim bCancelled As Boolean
    lngSupervisorID = 0
    If IsNumeric(iRow) Then
        If iRow <= 0 Then
            If oExchange.SaleLines.Count = 0 And oExchange.PaymentLines.Count = 0 Then
                Action_DeleteSaleLine = eStart
            Else
                Action_DeleteSaleLine = eSale
            End If
            Exit Function
        End If
        If IsNumeric(iRow) Then
            lngRow = CInt(iRow)
            If lngRow <= oExchange.SaleLines.Count And lngRow >= 0 Then
                If CheckThisPoint(M_DELETELINE) Then
                    If SecurityControl(enSECURITY_SALELINEDELETE, lngSupervisorID, strName, bCancelled, "Enter security code", "You are not entitled to remove a sale line.") = True Then
                        RemoveSaleLine lngRow
                    End If
                Else
                      RemoveSaleLine lngRow
                End If
            End If
        Else
            MsgBox "Invalid line to delete"
        End If
    End If
    If oExchange.SaleLines.Count = 0 And oExchange.PaymentLines.Count = 0 Then
        Action_DeleteSaleLine = eStart
    Else
        Action_DeleteSaleLine = eSale
    End If
    setInputBox "", "", "", True

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_DeleteSaleLine(iRow)", iRow
End Function

Private Function Action_Void() As eState
    On Error GoTo errHandler
10        On Error GoTo errHandler
      Dim Res As Boolean
      Dim frm As frmOD
      Dim bCancelled As Boolean
        lngSupervisorID = 0
        If CheckThisPoint(M_VOID) Then
          If Not SecurityControl(enSECURITY_VOIDEXCHANGE, lngSupervisorID, strName, , "Enter your signature.", "Your signature does not give you permission to void.", True) Then
              Action_Void = eStart
              Exit Function
          End If
        End If
20        If IsNumeric(strSuffix) Then
30            iToVoid = CLng(strSuffix)
40            If X4.UpperBound(1) <= 0 Then
50                MsgBox "This transaction number is out of range ", vbInformation, "Can't do this"
60                bValid = False
70            Else
80                If iToVoid >= CLng(X4(X4.UpperBound(1), 1)) And iToVoid < oExchange.ExchangeNumber Then
90                    If (X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 12) = 0) Then
100                       If Not oExchange.CanCancel(iToVoid) Then
                              iToVoid = 0
110                           MsgBox "This exhange cannot be voided.", vbInformation, "Can't do this"
120                       Else
                            Set frm = New frmOD
                            frm.component "Cancelling exchange - provide reason"
                            frm.Show vbModal
                            If frm.Cancelled Then
                                Unload frm
                                bCancelled = True
                                Action_Void = enPresentState
                                Exit Function
                            End If
270                         lngOPID = lngStaffID
280                         oExchange.SupervisorID = lngSupervisorID
                            oExchange.Note = "VOID: " & frm.Reason
290                         oExchange.Note = oExchange.Note '& "#" & CStr(iToVoid)
300                         oExchange.SetExchangeType eVoidAction
310                         AcceptSale False
320
330                       End If
340                   Else
350                       MsgBox "This transaction has been voided already by exchange number " & CStr(X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 12)), vbInformation, "Can't do this"
360                       bValid = False
370                   End If
380               Else
390                   MsgBox "This transaction number is out of range ", vbInformation, "Can't do this"
400                       bValid = False
410               End If
420           End If
430       End If
440       Action_Void = eStart
          
450       Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Void"
End Function

'====================================================================================================
'====================================================================================================
'====================================================================================================
'====================================================================================================
'====================================================================================================


Private Sub UpdateClientFromServerFiles(Optional rs As ADODB.Recordset, Optional msg As String)
10        On Error GoTo errHandler
      Dim ar() As String

20        ar = Split(SVRMsg.Label, ",")
30        If Not rs Is Nothing Then
40            UpdatingLocalDatabase True, rs.RecordCount
50        Else
60            UpdatingLocalDatabase True, 0
70        End If
80        Select Case ar(0)
              Case "PROD"
                      'Load product updates
90                    SaveProductUpdate rs
100           Case "STAF"
                      'Load Staff Member updates
110                   SaveStaffUpdate rs
120           Case "CUST"
                      'Load Customer updates
130                   SaveCustomerUpdate rs
140           Case "ORDR"
                      'Load Customer order updates
150                   SaveCustomerOrderUpdate rs
160           Case "MARK"
                      'Load marketing updates
170                   SaveMarketingUpdate rs
180           Case "APPL"
                      'Load appro updates
190                   SaveApproUpdate rs
200           Case "APPRL"
                      'Load appro updates
210                   SaveApproRUpdate rs
220           Case "ClearCustomers"
230                   oPC.OpenLocalDatabase
240                   oPC.DBLocalConn.Execute "Delete FROM tCustomer"
250                   oPC.CloseLocalDatabase
260           Case "ClearProducts"
270                   oPC.OpenLocalDatabase
280                   oPC.DBLocalConn.Execute "Delete FROM tProduct"
290                   oPC.CloseLocalDatabase
300           Case "ClearAppros"
310                   oPC.OpenLocalDatabase
320                   oPC.DBLocalConn.Execute "Delete FROM tAPPL"
330                   oPC.CloseLocalDatabase
340           Case "ClearCustomerOrders"
350                   oPC.OpenLocalDatabase
360                   oPC.DBLocalConn.Execute "Delete FROM tCOL"
370                   oPC.CloseLocalDatabase
380           Case "ClearMarketingRules"
390                   oPC.OpenLocalDatabase
400                   oPC.DBLocalConn.Execute "Delete FROM tMarketing"
410                   oPC.CloseLocalDatabase
420           Case "ClearStaffMembers"
430                   oPC.OpenLocalDatabase
440                   oPC.DBLocalConn.Execute "Delete FROM tStaffMembers"
450                   oPC.CloseLocalDatabase
460       End Select
470       UpdatingLocalDatabase False, 0
480       Exit Sub
errHandler:
490       If ErrMustStop Then Debug.Assert False: Resume
500       ErrorIn "frmPOSMain.UpdateClientFromServerFiles(rs,msg)", Array(rs, msg), , , "Error line", Erl()
End Sub

Private Function SaveProductUpdate(rs As ADODB.Recordset) As Boolean
10        On Error GoTo errHandler
      Dim cmd As ADODB.Command
      Dim par As ADODB.Parameter
      Dim lngCnt As Long
          
20        oPC.OpenLocalDatabase
          
30        bUpdating = True
40        rs.MoveFirst
50        lngCnt = 1
60        Do While Not rs.EOF
        '  If rs!PRU_EAN = "9771995429008" Then MsgBox "HERE"
70            lngCnt = lngCnt + 1
80            If lngCnt Mod 100 = 0 Then
90                Counter "Pr", lngCnt
100           End If
110           Set cmd = New ADODB.Command
120           cmd.CommandType = adCmdStoredProc
130           cmd.ActiveConnection = oPC.DBLocalConn
              
140           If FNS(rs!PRU_LOG_TYPE) = "DEL" Then
150               cmd.CommandText = "sp_DeleteProductOnFD"
160               Set par = cmd.CreateParameter("@PID", adGUID, , , rs!PRU_P_ID)
170               cmd.Parameters.Append par
180               Set par = Nothing
190               cmd.Execute
                  
200           Else
210               cmd.CommandText = "dbo.sp_InsertProductUpdateToFD"
                      
220               Set par = cmd.CreateParameter("@PID", adGUID, , , rs!PRU_P_ID)
230               cmd.Parameters.Append par
240               Set par = Nothing
250               Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInput, 50, FNS(rs!PRU_Code))
260               cmd.Parameters.Append par
270               Set par = Nothing
280               Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 50, Left(FNS(rs!PRU_EAN), 50))
290               cmd.Parameters.Append par
300               Set par = Nothing
310               Set par = cmd.CreateParameter("@PUBLISHER", adVarChar, adParamInput, 500, Left(FNS(rs!PRU_Publisher), 50))
320               cmd.Parameters.Append par
330               Set par = Nothing
340               Set par = cmd.CreateParameter("@SERIESTITLE", adVarChar, adParamInput, 225, Left(FNS(rs!PRU_SeriesTitle), 225))
350               cmd.Parameters.Append par
360               Set par = Nothing
370               Set par = cmd.CreateParameter("@AUTHOR", adVarChar, adParamInput, 225, Left(FNS(rs!PRU_MainAuthor), 225))
380               cmd.Parameters.Append par
390               Set par = Nothing
400               Set par = cmd.CreateParameter("@TITLE", adVarChar, adParamInput, 225, Left(FNS(rs!PRU_Title), 225))
410               cmd.Parameters.Append par
420               Set par = Nothing
430               Set par = cmd.CreateParameter("@SP", adInteger, adParamInput, , FNN(rs!PRU_SP))
440               cmd.Parameters.Append par
450               Set par = Nothing
460               Set par = cmd.CreateParameter("@SSP", adInteger, adParamInput, , FNN(rs!PRU_SSP))
470               cmd.Parameters.Append par
480               Set par = Nothing
490               Set par = cmd.CreateParameter("@VATRATE", adNumeric, adParamInput, 10, FNDBL(rs!PRU_VATRATE))
500               par.Precision = 8
510               par.NumericScale = 2
520               cmd.Parameters.Append par
530               Set par = Nothing
540               Set par = cmd.CreateParameter("@PTID", adInteger, adParamInput, , FNN(rs!PRU_PTID))
550               cmd.Parameters.Append par
560               Set par = Nothing
570               Set par = cmd.CreateParameter("@SECID", adInteger, adParamInput, , FNN(rs!PRU_SECID))
580               cmd.Parameters.Append par
590               Set par = Nothing
600               Set par = cmd.CreateParameter("@NDA", adBoolean, adParamInput, , FNB(rs!PRU_NDA))
610               cmd.Parameters.Append par
620               Set par = Nothing
630               Set par = cmd.CreateParameter("@MultibuyCode", adVarChar, adParamInput, 15, FNS(rs!PRU_MultibuyCode))
640               cmd.Parameters.Append par
650               Set par = Nothing
660               cmd.Execute
                  
670           End If

680           Set cmd = Nothing
690           rs.MoveNext
700       Loop

710       bUpdating = False
720       SaveProductUpdate = True
MEX:

730       If rs.State = adStateOpen Then rs.Close
740       Set rs = Nothing
          
750       oPC.CloseLocalDatabase
          
760       Exit Function
errHandler:
770       ErrPreserve
790       LogSaveToFile Error & ", line:" & Erl()
800       GoTo MEX
End Function

Private Function SaveStaffUpdate(rs As ADODB.Recordset) As Boolean
10        On Error GoTo errHandler
      Dim NewRS As New ADODB.Recordset
      Dim sSQL As String
      Dim sName As String
      Dim i As Integer
      Dim strPos As String

20        bUpdating = True
30        strPos = "1"
40        oPC.OpenLocalDatabase
50        rs.MoveFirst
60        Do While Not rs.EOF
70            sSQL = "SELECT * FROM tStaffMembers WHERE tStaffMembers.SM_ID =" & rs!SMU_ID
80            NewRS.LockType = adLockOptimistic
90            NewRS.CursorType = adOpenDynamic
100           Set NewRS.ActiveConnection = oPC.DBLocalConn
110           NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
120           If FNS(rs!SMU_NAME) = "X" Then
130               If Not NewRS.EOF Then NewRS.Delete
140           Else
150               If NewRS.EOF Then
160                   NewRS.AddNew
170               End If
180               If Not IsNull(rs!SMU_ID) Then NewRS!SM_ID = rs!SMU_ID
190               If Not IsNull(rs!SMU_NAME) Then NewRS!SM_Name = Trim$(rs!SMU_NAME)
200               If Not IsNull(rs!SMU_Role) Then NewRS!SM_Role = rs!SMU_Role 'DO NOT TRIM this field
210               If Not IsNull(rs!SMU_Telephone) Then NewRS!SM_Telephone = Trim$(rs!SMU_Telephone)
220               If Not IsNull(rs!SMU_Mobile) Then NewRS!SM_Mobile = Trim$(rs!SMU_Mobile)
230               If Not IsNull(rs!SMU_Password) Then NewRS!SM_Password = Trim$(rs!SMU_Password)
240               If Not IsNull(rs!SMU_Shortname) Then NewRS!SM_Shortname = rs!SMU_Shortname
250               sName = Trim$(rs!SMU_NAME)
260               i = 1
DoUpdate:
270               NewRS.Update
280               If Err = -2147217887 Then
290                 NewRS!SM_Code = Left(NewRS!SM_Code, 3) & CStr(i)
300                 i = i + 1
310                 Err.Clear
320                 GoTo DoUpdate
330               ElseIf Err <> 0 Then
340                   NewRS.Close
350                   bUpdating = False
360                   GoTo errHandler
370               End If
380               NewRS.Close
390           End If
400           rs.MoveNext
410       Loop
420       SaveStaffUpdate = True
430       bUpdating = False
MEX:
440       If rs.State = adStateOpen Then rs.Close
450       Set rs = Nothing
460       Set NewRS = Nothing
470       oPC.CloseLocalDatabase
480       Exit Function
errHandler:
490       If ErrMustStop Then Debug.Assert False: Resume
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
        LogSaveToFileLocal "COU_COLID = " & CStr(rs!COU_COLID)
        sSQL = "SELECT * FROM tCOL WHERE COL_COLID = " & CStr(rs!COU_COLID)
        NewRS.LockType = adLockPessimistic
        NewRS.CursorType = adOpenStatic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL
        If NewRS.EOF Then
            NewRS.AddNew
        End If
        LogSaveToFileLocal "Pos 1"
        NewRS!COL_COLID = FNN(rs!COU_COLID)
        NewRS!COL_TPID = FNN(rs!COU_TPID)
        NewRS!COL_TRID = FNN(rs!COU_TRID)
        NewRS!COL_Date = FND(rs!COU_Date)

        NewRS!COL_CODE = FNS(rs!COU_CODE)
        NewRS!COL_PID = FNS(rs!COU_PID)
        NewRS!COL_Qty = FNN(rs!COU_QTY)
        NewRS!COL_QTYDISPATCHED = FNN(rs!COU_QTYDISPATCHED)
        LogSaveToFileLocal "Pos 2"

        NewRS!COL_Price = FNN(rs!COU_PRICE)
        NewRS!COL_DiscountRate = FNN(rs!COU_DISCOUNTRATE)
        NewRS!COL_Deposit = FNN(rs!COU_DEPOSIT)
        NewRS!COL_Depositstatus = FNS(rs!COU_DEPOSITSTATUS)

        NewRS!COL_DELETE = FNS(rs!COU_DOCSTATUS)
        LogSaveToFileLocal "Pos 3"
DoUpdate:
        NewRS.Update
        LogSaveToFileLocal "Pos 4"
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
       ErrorIn "frmPOSMain.SaveCustomerOrderUpdate(rs)", Array(rs), , , "Error line", Erl()
        MsgBox ("In error handler" & CStr(rs!COU_COLID))
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
        sSQL = "SELECT * FROM PBKSFD.dbo.tAPPL WHERE APPL_APPLID = " & rs!APPL_APPLID
        Set NewRS = New ADODB.Recordset
        
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
10        On Error GoTo errHandler
      Dim NewRS As New ADODB.Recordset
      Dim sSQL As String
      Dim sName As String
      Dim i As Integer

20        bUpdating = True
          
30        oPC.OpenLocalDatabase
      ' MsgBox "SaveMarketingUpdate" & CStr(rs.RecordCount)
40        rs.MoveFirst
50        Do While Not rs.EOF
60            NewRS.LockType = adLockOptimistic
70            sSQL = "SELECT * FROM tMarketing WHERE M_ID = " & rs!MC_ID
80            NewRS.LockType = adLockOptimistic
90            NewRS.CursorType = adOpenDynamic
100           Set NewRS.ActiveConnection = oPC.DBLocalConn
110           NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
120           If rs!MC_TYPE = "DEL" Then   'must delete record
130               If Not NewRS.EOF Then
140                   NewRS.Delete
150                   NewRS.Update
160               End If
170           Else
180               If NewRS.EOF Then
190                   NewRS.AddNew
200               End If
210               NewRS!M_PT_ID = FNN(rs!MC_PT_ID)
220               NewRS!M_SECTION_ID = FNN(rs!MC_SECTION_ID)
230               NewRS!M_CUSTTYPE_ID = FNN(rs!MC_CUSTTYPE_ID)
240               NewRS!M_DISCOUNT = FND(rs!MC_DISCOUNT)
250               NewRS!M_MINVALUE = FNN(rs!MC_MINVALUE)
260               NewRS!M_DESCRIPTION = FNS(rs!MC_DESCRIPTION)
270               NewRS!M_NODISCOUNTALLOWABLE = FNS(rs!MC_NODISCOUNTALLOWABLE)
280               NewRS!M_IDENTIFYCUSTOMER = FNN(rs!MC_IDENTIFYCUSTOMER)
290               NewRS!M_ACTIVE = FNN(rs!MC_ACTIVE)
300               NewRS!M_ID = FNN(rs!MC_ID)
DoUpdate:
310               NewRS.Update
320               NewRS.Close
330           End If
340           rs.MoveNext
350       Loop
360       SaveMarketingUpdate = True
370       bUpdating = False
          
MEX:
380       If rs.State = adStateOpen Then rs.Close
390       Set rs = Nothing
400       Set NewRS = Nothing
          
410       oPC.CloseLocalDatabase
          
420       Exit Function
errHandler:
430    MsgBox "SaveMarketingUpdate In Error handler" & Error & "Line:" & CStr(Erl())
440       If ErrMustStop Then Debug.Assert False: Resume
450       ErrSaveToFile
460       GoTo MEX
End Function

Private Function CheckThisPoint(CheckPoint As Long) As Boolean
    On Error GoTo errHandler
    If (oPC.SecurityNumber And CheckPoint) = CheckPoint Then
        CheckThisPoint = True
    Else
        CheckThisPoint = False
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CheckThisPoint(CheckPoint)", CheckPoint
End Function

Private Sub AdjustHeight(txtBox As TextBox, iLines As Long)
    Dim sText As String
    Dim r As RECT
    Dim nHeight As Long
   Dim lngScaleMode As Long
   
    'adjust the scale mode for easier calculations
    lngScaleMode = Me.ScaleMode
    Me.ScaleMode = vbPixels
    r.Right = txtBox.Width - 4 ' -4 px for the border, assumes 3d style is used
    sText = txtBox.Text
    nHeight = DrawText(Me.hdc, sText, Len(sText), r, DT_CALCRECT)
    If nHeight Then
        txtBox.Height = ((nHeight + 15) * iLines)
      '  txtBox.Height = 200
    End If
    'Restore
    Me.ScaleMode = lngScaleMode
    
End Sub

'Sub TextTrans(MyTB As Object)
'Dim TempDC As Long
'Dim Temp As String
'Dim MyLoc As RECT
'On Error Resume Next
'    Temp = MyTB.Text
'    MyLoc.Left = MyTB.Left
'    MyLoc.Top = MyTB.Top
'    MyLoc.Right = MyLoc.Left + MyTB.Width
'    MyLoc.Bottom = MyLoc.Top + MyTB.Height
'    MyTB.Parent.Cls
'    MyTB.Parent.ForeColor = MyTB.ForeColor
'    Set MyTB.Parent.Font = MyTB.Font
'    DrawText MyTB.Parent.hdc, Temp, Len(Temp), MyLoc, DT_EDITCONTROL
'    TempDC = GetDC(MyTB.hWnd)
'    BitBlt TempDC, 0, 0, MyTB.Width, MyTB.Height, MyTB.Parent.hdc, MyTB.Left, MyTB.Top, vbSrcCopy
'End Sub

'Private Sub txtInput_Validate(Cancel As Boolean)
'    If bCanSenseDrawer = False Then Exit Sub
'    If OPOSCashDrawer.DrawerOpened = True Then
'        Cancel = True
'        MsgBox "Please close drawer before continuing."
'        Exit Sub
'    End If
'
'End Sub
