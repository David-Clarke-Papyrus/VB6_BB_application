VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{DA4E6F7B-F5EE-43C5-A9A1-6BCC726F901E}#1.8#0"; "StatusBarX5.OCX"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr19.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{A9CD2883-061D-11D4-B62B-00004C937F50}#1.0#0"; "CoCash19.ocx"
Begin VB.Form frmPOSMain 
   BackColor       =   &H00E1E1E1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DiscountSet"
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11655
   FillColor       =   &H00E1E1E1&
   Icon            =   "frmPOSMain6.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   885
      Left            =   9705
      TabIndex        =   18
      Top             =   7080
      Width           =   1935
      Begin VB.Label lblProg 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label lblUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Updating"
         ForeColor       =   &H80000010&
         Height          =   285
         Left            =   60
         TabIndex        =   19
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
      TabIndex        =   15
      Text            =   "frmPOSMain6.frx":08CA
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
      TabIndex        =   14
      Text            =   "frmPOSMain6.frx":08D0
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
      Left            =   60
      OleObjectBlob   =   "frmPOSMain6.frx":08D6
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   11445
   End
   Begin StatusBarXCtl.StatusBarX SB 
      Height          =   870
      Left            =   0
      Top             =   7125
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   1535
      Appearance      =   0
      BorderColor     =   14339533
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
      HighlightColor  =   -2147483632
      HighlightDkColor=   -2147483635
      PanelCount      =   1
      Panel1Key       =   "test"
      Panel1ForeColor =   7884871
      Panel1WordWrap  =   -1  'True
      Panel1Width     =   638
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
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2910
      Left            =   60
      OleObjectBlob   =   "frmPOSMain6.frx":6415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   11415
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Height          =   1380
      Left            =   30
      OleObjectBlob   =   "frmPOSMain6.frx":AE5C
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3705
      Width           =   4485
   End
   Begin TrueOleDBGrid60.TDBGrid G5 
      Height          =   2010
      Left            =   1950
      OleObjectBlob   =   "frmPOSMain6.frx":E4E7
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   855
      Visible         =   0   'False
      Width           =   7440
   End
   Begin TrueOleDBGrid60.TDBGrid G3 
      Height          =   2400
      Left            =   1500
      OleObjectBlob   =   "frmPOSMain6.frx":12792
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   8745
   End
   Begin COCASHLib.OPOSCashDrawer OPOSCashDrawer1 
      Left            =   1920
      Top             =   5355
      _Version        =   65536
      _ExtentX        =   556
      _ExtentY        =   582
      _StockProps     =   0
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
      TabIndex        =   12
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
      TabIndex        =   13
      Top             =   6240
      Width           =   5745
   End
   Begin VB.Label txtPaymentTotal 
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
      Left            =   180
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   3075
      Width           =   1170
   End
   Begin VB.Label txtExtTotal 
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
      Left            =   8370
      TabIndex        =   9
      Top             =   3075
      Width           =   1470
   End
   Begin VB.Label txtQtyTotal 
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
      Left            =   5520
      TabIndex        =   8
      Top             =   3075
      Width           =   825
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
      Left            =   4650
      TabIndex        =   7
      Top             =   3600
      Width           =   6060
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00DACDCD&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2
      X2              =   773
      Y1              =   237
      Y2              =   237
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
      Left            =   75
      TabIndex        =   5
      Top             =   5475
      Visible         =   0   'False
      Width           =   4920
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   5340
      Top             =   5265
      _Version        =   65536
      _ExtentX        =   820
      _ExtentY        =   609
      _StockProps     =   0
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
Private WithEvents SVREvent As MSMQEvent
Attribute SVREvent.VB_VarHelpID = -1
Private WithEvents POSACKEvent As MSMQEvent
Attribute POSACKEvent.VB_VarHelpID = -1
Dim arApproReturnLines() As ReturnRec
Dim strMsg As String
Dim strOrderedTitle As String


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
End Enum
Dim enMode As enModeType

Private Enum enPaymentMode
    ePaymentMode_Cash = 1
    ePaymentMode_Cheque = 2
    ePaymentMode_CreditCard = 3
    ePaymentMode_Voucher = 4
    ePaymentMode_RedeemedDeposit = 5
    ePaymentMode_CreditNote = 6
End Enum


Private Enum enumDocumentType
    eTypReceipt = 1
    eTypVoucher = 2
    eTypCashRefund = 3
    etypCreditNote = 4
    eTypDeposit = 5
    eTypDepositRefund = 6
    eTypAppro = 7
    eTypPettyCash = 8
    eTypPettyCashCredit = 9
    eTypApproReturn = 10
    eTypOrder = 11
    eTypChangeVoucher = 12
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

Dim WithEvents oExchange As a_Exchange
Attribute oExchange.VB_VarHelpID = -1
Dim oPAYMENTLine As a_Payment
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2
Dim cCOLS As c_COLS
Dim cApps As c_APPs
Dim oSALELine As a_Sale
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
    Else
        G4.Visible = False
        G1.Visible = True
     '   G1.Height = 223
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ShowTransactions(bShow)", bShow
End Sub






Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bUpdating Then
        MsgBox "Please wait, updating local database. Try again later.", vbInformation + vbOKOnly, "Can't close now"
        Cancel = 1
    End If
End Sub



Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
        If (X1(Bookmark, 3) < 1) And (X1(Bookmark, 3) <> "") Then
            RowStyle.BackColor = vbYellow
        Else
            RowStyle.BackColor = &HFFFFFF
        End If
End Sub

Private Sub mnusavecol_Click()
    On Error GoTo errHandler
    SaveLayout Me.G4, Me.Name
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
        strMsg = SB.Panels(1).Text
        lngRecordUpdateCount = lngCnt
        lblUpdate.Caption = "updating (" & CStr(lngCnt) & ")"
        lblUpdate.Visible = True
        Me.Refresh
     '   SB.Panels(1).Text = "Updating local database . . . (" & CStr(lngCnt) & " records)"
    Else
       ' SB.Panels(1).Text = strMsg
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
        InterpretState = "ePaymentType_CreditNote"
    Case 94
        InterpretState = "ePaymentType_CreditNoteRef"
    Case 95
        InterpretState = "ePaymentType_voucher"
    Case 96
        InterpretState = "ePaymentType_ChequeRef"
    Case 97
        InterpretState = "ePaymentType_CreditCardRef"
    Case 98
        InterpretState = "ePaymentType_voucherRef"
    Case 99
        InterpretState = "ePaymentType_RedeemDeposit"
    Case 100
        InterpretState = "eRefundType_CreditNote"
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
  '  lblStatus.Caption = "Sales for " & oPC.ZSession.NominalDateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetTitleBar(pShowExchangeNumber)", pShowExchangeNumber
End Sub

Sub POSACKEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim lngResult As Integer


    Set QPOSACK = Queue
    Set POSAckMsg = QPOSACK.Receive(ReceiveTimeout:=0)
    If Not (POSAckMsg Is Nothing) Then
        If lblProg.Caption > "" Then
            lblProg.Caption = Left(lblProg.Caption, Len(lblProg.Caption) - 1)
        End If
    End If
    QPOSACK.EnableNotification POSACKEvent
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.POSACKEvent_Arrived(Queue,Cursor)", Array(Queue, Cursor), EA_NORERAISE
    HandleError
End Sub

Sub SVREvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim lngResult As Integer


    Set QSVR = Queue
    Set SVRMsg = QSVR.Receive(ReceiveTimeout:=0)
    If Not (SVRMsg Is Nothing) Then
        If Left(SVRMsg.Label, 14) = "ClearCustomers" Or Left(SVRMsg.Label, 13) = "ClearProducts" Then
            UpdateClientFromServerFiles , SVRMsg.Label
        Else
            UpdateClientFromServerFiles SVRMsg.Body, ""
        End If
    End If
    QSVR.EnableNotification SVREvent
    
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
    QI.Create
    On Error GoTo errHandler
    Err.Clear
    Set QPOSACK = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set POSACKEvent = New MSMQEvent
    QPOSACK.EnableNotification POSACKEvent
    
    'Set up our SVR queue for receiving notifications about DB changes
    Set QI = Nothing
    Set QI = New MSMQQueueInfo
    QI.PathName = oPC.NameOfPC & "\Private$\qsvr"
    On Error Resume Next
    QI.Create
    Err.Clear
    On Error GoTo errHandler
    Set QSVR = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set SVREvent = New MSMQEvent
    QSVR.EnableNotification SVREvent
    
Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetupQueues"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim Result As Integer
Dim bLoggedOnAlready As Boolean
Dim strPos As String

    bEnvironmentOK = True
    ESC = Chr(27)
    iToVoid = 0
    
    'Try to load local DB connection
    If oPC Is Nothing Then
        Set oPC = New z_POSCLIConnection
        oPC.dbConnect
        oPC.LoadProperties
    End If
    
 strPos = "1"
    bLoggedOnAlready = False
    bLogonOK = True
    oPC.SetupZSession lngStaffID, strName
    If oPC.ZSession.SupervisorID = 0 Then
        LogonOperator
        If bLogonOK = False Then
            bCloseZsession = True
            GoTo EXITHANDLER
        End If
        oPC.ZSession.SupervisorID = lngStaffID
        oPC.ZSession.SupervisorName = strName
        bLoggedOnAlready = True
    End If
strPos = "2"
    If oPC.ZSession.LoadOpenXSession = False Then
        
        oPC.ZSession.OpSession.Start_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
        If oPC.ZSession.OpSession.SupervisorID = 0 Then
            If bLoggedOnAlready = False Then
                LogonOperator
                If bLogonOK = False Then
                    bCloseXsession = True
                    bCloseZsession = True
                    GoTo EXITHANDLER
                End If
            End If
            oPC.ZSession.OpSession.OperatorID = lngStaffID
            oPC.ZSession.OpSession.Name = strName
        End If
    End If
strPos = "3"
    SetForCOLSVisible False
'    SetForAPPSVisible False
    If oPC.DriveDrawer = True Then  'THere is a COM connected Cash Drawer
        MSComm1.Settings = oPC.COMPORTSettings
        MSComm1.CommPort = oPC.CashDrawerPort
        If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
    Else                            'There is a cash drawer connected to the Printer
        OPOSCashDrawer1.DeviceEnabled = True
    End If
strPos = "4"
    SetupQueues
strPos = "5"
    SetupPrinter
strPos = "6"
    LoadVoucherTypes
strPos = "7"
    LoadDiscountTypes
strPos = "8"
    G1.Array = X1
    G4.Height = 380
strPos = "9"
    ReSendExchanges
    
EXITHANDLER:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Load", , EA_NORERAISE, , "strpos", Array(strPos)
    HandleError
End Sub


Private Sub SetupPrinter()
Dim lngResult As Long
    If oPC.PrintSlips = True Then
        With OPOSPOSPrinter1
            lngResult = .Open(oPC.Printername)
            
            If lngResult = 0 Then
            lngResult = .ClaimDevice(1)
                If lngResult = OPOS_SUCCESS Then
                    .ClaimDevice 1000
                    .DeviceEnabled = True
                    .MapMode = PTR_MM_METRIC
                    .RecLetterQuality = True
                    .RecLineChars = 40
                Else
                    MsgBox "The till printer is not online. This application will close."
                    bEnvironmentOK = False
                    Exit Sub
                End If
            Else
                MsgBox "The till printer is not online. This application will close."
                bEnvironmentOK = False
                Exit Sub
            End If
        End With
        
        If oPC.DriveDrawer = False Then
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
        End If
        
        Me.lblState.Visible = False
    End If

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
                iDisc = CInt(Mid(Left(str, InStr(1, str, "%") - 1), k + 1))
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
Private Function LogonOperator() As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean
Dim Res As Boolean
    Res = False
    Do Until Res = True
        If Not SecurityControl(eOperator, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
            If bCancelled Then Res = True
            bLogonOK = False
        Else
            Res = True
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
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
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
    X4.Clear
    X4.ReDim 1, 1, 1, 13
    LoadExchanges
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
    SB.Panels(1).Text = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Stat(msg)", msg
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If Not bCloseXsession And Not bCloseZsession And Not bLogonOK = False Then
        If MsgBox("You want to close this application? Confirm", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
    If bEnvironmentOK = True Then
        bUnloading = True
        ConnectionTimer.Enabled = False
        CloseApplication Cancel
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
    Me.SB.Panels(1).Text = "Wait. The local data is being transmitted to the server."
    
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
    If IsNumeric(strSuffix) Then
        lngRow = CLng(strSuffix)
strPos = "02"
        If lngRow <= X4(X4.UpperBound(1) - 1, 1) And lngRow > 0 Then
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
                    PrintSalesSlip 1, True
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
            Stat "(X) Cancel,(CN)Credit note,(C)Cash"
        Case eInvoiceno
            setInputBox "", "", "", False
            lblInput.Caption = "Select line number of invoice to pay "
            Stat " .. to reverse"
        Case eCollect, eInvoiceMode
            G5.Visible = False
            setInputBox "", "", "", False
            If oExchange.transactionType = "RDEP" Or oExchange.TotalPayable < 0 Then
                lblInput.Caption = "Select refund type "
                Stat " .. to reverse,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque"
            ElseIf oExchange.transactionType = "AR" Then
                lblInput.Caption = "Select payment type "
                Stat " (X) to Cancel,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque"
            Else
                lblInput.Caption = "Select payment type "
                Stat " .. to reverse,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque"
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
            ClearTextFields
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
            End Select
        Case eDiscount                       ', eDiscountCashRefund, eDiscountCreditNote, eDiscountAppro
            Stat "   .. to reverse"
            setInputBox "", "", "", True
            lblInput.Caption = "Select discount type "
            txtDiscounts = Replace(oPC.DiscountSet, ";", vbCrLf)
            txtDiscounts.Visible = True
            
        Case elogin
            lblInput.Caption = "Staff code."
        
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
        Case ePaymentType_CreditNote
            lblInput.Caption = "Credit note value."
            Stat "'.. to reverse"
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case ePaymentType_CreditNoteRef
            setInputBox "", "", "", True
            lblInput.Caption = "Credit note reference."
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
                        Stat "Scan or (X)Cancel trans.,(C)Cash refund,(CC)Reverse credit card,(CN)Issue credit note,(Dn)Del sale,(DPn)Del paymt."
                    Else
                        Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(CN)Credit note,(Q)Cheque,(Dn)Del prod,(DPn)Del paymt."
                    End If
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(CN)Credit note,(Q)Cheque,(Dn)Del sale,(DPn)Del paymt."
                End If
            Else
                If oExchange.BalanceOwing < 0 Then
                    Stat "Scan or (X)Cancel trans.,(C)Cash refund,(CC)Reverse credit card,(CN)Issue credit note,(Dn)Del sale,(DPn)Del paymt.,(FC) Find customer"
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(CN)Credit note,(Q)Cheque,(Dn)Del sale,(DPn)Del paymt.,(FC) Find customer"
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
            Stat "Start by scan or (A)Appro,(AR)Appro return,(Vn)Void,(OR)Place order,(RDEP)Refund deposit,(PC)Petty cash,(PCR)Petty cash return"
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
    X1.DeleteRows (iRow)
    G1.ReBind
    iCurrentSaleLine = iCurrentSaleLine - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemoveSaleLine(iRow,pCurrent)", Array(iRow, pCurrent)
End Sub

Private Sub RejectSale()
    On Error GoTo errHandler
    oExchange.CancelEdit
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
    oExchange.SetExchangeType eSaleType
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
    txtInput = ""
    lblCustomername.Caption = ""
    lblReplacement.Visible = False
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    iToVoid = 0
    bCustomerVisible = False
    bSaleActive = False
    SetPresentState eStart
    enMode = emode_Sale
    SetTitleBar False
    SetForCOLSVisible False
'    SetForAPPSVisible False
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

'Save and send exchange
strPos = "00"
    bItemExchange = False
    If oExchange.NeedsCustomerInfo = True Then
strPos = "01"
        bCustomerOK = False
        Do Until bCustomerOK = True
strPos = "02"
            If GetCustomer() Then
                lblCustomername.Caption = DisplayCustomerDetails
            End If
            Set frmCustID = New frmIDCustomer
            If oExchange.Note > "" Then
                frmCustID.component oExchange.Note
            End If
strPos = "03"
            frmCustID.Show vbModal
            If oExchange.Customer.Name = "" Then
                oExchange.Note = frmCustID.CustomerName
                strMsg = "Confirm customer details:" & vbCrLf & "Name: " & frmCustID.CustomerName & vbCrLf & ""
            Else
                oExchange.Note = vbNullString
                strMsg = "Confirm customer details:" & vbCrLf & "Name: " & oExchange.Customer.Name & vbCrLf & "A/c No.;" & oExchange.Customer.AcNo
            End If
strPos = "04"
            oSALELine.Counterfoil = frmCustID.Counterfoil
strPos = "05"
            If MsgBox(strMsg, vbInformation + vbYesNo) = vbNo Then
                ClearCustomer
                bCustomerOK = False
            Else
                bCustomerOK = True
            End If
        Loop
    End If
strPos = "06"
        
    If oExchange.CustomerToBeCredited Then 'THis is to determine in the case of an exchange (not a RDEP) if money is to go out
        If oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_Cash) Then
            oExchange.SetExchangeType ereturntype
        ElseIf oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditNote) Then
            oExchange.SetExchangeType eCreditNoteType
        ElseIf oExchange.PaymentLines(1).PaymentType = ConvertPaymentStateToCode(ePaymentMode_CreditCard) Then
            oExchange.SetExchangeType eSaleType
        End If
    End If
strPos = "07"
    oExchange.CalculateTotals
    oExchange.OperatorID = lngOPID
    oExchange.StaffName = strName
strPos = "08"
    If iToVoid > 0 Then
        oExchange.ToVoid = iToVoid
    End If
strPos = "09"
    lngLowerBound = X4.LowerBound(1)
    lngRow = X4.Find(lngLowerBound, 1, iToVoid)
    Do While lngRow >= lngLowerBound
        X4(lngRow, 12) = oExchange.ExchangeNumber
        lngRow = X4.Find(lngRow + 1, 1, iToVoid)
    Loop
strPos = "10"
    G4.Refresh
    
    oExchange.ApplyEdit
strPos = "11"
    oPC.DBLocalConn.BeginTrans
strPos = "12"
    oExchange.PostExchange
strPos = "13"
    oPC.DBLocalConn.CommitTrans
    
    AddExchange
strPos = "14"
    SendPOSExchange oExchange.ExchangeID, oExchange.OPSID, oExchange.ZID
strPos = "15"

'Print Till Slip
    If oPC.PrintSlips = True Then
        Select Case oExchange.transactionType
        Case "S", "AR"
            If (oExchange.Customer.CustomerType = "L1" Or oExchange.Customer.CustomerType = "L2" Or oExchange.Customer.CustomerType = "L3") And (oExchange.LoyaltyValue > 0) Then
                PrintLoyaltyVoucher
            End If
            PrintSalesSlip oPC.InvoiceCopyCount
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
            PrintORDERSlip oPC.ApproCopyCount
        End Select
    
    'If there is a CN being paid out as change - we must print it
        If bIssueCreditNote Then
            PrintCNasChange oExchange.ChangeGivenF, 1, False
            bIssueCreditNote = False
        End If
    End If
strPos = "16"
'Start new exchange
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
    oExchange.SetExchangeType eSaleType
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AcceptSale", , , , "EXCH:IsEditing,strPOS", Array(oExchange.IsEditing, strPos)
End Sub

Private Sub ClearTextFields()
    On Error GoTo errHandler
    txtExtTotal = ""
    txtQtyTotal = ""
    txtVatValue = ""
    txtPaymentTotal = ""
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
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearSaleLines"
End Sub
Private Function LoadProductFromCode() As Boolean
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


    Set oLU = New z_Lookup
    strEAN = Trim$(txtInput)
    strCode = Trim$(txtInput)
    Call oLU.GetProduct(strEAN, strCode, strPID, strTitle, strAuthor, lngVatrate, lngPrice, lngDiscount, lngLoyaltyDiscount, bIdentifyCustomer, bNoDiscountAllowable, strDiscountRule)
     If strPID = "" Then
        LoadProductFromCode = False
        Exit Function
    End If
    If lngLoyaltyDiscount < oPC.LoyaltyRate Then
        lngLoyaltyDiscount = oPC.LoyaltyRate
    End If
    Set oLU = Nothing
    Set oSALELine = oExchange.SaleLines.Add
    oExchange.IdentifyCustomer = bIdentifyCustomer
     If oExchange.IdentifyCustomer = True And oExchange.Note = "" Then
         Set frmCustID = New frmIDCustomer
         frmCustID.component oExchange.Note
         frmCustID.Show vbModal
         oExchange.Note = frmCustID.CustomerName
         oSALELine.Counterfoil = frmCustID.Counterfoil
         Unload frmCustID
     End If
     
     
     iCurrentSaleLine = iCurrentSaleLine + 1
     X1.ReDim 1, iCurrentSaleLine, 1, 8
     oSALELine.Title = strTitle
     oSALELine.MainAuthor = strAuthor
     oSALELine.Price = lngPrice
     oSALELine.SetQty "1", False
     oSALELine.VATRATE = lngVatrate
     oSALELine.DiscountRate = CDbl(lngDiscount)   'Set the discount and loyalty discount from the product data
     oSALELine.LoyaltyDiscount = lngLoyaltyDiscount         ''
     oSALELine.NoDiscountAllowed = bNoDiscountAllowable     ''
     oSALELine.DiscountRule = strDiscountRule               ''
     oSALELine.PID = strPID
     If strCode > "" Then
         oSALELine.Code = strCode
     Else
         oSALELine.Code = strEAN
     End If
     
     
     oSALELine.ApplyEdit
     oSALELine.BeginEdit
     LoadProductFromCode = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadProductFromCode"
End Function
Private Sub oExchange_Recalculate()
    On Error GoTo errHandler
    If bUnloading Then Exit Sub
    DisplayTotals
    lblCustomername.Caption = DisplayCustomerDetails

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_Recalculate", , EA_NORERAISE
    HandleError
End Sub
Private Function GetCustomer() As Boolean
    On Error GoTo errHandler
Dim frm As New frmBrowseCustomers2
    GetCustomer = False
    frm.Show vbModal
    If frm.CustomerName > "" Then
        bCustomerVisible = True
        strCustomername = frm.CustomerName
        G3.Caption = frm.CustomerName
        lngCustomerID = frm.CustomerID
        oExchange.SetCustomer lngCustomerID
        oExchange.Note = frm.CustomerName
        GetCustomer = True
    Else
        ClearCustomer
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetCustomer"
End Function
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
    
    DisplayCustomerDetails = strDetails
    
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
'Private Sub SetForAPPSVisible(pYes As Boolean)
'    On Error GoTo errHandler
'    If pYes Then
'        G5.Visible = True
'        G5.ZOrder 0
'    Else
'        G5.Visible = False
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.SetForINVSVisible(pYes)", pYes
'End Sub

Private Sub LoadSaleRow(Index As Integer)
    On Error GoTo errHandler
Dim i As Long
    G1.Visible = True
    X1.Value(Index, 1) = oExchange.SaleLines(Index).CodeF
    X1.Value(Index, 2) = IIf(enPresentState = eSelectDepositLine, "(DEP)", "") & oExchange.SaleLines(Index).Title & " (" & oExchange.SaleLines(Index).MainAuthor & ")"
    X1.Value(Index, 3) = oExchange.SaleLines(Index).Qty
    X1.Value(Index, 4) = oExchange.SaleLines(Index).PriceF
    X1.Value(Index, 5) = oExchange.SaleLines(Index).DiscountRateF
    X1.Value(Index, 6) = oExchange.SaleLines(Index).PLessDiscExtF
    If oExchange.transactionType <> "INV" Then
        X1.Value(Index, 7) = oExchange.SaleLines(Index).PLessDiscExtVATF & "(" & oExchange.SaleLines(Index).VATRateF & ")"
    End If
    G1.ReBind
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadSaleRow(Index)", Index
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayTotals"
End Sub
Private Sub LoadPaymentRow(iIndex As Integer)
    On Error GoTo errHandler
Dim i As Long
    G2.Visible = True
    X2.Value(iIndex, 3) = oExchange.PaymentLines(iIndex).ReferenceComplete
    X2.Value(iIndex, 2) = oExchange.PaymentLines(iIndex).AmtF
    X2.Value(iIndex, 1) = oExchange.PaymentLines(iIndex).PaymentTypeF
    G2.Array = X2
    G2.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadPaymentRow(iIndex)", iIndex
End Sub
Private Sub RefreshSaleDisplay()
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To oExchange.SaleLines.Count
        LoadSaleRow i
    Next
    For i = 1 To oExchange.PaymentLines.Count
        LoadPaymentRow i
    Next
    oExchange.CalculateTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefreshSaleDisplay"
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

Private Sub PrintSalesSlip(pCopyCount As Integer, Optional bReprint As Boolean)
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
    
    ReDim idBuf(1 To oExchange.SaleLines.Count)
    For i = 1 To oExchange.SaleLines.Count
        If Not oExchange.SaleLines(i).IsDeleted Then
            idBuf(i).TType = IIf(oExchange.SaleLines(i).Qty < 0, "R ", "S ")
            idBuf(i).Name = oExchange.SaleLines(i).Title
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
    
    For c = 1 To pCopyCount
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
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                
            PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
            PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
            
        End With
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintSalesSlip(pCopyCount,bReprint)", Array(pCopyCount, bReprint)
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
            idBuf(i).Name = oExchange.SaleLines(i).Title
            idBuf(i).Disc = oExchange.SaleLines(i).DiscountRateF
            idBuf(i).Ext = oExchange.SaleLines(i).PLessDiscExtF
            idBuf(i).At = oExchange.SaleLines(i).QtyF & " @ " & oExchange.SaleLines(i).PriceF
            idBuf(i).Alteration = oExchange.SaleLines(i).PriceAlteration
            idBuf(i).DiscDesc = oExchange.SaleLines(i).DiscountRule
        End If
    Next i
    
    For c = 1 To pCopyCount
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
        With OPOSPOSPrinter1
            PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1, bReprint      'Print header
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
            
            
            sBuf = "Order:" & strOrderedTitle
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sBuf + vbLf
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|1200uF"     'create gap
            
            PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
            PrintFooter c, ConvertToType(oExchange.transactionType), OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
        End With
    Next

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

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
        With OPOSPOSPrinter1
            PrintHeader ConvertToType("CV"), OPOSPOSPrinter1, bReprint      'Print header
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                
            With OPOSPOSPrinter1
                
                
                sBuf = "Change Voucher"
                sExt = pAmtF
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
            End With
        
     '       .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(c)
            PrintFooter c, ConvertToType("C"), OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            .AsyncMode = False
        End With
        
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

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintLoyaltyVoucher"
End Sub
Private Sub PrintPettyCashVoucher(pCopyCount As Integer)
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
    
        With OPOSPOSPrinter1
            PrintHeader eTypPettyCash, OPOSPOSPrinter1           'Print header
            .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
            If oExchange.PaymentLines(1).PaymentType = "C" Then
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash: " & oExchange.PaymentLines(1).AmtF + vbLf
            Else
                .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash Credit: " & oExchange.PaymentLines(1).AmtF + vbLf
            End If
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
            .PrintNormal PTR_S_RECEIPT, oExchange.Note
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|5500uF"     'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
        End With
        
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
    MakePrintStringDetail = sType & sBuf & sAt & ESC + "|N" & sExt
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
        With OPOSPOSPrinter1
            PrintHeader eTypDepositRefund, OPOSPOSPrinter1           'Print header
            .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
           
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "REFUNDED.: " & oExchange.PaymentLines(1).AmtF + vbLf
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
strPos = "1"
    ReDim idBuf(1 To oExchange.SaleLines.Count)
    For j = 1 To oExchange.SaleLines.Count
        If Not oExchange.SaleLines(j).IsDeleted Then
strPos = "1a:" & CStr(j)
            idBuf(j).TType = IIf(oExchange.SaleLines(j).Qty < 0, "R ", "S ")
            idBuf(j).Name = oExchange.SaleLines(j).Title
            idBuf(j).Disc = oExchange.SaleLines(j).DiscountRateF
            idBuf(j).Ext = oExchange.SaleLines(j).PLessDiscExtF
            idBuf(j).At = oExchange.SaleLines(j).QtyF & " @ " & oExchange.SaleLines(j).PriceF
            idBuf(j).Alteration = oExchange.SaleLines(j).PriceAlteration
            idBuf(j).DiscDesc = oExchange.SaleLines(j).DiscountRule
        End If
    Next j
strPos = "2"
    For i = 1 To pCopyCount
        With OPOSPOSPrinter1
strPos = "3"
            PrintHeader eTypDeposit, OPOSPOSPrinter1             'Print header
strPos = "4"
            .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
            For j = LBound(idBuf) To UBound(idBuf)          'Print each line
                If .ResultCode <> OPOS_SUCCESS Then Exit For
                sAt = idBuf(j).At
                sBuf = idBuf(j).Name
                sExt = idBuf(j).Ext
                sType = idBuf(j).TType
                sDisc = idBuf(j).Disc
                bPriceAlteration = idBuf(j).Alteration
strPos = "5a " & CStr(i) & " " & CStr(.RecLineChars) & " , " & sType & " , " & sBuf & " , " & sAt & " , " & sExt & " , " & sDisc
                
                sValue = MakePrintStringDetail(.RecLineChars, sType, sBuf, sAt, sExt, sDisc, bPriceAlteration)
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
                .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(1).CodeF & ":DEPOSIT PAID" & vbLf
            Next j
strPos = "6"
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Deposit paid: " & oExchange.TotalPayableF + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Change: " & oExchange.ChangeGivenF + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
            .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(i)
            PrintFooter i, eTypDeposit, OPOSPOSPrinter1          'print footer
strPos = "7"
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .CutPaper 90
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
strPos = "8"
    
            'Back to the synchronous mode
            .AsyncMode = False
            
        End With
    Next i

' When a cursor is back to its default shape, it means the process ends.
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
Dim oSale As a_Sale
    Select Case oExchange.transactionTypeenum
    Case eSaleType, ereturntype, eCreditNoteType
        For Each oSale In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = oSale.CodeF  ' & " (" & oSale.QtyF & ") " & oSale.TitleF(30) & " " & oSale.PLessDiscF
            X4.Value(lngSalesItemCount, 6) = oSale.QtyF
            X4.Value(lngSalesItemCount, 7) = oSale.TitleF(30) & IIf(oExchange.ToVoid > 0, " (Voids:" & oExchange.ToVoid & ")", "")
            X4.Value(lngSalesItemCount, 8) = oSale.PriceF
            X4.Value(lngSalesItemCount, 9) = oSale.PLessDiscExtF
            
            'Add 4 columns
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 11) = oSale.PID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next
    Case eApproType
        For Each oSale In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = oSale.CodeF  ' & " (" & oSale.QtyF & ") " & oSale.TitleF(30) & " " & oSale.PLessDiscF
            X4.Value(lngSalesItemCount, 6) = oSale.QtyF
            X4.Value(lngSalesItemCount, 7) = oSale.TitleF(30) & IIf(oExchange.ToVoid > 0, " (Voids:" & oExchange.ToVoid & ")", "")
            X4.Value(lngSalesItemCount, 8) = oSale.PriceF
            X4.Value(lngSalesItemCount, 9) = oSale.PLessDiscExtF
            
            'Add 4 columns
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 11) = oSale.PID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next
    Case eDepositType, eOrderRequestType
        For Each oSale In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "DEPOSIT"
            X4.Value(lngSalesItemCount, 8) = oSale.PriceF
            X4.Value(lngSalesItemCount, 9) = oSale.PLessDiscExtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
        Next
    Case eApproReturnType
        For Each oSale In oExchange.SaleLines
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = oSale.Qty
            X4.Value(lngSalesItemCount, 7) = oSale.TitleF(30)
            X4.Value(lngSalesItemCount, 8) = oSale.PriceF
            X4.Value(lngSalesItemCount, 9) = oSale.PLessDiscExtF
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
    Case ePettyCashType
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "PETTY CASH" & ":" & oExchange.Note
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
    End Select
    G4.Array = X4
    G4.ReBind
    G4.Bookmark = lngSalesItemCount
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

    For i = 1 To G4.Columns.Count
        G4.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G4.Columns(i - 1).Width)
    Next
    
    ZID = oPC.ZSession.Current_Z_Session_ID

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = oPC.DBLocalConn
    cmd.CommandText = "q_ExchangeDetails"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , ZID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@TITLELENGTH", adInteger, adParamInput, , 50)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@CurrencyDivisor", adInteger, adParamInput, , 100)
    cmd.Parameters.Append prm
    Set prm = Nothing
   
   
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
    G4.Array = X4
    G4.ReBind
    G4.Bookmark = lngSalesItemCount


    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadExchanges"
End Sub

Private Sub PrintTotals(eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter)
    On Error GoTo errHandler
Dim sBuf As String
Dim sExt As String
Dim sValue As String
Dim oPmt As a_Payment

    Select Case eDocumentType
    Case eTypReceipt, eTypApproReturn
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
    Case eTypCashRefund
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
    Case eTypAppro
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
    Case etypCreditNote
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
            
            
            sBuf = "Credit note"
            sExt = oExchange.TotalLessDiscExtF
            sValue = MakePrintString(.RecLineChars, sBuf, sExt)
            .PrintNormal PTR_S_RECEIPT, sValue + vbLf
        End With
    Case eTypOrder
        With pPrinter
            sBuf = "Deposit received"
            sExt = oExchange.TotalPayableF
            sValue = MakePrintString(.RecLineChars, sBuf, sExt)
            .PrintNormal PTR_S_RECEIPT, ESC + "|bC" + sValue + vbLf
            
            
            
            sBuf = "Total"
            sExt = oExchange.TotalLessDiscExtF
            sValue = MakePrintString((.RecLineChars \ 2), sBuf, sExt)     'Because the width of characters of total is doubled, take this into consideration when computing.
            .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + sValue + vbLf
            
                sBuf = "Total"
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

    Select Case eDocumentType
    Case eTypReceipt
        With pPrinter
            .AsyncMode = True
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
            .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "TAX INVOICE" + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
            ar = Split(oPC.POSBranchAddress, ",")
            For i = 0 To UBound(ar)
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
            Next i
            fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
            .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eTypApproReturn
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
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eAppro
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
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
        
    Case eTypCashRefund
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eTypChangeVoucher
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case etypCreditNote
        With pPrinter
          .AsyncMode = True
          .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
          .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "CREDIT NOTE" + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
          ar = Split(oPC.POSBranchAddress, ",")
          For i = 0 To UBound(ar)
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
          Next i
            fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
          .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eTypDepositRefund
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    
    Case eTypDeposit
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eTypVoucher
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eTypAppro
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
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
    Case eTypPettyCash
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
          If bReprint = True Then
                .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
          End If
        End With
    Case eTypOrder
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.StationName & "," & oExchange.SalesPersonName & vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
          If bReprint = True Then
                .PrintNormal PTR_S_RECEIPT, ESC + "|uC" + ESC + "|bC" + ESC + "|2C" + "REPRINT" + vbLf
                .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
          End If
        End With
        
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
    Case eTypReceipt, eTypCashRefund, etypCreditNote, eTypDeposit, eTypDepositRefund, eTypPettyCash, eTypAppro, eTypApproReturn
        With pPrinter
            .AsyncMode = True
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
            .PrintNormal PTR_S_RECEIPT, ESC + "|700uF"
            .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
            If oExchange.Customer.ID > 0 Then
'            If oExchange.Customer.NameAndCodeandType(.RecLineChars) > "" Then
'                .PrintNormal PTR_S_RECEIPT, oExchange.Customer.NameAndCodeandType(.RecLineChars) + vbLf
'            End If
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
    Case "S"
        ConvertToType = eTypReceipt
    Case "R"
        ConvertToType = eTypCashRefund
    Case "C"
        ConvertToType = etypCreditNote
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

Public Sub SendPOSExchange(pEXCHID As String, pOPSID As String, pZID As String)
    On Error GoTo errHandler
Dim msg As String
Dim sFileName As String
Dim oShapeDB As New z_POSCLIConnectionShape
Dim sSQL As String
Dim tmprs As ADODB.Stream
Dim strPos As String


    Check (oShapeDB.dbConnecttoShape = 0), ERR_GENERAL, "Failed to create database connection!"
    If Not rsZSession Is Nothing Then
        If rsZSession.State <> 0 Then
            rsZSession.Close
        End If
    End If
    sSQL = "SHAPE {SELECT 'E' as TYP,tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession " _
        & " APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession " _
        & " APPEND (( SHAPE {SELECT EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE, " _
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
  '  Set rsZSession.ActiveConnection = Nothing
    DisconnectAll rsZSession
    
    
    
    'Place  POS MESSAGE in queue for server to fetch
    DispatchMessage rsZSession
    
    lblProg.Caption = lblProg.Caption & "X"
    If pEXCHID > "" Then
        SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
        oPC.DBLocalConn.Execute SQL
    End If
    oShapeDB.dbCloseConnectShape
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SendPOSExchange(pEXCHID,pOPSID,pZID)", Array(pEXCHID, pOPSID, pZID)
End Sub
Private Sub ReSendExchanges()
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strExchNum As String
Dim sr As New Scripting.FileSystemObject
Dim strFilename As String

    strFilename = "\\" & oPC.NameOfPC & "\PBKS_S\RESEND.TXT"
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
        strExchNum = oTF.ReadLinefromTextFile
        SendPOSExchange_ByExchangeNumber strExchNum
    
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

    Set rs = New ADODB.Recordset
    rs.Open "SELECT     EXCH_ID, EXCH_ZSessionID, EXCH_OPSESSIONID, EXCH_Voided, EXCH_VOIDS FROM dbo.tExchange WHERE EXCH_NUMBER = '" & pEXCHNumber & "'", oPC.DBLocalConn, adOpenStatic
    If Not rs.EOF Then
        pZID = rs.Fields("EXCH_ZSessionID")
        pOPSID = rs.Fields("EXCH_OPSESSIONID")
        pEXCHID = rs.Fields("EXCH_ID")
    End If
    rs.Close
    Set rs = Nothing
    Check (oShapeDB.dbConnecttoShape = 0), ERR_GENERAL, "Failed to create database connection!"
    If Not rsZSession Is Nothing Then
        If rsZSession.State <> 0 Then
            rsZSession.Close
        End If
    End If
    sSQL = "SHAPE {SELECT 'E' as TYP,tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession " _
        & " APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession " _
        & " APPEND (( SHAPE {SELECT EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE, " _
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
  '  Set rsZSession.ActiveConnection = Nothing
    DisconnectAll rsZSession
    
    
    
    'Place  POS MESSAGE in queue for server to fetch
    DispatchMessage rsZSession
    
    lblProg.Caption = lblProg.Caption & "X"
    If pEXCHID > "" Then
        SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
        oPC.DBLocalConn.Execute SQL
    End If
    oShapeDB.dbCloseConnectShape
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
Private Sub DispatchMessage(rs As ADODB.Recordset)
Dim oStream As ADODB.Stream
'Dim oTR As MSMQTransaction

    On Error GoTo errHandler
    Set QI = New MSMQQueueInfo
    QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
    Set QPOS = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    
    QI.FormatName = "DIRECT=OS:" & oPC.NameOfPC & "\Private$\QPOSAck"
    
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
    Case ePaymentMode_CreditNote
        ConvertPaymentStateToCode = "CN"
    Case ePaymentMode_Cash
        ConvertPaymentStateToCode = "C"
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
    
    Select Case enPresentState
    Case eAppro
        Select Case strPrefix
        Case "X"
            GetNewState = Action_CancelSale
        Case "F"
            GetNewState = eConfirmation
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
                    GetNewState = ePaymentType_CreditNote
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
    Case ePaymentType_CreditNote
        GetNewState = Action_PaymentType_CreditNote()
    Case ePaymentType_CreditNoteRef
        GetNewState = Action_PaymentType_CreditNoteRef()
    Case ePaymentType_RedeemDeposit
        GetNewState = Action_PaymentType_RedeemDeposit()
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
            GetNewState = Action_Price()
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
        Case "XEND"
            GetNewState = Action_eXTerminate
        Case "ZEND"
            GetNewState = Action_eZTerminate
        Case "A"
            enMode = emode_Appro
            GetNewState = Action_SearchCustomer("A")
        Case "RDEP"
            enMode = eMode_ReturnDeposit
            GetNewState = Action_SearchCustomer("R")
        Case "AR"
            enMode = eMode_ApproReturn
            GetNewState = Action_SearchCustomer("I")
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
        Case "X"
            GetNewState = Action_CancelSale
        Case Else
            If oExchange.transactionType = "RDEP" Or oExchange.TotalPayable < 0 = True Then    'This is a refund credit note
                    Select Case strPrefix
                    Case "C"
                        RefundPayment ePaymentMode_Cash
                        GetNewState = eConfirmation
                    Case "CN"
                        RefundPayment ePaymentMode_CreditNote
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
                Case "CN"
                    GetNewState = ePaymentType_CreditNote
'Commented 5/3/06 Deposits are redeemed after selecting a customer - with the order on the screen
'                Case "RD"
'                    If Valid_RedeemDeposit_Arg(strSuffix) Then
'                        GetNewState = ePaymentType_RedeemDeposit
'                    Else
'                        GetNewState = enPresentState
'                    End If
                Case Else
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
            Case "CN"
                RefundPayment ePaymentMode_CreditNote
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
        If enPresentState = eConfirmation Then
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
            c = Mid(pRaw, i, 1)
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
        For i = 1 To iMax
            c = Mid(pRaw, i, 1)
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
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SeparateInput(pRaw,pPrefix,pSuffix)", Array(pRaw, pPrefix, pSuffix)
End Function
Private Function Action_CancelSale() As eState
    If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        RejectSale
    Else
        Action_CancelSale = enPresentState
    End If
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
        oPC.ZSession.OpSession.Close_OP_Session
    End If
End Function
Private Function Action_eZTerminate() As eState
    If MsgBox("Close Z session?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
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
    LoadSaleRow iCurrentSaleLine
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

    On Error GoTo errHandler
    If oPC.DriveDrawer = True Then
        MSComm1.Output = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10)
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
        X1.ReDim 1, iCurrentSaleLine, 1, 7
        oSALELine.PID = X3(lngTmp, 13)
        oSALELine.Price = FNN(X3(lngTmp, 12))
        oSALELine.Title = X3(lngTmp, 5)
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
        X1.ReDim 1, iCurrentSaleLine, 1, 7
        oSALELine.PID = X3(lngTmp, 13)
        oSALELine.Price = FNN(X3(lngTmp, 12))
        oSALELine.Qty = (FNN(X3(lngTmp, 14))) * -1
        oSALELine.Title = X3(lngTmp, 5)
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

Private Function CreateApproReturnAndInvoice(pTPID As Long, pInvValue As Long) As Boolean
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
    ReDim arApproReturnLines(0)
   ' ReDim arApproReturnLines(1)
    Res = frm.ApproReturnData(lngAPPID, strAppCOde, dteAppDate, pInvValue, arApproReturnLines)
    Unload frm
    If Res = False Then
        Exit Function
    End If
    If UBound(arApproReturnLines) > 0 Then
    For i = 1 To UBound(arApproReturnLines)
        CreateApproReturnAndInvoice = True
        Set oSALELine = oExchange.SaleLines.Add
        oSALELine.ApplyEdit
        oSALELine.BeginEdit
        iCurrentSaleLine = iCurrentSaleLine + 1
        X1.ReDim 1, iCurrentSaleLine, 1, 7
        oSALELine.IsDepositItem = True '''Does not napply discount again
        oSALELine.Price = FNN(arApproReturnLines(i).Price)
        oSALELine.Qty = arApproReturnLines(i).APPLQtySold
        oExchange.TotalPayable = pInvValue
        oSALELine.DiscountRate = FNN(arApproReturnLines(i).DiscountRate) '   oExchange.Customer.DefaultDiscount  '
        oSALELine.Title = arApproReturnLines(i).Title
        oSALELine.Code = arApproReturnLines(i).Code
        oSALELine.VATRATE = arApproReturnLines(i).VATRATE
        oSALELine.PID = arApproReturnLines(i).PID
        oSALELine.COLID = arApproReturnLines(i).APPLID
        oSALELine.CalculateLine
        DisplayProduct
    Next i
    End If
    
    
    
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_eStart", , EA_NORERAISE
    HandleError
End Sub
Private Function RefundPayment(ePaymentMode As enPaymentMode)

    On Error GoTo errHandler
    PreparePaymentLine ePaymentMode
    oExchange.CalculateTotals
    oPAYMENTLine.Amt = oExchange.BalanceOwing
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
            Action_PaymentType_Creditcard = ePaymentType_CreditCardRef
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
            Action_PaymentType_CreditcardRef = ePaymentType_CreditCardRef
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
        If InStr(1, strValidVoucherTypes, strPrefix) > 0 And Len(strRaw) > 1 Then 'valid voucher type
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
Private Function Action_PaymentType_CreditNote() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditNote = DetermineReturnToState
    Else
        PreparePaymentLine ePaymentMode_CreditNote
        If oPAYMENTLine.SetAmt(Trim(strSuffix)) Then
            oExchange.CalculateTotals
            Action_PaymentType_CreditNote = ePaymentType_CreditNoteRef
        Else
            SetTip "Invalid payment amount."
            Action_PaymentType_CreditNote = ePaymentType_CreditNote
        End If
        DisplayPayment
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditNote", , EA_NORERAISE
    HandleError
End Function
Private Function Action_PaymentType_CreditNoteRef() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_PaymentType_CreditNoteRef = DetermineReturnToState
    Else
        If oPAYMENTLine.SetReference(Trim(strRaw)) Then
            oExchange.CalculateTotals
            If oExchange.PaymentsComplete(, strMsg) Then
                Action_PaymentType_CreditNoteRef = eConfirmation
            Else
                Action_PaymentType_CreditNoteRef = eSale
            End If
            DisplayPayment
        Else
            SetTip "Invalid Reference."
            Action_PaymentType_CreditNoteRef = ePaymentType_CreditNoteRef
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_PaymentType_CreditNoteRef", , EA_NORERAISE
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
        If lngDeposit <= 0 Or lngDeposit > 100000 Or FNS(X3(lngTmp, 7)) <> "P" Then
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
        If GetLevel(txtInput, strName, lngOPID) > 0 Then
            oExchange.SalesPersonID = lngOPID
            If oExchange.IssueCreditNoteForChange(strMsg) = True Then
                Set frm = New frmChangeToGive
                frm.component strMsg
                frm.Show vbModal
                bIssueCreditNote = frm.IssueChangeAsCreditNote
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
        If SecurityControl(2, lngOPID, strName, , "Enter your security key.", "Your key is invalid") Then
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
        If SecurityControl(2, lngOPID, strName, , "Enter your security key.", "Your key is invalid") Then
            oExchange.SetExchangeType ePettyCashCreditType
            oExchange.Note = frmPCC.Reason
            Set oPAYMENTLine = oExchange.PaymentLines.Add
            oPAYMENTLine.ApplyEdit
            oPAYMENTLine.BeginEdit
            oPAYMENTLine.SetAmt CStr(frmPCC.Amount)
            oPAYMENTLine.SetType "R"
            AcceptSale
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
            strRequestDetails = frm.Customer & "|" & frm.Item
            lngDepositValue = frm.Deposit
            oExchange.Note = strRequestDetails & vbCrLf & "Deposit:" & CStr(lngDepositValue)
            oExchange.SetExchangeType eOrderRequestType
        
            Set oSALELine = oExchange.SaleLines.Add
            oSALELine.ApplyEdit
            oSALELine.BeginEdit
            iCurrentSaleLine = iCurrentSaleLine + 1
            X1.ReDim 1, iCurrentSaleLine, 1, 7
            oSALELine.PID = ""
            oSALELine.Price = lngDepositValue
            oSALELine.Title = frm.Item
            oSALELine.Code = ""
            oSALELine.SetQty 1
            oSALELine.IsDepositItem = True
            oSALELine.CalculateLine
            oExchange.CalculateTotals
        
            DisplayProduct
            DisplayTotals
            Action_OrderRequest = eCollect
        End If
        strOrderedTitle = frm.Item & vbCrLf & "For: " & frm.Customer
        Unload frm
    End If
End Function
Private Function Action_Sale() As eState
    On Error GoTo errHandler
    If strPrefix = ".." Then
        Action_Sale = eStart
    ElseIf LoadProductFromCode Then
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Sale", , EA_NORERAISE
    HandleError
End Function
Private Function Action_Appro() As eState
    On Error GoTo errHandler
    If LoadProductFromCode Then
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
    If strPrefix = ".." Then
        Action_Price = DetermineReturnToState
        RemoveSaleLine iCurrentSaleLine
        DisplayTotals
        If oExchange.LoyaltyValue > 0 Then
            lblCustomername.Caption = DisplayCustomerDetails
        End If
    ElseIf bShiftDown Then
        If oSALELine.SetPrice(Trim(txtInput)) Then
            oExchange.CalculateTotals
            DisplayProduct
            Action_Price = eDiscount
        Else
            SetTip "Invalid price."
        End If
    Else
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
    ErrorIn "frmPOSMain.Action_Price", , EA_NORERAISE
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
                txtInput = Right(strSuffix, Len(strSuffix) - 1)
            End If
        End If

        If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(strSuffix), bItemExchange) Then
            oExchange.CalculateTotals
            DisplayProduct
            If oExchange.transactionTypeenum = eSaleType Then
                Action_Qty = eSale
            ElseIf oExchange.transactionTypeenum = eApproType Then
                Action_Qty = eAppro
            End If
            oSALELine.ApplyEdit
            oSALELine.BeginEdit
        Else
            SetTip "Invalid quantity."
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

    If strPrefix = ".." Then
        Action_Discount = DetermineReturnToState
    Else
        strDiscountCode = UCase(Left(txtInput, 1))
        If InStr(1, strValidDiscountTypes, strDiscountCode) > 0 Then 'valid discount type
            If strDiscountCode = "X" Then
                ConnectionTimer.Enabled = False
                If SecurityControl(3, lngStaffID, strName, , "Enter security code to allow discretionary discount") Then
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
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Action_Discount", , EA_NORERAISE
    HandleError
End Function
Private Function Action_OpenDrawer() As eState
    On Error GoTo errHandler
    If SecurityControl(3, lngStaffID, strName, , "Enter security code to open drawer") Then
        OpenDrawer
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

    If IsNumeric(strSuffix) Then
        iToVoid = CLng(strSuffix)
        If iToVoid >= CLng(X4(1, 1)) And iToVoid < oExchange.ExchangeNumber Then
            If (X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 12) = 0) Then
                If Not oExchange.CANCANCEL(iToVoid) Then
              '  If X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 7) = "DEPOSIT" Then
                    MsgBox "This exhange cannot be voided.", vbInformation, "Can't do this"
                Else
                    Res = False
                    bCancelled = False
                    Do Until Res = True Or bCancelled = True
                        If Not SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
                            If bCancelled Then Res = False
                        Else
                            Res = True
                        End If
                    Loop
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
    If Not rs Is Nothing Then UpdatingLocalDatabase True, rs.RecordCount
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
        Case "ClearCustomers"
                oPC.DBLocalConn.Execute "Delete FROM tCustomer"
        Case "ClearProducts"
                oPC.DBLocalConn.Execute "Delete FROM tProduct"
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
    
    bUpdating = True
    rs.MoveFirst
    lngCnt = 1
    Do While Not rs.EOF
        lngCnt = lngCnt + 1
        If lngCnt Mod 100 = 0 Then
            Counter "Pr", lngCnt
        End If
        Set cmd = New ADODB.Command
        cmd.CommandText = "dbo.sp_InsertProductUpdateToFD "
        cmd.CommandType = adCmdStoredProc
        
        Set par = cmd.CreateParameter("@PID", adGUID, , , rs!PRU_P_ID)
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@CODE", adVarChar, adParamInput, 50, FNS(rs!PRU_Code))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@EAN", adVarChar, adParamInput, 50, FNS(rs!PRU_EAN))
        cmd.Parameters.Append par
        Set par = Nothing
        Set par = cmd.CreateParameter("@PUBLISHER", adVarChar, adParamInput, 50, FNS(rs!PRU_Publisher))
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
        
        cmd.ActiveConnection = oPC.DBLocalConn
        cmd.Execute
        
        Set cmd = Nothing
        rs.MoveNext
        DoEvents
    Loop

    bUpdating = False
    SaveProductUpdate = True
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SaveProductUpdate(rs)", rs
End Function

Private Function SaveStaffUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True

    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tStaffMembers WHERE tStaffMembers.SM_ID =" & rs!SMU_ID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        
        If NewRS.EOF Then
            NewRS.AddNew
        End If
        If Not IsNull(rs!SMU_ID) Then NewRS!SM_ID = rs!SMU_ID
        If Not IsNull(rs!SMU_NAME) Then NewRS!SM_Name = Trim$(rs!SMU_NAME)
        If Not IsNull(rs!SMU_Role) Then NewRS!SM_Role = Trim$(rs!SMU_Role)
        If Not IsNull(rs!SMU_Telephone) Then NewRS!SM_Telephone = Trim$(rs!SMU_Telephone)
        If Not IsNull(rs!SMU_Mobile) Then NewRS!SM_Mobile = Trim$(rs!SMU_Mobile)
        If Not IsNull(rs!SMU_Password) Then NewRS!SM_Password = Trim$(rs!SMU_Password)
        If Not IsNull(rs!SMU_Level) Then NewRS!SM_Level = rs!SMU_Level
        If Not IsNull(rs!SMU_Shortname) Then NewRS!SM_Shortname = rs!SMU_Shortname
        sName = Trim$(rs!SMU_NAME)
        i = 1
DoUpdate:
        NewRS.Update
        If Err = -2147217887 Then
          NewRS!SM_Code = Left(NewRS!SM_Code, 3) & CStr(i)
          i = i + 1
          Err.Clear
          GoTo DoUpdate
        ElseIf Err <> 0 Then
          GoTo errHandler
        End If
        NewRS.Close
        DoEvents
        rs.MoveNext
    Loop
 '   oPC.DBConn.CommitTrans
    SaveStaffUpdate = True
        bUpdating = False

MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SaveStaffUpdate(rs)", rs
End Function

Private Function SaveCustomerUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
Dim lngCnt As Long

    bUpdating = True

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
            NewRS!C_Name = FNS((rs!CU_Name))
            NewRS!C_DefaultDiscount = FNN(rs!CU_DefaultDiscount)
            NewRS!C_Acno = FNS(rs!CU_Acno)
            NewRS!C_Initials = Left(FNS(rs!CU_Initials), 8)
            NewRS!C_Title = FNS(rs!CU_Title)
            NewRS!C_Phone = FNS(rs!CU_Phone)
            NewRS!C_VATABLE = FNN(rs!CU_VATABLE)
            NewRS!C_BALANCE = FNN(rs!CU_BALANCE)
            NewRS!C_Type = FNS(rs!CU_TYPE)
        End If
        NewRS.Update
        NewRS.Close
        DoEvents
        rs.MoveNext
    Loop
    SaveCustomerUpdate = True
        bUpdating = False
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SaveCustomerUpdate(rs)", rs
End Function

Private Function SaveCustomerOrderUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True
    
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
        DoEvents
        rs.MoveNext
    Loop
    
    SaveCustomerOrderUpdate = True
        bUpdating = False
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SaveCustomerOrderUpdate(rs)", rs
End Function
Private Function SaveApproUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer
    bUpdating = True

    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tAPPL WHERE APPL_APPLID = " & rs!APPL_APPLID
        NewRS.LockType = adLockOptimistic
        NewRS.CursorType = adOpenDynamic
        Set NewRS.ActiveConnection = oPC.DBLocalConn
        NewRS.Open sSQL  ', adOpenDynamic, adLockPessimistic
        
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
        DoEvents
        rs.MoveNext
    Loop
    
    SaveApproUpdate = True
    bUpdating = False
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SaveApproUpdate(rs)", rs
End Function

Private Function SaveMarketingUpdate(rs As ADODB.Recordset) As Boolean
    On Error GoTo errHandler
Dim NewRS As New ADODB.Recordset
Dim sSQL As String
Dim sName As String
Dim i As Integer

    bUpdating = True

    rs.MoveFirst
    Do While Not rs.EOF
        sSQL = "SELECT * FROM tMarketing WHERE M_PT_ID = " & rs!MC_PT_ID & " AND M_SECTION_ID = " & rs!MC_SECTION_ID
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
            NewRS!M_LOYALTYDISCOUNT = FNN(rs!MC_LOYALTYDISCOUNT)
            NewRS!M_DISCOUNT = FND(rs!MC_DISCOUNT)
    
            NewRS!M_DESCRIPTION = FNS(rs!MC_DESCRIPTION)
            NewRS!M_NODISCOUNTALLOWABLE = FNS(rs!MC_NODISCOUNTALLOWABLE)
            NewRS!M_IDENTIFYCUSTOMER = FNN(rs!MC_IDENTIFYCUSTOMER)
            NewRS!M_ACTIVE = FNN(rs!MC_ACTIVE)
DoUpdate:
            NewRS.Update
            NewRS.Close
        End If
        DoEvents
        rs.MoveNext
    Loop
    SaveMarketingUpdate = True
    bUpdating = False
    
MEX:
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set NewRS = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SaveMarketingUpdate(rs)", rs
End Function




