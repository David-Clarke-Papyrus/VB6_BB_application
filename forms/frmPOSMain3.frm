VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{DA4E6F7B-F5EE-43C5-A9A1-6BCC726F901E}#1.8#0"; "StatusBarX5.OCX"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr17.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPOSMain 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DiscountSet"
   ClientHeight    =   8115
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "frmPOSMain3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPettyCash 
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
      Left            =   5835
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "frmPOSMain3.frx":08CA
      Top             =   4485
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H80000000&
      Caption         =   "&Open"
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
      Left            =   10650
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7140
      Width           =   885
   End
   Begin VB.Timer ConnectionTimer 
      Interval        =   10000
      Left            =   225
      Top             =   5250
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
      Left            =   5805
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frmPOSMain3.frx":08D0
      Top             =   4470
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
      TabIndex        =   16
      Text            =   "frmPOSMain3.frx":08D6
      Top             =   4530
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
   Begin TrueOleDBGrid60.TDBGrid G3 
      Height          =   2040
      Left            =   5865
      OleObjectBlob   =   "frmPOSMain3.frx":08DC
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   5670
   End
   Begin TrueOleDBGrid60.TDBGrid G4 
      Height          =   2040
      Left            =   90
      OleObjectBlob   =   "frmPOSMain3.frx":5DF7
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   11385
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000000&
      Caption         =   "&Help"
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
      Left            =   9225
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6915
      Visible         =   0   'False
      Width           =   885
   End
   Begin StatusBarXCtl.StatusBarX SB 
      Height          =   870
      Left            =   150
      Top             =   7200
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   1535
      Appearance      =   0
      BorderColor     =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483648
      HighlightColor  =   -2147483632
      HighlightDkColor=   -2147483635
      PanelCount      =   1
      Panel1Key       =   "test"
      Panel1ForeColor =   7884871
      Panel1WordWrap  =   -1  'True
      Panel1Width     =   687
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
      Left            =   255
      TabIndex        =   1
      Top             =   6330
      Width           =   5385
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2880
      Left            =   135
      OleObjectBlob   =   "frmPOSMain3.frx":B936
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   11340
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Height          =   1380
      Left            =   165
      OleObjectBlob   =   "frmPOSMain3.frx":1037D
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3705
      Width           =   3300
   End
   Begin VB.Label lblOnlineStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Online"
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
      Height          =   300
      Left            =   10605
      TabIndex        =   18
      Top             =   7680
      Width           =   915
   End
   Begin VB.Label lblChange 
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
      Height          =   615
      Left            =   5700
      TabIndex        =   15
      Top             =   6495
      Width           =   4575
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
      Left            =   8685
      TabIndex        =   14
      Top             =   6330
      Width           =   2910
   End
   Begin VB.Label txtPaymentTotal 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   13
      Top             =   5100
      Width           =   2700
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      Height          =   960
      Left            =   3705
      TabIndex        =   9
      Top             =   3705
      Width           =   6210
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   3
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
      TabIndex        =   8
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
      Left            =   270
      TabIndex        =   7
      Top             =   5535
      Visible         =   0   'False
      Width           =   4860
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   3330
      Top             =   3615
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
      Top             =   5940
      Width           =   5400
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "App Setup"
      End
      Begin VB.Menu Line01 
         Caption         =   "-"
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
Public Enum enumDocumentType
    eTypReceipt = 1
    eTypVoucher = 2
    eTypCashRefund = 3
    etypCreditNote = 4
    eTypDeposit = 5
    etypDepositRefund = 6
    eTypAppro = 7
    eTypPettyCash = 8
End Enum
Dim lngPettyCashLine As Long
Dim arLineNumber() As String
Dim bShiftDown As Boolean
Dim flgUnloading As Boolean
Dim lngAmt As Long
Dim arDiscounts() As String
Dim arPettyCash() As String
Dim strValidVoucherTypes As String
Dim strValidDiscountTypes As String
Dim strPettyCashTypes As String
Dim WithEvents oExchange As a_Exchange
Attribute oExchange.VB_VarHelpID = -1
Dim oPAYMENTLine As a_Payment
Dim bValid As Boolean
Public WithEvents oPS As z_PollingServices_Client
Attribute oPS.VB_VarHelpID = -1
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2
Dim cCOLS As C_COLS
Dim oSALELine As a_Sale
Dim oCurrLine As ListItem
Dim ADOConn As ADODB.Connection
Dim frmExchange As frmExchange
Dim strDepositTitle As String
Dim bWaitForClearance As Boolean
Dim bEnvironmentOK As Boolean
Dim Result As String
Dim enRequestState As eState
Dim lngDeposit As Long
Dim iCOLForDeposit As Long
Dim ESC As String
Dim enPresentState As eState
Dim flgSaleActive As Boolean
Dim flgCustomerVisible As Boolean
Dim flgGDiscount As Boolean
Dim flgNewCode As Boolean
Dim flgEditItem As Boolean
Dim flgReturn As Boolean
Dim flgInvalidLine As Boolean
Dim flgVoidAndReplace As Boolean
Dim iCurLine As Integer
Dim flgLoading As Boolean
Dim bLoggedOn As Boolean
Dim sOldStat As String
Dim sOldCode As String
Dim X1 As New XArrayDB
Dim X2 As New XArrayDB
Dim X3 As New XArrayDB
Dim X4 As New XArrayDB
Dim bONLINE As Boolean
Dim strOpSessionID As String
Dim strSessionID As String
Dim sBill As String
Dim strCustomername As String
Dim lngCustomerID As Long
Dim sPaymentType As String
Dim iCurrentSaleLine As Integer
Dim iCurrentPaymentLine As Integer
Dim iCurrentCOL As Integer
Dim strName As String
Dim lngStaffID As Long
Dim strOperator As String
Dim bConnected As Boolean
Dim bCloseXsession As Boolean
Dim bCloseZsession As Boolean
Dim lngSMID As Long
Dim lngSalesItemCount As Long
Dim iToVoid As Long
Dim lngBalanceOwing As Long
Dim bLoyaltyCustomer As Boolean
Dim itmp As Integer
Dim strArg As String
Dim strArg2 As String
Dim frmDisc As frmDiscretionaryDiscount
Dim frmCustID As frmIDCustomer
Dim frmPC As frmPettyCash

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
    GetEnvironmentstatus = bEnvironmentOK
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







Private Sub cmdOpen_Click()
    OpenDrawer
End Sub

Private Sub oPS_UpdatingLocalDatabase(bOn As Boolean)
    On Error GoTo errHandler
Static strMsg As String
    If flgUnloading Then Exit Sub
    If bOn Then
        strMsg = SB.Panels(1).Text
        SB.Panels(1).Text = "Updating local database . . ."
    Else
        SB.Panels(1).Text = strMsg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_UpdatingLocalDatabase(bOn)", bOn
End Sub

Private Sub oPS_NoServer(MSG As String)
    On Error GoTo errHandler
    If flgUnloading Then Exit Sub
    bConnected = False
    lblOnlineStatus.Caption = "Offline"
    lblOnlineStatus.ForeColor = vbRed
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_LostServer(MSG)", MSG
End Sub

Private Sub SetPresentState(val As eState)
    On Error GoTo errHandler
    enPresentState = val
    Me.lblState.Caption = InterpretState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetPresentState(val)", val
End Sub
Private Function InterpretState() As String
    On Error GoTo errHandler
    Select Case enPresentState
    Case 0
        InterpretState = "eProductID"
    Case 1
        InterpretState = "eSearchProduct"
    Case 2
        InterpretState = "eTitle"
    Case 3
        InterpretState = "eAuthor"
    Case 4
        InterpretState = "eQty"
    Case 5
        InterpretState = "eDiscount"
    Case 6
        InterpretState = "ePrice"
    Case 7
        InterpretState = "elogin"
    Case 8
        InterpretState = "ePaymentType_Cash"
    Case 9
        InterpretState = "ePaymentAmt"
    Case 10
        InterpretState = "eConfirmation"
    Case 11
        InterpretState = "ePrevious"
    Case 12
        InterpretState = "eDelete"
    Case 13
        InterpretState = "ePaymentType_Cheque"
    Case 14
        InterpretState = "ePaymentType_CreditCard"
    Case 15
        InterpretState = "ePaymentType_voucher"
    Case 16
        InterpretState = "eXTerminate"
    Case 17
        InterpretState = "ePaymentType_ChequeRef"
    Case 18
        InterpretState = "ePaymentType_CreditCardRef"
    Case 19
        InterpretState = "ePaymentType_voucherRef"
    Case 20
        InterpretState = "eDeletePayment"
    Case 21
        InterpretState = "eSearchCustomer"
    Case 22
        InterpretState = "eRebuildIndexes"
    Case 23
        InterpretState = "ePaymentType_CreditNote"
    Case 24
        InterpretState = "eHelp"
    Case 25
        InterpretState = "eCancelSale"
    Case 26
        InterpretState = "ePaymentType_RedeemDeposit"
    Case 27
        InterpretState = "eCashRefund"
    Case 28
        InterpretState = "ePriceCashRefund"
    Case 29
        InterpretState = "eQtyCashRefund"
    Case 30
        InterpretState = "eDiscountCashRefund"
    Case 31
        InterpretState = "eConfirmationCashrefund"
    Case 32
        InterpretState = "eVoidandReplace"
    Case 33
        InterpretState = "eReviewExchanges"
    Case 34
        InterpretState = "eShowExchange"
    Case 35
        InterpretState = "eCreditNote"
    Case 36
        InterpretState = "eAcceptDeposit"
    Case 37
        InterpretState = "eDiscountCreditNote"
    Case 38
        InterpretState = "eConfirmationCreditNote"
    Case 39
        InterpretState = "eConfirmationDeposit"
    Case 40
        InterpretState = "ePriceCreditNote"
    Case 41
        InterpretState = "ePriceDeposit"
    Case 42
        InterpretState = "eQtyCreditNote"
    Case 43
        InterpretState = "eQtyDeposit"
    Case 44
        InterpretState = "eDiscountDeposit"
    Case 45
        InterpretState = "eAcceptDepositRef"
    Case 46
        InterpretState = "esearchcustomerfordeposit"
    Case 55
        InterpretState = "eAppro"
    Case 56
        InterpretState = "ePriceAppro"
    Case 57
        InterpretState = "eQtyAppro"
    Case 58
        InterpretState = "eConfirmationAppro"
    Case 59
        InterpretState = "eDiscountAppro"
    Case 60
        InterpretState = "eSearchCustomerforAppro"
    Case 61
        InterpretState = "ePettyCash"
    Case 62
        InterpretState = "ePettyCashAmt"
    Case 63
        InterpretState = "ePettyCashConfirmation"

    Case 99
        InterpretState = "enull"
    End Select

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.InterpretState"
End Function
'Private Sub Command1_Click()
'Dim result As Integer
'
''    result = IO1.Close
'    result = IO1.Open(oPC.cashDrawerPort, "baud=9600 parity=N data=8 stop=1")  'Set up scanner
'
'    result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
'
'    result = IO1.Close
'
'End Sub


Private Sub oPS_HasServer()
    On Error GoTo errHandler
    If flgUnloading Then Exit Sub
    bConnected = True
    lblOnlineStatus.Caption = "Online"
    lblOnlineStatus.ForeColor = &H80000012
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_HasServer"
End Sub


Private Sub oExchange_ContainsLines(pYesNo As Boolean)
    On Error GoTo errHandler
    If flgUnloading Then Exit Sub
    flgSaleActive = pYesNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_ContainsLines(pYesNo)", pYesNo
End Sub
Private Sub SetTitleBar(pShowExchangeNumber As Boolean)
    On Error GoTo errHandler
    Caption = "Papyrus Point-of-Sale       " & oPC.StationName & "      Supervisor: " & oPC.ZSession.SupervisorName & "/" & oPC.ZSession.Opsession.Name & IIf(pShowExchangeNumber = True, "              #" & oExchange.ExchangeNumber, "")
  '  lblStatus.Caption = "Sales for " & oPC.ZSession.NominalDateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetTitleBar(pShowExchangeNumber)", pShowExchangeNumber
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Dim frm As frmPOSHELP
    Set frm = New frmPOSHELP
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdClose_Click"
End Sub

Private Sub ConnectionTimer_Timer()
    If bConnected = False Then
        oPS.TryToStartPolling
    End If
    
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim Result As Integer
Dim lngResult As Long
Dim bLoggedOnAlready As Boolean
    flgSaleActive = False
    bEnvironmentOK = True
    ESC = Chr(27)
    iToVoid = 0
    
    'Try to load local DB connection
    If oPC Is Nothing Then
        Set oPC = New z_POSCLIConnection
        oPC.dbConnect
    End If
    
    If oPC.PrintSlips = True Then
        With OPOSPOSPrinter1
            lngResult = .Open(oPC.printername)
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
        Me.lblState.Visible = False
    End If
    
    

    Set oPS = New z_PollingServices_Client
    Check oPS.TryToStartPolling, EXC_SERVERUNAVAILABLE, "Cannot poll server"
    
    bLoggedOnAlready = False
    oPC.SetupZSession lngStaffID, strName, oPS.ClientOutbox
    If oPC.ZSession.supervisorID = 0 Then
        LogonOperator
        oPC.ZSession.supervisorID = lngStaffID
        oPC.ZSession.SupervisorName = strName
        bLoggedOnAlready = True
    End If
    If oPC.ZSession.loadOpenXSession = False Then
        oPC.ZSession.Opsession.START_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
        If oPC.ZSession.Opsession.supervisorID = 0 Then
            If bLoggedOnAlready = False Then
                LogonOperator
            End If
            oPC.ZSession.Opsession.OperatorID = lngStaffID
            oPC.ZSession.Opsession.Name = strName
        End If
    End If

    SetForCOLSVisible False
    If oPC.DriveDrawer = True Then
        MSComm1.Settings = oPC.COMPORTSettings
        MSComm1.CommPort = oPC.CashDrawerPort
        If MSComm1.PortOpen = False Then
            MSComm1.PortOpen = True
        End If
    End If
    
    LoadVoucherTypes
    LoadDiscountTypes
    LoadPettyCashTypes

    G4.Height = 380
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Load"
End Sub

Private Sub LoadVoucherTypes()
Dim ar() As String
Dim i As Integer
    ar = Split(oPC.VoucherSet, ";")
    
    strValidVoucherTypes = ""
    For i = 0 To UBound(ar)
        strValidVoucherTypes = strValidVoucherTypes & Left(ar(i), 1)
    Next
    
    txtVouchers = Replace(oPC.VoucherSet, ";", vbCrLf)
    
End Sub
Private Sub LoadDiscountTypes()
Dim i As Integer
    arDiscounts = Split(oPC.DiscountSet, ";")
    
    strValidDiscountTypes = ""
    For i = 0 To UBound(arDiscounts)
        strValidDiscountTypes = strValidDiscountTypes & Left(arDiscounts(i), 1)
    Next
    
    txtDiscounts = Replace(oPC.DiscountSet, ";", vbCrLf)
    
End Sub
Private Sub LoadPettyCashTypes()
Dim i As Integer
    arPettyCash = Split(oPC.PettyCashSet, ";")
    
    strPettyCashTypes = ""
    For i = 0 To UBound(arPettyCash)
        strPettyCashTypes = strPettyCashTypes & Left(arPettyCash(i), 1)
    Next
    
    txtPettyCash = Replace(oPC.PettyCashSet, ";", vbCrLf)
    
End Sub

Private Function GetDiscount(pCODE As String, pDescription) As Integer
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
End Function
Private Function LogonOperator() As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
        strOperator = strName
        bLoggedOn = True
    Else
       ' LockAll True
        SetPresentState elogin
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LogonOperator"
End Function

Private Function SwapOperator() As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean

    If oPC.ZSession.Opsession.InSession Then
        oPC.ZSession.Opsession.Close_OP_Session
    End If
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
        oPC.ZSession.Opsession.START_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
        strOperator = strName
        bLoggedOn = True
    Else
      '  LockAll True
        SetPresentState elogin
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SwapOperator"
End Function

Private Sub cmdZTotal_Click()
    On Error GoTo errHandler
Dim sPass As String
Dim frm As frmSecurity
Dim lngStaffID As Long
Dim strName As String

    If SecurityControl(4, lngStaffID, strName, , "Enter security code to close session") Then
        If oPC.ZSession.Opsession.InSession Then
            oPC.ZSession.Opsession.Close_OP_Session
        End If
        If oPC.ZSession.InSession Then
            oPC.ZSession.Close_Z_Session
        End If
        Unload Me
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdZTotal_Click"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Unload Me
        Else
            KeyCode = 0
        End If
    ElseIf Shift <> 0 Then
            KeyCode = 0
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_KeyPress(KeyAscii)", KeyAscii
End Sub



Private Sub StandbyMode()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtInput.Enabled = False
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StandbyMode"
End Sub

Public Sub StartSale()
    On Error GoTo errHandler
    Set oExchange = New a_Exchange
    oExchange.beginedit
    oExchange.SetExchangeType eSaleType
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    SetTitleBar False
    X4.Clear
    X4.ReDim 1, 1, 1, 13
    LoadExchanges

    SetPresentState eProductID
    SetMessages
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StartSale"
End Sub



Private Sub Stat(MSG As String)
    On Error GoTo errHandler
    SB.Panels(1).Text = MSG
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Stat(Msg)", MSG
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If bEnvironmentOK = True Then
        flgUnloading = True
        ConnectionTimer.Enabled = False
        CloseApplication Cancel
    End If
    Set oPS = Nothing
End Sub

Private Sub CloseApplication(bCancel As Integer)
    If flgSaleActive Then
        If MsgBox("There is still a transaction in process!" & vbLf & _
                  "Do you want to close this application anyway?", _
                  vbYesNo, "Transaction In Process!") = vbNo Then
            bCancel = True
            Exit Sub
        Else
            RejectSale
        End If
    Else
        If MsgBox("Closing Papyrus POS application?", vbYesNo + vbQuestion, "Close?") = vbNo Then
            bCancel = True
            Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    Me.SB.Panels(1).Text = "Wait. The local data is being transmitted to the server."
    
    oPS.RegisterWithServer (False)
    
    If oExchange.IsEditing Then oExchange.CancelEdit
    
    If bCloseXsession Then
        oPC.ZSession.Opsession.Close_OP_Session
    End If
    
    If bCloseZsession Then
        oPC.ZSession.Close_Z_Session
    End If
    
    With OPOSPOSPrinter1
        .DeviceEnabled = False
        .ReleaseDevice
        .Close
    End With
    
    If MSComm1.PortOpen = True Then
       MSComm1.PortOpen = False
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
End Sub


Private Sub mnuClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.mnuClose_Click"
End Sub


Private Sub oPS_PollingStoped(MSG As String)
    On Error GoTo errHandler
    If MsgBox("Automatic file transfer stopped!" & vbLf & _
               "Reason: " & MSG & vbLf & vbLf & _
               "Click YES to restart it.", vbYesNo + vbExclamation, "File Transfer Stopped!") = vbYes Then
        
        oPS.StartPolling
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_PollingStoped(Msg)", MSG
End Sub

Private Sub txtInput_GotFocus()
    On Error GoTo errHandler
   ' AutoSelect Controls("txtInput")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_GotFocus"
End Sub
Private Sub ShowExchange()
    On Error GoTo errHandler
Dim lngRow As Long
Dim lngTmp As Long

    Set frmExchange = New frmExchange
    If IsNumeric(txtInput) Then
        lngRow = CLng(txtInput)
        If lngRow <= X4(X4.UpperBound(1) - 1, 1) And lngRow > 0 Then
            lngTmp = X4.Find(1, 1, lngRow, , , XTYPE_LONG)
            If lngTmp > 0 Then
                frmExchange.component X4(lngTmp, 10)
                frmExchange.Show vbModal
                If frmExchange.MustPrint = True Then
                    Set oExchange = Nothing
                    Set oExchange = New a_Exchange
                    oExchange.Load (X4(lngTmp, 10)), True
                    PrintSalesSlip 1, True
                    Set oExchange = Nothing
                    Set oExchange = New a_Exchange
                    oExchange.beginedit
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
Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler

  '  txtInput = UCase(Trim(txtInput))
    bShiftDown = (Shift = 1)
    If KeyCode = 13 Then
        If bWaitForClearance Then
            lblChange.Caption = ""
            bWaitForClearance = False
            Exit Sub
        End If
        txtInput = UCase(Trim(txtInput))
        InterpretInput
        If bValid = True Then
            Statechange enRequestState, itmp, strArg, strArg2
        End If
     End If
       
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift)
End Sub

Private Sub InterpretInput()
    On Error GoTo errHandler
Dim sTmp As String
Dim i As Integer
Dim lngRow As Long
Dim strTmp As String
Dim lngTmp As Long
Dim oSM As New z_SM

        If FNS(txtInput) = "" Then
            bValid = False
            Exit Sub
        End If
        bValid = True
        enRequestState = enull
        If Left(Trim(txtInput), 6) = "Change" Then
            txtInput = ""
            bValid = False
            Exit Sub
        End If
        
        If flgSaleActive Then
            If UCase(txtInput) = "C" And oExchange.SaleLines.Count > 0 Then
                If enPresentState = eCashRefund Then
                    enRequestState = eRefundType_Cash
                Else
                    enRequestState = ePaymentType_Cash
                End If
            End If
        
            If UCase(txtInput) = "X" Then
                enRequestState = ecancelsale
            End If
        
            If UCase(txtInput) = "F" Then
                If (enPresentState = eCashRefund) And oExchange.SaleLines.Count > 0 Then
                    enRequestState = eConfirmationCashrefund
                ElseIf (enPresentState = eCreditNote) And oExchange.SaleLines.Count > 0 Then
                    enRequestState = eConfirmationCreditNote
                ElseIf (enPresentState = eAcceptDeposit) And oExchange.SaleLines.Count > 0 Then
                    enRequestState = eConfirmationDeposit
                ElseIf (enPresentState = eProductID) And oExchange.SaleLines.Count > 0 Then
                    enRequestState = eConfirmation
                ElseIf (enPresentState = eAppro) And oExchange.SaleLines.Count > 0 Then
                    enRequestState = eConfirmationAppro
                End If
            End If
            
            If bShiftDown = True Then
                If enPresentState = ePriceCashRefund Then
                    enRequestState = eDiscountCashRefund
                ElseIf enPresentState = ePriceCreditNote Then
                    enRequestState = eDiscountCreditNote
                ElseIf enPresentState = ePrice Then
                    enRequestState = eDiscount
                ElseIf enPresentState = ePriceAppro Then
                    enRequestState = eDiscountAppro
                End If
            End If
        
            If UCase(Left(txtInput, 2)) = "FC" Then
                enRequestState = eSearchCustomer
            End If
        
        Else
            If UCase(txtInput) = "OD" Then
                If SecurityControl(3, lngSMID, strName, , "Enter your security key.", "Your key is invalid") Then
                    OpenDrawer
                    bValid = False
                End If
            End If
            If Left(UCase(txtInput), 2) = "PC" Then
                If Len(txtInput) > 2 Then
                    If IsNumeric(Right(txtInput, Len(txtInput) - 2)) Then
                        lngPettyCashLine = CLng(Right(txtInput, Len(txtInput) - 2))
                        MsgBox "check that line is PC"
                            enRequestState = ePettyCashCredit
                    Else
                        MsgBox "Invalid line number for petty cash credit"
                    End If
                Else
                    enRequestState = ePettyCash
                End If
            End If
            If Left(UCase(txtInput), 2) = "VR" Then
                strTmp = Right(txtInput, Len(txtInput) - 2)
                If IsNumeric(strTmp) Then
                    iToVoid = CLng(strTmp)
                    If iToVoid >= CLng(X4(1, 1)) And iToVoid < oExchange.ExchangeNumber Then
                        If (X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 12) = 0) Then
                            If X4(X4.Find(1, 1, iToVoid, XORDER_ASCEND, XCOMP_EQ, XTYPE_LONG), 7) = "DEPOSIT" Then
                                MsgBox "Only sales transactions can be voided and replaced.", vbInformation, "Can't do this"
                            Else
                                enRequestState = eVoidandReplace
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
            End If
        

            If UCase(txtInput) = "L" Then
                enRequestState = elogin
            End If
            If UCase(txtInput) = "R" Then
                enRequestState = eCashRefund
            End If
            If UCase(txtInput) = "A" Then
                enRequestState = eSearchCustomerforAppro
            End If
            If UCase(txtInput) = "DEP" And Not flgSaleActive Then
                enRequestState = eSearchCustomerfordeposit
            End If
            If UCase(txtInput) = "RDEP" And Not flgSaleActive Then
                enRequestState = eSearchCustomerfordepositRefund
            End If
            If UCase(txtInput) = "CN" Then
                If flgSaleActive = True Then
                    enRequestState = ePaymentType_CreditNote
                Else
                    enRequestState = eCreditNote
                End If
            End If
            
            If UCase(txtInput) = "X" Then
                enRequestState = eXTerminate
            End If
            
            If UCase(txtInput) = "Z" Then
                enRequestState = eZTerminate
            End If
            
        End If
        
        
'        If UCase(txtInput) = "S" Then
'            enRequestState = eProductID
'        End If
'        If UCase(txtInput) = "H" Then
'            enRequestState = eHelp
'        End If
'        If UCase(txtInput) = "RBI" Then
'            enRequestState = eRebuildIndexes
'        End If
        If UCase(txtInput) = ".." Then
            Select Case enPresentState
            Case eProductID  'We may have set to V&R or some kind of return and .. should take us back to a default sale
                enRequestState = eProductID
            Case Else
                enRequestState = ePrevious
            End Select
        End If
        If UCase(txtInput) = "Q" And oExchange.SaleLines.Count > 0 Then
            enRequestState = ePaymentType_Cheque
        End If
        If UCase(txtInput) = "V" And oExchange.SaleLines.Count > 0 Then
            enRequestState = ePaymentType_voucher
        End If
        If UCase(txtInput) = "CC" And oExchange.SaleLines.Count > 0 Then
            If enPresentState = eCashRefund Then
                enRequestState = eRefundType_Creditcard
            Else
                enRequestState = ePaymentType_CreditCard
            End If
        End If
        If Left(UCase(txtInput), 2) = "RD" And Len(txtInput) > 2 And oExchange.SaleLines.Count > 0 And flgCustomerVisible = True Then
            If Len(Trim(txtInput)) - 2 > "" Then
                If IsNumeric(Right(Trim(txtInput), Len(Trim(txtInput)) - 2)) Then
                    iCOLForDeposit = CInt(Right(Trim(txtInput), Len(Trim(txtInput)) - 2))
                    If X3.UpperBound(1) >= iCOLForDeposit And X3.LowerBound(1) <= iCOLForDeposit And (X3(iCOLForDeposit, 7) = "P" Or oSM.CanRedeemDeposit(X3(iCOLForDeposit, 11), iToVoid)) Then
                        enRequestState = ePaymentType_RedeemDeposit
                    End If
                End If
            End If
        End If
        
        If UCase(txtInput) = "DD" Then
            enRequestState = eReviewExchanges
        ElseIf UCase(Left(txtInput, 2)) = "DP" Then
            enRequestState = eDeletePayment
            itmp = CInt(Right(Trim(txtInput), Len(Trim(txtInput)) - 2))
        ElseIf UCase(Left(txtInput, 1)) = "D" Then
            If Len(txtInput) > 1 Then
                If IsNumeric(Right(txtInput, Len(txtInput) - 1)) Then
                    enRequestState = eDelete
                    If Len(txtInput) > 1 Then
                        itmp = CInt(Right(Trim(txtInput), Len(Trim(txtInput)) - 1))
                    End If
                End If
            End If
        End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.InterpretInput"
End Sub
Private Sub SetMessages()
    On Error GoTo errHandler
    txtVouchers.Visible = False
    txtDiscounts.Visible = False
    txtPettyCash.Visible = False
    Select Case enPresentState
        Case ePettyCashCredit
            setInputBox "", "", "", False
            lblInput.Caption = "Pett cash return amount"
            Stat ".. Return"
        Case ePettyCashCreditAmt
            setInputBox "", "", "", False
            lblInput.Caption = "Petty cash amount returned"
            Stat ".. Return"
        Case ePettyCashCreditConfirmation
            setInputBox "OK", "*", "", True
            lblInput.Caption = "Confirm petty cash return"
            Stat ".. Return"
        Case ePettyCash
            setInputBox "", "", "", False
            txtPettyCash = Replace(oPC.PettyCashSet, ";", vbCrLf)
            txtPettyCash.Visible = True
            lblInput.Caption = "Select petty cash account"
            Stat ".. Return"
        Case ePettyCashAmt
            setInputBox "", "", "", False
            lblInput.Caption = "Petty cash amount"
            Stat ".. Return"
        Case ePettyCashReason
            setInputBox "", "", "", False
            lblInput.Caption = "Reason"
            Stat ".. Return"
        Case ePettyCashConfirmation
            setInputBox "OK", "*", "", True
            lblInput.Caption = "Confirm petty cash withdrawal"
            Stat ".. Return"
        Case eAcceptDeposit
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            lblInput.Caption = "Select order line number from list "
            Stat "'.. to reverse"
        Case eRefundDeposit
            lblInput.ForeColor = vbBlue
            txtInput.ForeColor = vbBlue
            lblInput.Caption = "Select order line number from list "
            Stat "'.. to reverse"
        Case eDepositMode
            setInputBox "", "", "", False
            lblInput.Caption = "Select payment type "
            Stat "(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque"
        Case eShowvoucherType
            lblInput.Caption = "Select voucher type "
            txtVouchers = Replace(oPC.VoucherSet, ";", vbCrLf)
            txtVouchers.Visible = True
            Stat "  .. to reverse"
        Case ecancelsale
            setInputBox "", "", "", True
        Case eCashRefund
            ClearTextFields
            setInputBox "", "", "", False
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            If flgSaleActive Then
                lblInput.Caption = "Product code."
                If flgCustomerVisible = True Then
                    Stat "Scan or (A) Appro,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(RDn)Redeem deposit,(Dn)Del prod,(DPn)Del paymt"
                Else
                    Stat "Scan or (C)Cash,(CC)Card, (X) Cancel transaction, (Dn)Del prod"
                End If
            Else
                lblInput.Caption = "Start cash refund "
                Stat "Start cash refund by entering product code,   .. to reverse"
            End If
        Case eAppro
            ClearTextFields
            setInputBox "", "", "", False
            lblInput.ForeColor = vbBlue
            txtInput.ForeColor = vbBlue
            If flgSaleActive Then
                lblInput.Caption = "Product code."
             '   If flgCustomerVisible = True Then
             '       Stat "Scan or (A)Appro,(Dn)Del prod,(DPn)Del paymt"
             '   Else
                    Stat "Scan or (F)Finalize,(X)Cancel transaction,(Dn)Del prod,(DPn)Del paymt"
             '   End If
            Else
                lblInput.Caption = "Start Appro "
                Stat "Start Appro by entering product code,   .. to reverse"
            End If
        Case eConfirmationDeposit
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm deposit payment"
            setInputBox "OK", "*", "", True
        Case eConfirmationAppro
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm Appro"
            setInputBox "OK", "*", "", True
        Case eConfirmationRefundDeposit
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm deposit refund"
            setInputBox "OK", "*", "", True
        Case eConfirmation
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm sale"
            setInputBox "OK", "*", "CHNG: " & oExchange.ChangeGivenF, True
        Case eConfirmationCashrefund
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm cash refund"
            setInputBox "OK", "*", "", True
            AutoSelect txtInput
        Case eConfirmationCreditNote
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm credit note"
            setInputBox "OK", "*", "", True
            AutoSelect txtInput
        Case eConfirmationDeposit
            Stat "'.. to reverse"
            lblInput.Caption = "Confirm deposit"
            setInputBox "OK", "*", "", True
            AutoSelect txtInput
        Case eCreditNote
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            setInputBox "", "", "", True
            If flgSaleActive Then
                lblInput.Caption = "Product code, payment code or special."
                If flgCustomerVisible = True Then
                    Stat "Scan or THis must be wrong (A)Appro,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(Dn)Del prod,(DPn)Del paymt"
                Else
                    Stat "Scan or (Dn)Del prod,(F)Finalize,(X)Cancel transaction"
                End If
            Else
                lblInput.Caption = "Start credit note "
                Stat "Start credit note by entering product code,   .. to reverse"
            End If
        
        Case eDiscount, eDiscountCashRefund, eDiscountCreditNote, eDiscountAppro
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
        Case ePriceCashRefund, ePriceCreditNote, ePriceAppro
            lblInput.Caption = "Price "
            Stat "Hold shift key down and press Enter for discount, '..' to reverse"
            setInputBox oSALELine.Price, "", "", True
        Case eProductID
            setInputBox "", "", "", True
            ShowTransactions False
            
            lblInput.ForeColor = &H714942
            txtInput.ForeColor = &H714942
            
            If flgSaleActive Then
                lblInput.Caption = "Product code, payment code or special."
                If flgCustomerVisible = True Then
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(Dn)Del prod,(DPn)Del paymt."
                Else
                    Stat "Scan or (X)Cancel trans.,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(Dn)Del prod,(DPn)Del paymt.,(FC) Find customer"
                End If
                DisplayTotals
            Else
                lblInput.Caption = "Start"
                Stat "Start by scan or (A)Appro,(R)Return,(CN)Credit note,(VRn)Void & replace,(DEP)Accept deposit,(RDEP)Refund deposit,(PC)Petty cash"
                ClearCustomer
                G3.Visible = False
            End If
            AutoSelect txtInput
        Case eqty, eQtyCashRefund, eQtyCreditNote, eQtyDeposit, eQtyAppro
            lblInput.Caption = "Qty "
            Stat "'.. to reverse"
            setInputBox oSALELine.Qty, "", "", True
        Case eRefundType_Cash
            lblInput.Caption = "Cash refunded."
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case eRefundType_Creditcard
            lblInput.Caption = "Credit card refund value."
            setInputBox CStr(oExchange.BalanceOwing), "", "", True
        Case eReviewExchanges
            lblInput.Caption = "Reviewing exchanges"
            Stat "Line number to print, DD to end review."
        Case eSearchCustomer, eSearchCustomerfordeposit, eSearchCustomerfordepositRefund, eSearchCustomerforAppro
            lblInput.Caption = "Search for . . . "
            strArg = Right(Trim(txtInput), Len(Trim(txtInput)) - 1)
            strArg2 = "Name"
            Stat ""
        Case eVoidandReplace
            lblInput.Caption = "Voiding #" & iToVoid & " and replacing"
            If flgSaleActive Then
                lblInput.Caption = "Product code, payment code or special."
                If flgCustomerVisible = True Then
                    Stat "Scan or (A)Appro,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(RDn)Redeem deposit,(Dn)Del prod,(DPn)Del paymt"
                Else
                    Stat "Scan or (A)Appro,(CN)Credit note,(C)Cash,(V)Voucher,(CC)Card,(Q)Cheque,(Dn)Del prod,(DPn)Del paymt,(FC)Find customer"
                End If
            Else
                lblInput.Caption = "Start cash refund "
                Stat "Start replacement by entering product code,   .. to reverse"
            End If
            flgVoidAndReplace = True
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetMessages"
End Sub
Private Sub SetDefaultsForStart()
    On Error GoTo errHandler
    lblReplacement.Caption = ""
    lblReplacement.Visible = False
    txtInput = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetDefaultsForStart"
End Sub
Private Sub DisplayProduct()
    On Error GoTo errHandler
    LoadSaleRow iCurrentSaleLine
    DisplayTotals
    
'    txtChange = "" 'erase any 'not on database' message
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
'Private Sub DisplayPayments()
'    On Error GoTo errHandler
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.DisplayPayments"
'End Sub
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
    On Error GoTo errHandler
    MsgBox "Open Drawer"
    If oPC.DriveDrawer = True Then
        MSComm1.Output = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.OpenDrawer"
End Sub
Private Function ValidRowNumberSet(pString As String)
Dim i As Integer
Dim bValid As Boolean

    arLineNumber = Split(pString, ",")
    bValid = True
    For i = 0 To UBound(arLineNumber)
        If Not IsNumeric(arLineNumber(i)) Then
            bValid = False
            Exit For
        End If
    Next i
    ValidRowNumberSet = bValid
    
End Function
Private Sub Statechange(pNewState As eState, Optional iRow As Integer, Optional pArg1 As String, Optional pArg2 As String)
    On Error GoTo errHandler
Dim Result As Integer
Dim lngCOLID As Long
Dim lngTmp As Long
Dim strDepositStatus As String
Dim bPrintCN As Boolean
Dim bValidRowSet  As Boolean
Dim dblDiscountRate As Double
Dim strDiscountCode As String
Dim strDiscountExplanation As String
Dim pDiscountDescription As String
Dim iDiscountRate As Double
Dim i As Integer
Dim lngTotalDeposit As Long

    Select Case enPresentState
    Case eAcceptDeposit
        Select Case pNewState
         Case ecancelsale
            If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                RejectSale
            End If
            SetPresentState eProductID
            SetMessages
        Case ePrevious
                ClearPayments
                oExchange.SetExchangeType eSaleType
                SetPresentState eProductID
                SetMessages
        Case Else
            If ValidRowNumberSet(txtInput) Then
                 bValidRowSet = True
                lngTotalDeposit = 0
                For i = 0 To UBound(arLineNumber)
                    iRow = CLng(arLineNumber(i))
                    If Not X3 Is Nothing Then
                        If iRow >= X3.LowerBound(1) And iRow <= X3.UpperBound(1) Then
                            lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                            lngDeposit = X3(lngTmp, 12)
                            lngCOLID = X3(lngTmp, 11)
                            strDepositTitle = X3(lngTmp, 5)
                            strDepositStatus = X3(lngTmp, 7)
                        Else
                            lngDeposit = 0
                        End If
                    End If
                    If lngDeposit > 0 And strDepositStatus = "O" Then
                        Set oSALELine = oExchange.SaleLines.Add
                        oSALELine.applyedit
                        oSALELine.beginedit
                        iCurrentSaleLine = iCurrentSaleLine + 1
                        X1.ReDim 1, iCurrentSaleLine, 1, 7
                        oSALELine.PID = X3(lngTmp, 13)
                        oSALELine.Price = lngDeposit
                        oSALELine.Title = X3(lngTmp, 5)
                        oSALELine.code = X3(lngTmp, 3)
                        lngTotalDeposit = lngTotalDeposit + lngDeposit
                        oSALELine.COLID = lngCOLID
                        DisplayProduct
                    End If
                Next
                Set oPAYMENTLine = oExchange.PaymentLines.Add
                oPAYMENTLine.applyedit
                oPAYMENTLine.beginedit
                iCurrentPaymentLine = iCurrentPaymentLine + 1
                X2.ReDim 1, iCurrentPaymentLine, 1, 3
                oPAYMENTLine.SetType "DEP"
                oPAYMENTLine.Amt = lngTotalDeposit
                DisplayPayment
                SetPresentState eDepositMode 'eConfirmationDeposit
                SetMessages
            Else
                MsgBox "Invalid row selection.", vbOKOnly, "Can't do this"
                SetMessages
            End If
        End Select
    Case eAcceptDepositRef
        Select Case pNewState
        Case ePrevious
            SetPresentState eAcceptDeposit
            SetMessages
        Case Else
            SetPresentState eConfirmation
            SetMessages
        End Select
        DisplayPayment
    Case eRefundDeposit
        Select Case pNewState
         Case ecancelsale
            If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                RejectSale
            End If
            oExchange.SetExchangeType eSaleType
            SetPresentState eProductID
            SetMessages
        Case ePrevious
                ClearPayments
                oExchange.SetExchangeType eSaleType
                SetPresentState eProductID
                SetMessages
        Case Else
            If ValidRowNumberSet(txtInput) Then
                bValidRowSet = True
                For i = 0 To UBound(arLineNumber)
                    iRow = CLng(arLineNumber(i))
                    If Not X3 Is Nothing Then
                        If iRow >= X3.LowerBound(1) And iRow <= X3.UpperBound(1) Then
                            lngTmp = X3.Find(1, 1, CStr(iRow), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
                            lngDeposit = X3(lngTmp, 12)
                            lngCOLID = X3(lngTmp, 11)
                            strDepositTitle = X3(lngTmp, 5)
                            strDepositStatus = X3(lngTmp, 7)
                        Else
                            lngDeposit = 0
                        End If
                    End If
                    If lngDeposit > 0 And strDepositStatus = "P" Then
                        Set oSALELine = oExchange.SaleLines.Add
                        oSALELine.applyedit
                        oSALELine.beginedit
                        iCurrentSaleLine = iCurrentSaleLine + 1
                        X1.ReDim 1, iCurrentSaleLine, 1, 7
                        oSALELine.PID = X3(lngTmp, 13)
                        oSALELine.Price = lngDeposit
                        oSALELine.Title = X3(lngTmp, 5)
                        oSALELine.code = X3(lngTmp, 3)
                        lngTotalDeposit = lngTotalDeposit + lngDeposit
                        oSALELine.COLID = lngCOLID
                        DisplayProduct
                    End If
                Next
                Set oPAYMENTLine = oExchange.PaymentLines.Add
                oPAYMENTLine.applyedit
                oPAYMENTLine.beginedit
                iCurrentPaymentLine = iCurrentPaymentLine + 1
                X2.ReDim 1, iCurrentPaymentLine, 1, 3
                oPAYMENTLine.SetType "RDEP"
                oPAYMENTLine.Amt = lngTotalDeposit * -1
                DisplayPayment
                SetPresentState eConfirmationRefundDeposit
                SetMessages
'''
            Else
                SetPresentState eProductID
                SetMessages
            End If
        End Select
    Case eConfirmation
            Select Case pNewState
            Case ePrevious
                RemovePaymentLine iCurrentPaymentLine
                SetPresentState eProductID
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    oExchange.SalesPersonName = strName
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
                    If oExchange.transactionType = "S" Then
                        lblChange.Caption = "CHANGE: " & oExchange.ChangeGivenF
                        bWaitForClearance = True
                    End If
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel sale?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to make a sale. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
              '  setInputBox "", "", "", True
            End Select
    Case eConfirmationAppro
            Select Case pNewState
            Case ePrevious
                SetPresentState eAppro
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm Appro?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to issue an Appro. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
            End Select
    Case eConfirmationCashrefund
            Select Case pNewState
            Case ePrevious
                SetPresentState eCashRefund
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel cash refund?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to issue a cash refund. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
            End Select
    Case eConfirmationCreditNote
            Select Case pNewState
            Case ePrevious
                SetPresentState eCreditNote
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel credit note?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to issue a credit note. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
                setInputBox "", "", "", True
            End Select
    Case eConfirmationDeposit
            Select Case pNewState
            Case ePrevious
                SetPresentState eAcceptDeposit
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel deposit?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to accept a deposit. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
            End Select
        '    setInputBox "", "", "", True
    Case eConfirmationRefundDeposit
            Select Case pNewState
            Case ePrevious
                SetPresentState eRefundDeposit
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel deposit refund?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to refund a deposit. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
            End Select
       '     setInputBox "", "", "", True
    
    Case enull
            Select Case pNewState
            Case elogin
                SetPresentState elogin
                SetMessages
                LogonOperator
            End Select
    Case eCashRefund
        Select Case pNewState
        Case eSearchCustomer
            If GetCustomer(pArg1, pArg2) Then
             '   FetchCOLS
             '   LoadCOLS
                oExchange.RecalculateAllLines
                oExchange.CalculateTotals
                RefreshSaleDisplay
                DisplayCustomerDetails
                txtInput = ""
                SetMessages
            End If
        Case ecancelsale
            If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                RejectSale
                setInputBox "", "", "", True
            End If
            SetPresentState eProductID
            SetMessages
        Case eRefundType_Cash
            Set oPAYMENTLine = oExchange.PaymentLines.Add
            oPAYMENTLine.applyedit
            oPAYMENTLine.beginedit
            iCurrentPaymentLine = iCurrentPaymentLine + 1
            X2.ReDim 1, iCurrentPaymentLine, 1, 3
            oPAYMENTLine.SetType "C"
            SetPresentState eRefundType_Cash
            SetMessages
        Case eRefundType_Creditcard
            Set oPAYMENTLine = oExchange.PaymentLines.Add
            oPAYMENTLine.applyedit
            oPAYMENTLine.beginedit
            iCurrentPaymentLine = iCurrentPaymentLine + 1
            X2.ReDim 1, iCurrentPaymentLine, 1, 3
            oPAYMENTLine.SetType "A"
            SetPresentState eRefundType_Creditcard
            SetMessages
        Case eConfirmationCashrefund
            SetPresentState eConfirmationCashrefund
            SetMessages
        Case eDelete
            If iRow <= oExchange.SaleLines.Count And iRow > 0 Then
                oExchange.SaleLines.Remove (iRow)
                oExchange.SaleLines.applyedit
                oExchange.SaleLines.beginedit
                oExchange.CalculateTotals
                X1.DeleteRows (iRow)
                G1.ReBind
                SetPresentState eCashRefund
                SetMessages
                iCurrentSaleLine = iCurrentSaleLine - 1
            End If
        Case ePrevious
            If flgSaleActive = False Then
                oExchange.SetExchangeType eSaleType
                SetPresentState eProductID
                SetMessages
            End If
        Case Else
            If LoadProductFromCode Then
                oExchange.CalculateTotals
                DisplayProduct
                SetPresentState ePriceCashRefund
                SetMessages
            Else
                MsgBox "Not on database or invalid action", vbInformation, "Status"
            '    txtChange = "Not on database."
            End If
        End Select
';;;;
    Case eAppro
        Select Case pNewState
        Case ecancelsale
            If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                RejectSale
                setInputBox "", "", "", True
                ClearCustomer
            End If
            SetPresentState eProductID
            SetMessages
        Case eConfirmationAppro
            SetPresentState eConfirmationAppro
            SetMessages
        Case eDelete
            If iRow <= oExchange.SaleLines.Count And iRow > 0 Then
                oExchange.SaleLines.Remove (iRow)
                oExchange.SaleLines.applyedit
                oExchange.SaleLines.beginedit
                oExchange.CalculateTotals
                X1.DeleteRows (iRow)
                G1.ReBind
                SetPresentState eAppro
                SetMessages
                iCurrentSaleLine = iCurrentSaleLine - 1
            End If
        Case ePrevious
            If flgSaleActive = False Then
                SetPresentState eProductID
                oExchange.SetExchangeType eSaleType
                SetMessages
                ClearCustomer
            End If
        Case Else
            If LoadProductFromCode Then
                oExchange.CalculateTotals
                DisplayProduct
                SetPresentState ePriceAppro
                SetMessages
            Else
                MsgBox "Not on database or invalid action", vbInformation, "Status"
            '    txtChange = "Not on database."
            End If
        End Select

    Case eCreditNote
        Select Case pNewState
        Case ecancelsale
            If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                RejectSale
                AutoSelect txtInput
            End If
            SetPresentState eProductID
            SetMessages
        Case eConfirmationCreditNote
            SetPresentState eConfirmationCreditNote
            SetMessages
        Case ePrevious
            If flgSaleActive = False Then
              SetPresentState eProductID
              oExchange.SetExchangeType eSaleType
              SetMessages
            End If
        Case eDelete
            If iRow <= oExchange.SaleLines.Count And iRow > 0 Then
                RemoveSaleLine iRow
                SetPresentState eCreditNote
                SetMessages
            End If
        Case Else
            If LoadProductFromCode Then
                oExchange.CalculateTotals
                DisplayProduct
                SetPresentState ePriceCreditNote
                SetMessages
            Else
                MsgBox "Not on database or invalid action", vbInformation, "Status"
            End If
        End Select
    Case eDepositMode
            Select Case pNewState
            Case ePrevious
                SetPresentState eAcceptDeposit
                SetMessages
            Case Else
                If txtInput = "CN" Or txtInput = "V" Or txtInput = "Q" Or txtInput = "CC" Or txtInput = "C" Then
                    oPAYMENTLine.SetType UCase(Trim(txtInput))
                    SetPresentState eConfirmationDeposit
                    SetMessages
                    DisplayPayment
                End If
            End Select
    
    Case eDiscount
            Select Case pNewState
            Case ePrevious
                SetPresentState ePrice
                SetMessages
            Case Else
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
                        dblDiscountRate = GetDiscount(strDiscountCode, pDiscountDescription)
                    End If
                    oSALELine.SetDiscountRateDbl dblDiscountRate, pDiscountDescription
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eqty
                    SetMessages
                End If
            End Select
            
    Case eDiscountAppro
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceAppro
                SetMessages
            Case Else
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
                        dblDiscountRate = GetDiscount(strDiscountCode, pDiscountDescription)
                    End If
                    If oSALELine.SetDiscountRateDbl(dblDiscountRate, pDiscountDescription) Then
                        oExchange.CalculateTotals
                        DisplayProduct
                        SetPresentState eQtyAppro
                        SetMessages
                    Else
                        SetTip "Invalid Discount."
                    End If
                End If
            End Select
    Case eDiscountCashRefund
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceCashRefund
                SetMessages
            Case Else
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
                        dblDiscountRate = GetDiscount(strDiscountCode, pDiscountDescription)
                    End If
                    If oSALELine.SetDiscountRateDbl(iDiscountRate, pDiscountDescription) Then
                        oExchange.CalculateTotals
                        DisplayProduct
                        SetPresentState eQtyCashRefund
                        SetMessages
                    Else
                        SetTip "Invalid Discount."
                    End If
                End If
            End Select
    Case eDiscountCreditNote
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceCreditNote
                SetMessages
            Case Else
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
                        dblDiscountRate = GetDiscount(strDiscountCode, pDiscountDescription)
                    End If
                    If oSALELine.SetDiscountRateDbl(iDiscountRate, pDiscountDescription) Then
                        oExchange.CalculateTotals
                        DisplayProduct
                        SetPresentState eQtyCreditNote
                        SetMessages
                    Else
                        SetTip "Invalid Discount."
                    End If
                End If
            End Select
        
    Case elogin
            SwapOperator
            SetPresentState eProductID
'    Case ePettyCashCredit
'            Select Case pNewState
'            Case ePrevious
'                SetPresentState eProductID
'                oExchange.SetExchangeType eSaleType
'                RemovePaymentLine iCurrentPaymentLine
'                SetMessages
'            Case Else
'                oExchange.SetExchangeType ePettyCashType
'                Set oPAYMENTLine = oExchange.PaymentLines.Add
'                oPAYMENTLine.applyedit
'                oPAYMENTLine.beginedit
'                oPAYMENTLine.setReference CStr(lngPettyCashLine)  'Exchange number of petty cash withdrawal being credited
'                iCurrentPaymentLine = iCurrentPaymentLine + 1
'                X2.ReDim 1, iCurrentPaymentLine, 1, 3
'                oPAYMENTLine.SetType "R"
'                DisplayPayment
'                SetPresentState ePettyCashCreditAmt
'                SetMessages
'            End Select
    Case ePettyCashCreditAmt
            Select Case pNewState
            Case ePrevious
                SetPresentState ePettyCashCredit
                SetMessages
            Case Else
                    If oPAYMENTLine.setAmt(txtInput) Then
                        SetPresentState ePettyCashCreditConfirmation
                        SetMessages
                    Else
                        SetTip "Invalid petty cash amount."
                    End If
              '  End If
            End Select
    Case ePettyCashCreditConfirmation
            Select Case pNewState
            Case ePrevious
                SetPresentState ePettyCashCreditAmt
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 3 Then
                    oExchange.SalesPersonID = lngSMID
                    OpenDrawer
                    AcceptSale
                Else
                    MsgBox "You do not have the authority to return a petty cash amount. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                    RejectSale
                End If
            End Select
    Case ePettyCash
            Select Case pNewState
            Case ePrevious
                SetPresentState eProductID
                oExchange.SetExchangeType eSaleType
                RemovePaymentLine iCurrentPaymentLine
                SetMessages
            Case Else
                If InStr(1, strPettyCashTypes, Left(txtInput, 1)) > 0 Then
                    oExchange.SetExchangeType ePettyCashType
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    oPAYMENTLine.setReference Trim(txtInput)
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "W"
                    DisplayPayment
                    SetPresentState ePettyCashReason
                    SetMessages
                Else
                    SetTip "Invalid petty cash account."
                End If
            End Select
    Case ePettyCashReason
            Select Case pNewState
            Case ePrevious
                SetPresentState ePettyCash
                SetMessages
            Case Else
                If oPAYMENTLine.setNote(txtInput) Then
                    SetPresentState ePettyCashAmt
                    SetMessages
                Else
                    SetTip "Invalid petty cash reason."
                End If
            End Select

    Case ePettyCashAmt
            Select Case pNewState
            Case ePrevious
                SetPresentState ePettyCash
                SetMessages
            Case Else
              '  If InStr(1, strPettyCashTypes, Left(txtInput, 1)) > 0 Then 'valid Petty cash type
              '      iDiscountRate = GetDiscount(Left(txtInput, 1), pDiscountDescription)
             '       If oSALELine.SetDiscountRateDbl(iDiscountRate, pDiscountDescription) Then
                    If oPAYMENTLine.setAmt(txtInput) Then
                        SetPresentState ePettyCashConfirmation
                        SetMessages
                    Else
                        SetTip "Invalid petty cash account."
                    End If
              '  End If
            End Select
    Case ePettyCashConfirmation
            Select Case pNewState
            Case ePrevious
                SetPresentState ePettyCashAmt
                SetMessages
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 3 Then
                    oExchange.SalesPersonID = lngSMID
                    OpenDrawer
                    AcceptSale
                Else
                    MsgBox "You do not have the authority to withdraw cash. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                    RejectSale
                End If
            End Select
    
    Case eProductID
            Select Case pNewState
            Case ePettyCash
                If SecurityControl(3, lngSMID, strName, , "Enter your security key.", "Your key is invalid") Then
                    SetPresentState ePettyCash
                    SetMessages
                Else
                    MsgBox "You do not have the authority to withdraw cash. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                   ' RejectSale
                End If
            Case ePettyCashCredit
                If SecurityControl(3, lngSMID, strName, , "Enter your security key.", "Your key is invalid") Then
                    oExchange.SetExchangeType ePettyCashType
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    oPAYMENTLine.setReference CStr(lngPettyCashLine)  'Exchange number of petty cash withdrawal being credited
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "R"
                    DisplayPayment
                    SetPresentState ePettyCashCreditAmt
                    SetMessages
                Else
                    MsgBox "You do not have the authority to deposit cash. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If

            Case eReviewExchanges
                setInputBox "", "", "", True
                If Not flgSaleActive Then
                    SetPresentState pNewState
                    SetMessages
                    ShowTransactions True
                End If
            Case eVoidandReplace
                lblReplacement.Caption = "Voiding and replacing Transaction #" & CStr(iToVoid)
                lblReplacement.Visible = True
                txtInput = ""
            Case eAppro
                If flgSaleActive = True Then
                    If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                        setInputBox "", "", "", True
                        oExchange.SetExchangeType eApproType
                        SetPresentState pNewState
                        SetMessages
                    End If
                Else
                    oExchange.SetExchangeType eApproType
                    SetPresentState pNewState
                    SetMessages
                End If
            Case eCashRefund
                If flgSaleActive = True Then
                    If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                        setInputBox "", "", "", True
                        oExchange.SetExchangeType ereturntype
                        SetPresentState pNewState
                        SetMessages
                    End If
                Else
                    oExchange.SetExchangeType ereturntype
                    SetPresentState pNewState
                    SetMessages
                End If
            Case eCreditNote
                If flgSaleActive = True Then
                    If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                        SetPresentState pNewState
                        SetMessages
                    End If
                Else
                    oExchange.SetExchangeType eCreditNoteType
                    SetPresentState pNewState
                    SetMessages
                End If
            Case ecancelsale
                If MsgBox("Cancel this transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                    RejectSale
                    setInputBox "", "", "", True
                End If
            Case eRebuildIndexes
                    Connect
                    RebuildIndexes
                    Disconnect
            Case eSearchCustomer
                    If GetCustomer(pArg1, pArg2) Then
                        FetchCOLS
                        LoadCOLS
                        oExchange.RecalculateAllLines
                        oExchange.CalculateTotals
                        RefreshSaleDisplay
                        DisplayCustomerDetails
                        txtInput = ""
                        SetMessages
                    End If
            Case eSearchCustomerfordeposit
                   If GetCustomer(pArg1, pArg2) Then
                        oExchange.SetExchangeType eDepositType
                    '    oExchange.transactionType = "DEP"
                        FetchCOLS
                        LoadCOLS
                        oExchange.CalculateTotals
                        RefreshSaleDisplay
                        DisplayCustomerDetails
                        txtInput = ""
                        SetPresentState eAcceptDeposit
                        SetMessages
                    Else
                        SetPresentState eProductID
                        SetMessages
                    End If
            Case eSearchCustomerforAppro
                   If GetCustomer(pArg1, pArg2) Then
                        oExchange.SetExchangeType eApproType
                   '     oExchange.transactionType = "APP"
                   '     FetchCOLS
                   '     LoadCOLS
                        oExchange.CalculateTotals
                        RefreshSaleDisplay
                        DisplayCustomerDetails
                        txtInput = ""
                        SetPresentState eAppro
                        SetMessages
                    Else
                        SetPresentState eProductID
                        SetMessages
                    End If
            Case eSearchCustomerfordepositRefund
                   If GetCustomer(pArg1, pArg2) Then
                        oExchange.SetExchangeType eReturnDepositType
                    '    oExchange.transactionType = "RDEP"
                        FetchCOLS
                        LoadCOLS
                        oExchange.CalculateTotals
                        RefreshSaleDisplay
                        DisplayCustomerDetails
                        txtInput = ""
                        SetPresentState eRefundDeposit
                        SetMessages
                    Else
                        SetPresentState eProductID
                        SetMessages
                    End If
            Case ePaymentType_Cash
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "C"
                    SetPresentState ePaymentType_Cash
                    SetMessages
            Case ePaymentType_Cheque
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "Q"
                    SetPresentState ePaymentType_Cheque
                    SetMessages
            Case ePaymentType_CreditCard
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "A"
                    SetPresentState ePaymentType_CreditCard
                    SetMessages
            Case ePaymentType_voucher
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "V"
                    SetPresentState ePaymentType_voucher
                    SetMessages
            Case ePaymentType_RedeemDeposit
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "RD"
                    oPAYMENTLine.COLID = Trim(X3(iCOLForDeposit, 11))
                    SetPresentState ePaymentType_RedeemDeposit
                    SetMessages
            Case ePaymentType_CreditNote
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.applyedit
                    oPAYMENTLine.beginedit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.SetType "CN"
                    SetPresentState ePaymentType_CreditNote
                    SetMessages
            Case eXTerminate
                If MsgBox("Confirm cash up?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                    SetPresentState eProductID
                    SetMessages
                Else
                    If oExchange.SaleLines.Count > 0 Then
                        oExchange.CancelEdit
                    End If
                    bCloseXsession = True
                    Unload Me
                End If
            Case eZTerminate
                If MsgBox("Close Z session?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                    SetPresentState eProductID
                    SetMessages
                Else
                    If oExchange.SaleLines.Count > 0 Then
                        oExchange.CancelEdit
                    End If
                    bCloseZsession = True
                    Unload Me
                End If
            Case elogin
                SwapOperator
                SetPresentState eProductID
                SetMessages
            Case eDelete
                If iRow <= oExchange.SaleLines.Count And iRow > 0 Then
                    oExchange.SaleLines.Remove (iRow)
                    oExchange.SaleLines.applyedit
                    oExchange.SaleLines.beginedit
                    oExchange.CalculateTotals
                    X1.DeleteRows (iRow)
                    G1.ReBind
                    SetPresentState eProductID
                    SetMessages
                    iCurrentSaleLine = iCurrentSaleLine - 1
                End If
            Case eDeletePayment
                If iRow <= oExchange.PaymentLines.Count And iRow > 0 Then
                    RemovePaymentLine iRow
                    SetPresentState eProductID
                    SetMessages
                End If
            Case eConfirmation
                If oExchange.PaymentsComplete() Then
                    SetPresentState eConfirmation
                    SetMessages
                End If
            Case Else
                If txtInput = ".." Then
                    SetDefaultsForStart
                Else
                    If LoadProductFromCode Then
                        oExchange.CalculateTotals
                        DisplayTotals
                        SetPresentState ePrice
                        SetMessages
                    Else
                        'txtChange = "Not on database."
                        MsgBox "Not on database or invalid action", vbInformation, "Status"
                    End If
                End If
            End Select
    Case ePaymentType_CreditNote
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
               ' oPAYMENTLine.setType "Credit Note"
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_Cash
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_Cheque
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                SetPresentState ePaymentType_ChequeRef
                SetMessages
                If oExchange.cashtransaction Then
                    OpenDrawer
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_CreditCard
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                SetPresentState ePaymentType_CreditCardRef
                SetMessages
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_voucher
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                SetPresentState ePaymentType_voucherRef
                SetMessages
                    If oExchange.cashtransaction Then
                        OpenDrawer
                    End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_ChequeRef
        Select Case pNewState
        Case ePrevious
            SetPresentState ePaymentType_Cheque
            SetMessages
'            setInputBox CStr(oPAYMENTLine.Amt), "", "", True
        Case Else
            If oPAYMENTLine.setReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
    Case ePaymentType_CreditCardRef
        Select Case pNewState
        Case ePrevious
            SetPresentState ePaymentType_CreditCard
            SetMessages
 '           txtInput = CStr(oPAYMENTLine.Amt)
        Case Else
            If oPAYMENTLine.setReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
    Case ePaymentType_voucherRef
        Select Case pNewState
        Case ePrevious
            SetPresentState ePaymentType_voucher
            SetMessages
 '           txtInput = CStr(oPAYMENTLine.Amt)
        Case Else
            If InStr(1, strValidVoucherTypes, Left(txtInput, 1)) > 0 And Len(txtInput) > 1 Then 'valid voucher type
                If oPAYMENTLine.setReference(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    If oExchange.PaymentsComplete Then
                        SetPresentState eConfirmation
                        SetMessages
                    Else
                        SetPresentState eProductID
                        SetMessages
                    End If
                    DisplayPayment
                Else
                    SetTip "Invalid Reference."
                End If
            End If
        End Select
    Case ePaymentType_RedeemDeposit
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
            
    Case ePrice   'expect txtinput to reflect .. or request for disc or actual price
            Select Case pNewState
            Case ePrevious
                RemoveSaleLine iCurrentSaleLine
                SetPresentState eProductID
                SetMessages
                If oExchange.LoyaltyValue > 0 Then
                    DisplayCustomerDetails
                End If
            Case eDiscount
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eDiscount
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eqty
                    SetMessages
                    DisplayCustomerDetails
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceAppro
            Select Case pNewState
            Case ePrevious
                RemoveSaleLine iCurrentSaleLine
                SetPresentState eAppro
                SetMessages
            Case eDiscountAppro
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eDiscountAppro
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eQtyAppro
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceCashRefund
            Select Case pNewState
            Case ePrevious
                RemoveSaleLine iCurrentSaleLine
                SetPresentState eCashRefund
                SetMessages
            Case eDiscountCashRefund
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eDiscountCashRefund
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eQtyCashRefund
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceCreditNote
            Select Case pNewState
            Case ePrevious
                RemoveSaleLine iCurrentSaleLine
                SetPresentState eCreditNote
                SetMessages
            Case eDiscountCreditNote
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eDiscountCreditNote
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eQtyCreditNote
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceDeposit
            Select Case pNewState
            Case ePrevious
                RemoveSaleLine iCurrentSaleLine
                SetPresentState eAcceptDeposit
                SetMessages
            Case eDiscountDeposit
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eDiscountDeposit
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eQtyDeposit
                    SetMessages
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case eqty
            Select Case pNewState
            Case ePrevious
                SetPresentState ePrice
                SetMessages
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), False) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eProductID
                    SetMessages
                    oSALELine.applyedit
                    oSALELine.beginedit
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyAppro
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceAppro
                SetMessages
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), False) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eAppro
                    SetMessages
                    oSALELine.applyedit
                    oSALELine.beginedit
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyCashRefund
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceCashRefund
                SetMessages
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), True) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eCashRefund
                    SetMessages
                    oSALELine.applyedit
                    oSALELine.beginedit
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyCreditNote
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceCreditNote
                SetMessages
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), True) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eCreditNote
                    SetMessages
                    oSALELine.applyedit
                    oSALELine.beginedit
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyDeposit
            Select Case pNewState
            Case ePrevious
                SetPresentState ePriceCreditNote
                SetMessages
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), True) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetPresentState eCreditNote
                    SetMessages
                    oSALELine.applyedit
                    oSALELine.beginedit
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eRefundType_Cash
        Select Case pNewState
        Case ePrevious
            SetPresentState eCashRefund
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case eRefundType_Creditcard
        Select Case pNewState
        Case ePrevious
            SetPresentState eCashRefund
            SetMessages
        Case Else
            If oPAYMENTLine.setAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetPresentState eConfirmation
                    SetMessages
                Else
                    SetPresentState eProductID
                    SetMessages
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select

    Case eReviewExchanges
        Select Case pNewState
        Case eReviewExchanges
            SetPresentState eProductID
            SetMessages
        Case Else
            ShowExchange
            setInputBox "", "", "", True
        End Select
    Case eSearchCustomerfordeposit
        Case eAcceptDeposit
            If flgSaleActive = True Then
                If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                    RejectSale
                    oExchange.SetExchangeType eDepositType
                    SetPresentState pNewState
                    SetMessages
                End If
            Else
                oExchange.SetExchangeType eDepositType
                SetPresentState pNewState
                SetMessages
            End If
    Case eVoidandReplace
        Select Case pNewState
        Case ePrevious
            SetPresentState eProductID
            SetMessages
        End Select
        
    End Select
    Exit Sub
MEX:
    flgLoading = True
    txtInput = ""
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Statechange(pNewState,iRow,pArg1,pArg2)", Array(pNewState, iRow, pArg1, _
         pArg2)
End Sub

Private Sub RemovePaymentLine(iRow As Integer)
    On Error GoTo errHandler
    If iRow = 0 Then Exit Sub
    oExchange.PaymentLines.Remove (iRow)
    oExchange.PaymentLines.applyedit
    oExchange.PaymentLines.beginedit
    oExchange.CalculateTotals
    txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF
    X2.DeleteRows (iRow)
    G2.ReBind
    iCurrentPaymentLine = iCurrentPaymentLine - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemovePaymentLine(iRow)", iRow
End Sub

Private Sub RemoveSaleLine(iRow As Integer)
    On Error GoTo errHandler
    oExchange.SaleLines.Remove (iRow)
    oExchange.SaleLines.applyedit
    oExchange.SaleLines.beginedit
    oExchange.CalculateTotals
    X1.DeleteRows (iRow)
    G1.ReBind
    iCurrentSaleLine = iCurrentSaleLine - 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RemoveSaleLine(iRow)", iRow
End Sub

Private Sub RejectSale()
    On Error GoTo errHandler
'    If Not oPAYMENTLine Is Nothing Then
'        If oPAYMENTLine.IsEditing Then oPAYMENTLine.CancelEdit
'    End If
    oExchange.CancelEdit
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    oExchange.beginedit
    oExchange.SalesPersonID = oPC.ZSession.Opsession.OperatorID
    oExchange.SetExchangeType eSaleType
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
    txtInput = ""
  '  txtChange = ""
    lblCustomername.Caption = ""
    lblReplacement.Visible = False
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    iToVoid = 0
    flgCustomerVisible = False
    flgSaleActive = False
    bLoyaltyCustomer = False
    SetPresentState eProductID
    SetMessages
    SetTitleBar False
    SetForCOLSVisible False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RejectSale"
End Sub
'Private Sub AcceptDeposit()
'    On Error GoTo errHandler
'
''Save and send exchange
'   ' oExchange.setExchangeType eSaleType
'    oExchange.OperatorID = lngSMID
'    oExchange.StaffName = strName
'    If iToVoid > 0 Then oExchange.ToVoid = iToVoid
'    oExchange.ApplyEdit
'    AddExchange
'    SendPOSExchange oExchange.ExchangeID, oExchange.OPSID, oExchange.ZID
''Print Till Slip
'   ' PrintDepositSlip
''Start new exchange
'    Set oExchange = Nothing
'    Set oExchange = New a_Exchange
'    oExchange.BeginEdit
'    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
'    oExchange.transactionType = "S"
'    ClearTextFields
'    X1.Clear
'    X1.ReDim 1, 1, 1, 8
'    G1.ReBind
'    X2.Clear
'    X2.ReDim 1, 1, 1, 3
'    G2.ReBind
'    txtInput = ""
'    txtChange = ""
'    lblReplacement.Visible = False
'    iCurrentSaleLine = 0
'    iCurrentPaymentLine = 0
'    bLoyaltyCustomer = False
'    iToVoid = 0
'    flgSaleActive = False
'    SetNewState eProductID
'    SetTitleBar False
'    SetForCOLSVisible False
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.AcceptSale"
'End Sub

Private Sub AcceptSale()
        On Error GoTo errHandler
Dim lngRow As Long
Dim lngLowerBound As Long
Dim strPos As String

'Save and send exchange

    If oExchange.transactionType = "R" Or oExchange.transactionType = "C" Then
        If GetCustomer("", "") Then
            DisplayCustomerDetails
        End If
        Set frmCustID = New frmIDCustomer
        frmCustID.component oExchange.CustomerName
        frmCustID.Show vbModal
        If oExchange.Customer.Name > "" Then
            oExchange.CustomerName = ""
        Else
            oExchange.CustomerName = frmCustID.CustomerName
        End If
        oSALELine.Counterfoil = frmCustID.Counterfoil
       
    End If

    oExchange.CalculateSalesTotal
    oExchange.OperatorID = lngSMID
    oExchange.StaffName = strName
    If iToVoid > 0 Then
        oExchange.ToVoid = iToVoid
    End If
    
    lngLowerBound = X4.LowerBound(1)
    lngRow = X4.Find(lngLowerBound, 1, iToVoid)
    Do While lngRow >= lngLowerBound
        X4(lngRow, 12) = oExchange.ExchangeNumber
        lngRow = X4.Find(lngRow + 1, 1, iToVoid)
    Loop
    G4.Refresh
    oExchange.applyedit
    oPC.DBLocalConn.BeginTrans
    oExchange.PostExchange
    oPC.DBLocalConn.CommitTrans
    
    AddExchange
    SendPOSExchange oExchange.ExchangeID, oExchange.OPSID, oExchange.ZID

'Print Till Slip
    If oPC.PrintSlips = True Then
        Select Case oExchange.transactionType
        Case "S"
            If (oExchange.Customer.CustomerType = "L1" Or oExchange.Customer.CustomerType = "L2" Or oExchange.Customer.CustomerType = "L3") And (oExchange.LoyaltyValue > 0) Then
                PrintLoyaltyVoucher
            End If
            PrintSalesSlip oPC.InvoiceCopyCount
        Case "R"
            PrintSalesSlip oPC.ReturnCopyCount
        Case "PC"
            PrintPettyCashVoucher oPC.PettyCashCount
        Case "C"
            PrintSalesSlip oPC.CreditNoteCopyCount
        Case "DEP"
            PrintDepositSlip oPC.DepositCopyCount
        Case "RDEP"
            PrintDepositRefundSlip oPC.DepositCopyCount
        Case "APP"
            PrintAPPROSlip oPC.ApproCopyCount
        End Select
    
    'If there is a CN being paid out as change - we must print it
        If oExchange.MustPrintCNasChange() Then
            PrintCNasChange oExchange.ChangeGiven, 1, False
        End If
    End If
    
'Start new exchange
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    oExchange.beginedit
    oExchange.SalesPersonID = oPC.ZSession.Opsession.OperatorID
    oExchange.SetExchangeType eSaleType
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
'    txtInput = ""
    lblCustomername.Caption = ""
    lblReplacement.Visible = False
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    bLoyaltyCustomer = False
    iToVoid = 0
    flgSaleActive = False
    flgCustomerVisible = False
    SetPresentState eProductID
    SetMessages
    SetTitleBar True
    SetForCOLSVisible False
    Exit Sub
errHandler:
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AcceptSale", , EA_DFTRBK, , "strPOS", Array(strPos)
End Sub
'Private Sub AcceptReturn()
'    On Error GoTo errHandler
'
''Save and send exchange
'    oExchange.setExchangeType eTypCashRefundType
'    oExchange.OperatorID = lngSMID
'    oExchange.StaffName = strName
'    If iToVoid > 0 Then oExchange.ToVoid = iToVoid
'    oExchange.ApplyEdit
'    SendPOSExchange oExchange.ExchangeID, oExchange.OPSID, oExchange.ZID
' 'Print Till Slip
'  '  oExchange.printTillSlip
''Start new exchange
'    Set oExchange = Nothing
'    Set oExchange = New a_Exchange
'    oExchange.BeginEdit
'    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
'    oExchange.transactionType = "S"
'    ClearTextFields
'    X1.Clear
'    X1.ReDim 1, 1, 1, 8
'    G1.ReBind
'    X2.Clear
'    X2.ReDim 1, 1, 1, 3
'    G2.ReBind
'    txtInput = ""
'    lblReplacement.Visible = False
'    iCurrentSaleLine = 0
'    iCurrentPaymentLine = 0
'    iToVoid = 0
'    bLoyaltyCustomer = False
'    SetNewState eProductID
'    SetTitleBar
'    flgSaleActive = False
'    SetForCOLSVisible False
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.AcceptSale"
'End Sub

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
    oExchange.PaymentLines.applyedit
    oExchange.PaymentLines.beginedit
    iCurrentPaymentLine = 0
    X2.Clear
    X2.ReDim 1, iCurrentPaymentLine, 1, 3
    G2.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearPayments"
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
     If oExchange.IdentifyCustomer = True And oExchange.CustomerName = "" Then
         Set frmCustID = New frmIDCustomer
         frmCustID.component oExchange.CustomerName
         frmCustID.Show vbModal
         oExchange.CustomerName = frmCustID.CustomerName
         oSALELine.Counterfoil = frmCustID.Counterfoil
         Unload frmCustID
     End If
     
     
     iCurrentSaleLine = iCurrentSaleLine + 1
     X1.ReDim 1, iCurrentSaleLine, 1, 8
     oSALELine.Title = strTitle
     oSALELine.MainAuthor = strAuthor
     oSALELine.Price = lngPrice
     oSALELine.SetQty "1", False
     oSALELine.VATRate = lngVatrate
     oSALELine.DiscountRate = CDbl(lngDiscount)   'Set the discount and loyalty discount from the product data
     oSALELine.LoyaltyDiscount = lngLoyaltyDiscount         ''
     oSALELine.NoDiscountAllowed = bNoDiscountAllowable     ''
     oSALELine.DiscountRule = strDiscountRule               ''
     oSALELine.PID = strPID
     If strCode > "" Then
         oSALELine.code = strCode
     Else
         oSALELine.code = strEAN
     End If
     
     
     oSALELine.applyedit
     oSALELine.beginedit
     LoadProductFromCode = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadProductFromCode", , , , "StrPos", Array(strPos)
End Function
Private Sub oExchange_Recalculate()
    On Error GoTo errHandler
    If flgUnloading Then Exit Sub
    '    MsgBox "Code commented"
'    If bLoyaltyCustomer = True Then
'        lblLoyaltyValue.Caption = oExchange.LoyaltyValueF
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_Recalculate"
End Sub
Private Function GetCustomer(pArg1 As String, pArg2 As String) As Boolean
    On Error GoTo errHandler
Dim frm As New frmBrowseCustomers2
    GetCustomer = False
    frm.Show vbModal
'    If frm.CustomerType > "" Then
    If frm.CustomerName > "" Then
        flgCustomerVisible = True
        strCustomername = frm.CustomerName
        G3.Caption = frm.CustomerName
        lngCustomerID = frm.CustomerID
        oExchange.SetCustomer lngCustomerID
        oExchange.CustomerName = frm.CustomerName
        If UCase(oExchange.Customer.CustomerType) = "L1" Then
            bLoyaltyCustomer = True
        End If
        GetCustomer = True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetCustomer(pArg1,pArg2)", Array(pArg1, pArg2)
End Function
Private Function ClearCustomer()
     '   oExchange.SetCustomer 0
        lngCustomerID = 0
        strCustomername = ""
        flgCustomerVisible = False
        G3.Caption = ""
        G3.Visible = False
        Me.lblCustomername.Caption = ""
End Function
Private Sub DisplayCustomerDetails()
    On Error GoTo errHandler
Dim strDetails As String
    Select Case UCase(oExchange.Customer.CustomerType)
    Case "L1"   'Loyalty club 1 member
            strDetails = oExchange.Customer.NameAndCode(99) & vbCrLf & "Loyalty value: " & oExchange.LoyaltyValueF
    Case "BC"
        strDetails = oExchange.Customer.NameAndCode(99) & vbCrLf & "(Book club)" & ":" & oExchange.DiscountRateF
    Case ""
        strDetails = oExchange.Customer.NameAndCode(99)
    End Select
    
    Me.lblCustomername.Caption = strDetails
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayCustomerDetails"
End Sub



Private Function GetLastFileNum() As Long
    On Error GoTo errHandler
Dim lNum As Long, lTmp As Long
Dim sFile As String

    sFile = Dir(App.Path & "\*.sbk")
    Do While sFile <> ""
        lTmp = val(Mid(sFile, 5, Len(sFile) - 5))
        If lNum < lTmp Then lNum = lTmp
        sFile = Dir
    Loop
    GetLastFileNum = lNum + 1
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetLastFileNum"
End Function

Private Sub LoadHelp()
    On Error GoTo errHandler
Dim fHelp As New frmHelp
    fHelp.Show vbModal
    Set fHelp = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadHelp"
End Sub


Private Sub CreateBill()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.CreateBill"
End Sub

Private Sub FetchCOLS()
    On Error GoTo errHandler
    Set cCOLS = New C_COLS
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
            X3.Value(i, 3) = .code
            X3.Value(i, 5) = .Description
            X3.Value(i, 4) = .Qty & "(" & .QtyDispatched & ")"
            X3.Value(i, 6) = .DepositF
            X3.Value(i, 7) = .DepositStatus
            X3.Value(i, 8) = .PriceF
            X3.Value(i, 9) = .DiscountRateF
            X3.Value(i, 10) = .COLDATEForSort
            X3.Value(i, 11) = .COLID
            X3.Value(i, 12) = .Deposit
            X3.Value(i, 13) = .PID
        End With
    Next
  '  X3.QuickSort 1, X3.UpperBound(1), 10, XORDER_DESCEND, XTYPE_STRING   'sorted in query - we need the ordinal poistion (column 1) to be in sequence
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
   '     Me.G1.Height = 180
   '     Me.frSaleTotal.Top = 230
        G3.Visible = True
    Else
   '     Me.G1.Height = 265
   '     Me.frSaleTotal.Top = 312
        G3.Visible = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetForCOLSVisible(pYes)", pYes
End Sub
Private Sub LoadSaleRow(iIndex As Integer)
    On Error GoTo errHandler
Dim i As Long
    G1.Visible = True
    X1.Value(iIndex, 1) = oExchange.SaleLines(iIndex).CodeF
    X1.Value(iIndex, 2) = oExchange.SaleLines(iIndex).Title & " (" & oExchange.SaleLines(iCurrentSaleLine).MainAuthor & ")"
    X1.Value(iIndex, 3) = oExchange.SaleLines(iIndex).Qty
    X1.Value(iIndex, 4) = oExchange.SaleLines(iIndex).PriceF
    X1.Value(iIndex, 5) = oExchange.SaleLines(iIndex).DiscountRateF
    X1.Value(iIndex, 6) = oExchange.SaleLines(iIndex).PLessDiscExtF
    X1.Value(iIndex, 7) = oExchange.SaleLines(iIndex).PLessDiscExtVATF & "(" & oExchange.SaleLines(iCurrentSaleLine).VATRateF & ")"
    G1.Array = X1
    Me.G1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadSaleRow(iIndex)", iIndex
End Sub
Private Sub DisplayTotals()
    On Error GoTo errHandler
    txtExtTotal = oExchange.TotalPayableF '"Total: " & oExchange.TotalPayableF
    txtQtyTotal = oExchange.TotalQty   '"(" & oExchange.TotalQty & " " & IIf(oExchange.TotalQty > 1, "Items", "Item") & ")"
    txtVatValue = oExchange.TotalVATF   '"Includes VAT of:  " & oExchange.TotalVATF
    txtPaymentTotal = "Total paid: " & oExchange.TotalPaymentF
    DisplayCustomerDetails
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayTotals"
End Sub
Private Sub LoadPaymentRow(iIndex As Integer)
    On Error GoTo errHandler
Dim i As Long
    G2.Visible = True
    X2.Value(iIndex, 1) = oExchange.PaymentLines(iIndex).PaymentTypeF
    X2.Value(iIndex, 2) = oExchange.PaymentLines(iIndex).AmtF
    X2.Value(iIndex, 3) = oExchange.PaymentLines(iIndex).PaymentType
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
    DisplayTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RefreshSaleDisplay"
End Sub
Public Sub SendPOSExchange(pEXCHID As String, pOPSID As String, pZID As String)
    On Error GoTo errHandler
Dim MSG As String
Dim sFileName As String
Dim oShapeDB As New z_POSCLIConnectionShape
Dim sSQL As String
Dim tmprs As ADODB.Recordset

    Check (oShapeDB.dbConnecttoShape = 0), ERR_GENERAL, "Failed to create database connection!"
    If Not rsZSession Is Nothing Then
        If rsZSession.State <> 0 Then
            rsZSession.Close
        End If
    End If
    sSQL = "SHAPE {SELECT 'E' as TYP,tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession APPEND (( SHAPE {SELECT EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE,EXCH_SALEVALUE,EXCH_DISCOUNTVALUE,EXCH_VATVALUE,EXCH_CHANGEGIVEN,EXCH_LOYALTYVALUE,EXCH_TYPE,EXCH_SUPERVISORID,EXCH_NUMBER,EXCH_VOIDS,EXCH_CUSTOMERNAME FROM tEXCHANGE WHERE EXCH_ID = '" & pEXCHID & "'}  AS POSExchange APPEND ({SELECT * FROM tCSL}  AS rsSALESLINES RELATE EXCH_ID TO CSL_EXCH_ID) AS SALESLINES,({SELECT * FROM tPayment}  AS rsPAYMENTS RELATE EXCH_ID TO PAY_EXCH_ID) AS PAYMENTS) AS POSExchange RELATE OPS_ID TO EXCH_OPSESSIONID) AS POSExchange) AS OPSession RELATE Z_ID TO OPS_Z_ID) AS OPSession"
    Set rsZSession = Nothing
    Set rsZSession = New ADODB.Recordset
    rsZSession.Open sSQL, oShapeDB.DBConn, adOpenStatic
    Set rsZSession.ActiveConnection = Nothing
    
    If Not rsZSession.EOF Then
        sFileName = oPC.StationName & "-" & Format(Now(), "DDHHNNSS") 'Format(oGD.GetNextFileNum(), "00000")
        sFileName = "\" & sFileName & ".pos"
        rsZSession.Save oPS.ClientOutbox & sFileName, adPersistADTG
    End If
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
'    Set ADOConn = New ADODB.Connection
'    If ADOConn.State <> adStateOpen Then
'        ADOConn.Provider = "sqloledb"
'        ADOConn.Open "Data Source=" & oPC.LocalSQLServerName & ";Initial Catalog=PBKSFD;User Id=sa;Password=; Network Library=dbmssocn;"
'    End If
'    LoadTriggers
'    strServerMachineName = GetIniKeyValue(strLocalPath & "\PBKS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
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
    ErrorIn "frmPOSMain.PrintSalesSlip"
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
    ErrorIn "frmPOSMain.PrintSalesSlip"
End Sub

Private Sub PrintCNasChange(pAMt As Long, pCopyCount As Integer, Optional bReprint As Boolean)
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
            PrintHeader ConvertToType("C"), OPOSPOSPrinter1, bReprint      'Print header
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
                
            With OPOSPOSPrinter1
                
                
                sBuf = "Credit note"
                sExt = Format(pAMt, "Currency")
                sValue = MakePrintString(.RecLineChars, sBuf, sExt)
                .PrintNormal PTR_S_RECEIPT, sValue + vbLf
            End With
        
        
            PrintFooter c, ConvertToType("C"), OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            .AsyncMode = False
        End With
    Next

    MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintCNasChange"
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
           
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + "Petty Cash: " & oExchange.PaymentLines(1).AmtF + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
            .PrintNormal PTR_S_RECEIPT, "Reference: " & oExchange.PaymentLines(1).ReferenceDescription + vbLf
            .PrintNormal PTR_S_RECEIPT, "Reason: " & oExchange.PaymentLines(1).Note + vbLf
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
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
    ErrorIn "frmPOSMain.PrintLoyaltyVoucher"
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
            PrintHeader eTypDeposit, OPOSPOSPrinter1           'Print header
            .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
           
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "Deposit Refunded: " & oExchange.PaymentLines(1).AmtF + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & strDepositTitle
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
          '  .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(i)
            PrintFooter i, etypDepositRefund, OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
        End With
    Next i

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Exit Sub


End Sub
Private Sub PrintDepositSlip(pCopyCount As Integer)
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
            PrintHeader eTypDeposit, OPOSPOSPrinter1           'Print header
            .PrintNormal PTR_S_RECEIPT, ESC + "|600uF"      'create gap
           
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + ESC + "|bC" + ESC + "|2C" + "Deposit paid: " & oExchange.PaymentLines(1).AmtF + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & strDepositTitle
            .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
            .PrintNormal PTR_S_RECEIPT, ESC + "|100uF" & "Copy number: " & CStr(i)
            PrintFooter i, eTypDeposit, OPOSPOSPrinter1          'print footer
            
            .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go
    
            'Back to the synchronous mode
            .AsyncMode = False
        End With
    Next i

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintDepositSlip"

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
    Select Case oExchange.transactionTypeEnum
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
    Case eDepositType
            lngSalesItemCount = lngSalesItemCount + 1
            X4.InsertRows (lngSalesItemCount)
            X4.Value(lngSalesItemCount, 1) = oPC.ExchangeNumber - 1  'lngSalesItemCount
            X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
            X4.Value(lngSalesItemCount, 4) = oExchange.transactionType
            X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
            X4.Value(lngSalesItemCount, 5) = ""
            X4.Value(lngSalesItemCount, 6) = "1"
            X4.Value(lngSalesItemCount, 7) = "DEPOSIT"
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
            X4.Value(lngSalesItemCount, 10) = oExchange.ExchangeID
            X4.Value(lngSalesItemCount, 13) = oExchange.ToVoid
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
            X4.Value(lngSalesItemCount, 7) = "PETTY CASH" & ":" & oExchange.PaymentLines(1).PaymentTypeF & ":" & oExchange.PaymentLines(1).ReferenceComplete
            X4.Value(lngSalesItemCount, 8) = ""
            X4.Value(lngSalesItemCount, 9) = oExchange.PaymentLines(1).AmtF
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
    ErrorIn "frmPOSMain.G4_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle)
End Sub

Private Sub LoadExchanges()
    On Error GoTo errHandler
Dim ZID As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    
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
    Case eTypReceipt
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
            
            sBuf = "Customer's payment"
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
            sBuf = "Change"
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, (.RecLineChars)) + vbLf
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, (.RecLineChars)) + vbLf
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, .RecLineChars) + vbLf
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, (.RecLineChars)) + vbLf
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, (.RecLineChars)) + vbLf
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, (.RecLineChars)) + vbLf
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
            ElseIf oExchange.CustomerName > "" Then
                .PrintNormal PTR_S_RECEIPT, Left(oExchange.CustomerName, (.RecLineChars)) + vbLf
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
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintHeader(eDocumentType,pPrinter)", Array(eDocumentType, pPrinter)
End Sub
Private Sub PrintFooter(pCopyNumber As Integer, eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter)
    On Error GoTo errHandler
Dim ar() As String
Dim i As Integer
Dim sValue As String
    Select Case eDocumentType
    Case eTypReceipt, eTypCashRefund, etypCreditNote, eTypDeposit, etypDepositRefund, eTypPettyCash, eTypAppro
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
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSemailAddress + vbLf
            If pCopyNumber > 1 Then
                .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "Copy number: " & CStr(pCopyNumber) + vbLf
            End If
            .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
        End With
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.PrintFooter(eDocumentType,pPrinter)", Array(eDocumentType, pPrinter)
End Sub
Private Function ConvertToType(val As String) As Integer
    On Error GoTo errHandler
    Select Case val
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
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ConvertToType(val)", val
End Function
Private Function validVoucherCode(pCODE As String) As Boolean
validVoucherCode = True
End Function
