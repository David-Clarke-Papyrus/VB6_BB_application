VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{DA4E6F7B-F5EE-43C5-A9A1-6BCC726F901E}#1.8#0"; "StatusBarX5.OCX"
Object = "{C9E1AFB0-1172-11D7-83AD-0050DA238ADA}#1.0#0"; "Coptr17.ocx"
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Begin VB.Form frmPOSMain 
   BackColor       =   &H00DFDED2&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Papyrus II    Point of Sale"
   ClientHeight    =   8115
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11655
   Icon            =   "frmPOSMain2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   541
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   777
   StartUpPosition =   1  'CenterOwner
   Begin IOLib.IO IO1 
      Left            =   6060
      Top             =   6180
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin TrueOleDBGrid60.TDBGrid G4 
      Height          =   5835
      Left            =   15
      OleObjectBlob   =   "frmPOSMain2.frx":08CA
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   11130
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0FFC0&
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
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6795
      Width           =   885
   End
   Begin VB.TextBox txtPaymentTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DFDED2&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   390
      Left            =   8310
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1290
   End
   Begin StatusBarXCtl.StatusBarX SB 
      Height          =   450
      Left            =   90
      Top             =   7665
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   794
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PanelCount      =   1
      Panel1Key       =   "test"
      Panel1BackColor =   13882315
      Panel1ForeColor =   7884871
      Panel1WordWrap  =   -1  'True
      Panel1Width     =   752
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
      Height          =   885
      Left            =   255
      TabIndex        =   1
      Top             =   6465
      Width           =   5385
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3495
      Left            =   15
      OleObjectBlob   =   "frmPOSMain2.frx":4575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   11445
   End
   Begin VB.Frame frSaleTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDED2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   6615
      TabIndex        =   3
      Top             =   3630
      Width           =   4815
      Begin VB.TextBox txtExtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DFDED2&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00714942&
         Height          =   390
         Left            =   1455
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   165
         Width           =   1440
      End
      Begin VB.TextBox txtQtyTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DFDED2&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00714942&
         Height          =   390
         Left            =   150
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   165
         Width           =   705
      End
      Begin VB.TextBox txtVatValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00DFDED2&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00714942&
         Height          =   390
         Left            =   3210
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   1470
      End
   End
   Begin TrueOleDBGrid60.TDBGrid G2 
      Height          =   1380
      Left            =   6735
      OleObjectBlob   =   "frmPOSMain2.frx":8FBC
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4680
      Width           =   4755
   End
   Begin TrueOleDBGrid60.TDBGrid G3 
      Height          =   1200
      Left            =   30
      OleObjectBlob   =   "frmPOSMain2.frx":C517
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4695
      Visible         =   0   'False
      Width           =   6255
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
      Left            =   3225
      TabIndex        =   15
      Top             =   3720
      Width           =   1965
   End
   Begin VB.Label lblCustomer 
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
      Left            =   45
      TabIndex        =   14
      Top             =   3720
      Width           =   3225
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
      Height          =   510
      Left            =   1845
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   75
      TabIndex        =   11
      Top             =   -225
      Width           =   6525
   End
   Begin COPTRLib.OPOSPOSPrinter OPOSPOSPrinter1 
      Left            =   6270
      Top             =   4695
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
      Top             =   6075
      Width           =   3555
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
End Enum
Dim ESC As String
Dim WithEvents oExchange As a_Exchange
Attribute oExchange.VB_VarHelpID = -1
Public WithEvents oPS As z_PollingServices_Client
Attribute oPS.VB_VarHelpID = -1
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2
Dim ADOConn As ADODB.Connection
Dim frmExchange As frmExchange
Dim enPresentState As eState
Dim oCurrLine As ListItem
Dim flgSaleActive As Boolean
Dim flgGDiscount As Boolean
Dim flgNewCode As Boolean
Dim flgEditItem As Boolean
Dim flgReturn As Boolean
Dim flgInvalidLine As Boolean
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
Dim strCustomerName As String
Dim lngCustomerID As Long
Dim sPaymentType As String
Dim iCurrentSaleLine As Integer
Dim iCurrentPaymentLine As Integer
Dim iCurrentCOL As Integer
Dim strName As String
Dim lngStaffID As Long
Dim cCOLS As C_COLS
Dim oSALELine As a_Sale
Dim oPAYMENTLine As a_Payment
Dim strOperator As String
Dim bConnected As Boolean
Dim bCloseZSession As Boolean
Dim lngSMID As Long
Dim lngSalesItemCount As Long
Dim iToVoid As Long
Dim lngBalanceOwing As Long
Dim bLoyaltyCustomer As Boolean
Private Type ITEMDATA
    TType As String
    Name As String
    Ext As String
    At As String
End Type

Private Sub ShowTransactions(bShow As Boolean)
    If bShow Then
        G4.Visible = True
        G1.Visible = False
    Else
        G4.Visible = False
        G1.Visible = True
    End If
End Sub
Private Sub oPS_UpdatingLocalDatabase(bOn As Boolean)
Static strMsg As String
    If bOn Then
        strMsg = SB.Panels(1).Text
        SB.Panels(1).Text = "Updating local database . . ."
    Else
        SB.Panels(1).Text = strMsg
    End If
End Sub

Private Sub oPS_LostServer(pMsg As String)
    bConnected = False
    lblStatus.Caption = pMsg
    lblStatus.ForeColor = vbRed
End Sub



Private Sub Command1_Click()
Dim result As Integer

'    result = IO1.Close
    result = IO1.Open(oPC.cashDrawerPort, "baud=9600 parity=N data=8 stop=1")  'Set up scanner
                    
    result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
    
    result = IO1.Close

End Sub

Private Sub G4_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If X4(Bookmark, 7) = -1 Then
        RowStyle.BackColor = RGB(192, 192, 192)
    ElseIf X4(Bookmark, 8) > 0 Then
        RowStyle.BackColor = RGB(176, 222, 173)
    Else
        RowStyle.BackColor = &HFFFFFF
    End If
End Sub

Private Sub oPS_HasServer()
    bConnected = True
    lblStatus.Caption = "Connected"
    lblStatus.ForeColor = &H80000012
End Sub


Private Sub oExchange_ContainsLines(pYesNo As Boolean)
    On Error GoTo errHandler
    flgSaleActive = pYesNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oExchange_ContainsLines(pYesNo)", pYesNo, EA_NORERAISE
    HandleError
End Sub
Private Sub SetTitleBar()
    On Error GoTo errHandler
    Caption = "Papyrus Point-of-Sale       " & oPC.NameOfPC & "      Supervisor: " & oPC.ZSession.Supervisorname & "/" & strOperator & "              #" & oExchange.ExchangeNumber
    lblStatus.Caption = "Sales for " & oPC.ZSession.NominalDateF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetTitleBar"
End Sub

Private Sub cmdClose_Click()
Dim frm As frmPOSHELP
    Set frm = New frmPOSHELP
    frm.Show
End Sub

Private Sub Form_Load()
Dim result As Integer

    On Error GoTo errHandler
    flgSaleActive = False
    ESC = Chr(27)
    'Try to load local DB connection
    If oPC Is Nothing Then
        Set oPC = New z_POSCLIConnection
        oPC.dbConnect
    End If

    Set oPS = New z_PollingServices_Client
    Check oPS.TryToStartPolling, EXC_SERVERUNAVAILABLE, "Cannot poll server"
    LogonOperator

    oPC.SetupZSession lngStaffID, strName
    oPC.ZSession.OpSession.START_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
    StartSale
    SetTitleBar
    With OPOSPOSPrinter1
        .Open oPC.Printername
        .ClaimDevice 1000
        .DeviceEnabled = True
        .MapMode = PTR_MM_METRIC
        .RecLetterQuality = True
        
'        If .CapRecBitmap = True Then
'            .SetBitmap 1, PTR_S_SLIP, App.Path + "\Logo.bmp", .RecLineWidth / 2, PTR_BM_CENTER
'        End If
    End With
    X4.Clear
    X4.ReDim 1, 1, 1, 8
    LoadExchanges
    SetForCOLSVisible False
    result = IO1.Open(oPC.cashDrawerPort, "baud=9600 parity=N data=8 stop=1")  'Set up scanner

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Function LogonOperator() As Boolean
    On Error GoTo errHandler
Dim bCancelled As Boolean
            
    If SecurityControl(2, lngStaffID, strName, bCancelled, "Enter your security key.", "Your key is invalid") Then
        strOperator = strName
        bLoggedOn = True
    Else
       ' LockAll True
        enPresentState = elogin
    End If
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
        oPC.ZSession.OpSession.START_OP_Session oPC.ZSession.Current_Z_Session_ID, lngStaffID
        strOperator = strName
        bLoggedOn = True
    Else
      '  LockAll True
        enPresentState = elogin
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
        If oPC.ZSession.OpSession.InSession Then
            oPC.ZSession.OpSession.Close_OP_Session
        End If
        If oPC.ZSession.InSession Then
            oPC.ZSession.Close_Z_Session
        End If
        Unload Me
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.cmdZTotal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
'    If Me.txtInput.Visible And Me.txtInput.Enabled Then Me.txtInput.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Activate", , EA_NORERAISE
    HandleError
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
    ErrorIn "frmPOSMain.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
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

Private Sub StartSale()
    On Error GoTo errHandler
'    flgSaleActive = True
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    iCurrentSaleLine = 0
    SetState eProductID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.StartSale"
End Sub



Private Sub Stat(Msg As String)
    On Error GoTo errHandler
    SB.Panels(1).Text = Msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Stat(msg)", Msg
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim result As Integer

    On Error GoTo errHandler
    If flgSaleActive Then
        If MsgBox("There is still a transaction in process!" & vbLf & _
                  "Do you want to close this Application anyway?", _
                  vbYesNo, "Transaction In Process!") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    Else
        RejectSale
        If MsgBox("Closing Papyrus POS application?", vbYesNo + vbQuestion, "Close?") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    oPS.RegisterWithServer (False)
    oPC.ZSession.OpSession.Close_OP_Session
    If oExchange.IsEditing Then oExchange.CancelEdit
    If bCloseZSession Then
        oPC.ZSession.Close_Z_Session
    End If
    If bConnected Then
        Screen.MousePointer = vbHourglass
        Me.SB.Panels(1).Text = "Wait. The local data is being transmitted to the server."
        EventPause 5
        Screen.MousePointer = vbDefault
    End If
    
    With OPOSPOSPrinter1
        .DeviceEnabled = False
        .ReleaseDevice
        .Close
    End With
    result = IO1.Close

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
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


Private Sub oPS_PollingStoped(Msg As String)
    On Error GoTo errHandler
    If MsgBox("Automatic file transfer stopped!" & vbLf & _
               "Reason: " & Msg & vbLf & vbLf & _
               "Click YES to restart it.", vbYesNo + vbExclamation, "File Transfer Stopped!") = vbYes Then
        
        oPS.StartPolling
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.oPS_PollingStoped(msg)", Msg, EA_NORERAISE
    HandleError
End Sub

Private Sub txtInput_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtInput
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowExchange()
Dim lngRow As Long
Dim lngTmp As Long

    Set frmExchange = New frmExchange
    If IsNumeric(txtInput) Then
        lngRow = CLng(txtInput)
        If lngRow <= X4(X4.UpperBound(1) - 1, 1) And lngRow > 0 Then
            lngTmp = X4.Find(1, 1, lngRow, , , XTYPE_LONG)
            If lngTmp > 0 Then
                frmExchange.Component X4(lngTmp, 5)
                frmExchange.Show vbModal
                Unload frmExchange
            End If
        End If
    End If
End Sub
Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
Dim sTmp As String
Dim i As Integer
Dim enRequestState As eState
Dim itmp As Integer
Dim strArg As String
Dim strArg2 As String
Dim lngRow As Long
Dim strTmp As String
Dim lngTmp As Long

    txtInput = Trim(txtInput)

    
    Select Case KeyCode
    Case 13
'        If UCase(Left(txtInput, 2)) = "SH" Then  'Show details of this exchange
'        Else
'            ShowTransactions False
'        End If
        enRequestState = enull
        If UCase(txtInput) = "P" Then
            enRequestState = eProductID
        End If
'''''''''''
        If Left(UCase(txtInput), 2) = "VR" Then
            strTmp = Right(txtInput, Len(txtInput) - 2)
            If IsNumeric(strTmp) Then
                iToVoid = CLng(strTmp)
                If iToVoid >= CLng(X4(1, 1)) And iToVoid < oExchange.ExchangeNumber Then
                    enRequestState = eVoidAndReplace
                End If
            End If
        End If
        If UCase(txtInput) = "X" Then
            enRequestState = eCancelSale
        End If
        If UCase(txtInput) = "F" Then
            If (enPresentState = eCashRefund) And oExchange.SaleLines.Count > 0 Then
                txtInput = "OK"
                txtInput.PasswordChar = "*"
                AutoSelect txtInput
                enRequestState = eConfirmationCashrefund
            ElseIf (enPresentState = eCreditNote) And oExchange.SaleLines.Count > 0 Then
                txtInput = "OK"
                txtInput.PasswordChar = "*"
                AutoSelect txtInput
                enRequestState = eConfirmationCreditNote
            ElseIf (enPresentState = eDeposit) And oExchange.SaleLines.Count > 0 Then
                txtInput = "OK"
                txtInput.PasswordChar = "*"
                AutoSelect txtInput
                enRequestState = eConfirmationDeposit
            ElseIf (enPresentState = eProductID) And oExchange.SaleLines.Count > 0 Then
                enRequestState = eConfirmation
            End If
        End If

        If UCase(Right(txtInput, 1)) = "/" Then
            enRequestState = eDiscount
        End If
        If UCase(Left(txtInput, 2)) = "CX" Then
            enRequestState = eSearchCustomer
            strArg = Right(Trim(txtInput), Len(Trim(txtInput)) - 1)
            strArg2 = "Name"
        End If
        If UCase(Left(txtInput, 3)) = "CX/" Then
            enRequestState = eSearchCustomer
            strArg = Right(Trim(txtInput), Len(Trim(txtInput)) - 2)
            strArg2 = "Phone"
        End If
        If UCase(txtInput) = "L" Then
            enRequestState = elogin
        End If
        If UCase(txtInput) = "R" Then
            enRequestState = eCashRefund
        End If
        If UCase(txtInput) = "DEP" Then
            enRequestState = eDeposit
        End If
        If UCase(txtInput) = "CN" Then
            Select Case enPresentState
                Case eProductID
                    enRequestState = eCreditNote
                Case Else
                    enRequestState = ePaymentType_CreditNote
            End Select
        End If
        If UCase(txtInput) = "S" Then
            enRequestState = eProductID
        End If
        If UCase(txtInput) = "H" Then
            enRequestState = eHelp
        End If
        If UCase(txtInput) = "RBI" Then
            enRequestState = eRebuildIndexes
        End If
        If UCase(txtInput) = "Z" Then
            enRequestState = eZTerminate
        End If
        If UCase(txtInput) = ".." Then
            enRequestState = ePrevious
        End If
        If UCase(txtInput) = "C" Then
            enRequestState = ePaymentType_Cash
        End If
        If UCase(txtInput) = "Q" Then
            enRequestState = ePaymentType_Cheque
        End If
        If UCase(txtInput) = "V" Then
            enRequestState = ePaymentType_voucher
        End If
        If UCase(txtInput) = "CC" Then
            enRequestState = ePaymentType_CreditCard
        End If
        If UCase(txtInput) = "CD" Then
            enRequestState = ePaymentType_CustomerDeposit
        End If
        If UCase(txtInput) = "DD" Then
            ShowTransactions True
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
        Statechange enRequestState, itmp, strArg, strArg2
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub
Private Sub SetState(pState As eState)
    On Error GoTo errHandler
    enPresentState = pState
    Select Case pState
        Case eReviewExchanges
            lblInput.Caption = "Reviewing exchanges"
            Stat "Line number to print, DD to end review."
        Case eCashRefund
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            If flgSaleActive Then
                lblInput.Caption = "Product code, payment code or special."
                Stat "(Dn)Del prod"
            Else
                lblInput.Caption = "Start cash refund "
                Stat "Start cash refund by entering product code,   .. to reverse"
            End If
        Case eCreditNote
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            If flgSaleActive Then
                lblInput.Caption = "Product code, payment code or special."
                Stat "(Dn)Del prod"
            Else
                lblInput.Caption = "Start credit note "
                Stat "Start cash refund by entering product code,   .. to reverse"
            End If
        Case ePriceCashRefund, ePriceCreditNote
            lblInput.Caption = "Price "
            Stat "'.. to reverse"
        Case eQtyCashRefund, eQtyCreditNote
            lblInput.Caption = "Qty "
            Stat "'.. to reverse"
            
        Case eDeposit
            lblInput.ForeColor = vbRed
            txtInput.ForeColor = vbRed
            lblInput.Caption = "Start deposit "
            Stat "'.. to reverse"
        Case eProductID
            lblInput.ForeColor = &H714942
            txtInput.ForeColor = &H714942
            If flgSaleActive Then
                lblInput.Caption = "Product code, payment code or special."
                Stat "(Dn)Del prod, (DPn)Del paymt, (C)Cash, (V)Voucher, (CC)Card, (Q)cheque"
            Else
                lblInput.Caption = "Start sale"
                Stat "Start sale by entering product code, or (R)eturn, (Z) cash up."
            End If
        Case eSearchCustomer
            lblInput.Caption = "Search for . . . "
            Stat ""
        Case eQty
            lblInput.Caption = "Quantity"
            Stat "'.. to reverse"
        Case eDiscount
            lblInput.Caption = "Discount."
            Stat ""
        Case ePrice
            lblInput.Caption = "Price"
            Stat "Append '/' to price for discount capture, '..' to reverse"
        Case elogin
            lblInput.Caption = "Staff code."
        Case ePaymentType_Cash
            lblInput.Caption = "Cash received."
            Stat ""
        Case ePaymentType_CreditCard
            lblInput.Caption = "Credit card charge value."
            Stat "'.. to reverse"
        Case ePaymentType_CreditNote
            lblInput.Caption = "Credit note value."
            Stat "'.. to reverse"
        Case ePaymentType_voucher
            lblInput.Caption = "Credit voucher value."
            Stat "'.. to reverse"
        Case ePaymentType_Cheque
            lblInput.Caption = "Cheque value."
            Stat "'.. to reverse"
        Case ePaymentType_CreditCardRef
            lblInput.Caption = "Credit card reference."
            Stat "'.. to reverse"
        Case ePaymentType_voucherRef
            lblInput.Caption = "Voucher number."
            Stat "'.. to reverse"
        Case ePaymentType_ChequeRef
            lblInput.Caption = "Cheque reference."
            Stat "'.. to reverse"
        Case eConfirmation
            lblInput.Caption = "Confirm sale."
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.SetState(pState)", pState
End Sub
Private Sub DisplayProduct()
    On Error GoTo errHandler
    LoadSaleRow
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.DisplayProduct"
End Sub
Private Sub DisplayPayment()
    On Error GoTo errHandler
    LoadPaymentRow
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

Private Sub Statechange(pNewState As eState, Optional iRow As Integer, Optional pArg1 As String, Optional pArg2 As String)
Dim result As Integer

    On Error GoTo errHandler
    Select Case enPresentState
    Case eReviewExchanges
        Select Case pNewState
        Case eReviewExchanges
            ShowTransactions False
            SetState eProductID
            txtInput = ""
            AutoSelect txtInput
        Case Else
            ShowExchange
            txtInput = ""
        End Select
        
    Case ePaymentType_CreditNote
        Select Case pNewState
        
        Case ePrevious
            SetState eProductID
        Case Else
            If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
                oPAYMENTLine.setType "Credit Note"
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                Else
                    SetState eProductID
                    txtInput = ""
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_Cash
        Select Case pNewState
        Case ePrevious
            SetState eProductID
        Case Else
            If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                Else
                    SetState eProductID
                    txtInput = ""
                End If
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_Cheque
        Select Case pNewState
        Case ePrevious
            SetState eProductID
        Case Else
            If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                SetState ePaymentType_ChequeRef
                result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
                txtInput = ""
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_CreditCard
        Select Case pNewState
        Case ePrevious
            SetState eProductID
        Case Else
            If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                SetState ePaymentType_CreditCardRef
                result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
                txtInput = ""
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_voucher
        Select Case pNewState
        Case ePrevious
            SetState eProductID
        Case Else
            If oPAYMENTLine.SetAmt(Trim(txtInput)) Then
                oExchange.CalculateTotals
                SetState ePaymentType_voucherRef
                result = IO1.WriteString("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" & Chr(13) & Chr(10))
                txtInput = ""
            Else
                SetTip "Invalid payment amount."
            End If
            DisplayPayment
        End Select
    Case ePaymentType_ChequeRef
        Select Case pNewState
        Case ePrevious
            SetState ePaymentType_Cheque
            txtInput = CStr(oPAYMENTLine.Amt)
            AutoSelect txtInput
        Case Else
            If oPAYMENTLine.setReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                Else
                    SetState eProductID
                    txtInput = ""
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
    Case ePaymentType_CreditCardRef
        Select Case pNewState
        Case ePrevious
            SetState ePaymentType_CreditCard
            txtInput = CStr(oPAYMENTLine.Amt)
            AutoSelect txtInput
        Case Else
            If oPAYMENTLine.setReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                Else
                    SetState eProductID
                    txtInput = ""
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
    Case ePaymentType_voucherRef
        Select Case pNewState
        Case ePrevious
            SetState ePaymentType_voucher
            txtInput = CStr(oPAYMENTLine.Amt)
            AutoSelect txtInput
        Case Else
            If oPAYMENTLine.setReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                Else
                    SetState eProductID
                    txtInput = ""
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
    Case ePaymentType_CustomerDeposit
        Select Case pNewState
        Case ePrevious
            SetState ePaymentType_CustomerDeposit
            txtInput = CStr(oPAYMENTLine.Amt)
            AutoSelect txtInput
        Case Else
            If oPAYMENTLine.setReference(Trim(txtInput)) Then
                oExchange.CalculateTotals
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                Else
                    SetState eProductID
                    txtInput = ""
                End If
                DisplayPayment
            Else
                SetTip "Invalid Reference."
            End If
        End Select
    Case enull
            Select Case pNewState
            Case elogin
                txtInput = ""
                SetState elogin
                LogonOperator
            End Select
    Case eCashRefund
        Select Case pNewState
        Case eConfirmationCashrefund
            SetState eConfirmationCashrefund
        Case ePrevious
            txtInput = ""
            AutoSelect txtInput
            SetState eProductID
        Case Else
            If LoadProductFromCode Then
                oExchange.CalculateTotals
                DisplayProduct
                SetState ePriceCashRefund
                txtInput = oSALELine.Price
                AutoSelect txtInput

            Else
                SetTip "Not on database."
            End If
        End Select
    Case eCreditNote
        Select Case pNewState
        Case eConfirmationCreditNote
            SetState eConfirmationCreditNote
        Case ePrevious
            txtInput = ""
            AutoSelect txtInput
            SetState eProductID
        Case Else
            If LoadProductFromCode Then
                oExchange.CalculateTotals
                DisplayProduct
                SetState ePriceCreditNote
                txtInput = oSALELine.Price
                AutoSelect txtInput

            Else
                SetTip "Not on database."
            End If
        End Select
    Case eDeposit
        Select Case pNewState
        Case eConfirmationDeposit
            SetState eConfirmationDeposit
'        Case eConfirmationDeposit
'            SetState eConfirmationDeposit
        Case ePrevious
            txtInput = ""
            AutoSelect txtInput
            SetState eProductID
        Case Else
            If LoadProductFromCode Then
                oExchange.CalculateTotals
                DisplayProduct
                SetState ePriceDeposit
                txtInput = oSALELine.Price
                AutoSelect txtInput

            Else
                SetTip "Not on database."
            End If
        End Select
    Case eProductID
            Select Case pNewState
            Case eReviewExchanges
                txtInput = ""
                SetState pNewState
            Case eVoidAndReplace
                'Show form as capturing a replacement
                lblReplacement.Caption = "Voiding and replacing Transaction #" & CStr(iToVoid)
                lblReplacement.Visible = True
                txtInput = ""
            Case eCashRefund
                If flgSaleActive = True Then
                    If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                        txtInput = ""
                        AutoSelect txtInput
                        SetState pNewState
                    End If
                Else
                    txtInput = ""
                    oExchange.transactionType = "R"
                    SetState pNewState
                End If
            Case eCreditNote
                If flgSaleActive = True Then
                    If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                        txtInput = ""
                        AutoSelect txtInput
                        SetState pNewState
                    End If
                Else
                    txtInput = ""
                    oExchange.transactionType = "C"
                    SetState pNewState
                End If
            Case eDeposit
                If flgSaleActive = True Then
                    If MsgBox("Cancel current transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                        txtInput = ""
                        AutoSelect txtInput
                        SetState pNewState
                    End If
                Else
                    txtInput = ""
                    oExchange.transactionType = "D"
                    SetState pNewState
                End If
            Case eCancelSale
                If MsgBox("Cancel this sale?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                    RejectSale
                    AutoSelect txtInput
                End If
            Case eRebuildIndexes
                    Connect
                    RebuildIndexes
                    Disconnect
            Case eSearchCustomer
                    If Not GetCustomer(pArg1, pArg2) Then
                        FetchCOLS
                        LoadCOLS
                        oExchange.CalculateTotals
                        txtInput = ""
                    End If
            Case ePaymentType_Cash
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.ApplyEdit
                    oPAYMENTLine.BeginEdit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.setType "C"
                    txtInput = CStr(oExchange.BalanceOwing)
                    AutoSelect txtInput
                    SetState ePaymentType_Cash
            Case ePaymentType_Cheque
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.ApplyEdit
                    oPAYMENTLine.BeginEdit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.setType "Q"
                    txtInput = CStr(oExchange.BalanceOwing)
                    AutoSelect txtInput
                    SetState ePaymentType_Cheque
            Case ePaymentType_CreditCard
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.ApplyEdit
                    oPAYMENTLine.BeginEdit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.setType "CC"
                    txtInput = CStr(oExchange.BalanceOwing)
                    AutoSelect txtInput
                    SetState ePaymentType_CreditCard
            Case ePaymentType_voucher
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.ApplyEdit
                    oPAYMENTLine.BeginEdit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.setType "V"
                    txtInput = CStr(oExchange.BalanceOwing)
                    AutoSelect txtInput
                    SetState ePaymentType_voucher
            Case ePaymentType_CustomerDeposit
                    Set oPAYMENTLine = oExchange.PaymentLines.Add
                    oPAYMENTLine.ApplyEdit
                    oPAYMENTLine.BeginEdit
                    iCurrentPaymentLine = iCurrentPaymentLine + 1
                    X2.ReDim 1, iCurrentPaymentLine, 1, 3
                    oPAYMENTLine.setType "V"
                    txtInput = CStr(oExchange.BalanceOwing)
                    AutoSelect txtInput
                    SetState ePaymentType_CustomerDeposit
            Case eZTerminate
                If MsgBox("Confirm cash up?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                    SetState eProductID
                Else
                    If oExchange.SaleLines.Count > 0 Then
                        oExchange.CancelEdit
                    End If
                    bCloseZSession = True
                    Unload Me
                End If
            Case elogin
                txtInput = ""
                enPresentState = elogin
                LogonOperator
            Case eDelete
                If iRow <= oExchange.SaleLines.Count And iRow > 0 Then
                    oExchange.SaleLines.Remove (iRow)
                    oExchange.SaleLines.ApplyEdit
                    oExchange.SaleLines.BeginEdit
                    oExchange.CalculateTotals
                    X1.DeleteRows (iRow)
                    G1.ReBind
                    SetState eProductID
                    txtInput = ""
                    ClearTextFields
                    iCurrentSaleLine = iCurrentSaleLine - 1
                End If
            Case eDeletePayment
                If iRow <= oExchange.PaymentLines.Count And iRow > 0 Then
                    oExchange.PaymentLines.Remove (iRow)
                    oExchange.PaymentLines.ApplyEdit
                    oExchange.PaymentLines.BeginEdit
                    oExchange.CalculateTotals
                    txtExtTotal = oExchange.TotalPaymentF
                    X2.DeleteRows (iRow)
                    G2.ReBind
                    SetState eProductID
                    txtInput = ""
                    iCurrentPaymentLine = iCurrentPaymentLine - 1
                End If
            Case eConfirmation
                If oExchange.PaymentsComplete Then
                    SetState eConfirmation
                    txtInput = "OK"
                    txtInput.PasswordChar = "*"
                    AutoSelect txtInput
                End If

            Case Else
                If LoadProductFromCode Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState ePrice
                    txtInput = oSALELine.Price
                    AutoSelect txtInput

                Else
                    SetTip "Not on database."
                End If
            End Select
    Case ePrice
            Select Case pNewState
            Case ePrevious
                enPresentState = eProductID
                txtInput = ""
                If oExchange.LoyaltyValue > 0 Then
                    lblCustomer.Caption = lblCustomer.Caption & "  ( " & oExchange.LoyaltyValueF & " )"
                End If
            Case eDiscount
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eDiscount
                    txtInput = ""
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQty
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceCashRefund
            Select Case pNewState
            Case ePrevious
                enPresentState = eCashRefund
                txtInput = ""
            Case eDiscountCashRefund
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eDiscountCashRefund
                    txtInput = ""
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQtyCashRefund
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceCreditNote
            Select Case pNewState
            Case ePrevious
                enPresentState = eCreditNote
                txtInput = ""
            Case eDiscountCreditNote
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eDiscountCreditNote
                    txtInput = ""
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQtyCreditNote
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case ePriceDeposit
            Select Case pNewState
            Case ePrevious
                enPresentState = eDeposit
                txtInput = ""
            Case eDiscountDeposit
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eDiscountDeposit
                    txtInput = ""
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            Case Else
                If oSALELine.SetPrice(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQtyDeposit
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid price."
                End If
            End Select
    Case eDiscount
                If oSALELine.SetDiscountRate(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQty
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid Discount."
                End If
    Case eDiscountCashRefund
                If oSALELine.SetDiscountRate(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQtyCashRefund
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid Discount."
                End If
    Case eDiscountCreditNote
                If oSALELine.SetDiscountRate(Trim(txtInput)) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eQtyCreditNote
                    txtInput = oSALELine.Qty
                    AutoSelect txtInput
                Else
                    SetTip "Invalid Discount."
                End If
    Case eQty
            Select Case pNewState
            Case ePrevious
                txtInput = ""
                SetState ePrice
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), False) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eProductID
                    oSALELine.ApplyEdit
                    oSALELine.BeginEdit
                    txtInput = ""
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyCashRefund
            Select Case pNewState
            Case ePrevious
                txtInput = ""
                SetState ePriceCashRefund
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), True) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eCashRefund
                    oSALELine.ApplyEdit
                    oSALELine.BeginEdit
                    txtInput = ""
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyCreditNote
            Select Case pNewState
            Case ePrevious
                txtInput = ""
                SetState ePriceCreditNote
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), True) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eCreditNote
                    oSALELine.ApplyEdit
                    oSALELine.BeginEdit
                    txtInput = ""
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case eQtyDeposit
            Select Case pNewState
            Case ePrevious
                txtInput = ""
                SetState ePriceCreditNote
            Case Else
                If oExchange.SaleLines(iCurrentSaleLine).SetQty(Trim(txtInput), True) Then
                    oExchange.CalculateTotals
                    DisplayProduct
                    SetState eCreditNote
                    oSALELine.ApplyEdit
                    oSALELine.BeginEdit
                    txtInput = ""
                Else
                    SetTip "Invalid qty."
                End If
            End Select
    Case elogin
            SwapOperator
            enPresentState = eProductID
    Case eConfirmation
            Select Case pNewState
            Case ePrevious
                SetState eProductID
                txtInput = ""
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    oExchange.SalesPersonName = strName
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel sale?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to issue a return. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
                txtInput.PasswordChar = ""
            End Select
            
    Case eConfirmationCashrefund
            Select Case pNewState
            Case ePrevious
                SetState eProductID
                txtInput = ""
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel cash refund?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to issue a cash refund. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
                txtInput.PasswordChar = ""
            End Select
    Case eConfirmationCreditNote
            Select Case pNewState
            Case ePrevious
                SetState eProductID
                txtInput = ""
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
                txtInput.PasswordChar = ""
            End Select
    Case eConfirmationDeposit
            Select Case pNewState
            Case ePrevious
                SetState eProductID
                txtInput = ""
            Case Else
                If GetLevel(txtInput, strName, lngSMID) > 0 Then
                    oExchange.SalesPersonID = lngSMID
                    AcceptSale
                ElseIf UCase(txtInput) = "XX" Then
                    If MsgBox("Confirm cancel deposit?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        RejectSale
                    End If
                Else
                    MsgBox "You do not have the authority to accept a deposit. Talk to your supervisor.", vbInformation + vbOKOnly, "Security"
                End If
                txtInput.PasswordChar = ""
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

Private Sub RejectSale()
    On Error GoTo errHandler
'    If Not oPAYMENTLine Is Nothing Then
'        If oPAYMENTLine.IsEditing Then oPAYMENTLine.CancelEdit
'    End If
    oExchange.CancelEdit
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
    oExchange.transactionType = "S"
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
    txtInput = ""
    lblReplacement.Visible = False
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    iToVoid = 0
    flgSaleActive = False
    bLoyaltyCustomer = False
    SetState eProductID
    SetTitleBar
    SetForCOLSVisible False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.RejectSale"
End Sub
Private Sub AcceptSale()
    On Error GoTo errHandler

'Save and send exchange
   ' oExchange.setExchangeType eSaleType
    oExchange.OperatorID = lngSMID
    oExchange.StaffName = strName
    If iToVoid > 0 Then oExchange.ToVoid = iToVoid
    oExchange.ApplyEdit
    AddExchange
    SendPOSExchange oExchange.ExchangeID, oExchange.OPSID, oExchange.ZID
 'Print Till Slip
    PrintSalesSlip
    
    
'Start new exchange
    Set oExchange = Nothing
    Set oExchange = New a_Exchange
    oExchange.BeginEdit
    oExchange.SalesPersonID = oPC.ZSession.OpSession.OperatorID
    oExchange.transactionType = "S"
    ClearTextFields
    X1.Clear
    X1.ReDim 1, 1, 1, 8
    G1.ReBind
    X2.Clear
    X2.ReDim 1, 1, 1, 3
    G2.ReBind
    txtInput = ""
    lblReplacement.Visible = False
    iCurrentSaleLine = 0
    iCurrentPaymentLine = 0
    bLoyaltyCustomer = False
    iToVoid = 0
    flgSaleActive = False
    SetState eProductID
    SetTitleBar
    SetForCOLSVisible False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.AcceptSale"
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
'    SetState eProductID
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
    oExchange.PaymentLines.ApplyEdit
    oExchange.PaymentLines.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.ClearPayments"
End Sub
'Private Sub LockAll(bLocked As Boolean)
'    On Error GoTo errHandler
'    Me.txtInput.Enabled = Not bLocked
''errHandler:
''    If ErrMustStop Then Debug.Assert False: Resume
''    ErrorIn "frmPOSMain.LockAll(bLocked)", bLocked
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOSMain.LockAll(bLocked)", bLocked
'End Sub



Private Function LoadProductFromCode() As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oGD As New z_GetData
    Set rs = oGD.GetProduct(Trim$(Me.txtInput))
    If Not rs Is Nothing Then
        Set oSALELine = oExchange.SaleLines.Add
 '       oSALELine.BeginEdit
        iCurrentSaleLine = iCurrentSaleLine + 1
        X1.ReDim 1, iCurrentSaleLine, 1, 8
       ' lstItems.ListItems.Add
        oSALELine.Title = NZS(rs!P_Title)
        oSALELine.MainAuthor = NZS(rs!P_MainAuthor)
        oSALELine.SetPrice FNN(rs!P_SAPrice)
        oSALELine.SetQty "1", False
        oSALELine.VATRate = FNDBL(rs!P_Vatrate)
        oSALELine.LoyaltyDiscount = FNN(rs!P_LCDiscount)
        oSALELine.PID = NZS(rs!P_ID)
        If NZS(rs!P_Code) > "" Then
            oSALELine.Code = NZS(rs!P_Code)
        Else
            oSALELine.Code = NZS(rs!P_EAN)
        End If
   '     CalculateAll
        oSALELine.ApplyEdit
        oSALELine.BeginEdit
        LoadProductFromCode = True
        rs.Close
        Set rs = Nothing
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadProductFromCode"
End Function
Private Sub oExchange_Recalculate()
    If bLoyaltyCustomer = True Then
        lblLoyaltyValue.Caption = oExchange.LoyaltyValueF
    End If
End Sub
Private Function GetCustomer(pArg1 As String, pArg2 As String)
    On Error GoTo errHandler
Dim frm As New frmBrowseCustomers2
    frm.Show vbModal
    strCustomerName = frm.CustomerName
    lngCustomerID = frm.CustomerID
    oExchange.SetCustomer lngCustomerID
    If frm.CustomerType = "L1" Then
        Me.lblCustomer = strCustomerName & " (Loyalty)"
        Me.lblLoyaltyValue.Caption = oExchange.LoyaltyValueF
        bLoyaltyCustomer = True
    Else
        lblCustomer = strCustomerName
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.GetCustomer(pArg1,pArg2)", Array(pArg1, pArg2)
End Function




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
    X3.ReDim 1, cCOLS.Count, 1, 9
    For i = 1 To cCOLS.Count
        With cCOLS(i)
            X3.Value(i, 1) = .COLDate
            X3.Value(i, 2) = .Code
            X3.Value(i, 4) = .Description
            X3.Value(i, 3) = .Qty & "(" & .QtyDispatched & ")"
            X3.Value(i, 5) = .DepositF
            X3.Value(i, 6) = .PriceF
            X3.Value(i, 7) = .DiscountRateF
            X3.Value(i, 8) = .COLDATEForSort
        End With
    Next
    X3.QuickSort 1, X3.UpperBound(1), 8, XORDER_DESCEND, XTYPE_STRING
    G3.Array = X3
    Me.G3.ReBind
    SetForCOLSVisible True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadCOLS"
End Sub
Private Sub SetForCOLSVisible(pYes As Boolean)
    If pYes Then
   '     Me.G1.Height = 180
   '     Me.frSaleTotal.Top = 230
        G3.Visible = True
    Else
   '     Me.G1.Height = 265
   '     Me.frSaleTotal.Top = 312
        G3.Visible = False
    End If
End Sub
Private Sub LoadSaleRow()
    On Error GoTo errHandler
Dim i As Long
    G1.Visible = True
    X1.Value(iCurrentSaleLine, 1) = oExchange.SaleLines(iCurrentSaleLine).CodeF
    X1.Value(iCurrentSaleLine, 2) = oExchange.SaleLines(iCurrentSaleLine).Title & " (" & oExchange.SaleLines(iCurrentSaleLine).MainAuthor & ")"
    X1.Value(iCurrentSaleLine, 3) = oExchange.SaleLines(iCurrentSaleLine).Qty
    X1.Value(iCurrentSaleLine, 4) = oExchange.SaleLines(iCurrentSaleLine).PriceF
    X1.Value(iCurrentSaleLine, 5) = oExchange.SaleLines(iCurrentSaleLine).DiscountRateF
    X1.Value(iCurrentSaleLine, 6) = oExchange.SaleLines(iCurrentSaleLine).PLessDiscExtF
    X1.Value(iCurrentSaleLine, 7) = oExchange.SaleLines(iCurrentSaleLine).PLessDiscExtVATF & "(" & oExchange.SaleLines(iCurrentSaleLine).VATRateF & ")"
    G1.Array = X1
    Me.G1.ReBind
    txtExtTotal = oExchange.TotalPayableF
    txtQtyTotal = oExchange.TotalQty
    Me.txtVatValue = oExchange.TotalVATF
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadSaleRow"
End Sub

Private Sub LoadPaymentRow()
    On Error GoTo errHandler
Dim i As Long
    G2.Visible = True
    X2.Value(iCurrentPaymentLine, 1) = oExchange.PaymentLines(iCurrentPaymentLine).PaymentTypeF
    X2.Value(iCurrentPaymentLine, 2) = oExchange.PaymentLines(iCurrentPaymentLine).AmtF
    X2.Value(iCurrentPaymentLine, 3) = oExchange.PaymentLines(iCurrentPaymentLine).PaymentType
    G2.Array = X2
    G2.ReBind
    Me.txtPaymentTotal = oExchange.TotalPaymentF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOSMain.LoadPaymentRow"
End Sub

Public Sub SendPOSExchange(pEXCHID As String, pOPSID As String, pZID As String)
    On Error GoTo errHandler
Dim Msg As String
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
    sSQL = "SHAPE {SELECT 'E' as TYP,tZSession.* FROM tZSession WHERE (Z_ID = '" & pZID & "')}  AS ZSession APPEND (( SHAPE {SELECT * FROM tOPSESSION WHERE OPS_ID = '" & pOPSID & "'}  AS OPSession APPEND (( SHAPE {SELECT EXCH_STATUS, EXCH_ID, EXCH_ZSESSIONID,EXCH_OPSESSIONID,EXCH_TP_ID,EXCH_TYPE,EXCH_SALEDATE,EXCH_SALEVALUE,EXCH_DISCOUNTVALUE,EXCH_VATVALUE,EXCH_CHANGEGIVEN,EXCH_TYPE,EXCH_NUMBER,EXCH_VOIDS FROM tEXCHANGE WHERE EXCH_ID = '" & pEXCHID & "'}  AS POSExchange APPEND ({SELECT * FROM tCSL}  AS rsSALESLINES RELATE EXCH_ID TO CSL_EXCH_ID) AS SALESLINES,({SELECT * FROM tPayment}  AS rsPAYMENTS RELATE EXCH_ID TO PAY_EXCH_ID) AS PAYMENTS) AS POSExchange RELATE OPS_ID TO EXCH_OPSESSIONID) AS POSExchange) AS OPSession RELATE Z_ID TO OPS_Z_ID) AS OPSession"
    Set rsZSession = Nothing
    Set rsZSession = New ADODB.Recordset
    rsZSession.Open sSQL, oShapeDB.DBConn, adOpenStatic
    Set rsZSession.ActiveConnection = Nothing
    
    If Not rsZSession.EOF Then
        sFileName = oPC.NameOfPC & "-" & Format(Now(), "DDHHNNSS") 'Format(oGD.GetNextFileNum(), "00000")
        sFileName = "\" & sFileName & ".pos"
        rsZSession.Save oPS.ClientOutbox & sFileName, adPersistADTG
    End If
    If pEXCHID > "" Then
        SQL = "UPDATE tExchange SET EXCH_STATUS = 'X' WHERE EXCH_ID = '" & pEXCHID & "'"
        oPC.DBConn.Execute SQL
    End If
    oShapeDB.dbCloseConnectShape
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSExchange.SendPOSExchange"
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
    ErrorIn "frmMain.Connect"
End Sub

Private Sub RebuildIndexes()
Dim oTable As SQLDMO.Table
    For Each oTable In oDatabase.Tables
        If Not oTable.SystemObject Then oTable.RebuildIndexes
    Next
End Sub
Private Function Disconnect()
    On Error Resume Next
    oSQLServer.Disconnect
    Set ADOConn = Nothing
End Function

Private Sub PrintSalesSlip()

    Dim lValue As Long
    Dim i As Integer
    Dim idBuf() As ITEMDATA
    Dim fDate As String
    Dim BcData  As String
    Dim sBuf As String
    Dim sExt As String
    Dim sType As String
    Dim sAt As String
    Dim sValue As String

' When outputting to a printer,a mouse cursor becomes like a hourglass.
    MousePointer = vbHourglass

    BcData = "4902720005074"
    
    ReDim idBuf(1 To oExchange.SaleLines.Count)
    For i = 1 To oExchange.SaleLines.Count
        If Not oExchange.SaleLines(i).IsDeleted Then
            idBuf(i).TType = IIf(oExchange.SaleLines(i).Qty < 0, "R ", "S ")
            idBuf(i).Name = oExchange.SaleLines(i).TitleIncDiscF(20)
            idBuf(i).Ext = oExchange.SaleLines(i).PLessDiscExtF
            idBuf(i).At = oExchange.SaleLines(i).QtyF & " @ " & oExchange.SaleLines(i).PriceF
        End If
    Next i
    
    With OPOSPOSPrinter1
        PrintHeader ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'Print header
        
        For i = LBound(idBuf) To UBound(idBuf)          'Print each line
            If .ResultCode <> OPOS_SUCCESS Then Exit For
            sAt = idBuf(i).At
            sBuf = idBuf(i).Name
            sExt = idBuf(i).Ext
            sType = idBuf(i).TType
            sValue = MakePrintStringDetail(.RecLineChars, sType, sBuf, sAt, sExt)
            .PrintNormal PTR_S_RECEIPT, sValue + vbLf
            .PrintNormal PTR_S_RECEIPT, oExchange.SaleLines(i).CodeF + vbLf
        Next
        .PrintNormal PTR_S_RECEIPT, ESC + "|200uF"      'create gap
            
        PrintTotals ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print totals
        PrintFooter ConvertToType(oExchange.transactionType), OPOSPOSPrinter1           'print footer
        
        .PrintNormal PTR_S_RECEIPT, ESC + "|2500uF"     'create gap
        .TransactionPrint PTR_S_RECEIPT, PTR_TP_NORMAL  'Go

        'Back to the synchronous mode
        .AsyncMode = False
    End With

' When a cursor is back to its default shape, it means the process ends.
    MousePointer = vbDefault

End Sub

Private Function MakePrintStringDetail(lRecLineChars As Long, sType As String, sBuf As String, sAt As String, sExt As String) As String
Dim sValue As String
Dim strNotChangeable As String
Dim iAvailable As Integer
    
    strNotChangeable = sAt & " " & sExt
    iAvailable = lRecLineChars - Len(strNotChangeable) - Len(sType)
    If iAvailable > 2 Then
        sValue = sType & Left(sBuf, iAvailable) & sAt & " " & sExt
    Else
        sValue = sType & sAt & " " & sExt
    End If
    MakePrintStringDetail = sValue
End Function
Private Function MakePrintString(lRecLineChars As Long, sBuf As String, sPrice As String) As String
Dim sValue As String
    If lRecLineChars < (Len(sBuf) + Len(sPrice)) Then
        sValue = sBuf + sPrice
    Else
        sValue = sBuf + Space(lRecLineChars - (Len(sBuf) + Len(sPrice))) + sPrice
    End If

    MakePrintString = sValue
End Function

Private Sub AddExchange()
Dim oSale As a_Sale

    For Each oSale In oExchange.SaleLines
        lngSalesItemCount = lngSalesItemCount + 1
        X4.InsertRows (lngSalesItemCount)
        X4.Value(lngSalesItemCount, 1) = lngSalesItemCount
        X4.Value(lngSalesItemCount, 2) = oExchange.ExchangeTimeF
        X4.Value(lngSalesItemCount, 3) = oExchange.StaffName
        X4.Value(lngSalesItemCount, 4) = oSale.CodeF & " (" & oSale.QtyF & ") " & oSale.TitleF(30) & " " & oSale.PLessDiscF
        X4.Value(lngSalesItemCount, 5) = oExchange.ExchangeID
        X4.Value(lngSalesItemCount, 6) = oSale.PID
    Next
    G4.Array = X4
    G4.ReBind
    G4.Bookmark = lngSalesItemCount
End Sub

Private Sub LoadExchanges()
Dim ZID As String
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

    
    ZID = oPC.ZSession.Current_Z_Session_ID

    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = oPC.DBConn
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
        X4.Value(lngSalesItemCount, 1) = FNN(rs.Fields("EXCH_NUMBER"))  'lngSalesItemCount
        X4.Value(lngSalesItemCount, 2) = Format(rs.Fields("EXCH_SaleDate"), "HH:NN")
        X4.Value(lngSalesItemCount, 3) = FNS(rs.Fields("SM_SHORTNAME"))
        X4.Value(lngSalesItemCount, 4) = FNS(rs.Fields("Code")) & " (" & FNN(rs.Fields("CSL_Qty")) & ") " & FNS(rs.Fields("TITLE")) & " " & rs.Fields("DiscountedValueIncVAT")
        X4.Value(lngSalesItemCount, 5) = FNS(rs.Fields("EXCH_ID"))
        X4.Value(lngSalesItemCount, 6) = FNS(rs.Fields("P_ID"))
        X4.Value(lngSalesItemCount, 7) = FNN(rs.Fields("EXCH_Voided"))
        X4.Value(lngSalesItemCount, 8) = FNN(rs.Fields("EXCH_Voids"))
    '    If FNN(rs.Fields("SM_ID")) > 0 Then MsgBox "Here"
        rs.MoveNext
    Loop
    G4.Array = X4
    G4.ReBind
    G4.Bookmark = lngSalesItemCount


    
End Sub

Private Sub PrintTotals(eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter)
Dim sBuf As String
Dim sExt As String
Dim sValue As String
    Select Case eDocumentType
    Case eTypReceipt
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
            
            sBuf = "Customer's payment"
            sExt = oExchange.TotalPaymentF
            sValue = MakePrintString(.RecLineChars, sBuf, sExt)
            .PrintNormal PTR_S_RECEIPT, ESC + "|N" + sValue + vbLf
            
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
            
            
            sBuf = "Cash refund"
            sExt = oExchange.TotalLessDiscExtF
            sValue = MakePrintString(.RecLineChars, sBuf, sExt)
            .PrintNormal PTR_S_RECEIPT, sValue + vbLf
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
End Sub

Private Sub PrintHeader(eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter)
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.TillCode & "," & oExchange.SalesPersonName & vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.TillCode & "," & oExchange.SalesPersonName & vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
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
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.TillCode & "," & oExchange.SalesPersonName & vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
        End With
    Case eTypDeposit
        With pPrinter
          .AsyncMode = True
          .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
          .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSCompanyName + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "Deposit paid" + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSBranchName + vbLf
          ar = Split(oPC.POSBranchAddress, ",")
          For i = 0 To UBound(ar)
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
          Next i
            fDate = Format(Now, "dd/mm/yy h:mm AM/PM ")
          .PrintNormal PTR_S_RECEIPT, ESC + "|300uF"
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + "#" & CStr(oExchange.ExchangeNumber) & "  " & oPC.TillCode & "," & oExchange.SalesPersonName & vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + fDate + vbLf
          .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
        End With
    End Select
End Sub
Private Sub PrintFooter(eDocumentType As enumDocumentType, pPrinter As OPOSPOSPrinter)
Dim ar() As String
Dim i As Integer

    Select Case eDocumentType
    Case eTypReceipt, eTypCashRefund, etypCreditNote, eTypDeposit
        With pPrinter
            .AsyncMode = True
            .TransactionPrint PTR_S_RECEIPT, PTR_TP_TRANSACTION
            .PrintNormal PTR_S_RECEIPT, ESC + "|700uF"
            .PrintNormal PTR_S_RECEIPT, ESC + "|1B"
            ar = Split(oPC.POSReceiptMessage, ",")
            For i = 0 To UBound(ar)
              .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + ar(i) + vbLf
            Next i
            .PrintNormal PTR_S_RECEIPT, ESC + "|cA" + oPC.POSemailAddress + vbLf
            .PrintNormal PTR_S_RECEIPT, ESC + "|500uF"
        End With
    End Select
End Sub
Private Function ConvertToType(val As String) As Integer
    Select Case val
    Case "S"
        ConvertToType = eTypReceipt
    Case "R"
        ConvertToType = eTypCashRefund
    Case "C"
        ConvertToType = etypCreditNote
    Case "D"
        ConvertToType = eTypDeposit
    End Select
End Function
