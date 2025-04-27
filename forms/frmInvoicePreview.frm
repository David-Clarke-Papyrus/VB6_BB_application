VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "COOLBU~1.OCX"
Begin VB.Form frmInvoicePreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmInvoicePreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdToReal 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Copy to real invoice"
      Height          =   345
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
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
      Height          =   1140
      Left            =   2865
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4905
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1950
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicePreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton cmdUP 
      BackColor       =   &H00C4BCA4&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11100
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4140
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00C4BCA4&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11100
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4470
      Width           =   330
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicePreview.frx":284D
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   150
      Width           =   1545
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   210
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicePreview.frx":2997
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtInvoiceNum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2835
      Left            =   240
      OleObjectBlob   =   "frmInvoicePreview.frx":2CA1
      TabIndex        =   18
      Top             =   1905
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   960
      Left            =   225
      TabIndex        =   23
      Top             =   825
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1693
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.Label txtStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   390
      Left            =   9420
      TabIndex        =   21
      Top             =   90
      Width           =   1770
   End
   Begin VB.Label txtComp 
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
      Height          =   270
      Left            =   4590
      TabIndex        =   20
      Top             =   105
      Width           =   3240
   End
   Begin VB.Label lblSI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Height          =   240
      Left            =   675
      TabIndex        =   19
      Top             =   480
      Width           =   2970
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1125
      X2              =   2595
      Y1              =   0
      Y2              =   825
   End
   Begin VB.Label lblTPFax 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   495
      TabIndex        =   16
      Top             =   1515
      Width           =   2895
   End
   Begin VB.Label lblTPPhone 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   495
      TabIndex        =   15
      Top             =   1185
      Width           =   2895
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   495
      TabIndex        =   14
      Top             =   855
      Width           =   2895
   End
   Begin VB.Label lblDelToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   9015
      TabIndex        =   13
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label lblBillToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   945
      Left            =   5865
      TabIndex        =   12
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5085
      TabIndex        =   11
      Top             =   780
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Goods to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7845
      TabIndex        =   10
      Top             =   780
      Width           =   1050
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   5505
      TabIndex        =   6
      Top             =   4860
      Width           =   3495
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   9090
      TabIndex        =   5
      Top             =   4860
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   720
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1365
   End
End
Attribute VB_Name = "frmInvoicePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cInv As c_Invoices
Dim oInvoice As a_Invoice
Dim dblTotal As Double
Dim XA As XArrayDB
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetMenu()
    Forms(0).mnuVoid.Enabled = (oInvoice.Status = stInProcess And oInvoice.IsNew = False)
    Forms(0).mnuCancel.Enabled = False '(oInvoice.Status = stCOMPLETE) Or (oInvoice.Status = stISSUED)
    Forms(0).mnuCancelLine.Enabled = False  '(oInvoice.Status = stISSUED)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = True
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuCreateCreditNote.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
End Sub
Public Sub CreateCreditNote()
Dim oNew As a_CN
Dim ofrm As New frmCN
Dim lngID As Long
Dim frm As frmGenCN

    Set frm = New frmGenCN
    frm.Component oInvoice, XA
    frm.Show vbModal
    Set oNew = New a_CN
    oNew.BeginEdit
    oNew.BuildFromInvoice oInvoice
    oNew.ApplyEdit
    Unload frmGenCN
'    lngID = oNew.trid
'    Set oNew = Nothing
'    Set oNew = New a_CN
'    oNew.Load lngID, False
'    ofrm.Component , oNew
'    ofrm.Show
'
End Sub

Private Sub cmdToReal_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer

    Screen.MousePointer = vbHourglass
    Set cmd = New ADODB.Command
    cmd.CommandText = "CopyInvoice"
    cmd.CommandType = adCmdStoredProc
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    Set par = cmd.CreateParameter("@INVID", adInteger, adParamInput, , oInvoice.InvoiceID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , oInvoice.COMPID)
    cmd.Parameters.Append par
    cmd.Execute
    Set par = Nothing
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    MsgBox "A new invoice has been created and will be found by browsing invoices.", , "Action complete"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.cmdToReal_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Activate()
    SetMenu
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub

Public Sub Component(pID As Long)
Dim lngID As Long
    lngID = pID
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    Me.Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
    If oInvoice.SalesRepName > "" Then
        Me.Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    Me.cmdToReal.Visible = oInvoice.proforma
    LoadControls
    SetMenu
End Sub
Public Sub ComponentObject(pInvoice As a_Invoice)
    Set oInvoice = pInvoice
    Me.Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.proforma, "    PRO-FORMA", "")
    If oInvoice.SalesRepName > "" Then
        Me.Caption = Me.Caption & "  (Rep: " & oInvoice.SalesRepName & ")"
    End If
    Me.cmdToReal.Visible = oInvoice.proforma
    LoadControls
End Sub
Private Sub LoadControls()
Dim dblVAT As Double
Dim dblConversionRate As Double
Dim strCurrencyFormat As String
Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strAddress As String
Dim strTotalCaption As String
Dim strTotalValues As String
    On Error GoTo ERR_Handler
    
        With oInvoice
            If (.Status = stInProcess) Or (.proforma = True) Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtDate = .DOCDate
            If DateDiff("d", .DOCDate, .CaptureDate) > 1 Then
                lblSI.Caption = "Issued: " & .CaptureDateF
            Else
                lblSI.Caption = ""
            End If
            Me.txtStatus.Caption = .statusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If Not .BillingCompany Is Nothing Then
                Me.txtComp = "From: " & .BillingCompany.CompanyName
            End If
            Me.txtInvoiceNum = .DocCode
            lblTPName.Caption = .Customer.Fullname & IIf(Len(.TPAccNum) > 0, " (" & .TPAccNum & ")", "")
            If Not .Customer.BillTOAddress Is Nothing Then
                lblTPPhone.Caption = .Customer.BillTOAddress.Phone
                lblTPFax.Caption = .Customer.BillTOAddress.Fax
            End If
            Me.txtTPMemo = IIf(Len(.Memo) > 0, "Note:  " & Trim$(.Memo), "")
            txtTPMemo.Visible = (txtTPMemo > "")
            If .BillToAddressID > 0 Then
                If Not .BillTOAddress Is Nothing Then
                    strAddress = .BillTOAddress.AddressMailing
                End If
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .DelToAddressID > 0 Then
                If Not .DelToAddress Is Nothing Then
                    strAddress = .DelToAddress.AddressMailing
                End If
            End If
            Me.lblDelToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            dblConversionRate = .CurrencyFactor
            If .CurrencyFormat > "" Then
                strCurrencyFormat = .CurrencyFormat
            Else
                strCurrencyFormat = "Currency"
            End If
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
'            If Not .ForeignCurrency Is oPC.Configuration.DefaultCurrency Then
'                Me.txtCurrency = .ForeignCurrency.Description
'            End If
        End With
        LoadGrid
      '  LoadListView
     '   LoadSummary oInvoice.Postage, oInvoice.VATRate, dblConversionRate, strCurrencyFormat, curTotalValue, curTotalDeposits
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
Resume
End Sub


Private Sub cbCust_Click()
Dim frm As New frmCustomerPreview
    frm.Component oInvoice.Customer
    frm.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim frm As frmPrintingOptions_Inv
Dim i As Long
    Set frm = New frmPrintingOptions_Inv
    frm.ComponentObject oInvoice
    frm.Show vbModal
    LoadGrid
EXIT_HANDLER:
 '   Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
Resume
End Sub
Private Sub cmdEdit_Click()
Dim blnEdit As Boolean
Dim frmInvoice As frmInvoice
Dim strPreviousStatusBarCaption As String
    On Error GoTo ERR_Handler
    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).Text
    Forms(0).SB1.Panels(2).Text = "LOADING . . ."
  '  WaitMsg "Loading . . .", True, Me
        Set frmInvoice = New frmInvoice
  '  End If
    blnEdit = True
    frmInvoice.Component , oInvoice
    Unload Me
    frmInvoice.Show
    Forms(0).SB1.Panels(2).Text = strPreviousStatusBarCaption
'    WaitMsg "", False, Me

EXIT_HANDLER:
   ' Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    Resume
End Sub
Private Sub cmdUP_Click()
Dim i As Long
   ' str = FNS(XA.Value(G1.Bookmark, 11))
    If G1.Bookmark > 1 Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oInvoice.BeginEdit
        oInvoice.InvoiceLines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark - 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
       'LoadListView
       ' lvwInvLines.Refresh
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub cmdDown_Click()
Dim i As Long
    If G1.Bookmark < XA.UpperBound(1) Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oInvoice.BeginEdit
        oInvoice.InvoiceLines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark + 1, 11))
        oInvoice.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
End Sub

'Private Sub LoadListView()
'Dim lstItem As ListItem
'Dim i As Integer
'Dim currDeposit As Currency
'Dim currPrice As Currency
'Dim dblVAT As Double
'Dim strSummaryDescription As String
'Dim strSummary As String
'Dim lngTotal As Long
'Dim lngDepositTotal As Long
'
'    On Error GoTo ERR_Handler
'    lvwInvLines.ListItems.Clear
'    For i = 1 To oInvoice.InvoiceLines.Count
'        If Not oInvoice.InvoiceLines(i).BottomOfDocument Then
'            Set lstItem = lvwInvLines.ListItems.Add
'            With lstItem
'               ' .Key = oInvoice.InvoiceLines(i).InvoiceLineID & "k"
'                .Key = oInvoice.InvoiceLines(i).Key
'                .Text = oInvoice.InvoiceLines(i).CodeF
'                .SubItems(1) = oInvoice.InvoiceLines(i).TitleAuthorPublisher
'                .SubItems(2) = oInvoice.InvoiceLines(i).Qty
'                If oInvoice.InvoiceLines(i).Deposit > 0 Then
'                    .SubItems(3) = oInvoice.InvoiceLines(i).DepositF(False)
'                Else
'                    .SubItems(3) = " "
'                End If
'                .SubItems(4) = oInvoice.InvoiceLines(i).PriceF(False)
'                .SubItems(5) = oInvoice.InvoiceLines(i).DiscountPercentF
'                .SubItems(6) = oInvoice.InvoiceLines(i).Ref
'                .SubItems(7) = oInvoice.InvoiceLines(i).PLessDiscExtF(False)
'                .SubItems(8) = oInvoice.InvoiceLines(i).Note
'                .SubItems(9) = oInvoice.InvoiceLines(i).sequence
'                If oInvoice.InvoiceLines(i).PIID = 0 Then
'                    .ForeColor = &H427182
'                    .ListSubItems(1).ForeColor = &H706034
'                    .ListSubItems(1).ForeColor = &H706034
'                    .ListSubItems(2).ForeColor = &H706034
'                    .ListSubItems(3).ForeColor = &H706034
'                    .ListSubItems(4).ForeColor = &H706034
'                    .ListSubItems(5).ForeColor = &H706034
'                    .ListSubItems(6).ForeColor = &H706034
'                    .ListSubItems(7).ForeColor = &H706034
'                ElseIf oInvoice.InvoiceLines(i).NonStock = True Then
'                    lstItem.ForeColor = &H427182
'                    lstItem.ListSubItems(1).ForeColor = &H427182
'                    lstItem.ListSubItems(2).ForeColor = &H427182
'                    lstItem.ListSubItems(3).ForeColor = &H427182
'                    lstItem.ListSubItems(4).ForeColor = &H427182
'                    lstItem.ListSubItems(5).ForeColor = &H427182
'                    lstItem.ListSubItems(6).ForeColor = &H427182
'                    lstItem.ListSubItems(7).ForeColor = &H427182
'                End If
'            End With
'            If oInvoice.InvoiceLines(i).Note > "" Then
'                lstItem.ToolTipText = "Note:  " & oInvoice.InvoiceLines(i).Note
'            End If
'        End If
'    Next i
'
'EXIT_Handler:
'    Exit Sub
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'End Sub
Private Sub LoadGrid()
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long

    Set XA = New XArrayDB
    XA.Clear
    On Error GoTo ERR_Handler
    XA.ReDim 1, oInvoice.InvoiceLines.Count, 1, 19
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    G1.Columns(8).Width = 1
    For i = 1 To oInvoice.InvoiceLines.Count
            XA(i, 11) = oInvoice.InvoiceLines(i).Key
            XA(i, 12) = oInvoice.InvoiceLines(i).code
            XA(i, 15) = oInvoice.InvoiceLines(i).pID
            XA(i, 16) = IIf(oInvoice.InvoiceLines(i).SubstitutesAvailable, "Y", "N")
            XA(i, 17) = oInvoice.InvoiceLines(i).InvoiceLineID
            XA(i, 18) = oInvoice.InvoiceLines(i).COLID
            XA(i, 19) = oInvoice.InvoiceLines(i).Ean
            XA(i, 1) = oInvoice.InvoiceLines(i).CodeF
            XA(i, 2) = oInvoice.InvoiceLines(i).TitleAuthorPublisher
            XA(i, 3) = oInvoice.InvoiceLines(i).Qty & IIf(oInvoice.InvoiceLines(i).CreditedQty > 0, "(" & oInvoice.InvoiceLines(i).CreditedQty & ")", "")
            If oInvoice.InvoiceLines(i).Deposit > 0 Then
                XA(i, 4) = oInvoice.InvoiceLines(i).DepositF(False)
            Else
                XA(i, 4) = " "
            End If
            XA(i, 5) = oInvoice.InvoiceLines(i).PriceF(False)
            XA(i, 6) = oInvoice.InvoiceLines(i).DiscountPercentF
            XA(i, 7) = oInvoice.InvoiceLines(i).Ref
            XA(i, 8) = oInvoice.InvoiceLines(i).PLessDiscExtF(False)
            XA(i, 9) = oInvoice.InvoiceLines(i).Note
            XA(i, 10) = oInvoice.InvoiceLines(i).Sequence
            If oInvoice.InvoiceLines(i).Note > "" Then
                If oInvoice.InvoiceLines(i).Note = "Substitute" Then
                    XA(i, 9) = "Note:  " & oInvoice.InvoiceLines(i).Note & "  (Operator: right-mouse click for substitution options!)"
                Else
                XA(i, 9) = "Note:  " & oInvoice.InvoiceLines(i).Note
                End If
                G1.Columns(8).Width = 4000
            End If
            XA(i, 13) = oInvoice.InvoiceLines(i).CreditedQty
            XA(i, 14) = oInvoice.InvoiceLines(i).Qty
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, 0, GetRowType(11)
    
    G1.Array = XA
    G1.ReBind

    
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    Resume
End Sub

'Private Sub LoadSummary(pPostage As Currency, pVAT As Double, pConversionRate As Double, pCurrFormat As String, curTotalValue As Currency, curTotalDeposits As Currency)
'Dim currPrice As Currency
'Dim strDiscount As String
'Dim dblVAT As Double
'Dim strSummaryDescription As String
'Dim strSummary As String
'    dblVAT = (curTotalValue / (1 + pVAT)) * pVAT
'    If pPostage = 0 And curTotalDeposits = 0 And oInvoice.VATAble Then
'        strSummaryDescription = "(Includes VAT of " & Format(dblVAT, pCurrFormat) & ")          Total: "
'        strSummary = Format(curTotalValue, pCurrFormat)
'    ElseIf oInvoice.VATAble Then
'        strSummaryDescription = "Subtotal:"
'        strSummary = Format(curTotalValue, pCurrFormat)
'        If curTotalDeposits <> 0 Then
'            Me.lblDeposits = "(Deposits paid : " & Format(curTotalDeposits, pCurrFormat) & ")                          "
'        End If
'        If pPostage <> 0 Then
'            strSummaryDescription = strSummaryDescription & vbCrLf & "Plus Postage && handling:"
'            strSummary = strSummary & vbCrLf & Format(pPostage, pCurrFormat)
'        End If
'        strSummaryDescription = strSummaryDescription & vbCrLf & "(Includes VAT of " & Format(dblVAT, pCurrFormat) & ")          Total: "
'        strSummary = strSummary & vbCrLf & Format((curTotalValue + pPostage), pCurrFormat)
'    ElseIf Not oInvoice.VATAble Then
'        strSummaryDescription = "Subtotal:"
'        strSummary = Format(curTotalValue, pCurrFormat)
'        strSummaryDescription = strSummaryDescription & vbCrLf & "Less VAT"
'        strSummary = strSummary & vbCrLf & Format(dblVAT, pCurrFormat)
'        If curTotalDeposits <> 0 Then
'            Me.lblDeposits = "(Deposits paid : " & Format(curTotalDeposits, pCurrFormat) & ")                          "
'        End If
'        If pPostage <> 0 Then
'            strSummaryDescription = strSummaryDescription & vbCrLf & "Postage && handling:"
'            strSummary = strSummary & vbCrLf & Format(pPostage, pCurrFormat)
'        End If
'        strSummaryDescription = strSummaryDescription & vbCrLf & "(Total excluding VAT " & Format((curTotalValue - dblVAT) + pPostage, pCurrFormat) & ")"
'    End If
'    Me.lblDescription = strSummaryDescription
'    Me.lblSummary = strSummary
'
'End Sub


Private Sub CoolButton1_MouseEnter()

End Sub

Private Sub Form_Load()
    
    Me.top = 50
    Me.left = 50
    Me.Height = 6500
    Me.Width = 11600
    If oInvoice.proforma Then
        Me.BackColor = 14542803
    End If
 '   SetLvw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
    If oInvoice.IsEditing And frmInvoice Is Nothing Then oInvoice.CancelEdit
    Set oInvoice = Nothing
End Sub

Private Sub G1_Click()
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
   ' str = FNS(XA.Value(G1.Bookmark, 12))
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
    Exit Sub
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuInvoicePreview   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub InsertSubstitutes()
Dim frm As frmInsertSubstitute
Dim oIL As a_InvoiceLine
Dim str As String
Dim lngQty As Long

    If FNS(XA.Value(G1.Bookmark, 16)) <> "Y" Then
        MsgBox "There are no substitutes available for this item.", vbOKOnly + vbInformation, "Status"
        Exit Sub
    End If
    Set frm = New frmInsertSubstitute
    str = FNS(XA.Value(G1.Bookmark, 15))
    lngQty = FNN(XA.Value(G1.Bookmark, 3))
   
    frm.Component oInvoice.Customer.NameAndCode(50), lngQty, XA.Value(G1.Bookmark, 15), XA.Value(G1.Bookmark, 18), XA.Value(G1.Bookmark, 17), oInvoice.InvoiceID
    frm.Show vbModal
    Unload frm
    Unload Me
    MsgBox "Substitutions have been made.", vbOKOnly, "Status"
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler

    If FNN(XA(Bookmark, 13)) > 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoicePreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
 '   Forms(0).mnuCancelLine.Enabled = oCO.COLines(str).QtyDispatched = 0
   ' str = FNS(XA.Value(G1.Bookmark, 12))
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
End Sub

Private Sub G1_SelChange(Cancel As Integer)
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
End Sub
Private Sub G1_DblClick()
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oInvoice.InvoiceLines(str).pID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.Component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 7, 9
            GetRowType = XTYPE_STRING
        Case 3, 4, 6, 5, 8
            GetRowType = XTYPE_INTEGER
    End Select
End Function


'Private Sub lvwInvLines_AfterLabelEdit(Cancel As Integer, NewString As String)
'Cancel = True
'End Sub
'
'Private Sub lvwInvLines_Click()
'    Clipboard.Clear
'    Clipboard.SetText oInvoice.InvoiceLines(lvwInvLines.SelectedItem.Key).code
'
'End Sub
'
'Private Sub lvwInvLines_BeforeLabelEdit(Cancel As Integer)
'Cancel = True
'End Sub
'

'Private Sub lvwInvLines_DblClick()
'Dim frmA As frmProductPrevAQ
'Dim frm As frmProductPrev
'Dim oP As a_Product
'    Screen.MousePointer = vbHourglass
'    Set oP = New a_Product
'   ' oP.Load oInvoice.InvoiceLines.FindLineByID(val(Me.lvwInvLines.SelectedItem.Key)).pID, 0
'    oP.Load oInvoice.InvoiceLines(Me.lvwInvLines.SelectedItem.Key).pID, 0
'    If oPC.Configuration.AntiquarianYN Then
'        Set frmA = New frmProductPrevAQ
'        frmA.Component oP
'        frmA.Show
'    Else
'        Set frm = New frmProductPrev
'        frm.Component oP
'        frm.Show
'    End If
'    Screen.MousePointer = vbDefault
'End Sub

'Private Sub mnuFF_Click()
'Dim frm As frmCOFF
'Dim strSQL As String
'        Set frm = New frmCOFF
'        frm.Component oInvoice
'        frm.Show vbModal
'End Sub

'Private Sub lvwInvLines_ItemClick(ByVal Item As MSComctlLib.ListItem)
'    Me.lvwInvLines.ToolTipText = Item.SubItems(8)
'End Sub

'Private Sub SetLvw()
'Dim style As Long
'Dim hHeader As Long
'
'  'get the handle to the listview header
'   hHeader = SendMessage(lvwInvLines.hwnd, LVM_GETHEADER, 0, ByVal 0&)
'
'  'get the current style attributes for the header
'   style = GetWindowLong(hHeader, GWL_STYLE)
'
'  'modify the style by toggling the HDS_BUTTONS style
'   style = style Xor HDS_BUTTONS
'
'  'set the new style and redraw the listview
'   If style Then
'      Call SetWindowLong(hHeader, GWL_STYLE, style)
'      Call SetWindowPos(lvwInvLines.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If
'
'
'End Sub
Public Sub mnuSalesComm()
Dim frm As New frmSalesComm
Dim OpenResult As Integer

    frm.Component oInvoice.SalesRepID, oInvoice.SalesRepName, oInvoice.CustPaid, oInvoice.CommPaid
    frm.Show vbModal
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If frm.CustPaid <> oInvoice.CustPaid Then
        oPC.COShort.Execute "EXECUTE dbo.MarkInvoicePaid " & oInvoice.InvoiceID & "," & IIf(frm.CustPaid, "1", "0")
        oInvoice.CustPaid = frm.CustPaid
    End If
    If frm.CommPaid <> oInvoice.CommPaid Then
        oPC.COShort.Execute "EXECUTE dbo.MarkCOmmissionPaid " & oInvoice.InvoiceID & "," & IIf(frm.CommPaid, "1", "0")
        oInvoice.CommPaid = frm.CommPaid
    End If
    
    
    If oInvoice.SalesRepID <> frm.SalesRepID Then
        oInvoice.SalesRepID = frm.SalesRepID
        oInvoice.SalesRepName = frm.SalesRepName
        oPC.COShort.Execute "Execute dbo.AllocateSalesCommission " & oInvoice.InvoiceID & "," & oInvoice.SalesRepID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Unload frm

End Sub

Public Sub mnuCancel()
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelInvoice oInvoice
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Public Sub mnuVoid()
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oInvoice.VoidDocument
    RefreshData
End Sub
Public Sub RefreshData()
    oInvoice.ReLoad
    LoadControls
End Sub

