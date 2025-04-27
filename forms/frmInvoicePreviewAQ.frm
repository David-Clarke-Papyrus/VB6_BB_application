VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInvoicePreviewAQ 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Invoice"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11430
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmInvoicePreviewAQ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1995
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Print the invoice"
      Top             =   5100
      Width           =   855
   End
   Begin VB.TextBox txtCurrency 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00706034&
      Height          =   255
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   420
      Width           =   1635
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   420
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5715
      Width           =   5970
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   4440
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
      Left            =   1125
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicePreviewAQ.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtDate 
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
      TabIndex        =   5
      Top             =   240
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
      Left            =   255
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmInvoicePreviewAQ.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Height          =   375
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   1545
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
      TabIndex        =   1
      Top             =   240
      Width           =   1545
   End
   Begin MSComctlLib.ListView lvwInvLines 
      Height          =   2505
      Left            =   240
      TabIndex        =   0
      Top             =   2310
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   4419
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ISBN"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   7231
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Dep."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Disc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1545
      Left            =   240
      TabIndex        =   6
      Top             =   675
      Width           =   10860
      Begin VB.TextBox txtTPFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   240
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   765
         Width           =   2000
      End
      Begin VB.TextBox txtTPPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   240
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   450
         Width           =   2000
      End
      Begin VB.TextBox txtTPName 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   150
         Width           =   4320
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   7515
         Picture         =   "frmInvoicePreviewAQ.frx":03DE
         Stretch         =   -1  'True
         Top             =   405
         Width           =   360
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   17
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Goods to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         TabIndex        =   16
         Top             =   390
         Width           =   1050
      End
      Begin VB.Label lblGoodsToAddress 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4695
         TabIndex        =   15
         Top             =   405
         Width           =   2490
      End
      Begin VB.Label lblBillToAddress 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   840
         TabIndex        =   14
         Top             =   405
         Width           =   2490
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bill to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   7
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   5505
      TabIndex        =   20
      Top             =   4935
      Width           =   3495
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2925
      TabIndex        =   11
      Top             =   4845
      Width           =   2325
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   9090
      TabIndex        =   10
      Top             =   4935
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   540
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
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
      TabIndex        =   2
      Top             =   240
      Width           =   1365
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmInvoicePreviewAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cInv As c_Invoices
Dim oInvoice As a_Invoice
Dim dblTotal As Double

Dim lngID As Long

Public Sub Component(pID As Long)
    lngID = pID
    Set oInvoice = New a_Invoice
    oInvoice.Load lngID, True
    Me.Caption = "Invoice for " & oInvoice.TPName
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
            Me.txtDate = .TransDate
            Me.txtStatus = .Statusf
            Me.txtComp = "From: " & .BillingCompany.CompanyName
            If .status = stInPROCESS Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtInvoiceNum = .TransCode
            Me.txtTPName = .TPName & IIf(Len(.TPAccNum) > 0, " (" & .TPAccNum & ")", "")
            Me.txtTPPhone = .TPPhone
            Me.txtTPFax = IIf(Len(.TPFax) > 0, .TPFax & "(fax)", "")
            Me.txtTPMemo = IIf(Len(.TPMemo) > 0, "Note:  " & Trim$(.TPMemo), "")
            If .InvDocAddressID > 0 Then
                strAddress = .BillToAddress.AddressMailing
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .InvGoodsAddressID > 0 Then
                strAddress = .GoodsToAddress.AddressMailing
            End If
            Me.lblGoodsToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
        ' .dblVAT = .VATRate
            dblConversionRate = .CurrencyFactor
            If .CurrencyFormat > "" Then
                strCurrencyFormat = .CurrencyFormat
            Else
                strCurrencyFormat = "Currency"
            End If
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
            If Not .ForeignCurrency Is oPC.Configuration.DefaultCurrency Then
                Me.txtCurrency = .ForeignCurrency.Description
            End If
        End With
        LoadListView
     '   LoadSummary oInvoice.Postage, oInvoice.VATRate, dblConversionRate, strCurrencyFormat, curTotalValue, curTotalDeposits
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
Resume
End Sub

Private Sub cmdPreview_Click()
    oInvoice.PrintInvoice_Display True
End Sub

Private Sub cmdPrint_Click()
Dim blnEdit As Boolean

'
    On Error GoTo ERR_Handler
    blnEdit = False
    oInvoice.PrintInvoice True, True
EXIT_Handler:
    Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
Resume
End Sub
Private Sub cmdEdit_Click()
Dim blnEdit As Boolean
Dim frmInvoice As frmInvoiceAQ
    On Error GoTo ERR_Handler
  '  If frmInvoiceAQ Is Nothing Then
        Set frmInvoice = New frmInvoiceAQ
  '  End If
    blnEdit = True
    frmInvoice.Component oInvoice ', lngID
    frmInvoice.Show

EXIT_Handler:
    Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

Private Sub LoadListView()
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long

    On Error GoTo ERR_Handler

    For i = 1 To oInvoice.InvoiceLines.Count
        Set lstItem = lvwInvLines.ListItems.Add
        With lstItem
            .Key = oInvoice.InvoiceLines(i).InvoiceLineID & "k"
            .Text = oInvoice.InvoiceLines(i).CodeF
            .SubItems(1) = oInvoice.InvoiceLines(i).TitleAuthorPublisher
            .SubItems(2) = oInvoice.InvoiceLines(i).Qty
            If oInvoice.InvoiceLines(i).Deposit > 0 Then
                .SubItems(3) = oInvoice.InvoiceLines(i).DepositF(False)
            Else
                .SubItems(3) = " "
            End If
            .SubItems(4) = oInvoice.InvoiceLines(i).PriceF(False)
            .SubItems(5) = oInvoice.InvoiceLines(i).DiscountPercentF
'            If oInvoice.InvoiceLines(i).DiscountPercent <> 0 Then
'              '  currPrice = oInvoice.InvoiceLines(i).Price / (1 + oInvoice.InvoiceLines(i).DiscountPercent)
'            Else
'                currPrice = oInvoice.InvoiceLines(i).Price
'            End If
            .SubItems(6) = oInvoice.InvoiceLines(i).ExtensionF(False)
            If oInvoice.InvoiceLines(i).CopyID = 0 Then
                .ForeColor = &H427182
                .ListSubItems(1).ForeColor = &H706034
                .ListSubItems(1).ForeColor = &H706034
                .ListSubItems(2).ForeColor = &H706034
                .ListSubItems(3).ForeColor = &H706034
                .ListSubItems(4).ForeColor = &H706034
                .ListSubItems(5).ForeColor = &H706034
                .ListSubItems(6).ForeColor = &H706034
            ElseIf oInvoice.InvoiceLines(i).NonStock = True Then
                lstItem.ForeColor = &H427182
                lstItem.ListSubItems(1).ForeColor = &H427182
                lstItem.ListSubItems(2).ForeColor = &H427182
                lstItem.ListSubItems(3).ForeColor = &H427182
                lstItem.ListSubItems(4).ForeColor = &H427182
                lstItem.ListSubItems(5).ForeColor = &H427182
                lstItem.ListSubItems(6).ForeColor = &H427182
            End If
        
        
        End With
        
        
        
        
        
        If oInvoice.InvoiceLines(i).Note > "" Then
            Set lstItem = lvwInvLines.ListItems.Add
            lstItem.SubItems(1) = "Note:  " & oInvoice.InvoiceLines(i).Note
        End If
    Next i
    

    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
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

Private Sub Form_Load()
    
    Me.Top = 50
    Me.Left = 50
    Me.Height = 6500
    Me.Width = 11500

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oInvoice = Nothing
    Me.Hide
End Sub


Private Sub lvwInvLines_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub


'Private Sub mnuFileEdit_Click()
'Dim ofrm As New frmInvoice
'Dim blnEdit As Boolean
'
'    On Error GoTo ERR_Handler
'
'    If mnuFileEdit.Caption = "&Edit" Then
'        blnEdit = True
'        ofrm.Component oInvoice ', lngID
'        ofrm.Show vbModal
'    Else
'        blnEdit = False
'        ofrm.PrintInvoice
'    End If
'
'EXIT_Handler:
'    Unload ofrm
'    Set ofrm = Nothing
'    Unload Me
'    Exit Sub
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'End Sub
'
Private Sub mnuFileExit_Click()
    Me.Hide
End Sub


