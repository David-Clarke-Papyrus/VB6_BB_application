VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExchange 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View transaction"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   75
      Picture         =   "frmExchange.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4380
      Width           =   1035
   End
   Begin VB.TextBox txtChangeGiven 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5610
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   4530
      Width           =   2430
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5625
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4170
      Width           =   2430
   End
   Begin VB.TextBox txtVAT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4155
      Width           =   2430
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5835
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   105
      Width           =   2430
   End
   Begin VB.TextBox txtOperator 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   2430
   End
   Begin MSComctlLib.ListView lvwPayments 
      Height          =   870
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3045
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   1535
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwSales 
      Height          =   1635
      Left            =   225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1335
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   2884
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4481
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Discount"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   1658
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment(s)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3060
      TabIndex        =   9
      Top             =   3075
      Width           =   1035
   End
End
Attribute VB_Name = "frmExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oEx As New a_Exchange
Dim bPrint As Boolean

Public Sub component(pExchangeID As String)
    oEx.Load pExchangeID, True
    
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
    bPrint = True
    Me.Hide
End Sub
Public Function MustPrint() As Boolean
    MustPrint = bPrint
End Function

Private Sub Form_Load()
    If oEx.transactionType = "CN" Then
        Me.Caption = "CREDIT NOTE for : " & oEx.Customer.NameAndCode(50)
        Me.txtOperator = oEx.StaffName
        Me.txtTotal = "Total: " & oEx.TotalPayableF
        txtDate = oEx.ExchangeDateTimeF
        txtChangeGiven = ""
    Else
        Me.Caption = "Sale"
        Me.txtOperator = oEx.StaffName
        Me.txtTotal = "Total: " & oEx.TotalPayableF
        txtDate = oEx.ExchangeDateTimeF
        txtChangeGiven = "Change: " & oEx.ChangeGivenF
    End If
    txtVAT = "VAT: " & oEx.TotalVATF
    lvwSales.ColumnHeaders(2).Width = 2800
    lvwPayments.ColumnHeaders(1).Width = 2500
    LoadPayments
    LoadSales
  '  cmdClose.SetFocus
End Sub

Private Sub LoadSales()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    lvwSales.ListItems.Clear
    For i = 1 To oEx.SaleLines.Count
        Set lstItem = lvwSales.ListItems.Add
        With oEx.SaleLines(i)
            lstItem.Text = .CodeF
'            If lstItem.Key = "" Then lstItem.Key = .Key
            lstItem.SubItems(1) = .title
            lstItem.SubItems(2) = .Qty
            lstItem.SubItems(3) = .PriceF
            lstItem.SubItems(4) = .DiscountRateF
            lstItem.SubItems(5) = .PLessDiscExtF
        End With

    Next i
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchange.LoadSales"
End Sub
Private Sub LoadPayments()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    lvwPayments.ListItems.Clear
    For i = 1 To oEx.PaymentLines.Count
        Set lstItem = lvwPayments.ListItems.Add
        With oEx.PaymentLines(i)
            lstItem.Text = .PaymentTypeF & IIf(.ReferenceComplete > "", "(" & .ReferenceComplete & ")", .ReferenceComplete)
            lstItem.SubItems(1) = .AmtF
        End With

    Next i
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExchange.LoadPayments"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oEx = Nothing
End Sub

Private Sub lvwPayments_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvwSales_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub
