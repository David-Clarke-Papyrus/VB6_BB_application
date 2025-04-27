VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Papyrus debtors, creditors and cashbook"
   ClientHeight    =   6960
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11490
   Icon            =   "MDI_Accounts.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   6525
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14605
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuDebtors 
         Caption         =   "Debtors"
      End
      Begin VB.Menu mnuCreditors 
         Caption         =   "Creditors"
      End
      Begin VB.Menu mnuCashbook 
         Caption         =   "Cashbook"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New"
      Begin VB.Menu mnuCustomerMaster 
         Caption         =   "Customer master record"
      End
      Begin VB.Menu mnuNewCRemittance 
         Caption         =   "Customer remittance"
      End
      Begin VB.Menu mnuNewCPayment 
         Caption         =   "Customer payment"
      End
      Begin VB.Menu mnuNewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewSupplierMaster 
         Caption         =   "Supplier master record"
      End
      Begin VB.Menu mnuNewSRemittance 
         Caption         =   "Supplier remittance"
      End
      Begin VB.Menu mnuNewSPayment 
         Caption         =   "Supplier payment"
      End
   End
   Begin VB.Menu mnuBrowse 
      Caption         =   "Browse"
      Begin VB.Menu mnuBrowseCustomers 
         Caption         =   "Customers"
      End
      Begin VB.Menu mnuBrowseSuppliers 
         Caption         =   "Suppliers"
      End
   End
   Begin VB.Menu mnuMonthEnd 
      Caption         =   "MonthEnd"
      Begin VB.Menu mnuPrepareStatements 
         Caption         =   "Prepare statements"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "Actions"
      Begin VB.Menu mnuPreparePayments 
         Caption         =   "Prepare payments"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuCashbookTemplate 
         Caption         =   "Cashbook template"
      End
      Begin VB.Menu mnuChartOfAccounts 
         Caption         =   "Chart of accounts"
      End
   End
   Begin VB.Menu mnuDebtorPopup 
      Caption         =   "DebtorPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuNewRemittance 
         Caption         =   "New remittance"
      End
   End
   Begin VB.Menu mnuCashbookPopup 
      Caption         =   "CashbookPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCashbookDebtorSelect 
         Caption         =   "Debtor"
      End
      Begin VB.Menu mnuPaymentMatch 
         Caption         =   "Match payments"
      End
      Begin VB.Menu mnuRemittance 
         Caption         =   "Remittance"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmDebtors As frmDebtors
Dim frmCreditors As frmCreditors
Dim frmCustomerRemittance As frmCustomerRemittance
Dim frmCRemittancePreview As frmCRemittancePreview
Dim frmCustJnl As frmCustJnl
Dim frmCashBook As frmCashBook
Dim frmCashBookMaintenenance As frmCashBookMaintenance
Dim frmStatementControl As frmStatementControl
Dim frmCustomer As frmCustomer
Dim frmBrowseCustomers As frmBrowseCustomers
Dim frmNewCustomer As frmNewCustomer
Dim frmChartOfAccounts As frmChartOfAccounts

Private Sub MDIForm_Load()
    Me.SB1.Panels(1) = "Last day-end: " & oPC.Configuration.LastUpdateDateF & "   "
    Me.SB1.Panels(2) = " " & oPC.NewQuotation
    Me.SB1.Panels(3) = "user:" & oPC.UserName & ", " & IIf(oPC.DatabaseName <> "PBKS", "server:" & oPC.servername & ", database:" & oPC.DatabaseName, "Server:" & oPC.servername)
    SB1.Panels(2).ToolTipText = SB1.Panels(2).Text
End Sub

Private Sub mnuBrowseCustomers_Click()
    Set frmBrowseCustomers = New frmBrowseCustomers
    frmBrowseCustomers.Show
End Sub

Private Sub mnuBrowseSuppliers_Click()

    Set frmBrowsesuppliers = New frmBrowsesuppliers
    frmBrowsesuppliers.Show

End Sub

Private Sub mnuCashbook_Click()
    Set frmCashBook = New frmCashBook
    frmCashBook.Show
End Sub

Private Sub mnuCashbookTemplate_Click()
    Set frmCashBookMaintenance = New frmCashBookMaintenance
    frmCashBookMaintenance.Show
End Sub

Private Sub mnuChartOfAccounts_Click()
    Set frmChartOfAccounts = New frmChartOfAccounts
    frmChartOfAccounts.Show
End Sub

Private Sub mnuCreditors_Click()
    Set frmCreditors = New frmCreditors
    frmCreditors.Show
End Sub

Private Sub mnuCustomerMaster_Click()
    Set frmNewCustomer = New frmNewCustomer
    frmNewCustomer.Show vbModal
    
End Sub

Private Sub mnuDebtors_Click()
    Set frmDebtors = New frmDebtors
    frmDebtors.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNewRemittance_Click()
    
    Me.ActiveForm.mnuNewRemittance
    
End Sub

Private Sub mnuNewCRemittance_Click()
'    Set frmCustomerRemittance = New frmCustomerRemittance
'
'    frmCustomerRemittance.Component
'    frmCustomerRemittance.Show
'
End Sub

Private Sub mnuPreparePayments_Click()
Dim frmCreditorsPayments As New frmCreditorsPayments
     
    frmCreditorsPayments.Show

End Sub

Private Sub mnuPrepareStatements_Click()
Set frmStatementControl = New frmStatementControl
    frmStatementControl.Show
End Sub

Public Sub NewCustomer(pType As enumCustomerType)
    On Error GoTo errHandler
Dim frm As frmCustomer
Dim oCust As a_Customer
    Set frm = New frmCustomer
    Set oCust = New a_Customer
    oCust.BeginEdit
    oCust.InitializeNewCustomer pType
    frm.Component oCust
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewCustomer(pType)", pType
End Sub
Private Sub mnuCashbookDebtorSelect_Click()
    Me.ActiveForm.mnuSelectDebtor
End Sub
Private Sub mnuPaymentMatch_Click()
    Me.ActiveForm.mnuPaymentMatch
End Sub
Private Sub mnuRemittance_Click()
    Me.ActiveForm.mnuLoadRemittance
End Sub

