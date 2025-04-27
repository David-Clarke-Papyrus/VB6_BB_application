VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCashBookMaintenance 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Cash book template"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   ClipControls    =   0   'False
   Icon            =   "frmCashbookMaintenance.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11850
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Close"
      Height          =   615
      Left            =   225
      Picture         =   "frmCashbookMaintenance.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5550
      Width           =   1000
   End
   Begin VB.ComboBox cboBankAccount 
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   345
      Width           =   3240
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   4830
      Left            =   210
      OleObjectBlob   =   "frmCashbookMaintenance.frx":0396
      TabIndex        =   0
      Top             =   690
      Width           =   11475
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Find file to import"
      Filter          =   "*.csv,*.txt"
      InitDir         =   "PBKS_S_"
   End
   Begin VB.Label lblBank 
      BackStyle       =   0  'Transparent
      Caption         =   "Bank account"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   315
      TabIndex        =   2
      Top             =   90
      Width           =   2835
   End
End
Attribute VB_Name = "frmCashBookMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flgLoading As Boolean

Dim strImportedFile As String
Dim x As New XArrayDB
Dim XGLA As XArrayDB
Dim rsImp As ADODB.Recordset
Dim rsAccs As ADODB.Recordset
Dim tlGLAccounts As New z_TextList
Dim rsDebtors  As ADODB.Recordset
Dim rsDRInvoices As ADODB.Recordset
Dim zD As New z_Debtors
Dim strFilename As String
'Dim frmRemittance As frmCustomerRemittance
Dim s As String
Dim iCur As Long
Dim iLCur As Long
Dim rs As ADODB.Recordset
Dim oFSO As New FileSystemObject
Dim rsBanks As ADODB.Recordset
Dim oSQL As z_SQL
Dim tlBanks As New z_TextList

Private Sub cboBankAccount_Click()
    If flgLoading Then Exit Sub
    LoadGridFromFile
    G.Enabled = True
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub G_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
    If Button = 2 Then   ' Check if right mouse button was clicked.
        mnuSelectDebtor
       ' PopupMenu Forms(0).mnuCashBookPopup  ' Display the File menu as a pop-up menu.
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCashBookMaintenance.G_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Public Sub mnuSelectDebtor()
Dim frm As New frmBrowseCustomers2
Dim TRID As Long

    G.Update
    frm.Show vbModal
    rsImp.Fields("CT_TPNAME") = Left(frm.CustomerName, 50)
    rsImp.Fields("CT_TradingPartnerID") = frm.CustomerID
    rsImp.Fields("CT_Account").Value = tlBanks.key(cboBankAccount)
    rsImp.Update

End Sub


Private Sub LoadGridFromFile()
 Dim i As Integer
 Dim vi As ValueItem
    oPC.OpenDBSHort
    Set rsAccs = New ADODB.Recordset
    rsAccs.CursorLocation = adUseClient
    rsAccs.Open "SELECT AC_ID,DESCR,VATRATE FROM vGLAccountsToView", oPC.COShort, adOpenStatic
    i = 0
    Do While Not rsAccs.eof
        i = i + 1
        Set vi = New ValueItem
        vi.Value = rsAccs.Fields("AC_ID")
        vi.DisplayValue = rsAccs.Fields("DESCR")
        G.Columns(5).ValueItems.Add vi
        rsAccs.MoveNext
    Loop
    G.Columns(5).ValueItems.Translate = True
    G.Columns(5).ValueItems.Presentation = dbgComboBox
    Set rsImp = New ADODB.Recordset
    rsImp.CursorLocation = adUseClient
    rsImp.Open "SELECT * FROM tCashbookTemplate WHERE CT_Account = '" & tlBanks.key(cboBankAccount) & "'", oPC.COShort, adOpenDynamic, adLockOptimistic
    
    G.DataSource = rsImp
    G.Refresh
    G.ReBind
End Sub


Private Sub Form_Load()
    flgLoading = True
    
    SetGridLayout Me.G, Me.Name
    SetFormSize Me
    Me.Top = 500
    Me.Left = 500
    
    
    Set rsBanks = New ADODB.Recordset
    Set oSQL = Nothing
    Set oSQL = New z_SQL
    Set tlBanks = New z_TextList
    tlBanks.Load ltBankAccounts
    
    LoadCombo cboBankAccount, tlBanks
'    oSQL.RunGetRecordset "SELECT Bank_Name + ' ' + Bank_AccountNumber as Description,Bank_AccountNumber FROM tBankAccount ORDER BY BANK_Name,BANK_AccountNumber", enText, Array(), "", rsBanks
'    LoadComboFromRecordset Me.cboBankAccount, rsBanks
    
    flgLoading = False
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
    If Me.Width > 7000 Then
        G.Width = NonNegative_Lng(Me.Width - 400)
    End If
    G.Height = NonNegative_Lng(Me.Height - 2400)
    cmdClose.Top = NonNegative_Lng(Me.Height - 1400)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
    
End Sub
'
