VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form frmCashBook 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Cash book"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15000
   ClipControls    =   0   'False
   Icon            =   "frmCashbook.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   15000
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   510
      Left            =   9870.001
      Top             =   60
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   900
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00F7EDE8&
      Height          =   660
      Left            =   225
      TabIndex        =   7
      Top             =   150
      Width           =   6045
      Begin VB.ComboBox cboBankAccount 
      ForeColor       =   &H8000000D&
      Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   210
         Width           =   2865
      End
      Begin MSComCtl2.DTPicker dpSince 
         Height          =   360
         Left            =   4605
      TabIndex        =   8
         Top             =   195
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Format          =   60817409
         CurrentDate     =   40525
   End
      Begin VB.Label lblBank 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank account"
      ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   75
         TabIndex        =   11
         Top             =   255
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4140
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdSplit 
      BackColor       =   &H00E7E6D8&
      Caption         =   "Split"
      Enabled         =   0   'False
      Height          =   540
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00E7E6D8&
      Caption         =   "Import"
      Enabled         =   0   'False
      Height          =   540
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdMatch 
      BackColor       =   &H00E7E6D8&
      Caption         =   "Auto allocate debtor/account"
      Height          =   540
      Left            =   2145
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Close"
      Height          =   615
      Left            =   225
      Picture         =   "frmCashbook.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5550
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   4680
      Left            =   240
      OleObjectBlob   =   "frmCashbook.frx":0396
      TabIndex        =   0
      Top             =   795
      Width           =   14370
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
   Begin VB.Label lblClosingBalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6420
      TabIndex        =   3
      Top             =   5760
      Width           =   3030
   End
   Begin VB.Label lblOpeningBalance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6420
      TabIndex        =   2
      Top             =   375
      Width           =   3030
   End
End
Attribute VB_Name = "frmCashBook"
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
Dim rsDups As ADODB.Recordset
Dim zD As New z_Debtors
Dim strFilename As String
'Dim frmRemittance As frmCustomerRemittance
Dim frmDups As frmDuplicatedStatementLines
Dim ofrmR As frmCRemittancePreview
Dim transRow As dStatementProps

Dim s As String
Dim iCur As Long
Dim iLCur As Long
Dim zFile As z_TextFile
Dim rs As ADODB.Recordset
Dim oFSO As New FileSystemObject
Dim rsBanks As ADODB.Recordset
Dim oSQL As z_SQL
Dim tlBanks As New z_TextList
Dim dblOpeningBalance As Double
Dim dblClosingBalance As Double


Private Sub cboBankAccount_Click()
    If flgLoading Then Exit Sub
    LoadGridFromFile
    dblOpeningBalance = oSQL.CalculateStatementBalance(tlBanks.key(cboBankAccount), dpSince, True)
    cmdImport.Enabled = True
    G.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdMatch_Click()
    Dim oSQL As New z_SQL
    oSQL.MatchCashbook
    LoadGridFromFile
End Sub



Private Sub dpSince_Change()
    LoadGridFromFile
    dblOpeningBalance = oSQL.CalculateStatementBalance(tlBanks.key(cboBankAccount), dpSince, True)
    lblOpeningBalance.Caption = "Opening balance: " & Format(dblOpeningBalance, oPC.Configuration.DefaultCurrency.FormatString)
    dblClosingBalance = oSQL.CalculateStatementBalance(tlBanks.key(cboBankAccount), CDate("2099-01-01"), True)
    lblClosingBalance.Caption = "Closing balance: " & Format(dblClosingBalance, oPC.Configuration.DefaultCurrency.FormatString)
End Sub

Private Sub dpSince_Click()
    LoadGridFromFile
    dblOpeningBalance = oSQL.CalculateStatementBalance(tlBanks.key(cboBankAccount), dpSince, True)
End Sub


Private Sub G_ButtonClick(ByVal ColIndex As Integer)
    mnuSelectDebtor
End Sub

Private Sub G_OnAddNew()
    G.Columns(5).Value = "0"
    G.Columns(6).Value = "0"
End Sub

Private Sub G_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
    If Button = 2 Then   ' Check if right mouse button was clicked.
        PopupMenu Forms(0).mnuCashbookPopup  ' Display the File menu as a pop-up menu.
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub mnuSelectDebtor()
Dim frm As New frmBrowseCustomers2
Dim TRID As Long

    G.Update
    frm.Show vbModal
    rsImp.Fields("ST_DebtorName") = Left(frm.CustomerName, 50)
    rsImp.Fields("ST_PostedToTPID") = frm.CustomerID
    rsImp.Update
    
    'Create customer payment if credit and if Debtor is selected and not already exists
    If UCase(rsImp.Fields("ST_TRANSACTIONTYPE")) = "CREDIT" And _
        FNN(rsImp.Fields("ST_PostedToTPID")) > 0 Then
        Set oSQL = New z_SQL
        TRID = oSQL.CreatePayment(FNN(rsImp.Fields("ST_PostedToTPID")), _
                            FNDBL(rsImp.Fields("ST_TransCredit")), _
                            FNN(rsImp.Fields("Statement_ID")), _
                            FNS(rsImp.Fields("ST_Reference")), _
                            FND(rsImp.Fields("ST_TransDatePosted")))
    End If
        
        
            
    
End Sub
Public Sub mnuPaymentMatch()
Dim frm As New frmPaymentMatch
Dim oCust As New a_Customer
Dim Res As Boolean

    Res = oCust.Load(FNN(rsImp.Fields("ST_PostedToTPID")))
    If Res Then
        frm.Component oCust.ID, oCust.NameAndCode(50), top, Left
        frm.Show
    End If

End Sub
Public Sub mnuLoadRemittance()

'    If FNN(rsImp.Fields("Statement_ID")) > 0 Then   'there is already a remittance
'        Set ofrmR = New frmCRemittancePreview
'        ofrmR.Component 0, FNS(rsImp.Fields("ST_DebtorName")), FNN(rsImp.Fields("Statement_ID"))
'        ofrmR.Show
'    Else
'        If FNN(rsImp.Fields("ST_PostedToTPID")) = 0 Then Exit Sub
'        Set frmRemittance = New frmCustomerRemittance
'        frmRemittance.Component rsImp.Fields("ST_PostedToTPID"), FNS(rsImp.Fields("ST_DebtorName")), Me.hWnd, Date, "", FNN(rsImp.Fields("Statement_ID"))
'        frmRemittance.Show
'    End If
End Sub
Private Sub cmdImport_Click()

    CD1.InitDir = GetSetting("PBKS" & "ImportCashBook", "ImportFromFile", "SourceFolder", oPC.SharedFolderRoot)
    CD1.ShowOpen
    SaveSetting "PBKS" & "ImportCashBook", "ImportFromFile", "SourceFolder", oFSO.GetParentFolderName(strFilename)
    LoadFromFileToTable (CD1.FileName)
    LoadGridFromFile
End Sub

Private Sub LoadGridFromFile()
 Dim i As Integer
 Dim vi As ValueItem
    oPC.OpenDBSHort
    Set rsAccs = New ADODB.Recordset
    rsAccs.CursorLocation = adUseClient
    rsAccs.Open "SELECT Category,Description,VAT,ID FROM vAccounts", oPC.COShort, adOpenStatic
    i = 0
    Do While Not rsAccs.eof
        i = i + 1
        Set vi = New ValueItem
        vi.Value = rsAccs.Fields("ID")
        vi.DisplayValue = rsAccs.Fields("Description")
        G.Columns(3).ValueItems.Add vi
        rsAccs.MoveNext
    Loop
    G.Columns(3).ValueItems.Translate = True
    G.Columns(3).ValueItems.Presentation = dbgComboBox
    Set rsImp = New ADODB.Recordset
    rsImp.CursorLocation = adUseClient
    rsImp.Open "SELECT * FROM tStatement WHERE ST_BANKACCOUNT = '" & tlBanks.key(cboBankAccount) & "' AND ST_TransDatePosted >= '" & Format(Me.dpSince, "YYYY-MM-DD") & "' ORDER BY ST_FITID ", oPC.COShort, adOpenDynamic, adLockOptimistic
    Set Adodc1.Recordset = rsImp
    Set G.DataSource = Me.Adodc1
    G.Refresh
    G.ReBind
End Sub


Private Sub Form_Load()
    flgLoading = True
    
    SetGridLayout Me.G, Me.Name
    SetFormSize Me
    Me.top = 500
    Me.Left = 500
    
    Me.dpSince = FirstOfMonth(DateAdd("m", -1, Date))
    
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
    cmdclose.top = NonNegative_Lng(Me.Height - 1400)
    cmdImport.top = cmdclose.top
    cmdMatch.top = cmdclose.top
    cmdSplit.top = cmdclose.top
    lblClosingBalance.top = cmdclose.top
    lblClosingBalance.Left = Me.lblOpeningBalance.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.G, Me.Name, Me.Height, Me.Width
    
End Sub

Private Sub G_AfterUpdate()
Dim i As Integer
'    x(G.Bookmark, 15) = XGLA.Find(1, 1, x(G.Bookmark, 3), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
    
'   ' rsImp.Find ""
'    rsImp.MoveFirst
'    rsImp.Find "Statement_ID = " & FNS(x(G.Bookmark, 10))
'    rsImp.Fields("ST_PostedToAccountID") = FNS(FNN(x(G.Bookmark, 15)))
'    rsImp.Update
End Sub

Private Sub G_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim i As Integer

 '   i = ColIndex + 1
 '   i = X.Find(1, 10, FNS(X(G.Text, 10)), XORDER_ASCEND, XCOMP_EQ, XTYPE_INTEGER)
    Select Case i
    Case 2
  '      oDC.style = Trim(Grid1.Text)
    Case 3
     '    X(G.Bookmark, 15) = XGLA.Find(1, 1, X(G.Bookmark, 3), XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
    Case 4
  '      oDC.SetPrinter oConfig.Printers.key(Trim(Grid1.Text)), Trim(Grid1.Text)
    Case 5
   '     If ConvertToLng(Grid1.Text, lngResult) Then
   '         oDC.QtyCopies = lngResult
   '     End If
    End Select

End Sub

Private Sub G_Click()

    zD.AlDebtors rsDebtors
   ' LoadCombo me.cboDebtor,
  ' me.cboDebtor.
End Sub

Private Sub LoadFromFileToTable(fpath As String)
Set zFile = New z_TextFile
Dim oSQL As New z_SQL
    
    If fpath = "" Then Exit Sub
    oSQL.RunSQL "Truncate TABLE tmpStatement"
    Set rs = New ADODB.Recordset
    oSQL.GetDynamicRecordset "SELECT * FROM tmpStatement", enText, Array(), "", rs
    
    iCur = 1
    zFile.OpenTextFileToRead fpath
    s = zFile.ReadWholeFile
    s = stripCRLF(s)
    transRow.BankID = GetTagValue(s, "<BANKID>")
    transRow.ACCTID = GetTagValue(s, "<ACCTID>")
    Do While GetDetailLine(s) > ""
    Loop
    zFile.CloseTextFileNoErrors
    rs.Close
    Set rs = Nothing
    
    'Check for duplicates
    oSQL.RunSQL "UPDATE tmpStatement SET Status = 'D' FROM tmpStatement a JOIN vPossibleDuplicatedStatementLines b ON  a.ID = b.ID"
    Set rsDups = Nothing
    Set rsDups = New ADODB.Recordset
    rsDups.CursorLocation = adUseClient
    oSQL.GetDynamicRecordset "SELECT ID,TRNTYPE,CASE WHEN TRNTYPE = 'CREDIT' THEN TRNAMT ELSE 0 END AS CREDIT,CASE WHEN TRNTYPE = 'DEBIT' THEN TRNAMT ELSE 0 END AS DEBIT,CONVERT(DATETIME,DTPOSTED,120) as DTPOSTED,MEMO,STATUS FROM tmpStatement WHERE STATUS = 'D'", enText, Array(), "", rsDups
    If rsDups.RecordCount > 0 Then
        Set frmDups = New frmDuplicatedStatementLines
        frmDups.Component rsDups, "Possible duplicates found in importing to account " & Me.cboBankAccount & "."
        frmDups.Show vbModal
    End If
End Sub

Private Function GetTagValue(s As String, Tag As String) As String
Dim i As Long
Dim iDelim As Long
Dim iLen As Integer

    iLen = Len(Tag)
    iCur = InStr(iCur, s, Tag)
    iDelim = InStr(iCur + 1, s, "<")
    GetTagValue = Mid(s, iCur + iLen, iDelim - (iCur + iLen))
    iCur = iDelim

End Function
Private Function GetPartialTagValue(s As String, Tag As String) As String
Dim i As Long
Dim iDelim As Long
Dim iLen As Integer

    iLen = Len(Tag)
    iLCur = InStr(iLCur, s, Tag)
    iDelim = InStr(iLCur + 1, s, "<")
    If iDelim > 0 Then
        GetPartialTagValue = Mid(s, iLCur + iLen, iDelim - (iLCur + iLen))
    Else
        GetPartialTagValue = Mid(s, iLCur + iLen)
    End If
    iLCur = iDelim

End Function

Private Function GetDetailLine(s As String) As String
Dim i As Long
Dim iDelim As Long
Dim d As String

    iLCur = 1
    d = ""
    iCur = InStr(iCur, s, "<STMTTRN>")
    If iCur > 0 Then
        iDelim = InStr(iCur + 1, s, "</STMTTRN>")
        d = Mid(s, iCur + 9, iDelim - (iCur + 10))
        iCur = iDelim
         transRow.TRNTYPE = GetPartialTagValue(d, "<TRNTYPE>")
         transRow.DTPOSTED = GetPartialTagValue(d, "<DTPOSTED>")
         transRow.TRNAMT = vbCrLf & GetPartialTagValue(d, "<TRNAMT>")
         transRow.FITID = vbCrLf & GetPartialTagValue(d, "<FITID>")
         transRow.MEMO = vbCrLf & GetPartialTagValue(d, "<MEMO>")
    End If
    If d > "" Then
        rs.AddNew
        rs.Fields("BANKID") = FNS(transRow.BankID)
        rs.Fields("ACCTID") = FNS(transRow.ACCTID)
        rs.Fields("FITID") = stripCRLF(FNS(transRow.FITID))
        rs.Fields("TRNTYPE") = FNS(transRow.TRNTYPE)
        rs.Fields("DTPOSTED") = FNS(transRow.DTPOSTED)
        rs.Fields("TRNAMT") = CDbl(stripCRLF(transRow.TRNAMT))
        rs.Fields("MEMO") = stripCRLF(FNS(transRow.MEMO))
        rs.Update
    End If
    GetDetailLine = d
End Function

Private Sub G_BeforeUpdate(Cancel As Integer)
    If Adodc1.Recordset.EditMode = adEditAdd Then
        Adodc1.Recordset.Fields("ST_BankAccount") = tlBanks.key(cboBankAccount)
        Adodc1.Recordset.Fields("ST_FITID") = "D333"
        If rsImp.Fields("ST_TRANSDEBIT") <> 0 Then
            Adodc1.Recordset.Fields("ST_TransactionType") = "DEBIT"
            rsImp.Fields("ST_TRANSCREDIT") = 0
        Else
            Adodc1.Recordset.Fields("ST_TransactionTYpe") = "CREDIT"
            rsImp.Fields("ST_TRANSDEBIT") = 0
        End If
    End If
End Sub

