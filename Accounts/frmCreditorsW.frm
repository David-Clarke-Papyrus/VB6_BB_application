VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{B4B5B73C-172E-47B1-BFC2-C6F740957D01}#1.0#0"; "VB Control Manager.ocx"
Begin VB.Form frmCreditors 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Creditors"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16470
   ControlBox      =   0   'False
   FillColor       =   &H00FCF2EB&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   16470
   WindowState     =   2  'Maximized
   Begin VBControlManager.ControlManager CM 
      Height          =   10530
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   18574
      BackColor       =   255
      Size            =   4
      TitleBar_CloseVisible=   0   'False
      Begin VB.Frame fr3 
         BackColor       =   &H00F7EDE8&
         Height          =   10350
         Left            =   8010
         TabIndex        =   30
         Top             =   180
         Width           =   7995
         Begin TrueOleDBGrid60.TDBGrid PaymentsGrid 
            Height          =   3240
            Left            =   180
            OleObjectBlob   =   "frmCreditorsW.frx":0000
            TabIndex        =   31
            Top             =   300
            Width           =   6300
         End
      End
      Begin VB.Frame fr2 
         BackColor       =   &H00F7EDE8&
         Height          =   6795
         Left            =   0
         TabIndex        =   27
         Top             =   3735
         Width           =   7965
         Begin VB.CommandButton cbSince 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7E6D8&
            Caption         =   "Last week"
            Height          =   450
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   630
            Width           =   7530
         End
         Begin TrueOleDBGrid60.TDBGrid LedgerGrids 
            Height          =   3240
            Left            =   585
            OleObjectBlob   =   "frmCreditorsW.frx":4C59
            TabIndex        =   29
            Top             =   1305
            Width           =   6855
         End
      End
      Begin VB.Frame fr1 
         BackColor       =   &H00F7EDE8&
         Height          =   3345
         Left            =   -15
         TabIndex        =   1
         Top             =   180
         Width           =   7950
         Begin VB.TextBox txtArg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00915A48&
            Height          =   375
            Left            =   720
            TabIndex        =   25
            Text            =   "<Creditor by name or A/C no>"
            Top             =   195
            Width           =   5265
         End
         Begin VB.Frame frBalances 
            BackColor       =   &H00F9F2EE&
            ForeColor       =   &H8000000D&
            Height          =   1200
            Left            =   15
            TabIndex        =   4
            Top             =   3855
            Width           =   7485
            Begin VB.TextBox txtCurBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
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
               Height          =   315
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   360
               Width           =   960
            End
            Begin VB.TextBox txt30Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
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
               Height          =   315
               Left            =   3255
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txt60Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
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
               Height          =   315
               Left            =   4290
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txt90Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
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
               Height          =   315
               Left            =   5325
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txt120PlusBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
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
               Height          =   315
               Left            =   6345
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txtBalance 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   360
               Width           =   1110
            End
            Begin VB.TextBox txtBF120PlusBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
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
               Height          =   315
               Left            =   6345
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBF90Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
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
               Height          =   315
               Left            =   5325
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBF60Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
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
               Height          =   315
               Left            =   4290
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBF30Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
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
               Height          =   315
               Left            =   3255
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBFCurBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
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
               Height          =   315
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   750
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.TextBox txtBFBalance 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   750
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total balance"
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
               Height          =   255
               Left            =   1140
               TabIndex        =   24
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "This month"
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
               Height          =   255
               Left            =   2250
               TabIndex        =   23
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "30 days"
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
               Height          =   255
               Left            =   3240
               TabIndex        =   22
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "60 days"
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
               Height          =   255
               Left            =   4260
               TabIndex        =   21
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "90 days"
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
               Height          =   255
               Left            =   5310
               TabIndex        =   20
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "120+ days"
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
               Height          =   255
               Left            =   6270
               TabIndex        =   19
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Current"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   330
               TabIndex        =   18
               Top             =   390
               Width           =   675
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Month start"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   60
               TabIndex        =   17
               Top             =   780
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdMatchPayments 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Match payments"
            Height          =   450
            Left            =   7620
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3870
            Width           =   870
         End
         Begin VB.CommandButton cmdInvoiceAge 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7E6D8&
            Caption         =   "Last week"
            Height          =   450
            Left            =   6015
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   165
            Width           =   1320
         End
         Begin TrueOleDBGrid60.TDBGrid Grid 
            Height          =   2265
            Left            =   705
            OleObjectBlob   =   "frmCreditorsW.frx":98B1
            TabIndex        =   26
            Top             =   765
            Width           =   7095
         End
      End
   End
End
Attribute VB_Name = "frmCreditors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmRemittancesBrowse As frmBrowseDBJNLs
Dim frmCashBookMaintenance As frmCashBookMaintenance
Dim frmRemittancePreview As frmCRemittancePreview
Dim frmCustJnl As frmCustJnl
Dim frmCashBook As frmCashBook
Dim cVendors As c_Supplier

Dim dteDate1 As Date
Dim dteDate2 As Date
Dim cInvoices As c_Invoices
Dim XA As New XArrayDB
Dim enSince As enumSince
Dim rs As New ADODB.Recordset
Dim oSQL As New z_SQL
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
Dim strCustomerName As String
Dim oTRs As c_DebtorsTransPerTP
Dim XB As New XArrayDB
Dim flgLoading As Boolean
Dim oVendor As a_Supplier
Dim Res As Boolean
Dim oSM As z_StockManager
Private Sub cbSince_Click()
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
End Sub

Private Sub CM_SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    resize
End Sub


Private Sub Form_Load()
    flgLoading = True
    Me.Grid.Top = 750
    Me.LedgerGrids.Top = 750
    Me.cbSince.Top = 270
    Me.Top = 200
    Me.Left = 50
    Me.Width = 6600
    
    Me.Height = 4000
    enSince = OptionLoop(enSince, 5)
    
    cbSince.Caption = TranslateSince(CInt(enSince))
    SetDateArgs
    SetGridLayout Me.Grid, Me.Name & Grid.Name
    SetFormSize Me
    Me.WindowState = vbNormal
    flgLoading = False
'    Me.CM.Controls(0).TitleBar_CloseVisible = False
'    Me.CM.Controls(1).TitleBar_CloseVisible = False
'    Me.CM.Controls(2).TitleBar_CloseVisible = False
    
End Sub
Private Sub LoadBrowse()
'    Find
'    LoadArray
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Supplier
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.ReDim 1, cVendors.Count, 1, 6
    For lngIndex = 1 To cVendors.Count
        With objItem
            Set objItem = cVendors.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Name
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
           ' XA.Value(lngIndex, 4) = objItem.Balance
            XA.Value(lngIndex, 5) = objItem.ID
          '  XA.Value(lngIndex, 6) = objItem.Blocked
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadArray"
End Sub

Private Sub Grid_DblClick()
Dim lngID As Long
Dim bNotFound As Boolean

    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Set oVendor = Nothing
    Set oVendor = New a_Supplier
    oVendor.Load FNN(XA(Grid.Bookmark, 5))
   
    LoadTransactions FNN(XA(Grid.Bookmark, 5))
    
    LoadLedgerGrid
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadTransactions(Optional CustID As Long, Optional pAcno As String)
    On Error GoTo errHandler
    Set oTRs = Nothing
    Set oTRs = New c_DebtorsTransPerTP
    oTRs.Load CustID, CDate("2000-01-01"), False, pAcno
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadTransactions"
End Sub


Private Sub LoadStatement(Optional pAcno As String)
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double

'    If oTRs.Count = 0 Then
'        XB.Clear
'        LedgerGrids.ReBind
'        Exit Sub
'    End If
'
'    For i = 1 To LedgerGrids.Columns.Count
'        LedgerGrids.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "LedgerGrids", CStr(i), LedgerGrids.Columns(i - 1).Width)
'    Next
'
'    XB.Clear
'    LedgerGrids.ReBind
'   ' Set oSQL = New z_SQL
'   ' oSM.RecalculateTPBalance oVendor.ID
''    Me.txtCurBal = oVendor.BalanceCurF
''    Me.txt30Bal = oVendor.Balance30F
''    Me.txt60Bal = oVendor.Balance60F
''    Me.txt90Bal = oVendor.Balance90F
''    Me.txt120PlusBal = oVendor.Balance120F
''    Me.txtBalance = oVendor.BalanceF
'
'    i = 1
'    j = 1
'    Do While i <= oTRs.Count
'        If oTRs.Item(i).DocType <> "BF" Then
'            XB.ReDim 1, j, 1, 8
'            XB.Value(j, 1) = oTRs.Item(i).DOCCode
'            XB.Value(j, 2) = oTRs.Item(i).DocType
'            XB.Value(j, 3) = oTRs.Item(i).DocDateF
'            XB.Value(j, 4) = oTRs.Item(i).DebitF
'            XB.Value(j, 5) = oTRs.Item(i).CreditF
'            XB.Value(j, 6) = oTRs.Item(i).MEMO
'            XB.Value(j, 7) = oTRs.Item(i).DOCID
'            XB.Value(j, 7) = oTRs.Item(i).DOCCaptureDate
' '           dblBal = dblBal + oTRs.Item(i).Debit
' '           dblBal = dblBal - oTRs.Item(i).Credit
'            j = j + 1
'        End If
'        i = i + 1
'    Loop
'    XB.QuickSort 1, XB.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
'    LedgerGrids.Array = XB
'    LedgerGrids.ReBind
'    LedgerGrids.Caption = oVendor.Fullname & " (" & oVendor.AcNo & ")"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.LoadStatement"
End Sub
Private Sub LoadTransactions(Optional CustID As Long, Optional pAcno As String)
    On Error GoTo errHandler
    Set oTRs = Nothing
    Set oTRs = New c_DebtorsTransPerTP
    oTRs.Load CustID, CDate("2000-01-01"), False, pAcno
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadTransactions"
End Sub


'Private Sub Find()
'    On Error GoTo errHandler
'Dim bNotFound As Boolean
'Dim frm As frmBrowseCustomers2
'Dim lngTPID As Long
'Dim byear As Boolean
'Dim yr As String
'Dim mth As String
'Dim strDate1 As String
'Dim strDate2 As String
'Dim lngCount As Long
'
'    bNotFound = False
'    If Left(txtArg, 3) = "yr=" Then byear = True
'
'    If txtArg > " " And Not (byear) Then
'        'Search for Reference
'        Set cJNL = Nothing
'        Set cJNL = New c_JNL
'        cJNL.Load bNotFound, 0, "", txtArg, dteDate1, dteDate2
'        If bNotFound Then
'            'Search for customer by AcJNLO
'            Set cJNL = Nothing
'            Set cJNL = New c_JNL
'            SetDateArgs
'            cJNL.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
'            If bNotFound Then
'               Set frm = New frmBrowseCustomers2
'               frm.component txtArg, lngCount
'               If lngCount > 1 Then
'                    frm.Show vbModal
'                    lngTPID = frm.CustomerID
'                    Unload frm
'                ElseIf lngCount = 1 Then
'                    lngTPID = frm.CustomerID
'                    Unload frm
'                End If
'               If lngTPID > 0 Then
'                    Set cJNL = Nothing
'                    Set cJNL = New c_JNL
'                    SetDateArgs
'                    cJNL.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
'               End If
'            End If
'        Else
'            enSince = 1
'            cbSince.Caption = TranslateSince(1)
'        End If
'    Else
'        Set cJNL = Nothing
'        Set cJNL = New c_JNL
'        If byear Then
'            yr = Mid(txtArg, 4, 4)
'            mth = Mid(txtArg, 9, 2)
'            If mth > "" Then
'                strDate1 = yr & "-" & mth & "-01"
'                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
'            Else
'                strDate1 = yr & "-01-01"
'                strDate2 = yr & "-12-31"
'            End If
'            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
'                SetDateArgs
'            Else
'                dteDate1 = CDate(strDate1)
'                dteDate2 = CDate(strDate2)
'            End If
'        Else
'            SetDateArgs
'        End If
'        cJNL.Load bNotFound, 0, "", "", dteDate1, dteDate2
'    End If
'
'EXIT_Handler:
'    mSetfocus Grid
'    MousePointer = vbDefault
'    Exit Sub
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmDebtors.Find"
'End Sub

Private Sub txtArg_GotFocus()
    AutoSelect txtArg
End Sub

Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    
    If KeyAscii = 13 Then  ' The ENTER key.
       HandleResults
        If cVendors.Count > 1 Then
            On Error Resume Next
            Grid.SetFocus
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub SetDateArgs()
    On Error GoTo errHandler
    Select Case enSince
    Case enAny
        dteDate1 = CDate("1995-01-01")
        dteDate2 = DateAdd("d", 1, Date)
    Case enWeek
        dteDate1 = DateAdd("d", -7, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enMonth
        dteDate1 = DateAdd("m", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enQuarter
        dteDate1 = DateAdd("q", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enYear
        dteDate1 = DateAdd("yyyy", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.SetDateArgs"
End Sub

Private Sub Form_Resize()
    resize
End Sub
Private Sub resize()
    Me.txtArg.Top = fr1.Top + 10
    Me.txtArg.Left = fr1.Left + 50
    
    Me.Grid.Left = fr1.Left + 50
    Me.Grid.Top = fr1.Top + 470
    Me.Grid.Width = NonNegative_Lng(fr1.Width - 400)
    Me.Grid.Height = NonNegative_Lng(fr1.Height - 1800)
    
    Me.cbSince.Top = fr2.Top - fr1.Height - 200
    Me.cbSince.Left = fr2.Left + 50
    
    frBalances.Left = fr1.Left + 50
    frBalances.Top = NonNegative_Lng(fr1.Height - 1000)
    frBalances.Height = 900
    Me.LedgerGrids.Left = fr2.Left + 50
    Me.LedgerGrids.Top = fr2.Top - fr1.Height + 300
    Me.LedgerGrids.Width = NonNegative_Lng(fr2.Width - 400)
    Me.LedgerGrids.Height = NonNegative_Lng(fr2.Height - 1000)

    Me.PaymentsGrid.Left = fr3.Left + 50
    Me.PaymentsGrid.Top = fr3.Top - fr3.Height + 300
    Me.PaymentsGrid.Width = NonNegative_Lng(fr3.Width - 400)
    Me.PaymentsGrid.Height = NonNegative_Lng(fr3.Height - 1000)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.Grid, Me.Name & Grid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        LoadBrowse
    End If
End Sub


Private Sub HandleResults(Optional plngCount As Long)
    On Error GoTo errHandler
    If txtArg = "" Then Exit Sub
    Set cVendors = Nothing
    Set cVendors = New c_Supplier
    Screen.MousePointer = vbHourglass
    
    cVendors.LoadEasy Replace(txtArg, "'", "''"), False ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    plngCount = cVendors.Count
    LoadArray
    Grid.ReBind
    Grid.Enabled = True
    If cVendors.Count = 1 Then
        LoadTransactions FNN(XA(Grid.Bookmark, 5))
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.HandleResults(plngCount)", plngCount
End Sub

