VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Object = "{B4B5B73C-172E-47B1-BFC2-C6F740957D01}#1.0#0"; "VB Control Manager.ocx"
Begin VB.Form frmDebtors 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Debtors system"
   ClientHeight    =   12705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   ControlBox      =   0   'False
   FillColor       =   &H00FCF2EB&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12705
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
   Begin VBControlManager.ControlManager CM1 
      Height          =   7095
      Left            =   2565
      TabIndex        =   27
      Top             =   5745
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   12515
      Size            =   4
      Begin VB.Frame fr2 
         Height          =   1664
         Left            =   0
         TabIndex        =   31
         Top             =   5431
         Width           =   6885
         Begin TrueOleDBGrid60.TDBGrid TDBGridled 
            Height          =   3240
            Left            =   0
            OleObjectBlob   =   "frmDebtors.frx":0000
            TabIndex        =   32
            Top             =   0
            Width           =   8445
         End
      End
      Begin VB.Frame frm1 
         Height          =   5067
         Left            =   0
         TabIndex        =   28
         Top             =   180
         Width           =   6885
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   420
            Left            =   2520
            MaxLength       =   20
            TabIndex        =   30
            Top             =   300
            Width           =   4020
         End
         Begin TrueOleDBGrid60.TDBGrid TDBGrid 
            Height          =   2940
            Left            =   150
            OleObjectBlob   =   "frmDebtors.frx":4C57
            TabIndex        =   29
            Top             =   720
            Width           =   6405
         End
      End
   End
   Begin CCubeX2.ContourCubeX CC 
      Height          =   2715
      Left            =   75
      TabIndex        =   1
      Top             =   5685
      Width           =   1905
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   16380654
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   0
      AllowInactiveDimArea=   0   'False
      AllowExpand     =   -1  'True
      AllowPivot      =   0   'False
      TotalsString    =   ""
      InactiveDimAreaBkColor=   16380654
      AutoSize        =   0   'False
      UnusedDataAreaColor=   16380654
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmDebtors.frx":8CCB
   End
   Begin VB.CommandButton cmdMatchPayments 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Match payments"
      Height          =   450
      Left            =   14220
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4140
      Width           =   870
   End
   Begin VB.TextBox txtArg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   420
      Left            =   2475
      MaxLength       =   20
      TabIndex        =   0
      Top             =   285
      Width           =   4020
   End
   Begin VB.Frame frBalances 
      BackColor       =   &H00F9F2EE&
      ForeColor       =   &H8000000D&
      Height          =   1200
      Left            =   6615
      TabIndex        =   4
      Top             =   4125
      Width           =   7485
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
         TabIndex        =   16
         Top             =   750
         Visible         =   0   'False
         Width           =   1110
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
         TabIndex        =   15
         Top             =   750
         Visible         =   0   'False
         Width           =   960
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   750
         Visible         =   0   'False
         Width           =   990
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
         TabIndex        =   11
         Top             =   750
         Visible         =   0   'False
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
         TabIndex        =   10
         Top             =   360
         Width           =   1110
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   360
         Width           =   990
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
         TabIndex        =   6
         Top             =   360
         Width           =   990
      End
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
         TabIndex        =   5
         Top             =   360
         Width           =   960
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
         TabIndex        =   24
         Top             =   780
         Visible         =   0   'False
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
         TabIndex        =   23
         Top             =   390
         Width           =   675
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   150
         Width           =   975
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
         TabIndex        =   17
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.CommandButton cbSince 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7E6D8&
      Caption         =   "Last week"
      Height          =   450
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   270
      Width           =   8460
   End
   Begin VBSmartXPMenu.SmartMenuXP SmartMenuXP 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Shadow          =   0   'False
   End
   Begin TrueOleDBGrid60.TDBGrid LedgerGrid 
      Height          =   3240
      Left            =   6615
      OleObjectBlob   =   "frmDebtors.frx":96D0
      TabIndex        =   3
      Top             =   780
      Width           =   8445
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   4035
      Left            =   105
      OleObjectBlob   =   "frmDebtors.frx":E327
      TabIndex        =   25
      Top             =   765
      Width           =   6405
   End
End
Attribute VB_Name = "frmDebtors"
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
Dim cCust As c_Customer

Dim dteDate1 As Date
Dim dteDate2 As Date
Dim cJNL As c_JNL
Dim dCN As d_JNL
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
Dim oCust As a_Customer
Dim Res As Boolean
Dim oSM As z_StockManager
Private Sub cbSince_Click()
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
End Sub

Private Sub CC_OnDimValueClick(ByVal AxisSection As CCubeX2.IAxisSection, ByVal Level As Long)
    Set oCust = Nothing
    Set oCust = New a_Customer
    Res = oCust.Load(0, AxisSection.getValue(Level))
    LoadTransactions AxisSection.getValue(Level)
    LoadStatement
End Sub

Private Sub cmdMatchPayments_Click()
Dim frm As New frmPaymentMatch

    frm.component oCust.ID, oCust.NameAndCode(50), Top, Left, Height, 12800
    frm.Show ' vbModal
    Screen.MousePointer = vbHourglass
    oCust.Reload
    LoadDebtorsStatement

End Sub
Private Sub LoadDebtorsStatement()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    oSQL.RunProc "[AgeInvoices]", Array(oCust.ID), ""
    oCust.Reload
    LoadTransactions
    LoadStatement
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadDebtorsStatement"
End Sub

Private Sub Form_Load()
    flgLoading = True
    Me.CC.Top = 5420
    Me.Grid.Top = 750
    Me.LedgerGrid.Top = 750
    Me.cbSince.Top = 270
    Me.Top = 200
    Me.Left = 50
    Me.Width = 6600
    
    Me.Height = 4000
    enSince = OptionLoop(enSince, 5)
    
    cbSince.Caption = TranslateSince(CInt(enSince))
    SetDateArgs
    pBuildMenus
 '   LoadBrowse
    SetGridLayout Me.Grid, Me.Name & Grid.Name
    SetFormSize Me
    Me.WindowState = vbNormal
    flgLoading = False
    
End Sub
Private Sub LoadBrowse()
    Find
    LoadArray
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.ReDim 1, cCust.Count, 1, 6
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Fullname2
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
            'XA.Value(lngIndex, 4) = objItem.Balance
            XA.Value(lngIndex, 5) = objItem.ID
            XA.Value(lngIndex, 6) = objItem.Blocked
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
    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oCust = Nothing
    Set oCust = New a_Customer
    oCust.Load FNN(XA(Grid.Bookmark, 5))
    LoadTransactions FNN(XA(Grid.Bookmark, 5))
'    If XA(Grid.Bookmark, 8) = "JNL" Then
''        Set frmCustJnl = New frmCustomerPreview
''        lngID = val(XA(Grid.Bookmark, 7))
''      '  frmCustJnl.component lngID, ""
''        frmCustJnl.Show
'    ElseIf XA(Grid.Bookmark, 8) = "REMIT" Then
'        Set rs = New ADODB.Recordset
'        oSQL.GetDynamicRecordset_Improved "Select * FROM vCRemittances WHERE TR_ID = " & CStr(val(XA.Value(Grid.Bookmark, 5))), enText, Array(), "", rs
'    '    SetFormSize Me
'        Preparecube
'
''        Set frmRemittancePreview = New frmCRemittancePreview
''        lngID = val(XA(Grid.Bookmark, 5))
''        frmRemittancePreview.component lngID, ""
''        frmRemittancePreview.Show
'    End If
    LoadStatement
    Screen.MousePointer = vbDefault

End Sub
Private Sub Preparecube()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    If rs Is Nothing Then Exit Sub
    rs.MoveFirst
    If rs.EOF Then
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.EOF Then
        
        CloseCube
        With CC.Cube
            .Dims.Add("Remittance Customer Name", "RemittanceCustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Remittance no", "RemittanceDocCode", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("CustomerName", "CustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Acno", "CustomerAcno", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("Payment no", "PaymentDocCode", , xda_vertical).MoveTo xda_vertical
            .BaseFacts.Add "Amount", "Amount"
            .Facts.Add "Amount", "Amount", xfaa_SUM
            .BaseFacts.Add "SettlementDiscount", "SettlementDiscount"
            .Facts.Add "SettlementDiscount", "SettlementDiscount", xfaa_SUM
            CC.Facts(0).Appearance.Format = "###,##0.00"
            CC.Facts(0).Caption = "Amount"
            CC.Facts(1).Appearance.Format = "###,##0.00"
            CC.Facts(1).Caption = "Sett.disc."
            CC.NoGrandTotals = False
            CC.Dims(1).NoTotals = True
            CC.Dims(2).NoTotals = True
            CC.Dims(4).NoTotals = True
            CC.TitleSettings.Visible = False
            CC.VAxis.DrillDownLevel = 4
            For Each Fact In CC.Facts
              Fact.Visible = True
            Next
            Set rs.ActiveConnection = Nothing
            .Open rs

        End With
        AfterOpen
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformance.Preparecube"
    HandleError
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
    CC.Visible = CC.Active
    CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.AfterOpen"
End Sub

Private Sub CheckVisible()
    On Error GoTo errHandler
    CC.Visible = True 'ContourCubeX.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.CheckVisible"
End Sub
Private Sub CloseCube()
    On Error GoTo errHandler
    With CC
      .Active = False
      .Cube.Dims.Clear
      .Cube.Facts.Clear
      .Cube.BaseFacts.Clear
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.CloseCube"
End Sub

Private Sub LoadStatement(Optional pAcno As String)
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double

    If oTRs.Count = 0 Then
        XB.Clear
        LedgerGrid.ReBind
        Exit Sub
    End If
    
    For i = 1 To LedgerGrid.Columns.Count
        LedgerGrid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "LedgerGrid", CStr(i), LedgerGrid.Columns(i - 1).Width)
    Next
    
    XB.Clear
    LedgerGrid.ReBind
   ' Set oSQL = New z_SQL
   ' oSM.RecalculateTPBalance oCust.ID
    Me.txtCurBal = oCust.BalanceCurF
    Me.txt30Bal = oCust.Balance30F
    Me.txt60Bal = oCust.Balance60F
    Me.txt90Bal = oCust.Balance90F
    Me.txt120PlusBal = oCust.Balance120F
    Me.txtBalance = oCust.BalanceF
    
    i = 1
    j = 1
    Do While i <= oTRs.Count
        If oTRs.Item(i).DocType <> "BF" Then
            XB.ReDim 1, j, 1, 8
            XB.Value(j, 1) = oTRs.Item(i).DOCCode
            XB.Value(j, 2) = oTRs.Item(i).DocType
            XB.Value(j, 3) = oTRs.Item(i).DocDateF
            XB.Value(j, 4) = oTRs.Item(i).DebitF
            XB.Value(j, 5) = oTRs.Item(i).CreditF
            XB.Value(j, 6) = oTRs.Item(i).MEMO
            XB.Value(j, 7) = oTRs.Item(i).DOCID
            XB.Value(j, 7) = oTRs.Item(i).DOCCaptureDate
 '           dblBal = dblBal + oTRs.Item(i).Debit
 '           dblBal = dblBal - oTRs.Item(i).Credit
            j = j + 1
        End If
        i = i + 1
    Loop
    XB.QuickSort 1, XB.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
    LedgerGrid.Array = XB
    LedgerGrid.ReBind
    LedgerGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"
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


Private Sub Find()
    On Error GoTo errHandler
Dim bNotFound As Boolean
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
Dim byear As Boolean
Dim yr As String
Dim mth As String
Dim strDate1 As String
Dim strDate2 As String
Dim lngCount As Long

    bNotFound = False
    If Left(txtArg, 3) = "yr=" Then byear = True
    
    If txtArg > " " And Not (byear) Then
        'Search for Reference
        Set cJNL = Nothing
        Set cJNL = New c_JNL
        cJNL.Load bNotFound, 0, "", txtArg, dteDate1, dteDate2
        If bNotFound Then
            'Search for customer by AcJNLO
            Set cJNL = Nothing
            Set cJNL = New c_JNL
            SetDateArgs
            cJNL.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
            If bNotFound Then
               Set frm = New frmBrowseCustomers2
               frm.component txtArg, lngCount
               If lngCount > 1 Then
                    frm.Show vbModal
                    lngTPID = frm.CustomerID
                    Unload frm
                ElseIf lngCount = 1 Then
                    lngTPID = frm.CustomerID
                    Unload frm
                End If
               If lngTPID > 0 Then
                    Set cJNL = Nothing
                    Set cJNL = New c_JNL
                    SetDateArgs
                    cJNL.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
               End If
            End If
        Else
            enSince = 1
            cbSince.Caption = TranslateSince(1)
        End If
    Else
        Set cJNL = Nothing
        Set cJNL = New c_JNL
        If byear Then
            yr = Mid(txtArg, 4, 4)
            mth = Mid(txtArg, 9, 2)
            If mth > "" Then
                strDate1 = yr & "-" & mth & "-01"
                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
            Else
                strDate1 = yr & "-01-01"
                strDate2 = yr & "-12-31"
            End If
            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
                SetDateArgs
            Else
                dteDate1 = CDate(strDate1)
                dteDate2 = CDate(strDate2)
            End If
        Else
            SetDateArgs
        End If
        cJNL.Load bNotFound, 0, "", "", dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    MousePointer = vbDefault
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.Find"
End Sub
Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    
    If KeyAscii = 13 Then  ' The ENTER key.
       HandleResults
        If cCust.Count > 1 Then
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
Private Sub pBuildMenus()
    
    With Me.SmartMenuXP.MenuItems
'        ' Root > File...
        .Add 0, "keyFile", , "&File"
        .Add "keyFile", "keyExit", , "E&xit", , vbAltMask, vbKeyQ
'
''Root>view
        .Add 0, "keyView", , "View"
        .Add "keyView", "KeyRemittances", , "Browse &remittances"
        .Add "keyView", "KeyJournals", , "Browse &journals"
        .Add "keyView", "KeyLedgerView", , "Browse customer &ledger"
        .Add "keyView", "KeyStatementView", , "&Statement"
    
''Root>Actions
        .Add 0, "keyActions", , "Actions"
        .Add "keyActions", "KeyCaptureRemittance", , "&Capture remittance"
        .Add "keyActions", "KeyMatch", , "&Matching"
        .Add "keyActions", "KeyProduceStatements", , "&Produce statements"
        .Add "keyActions", "KeyOpenCashBook", , "C&ash book"
        .Add "keyActions", "KeyTemplate", , "Cash book &template"
        

    End With
    
    SmartMenuXP.Font.Name = "Ms Sans Serif"
    SmartMenuXP.BackColor = &HF7EDE8
    SmartMenuXP.Font.Size = 9

End Sub

Private Sub SmartMenuXP_Click(ByVal ID As Long)
    With SmartMenuXP.MenuItems
        
        Select Case .key(ID)
            Case "keyExit"
                    Unload Me
            Case "KeyRemittances"
                Set frmRemittancesBrowse = New frmBrowseDBJNLs
                frmRemittancesBrowse.Show
            Case "KeyCaptureRemittance"
                Set frm = New frmBrowseCustomers2
                frm.Show vbModal
                lngTPID = frm.CustomerID
                strCustomerName = frm.CustomerName
                Unload frm
                If lngTPID = 0 Then Exit Sub
                Set frmCustomerRemittance = New frmCustomerRemittance
                frmCustomerRemittance.component lngTPID, strCustomerName, Me.hwnd, Date, "", 0
                frmCustomerRemittance.Show
                Set oTRs = Nothing
                Set oTRs = New c_DebtorsTransPerTP
                oTRs.Load lngTPID, CDate("2000-01-01")
                LedgerGrid.Caption = strCustomerName
            Case "KeyOpenCashBook"
                Set frmCashBook = New frmCashBook
                frmCashBook.Show
            Case "KeyTemplate"
                Set frmCashBookMaintenance = New frmCashBookMaintenance
                frmCashBookMaintenance.Show
        End Select
        
    End With
    
End Sub

Private Sub Form_Resize()
  '  CC.Height = NonNegative_Lng((Me.Height) - 6100)
  '  Grid.Height = NonNegative_Lng(Me.Height / 2)


  '  frBalances.top = NonNegative_Lng((Me.Height - 795) / 2)
    Me.LedgerGrid.Top = Me.Grid.Top
    Me.Grid.Height = NonNegative_Lng((Me.Height + frBalances.Height) / 2) - 1600
    Me.LedgerGrid.Height = NonNegative_Lng(Me.Grid.Height - 1200)
    frBalances.Top = NonNegative_Lng((LedgerGrid.Top + LedgerGrid.Height + 50))
    cmdMatchPayments.Top = frBalances.Top
    Me.CC.Top = Me.Height / 2
    Me.CC.Height = Me.Height / 2
'    Me.gAllocations.Height = Frame1.Height - 1000
'    GCredits.top = Frame1.Height + 1100
'    lblUnallocatedCredits.top = GCredits.top - 200
'    GCredits.Height = NonNegative_Lng((Me.Height - 3300) / 2)
'    Me.cmdNewPayment.top = Me.Height - 2000
'    Me.cmdJnls.top = Me.Height - 1000
'    cmdPost.top = Me.Height - 900

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.Grid, Me.Name & Grid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
End Sub

Private Function pGetPicture(sFilename As String) As StdPicture
    ' - This example uses LoadPicture() to load the menu images from disk
    ' - You can also use an ImageList object for this purpose...
    Set pGetPicture = LoadPicture("C:\Downloads\SmartMenu\Images\" + sFilename + ".ico")
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        LoadBrowse
    End If
End Sub


Private Sub HandleResults(Optional plngCount As Long)
    On Error GoTo errHandler
    If txtArg = "" Then Exit Sub
    Set cCust = Nothing
    Set cCust = New c_Customer
    Screen.MousePointer = vbHourglass
    
    cCust.LoadEasy Replace(txtArg, "'", "''"), False ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    plngCount = cCust.Count
    LoadArray
    Grid.ReBind
    Grid.Enabled = True
    If cCust.Count = 1 Then
        LoadTransactions FNN(XA(Grid.Bookmark, 5))
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.HandleResults(plngCount)", plngCount
End Sub

