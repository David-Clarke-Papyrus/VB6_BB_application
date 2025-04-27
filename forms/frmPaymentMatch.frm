VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmPaymentMatch 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Payment matching"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   ControlBox      =   0   'False
   Icon            =   "frmPaymentMatch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   12645
   Begin VB.Frame Fram2 
      BackColor       =   &H00F7EDE8&
      Height          =   675
      Left            =   4305
      TabIndex        =   16
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdSince 
         BackColor       =   &H00E7E6D8&
         Caption         =   "&Fetch"
         Height          =   465
         Left            =   4755
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   150
         Width           =   930
      End
      Begin VB.CheckBox chkOS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F7EDE8&
         Caption         =   "Outstanding debits only"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2505
         TabIndex        =   17
         Top             =   225
         Value           =   1  'Checked
         Width           =   1965
      End
      Begin MSComCtl2.DTPicker dtSince 
         Height          =   300
         Left            =   720
         TabIndex        =   18
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
         _Version        =   393216
         Format          =   221839361
         CurrentDate     =   39882
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "since "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   105
         TabIndex        =   19
         Top             =   255
         Width           =   555
      End
   End
   Begin VB.ComboBox cboRelatedDebtors 
      ForeColor       =   &H00915A48&
      Height          =   315
      Left            =   705
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   90
      Width           =   3540
   End
   Begin VB.CommandButton cmdAllocate 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Auto allocate"
      Height          =   600
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5130
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.CommandButton cmdJnls 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&New journal"
      Height          =   480
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1905
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11295
      Picture         =   "frmPaymentMatch.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   45
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F7EDE8&
      Caption         =   "Debits and allocated credits"
      ForeColor       =   &H8000000D&
      Height          =   2625
      Left            =   60
      TabIndex        =   5
      Top             =   795
      Width           =   12330
      Begin VB.CommandButton cmdURemovePosting 
         BackColor       =   &H00E7E6D8&
         Caption         =   "&Remove allocation"
         Height          =   390
         Left            =   10575
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1545
      End
      Begin TrueOleDBGrid60.TDBGrid gDebits 
         Height          =   1995
         Left            =   105
         OleObjectBlob   =   "frmPaymentMatch.frx":0714
         TabIndex        =   12
         Top             =   450
         Width           =   5940
      End
      Begin TrueOleDBGrid60.TDBGrid gAllocations 
         Height          =   1635
         Left            =   6150
         OleObjectBlob   =   "frmPaymentMatch.frx":4E04
         TabIndex        =   13
         Top             =   465
         Width           =   6015
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Matching allocated credits"
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
         Left            =   6030
         TabIndex        =   8
         Top             =   225
         Width           =   2910
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Debits"
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
         Left            =   285
         TabIndex        =   7
         Top             =   225
         Width           =   2985
      End
   End
   Begin VB.CommandButton cmdNewPayment 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&New payment"
      Height          =   480
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4260
      Width           =   1905
   End
   Begin VB.CommandButton cmdPost 
      BackColor       =   &H00E7E6D8&
      Caption         =   "&Allocate to selected debit"
      Height          =   390
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5010
      Width           =   1905
   End
   Begin TrueOleDBGrid60.TDBGrid GCredits 
      Height          =   1320
      Left            =   75
      OleObjectBlob   =   "frmPaymentMatch.frx":94F5
      TabIndex        =   0
      Top             =   3645
      Width           =   7380
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Debtor"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   150
      Width           =   510
   End
   Begin VB.Label lblUnallocatedCredits 
      BackStyle       =   0  'Transparent
      Caption         =   "Unallocated credits"
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
      Left            =   300
      TabIndex        =   3
      Top             =   3375
      Width           =   2370
   End
   Begin VB.Label lblInvoiceCode 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7185
      TabIndex        =   1
      Top             =   4575
      Width           =   1305
   End
End
Attribute VB_Name = "frmPaymentMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDebits As New ADODB.Recordset
Dim rsUnallocatedCredits As New ADODB.Recordset
Dim rsAllocations As New ADODB.Recordset
Dim XDebits As New XArrayDB
Dim XCredits As New XArrayDB
Dim XAllocations As New XArrayDB
Dim mTRID As Long
Dim PAYID As Long
Dim AmtToPost As Double
Dim TPID As Long
Dim OpenResult As Integer
Dim strName As String
Dim tlChildCustomers As z_TextList
Dim lngTop As Long
Dim lngLeft As Long
Dim strCustomerName As String


Private Sub LoadRelatedDebtors()
Dim oSQL As New z_SQL
Dim rs As ADODB.Recordset
Dim Res As Long

    LoadCombo Me.cboRelatedDebtors, tlChildCustomers
    cboRelatedDebtors.ListIndex = 0

End Sub
Private Sub LoadAllData()
    On Error GoTo errHandler
    LoadDebits 0
    LoadCredits
    mTRID = 0
    If XDebits.UpperBound(1) > 0 Then
        mTRID = XDebits(1, 6)
    End If
    LoadAllocationsRS

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadAllData"
End Sub

Private Sub cboRelatedDebtors_Change()
    TPID = tlChildCustomers.Key(cboRelatedDebtors)
End Sub

Private Sub cboRelatedDebtors_Click()
    TPID = tlChildCustomers.Key(cboRelatedDebtors)
    LoadAllData
End Sub

Private Sub cmdAllocate_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    Screen.MousePointer = vbHourglass
    oSQL.RunProc "MatchPaymentsAuto", Array(TPID), ""
    LoadDebits IIf(Me.chkOS = 1, 0, Me.dtSince)
    LoadCredits
    If XDebits.UpperBound(1) > 0 Then
        mTRID = XDebits(gDebits.Bookmark, 6)
    End If
    LoadAllocationsRS
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdAllocate_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdNewPayment_Click()
'    On Error GoTo errHandler
'Dim frm As frmCustPmt
'Dim oSQL As New z_SQL
'    Set frm = New frmCustPmt
'    frm.component TPID, strName
'    frm.Show vbModal
'  '  oSQL.RunProc "MatchPaymentsAuto", Array(TPID), ""
'  '  LoadAllData
'    LoadCredits
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPaymentMatch.cmdNewPayment_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub LoadDebits(pSince As Date)
    On Error GoTo errHandler
    If rsDebits.State = 1 Then rsDebits.Close
    rsDebits.CursorLocation = adUseClient
    If chkOS = 1 Then
        rsDebits.open "SELECT * FROM dsDebitsOS_perParent WHERE  TPID_c = " & TPID & " and IsOpen = 1 ORDER BY DTE DESC", oPC.COShort, adOpenDynamic
    Else
        rsDebits.open "SELECT * FROM dsDebitsOS_perParent WHERE  TPID_c = " & TPID & " and  Dte >= '" & ReverseDate(Me.dtSince) & "' ORDER BY DTE DESC", oPC.COShort, adOpenDynamic
    End If
    LoadXDebits
    Set gDebits.Array = XDebits
    gDebits.ReBind
    gDebits.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadDebits(pSince)", pSince
End Sub
Private Sub LoadCredits()
    On Error GoTo errHandler
    If rsUnallocatedCredits.State = 1 Then rsUnallocatedCredits.Close
    rsUnallocatedCredits.CursorLocation = adUseClient
    rsUnallocatedCredits.open "SELECT * FROM dsCreditsOS_perParent WHERE Credit - PostedAmount > 0  and TPID = " & TPID & " ORDER BY DTE", oPC.COShort, adOpenDynamic
    LoadXCredits
    Set GCredits.Array = XCredits
    GCredits.ReBind
    GCredits.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadCredits"
End Sub
Private Sub LoadAllocationsRS()
    On Error GoTo errHandler
    If rsAllocations.State = 1 Then rsAllocations.Close
    rsAllocations.CursorLocation = adUseClient
    rsAllocations.open "SELECT * FROM vMatchedCredits WHERE TargetTRID = " & mTRID & " ORDER BY DTE", oPC.COShort, adOpenDynamic
    'vMatchedPayments
    LoadXAllocations
    Set gAllocations.Array = XAllocations
    gAllocations.ReBind
    gAllocations.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadAllocationsRS"
End Sub
Private Sub LoadXDebits()
    On Error GoTo errHandler
Dim i As Integer

    XDebits.ReDim 1, rsDebits.RecordCount, 1, 6
    i = 0
    Do While Not rsDebits.eof
        i = i + 1
        XDebits(i, 1) = FNS(rsDebits.fields("DocCode"))
        XDebits(i, 2) = FNS(rsDebits.fields("dte"))
        XDebits(i, 3) = FNDBL(rsDebits.fields("Debit"))
        XDebits(i, 4) = FNDBL(rsDebits.fields("AmtPaid"))
        XDebits(i, 5) = FNDBL(rsDebits.fields("Debit")) - FNDBL(rsDebits.fields("AmtPaid")) '- FNDBL(rsDebits.Fields("SettDisc"))
        XDebits(i, 6) = FNN(rsDebits.fields("TRID"))
        
        rsDebits.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadXDebits"
End Sub
Private Sub LoadXCredits()
    On Error GoTo errHandler
Dim i As Integer

    XCredits.ReDim 1, rsUnallocatedCredits.RecordCount, 1, 7
    i = 0
    Do While Not rsUnallocatedCredits.eof
        i = i + 1
        XCredits(i, 1) = FNS(rsUnallocatedCredits.fields("DocCode"))
        XCredits(i, 2) = FNS(rsUnallocatedCredits.fields("dte"))
        XCredits(i, 3) = FNDBL(rsUnallocatedCredits.fields("Credit")) - FNDBL(rsUnallocatedCredits.fields("SettDisc"))
        XCredits(i, 4) = FNDBL(rsUnallocatedCredits.fields("SettDisc"))
        XCredits(i, 5) = FNDBL(rsUnallocatedCredits.fields("PostedAmount"))
        XCredits(i, 6) = FNDBL(rsUnallocatedCredits.fields("Credit")) - FNDBL(rsUnallocatedCredits.fields("PostedAmount"))
        XCredits(i, 7) = FNN(rsUnallocatedCredits.fields("TRID"))
        rsUnallocatedCredits.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadXCredits"
End Sub
Private Sub LoadXAllocations()
    On Error GoTo errHandler
Dim i As Integer

    XAllocations.ReDim 1, rsAllocations.RecordCount, 1, 7
    i = 0
    Do While Not rsAllocations.eof
        i = i + 1
        XAllocations(i, 1) = FNS(rsAllocations.fields("DocCode"))
        XAllocations(i, 2) = FNS(rsAllocations.fields("dte"))
        XAllocations(i, 3) = FNDBL(rsAllocations.fields("Credit"))
        XAllocations(i, 4) = FNS(rsAllocations.fields("CreditType"))
        XAllocations(i, 5) = FNDBL(rsAllocations.fields("PostedAmount"))
        XAllocations(i, 6) = FNN(rsAllocations.fields("TRID"))
        XAllocations(i, 7) = FNN(rsAllocations.fields("PA_ID"))
        rsAllocations.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.LoadXAllocations"
End Sub

Private Sub cmdOSOnly_Click()
    On Error GoTo errHandler
    dtSince = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdOSOnly_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdNewPayment_Click()
Dim frm1 As New frmCustPmt
Dim oSQL As New z_SQL

    Set frm1 = New frmCustPmt
    frm1.component TPID, strName
        frm1.Show vbModal
   ' oSQL.RunProc "MatchPaymentsAuto", Array(TPID), ""
    LoadAllData

End Sub

Private Sub cmdPost_Click()
    On Error GoTo errHandler
Dim oMatch As New a_PaymentMatches

    If IsNull(GCredits.Bookmark) Then Exit Sub
   oMatch.PostDebtorsPaymentOpenItem XDebits(gDebits.Bookmark, 6), XCredits(GCredits.Bookmark, 7)
    LoadDebits IIf(Me.chkOS = 1, 0, Me.dtSince)
    LoadCredits
    If XDebits.UpperBound(1) > 0 Then
        mTRID = XDebits(gDebits.Bookmark, 6)
    End If
    LoadAllocationsRS
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdPost_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSince_Click()
    On Error GoTo errHandler
    LoadDebits IIf(Me.chkOS = 1, 0, Me.dtSince)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdSince_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdURemovePosting_Click()
    On Error GoTo errHandler
Dim lngPaymentAllocationID As Long
Dim oMatch As New a_PaymentMatches

    If IsNull(gAllocations.Bookmark) Then Exit Sub
    lngPaymentAllocationID = XAllocations(gAllocations.Bookmark, 7)
    oMatch.DeletePosting lngPaymentAllocationID
    LoadDebits IIf(Me.chkOS = 1, 0, Me.dtSince)
    LoadCredits
    mTRID = XDebits(1, 6)
    LoadAllocationsRS
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdURemovePosting_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(pTPID As Long, pName As String, Optional plngTop As Long, Optional plngLeft As Long)
    On Error GoTo errHandler
    
    If plngTop > 0 Then lngTop = plngTop
    If plngLeft > 0 Then lngLeft = plngLeft
    strCustomerName = pName
    strName = pName
    TPID = pTPID
    Set tlChildCustomers = New z_TextList
    tlChildCustomers.Load ltChildCustomers, CStr(TPID)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.component(pTPID,pName)", Array(pTPID, pName)
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    DoEvents
    Caption = "Match payments for customer: " & strCustomerName
    dtSince = DateAdd("m", -2, Date)
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    LoadRelatedDebtors
    LoadAllData
    
    Me.gDebits.Enabled = True
    Me.gDebits.HeadBackColor = RGB(40, 40, 40)
    gDebits.Refresh
    SetFormSize Me
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    Frame1.Height = NonNegative_Lng((Me.Height - 795) / 2)
    Me.gDebits.TOP = Me.gAllocations.TOP
    Me.gDebits.Height = NonNegative_Lng(Frame1.Height - 600)
    Me.cmdURemovePosting.TOP = NonNegative_Lng(Frame1.Height - 450)
    Me.gAllocations.Height = NonNegative_Lng(Frame1.Height - 1000)
    GCredits.TOP = Frame1.Height + 1100
    lblUnallocatedCredits.TOP = NonNegative_Lng(GCredits.TOP - 200)
    GCredits.Height = NonNegative_Lng((Me.Height - 3300) / 2)
    Me.cmdNewPayment.TOP = NonNegative_Lng(Me.Height - 2000)
    Me.cmdJnls.TOP = NonNegative_Lng(Me.Height - 1000)
    cmdPost.TOP = NonNegative_Lng(Me.Height - 900)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    SaveFormSize Me.Name, Me.Height, Me.Width

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub gDebits_Click()
'    mTRID = XDebits(gDebits.Bookmark, 6)
'    LoadAllocationsRS

End Sub
Private Sub GDebits_SelChange(Cancel As Integer)
    On Error GoTo errHandler
    mTRID = XDebits(gDebits.Bookmark, 6)
    LoadAllocationsRS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.GDebits_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub GCredits_SelChange(Cancel As Integer)
    On Error GoTo errHandler
    PAYID = XCredits(GCredits.Bookmark, 6)
    LoadAllocationsRS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.GCredits_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub optOS_Click()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.optOS_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkOS_Click()
    On Error GoTo errHandler
    If chkOS = 1 Then
        LoadDebits 0
    Else
        LoadDebits dtSince
    End If
    LoadCredits
    mTRID = 0
    If XDebits.UpperBound(1) > 0 Then
        mTRID = XDebits(1, 6)
    End If
    LoadAllocationsRS


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.chkOS_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdJnls_Click()
    On Error GoTo errHandler
Dim frm1 As New frmCustJnl
Dim oSQL As New z_SQL

    Set frm1 = New frmCustJnl
    frm1.component TPID, strName
        frm1.Show vbModal
    oSQL.RunProc "MatchPaymentsAuto", Array(TPID), ""
    LoadAllData
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.cmdJnls_Click", , EA_NORERAISE
    HandleError
End Sub

