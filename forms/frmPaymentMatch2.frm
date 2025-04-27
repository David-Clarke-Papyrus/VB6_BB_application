VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B4B5B73C-172E-47B1-BFC2-C6F740957D01}#1.0#0"; "VB Control Manager.ocx"
Begin VB.Form frmPaymentMatch 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Payment and journal posting and matching"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21015
   ControlBox      =   0   'False
   Icon            =   "frmPaymentMatch2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   21015
   Begin VBControlManager.ControlManager CM 
      Height          =   6300
      Left            =   1635
      TabIndex        =   0
      Top             =   75
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   11113
      BackColor       =   11565924
      TitleBar_CloseVisible=   0   'False
      TitleBar_Height =   0
      TitleBar_Visible=   0   'False
      Begin VB.Frame fr2b 
         Caption         =   "Matching allocated credits"
         ForeColor       =   &H8000000D&
         Height          =   2400
         Left            =   6375
         TabIndex        =   16
         Top             =   1395
         Width           =   6195
         Begin VB.CommandButton cmdURemovePosting 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&Remove allocation"
            Height          =   390
            Left            =   4515
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1935
            Width           =   1545
         End
         Begin TrueOleDBGrid60.TDBGrid gAllocations 
            Height          =   1635
            Left            =   90
            OleObjectBlob   =   "frmPaymentMatch2.frx":038A
            TabIndex        =   18
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Frame fr1b 
         Caption         =   "Debits"
         ForeColor       =   &H8000000D&
         Height          =   2325
         Left            =   165
         TabIndex        =   12
         Top             =   1395
         Width           =   6030
         Begin VB.CommandButton cmdNewPayment 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&New payment"
            Height          =   300
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1845
            Width           =   1905
         End
         Begin VB.CommandButton cmdJnls 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&New journal"
            Height          =   345
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1815
            Width           =   1905
         End
         Begin TrueOleDBGrid60.TDBGrid gDebits 
            Height          =   1545
            Left            =   60
            OleObjectBlob   =   "frmPaymentMatch2.frx":4A0B
            TabIndex        =   15
            Top             =   270
            Width           =   5940
         End
      End
      Begin VB.Frame Fram2 
         BackColor       =   &H00F7EDE8&
         Height          =   930
         Left            =   0
         TabIndex        =   4
         Top             =   180
         Width           =   12570
         Begin VB.ComboBox cboRelatedDebtors 
            ForeColor       =   &H00915A48&
            Height          =   315
            Left            =   6585
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   390
            Width           =   3540
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
            Left            =   11430
            Picture         =   "frmPaymentMatch2.frx":908B
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   195
            Width           =   1000
         End
         Begin VB.CheckBox chkOS 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00F7EDE8&
            Caption         =   "Outstanding debits only"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2505
            TabIndex        =   6
            Top             =   225
            Value           =   1  'Checked
            Width           =   1965
         End
         Begin VB.CommandButton cmdSince 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&Fetch"
            Height          =   465
            Left            =   4755
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   150
            Width           =   930
         End
         Begin MSComCtl2.DTPicker dtSince 
            Height          =   300
            Left            =   720
            TabIndex        =   7
            Top             =   210
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            Format          =   58851329
            CurrentDate     =   39882
         End
         Begin VB.Label Label1 
            BackColor       =   &H00D3D3CB&
            BackStyle       =   0  'Transparent
            Caption         =   "Debtor"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   6000
            TabIndex        =   11
            Top             =   450
            Width           =   510
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00D3D3CB&
            BackStyle       =   0  'Transparent
            Caption         =   "Since "
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   105
            TabIndex        =   8
            Top             =   255
            Width           =   555
         End
      End
      Begin VB.Frame fr4 
         Caption         =   "Unallocated credits"
         ForeColor       =   &H8000000D&
         Height          =   2010
         Left            =   -30
         TabIndex        =   1
         Top             =   4290
         Width           =   12570
         Begin VB.CommandButton cmdPost 
            BackColor       =   &H00E7E6D8&
            Caption         =   "&Allocate to selected debit"
            Height          =   390
            Left            =   6090
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1695
            Width           =   1905
         End
         Begin TrueOleDBGrid60.TDBGrid GCredits 
            Height          =   1320
            Left            =   600
            OleObjectBlob   =   "frmPaymentMatch2.frx":9415
            TabIndex        =   3
            Top             =   330
            Width           =   7380
         End
      End
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
    TPID = tlChildCustomers.key(cboRelatedDebtors)
End Sub

Private Sub cboRelatedDebtors_Click()
    TPID = tlChildCustomers.key(cboRelatedDebtors)
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


Private Sub CM_SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Resize
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
    If pSince > CDate(0) Then
        rsDebits.Open "SELECT * FROM vDebtorsDebitsAll WHERE TPID = " & CStr(TPID) & " and Dte > '" & ReverseDate(pSince) & "' ORDER BY DTE", oPC.COShort, adOpenDynamic
    Else
        rsDebits.Open "SELECT * FROM vDebtorsDebitsOS WHERE TPID = " & TPID & " ORDER BY DTE", oPC.COShort, adOpenDynamic
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
    rsUnallocatedCredits.Open "SELECT * FROM vDebtorsCreditsOS WHERE Credit - PostedAmount > 0  and TPID = " & TPID & " ORDER BY DTE", oPC.COShort, adOpenDynamic
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
    rsAllocations.Open "SELECT * FROM vMatchedCredits WHERE TargetTRID = " & mTRID & " ORDER BY DTE", oPC.COShort, adOpenDynamic
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
    Do While Not rsDebits.EOF
        i = i + 1
        XDebits(i, 1) = FNS(rsDebits.Fields("DocCode"))
        XDebits(i, 2) = FNS(rsDebits.Fields("dte"))
        XDebits(i, 3) = FNDBL(rsDebits.Fields("Debit"))
        XDebits(i, 4) = FNDBL(rsDebits.Fields("AmtPaid"))
        XDebits(i, 5) = FNDBL(rsDebits.Fields("Debit")) - FNDBL(rsDebits.Fields("AmtPaid")) '- FNDBL(rsDebits.Fields("SettDisc"))
        XDebits(i, 6) = FNN(rsDebits.Fields("TRID"))
        
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
    Do While Not rsUnallocatedCredits.EOF
        i = i + 1
        XCredits(i, 1) = FNS(rsUnallocatedCredits.Fields("DocCode"))
        XCredits(i, 2) = FNS(rsUnallocatedCredits.Fields("dte"))
        XCredits(i, 3) = FNDBL(rsUnallocatedCredits.Fields("Credit")) - FNDBL(rsUnallocatedCredits.Fields("SettDisc"))
        XCredits(i, 4) = FNDBL(rsUnallocatedCredits.Fields("SettDisc"))
        XCredits(i, 5) = FNDBL(rsUnallocatedCredits.Fields("PostedAmount"))
        XCredits(i, 6) = FNDBL(rsUnallocatedCredits.Fields("Credit")) - FNDBL(rsUnallocatedCredits.Fields("PostedAmount"))
        XCredits(i, 7) = FNN(rsUnallocatedCredits.Fields("TRID"))
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
    Do While Not rsAllocations.EOF
        i = i + 1
        XAllocations(i, 1) = FNS(rsAllocations.Fields("DocCode"))
        XAllocations(i, 2) = FNS(rsAllocations.Fields("dte"))
        XAllocations(i, 3) = FNDBL(rsAllocations.Fields("Credit"))
        XAllocations(i, 4) = FNS(rsAllocations.Fields("CreditType"))
        XAllocations(i, 5) = FNDBL(rsAllocations.Fields("PostedAmount"))
        XAllocations(i, 6) = FNN(rsAllocations.Fields("TRID"))
        XAllocations(i, 7) = FNN(rsAllocations.Fields("PA_ID"))
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
    CM.FillContainer = True
    Screen.MousePointer = vbDefault
    SetCM Me, CM
    SetGridLayout Me.GCredits, Me.Name & GCredits.Name
    SetGridLayout Me.gDebits, Me.Name & gDebits.Name
    SetGridLayout Me.gAllocations, Me.Name & gAllocations.Name

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPaymentMatch.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    Resize
End Sub
Private Sub Resize()
    Me.gDebits.Left = 100
    Me.gDebits.Width = NonNegative_Lng(fr1b.Width - 200)
    Me.gDebits.Height = NonNegative_Lng(fr1b.Height - 800)
    Me.cmdNewPayment.Top = gDebits.Top + gDebits.Height
    Me.cmdNewPayment.Left = NonNegative_Lng(Me.gDebits.Width - 3800)
    Me.cmdJnls.Top = gDebits.Top + gDebits.Height
    Me.cmdJnls.Left = NonNegative_Lng(Me.gDebits.Width - 1800)
    Me.gAllocations.Left = 100
    Me.gAllocations.Width = NonNegative_Lng(fr2b.Width - 200)
    Me.gAllocations.Height = NonNegative_Lng(fr2b.Height - 800)
    Me.cmdURemovePosting.Top = gAllocations.Top + gAllocations.Height
    Me.cmdURemovePosting.Left = NonNegative_Lng(Me.gAllocations.Width - 1700)
    
    Me.GCredits.Left = 100
    Me.GCredits.Width = NonNegative_Lng(fr4.Width - 200)
    Me.GCredits.Height = NonNegative_Lng(fr4.Height - 800)
    Me.cmdPost.Top = GCredits.Top + GCredits.Height
    Me.cmdPost.Left = NonNegative_Lng(Me.GCredits.Width - 2100)
'    fr1b.Height = NonNegative_Lng((Me.Height - 795) / 2)
'    Me.gDebits.Top = Me.gAllocations.Top
'    Me.gDebits.Height = NonNegative_Lng(Frame1.Height - 600)
'    Me.cmdURemovePosting.Top = NonNegative_Lng(Frame1.Height - 450)
'    Me.gAllocations.Height = NonNegative_Lng(Frame1.Height - 1000)
'    GCredits.Top = Frame1.Height + 1100
'    GCredits.Height = NonNegative_Lng((Me.Height - 3300) / 2)
'    Me.cmdNewPayment.Top = NonNegative_Lng(Me.Height - 2000)
'    Me.cmdJnls.Top = NonNegative_Lng(Me.Height - 1000)
'    cmdPost.Top = NonNegative_Lng(Me.Height - 900)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    SaveLayout Me.gDebits, Me.Name & gDebits.Name
    SaveLayout Me.gAllocations, Me.Name & gAllocations.Name
    SaveLayout Me.GCredits, Me.Name & GCredits.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
    SaveSplits Me.Name, Me.CM
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

