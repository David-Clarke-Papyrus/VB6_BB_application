VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmBrowsePayments 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse customer payments"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15135
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowsePayments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin CCubeX2.ContourCubeX CC 
      Height          =   3345
      Left            =   8895.001
      TabIndex        =   9
      Top             =   1185
      Width           =   5850
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   13882315
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   ""
      InactiveDimAreaBkColor=   13882315
      AutoSize        =   0   'False
      UnusedDataAreaColor=   13882315
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmBrowsePayments.frx":058A
   End
   Begin VB.TextBox txtCustRemittanceCode 
      BackColor       =   &H00CDFAFA&
      Enabled         =   0   'False
      Height          =   345
      Left            =   9045.001
      MaxLength       =   30
      TabIndex        =   12
      Top             =   750
      Width           =   2070
   End
   Begin VB.TextBox txtDepositDate 
      BackColor       =   &H00CDFAFA&
      Enabled         =   0   'False
      Height          =   345
      Left            =   13170
      MaxLength       =   30
      TabIndex        =   11
      Top             =   750
      Width           =   1275
   End
   Begin VB.TextBox txtBatchTotal 
      BackColor       =   &H00CDFAFA&
      Enabled         =   0   'False
      Height          =   345
      Left            =   11475
      MaxLength       =   30
      TabIndex        =   10
      Top             =   750
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowsePayments.frx":0A23
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4620
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   7680
      Picture         =   "frmBrowsePayments.frx":0DAD
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   6810
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5055
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowsePayments.frx":1137
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1000
      End
      Begin VB.TextBox txtArg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   2490
         TabIndex        =   0
         Tag             =   "Enter product code,document number, Acc no.,or start of customer name followed by '*'. Hit ENTER to fetch."
         Top             =   240
         Width           =   2500
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   405
         Left            =   6180
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for . . ."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2760
         TabIndex        =   2
         Top             =   750
         Width           =   1755
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3705
      Left            =   90
      OleObjectBlob   =   "frmBrowsePayments.frx":14C1
      TabIndex        =   5
      Top             =   1185
      Width           =   8415.001
   End
   Begin VB.Label lblCustomerRemittance 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer remittance number"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   9090.001
      TabIndex        =   15
      Top             =   540
      Width           =   2010
   End
   Begin VB.Label lblDepositDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit date"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   13320
      TabIndex        =   14
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label lblBatchTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total deposited"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   11535
      TabIndex        =   13
      Top             =   540
      Width           =   1290
   End
End
Attribute VB_Name = "frmBrowsePayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Dim mcol As c_Payments
Dim tlCustomer As z_TextList
Dim lngTPID As Long
Dim strRef As String
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim XA As New XArrayDB
Dim xMLDoc As ujXML
Dim frmCustomer As frmCustomerPreview

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.Grid, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.SetMenu"
End Sub



Private Sub cbSince_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cbSince_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cbSince_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If KeyAscii = 13 Then
        Find
        LoadArray
        Grid.ReBind
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
 Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Screen.MousePointer = vbHourglass
    Find
    LoadCube
    LoadArray
    Grid.ReBind
    Grid.Bookmark = 1
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
   ' cmdFind1_Click
    If Grid.Enabled Then
        If XA.Count(1) > 0 Then
            mSetfocus Grid
        Else
            mSetfocus Me.txtArg
        End If
    Else
        mSetfocus Me.txtArg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    Grid.Width = NonNegative_Lng(Me.Width - 6800)
    lngDiff = Grid.Height
    Grid.Height = NonNegative_Lng(Me.Height - (Grid.top + 1220))
    lngDiff = (Grid.Height - lngDiff)
    cmdPrint.top = NonNegative_Lng(Me.Height - 1150)
    cmdClose.top = cmdPrint.top
    Me.lblCustomerRemittance.Left = NonNegative_Lng(Me.Width - 6300)
    Me.lblDepositDate.Left = NonNegative_Lng(Me.Width - 3900)
    Me.lblBatchTotal.Left = NonNegative_Lng(Me.Width - 2200)
    Me.txtCustRemittanceCode.Left = NonNegative_Lng(Me.Width - 6300)
    Me.txtDepositDate.Left = NonNegative_Lng(Me.Width - 3900)
    Me.txtBatchTotal.Left = NonNegative_Lng(Me.Width - 2200)
    Me.CC.Left = NonNegative_Lng(Me.Width - 6500)
    Me.CC.Height = NonNegative_Lng(Me.Height - 2390)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadCube()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    If rs Is Nothing Then Exit Sub
    rs.MoveFirst
    If rs.eof Then
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.eof Then
        
        CloseCube
        With CC.Cube
            .Dims.Add("CustomerName", "CustomerName", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("DepositCode", "DepositCode", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("RemittanceReference", "RemittanceReference", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("DepositDate", "DepositDate", , xda_vertical).MoveTo xda_vertical
            .BaseFacts.Add "DepositAmount", "DepositAmount"
            .Facts.Add "DepositAmount", "DepositAmount", xfaa_SUM
            .BaseFacts.Add "DepositSettlementAmount", "DepositSettlementAmount"
            .Facts.Add "DepositSettlementAmount", "DepositSettlementAmount", xfaa_SUM
            CC.Facts(0).Appearance.Format = "###,##0.00"
            CC.Facts(0).Caption = "Amount"
            CC.Facts(1).Appearance.Format = "###,##0.00"
            CC.Facts(1).Caption = "Sett.disc."
            CC.NoGrandTotals = False
            CC.TitleSettings.Text = "Customer payments summary"
            CC.VAxis.DrillDownLevel = 0
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
Private Sub CloseCube()
    On Error GoTo errHandler
 With CC
   .Active = False
   .Cube.Dims.Clear
   .Cube.Facts.Clear
   .Cube.BaseFacts.Clear
 End With
' CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CloseCube"
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
 CC.Visible = CC.Active
' CheckEnabled
 CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.AfterOpen"
End Sub
Private Sub CheckVisible()
    On Error GoTo errHandler
 CC.Visible = True 'ContourCubeX.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckVisible"
End Sub
Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If KeyAscii = 13 Then
        Find
        LoadArray
        Grid.ReBind
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

'Private Function ArgIsProductCode() As Boolean
'    On Error GoTo errHandler
'
'   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowsePayments.ArgIsProductCode"
'End Function
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
    ErrorIn "frmBrowsePayments.SetDateArgs"
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
        Set mcol = Nothing
        Set mcol = New c_Payments
            mcol.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2, , rs
            If bNotFound Then
               Set frm = New frmBrowseCustomers2
               frm.component txtArg, lngCount
               If lngCount > 1 Then
                    frm.Show vbModal
                    lngTPID = frm.CustomerID
  '                  Me.txtArg = frm.CustomerName
                    Unload frm
                ElseIf lngCount = 1 Then
                    lngTPID = frm.CustomerID
'                    Me.txtArg = frm.CustomerName
                    Unload frm
                End If
               If lngTPID > 0 Then
                   Set mcol = Nothing
                   Set mcol = New c_Payments
                   SetDateArgs
                   mcol.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2, , rs
               End If
        End If
    Else
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
        mcol.Load bNotFound, 0, "", "", dteDate1, dteDate2, , rs
    End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.lvw_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Set tlCustomer = New z_TextList
    Set mcol = New c_Payments
    If Me.WindowState <> 2 Then
        Me.top = 50
        Me.Left = 50
        Me.Width = 7300
        Me.Height = 6100
    End If
    SetMenu
    LoadControls
    SetFormSize Me
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set tlCustomer = Nothing
    Set mcol = Nothing
    
    SaveLayout Me.Grid, Me.Name & "Grid"
    SaveFormSize Me.Name, Me.Height, Me.Width
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = ""
    strDate1 = ""
    strDate2 = ""
    lngTPID = 0
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False
    SetGridLayout Me.Grid, Me.Name & "Grid"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, mcol.Count, 1, 8
    For i = 1 To mcol.Count
            XA.Value(i, 1) = mcol(i).DepositorName & (IIf(Len(Trim(mcol(i).DepositorAcNo)) <= 1, "", "(" & Trim(mcol(i).DepositorAcNo) & ")"))
            XA.Value(i, 2) = mcol(i).DepositCode & mcol(i).StaffNameB
            XA.Value(i, 3) = mcol(i).DepositDateF
            XA.Value(i, 4) = mcol(i).DepositAmountF
            XA.Value(i, 5) = mcol(i).DepositSettlementDiscountF
            XA.Value(i, 6) = mcol(i).DepositID & "K"
            XA.Value(i, 7) = mcol(i).statusF
            XA.Value(i, 8) = mcol(i).DepositorID
    Next
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If flgLoading Then Exit Sub
    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frmCustomer = New frmCustomerPreview
    lngID = val(XA(Grid.Bookmark, 8))
    frmCustomer.Component2 lngID    ', False
    frmCustomer.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Grid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 7) = "VOID" Or XA(Bookmark, 6) = "CANCELLED" Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    End If
    If XA(Bookmark, 7) = "IN PROCESS" Then
        RowStyle.BackColor = &H80FF80
    End If
    If XA(Bookmark, 7) = "COMPLETE" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
 '   If ColIndex = 2 Then ColIndex = 4
    If ColIndex = 2 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 4, Direction, GetRowType(4) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2
            GetRowType = XTYPE_STRING
        Case 3, 4
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    If flgLoading Then Exit Sub
    str = "Notes" & vbCrLf _
            & "Enter document number, Acc no.,or start of customer name followed by '*'." & vbCrLf _
            & "Hit ENTER to fetch. " & vbCrLf & vbCrLf _
            & "Search for old data like this . . . " & vbCrLf _
            & "yr=2002     fetches all records for 2002" & vbCrLf & vbCrLf _
            & "yr=2002-03     fetches all records for March 2002" & vbCrLf & vbCrLf _
            & "Maximum records returned is settable  (ask support person)" & vbCrLf _
            & "This is currently set at " & oPC.MaxBrowseRecs & " records" & vbCrLf
    MsgBox str, vbInformation, "Help"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.Label3_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Public Function ExportToXML() As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strFilename As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument30
Dim opXMLDOC As New MSXML2.DOMDocument30
Dim objXMLDOC  As New MSXML2.DOMDocument30
Dim strExecutable As String

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "CO_1"
        .chCreate "CO"
            .elText = "Customer orders at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        For i = 1 To mcol.Count
            
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = mcol(i).DepositorName & (IIf(Len(Trim(mcol(i).DepositorAcNo)) <= 1, "", "(" & Trim(mcol(i).DepositorAcNo) & ")"))
            .elCreateSibling "Col_2"
                .elText = mcol(i).DepositCode & mcol(i).StaffNameB
            .elCreateSibling "Col_3"
                .elText = mcol(i).DepositDateF
            .elCreateSibling "Col_4"
                .elText = mcol(i).statusF
                .navUP
        Next i
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\COs" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .HTML FILE
    objXSL.async = False
    objXSL.validateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\CO_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

    strFilename = oPC.LocalFolder & "\CO.RTF"
    If fs.FileExists(strFilename) Then
        fs.DeleteFile strFilename, True
    End If
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile
    
    strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
    Shell strExecutable, vbNormalFocus
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePayments.ExportToXML"
End Function

