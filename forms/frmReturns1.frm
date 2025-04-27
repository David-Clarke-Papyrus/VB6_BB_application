VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmReturn1 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Returns"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   6585
   Begin VB.Frame Fr 
      BackColor       =   &H00D3D3CB&
      Height          =   1845
      Left            =   60
      TabIndex        =   5
      Top             =   5370
      Width           =   6015
      Begin VB.CommandButton cmdSupp 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Generate &supplementary returns document"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1830
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmReturns1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   4005
      End
      Begin VB.CheckBox chkAllDels 
         BackColor       =   &H00D3D3CB&
         Caption         =   "All deliveries"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   345
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   13882315
         CustomFormat    =   "MM/yyy"
         Format          =   221839363
         CurrentDate     =   39090
      End
      Begin VB.CommandButton cmdGenerate 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Generate Returns document"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3015
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmReturns1.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   2805
      End
      Begin VB.Label lblOr 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "or"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   525
         TabIndex        =   13
         Top             =   765
         Width           =   270
      End
      Begin VB.Label Label2 
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
         Left            =   1170
         TabIndex        =   9
         Top             =   270
         Width           =   360
      End
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
      Height          =   5325
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   6000
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   270
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   4680
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmReturns1.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   270
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
         Left            =   2610
         TabIndex        =   2
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   270
         Width           =   1785
      End
      Begin TrueOleDBGrid60.TDBGrid Grid 
         Height          =   3945
         Left            =   150
         OleObjectBlob   =   "frmReturns1.frx":0A9E
         TabIndex        =   4
         Top             =   1200
         Width           =   5685
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for product code. . ."
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2580
         TabIndex        =   3
         Top             =   750
         Width           =   1980
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   240
      Left            =   1560
      TabIndex        =   0
      Top             =   7590
      Visible         =   0   'False
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmReturn1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cR As c_R
Dim dR As d_R
Dim tlSupplier As z_TextList
Dim lngTPID As Long
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim XA As New XArrayDB
Dim oSupp As a_Supplier
Dim WithEvents oSM As z_StockManager
Attribute oSM.VB_VarHelpID = -1


Public Sub component(Optional pTP As a_Supplier)
    On Error GoTo errHandler
Dim strCaption As String

    Set oSupp = pTP
    If Not oSupp Is Nothing Then
        strCaption = "Returns for " & oSupp.NameAndCode(24)
    Else
        strCaption = "Returns for all suppliers"
    End If
    Caption = strCaption
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.component(pTP)", pTP
End Sub

Private Sub cbSince_Click()
    On Error GoTo errHandler
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
    txtArg = ""
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cbSince_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbSince_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        Find
        LoadArray
        Grid.ReBind
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkAllDels_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Me.DT1.Enabled = Not Me.chkAllDels.Enabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.chkAllDels_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    Find
    LoadArray
    Grid.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Label2_Click()
    On Error GoTo errHandler
Dim str As String
    str = "You can search through all items received sale-or-return in the window period specified on the supplier record," _
    & " or you can look for items delivered on an invoice dated in a selected month." & vbCrLf
    MsgBox str, vbInformation, "Help"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Label2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oSM_Emax(pMax As Long)
    On Error GoTo errHandler
    PB1.Max = pMax
    PB1.Min = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.oSM_Emax(pMax)", pMax, EA_NORERAISE
    HandleError
End Sub

Private Sub oSM_eProgress(iProg As Long)
    On Error GoTo errHandler
    PB1.Value = iProg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.oSM_eProgress(iProg)", iProg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdGenerate_Click()
    On Error GoTo errHandler
Dim iresult As Integer


    Screen.MousePointer = vbHourglass
    Set oSM = New z_StockManager
    PB1.Visible = True
    If oSupp Is Nothing Then
        iresult = oSM.GenerateReturnsPerPub(0, IIf(chkAllDels = 0, Me.DT1, CDate(0)), gSTAFFID)
    Else
        iresult = oSM.GenerateReturnsPerPub(oSupp.ID, IIf(chkAllDels = 0, Me.DT1, CDate(0)), gSTAFFID)
    End If
    
    Set oSM = Nothing
    PB1.Visible = False
    If iresult > 0 Then
        MsgBox "Return generated", vbInformation, "Status"
    Else
        MsgBox "No returns possible", vbInformation, "Status"
    End If
    cmdFind1_Click
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cmdGenerate_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSupp_Click()
    On Error GoTo errHandler
Dim iresult As Integer

    Screen.MousePointer = vbHourglass
    Set oSM = New z_StockManager
    PB1.Visible = True
    iresult = oSM.GenerateReturnsPerPub(oSupp.ID, IIf(chkAllDels = 1, Me.DT1, CDate(0)), gSTAFFID, True)
    Set oSM = Nothing
    PB1.Visible = False
    If iresult > 0 Then
        MsgBox "Return generated", vbInformation, "Status"
    Else
        MsgBox "No returns possible", vbInformation, "Status"
    End If
    cmdFind1_Click
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cmdSupp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        Grid_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        Find
        LoadArray
        Grid.ReBind
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler

   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN13(txtArg))
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.ArgIsProductCode"
End Function
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
    ErrorIn "frmReturn1.SetDateArgs"
End Sub

Private Sub Find()
    On Error GoTo errHandler
Dim bNotFound As Boolean
Dim frm As frmBrowseSUppliers2
Dim lngTPID As Long
    Screen.MousePointer = vbHourglass
    bNotFound = False
    If txtArg > " " Then
        enSince = 1
        cbSince.Caption = TranslateSince(1)
        If ArgIsProductCode Then
            'Search for product code
            Set cR = Nothing
            Set cR = New c_R
            cR.Load bNotFound, 0, "", "", dteDate1, dteDate2, , txtArg
            Exit Sub
        End If
    Else
        If oSupp Is Nothing Then
            enSince = enMonth
            SetDateArgs
            cR.Load bNotFound, 0, "", "", dteDate1, dteDate2
        Else
            SetDateArgs
            cR.Load bNotFound, oSupp.ID, "", "", dteDate1, dteDate2
        End If
    End If

EXIT_Handler:
    Screen.MousePointer = vbDefault
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
'    Set tlSupplier = New z_TextList
    Set cR = New c_R
    Set dR = New d_R
    If Me.WindowState <> 2 Then
        Me.TOP = 250
        Me.Left = 250
        Me.Width = 6500
        Me.Height = 7770
    End If
    LoadControls
    cmdFind1_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
  '  Set tlSupplier = Nothing
    Set cR = Nothing
    Set dR = Nothing
  '  Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = ""
    lngTPID = 0
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_R
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cR.Count, 1, 6
    For i = 1 To cR.Count
        With objItem
            XA.Value(i, 1) = cR.Item(i).DocDateF
            XA.Value(i, 3) = cR.Item(i).StatusF
            XA.Value(i, 4) = cR.Item(i).DateForSort
            XA.Value(i, 2) = cR.Item(i).DOCCode & IIf(cR.Item(i).RType = "E", " (S)", "")
            XA.Value(i, 5) = cR.Item(i).TRID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
Dim frm As Form
    Screen.MousePointer = vbHourglass
    If cR.Item(XA(Grid.Bookmark, 5) & "k").Status = 2 Then
        Set frm = New frmReturn2
        frm.component cR.Item(XA(Grid.Bookmark, 5) & "k"), oSupp.NameAndCode(25)
        frm.Show
    Else
        Set frm = New frmReturn3
        frm.component cR.Item(XA(Grid.Bookmark, 5) & "k") ', oSupp.NameAndCode(25)
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmReturn1: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmReturn1: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Grid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 6) = "VOID" Or XA(Bookmark, 6) = "CANCELLED" Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    End If
    If XA(Bookmark, 6) = "IN PROCESS" Then
        RowStyle.BackColor = &H80FF80
    End If
    If XA(Bookmark, 6) = "COMPLETE" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn1.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub



