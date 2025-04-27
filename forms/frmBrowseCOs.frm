VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmBrowseCOs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse orders"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseCOs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowseCOs.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5910
      Picture         =   "frmBrowseCOs.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4530
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
      Height          =   1110
      Left            =   90
      TabIndex        =   1
      Top             =   -75
      Width           =   6870
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5265
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowseCOs.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   2610
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
         Left            =   6330
         TabIndex        =   7
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search by . . ."
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
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
   End
   Begin CoolButtonControl.CoolButton CB1 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "You can look for ordinary orders or wants."
      Top             =   4740
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   556
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Type: Ordinary orders"
      Style           =   1
      BackStyle       =   0
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3345
      Left            =   75
      OleObjectBlob   =   "frmBrowseCOs.frx":1028
      TabIndex        =   6
      Top             =   1110
      Width           =   6885
   End
End
Attribute VB_Name = "frmBrowseCOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mcol As c_COs
Dim dispCO As d_CO
Dim tlCustomer As z_TextList
Dim lngTPID As Long
Dim strRef As String
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim enCOType As enumCOType
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim ofrm As frmInvoicePreview
Dim ofrmCO As frmCOPreview
Dim XA As New XArrayDB
Dim xMLDoc As ujXML
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.mnuSaveLayout"
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
    ''Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.SetMenu"
End Sub


Private Sub CB1_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    enCOType = OptionLoop(enCOType, 2)
    Select Case enCOType
    Case enNormalCO
        cb1.Caption = "&Type: Ordinary orders"
    Case enWant
        cb1.Caption = "&Type: Wants"
    End Select
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.CB1_Click", , EA_NORERAISE
    HandleError
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
    ErrorIn "frmBrowseCOs.cbSince_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseCOs.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
 Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Screen.MousePointer = vbHourglass
    Find
    LoadArray
    Grid.ReBind
    Grid.Bookmark = 1
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
Dim bm As Variant
    SetMenu
'    bm = Grid.Bookmark
'    cmdFind1_Click
'    Grid.Bookmark = bm
    txtArg = ""
    mSetfocus Me.txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    Grid.Width = NonNegative_Lng(Me.Width - (Grid.Left + 400))
    lngDiff = Grid.Height
    Grid.Height = NonNegative_Lng(Me.Height - (Grid.TOP + 1220))
    lngDiff = (Grid.Height - lngDiff)
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdPrint.TOP + lngDiff
    cb1.TOP = cb1.TOP + lngDiff
    cmdClose.Left = NonNegative_Lng(Grid.Width - 1440)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If KeyAscii = 13 Then
        Grid_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
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
    ErrorIn "frmBrowseCOs.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler
   
   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.ArgIsProductCode"
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
    ErrorIn "frmBrowseCOs.SetDateArgs"
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
        If ArgIsProductCode Then
            'Search for product code
            enSince = 1
            cbSince.Caption = TranslateSince(1)
            Set mcol = Nothing
            Set mcol = New c_COs
            mcol.Load bNotFound, enCOType, 0, "", "", , , , , txtArg
            Exit Sub
        End If
        If txtArg = "\" Then
            'Search for unissued POs
            Set mcol = Nothing
            Set mcol = New c_COs
            mcol.Load bNotFound, enCOType, 0, "", "", , , , , , , , True
            GoTo EXIT_Handler
        End If
        'Search for Reference
        Set mcol = Nothing
        Set mcol = New c_COs
            'Search by Our document Code
        bNotFound = False
        mcol.Load bNotFound, enCOType, 0, "", txtArg
        If bNotFound Then
            Set mcol = Nothing
            Set mcol = New c_COs
            bNotFound = False
            'Search by Doc reference
            mcol.Load bNotFound, enCOType, 0, "", , , , , , , txtArg
            If bNotFound Then
                'Search for customer by ACCNO
                Set mcol = Nothing
                Set mcol = New c_COs
                SetDateArgs
                mcol.Load bNotFound, enCOType, 0, txtArg
'Search by line ref no
                If bNotFound Then
                    Set mcol = Nothing
                    Set mcol = New c_COs
                    SetDateArgs
                    mcol.Load bNotFound, enCOType, 0, , "", , , , , , , txtArg
                
'Search by customer name
                    If bNotFound Then
                        Set frm = New frmBrowseCustomers2
                        frm.component txtArg, lngCount
                        If lngCount > 1 Then
                            frm.Show vbModal
                            lngTPID = frm.CustomerID
                        ElseIf lngCount = 1 Then
                            lngTPID = frm.CustomerID
                        End If
                        Unload frm
                        If lngTPID > 0 Then
                           Set mcol = Nothing
                           Set mcol = New c_COs
                           SetDateArgs
                           mcol.Load bNotFound, enCOType, lngTPID
                        End If
                    End If
                End If
            End If
        Else
            enSince = 1
            cbSince.Caption = TranslateSince(1)
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
        mcol.Load bNotFound, enCOType, 0, "", "", dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCOs_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.lvwCOs_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCOs_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.lvwCOs_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    flgLoading = True
    enCOType = enNormalCO
    Set tlCustomer = New z_TextList
    Set mcol = New c_COs
    Set dispCO = New d_CO
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Width = 7250
        Me.Height = 5700
    End If
    LoadControls
    cmdFind1_Click

    cb1.Visible = oPC.Configuration.SupportsWants
    SetGridLayout Me.Grid, Me.Name
    SetFormSize Me
    
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set tlCustomer = Nothing
    Set mcol = Nothing
    Set dispCO = Nothing
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = "\"
    strDate1 = ""
    strDate2 = ""
    lngTPID = 0
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, mcol.Count, 1, 10
    For i = 1 To mcol.Count
        With objItem
            XA.Value(i, 1) = mcol(i).CustomerDisplay & (IIf(Len(Trim(mcol(i).StaffNameB)) <= 1, "", Trim(mcol(i).StaffNameB)))
          '  XA.Value(i, 1) = mcol(i).TPNAME & (IIf(Len(Trim(mcol(i).TPACCNo)) <= 1, "", "(" & Trim(mcol(i).TPACCNo) & ")"))
            XA.Value(i, 2) = mcol(i).DOCCode & mcol(i).StaffNameB
            XA.Value(i, 3) = mcol(i).DocDateF
            XA.Value(i, 4) = mcol(i).DOCDate
            XA.Value(i, 5) = mcol(i).TRID & "K"
            XA.Value(i, 6) = mcol(i).StatusF
            XA.Value(i, 7) = mcol(i).DateForSort
            XA.Value(i, 10) = mcol(i).TPName
        End With
    Next
    If enCOType = enNormalCO Then
        XA.QuickSort 1, XA.UpperBound(1), 7, XORDER_DESCEND, XTYPE_DATE
    Else
        XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    End If
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If flgLoading Then Exit Sub
    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set ofrmCO = New frmCOPreview
    lngID = val(XA(Grid.Bookmark, 5))
    ofrmCO.component lngID, False
    ofrmCO.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseCOs: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseCOs: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Grid_DblClick", , EA_NORERAISE
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
    ErrorIn "frmBrowseCOs.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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
    If ColIndex = 0 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, Direction, GetRowType(10) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    ElseIf ColIndex = 2 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 4, Direction, GetRowType(4) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 10
            GetRowType = XTYPE_STRING
        Case 3, 4
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    If flgLoading Then Exit Sub
    str = "Notes" & vbCrLf _
            & "Enter product code, document number, Acc no.,or start of customer name followed by '*'." & vbCrLf _
            & "Hit ENTER to fetch. " & vbCrLf & vbCrLf _
            & "Search for old data like this . . . " & vbCrLf _
            & "yr=2002     fetches all records for 2002" & vbCrLf & vbCrLf _
            & "yr=2002-03     fetches all records for March 2002" & vbCrLf & vbCrLf _
            & "'\'     fetches all unissued records" & vbCrLf & vbCrLf _
            & "Maximum records returned is settable in PBKS.INI file (ask support person)" & vbCrLf _
            & "This is currently set at " & oPC.MaxBrowseRecs & " records" & vbCrLf
    MsgBox str, vbInformation, "Help"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.Label3_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function IsAmongBookmarks(TRID As Long) As Boolean
    On Error GoTo errHandler
    Dim i As Integer
    IsAmongBookmarks = False
    For i = 1 To Grid.SelBookmarks.Count
        If val(XA.Value(Grid.SelBookmarks(i - 1), 5)) = TRID Then
            IsAmongBookmarks = True
            Exit For
        End If
    Next i
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.IsAmongBookmarks(TRID)", TRID
End Function
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
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strExecutable As String

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "CO_1"
        .chCreate "CO"
            .elText = "Customer orders at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        For i = 1 To mcol.Count
            If IsAmongBookmarks(mcol(i).TRID) Then
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = mcol(i).TPName & (IIf(Len(Trim(mcol(i).TPAccNo)) <= 1, "", "(" & Trim(mcol(i).TPAccNo) & ")"))
            .elCreateSibling "Col_2"
                .elText = mcol(i).DOCCode & mcol(i).StaffNameB
            .elCreateSibling "Col_3"
                .elText = mcol(i).DocDateF
            .elCreateSibling "Col_4"
                .elText = mcol(i).StatusF
                .navUP
            End If
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

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\CO_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "CO_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\CO_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

    strFilename = oPC.SharedFolderRoot & "\CO.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\CO" & "_" & CStr(i) & ".RTF"
    Loop
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile
    
    strExecutable = GetPDFExecutable(strFilename)
        If strExecutable = "" Then
            MsgBox "There is no application set on this computer to open the file: " & strFilename & ". The document cannot be displayed", vbOKOnly, "Can't do this"
        Else
            Shell strExecutable & " " & strFilename, vbNormalFocus
        End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCOs.ExportToXML"
End Function

