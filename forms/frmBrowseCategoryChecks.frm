VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseCategoryChecks 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse category checks"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12300
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseCategoryChecks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   60
      Picture         =   "frmBrowseCategoryChecks.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10785
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmBrowseCategoryChecks.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      UseMaskColor    =   -1  'True
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
      Height          =   1320
      Left            =   90
      TabIndex        =   1
      Top             =   -75
      Width           =   6810
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   210
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5220
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowseCategoryChecks.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   210
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
         Left            =   2580
         TabIndex        =   0
         ToolTipText     =   "Enter product code, document numberAcc no. or start of customer name followed by '*'. Hit ENTER to fetch."
         Top             =   210
         Width           =   2500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   2640
         TabIndex        =   2
         Top             =   660
         Width           =   1200
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3705
      Left            =   90
      OleObjectBlob   =   "frmBrowseCategoryChecks.frx":1028
      TabIndex        =   5
      Top             =   1290
      Width           =   11685
   End
End
Attribute VB_Name = "frmBrowseCategoryChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim XA As New XArrayDB
Dim xMLDoc As ujXML
Dim cCatChk As c_CATCHK
Dim frm As frmCategoryChecks

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.mnuSaveLayout"
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
    ErrorIn "frmBrowseCategoryChecks.SetMenu"
End Sub


Private Sub cbSince_Click()
    On Error GoTo errHandler
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.cbSince_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Find
    LoadArray
    Grid.ReBind
    Grid.Bookmark = 1
    
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Activate()
    On Error GoTo errHandler
Dim bm As Variant
    SetMenu
'    bm = Grid.Bookmark
'    cmdFind1_Click
'    Grid.Bookmark = bm
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
    ErrorIn "frmBrowseCategoryChecks.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Form_Deactivate", , EA_NORERAISE
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
    cmdClose.TOP = cmdClose.TOP + lngDiff
    cmdClose.Left = NonNegative_Lng(Grid.Width - 1140)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Form_Resize", , EA_NORERAISE
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
    ErrorIn "frmBrowseCategoryChecks.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
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
    ErrorIn "frmBrowseCategoryChecks.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler

   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.ArgIsProductCode"
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
    ErrorIn "frmBrowseCategoryChecks.SetDateArgs"
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
    If UCase(Left(txtArg, 3)) = "YR=" Then byear = True
    If txtArg > " " And Not (byear) Then
        If ArgIsProductCode Then
            enSince = 1
            cbSince.Caption = TranslateSince(1)
            'Search for product code
            Set cCatChk = Nothing
            Set cCatChk = New c_CATCHK
            cCatChk.Load bNotFound, 0, "", "", , , , txtArg
            GoTo EXIT_Handler
        End If
        'Search for category code
        Set cCatChk = Nothing
        Set cCatChk = New c_CATCHK
        cCatChk.Load bNotFound, , txtArg
        If bNotFound Then
            'Search for customer by category description
            Set cCatChk = Nothing
            Set cCatChk = New c_CATCHK
            SetDateArgs
            cCatChk.Load bNotFound, txtArg, ""
            If bNotFound Then
                'Search for customer by ACCNO
                Set cCatChk = Nothing
                Set cCatChk = New c_CATCHK
                SetDateArgs
                cCatChk.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
            Else
                enSince = 1
                cbSince.Caption = TranslateSince(1)
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
        Set cCatChk = Nothing
        Set cCatChk = New c_CATCHK
        cCatChk.Load bNotFound, "", "", , dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Width = 9000
        Me.Height = 6200
    End If
    SetMenu
    SetGridLayout Me.Grid, Me.Name
    SetFormSize Me
    LoadControls
    cmdFind1_Click
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = ""
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_CATCHK
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cCatChk.Count, 1, 14
    For i = 1 To cCatChk.Count
        With objItem
            XA.Value(i, 1) = cCatChk.Item(i).DocDateF
            XA.Value(i, 2) = cCatChk.Item(i).DOCCode
            XA.Value(i, 3) = cCatChk.Item(i).CategoryCode
            XA.Value(i, 4) = cCatChk.Item(i).CategoryName
            XA.Value(i, 5) = cCatChk.Item(i).OperatorName
            XA.Value(i, 6) = cCatChk.Item(i).SupervisorName
            XA.Value(i, 7) = cCatChk.Item(i).SupervisorName
            XA.Value(i, 8) = cCatChk.Item(i).CATCHKID & "K"
            XA.Value(i, 9) = cCatChk.Item(i).DateForSort
            XA.Value(i, 10) = cCatChk.Item(i).CategoryCode
            XA.Value(i, 11) = cCatChk.Item(i).SignedOFFID
            XA.Value(i, 12) = cCatChk.Item(i).StatusName
            XA.Value(i, 13) = cCatChk.Item(i).StatusasString
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 9, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
Dim s As String
Dim bEmpty As Boolean

    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frm = New frmCategoryChecks
    lngID = val(XA(Grid.Bookmark, 8))
    s = "Category check " & FNS(XA(Grid.Bookmark, 2)) & " for:" & FNS(XA(Grid.Bookmark, 4)) & "(" & FNS(XA(Grid.Bookmark, 5)) & ")" & "  started: " & FNS(XA.Value(Grid.Bookmark, 1))
    frm.component lngID, s, FNS(XA.Value(Grid.Bookmark, 7)), FNS(FNN(XA.Value(Grid.Bookmark, 11))), bEmpty
    If Not bEmpty Then
        frm.Show
    Else
        Set frm = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
  errRepeat = errRepeat + 1
  LogSaveToFile "Access violation in frmBrowseCategoryChecks: Grid_DblClick"  'unknown source
  If errRepeat < 5 Then
      Resume Next
  Else
      LogSaveToFile "Access violation in frmBrowseCategoryChecks: Grid_DblClick after 5 re-attempts"
      MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
      Err.Clear
      Exit Sub
  End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Grid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 12) = "VOID" Or XA(Bookmark, 12) = "CANCELLED" Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    End If
    If XA(Bookmark, 12) = "IN PROCESS" Then
        RowStyle.BackColor = &H80FF80
    End If
    If XA(Bookmark, 12) = "Op.Checked" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    If XA(Bookmark, 12) = "COMPLETE" Then
        RowStyle.BackColor = RGB(186, 200, 245) 'RGB(238, 238, 238)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, _
         Bookmark, RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
 '   If ColIndex = 2 Then ColIndex = 4
    If ColIndex = 0 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 8, Direction, XTYPE_DATE
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1)
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 4
            GetRowType = XTYPE_STRING
        Case 3
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCategoryChecks.GetRowType(ColIndex)", ColIndex
End Function
'Private Sub cmdPrint_Click()
'    ExportToXML
'End Sub
'Private Function IsAmongBookmarks(TRID As Long) As Boolean
'    Dim i As Integer
'    IsAmongBookmarks = False
'    For i = 1 To Grid.SelBookmarks.Count
'        If val(XA.Value(Grid.SelBookmarks(i - 1), 5)) = TRID Then
'            IsAmongBookmarks = True
'            Exit For
'        End If
'    Next i
'End Function
'Public Function ExportToXML() As Boolean
'    On Error GoTo errHandler
'Dim oTF As New z_TextFile
'Dim strPath As String
'Dim strBillto As String
'Dim strDelto As String
'Dim strFOFile As String
'Dim strFilename As String
'Dim strXML As String
'Dim strCommand As String
'Dim i As Integer
'Dim strHTML As String
'Dim fs As New FileSystemObject
'Dim objXSL As New MSXML2.DOMDocument60
'Dim opXMLDOC As New MSXML2.DOMDocument60
'Dim objXMLDOC  As New MSXML2.DOMDocument60
'Dim strExecutable As String
'
'    Set xMLDoc = New ujXML
'    With xMLDoc
'        .docProgID = "MSXML2.DOMDocument"
'        .docInit "COSR_1"
'        .chCreate "COSR"
'            .elText = "Invoices at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
'        For i = 1 To cCOSR.Count
'            If IsAmongBookmarks(cCOSR(i).TRID) Then
'            .elCreateSibling "DetailLine", True
'            .chCreate "Col_1"
'                .elText = cCOSR(i).TPNAME & (IIf(Len(Trim(cCOSR(i).TPACCNo)) <= 1, "", "(" & Trim(cCOSR(i).TPACCNo) & ")")) & (IIf(Len(Trim(cCOSR(i).SMSHortname)) <= 1, "", "(" & Trim(cCOSR(i).SMSHortname) & ")"))
'            .elCreateSibling "Col_2"
'                .elText = cCOSR(i).Ref
'            .elCreateSibling "Col_3"
'                .elText = cCOSR(i).TDateF
'            .elCreateSibling "Col_4"
'                .elText = cCOSR(i).statusF
'                .navUP
'            End If
'        Next i
'
'    End With
'
''FINALLY PRODUCE THE .XML FILE
'    strXML = oPC.SharedFolderRoot & "\TEMP\COSR" & ".xml"
'    With xMLDoc
'        If fs.FileExists(strXML) Then
'            fs.DeleteFile strXML
'        End If
'        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
'    End With
'
'''WRITE THE .RTF FILE
'    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\COSR_RTF_1.xslt") Then
'        MsgBox "You are missing the template file " & "COSR_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
'    End If
'    objXSL.async = False
'    objXSL.validateOnParse = False
'    objXSL.resolveExternals = False
'    strPath = oPC.SharedFolderRoot & "\Templates\COSR_RTF_1.xslt"
'    Set fs = New FileSystemObject
'    If fs.FileExists(strPath) Then
'        objXSL.Load strPath
'    End If
'
'    strFilename = oPC.SharedFolderRoot & "\COSR.RTF"
'    i = 0
'    Do Until fs.FileExists(strFilename) = False
'        i = i + 1
'        strFilename = oPC.SharedFolderRoot & "\COSR" & "_" & CStr(i) & ".RTF"
'    Loop
'    oTF.OpenTextFileToAppend strFilename
'    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
'    oTF.CloseTextFile
'
'    strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
'    Shell strExecutable, vbNormalFocus
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseInvoices.ExportToXML"
'End Function

