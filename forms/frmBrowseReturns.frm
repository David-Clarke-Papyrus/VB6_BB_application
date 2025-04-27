VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseReturns 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Returns"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   Icon            =   "frmBrowseReturns.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7260
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowseReturns.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4950
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5820
      Picture         =   "frmBrowseReturns.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4950
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
      Height          =   1140
      Left            =   90
      TabIndex        =   0
      Top             =   -45
      Width           =   6780
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   2310
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
         ToolTipText     =   "Enter product code or document number. Hit ENTER to fetch."
         Top             =   270
         Width           =   2500
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5265
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowseReturns.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1000
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
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for . . ."
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   1230
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3660
      Left            =   90
      OleObjectBlob   =   "frmBrowseReturns.frx":1028
      TabIndex        =   3
      Top             =   1230
      Width           =   6780
   End
End
Attribute VB_Name = "frmBrowseReturns"
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
Dim xMLDoc As ujXML

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.mnuSaveLayout"
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
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.SetMenu"
End Sub


Private Sub cbSince_Click()
    On Error GoTo errHandler
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
  '  txtArg = ""
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.cbSince_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseReturns.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.cmdClose_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseReturns.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub Grid_GotFocus()
'    Shape1.Visible = True
'End Sub
'Private Sub Grid_LostFocus()
'    Shape1.Visible = False
'End Sub






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
    ErrorIn "frmBrowseReturns.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Form_Deactivate", , EA_NORERAISE
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
    cmdClose.Left = NonNegative_Lng(Grid.Width - 1000)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Form_Resize", , EA_NORERAISE
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
    ErrorIn "frmBrowseReturns.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    str = "Notes" & vbCrLf _
            & "Enter product code or document number or start of supplier name followed by and asterisk." & vbCrLf _
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
    ErrorIn "frmBrowseReturns.Label3_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseReturns.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler

   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.ArgIsProductCode"
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
    ErrorIn "frmBrowseReturns.SetDateArgs"
End Sub

Private Sub Find()
    On Error GoTo errHandler
Dim bNotFound As Boolean
Dim frm As frmBrowseSUppliers2
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
            'Search for product code
            enSince = 1
            cbSince.Caption = TranslateSince(1)
            Set cR = Nothing
            Set cR = New c_R
            cR.Load bNotFound, 0, "", "", , , , txtArg
            Exit Sub
        End If
        If txtArg = "\" Then
            'Search for unissued POs
            Set cR = Nothing
            Set cR = New c_R
            cR.Load bNotFound, 0, "", "", , , , , True
            GoTo EXIT_Handler
        End If
        'Search for Reference
        Set cR = Nothing
        Set cR = New c_R
        cR.Load bNotFound, 0, "", txtArg, dteDate1, dteDate2, "", ""
        If bNotFound Then
            'Search for customer by ACCNO
            Set cR = Nothing
            Set cR = New c_R
            SetDateArgs
            cR.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2, "", ""
            If bNotFound Then
               Set frm = New frmBrowseSUppliers2
               frm.component txtArg, lngCount
                If lngCount > 1 Then
                    frm.Show vbModal
                    lngTPID = frm.SupplierID
                ElseIf lngCount = 1 Then
                    lngTPID = frm.SupplierID
                End If
                Unload frm
               If lngTPID > 0 Then
                    Set cR = Nothing
                    Set cR = New c_R
                    SetDateArgs
                    cR.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
'                    If bNotFound Then
'                        'Search for customer by Approval ref
'                        Set cR = Nothing
'                        Set cR = New c_R
'                        SetDateArgs
'                        cR.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2, "", ""
'                        If bNotFound Then
'
               End If
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
        cR.Load bNotFound, 0, "", "", dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    Set cR = New c_R
    Set dR = New d_R
    If Me.WindowState <> 2 Then
        Me.TOP = 250
        Me.Left = 250
        Me.Width = 7100
        Me.Height = 6100
    End If
    SetGridLayout Me.Grid, Me.Name
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set cR = Nothing
    Set dR = Nothing
  '  Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = "\"
    lngTPID = 0
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_R
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cR.Count, 1, 8
    For i = 1 To cR.Count
  '      With objItem
            XA.Value(i, 1) = cR.Item(i).TPName & IIf(cR.Item(i).RType = "E", " (S)", "")
            XA.Value(i, 2) = cR.Item(i).DOCCode & cR.Item(i).StaffNameB
            XA.Value(i, 3) = cR.Item(i).DocDateF
            XA.Value(i, 4) = cR.Item(i).IssDateF   'DateForSort
            XA.Value(i, 5) = cR.Item(i).TRID
            XA.Value(i, 6) = cR.Item(i).StatusF
            XA.Value(i, 7) = cR.Item(i).RType
            XA.Value(i, 8) = cR.Item(i).TRID
 '       End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
Dim frm As Form

    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    'lngID = XA(Grid.Bookmark, 5)
    If XA(Grid.Bookmark, 7) = "S" And cR.Item(XA(Grid.Bookmark, 5) & "k").STATUS = 2 Then
        Set frm = New frmReturnPreview
        frm.component cR.Item(XA(Grid.Bookmark, 5) & "k").TRID
        frm.Show
    ElseIf cR.Item(XA(Grid.Bookmark, 5) & "k").STATUS = 2 Then
        Set frm = New frmReturn2
        frm.component cR.Item(XA(Grid.Bookmark, 5) & "k"), XA.Value(Grid.Bookmark, 1)
        frm.Show
    Else
        Set frm = New frmReturn3
        frm.Component2 XA(Grid.Bookmark, 8)
      '  frm.Component cR.Item(XA(Grid.Bookmark, 5) & "k"), XA.Value(Grid.Bookmark, 1)
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseReturns: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseReturns: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Grid_DblClick", , EA_NORERAISE
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
    If XA(Bookmark, 6) = "REQUESTED" Then
        RowStyle.BackColor = &HDBFAFB
    End If
    If XA(Bookmark, 6) = "STOCK RETURNED" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
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
    If ColIndex = 1 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 4, Direction, GetRowType(4) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 2, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 3
            GetRowType = XTYPE_STRING
        Case 2, 4
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.GetRowType(ColIndex)", ColIndex
End Function
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseReturns.cmdPrint_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseReturns.IsAmongBookmarks(TRID)", TRID
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
        .docInit "RET_1"
        .chCreate "RET"
            .elText = "Returns at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        For i = 1 To cR.Count
            If IsAmongBookmarks(cR(i).TRID) Then
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = cR.Item(i).TPName & IIf(cR.Item(i).RType = "E", " (S)", "")
            .elCreateSibling "Col_2"
                .elText = cR.Item(i).DocDateF
            .elCreateSibling "Col_3"
                .elText = cR.Item(i).DOCCode
            .elCreateSibling "Col_4"
                .elText = cR.Item(i).StatusF
                .navUP
            End If
        Next i

    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\RET" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\RET_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "RET_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\RET_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

    strFilename = oPC.SharedFolderRoot & "\RET.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\RET" & "_" & CStr(i) & ".RTF"
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
    ErrorIn "frmBrowseReturns.ExportToXML"
End Function
