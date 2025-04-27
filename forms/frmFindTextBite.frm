VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmFindTextBite 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Text bites"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindTextBite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cnmdCopy 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Copy selected"
      Height          =   345
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1950
      Width           =   1485
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   510
      Left            =   0
      Picture         =   "frmFindTextBite.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1905
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   4935
      Picture         =   "frmFindTextBite.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1905
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      DragIcon        =   "frmFindTextBite.frx":0A9E
      Height          =   1905
      Left            =   -15
      OleObjectBlob   =   "frmFindTextBite.frx":0E28
      TabIndex        =   1
      Top             =   0
      Width           =   5940
   End
End
Attribute VB_Name = "frmFindTextBite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mTL As z_TextList
Dim strRef As String
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim XA As New XArrayDB
Dim xMLDoc As ujXML


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
    ErrorIn "frmFindTextBite.SetMenu"
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
 '   Find
    mTL.Load ltTextBite
    
    LoadArray
    Grid.ReBind
    Grid.Bookmark = 1

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cnmdCopy_Click()
    On Error GoTo errHandler

    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Grid.text  'XA.Value(Grid.SelBookmarks(1), 1)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.cnmdCopy_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
Dim i As Integer

    SetMenu
    For i = 1 To Grid.Columns.Count
        Grid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid.Columns(i - 1).Width)
    Next
    Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
    Me.Height = GetSetting("PBKS", Me.Name, "FormHeight", Me.Height)
    Me.top = GetSetting("PBKS", Me.Name, "FormTop", Me.top)
    Me.Left = GetSetting("PBKS", Me.Name, "FormLeft", Me.Left)
    Me.Grid.RowHeight = GetSetting("PBKS", Me.Name, "RowHeight", 350)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub




Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    Grid.Width = NonNegative_Lng(Me.Width - (Grid.Left + 400))
    lngDiff = Grid.Height
    Grid.Height = NonNegative_Lng(Me.Height - (Grid.top + 1220))
    lngDiff = (Grid.Height - lngDiff)
    cmdPrint.top = cmdPrint.top + lngDiff
    cnmdCopy.top = cnmdCopy.top + lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    cmdClose.Left = NonNegative_Lng(Grid.Width - 1440)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
    On Error GoTo errHandler
    Grid.Col = ColIndex
    Grid.Bookmark = RowBookmark
    
    ' Set up drag operation, such as creating visual effects by
    ' highlighting the cell or row being dragged.
    
    ' Use VB manual drag support (put TDBGrid1 into drag mode)
    Grid.Drag vbBeginDrag

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Grid_DragCell(SplitIndex,RowBookmark,ColIndex)", Array(SplitIndex, _
         RowBookmark, ColIndex), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    On Error GoTo errHandler
Effect = 0
DefaultCursors = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Grid_OLEGiveFeedback(Effect,DefaultCursors)", Array(Effect, _
         DefaultCursors), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_OLEStartDrag(ByVal Data As TrueOleDBGrid60.DataObject, AllowedEffects As Long)
    On Error GoTo errHandler
'    AllowedEffects = 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Grid_OLEStartDrag(Data,AllowedEffects)", Array(Data, AllowedEffects), _
         EA_NORERAISE
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
    ErrorIn "frmFindTextBite.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


'Private Sub txtArg_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       ' Find
'        LoadArray
'        Grid.ReBind
'    End If
'End Sub

'Private Function ArgIsProductCode() As Boolean
'
'   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
'
'End Function
'Private Sub SetDateArgs()
'    Select Case enSince
'    Case enAny
'        dteDate1 = CDate("1995-01-01")
'        dteDate2 = DateAdd("d", 1, Date)
'    Case enWeek
'        dteDate1 = DateAdd("d", -7, Date)
'        dteDate2 = DateAdd("d", 1, Date)
'    Case enMonth
'        dteDate1 = DateAdd("m", -1, Date)
'        dteDate2 = DateAdd("d", 1, Date)
'    Case enQuarter
'        dteDate1 = DateAdd("q", -1, Date)
'        dteDate2 = DateAdd("d", 1, Date)
'    Case enYear
'        dteDate1 = DateAdd("yyyy", -1, Date)
'        dteDate2 = DateAdd("d", 1, Date)
'    End Select
'
'End Sub

'Private Sub Find()
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
'    On Error GoTo ERR_Handler
'    bNotFound = False
'    If UCase(Left(txtArg, 3)) = "YR=" Then byear = True
'    If txtArg > " " And Not (byear) Then
'        If ArgIsProductCode Then
'            'Search for product code
'            enSince = 1
'            Set mTL = Nothing
'            Set mTL = New Z_Textlist
'            mTL.Load bNotFound, 0, "", "", dteDate1, dteDate2, , txtArg
'            GoTo EXIT_HANDLER
'        End If
'    End If
'    mSetfocus Grid
'
'EXIT_HANDLER:
'    Exit Sub
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_HANDLER
'    Resume
'End Sub


'Private Sub cmdFind_LostFocus()
'    LoadControls
'End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
   ' Set tlSupplier = New z_TextList
    Set mTL = New z_TextList
    If Me.WindowState <> 2 Then
        Me.top = 2000
        Me.Left = 1000
        Me.Width = 6120
        Me.Height = 2970
    End If
    SetGridLayout Me.Grid, Me.Name
    cmdFind1_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set mTL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


'Private Sub LoadControls()
'    flgLoading = True
'    txtArg = ""
'    flgLoading = False
'End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, mTL.Count, 1, 6
    For i = 1 To mTL.Count
            XA.Value(i, 1) = mTL.f3ByOrdinalIndex(i)
            XA.Value(i, 2) = mTL.ItemByOrdinalIndex(i)
    Next
    XA.QuickSort 1, XA.UpperBound(1), 4, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
'Dim lngID As Long
'Dim blnEdit As Boolean
'    If IsNull(Grid.Bookmark) Then Exit Sub
'    Set ofrm = New frmAPPRPreview
'    lngID = val(XA(Grid.Bookmark, 5))
'    ofrm.Component lngID    ', False
'    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Grid_DblClick", , EA_NORERAISE
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
    ErrorIn "frmFindTextBite.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
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
    ErrorIn "frmFindTextBite.GetRowType(ColIndex)", ColIndex
End Function
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.cmdPrint_Click", , EA_NORERAISE
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
    ErrorIn "frmFindTextBite.IsAmongBookmarks(TRID)", TRID
End Function

Public Function ExportToXML() As Boolean
    On Error GoTo errHandler
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
'        .docInit "APPR_1"
'        .chCreate "APPR"
'            .elText = "Appro returns at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
'        For i = 1 To mTL.Count
'            If IsAmongBookmarks(mTL(i).TRID) Then
'            .elCreateSibling "DetailLine", True
'            .chCreate "Col_1"
'                .elText = mTL(i).TPName
'            .elCreateSibling "Col_2"
'                .elText = mTL(i).DocCode
'            .elCreateSibling "Col_3"
'                .elText = mTL(i).DocDateF
'            .elCreateSibling "Col_4"
'                .elText = mTL(i).statusF
'                .navUP
'            End If
'        Next i
'
'    End With
'
''FINALLY PRODUCE THE .XML FILE
'    strXML = oPC.SharedFolderRoot & "\TEMP\APPR" & ".xml"
'    With xMLDoc
'        If fs.FileExists(strXML) Then
'            fs.DeleteFile strXML
'        End If
'        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
'    End With
'
'''WRITE THE .HTML FILE
'    objXSL.async = False
'    objXSL.validateOnParse = False
'    objXSL.resolveExternals = False
'    strPath = oPC.SharedFolderRoot & "\Templates\APPR_RTF_1.xslt"
'    Set fs = New FileSystemObject
'    If fs.FileExists(strPath) Then
'        objXSL.Load strPath
'    End If
'
'    strFilename = oPC.LocalFolder & "\APPR.RTF"
'    If fs.FileExists(strFilename) Then
'        fs.DeleteFile strFilename, True
'    End If
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
'    ErrorIn "frmBrowseAPPRs.ExportToXML"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.ExportToXML"
End Function

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name
    SaveSetting "PBKS", Me.Name, "Formwidth", Me.Width
    SaveSetting "PBKS", Me.Name, "Formheight", Me.Height
    SaveSetting "PBKS", Me.Name, "FormTop", Me.top
    SaveSetting "PBKS", Me.Name, "FormLeft", Me.Left
    SaveSetting "PBKS", Me.Name, "RowHeight", Grid.RowHeight
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.mnuSaveLayout"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To Grid.Columns.Count
        Grid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid.Columns(i - 1).Width)
    Next
    Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
    Me.Height = GetSetting("PBKS", Me.Name, "FormHeight", Me.Height)
    Me.top = GetSetting("PBKS", Me.Name, "FormTop", Me.top)
    Me.Left = GetSetting("PBKS", Me.Name, "FormLeft", Me.Left)
    Me.Grid.RowHeight = GetSetting("PBKS", Me.Name, "RowHeight", 350)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmFindTextBite.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

