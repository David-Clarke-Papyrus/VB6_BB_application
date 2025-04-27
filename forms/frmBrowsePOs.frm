VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowsePOs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse purchase orders"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowsePOs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmBrowsePOs.frx":058A
   ScaleHeight     =   5700
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowsePOs.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4935
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10125
      Picture         =   "frmBrowsePOs.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4935
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
      Height          =   1050
      Left            =   120
      TabIndex        =   1
      Top             =   -75
      Width           =   7680
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   210
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5310
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowsePOs.frx":1028
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   195
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
         Left            =   2550
         TabIndex        =   0
         ToolTipText     =   "Enter product code,  Acc/ no. or document number or start of supplier name followed by '*'. Hit ENTER to fetch."
         Top             =   210
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
         Left            =   6570
         TabIndex        =   6
         Top             =   270
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
         Top             =   690
         Width           =   1350
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3885
      Left            =   120
      OleObjectBlob   =   "frmBrowsePOs.frx":13B2
      TabIndex        =   5
      Top             =   1005
      Width           =   11010
   End
End
Attribute VB_Name = "frmBrowsePOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cPO As c_POs
Dim dPO As d_PO
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
Dim ofrm As frmPOPreview
Dim XA As New XArrayDB
Dim xMLDoc As ujXML
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.mnuSaveLayout"
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
    ErrorIn "frmBrowsePOs.SetMenu"
End Sub


Private Sub cbSince_Click()
    On Error GoTo errHandler
    enSince = OptionLoop(enSince, 7)
    cbSince.Caption = TranslateSince(CInt(enSince))
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.cbSince_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowsePOs.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.cmdClose_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowsePOs.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()

    On Error GoTo errHandler
Dim bm As Variant

    SetMenu
 '   bm = Grid.Bookmark
 '   cmdFind1_Click
 '   Grid.Bookmark = bm
    txtArg = ""
    mSetfocus Me.txtArg
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Form_Deactivate", , EA_NORERAISE
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
    cmdClose.Left = NonNegative_Lng(Grid.Width - 1440)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Grid_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

'Private Sub Grid_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
'If Col <> 0 Then Exit Sub
'    If FNN(XA(Bookmark, 12)) > 0 Then
'        CellStyle.ForegroundPicture = LoadResPicture(103, vbResBitmap)
'        'On Error Resume Next
'        'Set CellStyle.ForegroundPicture = LoadPicture(oPC.SharedFolderRoot & "\Templates\Pagelink.BMP")
'        CellStyle.ForegroundPicturePosition = dbgFPLeft
'    End If
'
'End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    
    If XA(Bookmark, 9) = "VOID" Or XA(Bookmark, 9) = "CANCELLED" Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    End If
    If XA(Bookmark, 9) = "IN PROCESS" Then
        RowStyle.BackColor = &H80FF80
    End If
    If XA(Bookmark, 9) = "COMPLETE" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
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
    ErrorIn "frmBrowsePOs.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
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
    ErrorIn "frmBrowsePOs.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler

   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.ArgIsProductCode"
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
    ErrorIn "frmBrowsePOs.SetDateArgs"
End Sub

Private Sub Find()
    On Error GoTo errHandler
Dim byear As Boolean
Dim yr As String
Dim mth As String
Dim strDate1 As String
Dim strDate2 As String
Dim bNotFound As Boolean
Dim frm As frmBrowseSUppliers2
Dim lngTPID As Long
Dim lngCount As Long


    bNotFound = False
    If Left(txtArg, 3) = "yr=" Then byear = True
    If txtArg > " " And Not (byear) Then
        If ArgIsProductCode Then
            'Search for product code
            enSince = 1
            cbSince.Caption = TranslateSince(1)
            Set cPO = Nothing
            Set cPO = New c_POs
            cPO.Load bNotFound, 0, "", "", , , , , txtArg
            GoTo EXIT_Handler
        End If
        If txtArg = "\" Then
            'Search for unissued POs
            Set cPO = Nothing
            Set cPO = New c_POs
            cPO.Load bNotFound, 0, "", "", , , , , , , True
            GoTo EXIT_Handler
        End If
        'Search for Reference
        Set cPO = Nothing
        Set cPO = New c_POs
            'Search by document reference
        cPO.Load bNotFound, 0, , , txtArg
        If bNotFound Then
           Set cPO = Nothing
           Set cPO = New c_POs
           SetDateArgs
            bNotFound = False
            'Search by line reference
           cPO.Load bNotFound, 0, "", txtArg
           If bNotFound Then
               'Search for customer by Supplier account number
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
                    Set cPO = Nothing
                    Set cPO = New c_POs
                    SetDateArgs
                    bNotFound = False
                    cPO.Load bNotFound, lngTPID, "", "", , dteDate1, dteDate2
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
        cPO.Load bNotFound, 0, "", "", , dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    Set tlCustomer = New z_TextList
    Set cPO = New c_POs
    Set dPO = New d_PO
    SetMenu
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Width = 11000
        Me.Height = 6200
    End If
    LoadControls
    cmdFind1_Click
    
    SetGridLayout Me.Grid, Me.Name
    SetFormSize Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set tlCustomer = Nothing
    Set cPO = Nothing
    Set dPO = Nothing
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = "\"
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    lngTPID = 0
    flgLoading = False
    AutoSelect txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_PO
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cPO.Count, 1, 14
    For i = 1 To cPO.Count
        With objItem
            XA.Value(i, 1) = cPO(i).TPName & (IIf(Len(Trim(cPO(i).TPAccNo)) <= 1, "", "(" & Trim(cPO(i).TPAccNo) & ")"))
            XA.Value(i, 2) = cPO(i).Ref & cPO(i).StaffNameB
            XA.Value(i, 3) = cPO(i).DocDateF & cPO(i).OrderType
            XA.Value(i, 4) = cPO(i).Totals
            XA.Value(i, 5) = cPO(i).DispatchMode
            XA.Value(i, 6) = cPO(i).Log
            XA.Value(i, 7) = cPO(i).DateForSort
            XA.Value(i, 8) = cPO(i).TRID & "K"
            XA.Value(i, 9) = cPO(i).StatusF
            XA.Value(i, 12) = cPO(i).ParentID
            XA.Value(i, 14) = cPO(i).TPName & (IIf(Len(Trim(cPO(i).TPAccNo)) <= 1, "", "(" & Trim(cPO(i).TPAccNo) & ")"))
            
            If FNN(XA(i, 12)) > 0 Then
                XA.Value(i, 1) = "<<- " & XA.Value(i, 1)
            End If
            
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 7, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set ofrm = New frmPOPreview
    lngID = val(XA(Grid.Bookmark, 8))
    ofrm.component lngID    ', False
    ofrm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowsePOs: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowsePOs: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Grid_DblClick", , EA_NORERAISE
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
    If ColIndex = 0 Then ColIndex = 13
    If ColIndex = 2 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 6, Direction, GetRowType(6) ', 6, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 13
            GetRowType = XTYPE_STRING
        Case 4, 6
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    str = "Notes" & vbCrLf _
            & "Enter product code, Acc/no. or document number or start of supplier name followed by '*'." & vbCrLf _
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
    ErrorIn "frmBrowsePOs.Label3_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function IsAmongBookmarks(TRID As Long) As Boolean
    On Error GoTo errHandler
    Dim i As Integer
    IsAmongBookmarks = False
    For i = 1 To Grid.SelBookmarks.Count
        If val(XA.Value(Grid.SelBookmarks(i - 1), 8)) = TRID Then
            IsAmongBookmarks = True
            Exit For
        End If
    Next i
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.IsAmongBookmarks(TRID)", TRID
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
        .docInit "PO_B_1"
        .chCreate "POB"
            .elText = "Purchase orders at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        For i = 1 To cPO.Count
            If IsAmongBookmarks(cPO(i).TRID) Then
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = cPO(i).TPName & (IIf(Len(Trim(cPO(i).TPAccNo)) <= 1, "", "(" & Trim(cPO(i).TPAccNo) & ")"))
            .elCreateSibling "Col_2"
                .elText = cPO(i).Ref & cPO(i).StaffNameB
            .elCreateSibling "Col_3"
                .elText = cPO(i).DocDateF
            .elCreateSibling "Col_4"
                .elText = cPO(i).Ref & cPO(i).Log
            .elCreateSibling "Col_5"
                .elText = cPO(i).StatusF
            .elCreateSibling "Col_6"
                .elText = cPO(i).DispatchMode
                .navUP
            End If
        Next i

        
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\PO_B" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\PO_B_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "PO_B_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\PO_B_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

'    strFilename = oPC.SharedFolderRoot & "\PO.RTF"
'    If fs.FileExists(strFilename) Then
'        fs.DeleteFile strFilename, True
'    End If
'
    
    strFilename = oPC.SharedFolderRoot & "\PO.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\PO" & "_" & CStr(i) & ".RTF"
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
    ErrorIn "frmBrowsePOs.ExportToXML"
End Function

Public Sub PrepareDetailList()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim sTableName As String
Dim i As Integer
Dim rpt As New arPOCaptureSheet
Dim rs As ADODB.Recordset
Dim ret As Long

    sTableName = "dbo.tTransactionList_" & Replace(oPC.WorkstationName, "-", "")

    oSQL.RunProc "PrepareTransactionListTable", Array(sTableName), ""
    oSQL.RunSQL ("TRUNCATE TABLE " & sTableName)
    For i = 0 To Grid.SelBookmarks.Count - 1
        If val(XA.Value(Grid.SelBookmarks(i), 12)) > 0 Then
            oSQL.RunSQL "INSERT INTO " & sTableName & " (TR_ID) VALUES (" & val(XA.Value(Grid.SelBookmarks(i), 8)) & ")"
        End If
    Next i
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    ret = oSQL.RunGetRecordset("SELECT * FROM vInternetSupplierSheet a JOIN " & sTableName & " b ON a.TR_ID = b.TR_ID ORDER BY PARENT,TITLE", enText, Array(), "", rs)
    If rs.eof = False Then
            rpt.component rs
            rpt.Show
        Else
            MsgBox "No records to return", vbOKOnly, "Status"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.PrepareDetailList"
End Sub
Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuActionTransactionList   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.Grid_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub


