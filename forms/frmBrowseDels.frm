VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseDels 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse Goods received notes"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseDels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   120
      Picture         =   "frmBrowseDels.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4995
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   1275
      Picture         =   "frmBrowseDels.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4995
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
      Top             =   -90
      Width           =   6840
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5220
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowseDels.frx":0C9E
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
         Left            =   2580
         TabIndex        =   0
         ToolTipText     =   "Enter product code, product number Acc no. or start of supplier name followed by '*'. Hit ENTER to fetch."
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
         Left            =   6270
         TabIndex        =   6
         Top             =   330
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
         Left            =   2730
         TabIndex        =   2
         Top             =   720
         Width           =   1980
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3825
      Left            =   105
      OleObjectBlob   =   "frmBrowseDels.frx":1028
      TabIndex        =   5
      Top             =   1110
      Width           =   9645
   End
End
Attribute VB_Name = "frmBrowseDels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cDEL As c_DELs
Dim dDel As d_DEL
Dim tlSupplier As z_TextList
Dim lngTPID As Long
Dim strRef As String
Dim enSince As enumSince
Dim dteDate1 As Date
Dim dteDate2 As Date
Dim strDate1 As String
Dim strDate2 As String
Dim blnNoRecordsReturned As Boolean
Dim flgLoading As Boolean
Dim ofrm As frmDELPreview
Dim XA As New XArrayDB
Dim xMLDoc As ujXML

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.mnuSaveLayout"
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
    ErrorIn "frmBrowseDels.SetMenu"
End Sub


Private Sub cbSince_Click()
    On Error GoTo errHandler
    enSince = OptionLoop(enSince, 6)
    cbSince.Caption = TranslateSince(CInt(enSince))
    mSetfocus txtArg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.cbSince_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseDels.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.cmdClose_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseDels.cmdFind1_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseDels.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Form_Deactivate", , EA_NORERAISE
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
    ErrorIn "frmBrowseDels.Form_Resize", , EA_NORERAISE
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
    ErrorIn "frmBrowseDels.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      PopupMenu Forms(0).mnuBrowseDeliveriesPopup   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Grid_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub PrintSelectedDeliveries()
    On Error GoTo errHandler

Dim IDset As String
Dim Doccodes As String
Dim SQL As String

Dim i As Long
Dim OpenResult As Integer
Dim arGRNC As New arGRNVerification
Dim rs As ADODB.Recordset
    
    If Grid = "" Or Grid = "No records" Then Exit Sub
    If Grid.Bookmark = 0 Then Exit Sub
    
    If Grid.SelBookmarks.Count = 0 Then
        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
        Exit Sub
    End If
    
    For i = 0 To Grid.SelBookmarks.Count - 1
        IDset = IDset & Replace(CStr(XA(Grid.SelBookmarks(i), 10)), "K", "") & ","
        Doccodes = Doccodes & FNS(XA(Grid.SelBookmarks(i), 12)) & ","
    Next i
    IDset = Left(IDset, Len(IDset) - 1)
    Doccodes = Left(Doccodes, Len(Doccodes) - 1)
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        Set rs = New ADODB.Recordset
        
        SQL = "Select * from vGRNVerification WHERE TRID in (" & IDset & ") " _
        & " ORDER BY SupplierDocument"

        
        rs.Open SQL, oPC.COShort, adOpenForwardOnly
        
        arGRNC.component rs, Doccodes
        arGRNC.Show vbModal
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.mnuBrowseInvoicesPopup_Print_Click"
End Sub
Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    str = "Notes" & vbCrLf _
            & "Enter product code, product number Acc no. or start of supplier name followed by '*'." & vbCrLf _
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
    ErrorIn "frmBrowseDels.Label3_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseDels.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler

   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.ArgIsProductCode"
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
    ErrorIn "frmBrowseDels.SetDateArgs"
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
  '  txtArg = Replace(txtArg, " ", "")
    If UCase(Left(txtArg, 3)) = "YR=" Then byear = True
    If txtArg > " " And Not (byear) Then
        If ArgIsProductCode Then
            'Search for product code
            enSince = 1
            cbSince.Caption = TranslateSince(1)
            Set cDEL = Nothing
            Set cDEL = New c_DELs
            cDEL.Load bNotFound, 0, "", "", , , , txtArg
            GoTo EXIT_Handler
        End If
        If txtArg = "\" Then
            'Search for unissued POs
            Set cDEL = Nothing
            Set cDEL = New c_DELs
            cDEL.Load bNotFound, 0, "", "", , , , , , , True
            GoTo EXIT_Handler
        End If
        'Search for Reference
        Set cDEL = Nothing
        Set cDEL = New c_DELs
        cDEL.Load bNotFound, 0, "", txtArg ', dteDate1, dteDate2
        If bNotFound Then
            'Search for customer by ACCNO
            Set cDEL = Nothing
            Set cDEL = New c_DELs
            cDEL.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
            If bNotFound Then
                'Search for Invoice by supplier invoice no
                Set cDEL = Nothing
                Set cDEL = New c_DELs
                SetDateArgs
                cDEL.Load bNotFound, 0, "", "", dteDate1, dteDate2, , , , txtArg
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
                        Set cDEL = Nothing
                        Set cDEL = New c_DELs
                        SetDateArgs
                        cDEL.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
                    End If
                End If
            
            Else
                enSince = 1
                cbSince.Caption = TranslateSince(1)
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
        cDEL.Load bNotFound, 0, "", "", dteDate1, dteDate2
    End If
    Grid.Visible = True

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    Set tlSupplier = New z_TextList
    Set cDEL = New c_DELs
    Set dDel = New d_DEL
    SetMenu
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Width = 9250
        Me.Height = 6200
    End If
    
    
    SetGridLayout Me.Grid, Me.Name
    SetFormSize Me
    LoadControls
    cmdFind1_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set tlSupplier = Nothing
    Set cDEL = Nothing
    Set dDel = Nothing
    Set ofrm = Nothing
    SaveLayout Me.Grid, Me.Name, Me.Height, Me.Width
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmBrowseDels.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_DEL
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Long
    XA.Clear
    XA.ReDim 1, cDEL.Count, 1, 13
    For i = 1 To cDEL.Count
        With objItem
            XA.Value(i, 1) = cDEL(i).TPNAME & (IIf(Len(Trim(cDEL(i).TPAccNo)) <= 1, "", "(" & Trim(cDEL(i).TPAccNo) & ")"))
            XA.Value(i, 2) = cDEL(i).Ref & cDEL(i).StaffNameB
            XA.Value(i, 3) = cDEL(i).DocDateF
            XA.Value(i, 4) = cDEL(i).InvoiceValueF & "(" & cDEL(i).InvoiceQtyF & ")"
            XA.Value(i, 5) = cDEL(i).InvoiceRef & ", " & cDEL(i).InvoiceDateF  'cDEL(i).InvoiceValueF & "(" & cDEL(i).InvoiceQtyF & ")"
            XA.Value(i, 6) = cDEL(i).InvoiceRef
            XA.Value(i, 7) = cDEL(i).InvoiceDateF
            XA.Value(i, 8) = cDEL(i).InvoiceShortF
            XA.Value(i, 9) = cDEL(i).DateForSort  'DateForSort
            XA.Value(i, 10) = cDEL(i).TRID & "K"
            XA.Value(i, 11) = cDEL(i).StatusF
            XA.Value(i, 12) = cDEL(i).DOCCode
            
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 9, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set ofrm = New frmDELPreview
    lngID = val(XA(Grid.Bookmark, 10))
    ofrm.component lngID    ', False
    ofrm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseDels: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseDels: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Grid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 11) = "VOID" Or XA(Bookmark, 11) = "CANCELLED" Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    End If
    If XA(Bookmark, 11) = "IN PROCESS" Then
        RowStyle.BackColor = &H80FF80
    End If
    If XA(Bookmark, 11) = "COMPLETE" Then
        RowStyle.BackColor = &HFFFFC0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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
    If ColIndex = 2 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 9, Direction, XTYPE_DATE
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 6
            GetRowType = XTYPE_STRING
        Case 3, 7
            GetRowType = XTYPE_DATE
        Case 8, 4, 5
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.GetRowType(ColIndex)", ColIndex
End Function
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function IsAmongBookmarks(TRID As Long) As Boolean
    On Error GoTo errHandler
    Dim i As Integer
    IsAmongBookmarks = False
    For i = 1 To Grid.SelBookmarks.Count
        If val(XA.Value(Grid.SelBookmarks(i - 1), 10)) = TRID Then
            IsAmongBookmarks = True
            Exit For
        End If
    Next i
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseDels.IsAmongBookmarks(TRID)", TRID
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
        .docInit "GRN_1"
        .chCreate "GRN"
            .elText = "Goods received at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        For i = 1 To cDEL.Count
            If IsAmongBookmarks(cDEL(i).TRID) Then
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = cDEL(i).TPNAME & (IIf(Len(Trim(cDEL(i).TPAccNo)) <= 1, "", "(" & Trim(cDEL(i).TPAccNo) & ")"))
            .elCreateSibling "Col_2"
                .elText = cDEL(i).Ref
            .elCreateSibling "Col_3"
                .elText = cDEL(i).DocDateF
            .elCreateSibling "Col_4"
                .elText = cDEL(i).StatusF
            .elCreateSibling "Col_5"
                .elText = cDEL(i).InvoiceRef
            .elCreateSibling "Col_6"
                .elText = cDEL(i).InvoiceDate
            .elCreateSibling "Col_7"
                .elText = cDEL(i).InvoiceValueF
            .elCreateSibling "Col_8"
                .elText = cDEL(i).InvoiceQty
            .elCreateSibling "Col_9"
                .elText = cDEL(i).InvoiceShort
                .navUP
            End If
        Next i
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\DEL" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\GRN_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "GRN_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\GRN_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

'    strFilename = oPC.LocalFolder & "\GRN_1.RTF"
'    If fs.FileExists(strFilename) Then
'        fs.DeleteFile strFilename, True
'    End If
    strFilename = oPC.SharedFolderRoot & "\GRN_1.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\GRN_1" & "_" & CStr(i) & ".RTF"
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
    ErrorIn "frmBrowseDels.ExportToXML"
End Function

