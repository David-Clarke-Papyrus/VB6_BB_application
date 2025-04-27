VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmBrowseCS 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse daily cash sales"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseCSs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   5610
      Picture         =   "frmBrowseCSs.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4710
      Width           =   1035
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
      Height          =   1170
      Left            =   120
      TabIndex        =   1
      Top             =   -75
      Width           =   6510
      Begin VB.CommandButton cbSince 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Since: Last week"
         Height          =   450
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   2310
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   615
         Left            =   5190
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmBrowseCSs.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   2550
         TabIndex        =   0
         ToolTipText     =   "Enter product code, document number,Acc no. Hit ENTER to fetch."
         Top             =   270
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
         Left            =   5850
         TabIndex        =   6
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for . . ."
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2670
         TabIndex        =   2
         Top             =   750
         Width           =   1980
      End
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3465
      Left            =   135
      OleObjectBlob   =   "frmBrowseCSs.frx":0C9E
      TabIndex        =   5
      Top             =   1170
      Width           =   6510
   End
End
Attribute VB_Name = "frmBrowseCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCS As c_CSs
Dim dCS As d_CS
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
Dim ofrm As frmCSPreview
Dim XA As New XArrayDB

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.mnuSaveLayout"
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
    ErrorIn "frmBrowseCS.SetMenu"
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
    ErrorIn "frmBrowseCS.cbSince_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseCS.cbSince_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind1_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Find
    LoadArray
    Grid.ReBind
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.cmdFind1_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseCS.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    Grid.Width = NonNegative_Lng(Me.Width - (Grid.Left + 400))
    lngDiff = Grid.Height
    Grid.Height = NonNegative_Lng(Me.Height - (Grid.TOP + 1220))
    lngDiff = (Grid.Height - lngDiff)
    cmdClose.TOP = cmdClose.TOP + lngDiff
    cmdClose.Left = NonNegative_Lng(Grid.Width - 1000)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Form_Resize", , EA_NORERAISE
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
    ErrorIn "frmBrowseCS.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub Label3_Click()
    On Error GoTo errHandler
Dim str As String
    str = "Notes" & vbCrLf _
            & "Enter product code, document number, Acc no. or start of customer name followed by '*'. " & vbCrLf _
            & "Hit ENTER to fetch. " & vbCrLf & vbCrLf _
            & "Search for old data like this . . . " & vbCrLf _
            & "yr=2002     fetches all records for 2002" & vbCrLf & vbCrLf _
            & "yr=2002-03     fetches all records for March 2002" & vbCrLf & vbCrLf _
            & "Maximum records returned is settable in PBKS.INI file (ask support person)" & vbCrLf _
            & "This is currently set at " & oPC.MaxBrowseRecs & " records" & vbCrLf
    MsgBox str, vbInformation, "Help"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Label3_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseCS.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Function ArgIsProductCode() As Boolean
    On Error GoTo errHandler

   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.ArgIsProductCode"
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
    ErrorIn "frmBrowseCS.SetDateArgs"
End Sub

Private Sub Find()
    On Error GoTo errHandler
Dim bNotFound As Boolean
Dim lngTPID As Long
Dim byear As Boolean
Dim yr As String
Dim mth As String
Dim strDate1 As String
Dim strDate2 As String

    bNotFound = False
    If UCase(Left(txtArg, 3)) = "YR=" Then byear = True
    If txtArg > " " And Not (byear) Then
        If ArgIsProductCode Then
            'Search for product code
            enSince = 1
            cbSince.Caption = TranslateSince(1)
            Set cCS = Nothing
            Set cCS = New c_CSs
            cCS.Load bNotFound, "", dteDate1, dteDate2, , txtArg
            GoTo EXIT_Handler
        End If
        'Search for Reference
        Set cCS = Nothing
        Set cCS = New c_CSs
        SetDateArgs
        cCS.Load bNotFound, txtArg, dteDate1, dteDate2
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
        cCS.Load bNotFound, "", dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    Set tlCustomer = New z_TextList
    Set cCS = New c_CSs
    Set dCS = New d_CS
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Width = 6900
        Me.Height = 5800
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set tlCustomer = Nothing
    Set cCS = Nothing
    Set dCS = Nothing
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtArg = ""
    enSince = enWeek
    cbSince.Caption = TranslateSince(CInt(enSince))
    lngTPID = 0
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_CS
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cCS.Count, 1, 6
    For i = 1 To cCS.Count
        With objItem
            XA.Value(i, 1) = cCS.Item(i).DOCCode
            XA.Value(i, 2) = Format(cCS(i).DOCDate, "dddd dd/mm/yyyy")
           ' XA.Value(i, 3) = cCS(i).statusF
            XA.Value(i, 4) = cCS(i).DateForSort
            XA.Value(i, 5) = cCS(i).TRID & "K"
            XA.Value(i, 6) = cCS(i).StatusF
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 4, XORDER_DESCEND, XTYPE_STRING
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set ofrm = New frmCSPreview
    lngID = val(XA(Grid.Bookmark, 5))
    ofrm.component lngID    ', False
    ofrm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseCS: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseCS: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Grid_DblClick", , EA_NORERAISE
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
    ErrorIn "frmBrowseCS.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    Else
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 2, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    End If
    
    Grid.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1
            GetRowType = XTYPE_STRING
        Case 2
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCS.GetRowType(ColIndex)", ColIndex
End Function

