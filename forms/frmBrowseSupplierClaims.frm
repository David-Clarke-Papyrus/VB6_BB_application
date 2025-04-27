VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseSupplierClaims 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Claims"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   Icon            =   "frmBrowseSupplierClaims.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   8550
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   555
      Left            =   3000
      TabIndex        =   5
      Top             =   -60
      Width           =   1980
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Current"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   885
         TabIndex        =   7
         Top             =   180
         Width           =   1020
      End
      Begin VB.OptionButton optAll 
         BackColor       =   &H00D3D3CB&
         Caption         =   "All"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdFind1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Find"
      Height          =   615
      Left            =   5115
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmBrowseSupplierClaims.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   -30
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   15
      Picture         =   "frmBrowseSupplierClaims.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4290
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5790
      Picture         =   "frmBrowseSupplierClaims.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4305
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Height          =   3555
      Left            =   15
      OleObjectBlob   =   "frmBrowseSupplierClaims.frx":1028
      TabIndex        =   0
      Top             =   570
      Width           =   8370
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
      Left            =   6285
      TabIndex        =   4
      Top             =   75
      Width           =   495
   End
End
Attribute VB_Name = "frmBrowseSupplierClaims"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSCL As c_SCL
Dim dSCL As d_SCL
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
Dim mArg As String

Public Sub component(arg As String)
    mArg = arg
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.mnuSaveLayout"
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
    ErrorIn "frmBrowseSupplierClaims.SetMenu"
End Sub



Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.cmdClose_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseSupplierClaims.cmdFind1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
Dim bm As Variant
    SetMenu
'    txtArg = ""
'    mSetfocus Me.txtArg
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Form_Deactivate", , EA_NORERAISE
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
    ErrorIn "frmBrowseSupplierClaims.Form_Resize", , EA_NORERAISE
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
    ErrorIn "frmBrowseSupplierClaims.Grid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtArg_KeyPress(KeyAscii As Integer)
'    On Error GoTo errHandler
'    If KeyAscii = 13 Then
'        Find
'        LoadArray
'        Grid.ReBind
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseSupplierClaims.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
'    HandleError
'End Sub

'Private Function ArgIsProductCode() As Boolean
'    On Error GoTo errHandler
'
'   ArgIsProductCode = (IsHashCode(txtArg) Or IsISBN10(txtArg) Or IsISBN13(txtArg))
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseSupplierClaims.ArgIsProductCode"
'End Function
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
 
        If (Me.optAll = False) Then
            Set cSCL = Nothing
            Set cSCL = New c_SCL
            cSCL.Load bNotFound, 0, "", "", , , True
            GoTo EXIT_Handler
        Else
            Set cSCL = Nothing
            Set cSCL = New c_SCL
            cSCL.Load bNotFound, 0, "", "", , , False
            GoTo EXIT_Handler
        End If
        'Search for Reference
'        Set cSCL = Nothing
'        Set cSCL = New c_SCL
'        cSCL.Load bNotFound, 0, "", txtArg, "", ""
'        If bNotFound Then
'            'Search for customer by ACCNO
'            Set cSCL = Nothing
'            Set cSCL = New c_SCL
'            cSCL.Load bNotFound, 0, txtArg, "", "", ""
'            If bNotFound Then
'               Set frm = New frmBrowseSUppliers2
'               frm.component txtArg, lngCount
'                If lngCount > 1 Then
'                    frm.Show vbModal
'                    lngTPID = frm.SupplierID
'                ElseIf lngCount = 1 Then
'                    lngTPID = frm.SupplierID
'                End If
'                Unload frm
'               If lngTPID > 0 Then
'                    Set cSCL = Nothing
'                    Set cSCL = New c_SCL
'                    cSCL.Load bNotFound, lngTPID, "", ""
'               End If
'            End If
'        End If
 '   Else
   '      cSCL.Load bNotFound, 0, "", ""
    
  '  End If

EXIT_Handler:
    mSetfocus Grid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Find"
End Sub


Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    Set cSCL = New c_SCL
    Set dSCL = New d_SCL
    If Me.WindowState <> 2 Then
        Me.TOP = 250
        Me.Left = 250
        Me.Width = 7100
        Me.Height = 6100
    End If
    SetGridLayout Me.Grid, Me.Name
    SetFormSize Me
    LoadControls
    If mArg > "" Then
  '      txtArg = mArg
        cmdFind1_Click
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout Me.Grid, Me.Name & Grid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
    UnsetMenu
    Set cSCL = Nothing
    Set dSCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
 '   txtArg = "\"
    lngTPID = 0
    enSince = enWeek
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.LoadControls"
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_SCL
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.Clear
    XA.ReDim 1, cSCL.Count, 1, 10
    For i = 1 To cSCL.Count
  '      With objItem
            XA.Value(i, 1) = cSCL.Item(i).SupplierName
            XA.Value(i, 2) = cSCL.Item(i).DOCCode
            XA.Value(i, 3) = cSCL.Item(i).DocDateF
            XA.Value(i, 4) = cSCL.Item(i).ClaimValueF
            XA.Value(i, 5) = cSCL.Item(i).DocStatusF
            XA.Value(i, 8) = cSCL.Item(i).TPID
            XA.Value(i, 9) = cSCL.Item(i).ClaimNeedsApproval
            XA.Value(i, 10) = cSCL.Item(i).TRID
 '       End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.LoadArray"
End Sub

Private Sub Grid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
Dim frm As frmSCPreview

    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set frm = New frmSCPreview
    frm.component cSCL.Item(XA(Grid.Bookmark, 10) & "k").TRID, FNS(XA.Value(Grid.Bookmark, 5)), FNS(XA.Value(Grid.Bookmark, 4)), FNS(XA.Value(Grid.Bookmark, 1)), _
        FNS(XA.Value(Grid.Bookmark, 9)), FNN(XA.Value(Grid.Bookmark, 8))
    frm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseSupplierClaims: Grid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseSupplierClaims: Grid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Grid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 5) = "Closed" Then
        RowStyle.BackColor = &HFFFFC0
  '      RowStyle.Font.Strikethrough = True
    End If
    If XA(Bookmark, 5) <> "Closed" Then
        RowStyle.BackColor = &H80FF80
    End If
'    If XA(Bookmark, 6) = "REQUESTED" Then
'        RowStyle.BackColor = &HDBFAFB
 '   End If
 '   If XA(Bookmark, 6) = "STOCK RETURNED" Then
 '       RowStyle.BackColor = &HFFFFC0
 '   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.Grid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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
    ErrorIn "frmBrowseSupplierClaims.Grid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
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
    ErrorIn "frmBrowseSupplierClaims.GetRowType(ColIndex)", ColIndex
End Function
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSupplierClaims.cmdPrint_Click", , EA_NORERAISE
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
    ErrorIn "frmBrowseSupplierClaims.IsAmongBookmarks(TRID)", TRID
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
        For i = 1 To cSCL.Count
            If IsAmongBookmarks(cSCL(i).TRID) Then
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = cSCL.Item(i).SupplierName
            .elCreateSibling "Col_2"
                .elText = cSCL.Item(i).DocDateF
            .elCreateSibling "Col_3"
                .elText = cSCL.Item(i).DOCCode
            .elCreateSibling "Col_4"
                .elText = cSCL.Item(i).DocStatusF
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
    ErrorIn "frmBrowseSupplierClaims.ExportToXML"
End Function
