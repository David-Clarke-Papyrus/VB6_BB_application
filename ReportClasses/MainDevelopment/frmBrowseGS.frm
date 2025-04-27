VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseGS 
   BackColor       =   &H00D3D3CB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse products"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   1005
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseGS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11490
   Begin VB.CheckBox chkServices 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Services"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7740
      TabIndex        =   14
      Top             =   1140
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CheckBox chkNewspapers 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Non stock-take item (e.g. newspaper)"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7740
      TabIndex        =   13
      Top             =   795
      Width           =   3630
   End
   Begin VB.CheckBox chkGeneralProducts 
      BackColor       =   &H00D3D3CB&
      Caption         =   "General products"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7740
      TabIndex        =   12
      Top             =   450
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.CheckBox chkBooks 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Books"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7740
      TabIndex        =   11
      Top             =   105
      Visible         =   0   'False
      Width           =   1890
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4095
      Left            =   45
      OleObjectBlob   =   "frmBrowseGS.frx":058A
      TabIndex        =   10
      Top             =   1710
      Width           =   11250
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   10275
      Picture         =   "frmBrowseGS.frx":5769
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5850
      Width           =   1035
   End
   Begin VB.CommandButton cmdSaveLayout 
      BackColor       =   &H00C4BCA4&
      Caption         =   "save layout"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5850
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   1605
      Left            =   45
      TabIndex        =   3
      Top             =   -75
      Width           =   7545
      Begin VB.CommandButton cmdClearSection 
         BackColor       =   &H00D3C9C0&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1230
         Width           =   255
      End
      Begin VB.CommandButton cmdClearPT 
         BackColor       =   &H00D3C9C0&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   795
         Width           =   255
      End
      Begin VB.CommandButton cmdClearSearch 
         BackColor       =   &H00D3C9C0&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   390
         Width           =   255
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         ItemData        =   "frmBrowseGS.frx":5814
         Left            =   1230
         List            =   "frmBrowseGS.frx":5816
         TabIndex        =   19
         Top             =   210
         Width           =   2940
      End
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   690
         Width           =   2565
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2580
      End
      Begin VB.TextBox txtRecsFound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   6540
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   660
         Width           =   900
      End
      Begin VB.TextBox txtmaxnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   6540
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   900
      End
      Begin VB.CheckBox chkCopies 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Stock on hand"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   5100
         TabIndex        =   1
         Top             =   1200
         Width           =   2010
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Search"
         Height          =   855
         Left            =   4590
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   810
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Product type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   330
         TabIndex        =   17
         Top             =   735
         Width           =   1230
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Found"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5895
         TabIndex        =   9
         Top             =   690
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Max"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6060
         TabIndex        =   5
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Search for"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   255
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmBrowseGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strthing As String
Dim tlkeys As z_TextList
Private oSearchEngine As z_SearchEngineB
Dim colList As Collection
Dim intShowCopies As Integer
Dim lslist As ListItem
Dim roProduct As a_Product
Dim enSource As enProductDataSource
Dim mnu As Menu
Dim XA As New XArrayDB
Dim XN As New XArrayDB
Dim strTime As String
Dim tlCats As z_TextList

Dim tlSuppliers As z_TextList
Dim bWithCopies As Boolean

Private Sub cboProductType_DblClick()
    cboProductType = ""
End Sub

Private Sub cboProductType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        mSetfocus GN
    End If
    Exit Sub
End Sub

Private Sub cboSection_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        mSetfocus GN
    End If
    Exit Sub
End Sub

Private Sub chkCopies_Click()
    On Error GoTo errHandler
    oSearchEngine.instock chkCopies
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.chkCopies_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Search(pSearchType As enSearchType, pCriteria As String, Optional pSection As String, Optional pProductType As String)
    On Error GoTo errHandler
Dim strParsedCriteria As String
Dim lngRecsFound As Long
Dim lngResult As Long
Dim lngrows As Long
Dim strArticle As String
Dim strNet As String
Dim strTypes As String
Dim lngSectionID As Long
Dim lngProductTypeID As Long

    strTypes = ""
    If chkBooks Then strTypes = "B"
    If chkGeneralProducts Then strTypes = strTypes & "G"
    If chkNewspapers Then strTypes = strTypes & "M"
    If chkServices Then strTypes = strTypes & "N"
    If strTypes = "" Then
        chkGeneralProducts = 1
        strTypes = "G"
    End If
    
    txtRecsFound = ""
    lngSectionID = 0
    lngProductTypeID = 0
    
    StripArticle pCriteria, strArticle, strNet
    pCriteria = strNet
    oSearchEngine.prisearch
    enSource = enLocalDB
    '--------------
    oPC.OpenDBSHort
    '--------------
    oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, CLng(txtmaxnum) + 1, strTypes '"NGM"
    
    If pSearchType = enSearchByCatalogue Then
        oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchNormal Then
        oSearchEngine.SimpleSearch pCriteria, lngRecsFound
'    ElseIf pSearchType = enSearchBF Then
'        enSource = enBF
'        oSearchEngine.BFSearchEx pCriteria, lngRecsFound, CLng(txtmaxnum), lngResult
    Else
        enSource = enLocalDB
        If pSection <> "<ALL>" Then
            lngSectionID = oPC.Configuration.Sections.Key(pSection)
        End If
        If pProductType <> "<ALL>" Then
            lngProductTypeID = oPC.Configuration.ProductTypes.Key(pProductType)
        End If
        oSearchEngine.AdvancedSearch lngRecsFound, pCriteria, lngSectionID, lngProductTypeID
    End If
    'If lngRecsFound > CLng(txtmaxnum) Then MsgBox "Too many records to return, refine your search.", vbInformation + vbOKOnly, "Search result"
    oSearchEngine.Execute IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
    Set colList = Nothing
    Set colList = oSearchEngine.getcols
    lngrows = oSearchEngine.Rows
    txtRecsFound = lngRecsFound
    LoadGrid
    If colList.Count = 0 Then
        Select Case enSource
        Case enLocalDB
            XN.ReDim 1, 1, 1, 12
            XN(1, 1) = "No records"
            GN.ReBind
        End Select
    End If
    If CLng(txtRecsFound) > CLng(txtmaxnum) Then
        MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
        Me.GN.Refresh
    End If
    '--------------
    oPC.DisconnectDBShort
    '--------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.Search(pSearchType,pCriteria)", Array(pSearchType, pCriteria)
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Long

    Select Case enSource
    Case enLocalDB
        GN.Visible = True
        XN.Clear
        XN.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
                XN.Value(i, 1) = colList.Item(i).CodeF
                XN.Value(i, 2) = colList.Item(i).statusF & " " & colList.Item(i).Title
                XN.Value(i, 3) = colList.Item(i).Author
                XN.Value(i, 4) = colList.Item(i).Publisher
                XN.Value(i, 5) = colList.Item(i).QtyOnHand
                XN.Value(i, 6) = colList.Item(i).QtyOnOrder
                XN.Value(i, 7) = colList.Item(i).QtyOnBackorder
                XN.Value(i, 8) = colList.Item(i).QtyTotalSold
                XN.Value(i, 10) = colList.Item(i).LastDateDelivered
                XN.Value(i, 9) = colList.Item(i).LocalPriceF
                XN.Value(i, 11) = colList.Item(i).pID
                XN.Value(i, 12) = colList.Item(i).code
        Next
        XN.QuickSort 1, XN.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
        GN.Array = XN
        Me.GN.ReBind
        
        
        
    End Select
'Errh:
'    MsgBox Error
'    Exit Sub
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.LoadGrid"
End Sub



Private Sub cmdClearSearch_Click()
    cboSearch = ""
End Sub
Private Sub cmdClearPT_Click()
    cboProductType = "<ALL>"
End Sub
Private Sub cmdClearSection_Click()
    cboSection = "<ALL>"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdSaveLayout_Click()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To GN.Columns.Count
        SaveSetting "PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo errHandler
    
    cboSearch.AddItem cboSearch, 0
    oSearchEngine.instock chkCopies
    
    Screen.MousePointer = vbHourglass
    
    cboSearch = FNS(cboSearch)
    If InStr(cboSearch, "/") > 0 Or cboSection <> "<ALL>" Or cboProductType <> "<ALL>" Then
        Search enSearchAdvanced, cboSearch, cboSection, cboProductType
        mSetfocus GN
    Else
        Search enSearchNormal, cboSearch
        mSetfocus GN
    End If
'    txtcritvalues = FNS(txtcritvalues)
'    If InStr(txtcritvalues, "/") > 0 Then
'        Search enSearchAdvanced, txtcritvalues
'        mSetfocus GN
'    Else
'        If txtcritvalues <> "" Then
'            Search enSearchNormal, txtcritvalues
'            mSetfocus GN
'        Else
'            MsgBox "You must specify a search text.", vbInformation, "Can't do this"
'        End If
'    End If
    mSetfocus cboSearch
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        mSetfocus GN
    End If
End Sub




Private Sub Form_Activate()
    On Error GoTo errHandler
    XA.Clear
    XA.ReDim 1, 1, 1, 7
    cboSearch.SetFocus
    bWithCopies = False
    chkCopies = IIf(bWithCopies, 1, 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim i As Integer
    Set oSearchEngine = New z_SearchEngineB
    Set colList = New Collection
    
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
    XA.ReDim 1, 1, 1, 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    Me.top = 20
    Me.left = 50
    Height = 6925
    Set tlSuppliers = New z_TextList
    tlSuppliers.Load ltSupplier, ""
    
'    Set tlCats = Nothing
'    Set tlCats = New z_TextList
'
'    tlCats.Load ltCatalogue
    
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    Me.cboSection = "<ALL>"
    Me.cboProductType = "<ALL>"
    
  '  LoadCombo cboCat, tlCats

    GN.Columns(3).Caption = "Supplier"
    txtmaxnum = 500
    
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oSearchEngine = Nothing
    Set roProduct = Nothing
    Set colList = Nothing
    Set tlkeys = Nothing
    Set lslist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub


Private Sub GN_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GN_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub


Private Sub GN_DblClick()
    On Error GoTo errHandler
'Dim frmA As frmProductPrevAQ
'Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev
Dim lngprod As Long
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 11))
    If str = "" Then Exit Sub
    Set roProduct = New a_Product
    WaitMsg "Loading . . .", True, Me
    roProduct.Load str, 0, "", strTime
    If roProduct.pID = "" Then Exit Sub
    
    Set frmNB = New frmProductNBPrev
    frmNB.Component roProduct, strTime
    frmNB.Show

    Set roProduct = Nothing
    WaitMsg "", False, Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GN_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub GN_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 0 Then ColIndex = 11
    
        XN.QuickSort XN.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    GN.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4, 12
            GetRowType = XTYPE_STRING
        Case 5, 6, 7, 8, 9
            GetRowType = XTYPE_INTEGER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GetRowType(ColIndex)", ColIndex
End Function

Private Sub GN_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    
    If KeyAscii = vbKeyReturn Then
        GN_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GN_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub GN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuFindForm   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.GN_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), EA_NORERAISE
    HandleError
End Sub

Public Sub AddToTempList()
    On Error GoTo errHandler
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 11))
    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
        If XA(XA.UpperBound(1), 1) > "" Then
            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
        End If
        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
        XA(XA.UpperBound(1), 4) = 1
        XA(XA.UpperBound(1), 5) = 0
        XA(XA.UpperBound(1), 6) = ""
        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.AddToTempList"
End Sub
Public Sub PlaceCO()
    On Error GoTo errHandler
Dim frm As New frmPlaceCO
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 1))
    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
        If XA(XA.UpperBound(1), 1) > "" Then
            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
        End If
        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
        XA(XA.UpperBound(1), 4) = 1
        XA(XA.UpperBound(1), 5) = 0
        XA(XA.UpperBound(1), 6) = ""
        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
    End If
    frm.Component XA, "ORDER"
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.PlaceCO"
End Sub
Public Sub PlaceOnReserve()
    On Error GoTo errHandler
Dim frm As New frmPlaceCO
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 11))
    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
        If XA(XA.UpperBound(1), 1) > "" Then
            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
        End If
        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
        XA(XA.UpperBound(1), 4) = 1
        XA(XA.UpperBound(1), 5) = 0
        XA(XA.UpperBound(1), 6) = ""
        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
    End If
    frm.Component XA, "RESERVE"
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.PlaceOnReserve"
End Sub
Public Sub StartNewList()
    On Error GoTo errHandler
    XA.Clear
    XA.ReDim 1, 1, 1, 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.StartNewList"
End Sub



Private Sub txtcritvalues_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseGS.txtcritvalues_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
