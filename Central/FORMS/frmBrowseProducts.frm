VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseProducts 
   BackColor       =   &H00D3D3CB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse books"
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
   Icon            =   "frmBrowseProducts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11490
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4455
      Left            =   75
      OleObjectBlob   =   "frmBrowseProducts.frx":0442
      TabIndex        =   15
      Top             =   1590
      Width           =   11250
   End
   Begin TrueOleDBGrid60.TDBGrid GBF 
      Height          =   4455
      Left            =   60
      OleObjectBlob   =   "frmBrowseProducts.frx":5741
      TabIndex        =   14
      Top             =   1575
      Width           =   11265
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   10320
      Picture         =   "frmBrowseProducts.frx":A419
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6030
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
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6015
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "By BIC"
      ForeColor       =   &H8000000D&
      Height          =   1125
      Left            =   7200
      TabIndex        =   9
      Top             =   135
      Width           =   1935
      Begin VB.CommandButton cmdBIC 
         BackColor       =   &H00C4BCA4&
         Caption         =   "BIC"
         Height          =   510
         Left            =   315
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   1335
      Left            =   45
      TabIndex        =   6
      Top             =   -75
      Width           =   6705
      Begin VB.TextBox txtRecsFound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   660
         Width           =   900
      End
      Begin VB.TextBox txtmaxnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   5700
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   210
         Width           =   900
      End
      Begin VB.CheckBox chkCopies 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Copies on hand"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2295
         TabIndex        =   4
         Top             =   930
         Width           =   2010
      End
      Begin VB.CheckBox chkAntiquarianOnly 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Antiquarian only"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   930
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Search"
         Height          =   435
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txtcritvalues 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   405
         Left            =   210
         TabIndex        =   0
         Top             =   435
         Width           =   3555
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   405
         Left            =   4665
         TabIndex        =   16
         Top             =   435
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Found"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5055
         TabIndex        =   13
         Top             =   690
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5220
         TabIndex        =   8
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   195
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmBrowseProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strthing As String
Dim tlkeys As z_TextList
Private oSearchEngine As z_SearchEngineC
Dim colList As Collection
Dim intShowCopies As Integer
'Dim rsdata As New ADOR.Recordset
'Dim lslist As ListItem
Dim roProduct As a_Product
Dim enSource As enProductDataSource
Dim mnu As Menu
Dim XA As New XArrayDB
Dim XBF As New XArrayDB
Dim XN As New XArrayDB
Dim strTime As String
Dim tlSuppliers As z_TextList
Dim tlCats As z_TextList

Private Sub chkCopies_Click()
    On Error GoTo errHandler
    oSearchEngine.instock chkCopies
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.chkCopies_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdBIC_Click()
    On Error GoTo errHandler
Dim frm As frmBICTree
Dim strBICCode As String
    Set frm = New frmBICTree
    frm.Show vbModal
    strBICCode = frm.SelectedCode
    Unload frm
    Search enSearchBIC, strBICCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdBIC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Search(pSearchType As enSearchType, pCriteria As String)
    On Error GoTo errHandler
Dim strParsedCriteria As String
Dim lngRecsFound As Long
Dim lngResult As Long
Dim lngrows As Long
Dim strArticle As String
Dim strNet As String
    
    txtRecsFound = ""
    
    StripArticle pCriteria, strArticle, strNet
    pCriteria = strNet
    oSearchEngine.prisearch
    oSearchEngine.SetupSQLwoCriteria True, pSearchType, Me.chkAntiquarianOnly, CLng(txtmaxnum), "B"
    enSource = enLocalDB
    If pSearchType = enSearchByCatalogue Then
        oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchNormal Then
        oSearchEngine.SimpleSearch pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchBF Then
        enSource = enBF
        oSearchEngine.BFSearchEx pCriteria, lngRecsFound, CLng(txtmaxnum), lngResult
    ElseIf pSearchType = enSearchBIC Then
        enSource = enLocalDB
        oSearchEngine.SearchBIC pCriteria, lngRecsFound
    Else
        oSearchEngine.AdvancedSearch lngRecsFound, pCriteria
    End If
    oSearchEngine.Execute (txtmaxnum)
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
        Case enBF
            XBF.ReDim 1, 1, 1, 12
            XBF(1, 1) = "No records"
            GBF.ReBind
        End Select
    End If
    If CLng(txtRecsFound) = CLng(txtmaxnum) Then
        MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
        Me.GN.Refresh
    End If
    Exit Sub
    
'Errh:
'    oError.SetError Err, Error, Now, "frmBrowseProductAQ:Search", "", ""
'    Exit Sub
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Search(pSearchType,pCriteria)", Array(pSearchType, pCriteria)
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Long

    Select Case enSource
    Case enLocalDB
        GBF.Visible = False
        GN.Visible = True
        XN.Clear
        XBF.Clear
        GBF.ReBind
        XN.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
'                If colList.Item(i).Code = "" Then
'                    XBF.Value(i, 10) = i & "h"
'                Else
'                    XBF.Value(i, 10) = colList.Item(i).Code & "k"
'                End If
                XN.Value(i, 1) = colList.Item(i).CodeF
                XN.Value(i, 2) = colList.Item(i).statusF & " " & colList.Item(i).Title
                XN.Value(i, 3) = colList.Item(i).Author
                XN.Value(i, 4) = colList.Item(i).Distributor
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
        
        
        
    Case enBF
        XN.Clear
        GN.ReBind
        XBF.Clear
        GBF.Visible = True
        GN.Visible = False
        XBF.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
'                If colList.Item(i).Code = "" Then
'                    XBF.Value(i, 10) = i & "h"
'                Else
'                    XBF.Value(i, 10) = colList.Item(i).Code & "k"
'                End If
                XBF.Value(i, 1) = colList.Item(i).code
                XBF.Value(i, 2) = colList.Item(i).Title
                XBF.Value(i, 3) = colList.Item(i).Author
                XBF.Value(i, 4) = IIf(colList.Item(i).DistributorByIdx(1) = "", "Pub by:" & colList.Item(i).Publisher, colList.Item(i).DistributorByIdx(1))
                XBF.Value(i, 5) = colList.Item(i).LocalPriceF
                XBF.Value(i, 6) = colList.Item(i).USPriceF
                XBF.Value(i, 7) = colList.Item(i).UKPriceF
                XBF.Value(i, 8) = colList.Item(i).DistributorByIdx(1)
                XBF.Value(i, 12) = colList.Item(i).code
        Next
        XBF.QuickSort 1, XBF.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
        GN.Array = XN
        Me.GN.ReBind
        GBF.Array = XBF
        Me.GBF.ReBind
    End Select
'Errh:
'    MsgBox Error
'    Exit Sub
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.LoadGrid"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSaveLayout_Click()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To GN.Columns.Count
        SaveSetting "CENTRAL", "SearchFormA", CStr(i), GN.Columns(i - 1).Width
    Next
    For i = 1 To GBF.Columns.Count
        SaveSetting "CENTRAL", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo errHandler
strTime = "Start:" & Now() & vbCrLf
    Screen.MousePointer = vbHourglass
    txtcritvalues = FNS(txtcritvalues)
    If UCase(right(txtcritvalues, 2)) = "+B" Or UCase(right(txtcritvalues, 2)) = "!!" Then
        Search enSearchBF, left(txtcritvalues, Len(txtcritvalues) - 2)
        GBF.SetFocus
    ElseIf InStr(txtcritvalues, "/") > 0 Then
        Search enSearchAdvanced, txtcritvalues
        GN.SetFocus
    Else
        Search enSearchNormal, txtcritvalues
        GN.SetFocus
    End If
    Screen.MousePointer = vbDefault
    
strTime = "Emd:" & Now() & vbCrLf
    Exit Sub
    
'errHandler:
'    oError.SetError Err, Error, Now, "frmBrowseProductAQ:cmdSearch", "", ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdconsearch_Click()
'On Error GoTo ERRHANDLER
'
''    If cbocrit = "Supplier" Or cbocrit = "Catalogue" Or cbocrit = "Category" Then
''        oSearchEngine.secsearch
''        oSearchEngine.SetupSQLwoCriteria True, enSearchNormal, Me.chkAntiquarianOnly
''        oSearchEngine.selectcriteria cbocrit, strthing
''        lblcrit.Caption = lblcrit.Caption & cbocrit & " " & "'" & Cbocritvalues & "'" & "\"
''    Else
'        oSearchEngine.secsearch
'        oSearchEngine.SetupSQLwoCriteria True, enSearchNormal, Me.chkAntiquarianOnly
'    '    oSearchEngine.selectcriteria cbocrit, txtcritvalues
'     '   lblcrit.Caption = lblcrit.Caption & cbocrit & " " & "'" & txtcritvalues & "'" & "\"
''    End If
'    Me.MousePointer = 11
'    oSearchEngine.Execute (txtmaxnum)
'    Me.MousePointer = 0
'    Set colList = Nothing
'    Set colList = oSearchEngine.getcols
'    Dim lngrows As Long
'    lngrows = oSearchEngine.Rows
'    txtRecsFound = lngrows
'    Dim i As Integer
'    Me.MousePointer = 11
'    For i = 1 To colList.Count
'        With lslist
'            .Key = colList.Item(i).pID
'            .Text = colList.Item(i).Code
'            .SubItems(1) = colList.Item(i).Title
'            .SubItems(2) = colList.Item(i).Author
'            .SubItems(3) = colList.Item(i).Publisher
'            .SubItems(4) = colList.Item(i).Stock
'        End With
'    Next
'    If colList.Count = 0 Then
'        Set lslist = lvwLines.ListItems.Add
'        lslist.Text = "No Records Found"
'    End If
'    Me.MousePointer = 0
'    If CLng(txtRecsFound) > CLng(txtmaxnum) Then
'        MsgBox "No. of records exceeds maximum, you must narrow down the search criteria.", , "Criteria too broad"
'    End If
'    Exit Sub
'ERRHANDLER:
'    If Err.Number = 3021 Then
'        'MsgBox "No records matching criteria found.", vbExclamation, "Records not Found"
'    Else
'        Err.Raise Err
'    End If
'
'    Set lslist = lvwLines.ListItems.Add
'    lslist.Text = "No Records"
'    'rsdata.Close
'     Exit Sub
'     Resume
'End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    XA.Clear
    XA.ReDim 1, 1, 1, 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Activate", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim i As Integer
    Set oSearchEngine = New z_SearchEngineC
    Set colList = New Collection
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
    
    For i = 1 To GBF.Columns.Count
        GBF.Columns(i - 1).Width = GetSetting("CENTRAL", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
    Next
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("CENTRAL", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
    XA.ReDim 1, 1, 1, 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Initialize", , EA_NORERAISE
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
'    tlCats.Load ltCatalogue
        GN.Columns(3).Caption = "Distributor"
    GBF.Columns(3).Caption = "Distributor"
    txtmaxnum = 500
    
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
    For i = 1 To GBF.Columns.Count
        GBF.Columns(i - 1).Width = GetSetting("CENTRAL", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
    Next
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("CENTRAL", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oSearchEngine = Nothing
    Set roProduct = Nothing
    Set colList = Nothing
    Set tlkeys = Nothing
   ' Set lslist = Nothing
   ' Set lllisttype = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub GBF_Click()
    On Error GoTo errHandler
Dim str As String
    str = FNS(XBF.Value(GBF.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText str
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GBF_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub GBF_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(GBF.Bookmark) Then Exit Sub
    str = FNS(XBF.Value(GBF.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText str
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GBF_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub GBF_DblClick()
    On Error GoTo errHandler
Dim oProd As a_Product
Dim str As String
    str = FNS(XBF.Value(GBF.Bookmark, 1))
    If str = "" Then Exit Sub
    If MsgBox("Do you want to create a record in the database from the Bookfind data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    Set oProd = Nothing
    Set oProd = New a_Product
    Screen.MousePointer = vbHourglass
    oProd.Load "", 0, str
    Screen.MousePointer = vbDefault
    MsgBox "Record added"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GBF_DblClick", , EA_NORERAISE
    HandleError
End Sub



Private Sub GBF_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        GBF_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GBF_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub GN_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText str
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText str
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub


Private Sub GN_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev
Dim lngprod As Long
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 11))
    If str = "" Then Exit Sub
    Set roProduct = New a_Product
    WaitMsg "Loading . . .", True, Me
    roProduct.Load str, 0, "", strTime
    If roProduct.pID = "" Then Exit Sub
    If roProduct.ProductType = "B" Then
        Set frm = New frmProductPrev
        frm.Component roProduct, strTime
        frm.Show
    Else
        Set frmNB = New frmProductNBPrev
        frmNB.Component roProduct, strTime
        frmNB.Show
    End If
    Set roProduct = Nothing
    WaitMsg "", False, Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_DblClick", , EA_NORERAISE
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
 '   Else
 '       XN.QuickSort XA.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
 '   End If
    
    GN.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
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
    ErrorIn "frmBrowseProductsAQ.GetRowType(ColIndex)", ColIndex
End Function

Private Sub GN_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    
    If KeyAscii = vbKeyReturn Then
        GN_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
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
    ErrorIn "frmBrowseProductsAQ.GN_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub

'Public Sub AddToTempList()
'    On Error GoTo errHandler
'Dim str As String
'    str = FNS(XN.Value(GN.Bookmark, 11))
'    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
'        If XA(XA.UpperBound(1), 1) > "" Then
'            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
'        End If
'        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
'        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
'        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
'        XA(XA.UpperBound(1), 4) = 1
'        XA(XA.UpperBound(1), 5) = 0
'        XA(XA.UpperBound(1), 6) = ""
'        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProductsAQ.AddToTempList"
'End Sub
'Public Sub PlaceCO()
'    On Error GoTo errHandler
'Dim frm As New frmPlaceCO
'Dim str As String
'    str = FNS(XN.Value(GN.Bookmark, 1))
'    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
'        If XA(XA.UpperBound(1), 1) > "" Then
'            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
'        End If
'        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
'        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
'        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
'        XA(XA.UpperBound(1), 4) = 1
'        XA(XA.UpperBound(1), 5) = 0
'        XA(XA.UpperBound(1), 6) = ""
'        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
'    End If
'    frm.Component XA, "ORDER"
'    frm.Show vbModal
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProductsAQ.PlaceCO"
'End Sub
'Public Sub PlaceOnReserve()
'    On Error GoTo errHandler
'Dim frm As New frmPlaceCO
'Dim str As String
'    str = FNS(XN.Value(GN.Bookmark, 11))
'    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
'        If XA(XA.UpperBound(1), 1) > "" Then
'            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
'        End If
'        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
'        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
'        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
'        XA(XA.UpperBound(1), 4) = 1
'        XA(XA.UpperBound(1), 5) = 0
'        XA(XA.UpperBound(1), 6) = ""
'        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
'    End If
'    frm.Component XA, "RESERVE"
'    frm.Show vbModal
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProductsAQ.PlaceOnReserve"
'End Sub
'Public Sub StartNewList()
'    On Error GoTo errHandler
'    XA.Clear
'    XA.ReDim 1, 1, 1, 7
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProductsAQ.StartNewList"
'End Sub


Private Sub Label2_Click()
    On Error GoTo errHandler
Dim str As String
    str = "To use search box . . ." & vbCrLf _
            & "Search on title . . . /Harry potter" & vbCrLf _
            & "   will yield all titles starting with 'Harry Potter'" & vbCrLf _
            & "Search on title . . . /*Harry potter" & vbCrLf _
            & "   will yield all titles containing 'Harry Potter'" & vbCrLf _
            & "Search on title . . . /*Harry * goblet" & vbCrLf _
            & "   will yield all titles containing 'Harry' and 'goblet' in that order" & vbCrLf & vbCrLf _
            & "Replacing '/' with '//' will search authors" & vbCrLf & vbCrLf _
            & "Replacing '/' with '///' will search publishers" & vbCrLf & vbCrLf _
            & "Adding '!!' at the end of the search string will search on Bookfind (if installed)" & vbCrLf
    MsgBox str, vbInformation, "Help"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Label2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtcritvalues_DblClick()
    On Error GoTo errHandler
    txtcritvalues = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.txtcritvalues_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtcritvalues_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.txtcritvalues_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
