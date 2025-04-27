VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseProductsAQ 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse books"
   ClientHeight    =   6825
   ClientLeft      =   240
   ClientTop       =   1020
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
   Icon            =   "frmBrowseProductsAQ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11490
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4455
      Left            =   60
      OleObjectBlob   =   "frmBrowseProductsAQ.frx":0442
      TabIndex        =   15
      Top             =   1575
      Width           =   11250
   End
   Begin TrueOleDBGrid60.TDBGrid GBF 
      Height          =   4455
      Left            =   60
      OleObjectBlob   =   "frmBrowseProductsAQ.frx":5169
      TabIndex        =   14
      Top             =   1575
      Width           =   11265
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10320
      Picture         =   "frmBrowseProductsAQ.frx":9D4D
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6090
      Width           =   1000
   End
   Begin VB.Frame frCatalogue 
      BackColor       =   &H00D3D3CB&
      Caption         =   "By catalogue"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1425
      Left            =   6855
      TabIndex        =   9
      Top             =   45
      Width           =   1935
      Begin VB.CommandButton cmdCAT 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Search"
         Height          =   405
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   840
         Width           =   1410
      End
      Begin VB.ComboBox cboCat 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmBrowseProductsAQ.frx":A0D7
         Left            =   255
         List            =   "frmBrowseProductsAQ.frx":A0D9
         TabIndex        =   2
         Text            =   "cboCat"
         Top             =   390
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   1335
      Left            =   45
      TabIndex        =   7
      Top             =   -30
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
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   405
         Left            =   4680
         TabIndex        =   16
         Top             =   435
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Found"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5220
         TabIndex        =   10
         Top             =   270
         Width           =   390
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   195
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmBrowseProductsAQ"
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
'Dim rsdata As New ADOR.Recordset
Dim lslist As ListItem
Dim roProduct As a_Product
Dim enSource As enProductDataSource
Dim mnu As Menu
Dim XA As XArrayDB
Dim XBF As XArrayDB
Dim XN As XArrayDB
Dim strTime As String
Dim tlSuppliers As z_TextList
Dim tlCats As z_TextList
Dim bShiftDown As Boolean
Dim flgLoading As Boolean

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
    If strBICCode > "" Then
        Screen.MousePointer = vbHourglass
        Me.Refresh
        DoEvents
        search enSearchBIC, strBICCode
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdBIC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCAT_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    
    Me.txtmaxnum = "9999999"
    '--------------
    oPC.OpenDBSHort
    '--------------
    search enSearchByCatalogue, cboCat
    '--------------
    oPC.DisconnectDBShort
    '--------------
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdCAT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub search(pSearchType As enSearchType, pCriteria As String)
    On Error GoTo errHandler
Dim strParsedCriteria As String
Dim lngRecsFound As Long
Dim lngResult As Long
Dim lngrows As Long
Dim strArticle As String
Dim strNet As String
Dim strErrPos As String
strErrPos = "1"
    txtRecsFound = ""
    If pSearchType <> enSearchBIC Then
        StripArticle pCriteria, strArticle, strNet
        pCriteria = strNet
    End If
strErrPos = "2"
    oSearchEngine.prisearch
    oSearchEngine.SetupSQLwoCriteria True, chkCopies, pSearchType, chkAntiquarianOnly, IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500), "B"
    If pSearchType = enSearchByCatalogue Then
        enSource = enLocalDB
        oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchNormal Then
        enSource = enLocalDB
        oSearchEngine.SimpleSearch pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchBF Then
        enSource = enBF
        oSearchEngine.BFSearchEx pCriteria, lngRecsFound, CLng(txtmaxnum), lngResult
    ElseIf pSearchType = enSearchBIC Then
        enSource = enLocalDB
        oSearchEngine.SearchBIC pCriteria, lngRecsFound
    Else
        enSource = enLocalDB
        oSearchEngine.AdvancedSearch lngRecsFound, pCriteria
    End If
    If lngRecsFound = -1 Then
            MsgBox "No records returned because the criteria are incorrectly expressed.", , "Criteria invalid"
    Else
    oSearchEngine.Execute IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
    Set colList = Nothing
    Set colList = oSearchEngine.getcols
    lngrows = oSearchEngine.Rows
    txtRecsFound = CStr(lngRecsFound)
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
    If lngRecsFound = IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500) Then
        MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
        Me.GN.Refresh
    End If
    End If
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
                XN.Value(i, 2) = colList.Item(i).StatusF & " " & colList.Item(i).Title
                XN.Value(i, 3) = colList.Item(i).Author
                XN.Value(i, 4) = colList.Item(i).Distributor
                XN.Value(i, 5) = colList.Item(i).SerialF
                XN.Value(i, 6) = colList.Item(i).CopyPriceF
                XN.Value(i, 7) = colList.Item(i).PurchaseDateF
                XN.Value(i, 8) = colList.Item(i).CopiesSold
                XN.Value(i, 10) = colList.Item(i).CopiesSold
                XN.Value(i, 9) = colList.Item(i).LocalPriceF
                XN.Value(i, 11) = colList.Item(i).PID
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

Private Sub cmdclose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.GN, "SearchAQ_A", Me.Height, Me.Width
    SaveLayout Me.GBF, "SearchAQ_B"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    txtcritvalues = FNS(txtcritvalues)
    '--------------
    oPC.OpenDBSHort
    '--------------
    chkAntiquarianOnly = 1
    If UCase(Right(txtcritvalues, 2)) = "+B" Or UCase(Right(txtcritvalues, 2)) = "!!" Then
        search enSearchBF, Left(txtcritvalues, Len(txtcritvalues) - 2)
        mSetfocus GBF
    ElseIf InStr(txtcritvalues, "/") > 0 Then
        search enSearchAdvanced, txtcritvalues
        mSetfocus GN
    Else
        search enSearchNormal, txtcritvalues
        mSetfocus GN
    End If
    Screen.MousePointer = vbDefault
    
    '--------------
    oPC.DisconnectDBShort
    '--------------
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
    XBF.Clear
    XBF.ReDim 1, 1, 1, 12
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Activate", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Deactivate()
    UnsetMenu

End Sub
Private Sub SetMenu()
    On Error GoTo errHandler

    Forms(0).mnuSaveColumnWidths.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.SetMenu"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim i As Integer
    Set XN = New XArrayDB
    Set XBF = New XArrayDB
    Set XA = New XArrayDB
    Set oSearchEngine = New z_SearchEngineB
    Set colList = New Collection
    If oPC.Configuration.AntiquarianYN Then
        chkAntiquarianOnly.Visible = True
        chkAntiquarianOnly = 1
    Else
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
    End If
    
    For i = 1 To GBF.Columns.Count
        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
    Next
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
    XA.ReDim 1, 1, 1, 7
    XBF.ReDim 1, 1, 1, 12
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    flgLoading = True
    If Me.WindowState <> 2 Then
        Me.Top = 20
        Me.Left = 50
    End If
    Set tlSuppliers = New z_TextList
    tlSuppliers.Load ltSupplier, ""
    Set tlCats = Nothing
    Set tlCats = New z_TextList
    tlCats.Load ltCatalogue
    LoadCombo cboCat, tlCats
    If oPC.Configuration.AntiquarianYN Then
        Me.GN.Columns(3).Caption = "Publisher"
    Else
        GN.Columns(3).Caption = "Distributor"
    End If
    GBF.Columns(3).Caption = "Distributor"
    txtmaxnum = 500
    
    If oPC.Configuration.AntiquarianYN Then
        chkAntiquarianOnly.Visible = True
        chkAntiquarianOnly = 1
    Else
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
    End If
    SetFormSize Me
    SetGridLayout GN, "SearchAQ_A"
    SetGridLayout GBF, "SearchAQ_B"
    flgLoading = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    If flgLoading Then Exit Sub
    GN.Width = NonNegative_Lng(Me.Width - 380)
    lngDiff = GN.Height
    GN.Height = NonNegative_Lng(Me.Height - (GN.Top + 1260))
    lngDiff = GN.Height - lngDiff
    cmdClose.Top = cmdClose.Top + lngDiff
    cmdClose.Left = NonNegative_Lng(GN.Width - 1000)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oSearchEngine = Nothing
    Set roProduct = Nothing
    Set colList = Nothing
    Set tlkeys = Nothing
    Set lslist = Nothing
    Set XN = Nothing
    Set XBF = Nothing
    Set XA = Nothing
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
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GBF_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub GBF_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If LastRow = "" Then Exit Sub
    If XBF.UpperBound(1) = 0 Then Exit Sub
    If IsNull(GBF.Bookmark) Then Exit Sub
    str = FNS(XBF.Value(GBF.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
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
    If str = "No records" Then Exit Sub
    If str = "" Then Exit Sub
    If MsgBox("Do you want to create a record in the database from the Bookfind data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    Set oProd = Nothing
    Set oProd = New a_Product
    Screen.MousePointer = vbHourglass
    oProd.Load "", 0, str
    Screen.MousePointer = vbDefault
    MsgBox "Record added", , "Status"
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseProductsAQ: GBF_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseProductsAQ: GBF_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
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
    If XN.Count(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub GN_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
If Shift = 1 Then
    bShiftDown = True
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub GN_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
'If KeyCode = 16 Then
'    bShiftDown = False
'End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_KeyUp(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If LastRow = "" Then Exit Sub
    If XN.UpperBound(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub


Private Sub GN_DblClick()
    On Error GoTo errHandler
Dim frmA As frmProductPrevAQ
Dim frmAEdit As frmProductAQ
Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev   'non book form
Dim lngprod As Long
Dim str As String
Dim strErrPos As String

    If XN.UpperBound(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
   
    strErrPos = "Position 1"
    str = FNS(XN.Value(GN.Bookmark, 11))
    strErrPos = "Position 2"
    If str = "" Then Exit Sub
    WaitMsg "Loading . . .", True, Me
    Set roProduct = New a_Product
    roProduct.Load str, 0, "", strTime
    If roProduct.PID = "" Then Exit Sub
    
    If roProduct.ProductType = "B" Then
            If bShiftDown = True Then
                Set frmAEdit = New frmProductAQ
                frmAEdit.Component roProduct
                frmAEdit.Show
            Else
                Set frmA = New frmProductPrevAQ
                frmA.Component roProduct
                frmA.Show
            End If
    Else
        Set frmNB = New frmProductNBPrev
        frmNB.Component roProduct, strTime
        frmNB.Show
    End If
    
    
    strErrPos = "Position 12"
    Set roProduct = Nothing
    WaitMsg "", False, Me
    bShiftDown = False

    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseProductsAQ: GN_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseProductsAQ: GN_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub GN_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    
    If XN.UpperBound(1) = 0 Then Exit Sub
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


Private Sub GN_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
    If Button = 2 Then   ' Check if right mouse button was clicked.
        PopupMenu Forms(0).mnuFindForm
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.GN_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
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
    ErrorIn "frmBrowseProductsAQ.AddToTempList"
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
    ErrorIn "frmBrowseProductsAQ.PlaceCO"
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
    ErrorIn "frmBrowseProductsAQ.PlaceOnReserve"
End Sub
Public Sub StartNewList()
    On Error GoTo errHandler
    XA.Clear
    XA.ReDim 1, 1, 1, 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.StartNewList"
End Sub


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


Private Sub txtmaxnum_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not IsNumeric(txtmaxnum) Then
        MsgBox "You must enter a number here, representing the maximum number of records you want to get back, a suggested value is 500.", , "Invalid value"
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProductsAQ.txtmaxnum_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
