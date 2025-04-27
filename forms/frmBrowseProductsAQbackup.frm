VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBrowseProductsAQbackup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse stock"
   ClientHeight    =   6150
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   11490
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "By catalogue"
      ForeColor       =   &H8000000D&
      Height          =   1650
      Left            =   9435
      TabIndex        =   11
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdCAT 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Search"
         Height          =   510
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1035
         Width           =   1695
      End
      Begin VB.ComboBox cboCat 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmBrowseProductsAQbackup.frx":0000
         Left            =   255
         List            =   "frmBrowseProductsAQbackup.frx":0002
         TabIndex        =   5
         Text            =   "cboCat"
         Top             =   525
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1635
      Left            =   90
      TabIndex        =   8
      Top             =   135
      Width           =   9135
      Begin VB.CheckBox chkAntiquarianOnly 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Antiquarian only"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   5910
         TabIndex        =   19
         Top             =   1095
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CommandButton cmdBIC 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&BIC"
         Height          =   570
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2295
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1080
         Width           =   915
      End
      Begin VB.TextBox txtmaxnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   555
         TabIndex        =   13
         Top             =   1095
         Width           =   900
      End
      Begin VB.CheckBox chkCopies 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copies on hand"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   3600
         TabIndex        =   12
         Top             =   1095
         Width           =   2010
      End
      Begin VB.ComboBox Cbocritvalues 
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmBrowseProductsAQbackup.frx":0004
         Left            =   3210
         List            =   "frmBrowseProductsAQbackup.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   495
         Width           =   1740
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Search"
         Default         =   -1  'True
         Height          =   570
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdconsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Continue Search"
         Height          =   570
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1230
      End
      Begin VB.ComboBox cbocrit 
         ForeColor       =   &H00800000&
         Height          =   360
         ItemData        =   "frmBrowseProductsAQbackup.frx":0008
         Left            =   180
         List            =   "frmBrowseProductsAQbackup.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   495
         Width           =   2895
      End
      Begin VB.TextBox txtcritvalues 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   3210
         TabIndex        =   1
         Top             =   480
         Width           =   1740
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Found"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1635
         TabIndex        =   16
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Max"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   45
         TabIndex        =   15
         Top             =   1095
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Search by"
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   195
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "search for"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3225
         TabIndex        =   9
         Top             =   195
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstest 
      Height          =   3975
      Left            =   75
      TabIndex        =   7
      Top             =   1980
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7011
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2382
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   8468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Author"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Publisher"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Stock"
         Object.Width           =   1501
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "booki"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblcrit 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   7140
      TabIndex        =   17
      Top             =   1245
      Width           =   4200
   End
End
Attribute VB_Name = "frmBrowseProductsAQbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strthing As String
Dim tlkeys As z_TextList
Dim sesearch As z_SearchEngine
Dim colList As Collection
Dim intvalstock As Integer
Dim intShowCopies As Integer
'Dim rsdata As New ADOR.Recordset
Dim lslist As ListItem
Dim roProduct As a_Product
Dim tlCategories As z_TextList
Dim tlCatalogues As z_TextList
Dim tlSuppliers As z_TextList

Private Sub cbocrit_Click()
    Select Case cbocrit
        Case "Category"
            txtcritvalues.Visible = False
            Cbocritvalues.Visible = True
            Cbocritvalues.Clear
            LoadCombo Me.Cbocritvalues, tlCategories
        Case "Catalogue"
            txtcritvalues.Visible = False
            Cbocritvalues.Visible = True
            Cbocritvalues.Clear
            LoadCombo Cbocritvalues, tlCatalogues
        Case "Supplier"
            txtcritvalues.Visible = False
            Cbocritvalues.Visible = True
            Cbocritvalues.Clear
            LoadCombo Cbocritvalues, tlSuppliers
        Case Else
            Cbocritvalues.Visible = False
            txtcritvalues.Visible = True
    End Select
End Sub


Private Sub chkCopies_Click()
    sesearch.instock chkCopies
End Sub

Private Sub cmdBIC_Click()
Dim frm As frmBICTree
Dim strBICCode As String
    Me.cbocrit = "9-BIC"
    Set frm = New frmBICTree
    frm.Show vbModal
    strBICCode = frm.SelectedCode
    Unload frm
    Search enSearchNormal, strBICCode
End Sub

Private Sub cmdCAT_Click()
    Search enSearchByCatalogue, cboCat
'    lblcrit.Caption = ""
'    lstest.ListItems.Clear
'    txtnum = ""
'    cmdconsearch.Enabled = True
'    intvalstock = 0
'        sesearch.prisearch
'        sesearch.lookforobject True, enSearchByCatalogue
'        sesearch.selectcriteria "Catalogue", Me.cboCat
'        lblcrit.Caption = lblcrit.Caption & "catalogue " & "'" & Cbocritvalues & "'" & "\"
'    Me.MousePointer = 11
'    sesearch.Execute (txtmaxnum)
'    Me.MousePointer = 0
'    Set colList = Nothing
'    Set colList = sesearch.getcols
'    Dim lngrows As Long
'    lngrows = sesearch.Rows
'    txtnum = lngrows
'    Dim i As Long
'    lstest.ListItems.Clear
'    Me.MousePointer = 11
'    For i = 1 To colList.Count
'        Set lslist = lstest.ListItems.Add
'        With lslist
'            .Key = Format$(colList.Item(i).pID)
'            .Text = colList.Item(i).ISBN
'            .SubItems(1) = colList.Item(i).Title
'            .SubItems(2) = colList.Item(i).Author
'            .SubItems(3) = colList.Item(i).Publisher
'            .SubItems(4) = colList.Item(i).Stock
'        End With
'    Next
'    If colList.Count = 0 Then
'        Set lslist = lstest.ListItems.Add
'        lslist.Text = "No Records Found"
'    End If
'    Me.MousePointer = 0
'    If CLng(txtnum) > CLng(txtmaxnum) Then
'        MsgBox "No. of records exceeds maximum, you must narrow down the search criteria.", , "Criteria to broad"
'    End If
    Exit Sub
    
ERRHANDLER:
    oError.SetError Err, Error, Now, "frmBrowseProductAQ:cmdCAT", "", ""
End Sub

Private Sub Search(pSearchType As enSearchType, pCriteria As String)
On Error GoTo ERRHANDLER
    
    lblcrit.Caption = ""
    lstest.ListItems.Clear
    txtnum = ""
    cmdconsearch.Enabled = True
    intvalstock = 0
    
    sesearch.prisearch
    sesearch.lookforobject True, pSearchType, Me.chkAntiquarianOnly
    If pSearchType = enSearchByCatalogue Then
        sesearch.selectcriteria "Catalogue", pCriteria
    Else
        sesearch.selectcriteria cbocrit, pCriteria
    End If
    lblcrit.Caption = lblcrit.Caption & cbocrit & " " & "'" & pCriteria & "'" & "\"
    
    sesearch.Execute (txtmaxnum)
    Set colList = Nothing
    Set colList = sesearch.getcols
    Dim lngrows As Long
    lngrows = sesearch.Rows
    txtnum = lngrows
    Dim i As Long
    lstest.ListItems.Clear
    For i = 1 To colList.Count
        Set lslist = lstest.ListItems.Add
        With lslist
            .Key = colList.Item(i).pID
            .Text = colList.Item(i).Code
            .SubItems(1) = colList.Item(i).Title
            .SubItems(2) = colList.Item(i).Author
            .SubItems(3) = colList.Item(i).Publisher
            .SubItems(4) = colList.Item(i).Stock
        End With
    Next
    If colList.Count = 0 Then
        Set lslist = lstest.ListItems.Add
        lslist.Text = "No Records Found"
    End If
    If CLng(txtnum) > CLng(txtmaxnum) Then
        MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
    End If
    Exit Sub
    
ERRHANDLER:
    oError.SetError Err, Error, Now, "frmBrowseProductAQ:Search", "", ""
    Exit Sub
    Resume
End Sub
Private Sub cmdsearch_click()
On Error GoTo ERRHANDLER
    
    Screen.MousePointer = vbHourglass
    If cbocrit = "Supplier" Or cbocrit = "Category" Then
        Search enSearchNormal, Cbocritvalues
    Else
        Search enSearchNormal, txtcritvalues
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ERRHANDLER:
    oError.SetError Err, Error, Now, "frmBrowseProductAQ:cmdSearch", "", ""
End Sub

Private Sub cmdconsearch_Click()
On Error GoTo ERRHANDLER
    
    intvalstock = 0
    If cbocrit = "Supplier" Or cbocrit = "Catalogue" Or cbocrit = "Category" Then
        sesearch.secsearch
        sesearch.lookforobject True, enSearchNormal, Me.chkAntiquarianOnly
        sesearch.selectcriteria cbocrit, strthing
                lblcrit.Caption = lblcrit.Caption & cbocrit & " " & "'" & Cbocritvalues & "'" & "\"
    Else
        sesearch.secsearch
        sesearch.lookforobject True, enSearchNormal, Me.chkAntiquarianOnly
        sesearch.selectcriteria cbocrit, txtcritvalues
                lblcrit.Caption = lblcrit.Caption & cbocrit & " " & "'" & txtcritvalues & "'" & "\"
    End If
    Me.MousePointer = 11
    sesearch.Execute (txtmaxnum)
    Me.MousePointer = 0
    Set colList = Nothing
    Set colList = sesearch.getcols
    Dim lngrows As Long
    lngrows = sesearch.Rows
    txtnum = lngrows
    Dim i As Integer
    lstest.ListItems.Clear
    Me.MousePointer = 11
    For i = 1 To colList.Count
        Set lslist = lstest.ListItems.Add
        With lslist
            .Key = colList.Item(i).pID
            .Text = colList.Item(i).Code
            .SubItems(1) = colList.Item(i).Title
            .SubItems(2) = colList.Item(i).Author
            .SubItems(3) = colList.Item(i).Publisher
            .SubItems(4) = colList.Item(i).Stock
        End With
    Next
    If colList.Count = 0 Then
        Set lslist = lstest.ListItems.Add
        lslist.Text = "No Records Found"
    End If
    Me.MousePointer = 0
    If CLng(txtnum) > CLng(txtmaxnum) Then
        MsgBox "No. of records exceeds maximum, you must narrow down the search criteria.", , "Criteria too broad"
    End If
    Exit Sub
ERRHANDLER:
    If Err.Number = 3021 Then
        'MsgBox "No records matching criteria found.", vbExclamation, "Records not Found"
    Else
        Err.Raise Err
    End If
        
    Set lslist = lstest.ListItems.Add
    lslist.Text = "No Records"
    'rsdata.Close
     Exit Sub
     Resume
End Sub

Private Sub Form_Initialize()
    Set sesearch = New z_SearchEngine
    Set colList = New Collection
    If oPC.Configuration.AntiquarianYN Then
        chkCopies = 0
        chkCopies.Visible = True
        chkAntiquarianOnly.Visible = True
        chkAntiquarianOnly = 1
    Else
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
        chkCopies = 0
        chkCopies.Visible = False
    End If
End Sub

Private Sub Form_Load()

'    Dim oADODBConn As New PapyConn
'
'    Set oADODBConn = New PapyConn
'
'    oADODBConn.Username = Trim$("myadmin")
'    oADODBConn.Password = Trim$("car")
'    oADODBConn.Database = "Papyrus"
'
'    oADODBConn.Connect
    
    Me.top = 20
    Me.left = 50
    Height = 6525
    Set tlCategories = New z_TextList
    tlCategories.Load ltCategory
    Set tlCatalogues = New z_TextList
    tlCatalogues.Load ltCatalogue
    Set tlSuppliers = New z_TextList
    tlSuppliers.Load ltSupplier, ""
    Set colList = sesearch.addcols
    LoadCombocol cbocrit, colList
    LoadCombo cboCat, tlCatalogues

    txtmaxnum = 500
    
    cmdconsearch.Enabled = False
 '   cmdnewsearch.Enabled = False

End Sub

Private Sub Form_Terminate()
    Set sesearch = Nothing
    Set roProduct = Nothing
    Set colList = Nothing
    Set tlkeys = Nothing
    Set lslist = Nothing
   ' Set lllisttype = Nothing
End Sub

Private Sub lstest_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lstest_Click()
    Clipboard.SetText lstest.SelectedItem.Text
End Sub

Private Sub lstest_DblClick()
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim lngprod As Long
    
    Set roProduct = New a_Product
    roProduct.Load lstest.SelectedItem.Key, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.Component roProduct
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component roProduct
        frm.Show
    End If
    Set roProduct = Nothing
End Sub
