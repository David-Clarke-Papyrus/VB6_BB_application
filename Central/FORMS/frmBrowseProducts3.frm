VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseProducts 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse books"
   ClientHeight    =   8520
   ClientLeft      =   240
   ClientTop       =   1020
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseProducts3.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11655
   Begin VB.CommandButton cmdGetFromSB 
      BackColor       =   &H00C4BCA4&
      Caption         =   "TEST "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7635
      Width           =   1830
   End
   Begin VB.CommandButton cmdPrintbarcode 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Barcode print"
      Height          =   480
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6210
      Width           =   1455
   End
   Begin VB.CommandButton cmdDebugOff 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Turn debug OFF"
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
      TabIndex        =   23
      Top             =   7740
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdDebugOn 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Turn debug ON"
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
      TabIndex        =   22
      Top             =   7440
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowseProducts3.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6210
      Width           =   1000
   End
   Begin VB.TextBox txtRecsFound 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   10620
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   885
      Width           =   675
   End
   Begin VB.TextBox txtmaxnum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   10620
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   390
      Width           =   690
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1560
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   2752
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483644
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Normal search"
      TabPicture(0)   =   "frmBrowseProducts3.frx":07CC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkCopies"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkAntiquarianOnly"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdsearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboProductType"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboSection"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Special searches"
      TabPicture(1)   =   "frmBrowseProducts3.frx":07E8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frCatalogue"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command2 
         BackColor       =   &H00F2E0D9&
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74550
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1140
         Width           =   405
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F2E0D9&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1140
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search by BIC codes (if captured)"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   5325
         TabIndex        =   28
         Top             =   375
         Width           =   3810
         Begin VB.CommandButton cmdBIC 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   1365
            Picture         =   "frmBrowseProducts3.frx":0804
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   300
            Width           =   1260
         End
      End
      Begin VB.Frame frCatalogue 
         Caption         =   "Search by catalogue"
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   555
         TabIndex        =   25
         Top             =   375
         Width           =   3735
         Begin VB.CommandButton cmdCAT 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   705
            Left            =   2235
            Picture         =   "frmBrowseProducts3.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   210
            Width           =   1260
         End
         Begin VB.ComboBox cboCat 
            ForeColor       =   &H00800000&
            Height          =   360
            ItemData        =   "frmBrowseProducts3.frx":0F18
            Left            =   165
            List            =   "frmBrowseProducts3.frx":0F1A
            TabIndex        =   26
            Top             =   390
            Width           =   1410
         End
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69465
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   990
         Width           =   2130
      End
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69480
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   810
         Left            =   -66975
         Picture         =   "frmBrowseProducts3.frx":0F1C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   510
         Width           =   1260
      End
      Begin VB.CheckBox chkAntiquarianOnly 
         Caption         =   "Antiquarian only"
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
         Height          =   315
         Left            =   -73875
         TabIndex        =   8
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1440
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
         ItemData        =   "frmBrowseProducts3.frx":12A6
         Left            =   -74160
         List            =   "frmBrowseProducts3.frx":12A8
         TabIndex        =   0
         Top             =   570
         Width           =   3180
      End
      Begin VB.CheckBox chkCopies 
         Appearance      =   0  'Flat
         Caption         =   "Copies on hand"
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
         Height          =   315
         Left            =   -72315
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2010
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Product type"
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
         Height          =   285
         Left            =   -70935
         TabIndex        =   19
         Top             =   555
         Width           =   1380
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
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
         Height          =   285
         Left            =   -70275
         TabIndex        =   14
         Top             =   1050
         Width           =   735
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
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
         Left            =   4350
         TabIndex        =   12
         Top             =   -300
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74160
         TabIndex        =   10
         Top             =   315
         Width           =   1455
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
         Left            =   -74580
         TabIndex        =   9
         Top             =   510
         Width           =   360
      End
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4455
      Left            =   105
      OleObjectBlob   =   "frmBrowseProducts3.frx":12AA
      TabIndex        =   3
      Top             =   1695
      Width           =   11340
   End
   Begin TrueOleDBGrid60.TDBGrid GBF 
      Height          =   4455
      Left            =   105
      OleObjectBlob   =   "frmBrowseProducts3.frx":6495
      TabIndex        =   6
      Top             =   1695
      Width           =   11265
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10320
      Picture         =   "frmBrowseProducts3.frx":B079
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6180
      Width           =   1000
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
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7110
      Visible         =   0   'False
      Width           =   1185
   End
   Begin TrueOleDBGrid60.TDBGrid GEX 
      Height          =   930
      Left            =   945
      OleObjectBlob   =   "frmBrowseProducts3.frx":B403
      TabIndex        =   21
      Top             =   4365
      Visible         =   0   'False
      Width           =   11250
   End
   Begin VB.Label lblHUBRESULT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   750
      Left            =   4740
      TabIndex        =   33
      Top             =   6225
      Width           =   2550
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   9975
      TabIndex        =   18
      Top             =   915
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      Caption         =   "Max"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   10140
      TabIndex        =   17
      Top             =   450
      Width           =   390
   End
End
Attribute VB_Name = "frmBrowseProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strthing As String
Dim xMLDoc As ujXML
Dim strSearchCombo As String
Dim tlkeys As z_TextList
Private oSearchEngine As z_SearchEngineC
Dim colList As Collection
Dim intShowCopies As Integer
'Dim rsdata As New ADOR.Recordset
Dim lslist As ListItem
Dim oProduct As a_Product
Dim enSource As enProductDataSource
Dim mnu As Menu
Dim XA As XArrayDB
Dim XBF As XArrayDB
Dim XN As XArrayDB
Dim strTime As String
Dim tlSuppliers As z_TextList
Dim tlCats As z_TextList
Dim bShiftDown As Boolean
Dim strPID As String
Dim bWithCopies As Boolean
Dim wdthCol(0 To 15) As Double
Dim OriginalwdthGN As Double
Dim flgLoading As Boolean
Dim BookmarkPointer As Long

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.GN, "SearchFormA"
    SaveLayout Me.GBF, "SearchFormB"
    SaveSetting "PBKS", Me.Name, "Formwidth", Me.Width
 SetColumnWidths
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetMenu()

    Forms(0).mnuSaveColumnWidths.Enabled = True
    
End Sub
Public Sub UnsetMenu()

    Forms(0).mnuSaveColumnWidths.Enabled = False
      
End Sub



Private Sub cboProductType_DblClick()
    cboProductType = ""
End Sub

Private Sub cboSearch_DblClick()
    cboSearch = ""
End Sub


Private Sub chkCopies_Click()
    On Error GoTo errHandler
    bWithCopies = (chkCopies = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.chkCopies_Click", , EA_NORERAISE
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
    Set frm = Nothing
    
    
    If strBICCode > "" Then
'        Me.Refresh
'        Screen.MousePointer = vbHourglass
'        Screen.MousePointer = vbHourglass
'        Me.Refresh
    '--------------
        oPC.OpenDBSHort
    '--------------
        Search enSearchBIC, strBICCode
    '--------------
        oPC.DisconnectDBShort
    '--------------
    End If
        Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdBIC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCAT_Click()
    On Error GoTo errHandler
    Me.txtmaxnum = "9999999"
    Screen.MousePointer = vbHourglass
        Search enSearchByCatalogue, cboCat
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdCAT_Click", , EA_NORERAISE
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
Dim strErrPos As String
Dim lngSectionID As Long
Dim lngProductTypeID As Long

strErrPos = "1"
        Screen.MousePointer = vbHourglass

    txtRecsFound = ""
    lngSectionID = 0
    lngProductTypeID = 0
strErrPos = "1a"
    If pSearchType <> enSearchBIC Then
        StripArticle pCriteria, strArticle, strNet
        pCriteria = strNet
    End If
    oSearchEngine.prisearch
    '--------------
    oPC.OpenDBSHort
    '--------------
strErrPos = "2"
    oSearchEngine.SetupSQLwoCriteria False, pSearchType, False, IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500), ""
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
   '     MsgBox "Skipping Products types etc"
        If pSection <> "<ALL>" Then
            lngSectionID = oPC.Configuration.Sections.Key(pSection)
        End If
        If pProductType <> "<ALL>" Then
            lngProductTypeID = oPC.Configuration.ProductTypes.Key(pProductType)
        End If
        oSearchEngine.AdvancedSearch lngRecsFound, pCriteria
    End If
strErrPos = "3"
    If lngRecsFound = -1 Then
            MsgBox "No records returned because the criteria are incorrectly expressed.", , "Criteria invalid"
    Else
        oSearchEngine.Execute IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
        Set colList = Nothing
        Set colList = oSearchEngine.getcols
        lngrows = oSearchEngine.Rows
        txtRecsFound = CStr(lngRecsFound)
strErrPos = "4"
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
strErrPos = "5"
        If lngRecsFound = IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500) Then
            MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
            Me.GN.Refresh
        End If
    End If
    '--------------
    oPC.DisconnectDBShort
    '--------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Search(pSearchType,pCriteria)", Array(pSearchType, pCriteria), , , "strErrPos", Array(strErrPos)
End Sub
Private Sub LoadGridEx()
Dim i As Long
Dim XEX As New XArrayDB
        XEX.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
                XEX.Value(i, 1) = colList.Item(i).CodeF
                XEX.Value(i, 2) = colList.Item(i).statusF & " " & colList.Item(i).Title
                XEX.Value(i, 3) = colList.Item(i).Author
                XEX.Value(i, 4) = colList.Item(i).Distributor
                XEX.Value(i, 5) = colList.Item(i).QtyOnHand
                XEX.Value(i, 6) = colList.Item(i).QtyonOrder
                XEX.Value(i, 7) = colList.Item(i).QtyOnBackorder
                XEX.Value(i, 8) = colList.Item(i).QtyTotalSold
                XEX.Value(i, 10) = colList.Item(i).LastDateDelivered
                XEX.Value(i, 9) = colList.Item(i).LocalPriceF
                XEX.Value(i, 11) = colList.Item(i).pID
                XEX.Value(i, 12) = colList.Item(i).code
        Next
        Set GEX.Array = XEX
        GEX.ReBind
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Long

    Screen.MousePointer = vbHourglass
    Select Case enSource
    Case enLocalDB
        GBF.Visible = False
        GN.Visible = True
        XN.Clear
        XBF.Clear
        GBF.ReBind
        XN.ReDim 1, colList.Count, 1, 15
        For i = 1 To colList.Count
                XN.Value(i, 1) = colList.Item(i).CodeF
                XN.Value(i, 2) = colList.Item(i).statusF & " " & colList.Item(i).Title
                XN.Value(i, 3) = colList.Item(i).Author
                XN.Value(i, 4) = colList.Item(i).Distributor
                XN.Value(i, 5) = colList.Item(i).saleslist
                XN.Value(i, 6) = colList.Item(i).Publisher
                XN.Value(i, 7) = colList.Item(i).PubDate
                XN.Value(i, 8) = colList.Item(i).QtyTotalSold
                XN.Value(i, 10) = colList.Item(i).LastDateDelivered
                XN.Value(i, 9) = colList.Item(i).LocalPriceF
                XN.Value(i, 11) = colList.Item(i).pID
                XN.Value(i, 12) = colList.Item(i).code
                XN.Value(i, 13) = colList.Item(i).EAN
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
                If oPC.BOOKFINDISBN13ENABLED Then
                    XBF.Value(i, 1) = colList.Item(i).EAN
                Else
                    XBF.Value(i, 1) = colList.Item(i).code
                End If
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
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.LoadGrid"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdDebugOff_Click()
'oSearchEngine.mbDebug = False
'End Sub
'
'Private Sub cmdDebugOn_Click()
'    oSearchEngine.mbDebug = True
'End Sub

Private Sub cmdGetFromSB_Click()
Dim ob As New z_SQL

    Screen.MousePointer = vbHourglass
    
    ob.RunProc "dbo.SENDBROKERMESSAGE", Array(Me.cboSearch), "TEST"
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdPrint_Click()
    ExportToXML
End Sub

'Private Sub cmdPrintbarcode_Click()
'    Dim ar As New arPrintBarcodeList
'
'    ar.Component XN
'    ar.Show vbModal
'    Set ar = Nothing
'
'End Sub
'
Private Sub cmdSearch_Click()
    On Error GoTo errHandler
    cboSearch.AddItem cboSearch, 0
    oSearchEngine.instock False
    
    Screen.MousePointer = vbHourglass
    
    cboSearch = FNS(cboSearch)
    If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" Then
        Search enSearchBF, Left(cboSearch, Len(cboSearch) - 2)
        mSetfocus GBF
    ElseIf InStr(cboSearch, "/") > 0 Or UCase(cboSection) <> "<ALL>" Or UCase(cboProductType) <> "<ALL>" Then
        Search enSearchAdvanced, cboSearch, cboSection, cboProductType
        mSetfocus GN
    Else
        Search enSearchNormal, cboSearch
        mSetfocus GN
    End If
    mSetfocus cboSearch
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Command1_Click()
Dim str As String
    'If oPC.InternetDialup = True Then Exit Sub
    Screen.MousePointer = vbHourglass
    If cboSearch.Text = "" Then Exit Sub
    If IsNumeric(Left(Me.cboSearch.Text, 9)) Then
        If IsNumeric(Left(Me.cboSearch.Text, 13)) Then
            OpenBrowser "http://books.google.com/books?isbn=" & Left(Me.cboSearch.Text, 13)
        Else
            OpenBrowser "http://books.google.com/books?isbn=" & Left(Me.cboSearch.Text, 10)
        End If
    Else
        str = Replace(FNS(cboSearch), "/", "")
        OpenBrowser "http://books.google.com/books?q=" & str
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
Dim str As String
Dim str2 As String
  '  If oPC.InternetDialup = True Then Exit Sub
    Screen.MousePointer = vbHourglass
    If cboSearch.Text = "" Then Exit Sub
    If IsNumeric(Left(Me.cboSearch.Text, 9)) Then
        If IsNumeric(Left(Me.cboSearch.Text, 13)) Then
            str = "http://www.amazon.co.uk/dp/XXX"
            str = Replace(str, "XXX", Left(Me.cboSearch.Text, 10))
            OpenBrowser str
        Else
            str = "http://www.amazon.co.uk/dp/XXX"
            str = Replace(str, "XXX", Left(Me.cboSearch.Text, 10))
            OpenBrowser str
        End If
    Else
        
        str = "http://www.amazon.co.uk/gp/search?search-alias=stripbooks&field-keywords=&author=&select-author=field-author-like&title=XXX&select-title=field-title&subject=&select-subject=field-subject&field-publisher=&field-isbn=&chooser-sort=rank%21%2Bsalesrank&node=&field-binding=&mysubmitbutton1.x=53&mysubmitbutton1.y=12"
        str2 = Replace(FNS(cboSearch), "/", "")
        str = Replace(str, "XXX", str2)
        OpenBrowser str
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    On Error GoTo errHandler
    SetMenu
    XA.Clear
    XA.ReDim 1, 1, 1, 7
    XBF.Clear
    XBF.ReDim 1, 1, 1, 12
   ' UnsetMenu
    SSTab1.Tab = 0
    cboSearch.SetFocus
    bWithCopies = False
    chkCopies = IIf(bWithCopies, 1, 0)
    Me.Command1.Enabled = True
    cmdGetFromSB.Visible = False
    lblHUBRESULT.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetColumnWidths()
Dim i As Integer
    For i = 1 To GBF.Columns.Count
        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
    Next
    GBF.Columns(GBF.Columns.Count - 1).Width = GBF.Columns(GBF.Columns.Count - 1).Width * 0.8
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
    GN.Columns(GN.Columns.Count - 1).Width = GN.Columns(GN.Columns.Count - 1).Width * 0.8
End Sub


Private Sub Form_Deactivate()
    UnsetMenu
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim i As Integer
    Set XN = New XArrayDB
    Set XBF = New XArrayDB
    Set XA = New XArrayDB
    Set oSearchEngine = New z_SearchEngineC
    Set colList = New Collection
'    If oPC.Configuration.AntiquarianYN Then
'        chkAntiquarianOnly.Visible = True
'        chkAntiquarianOnly = 1
'    Else
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
 '   End If
    SetColumnWidths
    Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
    XA.ReDim 1, 1, 1, 7
    XBF.ReDim 1, 1, 1, 12
    SSTab1.Tab = 0
    cboSearch.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    SetMenu
    flgLoading = True
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
    Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
    flgLoading = False
    OriginalwdthGN = 250
    For i = 1 To GN.Columns.Count - 1
        OriginalwdthGN = OriginalwdthGN + GN.Columns(i - 1).Width
        wdthCol(i - 1) = GN.Columns(i - 1).Width
    Next
    If Me.WindowState <> 2 Then
        Me.Top = 20
        Me.Left = 50
        Height = 7220
    End If
 '   Width = 11600
    
    Set tlCats = Nothing
    Set tlCats = New z_TextList
    
    If oPC.SupportsCatalogue Then
        tlCats.Load ltCatalogue
        LoadCombo cboCat, tlCats
    End If
    frCatalogue.Enabled = oPC.SupportsCatalogue
    
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    
    
'    If oPC.Configuration.AntiquarianYN Then
'        Me.GN.Columns(3).Caption = "Publisher"
'    Else
        GN.Columns(3).Caption = "Distributor"
'    End If
    GBF.Columns(3).Caption = "Distributor"
    txtmaxnum = 500
    
'    If oPC.Configuration.AntiquarianYN Then
'        chkAntiquarianOnly.Visible = True
'        chkAntiquarianOnly = 1
'    Else
        chkAntiquarianOnly = 0
        chkAntiquarianOnly.Visible = False
'    End If

    For i = 1 To GBF.Columns.Count
        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
    Next
    On Error Resume Next
    cboSection = "<ALL>"
    cboProductType = "<ALL>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    If flgLoading Then Exit Sub
    GN.Width = Me.Width - 380
    GBF.Width = GN.Width
    
    lngDiff = GN.Height
    GN.Height = Me.Height - (GN.Top + 1070)
    GBF.Height = GN.Height
    lngDiff = GN.Height - lngDiff
    
    cmdPrint.Top = cmdPrint.Top + lngDiff
    cmdClose.Top = cmdClose.Top + lngDiff
    cmdPrintbarcode.Top = cmdPrintbarcode.Top + lngDiff
    cmdClose.Left = GN.Width - 1000
    
    ResizeColumnsGN
End Sub
Sub ResizeColumnsGN()
Dim i As Integer
Dim newwidth As Long

    For i = 1 To GN.Columns.Count - 1
        GN.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GN.Width) / OriginalwdthGN) * 0.9)
    Next

End Sub
Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oSearchEngine = Nothing
    Set oProduct = Nothing
    Set colList = Nothing
    Set tlkeys = Nothing
    Set lslist = Nothing
    Set XN = Nothing
    Set XBF = Nothing
    Set XA = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
End Sub

Private Sub GBF_Click()
    On Error GoTo errHandler
Dim str As String
    str = FNS(XBF.Value(GBF.Bookmark, 1))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText Left(str, 13)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GBF_Click", , EA_NORERAISE
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
    Clipboard.SetText Left(str, 13)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GBF_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GBF_DblClick", , EA_NORERAISE
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
    ErrorIn "frmBrowseProducts.GBF_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub GN_Click()
    On Error GoTo errHandler
Dim str As String
    If XN Is Nothing Then Exit Sub
    On Error Resume Next
    If XN.UpperBound(1) = 0 Then Exit Sub
    If Err Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    str = IIf(FNS(XN.Value(GN.Bookmark, 13)) > "", FNS(XN.Value(GN.Bookmark, 13)), FNS(XN.Value(GN.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText Left(str, 13)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub GN_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyLeft) Then
        mSetfocus cboSearch
    End If
    If Shift = 1 Then
        bShiftDown = True
    End If

End Sub

Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If LastRow = "" Then Exit Sub
    If XN.UpperBound(1) = 0 Then Exit Sub
    On Error Resume Next
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
 '   str = FNS(XN.Value(GN.Bookmark, 12))
    str = IIf(FNS(XN.Value(GN.Bookmark, 13)) > "", FNS(XN.Value(GN.Bookmark, 13)), FNS(XN.Value(GN.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub
Public Property Get NextPID() As String
    If GN.Array Is Nothing Then
        NextPID = ""
        Exit Property
    End If
    If BookmarkPointer < GN.Array.UpperBound(1) Then
        BookmarkPointer = BookmarkPointer + 1
        NextPID = FNS(XN.Value(BookmarkPointer, 11))
    Else
        NextPID = ""
    End If
End Property
Public Property Get PrevPID() As String
    If GN.Array Is Nothing Then
        PrevPID = ""
        Exit Property
    End If
    If BookmarkPointer > 1 Then
        BookmarkPointer = BookmarkPointer - 1
        PrevPID = FNS(XN.Value(BookmarkPointer, 11))
    Else
        PrevPID = ""
    End If
        
End Property
Private Sub GN_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev   'non book form
Dim lngprod As Long
Dim strErrPos As String

    If XN.UpperBound(1) = 0 Then Exit Sub
    On Error Resume Next
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    On Error GoTo errHandler
    BookmarkPointer = GN.Bookmark
    strPID = FNS(XN.Value(GN.Bookmark, 11))
    If strPID = "" Then Exit Sub
    If bShiftDown Then
        ShowSalesPatterns
    Else
        Set oProduct = New a_Product
        Screen.MousePointer = vbHourglass
        oProduct.Load strPID, 0, "", strTime
        If oProduct.pID = "" Then Exit Sub
        If oProduct.ProductType <> "G" Then
'            If oPC.Configuration.AntiquarianYN Then
'                Set frmA = New frmProductPrevAQ
'                frmA.Component oProduct
'                frmA.Show
'            Else
                Set frm = New frmProductPrev
                frm.Component oProduct, strTime
                frm.Show
 '           End If
        Else
            Set frmNB = New frmProductNBPrev
            frmNB.Component oProduct, strTime
            frmNB.Show
        End If
    End If
    Set oProduct = Nothing
    Screen.MousePointer = vbDefault
    bShiftDown = False
    Exit Sub
errHandler:
    ErrPreserve
    bShiftDown = False
    If Err = 10005 Then Resume Next  'assume that this is the elusive vbcsExceptionFilter error that seems both harmless and untraceable

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_DblClick", , EA_NORERAISE, , "Error position", Array(strErrPos)
    HandleError
End Sub
Public Sub ShowSalesPatterns()
Dim frmSales As frmSalesCH
    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Set oProduct = New a_Product
    strPID = FNS(XN.Value(GN.Bookmark, 11))
    If strPID = "" Then Exit Sub

    oProduct.Load strPID, 0
    If oProduct.pID = "" Then Exit Sub
    Set frmSales = New frmSalesCH
    frmSales.Component oProduct
    frmSales.Show
    Screen.MousePointer = vbDefault
    Set frmSales = Nothing
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
    ErrorIn "frmBrowseProducts.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
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
    ErrorIn "frmBrowseProducts.GetRowType(ColIndex)", ColIndex
End Function

Private Sub GN_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    
    If KeyAscii = vbKeyReturn Then
        GN_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub GN_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuFindForm   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Public Sub AddToTempList()
    On Error GoTo errHandler
Dim str As String
    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
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
    ErrorIn "frmBrowseProducts.AddToTempList"
End Sub
'Public Sub PlaceCO()
'    On Error GoTo ErrHandler
'Dim frm As New frmPlaceCO
'Dim str As String
'    If GN = "" Or GN = "No records" Then Exit Sub
'    If GN.Bookmark = 0 Then Exit Sub
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
'    StartNewList
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProducts.PlaceCO"
'End Sub
'Public Sub PlaceOnReserve()
'    On Error GoTo ErrHandler
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
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProducts.PlaceOnReserve"
'End Sub
'Public Sub StartNewList()
'    On Error GoTo ErrHandler
'    XA.Clear
'    XA.ReDim 1, 1, 1, 7
'    Exit Sub
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProducts.StartNewList"
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
    ErrorIn "frmBrowseProducts.Label2_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtcritvalues_DblClick()
'    On Error GoTo errHandler
'    txtcritvalues = ""
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProducts.txtcritvalues_DblClick", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub txtcritvalues_KeyPress(KeyAscii As Integer)
'    On Error GoTo errHandler
'    If KeyAscii = vbKeyReturn Then
'        cmdSearch_Click
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseProducts.txtcritvalues_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub cboSearch_DblClick()
'    cboSearch = ""
'End Sub

Private Sub cboSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        If GN.Visible = True Then
            mSetfocus GN
        Else
            mSetfocus GBF
        End If
    End If
    Exit Sub
End Sub


Private Sub Label26_dblClick()
    cboSection = "<All>"
End Sub

Private Sub Label3_DblClick()
    cboSearch = ""
End Sub

Private Sub Label4_DblClick()
    cboProductType = "<All>"
End Sub

Private Sub txtmaxnum_Validate(Cancel As Boolean)
    If Not IsNumeric(txtmaxnum) Then
        MsgBox "You must enter a number here, representing the maximum number of records you want to get back, a suggested value is 500.", , "Invalid value"
        Cancel = True
    End If
End Sub

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
Dim objXSL As New MSXML2.DOMDocument30
Dim opXMLDOC As New MSXML2.DOMDocument30
Dim objXMLDOC  As New MSXML2.DOMDocument30
Dim strExecutable As String

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "SSI_1"
        .chCreate "SSI"
            .elText = "Selected stock items at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = "Code"
            .elCreateSibling "Col_2"
                .elText = "Title"
            .elCreateSibling "Col_3"
                .elText = "Author"
            .elCreateSibling "Col_4"
                .elText = "Publisher"
            .elCreateSibling "Col_5"
                .elText = "Price"
            .elCreateSibling "Col_6"
                .elText = "Qty"
                .navUP
        
        
        For i = 1 To colList.Count
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = colList.Item(i).CodeF
            .elCreateSibling "Col_2"
                .elText = colList.Item(i).statusF & " " & colList.Item(i).Title
            .elCreateSibling "Col_3"
                .elText = colList.Item(i).Author
            .elCreateSibling "Col_4"
                .elText = colList.Item(i).Publisher
            .elCreateSibling "Col_5"
                .elText = colList.Item(i).LocalPriceF
            .elCreateSibling "Col_6"
                .elText = colList.Item(i).QtyOnHand
                .navUP
        Next i

        
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\SSI" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .HTML FILE
    objXSL.async = False
    objXSL.validateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

    strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1.RTF"
    If fs.FileExists(strFilename) Then
        fs.DeleteFile strFilename, True
    End If
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile

    strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
    Shell strExecutable, vbNormalFocus
    
    Exit Function
errHandler:
    ErrPreserve
    If Err = 70 Then
        MsgBox "Cannot delete a temporary file (SSI_1.RTF). It is probably open in an editing program (e.g. Microsoft WORD)." & vbCrLf & "Please save the file if you need to and close the document, then try this print operation again.", vbInformation, "Can't do this"
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowsePOs.ExportToXML"
End Function

'Public Sub mnuSetPT()
'Dim IDs As String
'Dim frm As New frmSetProductType
'Dim i As Integer
'
'    ReDim strTitle(GN.SelBookmarks.Count)
'    IDs = ""
'    For i = 0 To GN.SelBookmarks.Count - 1
'        IDs = IDs & ",'" & XN(GN.SelBookmarks(i), 11) & "'"
'    Next i
'    If Left(IDs, 1) = "," Then
'        IDs = Right(IDs, Len(IDs) - 1)
'    End If
'    If IDs > "" Then
'        frm.Component IDs
'        frm.Show vbModal
'    Else
'        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
'    End If
'    Unload frm
'End Sub
Public Sub SetForWebExport()
Dim cnt As Integer
    cnt = 0
    ReDim strTitle(GN.SelBookmarks.Count)
    Dim i As Integer
    For i = 0 To GN.SelBookmarks.Count - 1
        MarkProductForWebExport XN(GN.SelBookmarks(i), 11)
        cnt = cnt + 1
    Next i
    If cnt > 0 Then
        MsgBox "Records have been marked for Web export", vbInformation, "Status"
    Else
        MsgBox "There are no rows selected. Click on the left margin to select rows before choosing option to mark for web export.", vbInformation, "Status"
    End If

End Sub
'Public Sub mnuSetSection()
'Dim IDs As String
'Dim frm As New frmSetSection
'Dim i As Integer
'
'    ReDim strTitle(GN.SelBookmarks.Count)
'    IDs = ""
'    For i = 0 To GN.SelBookmarks.Count - 1
'        IDs = IDs & ",'" & XN(GN.SelBookmarks(i), 11) & "'"
'    Next i
'    If Left(IDs, 1) = "," Then
'        IDs = Right(IDs, Len(IDs) - 1)
'    End If
'    If IDs > "" Then
'        frm.Component IDs
'        frm.Show vbModal
'    Else
'        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
'    End If
'    Unload frm
'End Sub

Public Sub mnuTouchRecord()
Dim cnt As Integer
    cnt = 0
    ReDim strTitle(GN.SelBookmarks.Count)
    Dim i As Integer
    For i = 0 To GN.SelBookmarks.Count - 1
        TouchRecord XN(GN.SelBookmarks(i), 11)
        cnt = cnt + 1
    Next i
    If cnt > 0 Then
        MsgBox "P.O.S. computers have been updated", vbInformation, "Status"
    Else
        MsgBox "There are no rows selected. Click on the left margin to select rows before choosing option to send to P.O.S. computers.", vbInformation, "Status"
    End If
End Sub
Private Sub TouchRecord(pPID As String)
Dim oSQL As New z_SQL

    oSQL.RunSQL "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_LoyaltyRATE," _
            & "PRU_PTID,PRU_SECID) " _
            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
            & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID " _
            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID " _
            & " WHERE P_ID = '" & pPID & "'"

End Sub

Private Sub MarkProductForWebExport(pPID As String)
Dim oSQL As New z_SQL

    oSQL.MarkForWebExport pPID

End Sub

