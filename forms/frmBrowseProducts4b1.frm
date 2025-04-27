VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseProducts 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse products"
   ClientHeight    =   8520
   ClientLeft      =   240
   ClientTop       =   1020
   ClientWidth     =   14265
   Icon            =   "frmBrowseProducts4b1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   14265
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   13380
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
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
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6195
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton cmdPrintbarcode 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Barcode print"
      Height          =   480
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowseProducts4b1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6195
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1560
      Left            =   120
      TabIndex        =   5
      Top             =   105
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   2752
      _Version        =   393216
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483644
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Normal search"
      TabPicture(0)   =   "frmBrowseProducts4b1.frx":07CC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkCopies"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboSearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdsearch"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboProductType"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboCategory"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtRecsFound"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtmaxnum"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkIncludeObsolete"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Special searches"
      TabPicture(1)   =   "frmBrowseProducts4b1.frx":07E8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "frCatalogue"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Google Books"
      TabPicture(2)   =   "frmBrowseProducts4b1.frx":0804
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkISBNOnly"
      Tab(2).Control(1)=   "cmdMoreGoo"
      Tab(2).Control(2)=   "cboGoogle"
      Tab(2).Control(3)=   "cmdGOO"
      Tab(2).Control(4)=   "Label6"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame1 
         Caption         =   "Course codes"
         ForeColor       =   &H8000000D&
         Height          =   1065
         Left            =   -67950
         TabIndex        =   38
         Top             =   360
         Width           =   4050
         Begin VB.CommandButton cmdCourseCodes 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   2940
            Picture         =   "frmBrowseProducts4b1.frx":0820
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   255
            Width           =   945
         End
         Begin VB.TextBox Text1 
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   195
            TabIndex        =   39
            Top             =   390
            Width           =   2580
         End
      End
      Begin VB.CheckBox chkIncludeObsolete 
         Appearance      =   0  'Flat
         Caption         =   "Include obsolete"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2580
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1590
      End
      Begin VB.CheckBox chkISBNOnly 
         Caption         =   "Show only books published since 1950"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -74745
         TabIndex        =   36
         Top             =   1110
         Width           =   3390
      End
      Begin VB.TextBox txtmaxnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   10335
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   690
      End
      Begin VB.TextBox txtRecsFound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   10335
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   975
         Width           =   675
      End
      Begin VB.CommandButton cmdMoreGoo 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&More"
         Height          =   435
         Left            =   -67050
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   630
         Width           =   1260
      End
      Begin VB.ComboBox cboGoogle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   390
         ItemData        =   "frmBrowseProducts4b1.frx":0BAA
         Left            =   -74865
         List            =   "frmBrowseProducts4b1.frx":0BAC
         TabIndex        =   28
         Top             =   660
         Width           =   6375
      End
      Begin VB.CommandButton cmdGOO 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&First"
         Height          =   435
         Left            =   -68370
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   630
         Width           =   1260
      End
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
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1140
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search by BIC codes (if captured)"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   -71415
         TabIndex        =   21
         Top             =   375
         Width           =   3300
         Begin VB.CommandButton cmdBIC 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   960
            Picture         =   "frmBrowseProducts4b1.frx":0BAE
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   255
            Width           =   1260
         End
      End
      Begin VB.Frame frCatalogue 
         Caption         =   "Search by catalogue"
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   -74790
         TabIndex        =   18
         Top             =   375
         Width           =   3240
         Begin VB.CommandButton cmdCAT 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   705
            Left            =   2070
            Picture         =   "frmBrowseProducts4b1.frx":0F38
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   210
            Width           =   975
         End
         Begin VB.ComboBox cboCat 
            ForeColor       =   &H00800000&
            Height          =   360
            ItemData        =   "frmBrowseProducts4b1.frx":12C2
            Left            =   165
            List            =   "frmBrowseProducts4b1.frx":12C4
            TabIndex        =   19
            Top             =   390
            Width           =   1410
         End
      End
      Begin VB.ComboBox cboCategory 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5535
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   990
         Width           =   2130
      End
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   810
         Left            =   8025
         Picture         =   "frmBrowseProducts4b1.frx":12C6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   510
         Width           =   1260
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         ItemData        =   "frmBrowseProducts4b1.frx":1650
         Left            =   840
         List            =   "frmBrowseProducts4b1.frx":1652
         TabIndex        =   0
         Top             =   570
         Width           =   3180
      End
      Begin VB.CheckBox chkCopies 
         Appearance      =   0  'Flat
         Caption         =   "Copies on hand"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   1035
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9840
         TabIndex        =   35
         Top             =   540
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
         Caption         =   "Found"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9690
         TabIndex        =   34
         Top             =   1005
         Width           =   555
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for... (Do not use wild cards (*) - this uses Google search engine)"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74850
         TabIndex        =   29
         Top             =   405
         Width           =   7725
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Product type"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   4065
         TabIndex        =   12
         Top             =   555
         Width           =   1380
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   4725
         TabIndex        =   11
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
         TabIndex        =   9
         Top             =   -300
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   840
         TabIndex        =   7
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
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   405
         Left            =   420
         TabIndex        =   6
         Top             =   510
         Width           =   360
      End
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4455
      Left            =   2820
      OleObjectBlob   =   "frmBrowseProducts4b1.frx":1654
      TabIndex        =   2
      Top             =   3720
      Width           =   11340
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10320
      Picture         =   "frmBrowseProducts4b1.frx":6957
      Style           =   1  'Graphical
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   7110
      Visible         =   0   'False
      Width           =   1185
   End
   Begin TrueOleDBGrid60.TDBGrid GEX 
      Height          =   930
      Left            =   2325
      OleObjectBlob   =   "frmBrowseProducts4b1.frx":6CE1
      TabIndex        =   14
      Top             =   3165
      Visible         =   0   'False
      Width           =   11250
   End
   Begin TrueOleDBGrid60.TDBGrid GOO 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmBrowseProducts4b1.frx":BFE1
      TabIndex        =   30
      Top             =   1725
      Width           =   11340
   End
   Begin VB.Label lblHUBRESULT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   750
      Left            =   4740
      TabIndex        =   26
      Top             =   6225
      Width           =   2550
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
Private oSearchEngine As z_SearchEngineB
Private oSearchEngineC As z_SearchEngineC
Dim colList As Collection
Dim rsResult As ADODB.Recordset
Dim lslist As ListItem
Dim oProduct As a_Product
Dim enSource As enProductDataSource
Dim mnu As Menu
Dim XA As XArrayDB
Dim XBF As XArrayDB
Dim XN As XArrayDB
Dim XGOO As XArrayDB
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
Dim strGOOXML As String
Dim mXML As ujXML
Dim GooIndex As Integer
Dim GooIndex2 As Integer
Dim GoogleButtonNo As Integer
Dim bISBNOnly As Boolean
Dim PrivateCnn As ADODB.Connection
Dim oColMap As New z_ColEx
Dim sCurrentGrid As String

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.GN, "SearchFormA"
    SaveLayout Me.GEX, "SearchFormB"
    SaveLayout Me.GOO, "SearchFormC"
    SaveSetting "PBKS", Me.Name, "Formwidth", Me.Width
 SetColumnWidths
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler

    Forms(0).mnuSaveColumnWidths.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.SetMenu"
End Sub
Public Sub UnsetMenu()
    On Error GoTo errHandler

    Forms(0).mnuSaveColumnWidths.Enabled = False
      
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.UnsetMenu"
End Sub

Private Sub cboProductType_DblClick()
    On Error GoTo errHandler
    cboProductType = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cboProductType_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboSearch_DblClick()
    On Error GoTo errHandler
    cboSearch = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cboSearch_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub Check1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Check1_Click", , EA_NORERAISE
    HandleError
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

Private Sub chkISBNOnly_Click()
    On Error GoTo errHandler
    bISBNOnly = (chkISBNOnly = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.chkISBNOnly_Click", , EA_NORERAISE
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
        search enSearchBIC, strBICCode
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
        search enSearchByCatalogue, cboCat
    Screen.MousePointer = vbDefault
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdCAT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo errHandler
    
    cboSearch.AddItem cboSearch, 0
    Screen.MousePointer = vbHourglass
    cboSearch = FNS(cboSearch)
    
    If cboSearch = "" And UCase(Me.cboCategory) = "<ALL>" And UCase(Me.cboProductType) = "<ALL>" Then
        Screen.MousePointer = vbDefault
        MsgBox "You must show what you are searching for.", vbOKOnly, "Enter a search request"
        Exit Sub
    End If
    
100       If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" Then
110      '     SetActiveGrid "GBF"
              DoEvents
'120           GBF.SearchArguments = Left(cboSearch, Len(cboSearch) - 2)
'130           GBF.search
              search enSearchBF, Left(cboSearch, Len(cboSearch) - 2)
140 '           mSetfocus GBF
150       Else
        SetActiveGrid "GN"
        search enSearchAdvanced, cboSearch, cboCategory, cboProductType
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

Private Sub search(pSearchType As enSearchType, pCriteria As String, Optional pSection As String, Optional pProductType As String)
    On Error GoTo errHandler
Dim strParsedCriteria As String
Dim lngRecsFound As Long
Dim lngResult As Long
Dim lngrows As Long
Dim strArticle As String
Dim strNet As String
Dim strErrPos As String
Dim lngSectionID As Long
Dim errRepeat As Integer
Dim lngProductTypeID As Long
Dim lngMaxRecs As Long
    errSysHandlerSet
    Set rsResult = Nothing
    Set rsResult = New ADODB.Recordset
    rsResult.CursorLocation = adUseClient
    rsResult.CursorType = adOpenStatic
    
    txtRecsFound = ""
    lngSectionID = 0
    lngProductTypeID = 0
    If pSearchType <> enSearchBIC Then
        StripArticle pCriteria, strArticle, strNet
        pCriteria = strNet
    End If
    '--------------
    oPC.OpenDBSHort
    '--------------
    
    If pSection <> "<ALL>" Then
        lngSectionID = oPC.Configuration.Sections.Key(pSection)
    End If
    If pProductType <> "<ALL>" Then
        lngProductTypeID = oPC.Configuration.ProductTypes.Key(pProductType)
    End If
        
    lngMaxRecs = IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
SearchAgain:
    If (Not Replace(pCriteria, "/", "") = "") Or lngSectionID > 0 Or lngProductTypeID > 0 Then
        If pSearchType = enSearchByCatalogue Then
            enSource = enLocalDB
            oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "", (chkIncludeObsolete = 1)
            oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
        ElseIf pSearchType = enSearchBIC Then
            enSource = enLocalDB
            oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "", (chkIncludeObsolete = 1)
            oSearchEngine.SearchBIC pCriteria, lngRecsFound
        ElseIf pSearchType = enCourseCode Then
            enSource = enLocalDB
            oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "", (chkIncludeObsolete = 1)
            oSearchEngine.SearchCourseCode pCriteria, lngRecsFound
        ElseIf pSearchType = enSearchBF Then
'360               oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "B", (chkIncludeObsolete = 1)
'370               enSource = enBF
'380               oSearchEngine.BFSearchEx pCriteria, lngRecsFound, txtmaxnum, lngResult
'          Dim bfo As NielsenLookup
'          Set bfo = New NielsenLookup
'              If bfo Is Nothing Then
'                  MsgBox "before setting bfo is nothing"
'              End If
          If oPC.GetProperty("NielsenUserID") > "" Then   'we are using the online service
          MsgBox ("POS 1")
            Dim oProduct As a_Product
            Set oProduct = New a_Product
            '========================================================
            lngRecsFound = oProduct.Load("", 0, pCriteria)
            If lngRecsFound = 1 Then
          MsgBox ("POS found")
                pCriteria = Left(pCriteria, 13)
                pSearchType = enSearchNormal
                GoTo SearchAgain
            Else
                          MsgBox ("POS NOT found")
            End If
            '========================================================
          End If

        Else
            enSource = enLocalDB
           Call oSearchEngineC.SearchOnServer(pSearchType, oPC.WorkstationID, False, False, False, IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500), 1, _
            (chkIncludeObsolete = 1), (Me.chkCopies = 1), pCriteria, lngRecsFound, rsResult, lngSectionID, _
            lngProductTypeID, IIf(pSection = "<NONE>", True, False), IIf(pProductType = "<NONE>", True, False))
            If rsResult Is Nothing Then
                lngRecsFound = 0
            Else
                If rsResult.State = 0 Then
                    lngRecsFound = 0
                Else
                    If rsResult.eof Or rsResult.Fields.Count = 1 Then
                        lngRecsFound = 0
                    Else
                        lngRecsFound = rsResult.RecordCount
                    End If
                End If
            End If
        End If
    Else
        lngRecsFound = 0
    End If
            
    If lngRecsFound > 0 Then
        If pSearchType = enSearchByCatalogue Or pSearchType = enSearchBIC Or pSearchType = enCourseCode Or pSearchType = enSearchBF Then
            oSearchEngine.execute lngMaxRecs
            Set colList = Nothing
            Set colList = oSearchEngine.getcols
            lngrows = oSearchEngine.rows
        Else
            oSearchEngineC.MassageRows rsResult
            Set colList = Nothing
            Set colList = oSearchEngineC.getcols
            lngrows = oSearchEngineC.rows
        End If
    End If
    txtRecsFound = CStr(lngRecsFound)
    
    If lngRecsFound = 0 Then
        Select Case enSource
        Case enLocalDB
            XN.Clear
            XN.ReDim 1, 1, 1, 16
            XN(1, 1) = "No records"
            GN.Array = XN
            GN.ReBind
        Case enBF
            XBF.Clear
            XBF.ReDim 1, 1, 1, 12
            XBF(1, 1) = "No records"
            GN.Array = XN
          '  GBF.ReBind
        End Select
    Else
        LoadGrid
    End If
    If lngRecsFound = IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500) Then
        MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
        Me.GN.Refresh
    End If
    '--------------
    oPC.DisconnectDBShort
    '--------------
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in BrowseProducts: Search, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in BrowseProducts: Search after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't run search."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Search(pSearchType,pCriteria,pSection,pProductType)", _
         Array(pSearchType, pCriteria, pSection, pProductType)
End Sub
Private Sub LoadGridEx()
    On Error GoTo errHandler
Dim i As Long
Dim XEX As New XArrayDB
        XEX.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
                XEX.Value(i, 1) = colList.Item(i).CodeF
                XEX.Value(i, 2) = colList.Item(i).StatusF & " " & colList.Item(i).Title
                XEX.Value(i, 3) = colList.Item(i).Author
                XEX.Value(i, 4) = colList.Item(i).Distributor
                XEX.Value(i, 5) = colList.Item(i).QtyOnHand
                XEX.Value(i, 6) = colList.Item(i).QtyonOrder
                XEX.Value(i, 7) = colList.Item(i).QtyOnBackorder
                XEX.Value(i, 8) = colList.Item(i).QtyTotalSold
                XEX.Value(i, 10) = colList.Item(i).LastDateDelivered
                XEX.Value(i, 9) = colList.Item(i).LocalPriceF
                XEX.Value(i, 11) = colList.Item(i).PID
                XEX.Value(i, 12) = colList.Item(i).code
        Next
        Set GEX.Array = XEX
        GEX.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.LoadGridEx"
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Long

    Screen.MousePointer = vbHourglass
    Select Case enSource
    Case enLocalDB
 '       GBF.Visible = False
  '      GBF.Width = 0
        GN.Visible = True
        XN.Clear
     '   XBF.Clear
    '    GBF.ReBind
        XN.ReDim 1, colList.Count, 1, 30
        For i = 1 To colList.Count
                XN.Value(i, val(oColMap.Key("Code"))) = colList.Item(i).CodeF
                XN.Value(i, val(oColMap.Key("Item"))) = UCase(colList.Item(i).StatusShortF(True, True)) & " " & colList.Item(i).Title
                XN.Value(i, val(oColMap.Key("Author"))) = colList.Item(i).Author
                XN.Value(i, val(oColMap.Key("Distributor"))) = colList.Item(i).Distributor
                XN.Value(i, val(oColMap.Key("OH/OO/CO"))) = colList.Item(i).QtyOnHand & " / " & colList.Item(i).QtyonOrder & " / " & colList.Item(i).QtyOnBackorder
                XN.Value(i, val(oColMap.Key("Publisher"))) = colList.Item(i).Publisher
                XN.Value(i, val(oColMap.Key("PublicationDate"))) = colList.Item(i).PubDate & IIf(colList.Item(i).Edition > "", "/", "") & colList.Item(i).Edition
                XN.Value(i, val(oColMap.Key("TotalSold"))) = colList.Item(i).QtyTotalSold
                XN.Value(i, val(oColMap.Key("LastDateDelivered"))) = colList.Item(i).LastDateDelivered
                XN.Value(i, val(oColMap.Key("S.P."))) = colList.Item(i).LocalPriceF
                XN.Value(i, val(oColMap.Key("Multibuy"))) = colList.Item(i).Multibuy
                XN.Value(i, val(oColMap.Key("Categories"))) = colList.Item(i).Categories
                XN.Value(i, 11) = colList.Item(i).PID
                XN.Value(i, 12) = colList.Item(i).code
                XN.Value(i, 13) = colList.Item(i).EAN
                XN.Value(i, 16) = colList.Item(i).LocalPrice
                XN.Value(i, 17) = colList.Item(i).Obsolete
                XN.Value(i, 18) = colList.Item(i).Title
                XN.Value(i, 20) = colList.Item(i).QtyOnHand
                XN.Value(i, 21) = colList.Item(i).QtyonOrder
                XN.Value(i, 22) = colList.Item(i).QtyOnBackorder
        Next
        XN.QuickSort 1, XN.UpperBound(1), 18, XORDER_ASCEND, XTYPE_STRING
        GN.Array = XN
        Me.GN.ReBind
        
        
        
'350       Case enBF
'360           XN.Clear
'370           GN.ReBind
'380           XBF.Clear
'390           GBF.Visible = True
'400           GN.Visible = False
'410           XBF.ReDim 1, colList.Count, 1, 12
'420           For i = 1 To colList.Count
'440                       XBF.Value(i, 1) = colList.Item(i).EAN
'480                   XBF.Value(i, 2) = colList.Item(i).Title
'490                   XBF.Value(i, 3) = colList.Item(i).Author
'500                   XBF.Value(i, 4) = IIf(colList.Item(i).DistributorByIdx(1) = "", "Pub by:" & colList.Item(i).Publisher, colList.Item(i).DistributorByIdx(1))
'510                   XBF.Value(i, 5) = colList.Item(i).LocalPriceF
'520                   XBF.Value(i, 6) = colList.Item(i).USPriceF
'530                   XBF.Value(i, 7) = colList.Item(i).UKPriceF
'540                   XBF.Value(i, 8) = colList.Item(i).DistributorCode & " : " & colList.Item(i).Distributor
'550                   XBF.Value(i, 12) = colList.Item(i).code
'560           Next
'570           XBF.QuickSort 1, XBF.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
'580           GN.Array = XN
'590           Me.GN.ReBind
     ''   GBF.Array = XBF
     ''   Me.GBF.ReBind
    End Select
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.LoadGrid"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
'Unload Me
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCourseCodes_Click()
    On Error GoTo errHandler
    Me.txtmaxnum = "9999999"
    If Me.Text1 > "" Then
        Screen.MousePointer = vbHourglass
        search enCourseCode, Me.Text1
        Screen.MousePointer = vbDefault
    End If
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdCourseCodes_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDebugOff_Click()
    On Error GoTo errHandler
oSearchEngine.mbDebug = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdDebugOff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDebugOn_Click()
    On Error GoTo errHandler
    oSearchEngine.mbDebug = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdDebugOn_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGetFromSB_Click()
    On Error GoTo errHandler
Dim ob As New z_SQL

    Screen.MousePointer = vbHourglass
    
    ob.RunProc "dbo.SENDBROKERMESSAGE", Array(Me.cboSearch), "TEST"
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdGetFromSB_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGOO_Click()
    On Error GoTo errHandler
    If Inet1.StillExecuting Then Exit Sub
    Screen.MousePointer = vbHourglass
    GoogleButtonNo = 1
    GooIndex = 0
    Set XGOO = New XArrayDB
    cboGoogle.AddItem cboGoogle, 0
    FetchFromGoogle
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdGOO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGOO_LostFocus()
    On Error GoTo errHandler
    
    If Inet1.StillExecuting And GoogleButtonNo = 1 Then
        cmdGOO.SetFocus
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdGOO_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdMoreGoo_LostFocus()
    On Error GoTo errHandler
    If Inet1.StillExecuting And GoogleButtonNo = 2 Then
        cmdMoreGoo.SetFocus
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdMoreGoo_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdMoreGoo_Click()
    On Error GoTo errHandler
    If Inet1.StillExecuting Then Exit Sub
    Screen.MousePointer = vbHourglass
    GoogleButtonNo = 2
    FetchFromGoogle
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdMoreGoo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub FetchFromGoogle()
    On Error GoTo errHandler
Dim se As New z_SearchEngineB
Dim SS As String
Dim r As String
Dim oG As New z_Google
    SS = se.GoogleSearchString(Me.cboGoogle)
    If bISBNOnly Then
        r = "+date:1950-2008"
    Else
        r = ""
    End If
    Inet1.URL = "http://books.google.com/books/feeds/volumes?q=" & Replace(SS, """", "%22") & r & "&max-results=20&start-index=" & CStr(GooIndex + 1)

    strGOOXML = Inet1.OpenURL
    If Left(strGOOXML, 7) = "invalid" Then
       ' MsgBox strGOOXML
        Exit Sub
    End If
    If strGOOXML > "" Then
        oG.LoadFromGoogle strGOOXML, XGOO, GooIndex
        Set GOO.Array = XGOO
        GOO.Refresh
        GOO.ReBind
        GOO.ZOrder 0
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.FetchFromGoogle"
End Sub
Private Function GetISBN(p As String) As String
    On Error GoTo errHandler
Dim a() As String
    GetISBN = ""
    a = Split(p, ":")
    If UBound(a) > 0 Then
        If a(0) = "ISBN" Then
            GetISBN = a(1)
        End If
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GetISBN(p)", p
End Function



Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    Dim fn As String
    ExportToSpreadsheet fn
    If MsgBox("Spreadsheet file saved in: " & fn & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
        OpenFileWithApplication fn, enExcel
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Public Function ExportToSpreadsheet(pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim s As String
Dim s2 As String
Dim lngNumberOfLines As Long
Dim i As Long
Dim fs As New FileSystemObject

    ExportToSpreadsheet = False
    
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder TEMP on local computer", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    
    pFilename = oPC.LocalFolder & "Temp\BrowseProducts_" & Format(Now(), "yyyymmddHHnn") & ".xls"
    
    oTF.OpenTextFile pFilename

    s = "SKU" & vbTab & "Title" & vbTab & "Author" & vbTab & "Distributor" & vbTab & "Publisher" & vbTab & "Pub Date" & vbTab & "Total sold" & vbTab & "Last Del." & vbTab _
        & "S.P." & vbTab & "Categories" & vbTab & "Obsolete" & vbTab & "On hand" & vbTab _
        & "QtyonOrder" & vbTab & "QtyOnBackorder" & vbTab & "QtyReserved"
    
    oTF.WriteToTextFile s
               
                
    lngNumberOfLines = 0
    For i = 1 To XN.Count(1)
        lngNumberOfLines = lngNumberOfLines + 1
        s = XN(i, val(oColMap.Key("Code")))
        s = s & vbTab & XN(i, val(oColMap.Key("Item")))
        s = s & vbTab & XN(i, val(oColMap.Key("Author")))
        s = s & vbTab & XN(i, val(oColMap.Key("Distributor")))
        s = s & vbTab & XN(i, val(oColMap.Key("Publisher")))
        s = s & vbTab & XN(i, val(oColMap.Key("PublicationDate")))
        s = s & vbTab & XN(i, val(oColMap.Key("TotalSold")))
        s = s & vbTab & XN(i, val(oColMap.Key("LastDateDelivered")))
        s = s & vbTab & XN(i, val(oColMap.Key("S.P.")))
     '   s = s & vbTab & XN(i, val(oColMap.key("Multibuy")))
        s = s & vbTab & XN(i, val(oColMap.Key("Categories")))
        s = s & vbTab & XN(i, 17)
        s = s & vbTab & XN(i, 20)
        s = s & vbTab & XN(i, 21)
        s = s & vbTab & XN(i, 22)
        s = s & vbTab & XN(i, 23)
        oTF.WriteToTextFile s
    Next
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
    Exit Function
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in BrowseProducts: ExportToSpreadsheet"  'unknown source
        If errRepeat < 5 Then
            Err.Clear
            Exit Function
        Else
            LogSaveToFile "Access violation in BrowseProducts: ExportToSpreadsheet after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Function
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.ExportToSpreadsheet(pFilename)", pFilename
End Function

Private Sub cmdPrintbarcode_Click()
    On Error GoTo errHandler
    Dim ar As New arPrintBarcodeList
    
    ar.component XN
    ar.Show vbModal
    Set ar = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cmdPrintbarcode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub GBF_mnuAddToDatabase()
'    str = FNS(XBF.Value(GBF.Bookmark, 1))
'    If str = "No records" Then Exit Sub
'    If str = "" Then Exit Sub
'    If MsgBox("Do you want to create a record in the database from the Bookfind data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
'
'    If CheckThisPoint(M_NEWPRODUCT) Then
'        If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then Exit Sub
'    End If
'
'    Set oProd = Nothing
'    Set oProd = New a_Product
'    Screen.MousePointer = vbHourglass
'    oProd.Load "", 0, str, , , True
'    Set frm = New frmProduct
'    frm.component oProd
'    frm.Show
'    Screen.MousePointer = vbDefault
End Sub
Private Sub Command1_Click()
    On Error GoTo errHandler
Dim str As String
    If oPC.InternetDialup = True Then Exit Sub
    Screen.MousePointer = vbHourglass
    If cboSearch.text = "" Then Exit Sub
    If IsNumeric(Left(Me.cboSearch.text, 9)) Then
        If IsNumeric(Left(Me.cboSearch.text, 13)) Then
            OpenBrowser "http://books.google.com/books?isbn=" & Left(Me.cboSearch.text, 13)
        Else
            OpenBrowser "http://books.google.com/books?isbn=" & Left(Me.cboSearch.text, 10)
        End If
    Else
        str = Replace(FNS(cboSearch), "/", "")
        OpenBrowser "http://books.google.com/books?q=" & str
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler
Dim str As String
Dim str2 As String
    If oPC.InternetDialup = True Then Exit Sub
    Screen.MousePointer = vbHourglass
    If cboSearch.text = "" Then Exit Sub
    If IsNumeric(Left(Me.cboSearch.text, 9)) Then
        If IsNumeric(Left(Me.cboSearch.text, 13)) Then
            str = "http://www.amazon.co.uk/dp/XXX"
            str = Replace(str, "XXX", Left(Me.cboSearch.text, 10))
            OpenBrowser str
        Else
            str = "http://www.amazon.co.uk/dp/XXX"
            str = Replace(str, "XXX", Left(Me.cboSearch.text, 10))
            OpenBrowser str
        End If
    Else
        
        str = "http://www.amazon.co.uk/gp/search?search-alias=stripbooks&field-keywords=&author=&select-author=field-author-like&title=XXX&select-title=field-title&subject=&select-subject=field-subject&field-publisher=&field-isbn=&chooser-sort=rank%21%2Bsalesrank&node=&field-binding=&mysubmitbutton1.x=53&mysubmitbutton1.y=12"
        str2 = Replace(FNS(cboSearch), "/", "")
        str = Replace(str, "XXX", str2)
        OpenBrowser str
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Command2_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub Form_Activate()
    On Error GoTo errHandler
' MsgBox "aCTIVATE 1"
    SetMenu
    p 1
    p 2
    p 3
'  MsgBox "aCTIVATE 2"
    bWithCopies = False
    chkCopies = IIf(bWithCopies, 1, 0)
    Me.Command1.Enabled = Not oPC.InternetDialup
'  MsgBox "aCTIVATE 3"
    cmdGetFromSB.Visible = oPC.GetProperty("UsesHUB") = "TRUE"
    lblHUBRESULT.Visible = cmdGetFromSB.Visible
'  MsgBox "aCTIVATE 4"
   Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Activate", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
    HandleError
End Sub
Private Sub SetColumnWidths()
    On Error Resume Next
Dim i As Integer
    On Error GoTo errHandler
    SaveLayout Me.GN, "SearchFormA"
    SaveLayout Me.GEX, "SearchFormB"
    SaveLayout Me.GOO, "SearchFormC"
    SaveSetting "PBKS", Me.Name, "Formwidth", Me.Width
    Exit Sub
errHandler:

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.SetColumnWidths"
End Sub


Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error Resume Next

    InitializeGrids
  '  cboSearch.SetFocus
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub InitializeGrids()
    On Error Resume Next
Dim i As Integer
   '   MsgBox "pOS 1"
    oColMap.Load oPC.GetProperty("BrowsePosition")
  '  MsgBox "pOS 2"

    Set XN = New XArrayDB
  '  Set XBF = New XArrayDB
    Set XA = New XArrayDB
    Set XGOO = New XArrayDB
    
    Set oSearchEngine = New z_SearchEngineB
    Set oSearchEngineC = New z_SearchEngineC
    
    Set colList = New Collection
'    If oPC.Configuration.AntiquarianYN Then
'        chkAntiquarianOnly.Visible = True
'        chkAntiquarianOnly = 1
'    Else
'        chkAntiquarianOnly = 0
'        chkAntiquarianOnly.Visible = False
'    End If
    SetColumnWidths
'    Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
   ' XN.ReDim 1, 1, 1, 9
    XA.ReDim 1, 0, 1, 9
  '  XBF.ReDim 1, 1, 1, 12
    XGOO.ReDim 1, 1, 1, 12
 '   SSTab1.Tab = 0
    mSetfocus cboSearch
  ''''^  GN.ZOrder 0
    FormatGN
   '   MsgBox "pOS 3"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.InitializeGrids"
End Sub
Private Sub FormatGN()
    On Error GoTo errHandler
Dim i As Integer
Dim lngAlignment As Long

 ' If Me.GN Is Null Then Exit Sub
  
    For i = 0 To oColMap.Count - 1
        Select Case oColMap.Item_2(i)
        Case "L"
            lngAlignment = 0
        Case "R"
            lngAlignment = 1
        Case "C"
            lngAlignment = 2
        End Select
        If CLng((oColMap.Item_3(i))) <= GN.Columns.Count Then
            GN.Columns(CLng((oColMap.Item_3(i))) - 1).Alignment = lngAlignment
            GN.Columns(CLng((oColMap.Item_3(i))) - 1).Caption = oColMap.Item_1(i)
        End If
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.FormatGN"
End Sub

Private Sub Form_Load()
10        On Error Resume Next
      Dim i As Integer
20         errSysHandlerSet
30        Set PrivateCnn = New ADODB.Connection
40        oPC.OpenSuppliedConnection PrivateCnn

50        GN.Left = 120
60        GOO.Left = 120
70        GN.TOP = 1727
80        GOO.TOP = 1727
90        XA.Clear
100       XA.ReDim 1, 0, 1, 9
110       XGOO.Clear
120       XGOO.ReDim 1, 1, 1, 12
130       flgLoading = True
140       Resize_GN
150   On Error Resume Next
160       Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
170       flgLoading = False
180       OriginalwdthGN = 250
190       For i = 1 To GN.Columns.Count - 1
200           OriginalwdthGN = OriginalwdthGN + GN.Columns(i - 1).Width
210           wdthCol(i - 1) = GN.Columns(i - 1).Width
220       Next
230       If Me.WindowState <> 2 Then
240           Me.TOP = 20
250           Me.Left = 50
260           Height = 7220
270       End If
280    On Error GoTo errHandler
         
290       Set tlCats = Nothing
300       Set tlCats = New z_TextList
          
310       If oPC.SupportsCatalogue Then
320           tlCats.Load ltCatalogue
330           LoadCombo cboCat, tlCats
340       End If
350       frCatalogue.Enabled = oPC.SupportsCatalogue
          
360       LoadCombo cboCategory, oPC.Configuration.Sections
370       Me.cboCategory.AddItem "<NONE>"
          
380       LoadCombo cboProductType, oPC.Configuration.ProductTypes
390       cboProductType.AddItem "<NONE>"
          
400       If oPC.Configuration.AntiquarianYN Then
410           Me.GN.Columns(3).Caption = "Publisher"
420       Else
430           GN.Columns(3).Caption = "Distributor"
440       End If
450       txtmaxnum = 500

460       Me.chkIncludeObsolete = IIf(GetSetting("PBKS", Me.Name, "IncludeObsolete", 0) = 1, CInt("1"), CInt("0"))
          
470       On Error Resume Next
480       cboCategory = "<ALL>"
490       cboProductType = "<ALL>"
         
500       SetActiveGrid "GN"

510       Exit Sub
errHandler:
520       If ErrMustStop Then Debug.Assert False: Resume
530       ErrorIn "frmBrowseProducts.Form_Load", , EA_NORERAISE
540       HandleError
End Sub
Private Sub Resize_GBF()
Dim i As Long
'MsgBox
'    For i = 1 To GBF.Columns.Count
'        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
'    Next
End Sub
Private Sub Resize_GOO()
    On Error Resume Next
Dim i As Long
    For i = 1 To GOO.Columns.Count
        GOO.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormC", CStr(i), GOO.Columns(i - 1).Width)
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Resize_GOO"
End Sub
Private Sub Resize_GN()
    On Error Resume Next
Dim i As Long
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Resize_GN"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
Dim lngDiff As Long
'MsgBox "Resize Pos 1"
    If flgLoading Then Exit Sub
    GN.Width = NonNegative_Lng(Me.Width - 380)
  '  GBF.Width = GN.Width
    GOO.Width = GN.Width
'MsgBox "Resize Pos 2"
    
    lngDiff = GN.Height
    GN.Height = NonNegative_Lng(Me.Height - (GN.TOP + 1070))
'    GBF.Height = GN.Height
    GOO.Height = GN.Height
'MsgBox "Resize Pos 3"
    lngDiff = GN.Height - lngDiff
    
    cmdPrint.TOP = GN.TOP + GN.Height + 30
    cmdClose.TOP = GN.TOP + GN.Height + 30
    cmdPrintbarcode.TOP = GN.TOP + GN.Height + 30
    cmdClose.Left = NonNegative_Lng(GN.Width - 1000)
'MsgBox "Resize Pos 4"
   ' ResizeColumns
       FormatGN

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Resize", , EA_NORERAISE
    HandleError
End Sub
Sub ResizeColumns()
    On Error Resume Next
Dim i As Integer
Dim newwidth As Long
'MsgBox "ResizeColumns Pos 1"

    For i = 1 To GN.Columns.Count - 1
        GN.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GN.Width) / OriginalwdthGN) * 0.9)
    Next
    For i = 1 To GOO.Columns.Count - 1
        GOO.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GOO.Width) / OriginalwdthGN) * 0.9)
    Next
'MsgBox "ResizeColumns Pos 2"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.ResizeColumns"
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
    Set XGOO = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    SaveSetting "PBKS", Me.Name, "IncludeObsolete", IIf(Me.chkIncludeObsolete = 1, "1", "0")
'---------------------------------------------------
    If PrivateCnn Is Nothing Then Exit Sub
    oPC.CloseSUppliedConnection PrivateCnn
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub GBF_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    
'MsgBox "GBF_RowColChange Pos 1"
    
    If XBF Is Nothing Then Exit Sub
    If XBF.UpperBound(1) = 0 Then Exit Sub
    If Err Then Exit Sub
 '   If IsNull(GBF.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
    'MsgBox "Code Commented"
'    str = FNS(XBF.Value(GBF.Bookmark, 1))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GBF_Click", , EA_NORERAISE, , "Line number", Array(Erl())
    HandleError
End Sub

Private Sub GBF_DblClick()
    On Error GoTo errHandler
Dim oProd As a_Product
Dim str As String
Dim frm As frmProduct
          'gBox "Code Commented"

'    str = FNS(XBF.Value(GBF.Bookmark, 1))
'    If str = "No records" Then Exit Sub
'    If str = "" Then Exit Sub
'    If MsgBox("Do you want to create a record in the database from the Bookfind data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
'
'    If CheckThisPoint(M_NEWPRODUCT) Then
'        If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then Exit Sub
'    End If
'
'    Set oProd = Nothing
'    Set oProd = New a_Product
'    Screen.MousePointer = vbHourglass
'    oProd.Load "", 0, str, , , True
'    Set frm = New frmProduct
'    frm.component oProd
'    frm.Show
'    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in BrowseProducts: GBF_DblCLick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in BrowseProducts: GBF_DblCLick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
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






Private Sub GN_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XN.Value(1, 5) = "" Then Exit Sub
    If XN(Bookmark, 17) = True Then
        RowStyle.Font.Strikethrough = True
    End If
    If Len(XN(Bookmark, 2)) > 1 Then
        If Left(XN(Bookmark, 2), 2) = "**" Then
            RowStyle.BackColor = RGB(220, 220, 220)
        End If
    End If
    If Len(XN(Bookmark, 2)) > 5 Then
        If Left(XN(Bookmark, 2), 6) = "**(OP)" Then
            RowStyle.BackColor = RGB(180, 180, 180)
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub GOO_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
'MsgBox "GOO_RowColChange Pos 1"

    On Error Resume Next
    If XGOO Is Nothing Then Exit Sub
    If XGOO.Count(1) = 0 Then Exit Sub
    If XGOO.UpperBound(1) = 0 Then Exit Sub
    If Err Then Exit Sub
    If IsNull(GOO.Bookmark) Then Exit Sub
    On Error GoTo errHandler
    
    On Error Resume Next
    str = FNS(XGOO.Value(GOO.Bookmark, 1))
    On Error GoTo errHandler
    If str = "" Then Exit Sub
    
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub

    Exit Sub
errHandler:
    ErrorIn "frmBrowseProducts.GOO_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE, , "Line number", Array(Erl())
    HandleError
    
End Sub



Private Sub GN_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If (KeyCode = vbKeyLeft) Then
        mSetfocus cboSearch
    End If
    If Shift = 1 Then
        bShiftDown = True
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim errRepeat As Integer
Dim str As String
'MsgBox "GN_RowColChange Pos 1"

    errRepeat = 0

    On Error Resume Next
    If LastRow = "" Then Exit Sub
    If XN.Count(1) = 0 Then Exit Sub
    If GN.VisibleRows < 1 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
On Error GoTo errHandler

    If IsNumeric(GN.Bookmark) Then
        If GN.Bookmark <= XN.UpperBound(1) Then
            str = IIf(FNS(XN.Value(GN.Bookmark, 13)) > "", FNS(XN.Value(GN.Bookmark, 13)), FNS(XN.Value(GN.Bookmark, 12)))
        End If
    End If

    If str = "" Then Exit Sub
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Or Err.Number = 2147227667 Then    'Access violation
  errRepeat = errRepeat + 1
  LogSaveToFile "Access violation in BrowseProducts: GN_RowColChanged"  'unknown source
  If errRepeat < 5 Then
      Resume Next
  Else
      LogSaveToFile "Access violation in BrowseProducts: GN_RowColChanged after 5 re-attempts"
      MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
      Err.Clear
      Exit Sub
  End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE, , "strErrPos,Line number", Array(strErrPos, Erl())
    HandleError
End Sub
Public Property Get NextPID() As String
    On Error GoTo errHandler
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
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.NextPID"
End Property
Public Property Get PrevPID() As String
    On Error GoTo errHandler
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
        
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.PrevPID"
End Property
Private Sub GN_DblClick()
    OpenInventoryRecord
End Sub

Public Sub OpenInventoryRecord()
    On Error GoTo errHandler
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim frmNB As frmProductNBPrev   'non book form
Dim lngprod As Long
Dim errRepeat As Integer

    errSysHandlerSet

    errRepeat = 0
    If XN.Count(1) = 0 Then
        LogSaveToFile "XN.Count(1) = 0"
        Exit Sub
    End If
    If IsNull(GN.Bookmark) Then
        LogSaveToFile "IsNull(GN.Bookmark) is true"
        Exit Sub
    End If
    If GN.Bookmark > XN.Count(1) Then
      LogSaveToFile "GN.Bookmark > XN.Count(1)"
      Exit Sub
    End If
    BookmarkPointer = GN.Bookmark
    strPID = FNS(XN.Value(GN.Bookmark, 11))
    If strPID = "" Then
        LogSaveToFile "strPID = ''"
        Exit Sub
    End If
    If bShiftDown Then
        ShowSalesPatterns
    Else
        Set oProduct = New a_Product
        Screen.MousePointer = vbHourglass
        oProduct.Load strPID, 0, "", strTime
        If oProduct.PID = "" Then Exit Sub
        If oProduct.ProductType = "B" Then
            If oPC.Configuration.AntiquarianYN Then
                  Set frmA = Nothing
                Set frmA = New frmProductPrevAQ
                frmA.component oProduct
                frmA.Show
            Else
                  Set frm = Nothing
                Set frm = New frmProductPrev
                frm.component oProduct, strTime
                frm.Show
            End If
        Else
            Set frm = Nothing
            Set frmNB = New frmProductNBPrev
            frmNB.component oProduct, strTime
            frmNB.Show
        End If
    End If
    Set oProduct = Nothing
    Screen.MousePointer = vbDefault
    bShiftDown = False
    Exit Sub
errHandler:
    ErrPreserve
    Set frm = Nothing
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in BrowseProducts: GN_DblCLick, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
        If errRepeat < 5 Then
              Set frm = Nothing
              Err.Clear
              Exit Sub
        Else
            LogSaveToFile "Access violation in BrowseProducts: GN_DblCLick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.OpenInventoryRecord", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
    HandleError
End Sub
Public Sub ShowSalesPatterns()
    On Error GoTo errHandler
Dim frmSales As frmSalesCH
    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub

    Screen.MousePointer = vbHourglass
    Set oProduct = New a_Product
    strPID = FNS(XN.Value(GN.Bookmark, 11))
    If strPID = "" Then Exit Sub

    oProduct.Load strPID, 0
    If oProduct.PID = "" Then Exit Sub
    Set frmSales = New frmSalesCH
    frmSales.component oProduct
    frmSales.Show
    Screen.MousePointer = vbDefault
    Set frmSales = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.ShowSalesPatterns"
End Sub
Private Sub GN_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If XN.Count(1) = 0 Then Exit Sub
    If XN(1, 1) = "No records" Then Exit Sub
    If XN.UpperBound(1) = 0 Then Exit Sub
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 0 Then ColIndex = 12
    If ColIndex = 1 Then ColIndex = 18
        XN.QuickSort XN.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
 '   Else
 '       XN.QuickSort XA.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
 '   End If
    
    GN.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE, , "colcount", CLng(XN.Count(1))
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
  '  Select Case ColIndex
        If ColIndex > 10 Then
            GetRowType = XTYPE_STRING
        Else
        If GN.Columns(ColIndex - 1).Alignment = dbgRight Then
            GetRowType = XTYPE_INTEGER
        Else
            GetRowType = XTYPE_STRING
        End If
        End If
'        Case 1, 2, 3, 4, 12
'            GetRowType = XTYPE_STRING
'        Case 5, 6, 7, 8, 9
'            GetRowType = XTYPE_INTEGER
   ' End Select
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

Public Sub mnuProductStatus()
    On Error GoTo errHandler
Dim frm As New frmPreDeliveryAdvice
Dim IDs As String
Dim i, j As Integer
Dim x As New XArrayDB
Dim XMLArgs As String
    
    If XN.Count(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    If GN.SelBookmarks.Count = 0 Then Exit Sub
    strPID = FNS(XN.Value(GN.Bookmark, 11))
    If strPID = "" Then Exit Sub
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_PRE_DEL_ADVICE"
            .chCreate "MessageType"
                .elText = "PRE_DEL_ADVICE"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "WORKSTATION"
                .elText = oPC.WorkstationName
            .elCreateSibling "DetailLines", True
            For i = 1 To GN.SelBookmarks.Count
                    .chCreate "ITEM"
                    .chCreate "PID"
                        .elText = CStr(XN(GN.SelBookmarks.Item(i - 1), 11))
                    .navUP
                    .navUP
            Next i

         XMLArgs = .docXML
    End With
    
    If XMLArgs > "" Then
        frm.component XMLArgs, IDs, "R", ""
        frm.Show vbModal
    Else
        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuProductStatus"
End Sub
Private Sub GN_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
Dim errRepeat As Integer
    errRepeat = 0
    If XN.Count(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      PopupMenu Forms(0).mnuFindForm   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in BrowseProducts: Mousedown"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in BrowseProducts: GN_MouseDown after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GN_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Public Sub AddToTempList()
    On Error GoTo errHandler
Dim str As String
Dim i As Integer
Dim TOP As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    
    'Get PID for current row
'    str = FNS(XN.Value(GN.Bookmark, 11))
'    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
'        If XA(XA.UpperBound(1), 1) > "" Then
'            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 9
'        End If
'        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
'        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
'        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
'        XA(XA.UpperBound(1), 4) = 1
'        XA(XA.UpperBound(1), 5) = 0
'        XA(XA.UpperBound(1), 6) = ""
'        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 9))
'        XA(XA.UpperBound(1), 8) = FNS(XN.Value(GN.Bookmark, 11))
'        XA(XA.UpperBound(1), 9) = FNS(XN.Value(GN.Bookmark, 16))
'    End If
    TOP = XA.UpperBound(1)
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(TOP + i, 4) = 1
        XA(TOP + i, 5) = 0
        XA(TOP + i, 6) = ""
        XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next
  
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.AddToTempList"
End Sub
Public Sub PlaceCO()
    On Error GoTo errHandler
Dim frm As New frmPlaceCO
Dim TOP As Integer
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    
    TOP = XA.UpperBound(1)
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(TOP + i, 4) = 1
        XA(TOP + i, 5) = 0
        XA(TOP + i, 6) = ""
        XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next
    
    frm.component XA, "ORDER"
    frm.Show 'vbModal
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.PlaceCO"
End Sub

Public Sub PrintLabels()
    On Error GoTo errHandler
Dim frm As frmPrintLabels
Dim str As String
Dim TOP As Integer
Dim i As Integer
    Set frm = New frmPrintLabels
    
    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    
    TOP = XA.UpperBound(1)
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        If Len(FNS(XN.Value(GN.SelBookmarks(i - 1), 1))) = 13 Then
            XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        Else
            XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 13))
        End If
        XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(TOP + i, 4) = 1
        XA(TOP + i, 5) = 0
        XA(TOP + i, 6) = FNS(XN.Value(GN.SelBookmarks(i - 1), 13))
        XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next
    
    frm.component "S", , XA
    frm.Show
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.PrintLabels"
End Sub


Public Sub PlacePF(strType As String)
    On Error GoTo errHandler
Dim frm As New frmPlacePF
Dim str As String
Dim TOP As Integer
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    
    TOP = XA.UpperBound(1)
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(TOP + i, 4) = 1
        XA(TOP + i, 5) = 0
        XA(TOP + i, 6) = ""
        XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next
    
    frm.component XA, strType
    frm.Show
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.PlacePF(strType)", strType
End Sub

Public Sub PlaceOnReserve()
    On Error GoTo errHandler
Dim frm As New frmPlaceCO
Dim str As String
Dim TOP As Integer
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    
    TOP = XA.UpperBound(1)
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(TOP + i, 4) = 1
        XA(TOP + i, 5) = 0
        XA(TOP + i, 6) = ""
        XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next
    frm.component XA, "RESERVE"
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.PlaceOnReserve"
End Sub
Public Sub StartNewList()
    On Error GoTo errHandler
    XA.Clear
    XA.ReDim 1, 0, 1, 9
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.StartNewList"
End Sub


Private Sub GOO_DblClick()
    On Error GoTo errHandler
Dim oProd As a_Product
Dim str As String
Dim sMsg As String
Dim lngRes As Long
Dim errRepeat As Integer

    errRepeat = 0
    str = FNS(XGOO.Value(GOO.Bookmark, 1))
    If str = "No records" Then Exit Sub
    If XGOO.Value(GOO.Bookmark, 2) = "" Then
        MsgBox "Item has no title. Cannot save", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If MsgBox("Do you want to create a record in the database from the Google data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    Set oProd = Nothing
    Set oProd = New a_Product
    oProd.BeginEdit
    oProd.SetProductType "B"
    oProd.SetTitle XGOO.Value(GOO.Bookmark, 2)
    If IsISBN13(XGOO.Value(GOO.Bookmark, 1), True) Then
        oProd.SetEAN XGOO.Value(GOO.Bookmark, 1)
    Else
        oProd.SetCode "#"
    End If
    oProd.SetAuthor XGOO.Value(GOO.Bookmark, 3)
    oProd.SetPublisher XGOO.Value(GOO.Bookmark, 4)
    oProd.SetPublicationDate XGOO.Value(GOO.Bookmark, 6)
    oProd.SetDescription XGOO.Value(GOO.Bookmark, 5)
    oProd.ApplyEdit lngRes, sMsg
    Screen.MousePointer = vbDefault
    MsgBox "Record added", , "Status"
    Exit Sub

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in BrowseProducts: GOO_DblCLick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in BrowseProducts: GOO_DblCLick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.GOO_DblClick", , EA_NORERAISE
    HandleError
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
    On Error GoTo errHandler
    If cboSearch = "" And UCase(Me.cboCategory) = "<ALL>" And UCase(Me.cboCategory) = "<ALL>" Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        If GN.Visible = True Then
            mSetfocus GN
        Else
          '  mSetfocus GBF
        End If
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.cboSearch_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub Label26_dblClick()
    On Error GoTo errHandler
    cboCategory = "<All>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Label26_dblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Label3_DblClick()
    On Error GoTo errHandler
    cboSearch = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Label3_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Label4_DblClick()
    On Error GoTo errHandler
    cboProductType = "<All>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Label4_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    
    'Attempting to avoid occasional windows crash when doublwclicking the GN grid. Reducing the size of the other grids so one does not overlay another 17/7/2010
    If SSTab1.Tab = 0 Then
        mSetfocus Me.cboSearch
    ElseIf SSTab1.Tab = 1 Then
        SetActiveGrid "GN"
      '  GBF.Visible = True
    Else
        SetActiveGrid "GOO"
        mSetfocus cboGoogle
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.SSTab1_Click(PreviousTab)", PreviousTab, EA_NORERAISE
    HandleError
End Sub
Private Sub SetActiveGrid(Gridcode As String)
    On Error GoTo errHandler
    If Gridcode = sCurrentGrid Then Exit Sub
    Select Case Gridcode
    Case "GN"
        GN.ZOrder 0
      '  GBF.ZOrder 1
        GOO.ZOrder 1
        If sCurrentGrid = "GBF" Then
'80                GN.Width = GBF.Width
'90                GN.Height = GBF.Height
'100               GBF.Width = 0
'110               GBF.Height = 0
        Else
            GN.Width = GOO.Width
            GN.Height = GOO.Height
            GOO.Width = 0
            GOO.Height = 0
        End If
        sCurrentGrid = "GN"
        GN.Visible = True
'200       Case "GBF"
'210      '     GBF.ZOrder 0
'220           GN.ZOrder 1
'230           GOO.ZOrder 1
'240           If sCurrentGrid = "GN" Then
'250               GBF.Width = GN.Width
'260               GBF.Height = GN.Height
'270               GN.Width = 0
'280               GN.Height = 0
'290           Else
'300               GBF.Width = GOO.Width
'310       '        GBF.Height = GOO.Height
'320               GOO.Width = 0
'330               GOO.Height = 0
'340           End If
'350           sCurrentGrid = "GBF"
'360           GBF.Visible = True
    Case "GOO"
        GOO.ZOrder 0
     '   GBF.ZOrder 1
        GN.ZOrder 1
        If sCurrentGrid = "GN" Then
            GOO.Width = GN.Width
            GOO.Height = GN.Height
            GN.Width = 0
            GN.Height = 0
        Else
        '    GOO.Width = GBF.Width
         '   GOO.Height = GBF.Height
        ' '   GBF.Width = 0
         '   GBF.Height = 0
        End If
        sCurrentGrid = "GOO"
        GOO.Visible = True
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.SetActiveGrid(Gridcode)", Gridcode
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
    ErrorIn "frmBrowseProducts.txtmaxnum_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'Private Function IsAmongBookmarks(X As XArrayDB, ID As String, G As TDBGrid, IDPosInGrid As Long) As Boolean
'    Dim i As Integer
'    IsAmongBookmarks = False
'    For i = 1 To G.SelBookmarks.Count
'        If (X.Value(G.SelBookmarks(i - 1), IDPosInGrid)) = ID Then
'            IsAmongBookmarks = True
'            Exit For
'        End If
'    Next i
'End Function

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
            .elCreateSibling "Col_7"
                .elText = "OH"
            .elCreateSibling "Col_8"
                .elText = "OO"
            .elCreateSibling "Col_9"
                .elText = "CO"
                .navUP
        
'        For i = 1 To colList.Count
'            If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" And Me.SSTab1.Tab = 0 Then
'                If mIsAmongBookmarks(XBF, colList(i).EAN, GBF, 1, "STRING") Then
'                    .elCreateSibling "DetailLine", True
'                    .chCreate "Col_1"
'                        .elText = colList.Item(i).EAN
'                    .elCreateSibling "Col_2"
'                        .elText = colList.Item(i).statusF & " " & colList.Item(i).Title
'                    .elCreateSibling "Col_3"
'                        .elText = colList.Item(i).Author
'                    .elCreateSibling "Col_4"
'                        .elText = colList.Item(i).Publisher
'                    .elCreateSibling "Col_5"
'                        .elText = colList.Item(i).UKPriceF & " / " & colList.Item(i).USPriceF
'                    .elCreateSibling "Col_6"
'                        .elText = "n/a" 'colList.Item(i).QtyOnHand
'                        .navUP
'                End If
'            Else
'                If mIsAmongBookmarks(XN, colList(i).pID, GN, 11, "UNIQUEIDENTIFIER") Then
'                    .elCreateSibling "DetailLine", True
'                    .chCreate "Col_1"
'                        .elText = colList.Item(i).CodeF
'                    .elCreateSibling "Col_2"
'                        .elText = colList.Item(i).statusF & " " & colList.Item(i).Title
'                    .elCreateSibling "Col_3"
'                        .elText = colList.Item(i).Author
'                    .elCreateSibling "Col_4"
'                        .elText = colList.Item(i).Publisher
'                    .elCreateSibling "Col_5"
'                        .elText = colList.Item(i).LocalPriceF
'                    .elCreateSibling "Col_6"
'                        .elText = colList.Item(i).QtyOnHand
'                        .navUP
'                End If
'            End If
'        Next i
            If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" And Me.SSTab1.Tab = 0 Then
'290                   For i = 1 To XBF.UpperBound(1)
'300                       If mIsAmongBookmarks(XBF, colList(i).EAN, GBF, 1, "STRING") Then
'310                           .elCreateSibling "DetailLine", True
'320                           .chCreate "Col_1"
'330                               .elText = XBF.Value(i, 1)
'340                           .elCreateSibling "Col_2"
'350                               .elText = XBF.Value(i, 2)
'360                           .elCreateSibling "Col_3"
'370                               .elText = XBF.Value(i, 3)
'380                           .elCreateSibling "Col_4"
'390                               .elText = XBF.Value(i, 6)
'400                           .elCreateSibling "Col_5"
'410                               .elText = XBF.Value(i, 9)
'420                           .elCreateSibling "Col_6"
'430                               .elText = "n/a" 'colList.Item(i).QtyOnHand
'440                               .navUP
'450                       End If
'460                   Next
            Else
                For i = 1 To XN.UpperBound(1)
                    If mIsAmongBookmarks(XN, XN.Value(i, 11), GN, 11, "UNIQUEIDENTIFIER") Then
                        .elCreateSibling "DetailLine", True
                        .chCreate "Col_1"
                            .elText = XN.Value(i, 1)
                        .elCreateSibling "Col_2"
                            .elText = XN.Value(i, 2)
                        .elCreateSibling "Col_3"
                            .elText = XN.Value(i, 3)
                        .elCreateSibling "Col_4"
                            .elText = XN.Value(i, 6)
                        .elCreateSibling "Col_5"
                            .elText = XN.Value(i, 9)
                        .elCreateSibling "Col_6"
                            .elText = XN.Value(i, 5)
                        .elCreateSibling "Col_7"
                            .elText = XN.Value(i, 20)
                        .elCreateSibling "Col_8"
                            .elText = XN.Value(i, 21)
                        .elCreateSibling "Col_9"
                            .elText = XN.Value(i, 22)
                            .navUP
                    End If
                Next
            End If

        
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\SSI" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "SSI_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

'    strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1.RTF"
'    If fs.FileExists(strFilename) Then
'        fs.DeleteFile strFilename, True
'    End If
    strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1" & "_" & CStr(i) & ".RTF"
    Loop
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile

    strExecutable = GetPDFExecutable(strFilename)
  If strExecutable = "" Then
      MsgBox "There is no application set on this computer to open the file: " & strFilename & ". The document cannot be displayed", vbOKOnly, "Can't do this"
  Else
      Shell strExecutable & " " & strFilename
  End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.ExportToXML"
End Function

Public Sub mnuSetPT()
    On Error GoTo errHandler
Dim IDs As String
Dim frm As New frmSetProductType
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    ReDim strTitle(GN.SelBookmarks.Count)
    IDs = ""
    For i = 0 To GN.SelBookmarks.Count - 1
        IDs = IDs & ",'" & XN(GN.SelBookmarks(i), 11) & "'"
    Next i
    If Left(IDs, 1) = "," Then
        IDs = Right(IDs, Len(IDs) - 1)
    End If
    If IDs > "" Then
        frm.component IDs
        frm.Show vbModal
    Else
        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
    End If
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuSetPT"
End Sub
Public Sub SetForWebExport()
    On Error GoTo errHandler
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.SetForWebExport"
End Sub
Public Sub mnuSetSection()
    On Error GoTo errHandler
Dim IDs As String
Dim frm As New frmSetSection
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    ReDim strTitle(GN.SelBookmarks.Count)
    IDs = ""
    For i = 0 To GN.SelBookmarks.Count - 1
        IDs = IDs & ",'" & XN(GN.SelBookmarks(i), 11) & "'"
    Next i
    If Left(IDs, 1) = "," Then
        IDs = Right(IDs, Len(IDs) - 1)
    End If
    If IDs > "" Then
        frm.component IDs
        frm.Show vbModal
    Else
        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
    End If
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuSetSection"
End Sub
Public Sub mnuFindAllSOH()
    On Error GoTo errHandler
Dim IDs As String
Dim frm As New frmSOHALL
Dim i As Integer
Dim strEAN As String
Dim strTitle As String
Dim oSQL As New z_SQL
Dim lngREQID As Long
    p 0

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    strEAN = XN(GN.Bookmark, 13)
    strTitle = XN(GN.Bookmark, 2)
    oSQL.Request_SOH_ALLBRANCHES strEAN, lngREQID
    MsgWaitObj 2000
    If strEAN > "" And lngREQID > 0 Then
    p 1
        frm.Caption = strEAN
        frm.component lngREQID, strTitle
        frm.Show vbModal
    Else
        MsgBox "You have not clicked on a row.", vbInformation, "No row"
    End If
    Unload frm
    p 2

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseProducts.mnuFindAllSOH"  'unknown source
        If errRepeat < 5 Then
            Err.Clear
            Exit Sub
        Else
            LogSaveToFile "Access violation in frmBrowseProducts.mnuFindAllSOH after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuFindAllSOH", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
    HandleError
End Sub
Public Sub mnuTouchRecord()
    On Error GoTo errHandler
Dim cnt As Integer
    cnt = 0
    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuTouchRecord"
End Sub
Public Sub mnuLoadReorderSlate()
    On Error GoTo errHandler
Dim frmREORDER_SAL As frmREORDER_CO
Dim oSQL As New z_SQL
Dim i As Integer
Dim TOP As Integer
p 1

    If GN = "" Or GN = "No records" Then Exit Sub
p 2
    If GN.Bookmark = 0 Then Exit Sub
p 3
    TOP = XA.UpperBound(1)
p 4
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
p 5
    For i = 1 To GN.SelBookmarks.Count
        XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(TOP + i, 4) = 1
        XA(TOP + i, 5) = 0
        XA(TOP + i, 6) = ""
        XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next
p 6
    oSQL.LoadBrowsedProductsToTempTable XA
p 7
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.mnuLoadReorderSlate"
End Sub

Private Sub TouchRecord(pPID As String)
    On Error GoTo errHandler
Dim oSQL As New z_SQL

'    oSQL.RunSQL "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
'            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_LoyaltyRATE," _
'            & "PRU_PTID,PRU_SECID,PRU_MULTIBUYCODE) " _
'            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
'            & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID,vMultibuyCode.DICT_System " _
'            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID   LEFT JOIN vMultibuyCode ON P_ID = vMultibuyCode.PSEC_P_ID" _
'            & " WHERE P_ID = '" & pPID & "'"
    oSQL.RunSQL "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_SSP,PRU_NDA,PRU_LoyaltyRATE," _
            & "PRU_PTID,PRU_SECID,PRU_MULTIBUYCODE) " _
            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
            & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_Special,P_NDA,P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID,P_MultibuyCode " _
            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID   LEFT JOIN vMultibuyCode ON P_ID = vMultibuyCode.PSEC_P_ID" _
            & " WHERE P_ID = '" & pPID & "'"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.TouchRecord(pPID)", pPID
End Sub

Private Sub MarkProductForWebExport(pPID As String)
    On Error GoTo errHandler
Dim oSQL As New z_StockManager

    oSQL.MarkForWebExport pPID

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.MarkProductForWebExport(pPID)", pPID
End Sub

Public Sub mnuAddToSpecialOrder(pSTAFFID As Long)
Dim oSM As New z_StockManager
Dim f As frmCOPreview
Dim XMLArgs As String
Dim i As Long

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_SpecialOrderAddition"
            .chCreate "MessageType"
                .elText = "SpecialOrderAddition"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "StaffMember"
                .elText = CStr(pSTAFFID)
            .elCreateSibling "DetailLines", True
            For i = 0 To GN.SelBookmarks.Count - 1
                    .chCreate "ITEM"
                    .chCreate "PID"
                        .elText = XN(GN.SelBookmarks(i), 11)
'                    .elCreateSibling "CodeF"
'                        .elText = colList.Item(GN.SelBookmarks(i)).CodeF
'                    .elCreateSibling "Description"
'                        .elText = Replace(UCase(colList.Item(GN.SelBookmarks(i)).StatusShortF(True, True)) & " " & colList.Item(GN.SelBookmarks(i)).Title, "'", "''")

                    .navUP
                    .navUP
            Next i

         XMLArgs = .docXML
    End With
    
    If XMLArgs > "" Then
        oSM.CreateSpecialOrder XMLArgs
'        frm.component XMLArgs, IDs, "R", ""
'        frm.Show vbModal
    Else
        MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
    End If
    Exit Sub
End Sub
