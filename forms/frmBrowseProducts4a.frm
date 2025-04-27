VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Object = "{8B07DDC0-1FC2-4D71-B114-D4F3E02F1F1A}#1.0#0"; "PBKS_Net_Controls.tlb"
Begin VB.Form frmBrowseProducts 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse products"
   ClientHeight    =   8520
   ClientLeft      =   240
   ClientTop       =   1020
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseProducts4a.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   14265
   Begin PBKS_Net_ControlsCtl.BookfindGrid GBF 
      Height          =   4125
      Left            =   4935
      TabIndex        =   42
      Top             =   2970
      Visible         =   0   'False
      Width           =   6690
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "13882315"
      SearchArguments =   ""
      FirstRownumberToReturn=   "0"
      BackColor       =   "203, 211, 211"
      ForeColor       =   "ControlText"
      Location        =   "329, 198"
      Name            =   "BookfindGrid"
      Size            =   "446, 275"
      Object.TabIndex        =   "0"
   End
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
      TabIndex        =   26
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   7440
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   90
      Picture         =   "frmBrowseProducts4a.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6195
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1560
      Left            =   120
      TabIndex        =   6
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
      TabPicture(0)   =   "frmBrowseProducts4a.frx":07CC
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
      TabPicture(1)   =   "frmBrowseProducts4a.frx":07E8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "frCatalogue"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Google Books"
      TabPicture(2)   =   "frmBrowseProducts4a.frx":0804
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
         TabIndex        =   39
         Top             =   360
         Width           =   4050
         Begin VB.CommandButton cmdCourseCodes 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   2940
            Picture         =   "frmBrowseProducts4a.frx":0820
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   255
            Width           =   945
         End
         Begin VB.TextBox Text1 
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   195
            TabIndex        =   40
            Top             =   390
            Width           =   2580
         End
      End
      Begin VB.CheckBox chkIncludeObsolete 
         Appearance      =   0  'Flat
         Caption         =   "Include obsolete"
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
         Left            =   2580
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1590
      End
      Begin VB.CheckBox chkISBNOnly 
         Caption         =   "Show only books published since 1950"
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
         Height          =   270
         Left            =   -74745
         TabIndex        =   37
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   630
         Width           =   1260
      End
      Begin VB.ComboBox cboGoogle 
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
         ItemData        =   "frmBrowseProducts4a.frx":0BAA
         Left            =   -74865
         List            =   "frmBrowseProducts4a.frx":0BAC
         TabIndex        =   29
         Top             =   660
         Width           =   6375
      End
      Begin VB.CommandButton cmdGOO 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&First"
         Height          =   435
         Left            =   -68370
         Style           =   1  'Graphical
         TabIndex        =   28
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   1140
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search by BIC codes (if captured)"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   -71415
         TabIndex        =   22
         Top             =   375
         Width           =   3300
         Begin VB.CommandButton cmdBIC 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   960
            Picture         =   "frmBrowseProducts4a.frx":0BAE
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   255
            Width           =   1260
         End
      End
      Begin VB.Frame frCatalogue 
         Caption         =   "Search by catalogue"
         ForeColor       =   &H8000000D&
         Height          =   1020
         Left            =   -74790
         TabIndex        =   19
         Top             =   375
         Width           =   3240
         Begin VB.CommandButton cmdCAT 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   705
            Left            =   2070
            Picture         =   "frmBrowseProducts4a.frx":0F38
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   210
            Width           =   975
         End
         Begin VB.ComboBox cboCat 
            ForeColor       =   &H00800000&
            Height          =   360
            ItemData        =   "frmBrowseProducts4a.frx":12C2
            Left            =   165
            List            =   "frmBrowseProducts4a.frx":12C4
            TabIndex        =   20
            Top             =   390
            Width           =   1410
         End
      End
      Begin VB.ComboBox cboCategory 
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
         Left            =   5535
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
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
         Left            =   5520
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   510
         Width           =   2115
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Height          =   810
         Left            =   8025
         Picture         =   "frmBrowseProducts4a.frx":12C6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   510
         Width           =   1260
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
         ItemData        =   "frmBrowseProducts4a.frx":1650
         Left            =   840
         List            =   "frmBrowseProducts4a.frx":1652
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
         Left            =   1035
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
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
         Left            =   9840
         TabIndex        =   36
         Top             =   540
         Width           =   390
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         BackStyle       =   0  'Transparent
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
         Left            =   9690
         TabIndex        =   35
         Top             =   1005
         Width           =   555
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search for... (Do not use wild cards (*) - this uses Google search engine)"
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
         Left            =   -74850
         TabIndex        =   30
         Top             =   405
         Width           =   7725
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
         Left            =   4065
         TabIndex        =   13
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
         Left            =   4725
         TabIndex        =   12
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
         TabIndex        =   10
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
         Left            =   840
         TabIndex        =   8
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
         Left            =   420
         TabIndex        =   7
         Top             =   510
         Width           =   360
      End
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4455
      Left            =   225
      OleObjectBlob   =   "frmBrowseProducts4a.frx":1654
      TabIndex        =   3
      Top             =   3135
      Width           =   11340
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10320
      Picture         =   "frmBrowseProducts4a.frx":6957
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
      Left            =   330
      OleObjectBlob   =   "frmBrowseProducts4a.frx":6CE1
      TabIndex        =   15
      Top             =   4245
      Visible         =   0   'False
      Width           =   11250
   End
   Begin TrueOleDBGrid60.TDBGrid GOO 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmBrowseProducts4a.frx":BFE1
      TabIndex        =   31
      Top             =   1725
      Width           =   11340
   End
   Begin VB.Label lblHUBRESULT 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   750
      Left            =   4740
      TabIndex        =   27
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
10        On Error GoTo errHandler
20        SaveLayout Me.GN, "SearchFormA"
30        SaveLayout Me.GBF, "SearchFormB"
40        SaveLayout Me.GOO, "SearchFormC"
50        SaveSetting "PBKS", Me.Name, "Formwidth", Me.Width
60     SetColumnWidths
70        Exit Sub
errHandler:
80        If ErrMustStop Then Debug.Assert False: Resume
90        ErrorIn "frmBrowseProducts.mnuSaveLayout"
End Sub

Private Sub SetMenu()
10        On Error GoTo errHandler

20        Forms(0).mnuSaveColumnWidths.Enabled = True
          
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmBrowseProducts.SetMenu"
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
10        On Error GoTo errHandler
      Dim frm As frmBICTree
      Dim strBICCode As String
20        Set frm = New frmBICTree
30        frm.Show vbModal
40        strBICCode = frm.SelectedCode
50        Unload frm
60        Set frm = Nothing
          
          
70        If strBICCode > "" Then
      '        Me.Refresh
      '        Screen.MousePointer = vbHourglass
      '        Screen.MousePointer = vbHourglass
      '        Me.Refresh
          '--------------
80            oPC.OpenDBSHort
          '--------------
90            search enSearchBIC, strBICCode
          '--------------
100           oPC.DisconnectDBShort
          '--------------
110       End If
120           Screen.MousePointer = vbDefault
          
130       Exit Sub
errHandler:
140       If ErrMustStop Then Debug.Assert False: Resume
150       ErrorIn "frmBrowseProducts.cmdBIC_Click", , EA_NORERAISE
160       HandleError
End Sub

Private Sub cmdCAT_Click()
10        On Error GoTo errHandler
20        Me.txtmaxnum = "9999999"
30        Screen.MousePointer = vbHourglass
40            search enSearchByCatalogue, cboCat
50        Screen.MousePointer = vbDefault
60        Exit Sub
          
70        Exit Sub
errHandler:
80        If ErrMustStop Then Debug.Assert False: Resume
90        ErrorIn "frmBrowseProducts.cmdCAT_Click", , EA_NORERAISE
100       HandleError
End Sub

Private Sub cmdSearch_Click()
10        On Error GoTo errHandler
          
20        cboSearch.AddItem cboSearch, 0
30        Screen.MousePointer = vbHourglass
40        cboSearch = FNS(cboSearch)
          
50        If cboSearch = "" And UCase(Me.cboCategory) = "<ALL>" And UCase(Me.cboProductType) = "<ALL>" Then
60            Screen.MousePointer = vbDefault
70            MsgBox "You must show what you are searching for.", vbOKOnly, "Enter a search request"
80            Exit Sub
90        End If
          
100       If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" Then
110           SetActiveGrid "GBF"
              DoEvents
120           GBF.SearchArguments = Left(cboSearch, Len(cboSearch) - 2)
130           GBF.search
          '    search enSearchBF, Left(cboSearch, Len(cboSearch) - 2)
140           mSetfocus GBF
150       Else
160           SetActiveGrid "GN"
170           search enSearchAdvanced, cboSearch, cboCategory, cboProductType
180           mSetfocus GN
190       End If
          
200       mSetfocus cboSearch
210       Screen.MousePointer = vbDefault
          
220       Exit Sub
errHandler:
230       If ErrMustStop Then Debug.Assert False: Resume
240       ErrorIn "frmBrowseProducts.cmdSearch_Click", , EA_NORERAISE
250       HandleError
End Sub

Private Sub search(pSearchType As enSearchType, pCriteria As String, Optional pSection As String, Optional pProductType As String)
10        On Error GoTo errHandler
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
20        errSysHandlerSet
30        Set rsResult = Nothing
40        Set rsResult = New ADODB.Recordset
50        rsResult.CursorLocation = adUseClient
60        rsResult.CursorType = adOpenStatic
          
70        txtRecsFound = ""
80        lngSectionID = 0
90        lngProductTypeID = 0
100       If pSearchType <> enSearchBIC Then
110           StripArticle pCriteria, strArticle, strNet
120           pCriteria = strNet
130       End If
          '--------------
140       oPC.OpenDBSHort
          '--------------
          
150       If pSection <> "<ALL>" Then
160           lngSectionID = oPC.Configuration.Sections.Key(pSection)
170       End If
180       If pProductType <> "<ALL>" Then
190           lngProductTypeID = oPC.Configuration.ProductTypes.Key(pProductType)
200       End If
              
210       lngMaxRecs = IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
220       If (Not Replace(pCriteria, "/", "") = "") Or lngSectionID > 0 Or lngProductTypeID > 0 Then
230           If pSearchType = enSearchByCatalogue Then
240               enSource = enLocalDB
250               oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "", (chkIncludeObsolete = 1)
260               oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
270           ElseIf pSearchType = enSearchBIC Then
280               enSource = enLocalDB
290               oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "", (chkIncludeObsolete = 1)
300               oSearchEngine.SearchBIC pCriteria, lngRecsFound
310           ElseIf pSearchType = enCourseCode Then
320               enSource = enLocalDB
330               oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "", (chkIncludeObsolete = 1)
340               oSearchEngine.SearchCourseCode pCriteria, lngRecsFound
'350           ElseIf pSearchType = enSearchBF Then
'360               oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "B", (chkIncludeObsolete = 1)
'370               enSource = enBF
'380               oSearchEngine.BFSearchEx pCriteria, lngRecsFound, CLng(txtmaxnum), lngResult
390           Else
400               enSource = enLocalDB
410              Call oSearchEngineC.SearchOnServer(pSearchType, oPC.WorkstationID, False, False, False, IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500), 1, _
                  (chkIncludeObsolete = 1), (Me.chkCopies = 1), pCriteria, lngRecsFound, rsResult, lngSectionID, _
                  lngProductTypeID, IIf(pSection = "<NONE>", True, False), IIf(pProductType = "<NONE>", True, False))
420               If rsResult Is Nothing Then
430                   lngRecsFound = 0
440               Else
450                   If rsResult.State = 0 Then
460                       lngRecsFound = 0
470                   Else
480                       If rsResult.eof Or rsResult.fields.Count = 1 Then
490                           lngRecsFound = 0
500                       Else
510                           lngRecsFound = rsResult.RecordCount
520                       End If
530                   End If
540               End If
550           End If
560       Else
570           lngRecsFound = 0
580       End If
                  
590       If lngRecsFound > 0 Then
600           If pSearchType = enSearchByCatalogue Or pSearchType = enSearchBIC Or pSearchType = enCourseCode Or pSearchType = enSearchBF Then
610               oSearchEngine.execute lngMaxRecs
620               Set colList = Nothing
630               Set colList = oSearchEngine.getcols
640               lngrows = oSearchEngine.rows
650           Else
660               oSearchEngineC.MassageRows rsResult
670               Set colList = Nothing
680               Set colList = oSearchEngineC.getcols
690               lngrows = oSearchEngineC.rows
700           End If
710       End If
720       txtRecsFound = CStr(lngRecsFound)
          
730       If lngRecsFound = 0 Then
740           Select Case enSource
              Case enLocalDB
750               XN.Clear
760               XN.ReDim 1, 1, 1, 16
770               XN(1, 1) = "No records"
780               GN.Array = XN
790               GN.ReBind
800           Case enBF
810               XBF.Clear
820               XBF.ReDim 1, 1, 1, 12
830               XBF(1, 1) = "No records"
840               GN.Array = XN
850             '  GBF.ReBind
860           End Select
870       Else
880           LoadGrid
890       End If
900       If lngRecsFound = IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500) Then
910           MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
920           Me.GN.Refresh
930       End If
          '--------------
940       oPC.DisconnectDBShort
          '--------------
950       Exit Sub
errHandler:
960       ErrPreserve
980       If Err.Number = -2147217407 Then   'Access violation
990           errRepeat = errRepeat + 1
1000          LogSaveToFile "Access violation in BrowseProducts: Search, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
1010          If errRepeat < 5 Then
1020              Resume Next
1030          Else
1040              LogSaveToFile "Access violation in BrowseProducts: Search after 5 re-attempts"
1050              MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't run search."
1060              Err.Clear
1070              Exit Sub
1080          End If
1090      End If
1100      If ErrMustStop Then Debug.Assert False: Resume
1110      ErrorIn "frmBrowseProducts.Search(pSearchType,pCriteria,pSection,pProductType)", _
               Array(pSearchType, pCriteria, pSection, pProductType)
End Sub
Private Sub LoadGridEx()
10        On Error GoTo errHandler
      Dim i As Long
      Dim XEX As New XArrayDB
20            XEX.ReDim 1, colList.Count, 1, 12
30            For i = 1 To colList.Count
40                    XEX.Value(i, 1) = colList.Item(i).CodeF
50                    XEX.Value(i, 2) = colList.Item(i).StatusF & " " & colList.Item(i).Title
60                    XEX.Value(i, 3) = colList.Item(i).Author
70                    XEX.Value(i, 4) = colList.Item(i).Distributor
80                    XEX.Value(i, 5) = colList.Item(i).QtyOnHand
90                    XEX.Value(i, 6) = colList.Item(i).QtyonOrder
100                   XEX.Value(i, 7) = colList.Item(i).QtyOnBackorder
110                   XEX.Value(i, 8) = colList.Item(i).QtyTotalSold
120                   XEX.Value(i, 10) = colList.Item(i).LastDateDelivered
130                   XEX.Value(i, 9) = colList.Item(i).LocalPriceF
140                   XEX.Value(i, 11) = colList.Item(i).PID
150                   XEX.Value(i, 12) = colList.Item(i).code
160           Next
170           Set GEX.Array = XEX
180           GEX.ReBind
190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.LoadGridEx"
End Sub

Private Sub LoadGrid()
10        On Error GoTo errHandler
      Dim i As Long

20        Screen.MousePointer = vbHourglass
30        Select Case enSource
          Case enLocalDB
40            GBF.Visible = False
50            GBF.Width = 0
60            GN.Visible = True
70            XN.Clear
           '   XBF.Clear
          '    GBF.ReBind
80            XN.ReDim 1, colList.Count, 1, 30
90            For i = 1 To colList.Count
100                   XN.Value(i, val(oColMap.Key("Code"))) = colList.Item(i).CodeF
110                   XN.Value(i, val(oColMap.Key("Item"))) = UCase(colList.Item(i).StatusShortF(True, True)) & " " & colList.Item(i).Title
120                   XN.Value(i, val(oColMap.Key("Author"))) = colList.Item(i).Author
130                   XN.Value(i, val(oColMap.Key("Distributor"))) = colList.Item(i).Distributor
140                   XN.Value(i, val(oColMap.Key("OH/OO/CO"))) = colList.Item(i).QtyOnHand & " / " & colList.Item(i).QtyonOrder & " / " & colList.Item(i).QtyOnBackorder
150                   XN.Value(i, val(oColMap.Key("Publisher"))) = colList.Item(i).Publisher
160                   XN.Value(i, val(oColMap.Key("PublicationDate"))) = colList.Item(i).PubDate & IIf(colList.Item(i).Edition > "", "/", "") & colList.Item(i).Edition
170                   XN.Value(i, val(oColMap.Key("TotalSold"))) = colList.Item(i).QtyTotalSold
180                   XN.Value(i, val(oColMap.Key("LastDateDelivered"))) = colList.Item(i).LastDateDelivered
190                   XN.Value(i, val(oColMap.Key("S.P."))) = colList.Item(i).LocalPriceF
200                   XN.Value(i, val(oColMap.Key("Multibuy"))) = colList.Item(i).Multibuy
210                   XN.Value(i, val(oColMap.Key("Categories"))) = colList.Item(i).Categories
220                   XN.Value(i, 11) = colList.Item(i).PID
230                   XN.Value(i, 12) = colList.Item(i).code
240                   XN.Value(i, 13) = colList.Item(i).EAN
250                   XN.Value(i, 16) = colList.Item(i).LocalPrice
260                   XN.Value(i, 17) = colList.Item(i).Obsolete
270                   XN.Value(i, 18) = colList.Item(i).Title
280                   XN.Value(i, 20) = colList.Item(i).QtyOnHand
290                   XN.Value(i, 21) = colList.Item(i).QtyonOrder
300                   XN.Value(i, 22) = colList.Item(i).QtyOnBackorder
310           Next
320           XN.QuickSort 1, XN.UpperBound(1), 18, XORDER_ASCEND, XTYPE_STRING
330           GN.Array = XN
340           Me.GN.ReBind
              
              
              
350       Case enBF
360           XN.Clear
370           GN.ReBind
380           XBF.Clear
390           GBF.Visible = True
400           GN.Visible = False
410           XBF.ReDim 1, colList.Count, 1, 12
420           For i = 1 To colList.Count
440                       XBF.Value(i, 1) = colList.Item(i).EAN
480                   XBF.Value(i, 2) = colList.Item(i).Title
490                   XBF.Value(i, 3) = colList.Item(i).Author
500                   XBF.Value(i, 4) = IIf(colList.Item(i).DistributorByIdx(1) = "", "Pub by:" & colList.Item(i).Publisher, colList.Item(i).DistributorByIdx(1))
510                   XBF.Value(i, 5) = colList.Item(i).LocalPriceF
520                   XBF.Value(i, 6) = colList.Item(i).USPriceF
530                   XBF.Value(i, 7) = colList.Item(i).UKPriceF
540                   XBF.Value(i, 8) = colList.Item(i).DistributorCode & " : " & colList.Item(i).Distributor
550                   XBF.Value(i, 12) = colList.Item(i).code
560           Next
570           XBF.QuickSort 1, XBF.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
580           GN.Array = XN
590           Me.GN.ReBind
           '   GBF.Array = XBF
           '   Me.GBF.ReBind
600       End Select
610       Screen.MousePointer = vbDefault
620       Exit Sub
errHandler:
630       If ErrMustStop Then Debug.Assert False: Resume
640       ErrorIn "frmBrowseProducts.LoadGrid"
End Sub

Private Sub cmdClose_Click()
10        On Error GoTo errHandler
      'Unload Me
20        Me.Hide
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmBrowseProducts.cmdClose_Click", , EA_NORERAISE
60        HandleError
End Sub

Private Sub cmdCourseCodes_Click()
10        On Error GoTo errHandler
20        Me.txtmaxnum = "9999999"
30        If Me.Text1 > "" Then
40            Screen.MousePointer = vbHourglass
50            search enCourseCode, Me.Text1
60            Screen.MousePointer = vbDefault
70        End If
80        Exit Sub

90        Exit Sub
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "frmBrowseProducts.cmdCourseCodes_Click", , EA_NORERAISE
120       HandleError
End Sub

Private Sub cmdDebugOff_Click()
10        On Error GoTo errHandler
20    oSearchEngine.mbDebug = False
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmBrowseProducts.cmdDebugOff_Click", , EA_NORERAISE
60        HandleError
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
10        On Error GoTo errHandler
      Dim ob As New z_SQL

20        Screen.MousePointer = vbHourglass
          
30        ob.RunProc "dbo.SENDBROKERMESSAGE", Array(Me.cboSearch), "TEST"
          
40        Screen.MousePointer = vbDefault
          
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.cmdGetFromSB_Click", , EA_NORERAISE
80        HandleError
End Sub

Private Sub cmdGOO_Click()
10        On Error GoTo errHandler
20        If Inet1.StillExecuting Then Exit Sub
30        Screen.MousePointer = vbHourglass
40        GoogleButtonNo = 1
50        GooIndex = 0
60        Set XGOO = New XArrayDB
70        cboGoogle.AddItem cboGoogle, 0
80        FetchFromGoogle
90        Exit Sub
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "frmBrowseProducts.cmdGOO_Click", , EA_NORERAISE
120       HandleError
End Sub

Private Sub cmdGOO_LostFocus()
10        On Error GoTo errHandler
          
20        If Inet1.StillExecuting And GoogleButtonNo = 1 Then
30            cmdGOO.SetFocus
40        End If
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.cmdGOO_LostFocus", , EA_NORERAISE
80        HandleError
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
10        On Error GoTo errHandler
      Dim se As New z_SearchEngineB
      Dim SS As String
      Dim r As String
      Dim oG As New z_Google
20        SS = se.GoogleSearchString(Me.cboGoogle)
30        If bISBNOnly Then
40            r = "+date:1950-2008"
50        Else
60            r = ""
70        End If
80        Inet1.URL = "http://books.google.com/books/feeds/volumes?q=" & Replace(SS, """", "%22") & r & "&max-results=20&start-index=" & CStr(GooIndex + 1)

90        strGOOXML = Inet1.OpenURL
100       If Left(strGOOXML, 7) = "invalid" Then
110          ' MsgBox strGOOXML
120           Exit Sub
130       End If
140       If strGOOXML > "" Then
150           oG.LoadFromGoogle strGOOXML, XGOO, GooIndex
160           Set GOO.Array = XGOO
170           GOO.Refresh
180           GOO.ReBind
190           GOO.ZOrder 0
200       End If
210       Screen.MousePointer = vbDefault
220       Exit Sub
errHandler:
230       If ErrMustStop Then Debug.Assert False: Resume
240       ErrorIn "frmBrowseProducts.FetchFromGoogle"
End Sub
Private Function GetISBN(p As String) As String
10        On Error GoTo errHandler
      Dim a() As String
20        GetISBN = ""
30        a = Split(p, ":")
40        If UBound(a) > 0 Then
50            If a(0) = "ISBN" Then
60                GetISBN = a(1)
70            End If
80        End If
90        Exit Function
errHandler:
100       If ErrMustStop Then Debug.Assert False: Resume
110       ErrorIn "frmBrowseProducts.GetISBN(p)", p
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
10        On Error GoTo errHandler
      Dim oTF As New z_TextFile
      Dim s As String
      Dim s2 As String
      Dim lngNumberOfLines As Long
      Dim i As Long
      Dim fs As New FileSystemObject

20        ExportToSpreadsheet = False
          
30        If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
40            fs.CreateFolder (oPC.LocalFolder & "\TEMP")
50            If Err <> 0 Then
60                MsgBox "Cannot create folder TEMP on local computer", vbInformation + vbOKOnly, "Can't do this"
70            End If
80        End If
          
90        pFilename = oPC.LocalFolder & "Temp\BrowseProducts_" & Format(Now(), "yyyymmddHHnn") & ".xls"
          
100       oTF.OpenTextFile pFilename

110       s = "SKU" & vbTab & "Title" & vbTab & "Author" & vbTab & "Distributor" & vbTab & "Publisher" & vbTab & "Pub Date" & vbTab & "Total sold" & vbTab & "Last Del." & vbTab _
              & "S.P." & vbTab & "Categories" & vbTab & "Obsolete" & vbTab & "On hand" & vbTab _
              & "QtyonOrder" & vbTab & "QtyOnBackorder" & vbTab & "QtyReserved"
          
120       oTF.WriteToTextFile s
                     
                      
130       lngNumberOfLines = 0
140       For i = 1 To XN.Count(1)
150           lngNumberOfLines = lngNumberOfLines + 1
160           s = XN(i, val(oColMap.Key("Code")))
170           s = s & vbTab & XN(i, val(oColMap.Key("Item")))
180           s = s & vbTab & XN(i, val(oColMap.Key("Author")))
190           s = s & vbTab & XN(i, val(oColMap.Key("Distributor")))
200           s = s & vbTab & XN(i, val(oColMap.Key("Publisher")))
210           s = s & vbTab & XN(i, val(oColMap.Key("PublicationDate")))
220           s = s & vbTab & XN(i, val(oColMap.Key("TotalSold")))
230           s = s & vbTab & XN(i, val(oColMap.Key("LastDateDelivered")))
240           s = s & vbTab & XN(i, val(oColMap.Key("S.P.")))
           '   s = s & vbTab & XN(i, val(oColMap.key("Multibuy")))
250           s = s & vbTab & XN(i, val(oColMap.Key("Categories")))
260           s = s & vbTab & XN(i, 17)
270           s = s & vbTab & XN(i, 20)
280           s = s & vbTab & XN(i, 21)
290           s = s & vbTab & XN(i, 22)
300           s = s & vbTab & XN(i, 23)
310           oTF.WriteToTextFile s
320       Next
330       oTF.CloseTextFile
340       ExportToSpreadsheet = True
          
350       Exit Function
errHandler:
360       ErrPreserve
370       If Err.Number = -2147217407 Then   'Access violation
380           errRepeat = errRepeat + 1
390           LogSaveToFile "Access violation in BrowseProducts: ExportToSpreadsheet"  'unknown source
400           If errRepeat < 5 Then
410               Err.Clear
420               Exit Function
430           Else
440               LogSaveToFile "Access violation in BrowseProducts: ExportToSpreadsheet after 5 re-attempts"
450               MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
460               Err.Clear
470               Exit Function
480           End If
490       End If
500       If ErrMustStop Then Debug.Assert False: Resume
510       ErrorIn "frmBrowseProducts.ExportToSpreadsheet(pFilename)", pFilename
End Function

Private Sub cmdPrintbarcode_Click()
10        On Error GoTo errHandler
          Dim ar As New arPrintBarcodeList
          
20        ar.component XN
30        ar.Show vbModal
40        Set ar = Nothing

50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.cmdPrintbarcode_Click", , EA_NORERAISE
80        HandleError
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
10        On Error GoTo errHandler
      Dim str As String
20        If oPC.InternetDialup = True Then Exit Sub
30        Screen.MousePointer = vbHourglass
40        If cboSearch.Text = "" Then Exit Sub
50        If IsNumeric(Left(Me.cboSearch.Text, 9)) Then
60            If IsNumeric(Left(Me.cboSearch.Text, 13)) Then
70                OpenBrowser "http://books.google.com/books?isbn=" & Left(Me.cboSearch.Text, 13)
80            Else
90                OpenBrowser "http://books.google.com/books?isbn=" & Left(Me.cboSearch.Text, 10)
100           End If
110       Else
120           str = Replace(FNS(cboSearch), "/", "")
130           OpenBrowser "http://books.google.com/books?q=" & str
140       End If
150       Screen.MousePointer = vbDefault
160       Exit Sub
errHandler:
170       If ErrMustStop Then Debug.Assert False: Resume
180       ErrorIn "frmBrowseProducts.Command1_Click", , EA_NORERAISE
190       HandleError
End Sub

Private Sub Command2_Click()
10        On Error GoTo errHandler
      Dim str As String
      Dim str2 As String
20        If oPC.InternetDialup = True Then Exit Sub
30        Screen.MousePointer = vbHourglass
40        If cboSearch.Text = "" Then Exit Sub
50        If IsNumeric(Left(Me.cboSearch.Text, 9)) Then
60            If IsNumeric(Left(Me.cboSearch.Text, 13)) Then
70                str = "http://www.amazon.co.uk/dp/XXX"
80                str = Replace(str, "XXX", Left(Me.cboSearch.Text, 10))
90                OpenBrowser str
100           Else
110               str = "http://www.amazon.co.uk/dp/XXX"
120               str = Replace(str, "XXX", Left(Me.cboSearch.Text, 10))
130               OpenBrowser str
140           End If
150       Else
              
160           str = "http://www.amazon.co.uk/gp/search?search-alias=stripbooks&field-keywords=&author=&select-author=field-author-like&title=XXX&select-title=field-title&subject=&select-subject=field-subject&field-publisher=&field-isbn=&chooser-sort=rank%21%2Bsalesrank&node=&field-binding=&mysubmitbutton1.x=53&mysubmitbutton1.y=12"
170           str2 = Replace(FNS(cboSearch), "/", "")
180           str = Replace(str, "XXX", str2)
190           OpenBrowser str
200       End If
210       Screen.MousePointer = vbDefault

220       Exit Sub
errHandler:
230       If ErrMustStop Then Debug.Assert False: Resume
240       ErrorIn "frmBrowseProducts.Command2_Click", , EA_NORERAISE
250       HandleError
End Sub




Private Sub Form_Activate()
10        On Error GoTo errHandler
20 ' MsgBox "aCTIVATE 1"
30        SetMenu
40        p 1
50        p 2
60        p 3
70  '  MsgBox "aCTIVATE 2"
80        bWithCopies = False
90        chkCopies = IIf(bWithCopies, 1, 0)
100       Me.Command1.Enabled = Not oPC.InternetDialup
110 '  MsgBox "aCTIVATE 3"
120       cmdGetFromSB.Visible = oPC.GetProperty("UsesHUB") = "TRUE"
130       lblHUBRESULT.Visible = cmdGetFromSB.Visible
140 '  MsgBox "aCTIVATE 4"
150      Exit Sub
errHandler:
160       If ErrMustStop Then Debug.Assert False: Resume
170       ErrorIn "frmBrowseProducts.Form_Activate", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
180       HandleError
End Sub
Private Sub SetColumnWidths()
10        On Error GoTo errHandler
      Dim i As Integer
      '    For i = 1 To GBF.Columns.Count
      '        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
      '    Next
      '    GBF.Columns(GBF.Columns.Count - 1).Width = GBF.Columns(GBF.Columns.Count - 1).Width * 0.8
'MsgBox "column count = "
'MsgBox "GN is null" & (GN Is Nothing)
'MsgBox "GN columns is null" & (GN.Columns Is Nothing)
'      MsgBox "column count = " & GN.Columns.Count
'      MsgBox "GetSetting = " & GetSetting("PBKS", "SearchFormA", CStr(0), GN.Columns(0).Width)
'20        For i = 1 To GN.Columns.Count
'30            GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
'40        Next
'50        GN.Columns(GN.Columns.Count - 1).Width = GN.Columns(GN.Columns.Count - 1).Width * 0.8
'
'60        For i = 1 To GOO.Columns.Count
'70            GOO.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormC", CStr(i), GOO.Columns(i - 1).Width)
'80        Next
'90        GOO.Columns(GOO.Columns.Count - 1).Width = GOO.Columns(GOO.Columns.Count - 1).Width * 0.8
100       Exit Sub
errHandler:

110       If ErrMustStop Then Debug.Assert False: Resume
120       ErrorIn "frmBrowseProducts.SetColumnWidths"
End Sub


Private Sub Form_Deactivate()
10        On Error GoTo errHandler
20        UnsetMenu
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmBrowseProducts.Form_Deactivate", , EA_NORERAISE
60        HandleError
End Sub

Private Sub Form_Initialize()
10        On Error GoTo errHandler

20        InitializeGrids
30      '  cboSearch.SetFocus
          
40        Exit Sub
errHandler:
50        If ErrMustStop Then Debug.Assert False: Resume
60        ErrorIn "frmBrowseProducts.Form_Initialize", , EA_NORERAISE
70        HandleError
End Sub
Private Sub InitializeGrids()
10        On Error GoTo errHandler
      Dim i As Integer
   '   MsgBox "pOS 1"
20        oColMap.Load oPC.GetProperty("BrowsePosition")
  '  MsgBox "pOS 2"

30        Set XN = New XArrayDB
40      '  Set XBF = New XArrayDB
50        Set XA = New XArrayDB
60        Set XGOO = New XArrayDB
          
70        Set oSearchEngine = New z_SearchEngineB
80        Set oSearchEngineC = New z_SearchEngineC
          
90        Set colList = New Collection
      '    If oPC.Configuration.AntiquarianYN Then
      '        chkAntiquarianOnly.Visible = True
      '        chkAntiquarianOnly = 1
      '    Else
      '        chkAntiquarianOnly = 0
      '        chkAntiquarianOnly.Visible = False
      '    End If
100       SetColumnWidths
110   '    Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
         ' XN.ReDim 1, 1, 1, 9
120       XA.ReDim 1, 0, 1, 9
130     '  XBF.ReDim 1, 1, 1, 12
140       XGOO.ReDim 1, 1, 1, 12
150    '   SSTab1.Tab = 0
160       mSetfocus cboSearch
170     ''''^  GN.ZOrder 0
180     '  FormatGN
   '   MsgBox "pOS 3"

190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.InitializeGrids"
End Sub
Private Sub FormatGN()
10        On Error GoTo errHandler
      Dim i As Integer

      Dim lngAlignment As Long

20        For i = 0 To oColMap.Count - 1
30            Select Case oColMap.Item_2(i)
              Case "L"
40                lngAlignment = 0
50            Case "R"
60                lngAlignment = 1
70            Case "C"
80                lngAlignment = 2
90            End Select
100           If CLng((oColMap.Item_3(i))) <= GN.Columns.Count Then
110               GN.Columns(CLng((oColMap.Item_3(i))) - 1).Alignment = lngAlignment
120               GN.Columns(CLng((oColMap.Item_3(i))) - 1).Caption = oColMap.Item_1(i)
130           End If
140       Next
150       Exit Sub
errHandler:
160       If ErrMustStop Then Debug.Assert False: Resume
170       ErrorIn "frmBrowseProducts.FormatGN"
End Sub

Private Sub Form_Load()
10        On Error GoTo errHandler
      Dim i As Integer
        ' MsgBox "pos A0"

20         errSysHandlerSet
       '  MsgBox "pos A1"

30        Set PrivateCnn = New ADODB.Connection
40        oPC.OpenSuppliedConnection PrivateCnn

50        GN.Left = 120
60        GBF.Left = 120
70        GOO.Left = 120
80        GN.TOP = 1727
90        GBF.TOP = 1725
100       GOO.TOP = 1727
110       XA.Clear
120       XA.ReDim 1, 0, 1, 9
130      ' XBF.Clear
140     '  XBF.ReDim 1, 1, 1, 12
150       XGOO.Clear
160       XGOO.ReDim 1, 1, 1, 12
170       flgLoading = True
180       Resize_GN
        '  MsgBox "pos A2"
         
190       Me.Width = GetSetting("PBKS", Me.Name, "Formwidth", Me.Width)
200       flgLoading = False
210       OriginalwdthGN = 250
220       For i = 1 To GN.Columns.Count - 1
230           OriginalwdthGN = OriginalwdthGN + GN.Columns(i - 1).Width
240           wdthCol(i - 1) = GN.Columns(i - 1).Width
250       Next
260       If Me.WindowState <> 2 Then
270           Me.TOP = 20
280           Me.Left = 50
290           Height = 7220
300       End If
       '   MsgBox "pos A3"
         
310       Set tlCats = Nothing
320       Set tlCats = New z_TextList
          
330       If oPC.SupportsCatalogue Then
340           tlCats.Load ltCatalogue
350           LoadCombo cboCat, tlCats
360       End If
370       frCatalogue.Enabled = oPC.SupportsCatalogue
          
380       LoadCombo cboCategory, oPC.Configuration.Sections
390       Me.cboCategory.AddItem "<NONE>"
          
400       LoadCombo cboProductType, oPC.Configuration.ProductTypes
410       cboProductType.AddItem "<NONE>"
          
420       If oPC.Configuration.AntiquarianYN Then
430           Me.GN.Columns(3).Caption = "Publisher"
440       Else
450           GN.Columns(3).Caption = "Distributor"
460       End If
         ' GBF.Columns(3).Caption = "Distributor"
470       txtmaxnum = 500
       '  MsgBox "pos A4"
          
      '    If oPC.Configuration.AntiquarianYN Then
      '        chkAntiquarianOnly.Visible = True
      '        chkAntiquarianOnly = 1
      '    Else
      '        chkAntiquarianOnly = 0
      '        chkAntiquarianOnly.Visible = False
      '    End If

480       Me.chkIncludeObsolete = IIf(GetSetting("PBKS", Me.Name, "IncludeObsolete", 0) = 1, CInt("1"), CInt("0"))
          
490       On Error Resume Next
500       cboCategory = "<ALL>"
510       cboProductType = "<ALL>"
       '   MsgBox "pos A5"
         
520       SetActiveGrid "GN"
       '  MsgBox "pos A6"

530       Exit Sub
errHandler:
540       If ErrMustStop Then Debug.Assert False: Resume
550       ErrorIn "frmBrowseProducts.Form_Load", , EA_NORERAISE
560       HandleError
End Sub
Private Sub Resize_GBF()
Dim i As Long
'MsgBox
'    For i = 1 To GBF.Columns.Count
'        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
'    Next
End Sub
Private Sub Resize_GOO()
10        On Error GoTo errHandler
      Dim i As Long
20        For i = 1 To GOO.Columns.Count
30            GOO.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormC", CStr(i), GOO.Columns(i - 1).Width)
40        Next
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.Resize_GOO"
End Sub
Private Sub Resize_GN()
10        On Error GoTo errHandler
      Dim i As Long
20        For i = 1 To GN.Columns.Count
30            GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
40        Next
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.Resize_GN"
End Sub

Private Sub Form_Resize()
10        On Error GoTo errHandler
      Dim lngDiff As Long
20    'MsgBox "Resize Pos 1"
30        If flgLoading Then Exit Sub
40        GN.Width = NonNegative_Lng(Me.Width - 380)
50        GBF.Width = GN.Width
60        GOO.Width = GN.Width
70    'MsgBox "Resize Pos 2"
          
80        lngDiff = GN.Height
90        GN.Height = NonNegative_Lng(Me.Height - (GN.TOP + 1070))
100       GBF.Height = GN.Height
110       GOO.Height = GN.Height
120   'MsgBox "Resize Pos 3"
130       lngDiff = GN.Height - lngDiff
          
140       cmdPrint.TOP = GN.TOP + GN.Height + 30
150       cmdClose.TOP = GN.TOP + GN.Height + 30
160       cmdPrintbarcode.TOP = GN.TOP + GN.Height + 30
170       cmdClose.Left = NonNegative_Lng(GN.Width - 1000)
180   'MsgBox "Resize Pos 4"
         ' ResizeColumns
             ' FormatGN

190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.Form_Resize", , EA_NORERAISE
220       HandleError
End Sub
Sub ResizeColumns()
10        On Error GoTo errHandler
      Dim i As Integer
      Dim newwidth As Long
20    'MsgBox "ResizeColumns Pos 1"

30        For i = 1 To GN.Columns.Count - 1
40            GN.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GN.Width) / OriginalwdthGN) * 0.9)
50        Next
60        For i = 1 To GOO.Columns.Count - 1
70            GOO.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GOO.Width) / OriginalwdthGN) * 0.9)
80        Next
90    'MsgBox "ResizeColumns Pos 2"

100       Exit Sub
errHandler:
110       If ErrMustStop Then Debug.Assert False: Resume
120       ErrorIn "frmBrowseProducts.ResizeColumns"
End Sub
Private Sub Form_Terminate()
10        On Error GoTo errHandler
20        Set oSearchEngine = Nothing
30        Set oProduct = Nothing
40        Set colList = Nothing
50        Set tlkeys = Nothing
60        Set lslist = Nothing
70        Set XN = Nothing
80        Set XBF = Nothing
90        Set XA = Nothing
100       Set XGOO = Nothing
110       Exit Sub
errHandler:
120       If ErrMustStop Then Debug.Assert False: Resume
130       ErrorIn "frmBrowseProducts.Form_Terminate", , EA_NORERAISE
140       HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
10        On Error GoTo errHandler
20        UnsetMenu
30        SaveSetting "PBKS", Me.Name, "IncludeObsolete", IIf(Me.chkIncludeObsolete = 1, "1", "0")
      '---------------------------------------------------
40        If PrivateCnn Is Nothing Then Exit Sub
50        oPC.CloseSUppliedConnection PrivateCnn
      '---------------------------------------------------

60        Exit Sub
errHandler:
70        If ErrMustStop Then Debug.Assert False: Resume
80        ErrorIn "frmBrowseProducts.Form_Unload(Cancel)", Cancel, EA_NORERAISE
90        HandleError
End Sub

Private Sub GBF_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
10        On Error GoTo errHandler
      Dim str As String
          
      'MsgBox "GBF_RowColChange Pos 1"
    
20        If XBF Is Nothing Then Exit Sub
30        If XBF.UpperBound(1) = 0 Then Exit Sub
40        If Err Then Exit Sub
50     '   If IsNull(GBF.Bookmark) Then Exit Sub
60        If Err Then Exit Sub
          
          'MsgBox "Code Commented"
70    '    str = FNS(XBF.Value(GBF.Bookmark, 1))
80        If str = "" Then Exit Sub
90        On Error Resume Next
100       Clipboard.Clear
110       Clipboard.SetText Left(str, ISBNLENGTH)
120       Exit Sub
errHandler:
130       If ErrMustStop Then Debug.Assert False: Resume
140       ErrorIn "frmBrowseProducts.GBF_Click", , EA_NORERAISE, , "Line number", Array(Erl())
150       HandleError
End Sub

Private Sub GBF_DblClick()
10        On Error GoTo errHandler
      Dim oProd As a_Product
      Dim str As String
      Dim frm As frmProduct
20              'gBox "Code Commented"

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
30        Exit Sub
errHandler:
40        ErrPreserve
50        If Err.Number = -2147217407 Then   'Access violation
60            errRepeat = errRepeat + 1
70            LogSaveToFile "Access violation in BrowseProducts: GBF_DblCLick"  'unknown source
80            If errRepeat < 5 Then
90                Resume Next
100           Else
110               LogSaveToFile "Access violation in BrowseProducts: GBF_DblCLick after 5 re-attempts"
120               MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
130               Err.Clear
140               Exit Sub
150           End If
160       End If
170       If ErrMustStop Then Debug.Assert False: Resume
180       ErrorIn "frmBrowseProducts.GBF_DblClick", , EA_NORERAISE
190       HandleError
End Sub



Private Sub GBF_KeyPress(KeyAscii As Integer)
10        On Error GoTo errHandler
20        If KeyAscii = vbKeyReturn Then
30            GBF_DblClick
40        End If
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.GBF_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
80        HandleError
End Sub






Private Sub GN_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
10        On Error GoTo errHandler
20        If XN.Value(1, 5) = "" Then Exit Sub
30        If XN(Bookmark, 17) = True Then
40            RowStyle.Font.Strikethrough = True
50        End If
60        If Len(XN(Bookmark, 2)) > 1 Then
70            If Left(XN(Bookmark, 2), 2) = "**" Then
80                RowStyle.BackColor = RGB(220, 220, 220)
90            End If
100       End If
110       If Len(XN(Bookmark, 2)) > 5 Then
120           If Left(XN(Bookmark, 2), 6) = "**(OP)" Then
130               RowStyle.BackColor = RGB(180, 180, 180)
140           End If
150       End If
160       Exit Sub
errHandler:
170       If ErrMustStop Then Debug.Assert False: Resume
180       ErrorIn "frmBrowseProducts.GN_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
               RowStyle), EA_NORERAISE
190       HandleError
End Sub

Private Sub GOO_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
10        On Error GoTo errHandler
      Dim str As String
      'MsgBox "GOO_RowColChange Pos 1"

20        On Error Resume Next
30        If XGOO Is Nothing Then Exit Sub
40        If XGOO.Count(1) = 0 Then Exit Sub
50        If XGOO.UpperBound(1) = 0 Then Exit Sub
60        If Err Then Exit Sub
70        If IsNull(GOO.Bookmark) Then Exit Sub
80        On Error GoTo errHandler
          
          On Error Resume Next
90        str = FNS(XGOO.Value(GOO.Bookmark, 1))
          On Error GoTo errHandler
100       If str = "" Then Exit Sub
          
110       Clipboard.Clear
120       Clipboard.SetText Left(str, ISBNLENGTH)
130       Exit Sub

140       Exit Sub
errHandler:
150       ErrorIn "frmBrowseProducts.GOO_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE, , "Line number", Array(Erl())
160       HandleError
          
End Sub



Private Sub GN_KeyDown(KeyCode As Integer, Shift As Integer)
10        On Error GoTo errHandler
20        If (KeyCode = vbKeyLeft) Then
30            mSetfocus cboSearch
40        End If
50        If Shift = 1 Then
60            bShiftDown = True
70        End If

80        Exit Sub
errHandler:
90        If ErrMustStop Then Debug.Assert False: Resume
100       ErrorIn "frmBrowseProducts.GN_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
110       HandleError
End Sub

Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
10        On Error GoTo errHandler
Dim errRepeat As Integer
Dim str As String
      'MsgBox "GN_RowColChange Pos 1"

20        errRepeat = 0
30
40        On Error Resume Next
50        If LastRow = "" Then Exit Sub
60        If XN.Count(1) = 0 Then Exit Sub
70        If GN.VisibleRows < 1 Then Exit Sub
80        If IsNull(GN.Bookmark) Then Exit Sub
90        If Err Then Exit Sub
100   On Error GoTo errHandler
110
120       If IsNumeric(GN.Bookmark) Then
130           If GN.Bookmark <= XN.UpperBound(1) Then
140               str = IIf(FNS(XN.Value(GN.Bookmark, 13)) > "", FNS(XN.Value(GN.Bookmark, 13)), FNS(XN.Value(GN.Bookmark, 12)))
150           End If
160       End If
170
180       If str = "" Then Exit Sub
          
190       On Error Resume Next
200       Clipboard.Clear
210       Clipboard.SetText Left(str, ISBNLENGTH)
220       Exit Sub
errHandler:
230       ErrPreserve
240       If Err.Number = -2147217407 Or Err.Number = 2147227667 Then    'Access violation
250     errRepeat = errRepeat + 1
260     LogSaveToFile "Access violation in BrowseProducts: GN_RowColChanged"  'unknown source
270     If errRepeat < 5 Then
280         Resume Next
290     Else
300         LogSaveToFile "Access violation in BrowseProducts: GN_RowColChanged after 5 re-attempts"
310         MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
320         Err.Clear
330         Exit Sub
340     End If
350       End If
360       If ErrMustStop Then Debug.Assert False: Resume
370       ErrorIn "frmBrowseProducts.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
               EA_NORERAISE, , "strErrPos,Line number", Array(strErrPos, Erl())
380       HandleError
End Sub
Public Property Get NextPID() As String
10        On Error GoTo errHandler
20        If GN.Array Is Nothing Then
30            NextPID = ""
40            Exit Property
50        End If
60        If BookmarkPointer < GN.Array.UpperBound(1) Then
70            BookmarkPointer = BookmarkPointer + 1
80            NextPID = FNS(XN.Value(BookmarkPointer, 11))
90        Else
100           NextPID = ""
110       End If
120       Exit Property
errHandler:
130       If ErrMustStop Then Debug.Assert False: Resume
140       ErrorIn "frmBrowseProducts.NextPID"
End Property
Public Property Get PrevPID() As String
10        On Error GoTo errHandler
20        If GN.Array Is Nothing Then
30            PrevPID = ""
40            Exit Property
50        End If
60        If BookmarkPointer > 1 Then
70            BookmarkPointer = BookmarkPointer - 1
80            PrevPID = FNS(XN.Value(BookmarkPointer, 11))
90        Else
100           PrevPID = ""
110       End If
              
120       Exit Property
errHandler:
130       If ErrMustStop Then Debug.Assert False: Resume
140       ErrorIn "frmBrowseProducts.PrevPID"
End Property
Private Sub GN_DblClick()
    OpenInventoryRecord
End Sub

Public Sub OpenInventoryRecord()
10        On Error GoTo errHandler
      Dim frmA As frmProductPrevAQ
      Dim frm As frmProductPrev
      Dim frmNB As frmProductNBPrev   'non book form
      Dim lngprod As Long
      Dim errRepeat As Integer

20        errSysHandlerSet

30        errRepeat = 0
40        If XN.Count(1) = 0 Then
50            LogSaveToFile "XN.Count(1) = 0"
60            Exit Sub
70        End If
80        If IsNull(GN.Bookmark) Then
90            LogSaveToFile "IsNull(GN.Bookmark) is true"
100           Exit Sub
110       End If
120       If GN.Bookmark > XN.Count(1) Then
130         LogSaveToFile "GN.Bookmark > XN.Count(1)"
140         Exit Sub
150       End If
160       BookmarkPointer = GN.Bookmark
170       strPID = FNS(XN.Value(GN.Bookmark, 11))
180       If strPID = "" Then
190           LogSaveToFile "strPID = ''"
200           Exit Sub
210       End If
220       If bShiftDown Then
230           ShowSalesPatterns
240       Else
250           Set oProduct = New a_Product
260           Screen.MousePointer = vbHourglass
270           oProduct.Load strPID, 0, "", strTime
280           If oProduct.PID = "" Then Exit Sub
290           If oProduct.ProductType = "B" Then
300               If oPC.Configuration.AntiquarianYN Then
310                     Set frmA = Nothing
320                   Set frmA = New frmProductPrevAQ
330                   frmA.component oProduct
340                   frmA.Show
350               Else
360                     Set frm = Nothing
370                   Set frm = New frmProductPrev
380                   frm.component oProduct, strTime
390                   frm.Show
400               End If
410           Else
420               Set frm = Nothing
430               Set frmNB = New frmProductNBPrev
440               frmNB.component oProduct, strTime
450               frmNB.Show
460           End If
470       End If
480       Set oProduct = Nothing
490       Screen.MousePointer = vbDefault
500       bShiftDown = False
510       Exit Sub
errHandler:
520       ErrPreserve
530       Set frm = Nothing
540       If Err.Number = -2147217407 Then   'Access violation
550           errRepeat = errRepeat + 1
560           LogSaveToFile "Access violation in BrowseProducts: GN_DblCLick, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
570           If errRepeat < 5 Then
580                 Set frm = Nothing
590                 Err.Clear
600                 Exit Sub
610           Else
620               LogSaveToFile "Access violation in BrowseProducts: GN_DblCLick after 5 re-attempts"
630               MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
640               Err.Clear
650               Exit Sub
660           End If
670       End If
680       If ErrMustStop Then Debug.Assert False: Resume
690       ErrorIn "frmBrowseProducts.OpenInventoryRecord", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
700       HandleError
End Sub
Public Sub ShowSalesPatterns()
10        On Error GoTo errHandler
      Dim frmSales As frmSalesCH
20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub

40        Screen.MousePointer = vbHourglass
50        Set oProduct = New a_Product
60        strPID = FNS(XN.Value(GN.Bookmark, 11))
70        If strPID = "" Then Exit Sub

80        oProduct.Load strPID, 0
90        If oProduct.PID = "" Then Exit Sub
100       Set frmSales = New frmSalesCH
110       frmSales.component oProduct
120       frmSales.Show
130       Screen.MousePointer = vbDefault
140       Set frmSales = Nothing
150       Exit Sub
errHandler:
160       If ErrMustStop Then Debug.Assert False: Resume
170       ErrorIn "frmBrowseProducts.ShowSalesPatterns"
End Sub
Private Sub GN_HeadClick(ByVal ColIndex As Integer)
10        On Error GoTo errHandler
      Static Direction As Variant
20        If XN.Count(1) = 0 Then Exit Sub
          If XN(1, 1) = "No records" Then Exit Sub
30        If XN.UpperBound(1) = 0 Then Exit Sub
40        If Direction = 0 Then
50            Direction = 1
60        Else
70            Direction = 0
80        End If
90        If ColIndex = 0 Then ColIndex = 12
100       If ColIndex = 1 Then ColIndex = 18
110           XN.QuickSort XN.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
       '   Else
       '       XN.QuickSort XA.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
       '   End If
          
120       GN.Refresh
130       Exit Sub
errHandler:
140       If ErrMustStop Then Debug.Assert False: Resume
150       ErrorIn "frmBrowseProducts.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE, , "colcount", CLng(XN.Count(1))
160       HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
10        On Error GoTo errHandler
        '  Select Case ColIndex
20            If ColIndex > 10 Then
30                GetRowType = XTYPE_STRING
40            Else
50            If GN.Columns(ColIndex - 1).Alignment = dbgRight Then
60                GetRowType = XTYPE_INTEGER
70            Else
80                GetRowType = XTYPE_STRING
90            End If
100           End If
      '        Case 1, 2, 3, 4, 12
      '            GetRowType = XTYPE_STRING
      '        Case 5, 6, 7, 8, 9
      '            GetRowType = XTYPE_INTEGER
         ' End Select
110       Exit Function
errHandler:
120       If ErrMustStop Then Debug.Assert False: Resume
130       ErrorIn "frmBrowseProducts.GetRowType(ColIndex)", ColIndex
End Function

Private Sub GN_KeyPress(KeyAscii As Integer)
10        On Error GoTo errHandler
          
20        If KeyAscii = vbKeyReturn Then
30            GN_DblClick
40        End If
50        Exit Sub
errHandler:
60        If ErrMustStop Then Debug.Assert False: Resume
70        ErrorIn "frmBrowseProducts.GN_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
80        HandleError
End Sub

Public Sub mnuProductStatus()
10        On Error GoTo errHandler
      Dim frm As New frmPreDeliveryAdvice
      Dim IDs As String
      Dim i, j As Integer
      Dim x As New XArrayDB
      Dim XMLArgs As String
          
20        If XN.Count(1) = 0 Then Exit Sub
30        If IsNull(GN.Bookmark) Then Exit Sub
40        If Err Then Exit Sub
50        If GN.SelBookmarks.Count = 0 Then Exit Sub
60        strPID = FNS(XN.Value(GN.Bookmark, 11))
70        If strPID = "" Then Exit Sub
          
80        Set xMLDoc = New ujXML
90        With xMLDoc
100           .docProgID = "MSXML2.DOMDocument"
110           .docInit "doc_PRE_DEL_ADVICE"
120               .chCreate "MessageType"
130                   .elText = "PRE_DEL_ADVICE"
140               .elCreateSibling "MessageCreationDate"
150                   .elText = Format(Now(), "yyyymmddHHNN")
160               .elCreateSibling "WORKSTATION"
170                   .elText = oPC.WorkstationName
180               .elCreateSibling "DetailLines", True
190               For i = 1 To GN.SelBookmarks.Count
200                       .chCreate "ITEM"
210                       .chCreate "PID"
220                           .elText = CStr(XN(GN.SelBookmarks.Item(i - 1), 11))
230                       .navUP
240                       .navUP
250               Next i

260            XMLArgs = .docXML
270       End With
          
280       If XMLArgs > "" Then
290           frm.component XMLArgs, IDs, "R", ""
300           frm.Show vbModal
310       Else
320           MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
330       End If

340       Exit Sub
errHandler:
350       If ErrMustStop Then Debug.Assert False: Resume
360       ErrorIn "frmBrowseProducts.mnuProductStatus"
End Sub
Private Sub GN_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
10        On Error GoTo errHandler
      Dim errRepeat As Integer
20        errRepeat = 0
30        If XN.Count(1) = 0 Then Exit Sub
40        If IsNull(GN.Bookmark) Then Exit Sub
50        If Err Then Exit Sub
          
60       If Button = 2 Then   ' Check if right mouse button
                              ' was clicked.
70          PopupMenu Forms(0).mnuFindForm   ' Display the File menu as a
                              ' pop-up menu.
80       End If
90        Exit Sub
errHandler:
100       ErrPreserve
110       If Err.Number = -2147217407 Then   'Access violation
120           errRepeat = errRepeat + 1
130           LogSaveToFile "Access violation in BrowseProducts: Mousedown"  'unknown source
140           If errRepeat < 5 Then
150               Resume Next
160           Else
170               LogSaveToFile "Access violation in BrowseProducts: GN_MouseDown after 5 re-attempts"
180               MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
190               Err.Clear
200               Exit Sub
210           End If
220       End If
230       If ErrMustStop Then Debug.Assert False: Resume
240       ErrorIn "frmBrowseProducts.GN_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
               EA_NORERAISE
250       HandleError
End Sub

Public Sub AddToTempList()
10        On Error GoTo errHandler
      Dim str As String
      Dim i As Integer
      Dim TOP As Integer

20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub
          
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
40        TOP = XA.UpperBound(1)
50        XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
60        For i = 1 To GN.SelBookmarks.Count
70            XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
80            XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
90            XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
100           XA(TOP + i, 4) = 1
110           XA(TOP + i, 5) = 0
120           XA(TOP + i, 6) = ""
130           XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
140           XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
150           XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
160       Next
        
          
          
170       Exit Sub
errHandler:
180       If ErrMustStop Then Debug.Assert False: Resume
190       ErrorIn "frmBrowseProducts.AddToTempList"
End Sub
Public Sub PlaceCO()
10        On Error GoTo errHandler
      Dim frm As New frmPlaceCO
      Dim TOP As Integer
      Dim i As Integer

20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub
          
40        TOP = XA.UpperBound(1)
50        XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
60        For i = 1 To GN.SelBookmarks.Count
70            XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
80            XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
90            XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
100           XA(TOP + i, 4) = 1
110           XA(TOP + i, 5) = 0
120           XA(TOP + i, 6) = ""
130           XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
140           XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
150           XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
160       Next
          
170       frm.component XA, "ORDER"
180       frm.Show 'vbModal
190       StartNewList
200       Exit Sub
errHandler:
210       If ErrMustStop Then Debug.Assert False: Resume
220       ErrorIn "frmBrowseProducts.PlaceCO"
End Sub

Public Sub PrintLabels()
10        On Error GoTo errHandler
      Dim frm As frmPrintLabels
      Dim str As String
      Dim TOP As Integer
      Dim i As Integer
20        Set frm = New frmPrintLabels
          
30        If GN = "" Or GN = "No records" Then Exit Sub
40        If GN.Bookmark = 0 Then Exit Sub
          
50        TOP = XA.UpperBound(1)
60        XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
70        For i = 1 To GN.SelBookmarks.Count
80            If Len(FNS(XN.Value(GN.SelBookmarks(i - 1), 1))) = 13 Then
90                XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
100           Else
110               XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 13))
120           End If
130           XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
140           XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
150           XA(TOP + i, 4) = 1
160           XA(TOP + i, 5) = 0
170           XA(TOP + i, 6) = FNS(XN.Value(GN.SelBookmarks(i - 1), 13))
180           XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
190           XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
200           XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
210       Next
          
220       frm.component "S", , XA
230       frm.Show
240       StartNewList
250       Exit Sub
errHandler:
260       If ErrMustStop Then Debug.Assert False: Resume
270       ErrorIn "frmBrowseProducts.PrintLabels"
End Sub


Public Sub PlacePF(strType As String)
10        On Error GoTo errHandler
      Dim frm As New frmPlacePF
      Dim str As String
      Dim TOP As Integer
      Dim i As Integer

20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub
          
40        TOP = XA.UpperBound(1)
50        XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
60        For i = 1 To GN.SelBookmarks.Count
70            XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
80            XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
90            XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
100           XA(TOP + i, 4) = 1
110           XA(TOP + i, 5) = 0
120           XA(TOP + i, 6) = ""
130           XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
140           XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
150           XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
160       Next
          
170       frm.component XA, strType
180       frm.Show
190       StartNewList
200       Exit Sub
errHandler:
210       If ErrMustStop Then Debug.Assert False: Resume
220       ErrorIn "frmBrowseProducts.PlacePF(strType)", strType
End Sub

Public Sub PlaceOnReserve()
10        On Error GoTo errHandler
      Dim frm As New frmPlaceCO
      Dim str As String
      Dim TOP As Integer
      Dim i As Integer

20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub
          
40        TOP = XA.UpperBound(1)
50        XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
60        For i = 1 To GN.SelBookmarks.Count
70            XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
80            XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
90            XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
100           XA(TOP + i, 4) = 1
110           XA(TOP + i, 5) = 0
120           XA(TOP + i, 6) = ""
130           XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
140           XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
150           XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
160       Next
170       frm.component XA, "RESERVE"
180       frm.Show vbModal
190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.PlaceOnReserve"
End Sub
Public Sub StartNewList()
10        On Error GoTo errHandler
20        XA.Clear
30        XA.ReDim 1, 0, 1, 9
40        Exit Sub
errHandler:
50        If ErrMustStop Then Debug.Assert False: Resume
60        ErrorIn "frmBrowseProducts.StartNewList"
End Sub


Private Sub GOO_DblClick()
10        On Error GoTo errHandler
      Dim oProd As a_Product
      Dim str As String
      Dim sMsg As String
      Dim lngRes As Long
      Dim errRepeat As Integer

20        errRepeat = 0
30        str = FNS(XGOO.Value(GOO.Bookmark, 1))
40        If str = "No records" Then Exit Sub
50        If XGOO.Value(GOO.Bookmark, 2) = "" Then
60            MsgBox "Item has no title. Cannot save", vbOKOnly, "Can't do this"
70            Exit Sub
80        End If
90        If MsgBox("Do you want to create a record in the database from the Google data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
100       Set oProd = Nothing
110       Set oProd = New a_Product
120       oProd.BeginEdit
130       oProd.SetProductType "B"
140       oProd.SetTitle XGOO.Value(GOO.Bookmark, 2)
150       If IsISBN13(XGOO.Value(GOO.Bookmark, 1), True) Then
160           oProd.SetEAN XGOO.Value(GOO.Bookmark, 1)
170       Else
180           oProd.SetCode "#"
190       End If
200       oProd.SetAuthor XGOO.Value(GOO.Bookmark, 3)
210       oProd.SetPublisher XGOO.Value(GOO.Bookmark, 4)
220       oProd.SetPublicationDate XGOO.Value(GOO.Bookmark, 6)
230       oProd.SetDescription XGOO.Value(GOO.Bookmark, 5)
240       oProd.ApplyEdit lngRes, sMsg
250       Screen.MousePointer = vbDefault
260       MsgBox "Record added", , "Status"
270       Exit Sub

280       Exit Sub
errHandler:
290       ErrPreserve
300       If Err.Number = -2147217407 Then   'Access violation
310           errRepeat = errRepeat + 1
320           LogSaveToFile "Access violation in BrowseProducts: GOO_DblCLick"  'unknown source
330           If errRepeat < 5 Then
340               Resume Next
350           Else
360               LogSaveToFile "Access violation in BrowseProducts: GOO_DblCLick after 5 re-attempts"
370               MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
380               Err.Clear
390               Exit Sub
400           End If
410       End If
420       If ErrMustStop Then Debug.Assert False: Resume
430       ErrorIn "frmBrowseProducts.GOO_DblClick", , EA_NORERAISE
440       HandleError
End Sub


Private Sub Label2_Click()
10        On Error GoTo errHandler
      Dim str As String
20        str = "To use search box . . ." & vbCrLf _
                  & "Search on title . . . /Harry potter" & vbCrLf _
                  & "   will yield all titles starting with 'Harry Potter'" & vbCrLf _
                  & "Search on title . . . /*Harry potter" & vbCrLf _
                  & "   will yield all titles containing 'Harry Potter'" & vbCrLf _
                  & "Search on title . . . /*Harry * goblet" & vbCrLf _
                  & "   will yield all titles containing 'Harry' and 'goblet' in that order" & vbCrLf & vbCrLf _
                  & "Replacing '/' with '//' will search authors" & vbCrLf & vbCrLf _
                  & "Replacing '/' with '///' will search publishers" & vbCrLf & vbCrLf _
                  & "Adding '!!' at the end of the search string will search on Bookfind (if installed)" & vbCrLf
30        MsgBox str, vbInformation, "Help"
          
40        Exit Sub
errHandler:
50        If ErrMustStop Then Debug.Assert False: Resume
60        ErrorIn "frmBrowseProducts.Label2_Click", , EA_NORERAISE
70        HandleError
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
10        On Error GoTo errHandler
20        If cboSearch = "" And UCase(Me.cboCategory) = "<ALL>" And UCase(Me.cboCategory) = "<ALL>" Then Exit Sub
30        If KeyAscii = vbKeyReturn Then
40            cmdSearch_Click
50            If GN.Visible = True Then
60                mSetfocus GN
70            Else
80                mSetfocus GBF
90            End If
100       End If
110       Exit Sub
120       Exit Sub
errHandler:
130       If ErrMustStop Then Debug.Assert False: Resume
140       ErrorIn "frmBrowseProducts.cboSearch_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
150       HandleError
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
10        On Error GoTo errHandler
          
          'Attempting to avoid occasional windows crash when doublwclicking the GN grid. Reducing the size of the other grids so one does not overlay another 17/7/2010
20        If SSTab1.Tab = 0 Then
30            mSetfocus Me.cboSearch
40        ElseIf SSTab1.Tab = 1 Then
50            SetActiveGrid "GN"
60            GBF.Visible = True
70        Else
80            SetActiveGrid "GOO"
90            mSetfocus cboGoogle
100       End If
110       Exit Sub
errHandler:
120       If ErrMustStop Then Debug.Assert False: Resume
130       ErrorIn "frmBrowseProducts.SSTab1_Click(PreviousTab)", PreviousTab, EA_NORERAISE
140       HandleError
End Sub
Private Sub SetActiveGrid(Gridcode As String)
10        On Error GoTo errHandler
20        If Gridcode = sCurrentGrid Then Exit Sub
30        Select Case Gridcode
          Case "GN"
40            GN.ZOrder 0
50            GBF.ZOrder 1
60            GOO.ZOrder 1
70            If sCurrentGrid = "GBF" Then
80                GN.Width = GBF.Width
90                GN.Height = GBF.Height
100               GBF.Width = 0
110               GBF.Height = 0
120           Else
130               GN.Width = GOO.Width
140               GN.Height = GOO.Height
150               GOO.Width = 0
160               GOO.Height = 0
170           End If
180           sCurrentGrid = "GN"
190           GN.Visible = True
200       Case "GBF"
210           GBF.ZOrder 0
220           GN.ZOrder 1
230           GOO.ZOrder 1
240           If sCurrentGrid = "GN" Then
250               GBF.Width = GN.Width
260               GBF.Height = GN.Height
270               GN.Width = 0
280               GN.Height = 0
290           Else
300               GBF.Width = GOO.Width
310               GBF.Height = GOO.Height
320               GOO.Width = 0
330               GOO.Height = 0
340           End If
350           sCurrentGrid = "GBF"
360           GBF.Visible = True
370       Case "GOO"
380           GOO.ZOrder 0
390           GBF.ZOrder 1
400           GN.ZOrder 1
410           If sCurrentGrid = "GN" Then
420               GOO.Width = GN.Width
430               GOO.Height = GN.Height
440               GN.Width = 0
450               GN.Height = 0
460           Else
470               GOO.Width = GBF.Width
480               GOO.Height = GBF.Height
490               GBF.Width = 0
500               GBF.Height = 0
510           End If
520           sCurrentGrid = "GOO"
530           GOO.Visible = True
540       End Select
550       Exit Sub
errHandler:
560       If ErrMustStop Then Debug.Assert False: Resume
570       ErrorIn "frmBrowseProducts.SetActiveGrid(Gridcode)", Gridcode
End Sub
Private Sub txtmaxnum_Validate(Cancel As Boolean)
10        On Error GoTo errHandler
20        If Not IsNumeric(txtmaxnum) Then
30            MsgBox "You must enter a number here, representing the maximum number of records you want to get back, a suggested value is 500.", , "Invalid value"
40            Cancel = True
50        End If
60        Exit Sub
errHandler:
70        If ErrMustStop Then Debug.Assert False: Resume
80        ErrorIn "frmBrowseProducts.txtmaxnum_Validate(Cancel)", Cancel, EA_NORERAISE
90        HandleError
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
10        On Error GoTo errHandler
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

20        Set xMLDoc = New ujXML
30        With xMLDoc
40            .docProgID = "MSXML2.DOMDocument"
50            .docInit "SSI_1"
60            .chCreate "SSI"
70                .elText = "Selected stock items at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
              
80                .elCreateSibling "DetailLine", True
90                .chCreate "Col_1"
100                   .elText = "Code"
110               .elCreateSibling "Col_2"
120                   .elText = "Title"
130               .elCreateSibling "Col_3"
140                   .elText = "Author"
150               .elCreateSibling "Col_4"
160                   .elText = "Publisher"
170               .elCreateSibling "Col_5"
180                   .elText = "Price"
190               .elCreateSibling "Col_6"
200                   .elText = "Qty"
210               .elCreateSibling "Col_7"
220                   .elText = "OH"
230               .elCreateSibling "Col_8"
240                   .elText = "OO"
250               .elCreateSibling "Col_9"
260                   .elText = "CO"
270                   .navUP
              
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
280               If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" And Me.SSTab1.Tab = 0 Then
290                   For i = 1 To XBF.UpperBound(1)
300                       If mIsAmongBookmarks(XBF, colList(i).EAN, GBF, 1, "STRING") Then
310                           .elCreateSibling "DetailLine", True
320                           .chCreate "Col_1"
330                               .elText = XBF.Value(i, 1)
340                           .elCreateSibling "Col_2"
350                               .elText = XBF.Value(i, 2)
360                           .elCreateSibling "Col_3"
370                               .elText = XBF.Value(i, 3)
380                           .elCreateSibling "Col_4"
390                               .elText = XBF.Value(i, 6)
400                           .elCreateSibling "Col_5"
410                               .elText = XBF.Value(i, 9)
420                           .elCreateSibling "Col_6"
430                               .elText = "n/a" 'colList.Item(i).QtyOnHand
440                               .navUP
450                       End If
460                   Next
470               Else
480                   For i = 1 To XN.UpperBound(1)
490                       If mIsAmongBookmarks(XN, XN.Value(i, 11), GN, 11, "UNIQUEIDENTIFIER") Then
500                           .elCreateSibling "DetailLine", True
510                           .chCreate "Col_1"
520                               .elText = XN.Value(i, 1)
530                           .elCreateSibling "Col_2"
540                               .elText = XN.Value(i, 2)
550                           .elCreateSibling "Col_3"
560                               .elText = XN.Value(i, 3)
570                           .elCreateSibling "Col_4"
580                               .elText = XN.Value(i, 6)
590                           .elCreateSibling "Col_5"
600                               .elText = XN.Value(i, 9)
610                           .elCreateSibling "Col_6"
620                               .elText = XN.Value(i, 5)
630                           .elCreateSibling "Col_7"
640                               .elText = XN.Value(i, 20)
650                           .elCreateSibling "Col_8"
660                               .elText = XN.Value(i, 21)
670                           .elCreateSibling "Col_9"
680                               .elText = XN.Value(i, 22)
690                               .navUP
700                       End If
710                   Next
720               End If

              
730       End With
          
      'FINALLY PRODUCE THE .XML FILE
740       strXML = oPC.SharedFolderRoot & "\TEMP\SSI" & ".xml"
750       With xMLDoc
760           If fs.FileExists(strXML) Then
770               fs.DeleteFile strXML
780           End If
790           .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
800       End With

      ''WRITE THE .RTF FILE
810       If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt") Then
820           MsgBox "You are missing the template file " & "SSI_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
830       End If
840       objXSL.async = False
850       objXSL.validateOnParse = False
860       objXSL.resolveExternals = False
870       strPath = oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt"
880       Set fs = New FileSystemObject
890       If fs.FileExists(strPath) Then
900           objXSL.Load strPath
910       End If

      '    strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1.RTF"
      '    If fs.FileExists(strFilename) Then
      '        fs.DeleteFile strFilename, True
      '    End If
920       strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1.RTF"
930       i = 0
940       Do Until fs.FileExists(strFilename) = False
950           i = i + 1
960           strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1" & "_" & CStr(i) & ".RTF"
970       Loop
980       oTF.OpenTextFileToAppend strFilename
990       oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
1000      oTF.CloseTextFile

1010      strExecutable = GetPDFExecutable(strFilename)
        If strExecutable = "" Then
            MsgBox "There is no application set on this computer to open the file: " & strFilename & ". The document cannot be displayed", vbOKOnly, "Can't do this"
        Else
            Shell strExecutable & " " & strFilename
        End If
          
1030      Exit Function
errHandler:
1040      If ErrMustStop Then Debug.Assert False: Resume
1050      ErrorIn "frmBrowseProducts.ExportToXML"
End Function

Public Sub mnuSetPT()
10        On Error GoTo errHandler
      Dim IDs As String
      Dim frm As New frmSetProductType
      Dim i As Integer

20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub
40        ReDim strTitle(GN.SelBookmarks.Count)
50        IDs = ""
60        For i = 0 To GN.SelBookmarks.Count - 1
70            IDs = IDs & ",'" & XN(GN.SelBookmarks(i), 11) & "'"
80        Next i
90        If Left(IDs, 1) = "," Then
100           IDs = Right(IDs, Len(IDs) - 1)
110       End If
120       If IDs > "" Then
130           frm.component IDs
140           frm.Show vbModal
150       Else
160           MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
170       End If
180       Unload frm
190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.mnuSetPT"
End Sub
Public Sub SetForWebExport()
10        On Error GoTo errHandler
      Dim cnt As Integer
20        cnt = 0
30        ReDim strTitle(GN.SelBookmarks.Count)
          Dim i As Integer
40        For i = 0 To GN.SelBookmarks.Count - 1
50            MarkProductForWebExport XN(GN.SelBookmarks(i), 11)
60            cnt = cnt + 1
70        Next i
80        If cnt > 0 Then
90            MsgBox "Records have been marked for Web export", vbInformation, "Status"
100       Else
110           MsgBox "There are no rows selected. Click on the left margin to select rows before choosing option to mark for web export.", vbInformation, "Status"
120       End If

130       Exit Sub
errHandler:
140       If ErrMustStop Then Debug.Assert False: Resume
150       ErrorIn "frmBrowseProducts.SetForWebExport"
End Sub
Public Sub mnuSetSection()
10        On Error GoTo errHandler
      Dim IDs As String
      Dim frm As New frmSetSection
      Dim i As Integer

20        If GN = "" Or GN = "No records" Then Exit Sub
30        If GN.Bookmark = 0 Then Exit Sub
40        ReDim strTitle(GN.SelBookmarks.Count)
50        IDs = ""
60        For i = 0 To GN.SelBookmarks.Count - 1
70            IDs = IDs & ",'" & XN(GN.SelBookmarks(i), 11) & "'"
80        Next i
90        If Left(IDs, 1) = "," Then
100           IDs = Right(IDs, Len(IDs) - 1)
110       End If
120       If IDs > "" Then
130           frm.component IDs
140           frm.Show vbModal
150       Else
160           MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
170       End If
180       Unload frm
190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.mnuSetSection"
End Sub
Public Sub mnuFindAllSOH()
10        On Error GoTo errHandler
      Dim IDs As String
      Dim frm As New frmSOHALL
      Dim i As Integer
      Dim strEAN As String
      Dim strTitle As String
      Dim oSQL As New z_SQL
      Dim lngREQID As Long
20        p 0

30        If GN = "" Or GN = "No records" Then Exit Sub
40        If GN.Bookmark = 0 Then Exit Sub
50        strEAN = XN(GN.Bookmark, 13)
60        strTitle = XN(GN.Bookmark, 2)
70        oSQL.Request_SOH_ALLBRANCHES strEAN, lngREQID
80        MsgWaitObj 2000
90        If strEAN > "" And lngREQID > 0 Then
100       p 1
110           frm.Caption = strEAN
120           frm.component lngREQID, strTitle
130           frm.Show vbModal
140       Else
150           MsgBox "You have not clicked on a row.", vbInformation, "No row"
160       End If
170       Unload frm
180       p 2

190       Exit Sub
errHandler:
200       ErrPreserve
210       If Err.Number = -2147217407 Then   'Access violation
220           errRepeat = errRepeat + 1
230           LogSaveToFile "Access violation in frmBrowseProducts.mnuFindAllSOH"  'unknown source
240           If errRepeat < 5 Then
250               Err.Clear
260               Exit Sub
270           Else
280               LogSaveToFile "Access violation in frmBrowseProducts.mnuFindAllSOH after 5 re-attempts"
290               MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
300               Err.Clear
310               Exit Sub
320           End If
330       End If
340       If ErrMustStop Then Debug.Assert False: Resume
350       ErrorIn "frmBrowseProducts.mnuFindAllSOH", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
360       HandleError
End Sub
Public Sub mnuTouchRecord()
10        On Error GoTo errHandler
      Dim cnt As Integer
20        cnt = 0
30        If GN = "" Or GN = "No records" Then Exit Sub
40        If GN.Bookmark = 0 Then Exit Sub
50        ReDim strTitle(GN.SelBookmarks.Count)
          Dim i As Integer
60        For i = 0 To GN.SelBookmarks.Count - 1
70            TouchRecord XN(GN.SelBookmarks(i), 11)
80            cnt = cnt + 1
90        Next i
100       If cnt > 0 Then
110           MsgBox "P.O.S. computers have been updated", vbInformation, "Status"
120       Else
130           MsgBox "There are no rows selected. Click on the left margin to select rows before choosing option to send to P.O.S. computers.", vbInformation, "Status"
140       End If
150       Exit Sub
errHandler:
160       If ErrMustStop Then Debug.Assert False: Resume
170       ErrorIn "frmBrowseProducts.mnuTouchRecord"
End Sub
Public Sub mnuLoadReorderSlate()
10        On Error GoTo errHandler
      Dim frmREORDER_SAL As frmREORDER_CO
      Dim oSQL As New z_SQL
      Dim i As Integer
      Dim TOP As Integer
20    p 1

30        If GN = "" Or GN = "No records" Then Exit Sub
40    p 2
50        If GN.Bookmark = 0 Then Exit Sub
60    p 3
70        TOP = XA.UpperBound(1)
80    p 4
90        XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
100   p 5
110       For i = 1 To GN.SelBookmarks.Count
120           XA(TOP + i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
130           XA(TOP + i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
140           XA(TOP + i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
150           XA(TOP + i, 4) = 1
160           XA(TOP + i, 5) = 0
170           XA(TOP + i, 6) = ""
180           XA(TOP + i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
190           XA(TOP + i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
200           XA(TOP + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
210       Next
220   p 6
230       oSQL.LoadBrowsedProductsToTempTable XA
240   p 7
250       StartNewList
260       Exit Sub
errHandler:
270       If ErrMustStop Then Debug.Assert False: Resume
280       ErrorIn "frmBrowseProducts.mnuLoadReorderSlate"
End Sub

Private Sub TouchRecord(pPID As String)
10        On Error GoTo errHandler
      Dim oSQL As New z_SQL

      '    oSQL.RunSQL "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
      '            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_LoyaltyRATE," _
      '            & "PRU_PTID,PRU_SECID,PRU_MULTIBUYCODE) " _
      '            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
      '            & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID,vMultibuyCode.DICT_System " _
      '            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID   LEFT JOIN vMultibuyCode ON P_ID = vMultibuyCode.PSEC_P_ID" _
      '            & " WHERE P_ID = '" & pPID & "'"
20        oSQL.RunSQL "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
                  & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_SSP,PRU_NDA,PRU_LoyaltyRATE," _
                  & "PRU_PTID,PRU_SECID,PRU_MULTIBUYCODE) " _
                  & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
                  & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_Special,P_NDA,P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID,P_MultibuyCode " _
                  & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID   LEFT JOIN vMultibuyCode ON P_ID = vMultibuyCode.PSEC_P_ID" _
                  & " WHERE P_ID = '" & pPID & "'"

30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmBrowseProducts.TouchRecord(pPID)", pPID
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

10        If GN = "" Or GN = "No records" Then Exit Sub
20        If GN.Bookmark = 0 Then Exit Sub
          
30        Set xMLDoc = New ujXML
40        With xMLDoc
50            .docProgID = "MSXML2.DOMDocument"
60            .docInit "doc_SpecialOrderAddition"
70                .chCreate "MessageType"
80                    .elText = "SpecialOrderAddition"
90                .elCreateSibling "MessageCreationDate"
100                   .elText = Format(Now(), "yyyymmddHHNN")
110               .elCreateSibling "StaffMember"
120                   .elText = CStr(pSTAFFID)
130               .elCreateSibling "DetailLines", True
140               For i = 0 To GN.SelBookmarks.Count - 1
150                       .chCreate "ITEM"
160                       .chCreate "PID"
170                           .elText = XN(GN.SelBookmarks(i), 11)
      '                    .elCreateSibling "CodeF"
      '                        .elText = colList.Item(GN.SelBookmarks(i)).CodeF
      '                    .elCreateSibling "Description"
      '                        .elText = Replace(UCase(colList.Item(GN.SelBookmarks(i)).StatusShortF(True, True)) & " " & colList.Item(GN.SelBookmarks(i)).Title, "'", "''")

180                       .navUP
190                       .navUP
200               Next i

210            XMLArgs = .docXML
220       End With
          
230       If XMLArgs > "" Then
240           oSM.CreateSpecialOrder XMLArgs
      '        frm.component XMLArgs, IDs, "R", ""
      '        frm.Show vbModal
250       Else
260           MsgBox "Make a selection by clicking on the margin. (The whole line will be marked in blue.)" & vbCrLf & "Remember, you can select many lines at once by holding the CTRL key as you make selections.", vbInformation, "No selection"
270       End If
280       Exit Sub
End Sub
