VERSION 5.00
Object = "{8B07DDC0-1FC2-4D71-B114-D4F3E02F1F1A}#1.0#0"; "InteropUserControlLibrary1.tlb"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmBrowseProducts 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse products"
   ClientHeight    =   8520.001
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
   Icon            =   "frmBrowseProducts4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520.001
   ScaleWidth      =   14265
   Begin InteropUserControlLibrary1Ctl.BookfindGrid GBF 
      Height          =   2985
      Left            =   4020
      TabIndex        =   42
      Top             =   2805
      Width           =   8550.001
      Object.Visible         =   "True"
      Enabled         =   "True"
      ForegroundColor =   "-2147483630"
      BackgroundColor =   "13882315"
      SearchArguments =   ""
      QtyRowsToReturn =   "6"
      FirstRownumberToReturn=   "0"
      Location        =   "268, 187"
      Name            =   "BookfindGrid"
      Size            =   "570, 199"
      Object.TabIndex        =   "0"
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
      Picture         =   "frmBrowseProducts4.frx":0442
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
      Tab             =   1
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483644
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Normal search"
      TabPicture(0)   =   "frmBrowseProducts4.frx":07CC
      Tab(0).ControlEnabled=   0   'False
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
      TabPicture(1)   =   "frmBrowseProducts4.frx":07E8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frCatalogue"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Google Books"
      TabPicture(2)   =   "frmBrowseProducts4.frx":0804
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "cmdGOO"
      Tab(2).Control(2)=   "cboGoogle"
      Tab(2).Control(3)=   "Inet1"
      Tab(2).Control(4)=   "cmdMoreGoo"
      Tab(2).Control(5)=   "chkISBNOnly"
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame1 
         Caption         =   "Course codes"
         ForeColor       =   &H8000000D&
         Height          =   1065
         Left            =   7050
         TabIndex        =   39
         Top             =   360
         Width           =   4050
         Begin VB.CommandButton cmdCourseCodes 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   2940
            Picture         =   "frmBrowseProducts4.frx":0820
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
         Left            =   -72420.01
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
         Left            =   -74745.01
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
         Left            =   -64665
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
         Left            =   -64665
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
         Left            =   -67050.01
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   630
         Width           =   1260
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   -71580.01
         Top             =   1050
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
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
         ItemData        =   "frmBrowseProducts4.frx":0BAA
         Left            =   -74865.01
         List            =   "frmBrowseProducts4.frx":0BAC
         TabIndex        =   29
         Top             =   660
         Width           =   6375
      End
      Begin VB.CommandButton cmdGOO 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&First"
         Height          =   435
         Left            =   -68370.01
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
         Left            =   -74550.01
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
         Left            =   -74940.01
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1140
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search by BIC codes (if captured)"
         ForeColor       =   &H8000000D&
         Height          =   1035
         Left            =   3585
         TabIndex        =   22
         Top             =   375
         Width           =   3300
         Begin VB.CommandButton cmdBIC 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   675
            Left            =   960
            Picture         =   "frmBrowseProducts4.frx":0BAE
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
         Left            =   210
         TabIndex        =   19
         Top             =   375
         Width           =   3240
         Begin VB.CommandButton cmdCAT 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Find"
            Height          =   705
            Left            =   2070
            Picture         =   "frmBrowseProducts4.frx":0F38
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   210
            Width           =   975
         End
         Begin VB.ComboBox cboCat 
            ForeColor       =   &H00800000&
            Height          =   360
            ItemData        =   "frmBrowseProducts4.frx":12C2
            Left            =   165
            List            =   "frmBrowseProducts4.frx":12C4
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
         Left            =   -69465.01
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
         Left            =   -69480.01
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
         Left            =   -66975.01
         Picture         =   "frmBrowseProducts4.frx":12C6
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
         ItemData        =   "frmBrowseProducts4.frx":1650
         Left            =   -74160.01
         List            =   "frmBrowseProducts4.frx":1652
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
         Left            =   -73965.01
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
         Left            =   -65160
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
         Left            =   -65310
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
         Left            =   -74850.01
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
         Left            =   -70935.01
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
         Left            =   -70275.01
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
         Left            =   -74160.01
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
         Left            =   -74580.01
         TabIndex        =   7
         Top             =   510
         Width           =   360
      End
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   4455
      Left            =   225
      OleObjectBlob   =   "frmBrowseProducts4.frx":1654
      TabIndex        =   3
      Top             =   3135
      Width           =   11340
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10320
      Picture         =   "frmBrowseProducts4.frx":6957
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
      OleObjectBlob   =   "frmBrowseProducts4.frx":6CE1
      TabIndex        =   15
      Top             =   4245
      Visible         =   0   'False
      Width           =   11250
   End
   Begin TrueOleDBGrid60.TDBGrid GOO 
      Height          =   4455
      Left            =   120
      OleObjectBlob   =   "frmBrowseProducts4.frx":BFE1
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
    On Error GoTo errHandler
    SaveLayout Me.GN, "SearchFormA"
    SaveLayout Me.GBF, "SearchFormB"
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
    
    If UCase(Right(cboSearch, 2)) = "+B" Or UCase(Right(cboSearch, 2)) = "!!" Then
        SetActiveGrid "GBF"
        GBF.SearchArguments = Left(cboSearch, Len(cboSearch) - 2)
        GBF.search
       ' search enSearchBF, Left(cboSearch, Len(cboSearch) - 2)
      '  mSetfocus GBF
    Else
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
160           lngSectionID = oPC.Configuration.Sections.key(pSection)
170       End If
180       If pProductType <> "<ALL>" Then
190           lngProductTypeID = oPC.Configuration.ProductTypes.key(pProductType)
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
350           ElseIf pSearchType = enSearchBF Then
360               oSearchEngine.SetupSQLwoCriteria False, False, pSearchType, False, lngMaxRecs, "B", (chkIncludeObsolete = 1)
370               enSource = enBF
380               oSearchEngine.BFSearchEx pCriteria, lngRecsFound, CLng(txtmaxnum), lngResult
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
610               oSearchEngine.Execute lngMaxRecs
620               Set colList = Nothing
630               Set colList = oSearchEngine.getcols
640               lngrows = oSearchEngine.Rows
650           Else
660               oSearchEngineC.MassageRows rsResult
670               Set colList = Nothing
680               Set colList = oSearchEngineC.getcols
690               lngrows = oSearchEngineC.Rows
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
850               GBF.ReBind
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
        GBF.Visible = False
        GBF.Width = 0
        GN.Visible = True
        XN.Clear
        XBF.Clear
        GBF.ReBind
        XN.ReDim 1, colList.Count, 1, 30
        For i = 1 To colList.Count
                XN.Value(i, val(oColMap.key("Code"))) = colList.Item(i).CodeF
                XN.Value(i, val(oColMap.key("Item"))) = UCase(colList.Item(i).StatusShortF(True, True)) & " " & colList.Item(i).Title
                XN.Value(i, val(oColMap.key("Author"))) = colList.Item(i).Author
                XN.Value(i, val(oColMap.key("Distributor"))) = colList.Item(i).Distributor
                XN.Value(i, val(oColMap.key("OH/OO/CO"))) = colList.Item(i).QtyOnHand & " / " & colList.Item(i).QtyonOrder & " / " & colList.Item(i).QtyOnBackorder
                XN.Value(i, val(oColMap.key("Publisher"))) = colList.Item(i).Publisher
                XN.Value(i, val(oColMap.key("PublicationDate"))) = colList.Item(i).PubDate & IIf(colList.Item(i).Edition > "", "/", "") & colList.Item(i).Edition
                XN.Value(i, val(oColMap.key("TotalSold"))) = colList.Item(i).QtyTotalSold
                XN.Value(i, val(oColMap.key("LastDateDelivered"))) = colList.Item(i).LastDateDelivered
                XN.Value(i, val(oColMap.key("S.P."))) = colList.Item(i).LocalPriceF
                XN.Value(i, val(oColMap.key("Multibuy"))) = colList.Item(i).Multibuy
                XN.Value(i, val(oColMap.key("Categories"))) = colList.Item(i).Categories
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
        
        
        
    Case enBF
        XN.Clear
        GN.ReBind
        XBF.Clear
        GBF.Visible = True
        GN.Visible = False
        XBF.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
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
                XBF.Value(i, 8) = colList.Item(i).DistributorCode & " : " & colList.Item(i).Distributor
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
        MsgBox strGOOXML
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
        s = XN(i, val(oColMap.key("Code")))
        s = s & vbTab & XN(i, val(oColMap.key("Item")))
        s = s & vbTab & XN(i, val(oColMap.key("Author")))
        s = s & vbTab & XN(i, val(oColMap.key("Distributor")))
        s = s & vbTab & XN(i, val(oColMap.key("Publisher")))
        s = s & vbTab & XN(i, val(oColMap.key("PublicationDate")))
        s = s & vbTab & XN(i, val(oColMap.key("TotalSold")))
        s = s & vbTab & XN(i, val(oColMap.key("LastDateDelivered")))
        s = s & vbTab & XN(i, val(oColMap.key("S.P.")))
     '   s = s & vbTab & XN(i, val(oColMap.key("Multibuy")))
        s = s & vbTab & XN(i, val(oColMap.key("Categories")))
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



Private Sub Command1_Click()
    On Error GoTo errHandler
Dim str As String
    If oPC.InternetDialup = True Then Exit Sub
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Command2_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub Form_Activate()
    On Error GoTo errHandler

    SetMenu
    p 1
    p 2
    p 3
    bWithCopies = False
    chkCopies = IIf(bWithCopies, 1, 0)
    Me.Command1.Enabled = Not oPC.InternetDialup
    p 4
    cmdGetFromSB.Visible = oPC.getProperty("UsesHUB") = "TRUE"
    lblHUBRESULT.Visible = cmdGetFromSB.Visible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Activate", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
    HandleError
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
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Deactivate", , EA_NORERAISE
    HandleError
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
20        oColMap.Load oPC.getProperty("BrowsePosition")

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

190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "frmBrowseProducts.InitializeGrids"
End Sub
Private Sub FormatGN()
Dim i As Integer

Dim lngAlignment As Long

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
End Sub

Private Sub Form_Load()
10        On Error GoTo errHandler
      Dim i As Integer

20         errSysHandlerSet
         

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
130       XBF.Clear
140       XBF.ReDim 1, 1, 1, 12
150       XGOO.Clear
160       XGOO.ReDim 1, 1, 1, 12
170       flgLoading = True
180       Resize_GN
          
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
          
520       SetActiveGrid "GN"
530       Exit Sub
errHandler:
540       If ErrMustStop Then Debug.Assert False: Resume
550       ErrorIn "frmBrowseProducts.Form_Load", , EA_NORERAISE
560       HandleError
End Sub
Private Sub Resize_GBF()
Dim i As Long
    For i = 1 To GBF.Columns.Count
        GBF.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormB", CStr(i), GBF.Columns(i - 1).Width)
    Next
End Sub
Private Sub Resize_GOO()
Dim i As Long
    For i = 1 To GOO.Columns.Count
        GOO.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormC", CStr(i), GOO.Columns(i - 1).Width)
    Next
End Sub
Private Sub Resize_GN()
Dim i As Long
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
    Next
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    If flgLoading Then Exit Sub
    GN.Width = NonNegative_Lng(Me.Width - 380)
    GBF.Width = GN.Width
    GOO.Width = GN.Width
    
    lngDiff = GN.Height
    GN.Height = NonNegative_Lng(Me.Height - (GN.TOP + 1070))
    GBF.Height = GN.Height
    GOO.Height = GN.Height
    
    lngDiff = GN.Height - lngDiff
    
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    cmdPrintbarcode.TOP = cmdPrintbarcode.TOP + lngDiff
    cmdClose.Left = NonNegative_Lng(GN.Width - 1000)
    
   ' ResizeColumns
       ' FormatGN

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.Form_Resize", , EA_NORERAISE
    HandleError
End Sub
Sub ResizeColumns()
    On Error GoTo errHandler
Dim i As Integer
Dim newwidth As Long

    For i = 1 To GN.Columns.Count - 1
        GN.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GN.Width) / OriginalwdthGN) * 0.9)
    Next
    For i = 1 To GOO.Columns.Count - 1
        GOO.Columns(i - 1).Width = wdthCol(i - 1) * ((CDbl(GOO.Width) / OriginalwdthGN) * 0.9)
    Next

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
10        On Error GoTo errHandler
      Dim str As String
          
          
20        If XBF Is Nothing Then Exit Sub
30        If XBF.UpperBound(1) = 0 Then Exit Sub
40        If Err Then Exit Sub
50        If IsNull(GBF.Bookmark) Then Exit Sub
60        If Err Then Exit Sub
          
          
70        str = FNS(XBF.Value(GBF.Bookmark, 1))
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
    On Error GoTo errHandler
Dim oProd As a_Product
Dim str As String
Dim frm As frmProduct

    str = FNS(XBF.Value(GBF.Bookmark, 1))
    If str = "No records" Then Exit Sub
    If str = "" Then Exit Sub
    If MsgBox("Do you want to create a record in the database from the Bookfind data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    
    If CheckThisPoint(M_NEWPRODUCT) Then
        If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then Exit Sub
    End If
    
    Set oProd = Nothing
    Set oProd = New a_Product
    Screen.MousePointer = vbHourglass
    oProd.Load "", 0, str, , , True
   ' oProd.BeginEdit
   ' oProd.SetCode oProd.code   'to force validation
    Set frm = New frmProduct
    frm.component oProd
    frm.Show
    Screen.MousePointer = vbDefault
   ' MsgBox "Record added", , "Status"
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
10        On Error GoTo errHandler
      Dim str As String

20        On Error Resume Next
30        If XGOO Is Nothing Then Exit Sub
40        If XGOO.Count(1) = 0 Then Exit Sub
50        If XGOO.UpperBound(1) = 0 Then Exit Sub
60        If Err Then Exit Sub
70        If IsNull(GOO.Bookmark) Then Exit Sub
80        On Error GoTo errHandler
          
          
90        str = FNS(XGOO.Value(GOO.Bookmark, 1))
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
10        On Error GoTo errHandler
Dim errRepeat As Integer
Dim str As String

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
600               Exit Sub
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
            mSetfocus GBF
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
        GBF.Visible = True
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
    If Gridcode = sCurrentGrid Then Exit Sub
    Select Case Gridcode
    Case "GN"
        GN.ZOrder 0
        GBF.ZOrder 1
        GOO.ZOrder 1
        If sCurrentGrid = "GBF" Then
            GN.Width = GBF.Width
            GN.Height = GBF.Height
            GBF.Width = 0
            GBF.Height = 0
        Else
            GN.Width = GOO.Width
            GN.Height = GOO.Height
            GOO.Width = 0
            GOO.Height = 0
        End If
        sCurrentGrid = "GN"
        GN.Visible = True
    Case "GBF"
        GBF.ZOrder 0
        GN.ZOrder 1
        GOO.ZOrder 1
        If sCurrentGrid = "GN" Then
            GBF.Width = GN.Width
            GBF.Height = GN.Height
            GN.Width = 0
            GN.Height = 0
        Else
            GBF.Width = GOO.Width
            GBF.Height = GOO.Height
            GOO.Width = 0
            GOO.Height = 0
        End If
        sCurrentGrid = "GBF"
        GBF.Visible = True
    Case "GOO"
        GOO.ZOrder 0
        GBF.ZOrder 1
        GN.ZOrder 1
        If sCurrentGrid = "GN" Then
            GOO.Width = GN.Width
            GOO.Height = GN.Height
            GN.Width = 0
            GN.Height = 0
        Else
            GOO.Width = GBF.Width
            GOO.Height = GBF.Height
            GBF.Width = 0
            GBF.Height = 0
        End If
        sCurrentGrid = "GOO"
        GOO.Visible = True
    End Select
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

1010      strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
1020      Shell strExecutable, vbNormalFocus
          
1030      Exit Function
errHandler:
1040      If ErrMustStop Then Debug.Assert False: Resume
1050      ErrorIn "frmBrowseProducts.ExportToXML"
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
