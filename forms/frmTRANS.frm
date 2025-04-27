VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTRANS 
   BackColor       =   &H00DFDDEA&
   Caption         =   "Credit note"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmTRANS.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.TextBox txtAccnum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton cmdSelectCustomer 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Find customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8625
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmTRANS.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDDEA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1050
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   5220
      Width           =   3390
   End
   Begin VB.TextBox txtCustName 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDDEA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00706034&
      Height          =   255
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   150
      Width           =   1680
   End
   Begin VB.Frame fr2 
      BackColor       =   &H00DFDDEA&
      Height          =   1275
      Left            =   675
      TabIndex        =   22
      Top             =   3885
      Width           =   10185
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4860
         TabIndex        =   23
         Top             =   870
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdNewRows 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3990
      Width           =   630
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7515
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmTRANS.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5370
      Width           =   1110
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDDEA&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00706034&
      Height          =   250
      Left            =   4215
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   495
      Width           =   1695
   End
   Begin VB.TextBox txtRunningTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   9525
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3915
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00E0E0E0&
      Height          =   1305
      Left            =   705
      TabIndex        =   12
      Top             =   3855
      Width           =   10110
      Begin EXCOMBOBOXLibCtl.ComboBox cboMatch 
         Height          =   315
         Left            =   1680
         OleObjectBlob   =   "frmTRANS.frx":0A5E
         TabIndex        =   4
         Top             =   480
         Width           =   3945
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   5610
         TabIndex        =   33
         Top             =   465
         Width           =   735
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4785
         TabIndex        =   7
         Top             =   825
         Width           =   4260
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   7320
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   9150
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   435
         Width           =   840
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   825
         Width           =   3975
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6330
         TabIndex        =   5
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   465
         Width           =   1560
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Original invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1785
         TabIndex        =   35
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   5490
         TabIndex        =   34
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4275
         TabIndex        =   31
         Top             =   870
         Width           =   480
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   7050
         TabIndex        =   16
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   6495
         TabIndex        =   14
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   360
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "IN PROCESS"
      Top             =   5535
      Width           =   1260
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9750
      Picture         =   "frmTRANS.frx":1C84
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2580
      Left            =   135
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1215
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4551
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Inv. code"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Discount"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8550
      TabIndex        =   30
      Top             =   75
      Width           =   300
   End
   Begin VB.Label lblb 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6195
      TabIndex        =   29
      Top             =   60
      Width           =   300
   End
   Begin VB.Label lblAddDel 
      BackColor       =   &H00DFDDEA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   870
      Left            =   8865
      TabIndex        =   28
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label lblAddBill 
      BackColor       =   &H00DFDDEA&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   870
      Left            =   6525
      TabIndex        =   27
      Top             =   90
      Width           =   1920
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   3705
      TabIndex        =   21
      Top             =   150
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "A/C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   135
      TabIndex        =   20
      Top             =   165
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3735
      Picture         =   "frmTRANS.frx":1DCE
      Stretch         =   -1  'True
      Top             =   495
      Width           =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOK 
         Caption         =   "OK"
      End
      Begin VB.Menu mnuFileCancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnuFileSaveNew 
         Caption         =   "Save / New"
      End
      Begin VB.Menu mnuFileVoid 
         Caption         =   "Void"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuDel 
         Caption         =   "&Delete selected row"
      End
      Begin VB.Menu mnuEditNote 
         Caption         =   "Cutomer Note"
      End
      Begin VB.Menu mnuAddresses 
         Caption         =   "&Addresses"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "frmTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCN As a_CN
Attribute oCN.VB_VarHelpID = -1
Dim WithEvents oCNLine As a_CNL
Attribute oCNLine.VB_VarHelpID = -1
Dim oInv As c_Invoices
'Dim frmOL As frmOverlay1
Dim oCustomer As a_Customer
Dim oProd As a_Product
Dim oCurrentCopy
Dim bValidCN As Boolean
Dim bValidCNLine As Boolean
Dim tlCustomer As z_TextList
Dim lngCurrentExtension As Long
Dim lngCurrentTotal As Long
Dim lngCurrentDepositTotal As Long
Dim lngCurrentVATTotal As Long

Dim lngSelectedRowIndex As String
Dim lngEditingIdx As String
Dim vMode As EnumMode  ' 1:TPExists,Adding row;  2:TPExists, not adding row;  3 TPAbsent,not adding row
Dim bFrameEnabled As Boolean
Dim lngStockBal As Long
Dim curDeposit As Currency
Dim curTotal As Double
Dim curPrice As Currency
Dim dblQty As Double
Dim lngCompanyID As Long
Dim currPrice As Currency

Dim blnReadOnly As Boolean
Dim flgLoading As Boolean
Dim WithEvents vCanAdd As z_BrokenRules
Attribute vCanAdd.VB_VarHelpID = -1
Dim WithEvents vCanIssue As z_BrokenRules
Attribute vCanIssue.VB_VarHelpID = -1




Private Sub cboMatch_SelectionChanged()
'MsgBox "Here"
    oCNLine.INVLineID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 4)
    oCNLine.INVLineCode = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 1)
End Sub

Private Sub cmdSelectCustomer_Click()
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    If oCN.SetCustomer(lngTPID) Then
        With oCN.customer
            txtPhone = .Phone
            txtAccnum = .AcNo
            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
            lblAddBill.Caption = .defaultaddress.AddressShort
            lblAddDel.Caption = .defaultaddress.AddressShort
        End With
        vCanAdd.RuleBroken "TP", False
    End If

End Sub


Private Sub Form_Terminate()
'        Set frmOL = Nothing
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub
Private Sub lvwLines_Click()
    Clipboard.SetText lvwLines.SelectedItem.Text
End Sub

Private Sub mnuAddresses_Click()
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.Component oCN
    frm.Show vbModal
    lblAddBill.Caption = oCN.BillTOAddress.AddressShort
    lblAddDel.Caption = oCN.DelToAddress.AddressShort

End Sub

Private Sub mnuDel_Click()
    RemoveDetailLine
End Sub


Private Sub mnuPrint_Click()
Dim frm As frmPrintingOptions_CN
    Set frm = New frmPrintingOptions_CN
    frm.Show vbModal

End Sub

Private Sub oCN_Valid(pMsg As String)
    bValidCN = (pMsg = "")
    cmdIssue.Enabled = (bValidCN And oCN.CNLines.Count > 0)
    cmdSave.Enabled = bValidCN
    Me.txtError = pMsg
End Sub

Sub oCNLine_ExtensionChange(lngExtension As Long, strExtension As String)
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
End Sub

Private Sub oCNLine_Valid(Msg As String)
    cmdEnter.Enabled = (Msg = "")
    txtError = Msg
End Sub

'Private Sub oCN_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
'    flgLoading = True
'    Me.txtRunningTotal = strtotal
'    lngCurrentTotal = lngTotal
''    Me.txtRunningDeposit = strTotalDeposit
'    lngCurrentDepositTotal = lngTotalDeposit
'    lngCurrentVATTotal = lngTotalVAT
'    flgLoading = False
'End Sub

Private Sub oCN_Reloadlist()
    LoadListView
End Sub
Private Sub oCN_Dirty(pVal As Boolean)
If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
End Sub
Private Sub oCN_CurrRowStatus(pMsg As String)
    MsgBox "CurrentRow Status = " & pMsg
End Sub






Sub vCanAdd_NobrokenRules()
    Me.cmdNewRows.Enabled = True
End Sub
Private Sub Form_Load()
Dim curTotalDeposit As Currency
    left = 10
    top = 10
    Width = 11100
    Height = 6700
    flgLoading = True
    oCN.GetStatus
    SetLvw
    SetEditFrameEnabled False, enNotEditing
    vMode = enNotEditing
    SetupCboMatch
    flgLoading = False
End Sub
Private Sub Form_Initialize()
    Set vCanAdd = New z_BrokenRules
'    Set frmOL = New frmOverlay1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If oCN.IsEditing Then oCN.CancelEdit
    
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oCN = Nothing
    Set tlCustomer = Nothing
    Set oCNLine = Nothing
End Sub

Public Sub Component(Optional pCO As a_CN)
    flgLoading = True
    If pCO Is Nothing Then
        Set oCN = New a_CN
        oCN.beginedit
        Me.lvwLines.Enabled = False
        SetControlsForNew
        vCanAdd.RuleBroken "TP", True
    Else
        Set oCN = pCO
        oCN.beginedit
        LoadCustomer
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        mnuFileCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        Me.lvwLines.Enabled = True
    End If
End Sub
Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
Dim lngColour As Long
    'A is adding, E is editing
    bFrameEnabled = pYesNo   'shared for use in all the form
    
    If (eMode = enAddingRow Or eMode = enNotEditing) And pYesNo Then
        Me.txtCode.Enabled = True
    Else
        Me.txtCode.Enabled = False
    End If
    Me.txtNote.Enabled = pYesNo
    Me.txtDiscount.Enabled = pYesNo
    Me.txtPrice.Enabled = pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtTotal.Enabled = pYesNo
    Me.txtAccnum.Enabled = Not pYesNo
    
    Me.cmdEnter.Enabled = pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidCN
    Me.cmdSave.Enabled = (Not pYesNo) And bValidCN And oCN.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
    Me.txtDiscount.BackColor = lngColour
End Sub
Private Sub SetControlsForNew()
    mnuFileCancel.Caption = "&Cancel"
    txtAccnum = ""
'    txtFax = ""
    txtPhone = ""
    txtStatus = "IN PROCESS"
End Sub

Private Sub cmdEnter_Click()
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
    
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        txtCode.SetFocus
        Exit Sub
    End If
    oCNLine.ApplyEdit
'    oCNLine.SetAsNEW
    oCNLine.beginedit

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add 1, oCNLine.Key
        LoadListViewLine oCNLine.Key, Me.lvwLines.ListItems(1)
        Set oCNLine = oCN.CNLines.Add
        oCNLine.SetQty 1
        oCNLine.TRID = oCN.TRID
        txtCode.SetFocus
    ElseIf vMode = enEditingRow Then
        LoadListViewLine lngSelectedRowIndex, Me.lvwLines.ListItems(lngSelectedRowIndex)
        cmdNewRows_Click
    End If
    
    ClearLineControls
 '   fSetTranslucency frmOL.hwnd, 200
End Sub


Private Sub cmdNewRows_Click()
Dim lr As Long
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = enEditingRow Then       'We have finished editing a row
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enNotEditing
        fr2.ZOrder 0
        fr1.ZOrder 1
        Me.lvwLines.Enabled = True
    ElseIf vMode = enAddingRow Then    'we are stopping adding rows
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enEditingRow
        fr2.ZOrder 0
        fr1.ZOrder 1
        Me.lvwLines.Enabled = True
    ElseIf vMode = enNotEditing Then  'we are starting to add rows
        cmdNewRows.Caption = "&Stop"
        SetEditFrameEnabled True, vMode
        vMode = enAddingRow
        fr2.ZOrder 1
        fr1.ZOrder 0
        Me.lvwLines.Enabled = False
        Me.txtCode.SetFocus
        Set oCNLine = oCN.CNLines.Add
        oCNLine.TRID = oCN.TRID
    End If

    ClearLineControls
End Sub
Private Sub LoadListView()
Dim lstItem As ListItem
Dim i As Long
    On Error GoTo ERR_Handler
    lvwLines.ListItems.Clear
    For i = 1 To oCN.CNLines.Count
        Set lstItem = lvwLines.ListItems.Add
        Set oCNLine = oCN.CNLines(i)
        LoadListViewLine i & "k", lstItem
    Next i
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub
Private Sub LoadListViewLine(i As String, lstItem As ListItem)
Dim currPrice As Currency
    With oCNLine
        lstItem.Text = .ProductCodeF
        If lstItem.Key = "" Then lstItem.Key = i
        lstItem.SubItems(1) = .TitleAuthorPublisher
        lstItem.SubItems(2) = .Qty
        lstItem.SubItems(3) = .INVLineCode
        lstItem.SubItems(4) = .PriceF(False)
        lstItem.SubItems(5) = .DiscountPercentF  ' Format(.DiscountPercent, "##0.0%")
        lstItem.SubItems(6) = .PLessDiscExtF(False)
    End With
End Sub
Private Sub lvwLines_DblClick()
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    lngEditingIdx = lvwLines.SelectedItem.Key
    Set oCNLine = oCN.CNLines(lngEditingIdx)
    If oCNLine.Product.DefaultCopy Is Nothing Then
        LoadMatchingInvoices oCN.customer.ID, oCNLine.Product.pID
    Else
        LoadMatchingInvoices oCN.customer.ID, "", oCNLine.Product.DefaultCopy.ID
    End If
    SetcboMatchToID (oCNLine.INVLineID)
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    
    txtTitle = oCNLine.Title
   ' txtPrice = oCNLine.PriceF
    txtQty = oCNLine.QtyF
    txtDiscount = oCNLine.DiscountPercentF
  '  txtPrice.SetFocus
    txtCode = oCNLine.ProductCode
    AutoSelect txtPrice
    
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    txtPrice.SetFocus
    fr2.ZOrder 1
    fr1.ZOrder 0
    cmdNewRows.Caption = "&Stop edit"
    
End Sub
Private Sub SetcboMatchToID(pID As Long)
    If pID > 0 Then
        cboMatch.Items.SelectItem(cboMatch.Items.FindItem(pID, 4)) = True
    End If
End Sub
'---------Companies code
'Private Sub LoadComps()
'Dim oCNmp As a_CNmpany
'Dim oItem As ListItem
'Dim i As Integer
'    If oCN.CompanyID > 0 Then
'        txtComp = oPC.Configuration.Companies(CStr(oCN.CompanyID)).CompanyName
'    Else
'        txtComp = oPC.Configuration.DefaultCompany.CompanyName
'        oCN.CompanyID = oPC.Configuration.DefaultCompanyID
'    End If
'End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    If oCN.customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
End Sub
'-------End Compsny code
'Private Sub txtOrdernum_Validate(Cancel As Boolean)
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oCNLine.COLineCode = txtOrdernum
'    If Err Then
'      Beep
'      intPos = txtOrdernum.SelStart
'      txtOrdernum = oCNLine.COLineCode
'      txtOrdernum.SelStart = intPos - 1
'    End If
'
'End Sub

Private Sub txtNote_Change()
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oCNLine.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCNLine.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    Cancel = Not oCNLine.setnote(txtNote)
End Sub
Private Sub txtNote_LostFocus()
    If flgLoading Then Exit Sub
    txtNote = oCNLine.Note
End Sub

Private Sub mnuEditNote_Click()
Dim ofrm As New frmNote
    ofrm.Component oCN
    ofrm.Show vbModal
    Unload ofrm
    Set ofrm = Nothing
End Sub

Private Sub mnuFileCancel_Click()
    If oCN.IsDirty Then
        oCN.CancelEdit
    End If
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    oCN.CancelEdit
    Unload Me
End Sub

Private Sub mnuFileOK_Click()
'    cmdOK_Click
End Sub

Private Sub mnuFilePrint_Click()
    cmdIssue_Click
End Sub
Private Sub mnuFileVoid_Click()
    oCN.SetStatus stVOID
    txtStatus = "Void"
End Sub
Private Sub txtAccNum_Validate(Cancel As Boolean)
Dim lngCustID As Long
Dim bResult As Boolean
    If Len(txtAccnum) > 0 Then
        bResult = oCN.SetCustomerFromAccNum(txtAccnum)
        If bResult Then
            With oCN.customer
                txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
                txtPhone = .Phone
                lblAddBill.Caption = .defaultaddress.AddressShort
                lblAddDel.Caption = .defaultaddress.AddressShort
            End With
            vCanAdd.RuleBroken "TP", False
        Else
            MsgBox "No such account number", , "Can't fetch customer"
            txtAccnum = ""
            Set oCustomer = Nothing
            Cancel = True
        End If
    End If
End Sub
'Private Sub txtComp_DblClick()
'Dim iCompIdx As Integer
'Dim i As Integer
'Start:
'    i = iCompIdx + 1
'    If i > oPC.Configuration.Companies.Count Then
'        i = 1
'    End If
'    If lngCompanyID = oPC.Configuration.Companies(i).ID Then
'        GoTo Start
'    End If
'    txtComp = oPC.Configuration.Companies(i).CompanyName
'    oCN.CompanyID = oPC.Configuration.Companies(i).ID
'    iCompIdx = i
'End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean

On Error GoTo ERR_Handler
    
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
  '  bOK = oCNLine.SetLineProduct("", txtCode)
    
    If Not oCNLine.Product.DefaultCopy Is Nothing Then
        LoadMatchingInvoices oCN.customer.ID, "", oCNLine.Product.DefaultCopy.ID
    ElseIf Not oCNLine.Product Is Nothing Then
        LoadMatchingInvoices oCN.customer.ID, oCNLine.Product.pID
    End If

    
    If bOK Then
        txtTitle = oCNLine.Title
    '    txtPrice = oCNLine.PriceF
        txtQty = oCNLine.QtyF
      '  txtDiscount = oCNLine.Discount
        txtPrice.SetFocus
        txtCode = oCNLine.ProductCode
        AutoSelect txtPrice
    Else
        MsgBox "Cannot find book on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
        GoTo EXIT_Handler
    End If
    oCNLine.GetStatus

EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub
Private Sub LoadMatchingInvoices(pTPID As Long, pPID As String, Optional pPIID As Long, Optional pCode As String)
Dim odINV As d_Invoice
Dim i As Integer
Dim bReturned As Boolean

    Set oInv = New c_Invoices
    oInv.Load bReturned, pTPID, , , , , pPID, pCode, pPIID
    If oInv.Count > 0 Then 'There are invoices for this item
        cboMatch.BeginUpdate
        ReDim ar(4, oInv.Count - 1)
        cboMatch.Items.RemoveAllItems
        i = 0
        For Each odINV In oInv
            ar(0, i) = odINV.TDateF
            ar(1, i) = odINV.Ref
            ar(2, i) = odINV.Qty
            ar(3, i) = odINV.PriceF
            ar(4, i) = odINV.InvoiceLineID
            i = i + 1
        Next
        cboMatch.PutItems ar
        cboMatch.EndUpdate
    End If
 '   fClearTranslucency frmOL.hwnd
End Sub
Private Sub cboMatch_Click()
 '   oCNLine.INVLineID = oInv.Item(1).InvoiceLineID
End Sub

Private Sub RemoveDetailLine()
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oCN.CNLines.Remove lvwLines.ListItems(i).Key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Exit Sub
    End If
    lvwLines.ListItems.Remove i
    lvwLines.Refresh
End Sub

Private Sub LoadCustomer()
    With oCN
        txtStatus = .statusF
        SetIssueButtonCaption
        txtAccnum = .TPAccNum
        txtPhone = .TPPhone
        txtPhone = .TPPhone
    End With
End Sub


Private Sub SaveCO()
On Error GoTo ERR_Handler
    
    oCN.post
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Public Sub PrintOrder()
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    On Error GoTo ERR_Handler
    
    Me.MousePointer = vbHourglass
    oCN.Load oCN.TRID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoCNLines Then
        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
        GoTo EXIT_Handler
    End If
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    Select Case Err
    Case 5941
        MsgBox "Book Mark on word document is missing", vbOKOnly + vbInformation, "Papyrus Information"
        Resume Next
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
    End Select
    Resume
End Sub
Private Sub cmdIssue_Click()
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmCNPreview

    If oCN.status = stInProcess Then
        If MsgBox("Issue this order?.  Confirm.", vbYesNo + vbQuestion, "Papyrus Invoicing Status") = vbNo Then
            Exit Sub
        End If
    End If
    oCN.SetStatus stISSUED
    
    strResult = oCN.post
    Set frm = New frmCNPreview
    frm.ComponentObject oCN
    frm.Show
    Unload Me
End Sub
Private Sub cmdSave_Click()
    oCN.SetStatus stInProcess
    SaveCO
    oCN.beginedit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    oCN.CancelEdit
    Unload Me
End Sub


Private Sub ClearLineControls()
    flgLoading = True
    Me.txtCode = ""
    Me.txtDiscount = ""
    Me.txtPrice = ""
    Me.txtTitle = ""
    Me.txtTotal = ""
    Me.txtNote = ""
  '  Me.txtdeposit = ""
    Me.txtQty = ""
    cboMatch.Items.RemoveAllItems
    cboMatch.SetFocus
    Me.cmdNewRows.SetFocus
    flgLoading = False
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
'Private Sub txtETA_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oCNLine.SetETA(txtETA) Then
'        Cancel = True
'    End If
'End Sub
'Private Sub txtETA_GotFocus()
'    AutoSelect Controls("txtETA")
'End Sub
'
'Private Sub txtETA_LostFocus()
'    txtETA = oCNLine.ETAF
'End Sub

Private Sub txtPrice_GotFocus()
    AutoSelect Controls("txtPrice")
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
  '  If Not oCNLine.SetPrice(txtPrice) Then
  '      Cancel = True
  '  End If
End Sub
Private Sub txtPrice_LostFocus()
  '  txtPrice = oCNLine.PriceF
End Sub
Private Sub txtQty_GotFocus()
    AutoSelect Controls("txtQty")
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oCNLine.SetQty(txtQty) Then
        Cancel = True
    End If
End Sub
Private Sub txtQty_LostFocus()
  '  txtQty = oCNLine.QtyF
End Sub
Private Sub txtDiscount_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oCNLine.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
End Sub
Private Sub txtDiscount_LostFocus()
  '  txtDiscount = oCNLine.DiscountPercentF
End Sub
Private Sub txtDiscount_GotFocus()
    AutoSelect Controls("txtDiscount")
End Sub
'Private Sub txtDeposit_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oCNLine.SetDeposit(txtdeposit) Then
'        Cancel = True
'    End If
'End Sub
'Private Sub txtDeposit_LostFocus()
' '   txtdeposit = oCNLine.DepositF
'End Sub
'Private Sub txtDeposit_GotFocus()
'    AutoSelect Controls("txtDeposit")
'End Sub


Private Sub SetIssueButtonCaption()
        If oCN.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oCN.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
End Sub
Private Sub txtAccNum_LostFocus()
    txtAccnum = UCase(txtAccnum)
End Sub


Private Sub lvwLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvwLines.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
    If lvwLines.SortOrder = lvwAscending Then
        lvwLines.SortOrder = lvwDescending
    Else
        lvwLines.SortOrder = lvwAscending
    End If
   lvwLines.Sorted = True
End Sub
Private Sub SetLvw()
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwLines.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvwLines.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


End Sub
Sub SetupCboMatch()
        cboMatch.BeginUpdate
        cboMatch.WidthList = 360
        cboMatch.HeightList = 162
        cboMatch.AllowSizeGrip = True
        cboMatch.AutoDropDown = True
        
        cboMatch.Columns.Add "Date"
        cboMatch.Columns.Add "Code"
        cboMatch.Columns.Add "Qty"
        cboMatch.Columns.Add "Price"
        cboMatch.Columns.Add "INVID"
        cboMatch.Columns(0).Width = 70
        cboMatch.Columns(1).Width = 70
        cboMatch.Columns(2).Width = 30
        cboMatch.Columns(3).Width = 70
        cboMatch.Columns(4).Width = 0
        cboMatch.BackColorLock = Me.BackColor
        
        cboMatch.EndUpdate
        

End Sub

