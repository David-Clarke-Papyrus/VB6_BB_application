VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmCN 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Credit note"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmCN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.CheckBox chkChargeVAT 
      BackColor       =   &H00D3D3CB&
      Caption         =   "&Discount VAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   810
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5430
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2205
      Left            =   180
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1230
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3889
      SortKey         =   7
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
      NumItems        =   9
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
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Sa&ve"
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
      Picture         =   "frmCN.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
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
      Height          =   885
      Left            =   3615
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5175
      Width           =   3795
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
      Height          =   705
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5415
      Width           =   630
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   7500
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmCN.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1110
   End
   Begin VB.TextBox txtRunningTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Left            =   9630
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3495
      Width           =   1200
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
      Picture         =   "frmCN.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1050
      Left            =   2850
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   60
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1852
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin CoolButtonControl.CoolButton cbComp 
      Height          =   360
      Left            =   735
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   90
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   635
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1350
      Left            =   135
      TabIndex        =   10
      Top             =   3705
      Width           =   10710
      Begin VB.CommandButton cmdDam 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6300
         Picture         =   "frmCN.frx":1028
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Click to mark quantity of damaged stock being returned"
         Top             =   465
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelMatch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Click to clear all matches and credit a missing invoice"
         Top             =   195
         Width           =   255
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboMatch 
         Height          =   375
         Left            =   1710
         OleObjectBlob   =   "frmCN.frx":13B2
         TabIndex        =   2
         Top             =   480
         Width           =   3945
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5655
         TabIndex        =   3
         Top             =   465
         Width           =   630
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4995
         TabIndex        =   6
         Top             =   885
         Width           =   4530
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7770
         TabIndex        =   5
         Top             =   465
         Width           =   735
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9585
         MaskColor       =   &H00C4BCA4&
         Picture         =   "frmCN.frx":275C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   645
         Width           =   1000
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
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
         Height          =   330
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   885
         Width           =   3975
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6750
         TabIndex        =   4
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label lblDam 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   6270
         TabIndex        =   34
         Top             =   165
         Width           =   525
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Original invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1815
         TabIndex        =   23
         Top             =   165
         Width           =   1470
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   5655
         TabIndex        =   22
         Top             =   150
         Width           =   525
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00DFDDEA&
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4425
         TabIndex        =   20
         Top             =   870
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Disc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   7560
         TabIndex        =   14
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label9 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   6870
         TabIndex        =   12
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "From"
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
      TabIndex        =   30
      Top             =   120
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3015
      Picture         =   "frmCN.frx":2AE6
      Stretch         =   -1  'True
      Top             =   615
      Width           =   360
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "To:"
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
      Left            =   2940
      TabIndex        =   28
      Top             =   135
      Width           =   375
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3465
      TabIndex        =   27
      Top             =   150
      Width           =   1545
   End
   Begin VB.Label lblTPPhone 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3465
      TabIndex        =   26
      Top             =   465
      Width           =   1545
   End
   Begin VB.Label lblTPFax 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3465
      TabIndex        =   25
      Top             =   780
      Width           =   1545
   End
   Begin VB.Label lblAddBill 
      BackColor       =   &H00D3D3CB&
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
      Left            =   5715
      TabIndex        =   19
      Top             =   150
      Width           =   1920
   End
End
Attribute VB_Name = "frmCN"
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
Dim bDamagedReturns As Boolean
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

Private Sub chkChargeVAT_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCN.ShowVAT = (chkChargeVAT = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.chkChargeVAT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuChangeCustomer_Click()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    LoadNewCustomer lngTPID
    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuChangeCustomer_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancelMatch_Click()
    On Error GoTo errHandler
Dim i As Integer

    If cboMatch.Items.ItemCount > 0 Then
        oCNLine.INVLineID = 0
        For i = 0 To cboMatch.Items.ItemCount - 1
            cboMatch.Items.SelectItem(cboMatch.Items(i)) = False
        Next
    End If
    
'    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
'        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
'        Cancel = True
'        GoTo EXIT_Handler
'    End If
'    bOK = oCNLine.SetLineProduct("", txtCode)
'    If bOK Then
'        txtTitle = oCOLine.Title
'        txtPrice = oCOLine.Price
'        If oPC.AllowsSSInvoicing Then
'            txtQty = oCOLine.QtyFirmF
'            txtQtySS = oCOLine.QtySSF
'        Else
'            txtQty = oCOLine.QtyF
'        End If
'        txtdeposit = oCOLine.Deposit
'        txtDiscount = oCOLine.Discount
'        mSetfocus txtPrice
'        txtCode = oCOLine.Ean
'        txtETA = oCOLine.ETAF
'        If oCO.OrderRef > "" Then
'            txtOrdernum = oCO.OrderRef
'        End If
'        AutoSelect txtPrice
'    Else
'        Dim frmAdHoc As frmAdHocProduct
'        Set frmAdHoc = New frmAdHocProduct
'        frmAdHoc.Component txtCode
'        frmAdHoc.Show vbModal
'        txtCode = frmAdHoc.code
'        Unload frmAdHoc
'        Set frmAdHoc = Nothing
'        Cancel = True
'        GoTo START
'    End If
    
    
    
    
    
    
    oCNLine.INVLineID = 0
    oCNLine.INVLineCode = ""
    oCNLine.VATRate = 14
    oCN.ForeignCurrencyID = 0
    oCN.CurrRate = 0
    oCNLine.InvPrice = 0
    oCNLine.InvLineDate = CDate(0)
    oCNLine.SetQty 1, 9999
    oCNLine.SetDiscount 0
    oCNLine.InvPrice = oCNLine.product.SP
    txtPrice = oCNLine.Price
    txtQty = oCNLine.QtyF
    txtTitle = oCNLine.TitleF(25)
    txtCode = oCNLine.EAN
    Me.txtDiscount = 0
    
    txtPrice.Enabled = True
    txtDiscount.Enabled = True
    txtPrice.BackColor = &H80000005
    txtDiscount.BackColor = &H80000005
    mSetfocus txtQty

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdCancelMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDam_Click()
    On Error GoTo errHandler
Dim frm As New frmDam
    frm.Show vbModal
    oCNLine.SetQtyDam frm.DamagedQty
    lblDam.Caption = oCNLine.QtyDamF
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdDam_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oCN.statusF = "IN PROCESS" And Not oCN.IsNew)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SetMenu"
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayoutLvw Me.lvwLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuSaveLayout"
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oCN.Memo
    ofrm.Show vbModal
    oCN.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuMemo"
End Sub



Private Sub cboMatch_SelectionChanged()
    On Error GoTo errHandler
'MsgBox "Here"
    If cboMatch.Items.SelectCount = 0 Then Exit Sub
    If Not oCNLine.CNParent Is Nothing Then
        If oCNLine.CNParent.ForeignCurrencyID <> cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9) And oCNLine.CNParent.ForeignCurrencyID > 0 And oCN.CNLines.Count > 0 Then
            MsgBox "You cannot mix credits for products from invoices issued in different currencies"
            Exit Sub
        End If
    End If

    oCNLine.INVLineID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 4)
    oCNLine.INVLineCode = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 1)
    oCNLine.VATRate = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 7)
    oCN.ForeignCurrencyID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9)
    oCN.CurrRate = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 10)
    oCNLine.InvPrice = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 5)
    oCNLine.InvLineDate = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 8)
    oCNLine.SetQty cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2), cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2)
    oCNLine.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 6)
    txtPrice = oCNLine.PriceF(False)
    txtQty = oCNLine.QtyF
    Me.txtDiscount = oCNLine.DiscountPercentF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cboMatch_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbComp_Click()
    On Error GoTo errHandler
    oCN.COMPID = OptionLoop(oCN.COMPID, oPC.Configuration.Companies.Count)
    cbComp.Caption = oPC.Configuration.Companies(oCN.COMPID).CompanyName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cbComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oCN.Customer.ID > 0 Then
        frm.component oCN.Customer
        frm.Show
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdSelectCustomer_Click()
'Dim lngTPID As Long
'Dim frm As frmBrowseCustomers2
'    Set frm = New frmBrowseCustomers2
'    frm.Show vbModal
'    lngTPID = frm.CustomerID
'    If oCN.SetCustomer(lngTPID) Then
'        With oCN.customer
'            txtPhone = .Phone
'            txtAccnum = .AcNo
'            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
'            lblAddBill.Caption = .BillToADdress.AddressShort
'            lblAddDel.Caption = .BillToADdress.AddressShort
'        End With
'        vCanAdd.RuleBroken "TP", False
'    End If
'
'End Sub
'
'
Private Sub Form_Terminate()
    On Error GoTo errHandler
'        Set frmOL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_Click()
    On Error GoTo errHandler
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(lvwLines.SelectedItem.SubItems(8), ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.lvwLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAddresses_Click()
    On Error GoTo errHandler
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oCN
    frm.Show vbModal
    lblAddBill.Caption = oCN.billtoaddress.AddressShort

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuAddresses_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuDel_Click()
'    RemoveDetailLine
'End Sub
Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuDelLine"
End Sub


Private Sub mnuPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_CN
    Set frm = New frmPrintingOptions_CN
    frm.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oCN_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidCN = (pMsg = "")
    cmdIssue.Enabled = (bValidCN And oCN.CNLines.Count > 0)
    cmdSave.Enabled = bValidCN
    Me.txtError = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCN_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oCNLine_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCNLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oCNLine_Valid(msg As String)
    On Error GoTo errHandler
    cmdEnter.Enabled = (msg = "")
    txtError = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCNLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oCN_TotalChange(lngTotal As Long, strtotal As String, lngTotalVAT As Long, strTotalVAT As String)
    On Error GoTo errHandler
    flgLoading = True
    Me.txtRunningTotal = strtotal
    lngCurrentTotal = lngTotal
    lngCurrentVATTotal = lngTotalVAT
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCN_TotalChange(lngTotal,strtotal,lngTotalVAT,strTotalVAT)", Array(lngTotal, _
         strtotal, lngTotalVAT, strTotalVAT), EA_NORERAISE
    HandleError
End Sub

Private Sub oCN_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCN_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oCN_Dirty(pVal As Boolean)
    On Error GoTo errHandler
If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCN_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oCN_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.oCN_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadComps()
    On Error GoTo errHandler
Dim oComp As a_Company
Dim oItem As ListItem
Dim i As Integer
    If oCN.COMPID > 0 Then
        cbComp.Caption = oPC.Configuration.Companies(CStr(oCN.COMPID)).CompanyName
    Else
        cbComp.Caption = oPC.Configuration.DefaultCompany.CompanyName
        oCN.COMPID = oPC.Configuration.DefaultCOMPID
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.LoadComps"
End Sub




Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim curTotalDeposit As Currency
Dim strAddress As String
    If Me.WindowState <> 2 Then
        Left = 10
        top = 10
        Width = 11100
        Height = 6700
    End If
    bDamagedReturns = (oPC.getProperty("DamagedReturns") = "TRUE")
    SetupcboMatch
    Me.cmdDam.Visible = bDamagedReturns
    flgLoading = True
    LoadComps
    If oCN.BillToAddressID > 0 Then
        strAddress = oCN.billtoaddress.AddressMailing
    End If
    Me.lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
    Me.chkChargeVAT = IIf(oCN.ShowVAT, 1, 0)
'    SetEditFrameEnabled False, enNotEditing
    vMode = enNotEditing
    SetupcboMatch
    flgLoading = False
    oCN.GetStatus
    SetLvw
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules

'    Set frmOL = New frmOverlay1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oCN.IsEditing Then oCN.CancelEdit
    UnsetMenu
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oCN = Nothing
    Set tlCustomer = Nothing
    Set oCNLine = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Sub component(Optional pCustID As Long, Optional pCN As a_CN)
    On Error GoTo errHandler
    
    flgLoading = True
    If pCN Is Nothing Then
        Set oCN = New a_CN
        oCN.BeginEdit
        lvwLines.Enabled = False
   '     SetControlsForNew
        lvwLines.Height = 2200
        vCanAdd.RuleBroken "TP", True
        If pCustID > 0 Then
            LoadNewCustomer pCustID
        End If
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        mSetfocus txtCode
        Set oCNLine = oCN.CNLines.Add ' Therefore oCNLine is in Editing mode
        oCNLine.SetQty 1, 1
    Else
        Set oCN = pCN
        oCN.BeginEdit
        LoadCustomer
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        lvwLines.Enabled = True
        lvwLines.Height = 4000
        SetEditFrameEnabled False, enNotEditing
        vMode = enNotEditing
    End If
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.component(pCustID,pCN)", Array(pCustID, pCN)
End Sub
Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
    'A is adding, E is editing
    bFrameEnabled = pYesNo   'shared for use in all the form
    
    If (eMode = enAddingRow Or eMode = enNotEditing) And pYesNo Then
        Me.txtCode.Enabled = True
    Else
        Me.txtCode.Enabled = False
    End If
    Me.txtNote.Enabled = pYesNo
    Me.txtDiscount.Enabled = False  'pYesNo
    Me.txtPrice.Enabled = False ''pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtTotal.Enabled = pYesNo
    Me.txtQty.Enabled = pYesNo
    
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
    Me.txtPrice.BackColor = 14416635  'lngColour
    Me.txtDiscount.BackColor = 14416635  'lngColour
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub
'Private Sub SetControlsForNew()
'    mnuFileCancel.Caption = "&Cancel"
'    Me.lblTPPhone = ""
'End Sub
Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oCN.SetCustomer(plngTPID) Then
        With oCN.Customer
            If Not .billtoaddress Is Nothing Then
                lblTPPhone.Caption = .billtoaddress.Phone
                lblTPFax.Caption = .billtoaddress.Fax
                oCN.SetBillToAddress .billtoaddress
                oCN.setDelTOAddress .DelTOAddress
                lblAddBill.Caption = .billtoaddress.AddressShort
            End If
            lblTPName.Caption = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
        End With
        vCanAdd.RuleBroken "TP", False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.LoadNewCustomer(plngTPID)", plngTPID
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
    
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        mSetfocus txtCode
        Exit Sub
    End If
    oCNLine.ApplyEdit  'adds oCNLine to mcolItems
    oCNLine.BeginEdit

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add 1, oCNLine.key
        LoadListViewLine lvwLines.ListItems(lvwLines.ListItems.Count), oCNLine
        Set oCNLine = Nothing
        Set oCNLine = oCN.CNLines.Add
        oCNLine.SetQty 1, 1
        oCNLine.TRID = oCN.TRID
        'ocnline.i
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwLines.ListItems(lngEditingIdx), oCNLine
        cmdNewRows_Click
    End If
    oCN.CalculateTotals
    Me.txtRunningTotal = oCN.TotalPayableF(False)
    oCN.GetStatus
    ClearLineControls
    txtPrice.BackColor = &HDBFAFB
    txtDiscount.BackColor = &HDBFAFB
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdEnter_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
Dim lr As Long
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
        LogSaveToFile "Invoice New row button:enEditingRow"
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        LogSaveToFile "Invoice New row button:enAddingRow"
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        LogSaveToFile "Invoice New row button:enNotEditing"
        ChangeState enAddingRow
    End If

    ClearLineControls
'    If vMode = enEditingRow Then       'We have finished editing a row
'        cmdNewRows.Caption = "&Add"
'        SetEditFrameEnabled False, vMode
'        vMode = enNotEditing
'        lvwLines.Height = 4000
'        lvwLines.Enabled = True
'    ElseIf vMode = enAddingRow Then    'we are stopping adding rows
'        cmdNewRows.Caption = "&Add"
'        SetEditFrameEnabled False, vMode
'        vMode = enNotEditing 'enEditingRow
'        lvwLines.Enabled = True
'        lvwLines.Height = 4000
'        txtError = ""
'    ElseIf vMode = enNotEditing Then  'we are starting to add rows
'        cmdNewRows.Caption = "&Stop"
'        SetEditFrameEnabled True, vMode
'        vMode = enAddingRow
'        lvwLines.Enabled = False
'        lvwLines.Height = 2200
'        mSetfocus txtCode
'        Set oCNLine = oCN.CNLines.Add
'        mSetfocus txtCode
'        oCNLine.SetQty 1, 1
'        oCNLine.TRID = oCN.TRID
'    End If
'    ClearLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ChangeState(pToMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
    vMode = pToMode

    Select Case pToMode
    Case eneditingrow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQty.Enabled = True
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oCN.CNLines.Count > 0)
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        UnsetMenu
        fr1.ZOrder 1
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQty.Enabled = True
        txtError = ""
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oCN.CNLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        lblTPPhone.Caption = ""
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        ClearLineControls
        fr1.ZOrder 1
        mSetfocus txtCode
        Set oCNLine = oCN.CNLines.Add
        oCNLine.TRID = oCN.TRID
        oCNLine.SetQty 1, 1
        UnsetMenu
    Case enNotEditing
        flgLoading = True
        fr1.Visible = False
        txtError = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = True  '(oInvoice.InvoiceLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        lvwLines.Enabled = True
        lvwLines.Height = 4000
        SetMenu
        fr1.ZOrder 1
    End Select
    If Not oCN.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.ChangeState(pToMode)", pToMode
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To lvwLines.ColumnHeaders.Count
        lvwLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwLines.ColumnHeaders(i).Width)
    Next
    lvwLines.ListItems.Clear
    For i = 1 To oCN.CNLines.Count
        Set lstItem = lvwLines.ListItems.Add
        LoadListViewLine lstItem, oCN.CNLines(i)
    Next i
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oCNLine As a_CNL)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oCNLine
        lstItem.Text = .ProductCodeF
        lstItem.key = .key  '"" Then lstItem.Key = i
        lstItem.SubItems(1) = .TitleAuthorPublisher
        lstItem.SubItems(2) = .QtyComboF
        lstItem.SubItems(3) = .INVLineCode
        lstItem.SubItems(4) = .PriceF(False)
        lstItem.SubItems(5) = .DiscountPercentF  ' Format(.DiscountPercent, "##0.0%")
        lstItem.SubItems(6) = .PLessDiscExtF(False)
        lstItem.SubItems(8) = .EAN
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.LoadListViewLine(lstItem,oCNLine)", Array(lstItem, oCNLine)
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    If lvwLines.SelectedItem.Index < 1 Then Exit Sub
    
    lngEditingIdx = lvwLines.SelectedItem.key
    
    Set oCNLine = Nothing
    Set oCNLine = oCN.CNLines(lngEditingIdx)
    
    ChangeState eneditingrow
    
    cboMatch.BeginUpdate
    cboMatch.Items.RemoveAllItems
    cboMatch.EndUpdate
    
    If oCNLine.CopyID = 0 Then
        LoadMatchingInvoices oCN.Customer.ID, oCNLine.PID
    Else
        LoadMatchingInvoices oCN.Customer.ID, "", oCNLine.CopyID
    End If
    If cboMatch.Items.ItemCount > 0 Then
        SetcboMatchToID oCNLine.INVLineID
    End If
    
    If oCNLine.INVLineID > 0 Then
        SetcboMatchToID (oCNLine.INVLineID)
        txtNote = oCNLine.Note
        mSetfocus txtNote
    Else
        txtPrice = oCNLine.Price
        txtQty = oCNLine.QtyF
        txtDiscount = oCNLine.Discount
        mSetfocus txtPrice
        txtNote = oCNLine.Note
        txtPrice.BackColor = &H80000005
        txtDiscount.BackColor = &H80000005
        AutoSelect txtPrice
    End If
    If bDamagedReturns Then
        lblDam.Caption = oCNLine.QtyDamF
    End If
    txtTitle = oCNLine.TitleF(25)
    txtCode = oCNLine.EAN
    
    oCNLine.GetStatus
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetcboMatchToID(pILID As Long)
    On Error GoTo errHandler
    If pILID > 0 Then
        cboMatch.Items.SelectItem(cboMatch.Items.FindItem(pILID, 4)) = True
    '    oCNLine.SetPrice cboMatch.Items.CellCaption(cboMatch.Items(0), 5)
        If Not oCNLine.qty > 0 Then
            oCNLine.SetQty cboMatch.Items.CellCaption(cboMatch.Items(0), 2), cboMatch.Items.CellCaption(cboMatch.Items(0), 2)
        End If
        oCNLine.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items(0), 6)
        If oPC.Configuration.CaptureDecimal Then
            txtPrice = oCNLine.PriceF(False)
        Else
            txtPrice = oCNLine.Price
        End If
        txtQty = oCNLine.QtyComboF
        txtDiscount = oCNLine.DiscountPercentF
      '  txtQty.Enabled = False
        txtPrice.Enabled = False
        txtDiscount.Enabled = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SetcboMatchToID(pILID)", pILID
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
    On Error GoTo errHandler
    If oCN.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
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
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    txtNote = HandleTextWithBites(txtNote)
    On Error Resume Next
    oCNLine.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCNLine.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCNLine.setnote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oCNLine.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuEditNote_Click()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oCN
    ofrm.Show vbModal
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuEditNote_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileCancel_Click()
    On Error GoTo errHandler
    If oCN.IsDirty Then
        oCN.CancelEdit
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuFileCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo errHandler
    oCN.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuFileExit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFileOK_Click()
    On Error GoTo errHandler
'    cmdOK_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuFileOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuFilePrint_Click()
    On Error GoTo errHandler
    cmdIssue_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuFilePrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuFileVoid_Click()
    On Error GoTo errHandler
    oCN.SetStatus stVOID
    oCN.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.mnuFileVoid_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub txtAccNum_Validate(Cancel As Boolean)
'Dim lngCustID As Long
'Dim bResult As Boolean
'    If Len(txtAccnum) > 0 Then
'        bResult = oCN.SetCustomerFromAccNum(txtAccnum)
'        If bResult Then
'            With oCN.customer
'                txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
'                txtPhone = .Phone
'                lblAddBill.Caption = .BillToADdress.AddressShort
'                lblAddDel.Caption = .BillToADdress.AddressShort
'            End With
'            vCanAdd.RuleBroken "TP", False
'        Else
'            MsgBox "No such account number", , "Can't fetch customer"
'            txtAccnum = ""
'            Set oCustomer = Nothing
'            Cancel = True
'        End If
'    End If
'End Sub
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
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean

    
    If txtCode = "" Or vMode = eneditingrow Then Exit Sub
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_Handler
    End If
    bOK = oCNLine.SetLineProduct("", txtCode)
    cboMatch.BeginUpdate
    cboMatch.Items.RemoveAllItems
    cboMatch.EndUpdate

    If Not oCNLine.product.DefaultCopy Is Nothing Then
        LoadMatchingInvoices oCN.Customer.ID, "", oCNLine.product.DefaultCopy.ID
    ElseIf Not oCNLine.product Is Nothing Then
        LoadMatchingInvoices oCN.Customer.ID, oCNLine.PID
    End If
    
    If bOK Then
    Else
        MsgBox "Cannot find book on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
        GoTo EXIT_Handler
    End If
    If cboMatch.Items.ItemCount > 0 Then
            cboMatch.Items.SelectItem(cboMatch.Items(0)) = True
        oCNLine.InvPrice = cboMatch.Items.CellCaption(cboMatch.Items(0), 5)
        oCNLine.SetQty cboMatch.Items.CellCaption(cboMatch.Items(0), 2), cboMatch.Items.CellCaption(cboMatch.Items(0), 2)
        oCNLine.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items(0), 6)
        oCNLine.InvLineDate = cboMatch.Items.CellCaption(cboMatch.Items(0), 8)
        If oPC.Configuration.CaptureDecimal Then
            txtPrice = oCNLine.PriceF(False)
        Else
            txtPrice = oCNLine.Price
        End If
        txtTitle = oCNLine.TitleF(25)
        txtCode = oCNLine.EAN
        txtQty = oCNLine.qty
        txtDiscount = oCNLine.DiscountPercentF
        txtPrice.Enabled = False
        txtDiscount.Enabled = False
        mSetfocus txtNote
    End If
    oCNLine.GetStatus

EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadMatchingInvoices(pTPID As Long, pPID As String, Optional pPIID As Long, Optional pCode As String)
    On Error GoTo errHandler
Dim odINV As d_Invoice
Dim i As Integer
Dim bReturned As Boolean

    Set oInv = New c_Invoices
    oInv.Load bReturned, pTPID, "", "", , , , pPID, pCode, pPIID, "N"
    If oInv.Count > 0 Then 'There are invoices for this item
        cboMatch.BeginUpdate
        ReDim ar(10, oInv.Count - 1)
        cboMatch.Items.RemoveAllItems
        i = 0
        For Each odINV In oInv
            If odINV.statusF = "COMPLETE" Then
                ar(0, i) = odINV.TDateF
                ar(1, i) = odINV.Ref  'odINV.Doccode & " : " &
                ar(2, i) = odINV.qty
                ar(3, i) = odINV.PriceF
                ar(4, i) = odINV.InvoiceLineID
                ar(5, i) = odINV.Price
                ar(6, i) = odINV.Discount
                ar(7, i) = odINV.VATRate
                ar(8, i) = odINV.tDate
                ar(9, i) = odINV.CURRID
                ar(10, i) = odINV.CurrRate
                i = i + 1
            End If
        Next
        cboMatch.PutItems ar
        cboMatch.EndUpdate
    End If
 '   fClearTranslucency frmOL.hwnd
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.LoadMatchingInvoices(pTPID,pPID,pPIID,pCODE)", Array(pTPID, pPID, pPIID, pCode)
End Sub


Private Sub RemoveDetailLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oCN.CNLines.Remove lvwLines.ListItems(i).key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Exit Sub
    End If
    lvwLines.ListItems.Remove i
    lvwLines.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.RemoveDetailLine"
End Sub

Private Sub LoadCustomer()
    On Error GoTo errHandler
    With oCN
        SetIssueButtonCaption
        Me.lblTPName.Caption = .TPNAME
        Me.lblTPPhone = .TPPhone
        Me.lblTPFax = .TPFax
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.LoadCustomer"
End Sub


Private Sub SaveCO()
    On Error GoTo errHandler
    
    oCN.Post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SaveCO"
End Sub

Public Sub PrintOrder()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oCN.Load oCN.TRID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoCNLines Then
        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
        GoTo EXIT_Handler
    End If
    
EXIT_Handler:
    Me.MousePointer = vbDefault
'ERR_Handler:
'    Select Case Err
'    Case 5941
'        MsgBox "Book Mark on word document is missing", vbOKOnly + vbInformation, "Papyrus Information"
'        Resume Next
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'    End Select
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.PrintOrder"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmCNPreview

    If oPC.Configuration.Signtransactions = True Then
        If SecurityControl(enSECURITY_CN_SIGN, , "Sign this credit note", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oCN.Status = stInProcess Then
            If MsgBox("Issue this credit note?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If


    WaitMsg "Issuing credit note  . . .", True, Me
    oCN.SetStatus stISSUED
    oCN.StaffID = gSTAFFID
    
    strResult = oCN.Post
    Set frm = New frmCNPreview
    frm.ComponentObject oCN
    frm.Show
    WaitMsg "", False, Me
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
    oCN.SetStatus stInProcess
    oCN.RecalculateAllLines
    oCN.CalculateTotals
    SaveCO
    oCN.BeginEdit
    Set oCNLine = oCN.CNLines.Add
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmCNPreview

    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmCNPreview
        frm.ComponentObject oCN
        frm.Show
    End If
    
    If cmdCancel.Caption <> "&Close" Then
        If oCN.IsEditing And oCN.IsDirty Then
            If MsgBox("You wish to cancel?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
            oCN.CancelEdit
        End If
    End If
    
    If Not oCNLine Is Nothing Then
        If oCNLine.IsEditing Then oCNLine.CancelEdit
    End If

    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    txtPrice.Enabled = True
    txtQty.Enabled = True
    txtDiscount.Enabled = True
    Me.txtCode = ""
    Me.txtDiscount = ""
    Me.txtPrice = ""
    Me.txtTitle = ""
    Me.txtTotal = ""
    Me.txtNote = ""
    Me.txtQty = ""
    cboMatch.Items.RemoveAllItems
    mSetfocus cmdNewRows
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.ClearLineControls"
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
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
    On Error GoTo errHandler
    AutoSelect Controls("txtPrice")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
   ' MsgBox "Code commented"
    If Not oCNLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oCNLine.CalculateLine
    txtTotal = oCNLine.PLessDiscExtF(False)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_LostFocus()
    On Error GoTo errHandler
   ' txtPrice = oCNLine.PriceF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtPrice_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtQty")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If cboMatch.Items.SelectCount <> 0 Then
        If Not oCNLine.SetQty(txtQty, cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2)) Then
            Cancel = True
        End If
    Else
        oCNLine.SetQty txtQty, 20000
    End If
    oCNLine.CalculateLine
    txtTotal = oCNLine.PLessDiscExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_LostFocus()
    On Error GoTo errHandler
    txtQty = oCNLine.QtyF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtQty_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCNLine.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
    oCNLine.CalculateLine
    txtTotal = oCNLine.PLessDiscExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_LostFocus()
    On Error GoTo errHandler
  '  txtDiscount = oCNLine.DiscountPercentF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtDiscount_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDiscount")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
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
    On Error GoTo errHandler
        If oCN.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oCN.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SetIssueButtonCaption"
End Sub
'Private Sub txtAccNum_LostFocus()
'    txtAccnum = UCase(txtAccnum)
'End Sub


Private Sub lvwLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
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


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SetLvw"
End Sub
Sub SetupcboMatch()
    On Error GoTo errHandler
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
    cboMatch.Columns.Add "lngPRICE"
    cboMatch.Columns.Add "dblDiscount"
    cboMatch.Columns.Add "dblVATRate"
    cboMatch.Columns.Add "dteInvDate"
    
    cboMatch.Columns(0).Width = 70
    cboMatch.Columns(1).Width = 70
    cboMatch.Columns(2).Width = 30
    cboMatch.Columns(3).Width = 70
    cboMatch.Columns(4).Width = 0
    cboMatch.Columns(5).Width = 0
    cboMatch.Columns(6).Width = 0
    cboMatch.Columns(7).Width = 0
    cboMatch.Columns(8).Width = 0
    cboMatch.BackColorLock = Me.BackColor
    cboMatch.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCN.SetupcboMatch"
End Sub

