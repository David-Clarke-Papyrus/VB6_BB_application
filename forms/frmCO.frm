VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "COOLBU~1.OCX"
Begin VB.Form frmCO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Order"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmCO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.TextBox txtMemo 
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
      ForeColor       =   &H8000000D&
      Height          =   960
      Left            =   3330
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   5325
      Width           =   3975
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2265
      Left            =   90
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1215
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3995
      SortKey         =   9
      View            =   3
      SortOrder       =   -1  'True
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ef."
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
         Text            =   "Deposit"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "E.T.A"
         Object.Width           =   2048
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
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
      Left            =   8550
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmCO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5295
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
      Height          =   960
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdNewRows 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5370
      Width           =   780
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
      Left            =   7455
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmCO.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5295
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
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6630
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3480
      Width           =   4140
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1605
      Left            =   135
      TabIndex        =   12
      Top             =   3660
      Width           =   10650
      Begin VB.TextBox txtOrdernum 
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
         Left            =   2700
         TabIndex        =   3
         Top             =   450
         Width           =   1365
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
         Left            =   8940
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   450
         Width           =   1620
      End
      Begin VB.TextBox txtETA 
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
         Left            =   6675
         TabIndex        =   7
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtQty 
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
         Left            =   1740
         TabIndex        =   2
         Top             =   450
         Width           =   915
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
         Left            =   4830
         TabIndex        =   8
         Top             =   1140
         Width           =   4815
      End
      Begin VB.TextBox txtDiscount 
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
         Left            =   5130
         TabIndex        =   5
         Top             =   450
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
         Height          =   600
         Left            =   9735
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   930
         Width           =   840
      End
      Begin VB.TextBox txtdeposit 
         Alignment       =   1  'Right Justify
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
         Left            =   5895
         TabIndex        =   6
         Top             =   450
         Width           =   750
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
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1140
         Width           =   3975
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
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
         Left            =   4095
         TabIndex        =   4
         Top             =   450
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
         Top             =   450
         Width           =   1650
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Total"
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
         Height          =   240
         Left            =   9495
         TabIndex        =   37
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "ETA"
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
         Height          =   300
         Left            =   6975
         TabIndex        =   29
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   195
         Width           =   645
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   300
         Left            =   4215
         TabIndex        =   28
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   1950
         TabIndex        =   27
         Top             =   195
         Width           =   570
      End
      Begin VB.Label Label3 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Ref."
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
         Height          =   300
         Left            =   3165
         TabIndex        =   26
         Top             =   195
         Width           =   675
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
         Height          =   300
         Left            =   5175
         TabIndex        =   17
         Top             =   195
         Width           =   585
      End
      Begin VB.Label Label11 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Deposit"
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
         Height          =   300
         Left            =   5865
         TabIndex        =   16
         Top             =   195
         Width           =   810
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
         Height          =   285
         Left            =   660
         TabIndex        =   15
         Top             =   210
         Width           =   540
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
         Height          =   300
         Left            =   4305
         TabIndex        =   14
         Top             =   195
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issu&e"
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
      Left            =   9690
      Picture         =   "frmCO.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5295
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cmdBill 
      Height          =   1050
      Left            =   6105
      TabIndex        =   30
      Top             =   60
      Width           =   2085
      _ExtentX        =   3678
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
   Begin CoolButtonControl.CoolButton cmdDel 
      Height          =   1065
      Left            =   8730
      TabIndex        =   31
      Top             =   45
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1879
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
   Begin CoolButtonControl.CoolButton cmdSelectCustomer 
      Height          =   1080
      Left            =   135
      TabIndex        =   32
      Top             =   15
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1905
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
   Begin VB.Label txtPhone 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   765
      TabIndex        =   34
      Top             =   615
      Width           =   1575
   End
   Begin VB.Label txtCustName 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   765
      TabIndex        =   33
      Top             =   195
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   8400
      TabIndex        =   25
      Top             =   75
      Width           =   300
   End
   Begin VB.Label lblb 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   5775
      TabIndex        =   24
      Top             =   60
      Width           =   300
   End
   Begin VB.Label lblAddDel 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   870
      Left            =   8865
      TabIndex        =   23
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label lblAddBill 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   810
      Left            =   6180
      TabIndex        =   22
      Top             =   90
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   285
      Picture         =   "frmCO.frx":0BA8
      Stretch         =   -1  'True
      Top             =   645
      Width           =   360
   End
End
Attribute VB_Name = "frmCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCO As a_CO
Attribute oCO.VB_VarHelpID = -1
Dim WithEvents oCOLine As a_COL
Attribute oCOLine.VB_VarHelpID = -1
Dim oCustomer As a_Customer
Dim oProd As a_Product
Dim oCurrentCopy
Dim bValidCO As Boolean
Dim bValidCOLine As Boolean
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
Dim strCOErrMsg As String
Dim strCOLErrMsg As String

Public Sub Component(pCancel As Boolean, Optional pCO As a_CO, Optional pCustID As Long)
    On Error GoTo errHandler
Dim frm As frmHeader_CO

    If pCO Is Nothing Then
        Set oCO = New a_CO
        oCO.BeginEdit
        oCO.OrderType = enNormalCO
        oCO.SetStatus stInProcess
        lvwLines.Enabled = False
        If pCustID > 0 Then
            flgLoading = True
            LoadNewCustomer pCustID
            flgLoading = False
        End If
'''
        Set frm = New frmHeader_CO
        frm.Component oCO
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Unload Me
            pCancel = True
            Exit Sub
        End If
        Unload frm
'''
        
        ChangeState enAddingRow
    Else
        Set oCO = pCO
        oCO.BeginEdit
        oCO.GetStatus
        ChangeState enNotEditing
        
    End If
    
    If oCO.OrderType = enWant Then
        Me.Caption = "Wants for " & oCO.TPName
        Me.cmdNewRows.Enabled = False
    ElseIf oCO.OrderType = enNormalCO Then
        Me.Caption = "Order from " & oCO.Customer.Fullname & oCO.StaffNameB & IIf(oCO.OrderRef > "", "  (ref:" & oCO.OrderRef & ")", "")
    End If
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Component(pCO,pCustID)", Array(pCO, pCustID)
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmHeader_CO
Dim oSM As New z_StockManager
Dim oCOL As a_COL
    ofrm.Component oCO
    ofrm.Show vbModal
    
    txtMemo.Visible = (ofrm.Memo > "")
    txtMemo = "Note: " & ofrm.Memo
    
    oCO.setMemo ofrm.Memo
    For Each oCOL In oCO.COLines
        If oCOL.Ref = "" Then oCOL.SetRef ofrm.Ref
    Next
    LoadListView
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPP.mnuMemo"
End Sub


Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oCO.statusF = "IN PROCESS" And oCO.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuCancelLine.Enabled = (oCO.statusF = "ISSUED") 'And oCO.IsNew = False)
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.SetMenu"
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayoutLvw Me.lvwLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSelectCustomer_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oCO.Customer.ID > 0 Then
        frm.Component oCO.Customer
        frm.Show
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdSelectCustomer_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdBill_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oCO Is Nothing Then Exit Sub
    If oCO.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oCO.Customer.Addresses.Count Then
        i = 1
    End If
    If oCO.Customer.Addresses.Count >= i Then
        lblAddBill.Caption = oCO.Customer.Addresses(i).AddressMailing & vbCrLf & oCO.Customer.Addresses(i).EMail
        oCO.SetBillToAddress oCO.Customer.Addresses(i)
        iBillIdx = i
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdBill_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oCO Is Nothing Then Exit Sub
    If oCO.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oCO.Customer.Addresses.Count Then
        i = 1
    End If
    If oCO.Customer.Addresses.Count >= i Then
        lblAddDel.Caption = oCO.Customer.Addresses(i).AddressMailing & vbCrLf & oCO.Customer.Addresses(i).EMail
        oCO.setDelTOAddress oCO.Customer.Addresses(i)
        iBillIdx = i
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdDel_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oCO.SetCustomer(plngTPID) Then
        vCanAdd.RuleBroken "TP", False
        LoadCustomerDetailsToForm
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.LoadNewCustomer(plngTPID)", plngTPID
End Sub

Private Sub LoadCustomerDetailsToForm()
    With oCO.Customer
        txtMemo = "Note: " & oCO.Memo
        txtPhone = .Phone
        txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
        If Not (.BillTOAddress Is Nothing) Then
            If oCO.BillingCompany Is Nothing Then
                oCO.SetBillToAddress .BillTOAddress
            End If
            lblAddBill.Caption = .BillTOAddress.AddressShort
        End If
        If Not (.DelToAddress Is Nothing) Then
            If oCO.DelToAddress Is Nothing Then
                oCO.setDelTOAddress .DelToAddress
            End If
            lblAddDel.Caption = oCO.DelToAddress.AddressShort
        End If
    End With
End Sub


Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set vCanAdd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_Click()
    On Error GoTo errHandler
    If lvwLines.SelectedItem.Index > 0 Then
        Clipboard.Clear
        Clipboard.SetText left(lvwLines.SelectedItem.SubItems(9), ISBNLENGTH)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.lvwLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAddresses_Click()
    On Error GoTo errHandler
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.Component oCO
    frm.Show vbModal
    lblAddBill.Caption = oCO.BillTOAddress.AddressShort
    lblAddDel.Caption = oCO.DelToAddress.AddressShort

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuAddresses_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuChangeCustomer_Click()
'    On Error GoTo errHandler
'Dim lngTPID As Long
'Dim frm As frmBrowseCustomers2
'    Set frm = New frmBrowseCustomers2
'    frm.Show vbModal
'    lngTPID = frm.CustomerID
'    If oCO.SetCustomer(lngTPID) Then
'        With oCO.Customer
'            txtPhone = .Phone
'            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
'            lblAddBill.Caption = .BillTOAddress.AddressShort
'            lblAddDel.Caption = .BillTOAddress.AddressShort
'        End With
'        vCanAdd.RuleBroken "TP", False
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCO.mnuChangeCustomer_Click", , EA_NORERAISE
'    HandleError
'End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuDelLine"
End Sub



Private Sub oCO_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidCO = (pMsg = "")
    cmdIssue.Enabled = (bValidCO And oCO.COLines.Count > 0 And oCO.OrderType = enNormalCO And vMode = enNotEditing)
    cmdSave.Enabled = (bValidCO And oCO.COLines.Count > 0 And oCO.OrderType = enNormalCO And vMode = enNotEditing)
    strCOErrMsg = pMsg
    If vMode = enNotEditing Then
        txtError = strCOErrMsg
    Else
        txtError = strCOLErrMsg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCO_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oCOLine_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
 '   Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCOLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oCOLine_Valid(pMsg As String)
    On Error GoTo errHandler
    cmdEnter.Enabled = (pMsg = "")
    strCOLErrMsg = pMsg
    If vMode = enNotEditing Then
        txtError = strCOErrMsg
    Else
        txtError = strCOLErrMsg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCOLine_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub oCO_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
    On Error GoTo errHandler
    
    flgLoading = True
    
    txtRunningTotal = strtotal & IIf(lngTotalDeposit > 0, " less deposit of " & strTotalDeposit & " paid", "")
    lngCurrentTotal = lngTotal
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
   ' cmdNewRows.Enabled = (oCO.COLines.Count > 0)
    flgLoading = False
    
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCO_TotalChange(lngTotal,strtotal,lngTotalDeposit,strTotalDeposit,lngTotalVAT," & _
        "strTotalVAT)", Array(lngTotal, strtotal, lngTotalDeposit, strTotalDeposit, lngTotalVAT, strTotalVAT), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub oCO_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCO_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oCO_Dirty(pVal As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If pVal = True Then
        Me.cmdSave.Enabled = pVal And vMode = enNotEditing
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCO_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oCO_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.oCO_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub SetFocusFromCode()
    On Error GoTo errHandler
Dim strMsg As String
    
    If LenB(txtCode) > 0 Then
        If (oPC.Configuration.AntiquarianYN) Then
            mSetfocus txtPrice
        ElseIf txtOrdernum.Visible = False Then
            mSetfocus txtQty
        Else
            mSetfocus txtOrdernum
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetCursorFromCode"
End Sub

Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdSave.Enabled = True
    Me.cmdIssue.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim curTotalDeposit As Currency
    left = 10
    top = 10
    Width = 11100
    Height = 6700
    LoadControls
'    flgLoading = True
'    oCO.GetStatus
'    SetLvw
'    SetEditFrameEnabled False, enNotEditing
'    vMode = enNotEditing
'    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    LoadCustomerDetailsToForm
    LoadListView
End Sub
Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oCO.IsEditing Then oCO.CancelEdit
    UnsetMenu
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oCO = Nothing
    Set tlCustomer = Nothing
    Set oCOLine = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
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
    Me.txtDiscount.Enabled = pYesNo
    Me.txtPrice.Enabled = pYesNo
    txtOrdernum.Enabled = pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtTotal.Enabled = pYesNo
    Me.txtQty.Enabled = pYesNo
    
    Me.cmdEnter.Enabled = pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidCO And oCO.OrderType <> enWant
    Me.cmdSave.Enabled = (Not pYesNo) And bValidCO And oCO.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
    Me.txtDiscount.BackColor = lngColour
    txtOrdernum.BackColor = lngColour
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim strETACode As String
Dim strPos As String
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        If txtCode.Enabled Then mSetfocus txtCode
        Exit Sub
    End If
    If oPC.Configuration.COLAllocationStyle = "R" Then
        If oCOLine.Deposit = 0 Then
            If MsgBox("You have not indicated a deposit for this item. Do you wish to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    oCOLine.ApplyEdit
    oCOLine.BeginEdit

    If vMode = enAddingRow Then
        strETACode = oCOLine.ETACode
        lvwLines.ListItems.Add Index:=1, Key:=oCOLine.Key
        LoadListViewLine lvwLines.ListItems(oCOLine.Key), oCOLine
        lvwLines.Refresh
        ChangeState enAddingRow
        oCOLine.SetETA strETACode
        mSetfocus txtCode
    ElseIf vMode = enEditingRow Then
        LoadListViewLine lvwLines.ListItems(lngSelectedRowIndex), oCOLine
        ChangeState enNotEditing
    End If
    oCO.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdEnter_Click", , EA_NORERAISE, , "strpos", Array(strPos)
    HandleError
End Sub
Private Sub ChangeState(pToMode As EnumMode)
Dim lngColour As Long
    vMode = pToMode

    Select Case pToMode
    Case enEditingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtOrdernum.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQty.Enabled = True
        txtOrdernum.Enabled = True
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oCO.COLines.Count > 0)
        cmdCancel.Caption = "&Close"
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        fr1.ZOrder 1
        
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtOrdernum.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQty.Enabled = True
        txtError = ""
        flgLoading = True
        txtOrdernum = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oCO.COLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        
        Me.txtPhone.Caption = ""
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        ClearLineControls
        fr1.ZOrder 1
        
        mSetfocus txtCode
        
        Set oCOLine = oCO.COLines.Add
        oCOLine.TRID = oCO.TRID
        oCOLine.SetQty 1
        oCOLine.SetFulfilled "OS"

    Case enNotEditing
        flgLoading = True
        fr1.Visible = False
        txtError = ""
        txtOrdernum = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = (oCO.COLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        
        lvwLines.Enabled = True
        lvwLines.Height = 4000
        
        fr1.ZOrder 1
    
    
    End Select
End Sub

Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = enEditingRow Then
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        ChangeState enAddingRow
    End If

    ClearLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To lvwLines.ColumnHeaders.Count
        lvwLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwLines.ColumnHeaders(i).Width)
    Next
    If oCO.OrderType = enWant Then
        lvwLines.ColumnHeaders(2).Width = 4500
        lvwLines.ColumnHeaders(3).Width = 4200
        lvwLines.ColumnHeaders(4).Width = 0
        lvwLines.ColumnHeaders(5).Width = 0
        lvwLines.ColumnHeaders(6).Width = 0
        lvwLines.ColumnHeaders(7).Width = 0
        lvwLines.ColumnHeaders(8).Width = 0
        lvwLines.ColumnHeaders(9).Width = 0
        lvwLines.ColumnHeaders(3).Text = "Note"
    End If
    lvwLines.ListItems.Clear
    For i = 1 To oCO.COLines.Count
        If oCO.COLines(i).Fulfilled = "OS" Then
            Set lstItem = lvwLines.ListItems.Add
            lstItem.SubItems(9) = Format(oCO.COLines(i).Key, "@@@@@@@@@@")
            LoadListViewLine lstItem, oCO.COLines(i)
        End If
    Next i
EXIT_HANDLER:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oCOLine As a_COL)
    On Error GoTo errHandler
Dim currPrice As Currency

    With oCOLine
        lstItem.Text = .CodeF
        lstItem.Key = .Key
        lstItem.SubItems(1) = .TitleAuthorPublisher
        If oCO.OrderType = enNormalCO Then
            lstItem.SubItems(2) = .Qty
        Else
            lstItem.SubItems(2) = .Note
        End If
        
        lstItem.SubItems(3) = .Ref
        lstItem.SubItems(4) = .PriceF
        lstItem.SubItems(5) = .DiscountF
        If .Deposit <> 0 Then
            lstItem.SubItems(6) = .DepositF
        Else
            lstItem.SubItems(6) = " "
        End If
        lstItem.SubItems(7) = .ETAF
        lstItem.SubItems(8) = .ExtensionF
        lstItem.SubItems(9) = .Ean
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.LoadListViewLine(lstItem,oCOLine)", Array(lstItem, oCOLine)
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
Dim tmpCOLine As a_COL
Dim lngCOLID As Long

'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    If lvwLines.SelectedItem.Index < 1 Then Exit Sub
    
    lngEditingIdx = lvwLines.SelectedItem.Key
    
    If oCO.Status <> stInProcess Then
        'Store Current Row's ID
        lngCOLID = oCO.COLines(lvwLines.SelectedItem.Key).COLineID
        'point to exisiting line
        Set tmpCOLine = oCO.COLines(lvwLines.SelectedItem.Key)
        'Create new oCOLINE and copy values from old
        Set oCOLine = Nothing
        Set oCOLine = oCO.COLines.Add
        oCOLine.SetReplacementForLineID lngCOLID
        oCOLine.TRID = tmpCOLine.TRID
        oCOLine.SetQty tmpCOLine.Qty
 '       oCOLine.SetQtySS tmpCOLine.QtySS
        oCOLine.SetETA tmpCOLine.ETA
        oCOLine.SetDiscount tmpCOLine.Discount
        oCOLine.Ean = tmpCOLine.Ean
        oCOLine.code = tmpCOLine.code
        oCOLine.CodeF = tmpCOLine.CodeF
      '  oCOLine.ForeignPrice = tmpCOLine.ForeignPrice
        oCOLine.SetPrice tmpCOLine.Price
 '       oCOLine.lastaction = tmpCOLine.lastaction
        oCOLine.MainAuthor = tmpCOLine.MainAuthor
        oCOLine.Note = tmpCOLine.Note
        oCOLine.pID = tmpCOLine.pID
       ' oCOLine.Product = tmpCOLine.Product
        oCOLine.VATRate = tmpCOLine.VATRate
      '  oCOLine.SetSection tmpCOLine.Section
        oCOLine.SetRef tmpCOLine.Ref
        oCOLine.SetFulfilled "OS"
        oCOLine.Title = tmpCOLine.Title
        Set tmpCOLine = Nothing
    Else
        Set oCOLine = Nothing
        Set oCOLine = oCO.COLines(lvwLines.SelectedItem.Key)
    End If
    
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    
    ChangeState enEditingRow
    
    txtCode = CStr(oCOLine.Ean)
    txtTitle = oCOLine.Title
    txtPrice = CStr(oCOLine.Price)
    txtDiscount = CStr(oCOLine.Discount)
    txtdeposit = oCOLine.Deposit
    txtOrdernum = oCOLine.Ref
    txtQty = oCOLine.Qty
    txtNote = oCOLine.Note
    txtETA = oCOLine.ETAF
    If oCOLine.Qty > 1 Then
        mSetfocus txtQty
    Else
        mSetfocus txtPrice
    End If
    oCOLine.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If oCO.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'-------End Compsny code
Private Sub txtOrdernum_Validate(Cancel As Boolean)
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCOLine.SetRef txtOrdernum
    If Err Then
      Beep
      intPos = txtOrdernum.SelStart
      txtOrdernum = oCOLine.Ref
      txtOrdernum.SelStart = intPos - 1
    End If

End Sub

Private Sub txtNote_Change()
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oCOLine.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCOLine.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCOLine.setnote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oCOLine.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub




Private Sub mnuFileOK_Click()
    On Error GoTo errHandler
'    cmdOK_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuFileOK_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
Dim strResult As String
    oCO.SetStatus stVOID
    oCO.ApplyEdit strResult
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuVoid"
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean
Dim oProdCode As New z_ProdCode

START:
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_HANDLER
    End If
    bOK = oCOLine.SetLineProduct("", txtCode)
    If bOK Then
        txtTitle = oCOLine.Title
        txtPrice = oCOLine.Price
        txtQty = oCOLine.QtyF
        txtdeposit = oCOLine.Deposit
        txtDiscount = oCOLine.Discount
        mSetfocus txtPrice
        txtCode = oCOLine.Ean
        txtETA = oCOLine.ETAF
        If oCO.OrderRef > "" Then
            txtOrdernum = oCO.OrderRef
        End If
        AutoSelect txtPrice
    Else
        Dim frmAdHoc As frmAdHocProduct
        Set frmAdHoc = New frmAdHocProduct
        frmAdHoc.Component txtCode
        frmAdHoc.Show vbModal
        txtCode = frmAdHoc.code
        Unload frmAdHoc
        Set frmAdHoc = Nothing
        Cancel = True
        GoTo START
    End If

EXIT_HANDLER:
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub RemoveDetailLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oCO.COLines.Remove lvwLines.ListItems(i).Key
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
    ErrorIn "frmCO.RemoveDetailLine"
End Sub

Private Sub LoadCustomer()
    On Error GoTo errHandler
Dim strAddress As String
    With oCO
    '    txtStatus = .statusF
        SetIssueButtonCaption
        Me.txtCustName = .TPName
  '      txtAccnum = .TPAccNum
        txtPhone = .TPPhone
       ' txtPhone = .TPPhone
            If oCO.BillToAddressID > 0 Then
                strAddress = oCO.BillTOAddress.AddressMailing
            End If
            Me.lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
            If oCO.GoodsAddressID > 0 Then
                strAddress = oCO.DelToAddress.AddressMailing
            End If
            Me.lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")

    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.LoadCustomer"
End Sub


Private Sub SaveCO()
    On Error GoTo errHandler
    
    oCO.post
    
EXIT_HANDLER:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
  '  Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.SaveCO"
End Sub

Public Sub PrintOrder()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCOLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oCO.Load oCO.TRID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoCOLines Then
        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
        GoTo EXIT_HANDLER
    End If
    
EXIT_HANDLER:
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
    ErrorIn "frmCO.PrintOrder"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCOLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmCOPreview

    If oPC.Configuration.Signtransactions = True Then
        If SecurityControl(enSECURITY_CO_SIGN, , "Sign this order", DOCAPPROVAL, , , gSTAFFID) = False Then
               Exit Sub
        End If
    Else
        If oCO.Status = stInProcess Then
            If MsgBox("Issue this order?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    
    WaitMsg "Issuing customer order  . . .", True, Me
    oCO.SetStatus stISSUED
    oCO.StaffID = gSTAFFID
    strResult = oCO.post
    Set frm = New frmCOPreview
    frm.ComponentObject oCO
    frm.Show
    WaitMsg "", False, Me
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
    If oCO.Status <> stISSUED And oCO.Status <> stCOMPLETE Then
        oCO.SetStatus stInProcess
    End If
    SaveCO
    oCO.BeginEdit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmCOPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel this order?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    oCO.CancelEdit
    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmCOPreview
        frm.Component oCO.TRID
        frm.Show
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = ""
   ' Me.txtDiscount = ""
    Me.txtPrice = ""
    Me.txtTitle = ""
   ' Me.txtTotal = ""
    Me.txtNote = ""
    Me.txtdeposit = ""
    Me.txtQty = ""
    Me.txtOrdernum = ""
    Me.txtETA = ""
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.ClearLineControls"
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtETA_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If FNS(txtETA) = "" Then
        Cancel = True
        Exit Sub
    End If
    If Not oCOLine.SetETA(txtETA) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtETA_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtETA_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtETA")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtETA_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtETA_LostFocus()
    On Error GoTo errHandler
    txtETA = oCOLine.ETAF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtETA_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtPrice")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCOLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_LostFocus()
    On Error GoTo errHandler
  '  txtPrice = oCOLine.PriceF
    txtTotal = oCOLine.ExtensionInclDepositF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtPrice_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCOLine.SetQty(txtQty) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_LostFocus()
    On Error GoTo errHandler
    txtTotal = oCOLine.ExtensionInclDepositF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQty_LostFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCOLine.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_LostFocus()
    On Error GoTo errHandler
  '  txtDiscount = oCOLine.DiscountPercentF
    txtTotal = oCOLine.ExtensionInclDepositF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtDiscount_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDiscount")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDeposit_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCOLine.SetDeposit(txtdeposit) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtDeposit_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDeposit_LostFocus()
    On Error GoTo errHandler
 '   txtdeposit = oCOLine.DepositF
    txtTotal = oCOLine.ExtensionInclDepositF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtDeposit_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDeposit_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtDeposit")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtDeposit_GotFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oCO.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oCO.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.SetIssueButtonCaption"
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
    ErrorIn "frmCO.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
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
    ErrorIn "frmCO.SetLvw"
End Sub

