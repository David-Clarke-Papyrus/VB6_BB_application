VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmCO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Order"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmCOwSS2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5325
      Visible         =   0   'False
      Width           =   2745
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2265
      Left            =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1200
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
         Text            =   "Ref."
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
      Picture         =   "frmCOwSS2.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
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
      TabIndex        =   24
      TabStop         =   0   'False
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
      TabStop         =   0   'False
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
      Left            =   7440
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmCOwSS2.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3480
      Width           =   4140
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1605
      Left            =   120
      TabIndex        =   16
      Top             =   3660
      Width           =   10650
      Begin VB.TextBox lblSupplierDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5355
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   705
         Width           =   4170
      End
      Begin VB.CommandButton cmdEditProduct 
         Height          =   270
         Left            =   4890
         Picture         =   "frmCOwSS2.frx":085E
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1260
         Width           =   315
      End
      Begin VB.CommandButton cmdFind 
         Height          =   345
         Left            =   90
         Picture         =   "frmCOwSS2.frx":0BE8
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   315
         Width           =   375
      End
      Begin VB.CommandButton cmdExtraCharge 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Add.chrg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2025
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   900
         Width           =   645
      End
      Begin VB.TextBox txtExtraCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1005
         TabIndex        =   10
         Top             =   1230
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox txtExtraCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   60
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1230
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtQtySS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2955
         TabIndex        =   3
         Top             =   375
         Width           =   765
      End
      Begin VB.TextBox txtOrdernum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3735
         TabIndex        =   4
         Top             =   375
         Width           =   1365
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   9045
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   375
         Width           =   1515
      End
      Begin VB.TextBox txtETA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7725
         TabIndex        =   8
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   375
         Width           =   1305
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2175
         TabIndex        =   2
         Top             =   375
         Width           =   765
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4860
         TabIndex        =   11
         Top             =   930
         Width           =   4575
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6180
         TabIndex        =   6
         Top             =   375
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
         Left            =   9525
         MaskColor       =   &H00C4BCA4&
         Picture         =   "frmCOwSS2.frx":0F72
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   900
         Width           =   1000
      End
      Begin VB.TextBox txtdeposit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6945
         TabIndex        =   7
         Top             =   375
         Width           =   750
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1230
         Width           =   4170
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5130
         TabIndex        =   5
         Top             =   375
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   510
         TabIndex        =   1
         Top             =   375
         Width           =   1650
      End
      Begin VB.Label lblFCTerms 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5535
         TabIndex        =   47
         Top             =   660
         Width           =   2265
      End
      Begin VB.Label lblExtraCharge 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   2070
         TabIndex        =   45
         Top             =   1230
         Width           =   2085
      End
      Begin VB.Label lblExtra2 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Ex. chrge."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1020
         TabIndex        =   44
         Top             =   1005
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblExtra1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   195
         TabIndex        =   43
         Top             =   1005
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblQtySS 
         BackColor       =   &H00D3D3CB&
         Caption         =   "QtySS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   3060
         TabIndex        =   42
         Top             =   150
         Width           =   570
      End
      Begin VB.Label lblTotal 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   9495
         TabIndex        =   41
         Top             =   150
         Width           =   645
      End
      Begin VB.Label lblETA 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "ETA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   8040
         TabIndex        =   33
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   150
         Width           =   645
      End
      Begin VB.Label lblNote 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   4260
         TabIndex        =   32
         Top             =   975
         Width           =   585
      End
      Begin VB.Label lblQty 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty firm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2205
         TabIndex        =   31
         Top             =   150
         Width           =   675
      End
      Begin VB.Label lblOrdernum 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Ref."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4230
         TabIndex        =   30
         Top             =   150
         Width           =   675
      End
      Begin VB.Label lblDiscount 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Disc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   6240
         TabIndex        =   21
         Top             =   150
         Width           =   585
      End
      Begin VB.Label lblDeposit 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Deposit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   6930
         TabIndex        =   20
         Top             =   150
         Width           =   810
      End
      Begin VB.Label lblCode 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1110
         TabIndex        =   19
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblPrice 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   5370
         TabIndex        =   18
         Top             =   150
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
      Picture         =   "frmCOwSS2.frx":12FC
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5295
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cmdBill 
      Height          =   1050
      Left            =   6105
      TabIndex        =   34
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
      TabIndex        =   35
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
      TabIndex        =   36
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
      TabIndex        =   38
      Top             =   615
      Width           =   1575
   End
   Begin VB.Label txtCustName 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   765
      TabIndex        =   37
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   90
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   285
      Picture         =   "frmCOwSS2.frx":1686
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
Dim bShowExtracharges As Boolean

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
Dim bMemoExpanded As Boolean
Public Sub component(pCancel As Boolean, Optional pCO As a_CO, Optional pCustID As Long)
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
        frm.component oCO
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Unload Me
            pCancel = True
            Exit Sub
        End If
        Unload frm
        ChangeState enAddingRow
    Else
        Set oCO = pCO
        oCO.BeginEdit
        ChangeState enNotEditing
        SetIssueButtonCaption
    End If
    oCO.GetStatus
    If oCO.OrderType = enWant Then
        Me.Caption = "Wants for " & oCO.TPNAME
        Me.cmdNewRows.Enabled = False
    ElseIf oCO.OrderType = enNormalCO Then
      '  Me.Caption = "Order from " & oCO.Customer.FullName & oCO.StaffNameB & IIf(oCO.OrderRef > "", "  (ref:" & oCO.OrderRef & ")", "")
        Me.Caption = "Sales order (Edit) " & "  " & oCO.DOCCode & "    " & oCO.DOCDate & " " & oCO.StaffNameB & IIf(oCO.OrderRef > "", "  (ref:" & oCO.OrderRef & ")", "") & "   " & oCO.DOCCode
    End If
    SetMenu
    If oPC.AllowsSSInvoicing Then
        lblqty = "Qty firm"
        lblQtySS.Visible = True
        txtQtySS.Visible = True
        txtQty.Width = 885
    Else
        lblqty = "Qty"
        lblQtySS.Visible = False
        txtQtySS.Visible = False
        txtQty.Width = 1500
    End If
    bShowExtracharges = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.component(pCancel,pCO,pCustID)", Array(pCancel, pCO, pCustID)
End Sub

Private Sub cmdEditProduct_Click()
Dim f As New frmSupplierBookDetails
Dim X As Long
Dim Y As Long
Dim sSYS As String
Dim sSupplierNote As String

    If oCOLine Is Nothing Then Exit Sub
    If FNS(oCOLine.PID) = "" Then Exit Sub
    Set oProd = New a_Product
    oProd.Load oCOLine.PID, 0
    If oProd.SupplierCurrencyID > 0 Then
        sSYS = oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).SYSNAME
    Else
        sSYS = oPC.Configuration.DefaultCurrency.SYSNAME
    End If
    
    f.component sSYS, oProd.SupplierID, oProd.LastSupplierName, oProd, Me.Left, Me.top + 1500
    f.Show vbModal
'    If Not f.Cancelled Then
'        oProd.BeginEdit
'        oProd.DealID = f.DealID
'        oProd.ApplyEdit
'    End If
    Select Case sSYS
    Case "EUR"
        sSupplierNote = oProd.EUPriceF
    Case "USD"
        sSupplierNote = oProd.USPriceF
    Case "GBP"
        sSupplierNote = oProd.UKPriceF
    Case Else
        sSupplierNote = oProd.RRPF
    End Select
    lblSupplierDetails.Visible = True
    lblSupplierDetails.text = sSupplierNote & " / " & Format(f.DiscountRate, "##.00")
    mSetfocus Me.txtPrice
    Set oProd = Nothing
End Sub


Private Sub cmdExtraCharge_Click()
    On Error GoTo errHandler
    
    bShowExtracharges = Not bShowExtracharges
    ControlCaptureFrame bShowExtracharges
    If bShowExtracharges Then mSetfocus txtExtraCode
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdExtraCharge_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ControlCaptureFrame(bExtra As Boolean)
    On Error GoTo errHandler

   ' Me.cmdExtraCharge.Visible = False
    Me.txtExtraCharge.Visible = bExtra
    Me.txtExtraCharge.Enabled = bExtra
    Me.txtExtraCode.Visible = bExtra
    Me.lblExtraCharge.Visible = bExtra
    Me.lblExtra1.Visible = bExtra
    Me.lblExtra2.Visible = bExtra
    
    txtCode.Enabled = Not bExtra
    lblCode.Enabled = Not bExtra
    txtQty.Enabled = Not bExtra
    lblqty.Enabled = Not bExtra
    txtQtySS.Enabled = Not bExtra
    lblQtySS.Enabled = Not bExtra
    txtOrdernum.Enabled = Not bExtra
    lblOrdernum.Enabled = Not bExtra
    txtPrice.Enabled = Not bExtra
    lblPrice.Enabled = Not bExtra
    txtDiscount.Enabled = Not bExtra
    lblDiscount.Enabled = Not bExtra
    txtdeposit.Enabled = Not bExtra
    lblDeposit.Enabled = Not bExtra
    txtETA.Enabled = Not bExtra
    lblETA.Enabled = Not bExtra
    txtTotal.Enabled = Not bExtra
    lblTotal.Enabled = Not bExtra
    txtNote.Enabled = Not bExtra
    lblNote.Enabled = Not bExtra
  '  txtExtraCode = ""
  '  txtExtraCharge = ""
    If bExtra Then
        If oCOLine.ExtraCode > "" Then
            txtExtraCode = oCOLine.ExtraCode
            txtExtraCharge = oCOLine.ExtraCharge
            Me.lblExtraCharge.Caption = oCOLine.ExtraChargeDescription
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.ControlCaptureFrame(bExtra)", bExtra
End Sub

Private Sub cmdFind_Click()
    On Error GoTo errHandler
Dim frm As New frmQuickProductFind
Dim strCode As String
    strCode = txtCode
    
    frm.Show vbModal
    If frm.QtyQuickFound = 0 Then
        MsgBox "Nothing found", vbInformation, "Status"
    End If
    If frm.Cancelled = False Then
        If frm.EAN > "" Then txtCode = frm.EAN
    End If
    txtCode.SetFocus
    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdFind_Click", , EA_NORERAISE
    HandleError
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
    ofrm.component oCO
    ofrm.Show vbModal
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = ofrm.Memo
    
    oCO.SetMemo ofrm.Memo
    For Each oCOL In oCO.COLines
        If oCOL.Ref = "" Then oCOL.SetRef ofrm.Ref
    Next
    LoadListView
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuMemo"
End Sub


Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oCO.StatusF = "IN PROCESS" And oCO.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuCancelLine.Enabled = (oCO.StatusF = "ISSUED") 'And oCO.IsNew = False)
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = False
    Forms(0).mnuPastelines.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.SetMenu"
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvwLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuSaveLayout"
End Sub
Private Sub cmdSelectCustomer_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oCO.Customer.ID > 0 Then
        frm.component oCO.Customer
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
        oCO.setDelToAddress oCO.Customer.Addresses(i)
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
    On Error GoTo errHandler
    With oCO.Customer
        txtTPMemo = oCO.Memo
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
                oCO.setDelToAddress .DelToAddress
            End If
            lblAddDel.Caption = oCO.DelToAddress.AddressShort
        End If
    End With
    oCO.SetDirty False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.LoadCustomerDetailsToForm"
End Sub


Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set vCanAdd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Label4_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Label4_Click", , EA_NORERAISE
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
    If lvwLines Is Nothing Then Exit Sub
    If lvwLines.SelectedItem Is Nothing Then Exit Sub
    
    If lvwLines.SelectedItem.Index > 0 Then
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText Left(lvwLines.SelectedItem.SubItems(9), ISBNLENGTH)
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
    frm.component oCO
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
    cmdIssue.Enabled = (bValidCO And oCO.COLines.Count > 0 And oCO.OrderType = enNormalCO And vMode = enNotEditing And oCO.Status <> stISSUED And oCO.Status <> stCOMPLETE)
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
    'If flgLoading Then Exit Sub
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
    ErrorIn "frmCO.SetFocusFromCode"
End Sub

Private Sub txtCode_GotFocus()
    Me.cmdEditProduct.Enabled = False
End Sub

Private Sub txtCode_LostFocus()
 cmdEditProduct.Enabled = True
End Sub

Private Sub txtExtraCode_LostFocus()
    On Error GoTo errHandler
                mSetfocus txtExtraCharge

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtExtraCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtExtraCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim bOK  As Boolean
'Dim oProdCode As New z_ProdCode

 '   If txtExtraCode = "" Then Exit Sub
    
    If Not (IsISBN13(txtExtraCode) Or IsISBN10(txtExtraCode) Or IsHashCode(txtExtraCode) Or IsPrivateCode(txtExtraCode) Or txtExtraCode = "") Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        Exit Sub
    End If
    If txtExtraCode > "" Then
        bOK = oCOLine.SetLineExtraProduct(txtExtraCode)
        If bOK Then
                Me.lblExtraCharge = oCOLine.ExtraChargeDescription
                txtPrice = oCOLine.Price
                txtExtraCharge.Enabled = True
                AutoSelect txtExtraCharge
                mSetfocus txtExtraCharge
        Else
            Cancel = True
            Me.txtExtraCharge.Enabled = False
        End If
    Else
        oCOLine.SetExtraCharge 0
        oCOLine.ExtraPID = ""
        txtExtraCharge.Enabled = False
        
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtExtraCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_DblClick()
    On Error GoTo errHandler
    If txtNote.Height > 285 Then
        txtNote.Height = 285
    Else
        txtNote.Height = txtNote.Height * 4
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtNote_DblClick", , EA_NORERAISE
    HandleError
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
    If Me.WindowState <> 2 Then
        Left = 10
        top = 10
        Width = 11100
        Height = 6700
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    cmdEditProduct.Visible = oPC.GetProperty("AllowSupplierDetailsCaptureInCustomerOrder") = "TRUE"
    lblSupplierDetails.Visible = cmdEditProduct.Visible
    LoadCustomerDetailsToForm
    cmdEditProduct.Visible = oPC.IncludeSupplierFeatures
    
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.LoadControls"
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
Dim oSQL As New z_SQL
Dim rsDup As ADODB.Recordset
Dim frmDups As frmPossibleDuplicateCOLS

    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        If txtCode.Enabled Then mSetfocus txtCode
        Exit Sub
    End If
    If Not oCOLine.oProd Is Nothing Then
        If oCOLine.Qty <= oCOLine.oProd.QtyOnHand Then
            If MsgBox("There is quantity on hand of :" & CStr(oCOLine.oProd.QtyOnHand) & vbCrLf & "Do you want to continue to post?", vbInformation + vbOKCancel, "Warning") = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    If Not oPC.IncludeSupplierFeatures Then
        If oCOLine.Deposit = 0 Then
            If MsgBox("You have not indicated a deposit for this item. Do you wish to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    If oCOLine.Ref > "" Then
        Set rsDup = oSQL.GetPossibleDuplicateCOLByRef(oCO.Customer.ID, oCOLine.Ref)
        If rsDup.RecordCount > 0 Then
            Set frmDups = New frmPossibleDuplicateCOLS
            frmDups.component rsDup
            frmDups.Show vbModal
            If frmDups.OKToContinue = False Then
                Unload frmDups
                Set frmDups = Nothing
                Exit Sub
            End If
            Unload frmDups
            Set frmDups = Nothing
        End If
    End If
    
    oCOLine.ApplyEdit
    oCOLine.BeginEdit

    If vMode = enAddingRow Then
        strETACode = oCOLine.ETACode
        If lvwLines.ListItems.Count < val(oCOLine.Key) Then
            lvwLines.ListItems.Add Index:=1, Key:=oCOLine.Key
            LoadListViewLine lvwLines.ListItems(oCOLine.Key), oCOLine
        End If
        lvwLines.Refresh
        ChangeState enAddingRow
        oCOLine.SetETA strETACode
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwLines.ListItems(lngSelectedRowIndex), oCOLine
        ChangeState enNotEditing
    End If
    oCO.GetStatus
    ControlCaptureFrame False
    bShowExtracharges = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.cmdEnter_Click", , EA_NORERAISE
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
    If Not oCO.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
    
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.ChangeState(pToMode)", pToMode
End Sub

Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        If txtCode > "" Then  'THis is not after a post but is an aborted  add row action
            oCO.COLines.DecrementMaxKeyUsed
        End If
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        ChangeState enAddingRow
    End If

    ClearLineControls
    ControlCaptureFrame False
    bShowExtracharges = False
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
        lvwLines.ColumnHeaders(3).text = "Note"
    End If
    lvwLines.ListItems.Clear
    For i = 1 To oCO.COLines.Count
        If oCO.COLines(i).Fulfilled = "OS" Then
            Set lstItem = lvwLines.ListItems.Add
            lstItem.SubItems(9) = Format(oCO.COLines(i).Key, "@@@@@@@@@@")
            LoadListViewLine lstItem, oCO.COLines(i)
        End If
    Next i
EXIT_Handler:
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
        lstItem.text = .CodeForEditing
        lstItem.Key = .Key
        lstItem.SubItems(1) = .TitleAuthorPublisher
        If oCO.OrderType = enNormalCO Then
            If oPC.AllowsSSInvoicing = True Then
                lstItem.SubItems(2) = .QtyFirm & "/" & .QtySS
            Else
                lstItem.SubItems(2) = .Qty
            End If
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
        lstItem.SubItems(9) = .EAN
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
Dim dblDiscountRate As Double
Dim sDescription As String
Dim bCreated As Boolean
Dim oProd As a_Product
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
        If oPC.AllowsSSInvoicing = True Then
            oCOLine.SetQtyFirm tmpCOLine.QtyFirm
            oCOLine.SetQtySS tmpCOLine.QtySS
        Else
            oCOLine.SetQty tmpCOLine.Qty
        End If
 '       oCOLine.SetQtySS tmpCOLine.QtySS
        oCOLine.SetETA tmpCOLine.ETA
        oCOLine.SetDiscount tmpCOLine.Discount
        oCOLine.EAN = tmpCOLine.EAN
        oCOLine.code = tmpCOLine.code
        oCOLine.CodeF = tmpCOLine.CodeF
        oCOLine.ExtraChargeDescription = tmpCOLine.ExtraChargeDescription
        oCOLine.ExtraCode = tmpCOLine.ExtraCode
        oCOLine.SetActionTaken tmpCOLine.ActionTaken
        oCOLine.ExtraPID = tmpCOLine.ExtraPID
        oCOLine.SetExtraCharge tmpCOLine.ExtraCharge
        oCOLine.ExtraCode = tmpCOLine.ExtraCode
      '  oCOLine.ForeignPrice = tmpCOLine.ForeignPrice
        oCOLine.SetPrice tmpCOLine.Price
 '       oCOLine.lastaction = tmpCOLine.lastaction
        oCOLine.MainAuthor = tmpCOLine.MainAuthor
        oCOLine.Note = tmpCOLine.Note
        oCOLine.PID = tmpCOLine.PID
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
        If oCOLine.PID > "" Then oCOLine.LoadProduct oCOLine.PID
    End If
    
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    
    ChangeState eneditingrow
    
    txtCode = CStr(oCOLine.EAN)
    txtTitle = oCOLine.Title
    txtPrice = CStr(oCOLine.Price)
    txtDiscount = CStr(oCOLine.Discount)
    txtdeposit = oCOLine.Deposit
    txtOrdernum = oCOLine.Ref
    txtExtraCharge = oCOLine.ExtraCharge
    txtExtraCode = oCOLine.ExtraCode
'    lblExtraCharge.Caption = oCOLine.ExtraChargeDescription
'    If txtExtraCode > "" Then
'        txtExtraCode.Visible = True
'        txtExtraCharge.Visible = True
'        lblExtra1.Visible = True
'        lblExtra2.Visible = True
'        lblExtraCharge.Visible = True
'       ' Me.cmdExtraCharge.Visible = False
'    Else
'        txtExtraCharge.Visible = False
'        txtExtraCode.Visible = False
'        lblExtra1.Visible = False
'        lblExtra2.Visible = False
'        lblExtraCharge.Visible = False
'       ' Me.cmdExtraCharge.Visible = True
'    End If
    If oPC.GetProperty("AllowSupplierDetailsCaptureInCustomerOrder") = "TRUE" Then
          Set oProd = New a_Product
          bCreated = True
          oProd.Load oCOLine.PID, 0
          oProd.DealDiscount dblDiscountRate, sDescription
          Set oProd = Nothing
          lblSupplierDetails = oCOLine.ForeignPriceF & " / " & Format(dblDiscountRate, "##.00")
    End If
    
    If oCOLine.FCID <> oPC.Configuration.DefaultCurrencyID And oCOLine.FCID > 0 Then
        Me.lblFCTerms = oCOLine.ForeignPriceF & "/" & oCOLine.FCFactorInvF
    End If
    If oPC.AllowsSSInvoicing Then
        txtQty = oCOLine.QtyFirm
        txtQtySS = oCOLine.QtySS
    Else
        txtQty = oCOLine.Qty
    End If
    txtNote = oCOLine.Note
    txtETA = oCOLine.ETAF
    If oCOLine.Qty > 1 Then
        mSetfocus txtQty
    Else
        mSetfocus txtPrice
    End If
    oCOLine.GetStatus
    If oCOLine.ExtraCode > "" Then
        ControlCaptureFrame True
    End If
    bShowExtracharges = True
    
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
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oCOLine.SetRef txtOrdernum
    If Err Then
      Beep
      intPos = txtOrdernum.SelStart
      txtOrdernum = oCOLine.Ref
      txtOrdernum.SelStart = intPos - 1
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtOrdernum_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    
    txtNote = HandleTextWithBites(txtNote)
    
    On Error Resume Next
    oCOLine.SetNote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCOLine.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oCOLine.SetNote(txtNote)
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
    If txtCode = "" Or vMode = eneditingrow Then Exit Sub
    
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_Handler
    End If
    bOK = oCOLine.SetLineProduct("", txtCode)
    If bOK Then
'        If oCOLine.COLType <> "A" Then
            txtTitle = oCOLine.Title
            txtPrice = oCOLine.Price
            If oPC.AllowsSSInvoicing Then
                txtQty = oCOLine.QtyFirmF
                txtQtySS = oCOLine.QtySSF
            Else
                txtQty = oCOLine.QtyF
            End If
            txtdeposit = oCOLine.Deposit
            txtDiscount = oCOLine.Discount
            mSetfocus txtPrice
            txtCode = oCOLine.CodeForEditing
            txtETA = oCOLine.ETAF
            If oCO.OrderRef > "" Then
                txtOrdernum = oCO.OrderRef
            End If
'            If oPOLine.Product.QtyCopiesOnHand > 0 Or oPOLine.Product.QtyOnOrder > 0 Then
'                txtOHOO = "OH: " & oPOLine.Product.QtyOnHandF & " / " & "OO: " & oPOLine.Product.QtyOnOrderF
'            Else
'                txtOHOO = ""
'            End If
            AutoSelect txtPrice
    Else
        If CheckThisPoint(M_NEWPRODUCTINADHOCFORM) Then
            If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then
                Cancel = True
                Exit Sub
            End If
        End If
        If GetAdhocDetails() Then
            GoTo START
        Else
            MsgBox "Cannot find item", vbOKOnly + vbInformation, "Finding stock item"
            Cancel = True
            Exit Sub
        End If
    End If

EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Function GetAdhocDetails() As Boolean
        If CheckThisPoint(M_NEWPRODUCTINADHOCFORM) Then
            If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then
                GetAdhocDetails = False
                Exit Function
            End If
        End If
    
        Dim frmAdHoc As frmAdHocProduct
        Set frmAdHoc = New frmAdHocProduct
        frmAdHoc.component txtCode
        frmAdHoc.Show vbModal
        txtCode = frmAdHoc.code
        GetAdhocDetails = Not frmAdHoc.IsCancelled
        Unload frmAdHoc
        Set frmAdHoc = Nothing

End Function

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
        Me.txtCustName = .TPNAME
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
    
    oCO.Post
    
EXIT_Handler:
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
        GoTo EXIT_Handler
    End If
    
EXIT_Handler:
    Me.MousePointer = vbDefault
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
Dim strResult As String
Dim frm As frmCOPreview

    If oPC.Configuration.SignTransactions = True Then
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
    strResult = oCO.Post
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
Dim lngTRID As Long

    On Error GoTo errHandler
    If oCO.Status <> stISSUED And oCO.Status <> stCOMPLETE Then
        oCO.SetStatus stInProcess
    End If
    SaveCO
    lngTRID = oCO.TRID
    Set oCO = Nothing
    Set oCO = New a_CO
    oCO.Load lngTRID, False
    LoadListView
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
    If cmdCancel.Caption = "&Close" And oCO.TRID > 0 Then
        Set frm = New frmCOPreview
        frm.component oCO.TRID, False
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
    Me.txtPrice = ""
    Me.txtTitle = ""
    Me.txtNote = ""
    Me.txtdeposit = ""
    Me.txtQty = ""
    Me.txtQtySS = ""
    Me.txtOrdernum = ""
    Me.txtETA = ""
    Me.txtExtraCode = ""
    Me.txtExtraCharge = ""
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




Private Sub txtExtraCharge_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtExtraCharge")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtExtraCharge_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtExtraCharge_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCOLine.SetExtraCharge(txtExtraCharge) Then
       ' Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtExtraCharge_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtExtraCharge_LostFocus()
    On Error GoTo errHandler
  '  txtExtraCharge = oCOLine.PriceF
    txtTotal = oCOLine.ExtensionInclDepositF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtExtraCharge_LostFocus", , EA_NORERAISE
    HandleError
End Sub













Private Sub txtPrice_DblClick()
    On Error GoTo errHandler
Dim f As New frmFCPrice
    
Dim X As Long
Dim Y As Long
    
    If Not oPC.SupportsUNISA Then Exit Sub
    
    If IIf(oCOLine.FCID = 0, oPC.Configuration.DefaultCurrencyID, oCOLine.FCID) = oPC.Configuration.DefaultCurrencyID Then
        f.component Me.Left + 2000, Me.top + 2000, oCOLine.Price, oCOLine.Price, oCOLine.FCID, oCOLine.VATRate, oCOLine.FCFactor
    Else
        f.component Me.Left + 2000, Me.top + 2000, oCOLine.ForeignPrice, oCOLine.Price, oCOLine.FCID, oCOLine.VATRate, oCOLine.FCFactor
    End If

    f.Show vbModal
    If f.UserCancelled Then
        Unload f
        Exit Sub
    End If
    oCOLine.SetFCFactor Round(f.FCFactor, 6)
    oCOLine.SetForeignPrice CStr(f.ForeignPrice)
    oCOLine.FCID = f.FCID
    oCOLine.Price = f.LocalPriceIncVAT
    Me.txtPrice = oCOLine.Price
    lblFCTerms.Caption = oCOLine.ForeignPriceF & "/" & f.FCFactorINV
    Unload f

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtPrice_DblClick", , EA_NORERAISE
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
Private Sub txtQty_DblClick()
    On Error GoTo errHandler
    If Not oPC.SupportsUNISA Then Exit Sub
    txtQty = oCO.TotalQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQty_DblClick", , EA_NORERAISE
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
    If oPC.AllowsSSInvoicing = True Then
        If Not oCOLine.SetQtyFirm(txtQty) Then
            Cancel = True
        End If
    Else
        If Not oCOLine.SetQty(txtQty) Then
            Cancel = True
        End If
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
'========
Private Sub txtQtySS_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtySS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQtySS_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtySs_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oCOLine.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtySS_LostFocus()
    On Error GoTo errHandler
    txtTotal = oCOLine.ExtensionInclDepositF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtQtySS_LostFocus", , EA_NORERAISE
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
        If oCO.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oCO.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Enabled = False
            'cmdIssue.Caption = "Print"
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
Dim Style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwLines.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   Style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   Style = Style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If Style Then
      Call SetWindowLong(hHeader, GWL_STYLE, Style)
      Call SetWindowPos(lvwLines.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.SetLvw"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
Dim strArg As String
Dim iStart As Integer
Dim iEnd As Integer
Dim oU As New z_UTIL
Dim strResult As String
Dim f As frmFindTextBite

    iStart = 0
    iEnd = 0
    iStart = InStr(1, txtTPMemo, "?") + 1
    If iStart = 0 Then Exit Sub
    strResult = ""
    iEnd = InStr(iStart, txtTPMemo, "?")
    If iStart > 0 And iEnd > iStart Then
        strArg = Trim(Mid(txtTPMemo, iStart, iEnd - iStart))
        strResult = oU.GetTextBite(strArg)
        If strResult > "" Then
            txtTPMemo = Replace(txtTPMemo, "?" & strArg & "?", strResult)
        End If
    Else
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtTPMemo_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DblClick()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.top = txtTPMemo.top + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    Else
        bMemoExpanded = True
        txtTPMemo.Height = txtTPMemo.Height + 800
        txtTPMemo.Width = txtTPMemo.Width + 800
        txtTPMemo.top = txtTPMemo.top - 800
        txtTPMemo.ZOrder 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtTPMemo_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_LostFocus()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.top = txtTPMemo.top + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtTPMemo_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)

'    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oCO.TRID
    oCO.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, X As Single, _
    Y As Single, State As Integer)
    On Error GoTo errHandler
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, X, Y)
        txtTPMemo.SelLength = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, X, Y, State), EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, X As Single, _
    Y As Single)
    On Error GoTo errHandler
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, X, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oCO.TRID
    oCO.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, X, Y), EA_NORERAISE
    HandleError
End Sub


Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oC As a_COL
Dim s As String

    Set rs = oPC.LinesClipboard
    If rs.State = 0 Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oCO.DOCCode, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
   ' rs.Open
    If rs.BOF And rs.eof Then Exit Sub
    rs.MoveFirst
    Do While Not rs.eof
        Set oC = oCO.COLines.Add
        oC.BeginEdit
        oC.PID = rs.fields("PID")
        oC.SetRef FNS(rs.fields("REF"))
        oC.QtyFirm = FNDBL(rs.fields("QtyFirm"))
        oC.QtySS = FNDBL(rs.fields("QTYSS"))
        oC.Qty = FNDBL(oC.QtySS) + FNDBL(oC.QtyFirm)
        oC.Price = FNDBL(rs.fields("Price"))
        oC.SetDiscount FNDBL(rs.fields("DISCOUNTRATE"))
        oC.CodeF = FNS(rs.fields("CODEF"))
        oC.EAN = FNS(rs.fields("EAN"))
        oC.SetFulfilled "OS"
        oC.Title = FNS(rs.fields("TITLE"))
        oC.VATRate = FNS(rs.fields("VATRATE"))
        oC.ApplyEdit
        rs.MoveNext
    Loop
    rs.Close
    oCO.ApplyEdit s
    oCO.BeginEdit
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuPastelines"
End Sub


