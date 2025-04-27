VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmQuotation 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Quotation"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11010
   ControlBox      =   0   'False
   Icon            =   "frmQuotation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11010
   Begin MSComctlLib.ListView lvwDocLines 
      Height          =   2475
      Left            =   75
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1200
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   4366
      SortKey         =   7
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Disc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ref"
         Object.Width           =   1834
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
   Begin VB.CheckBox chkChargeVAT 
      BackColor       =   &H00D3D3CB&
      Caption         =   "&Deduct VAT on foreign invoice"
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
      Left            =   900
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5880
      Width           =   3045
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
      Height          =   650
      Left            =   8730
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmQuotation.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5355
      UseMaskColor    =   -1  'True
      Width           =   1020
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
      Height          =   975
      Left            =   3990
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5325
      Width           =   2685
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
      Height          =   650
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5355
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Height          =   650
      Left            =   7725
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmQuotation.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1020
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1650
      Left            =   60
      TabIndex        =   13
      Top             =   3675
      Width           =   10725
      Begin VB.TextBox txtExtraCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   105
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1185
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtExtraCharge 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1065
         TabIndex        =   7
         Top             =   1185
         Visible         =   0   'False
         Width           =   1000
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
         Left            =   2055
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   855
         Width           =   645
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5475
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   435
         Width           =   3555
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1725
         TabIndex        =   1
         Top             =   435
         Width           =   765
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   3
         Top             =   435
         Width           =   1125
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4695
         TabIndex        =   4
         Top             =   435
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
         Height          =   650
         Left            =   9675
         MaskColor       =   &H00C4BCA4&
         Picture         =   "frmQuotation.frx":2EB6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   9075
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   435
         Width           =   1620
      End
      Begin VB.TextBox txtTitle 
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
         Height          =   330
         Left            =   4470
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1215
         Width           =   5115
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2505
         TabIndex        =   2
         Top             =   435
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   435
         Width           =   1635
      End
      Begin VB.Label lblFCTerms 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2880
         TabIndex        =   45
         Top             =   720
         Width           =   2265
      End
      Begin VB.Label lblExtraCharge 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   2100
         TabIndex        =   44
         Top             =   1215
         Width           =   2145
      End
      Begin VB.Label lblExtra1 
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Visible         =   0   'False
         Width           =   825
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
         Left            =   1080
         TabIndex        =   42
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "dbl-click to enlarge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   240
         Left            =   7050
         TabIndex        =   39
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label lblNote 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Note"
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
         Left            =   5475
         TabIndex        =   38
         Top             =   210
         Width           =   1440
      End
      Begin VB.Label lblRef 
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
         Left            =   3795
         TabIndex        =   30
         Top             =   210
         Width           =   555
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
         Left            =   4530
         TabIndex        =   19
         Top             =   210
         Width           =   1005
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
         Left            =   9630
         TabIndex        =   18
         Top             =   210
         Width           =   645
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
         Height          =   225
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label lblqty 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Qty"
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
         Left            =   1875
         TabIndex        =   16
         Top             =   210
         Width           =   600
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
         Left            =   2625
         TabIndex        =   15
         Top             =   210
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Issue"
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
      Height          =   650
      Left            =   9780
      Picture         =   "frmQuotation.frx":3240
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5355
      UseMaskColor    =   -1  'True
      Width           =   1020
   End
   Begin CoolButtonControl.CoolButton cmdBill 
      Height          =   1065
      Left            =   6390
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   60
      Width           =   1980
      _ExtentX        =   3493
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
   Begin CoolButtonControl.CoolButton cmdDel 
      Height          =   1065
      Left            =   8820
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   60
      Width           =   1980
      _ExtentX        =   3493
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
   Begin CoolButtonControl.CoolButton cbComp 
      Height          =   360
      Left            =   1260
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   75
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
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1050
      Left            =   3270
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   60
      Width           =   2340
      _ExtentX        =   4128
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
   Begin CoolButtonControl.CoolButton cbHeader 
      Height          =   390
      Left            =   30
      TabIndex        =   40
      ToolTipText     =   "Show header information"
      Top             =   75
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   688
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmQuotation.frx":35CA
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.Label lblTPFax 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   3915
      TabIndex        =   37
      Top             =   780
      Width           =   1545
   End
   Begin VB.Label lblTPPhone 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   3900
      TabIndex        =   36
      Top             =   465
      Width           =   1545
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   3915
      TabIndex        =   35
      Top             =   150
      Width           =   1545
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
      Left            =   8505
      TabIndex        =   29
      Top             =   60
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
      Left            =   6090
      TabIndex        =   28
      Top             =   60
      Width           =   300
   End
   Begin VB.Label lblAddDel 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   8895
      TabIndex        =   27
      Top             =   90
      Width           =   1710
   End
   Begin VB.Label lblAddBill 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   26
      Top             =   90
      Width           =   1710
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   3390
      TabIndex        =   22
      Top             =   135
      Width           =   375
   End
   Begin VB.Label Label3 
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
      Left            =   660
      TabIndex        =   21
      Top             =   135
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3465
      Picture         =   "frmQuotation.frx":37A4
      Stretch         =   -1  'True
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "frmQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oDOC As a_QU
Attribute oDOC.VB_VarHelpID = -1
Dim WithEvents oDOCL As a_QUL
Attribute oDOCL.VB_VarHelpID = -1
Dim oCustomer As a_Customer
Dim oProd As a_Product
Dim bCancelled As Boolean
Dim bValidInvoice As Boolean
Dim bValidInvoiceLine As Boolean
Dim tlCustomer As z_TextList
Dim oCurrentForeignCurrency As a_Currency
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
Dim oSM As z_StockManager
Dim bShowExtracharges As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property
Public Sub component(Optional pCustID As Long, Optional pDoc As a_QU)
    On Error GoTo errHandler
Dim strAddress As String
Dim frm As frmHeader
Dim strOrderNumber As String
Dim strOrderDate As String
Dim strMemo As String
Dim strForAttn As String

    If pDoc Is Nothing Then   'we create new invoice
        Set frm = New frmHeader

        frm.component
        frm.Show vbModal
        strMemo = frm.Memo
        strForAttn = frm.ForAttn
        Unload frm
        
        Set oDOC = New a_QU
        oDOC.BeginEdit
        oDOC.SetMemo strMemo
        oDOC.SetForAttn strForAttn
        
        
        If pCustID > 0 Then
            flgLoading = True
            LoadNewCustomer pCustID
            If Not oDOC.BillTOAddress Is Nothing Then
            End If
            If Not oDOC.DelToAddress Is Nothing Then
                strAddress = oDOC.DelToAddress.AddressMailing
                lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
            End If
            If Not oDOC.BillTOAddress Is Nothing Then
                strAddress = oDOC.BillTOAddress.AddressMailing
                lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
            End If
            flgLoading = False
        End If
'Handle interface
        ChangeState enAddingRow
        
    Else   'we are provided with a loaded invoice
        Set oDOC = pDoc
        oDOC.BeginEdit
        WaitMsg "Preparing to edit quotation  . . .", True, Me
        flgLoading = True
        If Not oDOC.BillTOAddress Is Nothing Then
            strAddress = oDOC.BillTOAddress.AddressMailing
            lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
            If oDOC.BillTOAddress.CountryID <> oPC.Configuration.LocalCountryID Then
                chkChargeVAT.Enabled = True
            Else
                chkChargeVAT.Enabled = False
            End If
        End If
        If Not oDOC.DelToAddress Is Nothing Then
            strAddress = oDOC.DelToAddress.AddressMailing
            lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
        End If
        
        flgLoading = False
'Handle interface
        oDOC.GetStatus
        oDOC.SetDirty False
        ChangeState enNotEditing
    End If
    
    oDOC.GetStatus
    
  '  cboRef.Visible = oPC.Configuration.SupportsWants
    SetMenu
        lblqty.Caption = "Qty"
      '  lblQtySS.Visible = False
      '  txtQtySS.Visible = False
        txtQty.Left = 1755
        txtQty.Width = 615
        lblqty.Left = 1860
        lblqty.Width = 375
        txtPrice.Left = 2400
        txtPrice.Width = 1000
        lblPrice.Left = 2520
        lblPrice.Width = 555
        txtRef.Left = 3435
        txtRef.Width = 1125
        lblRef.Left = 3690
        lblRef.Width = 555
        txtDiscount.Left = 4590
        txtDiscount.Width = 735
        lblDiscount.Left = 4425
        lblDiscount.Width = 1005
        txtNote.Left = 5355
        txtNote.Width = 3615
        lblNote.Left = 5355
        lblNote.Width = 1440
        cmdCancel.Left = 7740
        cmdSave.Left = 8760
      ' cmdPick.Visible = False

        Caption = "Quotation for " & oDOC.Customer.NameAndCode(25) & oDOC.StaffNameB
    bShowExtracharges = True
Exithandler:
    WaitMsg "", False, Me
    
        
    Exit Sub
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Component(pCustID,pDOC)", Array(pCustID, pDoc)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.component(pCustID,pDoc)", Array(pCustID, pDoc)
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim strAddress As String

    LoadComps
    
    If Not oDOC.IsNew Then
        chkChargeVAT = IIf(oDOC.ShowVAT, 1, 0)
    Else
        chkChargeVAT = IIf(oPC.Configuration.DiscountVATDefault, 1, 0)
    End If
    
  '  SetLvw
    LoadCustomerDetailsToForm
    LoadListView
    
    oDOC.SetDirty False

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.LoadControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.LoadControls"
End Sub


'Private Sub cbHeader_Click()
'Dim frm As New frmHeader
'Dim strOrderNumber As String
'Dim strOrderDate As String
'Dim strMemo As String
'
'    frm.Component False, "Customer reference number", "Date", oDOC.OrderNumber, oDOC.OrderDateF, oDOC.Memo
'    frm.Show vbModal
'    strOrderNumber = frm.OrderNumber
'    strOrderDate = frm.OrderDate
'    strMemo = frm.Memo
'    Unload frm
'    oDOC.setOrderNumber strOrderNumber
'    oDOC.setMemo strMemo
'    If strOrderDate > "" Then oDOC.SetOrderDate CDate(strOrderDate)
'End Sub


Private Sub cmdExtraCharge_Click()
    On Error GoTo errHandler
    
    bShowExtracharges = Not bShowExtracharges
    ControlCaptureFrame bShowExtracharges
    If bShowExtracharges Then mSetfocus txtExtraCode
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdExtraCharge_Click", , EA_NORERAISE
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
    txtPrice.Enabled = Not bExtra
    lblPrice.Enabled = Not bExtra
    txtDiscount.Enabled = Not bExtra
    lblDiscount.Enabled = Not bExtra
    txtTotal.Enabled = Not bExtra
    lblTotal.Enabled = Not bExtra
    Me.lblRef.Enabled = Not bExtra
    Me.txtRef.Enabled = Not bExtra
    txtNote.Enabled = Not bExtra
    lblNote.Enabled = Not bExtra
 '   txtExtraCode = ""
 '   txtExtraCharge = ""
    If bExtra Then
        If oDOCL.ExtraCode > "" Then
            txtExtraCode = oDOCL.ExtraCode
            txtExtraCharge = oDOCL.ExtraCharge
            Me.lblExtraCharge.Caption = oDOCL.ExtraChargeDescription
        End If
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.ControlCaptureFrame(bExtra)", bExtra
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.ControlCaptureFrame(bExtra)", bExtra
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
  ' cmdAppro.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Form_Activate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Form_Deactivate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oDOC.StatusF = "IN PROCESS" And oDOC.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oDOC.StatusF = "ISSUED") ' And oDOC.CanCancel = True
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.SetMenu"
End Sub



Private Sub cbComp_Click()
    On Error GoTo errHandler
    oDOC.COMPID = OptionLoop(oDOC.COMPID, oPC.Configuration.Companies.Count)
    cbComp.Caption = oPC.Configuration.Companies(oDOC.COMPID).CompanyName
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cbComp_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cbComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oDOC.Customer.ID > 0 Then
        frm.component oDOC.Customer
        frm.Show
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cbCust_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub



'Private Sub cboRef_SelectionChanged()
'    On Error GoTo errHandler
'    oDOCL.COLID = cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 3)
'    oDOCL.SetQty cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 2)
'    If oDOC.Customer.UseQuotedPrice Then
'        oDOCL.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4)
'        oDOCL.SetPrice cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6)
'    Else
'        If cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4) <> oDOC.Customer.DefaultDiscount Then
'            DoEvents
'            If MsgBox("The customer discount is different than the discount on this line. Use the customer discount?", vbYesNo, "Warning") = vbYes Then
'                oDOCL.SetDiscountPercent oDOC.Customer.DefaultDiscount
'            Else
'                oDOCL.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4)
'            End If
'        End If
'        If cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6) <> oProd.SP Then
'            If MsgBox("The product price is different than the price on this line. Use the product price?", vbYesNo, "Warning") = vbYes Then
'                oDOCL.SetPrice oProd.SP
'            Else
'                oDOCL.SetPrice cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6)
'            End If
'        End If
'    End If
'
'    oDOCL.SetRef cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 0)
'
'    txtDiscount = oDOCL.DiscountPercent
'    txtPrice = oDOCL.Price
'    txtRef = oDOCL.Ref
'        txtQty = oDOCL.Qty
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cboRef_SelectionChanged", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub chkChargeVAT_Click()
    On Error GoTo errHandler
    oDOC.ShowVAT = (chkChargeVAT = 1)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.chkChargeVAT_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.chkChargeVAT_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdBill_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oDOC.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oDOC.Customer.Addresses.Count Then
        i = 1
    End If
    lblAddBill.Caption = oDOC.Customer.Addresses(i).AddressMailing & vbCrLf & oDOC.Customer.Addresses(i).EMail
    oDOC.SetBillToAddress oDOC.Customer.Addresses(i)
    iBillIdx = i
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdBill_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdBill_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oDOC.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oDOC.Customer.Addresses.Count Then
        i = 1
    End If
    lblAddDel.Caption = oDOC.Customer.Addresses(i).AddressMailing & vbCrLf & oDOC.Customer.Addresses(i).EMail
    oDOC.setDelToAddress oDOC.Customer.Addresses(i)
    iBillIdx = i

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdDel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdDel_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oDOC.SetCustomer(plngTPID) Then
        vCanAdd.RuleBroken "TP", False
        LoadCustomerDetailsToForm
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.LoadNewCustomer(plngTPID)", plngTPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.LoadNewCustomer(plngTPID)", plngTPID
End Sub
Private Sub LoadCustomerDetailsToForm()
    On Error GoTo errHandler
    With oDOC.Customer
        If Not .BillTOAddress Is Nothing Then
            lblTPPhone.Caption = .BillTOAddress.Phone
            lblTPFax.Caption = .BillTOAddress.Fax
        End If
        lblTPName.Caption = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
        If Not .BillTOAddress Is Nothing Then
            If oDOC.BillTOAddress Is Nothing Then
                oDOC.SetBillToAddress .BillTOAddress
                lblAddBill.Caption = .BillTOAddress.AddressShort
            End If
        End If
        If Not .DelToAddress Is Nothing Then
            If oDOC.DelToAddress Is Nothing Then
                oDOC.setDelToAddress .DelToAddress
                lblAddDel.Caption = .DelToAddress.AddressShort
            End If
        End If
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.LoadCustomerDetailsToForm"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.LoadCustomerDetailsToForm"
End Sub
Private Sub cmdNote()
    On Error GoTo errHandler
Dim frm As New frmILNote
    frm.component oDOCL
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdNote"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdNote"
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set vCanAdd = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Form_Terminate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAddresses()
    On Error GoTo errHandler
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oDOC
    frm.Show vbModal
    lblAddBill.Caption = oDOC.BillTOAddress.AddressShort
    lblAddDel.Caption = oDOC.DelToAddress.AddressShort
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.mnuAddresses"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuAddresses"
End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveLine
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.mnuDelLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuDelLine"
End Sub

Private Sub oDOC_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidInvoice = (pMsg = "")
    
    cmdIssue.Enabled = (bValidInvoice And oDOC.QuoteLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidInvoice) And oDOC.IsDirty
    cmdCancel.Enabled = True  'oDOC.Status = stISSUED Or oDOC.Status = stCOMPLETE
    
    Me.txtError = pMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOC_Valid(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOC_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oDOCL_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOCL_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
'         strExtension), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOCL_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oDOCL_Valid(msg As String)
    On Error GoTo errHandler
        Me.cmdEnter.Enabled = (msg = "")
        Me.txtError = msg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOCL_Valid(Msg)", msg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOCL_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oDOC_TotalChange(lngTotalExt As Long, lngTotalDeposit As Long, lngTotalVAT As Long)
    On Error GoTo errHandler
    
    flgLoading = True
    
    lngCurrentTotal = lngTotalExt
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
  '  If vMode = enEditingRow Then
  '      cmdNewRows.Enabled = (oDOC.QuoteLines.Count > 0)
  '  End If

    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOC_TotalChange(lngTotalExt,lngTotalDeposit,lngTotalVAT)", _
'         Array(lngTotalExt, lngTotalDeposit, lngTotalVAT), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOC_TotalChange(lngTotalExt,lngTotalDeposit,lngTotalVAT)", _
         Array(lngTotalExt, lngTotalDeposit, lngTotalVAT), EA_NORERAISE
    HandleError
End Sub

Private Sub oDOC_Reloadlist()
    On Error GoTo errHandler
    LoadListView
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOC_Reloadlist", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOC_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oDOC_Dirty(pVal As Boolean)
    On Error GoTo errHandler
    If flgLoading And pVal Then Exit Sub
    If pVal = True Then
        cmdSave.Enabled = pVal 'And vMode = enNotEditing
        Me.cmdIssue.Enabled = pVal And (oDOC.Status <> stCOMPLETE And oDOC.Status <> stISSUED) 'And vMode = enNotEditing
        cmdCancel.Caption = "&Cancel"
    Else
        cmdSave.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdCancel.Enabled = True
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOC_Dirty(pVal)", pVal, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOC_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oDOC_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.oDOC_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOC_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub


Private Sub SetFocusFromCode()
    On Error GoTo errHandler
Dim strMsg As String
    
    If LenB(txtCode) > 0 Then
            mSetfocus txtQty
    End If

    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.SetCursorFromCode"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.SetFocusFromCode"
End Sub

Private Sub txtExtraCharge_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtExtraCharge")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtExtraCharge_GotFocus"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtExtraCharge_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtExtraCharge_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDOCL.SetExtraCharge(txtExtraCharge) Then
      '  Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtExtraCharge_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtExtraCharge_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtExtraCharge_LostFocus()
    On Error GoTo errHandler
  '  txtExtraCharge = oCOLine.PriceF
    txtTotal = oDOCL.ExtF(False)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtExtraCharge_LostFocus"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtExtraCharge_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtExtraCode_LostFocus()
    On Error GoTo errHandler
 '   mSetfocus txtExtraCharge
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtExtraCode_LostFocus", , EA_NORERAISE
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
        bOK = oDOCL.SetLineExtraProduct(txtExtraCode)
        If bOK Then
                Me.lblExtraCharge = oDOCL.ExtraChargeDescription
                txtPrice = oDOCL.Price
                txtExtraCharge.Enabled = True
            '    AutoSelect txtExtraCharge
            '    mSetfocus txtExtraCharge
        Else
            Cancel = True
            Me.txtExtraCharge.Enabled = False
        End If
    Else
        oDOCL.SetExtraCharge 0
        oDOCL.ExtraPID = ""
        oDOCL.ExtraCode = ""
        txtExtraCharge.Enabled = False
        
    End If

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtExtraCode_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtExtraCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_DblClick()
    On Error GoTo errHandler
    txtNote = HandleTextWithBites(txtNote)
    If txtNote.Height = 1125 Then
        txtNote.Height = 375
    Else
        txtNote.Height = 1125
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtNote_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_DblClick()
    On Error GoTo errHandler
Dim f As New frmFCPrice
    
Dim X As Long
Dim Y As Long
    
    If Not oPC.SupportsUNISA Then Exit Sub
    
    f.component Me.Left + 2000, Me.top + 2000, oDOCL.ForeignPrice, oDOCL.Price, oDOCL.FCID, oDOCL.VATRate, oDOCL.FCFactor

    f.Show vbModal
    If f.UserCancelled Then
        Unload f
        Exit Sub
    End If
    oDOCL.SetFCFactor Round(f.FCFactor, 6)
    oDOCL.SetForeignPrice CStr(f.ForeignPrice)
    oDOCL.FCID = f.FCID
    oDOCL.Price = f.LocalPriceIncVAT
    Me.txtPrice = oDOCL.Price
    lblFCTerms.Caption = oDOCL.ForeignPriceF & "/" & f.FCFactorINV
    Unload f
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtPrice_DblClick"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtPrice_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQty_DblClick()
    On Error GoTo errHandler
    If Not oPC.SupportsUNISA Then Exit Sub
    txtQty = oDOC.TotalQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtQty_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQty
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtQty_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If vMode = enNotEditing Then Exit Sub
        If Not oDOCL.SetQty(txtQty) Then
            Cancel = True
        End If
    oDOCL.CalculateLine
    txtTotal = oDOCL.ExtF(False)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtRef
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtRef_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtRef_GotFocus", , EA_NORERAISE
    HandleError
End Sub
''============
'Private Sub txtQtySS_GotFocus()
'    On Error GoTo errHandler
'    AutoSelect txtQtySS
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtQtySS_GotFocus", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtQtySs_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    If vMode = enNotEditing Then Exit Sub
'    If Not oDOCL.SetQtySS(txtQtySS) Then
'        Cancel = True
'    End If
'    oDOCL.CalculateLine
'    txtTotal = oDOCL.PLessDiscExtF(False)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtQtySS_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub


'========
Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    Me.cmdCancel.Enabled = True
   ' Me.cmdPick.Enabled = True
    Me.cmdSave.Enabled = True
    Me.cmdIssue.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.vCanAdd_NobrokenRules", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.vCanAdd_NobrokenRules", , EA_NORERAISE
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
    LoadControls
            'LogSaveToFile "Invoice form Loaded"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Form_Initialize", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
            'LogSaveToFile "Invoice form unloaded"
    If oDOC.IsEditing Then oDOC.CancelEdit
    UnsetMenu
    Set oCustomer = Nothing
    Set oDOC = Nothing
    Set tlCustomer = Nothing
    Set oDOCL = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmQuotationPreview

    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmQuotationPreview
        frm.ComponentObject oDOC
        frm.Show
    End If
    If cmdCancel.Caption <> "&Close" Then
        If oDOC.IsEditing And oDOC.IsDirty Then
            If MsgBox("You wish to cancel your changes?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
            oDOC.CancelEdit
        End If
    End If
    If Not oDOCL Is Nothing Then
        If oDOCL.IsEditing Then oDOCL.CancelEdit
    End If
        'LogSaveToFile "Invoice Cancel button"
    Unload Me
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim strLine As String
Dim strItemsDebug As String
Dim i As Integer

    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        mSetfocus txtCode
        Exit Sub
    End If
    If oDOCL.ServiceItem Then oDOCL.DiscountPercent = 0
    
    oDOCL.ApplyEdit
    oDOCL.BeginEdit
    
    If vMode = enAddingRow Then
        If lvwDocLines.ListItems.Count < val(oDOCL.Key) Then
            lvwDocLines.ListItems.Add Key:=oDOCL.Key
            LoadListViewLine lvwDocLines.ListItems(lvwDocLines.ListItems.Count), oDOCL
        End If
        lvwDocLines.Refresh
        ChangeState enAddingRow
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwDocLines.ListItems(lngSelectedRowIndex), oDOCL
        ChangeState enNotEditing
    End If
    oDOC.GetStatus
    ControlCaptureFrame False
    bShowExtracharges = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdEnter_Click", , EA_NORERAISE, , "vMode,Linecount,strItemsDebug,oDOCL.Key", Array(vMode, oDOC.QuoteLines.Count, strItemsDebug, oDOCL.key)
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdEnter_Click", , EA_NORERAISE
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
        txtRef.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQty.Enabled = True
       ' cboRef.Visible = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
      '  cmdPick.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oDOC.QuoteLines.Count > 0)
        lvwDocLines.Enabled = False
        lvwDocLines.Height = 2200
        UnsetMenu
        fr1.ZOrder 1
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtRef.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQty.Enabled = True
        txtError = ""
        flgLoading = True
        txtRef = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
       ' cmdPick.Enabled = False
        cmdSave.Enabled = True 'TEST change 18-09-2007
        cmdNewRows.Enabled = (oDOC.QuoteLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        lblTPPhone.Caption = ""
        lvwDocLines.Enabled = False
        lvwDocLines.Height = 2200
        ClearInvLineControls
        fr1.ZOrder 1
        mSetfocus txtCode
        Set oDOCL = oDOC.QuoteLines.Add
        oDOCL.QuoteID = oDOC.QuoteID
        oDOCL.SetQty 1
        UnsetMenu
    Case enNotEditing
        flgLoading = True
        fr1.Visible = False
        txtError = ""
        txtRef = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True '(oDOC.Status <> stCOMPLETE And oDOC.Status <> stISSUED)
       ' cmdPick.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = True  '(oDOC.QuoteLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        lvwDocLines.Enabled = True
        lvwDocLines.Height = 4000
        SetMenu
        fr1.ZOrder 1
    End Select
    If Not oDOC.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
'    If oDOC.Status = stISSUED Then
'        Me.cmdCancel.Enabled = False
'        Me.cmdSave.Enabled = False
'    End If
    
   ' lblAppro.Caption = ""
   ' cboRef.Visible = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.ChangeState(pToMode)", pToMode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.ChangeState(pToMode)", pToMode
End Sub
Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
        'LogSaveToFile "Invoice New row button:enEditingRow"
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        'LogSaveToFile "Invoice New row button:enAddingRow"
        If txtCode > "" Then  'THis is not after a post but is an aborted  add row action
            oDOC.QuoteLines.DecrementMaxKeyUsed
        End If
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        'LogSaveToFile "Invoice New row button:enNotEditing"
        ChangeState enAddingRow
    End If

    ClearInvLineControls
    ControlCaptureFrame False
    bShowExtracharges = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdNewRows_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
Dim strItemsDebug As String

    For i = 1 To lvwDocLines.ColumnHeaders.Count
        lvwDocLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwDocLines.ColumnHeaders(i).Width)
    Next
    lvwDocLines.ListItems.Clear
    For i = 1 To oDOC.QuoteLines.Count
        Set lstItem = lvwDocLines.ListItems.Add
        LoadListViewLine lstItem, oDOC.QuoteLines.Item(i)
        'strItemsDebug = strItemsDebug & "," & lvwDocLines.ListItems(i).Key
    Next i
    'Debug.Print strItemsDebug
EXIT_Handler:
    Set lstItem = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.LoadListView"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oDOCL As a_QUL)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oDOCL
        lstItem.text = .code
        lstItem.Key = .Key
        lstItem.SubItems(1) = .TitleAuthorPublisher
            lstItem.SubItems(2) = .Qty
        lstItem.SubItems(3) = .PriceF(False)
        lstItem.SubItems(4) = .DiscountPercentF  ' Format(.DiscountPercent, "##0.0%")
        lstItem.SubItems(5) = .Ref
        lstItem.SubItems(6) = .ExtF(False)
        lstItem.SubItems(7) = Format(.Key, "@@@@@@@@@@")
        lstItem.SubItems(8) = .EAN
        If .ServiceItem = True Then
            lstItem.ForeColor = &H427182
            lstItem.ListSubItems(1).ForeColor = &H427182
            lstItem.ListSubItems(2).ForeColor = &H427182
            lstItem.ListSubItems(3).ForeColor = &H427182
            lstItem.ListSubItems(4).ForeColor = &H427182
            lstItem.ListSubItems(5).ForeColor = &H427182
            lstItem.ListSubItems(6).ForeColor = &H427182
            lstItem.ListSubItems(7).ForeColor = &H427182
        Else
            lstItem.ListSubItems(1).ForeColor = &H706034
            lstItem.ListSubItems(2).ForeColor = &H706034
            lstItem.ListSubItems(3).ForeColor = &H706034
            lstItem.ListSubItems(4).ForeColor = &H706034
            lstItem.ListSubItems(5).ForeColor = &H706034
            lstItem.ListSubItems(6).ForeColor = &H706034
            lstItem.ListSubItems(7).ForeColor = &H706034
        End If
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.LoadListViewLine(lstItem)", Array(lstItem)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.LoadListViewLine(lstItem,oDOCL)", Array(lstItem, oDOCL)
End Sub
Private Sub lvwDocLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwDocLines.ListItems.Count = 0 Then Exit Sub
    If lvwDocLines.SelectedItem.Index < 1 Then Exit Sub
    
    lngEditingIdx = lvwDocLines.SelectedItem.Key
    
    Set oDOCL = Nothing
    Set oDOCL = oDOC.QuoteLines.Item(lngEditingIdx)
    
    lngSelectedRowIndex = lvwDocLines.SelectedItem.Key

    ChangeState eneditingrow

    Set oProd = Nothing
    Set oProd = New a_Product
    oProd.Load oDOCL.PID, 0
    
    txtDiscount = oDOCL.DiscountPercentF
        txtQty = oDOCL.QtyF
    txtNote = oDOCL.Note
    txtRef = oDOCL.Ref
    txtDiscount = oDOCL.DiscountPercent
    txtCode = IIf(oDOCL.code > "", CStr(oDOCL.code), CStr(oProd.code))
    txtTitle = oDOCL.Title
    txtQty = oDOCL.Qty
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oDOCL.PriceF(False)
    Else
        txtPrice = oDOCL.Price
    End If
    If oDOCL.Qty > 1 Then
        mSetfocus txtQty
    Else
        mSetfocus txtPrice
    End If
    If oDOCL.FCID <> oPC.Configuration.DefaultCurrencyID And oDOCL.FCID > 0 Then
        Me.lblFCTerms = oDOCL.ForeignPriceF & "/" & oDOCL.FCFactorInvF
    End If
    
    oDOCL.GetStatus
    If oDOCL.ExtraCode > "" Then
        ControlCaptureFrame True
    End If
    bShowExtracharges = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.lvwDocLines_DblClick", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.lvwDocLines_DblClick", , EA_NORERAISE
    HandleError
End Sub

'---------Companies code
Private Sub LoadComps()
    On Error GoTo errHandler
Dim oComp As a_Company
Dim oItem As ListItem
Dim i As Integer
    If oDOC.COMPID > 0 Then
        cbComp.Caption = oPC.Configuration.Companies(CStr(oDOC.COMPID)).CompanyName
    Else
        cbComp.Caption = oPC.Configuration.DefaultCompany.CompanyName
        oDOC.COMPID = oPC.Configuration.DefaultCOMPID
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.LoadComps"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.LoadComps"
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If oDOC.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oDOCL.SetNote(txtNote)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oDOCL.Note
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtNote_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oDOC.Memo
    ofrm.Show vbModal
    oDOC.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuMemo"
End Sub

Public Sub mnuCancel()
    On Error GoTo errHandler
    If oDOC.IsDirty Then
        oDOC.CancelEdit
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.mnuCancel"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuCancel"
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    oDOC.SetStatus stVOID
    oDOC.ApplyEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuVoid"
End Sub


Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim lngResult As Long
Dim pNumOfApproLines As Long
Dim ErrorPos As String
Dim i As Integer
Dim rsPreviousBillings As ADODB.Recordset
Dim strPreviousBillings As String
Dim oSQL As New z_SQL
Dim strPos As String
Dim oSM As New z_StockManager
strPos = "Pos 1"
    If txtCode = "" Or vMode = eneditingrow Then Exit Sub
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_Handler
    End If
    
START:
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
        .Load 0, 0, Trim$(txtCode)
        If oProd.PID > "" Then
        End If
        If Len(FNS(.PID)) <> 0 Then   'Book in database
                oDOCL.DiscountPercent = oDOC.Customer.DefaultDiscount
            If oProd.IsServiceItem Then   'No copy identified but product is a non-stock product (e.g. postage or insurance etc.)
                oDOCL.Price = oProd.SP
                mSetfocus txtPrice
                AutoSelect txtPrice
            Else    ' we may reach here is a copy is requested and not found
                    ' OR No copy is requested and the Title is found
                oDOCL.Price = oProd.SP
                oDOCL.CodeForExport = oProd.CodeForExport
                oDOCL.CodeF = oProd.CodeF
                oDOCL.code = oProd.EAN
                If oPC.Configuration.AllowCopyInfo And InStr(txtCode, "/") > 0 Then
                    If MsgBox("There is no copy with this serial number" & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Papyrus Invoicing Information") = vbNo Then
                        Cancel = True
                        Exit Sub
                    End If
                End If
            End If
            oDOCL.Title = .TitleAuthor  'L(35)
            oDOCL.PID = .PID
            oDOCL.ServiceItem = .IsServiceItem
            oDOCL.VATRate = .VATRateToUse
            If oDOCL.IsNew And oDOCL.DiscountPercent = 0 Then
                oDOCL.DiscountPercent = oDOC.Customer.DefaultDiscount
            End If
            If oDOCL.DiscountPercent <> oDOC.Customer.DefaultDiscount And oDOCL.IsNew = False Then
                If MsgBox("The discount on the invoice differs from the customer's usual discount. " & vbCrLf & "Use discount on order?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                    oDOCL.DiscountPercent = oDOC.Customer.DefaultDiscount
                End If
            End If
        Else   'Book nof found on database
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
    End With
    txtTitle = oDOCL.TitleAuthor
    txtPrice = oDOCL.Price
    txtQty = oDOCL.Qty
    txtRef = oDOCL.Ref
    txtDiscount = oDOCL.DiscountPercentF
    oDOCL.GetStatus
    SetFocusFromCode
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
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

Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not oDOCL.SetDiscountPercent(txtDiscount) Then
        Cancel = True
    End If
    oDOCL.CalculateLine
    txtTotal = oDOCL.ExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDOCL.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oDOCL.CalculateLine
    txtTotal = oDOCL.ExtF(False)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtPrice_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_Change()
        On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oDOCL.SetRef (txtRef)
    If Err Then
      Beep
      intPos = txtRef.SelStart
      txtRef = oDOCL.Ref
      txtRef.SelStart = intPos - 1
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtRef_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oDOCL.SetRef(txtRef)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtRef = oDOCL.Ref
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.txtRef_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.txtRef_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub RemoveLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwDocLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwDocLines.ListItems(i).Selected Then
            oDOC.QuoteLines.Remove lvwDocLines.ListItems(i).Key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Exit Sub
    End If
    lvwDocLines.ListItems.Remove i
    lvwDocLines.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.RemoveLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.RemoveLine"
End Sub

Private Sub SaveDoc()
    On Error GoTo errHandler
Dim strErrPos As String

'strErrPos = "Pos 1"
'If oDOC Is Nothing Then
'        MsgBox "SaveDoc: oDOC is nothing"
'End If
'MsgBox "Pos 1a"
    oDOC.ApplyEdit
'strErrPos = "Pos 2"
'If oDOC Is Nothing Then
'        MsgBox "SaveDoc: oDOC is nothing"
'End If
    oDOC.BeginEdit
'strErrPos = "Pos 3"
'If oDOC.QuoteLines Is Nothing Then
'        MsgBox "SaveDoc: oDOC is nothing"
'End If
    Set oDOCL = oDOC.QuoteLines.Add
    
EXIT_Handler:
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.SaveDoc", , , , "strErrPos", Array(strErrPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.SaveDoc"
End Sub

'Public Sub PrintDoc()
'    On Error GoTo errHandler
'Dim blnDeposit As Boolean
'Dim blnDiscount As Boolean
'Dim blnRoundedUp As Boolean
'Dim blnNoDocLines As Boolean
'Dim blnHideVAT As Boolean
'Dim iCurrency As Integer
'
'
'    Me.MousePointer = vbHourglass
'    oDOC.Load oDOC.InvoiceID, False
'    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
'
'    If blnNoDocLines Then
'        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
'        GoTo EXIT_HANDLER
'    End If
'
'EXIT_HANDLER:
'    Me.MousePointer = vbDefault
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.PrintDoc"
'End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoDocLines As Boolean
Dim iCurrency As Integer
Dim strResult As String
Dim frm As frmQuotationPreview
Dim frmDte As frmTRDate
    If oDOC.QtyNonStandardVAT > 0 Then
        If MsgBox("There are items with non-standard VAT in this invoice, continue?", vbYesNo + vbInformation, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
'    If oDOC.Status = stInProcess Then
'            If oPC.Configuration.Signtransactions = True Then
'                If SecurityControl(enSECURITY_INV_SIGN, , "Sign this invoice.", DOCAPPROVAL) = False Then
'                       Exit Sub
'                End If
'            End If
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Sign this quotation", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oDOC.Status = stInProcess Then
            If MsgBox("Issue this quotation?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
            
           ' If oDOC.DocDate < CDate("1950-01-01") Then
                oDOC.DOCDate = Date
                oDOC.CaptureDate = Now()
           ' End If
            
            WaitMsg "Issuing quotation  . . .", True, Me
            oDOC.VATable = oDOC.Customer.VATable
            oDOC.StaffID = gSTAFFID
            oDOC.RecalculateAllLines
            oDOC.CalculateTotals
            strResult = oDOC.Post(stCOMPLETE)
            
            If strResult = "" Then
                Set frm = New frmQuotationPreview
                frm.ComponentObject oDOC
                frm.Show
            ElseIf strResult > "" Then
                MsgBox "The document cannot be issued now, try later. The record is probably locked by another user. The message is: " & strResult & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                oDOC.BeginEdit
                WaitMsg "", False, Me
                Exit Sub
            End If
   ' End If
EXITH:
    WaitMsg "", False, Me
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdIssueClick", , EA_NORERAISE
'    HandleError
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oDOC_DirtyStatus(pDirty As Boolean)
    On Error GoTo errHandler
    If pDirty = True Then
        cmdCancel.Caption = "&Cancel"
    Else
        cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.oDOC_DirtyStatus(pDirty)", pDirty, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oIL As a_QUL
    If oDOC.Status <> stCOMPLETE And oDOC.Status <> stISSUED And oDOC.Status <> stCANCELLED And oDOC.Status <> stVOID Then
        oDOC.SetStatus stInProcess
    End If
    If oDOC.DOCDate < CDate("1950-01-01") Then
        oDOC.DOCDate = Date
        oDOC.CaptureDate = Now()
    End If
    oDOC.RecalculateAllLines
    oDOC.CalculateTotals
        'LogSaveToFile "Invoice Saving button"
    SaveDoc
    LoadListView
    cmdSave.Enabled = False
    cmdCancel.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.cmdSave_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub ClearInvLineControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = ""
    Me.txtDiscount = ""
    Me.txtPrice = ""
    Me.txtTitle = ""
    Me.txtTotal = ""
    Me.txtNote = ""
    Me.txtRef = ""
    lblFCTerms.Caption = ""
    txtQty = ""
   ' cmdAppro.BackColor = &HC4BCA4
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.ClearInvLineControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.ClearInvLineControls"
End Sub

Private Sub lvwDocLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.lvwDocLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.lvwDocLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oDOC.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oDOC.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.SetIssueButtonCaption"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.SetIssueButtonCaption"
End Sub

Private Sub lvwDocLines_Click()
    On Error GoTo errHandler
    If lvwDocLines Is Nothing Then Exit Sub
    If lvwDocLines.SelectedItem Is Nothing Then Exit Sub
    
    If lvwDocLines.SelectedItem.Index > 0 And Left(lvwDocLines.SelectedItem.SubItems(8), ISBNLENGTH) > "" Then
    On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText Left(lvwDocLines.SelectedItem.SubItems(8), ISBNLENGTH)
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.lvwDocLines_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.lvwDocLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDocLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvwDocLines.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
    If lvwDocLines.SortOrder = lvwAscending Then
        lvwDocLines.SortOrder = lvwDescending
    Else
        lvwDocLines.SortOrder = lvwAscending
    End If
   lvwDocLines.Sorted = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.lvwDocLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.lvwDocLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim Style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwDocLines.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   Style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   Style = Style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If Style Then
      Call SetWindowLong(hHeader, GWL_STYLE, Style)
      Call SetWindowPos(lvwDocLines.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.SetLvw"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.SetLvw"
End Sub

Private Sub vCanAdd_Status(errors As String)
    On Error GoTo errHandler
MsgBox errors & "CANAADD"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotation.vCanAdd_Status(errors)", errors, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.vCanAdd_Status(errors)", errors, EA_NORERAISE
    HandleError
End Sub


Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvwDocLines, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuSaveLayout"
End Sub
Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_QUL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In oDOC.QuoteLines
 '       If Not oLine.Product.IsServiceItem Then
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        rs.fields("Qty") = oLine.Qty
        rs.fields("QtyFirm") = oLine.Qty
        rs.fields("Price") = oLine.Price
        rs.fields("DISCOUNTRATE") = oLine.DiscountPercent
        rs.fields("CODEF") = oLine.CodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("EAN") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oPC.Configuration.VATRate
        rs.fields("REF") = oLine.Ref
        rs.fields("ETA") = CDate(0)
        rs.Update
  '      End If
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
            rs.Save oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.mnuCopyLines"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_QUL
Dim s As String

    If rs Is Nothing Then Exit Sub
    If rs.State = 0 Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oDOC.DOCCode, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
   ' rs.Open
    If rs.BOF And rs.eof Then Exit Sub
    rs.MoveFirst
    Do While Not rs.eof
        Set oLine = oDOC.QuoteLines.Add
        oLine.BeginEdit
        oLine.PID = rs.fields("PID")
        oLine.Ref = FNS(rs.fields("REF"))
        oLine.Qty = FNDBL(rs.fields("Qty"))
        oLine.Price = FNDBL(rs.fields("Price"))
        oLine.DiscountPercent = FNDBL(rs.fields("DISCOUNTRATE"))
        oLine.CodeF = FNS(rs.fields("CODEF"))
        oLine.EAN = FNS(rs.fields("EAN"))
        oLine.Title = FNS(rs.fields("TITLE"))
        oLine.ApplyEdit
        rs.MoveNext
    Loop
    rs.Close
    oDOC.ApplyEdit
    oDOC.BeginEdit
    LoadControls
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotation.mnuPastelines"
End Sub
