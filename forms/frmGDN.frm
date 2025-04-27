VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmGDN 
   BackColor       =   &H00D3D3CB&
   Caption         =   "GDN"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11010
   ControlBox      =   0   'False
   Icon            =   "frmGDN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11010
   Begin VB.CommandButton cmdConvertToNonVat 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Convert to Non-VAT invoice"
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
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   6690
      Width           =   1875
   End
   Begin VB.CommandButton cmdPick 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Pick"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1020
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
      Left            =   1245
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5910
      Visible         =   0   'False
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
      Left            =   7740
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmGDN.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5370
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
      Left            =   4020
      MultiLine       =   -1  'True
      TabIndex        =   24
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
      TabIndex        =   10
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
      Left            =   6720
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmGDN.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1020
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1950
      Left            =   60
      TabIndex        =   14
      Top             =   3405
      Width           =   10725
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
         Left            =   7050
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1350
         Width           =   255
      End
      Begin VB.CommandButton cmdFind 
         Height          =   345
         Left            =   1755
         Picture         =   "frmGDN.frx":2EB6
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   135
         Width           =   375
      End
      Begin VB.TextBox txtQtySS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2863
         TabIndex        =   4
         Top             =   495
         Width           =   675
      End
      Begin VB.CommandButton cmdAppro 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Appro"
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
         Left            =   8250
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1320
         Width           =   720
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6495
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   495
         Width           =   2610
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2130
         TabIndex        =   3
         Top             =   495
         Width           =   705
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4594
         TabIndex        =   6
         Top             =   495
         Width           =   1125
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5745
         TabIndex        =   7
         Top             =   495
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
         Picture         =   "frmGDN.frx":3240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1230
         Width           =   1050
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   9135
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   495
         Width           =   1545
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
         Left            =   75
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1590
         Width           =   7110
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3566
         TabIndex        =   5
         Top             =   495
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   75
         TabIndex        =   0
         Top             =   495
         Width           =   2040
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboRef 
         Height          =   345
         Left            =   165
         OleObjectBlob   =   "frmGDN.frx":35CA
         TabIndex        =   2
         Top             =   1275
         Visible         =   0   'False
         Width           =   6885
      End
      Begin VB.Label lblFCTerms 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3915
         TabIndex        =   50
         Top             =   780
         Width           =   2265
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
         Left            =   7455
         TabIndex        =   49
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Height          =   240
         Left            =   5565
         TabIndex        =   48
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label lblQtySS 
         BackColor       =   &H00D3D3CB&
         Caption         =   "SOR"
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
         Left            =   2970
         TabIndex        =   46
         Top             =   255
         Width           =   495
      End
      Begin VB.Label lblAppro 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   7395
         TabIndex        =   45
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label lblO1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Order ref"
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
         Height          =   240
         Left            =   225
         TabIndex        =   43
         Top             =   1035
         Width           =   1515
      End
      Begin VB.Label lblO4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
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
         Height          =   240
         Left            =   4515
         TabIndex        =   42
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label lblO3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Height          =   240
         Left            =   3585
         TabIndex        =   41
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label lblO2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Document"
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
         Height          =   240
         Left            =   2175
         TabIndex        =   40
         Top             =   1035
         Width           =   1515
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
         Left            =   6525
         TabIndex        =   39
         Top             =   255
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
         Left            =   4845
         TabIndex        =   31
         Top             =   255
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
         Left            =   5580
         TabIndex        =   20
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label11 
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
         TabIndex        =   19
         Top             =   255
         Width           =   645
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   18
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblqty 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Firm"
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
         Left            =   2250
         TabIndex        =   17
         Top             =   255
         Width           =   585
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
         Left            =   3675
         TabIndex        =   16
         Top             =   255
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
      Left            =   9765
      Picture         =   "frmGDN.frx":4974
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1020
   End
   Begin CoolButtonControl.CoolButton cmdBill 
      Height          =   1065
      Left            =   6390
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   15
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1879
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
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   15
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1879
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
      Left            =   780
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   105
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   635
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
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1050
      Left            =   3270
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   15
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1852
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
   Begin MSComctlLib.ListView lvwInvLines 
      Height          =   2295
      Left            =   75
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1065
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   4048
      SortKey         =   8
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
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
         Text            =   "Deposit"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Disc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ref"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblBlocked 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT BLOCKED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   825
      Left            =   3000
      TabIndex        =   55
      Top             =   180
      Width           =   6015
   End
   Begin VB.Label lblNonVATQuestion 
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
      Left            =   3555
      TabIndex        =   52
      Top             =   5370
      Width           =   315
   End
   Begin VB.Label lblNonVAT 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This invoice is a non-VAT invoice. All prices shown are VAT exclusive."
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   780
      TabIndex        =   51
      Top             =   5355
      Visible         =   0   'False
      Width           =   2790
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
      TabIndex        =   38
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
      TabIndex        =   37
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
      TabIndex        =   36
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   75
      Width           =   1920
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
      TabIndex        =   23
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
      Left            =   180
      TabIndex        =   22
      Top             =   135
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3465
      Picture         =   "frmGDN.frx":4CFE
      Stretch         =   -1  'True
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "frmGDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oGDN As a_GDN
Attribute oGDN.VB_VarHelpID = -1
Dim WithEvents oGDNLine As a_InvoiceLine
Attribute oGDNLine.VB_VarHelpID = -1
Dim oCustomer As a_Customer
Dim oProd As a_Product
Dim lngQtyQuickFound As Long

Dim bValidInvoice As Boolean
Dim bValidInvoiceLine As Boolean
Dim tlCustomer As z_TextList
Dim oCurrentCopy As a_Copy
Dim oCurrentForeignCurrency As a_Currency
Dim lngCurrentExtension As Long
Dim lngCurrentTotal As Long
Dim lngCurrentDepositTotal As Long
Dim lngCurrentVATTotal As Long

Dim lngSelectedRowIndex As String
Dim lngILEditingIdx As String
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
Public Sub component(Optional pCustID As Long, Optional pInvoice As a_GDN)
    On Error GoTo errHandler
Dim strAddress As String
    If pInvoice Is Nothing Then   'we create new invoice
'Handle objects
        Set oGDN = New a_GDN
        oGDN.BeginEdit
        
        If pCustID > 0 Then
            flgLoading = True
            LoadNewCustomer pCustID
            If Not oGDN.BillTOAddress Is Nothing Then
                strAddress = oGDN.BillTOAddress.AddressMailing
                lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
            End If
            If Not oGDN.DelToAddress Is Nothing Then
                strAddress = oGDN.DelToAddress.AddressMailing
                lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
            End If
            flgLoading = False
        End If
 '       oGDN.NonVATDocument = (oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False)
        oGDN.VATable = oGDN.Customer.VATable
        ChangeState enAddingRow
        
        
    Else   'we are provided with a loaded invoice
        Set oGDN = pInvoice
        oGDN.BeginEdit
        WaitMsg "Preparing to edit invoice  . . .", True, Me
        flgLoading = True
        If Not oGDN.BillTOAddress Is Nothing Then
            strAddress = oGDN.BillTOAddress.AddressMailing
            lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
'            If oGDN.billtoaddress.CountryID <> oPC.Configuration.LocalCountryID Then
'                chkChargeVAT.Enabled = True
'            Else
'                chkChargeVAT.Enabled = False
'            End If
        End If
        If Not oGDN.DelToAddress Is Nothing Then
            strAddress = oGDN.DelToAddress.AddressMailing
            lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
        End If
        If Not oGDN.BillTOAddress Is Nothing Then
            strAddress = oGDN.BillTOAddress.AddressMailing
            lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
        End If
        
        flgLoading = False
        oGDN.GetStatus
        oGDN.SetDirty False
        ChangeState enNotEditing
    End If
    oGDN.GetStatus
    lblBlocked.Visible = oGDN.Customer.Blocked
    cboRef.Visible = oPC.Configuration.SupportsWants
    SetMenu
    If oPC.AllowsSSInvoicing Then
        lblqty = "Firm"
        lblQtySS.Visible = True
        txtQtySS.Visible = True
        
        txtQty.Left = 2130
        txtQty.Width = 675
        lblqty.Left = 2130
        lblqty.Width = 855
        txtQtySS.Left = 2865
        txtQtySS.Width = 675
        lblQtySS.Left = 2900
        lblQtySS.Width = 500
        
        txtPrice.Left = 3555
        txtPrice.Width = 1000
        lblPrice.Left = 3600
        lblPrice.Width = 555
        txtRef.Left = 4590
        txtRef.Width = 1125
        lblRef.Left = 4590
        lblRef.Width = 555
        txtDiscount.Left = 5745
        txtDiscount.Width = 735
        lblDiscount.Left = 5545
        lblDiscount.Width = 1005
        txtNote.Left = 6510
        txtNote.Width = 2610
        lblNote.Left = 6610
        lblNote.Width = 1440
    Else
        lblqty.Caption = "Qty"
        lblQtySS.Visible = False
        txtQtySS.Visible = False
        txtQty.Left = 2155
        txtQty.Width = 615
        lblqty.Left = 2260
        lblqty.Width = 375
        txtPrice.Left = 2800
        txtPrice.Width = 1000
        lblPrice.Left = 2920
        lblPrice.Width = 555
        txtRef.Left = 3835
        txtRef.Width = 1125
        lblRef.Left = 4090
        lblRef.Width = 555
        txtDiscount.Left = 4990
        txtDiscount.Width = 735
        lblDiscount.Left = 4825
        lblDiscount.Width = 1005
        txtNote.Left = 5755
        txtNote.Width = 3215
        lblNote.Left = 5755
        lblNote.Width = 1440
    End If
    If oPC.AllowsInvoicePicking Then
        cmdCancel.Left = 6720
        cmdSave.Left = 7740
        cmdPick.Left = 8760
        cmdPick.Visible = True
        If oGDN.Status = stISSUED Then
            Me.cmdCancel.Enabled = False
            Me.cmdSave.Enabled = False
        End If
    Else
        cmdCancel.Left = 7740
        cmdSave.Left = 8760
        cmdPick.Visible = False
    End If

    If oPC.AllowsInvoicePicking Then
            Caption = IIf(oGDN.Status = stCOMPLETE, "Invoice for ", "Picking slip for ") & oGDN.Customer.NameAndCode(25) & oGDN.StaffNameB
    Else
        Caption = "Goods delivery note for " & oGDN.Customer.NameAndCode(25) & oGDN.StaffNameB
    End If

Exithandler:
    WaitMsg "", False, Me
    
        
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.component(pCustID,pInvoice)", Array(pCustID, pInvoice)
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim strAddress As String

    SetupcboMatch
    LoadComps
    
    If Not oGDN.IsNew Then
        chkChargeVAT = IIf(oGDN.ShowVAT, 1, 0)
    Else
        chkChargeVAT = IIf(oPC.Configuration.DiscountVATDefault, 1, 0)
    End If
    
 '   SetLvw
    LoadCustomerDetailsToForm
    LoadListView
    
    lblNonVATQuestion.Visible = (oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False)
    lblNonVAT.Visible = (oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False)
    
    oGDN.SetDirty False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadControls"
End Sub

Private Sub cmdCancelMatch_Click()
    On Error GoTo errHandler
Dim i As Integer

    If cboRef.Items.ItemCount = 0 Then Exit Sub
    oGDNLine.COLID = 0
    For i = 0 To cboRef.Items.ItemCount - 1
        cboRef.Items.SelectItem(cboRef.Items(i)) = False
    Next
    mSetfocus cmdEnter

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdCancelMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdConvertToNonVat_Click()
    On Error GoTo errHandler
    oGDN.ConvertToNonVATGDN
    LoadControls

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdConvertToNonVat_Click", , EA_NORERAISE
    HandleError
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
    ErrorIn "frmGDN.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPick_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoGDNLines As Boolean
Dim iCurrency As Integer
Dim strResult As String
Dim frm As frmGDNPreview
Dim frmDte As frmTRDate
    If (oPC.AllowsInvoicePicking = False And oGDN.Status = stInProcess) Or oPC.AllowsInvoicePicking = True Then
    
            WaitMsg "Picking GDN  . . .", True, Me
            oGDN.VATable = oGDN.Customer.VATable
            oGDN.StaffID = gSTAFFID
            oGDN.RecalculateAllLines
            oGDN.CalculateTotals
            
            strResult = oGDN.Post(stISSUED)
            
            If strResult = "" Then
                Set frm = New frmGDNPreview
                frm.ComponentObject oGDN
                frm.Show
            ElseIf strResult > "" Then
                MsgBox "The GDN cannot be issued now, try later. The record is probably locked by another user. The message is: " & strResult & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                oGDN.BeginEdit
                WaitMsg "", False, Me
                Exit Sub
            End If
    End If
EXITH:
    WaitMsg "", False, Me
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdPick_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    cmdAppro.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oGDN.StatusF = "IN PROCESS" And oGDN.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oGDN.StatusF = "ISSUED")      ' And oGDN.CanCancel = True
    Forms(0).mnuDelLine.Enabled = (oGDN.Status = stInProcess)   'This should not happen if the GDN is in Picking state ONly delete from an in process GDN
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.SetMenu"
End Sub



Private Sub cbComp_Click()
    On Error GoTo errHandler
    oGDN.COMPID = OptionLoop(oGDN.COMPID, oPC.Configuration.Companies.Count)
    cbComp.Caption = oPC.Configuration.Companies(oGDN.COMPID).CompanyName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cbComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oGDN.Customer.ID > 0 Then
        frm.component oGDN.Customer
        frm.Show
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cboRef_SelectionChanged()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oGDNLine.COLID = cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 3)
    oGDNLine.SetQty cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 2)
    If oGDN.Customer.UseQuotedPrice Then
        oGDNLine.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4)
        oGDNLine.SetPrice cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6)
    Else
        If cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4) <> oGDN.Customer.DefaultDiscount Then
            DoEvents
            If MsgBox("The customer discount is different than the discount on the customer order. Use the customer discount?", vbYesNo, "Warning") = vbYes Then
                oGDNLine.SetDiscountPercent oGDN.Customer.DefaultDiscount
            Else
                oGDNLine.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4)
            End If
        End If
        If cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6) <> oProd.SP Then
            If MsgBox("The product price is different than the price on the customer order. Use the product price?", vbYesNo, "Warning") = vbYes Then
                oGDNLine.SetPrice oProd.SP
            Else
                oGDNLine.SetPrice cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6)
            End If
        End If
    End If
    
    oGDNLine.SetRef cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 0)
    
    txtDiscount = oGDNLine.DiscountPercent
    txtPrice = oGDNLine.Price
    txtRef = oGDNLine.Ref
    If oPC.AllowsSSInvoicing Then
        Me.txtQty = oGDNLine.QtyFirm
        Me.txtQtySS = oGDNLine.QtySS
    Else
        txtQty = oGDNLine.Qty
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cboRef_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkChargeVAT_Click()
    On Error GoTo errHandler
    oGDN.ShowVAT = (chkChargeVAT = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.chkChargeVAT_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdBill_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oGDN.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oGDN.Customer.Addresses.Count Then
        i = 1
    End If
    lblAddBill.Caption = oGDN.Customer.Addresses(i).AddressMailing & vbCrLf & oGDN.Customer.Addresses(i).EMail
    oGDN.SetBillToAddress oGDN.Customer.Addresses(i)
    iBillIdx = i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdBill_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oGDN.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oGDN.Customer.Addresses.Count Then
        i = 1
    End If
    lblAddDel.Caption = oGDN.Customer.Addresses(i).AddressMailing & vbCrLf & oGDN.Customer.Addresses(i).EMail
    oGDN.setDelToAddress oGDN.Customer.Addresses(i)
    iBillIdx = i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdDel_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oGDN.SetCustomer(plngTPID) Then
        vCanAdd.RuleBroken "TP", False
        LoadCustomerDetailsToForm
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadNewCustomer(plngTPID)", plngTPID
End Sub
Private Sub LoadCustomerDetailsToForm()
    On Error GoTo errHandler
    With oGDN.Customer
        If Not .BillTOAddress Is Nothing Then
            lblTPPhone.Caption = .BillTOAddress.Phone
            lblTPFax.Caption = .BillTOAddress.Fax
        End If
        lblTPName.Caption = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
        If Not .BillTOAddress Is Nothing Then
            If oGDN.BillTOAddress Is Nothing Then
                oGDN.SetBillToAddress .BillTOAddress
                lblAddBill.Caption = .BillTOAddress.AddressShort
            End If
        End If
        If Not .DelToAddress Is Nothing Then
            If oGDN.DelToAddress Is Nothing Then
                oGDN.setDelToAddress .DelToAddress
                lblAddDel.Caption = .DelToAddress.AddressShort
            End If
        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadCustomerDetailsToForm"
End Sub
Private Sub cmdNote()
    On Error GoTo errHandler
Dim frm As New frmILNote
    frm.component oGDNLine
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdNote"
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set vCanAdd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAddresses()
    On Error GoTo errHandler
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oGDN
    frm.Show vbModal
    lblAddBill.Caption = oGDN.BillTOAddress.AddressShort
    lblAddDel.Caption = oGDN.DelToAddress.AddressShort
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuAddresses"
End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveInvoiceLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuDelLine"
End Sub

Private Sub lblNonVATQuestion_Click()
    On Error GoTo errHandler
Dim s As String
    s = "This message is shown because the 'Show VAT' check box has not been ticked in the customer record." _
    & " This is only applicable to customers who do not pay local VAT." & vbCrLf _
    & " The GDN will calculate values based on the ex VAT price and will make no reference to VAT on the printed document."
    MsgBox s, , "Non-VAT pricing"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.lblNonVATQuestion_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oGDN_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidInvoice = (pMsg = "")
  '  cmdIssue.Enabled = (bValidInvoice And oGDN.GDNLines.Count > 0 And vMode = enNotEditing)
  '  cmdPick.Enabled = (bValidInvoice And oGDN.GDNLines.Count > 0 And vMode = enNotEditing)
  '  cmdSave.Enabled = (bValidInvoice And oGDN.GDNLines.Count > 0 And vMode = enNotEditing)
    
    cmdIssue.Enabled = (bValidInvoice And oGDN.GDNLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidInvoice) And oGDN.IsDirty
    cmdCancel.Enabled = True
    
    lblNonVATQuestion.Visible = (oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False)
    lblNonVAT.Visible = (oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False)
    
    If oGDN.Status = stISSUED Then
        Me.cmdCancel.Enabled = False
        Me.cmdSave.Enabled = False
    End If
    Me.txtError = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDN_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oGDNLine_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDNLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oGDNLine_Valid(msg As String)
    On Error GoTo errHandler
        Me.cmdEnter.Enabled = (msg = "")
        Me.txtError = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDNLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oGDN_TotalChange(lngTotalExt As Long, lngTotalDeposit As Long, lngTotalVAT As Long)
    On Error GoTo errHandler
    
    flgLoading = True
    
    lngCurrentTotal = lngTotalExt
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
  '  If vMode = enEditingRow Then
  '      cmdNewRows.Enabled = (oGDN.GDNLines.Count > 0)
  '  End If

    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDN_TotalChange(lngTotalExt,lngTotalDeposit,lngTotalVAT)", _
         Array(lngTotalExt, lngTotalDeposit, lngTotalVAT), EA_NORERAISE
    HandleError
End Sub

Private Sub oGDN_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDN_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oGDN_Dirty(pVal As Boolean)
    On Error GoTo errHandler
    If flgLoading And pVal Then Exit Sub
    If pVal = True Then
        cmdSave.Enabled = pVal And vMode = enNotEditing And oPC.AllowsInvoicePicking = False
        cmdCancel.Caption = "&Cancel"
    Else
        cmdSave.Enabled = False
        cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDN_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oGDN_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDN_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub


Private Sub SetFocusFromCode()
    On Error GoTo errHandler
Dim strMsg As String
    
    
    If LenB(txtCode) > 0 Then
        If oGDNLine.ServiceItem = False Then
            If (oPC.Configuration.AntiquarianYN) And (Not oGDNLine.Product.DefaultCopy Is Nothing) Then
                mSetfocus txtPrice
            ElseIf cboRef.Visible = False Then
                mSetfocus txtQty
            Else
                mSetfocus cboRef
            End If
        Else
            txtQty = ""
            txtQtySS = ""
            
            txtPrice.Enabled = True
            mSetfocus txtPrice
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.SetFocusFromCode"
End Sub

Private Sub txtCode_DblClick()
    On Error GoTo errHandler
   ' cmdFind_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtCode_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
    txtNote = HandleTextWithBites(txtNote)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_DblClick()
    On Error GoTo errHandler
    If txtNote.Height = 1125 Then
        txtNote.Height = 285
    Else
        txtNote.Height = 1125
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtNote_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_DblClick()
    On Error GoTo errHandler
Dim f As New frmFCPrice
Dim oGDNL As a_InvoiceLine
Dim x As Long
Dim Y As Long
    
    If Not oPC.SupportsUNISA Then Exit Sub
    
    f.component Me.Left + 2000, Me.TOP + 1400, oGDNLine.ForeignPrice, oGDNLine.Price, oGDNLine.FCID, oGDNLine.VATRate, oGDNLine.FCFactor

    f.Show vbModal
    If f.UserCancelled Then
        Unload f
        Exit Sub
    End If
    oGDNLine.SetForeignPrice CStr(f.ForeignPrice)
    oGDNLine.FCID = f.FCID
    If Round(f.FCFactor, 6) <> Round(oGDNLine.FCFactor, 6) Then
        For Each oGDNL In oGDN.GDNLines
            If oGDNL.FCID = oGDNLine.FCID Then
                oGDNL.BeginEdit
                oGDNL.SetFCFactor Round(f.FCFactor, 6)
                oGDNL.ApplyEdit
            End If
        Next
    End If
    oGDNLine.Price = f.LocalPriceIncVAT
    Me.txtPrice = oGDNLine.Price
    lblFCTerms.Caption = oGDNLine.ForeignPriceF & "/" & f.FCFactorINV
    Unload f

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtPrice_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_DblClick()
    On Error GoTo errHandler
    If Not oPC.SupportsUNISA Then Exit Sub
    txtQty = oGDN.TotalQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtQty_DblClick", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtQty_GotFocus()
'    On Error GoTo errHandler
'    AutoSelect txtQty
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmGDN.txtQty_GotFocus", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If vMode = enNotEditing Then Exit Sub
    If oPC.AllowsSSInvoicing Then
        If Not oGDNLine.SetQtyFirm(txtQty) Then
            Cancel = True
        End If
    Else
        If Not oGDNLine.SetQty(txtQty) Then
            Cancel = True
        End If
    End If
    oGDNLine.CalculateLine
    txtTotal = oGDNLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtRef_GotFocus", , EA_NORERAISE
    HandleError
End Sub
'============
Private Sub txtQtySS_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtySS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtQtySS_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySs_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If vMode = enNotEditing Then Exit Sub
    If Not oGDNLine.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    oGDNLine.CalculateLine
    txtTotal = oGDNLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


'========
Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdPick.Enabled = True
    Me.cmdSave.Enabled = True
    Me.cmdIssue.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim curTotalDeposit As Currency
Dim strAddress As String
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11100
        Height = 6700
    End If
    LoadControls
    
    RefreshFCExchangeRates
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub RefreshFCExchangeRates()
    On Error GoTo errHandler
 Dim oGDNL As a_InvoiceLine
 
    If Not oPC.SupportsUNISA Then Exit Sub
 
    For Each oGDNL In oGDN.GDNLines
        If oGDNL.FCID > 0 Then
            oGDNL.BeginEdit
            oGDNL.SetFCFactor Round(oPC.Configuration.Currencies.FindCurrencyByID(oGDNL.FCID).Factor, 6)
            oGDNL.ApplyEdit
        End If
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.RefreshFCExchangeRates"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
            'LogSaveToFile "Invoice form unloaded"
    If Not oCurrentCopy Is Nothing Then
        If oCurrentCopy.IsEditing Then oCurrentCopy.CancelEdit
    End If
    If oGDN.IsEditing Then oGDN.CancelEdit
    UnsetMenu
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oGDN = Nothing
    Set tlCustomer = Nothing
    Set oGDNLine = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
Dim lngMax As Long

    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        mSetfocus txtCode
        Exit Sub
    End If
    If oGDNLine.ServiceItem Then oGDNLine.DiscountPercent = 0
    If oPC.GetProperty("WarnOverInvoicing") = "TRUE" Then
        If oProd.QtyOnHand - (oGDN.TotalQtyPerProduct(oGDNLine.PID) + IIf(oGDNLine.IsNew, oGDNLine.Qty, 0)) < 0 And Not oProd.IsServiceItem Then
            MsgBox "There are not enough items in stock to GDN this quantity." & vbCrLf & "The maximum you can GDN is " _
            & CStr(GetMax(0, oProd.QtyOnHand)), vbInformation, "Warning"
        End If
    End If
    If oPC.GetProperty("StopOverInvoicing") = "TRUE" Then
        If oProd.QtyOnHand - (oGDN.TotalQtyPerProduct(oGDNLine.PID) + IIf(oGDNLine.IsNew, oGDNLine.Qty, 0)) < 0 And Not oProd.IsServiceItem Then
            MsgBox "There are not enough items in stock to GDN this quantity." & vbCrLf & "The maximum you can GDN is " _
            & CStr(GetMax(0, oProd.QtyOnHand)), vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
    oGDNLine.ApplyEdit
    oGDNLine.BeginEdit
    Dim x As ListItem
    If vMode = enAddingRow Then
        'LogSaveToFile "GDN line added:" & oGDNLine.Key
        For i = 1 To lvwInvLines.ListItems.Count
            strItemsDebug = strItemsDebug & "," & lvwInvLines.ListItems(i).Key
        Next
       ' On Error Resume Next
        If lvwInvLines.ListItems.Count < val(oGDNLine.Key) Then
            lvwInvLines.ListItems.Add Key:=oGDNLine.Key
            LoadListViewLine lvwInvLines.ListItems(lvwInvLines.ListItems.Count), oGDNLine
        End If
        
        lvwInvLines.Refresh
        ChangeState enAddingRow
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwInvLines.ListItems(lngSelectedRowIndex), oGDNLine
        ChangeState enNotEditing
    End If
    oGDN.GetStatus
    cboRef.Items.RemoveAllItems

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdEnter_Click", , EA_NORERAISE
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
        cboRef.Visible = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdPick.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oGDN.GDNLines.Count > 0)
        lvwInvLines.Enabled = False
        lvwInvLines.Height = 2200
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
        cmdPick.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oGDN.GDNLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        lblTPPhone.Caption = ""
        lvwInvLines.Enabled = False
        lvwInvLines.Height = 2200
        ClearInvLineControls
        fr1.ZOrder 1
        mSetfocus txtCode
        Set oGDNLine = oGDN.GDNLines.Add
        oGDNLine.SetParentInvoice oGDN
        oGDNLine.InvoiceID = oGDN.GDNID
        oGDNLine.SetQty 1
        UnsetMenu
    Case enNotEditing
        flgLoading = True
        fr1.Visible = False
        txtError = ""
        txtRef = ""
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True
        cmdPick.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = True  '(oGDN.GDNLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        lvwInvLines.Enabled = True
        lvwInvLines.Height = 4000
        '''''''
'        If Not oGDNLine Is Nothing Then oGDNLine.CancelEdit
'        oGDN.GDNLines.CancelEdit
'        oGDN.GDNLines.BeginEdit
        
        SetMenu
        fr1.ZOrder 1
    End Select
    If Not oGDN.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
    If oGDN.Status = stISSUED Then
        Me.cmdCancel.Enabled = False
        Me.cmdSave.Enabled = False
    End If
    
    lblAppro.Caption = ""
    cboRef.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.ChangeState(pToMode)", pToMode
End Sub
Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
        'LogSaveToFile "Invoice New row button:enEditingRow"
        If oGDNLine.IsEditing Then
            oGDNLine.CancelEdit
            oGDNLine.BeginEdit
        End If
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        If txtCode > "" Then  'THis is not after a post but is an aborted  add row action
           oGDN.GDNLines.DecrementMaxKeyUsed
        End If
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        'LogSaveToFile "Invoice New row button:enNotEditing"
        ChangeState enAddingRow
    End If

    ClearInvLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
Dim strItemsDebug As String

    For i = 1 To lvwInvLines.ColumnHeaders.Count
        lvwInvLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwInvLines.ColumnHeaders(i).Width)
    Next
    lvwInvLines.ListItems.Clear
    For i = 1 To oGDN.GDNLines.Count
        Set lstItem = lvwInvLines.ListItems.Add
        LoadListViewLine lstItem, oGDN.GDNLines(i)
        strItemsDebug = strItemsDebug & "," & lvwInvLines.ListItems(i).Key
    Next i
    Debug.Print strItemsDebug
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oGDNLine As a_InvoiceLine)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oGDNLine
        lstItem.text = .CodeForEditing
        lstItem.Key = .Key
        lstItem.SubItems(1) = .TitleAuthorPublisher
        If oPC.AllowsSSInvoicing Then
            If .ServiceItem Then
                lstItem.SubItems(2) = ""
            Else
                lstItem.SubItems(2) = .QtyFirmF & "/" & .QtySSF
            End If
        Else
            lstItem.SubItems(2) = .Qty
        End If
        If .Deposit <> 0 Then
            lstItem.SubItems(3) = .DepositF(False)
        Else
            lstItem.SubItems(3) = " "
        End If
        lstItem.SubItems(4) = .PriceF(False)
        lstItem.SubItems(5) = .DiscountPercentF  ' Format(.DiscountPercent, "##0.0%")
        lstItem.SubItems(6) = .Ref
        lstItem.SubItems(7) = .PAfterDiscountExtF(False)
        lstItem.SubItems(8) = Format(.Key, "@@@@@@@@@@")
        lstItem.SubItems(9) = .EAN
        If .ServiceItem = True Then
            lstItem.ForeColor = &H427182
            lstItem.ListSubItems(1).ForeColor = &H427182
            lstItem.ListSubItems(2).ForeColor = &H427182
            lstItem.ListSubItems(3).ForeColor = &H427182
            lstItem.ListSubItems(4).ForeColor = &H427182
            lstItem.ListSubItems(5).ForeColor = &H427182
            lstItem.ListSubItems(6).ForeColor = &H427182
            lstItem.ListSubItems(7).ForeColor = &H427182
        ElseIf .PIID = 0 Then
            lstItem.ListSubItems(1).ForeColor = &H706034
            lstItem.ListSubItems(2).ForeColor = &H706034
            lstItem.ListSubItems(3).ForeColor = &H706034
            lstItem.ListSubItems(4).ForeColor = &H706034
            lstItem.ListSubItems(5).ForeColor = &H706034
            lstItem.ListSubItems(6).ForeColor = &H706034
            lstItem.ListSubItems(7).ForeColor = &H706034
        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadListViewLine(lstItem,oGDNLine)", Array(lstItem, oGDNLine)
End Sub
Private Sub lvwInvLines_DblClick()
    On Error GoTo errHandler
    Dim strPos As String
    
'This must load the editing line with the current line's data
    If lvwInvLines.ListItems.Count = 0 Then Exit Sub
    If lvwInvLines.SelectedItem.Index < 1 Then Exit Sub
    
    lngILEditingIdx = lvwInvLines.SelectedItem.Key
    Set oGDNLine = Nothing
    Set oGDNLine = oGDN.GDNLines(lngILEditingIdx)
    
    lngSelectedRowIndex = lvwInvLines.SelectedItem.Key

    ChangeState eneditingrow
    Set oProd = Nothing
    Set oProd = New a_Product
    oProd.Load oGDNLine.PID, 0
  '  If oGDNLine.COLID > 0 Then
        LoadandSHowcboRef "", oGDNLine.COLID
            If Me.cboRef.Items.ItemCount > 0 And oGDNLine.COLID > 0 Then
                flgLoading = True
                cboRef.Items.SelectItem(cboRef.Items.FindItem(oGDNLine.COLID, 3)) = True
                oGDNLine.COLID = cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 3)
                flgLoading = False
            Else
                cboRef.DropDown() = True
            End If
    HandlePossibleApprosOS oGDNLine.PID
    
'Set up screen values
    txtDiscount = oGDNLine.DiscountPercentF
    If oPC.AllowsSSInvoicing Then
        txtQty = oGDNLine.QtyFirmF
        txtQtySS = oGDNLine.QtySSF
    Else
        txtQty = oGDNLine.QtyF
    End If
strPos = "position 7"
    txtNote = oGDNLine.Note
    txtRef = oGDNLine.Ref
    txtDiscount = oGDNLine.DiscountPercent
        txtCode = oGDNLine.CodeForEditing
    txtTitle = oGDNLine.Title
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oGDNLine.PriceF(False)
    Else
        txtPrice = oGDNLine.Price
    End If
    lblAppro.Caption = oGDNLine.APPLQTY
    If oGDNLine.Qty > 1 Then
        mSetfocus txtQty
    Else
        mSetfocus txtPrice
    End If
strPos = "position 8"
    If oGDNLine.FCID <> oPC.Configuration.DefaultCurrencyID And oGDNLine.FCID > 0 Then
        Me.lblFCTerms = oGDNLine.ForeignPriceF & "/" & oGDNLine.FCFactorInvF
    End If
    
    oGDNLine.GetStatus
strPos = "position 9"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.lvwInvLines_DblClick", , EA_NORERAISE
    HandleError
End Sub

'---------Companies code
Private Sub LoadComps()
    On Error GoTo errHandler
Dim oComp As a_Company
Dim oItem As ListItem
Dim i As Integer
    If oGDN.COMPID > 0 Then
        cbComp.Caption = oPC.Configuration.Companies(CStr(oGDN.COMPID)).CompanyName
    Else
        cbComp.Caption = oPC.Configuration.DefaultCompany.CompanyName
        oGDN.COMPID = oPC.Configuration.DefaultCOMPID
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadComps"
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If oGDN.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'-------End Compsny code
'Private Sub txtNote_Change()
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oGDNLine.setnote (txtNote)
'    If Err Then
'      Beep
'      intPos = txtNote.SelStart
'      txtNote = oGDNLine.Note
'      txtNote.SelStart = intPos - 1
'    End If
'End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oGDNLine.SetNote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oGDNLine.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oGDN.Memo
    ofrm.Show vbModal
    oGDN.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuMemo"
End Sub

Public Sub mnuCancel()
    On Error GoTo errHandler
    If oGDN.IsDirty Then
        oGDN.CancelEdit
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuCancel"
End Sub






Public Sub mnuVoid()
    On Error GoTo errHandler
    oGDN.SetStatus stVOID
    oGDN.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuVoid"
End Sub


Private Sub txtCode_LostFocus()
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
        'Cancel = True
        txtCode.SetFocus
        GoTo EXIT_Handler
    End If
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
        .Load 0, 0, FNS(txtCode)
        If oProd.PID > "" Then
            Set rsPreviousBillings = oSM.PreviousBillings(oProd.PID, oGDN.Customer.ID)
            'lngResult = oSQL.RunGetRecordset("SELECT TR_CODE,TR_DATE FROM tTR JOIN tGDN ON TR_ID = I_ID JOIN tILINE ON IL_TR_ID = TR_ID WHERE IL_P_ID = '" & oProd.pID & "' AND TR_TP_ID = " & oGDN.Customer.ID, enText, Array(), "", rsPreviousBillings)
            If Not rsPreviousBillings Is Nothing Then
                If rsPreviousBillings.State <> 0 Then
                    If Not rsPreviousBillings.eof Then
                        strPreviousBillings = ""
                        Do While Not rsPreviousBillings.eof
                            strPreviousBillings = rsPreviousBillings.Fields(0) & "   " & Format(rsPreviousBillings.Fields(1), "dd/mm/yyyy") & vbCrLf
                            rsPreviousBillings.MoveNext
                        Loop
                        If strPreviousBillings > "" And oProd.IsServiceItem = False Then
                            MsgBox "This item is on a previous GDN to this client." & vbCrLf & strPreviousBillings, vbInformation, "Warning"
                        End If
                    End If
                End If
            End If
        End If
        'Check to see if copy sold
        If Not oProd.DefaultCopy Is Nothing Then 'Book in database and copy requested
            If oProd.DefaultCopy.SoldDate > CDate(0) Then  'Copy is sold
                MsgBox "Copy already sold", vbInformation, "Check"
               ' Cancel = True
               ' Exit Sub
                txtCode.SetFocus
                GoTo EXIT_Handler
               
            End If
        ElseIf Not oProd.IsServiceItem And InStr(txtCode, "/") > 0 Then
            ' we may reach here is a copy is requested and not found
            MsgBox "No such copy exists", vbInformation, "Check"
    '        Cancel = True
    '        Exit Sub
        txtCode.SetFocus
        GoTo EXIT_Handler
    
        End If
        If Len(FNS(.PID)) <> 0 Then   'Book in database
            If Not oProd.DefaultCopy Is Nothing Then  'Copy requested and identified
                Set oCurrentCopy = oProd.DefaultCopy
                oGDNLine.Price = oCurrentCopy.Price
                oGDNLine.PIID = oCurrentCopy.ID
                oGDNLine.DiscountPercent = oGDN.Customer.DefaultDiscount
            ElseIf oProd.IsServiceItem Then   'No copy identified but product is a non-stock product (e.g. postage or insurance etc.)
               ' mSetfocus txtPrice
               ' AutoSelect txtPrice
               ' oGDNLine.Qty = 1
                oGDNLine.Qty = 1
                oGDNLine.Price = oProd.SPex((oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False))
                oGDNLine.CodeForExport = oProd.CodeForExport
                oGDNLine.CodeF = oProd.CodeF
                oGDNLine.code = oProd.EAN
            Else    ' we may reach here is a copy is requested and not found
                    ' OR No copy is requested and the Title is found
                oGDNLine.Price = oProd.SPex((oGDN.Customer.VATable = False And oGDN.Customer.ShowVAT = False))
                oGDNLine.CodeF = oProd.CodeF
               ' oGDNLine.EAN oProd.EAN
                oGDNLine.code = oProd.code
                oGDNLine.CodeForExport = oProd.CodeForExport
                If oPC.Configuration.AllowCopyInfo And InStr(txtCode, "/") > 0 Then
                    If MsgBox("There is no copy with this serial number" & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Papyrus Invoicing Information") = vbNo Then
                        txtCode.SetFocus
                        GoTo EXIT_Handler
                    End If
                End If
            End If
strPos = "Pos 4"
            LoadandSHowcboRef .PID
strPos = "Pos 5"
            HandlePossibleApprosOS .PID
strPos = "Pos 6"
            oGDNLine.Title = .TitleAuthor  'L(35)
            oGDNLine.PID = .PID
            oGDNLine.ServiceItem = .IsServiceItem
            oGDNLine.VATRate = .VATRateToUse
            oGDNLine.Cost = .Cost
            If oGDNLine.IsNew And oGDNLine.DiscountPercent = 0 Then
                oGDNLine.DiscountPercent = oGDN.Customer.DefaultDiscount
            End If
            If oGDNLine.DiscountPercent <> oGDN.Customer.DefaultDiscount And oGDNLine.IsNew = False Then
                If MsgBox("The discount on the GDN differs from the customer's usual discount. " & vbCrLf & "Use discount on order?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                    oGDNLine.DiscountPercent = oGDN.Customer.DefaultDiscount
                End If
            End If
        Else   'Book nof found on database
            MsgBox "Cannot find book", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
       '     Cancel = True
       '     Exit Sub
        txtCode.SetFocus
        GoTo EXIT_Handler
       
        End If
       ' oGDNLine.ServiceItem = .IsServiceItem
        
        If Not .DefaultCopy Is Nothing Then
            oGDNLine.CodeF = .code & .DefaultCopy.SerialF
        End If
    End With
strPos = "Pos 6"
    txtTitle = oGDNLine.TitleAuthor
 '   If oPC.Configuration.CaptureDecimal Then
 '       txtPrice = oGDNLine.PriceF(False)
 '   Else
        txtPrice = oGDNLine.Price
 '   End If
        txtQty = oGDNLine.Qty
    txtRef = oGDNLine.Ref
    txtDiscount = oGDNLine.DiscountPercentF
    oGDNLine.GetStatus
        If cboRef.Items.ItemCount = 0 Then
            SetFocusFromCode
        Else
            AutoSelect txtQty
        End If
EXIT_Handler:
   ' Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not oGDNLine.SetDiscountPercent(txtDiscount) Then
        Cancel = True
    End If
    oGDNLine.CalculateLine
    txtTotal = oGDNLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oGDNLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oGDNLine.CalculateLine
    txtTotal = oGDNLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oGDNLine.SetRef (txtRef)
    If Err Then
      Beep
      intPos = txtRef.SelStart
      txtRef = oGDNLine.Ref
      txtRef.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtRef_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oGDNLine.SetRef(txtRef)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtRef = oGDNLine.Ref
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.txtRef_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub RemoveInvoiceLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwInvLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwInvLines.ListItems(i).Selected Then
            oGDN.GDNLines.Remove lvwInvLines.ListItems(i).Key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Exit Sub
    End If
    lvwInvLines.ListItems.Remove i
    lvwInvLines.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.RemoveInvoiceLine"
End Sub

Private Sub SaveGDN()
    On Error GoTo errHandler
Dim strErrPos As String

strErrPos = "Pos 1"
If oGDN Is Nothing Then
        'LogSaveToFile "SaveGDN: oGDN is nothing"
End If
    oGDN.ApplyEdit
strErrPos = "Pos 2"
If oGDN Is Nothing Then
        'LogSaveToFile "SaveGDN: oGDN is nothing"
End If
    oGDN.BeginEdit
strErrPos = "Pos 3"
If oGDN.GDNLines Is Nothing Then
        'LogSaveToFile "SaveGDN: oGDN.GDNLines is nothing"
End If
    Set oGDNLine = oGDN.GDNLines.Add
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.SaveGDN"
End Sub

Public Sub PrintGDN()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoGDNLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oGDN.Load oGDN.GDNID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoGDNLines Then
        MsgBox "There are no records to print on this GDN.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
        GoTo EXIT_Handler
    End If
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.PrintGDN"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoGDNLines As Boolean
Dim iCurrency As Integer
Dim strResult As String
Dim frm As frmGDNPreview
Dim frmDte As frmTRDate
Dim strOverInvoicedSet As String
Dim oSM As New z_StockManager
    If oGDN.Customer.Blocked Then
        MsgBox "This customer is blocked. You cannot issue this GDN.", vbCritical + vbOKOnly, "Can't do this"
        Exit Sub
    End If
        If oGDN.QtyNonStandardVAT > 0 Then
            If MsgBox("There are items with non-standard VAT in this GDN, continue?", vbYesNo + vbInformation, "Warning") = vbNo Then
                Exit Sub
            End If
        End If
    If (oPC.AllowsInvoicePicking = False And oGDN.Status = stInProcess) Or oPC.AllowsInvoicePicking = True Then
            If oPC.Configuration.SignTransactions = True Then
                If SecurityControl(enSECURITY_INV_SIGN, , "Sign this GDN.", DOCAPPROVAL) = False Then
                       Exit Sub
                End If
            End If
            
                If oPC.GetProperty("WarnOverInvoicing") = "TRUE" Then
                    strOverInvoicedSet = oSM.GetOverInvoicedIems(oGDN.GDNID)
                    If strOverInvoicedSet <> "" Then
                        If MsgBox("You are over-invoicing the following items." & vbCrLf & strOverInvoicedSet & vbCrLf & "Do you want to continue?", vbInformation + vbOKCancel, "Warning") = vbCancel Then
                            GoTo Redisplay
                        End If
                    End If
                End If
                If oPC.GetProperty("StopOverInvoicing") = "TRUE" Then
                    strOverInvoicedSet = oSM.GetOverInvoicedIems(oGDN.GDNID)
                    If strOverInvoicedSet <> "" Then
                        MsgBox "You are over-delivering the following items. You cannot issue this GDN until you have corrected it." & vbCrLf & strOverInvoicedSet, vbCritical, "Can't do this"
                        GoTo Redisplay
                    End If
                End If
            If oPC.AllowInvoiceDateOverride Then
                Set frmDte = New frmTRDate
                frmDte.component Date
                frmDte.Show vbModal
                oGDN.DOCDate = StartOfDay(frmDte.InvoiceDate)
                Unload frmDte
                oGDN.CaptureDate = Now()
            Else
                If oGDN.DOCDate < CDate("1950-01-01") Then
                    oGDN.DOCDate = Date
                    oGDN.CaptureDate = Now()
                End If
            End If
            
            WaitMsg "Issuing Goods delivery note  . . .", True, Me
            oGDN.VATable = oGDN.Customer.VATable
            oGDN.StaffID = gSTAFFID
            oGDN.RecalculateAllLines
            oGDN.CalculateTotals
            
            If Not oGDN.IsEditing Then oGDN.BeginEdit
            strResult = oGDN.Post(stCOMPLETE)
            
Redisplay:
            If strResult = "" Or strResult = "In Process" Then
                Set frm = New frmGDNPreview
                frm.ComponentObject oGDN
                frm.Show
            ElseIf strResult > "" Then
                MsgBox "The GDN cannot be issued now, try later. The record is probably locked by another user. The message is: " & strResult & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                If Not oGDN.IsEditing Then oGDN.BeginEdit
                WaitMsg "", False, Me
                Exit Sub
            End If
    End If
EXITH:
    WaitMsg "", False, Me
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oGDN_DirtyStatus(pDirty As Boolean)
    On Error GoTo errHandler
    If pDirty = True Then
        cmdCancel.Caption = "&Cancel"
    Else
        cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.oGDN_DirtyStatus(pDirty)", pDirty, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oGDNL As a_InvoiceLine
    oGDN.SetStatus stInProcess
    If oGDN.DOCDate < CDate("1950-01-01") Then
        oGDN.DOCDate = Date
        oGDN.CaptureDate = Now()
    End If
    oGDN.RecalculateAllLines
    oGDN.CalculateTotals
        'LogSaveToFile "Invoice Saving button"
    SaveGDN
    cmdSave.Enabled = False
    cmdCancel.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmGDNPreview

    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmGDNPreview
        frm.ComponentObject oGDN
        frm.Show
    End If
    If cmdCancel.Caption <> "&Close" Then
        If oGDN.IsEditing And oGDN.IsDirty Then
            If MsgBox("You wish to cancel your changes?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
            oGDN.CancelEdit
        End If
    End If
    If Not oGDNLine Is Nothing Then
        If oGDNLine.IsEditing Then oGDNLine.CancelEdit
    End If
        'LogSaveToFile "GDN Cancel button"
    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdCancel_Click", , EA_NORERAISE
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
    txtQty = ""
    cmdAppro.BackColor = &HC4BCA4
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.ClearInvLineControls"
End Sub

Private Sub lvwInvLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.lvwInvLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oGDN.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oGDN.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.SetIssueButtonCaption"
End Sub

Private Sub lvwInvLines_Click()
    On Error GoTo errHandler
    If lvwInvLines Is Nothing Then Exit Sub
    If lvwInvLines.SelectedItem Is Nothing Then Exit Sub
    
    If lvwInvLines.SelectedItem.Index > 0 And Left(lvwInvLines.SelectedItem.SubItems(9), ISBNLENGTH) > "" Then
    
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText Left(lvwInvLines.SelectedItem.SubItems(9), ISBNLENGTH)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.lvwInvLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwInvLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvwInvLines.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
    If lvwInvLines.SortOrder = lvwAscending Then
        lvwInvLines.SortOrder = lvwDescending
    Else
        lvwInvLines.SortOrder = lvwAscending
    End If
   lvwInvLines.Sorted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.lvwInvLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim Style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwInvLines.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   Style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   Style = Style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If Style Then
      Call SetWindowLong(hHeader, GWL_STYLE, Style)
      Call SetWindowPos(lvwInvLines.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.SetLvw"
End Sub

Private Sub vCanAdd_Status(errors As String)
    On Error GoTo errHandler
MsgBox errors & "CANAADD"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.vCanAdd_Status(errors)", errors, EA_NORERAISE
    HandleError
End Sub

'Private Sub ReconcileWithCOs()
'Dim frm As frmCOFF
'Dim oInv As a_GDN
'    Set oInv = New a_GDN
'    oInv.Load oGDN.InvoiceID, True
'    If oInv.hasCoffs Then
'        Set frm = New frmCOFF
'        frm.Component oGDN
'        frm.Show vbModal
'    End If
'    Set oInv = Nothing
'End Sub
Sub SetupcboMatch()
    On Error GoTo errHandler
    cboRef.BeginUpdate
    cboRef.WidthList = 500
    cboRef.HeightList = 162
    cboRef.AllowSizeGrip = True
   ' cboRef.AutoDropDown = True
    cboRef.Columns.Add "Ref"
    cboRef.Columns.Add "Order"
    cboRef.Columns.Add "Qty"
    cboRef.Columns.Add "COLID"
    cboRef.Columns.Add "Discount"
    cboRef.Columns.Add "Price"
    cboRef.Columns.Add "PriceLng"
    cboRef.Columns(0).Width = 100
    cboRef.Columns(1).Width = 70
    cboRef.Columns(2).Width = 70
    cboRef.Columns(3).Width = 0
    cboRef.Columns(4).Width = 70
    cboRef.Columns(5).Width = 70
    cboRef.Columns(6).Width = 0
    cboRef.BackColorLock = Me.BackColor
    cboRef.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.SetupcboMatch"
End Sub
Private Sub LoadMatches()
    On Error GoTo errHandler
Dim i As Integer
Dim oD As d_COLine
    If oGDN.COLsOSPerCUST.Count = 0 Then Exit Sub
    cboRef.Items.RemoveAllItems
    cboRef.BeginUpdate
    ReDim ar(6, oGDN.COLsOSPerCUST.Count)
    cboRef.Items.RemoveAllItems
    i = 0
    For Each oD In oGDN.COLsOSPerCUST
        If oGDN.GDNLines.FindLineByCOID(oD.COLID) Is Nothing Or vMode <> enAddingRow Then
            ReDim Preserve ar(6, i)
    
            ar(0, i) = oD.Ref
            ar(1, i) = oD.DOCCode
            ar(2, i) = oD.Qty
            ar(3, i) = oD.COLID
            ar(4, i) = oD.DiscountRate
            ar(5, i) = oD.Price
            ar(6, i) = oD.lngPrice
            i = i + 1
        Else
            If vMode = enAddingRow Then
                MsgBox "There are one or more lines entered for this item in this GDN already.", vbInformation + vbOKOnly, "Warning"
            End If
        End If
    Next
    If i > 0 Then
        cboRef.PutItems ar
    End If
    cboRef.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadMatches"
End Sub

Private Sub LoadandSHowcboRef(PID As String, Optional pCOLID As Long)
    On Error GoTo errHandler
            oGDN.LoadCOLsOS , PID, pCOLID
            If oGDN.COLsOSPerCUST.Count > 0 Then
                LoadMatches
                If cboRef.Items.ItemCount > 0 Then
                    cboRef.Visible = True
                    lblO1.Visible = True
                    lblO2.Visible = True
                    lblO3.Visible = True
                    lblO4.Visible = True
                    cboRef.Enabled = True
                    If Not oGDNLine.COLID > 0 Then
                        cboRef.DropDown() = True
                        cboRef.Items.SelectItem(cboRef.Items(0)) = True
                    End If
                End If
            Else
                cboRef.Items.RemoveAllItems
                lblO1.Visible = False
                lblO2.Visible = False
                lblO3.Visible = False
                lblO4.Visible = False
                cboRef.Visible = False
            End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.LoadandSHowcboRef(PID,pCOLID)", Array(PID, pCOLID)
End Sub

Private Sub HandlePossibleApprosOS(PID As String)
    On Error GoTo errHandler
    
        oGDN.LoadAPPLsOS oGDN.Customer.ID, PID
        If oGDN.APPLsOSPerCUST.Count > 0 Then
            Me.cmdAppro.BackColor = vbRed
            cmdAppro.Enabled = True
        Else
            Me.cmdAppro.BackColor = &HC4BCA4
            cmdAppro.Enabled = False
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.HandlePossibleApprosOS(PID)", PID
End Sub
Private Sub cboRef_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If vMode = enNotEditing Then
        Exit Sub
    End If
    txtDiscount = oGDNLine.DiscountPercent
    If oPC.AllowsSSInvoicing Then
        Me.txtQty = oGDNLine.QtyFirm
        Me.txtQtySS = oGDNLine.QtySS
    Else
        txtQty = oGDNLine.Qty
    End If
    txtRef = oGDNLine.Ref
    txtDiscount = oGDNLine.DiscountPercent
  '  txtPrice = oGDNLine.Price
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cboRef_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub GetAppros()
    On Error GoTo errHandler
Dim i As Integer
Dim frm As frmAPPOS
Dim tmpQty As Long
    If oGDN.APPLsOSPerCUST.Count > 0 Then
        Set frm = New frmAPPOS
        frm.component oGDN.APPLsOSPerCUST, "There is an appro outstanding for this item. " & vbCrLf & "Do you wish to leave it for return later or include it in this GDN?" & vbCrLf _
                & "By entering a non-zero quantity you will cause an appro return for that quantity to be issued when this GDN is issued.", oGDNLine.APPLID, oGDNLine.APPLQTY
        frm.Show vbModal
        oGDNLine.APPLID = frm.APPLID
        oGDNLine.DiscountPercent = frm.APPLDiscountRate
        tmpQty = frm.APPLQTY
        If tmpQty > oGDNLine.Qty Then
            MsgBox "You are wanting to return more appro items than you have delivered. The appro return quantity will be reduced to match the quantity delivered.", vbInformation, "Warning"
            oGDNLine.APPLQTY = oGDNLine.Qty
            
        Else
            oGDNLine.APPLQTY = tmpQty
        End If
        Unload frm
        Set frm = Nothing
    End If
    lblAppro.Caption = oGDNLine.APPLQTY
    txtDiscount = PBKSPercentF(oGDNLine.DiscountPercent)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.GetAppros"
End Sub
Private Sub cmdAppro_Click()
    On Error GoTo errHandler
    GetAppros
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.cmdAppro_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvwInvLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuSaveLayout"
End Sub

Public Sub mnuSalesComm()
    On Error GoTo errHandler
Dim frm As New frmSalesComm
Dim OpenResult As Integer

    frm.component oGDN.SalesRepID, oGDN.SalesRepName, oGDN.CustPaid, oGDN.CommPaid
    frm.Show vbModal
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If frm.CustPaid <> oGDN.CustPaid Then
        oPC.COShort.execute "EXECUTE dbo.MarkInvoicePaid " & oGDN.GDNID & "," & IIf(frm.CustPaid, "1", "0")
        oGDN.CustPaid = frm.CustPaid
    End If
    If frm.CommPaid <> oGDN.CommPaid Then
        oPC.COShort.execute "EXECUTE dbo.MarkCOmmissionPaid " & oGDN.GDNID & "," & IIf(frm.CommPaid, "1", "0")
        oGDN.CommPaid = frm.CommPaid
    End If
    
    
    If oGDN.SalesRepID <> frm.SalesRepID Then
        oGDN.SalesRepID = frm.SalesRepID
        oGDN.SalesRepName = frm.SalesRepName
        oPC.COShort.execute "Execute dbo.AllocateSalesCommission " & oGDN.GDNID & "," & oGDN.SalesRepID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuSalesComm"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset

    Set rs = oPC.LinesClipboard
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oGDN.DOCCodeF, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
   ' rs.Open
    If rs.BOF And rs.eof Then Exit Sub
    rs.MoveFirst
    Do While Not rs.eof
        MsgBox (rs.Fields("TITLE"))
        rs.MoveNext
    Loop
    rs.Close
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmGDN.mnuPastelines"
End Sub
