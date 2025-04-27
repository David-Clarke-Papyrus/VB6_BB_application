VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmInvoice 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice"
   ClientHeight    =   7455
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11010
   ControlBox      =   0   'False
   Icon            =   "frmInvoicewSS1.frx":0000
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
      Top             =   5355
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
      Picture         =   "frmInvoicewSS1.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   25
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
      Picture         =   "frmInvoicewSS1.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5355
      Width           =   1020
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1950
      Left            =   60
      TabIndex        =   14
      Top             =   3405
      Width           =   10725
      Begin EXCOMBOBOXLibCtl.ComboBox cboRef 
         Height          =   315
         Left            =   135
         OleObjectBlob   =   "frmInvoicewSS1.frx":2EB6
         TabIndex        =   55
         Top             =   1290
         Width           =   6810
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
         Picture         =   "frmInvoicewSS1.frx":4260
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
         Picture         =   "frmInvoicewSS1.frx":45EA
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
      Left            =   9780
      Picture         =   "frmInvoicewSS1.frx":4974
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5355
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
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
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
      TabIndex        =   2
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
      Picture         =   "frmInvoicewSS1.frx":4CFE
      Stretch         =   -1  'True
      Top             =   615
      Width           =   360
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oInvoice As a_Invoice
Attribute oInvoice.VB_VarHelpID = -1
Dim WithEvents oInvLine As a_InvoiceLine
Attribute oInvLine.VB_VarHelpID = -1
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
Dim bIsPreDelivery As Boolean

Public Sub component(PreDelivery As Boolean, Optional pCustID As Long, Optional pInvoice As a_Invoice, Optional Proforma As Boolean)
    On Error GoTo errHandler
Dim strAddress As String
    If pInvoice Is Nothing Then   'we create new invoice
'Handle objects
        bIsPreDelivery = PreDelivery
        Set oInvoice = New a_Invoice
        oInvoice.BeginEdit
        oInvoice.IsPreDelivery = bIsPreDelivery
        If Proforma = True Then
            oInvoice.SetProforma
            Caption = Caption & "     PRO-FORMA"
        End If
        
        If pCustID > 0 Then
            flgLoading = True
            LoadNewCustomer pCustID
            If Not oInvoice.BillTOAddress Is Nothing Then
                strAddress = oInvoice.BillTOAddress.AddressMailing
                lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
            End If
            If Not oInvoice.DelToAddress Is Nothing Then
                strAddress = oInvoice.DelToAddress.AddressMailing
                lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
            End If
            flgLoading = False
        End If
        oInvoice.VATable = oInvoice.Customer.VATable
        ChangeState enAddingRow
        
    Else   'we are provided with a loaded invoice
        Set oInvoice = pInvoice
        oInvoice.BeginEdit
        WaitMsg "Preparing to edit invoice  . . .", True, Me
        flgLoading = True
        If Not oInvoice.BillTOAddress Is Nothing Then
            strAddress = oInvoice.BillTOAddress.AddressMailing
            lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
        End If
        If Not oInvoice.DelToAddress Is Nothing Then
            strAddress = oInvoice.DelToAddress.AddressMailing
            lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
        End If
        If Not oInvoice.BillTOAddress Is Nothing Then
            strAddress = oInvoice.BillTOAddress.AddressMailing
            lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
        End If
        
        flgLoading = False
        oInvoice.GetSTatus
        oInvoice.SetDirty False
        ChangeState enNotEditing
    End If
    oInvoice.GetSTatus
    lblBlocked.Visible = oInvoice.Customer.Blocked
    cboRef.Visible = oPC.Configuration.SupportsWants
    If bIsPreDelivery Then
        cboRef.Visible = True
        Me.cmdCancelMatch.Enabled = False
        
    End If
    
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
    If oPC.AllowsInvoicePicking And Not oInvoice.Proforma Then
        cmdCancel.Left = 6720
        cmdSave.Left = 7740
        cmdPick.Left = 8760
        cmdPick.Visible = True
        If oInvoice.STATUS = stISSUED Then
            Me.cmdCancel.Enabled = False
            Me.cmdSave.Enabled = False
        End If
    Else
        cmdCancel.Left = 7740
        cmdSave.Left = 8760
        cmdPick.Visible = False
    End If

    If oPC.AllowsInvoicePicking Then
        If oInvoice.Proforma Then
            Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
        Else
            Caption = IIf(oInvoice.STATUS = stCOMPLETE, "Invoice for ", "Picking slip for ") & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
        End If
    Else
        Caption = "Invoice for " & oInvoice.Customer.NameAndCode(25) & oInvoice.StaffNameB & IIf(oInvoice.Proforma, "    PRO-FORMA", "")
    End If

Exithandler:
    WaitMsg "", False, Me
    
        
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.component(pCustID,pInvoice,Proforma)", Array(pCustID, pInvoice, Proforma)
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim strAddress As String

    SetupcboMatch
    LoadComps
    
    If Not oInvoice.IsNew Then
        chkChargeVAT = IIf(oInvoice.ShowVAT, 1, 0)
    Else
        chkChargeVAT = IIf(oPC.Configuration.DiscountVATDefault, 1, 0)
    End If
    
  '  SetLvw
    LoadCustomerDetailsToForm
    LoadListView
    
    lblNonVATQuestion.Visible = (oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False)
    lblNonVAT.Visible = (oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False)
    
    oInvoice.SetDirty False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadControls"
End Sub

Private Sub cmdCancelMatch_Click()
    On Error GoTo errHandler
Dim i As Integer

    If cboRef.Items.ItemCount = 0 Then Exit Sub
    oInvLine.COLID = 0
    For i = 0 To cboRef.Items.ItemCount - 1
        cboRef.Items.SelectItem(cboRef.Items(i)) = False
    Next
    Me.txtRef = ""
    mSetfocus cmdEnter

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdCancelMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdConvertToNonVat_Click()
    On Error GoTo errHandler
    oInvoice.ConvertToNonVATInvoice
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
    ErrorIn "frmInvoice.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPick_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoInvLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmInvoicePreview
Dim frmDte As frmTRDate
    'LogSaveToFile "Invoice Issue button"
    If (oPC.AllowsInvoicePicking = False And oInvoice.STATUS = stInProcess) Or oPC.AllowsInvoicePicking = True Then
    
'            If oPC.Configuration.Signtransactions = True Then
'                If SecurityControl(enSECURITY_INV_SIGN, , "Sign this invoice.", DOCAPPROVAL) = False Then
'                       Exit Sub
'                End If
'            End If
            
'            If oPC.AllowInvoiceDateOverride Then
'                Set frmDte = New frmTRDate
'                frmDte.Component Date
'                frmDte.Show vbModal
'                oInvoice.DocDate = StartOfDay(frmDte.InvoiceDate)
'                Unload frmDte
'                oInvoice.CaptureDate = Now()
'            Else
'                If oInvoice.DocDate < CDate("1950-01-01") Then
'                    oInvoice.DocDate = Date
'                    oInvoice.CaptureDate = Now()
'                End If
'            End If
            
            WaitMsg "Picking invoice  . . .", True, Me
            oInvoice.VATable = oInvoice.Customer.VATable
            oInvoice.StaffID = gSTAFFID
            oInvoice.RecalculateAllLines
            oInvoice.CalculateTotals
            If oInvoice.Proforma Then
                oInvoice.SetStatus stISSUED
            End If
            
            strResult = oInvoice.Post(stISSUED)
            
            If strResult = "" Then
                Set frm = New frmInvoicePreview
                frm.ComponentObject oInvoice
                frm.Show
            ElseIf strResult > "" Then
                MsgBox "The invoice cannot be issued now, try later. The record is probably locked by another user. The message is: " & strResult & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                oInvoice.BeginEdit
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
    ErrorIn "frmInvoice.cmdPick_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    cmdAppro.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oInvoice.StatusF = "IN PROCESS" And oInvoice.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oInvoice.StatusF = "ISSUED")      ' And oInvoice.CanCancel = True
    If oInvoice.Proforma = False Then
        Forms(0).mnuDelLine.Enabled = (oInvoice.STATUS = stInProcess)   'This should not happen if the invoice is in Picking state ONly delete from an in process invoice
    Else
        Forms(0).mnuDelLine.Enabled = True
    End If
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetMenu"
End Sub



Private Sub cbComp_Click()
    On Error GoTo errHandler
    oInvoice.COMPID = OptionLoop(oInvoice.COMPID, oPC.Configuration.Companies.Count)
    cbComp.Caption = oPC.Configuration.Companies(oInvoice.COMPID).CompanyName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cbComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oInvoice.Customer.ID > 0 Then
        frm.component oInvoice.Customer
        frm.Show
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cboRef_SelectionChanged()
    On Error GoTo errHandler
    Dim sQty As String
    Dim aq() As String
    
    If flgLoading Then Exit Sub
    If cboRef.Items.SelectCount = 0 Then
        oInvLine.COLID = 0
        Exit Sub
    End If
    oInvLine.COLID = cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 3)
    If oPC.AllowsSSInvoicing Then
        sQty = cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 2)
        aq = Split(sQty, "/")
        oInvLine.SetQtyFirm aq(0)
        oInvLine.SetQtySS aq(1)
    Else
        oInvLine.SetQty cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 2)
    End If
    If oInvoice.Customer.UseQuotedPrice Then
        oInvLine.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4)
        oInvLine.SetPrice cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6)
    Else
        If cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4) <> oInvoice.Customer.DefaultDiscount Then
            DoEvents
            If MsgBox("The customer discount is different than the discount on the customer order. Use the customer discount?", vbYesNo, "Warning") = vbYes Then
                oInvLine.SetDiscountPercent oInvoice.Customer.DefaultDiscount
            Else
                oInvLine.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 4)
            End If
        End If
        If cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6) <> oProd.SP Then
            If MsgBox("The product price is different than the price on the customer order. Use the product price?", vbYesNo, "Warning") = vbYes Then
                oInvLine.SetPrice oProd.SP
            Else
                oInvLine.SetPrice cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 6)
            End If
        End If
    End If
    
    oInvLine.SetRef cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 0)
    
    txtDiscount = oInvLine.DiscountPercent
    txtPrice = oInvLine.Price
    txtRef = oInvLine.Ref
    If oPC.AllowsSSInvoicing Then
        Me.txtQty = oInvLine.QtyFirm
        Me.txtQtySS = oInvLine.QtySS
    Else
        txtQty = oInvLine.Qty
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cboRef_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkChargeVAT_Click()
    On Error GoTo errHandler
    oInvoice.ShowVAT = (chkChargeVAT = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.chkChargeVAT_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdBill_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oInvoice.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oInvoice.Customer.Addresses.Count Then
        i = 1
    End If
    lblAddBill.Caption = oInvoice.Customer.Addresses(i).AddressMailing & vbCrLf & oInvoice.Customer.Addresses(i).EMail
    oInvoice.SetBillToAddress oInvoice.Customer.Addresses(i)
    iBillIdx = i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdBill_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDel_Click()
    On Error GoTo errHandler
Static iBillIdx As Integer
Dim i As Integer
START:
    If oInvoice.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oInvoice.Customer.Addresses.Count Then
        i = 1
    End If
    lblAddDel.Caption = oInvoice.Customer.Addresses(i).AddressMailing & vbCrLf & oInvoice.Customer.Addresses(i).EMail
    oInvoice.setDelToAddress oInvoice.Customer.Addresses(i)
    iBillIdx = i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdDel_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadNewCustomer(plngTPID As Long)
    On Error GoTo errHandler
    If oInvoice.SetCustomer(plngTPID) Then
        vCanAdd.RuleBroken "TP", False
        LoadCustomerDetailsToForm
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadNewCustomer(plngTPID)", plngTPID
End Sub
Private Sub LoadCustomerDetailsToForm()
    On Error GoTo errHandler
    With oInvoice.Customer
        If Not .BillTOAddress Is Nothing Then
            lblTPPhone.Caption = .BillTOAddress.Phone
            lblTPFax.Caption = .BillTOAddress.Fax
        End If
        lblTPName.Caption = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
        If Not .BillTOAddress Is Nothing Then
            If oInvoice.BillTOAddress Is Nothing Then
                oInvoice.SetBillToAddress .BillTOAddress
                lblAddBill.Caption = .BillTOAddress.AddressShort
            End If
        End If
        If Not .DelToAddress Is Nothing Then
            If oInvoice.DelToAddress Is Nothing Then
                oInvoice.setDelToAddress .DelToAddress
                lblAddDel.Caption = .DelToAddress.AddressShort
            End If
        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadCustomerDetailsToForm"
End Sub
Private Sub cmdNote()
    On Error GoTo errHandler
Dim frm As New frmILNote
    frm.component oInvLine
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdNote"
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

Private Sub mnuAddresses()
    On Error GoTo errHandler
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oInvoice
    frm.Show vbModal
    lblAddBill.Caption = oInvoice.BillTOAddress.AddressShort
    lblAddDel.Caption = oInvoice.DelToAddress.AddressShort
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuAddresses"
End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveInvoiceLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuDelLine"
End Sub

Private Sub lblNonVATQuestion_Click()
    On Error GoTo errHandler
Dim s As String
    s = "This message is shown because the 'Show VAT' check box has not been ticked in the customer record." _
    & " This is only applicable to customers who do not pay local VAT." & vbCrLf _
    & " The invoice will calculate values based on the ex VAT price and will make no reference to VAT on the printed document."
    MsgBox s, , "Non-VAT pricing"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.lblNonVATQuestion_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oInvoice_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidInvoice = (pMsg = "")
  '  cmdIssue.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0 And vMode = enNotEditing)
  '  cmdPick.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0 And vMode = enNotEditing)
  '  cmdSave.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0 And vMode = enNotEditing)
    
    cmdIssue.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidInvoice) And oInvoice.IsDirty
    cmdCancel.Enabled = True
    
    lblNonVATQuestion.Visible = (oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False)
    lblNonVAT.Visible = (oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False)
    
    If oInvoice.STATUS = stISSUED Then
        Me.cmdCancel.Enabled = False
        Me.cmdSave.Enabled = False
    End If
    Me.txtError = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvoice_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oInvLine_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvLine_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oInvLine_Valid(msg As String)
    On Error GoTo errHandler
        Me.cmdEnter.Enabled = (msg = "")
        Me.txtError = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oInvoice_TotalChange(lngTotalExt As Double, lngTotalDeposit As Double, lngTotalVAT As Double)
    On Error GoTo errHandler
    
    flgLoading = True
    
    lngCurrentTotal = lngTotalExt
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
  '  If vMode = enEditingRow Then
  '      cmdNewRows.Enabled = (oInvoice.InvoiceLines.Count > 0)
  '  End If

    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvoice_TotalChange(lngTotalExt,lngTotalDeposit,lngTotalVAT)", _
         Array(lngTotalExt, lngTotalDeposit, lngTotalVAT), EA_NORERAISE
    HandleError
End Sub

Private Sub oInvoice_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvoice_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oInvoice_Dirty(pVal As Boolean)
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
    ErrorIn "frmInvoice.oInvoice_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oInvoice_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvoice_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub


Private Sub SetFocusFromCode()
    On Error GoTo errHandler
Dim strMsg As String
    
    
    If LenB(txtCode) > 0 Then
        If oInvLine.ServiceItem = False Then
            If (oPC.Configuration.AntiquarianYN) And (Not oInvLine.Product.DefaultCopy Is Nothing) Then
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
    ErrorIn "frmInvoice.SetFocusFromCode"
End Sub

Private Sub txtCode_DblClick()
    On Error GoTo errHandler
   ' cmdFind_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtCode_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
    txtNote = HandleTextWithBites(txtNote)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtNote_Change", , EA_NORERAISE
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
    ErrorIn "frmInvoice.txtNote_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_DblClick()
    On Error GoTo errHandler
Dim f As New frmFCPrice
Dim oIL As a_InvoiceLine
Dim x As Long
Dim Y As Long
    
    If Not oPC.SupportsUNISA Then Exit Sub
    
    f.component Me.Left + 2000, Me.TOP + 1400, oInvLine.ForeignPrice, oInvLine.Price, oInvLine.FCID, oInvLine.VATRate, oInvLine.FCFactor

    f.Show vbModal
    If f.UserCancelled Then
        Unload f
        Exit Sub
    End If
    oInvLine.SetForeignPrice CStr(f.ForeignPrice)
    oInvLine.FCID = f.FCID
    If Round(f.FCFactor, 6) <> Round(oInvLine.FCFactor, 6) Then
        For Each oIL In oInvoice.InvoiceLines
            If oIL.FCID = oInvLine.FCID Then
                oIL.BeginEdit
                oIL.SetFCFactor Round(f.FCFactor, 6)
                oIL.ApplyEdit
            End If
        Next
    End If
    oInvLine.Price = f.LocalPriceIncVAT
    Me.txtPrice = oInvLine.Price
    lblFCTerms.Caption = oInvLine.ForeignPriceF & "/" & f.FCFactorINV
    Unload f

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtPrice_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_DblClick()
    On Error GoTo errHandler
    If Not oPC.SupportsUNISA Then Exit Sub
    txtQty = oInvoice.TotalQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtQty_DblClick", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtQty_GotFocus()
'    On Error GoTo errHandler
'    AutoSelect txtQty
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoice.txtQty_GotFocus", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If vMode = enNotEditing Then Exit Sub
    If oPC.AllowsSSInvoicing Then
        If Not oInvLine.SetQtyFirm(txtQty) Then
            Cancel = True
        End If
    Else
        If Not oInvLine.SetQty(txtQty) Then
            Cancel = True
        End If
    End If
    oInvLine.CalculateLine
    txtTotal = oInvLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtRef_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtRef
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtRef_GotFocus", , EA_NORERAISE
    HandleError
End Sub
'============
Private Sub txtQtySS_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtySS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtQtySS_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySs_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If vMode = enNotEditing Then Exit Sub
    If Not oInvLine.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    oInvLine.CalculateLine
    txtTotal = oInvLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmInvoice.vCanAdd_NobrokenRules", , EA_NORERAISE
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
    ErrorIn "frmInvoice.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub RefreshFCExchangeRates()
    On Error GoTo errHandler
 Dim oIL As a_InvoiceLine
 
    If Not oPC.SupportsUNISA Then Exit Sub
 
    For Each oIL In oInvoice.InvoiceLines
        If oIL.FCID > 0 Then
            oIL.BeginEdit
            oIL.SetFCFactor Round(oPC.Configuration.Currencies.FindCurrencyByID(oIL.FCID).Factor, 6)
            oIL.ApplyEdit
        End If
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.RefreshFCExchangeRates"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    Set vCanAdd = New z_BrokenRules
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
            'LogSaveToFile "Invoice form unloaded"
    If Not oCurrentCopy Is Nothing Then
        If oCurrentCopy.IsEditing Then oCurrentCopy.CancelEdit
    End If
    If oInvoice.IsEditing Then oInvoice.CancelEdit
    UnsetMenu
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oInvoice = Nothing
    Set tlCustomer = Nothing
    Set oInvLine = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
    If oInvLine.ServiceItem Then oInvLine.DiscountPercent = 0
    If oInvoice.Proforma = False Then
        If oPC.GetProperty("WarnOverInvoicing") = "TRUE" Then
            If oProd.QtyOnHand - (oInvoice.TotalQtyPerProduct(oInvLine.PID) + IIf(oInvLine.IsNew, oInvLine.Qty, 0)) < 0 And Not oProd.IsServiceItem Then
                MsgBox "There are not enough items in stock to invoice this quantity." & vbCrLf & "The maximum you can invoice is " _
                & CStr(GetMax(0, oProd.QtyOnHand)), vbInformation, "Warning"
            End If
        End If
        If oPC.GetProperty("StopOverInvoicing") = "TRUE" Then
            If oProd.QtyOnHand - (oInvoice.TotalQtyPerProduct(oInvLine.PID) + IIf(oInvLine.IsNew, oInvLine.Qty, 0)) < 0 And Not oProd.IsServiceItem Then
                MsgBox "There are not enough items in stock to invoice this quantity." & vbCrLf & "The maximum you can invoice is " _
                & CStr(GetMax(0, oProd.QtyOnHand)), vbInformation, "Can't do this"
                Exit Sub
            End If
        End If
    End If
    oInvLine.ApplyEdit
    oInvLine.BeginEdit
    Dim x As ListItem
    If vMode = enAddingRow Then
        For i = 1 To lvwInvLines.ListItems.Count
            strItemsDebug = strItemsDebug & "," & lvwInvLines.ListItems(i).Key
        Next
       ' On Error Resume Next
        If lvwInvLines.ListItems.Count < val(oInvLine.Key) Then
            lvwInvLines.ListItems.Add Key:=oInvLine.Key
            LoadListViewLine lvwInvLines.ListItems(lvwInvLines.ListItems.Count), oInvLine
        End If
        
        lvwInvLines.Refresh
        ChangeState enAddingRow
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvwInvLines.ListItems(lngSelectedRowIndex), oInvLine
        ChangeState enNotEditing
    End If
    oInvoice.GetSTatus
    cboRef.Items.RemoveAllItems

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdEnter_Click", , EA_NORERAISE
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
        cmdNewRows.Enabled = (oInvoice.InvoiceLines.Count > 0)
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
        cmdNewRows.Enabled = (oInvoice.InvoiceLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        lblTPPhone.Caption = ""
        lvwInvLines.Enabled = False
        lvwInvLines.Height = 2200
        ClearInvLineControls
        fr1.ZOrder 1
        mSetfocus txtCode
        Set oInvLine = oInvoice.InvoiceLines.Add
        oInvLine.SetParentInvoice oInvoice
        oInvLine.InvoiceID = oInvoice.InvoiceID
        oInvLine.SetQty 1
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
        cmdNewRows.Enabled = True  '(oInvoice.InvoiceLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        lvwInvLines.Enabled = True
        lvwInvLines.Height = 4000
        '''''''
'        If Not oInvLine Is Nothing Then oInvLine.CancelEdit
'        oInvoice.InvoiceLines.CancelEdit
'        oInvoice.InvoiceLines.BeginEdit
        
        SetMenu
        fr1.ZOrder 1
    End Select
    If Not oInvoice.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
    If oInvoice.STATUS = stISSUED Then
        Me.cmdCancel.Enabled = False
        Me.cmdSave.Enabled = False
    End If
    
    lblAppro.Caption = ""
    cboRef.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.ChangeState(pToMode)", pToMode
End Sub
Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
        If oInvLine.IsEditing Then
            oInvLine.CancelEdit
            oInvLine.BeginEdit
        End If
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        If txtCode > "" Then  'THis is not after a post but is an aborted  add row action
           oInvoice.InvoiceLines.DecrementMaxKeyUsed
        End If
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        ChangeState enAddingRow
    End If

    ClearInvLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdNewRows_Click", , EA_NORERAISE
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
    For i = 1 To oInvoice.InvoiceLines.Count
        Set lstItem = lvwInvLines.ListItems.Add
        LoadListViewLine lstItem, oInvoice.InvoiceLines(i)
        strItemsDebug = strItemsDebug & "," & lvwInvLines.ListItems(i).Key
    Next i
    Debug.Print strItemsDebug
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oInvLine As a_InvoiceLine)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oInvLine
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
    ErrorIn "frmInvoice.LoadListViewLine(lstItem,oInvLine)", Array(lstItem, oInvLine)
End Sub
Private Sub lvwInvLines_DblClick()
    On Error GoTo errHandler
    Dim strPos As String
    
'This must load the editing line with the current line's data
    If lvwInvLines.ListItems.Count = 0 Then Exit Sub
    If lvwInvLines.SelectedItem.Index < 1 Then Exit Sub
    
    lngILEditingIdx = lvwInvLines.SelectedItem.Key
    Set oInvLine = Nothing
    Set oInvLine = oInvoice.InvoiceLines(lngILEditingIdx)
    
    lngSelectedRowIndex = lvwInvLines.SelectedItem.Key

    ChangeState eneditingrow
    Set oProd = Nothing
    Set oProd = New a_Product
    oProd.Load oInvLine.PID, 0
  '  If oInvLine.COLID > 0 Then
        LoadandSHowcboRef "", oInvLine.COLID
        If Not oInvoice.Proforma Then
            If Me.cboRef.Items.ItemCount > 0 And oInvLine.COLID > 0 Then
                flgLoading = True
                cboRef.Items.SelectItem(cboRef.Items.FindItem(oInvLine.COLID, 3)) = True
                oInvLine.COLID = cboRef.Items.CellCaption(cboRef.Items.SelectedItem(0), 3)
                flgLoading = False
            Else
                cboRef.DropDown() = True
            End If
        End If
    HandlePossibleApprosOS oInvLine.PID
    
'Set up screen values
    txtDiscount = oInvLine.DiscountPercentF
    If oPC.AllowsSSInvoicing Then
        txtQty = oInvLine.QtyFirmF
        txtQtySS = oInvLine.QtySSF
    Else
        txtQty = oInvLine.QtyF
    End If
strPos = "position 7"
    txtNote = oInvLine.Note
    txtRef = oInvLine.Ref
    txtDiscount = oInvLine.DiscountPercent
        txtCode = oInvLine.CodeForEditing
    txtTitle = oInvLine.Title
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oInvLine.PriceF(False)
    Else
        txtPrice = oInvLine.Price
    End If
    lblAppro.Caption = oInvLine.APPLQTY
    If oInvLine.Qty > 1 Then
        mSetfocus txtQty
    Else
        mSetfocus txtPrice
    End If
strPos = "position 8"
    If oInvLine.FCID <> oPC.Configuration.DefaultCurrencyID And oInvLine.FCID > 0 Then
        Me.lblFCTerms = oInvLine.ForeignPriceF & "/" & oInvLine.FCFactorInvF
    End If
    
    oInvLine.GetSTatus
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
    If oInvoice.COMPID > 0 Then
        cbComp.Caption = oPC.Configuration.Companies(CStr(oInvoice.COMPID)).CompanyName
    Else
        cbComp.Caption = oPC.Configuration.DefaultCompany.CompanyName
        oInvoice.COMPID = oPC.Configuration.DefaultCOMPID
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadComps"
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If oInvoice.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'-------End Compsny code
'Private Sub txtNote_Change()
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oInvLine.setnote (txtNote)
'    If Err Then
'      Beep
'      intPos = txtNote.SelStart
'      txtNote = oInvLine.Note
'      txtNote.SelStart = intPos - 1
'    End If
'End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oInvLine.SetNote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oInvLine.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oInvoice.Memo
    ofrm.Show vbModal
    oInvoice.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuMemo"
End Sub

Public Sub mnuCancel()
    On Error GoTo errHandler
    If oInvoice.IsDirty Then
        oInvoice.CancelEdit
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuCancel"
End Sub






Public Sub mnuVoid()
    On Error GoTo errHandler
    oInvoice.SetStatus stVOID
    oInvoice.ApplyEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuVoid"
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

START:
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

        If oProd.PID > "" And Not oInvoice.Proforma Then
            Set rsPreviousBillings = oSM.PreviousBillings(oProd.PID, oInvoice.Customer.ID)
            'lngResult = oSQL.RunGetRecordset("SELECT TR_CODE,TR_DATE FROM tTR JOIN tINVOICE ON TR_ID = I_ID JOIN tILINE ON IL_TR_ID = TR_ID WHERE IL_P_ID = '" & oProd.pID & "' AND TR_TP_ID = " & oInvoice.Customer.ID, enText, Array(), "", rsPreviousBillings)
            If Not rsPreviousBillings Is Nothing Then
                If rsPreviousBillings.State <> 0 Then
                    If Not rsPreviousBillings.eof Then
                        strPreviousBillings = ""
                        Do While Not rsPreviousBillings.eof
                            strPreviousBillings = rsPreviousBillings.Fields(0) & "   " & Format(rsPreviousBillings.Fields(1), "dd/mm/yyyy") & vbCrLf
                            rsPreviousBillings.MoveNext
                        Loop
                        If strPreviousBillings > "" And oProd.IsServiceItem = False Then
                            MsgBox "This item is on a previous invoice to this client." & vbCrLf & strPreviousBillings, vbInformation, "Warning"
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
        Else
            If Not oProd.IsServiceItem And InStr(txtCode, "/") > 0 Then
                ' we may reach here is a copy is requested and not found
                MsgBox "No such copy exists", vbInformation, "Check"
                txtCode.SetFocus
                GoTo EXIT_Handler
            End If
        End If
        If Len(FNS(.PID)) <> 0 Then   'Book in database
            If Not oProd.DefaultCopy Is Nothing Then  'Copy requested and identified
                Set oCurrentCopy = oProd.DefaultCopy
                oInvLine.Price = oCurrentCopy.Price
                oInvLine.PIID = oCurrentCopy.ID
                oInvLine.DiscountPercent = oInvoice.Customer.DefaultDiscount
            Else
                If oProd.IsServiceItem Then   'No copy identified but product is a non-stock product (e.g. postage or insurance etc.)
                    ' mSetfocus txtPrice
                    ' AutoSelect txtPrice
                    ' oInvLine.Qty = 1
                     oInvLine.Qty = 1
                     oInvLine.Price = oProd.SPex((oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False))
                     oInvLine.CodeForExport = oProd.CodeForExport
                     oInvLine.CodeF = oProd.CodeF
                     oInvLine.code = oProd.EAN
                Else    ' we may reach here is a copy is requested and not found
                        ' OR No copy is requested and the Title is found
                    oInvLine.Price = oProd.SPex((oInvoice.Customer.VATable = False And oInvoice.Customer.ShowVAT = False))
                    oInvLine.CodeF = oProd.CodeF
                   ' oInvLine.EAN oProd.EAN
                    oInvLine.code = oProd.code
                    oInvLine.CodeForExport = oProd.CodeForExport
                    If oPC.Configuration.AllowCopyInfo And InStr(txtCode, "/") > 0 Then
                        If MsgBox("There is no copy with this serial number" & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Papyrus Invoicing Information") = vbNo Then
                            txtCode.SetFocus
                            GoTo EXIT_Handler
                        End If
                    End If
                End If
            End If
strPos = "Pos 4"
            LoadandSHowcboRef .PID
strPos = "Pos 5"
            HandlePossibleApprosOS .PID
strPos = "Pos 6"
            oInvLine.Title = .TitleAuthor  'L(35)
            oInvLine.PID = .PID
            oInvLine.ServiceItem = .IsServiceItem
            oInvLine.VATRate = .VATRateToUse
            oInvLine.Cost = .Cost
            If oInvLine.IsNew And oInvLine.DiscountPercent = 0 Then
                oInvLine.DiscountPercent = oInvoice.Customer.DefaultDiscount
            End If
            If oInvLine.DiscountPercent <> oInvoice.Customer.DefaultDiscount And oInvLine.IsNew = False Then
                If MsgBox("The discount on the invoice differs from the customer's usual discount. " & vbCrLf & "Use discount on order?", vbQuestion + vbYesNo, "Warning") = vbNo Then
                    oInvLine.DiscountPercent = oInvoice.Customer.DefaultDiscount
                End If
            End If
        Else   'Book nof found on database

            

        GetAdhocDetails
           txtCode.SetFocus
            GoTo START


            
            
        End If
        If Not .DefaultCopy Is Nothing Then
            oInvLine.CodeF = .code & .DefaultCopy.SerialF
        End If
    End With
    txtTitle = oInvLine.TitleAuthor
    txtPrice = oInvLine.Price
    txtQty = oInvLine.Qty
    txtRef = oInvLine.Ref
    txtDiscount = oInvLine.DiscountPercentF
    oInvLine.GetSTatus
    If Not oInvoice.Proforma Then
        If cboRef.Items.ItemCount = 0 Then
            SetFocusFromCode
        Else
            AutoSelect txtQty
        End If
    End If
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtCode_LostFocus", , EA_NORERAISE
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
        Unload frmAdHoc
        Set frmAdHoc = Nothing

End Function
Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not oInvLine.SetDiscountPercent(txtDiscount) Then
        Cancel = True
    End If
    oInvLine.CalculateLine
    txtTotal = oInvLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oInvLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oInvLine.CalculateLine
    txtTotal = oInvLine.PAfterDiscountExtF(False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oInvLine.SetRef (txtRef)
    If Err Then
      Beep
      intPos = txtRef.SelStart
      txtRef = oInvLine.Ref
      txtRef.SelStart = intPos - 1
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtRef_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oInvLine.SetRef(txtRef)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtRef_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtRef = oInvLine.Ref
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtRef_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub RemoveInvoiceLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvwInvLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwInvLines.ListItems(i).Selected Then
            oInvoice.InvoiceLines.Remove lvwInvLines.ListItems(i).Key
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
    ErrorIn "frmInvoice.RemoveInvoiceLine"
End Sub

Private Sub SaveInvoice()
    On Error GoTo errHandler
Dim strErrPos As String

strErrPos = "Pos 1"
If oInvoice Is Nothing Then
        'LogSaveToFile "SaveInvoice: oInvoice is nothing"
End If
    oInvoice.ApplyEdit
strErrPos = "Pos 2"
If oInvoice Is Nothing Then
        'LogSaveToFile "SaveInvoice: oInvoice is nothing"
End If
    oInvoice.BeginEdit
strErrPos = "Pos 3"
If oInvoice.InvoiceLines Is Nothing Then
        'LogSaveToFile "SaveInvoice: oInvoice.InvoiceLines is nothing"
End If
    Set oInvLine = oInvoice.InvoiceLines.Add
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SaveInvoice"
End Sub

Public Sub PrintInvoice()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoInvLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oInvoice.Load oInvoice.InvoiceID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoInvLines Then
        MsgBox "There are no records to print on this invoice.", vbOKOnly + vbInformation, "Papyrus Invoicing Status"
        GoTo EXIT_Handler
    End If
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.PrintInvoice"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoInvLines As Boolean
Dim iCurrency As Integer
Dim strResult As String
Dim frm As frmInvoicePreview
Dim frmDte As frmTRDate
Dim strOverInvoicedSet As String
Dim oSM As New z_StockManager
    If oInvoice.Customer.Blocked Then
        MsgBox "This customer is blocked. You cannot issue this invoice.", vbCritical + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If oInvoice.Proforma = False Then
        If oInvoice.QtyNonStandardVAT > 0 Then
            If MsgBox("There are items with non-standard VAT in this invoice, continue?", vbYesNo + vbInformation, "Warning") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    If (oPC.AllowsInvoicePicking = False And oInvoice.STATUS = stInProcess) Or oPC.AllowsInvoicePicking = True Or (oInvoice.Proforma And oInvoice.STATUS <> stCOMPLETE) Then
            If oPC.Configuration.SignTransactions = True Then
                If SecurityControl(enSECURITY_INV_SIGN, , "Sign this invoice.", DOCAPPROVAL) = False Then
                       Exit Sub
                End If
            End If
            
            If oInvoice.Proforma = False Then
                If oPC.GetProperty("WarnOverInvoicing") = "TRUE" Then
                    strOverInvoicedSet = oSM.GetOverInvoicedIems(oInvoice.InvoiceID)
                    If strOverInvoicedSet <> "" Then
                        If MsgBox("You are over-invoicing the following items." & vbCrLf & strOverInvoicedSet & vbCrLf & "Do you want to continue?", vbInformation + vbOKCancel, "Warning") = vbCancel Then
                            GoTo Redisplay
                        End If
                    End If
                End If
                If oPC.GetProperty("StopOverInvoicing") = "TRUE" Then
                    strOverInvoicedSet = oSM.GetOverInvoicedIems(oInvoice.InvoiceID)
                    If strOverInvoicedSet <> "" Then
                        MsgBox "You are over-invoicing the following items. You cannot issue this invoice until you have corrected it." & vbCrLf & strOverInvoicedSet, vbCritical, "Can't do this"
                        GoTo Redisplay
                    End If
                End If
            End If
            If oPC.AllowInvoiceDateOverride Then
                Set frmDte = New frmTRDate
                frmDte.component Date
                frmDte.Show vbModal
                oInvoice.DOCDate = StartOfDay(frmDte.InvoiceDate)
                Unload frmDte
                oInvoice.CaptureDate = Now()
            Else
                If oInvoice.DOCDate < CDate("1950-01-01") Then
                    oInvoice.DOCDate = Date
                    oInvoice.CaptureDate = Now()
                End If
            End If
            
            WaitMsg "Issuing invoice  . . .", True, Me
            oInvoice.VATable = oInvoice.Customer.VATable
            oInvoice.StaffID = gSTAFFID
            oInvoice.RecalculateAllLines
            oInvoice.CalculateTotals
           ' oInvoice.IsPreDelivery = bIsPreDelivery
            If oInvoice.Proforma Then
                oInvoice.SetStatus stISSUED
            End If
            
            If Not oInvoice.IsEditing Then oInvoice.BeginEdit
            strResult = oInvoice.Post(stCOMPLETE)
            
Redisplay:
            If strResult = "" Or strResult = "In Process" Then
                Set frm = New frmInvoicePreview
                frm.ComponentObject oInvoice
                frm.Show
            ElseIf strResult > "" Then
                MsgBox "The invoice cannot be issued now, try later. The record is probably locked by another user. The message is: " & strResult & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                If Not oInvoice.IsEditing Then oInvoice.BeginEdit
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
    ErrorIn "frmInvoice.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oInvoice_DirtyStatus(pDirty As Boolean)
    On Error GoTo errHandler
    If pDirty = True Then
        cmdCancel.Caption = "&Cancel"
    Else
        cmdCancel.Caption = "&Close"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvoice_DirtyStatus(pDirty)", pDirty, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
    oInvoice.SetStatus stInProcess
    If oInvoice.DOCDate < CDate("1950-01-01") Then
        oInvoice.DOCDate = Date
        oInvoice.CaptureDate = Now()
    End If
    oInvoice.RecalculateAllLines
    oInvoice.CalculateTotals
        'LogSaveToFile "Invoice Saving button"
    SaveInvoice
    LoadListView
    cmdSave.Enabled = False
    cmdCancel.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmInvoicePreview

    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmInvoicePreview
        frm.ComponentObject oInvoice
        frm.Show
    End If
    If cmdCancel.Caption <> "&Close" Then
        If oInvoice.IsEditing And oInvoice.IsDirty Then
            If MsgBox("You wish to cancel your changes?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
            oInvoice.CancelEdit
        End If
    End If
    If Not oInvLine Is Nothing Then
        If oInvLine.IsEditing Then oInvLine.CancelEdit
    End If
        'LogSaveToFile "Invoice Cancel button"
    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdCancel_Click", , EA_NORERAISE
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
    ErrorIn "frmInvoice.ClearInvLineControls"
End Sub

Private Sub lvwInvLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.lvwInvLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oInvoice.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oInvoice.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetIssueButtonCaption"
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
    ErrorIn "frmInvoice.lvwInvLines_Click", , EA_NORERAISE
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
    ErrorIn "frmInvoice.lvwInvLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
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
    ErrorIn "frmInvoice.SetLvw"
End Sub

Private Sub vCanAdd_Status(errors As String)
    On Error GoTo errHandler
MsgBox errors & "CANAADD"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.vCanAdd_Status(errors)", errors, EA_NORERAISE
    HandleError
End Sub

'Private Sub ReconcileWithCOs()
'Dim frm As frmCOFF
'Dim oInv As a_Invoice
'    Set oInv = New a_Invoice
'    oInv.Load oInvoice.InvoiceID, True
'    If oInv.hasCoffs Then
'        Set frm = New frmCOFF
'        frm.Component oInvoice
'        frm.Show vbModal
'    End If
'    Set oInv = Nothing
'End Sub
Sub SetupcboMatch()
    On Error GoTo errHandler
    cboRef.BeginUpdate
    cboRef.WidthList = 500
    cboRef.HeightList = 162
    cboRef.AllowSizeGrip = False
    cboRef.AllowHResize = False
    cboRef.BackColorLock = Me.BackColor
    cboRef.FullRowSelect = True
    cboRef.UseTabKey = False
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
    cboRef.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetupcboMatch"
End Sub
Private Sub LoadMatches()
    On Error GoTo errHandler
Dim i As Integer
Dim oD As d_COLine
    If oInvoice.COLsOSPerCUST.Count = 0 Then Exit Sub
    cboRef.Items.RemoveAllItems
    cboRef.BeginUpdate
    ReDim ar(6, oInvoice.COLsOSPerCUST.Count)
    cboRef.Items.RemoveAllItems
    i = 0
    For Each oD In oInvoice.COLsOSPerCUST
        If oInvoice.InvoiceLines.FindLineByCOID(oD.COLID) Is Nothing Or vMode <> enAddingRow Then
            ReDim Preserve ar(6, i)
    
            ar(0, i) = oD.Ref
            ar(1, i) = oD.DOCCode
            If oPC.AllowsSSInvoicing Then
                ar(2, i) = oD.QtyFirm & "/" & oD.QtySS
            Else
                ar(2, i) = oD.Qty
            End If
            ar(3, i) = oD.COLID
            ar(4, i) = oD.DiscountRate
            ar(5, i) = oD.Price
            ar(6, i) = oD.lngPrice
            
            i = i + 1
        Else
            If vMode = enAddingRow Then
                MsgBox "There are one or more lines entered for this item in this invoice already.", vbInformation + vbOKOnly, "Warning"
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
    ErrorIn "frmInvoice.LoadMatches"
End Sub

Private Sub LoadandSHowcboRef(PID As String, Optional pCOLID As Long)
    On Error GoTo errHandler
        If Not oInvoice.Proforma Then
            oInvoice.LoadCOLsOS , PID, pCOLID
            If oInvoice.COLsOSPerCUST.Count > 0 Then
                LoadMatches
                If cboRef.Items.ItemCount > 0 Then
                    cboRef.Visible = True
                    lblO1.Visible = True
                    lblO2.Visible = True
                    lblO3.Visible = True
                    lblO4.Visible = True
                    cboRef.Enabled = True
                    If Not oInvLine.COLID > 0 Then
                        cboRef.DropDown() = True
                        cboRef.Items.SelectItem(cboRef.Items(0)) = True
                    End If
                End If
            Else
                cboRef.Items.RemoveAllItems
                If Not bIsPreDelivery Then
                    lblO1.Visible = False
                    lblO2.Visible = False
                    lblO3.Visible = False
                    lblO4.Visible = False
                    cboRef.Visible = False
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
    ErrorIn "frmInvoice.LoadandSHowcboRef(PID,pCOLID)", Array(PID, pCOLID)
End Sub

Private Sub HandlePossibleApprosOS(PID As String)
    On Error GoTo errHandler
    
    If Not oInvoice.Proforma Then
        oInvoice.LoadAPPLsOS oInvoice.Customer.ID, PID
        If oInvoice.APPLsOSPerCUST.Count > 0 Then
            Me.cmdAppro.BackColor = vbRed
            cmdAppro.Enabled = True
        Else
            Me.cmdAppro.BackColor = &HC4BCA4
            cmdAppro.Enabled = False
        End If
    Else
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.HandlePossibleApprosOS(PID)", PID
End Sub
Private Sub cboRef_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If vMode = enNotEditing Then
        Exit Sub
    End If
    txtDiscount = oInvLine.DiscountPercent
    If oPC.AllowsSSInvoicing Then
        Me.txtQty = oInvLine.QtyFirm
        Me.txtQtySS = oInvLine.QtySS
    Else
        txtQty = oInvLine.Qty
    End If
    txtRef = oInvLine.Ref
    txtDiscount = oInvLine.DiscountPercent
  '  txtPrice = oInvLine.Price
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cboRef_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub GetAppros()
    On Error GoTo errHandler
Dim i As Integer
Dim frm As frmAPPOS
Dim tmpQty As Long
    If oInvoice.APPLsOSPerCUST.Count > 0 Then
        Set frm = New frmAPPOS
        frm.component oInvoice.APPLsOSPerCUST, "There is an appro outstanding for this item. " & vbCrLf & "Do you wish to leave it for return later or include it in this invoice?" & vbCrLf _
                & "By entering a non-zero quantity you will cause an appro return for that quantity to be issued when this invoice is issued.", oInvLine.APPLID, oInvLine.APPLQTY
        frm.Show vbModal
        oInvLine.APPLID = frm.APPLID
        oInvLine.DiscountPercent = frm.APPLDiscountRate
        tmpQty = frm.APPLQTY
        If tmpQty > oInvLine.Qty Then
            MsgBox "You are wanting to return more appro items than you have invoiced. The appro return quantity will be reduced to match the quantity invoiced.", vbInformation, "Warning"
            oInvLine.APPLQTY = oInvLine.Qty
            
        Else
            oInvLine.APPLQTY = tmpQty
        End If
        Unload frm
        Set frm = Nothing
    End If
    lblAppro.Caption = oInvLine.APPLQTY
    txtDiscount = PBKSPercentF(oInvLine.DiscountPercent)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.GetAppros"
End Sub
Private Sub cmdAppro_Click()
    On Error GoTo errHandler
    GetAppros
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdAppro_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvwInvLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuSaveLayout"
End Sub

Public Sub mnuSalesComm()
    On Error GoTo errHandler
Dim frm As New frmSalesComm
Dim OpenResult As Integer

    frm.component oInvoice.SalesRepID, oInvoice.SalesRepName, oInvoice.CustPaid, oInvoice.CommPaid
    frm.Show vbModal
    If frm.Cancelled Then
        Unload frm
        Exit Sub
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If frm.CustPaid <> oInvoice.CustPaid Then
        oPC.COShort.execute "EXECUTE dbo.MarkInvoicePaid " & oInvoice.InvoiceID & "," & IIf(frm.CustPaid, "1", "0")
        oInvoice.CustPaid = frm.CustPaid
    End If
    If frm.CommPaid <> oInvoice.CommPaid Then
        oPC.COShort.execute "EXECUTE dbo.MarkCOmmissionPaid " & oInvoice.InvoiceID & "," & IIf(frm.CommPaid, "1", "0")
        oInvoice.CommPaid = frm.CommPaid
    End If
    
    
    If oInvoice.SalesRepID <> frm.SalesRepID Then
        oInvoice.SalesRepID = frm.SalesRepID
        oInvoice.SalesRepName = frm.SalesRepName
        oPC.COShort.execute "Execute dbo.AllocateSalesCommission " & oInvoice.InvoiceID & "," & oInvoice.SalesRepID
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuSalesComm"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oI As a_InvoiceLine

    Set rs = oPC.LinesClipboard
    If rs.State = 0 Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oInvoice.DOCCodeF, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
   ' rs.Open
    If rs.BOF And rs.eof Then Exit Sub
    rs.MoveFirst
    Do While Not rs.eof
        Set oI = oInvoice.InvoiceLines.Add
        oI.BeginEdit
        oI.PID = rs.Fields("PID")
        oI.Ref = FNS(rs.Fields("REF"))
        oI.QtyFirm = FNDBL(rs.Fields("QtyFirm"))
        oI.QtySS = FNDBL(rs.Fields("QTYSS"))
        oI.Qty = FNDBL(oI.QtySS) + FNDBL(oI.QtyFirm)
        oI.Price = FNDBL(rs.Fields("Price"))
        oI.DiscountPercent = FNDBL(rs.Fields("DISCOUNTRATE"))
        oI.CodeF = FNS(rs.Fields("CODEF"))
      '  oI.EAN = rs.Fields("EANF")
        oI.Title = FNS(rs.Fields("TITLE"))
        oI.VATRate = FNDBL(rs.Fields("VATRATE"))
        oI.ApplyEdit
     '   MsgBox (rs.Fields("TITLE"))
        rs.MoveNext
    Loop
    rs.Close
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuPastelines"
End Sub
