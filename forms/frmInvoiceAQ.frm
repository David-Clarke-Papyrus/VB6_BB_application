VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmInvoice 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Invoice"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10770
   ControlBox      =   0   'False
   Icon            =   "frmInvoiceAQ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   10770
   Begin MSComctlLib.ListView lvwInvLines 
      Height          =   2145
      Left            =   75
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1200
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   3784
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
      TabIndex        =   23
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
      Height          =   765
      Left            =   8550
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmInvoiceAQ.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5355
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
      Height          =   975
      Left            =   4335
      MultiLine       =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5295
      Width           =   2865
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
      Height          =   795
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5340
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Height          =   765
      Left            =   7440
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmInvoiceAQ.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5355
      Width           =   1110
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   1875
      Left            =   75
      TabIndex        =   11
      Top             =   3420
      Width           =   10725
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
         TabIndex        =   42
         Top             =   1320
         Width           =   720
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
         Height          =   780
         Left            =   5355
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   465
         Width           =   3615
      End
      Begin VB.TextBox txtQty 
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
         Left            =   1755
         TabIndex        =   2
         Top             =   465
         Width           =   615
      End
      Begin VB.TextBox txtRef 
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
         Left            =   3435
         TabIndex        =   4
         Top             =   465
         Width           =   1125
      End
      Begin VB.TextBox txtDiscount 
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
         Left            =   4590
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
         Height          =   750
         Left            =   9675
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   975
         Width           =   975
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
         Left            =   9075
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   465
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
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1515
         Width           =   7110
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
         Left            =   2400
         TabIndex        =   3
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
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
         Left            =   105
         TabIndex        =   0
         Top             =   450
         Width           =   1605
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboRef 
         Height          =   345
         Left            =   165
         OleObjectBlob   =   "frmInvoiceAQ.frx":30B6
         TabIndex        =   1
         Top             =   1170
         Visible         =   0   'False
         Width           =   5160
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
         Left            =   7200
         TabIndex        =   43
         Top             =   1380
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
         Left            =   210
         TabIndex        =   41
         Top             =   915
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
         TabIndex        =   40
         Top             =   915
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
         TabIndex        =   39
         Top             =   915
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
         TabIndex        =   38
         Top             =   915
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Note"
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
         Left            =   5355
         TabIndex        =   37
         Top             =   195
         Width           =   1440
      End
      Begin VB.Label Label4 
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
         Height          =   240
         Left            =   3690
         TabIndex        =   28
         Top             =   195
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
         Height          =   240
         Left            =   4425
         TabIndex        =   17
         Top             =   195
         Width           =   1005
      End
      Begin VB.Label Label11 
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
         Left            =   9630
         TabIndex        =   16
         Top             =   195
         Width           =   645
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
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label8 
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
         Height          =   240
         Left            =   1860
         TabIndex        =   14
         Top             =   195
         Width           =   375
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
         Height          =   240
         Left            =   2550
         TabIndex        =   13
         Top             =   195
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
      Height          =   765
      Left            =   9675
      Picture         =   "frmInvoiceAQ.frx":4460
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5355
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cmdBill 
      Height          =   1065
      Left            =   6390
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   15
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
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   15
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
      Left            =   780
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   105
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   60
      Width           =   2460
      _ExtentX        =   4339
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   135
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3465
      Picture         =   "frmInvoiceAQ.frx":45AA
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

Private Sub cboRef_LostFocus()
    If vMode = enNotEditing Then
        Exit Sub
    End If
    txtDiscount = oInvLine.DiscountPercent

End Sub
Private Sub GetAppros()
Dim i As Integer
Dim frm As frmAPPOS
Dim tmpQty As Long
    If oInvoice.APPLsOSPerCUST.Count > 0 Then
        Set frm = New frmAPPOS
        frm.Component oInvoice.APPLsOSPerCUST, "There is an appro outstanding for this item. " & vbCrLf & "Do you wish to leave it for return later or include it in this invoice?" & vbCrLf _
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
End Sub

Private Sub cmdAppro_Click()
    GetAppros
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
    Forms(0).mnuVoid.Enabled = (oInvoice.statusF = "IN PROCESS" And oInvoice.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oInvoice.statusF = "ISSUED") ' And oInvoice.CanCancel = True
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True

    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuGenDisc.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetMenu"
End Sub


Public Sub Component(Optional pCustID As Long, Optional pInvoice As a_Invoice, Optional proforma As Boolean)
    On Error GoTo errHandler
    flgLoading = True
    If pInvoice Is Nothing Then
        Set oInvoice = New a_Invoice
        oInvoice.BeginEdit
        If proforma = True Then
            oInvoice.SetProforma
        End If
        If pCustID > 0 Then
            LoadNewCustomer pCustID
            If Not oInvoice.Customer.BillTOAddress Is Nothing Then
            If oInvoice.Customer.BillTOAddress.CountryID <> oPC.Configuration.LocalCountryID Then
                Me.chkChargeVAT.Enabled = True
            Else
                Me.chkChargeVAT.Enabled = False
            End If
            End If
        End If
        lvwInvLines.Enabled = False
        SetControlsForNew
        lvwInvLines.Height = 2200
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        mSetfocus txtCode
        Set oInvLine = oInvoice.InvoiceLines.Add
        oInvLine.SetQty 1
    Else
        Set oInvoice = pInvoice
        oInvoice.BeginEdit
        vMode = enNotEditing
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        If Not oInvoice.Customer.BillTOAddress Is Nothing Then
        If oInvoice.Customer.BillTOAddress.CountryID <> oPC.Configuration.LocalCountryID Then
            Me.chkChargeVAT.Enabled = True
        Else
            Me.chkChargeVAT.Enabled = False
        End If
        End If
        lvwInvLines.Enabled = True
        lvwInvLines.Height = 4000
        SetEditFrameEnabled False, enNotEditing
    End If
    flgLoading = False
    cboRef.Visible = oPC.Configuration.SupportsWants
    SetMenu
   ' LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Component(pCustID,pInvoice)", Array(pCustID, pInvoice)
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
        frm.Component oInvoice.Customer
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
    oInvLine.COLID = cboRef.Items.CellCaption(cboRef.Items(0), 3)
    oInvLine.SetDiscountPercent cboRef.Items.CellCaption(cboRef.Items(0), 4)
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
Start:
    If oInvoice.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oInvoice.Customer.Addresses.Count Then
        i = 1
    End If
    Me.lblAddBill.Caption = oInvoice.Customer.Addresses(i).AddressMailing & vbCrLf & oInvoice.Customer.Addresses(i).EMail
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
Start:
    If oInvoice.Customer.ID = 0 Then Exit Sub
    i = iBillIdx + 1
    If i > oInvoice.Customer.Addresses.Count Then
        i = 1
    End If
    Me.lblAddDel.Caption = oInvoice.Customer.Addresses(i).AddressMailing & vbCrLf & oInvoice.Customer.Addresses(i).EMail
    oInvoice.setDelTOAddress oInvoice.Customer.Addresses(i)
    iBillIdx = i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdDel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuChangeCustomer()
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
    ErrorIn "frmInvoice.mnuChangeCustomer"
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
     '   SetIssueButtonCaption
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
                oInvoice.setDelTOAddress .DelToAddress
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
    frm.Component oInvLine
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
    frm.Component oInvoice
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

Private Sub mnuGenDisc()
    On Error GoTo errHandler
Dim frm As frmGeneralDiscount
    Set frm = New frmGeneralDiscount
    frm.Component oInvoice
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.mnuGenDisc"
End Sub

Private Sub lblWants_Click()

End Sub

'Private Sub mnuPrint_Click()
'Dim frm As frmPrintingOptions_Inv
'    Set frm = New frmPrintingOptions_Inv
'    frm.Show vbModal
'
'End Sub

Private Sub oInvoice_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidInvoice = (pMsg = "")
    cmdIssue.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0 And vMode = enNotEditing)
    cmdSave.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0 And vMode = enNotEditing)
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

Private Sub oInvLine_Valid(MSG As String)
    On Error GoTo errHandler
        Me.cmdEnter.Enabled = (MSG = "")
        Me.txtError = MSG
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.oInvLine_Valid(Msg)", MSG, EA_NORERAISE
    HandleError
End Sub

Private Sub oInvoice_TotalChange(lngTotalExt As Long, lngTotalDeposit As Long, lngTotalVAT As Long)
    On Error GoTo errHandler
    flgLoading = True
    lngCurrentTotal = lngTotalExt
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
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


Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
Dim strMsg As String
    
    strMsg = " Pos 1:txtQty.enabled:" & IIf(txtQty.Enabled, "TRUE", "FALSE") & "   " & "txtQty.Visible:" & IIf(txtQty.Visible, "TRUE", "FALSE") & "   " _
        & "txtPrice.enabled:" & IIf(txtPrice.Enabled, "TRUE", "FALSE") & "   " & "txtPrice.Visible:" & IIf(txtPrice.Visible, "TRUE", "FALSE") & "   " _
        & "cboRef.enabled:" & IIf(cboRef.Enabled, "TRUE", "FALSE") & "   " & "cboRef.Visible:" & IIf(cboRef.Visible, "TRUE", "FALSE") & "   "
    If LenB(txtCode) > 0 Then
        If (oPC.Configuration.AntiquarianYN) And (Not oInvLine.Product.DefaultCopy Is Nothing) Then
    strMsg = "Pos 2: txtQty.enabled:" & IIf(txtQty.Enabled, "TRUE", "FALSE") & "   " & "txtQty.Visible:" & IIf(txtQty.Visible, "TRUE", "FALSE") & "   " _
        & "txtPrice.enabled:" & IIf(txtPrice.Enabled, "TRUE", "FALSE") & "   " & "txtPrice.Visible:" & IIf(txtPrice.Visible, "TRUE", "FALSE") & "   " _
        & "cboRef.enabled:" & IIf(cboRef.Enabled, "TRUE", "FALSE") & "   " & "cboRef.Visible:" & IIf(cboRef.Visible, "TRUE", "FALSE") & "   "
            txtQty.Enabled = False
            mSetfocus txtPrice
        ElseIf cboRef.Visible = False Then
    strMsg = "Pos 3: txtQty.enabled:" & IIf(txtQty.Enabled, "TRUE", "FALSE") & "   " & "txtQty.Visible:" & IIf(txtQty.Visible, "TRUE", "FALSE") & "   " _
        & "txtPrice.enabled:" & IIf(txtPrice.Enabled, "TRUE", "FALSE") & "   " & "txtPrice.Visible:" & IIf(txtPrice.Visible, "TRUE", "FALSE") & "   " _
        & "cboRef.enabled:" & IIf(cboRef.Enabled, "TRUE", "FALSE") & "   " & "cboRef.Visible:" & IIf(cboRef.Visible, "TRUE", "FALSE") & "   "
            mSetfocus txtQty
        Else
    strMsg = "Pos 4: txtQty.enabled:" & IIf(txtQty.Enabled, "TRUE", "FALSE") & "   " & "txtQty.Visible:" & IIf(txtQty.Visible, "TRUE", "FALSE") & "   " _
        & "txtPrice.enabled:" & IIf(txtPrice.Enabled, "TRUE", "FALSE") & "   " & "txtPrice.Visible:" & IIf(txtPrice.Visible, "TRUE", "FALSE") & "   " _
        & "cboRef.enabled:" & IIf(cboRef.Enabled, "TRUE", "FALSE") & "   " & "cboRef.Visible:" & IIf(cboRef.Visible, "TRUE", "FALSE") & "   "
            mSetfocus cboRef
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtCode_LostFocus", , EA_NORERAISE, , "strMsg", Array(strMsg)
    HandleError
End Sub

Private Sub txtQty_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtQty_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If vMode = enNotEditing Then Exit Sub
    If Not oInvLine.SetQty(txtQty) Then
        Cancel = True
    End If
    oInvLine.CalculateLine
    txtTotal = oInvLine.PLessDiscExtF(False)
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

Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    Me.cmdNewRows.Enabled = True
    Me.cmdCancel.Enabled = True
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
    left = 10
    top = 10
    Width = 11100
    Height = 6700
   ' SetupcboMatch
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
Dim strAddress As String

    flgLoading = True
    SetupcboMatch
    LoadComps
    If Not oInvoice.BillTOAddress Is Nothing Then
        strAddress = oInvoice.BillTOAddress.AddressMailing
    End If
    Me.lblAddBill.Caption = IIf(strAddress > "", strAddress, "unknown")
    If oInvoice.proforma Then
        Me.Caption = Me.Caption & "     PRO-FORMA"
    End If
    If Not oInvoice.DelToAddress Is Nothing Then
        strAddress = oInvoice.DelToAddress.AddressMailing
    End If
    Me.lblAddDel.Caption = IIf(strAddress > "", strAddress, "unknown")
    oInvoice.GetStatus
    If Not oInvoice.IsNew Then
        Me.chkChargeVAT = IIf(oInvoice.ShowVAT, 1, 0)
    Else
        chkChargeVAT = oPC.Configuration.DiscountVATDefault
    End If
    SetLvw
    LoadCustomerDetailsToForm
    LoadListView
    flgLoading = False

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
    Me.txtRef.Enabled = pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtTotal.Enabled = pYesNo
    Me.txtQty.Enabled = pYesNo
    Me.cmdEnter.Enabled = Not pYesNo
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidInvoice
    Me.cmdSave.Enabled = (Not pYesNo) And bValidInvoice And oInvoice.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
    Me.txtDiscount.BackColor = lngColour
    Me.txtRef.BackColor = lngColour
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub
Private Sub SetControlsForNew()
    On Error GoTo errHandler
    txtRef = ""
    lblTPPhone.Caption = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.SetControlsForNew"
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
'Dim frm As frmAPPOS
'Dim tmpQty As Long

    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        mSetfocus txtCode
        Exit Sub
    End If
    If oInvLine.NonStock Then oInvLine.DiscountPercent = 0
    
'    oInvoice.LoadAPPLsOs oInvoice.Customer.ID, oInvLine.pID
'    If oInvoice.APPLsOSPerCUST.Count > 0 Then
'        Set frm = frmAPPOS
'        frm.Component oInvoice.APPLsOSPerCUST, "There is an appro for this item. " & vbCrLf & "Do you wish to leave it for return later or include it in this invoice?"
'        frm.Show vbModal
'        oInvLine.APPLID = frm.APPLID
'        tmpQty = frm.APPLQTY
'        oInvLine.APPLQTY = tmpQty
'        Unload frm
'    End If
    
    
    oInvLine.ApplyEdit
    oInvLine.BeginEdit
    If vMode = enAddingRow Then
        For i = 1 To lvwInvLines.ListItems.Count
            strItemsDebug = strItemsDebug & "," & lvwInvLines.ListItems(i).Key
        Next
        lvwInvLines.ListItems.Add Key:=oInvLine.Key
        LoadListViewLine lvwInvLines.ListItems(lvwInvLines.ListItems.Count), oInvLine
        lvwInvLines.Refresh
        Set oInvLine = oInvoice.InvoiceLines.Add
        oInvLine.SetQty 1
        oInvLine.InvoiceID = oInvoice.InvoiceID
        mSetfocus txtCode
    ElseIf vMode = enEditingRow Then
        LoadListViewLine Me.lvwInvLines.ListItems(lngSelectedRowIndex), oInvLine
        cmdNewRows_Click
    End If
    oInvoice.GetStatus
    ClearInvLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.cmdEnter_Click", , EA_NORERAISE, , "vMode,Linecount,strItemsDebug,oInvLine.Key", Array(vMode, oInvoice.InvoiceLines.Count, strItemsDebug, oInvLine.Key)
    HandleError
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = enEditingRow Then
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enNotEditing
        lvwInvLines.Height = 4000
        lvwInvLines.Enabled = True
        fr1.ZOrder 1
    ElseIf vMode = enAddingRow Then
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enNotEditing 'enEditingRow
        lvwInvLines.Enabled = True
        lvwInvLines.Height = 4000
        txtError = ""
        fr1.ZOrder 1
    ElseIf vMode = enNotEditing Then
        cmdNewRows.Caption = "&Stop"
        SetEditFrameEnabled True, vMode
        vMode = enAddingRow
        lvwInvLines.Enabled = False
        lvwInvLines.Height = 2200
        mSetfocus txtCode
        Set oInvLine = oInvoice.InvoiceLines.Add
        oInvLine.SetQty 1
        oInvLine.InvoiceID = oInvoice.InvoiceID
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

    lvwInvLines.ListItems.Clear
    For i = 1 To oInvoice.InvoiceLines.Count
        Set lstItem = lvwInvLines.ListItems.Add
'        lstItem.SubItems(8) = oInvoice.InvoiceLines(i).Key
        LoadListViewLine lstItem, oInvoice.InvoiceLines(i)
'        With oInvoice.InvoiceLines(i)
'            lstItem.Text = .CodeF
'           ' lstItem.Key = .Key
'            lstItem.SubItems(1) = .TitleAuthorPublisher
'            lstItem.SubItems(2) = .Qty
'            If .Deposit <> 0 Then
'                lstItem.SubItems(3) = .DepositF(False)
'            Else
'                lstItem.SubItems(3) = " "
'            End If
'            lstItem.SubItems(4) = .PriceF(False)
'            lstItem.SubItems(5) = .DiscountPercentF  ' Format(.DiscountPercent, "##0.0%")
'            lstItem.SubItems(6) = .Ref
'            lstItem.SubItems(7) = .PLessDiscExtF(False)
'            lstItem.SubItems(8) = Format(.Key, "@@@@@@@@@@")
'            If .NonStock = True Then
'                lstItem.ForeColor = &H427182
'                lstItem.ListSubItems(1).ForeColor = &H427182
'                lstItem.ListSubItems(2).ForeColor = &H427182
'                lstItem.ListSubItems(3).ForeColor = &H427182
'                lstItem.ListSubItems(4).ForeColor = &H427182
'                lstItem.ListSubItems(5).ForeColor = &H427182
'                lstItem.ListSubItems(6).ForeColor = &H427182
'                lstItem.ListSubItems(7).ForeColor = &H427182
'            ElseIf .PIID = 0 Then
'                lstItem.ListSubItems(1).ForeColor = &H706034
'                lstItem.ListSubItems(2).ForeColor = &H706034
'                lstItem.ListSubItems(3).ForeColor = &H706034
'                lstItem.ListSubItems(4).ForeColor = &H706034
'                lstItem.ListSubItems(5).ForeColor = &H706034
'                lstItem.ListSubItems(6).ForeColor = &H706034
'                lstItem.ListSubItems(7).ForeColor = &H706034
'            End If
'        End With
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
        lstItem.Text = .CodeF
        lstItem.Key = .Key
        lstItem.SubItems(1) = .TitleAuthorPublisher
        lstItem.SubItems(2) = .Qty
        If .Deposit <> 0 Then
            lstItem.SubItems(3) = .DepositF(False)
        Else
            lstItem.SubItems(3) = " "
        End If
        lstItem.SubItems(4) = .PriceF(False)
        lstItem.SubItems(5) = .DiscountPercentF  ' Format(.DiscountPercent, "##0.0%")
        lstItem.SubItems(6) = .Ref
        lstItem.SubItems(7) = .PLessDiscExtF(False)
        lstItem.SubItems(8) = Format(.Key, "@@@@@@@@@@")
        If .NonStock = True Then
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
    ErrorIn "frmInvoice.LoadListViewLine(lstItem)", Array(lstItem)
End Sub
Private Sub lvwInvLines_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvwInvLines.ListItems.Count = 0 Then Exit Sub
    If lvwInvLines.SelectedItem.Index < 1 Then Exit Sub
    lngILEditingIdx = lvwInvLines.SelectedItem.Key
    Set oInvLine = oInvoice.InvoiceLines(lvwInvLines.SelectedItem.Key)
    
    '''''77777
    lngSelectedRowIndex = lvwInvLines.SelectedItem.Key
    Me.txtCode = CStr(oInvLine.CodeF)
    Me.txtTitle = oInvLine.Title
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oInvLine.PriceF(False)
    Else
        txtPrice = oInvLine.Price
    End If

'    Me.txtPrice = CStr(oInvLine.PriceF(False))
    Me.txtDiscount = CStr(oInvLine.DiscountPercentF)
    txtQty = oInvLine.Qty
    txtNote = oInvLine.Note
    Me.txtRef = oInvLine.Ref
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    fr1.ZOrder 0

    LoadandSHowcboRef oInvLine.pID
    HandlePossibleApprosOS oInvLine.pID
    lblAppro.Caption = oInvLine.APPLQTY
    If oInvLine.Qty > 1 Then
        mSetfocus txtQty
    Else
        mSetfocus txtPrice
    End If

    cmdNewRows.Caption = "&Stop edit"
    oInvLine.GetStatus
    
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
    Cancel = Not oInvLine.setnote(txtNote)
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
    ofrm.Component oInvoice.Memo
    ofrm.Show vbModal
    oInvoice.setMemo ofrm.Memo
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


Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim pNumOfApproLines As Long
Dim ErrorPos As String
Dim i As Integer



    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    Set oProd = New a_Product
    With oProd
        .Load 0, 0, Trim$(txtCode)
        'Check to see if copy sold
        If Not oProd.DefaultCopy Is Nothing Then 'Book in database and copy requested
            If oProd.DefaultCopy.SoldDate > CDate(0) Then  'Copy is sold
                MsgBox "Copy already sold", vbInformation, "Check"
                Cancel = True
                Exit Sub
            End If
        ElseIf Not oProd.isNonStock And InStr(txtCode, "/") > 0 Then
            ' we may reach here is a copy is requested and not found
            MsgBox "No such copy exists", vbInformation, "Check"
            Cancel = True
            Exit Sub
        End If
        If Len(FNS(.pID)) <> 0 Then   'Book in database
            If Not oProd.DefaultCopy Is Nothing Then  'Copy requested and identified
                Set oCurrentCopy = oProd.DefaultCopy
                oInvLine.Price = oCurrentCopy.Price
                oInvLine.PIID = oCurrentCopy.ID
            ElseIf oProd.isNonStock Then   'No copy identified but product is a non-stock product (e.g. postage or insurance etc.)
                mSetfocus txtPrice
                AutoSelect txtPrice
            Else    ' we may reach here is a copy is requested and not found
                    ' OR No copy is requested and the Title is found
                oInvLine.Price = oProd.SP
                If oPC.Configuration.AllowCopyInfo And InStr(txtCode, "/") > 0 Then
                    If MsgBox("There is no copy with this serial number" & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Papyrus Invoicing Information") = vbNo Then
                        Cancel = True
                        Exit Sub
                    End If
                End If
            End If
            
            LoadandSHowcboRef .pID
            
            HandlePossibleApprosOS .pID
            
            oInvLine.Title = .TitleAuthorPublisher  'L(35)
            oInvLine.pID = .pID
            oInvLine.NonStock = .isNonStock
            oInvLine.VATRate = .VATRateToUse
        Else   'Book nof found on database
            MsgBox "Cannot find book", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
            Cancel = True
            Exit Sub
        End If
        
        If .DefaultCopy Is Nothing Then
            oInvLine.code = .code
        Else
            oInvLine.code = .code
            oInvLine.CodeF = .code & .DefaultCopy.SerialF
        End If
    End With

    txtTitle = oInvLine.TitleAuthorPublisherL(35)
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oInvLine.PriceF(False)
    Else
        txtPrice = oInvLine.Price
    End If
    txtQty = oInvLine.Qty
    txtRef = oInvLine.Ref
    oInvLine.GetStatus
EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE, , "ErrorPos", Array(ErrorPos)
    HandleError
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not oInvLine.SetDiscountPercent(txtDiscount) Then
        Cancel = True
    End If
    oInvLine.CalculateLine
    txtTotal = oInvLine.PLessDiscExtF(False)
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
    txtTotal = oInvLine.PLessDiscExtF(False)
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
    On Error Resume Next
Dim intPos As Integer
    If flgLoading Then Exit Sub
    oInvLine.SetRef (txtRef)
    If Err Then
      Beep
      intPos = txtRef.SelStart
      txtRef = oInvLine.Ref
      txtRef.SelStart = intPos - 1
    End If
    Exit Sub
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
'*    oInvoice.ApplyEdit
'*    oInvoice.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.RemoveInvoiceLine"
End Sub

'Private Sub LoadCustomer()
'    With oInvoice
'        SetIssueButtonCaption
'        lblTPName.Caption = .customer.Fullname & IIf(.customer.AcNo > "", "  (" & .customer.AcNo & ")", "")
'        lblTPPhone.Caption = .customer.BillToAddress.Phone
'        lblTPFax.Caption = .customer.BillToAddress.Fax
'        Me.lblAddBill.Caption = .customer.BillToAddress.AddressMailing
'        Me.lblAddDel.Caption = .customer.DelToAddress.AddressMailing
'    End With
'End Sub


Private Sub SaveInvoice()
    On Error GoTo errHandler
  '  If oInvLine.IsEditing Then oInvLine.ApplyEdit
    oInvoice.ApplyEdit
    oInvoice.BeginEdit
    Set oInvLine = oInvoice.InvoiceLines.Add
    oInvoice.post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
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
    'oInvoice.PrintInvoice blnNoInvLines, blnHideVAT, blnDeposit, blnDiscount, oInvoice.CurrencyID '   iCurrency, blnRoundedUp
    
    If blnNoInvLines Then
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
    ErrorIn "frmInvoice.PrintInvoice"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoInvLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strresult As String
Dim frm As frmInvoicePreview

    If oInvoice.Status = stInProcess Then
        If MsgBox("Issue this invoice?", vbYesNo + vbQuestion, "Confirm") = vbYes Then
            If oPC.Configuration.Signtransactions = True Then
                If SecurityControl(2, gSTAFFID, , "Sign this document.", DOCAPPROVAL) = False Then
                       Exit Sub
                End If
            End If
      
            '  ReconcileWithCOs
            WaitMsg "Issuing invoice  . . .", True, Me
            oInvoice.Vatable = oInvoice.Customer.Vatable
            oInvoice.StaffID = gSTAFFID
            oInvoice.RecalculateAllLines
            oInvoice.CalculateTotals
            If oInvoice.proforma Then
                oInvoice.SetStatus stISSUED
            End If
            strresult = oInvoice.ApplyEdit
            If strresult = "TIMEOUT" Then
                MsgBox "The invoice cannot be saved. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                oInvoice.BeginEdit
                WaitMsg "", False, Me
                Exit Sub
            End If
            If Not oInvoice.proforma Then
                strresult = oInvoice.post(stCOMPLETE)
            End If
            If strresult = "" Then
                Set frm = New frmInvoicePreview
                frm.ComponentObject oInvoice
                frm.Show
            ElseIf strresult = "TIMEOUT" Then
                MsgBox "The invoice cannot be issued now, try later. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
                oInvoice.BeginEdit
                WaitMsg "", False, Me
                Exit Sub
            End If
'            If oPC.POSActive Then
'                oInvoice.InformLocalPOSdb
'            End If
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


Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
    oInvoice.SetStatus stInProcess
    oInvoice.RecalculateAllLines
    oInvoice.CalculateTotals
    SaveInvoice
 '   oInvoice.BeginEdit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoice.cmdSave_Click"
'    HandleError
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

    oInvoice.CancelEdit
    If Not oInvLine Is Nothing Then
        If oInvLine.IsEditing Then oInvLine.CancelEdit
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoice.cmdCancel_Click"
'    HandleError
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
        If oInvoice.statusF = "IN PROCESS" Then
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
    If lvwInvLines.SelectedItem.Index > 0 Then
        Clipboard.Clear
        Clipboard.SetText left(lvwInvLines.SelectedItem.Text, ISBNLENGTH)
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
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvwInvLines.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvwInvLines.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
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
    cboRef.AllowSizeGrip = True
    cboRef.AutoDropDown = True
    cboRef.Columns.Add "Ref"
    cboRef.Columns.Add "Order"
    cboRef.Columns.Add "Qty"
    cboRef.Columns.Add "COLID"
    cboRef.Columns.Add "Discount"
    cboRef.Columns(0).Width = 150
  '  MsgBox cboRef.Columns(0).Visible
    
    cboRef.Columns(1).Width = 70
    cboRef.Columns(2).Width = 70
    cboRef.Columns(3).Width = 0
    cboRef.Columns(4).Width = 70
    cboRef.BackColorLock = Me.BackColor
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
 '   If oInvoice.COLsOSPerCUST.CountFilter < 1 Then Exit Sub
    If oInvoice.COLsOSPerCUST.Count = 0 Then Exit Sub
    cboRef.BeginUpdate
 '   ReDim ar(3, oInvoice.COLsOSPerCUST.CountFilter - 1)
    ReDim ar(4, oInvoice.COLsOSPerCUST.Count)
    cboRef.Items.RemoveAllItems
    i = 0
    For Each oD In oInvoice.COLsOSPerCUST
        ar(1, i) = oD.DocCode
        ar(0, i) = oD.Ref
        ar(2, i) = oD.Qty
        ar(3, i) = oD.COLID
        ar(4, i) = oD.discountRate
        i = i + 1
    Next
    cboRef.PutItems ar
    cboRef.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmInvoice.LoadMatches"
End Sub

Private Sub LoadandSHowcboRef(pID As String)
    oInvoice.LoadCOLsOS , pID
    If oInvoice.COLsOSPerCUST.Count > 0 Then
        LoadMatches
        cboRef.Visible = True
        lblO1.Visible = True
        lblO2.Visible = True
        lblO3.Visible = True
        lblO4.Visible = True
        cboRef.Items.SelectItem(cboRef.Items(0)) = True
        oInvLine.Qty = cboRef.Items.CellCaption(cboRef.Items(0), 2)
        oInvLine.COLID = cboRef.Items.CellCaption(cboRef.Items(0), 3)
        oInvLine.Ref = cboRef.Items.CellCaption(cboRef.Items(0), 0)
        cboRef.Enabled = True
    Else
        oInvLine.Qty = 1
        cboRef.Items.RemoveAllItems
        lblO1.Visible = False
        lblO2.Visible = False
        lblO3.Visible = False
        lblO4.Visible = False
        cboRef.Visible = False
    End If

End Sub

Private Sub HandlePossibleApprosOS(pID As String)
    oInvoice.LoadAPPLsOs oInvoice.Customer.ID, pID
    If oInvoice.APPLsOSPerCUST.Count > 0 Then
        Me.cmdAppro.BackColor = vbRed
        cmdAppro.Enabled = True
    Else
        Me.cmdAppro.BackColor = &HC4BCA4
        cmdAppro.Enabled = False
    End If

End Sub
