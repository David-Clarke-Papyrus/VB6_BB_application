VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInvoice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Invoice"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmInvoiceAQbackup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.CommandButton cmdFulfilments 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Fulfilments"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6210
      MaskColor       =   &H00C4BCA4&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5625
      Width           =   1215
   End
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
      Left            =   795
      TabIndex        =   0
      Top             =   540
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
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   1485
   End
   Begin VB.CheckBox chkProforma 
      BackColor       =   &H00D8E2E7&
      Caption         =   "Pro-forma only"
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
      Left            =   75
      TabIndex        =   33
      Top             =   5685
      Width           =   1665
   End
   Begin VB.CheckBox chkChargeVAT 
      BackColor       =   &H00D8E2E7&
      Caption         =   "Discount VAT"
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
      Left            =   75
      TabIndex        =   32
      Top             =   5340
      Width           =   1575
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
      Height          =   765
      Left            =   8550
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmInvoiceAQbackup.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1830
      MultiLine       =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5220
      Width           =   2865
   End
   Begin VB.TextBox txtCustName 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   150
      Width           =   1680
   End
   Begin VB.Frame fr2 
      BackColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   675
      TabIndex        =   27
      Top             =   3945
      Width           =   10125
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
         TabIndex        =   28
         Top             =   870
         Visible         =   0   'False
         Width           =   480
      End
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
      Height          =   1110
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4035
      Width           =   630
   End
   Begin VB.TextBox txtComp 
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
      Left            =   825
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   105
      Width           =   2580
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Cancel"
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
      Picture         =   "frmInvoiceAQbackup.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1110
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   23
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3915
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00E0E0E0&
      Height          =   1305
      Left            =   690
      TabIndex        =   14
      Top             =   3870
      Width           =   10110
      Begin VB.TextBox txtRef 
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
         Left            =   6180
         TabIndex        =   5
         Top             =   465
         Width           =   1125
      End
      Begin VB.TextBox txtNote 
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   8910
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
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
         Left            =   7290
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
         Left            =   9060
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   435
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   465
         Width           =   1020
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   1785
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   540
         Width           =   3375
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
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
         Left            =   5190
         TabIndex        =   4
         Top             =   465
         Width           =   1000
      End
      Begin VB.TextBox txtCode 
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
         Left            =   120
         TabIndex        =   3
         Top             =   465
         Width           =   1605
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ref."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   6435
         TabIndex        =   38
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Disc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   7125
         TabIndex        =   20
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   8265
         TabIndex        =   19
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   135
         TabIndex        =   18
         Top             =   225
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2505
         TabIndex        =   17
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   5340
         TabIndex        =   16
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
      Height          =   315
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "IN PROCESS"
      Top             =   5250
      Width           =   1260
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
      Height          =   750
      Left            =   9675
      Picture         =   "frmInvoiceAQbackup.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5295
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwInvLines 
      Height          =   2580
      Left            =   75
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1215
      Width           =   10725
      _ExtentX        =   18918
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
      NumItems        =   8
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
      TabIndex        =   37
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
      TabIndex        =   36
      Top             =   60
      Width           =   300
   End
   Begin VB.Label lblAddDel 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   35
      Top             =   90
      Width           =   1950
   End
   Begin VB.Label lblAddBill 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   34
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
      TabIndex        =   26
      Top             =   150
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   3450
      X2              =   3450
      Y1              =   105
      Y2              =   465
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   25
      Top             =   135
      Width           =   555
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
      Left            =   315
      TabIndex        =   24
      Top             =   615
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3735
      Picture         =   "frmInvoiceAQbackup.frx":0BA8
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
      Begin VB.Menu mnuDiscount 
         Caption         =   "&General discount"
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
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy this invoice to new invoice"
      End
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




Private Sub chkChargeVAT_Click()
    oInvoice.ShowVAT = (chkChargeVAT = 1)
End Sub


Private Sub cmdFulfilments_Click()
    ReconcileWithCOs
End Sub

Private Sub cmdSelectCustomer_Click()
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    If oInvoice.SetCustomer(lngTPID) Then
        With oInvoice.Customer
            txtPhone = .Phone
            txtAccnum = .AcNo
            txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
            lblAddBill.Caption = .DefaultAddress.AddressShort
            lblAddDel.Caption = .DefaultAddress.AddressShort
        End With
        vCanAdd.RuleBroken "TP", False
    End If

End Sub


Private Sub mnuAddresses_Click()
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oInvoice
    frm.Show vbModal
    lblAddBill.Caption = oInvoice.BillToAddress.AddressShort
    lblAddDel.Caption = oInvoice.GoodsToAddress.AddressShort

End Sub

Private Sub mnuDel_Click()
    RemoveInvoiceLine
End Sub

Private Sub mnuDiscount_Click()
Dim frm As frmGeneralDiscount
    Set frm = New frmGeneralDiscount
    frm.component oInvoice
    frm.Show vbModal
End Sub

Private Sub mnuPrint_Click()
Dim frm As frmPrintingOptions_Inv
    Set frm = New frmPrintingOptions_Inv
    frm.Show vbModal

End Sub

Private Sub oInvoice_Valid(pMsg As String)
    bValidInvoice = (pMsg = "")
    cmdIssue.Enabled = (bValidInvoice And oInvoice.InvoiceLines.Count > 0)
    cmdSave.Enabled = bValidInvoice
    Me.txtError = pMsg
End Sub

Sub oInvLine_ExtensionChange(lngExtension As Long, strExtension As String)
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
End Sub

Private Sub oInvLine_Valid(Msg As String)
        Me.cmdEnter.Enabled = (Msg = "")
        Me.txtError = Msg
End Sub

Private Sub oInvoice_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
    flgLoading = True
    Me.txtRunningTotal = strtotal
    lngCurrentTotal = lngTotal
'    Me.txtRunningDeposit = strTotalDeposit
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
    flgLoading = False
End Sub

Private Sub oInvoice_Reloadlist()
    LoadListView
End Sub
Private Sub oInvoice_Dirty(pVal As Boolean)
If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
End Sub
Private Sub oInvoice_CurrRowStatus(pMsg As String)
    MsgBox "CurrentRow Status = " & pMsg
End Sub



'Private Sub txtDeposit_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oInvLine.SetDeposit(txtdeposit) Then
'        Cancel = True
'    End If
'    txtTotal = oInvLine.ExtensionF(False)
'End Sub
'Private Sub txtDeposit_GotFocus()
'    AutoSelect txtdeposit
'End Sub

Sub vCanAdd_NobrokenRules()
    Me.cmdNewRows.Enabled = True
    Me.cmdCancel.Enabled = True
    Me.cmdSave.Enabled = True
    Me.cmdIssue.Enabled = True
End Sub
Private Sub Form_Load()
Dim curTotalDeposit As Currency
    Left = 10
    Top = 10
    Width = 11100
    Height = 6700
    fr2.Height = 1230
    flgLoading = True
    LoadComps
 '   LoadCurrs
    flgLoading = False
    oInvoice.GetStatus
'    Me.cboCurr = oInvoice.ForeignCurrency.Description
    Me.chkChargeVAT = IIf(oInvoice.ShowVAT, 1, 0)
    SetLvw
End Sub
Private Sub Form_Initialize()
    Set vCanAdd = New z_BrokenRules
End Sub
Private Sub Form_Unload(cancel As Integer)
  '  If oInvoice.Customer.IsEditing Then oCustomer.CancelEdit
    If Not oCurrentCopy Is Nothing Then
        If oCurrentCopy.IsEditing Then oCurrentCopy.CancelEdit
    End If
  '  If Not oInvLine Is Nothing Then
  '      If oInvLine.IsEditing Then oInvLine.CancelEdit
  '  End If
    If oInvoice.IsEditing Then oInvoice.CancelEdit
    
    Set oCustomer = Nothing
    Set oCurrentCopy = Nothing
    Set oInvoice = Nothing
    Set tlCustomer = Nothing
    Set oInvLine = Nothing
End Sub

Public Sub component(Optional pInvoice As a_Invoice)
    flgLoading = True
    If pInvoice Is Nothing Then
        Set oInvoice = New a_Invoice
        oInvoice.BeginEdit
        Me.lvwInvLines.Enabled = False
        SetControlsForNew
        vCanAdd.RuleBroken "TP", True
    Else
        Set oInvoice = pInvoice
        oInvoice.BeginEdit
        LoadCustomer
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        mnuFileCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        Me.lvwInvLines.Enabled = True
    End If
    Me.txtAccnum.SetFocus
    SetEditFrameEnabled False, enNotEditing
    Me.txtAccnum.SetFocus
    vMode = enNotEditing
    flgLoading = False
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
    Me.txtRef.Enabled = pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtTotal.Enabled = pYesNo
    Me.txtAccnum.Enabled = Not pYesNo
    Me.txtComp.Enabled = Not pYesNo
'    Me.cmdNewRows.Enabled = Not pYesNo
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
End Sub
Private Sub SetControlsForNew()
    mnuFileCancel.Caption = "&Cancel"
    txtAccnum = ""
    txtRef = ""
    txtPhone = ""
    txtStatus = "IN PROCESS"
End Sub

Private Sub cmdEnter_Click()
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
    
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        txtCode.SetFocus
        Exit Sub
    End If
    If oInvLine.NonStock Then oInvLine.DiscountPercent = 0
    oInvLine.ApplyEdit
    oInvLine.BeginEdit

    If vMode = enAddingRow Then
        lvwInvLines.ListItems.Add 1, oInvLine.Key
        LoadListViewLine oInvLine.Key, Me.lvwInvLines.ListItems(1)
        Set oInvLine = oInvoice.InvoiceLines.Add
        oInvLine.SetQty 1
        oInvLine.InvoiceID = oInvoice.InvoiceID
        txtCode.SetFocus
    ElseIf vMode = enEditingRow Then
        LoadListViewLine lngSelectedRowIndex, Me.lvwInvLines.ListItems(lngSelectedRowIndex)
    End If
    
    ClearInvLineControls

End Sub


Private Sub cmdNewRows_Click()
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = enEditingRow Then
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enNotEditing
        fr2.ZOrder 0
        fr1.ZOrder 1
        Me.lvwInvLines.Enabled = True
    ElseIf vMode = enAddingRow Then
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enNotEditing 'enEditingRow
        fr2.ZOrder 0
        fr1.ZOrder 1
        Me.lvwInvLines.Enabled = True
        Me.txtError = ""
    ElseIf vMode = enNotEditing Then
        cmdNewRows.Caption = "&Stop"
        SetEditFrameEnabled True, vMode
        vMode = enAddingRow
        fr2.ZOrder 1
        fr1.ZOrder 0
        Me.lvwInvLines.Enabled = False
        Me.txtCode.SetFocus
        Set oInvLine = oInvoice.InvoiceLines.Add
        oInvLine.SetQty 1
        oInvLine.InvoiceID = oInvoice.InvoiceID
   
    End If

    ClearInvLineControls
End Sub
Private Sub LoadListView()
Dim lstItem As ListItem
Dim i As Long
    On Error GoTo ERR_Handler
    lvwInvLines.ListItems.Clear
    For i = 1 To oInvoice.InvoiceLines.Count
        Set lstItem = lvwInvLines.ListItems.Add
        Set oInvLine = oInvoice.InvoiceLines(i)
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
    With oInvLine
        lstItem.Text = .CodeF
        If lstItem.Key = "" Then lstItem.Key = i
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
        lstItem.SubItems(7) = .ExtensionF(False)
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
End Sub
Private Sub lvwInvLines_DblClick()
'This must load the editing line with the current line's data
    If lvwInvLines.ListItems.Count = 0 Then Exit Sub
    lngILEditingIdx = lvwInvLines.SelectedItem.Key
    Set oInvLine = oInvoice.InvoiceLines(lngILEditingIdx)
    
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
    txtNote = oInvLine.Note
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    txtPrice.SetFocus
    fr2.ZOrder 1
    fr1.ZOrder 0
    cmdNewRows.Caption = "&Stop edit"
    oInvLine.GetStatus
End Sub

'---------Companies code
Private Sub LoadComps()
Dim oComp As a_Company
Dim oItem As ListItem
Dim i As Integer
    If oInvoice.COMPID > 0 Then
        txtComp = oPC.Configuration.Companies(CStr(oInvoice.COMPID)).CompanyName
    Else
        txtComp = oPC.Configuration.DefaultCompany.CompanyName
        oInvoice.COMPID = oPC.Configuration.DefaultCOMPID
    End If
End Sub

Private Sub cboTP_Validate(cancel As Boolean)
    If oInvoice.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        cancel = True
    End If
End Sub
'-------End Compsny code
Private Sub txtNote_Change()
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oInvLine.SetNote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oInvLine.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub
Private Sub txtNote_Validate(cancel As Boolean)
    cancel = Not oInvLine.SetNote(txtNote)
End Sub
Private Sub txtNote_LostFocus()
    If flgLoading Then Exit Sub
    txtNote = oInvLine.Note
End Sub

Private Sub mnuEditNote_Click()
Dim ofrm As New frmNote
    ofrm.component oInvoice
    ofrm.Show vbModal
    Unload ofrm
    Set ofrm = Nothing
End Sub

Private Sub mnuFileCancel_Click()
    If oInvoice.IsDirty Then
        oInvoice.CancelEdit
    End If
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    oInvoice.CancelEdit
    Unload Me
End Sub

Private Sub mnuFileOK_Click()
'    cmdOK_Click
End Sub

Private Sub mnuFilePrint_Click()
    cmdIssue_Click
End Sub
Private Sub mnuFileVoid_Click()
    oInvoice.setStatus stVOID
    txtStatus = "Void"
End Sub
Private Sub txtAccNum_Validate(cancel As Boolean)
Dim lngCustID As Long
Dim bResult As Boolean
    If Len(txtAccnum) > 0 Then
        bResult = oInvoice.SetCustomerFromAccNum(txtAccnum)
        If bResult Then
            With oInvoice.Customer
                txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
                txtPhone = .Phone
                lblAddBill.Caption = .DefaultAddress.AddressShort
                lblAddDel.Caption = .DefaultAddress.AddressShort
            End With
            vCanAdd.RuleBroken "TP", False
            Me.cmdNewRows.Enabled = True
        Else
            MsgBox "No such account number", , "Can't fetch customer"
            txtAccnum = ""
            Set oCustomer = Nothing
            cancel = True
        End If
    End If
End Sub
Private Sub txtAccNum_LostFocus()
    txtAccnum = UCase(txtAccnum)
End Sub

Private Sub txtComp_DblClick()
Dim iCompIdx As Integer
Dim i As Integer
Start:
    i = iCompIdx + 1
    If i > oPC.Configuration.Companies.Count Then
        i = 1
    End If
    If lngCompanyID = oPC.Configuration.Companies(i).ID Then
        GoTo Start
    End If
    txtComp = oPC.Configuration.Companies(i).CompanyName
    oInvoice.COMPID = oPC.Configuration.Companies(i).ID
    iCompIdx = i
End Sub

Private Sub txtCode_Validate(cancel As Boolean)
Dim pQty As Integer
Dim pApproID As Long
Dim pNumOfApproLines As Long
    On Error GoTo ERR_Handler
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    Set oProd = New a_Product
    With oProd
        .Load 0, 0, Trim$(txtCode)
        If Not oProd.DefaultCopy Is Nothing Then
            If oProd.DefaultCopy.SoldDate > CDate(0) Then
                MsgBox "Copy already sold", vbInformation, "Check"
                cancel = True
                Exit Sub
            End If
        End If
            
        If Len(FixNullsString(.pID)) <> 0 Then
            If Not oProd.DefaultCopy Is Nothing Then
                Set oCurrentCopy = oProd.DefaultCopy
                oInvLine.Price = oCurrentCopy.Price
                oInvLine.PIID = oCurrentCopy.ID
                AutoSelect txtPrice
            ElseIf oProd.NonStock Then
                Me.txtPrice.SetFocus
                AutoSelect txtPrice
            Else
                oInvLine.Price = oProd.RRP
                If oPC.Configuration.AllowCopyInfo Then
                    If MsgBox("There is no copy with this serial number" & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Papyrus Invoicing Information") = vbNo Then
                        cancel = True
                        Exit Sub
                    End If
                End If
            End If
            oInvLine.Title = .TitleAuthorPublisher
            oInvLine.pID = .pID
            oInvLine.NonStock = .NonStock
            oInvLine.VATRate = .VATRateToUse
        Else
            If InStr(1, txtCode, "/") > 0 Then
                MsgBox "Cannot find copy on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
                cancel = True
                Exit Sub
            Else
                MsgBox "Cannot find book on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
                cancel = True
                Exit Sub
            End If
        End If
        If .DefaultCopy Is Nothing Then
            oInvLine.Code = .Code
        Else
            oInvLine.Code = .Code
            oInvLine.CodeF = .Code & .DefaultCopy.SerialF
        End If
    End With

    txtTitle = oInvLine.Title
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oInvLine.PriceF(False)
    Else
        txtPrice = oInvLine.Price
    End If
    txtPrice.SetFocus
    oInvLine.GetStatus
    
EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub txtDiscount_Validate(cancel As Boolean)
    If Not oInvLine.SetDiscountPercent(txtDiscount) Then
        cancel = True
    End If
    txtTotal = oInvLine.ExtensionF(False)
End Sub
Private Sub txtPrice_Validate(cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oInvLine.SetPrice(txtPrice) Then
        cancel = True
    End If
    txtTotal = oInvLine.ExtensionF(False)
End Sub
Private Sub txtPrice_GotFocus()
    AutoSelect txtPrice
End Sub
Private Sub txtRef_Change()
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
End Sub
Private Sub txtRef_Validate(cancel As Boolean)
    If flgLoading Then Exit Sub
    cancel = Not oInvLine.SetRef(txtRef)
End Sub
Private Sub txtRef_LostFocus()
    If flgLoading Then Exit Sub
    txtRef = oInvLine.Ref
End Sub

Private Sub RemoveInvoiceLine()
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
End Sub

Private Sub LoadCustomer()
    With oInvoice
        txtStatus = .StatusF
        SetIssueButtonCaption
        txtAccnum = .TPAccNum
 '       cboTP = Trim$(.TPName)
        txtPhone = .TPPhone
        txtPhone = .TPPhone
    '    txtFax = .TPFax
    End With
End Sub


Private Sub SaveInvoice()
On Error GoTo ERR_Handler
    
    oInvoice.post
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Public Sub PrintInvoice()
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoInvLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    On Error GoTo ERR_Handler
    
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
Dim blnNoInvLines As Boolean
Dim iCurrency As Integer
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmInvoicePreview

    If oInvoice.status = stInProcess Then
        If MsgBox("Issue this invoice?.  Confirm.", vbYesNo + vbQuestion, "Papyrus Invoicing Status") = vbNo Then
            Exit Sub
        End If
    End If
  '  ReconcileWithCOs
    If Me.chkProforma = 0 Then  'Unchecked
        oInvoice.setStatus stCOMPLETE
    Else
        oInvoice.setStatus stPROFORMA
    End If
    
    strResult = oInvoice.post
    ReconcileWithCOs
    Set frm = New frmInvoicePreview
    frm.ComponentObject oInvoice
    frm.Show
    Unload Me
End Sub
Private Sub cmdSave_Click()
    oInvoice.setStatus stInProcess
    SaveInvoice
    ReconcileWithCOs
    oInvoice.BeginEdit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    oInvoice.CancelEdit
    Unload Me
End Sub
'Private Sub cmdCancelPost_Click()
'    If vMode = enAddingRow Then
'        oInvoice.InvoiceLines().CancelEdit
'        oInvoice.InvoiceLines.Remove lngILEditingIdx
'    ElseIf vMode = enEditingRow Then
'        oInvoice.InvoiceLines().CancelEdit
'        oInvoice.InvoiceLines().BeginEdit
'    End If
'    SetEditFrameEnabled False
'    ClearInvLineControls
'  '  Me.cmdCancelPost.Enabled = False
'End Sub


Private Sub ClearInvLineControls()
    flgLoading = True
    Me.txtCode = ""
    Me.txtDiscount = ""
    Me.txtPrice = ""
    Me.txtTitle = ""
    Me.txtTotal = ""
    Me.txtNote = ""
    flgLoading = False
End Sub

Private Sub lvwInvLines_BeforeLabelEdit(cancel As Integer)
    cancel = True
End Sub
Private Sub SetIssueButtonCaption()
        If oInvoice.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oInvoice.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
End Sub


Private Sub lvwInvLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
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
End Sub
Private Sub SetLvw()
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


End Sub

Private Sub vCanAdd_Status(errors As String)
MsgBox errors & "CANAADD"
End Sub

Private Sub ReconcileWithCOs()
Dim frm As frmCOFF
Dim oIL As a_InvoiceLine
Dim bCOFFsExist  As Boolean
Dim strSQL As String
Dim rs As ADODB.Recordset
'Check to see if there are any COLs outstanding for any titles on this invoice
    'first look for any matches already recorded
    bCOFFsExist = False
    For Each oIL In oInvoice.InvoiceLines
        If oIL.COFFs.Count > 0 Then
            bCOFFsExist = True
            Exit For
        End If
    Next
    'next look for any COLs with that aren't already in COFFs
    Set rs = New ADODB.Recordset
    strSQL = "SELECT COL_QTY,COL_QtyDispatched,P_Title,P_Code,TR_Code,IL_Qty,COL_ID,IL_ID FROM tCOL JOIN tTR ON COL_TR_ID = TR_ID JOIN tILine ON COL_P_ID = IL_P_ID JOIN tProduct on IL_P_ID = P_ID" _
            & " WHERE TR_TP_ID = " & oInvoice.Customer.ID & " AND IL_TR_ID = " & oInvoice.InvoiceID '& " AND COL_QtyDispatched <> COL_Qty"
    rs.Open strSQL, oPC.CO
    
    If Not rs.EOF Then
        Set frm = New frmCOFF
        frm.component oInvoice, rs
        frm.Show vbModal
    End If
    
End Sub
