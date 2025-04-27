VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPO 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Invoice"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmSO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.CheckBox chkProforma 
      Alignment       =   1  'Right Justify
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
      Left            =   9210
      TabIndex        =   45
      Top             =   6135
      Width           =   1665
   End
   Begin VB.CheckBox chkChargeVAT 
      BackColor       =   &H00D8E2E7&
      Caption         =   "Discount VAT for overseas customers"
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
      Left            =   150
      TabIndex        =   44
      Top             =   6030
      Width           =   3975
   End
   Begin VB.ComboBox cboCurr 
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
      Left            =   5535
      TabIndex        =   42
      Top             =   5700
      Width           =   1860
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
      Picture         =   "frmSO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.CommandButton cmdAddrNotes 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Addresses and notes"
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
      Left            =   120
      MaskColor       =   &H00C4BCA4&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5265
      Width           =   2010
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
      Height          =   705
      Left            =   2265
      TabIndex        =   38
      Top             =   5295
      Width           =   2745
   End
   Begin VB.TextBox txtCustName 
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
      Height          =   255
      Left            =   3945
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   165
      Width           =   2985
   End
   Begin VB.Frame fr2 
      BackColor       =   &H00E0E0E0&
      Height          =   1275
      Left            =   675
      TabIndex        =   35
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
         TabIndex        =   36
         Top             =   870
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox txtTP 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      Left            =   3165
      TabIndex        =   2
      Top             =   570
      Width           =   615
   End
   Begin VB.ComboBox cboTP 
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
      Left            =   3855
      TabIndex        =   3
      Top             =   570
      Width           =   3165
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3990
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   2580
   End
   Begin VB.TextBox txtGeneralDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2220
      TabIndex        =   10
      Top             =   3450
      Width           =   1395
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
      Picture         =   "frmSO.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5370
      Width           =   1110
   End
   Begin VB.TextBox txtPhone 
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
      Height          =   250
      Left            =   7770
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   285
      Width           =   1410
   End
   Begin VB.TextBox txtRunningDeposit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6720
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3450
      Visible         =   0   'False
      Width           =   1035
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3915
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00E0E0E0&
      Height          =   1305
      Left            =   690
      TabIndex        =   16
      Top             =   3870
      Width           =   10110
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
         TabIndex        =   39
         Top             =   840
         Width           =   8550
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
         Left            =   6885
         TabIndex        =   8
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
         Left            =   8730
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   7650
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   465
         Width           =   1020
      End
      Begin VB.TextBox txtTitle 
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   465
         Width           =   4125
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
         Left            =   5850
         TabIndex        =   7
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
         TabIndex        =   5
         Top             =   465
         Width           =   1515
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
         Height          =   240
         Left            =   6645
         TabIndex        =   22
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
         Height          =   240
         Left            =   7905
         TabIndex        =   21
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
         Height          =   225
         Left            =   135
         TabIndex        =   20
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
         Height          =   240
         Left            =   2505
         TabIndex        =   19
         Top             =   225
         Width           =   2160
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
         Height          =   240
         Left            =   6090
         TabIndex        =   18
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Delete"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3435
      Width           =   870
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
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "IN PROCESS"
      Top             =   105
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
      Picture         =   "frmSO.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin VB.TextBox txtFax 
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
      Height          =   255
      Left            =   7785
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   645
      Width           =   1395
   End
   Begin VB.TextBox txtAccNum 
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
      Left            =   1050
      TabIndex        =   1
      Top             =   570
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwInvLines 
      Height          =   2280
      Left            =   135
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4022
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
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
         Text            =   " ISBN"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   7584
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
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Print in this currency"
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
      Height          =   285
      Left            =   5580
      TabIndex        =   43
      Top             =   5415
      Width           =   1800
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "To"
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
      Left            =   75
      TabIndex        =   34
      Top             =   600
      Width           =   555
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
      Left            =   75
      TabIndex        =   33
      Top             =   135
      Width           =   555
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Discount"
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
      Height          =   285
      Left            =   1380
      TabIndex        =   32
      Top             =   3465
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Name"
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
      Height          =   240
      Left            =   2445
      TabIndex        =   31
      Top             =   585
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "A/C"
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
      Height          =   240
      Left            =   630
      TabIndex        =   30
      Top             =   615
      Width           =   360
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fax"
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
      Height          =   225
      Left            =   7290
      TabIndex        =   29
      Top             =   645
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7290
      Picture         =   "frmSO.frx":0BA8
      Stretch         =   -1  'True
      Top             =   285
      Width           =   360
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Deposit:"
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
      Left            =   5520
      TabIndex        =   27
      Top             =   3465
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Total:"
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
      Left            =   8325
      TabIndex        =   26
      Top             =   3465
      Width           =   1035
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
      Begin VB.Menu mnuEditNote 
         Caption         =   "Cutomer Note"
      End
   End
End
Attribute VB_Name = "frmPO"
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



Private Sub cboCurr_Click()
    Set oCurrentForeignCurrency = oPC.Configuration.Currencies.FindByDescription(cboCurr)
    oInvoice.CurrencyID = oCurrentForeignCurrency.ID
End Sub

Private Sub chkChargeVAT_Click()
    oInvoice.ShowVAT = (chkChargeVAT = 1)
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
End Sub

Private Sub oInvoice_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
    flgLoading = True
    Me.txtRunningTotal = strtotal
    lngCurrentTotal = lngTotal
    Me.txtRunningDeposit = strTotalDeposit
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

Sub vCanAdd_NobrokenRules()
    Me.cmdNewRows.Enabled = True
End Sub
Private Sub Form_Load()
Dim curTotalDeposit As Currency
    Left = 10
    Top = 10
    Width = 11100
    Height = 6700
    flgLoading = True
    LoadComps
    LoadCurrs
    flgLoading = False
    oInvoice.GetStatus
    Me.cboCurr = oInvoice.ForeignCurrency.Description
    Me.chkChargeVAT = IIf(oInvoice.ShowVAT, 1, 0)
End Sub
Private Sub Form_Initialize()
    Set vCanAdd = New z_BrokenRules
End Sub
Private Sub Form_Unload(Cancel As Integer)
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

Public Sub Component(Optional pInvoice As a_Invoice)
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
    Me.txtTP.SetFocus
    SetEditFrameEnabled False, enNotEditing
    Me.txtAccnum.SetFocus
    vMode = enNotEditing
    flgLoading = False
End Sub
Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
Dim lngColour As Long
    'A is adding, E is editing
    bFrameEnabled = pYesNo   'shared for use in all the form
    
    Me.cboTP.Enabled = Not pYesNo
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
    Me.txtComp.Enabled = Not pYesNo
    
    Me.cmdEnter.Enabled = pYesNo
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
End Sub
Private Sub SetControlsForNew()
    mnuFileCancel.Caption = "&Cancel"
    txtAccnum = ""
    txtFax = ""
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
        'lngILEditingIdx = oInvoice.InvoiceLines.Count
       ' lngSelectedRowIndex = oInvLine.Key 'lvwInvLines.ListItems.Count
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
        vMode = enEditingRow
        fr2.ZOrder 0
        fr1.ZOrder 1
        Me.lvwInvLines.Enabled = True

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
        lstItem.SubItems(6) = .ExtensionF(False)
        If .NonStock = True Then
            lstItem.ForeColor = &H427182
            lstItem.ListSubItems(1).ForeColor = &H427182
            lstItem.ListSubItems(2).ForeColor = &H427182
            lstItem.ListSubItems(3).ForeColor = &H427182
            lstItem.ListSubItems(4).ForeColor = &H427182
            lstItem.ListSubItems(5).ForeColor = &H427182
            lstItem.ListSubItems(6).ForeColor = &H427182
        ElseIf .CopyID = 0 Then
            lstItem.ListSubItems(1).ForeColor = &H706034
            lstItem.ListSubItems(2).ForeColor = &H706034
            lstItem.ListSubItems(3).ForeColor = &H706034
            lstItem.ListSubItems(4).ForeColor = &H706034
            lstItem.ListSubItems(5).ForeColor = &H706034
            lstItem.ListSubItems(6).ForeColor = &H706034
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
    Me.txtPrice = CStr(oInvLine.PriceF(False))
    Me.txtDiscount = CStr(oInvLine.DiscountPercentF)
    txtNote = oInvLine.Note
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    txtPrice.SetFocus
    fr2.ZOrder 1
    fr1.ZOrder 0
    cmdNewRows.Caption = "&Stop edit"
End Sub

'---------Companies code
Private Sub LoadComps()
Dim oCOmp As a_Company
Dim oItem As ListItem
Dim i As Integer
    If oInvoice.CompanyID > 0 Then
        txtComp = oPC.Configuration.Companies(CStr(oInvoice.CompanyID)).CompanyName
    Else
        txtComp = oPC.Configuration.DefaultCompany.CompanyName
        oInvoice.CompanyID = oPC.Configuration.DefaultCompanyID
    End If
End Sub
Private Sub LoadCurrs()
Dim oCurr As a_Currency
Dim oItem As ListItem
Dim i As Integer
    For Each oCurr In oPC.Configuration.Currencies
        Me.cboCurr.AddItem oCurr.Description
    Next
End Sub
Private Sub cboTP_LostFocus()
Dim oCustomer As a_Customer
Dim lngTPID As Long
    If cboTP.ListIndex > -1 Then
        lngTPID = tlCustomer.Key(cboTP)
        Set oCustomer = New a_Customer
        With oCustomer
            .Load lngTPID
            txtFax = .DefaultAddress.Fax
            txtPhone = .Name
            txtAccnum = .AcNo
        End With
        txtTP = ""
        cmdIssue.Enabled = oInvoice.IsDirty
       ' cmdNewRows.SetFocus
        oInvoice.SetCustomer lngTPID
        vCanAdd.RuleBroken "TP", False
    End If
    Set oCustomer = Nothing
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    If oInvoice.Customer Is Nothing Then
        MsgBox "Please enter a customer before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
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
Private Sub txtNote_Validate(Cancel As Boolean)
    Cancel = Not oInvLine.SetNote(txtNote)
End Sub
Private Sub txtNote_LostFocus()
    If flgLoading Then Exit Sub
    txtNote = oInvLine.Note
End Sub

Private Sub mnuEditNote_Click()
Dim ofrm As New frmNote
    ofrm.Component oInvoice
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
Private Sub txtAccNum_Validate(Cancel As Boolean)
Dim lngCustID As Long
Dim bResult As Boolean
    If Len(txtAccnum) > 0 Then
         bResult = oInvoice.SetCustomerFromAccNum(txtAccnum)
        If bResult Then
            Me.txtCustName = oInvoice.Customer.Name
            cboTP.Text = oInvoice.Customer.Name
            txtPhone = oInvoice.Customer.Phone
            txtFax = oInvoice.Customer.DefaultAddress.Fax
            vCanAdd.RuleBroken "TP", False
        Else
            MsgBox "No such account number", , "Can't fetch customer"
            txtAccnum = ""
            Set oCustomer = Nothing
            Cancel = True
        End If
    End If
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
    oInvoice.CompanyID = oPC.Configuration.Companies(i).ID
    iCompIdx = i
End Sub

Private Sub txtCode_LostFocus()
Dim pQty As Integer
Dim pApproID As Long
Dim pNumOfApproLines As Long
    On Error GoTo ERR_Handler
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    Set oProd = New a_Product
    With oProd
        .Load 0, 0, Trim$(txtCode)
        If Len(FixNullsString(.pID)) <> 0 Then
            oInvLine.Title = .TitleAuthorPublisher
            oInvLine.VATRate = oProd.VATRateToUse
            oInvLine.ProductID = .pID
            oInvLine.NonStock = .NonStock
            oInvLine.VATRate = .VATRateToUse
            If Not oProd.DefaultCopy Is Nothing Then
                Set oCurrentCopy = oProd.DefaultCopy
                oInvLine.Price = oCurrentCopy.Price
                oInvLine.CopyID = oCurrentCopy.ID
                
            Else
                oInvLine.Price = oProd.RRP
                MsgBox "There is no copy with this serial number", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
            End If
        Else
            MsgBox "Cannot find book on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
            GoTo EXIT_Handler
        End If
        If .DefaultCopy Is Nothing Then
            oInvLine.Code = .Code
        Else
            oInvLine.Code = .Code
            oInvLine.CodeF = .Code & .DefaultCopy.SerialF
        End If
    End With

    txtTitle = oInvLine.Title
    txtPrice = oInvLine.PriceF(False)
    txtPrice.SetFocus
    
EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub txtDiscount_Validate(Cancel As Boolean)
    If Not oInvLine.SetDiscountPercent(txtDiscount) Then
        Cancel = True
    End If
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oInvLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
End Sub


Private Sub txtTP_Validate(Cancel As Boolean)
    If Len(txtTP) <> 0 Then
        Set tlCustomer = Nothing
        Set tlCustomer = New z_TextList
        tlCustomer.Load ltCustomer, Me.txtTP
        LoadCombo Me.cboTP, tlCustomer
    End If
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
        txtStatus = .status
        SetIssueButtonCaption
        txtAccnum = .TPAccNum
        cboTP = Trim$(.TPName)
        txtPhone = .TPPhone
        txtPhone = .TPPhone
        txtFax = .TPFax
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
Dim ViewOrPrint As PreviewPrint
Dim strResult As String

    If oInvoice.status = stInPROCESS Then
        If MsgBox("Issue this invoice?.  Confirm.", vbYesNo + vbQuestion, "Papyrus Invoicing Status") = vbNo Then
            Exit Sub
        End If
    End If
    If Me.chkProforma = 0 Then  'Unchecked
        oInvoice.setStatus stCOMPLETE
    Else
        oInvoice.setStatus stPROFORMA
    End If
    
    strResult = oInvoice.post
    
    If MsgBox("Do you want to print?", vbYesNo + vbQuestion, "Papyrus Invoicing Status") = vbYes Then
        oInvoice.PrintInvoice True, True  'iCurrency, blnRoundedUp
    End If
    Unload Me
End Sub
Private Sub cmdSave_Click()
    oInvoice.setStatus stInPROCESS
    SaveInvoice
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
Private Sub cmdDelete_Click()
    RemoveInvoiceLine
End Sub


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

Private Sub lvwInvLines_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
Private Sub txtDeposit_GotFocus()
    AutoSelect Controls("txtDeposit")
End Sub
Private Sub txtDiscount_GotFocus()
    AutoSelect Controls("txtDiscount")
End Sub
Private Sub txtPrice_GotFocus()
    AutoSelect Controls("txtPrice")
End Sub
Private Sub SetIssueButtonCaption()
        If oInvoice.status = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oInvoice.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
End Sub
Private Sub cmdAddrNotes_Click()
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.Component oInvoice
    frm.Show vbModal
    
End Sub
Private Sub txtAccNum_LostFocus()
    txtAccnum = UCase(txtAccnum)
End Sub


Private Sub txtDiscount_LostFocus()
 '   txtDiscount = oInvLine.DiscountPercentF
End Sub

Private Sub txtGeneralDiscount_LostFocus()
    txtGeneralDiscount = oInvoice.InvoiceDiscountRateF
End Sub

Private Sub txtGeneralDiscount_change()
    oInvoice.SetGeneralDiscount (txtGeneralDiscount)
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

