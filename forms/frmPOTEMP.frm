VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPOTEMP 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase order"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmPOTEMP.frx":0000
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
   Begin VB.CommandButton cmdSelectTP 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Find supplier"
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
      Picture         =   "frmPOTEMP.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5370
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
      Height          =   1050
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   5295
      Width           =   3390
   End
   Begin VB.TextBox txtName 
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   75
      Width           =   1815
   End
   Begin VB.Frame fr2 
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   675
      TabIndex        =   21
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
         TabIndex        =   22
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
      Picture         =   "frmPOTEMP.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5370
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
      TabIndex        =   18
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3915
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00E0E0E0&
      Height          =   1305
      Left            =   705
      TabIndex        =   11
      Top             =   3855
      Width           =   10110
      Begin VB.ComboBox cboRef 
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
         Height          =   345
         Left            =   1665
         TabIndex        =   39
         Top             =   465
         Width           =   1140
      End
      Begin VB.ComboBox cboCat 
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
         Height          =   345
         Left            =   2760
         TabIndex        =   37
         Top             =   480
         Width           =   1605
      End
      Begin VB.ComboBox cboDeal 
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
         Height          =   345
         Left            =   6030
         TabIndex        =   35
         Top             =   465
         Width           =   1395
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
         Left            =   4350
         TabIndex        =   32
         Top             =   480
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
         Left            =   8145
         Locked          =   -1  'True
         TabIndex        =   31
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
         TabIndex        =   6
         Top             =   825
         Width           =   4365
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
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
         Left            =   7425
         Locked          =   -1  'True
         TabIndex        =   5
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
         Left            =   9180
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   7
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
         TabIndex        =   12
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
         Left            =   5070
         TabIndex        =   4
         Top             =   480
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
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total"
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
         Left            =   8280
         TabIndex        =   40
         Top             =   225
         Width           =   630
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Category"
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
         Left            =   3030
         TabIndex        =   38
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deal"
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
         Left            =   6360
         TabIndex        =   36
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ref."
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
         TabIndex        =   34
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
         Left            =   4110
         TabIndex        =   33
         Top             =   240
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
         TabIndex        =   30
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
         Left            =   7365
         TabIndex        =   15
         Top             =   225
         Width           =   630
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
         TabIndex        =   14
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
         Left            =   5115
         TabIndex        =   13
         Top             =   240
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
      TabIndex        =   10
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
      Picture         =   "frmPOTEMP.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2580
      Left            =   135
      TabIndex        =   16
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      Height          =   270
      Left            =   3435
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   645
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
      Top             =   90
      Width           =   1920
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Left            =   3465
      TabIndex        =   20
      Top             =   60
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "S/code"
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
      Left            =   15
      TabIndex        =   19
      Top             =   165
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3735
      Picture         =   "frmPOTEMP.frx":0BA8
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
Attribute VB_Name = "frmPOTEMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPO As a_PO
Attribute oPO.VB_VarHelpID = -1
Dim WithEvents oPOLine As a_POL
Attribute oPOLine.VB_VarHelpID = -1
Dim oTP As a_Supplier
Dim oProd As a_Product
Dim oCurrentCopy
Dim bValidCN As Boolean
Dim bValidCNLine As Boolean
Dim tlSupplier As z_TextList
Dim lngCurrentExtension As Long
Dim lngCurrentTotal As Long
Dim lngCurrentDepositTotal As Long
Dim lngCurrentVATTotal As Long
Dim tlCategories As z_TextList
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
    oPOLine.Discount = cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem(0), 2)
    oPOLine.DEALID = cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem(0), 1)
End Sub



Private Sub cmdSelectTP_Click()
Dim lngTPID As Long
Dim frm As frmBrowseSuppliers2
    Set frm = New frmBrowseSuppliers2
    frm.Show vbModal
    lngTPID = frm.TPID
    If oPO.SetTP(lngTPID) Then
        With oPO.Supplier
            txtPhone = .Phone
            txtAccnum = .AcNo
            txtName = .Name
            lblAddBill.Caption = .DefaultAddress.AddressShort
            lblAddDel.Caption = .DefaultAddress.AddressShort
            LoadDeals
        End With
        vCanAdd.RuleBroken "TP", False
    End If

End Sub

Private Sub Form_Terminate()
    Set tlCategories = Nothing
'        Set frmOL = Nothing
End Sub

Private Sub Label5_DblClick()
Dim frm As frmSupplierPreview
    Set frm = New frmSupplierPreview
    frm.component oPO.Supplier
    frm.Show
End Sub

Private Sub lvwLines_AfterLabelEdit(cancel As Integer, NewString As String)
cancel = True
End Sub

Private Sub mnuAddresses_Click()
Dim frm As frmInvAddr
    Set frm = New frmInvAddr
    frm.component oPO
    frm.Show vbModal
    lblAddBill.Caption = oPO.BillToAddress.AddressShort
    lblAddDel.Caption = oPO.GoodsToAddress.AddressShort

End Sub

Private Sub mnuDel_Click()
    RemoveDetailLine
End Sub


Private Sub mnuPrint_Click()
Dim frm As frmPrintingOptions_CN
    Set frm = New frmPrintingOptions_CN
    frm.Show vbModal

End Sub

Private Sub oPO_Valid(pMsg As String)
    bValidCN = (pMsg = "")
    cmdIssue.Enabled = (bValidCN And oPO.POLines.Count > 0)
    cmdSave.Enabled = bValidCN
    Me.txtError = pMsg
End Sub

Sub oPOLine_ExtensionChange(lngExtension As Long, strExtension As String)
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
End Sub

Private Sub oPOLine_Valid(Msg As String)
    cmdEnter.Enabled = (Msg = "")
    txtError = Msg
End Sub

Private Sub oPO_TotalChange(lngTotal As Long, strtotal As String, lngTotalDeposit As Long, strTotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
    flgLoading = True
    Me.txtRunningTotal = strtotal
    lngCurrentTotal = lngTotal
'    Me.txtRunningDeposit = strTotalDeposit
    lngCurrentDepositTotal = lngTotalDeposit
    lngCurrentVATTotal = lngTotalVAT
    flgLoading = False
End Sub

Private Sub oPO_Reloadlist()
    LoadListView
End Sub
Private Sub oPO_Dirty(pVal As Boolean)
If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
End Sub
Private Sub oPO_CurrRowStatus(pMsg As String)
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
    oPO.GetStatus
    SetLvw
    SetEditFrameEnabled False, enNotEditing
    vMode = enNotEditing
    SetupcboDeal
    LoadCombo cboCAT, tlCategories
    flgLoading = False
End Sub
Private Sub Form_Initialize()
    Set vCanAdd = New z_BrokenRules
    Set tlCategories = New z_TextList
    tlCategories.Load ltCategory
End Sub
Private Sub Form_Unload(cancel As Integer)
    If oPO.IsEditing Then oPO.CancelEdit
    
    Set oTP = Nothing
    Set oCurrentCopy = Nothing
    Set oPO = Nothing
    Set tlSupplier = Nothing
    Set oPOLine = Nothing
End Sub

Public Sub component(Optional pPO As a_PO)
    flgLoading = True
    If pPO Is Nothing Then
        Set oPO = New a_PO
        oPO.BeginEdit
        Me.lvwLines.Enabled = False
        SetControlsForNew
        vCanAdd.RuleBroken "TP", True
    Else
        Set oPO = pPO
        oPO.BeginEdit
        LoadSupplier
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
    Me.cmdSave.Enabled = (Not pYesNo) And bValidCN And oPO.IsDirty
    
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
    oPOLine.ApplyEdit
'    oPOLine.SetAsNEW
    oPOLine.BeginEdit

    If vMode = enAddingRow Then
        lvwLines.ListItems.Add 1, oPOLine.Key
        LoadListViewLine oPOLine.Key, Me.lvwLines.ListItems(1)
        Set oPOLine = oPO.POLines.Add
        oPOLine.SetQty 1
        oPOLine.TRID = oPO.TRID
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
        Set oPOLine = oPO.POLines.Add
        oPOLine.TRID = oPO.TRID
    End If

    ClearLineControls
End Sub
Private Sub LoadListView()
Dim lstItem As ListItem
Dim i As Long
    On Error GoTo ERR_Handler
    lvwLines.ListItems.Clear
    For i = 1 To oPO.POLines.Count
        Set lstItem = lvwLines.ListItems.Add
        Set oPOLine = oPO.POLines(i)
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
    With oPOLine
        lstItem.Text = .ProductCodeF
        If lstItem.Key = "" Then lstItem.Key = i
        lstItem.SubItems(1) = .TitleAuthor
        lstItem.SubItems(2) = .Qty
        lstItem.SubItems(3) = .Ref
        lstItem.SubItems(4) = .PriceF
        lstItem.SubItems(5) = .DiscountF  ' Format(.Discount, "##0.0%")
        lstItem.SubItems(6) = .ExtensionInclDeposit
    End With
End Sub
Private Sub lvwLines_DblClick()
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    lngEditingIdx = lvwLines.SelectedItem.Key
    Set oPOLine = oPO.POLines(lngEditingIdx)
'    If oPOLine.Product.DefaultCopy Is Nothing Then
'        LoadMatchingInvoices oPO.Supplier.ID, oPOLine.Product.pID
'    Else
'        LoadMatchingInvoices oPO.Supplier.ID, "", oPOLine.Product.DefaultCopy.ID
'    End If
'    SetcboMatchToID (oPOLine.POLID)
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    
    txtTitle = oPOLine.Title
    txtPrice = oPOLine.PriceF
    txtQty = oPOLine.QtyF
    txtDiscount = oPOLine.Discount
  '  txtPrice.SetFocus
    txtCode = oPOLine.ProductCode
    AutoSelect txtPrice
    
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    txtPrice.SetFocus
    fr2.ZOrder 1
    fr1.ZOrder 0
    cmdNewRows.Caption = "&Stop edit"
    
End Sub
Private Sub SetcboDealToID(pID As Long)
'    If pID > 0 Then
'        cboDeal.Items.SelectItem(cboDeal.Items.FindItem(pID, 4)) = True
'    End If
End Sub
'---------Companies code
'Private Sub LoadComps()
'Dim oCNmp As a_CNmpany
'Dim oItem As ListItem
'Dim i As Integer
'    If oPO.CompanyID > 0 Then
'        txtComp = oPC.Configuration.Companies(CStr(oPO.CompanyID)).CompanyName
'    Else
'        txtComp = oPC.Configuration.DefaultCompany.CompanyName
'        oPO.CompanyID = oPC.Configuration.DefaultCompanyID
'    End If
'End Sub

Private Sub cboTP_Validate(cancel As Boolean)
    If oPO.Supplier Is Nothing Then
        MsgBox "Please enter a Supplier before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        cancel = True
    End If
End Sub
'-------End Compsny code
'Private Sub txtOrdernum_Validate(Cancel As Boolean)
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oPOLine.COLineCode = txtOrdernum
'    If Err Then
'      Beep
'      intPos = txtOrdernum.SelStart
'      txtOrdernum = oPOLine.COLineCode
'      txtOrdernum.SelStart = intPos - 1
'    End If
'
'End Sub

Private Sub txtNote_Change()
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oPOLine.SetNote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oPOLine.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub
Private Sub txtNote_Validate(cancel As Boolean)
    cancel = Not oPOLine.SetNote(txtNote)
End Sub
Private Sub txtNote_LostFocus()
    If flgLoading Then Exit Sub
    txtNote = oPOLine.Note
End Sub

Private Sub mnuEditNote_Click()
Dim ofrm As New frmNote
    ofrm.component oPO
    ofrm.Show vbModal
    Unload ofrm
    Set ofrm = Nothing
End Sub

Private Sub mnuFileCancel_Click()
    If oPO.IsDirty Then
        oPO.CancelEdit
    End If
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    oPO.CancelEdit
    Unload Me
End Sub

Private Sub mnuFileOK_Click()
'    cmdOK_Click
End Sub

Private Sub mnuFilePrint_Click()
    cmdIssue_Click
End Sub
Private Sub mnuFileVoid_Click()
    oPO.setStatus stVOID
    txtStatus = "Void"
End Sub
Private Sub txtAccNum_Validate(cancel As Boolean)
Dim lngCustID As Long
Dim bResult As Boolean
    If Len(txtAccnum) > 0 Then
        bResult = oPO.SetSupplierFromAccNum(txtAccnum)
        If bResult Then
            With oPO.Supplier
                txtName = .Name
                txtPhone = .Phone
                lblAddBill.Caption = .DefaultAddress.AddressShort
                lblAddDel.Caption = .DefaultAddress.AddressShort
                LoadDeals
            End With
            vCanAdd.RuleBroken "TP", False
        Else
            MsgBox "No such account number", , "Can't fetch Supplier"
            txtAccnum = ""
            Set oTP = Nothing
            cancel = True
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
'    oPO.CompanyID = oPC.Configuration.Companies(i).ID
'    iCompIdx = i
'End Sub

Private Sub txtCode_Validate(cancel As Boolean)
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean

On Error GoTo ERR_Handler
    
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    bOK = oPOLine.SetLineProduct("", txtCode)
    LoadCustomerRefs
'    If Not oPOLine.Product.DefaultCopy Is Nothing Then
'        LoadMatchingInvoices oPO.Supplier.ID, "", oPOLine.Product.DefaultCopy.ID
'    ElseIf Not oPOLine.Product Is Nothing Then
'        LoadMatchingInvoices oPO.Supplier.ID, oPOLine.Product.pID
'    End If
    
    
    If bOK Then
        txtTitle = oPOLine.Title
        txtPrice = oPOLine.PriceF
        txtQty = oPOLine.QtyF
        txtDiscount = oPOLine.Discount
        txtPrice.SetFocus
        txtCode = oPOLine.ProductCode
        AutoSelect txtPrice
    Else
        MsgBox "Cannot find book on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        cancel = True
        GoTo EXIT_Handler
    End If
    oPOLine.GetStatus

EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub
'Private Sub LoadMatchingInvoices(pTPID As Long, pPID As String, Optional pPIID As Long, Optional pCode As String)
'Dim odPO As d_Invoice
'Dim i As Integer
'Dim bReturned As Boolean
'
'    Set oPOS = New c_POs
'    oPOS.Load bReturned, pTPID, , , pPID, pCode, pPIID
'    If oPO.Count > 0 Then 'There are invoices for this item
'        cboMatch.BeginUpdate
'        ReDim ar(4, oPO.Count - 1)
'        cboMatch.Items.RemoveAllItems
'        i = 0
'        For Each odPO In oPO
'            ar(0, i) = odPO.TDateFormatted
'            ar(1, i) = odPO.InvoiceNumber
'            ar(2, i) = odPO.Qty
'            ar(3, i) = odPO.PriceF
'            ar(4, i) = odPO.InvoiceLineID
'            i = i + 1
'        Next
'        cboMatch.PutItems ar
'        cboMatch.EndUpdate
'    End If
' '   fClearTranslucency frmOL.hwnd
'End Sub
Private Sub cboMatch_Click()
 '   oPOLine.INVLineID = oPO.Item(1).InvoiceLineID
End Sub

Private Sub RemoveDetailLine()
Dim i As Integer
Dim iMax As Integer
    iMax = lvwLines.ListItems.Count
    For i = iMax To 1 Step -1
        If lvwLines.ListItems(i).Selected Then
            oPO.POLines.Remove lvwLines.ListItems(i).Key
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

Private Sub LoadSupplier()
    With oPO
        txtStatus = .statusF
        SetIssueButtonCaption
        txtAccnum = .TPAccNum
        txtPhone = .TPPhone
        txtPhone = .TPPhone
    End With
End Sub


Private Sub SaveCO()
On Error GoTo ERR_Handler
    
    oPO.post
    
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
    oPO.Load oPO.TRID, False
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
Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmCNPreview

    If oPO.status = stInProcess Then
        If MsgBox("Issue this order?.  Confirm.", vbYesNo + vbQuestion, "Papyrus Invoicing Status") = vbNo Then
            Exit Sub
        End If
    End If
    oPO.setStatus stISSUED
    
    strResult = oPO.post
    Set frm = New frmCNPreview
    frm.ComponentObject oPO
    frm.Show
    Unload Me
End Sub
Private Sub cmdSave_Click()
    oPO.setStatus stInProcess
    SaveCO
    oPO.BeginEdit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    oPO.CancelEdit
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
'    cboDeal.Items.RemoveAllItems
'    cboDeal.SetFocus
    Me.cmdNewRows.SetFocus
    flgLoading = False
End Sub

Private Sub lvwLines_BeforeLabelEdit(cancel As Integer)
    cancel = True
End Sub
'Private Sub txtETA_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oPOLine.SetETA(txtETA) Then
'        Cancel = True
'    End If
'End Sub
'Private Sub txtETA_GotFocus()
'    AutoSelect Controls("txtETA")
'End Sub
'
'Private Sub txtETA_LostFocus()
'    txtETA = oPOLine.ETAF
'End Sub

Private Sub txtPrice_GotFocus()
    AutoSelect Controls("txtPrice")
End Sub
Private Sub txtPrice_Validate(cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetPrice(txtPrice) Then
        cancel = True
    End If
End Sub
Private Sub txtPrice_LostFocus()
  '  txtPrice = oPOLine.PriceF
End Sub
Private Sub txtQty_GotFocus()
    AutoSelect Controls("txtQty")
End Sub
Private Sub txtQty_Validate(cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetQty(txtQty) Then
        cancel = True
    End If
End Sub
Private Sub txtQty_LostFocus()
  '  txtQty = oPOLine.QtyF
End Sub
Private Sub txtDiscount_Validate(cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetDiscount(txtDiscount) Then
        cancel = True
    End If
End Sub
Private Sub txtDiscount_LostFocus()
  '  txtDiscount = oPOLine.DiscountF
End Sub
Private Sub txtDiscount_GotFocus()
    AutoSelect Controls("txtDiscount")
End Sub
'Private Sub txtDeposit_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oPOLine.SetDeposit(txtdeposit) Then
'        Cancel = True
'    End If
'End Sub
'Private Sub txtDeposit_LostFocus()
' '   txtdeposit = oPOLine.DepositF
'End Sub
'Private Sub txtDeposit_GotFocus()
'    AutoSelect Controls("txtDeposit")
'End Sub


Private Sub SetIssueButtonCaption()
        If oPO.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oPO.IsDirty Then
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
Sub SetupcboDeal()
        cboDeal.BeginUpdate
        cboDeal.WidthList = 180
        cboDeal.HeightList = 162
        cboDeal.AllowSizeGrip = True
        cboDeal.AutoDropDown = True
        
        cboDeal.Columns.Add "Description"
        cboDeal.Columns.Add "Discount"
        cboDeal.Columns.Add "DLID"
        cboDeal.Columns(0).Width = 70
        cboDeal.Columns(1).Width = 70
        cboDeal.Columns(2).Width = 0
        cboDeal.BackColorLock = Me.BackColor
        
        cboDeal.EndUpdate
End Sub

Private Sub LoadDeals()
Dim oDL As a_Deal
Dim i As Integer
Dim ar()

    i = 0
    If oPO.Supplier.Deals.Count < 1 Then
        Exit Sub
    End If
    For Each oDL In oPO.Supplier.Deals
        i = i + 1
    Next
    
    ReDim ar(2, i - 1)
    i = 0
    cboDeal.BeginUpdate
    cboDeal.Items.RemoveAllItems
    For Each oDL In oPO.Supplier.Deals
        ar(0, i) = oDL.Description
        ar(1, i) = oDL.DiscountF
        ar(2, i) = oDL.ID
        i = i + 1
    Next
    cboDeal.PutItems ar
    cboDeal.EndUpdate
End Sub
