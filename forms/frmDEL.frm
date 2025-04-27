VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "COOLBU~1.OCX"
Begin VB.Form frmdel 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Goods received note"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmDEL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11595
   Begin VB.TextBox txtCurrencyRates 
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
      ForeColor       =   &H00008000&
      Height          =   250
      Left            =   75
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2835
      Width           =   7230
   End
   Begin VB.CommandButton cmdBatch 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Batch"
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
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5685
      Width           =   690
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2340
      Left            =   90
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   450
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4128
      SortKey         =   1
      View            =   3
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Firm"
         Object.Width           =   883
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "SS"
         Object.Width           =   882
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
      Picture         =   "frmDEL.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5280
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
      Left            =   1050
      MultiLine       =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5235
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
      Height          =   690
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5340
      Width           =   795
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
      Picture         =   "frmDEL.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1110
   End
   Begin VB.TextBox txtRunningTotal 
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
      Height          =   360
      Left            =   4515
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5685
      Width           =   1530
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   2145
      Left            =   75
      TabIndex        =   13
      Top             =   3045
      Width           =   10710
      Begin VB.TextBox txtSP 
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
         Left            =   2685
         TabIndex        =   6
         Top             =   1215
         Width           =   1260
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
         Left            =   9420
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   600
         Width           =   255
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
         Height          =   840
         Left            =   6810
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1215
         Width           =   2610
      End
      Begin VB.TextBox txtQtyFirm 
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
         Left            =   90
         TabIndex        =   3
         Top             =   1215
         Width           =   615
      End
      Begin VB.TextBox txtQtySS 
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
         Left            =   735
         TabIndex        =   4
         Top             =   1215
         Width           =   615
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
         Left            =   4005
         TabIndex        =   7
         Top             =   1215
         Width           =   885
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
         Height          =   555
         Left            =   9645
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1530
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
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
         Height          =   390
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1215
         Width           =   1215
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
         Height          =   240
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1785
         Width           =   6600
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
         Left            =   1380
         TabIndex        =   5
         Top             =   1230
         Width           =   1260
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
         TabIndex        =   1
         Top             =   435
         Width           =   1725
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboMatch 
         Height          =   375
         Left            =   2130
         OleObjectBlob   =   "frmDEL.frx":30B6
         TabIndex        =   2
         Top             =   435
         Width           =   7260
      End
      Begin VB.Label Label10 
         BackColor       =   &H00D3D3CB&
         Caption         =   "S.P."
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
         Left            =   3135
         TabIndex        =   37
         Top             =   945
         Width           =   705
      End
      Begin VB.Label Label3 
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
         Height          =   210
         Left            =   6840
         TabIndex        =   32
         Top             =   975
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Firm"
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
         Left            =   210
         TabIndex        =   30
         Top             =   945
         Width           =   720
      End
      Begin VB.Label lblWants 
         BackColor       =   &H00D3D3CB&
         Caption         =   "fulfilment of . . ."
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
         Left            =   2160
         TabIndex        =   25
         Top             =   210
         Width           =   1845
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
         Left            =   3855
         TabIndex        =   19
         Top             =   945
         Width           =   1350
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
         Left            =   5310
         TabIndex        =   18
         Top             =   945
         Width           =   990
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
         Left            =   105
         TabIndex        =   17
         Top             =   195
         Width           =   1410
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SS"
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
         Left            =   930
         TabIndex        =   16
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label6 
         BackColor       =   &H00D3D3CB&
         Caption         =   "R.R.P."
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
         Left            =   1725
         TabIndex        =   15
         Top             =   945
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issu&e"
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
      Left            =   9660
      Picture         =   "frmDEL.frx":4460
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5280
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   345
      Left            =   30
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   60
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   609
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
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   4980
      TabIndex        =   34
      Top             =   5400
      Width           =   720
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Left            =   7230
      TabIndex        =   31
      Top             =   30
      Width           =   525
   End
   Begin VB.Label txtFax 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   7950
      TabIndex        =   29
      Top             =   45
      Width           =   2250
   End
   Begin VB.Label txtPhone 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   210
      Left            =   4995
      TabIndex        =   28
      Top             =   45
      Width           =   2250
   End
   Begin VB.Label txtSuppname 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   1065
      TabIndex        =   27
      Top             =   45
      Width           =   3135
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   210
      TabIndex        =   22
      Top             =   60
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4380
      Picture         =   "frmDEL.frx":45AA
      Stretch         =   -1  'True
      Top             =   60
      Width           =   360
   End
End
Attribute VB_Name = "frmdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oDel As a_Delivery
Attribute oDel.VB_VarHelpID = -1
Dim WithEvents oDELL As a_DeliveryLine
Attribute oDELL.VB_VarHelpID = -1
Dim oSupplier As a_Supplier
Private tlSections As z_TextList

Dim bValidDEL As Boolean
Dim bValidDELLine As Boolean
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
Dim strDELErrMsg As String
Dim strDELLErrMsg As String

'Private Sub Text1_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then KeyAscii = 0
'End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        keybd_event &H9, 1, 1, 1
'    End If
'End Sub


Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oDel.statusF = "IN PROCESS" And oDel.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SetMenu"
End Sub



Public Sub Component(pCancel As Boolean, Optional pTPID As Long, Optional pDel As a_Delivery)
    On Error GoTo errHandler
Dim frm As frmHeader_GRN

    pCancel = False
    flgLoading = True
    If pDel Is Nothing Then
        Set oDel = New a_Delivery
        oDel.BeginEdit
        oDel.SetStatus stInProcess
        oDel.SetSupplier pTPID
        LoadSupplier
        Set frm = New frmHeader_GRN
        frm.Component oDel
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Unload Me
            pCancel = True
            Exit Sub
        End If
        Unload frm
        Me.lvw.Enabled = False
        If oDel.Supplier.Deals.Count < 1 Then
            MsgBox "There are no deals for this supplier. You cannot continue"
            pCancel = True
        End If
        lvw.Height = 2340
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        Set oDELL = oDel.DeliveryLines.Add
        oDELL.SetQtyFirm 1
        ClearLineControls
        oDel.GetStatus
        mSetfocus txtCode
    Else
        Set oDel = pDel
        oDel.BeginEdit
        LoadSupplier
        LoadListView
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        Me.lvw.Enabled = True
        lvw.Height = 4720
        vMode = enNotEditing
        SetEditFrameEnabled False, vMode
        ClearLineControls
    End If
    oDel.GetStatus
    If oDel.isFOreignCurrency Then
        Me.txtRunningTotal = oDel.TotalLessDiscExtF(False)
        txtCurrencyRates = oDel.CurrencyConversionAsText & "     Value is : " & oDel.TotalLessDiscExtF(True)
        txtCurrencyRates.Visible = True
    Else
        Me.txtRunningTotal = oDel.TotalLessDiscExtF(False)
        txtCurrencyRates.Visible = False
    End If
    SetMenu
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Component(pCancel,pTPID,pDel)", Array(pCancel, pTPID, pDel)
End Sub


Private Sub cboMatch_SelectionChanged()
    On Error GoTo errHandler
Dim h As HITEM
Dim tmp As String

    If cboMatch.Items.SelectCount > 0 Then
   '     oDELL.SetQtySS cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 3)
   '     oDELL.SetQtyFirm cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2)
        
                tmp = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 3)
                oDELL.SetQtySS Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, "(") - 1)
                tmp = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2)
                oDELL.SetQtyFirm Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, "(") - 1)
        
        oDELL.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 10)
        oDELL.SetPrice cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9)
        
        oDELL.POLID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 8)
        oDELL.COLID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 11)
        oDELL.Ref = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 6)
    Else
        oDELL.SetQtySS 0
        oDELL.SetQtyFirm 0
        oDELL.SetDiscount 0
        oDELL.SetPrice 0
        oDELL.POLID = 0
        oDELL.COLID = 0
        oDELL.Ref = ""
    End If
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oDELL.PriceF(oDel.isFOreignCurrency)
        txtSP = oDELL.PriceSell
    Else
        txtPrice = oDELL.Price(oDel.isFOreignCurrency)
        txtSP = oDELL.PriceSell
    End If
    txtQtyFirm = oDELL.QtyFirmF
    txtQtySS = oDELL.QtySSF
    txtDiscount = oDELL.DiscountF
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cboMatch_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdBatch_Click()
    On Error GoTo errHandler
Dim frm As New frmHeader_GRN
    frm.Component oDel
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdBatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancelMatch_Click()
Dim i As Integer

    On Error GoTo errHandler
    If cboMatch.Items.ItemCount = 0 Then Exit Sub
    oDELL.POLID = 0
    For i = 0 To cboMatch.Items.ItemCount - 1
        cboMatch.Items.SelectItem(cboMatch.Items(i)) = False
    Next
    mSetfocus txtQtyFirm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdCancelMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As New frmSupplierPreview
    
    If oDel.Supplier.ID > 0 Then
        frm.Component oDel.Supplier
        frm.Show
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbSupp_Click()
    On Error GoTo errHandler
Dim frm As frmSupplierPreview
    If oDel.Supplier.Name = "" Then Exit Sub
    Set frm = New frmSupplierPreview
    frm.Component oDel.Supplier
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cbSupp_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdFulfilments_Click()
    On Error GoTo errHandler
   ' ReconcileWithCOs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdFulfilments_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadNewSupplier(plngTPID As Long)
    On Error GoTo errHandler
    If oDel.SetSupplier(plngTPID) Then
        With oDel.Supplier
            txtPhone = .OrderToAddress.Phone
            txtSuppname = .NameAndCode(18)
            txtFax = .OrderToAddress.Fax
        End With
        vCanAdd.RuleBroken "TP", False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.LoadNewSupplier(plngTPID)", plngTPID
End Sub

Private Sub cmdNote_Click()
    On Error GoTo errHandler
Dim frm As New frmILNote
    frm.Component oDELL
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdNote_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set vCanAdd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.mnuDelLine"
End Sub
Private Sub oDEL_ValidToSave(pOK As Boolean)
    On Error GoTo errHandler
    cmdSave.Enabled = (pOK And oDel.DeliveryLines.Count > 0 And vMode = enNotEditing And oDel.IsDirty)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDEL_ValidToSave(pOK)", pOK, EA_NORERAISE
    HandleError
End Sub


Private Sub oDEL_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidDEL = (pMsg = "")
    cmdIssue.Enabled = (bValidDEL And oDel.DeliveryLines.Count > 0 And vMode = enNotEditing)
 '   cmdSave.Enabled = (bValidDEL And oDel.DeliveryLines.Count > 0 And vMode = enNotEditing)
    strDELErrMsg = pMsg
    If vMode = enNotEditing Then
        txtError = strDELErrMsg
    Else
        txtError = strDELLErrMsg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDEL_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub

Sub oDELL_ExtensionChange(lngExtension As Long, strExtension As String)
    On Error GoTo errHandler
MsgBox "Is this being used?"
    flgLoading = True
    Me.txtTotal = strExtension
    flgLoading = False
    lngCurrentExtension = lngExtension
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDELL_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
         strExtension), EA_NORERAISE
    HandleError
End Sub

Private Sub oDELL_Valid(msg As String)
    On Error GoTo errHandler
    Me.cmdEnter.Enabled = (msg = "")
    strDELLErrMsg = msg
    If vMode = enNotEditing Then
        txtError = strDELErrMsg
    Else
        txtError = strDELLErrMsg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDELL_Valid(Msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oDEL_TotalChange(strtotal As String, strTotalForeign As String)
    On Error GoTo errHandler
    flgLoading = True
    If oDel.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
        Me.txtRunningTotal = strtotal
    Else
        Me.txtRunningTotal = strTotalForeign
        txtCurrencyRates = oDel.CurrencyConversionAsText & "     Value is : " & strtotal
    End If
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDEL_TotalChange(strtotal,strTotalForeign)", Array(strtotal, strTotalForeign), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub oDEL_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDEL_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oDEL_Dirty(pVal As Boolean)
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
    ErrorIn "frmdel.oDEL_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oDEL_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.oDEL_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub



Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
    If txtCode > "" Then SendKeys "{DOWN}", True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDiscount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtyFirm_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtyFirm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtQtyFirm_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySS_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtySS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtQtySS_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtyFirm_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetQtyFirm(txtQtyFirm) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtQtyFirm_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySs_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE, , "Rowcount,ODELL=NOTHING,oDEL=Nothing", Array(oDel.DeliveryLines.Count, oDel Is Nothing, oDELL Is Nothing)
    HandleError
End Sub





Private Sub txtSP_Validate(Cancel As Boolean)
Dim lngTmp As Long

    If ConvertToLng(Trim(txtSP), lngTmp) Then
        oDELL.SetPriceSell Trim(txtSP)
    End If
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
    ErrorIn "frmdel.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim curTotalDeposit As Currency
Dim strAddress As String
    SetupcboMatch
    left = 10
    top = 10
    Width = 11100
    Height = 6700
    flgLoading = True
    flgLoading = False
    oDel.GetStatus
    SetLvw
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.Component oDel.Memo
    ofrm.Show vbModal
    oDel.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.mnuMemo"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    
    Set vCanAdd = New z_BrokenRules
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oDel.IsEditing Then oDel.CancelEdit
    UnsetMenu
    
    Set oSupplier = Nothing
    Set oDel = Nothing
    Set oDELL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
    txtCurrencyRates.Enabled = pYesNo
    Me.txtDiscount.Enabled = pYesNo
    Me.txtPrice.Enabled = pYesNo
    Me.txtQtyFirm.Enabled = pYesNo
    Me.txtQtySS.Enabled = pYesNo
    Me.txtTitle.Enabled = pYesNo
    Me.txtTotal.Enabled = pYesNo
    Me.cboMatch.Enabled = pYesNo
    Me.cmdEnter.Enabled = pYesNo And eMode <> enNotEditing
    Me.cmdCancel.Enabled = Not pYesNo
    Me.cmdIssue.Enabled = (Not pYesNo) And bValidDEL
    Me.cmdSave.Enabled = (Not pYesNo) And bValidDEL And oDel.IsDirty
    
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
    Me.txtDiscount.BackColor = lngColour
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim i As Integer
Dim iDiff As Integer

    If oDELL Is Nothing Then Exit Sub
    If oDel Is Nothing Then Exit Sub
    
    If cboMatch.Items.SelectCount > 0 Then
        i = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 12)
        If (oDELL.QtyFirm + oDELL.QtySS) > i Then
            If i > 1 Then
                MsgBox "There are only " & i & " items outstanding on this purchase order line."
            Else
                MsgBox "There is only one item outstanding on this purchase order line."
            End If
            Exit Sub
        End If
    End If
    
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        If txtCode.Enabled Then mSetfocus txtCode
        Exit Sub
    End If
    oDELL.ApplyEdit
    oDELL.BeginEdit

    If vMode = enAddingRow Then
        lvw.ListItems.Add 1, oDELL.Key
        LoadListViewLine oDELL.Key, Me.lvw.ListItems(1)
        Set oDELL = oDel.DeliveryLines.Add
       '''' oDELL.SetQtyFirm 1
        oDELL.trid = oDel.trid
        ClearLineControls
        mSetfocus txtCode
    ElseIf vMode = enEditingRow Then
        LoadListViewLine lngSelectedRowIndex, Me.lvw.ListItems(lngSelectedRowIndex)
        ClearLineControls
        lvw.Enabled = True
        lvw.Height = 4720
        vMode = enNotEditing
        SetEditFrameEnabled False, vMode
        cmdNewRows.Caption = "&Add"
        fr1.ZOrder 1
        txtCurrencyRates.ZOrder 1
    End If
    oDel.CalculateTotals
 '   oDel.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdEnter_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
'    'Change state
'    If vMode = enEditingRow Then
'        vMode = enNotEditing
'    ElseIf vMode = enAddingRow Then
'        vMode = enNotEditing 'enEditingRow
'    ElseIf vMode = enNotEditing Then
'        vMode = enAddingRow
'    End If
    
    If vMode = enEditingRow Then
        vMode = enNotEditing
        cmdNewRows.Caption = "&Add"
        Me.lvw.Enabled = True
        lvw.Height = 4720
        fr1.ZOrder 1
        txtCurrencyRates.ZOrder 1
        SetEditFrameEnabled False, vMode
    ElseIf vMode = enAddingRow Then
        vMode = enNotEditing 'enEditingRow
        cmdNewRows.Caption = "&Add"
        lvw.Enabled = True
        lvw.Height = 4720
        fr1.ZOrder 1
        txtCurrencyRates.ZOrder 1
        SetEditFrameEnabled False, vMode
        oDel.GetStatus

    ElseIf vMode = enNotEditing Then
        vMode = enAddingRow
        Set oDELL = oDel.DeliveryLines.Add
        oDELL.SetQtyFirm 1
        oDELL.trid = oDel.trid
        cmdNewRows.Caption = "&Stop"
        lvw.Enabled = False
        lvw.Height = 2340
        SetEditFrameEnabled True, vMode
        mSetfocus txtCode
    End If
    ClearLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    lvw.ListItems.Clear
    For i = 1 To oDel.DeliveryLines.Count
        Set lstItem = lvw.ListItems.Add
        Set oDELL = oDel.DeliveryLines(i)
        LoadListViewLine i & "k", lstItem
    Next i
EXIT_HANDLER:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.LoadListView"
End Sub
Private Sub LoadListViewLine(i As String, lstItem As ListItem)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oDELL
        lstItem.Text = .CodeF
        If lstItem.Key = "" Then lstItem.Key = .Key
        lstItem.SubItems(1) = .Title
        lstItem.SubItems(2) = .QtyFirmF
        lstItem.SubItems(3) = .QtySSF
        lstItem.SubItems(4) = .PriceF(oDel.isFOreignCurrency)
        lstItem.SubItems(5) = .DiscountF
        lstItem.SubItems(6) = .Ref
        lstItem.SubItems(7) = .PLessDiscExtF(oDel.isFOreignCurrency)
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.LoadListViewLine(i,lstItem)", Array(i, lstItem)
End Sub
Private Sub Lvw_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    If lvw.ListItems.Count = 0 Then Exit Sub
    lngILEditingIdx = lvw.SelectedItem.Key
    Set oDELL = oDel.DeliveryLines.Item(lngILEditingIdx)
    oDELL.SetLineProduct oDELL.pID
    
    oDel.ReloadMatches oDELL.pID  'loads only POLSOS for this product into oDEL.POLsOSPersSUPP
    CheckForPreviousMatchesInInvoice oDELL.Key  'marks up the qty outstanding to inclide any qtys already captured against that POL
    LoadMatches
    If cboMatch.Items.ItemCount > 0 Then
        If oDELL.POLID > 0 Then
            cboMatch.Items.SelectItem(cboMatch.Items.FindItem(oDELL.POLID, 8)) = True
        End If
    End If
    
    lngSelectedRowIndex = lvw.SelectedItem.Key
    Me.txtCode = CStr(oDELL.CodeF)
    Me.txtTitle = oDELL.Title
    Me.txtQtySS = oDELL.QtySS
    Me.txtQtyFirm = oDELL.QtyFirm
    Me.txtNote = oDELL.Note
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oDELL.PriceF(oDel.isFOreignCurrency)
    Else
        txtPrice = oDELL.Price(oDel.isFOreignCurrency)
    End If
    oDELL.GetStatus
    Me.txtSP = oDELL.PriceSell

    Me.txtDiscount = CStr(oDELL.DiscountF)
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    mSetfocus txtPrice
    lvw.Height = 2340
    fr1.ZOrder 0
    txtCurrencyRates.ZOrder 0
    cmdNewRows.Caption = "&Stop edit"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Lvw_DblClick", , EA_NORERAISE
    HandleError
End Sub

''---------Companies code
'Private Sub LoadComps()
'Dim oComp As a_Company
'Dim oItem As ListItem
'Dim i As Integer
'    If oDEL.COMPID > 0 Then
'        cbComp.Caption = oPC.Configuration.Companies(CStr(oDEL.COMPID)).CompanyName
'    Else
'        cbComp.Caption = oPC.Configuration.DefaultCompany.CompanyName
'        oDEL.COMPID = oPC.Configuration.DefaultCOMPID
'    End If
'End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If oDel.Supplier Is Nothing Then
        MsgBox "Please enter a Supplier before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'-------End Compsny code
'Private Sub txtNote_Change()
'Dim intPos As Integer
'    If flgLoading Then Exit Sub
'    On Error Resume Next
'    oDELL.setnote (txtNote)
'    If Err Then
'      Beep
'      intPos = txtNote.SelStart
'      txtNote = oDELL.Note
'      txtNote.SelStart = intPos - 1
'    End If
'End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    Cancel = Not oDELL.setnote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oDELL.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuEditNote_Click()
'Dim ofrm As New frmNote
'    ofrm.Component oDel
'    ofrm.Show vbModal
'    Unload ofrm
'    Set ofrm = Nothing
'End Sub

'Private Sub mnuFileCancel_Click()
'    If oDel.IsDirty Then
'        oDel.CancelEdit
'    End If
'    Unload Me
'End Sub

'Private Sub mnuFileExit_Click()
'    oDel.CancelEdit
'    Unload Me
'End Sub

'Private Sub mnuFileOK_Click()
''    cmdOK_Click
'End Sub
'
'Private Sub mnuFilePrint_Click()
'    cmdIssue_Click
'End Sub
Private Sub mnuFile()
    On Error GoTo errHandler
    oDel.SetStatus stVOID
    oDel.ApplyEdit
    Unload Me
'    txtStatus = "Void"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.mnuFile"
End Sub
'Private Sub txtAccNum_Validate(Cancel As Boolean)
'Dim lngCustID As Long
'Dim bResult As Boolean
'    If Len(txtAccnum) > 0 Then
'        bResult = oDEL.SetSupplierFromAccNum(txtAccnum)
'        If bResult Then
'            With oDEL.Supplier
'                txtCustName = .Title & IIf(Len(.Title) > 0, " ", "") & .Initials & IIf(Len(.Initials) > 0, " ", "") & .Name
'                txtPhone = .Phone
'                lblAddBill.Caption = .BillToADdress.AddressShort
'                lblAddDel.Caption = .BillToADdress.AddressShort
'            End With
'            vCanAdd.RuleBroken "TP", False
'            Me.cmdNewRows.Enabled = True
'        Else
'            MsgBox "No such account number", , "Can't fetch Supplier"
'            txtAccnum = ""
'            Set oSupplier = Nothing
'            Cancel = True
'        End If
'    End If
'End Sub
'Private Sub txtAccNum_LostFocus()
'    txtAccnum = UCase(txtAccnum)
'End Sub


Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim pNumOfApproLines As Long
'Dim frmSection As frmSection
Dim bOK As Boolean
Dim frmPubRep As frmPublishersReport
Dim oPCode As New z_ProdCode
Dim tmp As String
    If flgLoading Then Exit Sub
'    If vMode = enAddingRow And InStr(1, txtCode, "-") > 0 Then
'        Set frmPubRep = New frmPublishersReport
'        frmPubRep.Show vbModal
'        txtCode = ""
'        Cancel = True
'        Exit Sub
'    End If
START:
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        GoTo EXIT_HANDLER
    End If

    bOK = oDELL.SetLineProduct("", txtCode)
    If bOK Then
            oDELL.Title = oDELL.Product.TitleAuthorPublisherL(35)
            oDel.ReloadMatches oDELL.pID
            CheckForPreviousMatchesInInvoice oDELL.DELLID
            LoadMatches
            If cboMatch.Items.ItemCount > 0 Then
                cboMatch.Items.SelectItem(cboMatch.Items(0)) = True
                tmp = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 3)
                oDELL.SetQtySS Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, "(") - 1)
                tmp = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2)
                oDELL.SetQtyFirm Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, "(") - 1)
                oDELL.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 10)
                oDELL.SetPrice cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9)
                oDELL.SetRef cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 6)
                oDELL.COLID = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 11)
             '   cboMatch.s
            Else
                oDELL.SetQtyFirm 1
                oDELL.SetQtySS 0
                oDELL.SetDiscount 0
              '  oDELL.setPrice 0
            End If
    Else   'Book nof found on database
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

    txtTitle = oDELL.Title
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oDELL.PriceF(oDel.isFOreignCurrency)
    Else
        txtPrice = oDELL.Price(oDel.isFOreignCurrency)
    End If
    txtQtyFirm = oDELL.QtyFirmF
    txtQtySS = oDELL.QtySSF
    txtDiscount = oDELL.DiscountF
    mSetfocus txtPrice
    oDELL.GetStatus
    
EXIT_HANDLER:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub CheckForPreviousMatchesInInvoice(pKey As String)
    On Error GoTo errHandler
Dim dPOLOS As d_POLine
Dim oDELL As a_DeliveryLine
Dim iQtyUnMatched As Integer
Dim iQtyUnMatchedSS As Integer
Dim iQtyUnMatchedFIRM As Integer

    For Each dPOLOS In oDel.POLsOSPersSUPP
        iQtyUnMatched = dPOLOS.qtyTotal - dPOLOS.ReceivedSoFar
        iQtyUnMatchedSS = dPOLOS.QtySSOS
        iQtyUnMatchedFIRM = dPOLOS.QtyFIRMOS
        For Each oDELL In oDel.DeliveryLines
            If pKey <> oDELL.Key Then
                If oDELL.IsDeleted = False And oDELL.POLID = dPOLOS.POLID Then
                    iQtyUnMatched = iQtyUnMatched - (oDELL.QtyFirm + oDELL.QtySS)
                    iQtyUnMatchedSS = iQtyUnMatchedSS - oDELL.QtySS
                    iQtyUnMatchedFIRM = iQtyUnMatchedFIRM - oDELL.QtyFirm
                End If
            End If
        Next
        dPOLOS.QtyUnMatchedTmp = iQtyUnMatched
        dPOLOS.QtySSUnMatchedTmp = iQtyUnMatchedSS
        dPOLOS.QtyFIRMUnMatchedTmp = iQtyUnMatchedFIRM
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.CheckForPreviousMatchesInInvoice(pKey)", pKey
End Sub
Private Sub LoadMatches()
    On Error GoTo errHandler
Dim oPOL As d_POLine
Dim i As Long
    If oDel.POLsOSPersSUPP.Count < 1 Then Exit Sub
    cboMatch.BeginUpdate
    i = 0
    For Each oPOL In oDel.POLsOSPersSUPP 'count the unfulfilled ones
        If oPOL.QtyUnMatchedTmp > 0 Then
            i = i + 1
        End If
    Next
    If i = 0 Then
        cboMatch.EndUpdate
        cboMatch.Items.RemoveAllItems
        Exit Sub
    End If
    ReDim ar(12, 0)
    cboMatch.Items.RemoveAllItems
    i = 0
    For Each oPOL In oDel.POLsOSPersSUPP
        If oPOL.QtyUnMatchedTmp > 0 Then
        
            ReDim Preserve ar(12, i)
            ar(0, i) = oPOL.DocDateF
            ar(1, i) = oPOL.DocCode
            ar(2, i) = oPOL.QtyFirm & "(" & oPOL.QtyFIRMUnMatchedTmp & ")"
            ar(3, i) = oPOL.QtySS & "(" & oPOL.QtySSUnMatchedTmp & ")"
            ar(5, i) = oPOL.ReceivedSoFar & "(" & oPOL.QtyUnMatchedTmp & ")"
            ar(6, i) = oPOL.Ref
            ar(7, i) = oPOL.DiscountF
            ar(8, i) = oPOL.POLID
            If Not oDel.isFOreignCurrency Then
                ar(4, i) = oPOL.POLPriceF
                ar(9, i) = oPOL.POLPrice
            Else
                ar(4, i) = oPOL.POLForeignPriceF
                ar(9, i) = oPOL.POLForeignPrice
            End If
            ar(10, i) = oPOL.Discount
            ar(11, i) = oPOL.COLID
            ar(12, i) = oPOL.QtyUnMatchedTmp
            i = i + 1
        End If
    Next
    cboMatch.PutItems ar
    cboMatch.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.LoadMatches"
End Sub
Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Or oDELL.Product Is Nothing Then Exit Sub
    If Not oDELL.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
    txtSP = oDELL.PriceSell
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub RemoveLine()
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
    iMax = lvw.ListItems.Count
    For i = iMax To 1 Step -1
        If lvw.ListItems(i).Selected Then
            If oDel.DeliveryLines.Item(lvw.ListItems(i).Key).POLID > 0 Then
                oDel.DeliveryLines.Item(lvw.ListItems(i).Key).POLID = 0  'In order to recycle the POLID for later selection in cbomatch ( i.e. it is no longer claimed)
            End If
            oDel.DeliveryLines.Remove lvw.ListItems(i).Key
            Exit For
        End If
    Next i
    If i = 0 Then
        MsgBox "Select an item prior to deleting.", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Exit Sub
    End If
    lvw.ListItems.Remove i
    lvw.Refresh
    oDel.CalculateTotals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.RemoveLine"
End Sub

Private Sub LoadSupplier()
    On Error GoTo errHandler
    With oDel
        SetIssueButtonCaption
        Me.txtSuppname = .Supplier.NameAndCode(20)
        If Not .Supplier.BillTOAddress Is Nothing Then
            txtPhone = .Supplier.BillTOAddress.Phone
            txtFax = .Supplier.BillTOAddress.Fax
        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.LoadSupplier"
End Sub


Private Sub SaveInvoice()
    On Error GoTo errHandler
    
    oDel.post
    
EXIT_HANDLER:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SaveInvoice"
End Sub

Public Sub PrintDelivery()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoDELLs As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oDel.Load oDel.trid
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoDELLs Then
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
    ErrorIn "frmdel.PrintDelivery"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoDELLs As Boolean
Dim iCurrency As Integer
Dim strResult As String
Dim frm As frmDELPreview
Dim cCOLALLOC As chex_COLAllocation
Dim frmAlloc As frmCOLAllocation_FromDel

    If oPC.Configuration.Signtransactions = True Then
        If SecurityControl(enSECURITY_GRN_SIGN, , "Sign this G.R.N.", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oDel.Status = stInProcess Then
            If MsgBox("Issue this G.R.N.? ", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    WaitMsg "Issuing delivery. . .", True, Me
    
    oDel.SetStatus stISSUED
    oDel.StaffID = gSTAFFID
    strResult = oDel.post
    If strResult = "" Then
        Set frm = New frmDELPreview
        frm.Component oDel.trid
        frm.Show
    End If
    If oPC.Configuration.COLAllocationStyle <> "S" Then  ' this is a retail environment and customer orders are held back at counter, not invoiced immediately
        Set cCOLALLOC = Nothing
        Set cCOLALLOC = New chex_COLAllocation
        cCOLALLOC.GenerateCOLALLOCationset oDel.trid
        cCOLALLOC.Load oDel.trid
        If cCOLALLOC.Count > 0 Then
            Set frmAlloc = New frmCOLAllocation_FromDel
            frmAlloc.Component cCOLALLOC, "DELIVERY", False
            frmAlloc.Show
        End If
        Set cCOLALLOC = Nothing
    End If
    WaitMsg "", False, Me
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub PrintProductsAwaitedBYCOs()
'Dim oC As c_COLSPERDEL
'Dim ar As New arCOLSFulfilled
'    Set oC = oDel.CustomerOrdersFulfilled
'    ar.Component oC
'    ar.Show
'End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
    oDel.SetStatus stInProcess
    SaveDEL
    oDel.BeginEdit
    Set oDELL = oDel.DeliveryLines.Add
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmDELPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel this G.R.N.?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    oDel.CancelEdit
    If cmdCancel.Caption = "&Close" Then
        Set frm = New frmDELPreview
        frm.ComponentObject oDel
        frm.Show
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.txtCode = ""
    Me.txtDiscount = ""
    Me.txtQtyFirm = ""
    Me.txtQtySS = ""
    Me.txtPrice = ""
    txtSP = ""
    Me.txtTitle = ""
    Me.txtTotal = ""
    Me.txtNote = ""
    Me.cboMatch.Items.RemoveAllItems
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.ClearLineControls"
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oDel.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oDel.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SetIssueButtonCaption"
End Sub


Private Sub Lvw_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvw.SortKey = ColumnHeader.Index - 1
   ' Set Sorted to True to sort the list.
    If lvw.SortOrder = lvwAscending Then
        lvw.SortOrder = lvwDescending
    Else
        lvw.SortOrder = lvwAscending
    End If
   lvw.Sorted = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.Lvw_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvw.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvw.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SetLvw"
End Sub

Private Sub vCanAdd_Status(errors As String)
    On Error GoTo errHandler
MsgBox errors & "CANAADD"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.vCanAdd_Status(errors)", errors, EA_NORERAISE
    HandleError
End Sub

Private Sub SaveDEL()
    On Error GoTo errHandler
    
    oDel.post
    
EXIT_HANDLER:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
  '  Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SaveDEL"
End Sub
Sub SetupcboMatch()
    On Error GoTo errHandler
    cboMatch.BeginUpdate
    cboMatch.WidthList = 540
    cboMatch.HeightList = 162
    cboMatch.AllowSizeGrip = True
    cboMatch.AutoDropDown = True
    
    cboMatch.Columns.Add "Date"
    cboMatch.Columns.Add "Code"
    cboMatch.Columns.Add "Firm"
    cboMatch.Columns.Add "SS"
    cboMatch.Columns.Add "Price"
    cboMatch.Columns.Add "RecSoFar"
    cboMatch.Columns.Add "Ref"
    cboMatch.Columns.Add "Discount"
    cboMatch.Columns.Add "lngPOLID"
    cboMatch.Columns.Add "lngPrice"
    cboMatch.Columns.Add "dblDiscount"
    cboMatch.Columns.Add "lngCOLID"
    cboMatch.Columns.Add "lngQTYnNAllocated"
    
    cboMatch.Columns(0).Width = 90
    cboMatch.Columns(1).Width = 90
    cboMatch.Columns(2).Width = 50
    cboMatch.Columns(3).Width = 50
    cboMatch.Columns(4).Width = 70
    cboMatch.Columns(5).Width = 50
    cboMatch.Columns(6).Width = 70
    cboMatch.Columns(7).Width = 70
    cboMatch.Columns(8).Width = 0
    cboMatch.Columns(9).Width = 0
    cboMatch.Columns(10).Width = 0
    cboMatch.Columns(11).Width = 0
    cboMatch.Columns(12).Width = 0
    cboMatch.BackColorLock = Me.BackColor
    cboMatch.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel.SetupcboMatch"
End Sub

