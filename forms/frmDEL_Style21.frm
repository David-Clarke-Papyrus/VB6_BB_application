VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmdel_Style2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Goods received note"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11595
   ControlBox      =   0   'False
   Icon            =   "frmDEL_Style21.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11595
   Begin VB.TextBox txtRunningqty 
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
      Left            =   3750
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5685
      Width           =   750
   End
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
      TabIndex        =   43
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
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5685
      Width           =   690
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2340
      Left            =   90
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   450
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4128
      SortKey         =   8
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   4410
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
         Text            =   "Short"
         Object.Width           =   1005
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Price"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Disc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Ref"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
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
      Height          =   615
      Left            =   8730
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmDEL_Style21.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   1000
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
      Height          =   435
      Left            =   975
      MultiLine       =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2565
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
      Height          =   495
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5535
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
      Height          =   615
      Left            =   7710
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmDEL_Style21.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5460
      Width           =   1000
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5685
      Width           =   1530
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   2355
      Left            =   75
      TabIndex        =   20
      Top             =   3060
      Width           =   10710
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   195
         Left            =   1605
         TabIndex        =   9
         Top             =   195
         Width           =   345
      End
      Begin VB.ComboBox cboMultibuys 
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
         Left            =   675
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   885
         Width           =   2565
      End
      Begin VB.TextBox txtSections 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   7785
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   900
         Width           =   2790
      End
      Begin VB.CommandButton cmdSection 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7365
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   885
         Width           =   390
      End
      Begin VB.TextBox txtMargin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4395
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1815
         Width           =   1260
      End
      Begin VB.TextBox txtQtyShort 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   1485
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1470
         Width           =   255
      End
      Begin VB.CommandButton cmdShort 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1785
         Picture         =   "frmDEL_Style21.frx":2EB6
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Click to mark quantity of damaged stock being returned"
         Top             =   1455
         Width           =   390
      End
      Begin VB.CommandButton cmdListSubstitutions 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&List subst. matches"
         Height          =   285
         Left            =   4800
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   150
         Width           =   2085
      End
      Begin VB.CommandButton cmdSub 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Create subst."
         Height          =   525
         Left            =   9810
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   300
         Width           =   765
      End
      Begin VB.TextBox txtSP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4395
         TabIndex        =   15
         Top             =   1470
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
         Left            =   9300
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   555
         Width           =   255
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   750
         Left            =   6960
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Width           =   2610
      End
      Begin VB.TextBox txtQtyFirm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1485
         Width           =   660
      End
      Begin VB.TextBox txtQtySS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   795
         TabIndex        =   8
         Top             =   1485
         Width           =   660
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3495
         TabIndex        =   14
         Top             =   1485
         Width           =   870
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
         Left            =   9615
         MaskColor       =   &H00C4BCA4&
         Picture         =   "frmDEL_Style21.frx":3240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1485
         Width           =   1000
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Left            =   5745
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1110
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2070
         Width           =   6495
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2205
         TabIndex        =   11
         Top             =   1485
         Width           =   1260
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   435
         Width           =   1725
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboMatch 
         Height          =   315
         Left            =   1905
         OleObjectBlob   =   "frmDEL_Style21.frx":35CA
         TabIndex        =   2
         Top             =   465
         Width           =   7335
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
         Height          =   315
         Left            =   4695
         OleObjectBlob   =   "frmDEL_Style21.frx":4974
         TabIndex        =   4
         Top             =   885
         Width           =   2235
      End
      Begin VB.Label lblMulti 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Multi."
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
         Height          =   210
         Left            =   135
         TabIndex        =   10
         Top             =   945
         Width           =   480
      End
      Begin VB.Label lblSections 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cat."
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
         Height          =   315
         Left            =   6960
         TabIndex        =   53
         Top             =   930
         Width           =   480
      End
      Begin VB.Label lblPT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Product type"
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
         Height          =   270
         Left            =   3480
         TabIndex        =   52
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label lblMargin 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Margin:"
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
         Left            =   3690
         TabIndex        =   51
         Top             =   1815
         Width           =   705
      End
      Begin VB.Label lblShort 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Short"
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
         Left            =   1575
         TabIndex        =   48
         Top             =   1245
         Width           =   570
      End
      Begin VB.Label lblSP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sell.Pr."
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
         Left            =   4815
         TabIndex        =   44
         Top             =   1245
         Width           =   705
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
         Height          =   210
         Left            =   6975
         TabIndex        =   39
         Top             =   1275
         Width           =   510
      End
      Begin VB.Label lblFirm 
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
         Left            =   210
         TabIndex        =   37
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label lblWants 
         BackColor       =   &H00D3D3CB&
         Caption         =   "fulfilment of . . ."
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
         Left            =   2160
         TabIndex        =   32
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label lblSupplDisc 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Suppl.disc."
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
         Left            =   3495
         TabIndex        =   26
         Top             =   1245
         Width           =   930
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
         Left            =   6045
         TabIndex        =   25
         Top             =   1245
         Width           =   660
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
         Left            =   105
         TabIndex        =   24
         Top             =   195
         Width           =   1410
      End
      Begin VB.Label lblSS 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SS"
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
         Left            =   930
         TabIndex        =   23
         Top             =   1245
         Width           =   360
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
         Left            =   2520
         TabIndex        =   22
         Top             =   1245
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
      Height          =   615
      Left            =   9750
      Picture         =   "frmDEL_Style21.frx":5D1E
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5460
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   345
      Left            =   90
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   30
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
      Caption         =   "Calculated totals"
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
      Left            =   3810
      TabIndex        =   41
      Top             =   5445
      Width           =   2205
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
      TabIndex        =   38
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
      TabIndex        =   36
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
      TabIndex        =   35
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
      TabIndex        =   34
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
      TabIndex        =   29
      Top             =   60
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4380
      Picture         =   "frmDEL_Style21.frx":60A8
      Stretch         =   -1  'True
      Top             =   60
      Width           =   360
   End
End
Attribute VB_Name = "frmdel_Style2"
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
Dim bDroppingDown As Boolean
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
Dim lngProductTypeID As Long
Dim blnReadOnly As Boolean
Dim flgLoading As Boolean
Dim WithEvents vCanAdd As z_BrokenRules
Attribute vCanAdd.VB_VarHelpID = -1
Dim WithEvents vCanIssue As z_BrokenRules
Attribute vCanIssue.VB_VarHelpID = -1
Dim strDELErrMsg As String
Dim strDELLErrMsg As String
Dim bSubstitute As Boolean
Dim dblCurrentMargin As Double
Dim dblCurrentPrice As Double
Dim flgLoadingPrice As Boolean
Dim flgLoadingMargin As Boolean

Public Sub component(pCancel As Boolean, Optional pTPID As Long, Optional pDel As a_Delivery)
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
        If oDel.Supplier.Deals.Count < 1 Then
            MsgBox "There are no deals for this supplier. You cannot continue"
            pCancel = True
            Exit Sub
        End If
        Set frm = New frmHeader_GRN
        frm.component oDel
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
        

        ChangeState enAddingRow
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
        oDel.GetStatus
        ChangeState enNotEditing
    End If
    oDel.GetStatus
    If oDel.ISForeignCurrency Then
        Me.txtRunningTotal = oDel.TotalLessDiscExtF(True)
        txtCurrencyRates = oDel.CurrencyConversionAsText & "     Value is : " & oDel.TotalLessDiscExtF(True)
        txtCurrencyRates.Visible = True
    Else
        Me.txtRunningTotal = oDel.TotalLessDiscExtF(False)
        txtCurrencyRates.Visible = False
    End If
    Me.txtRunningqty = oDel.TotalQuantityNetF
    SetMenu
    txtSP.Visible = oPC.SetPricesInGRN

    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.component(pCancel,pTPID,pDel)", Array(pCancel, pTPID, pDel)
End Sub


Private Sub cboProductType_SelectionChanged()
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oDELL.ProductTypeID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    If oPC.Configuration.ProductTypes.f3(oDELL.ProductTypeID) = "" Or oPC.Configuration.ProductTypes.f3(oDELL.ProductTypeID) = "False" Then
        oDELL.Product.Seesafe = 0
    Else
        oDELL.Product.Seesafe = 1
    End If
    lngProductTypeID = oDELL.ProductTypeID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cboProductType_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSection_Click()
    On Error GoTo errHandler
Dim frm As frmSection2
    If Not oDELL.Product.PID > "" Then Exit Sub
    Set frm = New frmSection2
    frm.component oDELL.Product
    frm.Show vbModal
    txtSections = oDELL.Product.ProductSections.SectionsAsList
    mSetfocus txtQtyFirm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdSection_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdShort_Click()
    On Error GoTo errHandler
    txtQtyShort.Enabled = True
Dim frm As New frmSupplierRetFromDelivery
    frm.SetParentCoords Me.TOP, Me.Left
    frm.Show vbModal
    If frm.QtyClaim > 0 Then
        oDELL.ReasonID = frm.Reasons
        oDELL.SetQtyShort frm.QtyClaim
        txtQtyShort = frm.QtyClaim
    End If
    Unload frm
    txtPrice.SetFocus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdShort_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSub_Click()
    On Error GoTo errHandler
Dim frm As New frmSubstitute
    frm.component Trim(txtCode)
    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdSub_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler
Dim frm As frmProductSingles
Dim oProd As a_Product

    Screen.MousePointer = vbHourglass
    Set oProd = Constructor.CreateProduct(True)
    Set frm = New frmProductSingles
    frm.component oProd
    frm.Show 'vbModal
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oDel.StatusF = "IN PROCESS" And oDel.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
   ' Forms(0).mnuDelact.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SetMenu"
End Sub



Private Sub cboMatch_SelectionChanged()
    On Error GoTo errHandler
Dim H As HITEM
Dim tmp As String

    If cboMatch.Items.SelectCount > 0 Then
        If (cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 13) <> oDel.Supplier.ID) And bDroppingDown = False Then
            MsgBox "This order line is on a supplier other than the supplier of the present goods. ", vbOKOnly, "Warning"
        End If
        tmp = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 3)
        If vMode <> eneditingrow Then
            oDELL.SetQtySS Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, "(") - 1)
        End If
        tmp = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 2)
        If vMode <> eneditingrow Then
            oDELL.SetQtyFirm Mid(tmp, InStr(1, tmp, "(") + 1, InStr(1, tmp, "(") - 1)
            'New 20/3/2009 ---
            oDELL.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 10)
            oDELL.SetPrice cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9)
            '---
        End If
        
'        oDELL.SetDiscount cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 10)
'        oDELL.SetPrice cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 9)
        
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
        txtPrice = oDELL.PriceF(oDel.ISForeignCurrency)
        txtSP = oDELL.PriceSell
    Else
        txtPrice = oDELL.Price(oDel.ISForeignCurrency)
        txtSP = oDELL.PriceSell
    End If
    txtQtyFirm = oDELL.QtyFirmF
    txtQtySS = oDELL.QtySSF
    txtDiscount = oDELL.DiscountF
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.ISForeignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cboMatch_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdBatch_Click()
    On Error GoTo errHandler
Dim frm As New frmHeader_GRN
    frm.component oDel
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdBatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancelMatch_Click()
    On Error GoTo errHandler
Dim i As Integer

    If cboMatch.Items.ItemCount = 0 Then Exit Sub
    oDELL.POLID = 0
    For i = 0 To cboMatch.Items.ItemCount - 1
        cboMatch.Items.SelectItem(cboMatch.Items(i)) = False
    Next
    mSetfocus txtQtyFirm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdCancelMatch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As New frmSupplierPreview
    
    If oDel.Supplier.ID > 0 Then
        frm.component oDel.Supplier
        frm.Show
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbSupp_Click()
    On Error GoTo errHandler
Dim frm As frmSupplierPreview
    If oDel.Supplier.Name = "" Then Exit Sub
    Set frm = New frmSupplierPreview
    frm.component oDel.Supplier
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cbSupp_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdFulfilments_Click()
    On Error GoTo errHandler
   ' ReconcileWithCOs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdFulfilments_Click", , EA_NORERAISE
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
    ErrorIn "frmdel_Style2.LoadNewSupplier(plngTPID)", plngTPID
End Sub

Private Sub cmdNote_Click()
    On Error GoTo errHandler
Dim frm As New frmILNote
    frm.component oDELL
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdNote_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set vCanAdd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.mnuDelLine"
End Sub

Private Sub lvw_Click()
    On Error GoTo errHandler
    
    If lvw Is Nothing Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    
    If Me.lvw.SelectedItem.Index > 0 Then
    
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText Left(lvw.SelectedItem.SubItems(10), ISBNLENGTH)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.lvw_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oDEL_ValidToSave(pOK As Boolean)
    On Error GoTo errHandler
    cmdSave.Enabled = (pOK And oDel.DeliveryLines.Count > 0 And vMode = enNotEditing And oDel.IsDirty)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.oDEL_ValidToSave(pOK)", pOK, EA_NORERAISE
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
    ErrorIn "frmdel_Style2.oDEL_Valid(pMsg)", pMsg, EA_NORERAISE
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
    ErrorIn "frmdel_Style2.oDELL_ExtensionChange(lngExtension,strExtension)", Array(lngExtension, _
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
    ErrorIn "frmdel_Style2.oDELL_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oDEL_TotalChange(strtotal As String, strTotalForeign As String, strQtyTotal As String)
    On Error GoTo errHandler
    flgLoading = True
    If oDel.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
        Me.txtRunningTotal = strtotal
    Else
        Me.txtRunningTotal = strTotalForeign
        txtCurrencyRates = oDel.CurrencyConversionAsText & "     Value is : " & strtotal
    End If
    Me.txtRunningqty = strQtyTotal
    
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.oDEL_TotalChange(strtotal,strTotalForeign,strQtyTotal)", Array(strtotal, _
         strTotalForeign, strQtyTotal), EA_NORERAISE
    HandleError
End Sub

Private Sub oDEL_Reloadlist()
    On Error GoTo errHandler
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.oDEL_Reloadlist", , EA_NORERAISE
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
    ErrorIn "frmdel_Style2.oDEL_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub
Private Sub oDEL_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    MsgBox "CurrentRow Status = " & pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.oDEL_CurrRowStatus(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub



Private Sub txtCode_LostFocus()
    On Error GoTo errHandler
    If txtCode > "" Then
        bDroppingDown = True
        Sendkeys "+({F4})", True
        bDroppingDown = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtCode_LostFocus", , EA_NORERAISE
    HandleError
End Sub






Private Sub txtNote_Change()
    On Error GoTo errHandler
    txtNote = HandleTextWithBites(txtNote)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtyShort_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtyShort
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtQtyShort_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtyShort_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetQtyShort(txtQtyShort) Then
        Cancel = True
    End If
   ' oDel.CalculateTotals
   ' txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtQtyShort_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtQtyShort_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    If Not oDELL.SetQtyShort(txtQtyShort) Then
'        Cancel = True
'    End If
'    oDel.CalculateTotals
'    txtTotal = oDELL.PLessDiscExtF(oDel.isFOreignCurrency)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmdel.txtQtyShort_Validate(Cancel)", Cancel, EA_NORERAISE, , "Rowcount,ODELL=NOTHING,oDEL=Nothing", Array(oDel.DeliveryLines.Count, oDel Is Nothing, oDELL Is Nothing)
'    HandleError
'End Sub



Private Sub txtQtyFirm_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtyFirm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtQtyFirm_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySS_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtySS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtQtySS_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtyFirm_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetQtyFirm(txtQtyFirm) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.ISForeignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtQtyFirm_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySs_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.ISForeignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSP_Change()
    On Error GoTo errHandler
    If flgLoadingPrice Then Exit Sub
    If IsNumeric(txtSP) Then
        dblCurrentPrice = CDbl(txtSP)
        dblCurrentMargin = CalculateMargin()
        flgLoadingMargin = True
        txtMargin = PBKSPercentF(dblCurrentMargin * 100)
        flgLoadingMargin = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtSP_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtMargin_Change()
    On Error GoTo errHandler
Dim lngTmp As Long

    If flgLoadingMargin Then Exit Sub
    If IsNumeric(txtMargin) Then
        If CDbl(txtMargin) > 99 Then Exit Sub
        dblCurrentMargin = CDbl(txtMargin)
        dblCurrentPrice = CalculatePrice()
        flgLoadingPrice = True
        txtSP = CStr(CLng(dblCurrentPrice))
        If ConvertToLng(Trim(txtSP), lngTmp) Then
            oDELL.SetPriceSell Trim(txtSP)
        End If
        
        flgLoadingPrice = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtMargin_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtMargin_GotFocus()
    On Error GoTo errHandler
    dblCurrentMargin = CalculateMargin()
    txtMargin = Format(dblCurrentMargin * 100, "###,##0.00")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtMargin_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Function CalculatePrice() As Long
    On Error GoTo errHandler
   CalculatePrice = (oDELL.PLessDisc(oDel.ISForeignCurrency)) / ((100 - dblCurrentMargin) / 100)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.CalculatePrice"
End Function
Private Function CalculateMargin() As Double
    On Error GoTo errHandler
    If dblCurrentPrice > 0 Then
        CalculateMargin = (dblCurrentPrice - oDELL.PLessDisc(oDel.ISForeignCurrency)) / dblCurrentPrice
    Else
        CalculateMargin = 0
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.CalculateMargin"
End Function

Private Sub txtSP_GotFocus()
    On Error GoTo errHandler
    If IsNumeric(txtSP) Then
        dblCurrentPrice = CDbl(txtSP)
        dblCurrentMargin = CalculateMargin()
        txtMargin = PBKSPercentF(dblCurrentMargin * 100)
    End If
    AutoSelect txtSP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtSP_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngTmp As Long

    If ConvertToLng(Trim(txtSP), lngTmp) Then
        oDELL.SetPriceSell Trim(txtSP)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtSP_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmdel_Style2.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
Dim curTotalDeposit As Currency
Dim strAddress As String
Dim ar() As String
    flgLoading = True
    SetupcboMatch
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11100
        Height = 6700
    End If
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add "Seesafe"
    cboProductType.Columns(0).Width = 190
    cboProductType.Columns(1).Width = 0
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate
    
'    flgLoading = True
'    flgLoading = False
    oDel.GetStatus
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate
    LoadCombo Me.cboMultibuys, oPC.Configuration.Multibuys
    
  '  SetLvw
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oDel.Memo
    ofrm.Show vbModal
    oDel.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.mnuMemo"
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
   ' SetFormLayout
    Set vCanAdd = New z_BrokenRules
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetFormLayout()
    On Error GoTo errHandler

    If oPC.GetProperty("DeliveryStyle") = "BB" Then
        fr1.Height = 3075
        cboMultibuys.Visible = True
        cboMultibuys.TOP = 885
        lblMulti.Visible = True
        lblMulti.TOP = 945
        
        cboProductType.Visible = True
        cboProductType.TOP = 885
        lblPT.Visible = True
        lblPT.TOP = 945
        
        txtSections.Visible = True
        txtSections.TOP = 885
        lblSections.Visible = True
        lblSections.TOP = 945
        cmdSection.Visible = True
        cmdSection.TOP = 885
        
        lblFirm.TOP = 1275
        lblSS.TOP = 1275
        lblShort.TOP = 1275
        lblPrice.TOP = 1275
        lblSupplDisc.TOP = 1275
        lblSP.TOP = 1275
        lblTotal.TOP = 1275
        lblNote.TOP = 1275
        txtQtyFirm.TOP = 1485
        txtQtySS.TOP = 1485
        txtQtyShort.TOP = 1485
        cmdShort.TOP = 1455
        txtPrice.TOP = 1485
        txtDiscount.TOP = 1485
        txtSP.TOP = 1485
        txtNote.TOP = 1485
        txtTotal.TOP = 1485

        txtTitle.TOP = 2070
        lblMargin.TOP = 1815
        txtMargin.TOP = 1815
        
        cmdEnter.TOP = 1470
    Else
        fr1.Height = 2300
        cboMultibuys.Visible = False
        cboMultibuys.TOP = fr1.TOP + 885
        lblMulti.Visible = False
        lblMulti.TOP = 945
        
        cboProductType.Visible = False
        cboProductType.TOP = fr1.TOP + 885
        lblPT.Visible = False
        lblPT.TOP = 945
        
        txtSections.Visible = False
        txtSections.TOP = fr1.TOP + 885
        lblSections.Visible = False
        lblSections.TOP = 945
        cmdSection.Visible = False
        cmdSection.TOP = 885
        
        lblFirm.TOP = 945
        lblSS.TOP = 945
        lblShort.TOP = 945
        lblPrice.TOP = 945
        lblSupplDisc.TOP = 945
        lblSP.TOP = 945
        lblTotal.TOP = 945
        lblNote.TOP = 945
        txtQtyFirm.TOP = 1185
        txtQtySS.TOP = 1185
        txtQtyShort.TOP = 1185
        cmdShort.TOP = 1185
        txtPrice.TOP = 1185
        txtDiscount.TOP = 1185
        txtSP.TOP = 1185
        txtNote.TOP = 1185
        txtTotal.TOP = 1185

        txtTitle.TOP = 1785
        lblMargin.TOP = 1530
        txtMargin.TOP = 1515
        
        cmdEnter.TOP = 1320
    
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SetFormLayout"
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
    ErrorIn "frmdel_Style2.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim i As Integer
Dim iDiff As Integer
Dim dblMU As Double


    If oDELL Is Nothing Then Exit Sub
    If oDel Is Nothing Then Exit Sub
    If Not oDel.ISForeignCurrency Then
    dblMU = Markup(oDELL.PriceSell, (oDELL.Price(False) * (100 - oDELL.Discount)) / 100)
       If dblMU < oPC.Configuration.MinMU Then
           MsgBox "The markup indicated is only " & PBKSPercentF(dblMU) & " percent and less than the minimum markup allowed. (" & oPC.Configuration.MinMU & "%)" & vbCrLf & "Please change the selling price to at least " & Format(MinimumSP((oDELL.Price(False) * (100 - oDELL.Discount)) / 100) / oPC.Configuration.DefaultCurrency.Divisor, "R#,##0.00"), vbInformation, "Warning"
       End If
    End If
    If cboMatch.Items.SelectCount > 0 Then
        i = cboMatch.Items.CellCaption(cboMatch.Items.SelectedItem(0), 12)
        If (oDELL.QtyFirm + oDELL.QtySS + oDELL.QtyShort) > i Then
            If i > 1 Then
                MsgBox "There are only " & i & " items outstanding on this purchase order line.", vbInformation, "Warning"
            Else
                MsgBox "There is only one item outstanding on this purchase order line.", vbInformation, "Warning"
            End If
            Exit Sub
        End If
    End If
    txtQtyShort.Enabled = False
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        If txtCode.Enabled Then mSetfocus txtCode
        Exit Sub
    End If
    oDELL.ApplyEdit
    oDELL.BeginEdit

    If vMode = enAddingRow Then
        If lvw.ListItems.Count < val(oDELL.Key) Then
            lvw.ListItems.Add Key:=oDELL.Key
            LoadListViewLine lvw.ListItems(lvw.ListItems.Count), oDELL
        End If
        lvw.Refresh
        ChangeState enAddingRow
        mSetfocus txtCode
    ElseIf vMode = eneditingrow Then
        LoadListViewLine lvw.ListItems(lngSelectedRowIndex), oDELL
        ChangeState enNotEditing
        txtCurrencyRates.ZOrder 1
    End If
    oDel.CalculateTotals
    oDel.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdEnter_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
    
    If vMode = eneditingrow Then
       ' LogSaveToFile "GRN New row button:enEditingRow"
        If oDELL.IsEditing Then
            oDELL.CancelEdit
            oDELL.BeginEdit
        End If
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
      '  LogSaveToFile "GRN New row button:enAddingRow"
        If txtCode > "" Then  'THis is not after a post but is an aborted  add row action
            oDel.DeliveryLines.DecrementMaxKeyUsed
        End If
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
       ' LogSaveToFile "GRN New row button:enNotEditing"
        ChangeState enAddingRow
    End If


    ClearLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To lvw.ColumnHeaders.Count
        lvw.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvw.ColumnHeaders(i).Width)
    Next
    lvw.ListItems.Clear
    For i = 1 To oDel.DeliveryLines.Count
        Set lstItem = lvw.ListItems.Add
        LoadListViewLine lstItem, oDel.DeliveryLines(i)
    Next i
EXIT_Handler:
    Set lstItem = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.LoadListView"
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvw, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.mnuSaveLayout"
End Sub

Private Sub LoadListViewLine(lstItem As ListItem, oDELL As a_DeliveryLine)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oDELL
        lstItem.text = .CodeF
        lstItem.Key = .Key
        lstItem.SubItems(1) = .Title
        lstItem.SubItems(2) = .QtyFirmF
        lstItem.SubItems(3) = .QtySSF
        lstItem.SubItems(4) = .QtyShortF
        lstItem.SubItems(5) = .PriceF(oDel.ISForeignCurrency)
        lstItem.SubItems(6) = .DiscountF
        lstItem.SubItems(7) = .Ref
        lstItem.SubItems(8) = .PLessDiscExtF(oDel.ISForeignCurrency)
        lstItem.SubItems(9) = Format(.Key, "@@@@@@@@@@")
        lstItem.SubItems(10) = .EAN
        
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.LoadListViewLine(lstItem,oDELL)", Array(lstItem, oDELL)
End Sub
Private Sub Lvw_DblClick()
    On Error GoTo errHandler
'This must load the editing line with the current line's data
    flgLoading = True
    If lvw.ListItems.Count = 0 Then Exit Sub
    If lvw.SelectedItem.Index < 1 Then Exit Sub
    
    lngILEditingIdx = lvw.SelectedItem.Key
    Set oDELL = oDel.DeliveryLines(lvw.SelectedItem.Key)    'oDEL.DeliveryLines(lngILEditingIdx)
    lngSelectedRowIndex = lvw.SelectedItem.Key
    
    ChangeState eneditingrow
    
    oDELL.SetLineProduct oDELL.PID, , True
    
    oDel.ReloadMatches oDELL.PID  'loads only POLSOS for this product into oDEL.POLsOSPersSUPP
    CheckForPreviousMatchesInInvoice oDELL.Key  'marks up the qty outstanding to inclide any qtys already captured against that POL
    LoadMatches
    SetMultibuyCode oDELL.MBCode
    If cboMatch.Items.ItemCount > 0 Then
        If oDELL.POLID > 0 Then
            cboMatch.Items.SelectItem(cboMatch.Items.FindItem(oDELL.POLID, 8)) = True
        End If
    End If
    Me.txtCode = IIf(Len(CStr(oDELL.EAN)) > 0, CStr(oDELL.EAN), CStr(oDELL.code))
    Me.txtTitle = oDELL.Title
    Me.txtQtySS = oDELL.QtySS
    Me.txtQtyFirm = oDELL.QtyFirm
    Me.txtQtyShort = oDELL.QtyShortF
    Me.txtNote = oDELL.Note
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oDELL.PriceF(oDel.ISForeignCurrency)
    Else
        txtPrice = oDELL.Price(oDel.ISForeignCurrency)
    End If
    txtSP = oDELL.PriceSell
    oDELL.GetStatus
    cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(oDELL.ProductTypeID), 0)) = True

    Me.txtDiscount = CStr(oDELL.DiscountF)
 '   SetEditFrameEnabled True, enEditingRow
 '   vMode = enEditingRow
    If oDELL.QtyFirm > 1 Then
        mSetfocus Me.txtQtyFirm
    Else
        mSetfocus txtPrice
    End If
    flgLoading = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.Lvw_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub ChangeState(pToMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
    SetFormLayout
    vMode = pToMode
    lngProductTypeID = 0

    Select Case pToMode
    Case eneditingrow
        fr1.Visible = True
        Me.txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        txtCurrencyRates.Visible = False
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQtyFirm.Enabled = True
        Me.txtCurrencyRates.Visible = True
        cboProductType.Enabled = True
        txtQtySS = True
        UnsetMenu
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oDel.DeliveryLines.Count > 0)
        cmdCancel.Caption = "&Close"
        Me.lvw.Enabled = False
        lvw.Height = 2200
        fr1.ZOrder 1
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtDiscount.Enabled = True
        txtPrice.Enabled = True
        'txtRef.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQtyFirm.Enabled = True
        txtQtySS = True
        Me.txtCurrencyRates.Visible = True
        txtError = ""
        flgLoading = True
        UnsetMenu
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oDel.DeliveryLines.Count > 0)
        cmdNewRows.Caption = "&Stop"
        
        lvw.Enabled = False
        lvw.Height = 2200
        ClearLineControls
        fr1.ZOrder 1
        mSetfocus txtCode
'        If Not oDELL Is Nothing Then
'            If oDELL.IsEditing Then oDELL.CancelEdit
'            Set oDELL = Nothing
'        End If
        Set oDELL = oDel.DeliveryLines.Add
        oDELL.TRID = oDel.TRID
        oDELL.ProductTypeID = lngProductTypeID
        oDELL.SetQtyFirm 1
        
    Case enNotEditing
        flgLoading = True
        fr1.Visible = False
        txtError = ""
        SetMenu
        flgLoading = False
        cmdEnter.Enabled = False
        cmdCancel.Enabled = True
        cmdIssue.Enabled = True
        cmdSave.Enabled = True
        cmdNewRows.Enabled = True ' (oDel.DeliveryLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        Me.txtCurrencyRates.Visible = False
        lvw.Enabled = True
        lvw.Height = 4600
'        If oDel.DeliveryLines.IsEditing Then
'            oDel.DeliveryLines.CancelEdit
'            oDel.DeliveryLines.BeginEdit
'        End If
'        If Not oDELL Is Nothing Then
'            If oDELL.IsEditing Then   '
'                oDELL.CancelEdit
'                Set oDELL = Nothing
'            End If
'        End If
        fr1.ZOrder 1
    End Select
    oDel.GetStatus
        If Not oDel.IsDirty Then
            cmdCancel.Caption = "&Close"
        Else
            cmdCancel.Caption = "&Cancel"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.ChangeState(pToMode)", pToMode
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
    ErrorIn "frmdel_Style2.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
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
    Cancel = Not oDELL.SetNote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oDELL.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtNote_LostFocus", , EA_NORERAISE
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
    ErrorIn "frmdel_Style2.mnuFile"
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
Private Sub cmdListSubstitutions_Click()
    On Error GoTo errHandler
    bSubstitute = True
    AcceptCode
    bSubstitute = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdListSubstitutions_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function AcceptCode() As Boolean
    On Error GoTo errHandler
Dim pQty As Integer
Dim pApproID As Long
Dim pNumOfApproLines As Long
'Dim frmSection As frmSection
Dim bOK As Boolean
Dim frmPubRep As frmPublishersReport
Dim oPCode As New z_ProdCode
Dim tmp As String
Dim j As Integer

    If flgLoading Then Exit Function
START:
    AcceptCode = True
    If txtCode = "" Or vMode = eneditingrow Then Exit Function
    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        AcceptCode = False
        GoTo EXIT_Handler
    End If

    bOK = oDELL.SetLineProduct("", txtCode, True)
        If bOK Then
            txtSections = oDELL.Product.ProductSections.SectionsAsList
            If Me.cboProductType.Items.ItemCount > 0 Then
                If lngProductTypeID > 0 Then
                    cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(lngProductTypeID), 0)) = True
                ElseIf oDELL.ProductTypeID > 0 Then
                    cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(oDELL.ProductTypeID), 0)) = True
                ElseIf oDELL.Product.ProductTypeID > 0 Then
                    cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(oDELL.Product.ProductTypeID), 0)) = True
                Else
                    cboProductType.Items.SelectItem(cboProductType.Items(0)) = True
                    lngProductTypeID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
                End If
            End If
            oDELL.Title = oDELL.Product.TitleAuthorPublisherL(35)
            If bSubstitute = False Then
                oDel.ReloadMatches oDELL.PID
            Else
                oDel.ReloadMatches_forSubstitutions oDELL.PID
            End If
            CheckForPreviousMatchesInInvoice oDELL.DELLID
            LoadMatches
            SetMultibuyCode oDELL.Product.MultibuyCode
            If cboMatch.Items.ItemCount > 0 Then
                For j = 0 To cboMatch.Items.ItemCount - 1
                    If cboMatch.Items.CellCaption(cboMatch.Items(j), 13) = oDel.Supplier.ID Then
                        cboMatch.Items.SelectItem(cboMatch.Items(j)) = True
                    End If
                Next j
                If cboMatch.Items.SelectCount = 0 Then cboMatch.Items.SelectItem(cboMatch.Items(0)) = True
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
                If oDELL.QtyFirm = 0 And oDELL.QtySS = 0 Then
                    oDELL.SetQtyFirm 1
                    oDELL.SetQtySS 0
                End If
                   ' oDELL.SetDiscount 0
                
              '  oDELL.setPrice 0
            End If
    Else   'Book nof found on database
        If CheckThisPoint(M_NEWPRODUCTINADHOCFORM) Then
            If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then
                AcceptCode = False
                GoTo START
            End If
        End If
    
        Dim frmAdHoc As frmAdHocProduct
        Set frmAdHoc = New frmAdHocProduct
        frmAdHoc.component txtCode
        frmAdHoc.Show vbModal
        txtCode = frmAdHoc.code
        Unload frmAdHoc
        Set frmAdHoc = Nothing
        AcceptCode = False
        GoTo START
    End If

    txtTitle = oDELL.Title
    If oPC.Configuration.CaptureDecimal Then
        txtPrice = oDELL.PriceF(oDel.ISForeignCurrency)
    Else
        txtPrice = oDELL.Price(oDel.ISForeignCurrency)
    End If
    txtQtyFirm = oDELL.QtyFirmF
    txtQtySS = oDELL.QtySSF
    txtDiscount = oDELL.DiscountF
    mSetfocus txtPrice
    oDELL.GetStatus
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.AcceptCode"
End Function
Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not AcceptCode
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
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
        iQtyUnMatched = dPOLOS.QtyTotal - dPOLOS.ReceivedSoFar
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
    ErrorIn "frmdel_Style2.CheckForPreviousMatchesInInvoice(pKey)", pKey
End Sub
Private Sub LoadMatches()
    On Error GoTo errHandler
Dim oPOL As d_POLine
Dim i As Long
    If oDel.POLsOSPersSUPP.Count < 1 Then
        cboMatch.Items.RemoveAllItems
        Exit Sub
    End If
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
    ReDim ar(13, 0)
    cboMatch.Items.RemoveAllItems
    i = 0
    For Each oPOL In oDel.POLsOSPersSUPP
        If oPOL.QtyUnMatchedTmp > 0 Then
        
            ReDim Preserve ar(13, i)
            ar(0, i) = oPOL.DocDateF
            ar(1, i) = oPOL.DOCCode
            ar(2, i) = oPOL.QtyFirm & "(" & oPOL.QtyFIRMUnMatchedTmp & ")"
            ar(3, i) = oPOL.QtySS & "(" & oPOL.QtySSUnMatchedTmp & ")"
            ar(5, i) = oPOL.ReceivedSoFar & "(" & oPOL.QtyUnMatchedTmp & ")"
            ar(6, i) = oPOL.Ref
            ar(7, i) = oPOL.DiscountF
            ar(8, i) = oPOL.POLID
            If Not oDel.ISForeignCurrency Then
                ar(4, i) = oPOL.POLPriceF
                ar(9, i) = oPOL.POLPrice
            Else
                ar(4, i) = oPOL.POLForeignPriceF
                ar(9, i) = oPOL.POLForeignPrice
            End If
            ar(10, i) = oPOL.Discount
            ar(11, i) = oPOL.COLID
            ar(12, i) = oPOL.QtyUnMatchedTmp
            ar(13, i) = oPOL.SupplierID
            i = i + 1
        End If
    Next
    cboMatch.PutItems ar
    cboMatch.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.LoadMatches"
End Sub
Private Sub txtDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oDELL.SetDiscount(txtDiscount) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.ISForeignCurrency)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDiscount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Or oDELL.Product Is Nothing Then Exit Sub
    If Not oDELL.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oDel.CalculateTotals
    txtTotal = oDELL.PLessDiscExtF(oDel.ISForeignCurrency)
    txtSP = oDELL.PriceSell
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.txtPrice_GotFocus", , EA_NORERAISE
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
    ErrorIn "frmdel_Style2.RemoveLine"
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
    ErrorIn "frmdel_Style2.LoadSupplier"
End Sub


Private Sub SaveInvoice()
    On Error GoTo errHandler
    
    oDel.Post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SaveInvoice"
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
    oDel.Load oDel.TRID
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoDELLs Then
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
    ErrorIn "frmdel_Style2.PrintDelivery"
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

    If oPC.Configuration.SignTransactions = True Then
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
    strResult = oDel.Post
    If strResult = "" Then
        Set frm = New frmDELPreview
        frm.component oDel.TRID
        frm.Show
    End If
    If Not oPC.IncludeSupplierFeatures Then  ' this is a retail environment and customer orders are held back at counter, not invoiced immediately
        Set cCOLALLOC = Nothing
        Set cCOLALLOC = New chex_COLAllocation
        cCOLALLOC.GenerateCOLAllocationset oDel.TRID
        cCOLALLOC.Load oDel.TRID
        If cCOLALLOC.Count > 0 Then
            Set frmAlloc = New frmCOLAllocation_FromDel
            frmAlloc.component cCOLALLOC, "DELIVERY", False
            frmAlloc.Show
        End If
        Set cCOLALLOC = Nothing
    End If
    WaitMsg "", False, Me
'    Dim rs As ADODB.Recordset
'    Set rs = oDel.GetRepricingRequirement()
'    If rs.RecordCount > 0 Then
'        PrintRepriceSheet rs
'    End If

    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub PrintRepriceSheet(rs As ADODB.Recordset)
Dim ar As New arRepriceList
    ar.Show vbModal
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
    LoadListView
    oDel.BeginEdit
    Set oDELL = oDel.DeliveryLines.Add
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmDELPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel your changes?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
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
    ErrorIn "frmdel_Style2.cmdCancel_Click", , EA_NORERAISE
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
    ErrorIn "frmdel_Style2.ClearLineControls"
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        If oDel.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oDel.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SetIssueButtonCaption"
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
    ErrorIn "frmdel_Style2.Lvw_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub
Private Sub SetLvw()
    On Error GoTo errHandler
Dim Style As Long
Dim hHeader As Long
   
'  'get the handle to the listview header
'   hHeader = SendMessage(lvw.hwnd, LVM_GETHEADER, 0, ByVal 0&)
'
'  'get the current style attributes for the header
'   style = GetWindowLong(hHeader, GWL_STYLE)
'
'  'modify the style by toggling the HDS_BUTTONS style
'   style = style Xor HDS_BUTTONS
'
'  'set the new style and redraw the listview
'   If style Then
'      Call SetWindowLong(hHeader, GWL_STYLE, style)
'      Call SetWindowPos(lvw.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SetLvw"
End Sub

Private Sub vCanAdd_Status(errors As String)
    On Error GoTo errHandler
MsgBox errors & "CANAADD"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.vCanAdd_Status(errors)", errors, EA_NORERAISE
    HandleError
End Sub

Private Sub SaveDEL()
    On Error GoTo errHandler
    
    oDel.Post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
  '  Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SaveDEL"
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
    cboMatch.Columns.Add "Supplier"
    
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
    cboMatch.Columns(13).Width = 0
    cboMatch.BackColorLock = Me.BackColor
    cboMatch.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SetupcboMatch"
End Sub

Private Sub cboMultibuys_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oDELL.MBCode = oPC.Configuration.Multibuys.f4(oPC.Configuration.Multibuys.Key(cboMultibuys))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.cboMultibuys_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMultibuyCode(MBCode As String)
    On Error GoTo errHandler
    If MBCode > "" Then
        Me.cboMultibuys = oPC.Configuration.Multibuys.ItemByF4(MBCode)
    Else
        cboMultibuys = "<N/A>"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmdel_Style2.SetMultibuyCode(MBCode)", MBCode
End Sub

