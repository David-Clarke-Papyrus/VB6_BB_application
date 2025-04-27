VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmPO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "frmPO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11400
   Begin VB.CommandButton cmdShowBudgets 
      Height          =   330
      Left            =   10365
      Picture         =   "frmPO.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6180
      Width           =   390
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
      Left            =   9660
      Picture         =   "frmPO.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1110
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
      Left            =   7410
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmPO.frx":0C9E
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1110
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
      Left            =   8535
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmPO.frx":1028
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5400
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2280
      Left            =   90
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   810
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4022
      SortKey         =   8
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   5468
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Firm"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "SS"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ref"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Discount"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Key"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   0
      EndProperty
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
      Height          =   1065
      Left            =   2115
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5325
      Width           =   3540
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
      Height          =   615
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1000
   End
   Begin VB.TextBox txtRunningTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   250
      Left            =   9435
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3075
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00D3D3CB&
      Height          =   2040
      Left            =   105
      TabIndex        =   16
      Top             =   3255
      Width           =   10710
      Begin VB.CommandButton cmdFind 
         Height          =   345
         Left            =   105
         Picture         =   "frmPO.frx":13B2
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   375
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
         Picture         =   "frmPO.frx":173C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   870
         Width           =   1000
      End
      Begin VB.TextBox txtOHOO 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   9375
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1575
         Width           =   1260
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
         Left            =   5955
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   405
         Width           =   390
      End
      Begin VB.TextBox txtSections 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   405
         Width           =   1770
      End
      Begin VB.TextBox txtETA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3735
         TabIndex        =   10
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   1095
         Width           =   1410
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1560
      End
      Begin VB.TextBox txtRef 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   2205
         TabIndex        =   2
         Top             =   405
         Width           =   1470
      End
      Begin VB.TextBox txtQtySS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   8835
         TabIndex        =   7
         Top             =   405
         Width           =   660
      End
      Begin VB.TextBox txtQtyFirm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   8145
         TabIndex        =   6
         Top             =   405
         Width           =   660
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   570
         Left            =   5955
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   900
         Width           =   3405
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1575
         Width           =   10455
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   9525
         TabIndex        =   8
         Top             =   405
         Width           =   1095
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   570
         TabIndex        =   1
         Top             =   405
         Width           =   1620
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
         Height          =   315
         Left            =   3690
         OleObjectBlob   =   "frmPO.frx":1AC6
         TabIndex        =   3
         Top             =   405
         Width           =   2235
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboDeal 
         Height          =   315
         Left            =   90
         OleObjectBlob   =   "frmPO.frx":2E70
         TabIndex        =   9
         Top             =   1095
         Width           =   2010
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         Left            =   3840
         TabIndex        =   40
         Top             =   195
         Width           =   1875
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   4035
         TabIndex        =   38
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   885
         Width           =   645
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   2520
         TabIndex        =   36
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   315
         Left            =   2595
         TabIndex        =   34
         Top             =   195
         Width           =   570
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   8760
         TabIndex        =   27
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Category(s)"
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
         Left            =   6585
         TabIndex        =   26
         Top             =   195
         Width           =   1350
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deal"
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
         Left            =   855
         TabIndex        =   25
         Top             =   885
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   315
         Left            =   8220
         TabIndex        =   24
         Top             =   195
         Width           =   525
      End
      Begin VB.Label Label8 
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
         Height          =   315
         Left            =   5340
         TabIndex        =   23
         Top             =   1050
         Width           =   570
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   1050
         TabIndex        =   19
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Height          =   315
         Left            =   9855
         TabIndex        =   18
         Top             =   195
         Width           =   555
      End
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
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
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "IN PROCESS"
      Top             =   5535
      Width           =   1260
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   765
      Left            =   90
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   15
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   1349
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
   Begin CoolButtonControl.CoolButton cbDelTo 
      Height          =   480
      Left            =   7110
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   300
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   847
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   90
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3045
      Width           =   7230
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Deliver to"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   8445
      TabIndex        =   46
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tel"
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
      Left            =   4740
      TabIndex        =   45
      Top             =   45
      Width           =   390
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   33
      Top             =   60
      Width           =   225
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
      Height          =   540
      Left            =   495
      TabIndex        =   32
      Top             =   45
      Width           =   4020
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
      Height          =   250
      Left            =   5280
      TabIndex        =   31
      Top             =   45
      Width           =   1530
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
      Height          =   255
      Left            =   5340
      TabIndex        =   30
      Top             =   285
      Width           =   1380
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
      Left            =   4755
      TabIndex        =   29
      Top             =   270
      Width           =   390
   End
End
Attribute VB_Name = "frmPO"
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
Dim bValidPO As Boolean
Dim bValidPOLine As Boolean
Dim lngProductTypeID As Long
Dim lngCategoryID As Long
Dim tlSupplier As z_TextList
Dim lngCurrentExtension As Long
Dim lngCurrentTotal As Long
Dim lngCurrentDepositTotal As Long
Dim lngCurrentVATTotal As Long
Dim lngCurrencyID As Long
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
Dim strPOErrMsg As String
Dim strPOLErrMsg As String
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
    ofrm.component oPO.Memo
    ofrm.Show vbModal
    oPO.SetMemo ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.mnuMemo"
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
    ErrorIn "frmPO.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSection_Click()
    On Error GoTo errHandler
Dim frm As frmSection2
    If Not oPOLine.Product.PID > "" Then Exit Sub
    Set frm = New frmSection2
    frm.component oPOLine.Product
    frm.Show vbModal
    oPOLine.SetSection oPOLine.Product.ProductSections.SectionsAsList
    txtSections = oPOLine.Product.ProductSections.SectionsAsList
    oPOLine.ForceValidation
    mSetfocus txtQtyFirm
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cmdSection_Click", , EA_NORERAISE
'    HandleError
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdShowBudgets_Click()
Dim f As frmBudgetPreview
    If f Is Nothing Then
        Set f = New frmBudgetPreview
        f.Show
    Else
        Unload f
        Set f = Nothing
    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oPO.Status = stInProcess And oPO.IsNew = False)
    Forms(0).mnuDelLine.Enabled = True
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.SetMenu"
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayoutLvw Me.lvwLines, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.mnuSaveLayout"
End Sub


Public Sub component(pSubscriptionOrReplenishment As String, pCancel As Boolean, Optional pPO As a_PO, Optional PID As Long)
    On Error GoTo errHandler
Dim f As frmSupplier
Dim strPos As String
Dim fh As frmSubscriptionOrderHeader
Dim sNotes As String
Dim dteETA As Date

    pCancel = False
    flgLoading = True
    
    If pSubscriptionOrReplenishment = "NS" And pPO Is Nothing Then
        Set fh = New frmSubscriptionOrderHeader
        fh.Show vbModal
        sNotes = fh.Notes
        dteETA = fh.DueDate
        Unload fh
    End If
    
    If pPO Is Nothing Then
        Set oPO = New a_PO
        oPO.BeginEdit
        oPO.SetStatus stInProcess
        oPO.SetMemo sNotes
        oPO.tmpETA = dteETA
        oPO.OrderType = pSubscriptionOrReplenishment
        If PID > 0 Then
            LoadNewSupplier PID
            If oPO.Supplier.Addresses.Count < 1 Then
                MsgBox "There are no addresses for this supplier. You must have at least one address.", , "Cannot do this"
                pCancel = True
                Exit Sub
            Else
                If Not oPO.Supplier.DefaultCurrency Is oPC.Configuration.DefaultCurrency Then
                    If MsgBox("Usually you order from this supplier in " & oPO.Supplier.DefaultCurrency.Description & "." & vbCrLf & "Click OK to continue as usual or Cancel to use " & oPC.Configuration.DefaultCurrency.Description & ".", vbInformation + vbOKCancel, "Confirm currency") = vbOK Then
                        oPO.SetCaptureCurrency oPO.Supplier.DefaultCurrency
                    Else
                        oPO.SetCaptureCurrency oPC.Configuration.DefaultCurrency
                    End If
                End If
            End If
        End If
        If oPO.Supplier.Deals.Count < 1 Then
            If MsgBox("There are no deals for this supplier. Do you want to add a deal?", vbQuestion + vbYesNo, "Can't do this") = vbNo Then
                pCancel = True
                oPO.CancelEdit
                Set oPO = Nothing
                Exit Sub
            Else
                Set f = New frmSupplier
                f.component oPO.Supplier
                f.Show
                oPO.CancelEdit
                pCancel = True
                Set oPO = Nothing
                Exit Sub

            End If
        End If

        ChangeState enAddingRow
    Else
        Set oPO = pPO
        oPO.BeginEdit
        flgLoading = True
        LoadSupplier
        If oPO.DELTOStoreID = 0 Then
            oPO.setDelToStoreID oPC.Configuration.Stores(1).ID
        End If
        flgLoading = False
        If Not oPC.Configuration.Stores.FindStoreByID(oPO.DELTOStoreID) Is Nothing Then
            Me.cbDelTo.Caption = oPC.Configuration.Stores.FindStoreByID(oPO.DELTOStoreID).Description
        End If
        ChangeState enNotEditing
        SetIssueButtonCaption
        Me.Caption = "Purchase order (edit): " & pPO.DOCCode & " to " & pPO.Supplier.NameAndCode(30)
    End If
    oPO.GetStatus
    If oPO.ISForeignCurrency Then
        Me.txtRunningTotal = oPO.TotalLessDiscExtF(False)
        txtCurrencyRates = oPO.CurrencyConversionAsText & "     Value is : " & oPO.TotalLessDiscExtF(True)
        txtCurrencyRates.Visible = True
    Else
        Me.txtRunningTotal = oPO.TotalLessDiscExtF(False)
        txtCurrencyRates.Visible = False
    End If
        SetMenu
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.component(pCancel,pPO,PID)", Array(pCancel, pPO, PID)
End Sub

Private Sub cbDelTo_Click()
    On Error GoTo errHandler
Static iDelIdx As Integer
START:
    If oPO.Supplier.ID = 0 Then Exit Sub
    If iDelIdx = 0 Then iDelIdx = setCurrentAddressIndex("DEL")
    iDelIdx = iDelIdx + 1
    If iDelIdx > oPC.Configuration.Stores.Count Then
        iDelIdx = 1
    End If
    oPO.setDelToStoreID oPC.Configuration.Stores(iDelIdx).ID
    Me.cbDelTo.Caption = oPC.Configuration.Stores(iDelIdx).Description
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cbDelTo_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Function setCurrentAddressIndex(pType As String) As Integer
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To oPC.Configuration.Stores.Count
        If pType = "BILL" Then
            If oPO.BillToAddressID = oPC.Configuration.Stores(i).ID Then
                setCurrentAddressIndex = i
            End If
        Else
            If oPO.DELTOStoreID = oPC.Configuration.Stores(i).ID Then
                setCurrentAddressIndex = i
            End If
        
        End If
    Next
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.setCurrentAddressIndex(pType)", pType
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.setCurrentAddressIndex(pType)", pType
End Function

Private Sub cboProductType_SelectionChanged()
    On Error GoTo errHandler
If flgLoading Then Exit Sub
    oPOLine.ProductTypeID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    If oPC.Configuration.ProductTypes.f3(oPOLine.ProductTypeID) = "" Or oPC.Configuration.ProductTypes.f3(oPOLine.ProductTypeID) = "False" Then
        oPOLine.Product.Seesafe = 0
    Else
        oPOLine.Product.Seesafe = 1
    End If
    lngProductTypeID = oPOLine.ProductTypeID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cboProductType_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboDeal_SelectionChanged()
    On Error GoTo errHandler
    oPOLine.SetDiscount cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 0)
    oPOLine.DealID = CLng(cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 2))
    oPOLine.RecalculateLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cboDeal_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSupplier_Click()
    On Error GoTo errHandler
Dim frm As frmSupplierPreview
    If oPO.Supplier.Name = "" Then Exit Sub
    Set frm = New frmSupplierPreview
    frm.component oPO.Supplier
    frm.Show
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cmdSupplier_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdSupplier_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As New frmSupplierPreview
    If oPO.Supplier.ID > 0 Then
        frm.component oPO.Supplier
        frm.Show
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cbTP_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oTP = Nothing
    Set oCurrentCopy = Nothing
    Set oPO = Nothing
    Set tlSupplier = Nothing
    Set oPOLine = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.Form_Terminate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Sub mnuDelLine()
    On Error GoTo errHandler
    RemoveDetailLine
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuDelLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.mnuDelLine"
End Sub
Private Sub mnuPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_CN
    Set frm = New frmPrintingOptions_CN
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oPO_Valid(pMsg As String)
    On Error GoTo errHandler
    bValidPO = (pMsg = "")
    cmdIssue.Enabled = (bValidPO And oPO.POLines.Count > 0 And vMode = enNotEditing And oPC.ISSUE_PO_ON_THIS_WS And oPO.Status <> stISSUED And oPO.Status <> stCOMPLETE)
 '   cmdSave.Enabled = (bValidPO And oPO.POLines.Count > 0)
    cmdSave.Enabled = (bValidPO And vMode = enNotEditing)
    strPOErrMsg = pMsg
    If vMode = enNotEditing Then
        txtError = strPOErrMsg
    Else
        txtError = strPOLErrMsg
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.oPO_Valid(pMsg)", pMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.oPO_Valid(pMsg)", pMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub oPOLine_Valid(msg As String)
    On Error GoTo errHandler
    cmdEnter.Enabled = (msg = "")
    strPOLErrMsg = msg
    If vMode = enNotEditing Then
        txtError = strPOErrMsg
    Else
        txtError = strPOLErrMsg
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.oPOLine_Valid(Msg)", msg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.oPOLine_Valid(msg)", msg, EA_NORERAISE
    HandleError
End Sub

Private Sub oPO_TotalChange(strtotal As String, strTotalForeign As String)
    On Error GoTo errHandler
    flgLoading = True
    If oPO.CaptureCurrency Is oPC.Configuration.DefaultCurrency Then
        Me.txtRunningTotal = strtotal
    Else
        Me.txtRunningTotal = strTotalForeign
        txtCurrencyRates = oPO.CurrencyConversionAsText & "     Value is : " & strtotal
    End If
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.oPO_TotalChange(strtotal,strTotalForeign)", Array(strtotal, strTotalForeign), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.oPO_TotalChange(strtotal,strTotalForeign)", Array(strtotal, strTotalForeign), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub oPO_Reloadlist()
    On Error GoTo errHandler
    LoadListView
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.oPO_Reloadlist", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.oPO_Reloadlist", , EA_NORERAISE
    HandleError
End Sub
Private Sub oPO_Dirty(pVal As Boolean)
    On Error GoTo errHandler
    If pVal = True Then
        Me.cmdSave.Enabled = (True And Not bFrameEnabled)
        Me.cmdCancel.Caption = "&Cancel"
    Else
        Me.cmdSave.Enabled = False
        Me.cmdCancel.Caption = "&Close"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.oPO_Dirty(pVal)", pVal, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.oPO_Dirty(pVal)", pVal, EA_NORERAISE
    HandleError
End Sub

Private Sub oPOLine_ValueChanges()
    On Error GoTo errHandler
    txtTotal = oPOLine.PLessDiscExtF(False)

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.oPOLine_ValueChanges", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.oPOLine_ValueChanges", , EA_NORERAISE
    HandleError
End Sub


Sub vCanAdd_NobrokenRules()
    On Error GoTo errHandler
    mSetfocus cmdNewRows
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.vCanAdd_NobrokenRules", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.vCanAdd_NobrokenRules", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtRef_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If oPOLine Is Nothing Then Exit Sub
    oPOLine.Ref = txtRef
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtRef_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11100
        Height = 6700
    End If
    LoadControls
    lvwLines.Height = 4850
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
Dim ar() As String

    SetupcboDeal
    LoadDeals
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate
  '  SetLvw
    vMode = enNotEditing
    LoadListView
    
    oPO.SetDirty False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoice.LoadControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.LoadControls"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oProd = Nothing
    UnsetMenu
    If oPO Is Nothing Then Exit Sub
    If oPO.IsEditing Then oPO.CancelEdit
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadNewSupplier(plngTPID As Long)
    On Error GoTo errHandler
    If oPO.SetTP(plngTPID) Then
        With oPO.Supplier
            If Not .OrderToAddress Is Nothing Then
                Me.txtPhone = .OrderToAddress.Phone
                Me.txtFax = .OrderToAddress.Fax
            End If
            Me.txtSuppname = .NameAndCode(15)

        End With
        LoadDeals
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.LoadNewSupplier(plngTPID)", plngTPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.LoadNewSupplier(plngTPID)", plngTPID
End Sub
Public Function SetSupplier(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = oPO.Supplier.Load(pTPID)
    SetSupplier = bSuccess
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.SetSupplier(pTPID)", pTPID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.SetSupplier(pTPID)", pTPID
End Function

'Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
'    On Error GoTo errHandler
'Dim lngColour As Long
'    'A is adding, E is editing
'    bFrameEnabled = pYesNo      'shared for use in all the form
'    If (eMode = enAddingRow Or eMode = enNotEditing) And pYesNo Then
'        Me.txtCode.Enabled = True
'    Else
'        Me.txtCode.Enabled = False
'    End If
'    txtNote.Enabled = pYesNo
'    txtCurrencyRates.Enabled = pYesNo
'    txtPrice.Enabled = pYesNo
'    txtTitle.Enabled = pYesNo
'    txtQtyFirm.Enabled = pYesNo
'    txtQtySS.Enabled = pYesNo
'    txtSections.Enabled = pYesNo
'    cboProductType.Enabled = pYesNo
'    cboDeal.Enabled = pYesNo
'    cboProductType.Enabled = pYesNo
'    cmdEnter.Enabled = pYesNo
'    cmdCancel.Enabled = Not pYesNo
'
'    Me.cmdEnter.Enabled = Not pYesNo
'    Me.cmdCancel.Enabled = Not pYesNo
'    Me.cmdIssue.Enabled = (Not pYesNo) And bValidPO And oPC.ISSUE_PO_ON_THIS_WS
'    Me.cmdSave.Enabled = (Not pYesNo) And bValidPO And oPO.IsDirty
'
'
'    If pYesNo Then
'        lngColour = &HFFFFFF
'    Else
'        lngColour = 14416635
'    End If
'
'    Me.txtCode.BackColor = lngColour
'    Me.txtPrice.BackColor = lngColour
''errHandler:
''    If ErrMustStop Then Debug.Assert False: Resume
''    ErrorIn "frmPO.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.SetEditFrameEnabled(pYesNo,eMode)", Array(pYesNo, eMode)
'End Sub

Private Sub cmdEnter_Click()
    On Error GoTo errHandler
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim strETACode As String
'LogSaveToFile "cmdEnter_Click pos 1"
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        If txtCode.Enabled = True Then mSetfocus txtCode
        Exit Sub
    End If
'LogSaveToFile "cmdEnter_Click pos 2"
    If oPOLine.QtyFirm + oPOLine.QtySS <= oPOLine.Product.QtyOnHand Then
        If MsgBox("There is quantity on hand of :" & CStr(oPOLine.Product.QtyOnHand) & vbCrLf & "Do you want to continue to post?", vbInformation + vbOKCancel, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    oPOLine.ApplyEdit
'LogSaveToFile "cmdEnter_Click pos 4"
    oPOLine.BeginEdit
'LogSaveToFile "cmdEnter_Click pos 5"

    If vMode = enAddingRow Then
'LogSaveToFile "cmdEnter_Click pos 6"
        lvwLines.ListItems.Add Key:=oPOLine.Key
'LogSaveToFile "cmdEnter_Click pos 6.1"
        LoadListViewLine lvwLines.ListItems(lvwLines.ListItems.Count), oPOLine
'LogSaveToFile "cmdEnter_Click pos 6.2"
        lvwLines.Refresh
'LogSaveToFile "cmdEnter_Click pos 6.3"
        ChangeState enAddingRow
'LogSaveToFile "cmdEnter_Click pos 6.4"
        mSetfocus txtCode
    
    ElseIf vMode = eneditingrow Then
'LogSaveToFile "cmdEnter_Click pos 7"
        LoadListViewLine lvwLines.ListItems(lngSelectedRowIndex), oPOLine
        ChangeState enNotEditing
    
    End If
'LogSaveToFile "cmdEnter_Click pos 8"
    oPO.LevelNewValues oPOLine
'LogSaveToFile "cmdEnter_Click pos 9"
    oPO.CalculateTotals
'LogSaveToFile "cmdEnter_Click pos 10"
    oPO.GetStatus
'LogSaveToFile "cmdEnter_Click pos 11"
    lvwLines.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdEnter_Click", , EA_NORERAISE, , , "line number", Array(Erl())
    HandleError
End Sub
Private Sub ChangeState(pToMode As EnumMode)
    On Error GoTo errHandler
Dim lngColour As Long
Dim strETACode As String

    vMode = pToMode

    Select Case pToMode
    Case eneditingrow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtPrice.Enabled = True
        txtRef.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQtyFirm.Enabled = True
        txtQtySS.Enabled = True
        cmdEnter.Enabled = False
        cmdCancel.Enabled = False
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Caption = "&Stop"
        cmdNewRows.Enabled = (oPO.POLines.Count > 0)
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        UnsetMenu
        fr1.ZOrder 1
    Case enAddingRow
        fr1.Visible = True
        txtCode.Enabled = True
        txtNote.Enabled = True
        txtPrice.Enabled = True
        txtRef.Enabled = True
        txtTitle.Enabled = True
        txtTotal.Enabled = True
        txtQtyFirm.Enabled = True
        txtQtySS.Enabled = True
        txtETA.Visible = Not (oPO.OrderType = "NS")
        Label10.Visible = txtETA.Visible
        txtError = ""
        flgLoading = True
        txtRef = ""
        flgLoading = False
        cmdEnter.Enabled = IIf(oPC.AllowZeropricedPOLines, True, False)
        cmdCancel.Enabled = True
        cmdIssue.Enabled = False
        cmdSave.Enabled = False
        cmdNewRows.Enabled = (oPO.POLines.Count > 0) Or oPO.OrderType = "NS"
        cmdNewRows.Caption = "&Stop"
        lvwLines.Enabled = False
        lvwLines.Height = 2200
        ClearLineControls
        fr1.ZOrder 1
        If Not oPOLine Is Nothing Then
            strETACode = oPOLine.ETACode
        End If
        mSetfocus txtCode
        Set oPOLine = oPO.POLines.Add
        If oPO.tmpETA < Now() Then
            oPOLine.SetETA DateAdd("d", 1, Date)
        Else
            oPOLine.SetETA oPO.tmpETA
        End If
        oPOLine.TRID = oPO.TRID
        oPOLine.ProductTypeID = lngProductTypeID
        oPOLine.SetQtyFirm 1
        oPOLine.SetQtySS 0
        oPOLine.SetETA strETACode
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
        cmdSave.Enabled = True
        cmdNewRows.Enabled = True  '(oInvoice.InvoiceLines.Count > 0)
        cmdNewRows.Caption = "&Add"
        lvwLines.Enabled = True
        lvwLines.Height = 4000
        SetMenu
        fr1.ZOrder 1
    End Select
    If Not oPO.IsDirty Then
        cmdCancel.Caption = "&Close"
    Else
        cmdCancel.Caption = "&Cancel"
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoice.ChangeState(pToMode)", pToMode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.ChangeState(pToMode)", pToMode
End Sub


Private Sub cmdNewRows_Click()
    On Error GoTo errHandler
Dim lr As Long
    'Editing: A line has been seleted from the listview for editing
    'Adding: a new line is being prepared
    'notediting: in editing mode but no row selected
'    If Not oProd Is Nothing Then
'        If oProd.IsEditing Then oProd.CancelEdit
'    End If
    If vMode = eneditingrow Then
        ChangeState enNotEditing
    ElseIf vMode = enAddingRow Then
        ChangeState enNotEditing
    ElseIf vMode = enNotEditing Then
        ChangeState enAddingRow
    End If

    ClearLineControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdNewRows_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    For i = 1 To lvwLines.ColumnHeaders.Count
        lvwLines.ColumnHeaders(i).Width = GetSetting("PBKS", Me.Name, CStr(i), lvwLines.ColumnHeaders(i).Width)
    Next
    lvwLines.ListItems.Clear
    
    For i = 1 To oPO.POLines.Count
        If oPO.POLines(i).Fulfilled = "OS" Or oPO.POLines(i).Fulfilled = "" Then
        Set lstItem = lvwLines.ListItems.Add
        LoadListViewLine lstItem, oPO.POLines(i)
        End If
    Next i
EXIT_Handler:
    Set lstItem = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.LoadListView"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.LoadListView"
End Sub
Private Sub LoadListViewLine(lstItem As ListItem, oPOLine As a_POL)
    On Error GoTo errHandler
Dim currPrice As Currency
    With oPOLine
        lstItem.text = .ProductCodeF
        lstItem.Key = .Key
        lstItem.SubItems(4) = .Ref
        lstItem.SubItems(1) = .TitleAuthor
        lstItem.SubItems(2) = .QtyFirm
        lstItem.SubItems(3) = .QtySS
        lstItem.SubItems(6) = .DiscountF
        lstItem.SubItems(8) = Format(.Key, "@@@@@@@@@@")
        lstItem.SubItems(9) = .EAN
        If oPC.Configuration.DefaultCurrency Is oPO.CaptureCurrency Then
            lstItem.SubItems(5) = .PriceF(False)
            lstItem.SubItems(7) = .PLessDiscExtF(False)
        Else
            lstItem.SubItems(5) = .PriceF(True)
            lstItem.SubItems(7) = .PLessDiscExtF(True)
        End If
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.LoadListViewLine(lstItem,oPOLine)", Array(lstItem, oPOLine), , , "Line number", Array(Erl())
End Sub
Private Sub lvwLines_DblClick()
    On Error GoTo errHandler
Dim strPos As String
Dim tmpPOLine As a_POL
Dim lngPOLID As Long
Dim bOK As Boolean

'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    If lvwLines.SelectedItem.index < 1 Then Exit Sub
    
    lngEditingIdx = lvwLines.SelectedItem.Key
    If oPO.Status <> stInProcess Then
        'Store Current Row's ID
        lngPOLID = oPO.POLines(lvwLines.SelectedItem.Key).POLID
        'point to exisiting line
        Set tmpPOLine = oPO.POLines(lvwLines.SelectedItem.Key)
        'Create new oPOLINE and copy values from old
        Set oPOLine = Nothing
        Set oPOLine = oPO.POLines.Add
        oPOLine.SetReplacementForLineID lngPOLID
        oPOLine.TRID = tmpPOLine.TRID
        oPOLine.ProductTypeID = tmpPOLine.ProductTypeID
        oPOLine.SetQtyFirm tmpPOLine.QtyFirm
        oPOLine.SetQtySS tmpPOLine.QtySS
        oPOLine.SetETA tmpPOLine.ETA
        oPOLine.COLID = tmpPOLine.COLID
        oPOLine.DealID = tmpPOLine.DealID
        oPOLine.SetDiscount tmpPOLine.Discount
        oPOLine.EAN = tmpPOLine.EAN
        oPOLine.ProductCode = tmpPOLine.ProductCode
        oPOLine.ProductCodeF = tmpPOLine.ProductCodeF
        oPOLine.ForeignPrice = tmpPOLine.ForeignPrice
        oPOLine.SetPriceLong tmpPOLine.Price(oPO.ISForeignCurrency), oPO.ISForeignCurrency
        oPOLine.LastAction = tmpPOLine.LastAction
        oPOLine.MainAuthor = tmpPOLine.MainAuthor
        oPOLine.Note = tmpPOLine.Note
        oPOLine.PID = tmpPOLine.PID
        oPOLine.Product = tmpPOLine.Product
        oPOLine.VATRate = tmpPOLine.VATRate
        oPOLine.SetSection tmpPOLine.Section
        oPOLine.Ref = tmpPOLine.Ref
        oPOLine.Fulfilled = "OS"
        oPOLine.Title = tmpPOLine.Title
        Set tmpPOLine = Nothing
    Else
        Set oPOLine = Nothing
        Set oPOLine = oPO.POLines(lvwLines.SelectedItem.Key)
    End If
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    
    txtOHOO = ""
    bOK = oPOLine.SetLineProduct(oPOLine.PID, , , True)
    If bOK Then
        If oPOLine.Product.QtyCopiesOnHand > 0 Or oPOLine.Product.QtyonOrder > 0 Then
            txtOHOO = "OH: " & oPOLine.Product.QtyOnHandF & " / " & "OO: " & oPOLine.Product.QtyOnOrderF
        End If
    End If
        
    
    
    ChangeState eneditingrow
    
    strPos = "2"
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    txtCode = oPOLine.CodeForEditing
    txtTitle = oPOLine.Title
    txtPrice = oPOLine.Price(oPO.ISForeignCurrency)
    txtQtyFirm = oPOLine.QtyFirm
    txtQtySS = oPOLine.QtySS
    txtNote = oPOLine.Note
    txtRef = oPOLine.Ref
    txtTotal = oPOLine.PLessDiscExtF(oPO.ISForeignCurrency)
    Me.txtETA = oPOLine.ETAF
    Me.txtSections = oPOLine.Product.ProductSections.SectionsAsList
    If oPOLine.DealID > 0 Then
        On Error Resume Next
        cboDeal.Items.SelectItem(cboDeal.Items.FindItem(oPOLine.DealID, 2)) = True
        On Error GoTo errHandler
    Else
        If cboDeal.Items.ItemCount > 0 Then
            On Error Resume Next
            cboDeal.Items.SelectItem(cboDeal.Items.FirstVisibleItem) = True
            On Error GoTo errHandler
        End If
    End If
      On Error Resume Next
    cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(oPOLine.ProductTypeID), 0)) = True
      On Error GoTo errHandler
    mSetfocus txtPrice
    lvwLines.Height = 2600
    cmdNewRows.Caption = "&Stop edit"
    oPOLine.GetStatus
    AutoSelect txtPrice
    
    If cboDeal.Items.SelectCount > 0 Then
        oPOLine.SetDiscount cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 0)
        oPOLine.DealID = CLng(cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 2))
        oPOLine.RecalculateLine
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.lvwLines_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_Click()
    On Error GoTo errHandler
    If lvwLines Is Nothing Then Exit Sub
    If lvwLines.SelectedItem Is Nothing Then Exit Sub
    
    If lvwLines.SelectedItem.index > 0 Then
    On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText Left(lvwLines.SelectedItem.SubItems(9), ISBNLENGTH)
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.lvwLines_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.lvwLines_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If oPO.Supplier Is Nothing Then
        MsgBox "Please enter a Supplier before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cboTP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    txtNote = HandleTextWithBites(txtNote)
    On Error Resume Next
    oPOLine.SetNote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oPOLine.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oPOLine.SetNote(txtNote)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oPOLine.Note
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtNote_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    oPO.SetStatus stVOID
    txtStatus = "Void"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.mnuVoid"
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler

    txtCode = FNS(txtCode)
    If txtCode = "" Then Exit Sub

    If Not IsRecognizedCode(txtCode) Then
        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
        Cancel = True
        Exit Sub
    End If

    If Not ValidatePOLineFromCode(txtCode, Cancel) Then
        Cancel = True
    End If

EXIT_Handler:
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
    Resume EXIT_Handler
End Sub
'Private Sub txtCode_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'Dim pQty As Integer
'Dim pApproID As Long
'Dim bOK  As Boolean
'Dim oPCode As New z_ProdCode
'Dim lngQtyAlreadyOnThisOrder As Long
'Dim msg As String
'Dim frm As frmProduct
'START:
'    txtCode = FNS(txtCode)
'    If txtCode = "" Then Exit Sub
'    If Not (IsISBN13(txtCode) Or IsISBN10(txtCode) Or IsHashCode(txtCode) Or IsPrivateCode(txtCode)) Then
'        MsgBox "This is an invalid code, retry.", vbInformation, "Warning"
'        Cancel = True
'        GoTo EXIT_Handler
'    End If
'
'    lngQtyAlreadyOnThisOrder = CountQtyOnExistingRowsForSameCode
'    If lngQtyAlreadyOnThisOrder > 0 Then
'        msg = IIf(lngQtyAlreadyOnThisOrder = 1, "There is already an item with this code on this order. Continue?", "There are already " & CStr(lngQtyAlreadyOnThisOrder) & " items with this code on this order. Continue?")
'        If MsgBox(msg, vbInformation + vbYesNo, "Warning - already on order") = vbNo Then
'            Cancel = True
'            GoTo EXIT_Handler
'        End If
'    End If
''
'    bOK = oPOLine.SetLineProduct("", txtCode, oPO.Supplier.DefaultETA)
'    If bOK Then
'        If oPOLine.Product.Obsolete Then
'            MsgBox "Please note: this product is obsolete and may not be reordered. To reorder, change the obsolete status on the product record.", vbInformation, "Can't order this product"
'            Cancel = True
'            AutoSelect txtCode
'            Exit Sub
'        End If
'        If oPOLine.Product.Status <> "IP" Then
'            If MsgBox("The product's status is not 'In Print'.", vbInformation + vbOKCancel, "CWarning") = vbCancel Then
'               Cancel = True
'               AutoSelect txtCode
'               Exit Sub
'             End If
'        End If
'    End If
'    If bOK Then
'        txtQtyFirm = oPOLine.QtyFirmF
'        txtQtySS = oPOLine.QtySSF
'        txtPrice = oPOLine.Price(oPO.ISForeignCurrency)
'        txtTotal = oPOLine.PLessDiscExtF(oPO.ISForeignCurrency)
'        txtCode = oPOLine.EAN
'        txtETA = oPOLine.ETAF
'        txtTitle = oPOLine.TitleAuthor
'        txtSections = oPOLine.Product.ProductSections.SectionsAsList
'        txtOHOO = "OH: " & oPOLine.Product.QtyOnHandF & " / " & "OO: " & oPOLine.Product.QtyOnOrderF
'    Else
'        If CheckThisPoint(M_NEWPRODUCTINADHOCFORM) Then
'            If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then
'                Cancel = True
'                Exit Sub
'            End If
'        End If
'        If GetAdhocDetails() Then
'            GoTo START
'        Else
'           MsgBox "Cannot find item", vbOKOnly + vbInformation, "Finding stock item"
'           Cancel = True
'           Exit Sub
'        End If
'    End If
'    If cboDeal.Items.ItemCount > 0 Then
'        If oPOLine.DealID > 0 Then
'            If cboDeal.Items.FindItem(oPOLine.DealID, 2) = 0 Then
'                MsgBox "The deal by which this title was previously ordered is not available. " & vbCrLf & "Possibly this is a different supplier.", vbInformation, "Warning"
'                oPOLine.DealID = oPO.Supplier.Deals(1).ID 'CLng(cboDeal.Items.CellCaption(cboDeal.Items(1), 2))
'                cboDeal.Items.SelectItem(cboDeal.Items.FindItem(oPO.Supplier.Deals(1).ID, 2)) = True
'            Else
'                cboDeal.Items.SelectItem(cboDeal.Items.FindItem(oPOLine.DealID, 2)) = True
'            End If
'        Else
'            cboDeal.Items.SelectItem(cboDeal.Items(0)) = True
'        End If
'    End If
'    If Me.cboProductType.Items.ItemCount > 0 Then
'        If oPOLine.ProductTypeID > 0 Then
'            cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(oPOLine.ProductTypeID), 0)) = True
'        ElseIf lngProductTypeID > 0 Then
'            cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(lngProductTypeID), 0)) = True
'        Else
'            cboProductType.Items.SelectItem(cboProductType.Items(0)) = True
'            lngProductTypeID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
'        End If
'    End If
'    oPOLine.GetStatus
'
'EXIT_Handler:
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
Private Function IsRecognizedCode(code As String) As Boolean
    IsRecognizedCode = IsISBN13(code) Or IsISBN10(code) Or IsHashCode(code) Or IsPrivateCode(code)
End Function
Private Function ValidatePOLineFromCode(code As String, ByRef Cancel As Boolean) As Boolean
    On Error GoTo errHandler

    Dim lngQtyAlreadyOnThisOrder As Long
    Dim msg As String
    Dim bOK As Boolean

    ValidatePOLineFromCode = False

Retry:
    lngQtyAlreadyOnThisOrder = CountQtyOnExistingRowsForSameCode
    If lngQtyAlreadyOnThisOrder > 0 Then
        msg = "There " & IIf(lngQtyAlreadyOnThisOrder = 1, "is", "are") & " already " & _
              lngQtyAlreadyOnThisOrder & " item(s) with this code on this order. Continue?"
        If MsgBox(msg, vbInformation + vbYesNo, "Warning - already on order") = vbNo Then
            Exit Function
        End If
    End If

    bOK = oPOLine.SetLineProduct("", code, oPO.Supplier.DefaultETA)

    If bOK Then
        If Not HandleProductWarnings(Cancel) Then Exit Function
        FillPOLineFields
        ValidatePOLineFromCode = True
        Exit Function
    End If

    If CheckThisPoint(M_NEWPRODUCTINADHOCFORM) Then
        If Not SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", _
               "You do not have permission to create new stock items (or your signature is invalid).") Then
            Exit Function
        End If
    End If

    If GetAdhocDetails() Then
        code = FNS(txtCode)
        GoTo Retry
    Else
        MsgBox "Cannot find item", vbOKOnly + vbInformation, "Finding stock item"
        Exit Function
    End If

Exit Function

errHandler:
    HandleError
    Cancel = True
End Function

Private Function HandleProductWarnings(ByRef Cancel As Boolean) As Boolean
    HandleProductWarnings = False

    If oPOLine.Product Is Nothing Then Exit Function

    If oPOLine.Product.Obsolete Then
        MsgBox "This product is obsolete and may not be reordered. To reorder, change the obsolete status.", vbInformation
        Cancel = True
        AutoSelect txtCode
        Exit Function
    End If

    If oPOLine.Product.Status <> "IP" Then
        If MsgBox("The product's status is not 'In Print'.", vbInformation + vbOKCancel, "Warning") = vbCancel Then
            Cancel = True
            AutoSelect txtCode
            Exit Function
        End If
    End If

    HandleProductWarnings = True
End Function
Private Sub FillPOLineFields()
    On Error Resume Next

    txtQtyFirm = oPOLine.QtyFirmF
    txtQtySS = oPOLine.QtySSF
    txtPrice = oPOLine.Price(oPO.ISForeignCurrency)
    txtTotal = oPOLine.PLessDiscExtF(oPO.ISForeignCurrency)
    txtCode = oPOLine.EAN
    txtETA = oPOLine.ETAF
    txtTitle = oPOLine.TitleAuthor

    If Not oPOLine.Product Is Nothing Then
        txtSections = oPOLine.Product.ProductSections.SectionsAsList
        txtOHOO = "OH: " & oPOLine.Product.QtyOnHandF & " / OO: " & oPOLine.Product.QtyOnOrderF
    End If

    SetupDeals
    oPOLine.SetDiscount cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 0)
    oPOLine.DealID = CLng(cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 2))

    
    SetupProductTypes

    oPOLine.GetStatus
End Sub
Private Sub SetupDeals()
    If cboDeal.Items.ItemCount = 0 Then Exit Sub

    If oPOLine.DealID > 0 Then
        Dim index As Long
        index = cboDeal.Items.FindItem(oPOLine.DealID, 2)
        If index = 0 Then
            MsgBox "The previous deal is not available. Possibly different supplier.", vbInformation
            cboDeal.Items.SelectItem(0) = True
        Else
            cboDeal.Items.SelectItem(index) = True
        End If
    Else
        If cboDeal.Items.SelectedItem = 0 Then
            cboDeal.Items.SelectItem(0) = True
        End If

    End If
End Sub
Private Sub SetupProductTypes()
    If Me.cboProductType.Items.ItemCount = 0 Then Exit Sub

    If oPOLine.ProductTypeID > 0 Then
        cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(oPOLine.ProductTypeID), 0)) = True
    ElseIf lngProductTypeID > 0 Then
        cboProductType.Items.SelectItem(cboProductType.Items.FindItem(oPC.Configuration.ProductTypes.Item(lngProductTypeID), 0)) = True
    Else
        cboProductType.Items.SelectItem(0) = True
        lngProductTypeID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    End If
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

Private Function CountQtyOnExistingRowsForSameCode() As Long
    On Error GoTo errHandler
Dim tPOL As a_POL
Dim lngCount As Long

    lngCount = 0
    For Each tPOL In oPO.POLines
        If tPOL.EAN = FNS(txtCode) Or tPOL.ProductCode = FNS(txtCode) Then
            If tPOL.IsDeleted = False And tPOL.DateReplaced = CDate(0) Then
                lngCount = tPOL.QtyFirm + tPOL.QtySS
            End If
        End If
    Next
    
    CountQtyOnExistingRowsForSameCode = lngCount
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.CountQtyOnExistingRowsForSameCode"
End Function
Private Sub RemoveDetailLine()
    On Error GoTo errHandler
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.RemoveDetailLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.RemoveDetailLine"
End Sub

Private Sub LoadSupplier()
    On Error GoTo errHandler
    With oPO
        txtStatus = .StatusF
        Me.txtSuppname = .Supplier.NameAndCode(20)
        If Not .Supplier.BillTOAddress Is Nothing Then
            Me.txtPhone = .Supplier.BillTOAddress.Phone
            Me.txtFax = .Supplier.BillTOAddress.Fax
        End If
    End With
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.LoadSupplier"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.LoadSupplier"
End Sub


Private Sub SavePO()
    On Error GoTo errHandler
  '  If oPO.polines.IsEditing Then oPO.polines.ApplyEdit
    oPO.Post
    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
   ' Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.SavePO"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.SavePO"
End Sub

Public Sub PrintOrder()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim blnHideVAT As Boolean
Dim iCurrency As Integer

    
    Me.MousePointer = vbHourglass
    oPO.Load oPO.TRID, False
    blnDiscount = False ' TO BE REMOVED ON COMPLETION????
    
    If blnNoCNLines Then
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.PrintOrder"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.PrintOrder"
End Sub
Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim blnDeposit As Boolean
Dim blnDiscount As Boolean
Dim blnRoundedUp As Boolean
Dim blnNoCNLines As Boolean
Dim iCurrency As Integer
Dim frmDte As frmTRDate
Dim strResult As String
Dim frm As frmPOPreview
    If oPO.TotalPayable(False) = 0 Then
        If MsgBox("The total for this order is zero. Do you wish to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Sign this purchase order", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oPO.Status = stInProcess Then
            If MsgBox("Issue this purchase order?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    If oPC.AllowPODateOverride Then
        Set frmDte = New frmTRDate
        frmDte.component Date
        frmDte.Show vbModal
        oPO.DOCDate = StartOfDay(frmDte.InvoiceDate)
        Unload frmDte
        oPO.CaptureDate = Now()
    Else
        If oPO.DOCDate < CDate("1950-01-01") Then
            oPO.DOCDate = Date
            oPO.CaptureDate = Now()
        End If
    End If
    
    WaitMsg "Issuing purchase order  . . .", True, Me
    oPO.SetStatus stISSUED
    oPO.StaffID = gSTAFFID
    strResult = oPO.Post  'contains the SAVE action
    If strResult = "ERROR" Then
        MsgBox "This action has failed. Contact support"
        Exit Sub
    End If
    Set frm = New frmPOPreview
    frm.component oPO.TRID
    frm.Show
    WaitMsg "", False, Me
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cmdIssue_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdSave_Click()
    On Error GoTo errHandler
   ' If oPOLine.IsEditing Then oPOLine.CancelEdit
    If oPO.Status <> stISSUED And oPO.Status <> stCOMPLETE Then
        oPO.SetStatus stInProcess
    End If
    SavePO
    LoadListView
    oPO.BeginEdit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    mSetfocus cmdNewRows
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cmdSave_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
Dim frm As frmPOPreview
    If cmdCancel.Caption <> "&Close" Then
        If MsgBox("You wish to cancel this purchase order?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
        oPO.CancelEdit
    End If
    If cmdCancel.Caption = "&Close" And oPO.TRID > 0 Then
        Set frm = New frmPOPreview
        frm.component oPO.TRID
        frm.Show
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub ClearLineControls()
    On Error GoTo errHandler
    flgLoading = True
    txtCode = ""
    txtPrice = ""
    txtTitle = ""
    txtRef = ""
    txtTotal = ""
    txtNote = ""
    txtQtyFirm = ""
    txtQtySS = ""
    txtSections = ""
    cboProductType.Items.SelectItem(cboProductType.Items(0)) = True
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.ClearLineControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.ClearLineControls"
End Sub


Private Sub txtPrice_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPrice
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtPrice_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oPOLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    oPO.CalculateTotals
    txtTotal = oPOLine.PLessDiscExtF(oPO.ISForeignCurrency)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtyFirm_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtyFirm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtQtyFirm_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtyFirm_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oPOLine.SetQtyFirm(txtQtyFirm) Then
        Cancel = True
    End If
    If oPOLine.QtyFirm > 0 And oPOLine.Product.Seesafe Then
        MsgBox "This product is usually ordered seesafe only", vbExclamation, "Warning"
    End If
    oPO.CalculateTotals
    txtTotal = oPOLine.PLessDiscExtF(oPO.ISForeignCurrency)
    Me.txtRunningTotal = oPO.TotalLessDiscExtF(oPO.ISForeignCurrency)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtQtyFirm_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtQtyFirm_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtQtySS_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtQtySS
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtQtySS_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtQtySS_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQtySs_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oPOLine.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    oPO.CalculateTotals
    txtTotal = oPOLine.PLessDiscExtF(oPO.ISForeignCurrency)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtQtySs_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub SetIssueButtonCaption()
    On Error GoTo errHandler
        cmdIssue.Enabled = True
       If oPO.StatusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oPO.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Enabled = False
            'cmdIssue.Caption = "Print"
        End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.SetIssueButtonCaption"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.SetIssueButtonCaption"
End Sub

Private Sub lvwLines_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error GoTo errHandler
   ' When a ColumnHeader object is clicked, the ListView control is
   ' sorted by the subitems of that column.
   ' Set the SortKey to the Index of the ColumnHeader - 1
   lvwLines.SortKey = ColumnHeader.index - 1
   ' Set Sorted to True to sort the list.
    If lvwLines.SortOrder = lvwAscending Then
        lvwLines.SortOrder = lvwDescending
    Else
        lvwLines.SortOrder = lvwAscending
    End If
   lvwLines.Sorted = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.lvwLines_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
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


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.SetLvw"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.SetLvw"
End Sub
Sub SetupcboDeal()
    On Error GoTo errHandler
    cboDeal.BeginUpdate
    cboDeal.WidthList = 152
    cboDeal.HeightList = 162
    cboDeal.AutoDropDown = True
    cboDeal.AllowSizeGrip = False
    cboDeal.AllowHResize = False
    cboDeal.FullRowSelect = True
    cboDeal.BackColorLock = Me.BackColor
    cboDeal.Alignment = LeftAlignment
    cboDeal.UseTabKey = False
    cboDeal.Columns.Add "Discount"
    cboDeal.Columns.Add "Description"
    cboDeal.Columns.Add ""
    cboDeal.Columns(0).Width = 45
    cboDeal.Columns(0).AllowSizing = False
    cboDeal.Columns(1).Width = 100
    cboDeal.Columns(1).AllowSizing = False
    cboDeal.Columns(2).Width = 0
    cboDeal.Columns(2).AllowSizing = False
    cboDeal.Columns(2).Visible = False
    cboDeal.EndUpdate
    
    
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AutoDropDown = True
    cboProductType.AllowSizeGrip = False
    cboProductType.AllowHResize = False
    cboProductType.FullRowSelect = True
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.Alignment = LeftAlignment
    cboProductType.UseTabKey = False
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add "Seesafe"
    cboProductType.Columns(0).Width = 190
    cboProductType.Columns(0).AllowSizing = False
    cboProductType.Columns(1).Width = 0
    cboProductType.Columns(1).AllowSizing = False
    cboProductType.EndUpdate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.SetupcboDeal"
End Sub

Private Sub LoadDeals()
    On Error GoTo errHandler
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
        ar(1, i) = oDL.Description
        ar(0, i) = oDL.DiscountF
        ar(2, i) = oDL.ID
        i = i + 1
    Next
    cboDeal.PutItems ar
    cboDeal.EndUpdate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.LoadDeals"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.LoadDeals"
End Sub
'Private Sub LoadcboRef()
'Dim i As Integer
'Dim oD As d_COLine
'Dim ar()
'    i = 0
'    If oPOLine.COLsPerPID.Count < 1 Then
'        Exit Sub
'    End If
'    For Each oD In oPOLine.COLsPerPID
'        i = i + 1
'    Next
'    ReDim ar(3, i - 1)
'    i = 0
'    cboRef.BeginUpdate
'    cboRef.Items.RemoveAllItems
'    For Each oD In oPOLine.COLsPerPID
'        ar(1, i) = oD.DocCode
'        ar(0, i) = oD.Ref
'        ar(2, i) = oD.Qty
'        ar(3, i) = oD.COLID
'        i = i + 1
'    Next
'    cboRef.PutItems ar
'    cboRef.EndUpdate
'End Sub

Private Sub txtETA_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtETA")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtETA_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtETA_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtETA_LostFocus()
    On Error GoTo errHandler
    txtETA = oPOLine.ETAF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtETA_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtETA_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtETA_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
        If Not oPOLine.SetETA(txtETA) Then
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.txtETA_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPO.txtETA_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


