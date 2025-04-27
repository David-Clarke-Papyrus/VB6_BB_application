VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPO 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase order"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   Icon            =   "frmPObackup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   11400
   Begin MSComctlLib.ListView lvwLines 
      Height          =   2580
      Left            =   105
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   435
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
      NumItems        =   8
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
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Save"
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
      Picture         =   "frmPObackup.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
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
      Left            =   2115
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5340
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
      Height          =   675
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5355
      Width           =   870
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
      Left            =   7500
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmPObackup.frx":0914
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5370
      Width           =   1110
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
      ForeColor       =   &H8000000D&
      Height          =   250
      Left            =   9585
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Frame fr1 
      BackColor       =   &H00E0E0E0&
      Height          =   2025
      Left            =   120
      TabIndex        =   15
      Top             =   3285
      Width           =   10710
      Begin VB.TextBox txtETA 
         Alignment       =   2  'Center
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
         Left            =   3990
         TabIndex        =   9
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   1155
         Width           =   1410
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1560
      End
      Begin VB.TextBox txtRef 
         Alignment       =   2  'Center
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   3630
         TabIndex        =   3
         Top             =   450
         Width           =   1560
      End
      Begin VB.TextBox txtQtySS 
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   7905
         TabIndex        =   6
         Top             =   450
         Width           =   735
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboDeal 
         Height          =   375
         Left            =   135
         OleObjectBlob   =   "frmPObackup.frx":0E9E
         TabIndex        =   8
         Top             =   1170
         Width           =   1815
      End
      Begin VB.TextBox txtQtyFirm 
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   7140
         TabIndex        =   5
         Top             =   450
         Width           =   735
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
         ForeColor       =   &H8000000D&
         Height          =   720
         Left            =   5475
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1170
         Width           =   4260
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
         Height          =   735
         Left            =   9780
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1215
         Width           =   840
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1590
         Width           =   5220
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   8670
         TabIndex        =   7
         Top             =   450
         Width           =   1305
      End
      Begin VB.TextBox txtCode 
         Alignment       =   2  'Center
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   105
         TabIndex        =   1
         Top             =   465
         Width           =   1560
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboRef 
         Height          =   390
         Left            =   1710
         OleObjectBlob   =   "frmPObackup.frx":20C4
         TabIndex        =   2
         Top             =   465
         Width           =   1905
      End
      Begin EXCOMBOBOXLibCtl.ComboBox cboCAT 
         Height          =   390
         Left            =   5220
         OleObjectBlob   =   "frmPObackup.frx":32EA
         TabIndex        =   4
         Top             =   450
         Width           =   1905
      End
      Begin VB.Label lblCurrName 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   9645
         TabIndex        =   39
         Top             =   195
         Width           =   960
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "ETA"
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
         Height          =   315
         Left            =   4230
         TabIndex        =   38
         ToolTipText     =   "e.g. 3w = 3 weeks, 1m = 1 month"
         Top             =   885
         Width           =   645
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
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
         Height          =   255
         Left            =   2670
         TabIndex        =   36
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   4125
         TabIndex        =   34
         Top             =   195
         Width           =   570
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   7830
         TabIndex        =   27
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Category."
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
         Height          =   315
         Left            =   5715
         TabIndex        =   26
         Top             =   195
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deal"
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
         Height          =   315
         Left            =   855
         TabIndex        =   25
         Top             =   885
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   7200
         TabIndex        =   24
         Top             =   195
         Width           =   525
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Note:"
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
         Height          =   315
         Left            =   5460
         TabIndex        =   23
         Top             =   885
         Width           =   570
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   660
         TabIndex        =   18
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
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
         Height          =   315
         Left            =   9000
         TabIndex        =   17
         Top             =   195
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
      TabIndex        =   14
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
      Left            =   9735
      Picture         =   "frmPObackup.frx":4510
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5370
      UseMaskColor    =   -1  'True
      Width           =   1110
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   375
      Left            =   60
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   30
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   661
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
      Height          =   375
      Left            =   9315
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   30
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      ForeColor       =   &H00008000&
      Height          =   250
      Left            =   90
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3060
      Width           =   7230
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4725
      Picture         =   "frmPObackup.frx":465A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   360
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
      Left            =   120
      TabIndex        =   33
      Top             =   60
      Width           =   225
   End
   Begin VB.Label txtSuppname 
      BackColor       =   &H00E0E0E0&
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
      Left            =   495
      TabIndex        =   32
      Top             =   45
      Width           =   4020
   End
   Begin VB.Label txtPhone 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5280
      TabIndex        =   31
      Top             =   45
      Width           =   1530
   End
   Begin VB.Label txtFax 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7605
      TabIndex        =   30
      Top             =   45
      Width           =   1590
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   7020
      TabIndex        =   29
      Top             =   30
      Width           =   390
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
Dim tlCategories As z_TextList
Dim oCurrentCopy
Dim bValidCN As Boolean
Dim bValidCNLine As Boolean
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
Dim WithEvents vCanAdd As z_BrokenRules
Attribute vCanAdd.VB_VarHelpID = -1
Dim WithEvents vCanIssue As z_BrokenRules
Attribute vCanIssue.VB_VarHelpID = -1

'Private Sub cboMatch_SelectionChanged()
'    oPOLine.Discount = cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem(0), 2)
'    oPOLine.DEALID = cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem(0), 1)
'End Sub
'
Public Sub Component(pCancel As Boolean, Optional pPO As a_PO, Optional pID As Long)
Dim ar() As String
    pCancel = False
    flgLoading = True
    SetupcboDeal
    cboCat.BeginUpdate
    tlCategories.CollectionAsArray ar
    cboCat.PutItems ar
    cboCat.EndUpdate
    If pPO Is Nothing Then
        Set oPO = New a_PO
        oPO.beginedit
        lvwLines.Enabled = False
        SetControlsForNew
        If pID > 0 Then
            LoadNewSupplier pID
        End If
        If oPO.Supplier.Deals.Count < 1 Then
            MsgBox "There are no deals for this supplier. You cannot continue"
            pCancel = True
        End If
        lvwLines.Height = 2200
        cmdNewRows.Caption = "&Stop"
        vMode = enAddingRow
        SetEditFrameEnabled True, vMode
        Me.txtCode.SetFocus
        Set oPOLine = oPO.POLines.Add
        oPOLine.SetETA DateAdd("d", oPO.Supplier.DefaultETA, Date)
    Else
        Set oPO = pPO
        oPO.beginedit
        LoadSupplier
        If oPO.DELTOStoreID = 0 Then
            oPO.SetDELTOStoreID oPC.Configuration.Stores(1).ID
        End If
        cbDelTo.Caption = oPC.Configuration.Stores.FindStoreByID(oPO.DELTOStoreID).Description
        LoadListView
        LoadDeals
        cmdSave.Enabled = False
        cmdIssue.Enabled = False
        cmdCancel.Caption = "&Close"
        mnuFileCancel.Caption = "&Close"
        cmdNewRows.Enabled = True
        lvwLines.Enabled = True
        lvwLines.Height = 4850
        SetEditFrameEnabled False, enNotEditing
        vMode = enNotEditing
    End If
    oPO.SetDELTOStoreID oPC.Configuration.defaultStoreID
    Me.cbDelTo.Caption = oPC.Configuration.Stores.FindStoreByID(oPC.Configuration.defaultStoreID).Description
    oPO.GetStatus
    If oPO.ISForeignCurrency Then
        lblCurrName.Caption = "(" & oPO.Supplier.DefaultCurrency.Symbol & ")"
    Else
        lblCurrName.Caption = ""
    End If
'    If oPO.ISForeignCurrency Then
'        Me.txtRunningTotal = strTotalForeign
'    Else
'        Me.txtRunningTotal = strtotal
'    End If
'    txtCurrencyRates = oPO.CurrencyConversionAsText & "     Value is : " & oPO.TotalPayableF
    flgLoading = False
End Sub

Private Sub cbDelTo_Click()
Static i As Long
    i = OptionLoop(GetMax(i, 1), oPC.Configuration.Stores.Count)
    oPO.SetDELTOStoreID oPC.Configuration.Stores(i).ID
    Me.cbDelTo.Caption = oPC.Configuration.Stores(i).Description
End Sub

Private Sub cboCAT_SelectionChanged()
If flgLoading Then Exit Sub
    oPOLine.CategoryID = tlCategories.Key(cboCat.Items.CellCaption(cboCat.Items.SelectedItem, 0))
End Sub

Private Sub cboDeal_SelectionChanged()
    oPOLine.SetDiscount cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 0)
    oPOLine.DEALID = CLng(cboDeal.Items.CellCaption(cboDeal.Items.SelectedItem, 2))
End Sub

Private Sub cmdSupplier_Click()
Dim frm As frmSupplierPreview
    If oPO.Supplier.Name = "" Then Exit Sub
    Set frm = New frmSupplierPreview
    frm.Component oPO.Supplier
    frm.Show
End Sub

Private Sub cboRef_SelectionChanged()
    oPOLine.Ref = cboRef.Items.CellCaption(cboRef.Items.SelectedItem, 0)
    oPOLine.COLID = cboRef.Items.CellCaption(cboRef.Items.SelectedItem, 3)
    txtRef = oPOLine.Ref
End Sub

Private Sub cbTP_Click()
Dim frm As New frmSupplierPreview
    If oPO.Supplier.ID > 0 Then
        frm.Component oPO.Supplier
        frm.Show
    End If
End Sub
Private Sub Form_Initialize()
    Set vCanAdd = New z_BrokenRules
    Set tlCategories = New z_TextList
    tlCategories.Load ltCategory
End Sub
Private Sub Form_Terminate()
    Set tlCategories = Nothing
    Set oTP = Nothing
    Set oCurrentCopy = Nothing
    Set oPO = Nothing
    Set tlSupplier = Nothing
    Set oPOLine = Nothing
End Sub
Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    Cancel = True
End Sub
Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
'
'Private Sub mnuAddresses_Click()
'Dim frm As frmInvAddr
'    Set frm = New frmInvAddr
'    frm.Component oPO
'    frm.Show vbModal
'End Sub

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
    txtError = pMsg
End Sub
Private Sub oPOLine_Valid(Msg As String)
    cmdEnter.Enabled = (Msg = "")
    txtError = Msg
End Sub

'Sub oPOLine_ExtensionChange(lngExtension As Long, strExtension As String)
'    flgLoading = True
'   ' Me.txtTotal = strExtension
'    flgLoading = False
'    lngCurrentExtension = lngExtension
'End Sub


Private Sub oPO_TotalChange(strtotal As String, strTotalForeign As String)
    flgLoading = True
    If oPO.capturecurrency Is oPC.Configuration.DefaultCurrency Then
        Me.txtRunningTotal = strtotal
    Else
        Me.txtRunningTotal = strTotalForeign
        txtCurrencyRates = oPO.CurrencyConversionAsText & "     Value is : " & oPO.TotalPayableF
    End If
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

Private Sub txtRef_GotFocus()
'MsgBox "Gotfocus"
End Sub

Sub vCanAdd_NobrokenRules()
    Me.cmdNewRows.Enabled = True
    Me.cmdNewRows.SetFocus
End Sub

Private Sub txtRef_Validate(Cancel As Boolean)
    If oPOLine Is Nothing Then Exit Sub
    oPOLine.Ref = txtRef
End Sub

Private Sub Form_Load()
    left = 10
    top = 10
    Width = 11100
    Height = 6700
    SetLvw
    lvwLines.Height = 4850
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If oPO.IsEditing Then oPO.CancelEdit
End Sub
Private Sub LoadNewSupplier(plngTPID As Long)
    If oPO.SetTP(plngTPID) Then
        With oPO.Supplier
            Me.txtPhone = .OrderToAddress.Phone
            Me.txtSuppname = .NameAndCode(15)
            Me.txtFax = .OrderToAddress.Fax
        End With
        vCanAdd.RuleBroken "TP", False
        LoadDeals
    End If
End Sub
Public Function SetSupplier(pTPID As Long) As Boolean
Dim bSuccess As Boolean
    bSuccess = oPO.Supplier.Load(pTPID)
    SetSupplier = bSuccess
    If bSuccess Then
        vCanIssue.RuleBroken "TP", False
'        oPO.SetBillToAddress oPO.Supplier.BillTOAddress
'        oPO.setDelToAddress oPO.Supplier.DelToAddress
    End If
End Function

Private Sub SetEditFrameEnabled(pYesNo As Boolean, eMode As EnumMode)
Dim lngColour As Long
    'A is adding, E is editing
    bFrameEnabled = pYesNo   'shared for use in all the form
    If (eMode = enAddingRow Or eMode = enNotEditing) And pYesNo Then
        Me.txtCode.Enabled = True
    Else
        Me.txtCode.Enabled = False
    End If
    txtNote.Enabled = pYesNo
    Me.txtCurrencyRates.Enabled = pYesNo
    txtPrice.Enabled = pYesNo
    txtTitle.Enabled = pYesNo
    txtQtyFirm.Enabled = pYesNo
    txtQtySS.Enabled = pYesNo
    cboCat.Enabled = pYesNo
    cboDeal.Enabled = pYesNo
    cboRef.Enabled = pYesNo
    cmdEnter.Enabled = pYesNo
    cmdCancel.Enabled = Not pYesNo
    cmdIssue.Enabled = (Not pYesNo) And bValidCN
    cmdSave.Enabled = (Not pYesNo) And bValidCN And oPO.IsDirty
'    cmdNewRows.Enabled = pYesNo
    If pYesNo Then
        lngColour = &HFFFFFF
    Else
        lngColour = 14416635
    End If
    
    Me.txtCode.BackColor = lngColour
    Me.txtPrice.BackColor = lngColour
End Sub
Private Sub SetControlsForNew()
    mnuFileCancel.Caption = "&Cancel"
    txtStatus = "IN PROCESS"
End Sub

Private Sub cmdEnter_Click()
Dim currDeposit As Currency
Dim blnResult As Boolean
Dim strCurrFormat As String
Dim curTotalDeposit As Currency
Dim strETACode As String
    If txtCode = "" Then
        MsgBox "Enter a code", vbOKOnly + vbInformation, "Papyrus Ordering Information"
        Exit Sub
    End If
    oPOLine.ApplyEdit
    oPOLine.beginedit

    If vMode = enAddingRow Then
        strETACode = oPOLine.ETACode
        lvwLines.ListItems.Add 1, oPOLine.Key
        LoadListViewLine oPOLine.Key, Me.lvwLines.ListItems(1)
        Set oPOLine = oPO.POLines.Add
        oPOLine.SetQtyFirm 1
        oPOLine.SetQtySS 0
        oPOLine.SetETA strETACode
  '      txtETA = strETACode
        oPOLine.TRID = oPO.TRID
        txtCode.SetFocus
    ElseIf vMode = enEditingRow Then
        LoadListViewLine lngSelectedRowIndex, Me.lvwLines.ListItems(lngSelectedRowIndex)
        cmdNewRows_Click
    End If
    oPO.GetStatus
    ClearLineControls
    cmdNewRows.SetFocus
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
        Me.lvwLines.Enabled = True
        lvwLines.Height = 4850
        ClearLineControls
    ElseIf vMode = enAddingRow Then    'we are stopping adding rows
        cmdNewRows.Caption = "&Add"
        SetEditFrameEnabled False, vMode
        vMode = enEditingRow
        Me.lvwLines.Enabled = True
        lvwLines.Height = 4850
        ClearLineControls
    ElseIf vMode = enNotEditing Then  'we are starting to add rows
        cmdNewRows.Caption = "&Stop"
        SetEditFrameEnabled True, vMode
        vMode = enAddingRow
        Me.lvwLines.Enabled = False
        lvwLines.Height = 2600
        ClearLineControls
        Me.txtCode.SetFocus
        Set oPOLine = oPO.POLines.Add
        oPOLine.TRID = oPO.TRID
        oPOLine.SetETA DateAdd("d", oPO.Supplier.DefaultETA, Date)
        Me.txtETA = oPOLine.ETAF
    End If

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
        lstItem.SubItems(4) = .Ref
        lstItem.SubItems(1) = .TitleAuthor
        lstItem.SubItems(2) = .QtyFirm
        lstItem.SubItems(3) = .Qtyseesafe
        lstItem.SubItems(6) = .DiscountF  ' Format(.Discount, "##0.0%")
        If oPC.Configuration.DefaultCurrency Is oPO.capturecurrency Then
            lstItem.SubItems(5) = .PriceF
            lstItem.SubItems(7) = .ExtensionInclDepositF
        Else
            lstItem.SubItems(5) = .PriceF_Foreign
            lstItem.SubItems(7) = .ExtensionInclDepositF_Foreign
        End If
    End With
End Sub
Private Sub lvwLines_DblClick()
'This must load the editing line with the current line's data
    If lvwLines.ListItems.Count = 0 Then Exit Sub
    lngEditingIdx = lvwLines.SelectedItem.Key
    Set oPOLine = oPO.POLines(lngEditingIdx)
    lngSelectedRowIndex = lvwLines.SelectedItem.Key
    txtCode = oPOLine.code
    txtTitle = oPOLine.Title
    txtPrice = oPOLine.Price
    txtQtyFirm = oPOLine.QtyFirm
    txtQtySS = oPOLine.Qtyseesafe
    txtNote = oPOLine.Note
    txtRef = oPOLine.Ref
    txtTotal = oPOLine.ExtensionSimpleF
    Me.txtETA = oPOLine.ETAF
  '''''''  oPOLine.beginedit
    oPOLine.LoadColsPerPID
    cboDeal.Items.SelectItem(cboDeal.Items.FindItem(oPOLine.DEALID, 2)) = True
    LoadcboRef
    On Error Resume Next
    cboRef.Items.SelectItem(cboRef.Items.FindItem(oPOLine.Ref, 0)) = True
    On Error GoTo 0
'    txtRef = oPOLine.Ref
    tlCategories.Item (oPOLine.CategoryID)
    On Error Resume Next
    cboCat.Items.SelectItem(cboCat.Items.FindItem(tlCategories.Item(oPOLine.CategoryID), 0)) = True
    On Error GoTo 0
'    txtCode = oPOLine.ProductCode
    AutoSelect txtPrice
    
    SetEditFrameEnabled True, enEditingRow
    vMode = enEditingRow
    txtPrice.SetFocus
    lvwLines.Height = 2600
    cmdNewRows.Caption = "&Stop edit"
    oPOLine.GetStatus
End Sub

Private Sub cboTP_Validate(Cancel As Boolean)
    If oPO.Supplier Is Nothing Then
        MsgBox "Please enter a Supplier before continuing", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
    End If
End Sub

Private Sub txtNote_Change()
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    oPOLine.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oPOLine.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    Cancel = Not oPOLine.setnote(txtNote)
End Sub
Private Sub txtNote_LostFocus()
    If flgLoading Then Exit Sub
    txtNote = oPOLine.Note
End Sub

Private Sub mnuEditNote_Click()
Dim ofrm As New frmNote
    ofrm.Component oPO
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
    oPO.SetStatus stVOID
    txtStatus = "Void"
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
Dim pQty As Integer
Dim pApproID As Long
Dim bOK  As Boolean

On Error GoTo ERR_Handler
    
    If txtCode = "" Or vMode = enEditingRow Then Exit Sub
    bOK = oPOLine.SetLineProduct("", txtCode, oPO.Supplier.DefaultETA)
    If bOK Then
        txtTitle = oPOLine.Title
        txtPrice = oPOLine.Price
        txtQtyFirm = oPOLine.QtyFirmF
        txtQtySS = oPOLine.QtySeesafeF
        txtPrice.SetFocus
        txtCode = oPOLine.ProductCode
        txtETA = oPOLine.ETAF
        Me.txtTotal = oPOLine.ExtensionSimpleF
    Else
        MsgBox "Cannot find book on database or or bookfind", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        Cancel = True
        GoTo EXIT_Handler
    End If
    If cboDeal.Items.ItemCount > 0 Then
        cboDeal.Items.SelectItem(cboDeal.Items(0)) = True
    End If
    LoadcboRef
    If cboRef.Items.ItemCount > 0 Then
        cboRef.Items.SelectItem(cboRef.Items(0)) = True
        oPOLine.Ref = cboRef.Items.CellCaption(cboRef.Items.SelectedItem, 0)
        txtRef = oPOLine.Ref
    End If
    oPOLine.GetStatus
    txtPrice.SetFocus

EXIT_Handler:
    Set oProd = Nothing
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
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
        Me.txtSuppname = .Supplier.NameAndCode(20)
        Me.txtPhone = .BillTOAddress.Phone
        Me.txtFax = .BillTOAddress.Fax
    End With
End Sub


Private Sub SavePO()
On Error GoTo ERR_Handler
    
    oPO.post
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
   ' Resume
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
'Dim ViewOrPrint As PreviewPrint
Dim strResult As String
Dim frm As frmPOPreview

    If oPO.status = stInProcess Then
        If MsgBox("Issue this order?.  Confirm.", vbYesNo + vbQuestion, "Papyrus Invoicing Status") = vbNo Then
            Exit Sub
        End If
    End If
    oPO.SetStatus stISSUED
    
    strResult = oPO.post
    If strResult = "ERROR" Then
        MsgBox "This action has failed. Contact support"
        Exit Sub
    End If
    Set frm = New frmPOPreview
    frm.Component oPO.TRID
    frm.Show
    Unload Me
End Sub
Private Sub cmdSave_Click()
    oPO.SetStatus stInProcess
    SavePO
    oPO.beginedit
    cmdCancel.Caption = "&Close"
    cmdSave.Enabled = False
    cmdNewRows.SetFocus
End Sub

Private Sub cmdCancel_Click()
    oPO.CancelEdit
    Unload Me
End Sub


Private Sub ClearLineControls()
    flgLoading = True
    txtCode = ""
    txtPrice = ""
    txtTitle = ""
    txtRef = ""
    txtTotal = ""
    txtNote = ""
    txtQtyFirm = ""
    txtQtySS = ""
    cboRef.Items.RemoveAllItems
    cboCat.Items.SelectItem(cboCat.Items(0)) = True
    flgLoading = False
End Sub


Private Sub txtPrice_GotFocus()
    AutoSelect txtPrice
End Sub
Private Sub txtPrice_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetPrice(txtPrice) Then
        Cancel = True
    End If
    If oPO.ISForeignCurrency Then
        txtTotal = oPOLine.ExtensionSimpleF_Foreign
    Else
        txtTotal = oPOLine.ExtensionSimpleF
    End If
End Sub
Private Sub txtQtyFirm_GotFocus()
    AutoSelect txtQtyFirm
End Sub
Private Sub txtQtyFirm_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetQtyFirm(txtQtyFirm) Then
        Cancel = True
    End If
    If oPO.ISForeignCurrency Then
        txtTotal = oPOLine.ExtensionSimpleF_Foreign
    Else
        txtTotal = oPOLine.ExtensionSimpleF
    End If
End Sub

Private Sub txtQtySS_GotFocus()
    AutoSelect txtQtySS
End Sub
Private Sub txtQtySs_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetQtySS(txtQtySS) Then
        Cancel = True
    End If
    If oPO.ISForeignCurrency Then
        txtTotal = oPOLine.ExtensionSimpleF_Foreign
    Else
        txtTotal = oPOLine.ExtensionSimpleF
    End If
End Sub

Private Sub SetIssueButtonCaption()
        If oPO.statusF = "IN PROCESS" Then
            cmdIssue.Caption = "Issue"
        ElseIf oPO.IsDirty Then
            cmdIssue.Caption = "Save"
        Else
            cmdIssue.Caption = "Print"
        End If
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
    cboDeal.WidthList = 190
    cboDeal.HeightList = 162
    cboDeal.AllowSizeGrip = True
    cboDeal.AutoDropDown = True
    cboDeal.Columns.Add "Discount"
    cboDeal.Columns.Add "Description"
    cboDeal.Columns.Add ""
    cboDeal.Columns(0).Width = 45
    cboDeal.Columns(1).Width = 110
    cboDeal.Columns(2).Width = 0
    cboDeal.BackColorLock = Me.BackColor
    cboDeal.EndUpdate
    
    cboRef.BeginUpdate
    cboRef.WidthList = 190
    cboRef.HeightList = 162
    cboRef.AllowSizeGrip = True
    cboRef.AutoDropDown = True
    cboRef.Columns.Add "Ref"
    cboRef.Columns.Add "Order"
    cboRef.Columns.Add "Qty"
    cboRef.Columns.Add "COLID"
    cboRef.Columns(0).Width = 70
    cboRef.Columns(1).Width = 70
    cboRef.Columns(2).Width = 40
    cboRef.Columns(3).Width = 0
    cboRef.BackColorLock = Me.BackColor
    cboRef.EndUpdate

    cboCat.BeginUpdate
    cboCat.WidthList = 190
    cboCat.HeightList = 162
    cboCat.AllowSizeGrip = True
    cboCat.AutoDropDown = True
    cboCat.Columns.Add "Category"
    cboCat.Columns(0).Width = 190
    cboCat.BackColorLock = Me.BackColor
    cboCat.EndUpdate
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
        ar(1, i) = oDL.Description
        ar(0, i) = oDL.DiscountF
        ar(2, i) = oDL.ID
        i = i + 1
    Next
    cboDeal.PutItems ar
    cboDeal.EndUpdate
End Sub
Private Sub LoadcboRef()
Dim i As Integer
Dim oD As d_COLine
Dim ar()
    i = 0
    If oPOLine.COLsPerPID.Count < 1 Then
        Exit Sub
    End If
    For Each oD In oPOLine.COLsPerPID
        i = i + 1
    Next
    ReDim ar(3, i - 1)
    i = 0
    cboRef.BeginUpdate
    cboRef.Items.RemoveAllItems
    For Each oD In oPOLine.COLsPerPID
        ar(1, i) = oD.DocCode
        ar(0, i) = oD.Ref
        ar(2, i) = oD.Qty
        ar(3, i) = oD.COLID
        i = i + 1
    Next
    cboRef.PutItems ar
    cboRef.EndUpdate
End Sub

Private Sub txtETA_GotFocus()
    AutoSelect Controls("txtETA")
End Sub

Private Sub txtETA_LostFocus()
    txtETA = oPOLine.ETAF
End Sub

Private Sub txtETA_Validate(Cancel As Boolean)
    If flgLoading Then Exit Sub
    If Not oPOLine.SetETA(txtETA) Then
        Cancel = True
    End If
End Sub

