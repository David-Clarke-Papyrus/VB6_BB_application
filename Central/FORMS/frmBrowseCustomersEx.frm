VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBrowseCustomersEx 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse customers"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15000
   Icon            =   "frmBrowseCustomersEx.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7185
      Left            =   10710
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Text            =   "frmBrowseCustomersEx.frx":038A
      Top             =   105
      Width           =   4095
   End
   Begin VB.CommandButton cmdExportToFile 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export to file"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8025
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   7170
      UseMaskColor    =   -1  'True
      Width           =   1470
   End
   Begin VB.CommandButton cmdSendToBranches 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Send to branches"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1950
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   7725
      UseMaskColor    =   -1  'True
      Width           =   2520
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   615
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImportList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Import customerlist from CSV file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1965
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   8100
      UseMaskColor    =   -1  'True
      Width           =   2520
   End
   Begin VB.CommandButton cmdSetlabels 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Set labels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7995
      Width           =   825
   End
   Begin VB.CommandButton cmdLabels 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print labels"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8025
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   6675
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton cmdSimpleSearch 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Simple search"
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
      Left            =   6660
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers mwith an address containing . . ."
      Top             =   3900
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton cmdFix 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Fix Addresses"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   12390
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers matching the retrictions entered."
      Top             =   7620
      UseMaskColor    =   -1  'True
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Rule set"
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
      Height          =   3750
      Left            =   105
      TabIndex        =   10
      Top             =   105
      Width           =   8100
      Begin VB.Frame Frame3 
         BackColor       =   &H00D3D3CB&
         Height          =   645
         Left            =   105
         TabIndex        =   24
         Top             =   225
         Width           =   7770
         Begin VB.CommandButton cmdClearRuleSet 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Clear rule set"
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
            Left            =   6105
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers mwith an address containing . . ."
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   1515
         End
         Begin VB.CommandButton cmdLoadRuleSet 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Load rule set"
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
            Left            =   3165
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers mwith an address containing . . ."
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.ComboBox cboRuleSets 
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
            ForeColor       =   &H8000000D&
            Height          =   345
            ItemData        =   "frmBrowseCustomersEx.frx":0391
            Left            =   75
            List            =   "frmBrowseCustomersEx.frx":0393
            Style           =   2  'Dropdown List
            TabIndex        =   26
            ToolTipText     =   "Select a customer grouping"
            Top             =   195
            Width           =   2955
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Delete rule set"
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
            Left            =   4515
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers mwith an address containing . . ."
            Top             =   195
            UseMaskColor    =   -1  'True
            Width           =   1515
         End
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1275
         Left            =   105
         TabIndex        =   23
         Top             =   1380
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   2249
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Criterion"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operator"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Argument"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "RuleID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdRemoveRule 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Remove rule"
         Enabled         =   0   'False
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
         Left            =   6390
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers mwith an address containing . . ."
         Top             =   2310
         UseMaskColor    =   -1  'True
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D3D3CB&
         Height          =   915
         Left            =   135
         TabIndex        =   17
         Top             =   2655
         Width           =   4680
         Begin VB.CommandButton cmdSaveRule 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Save rule set"
            Enabled         =   0   'False
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
            Left            =   3105
            MaskColor       =   &H00E0E0E0&
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "Click to find all customers mwith an address containing . . ."
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.TextBox txtRulesetName 
            Height          =   330
            Left            =   135
            TabIndex        =   18
            Top             =   480
            Width           =   2910
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rule set name (min. len 5 chars)"
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
            Height          =   285
            Left            =   135
            TabIndex        =   20
            Top             =   195
            Width           =   2895
         End
      End
      Begin VB.CommandButton cmdFind1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Add rule set result to customer list"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   5400
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers matching the retrictions entered."
         Top             =   2820
         UseMaskColor    =   -1  'True
         Width           =   2595
      End
      Begin VB.ComboBox cboCriterion 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         ItemData        =   "frmBrowseCustomersEx.frx":0395
         Left            =   120
         List            =   "frmBrowseCustomersEx.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Select a customer grouping"
         Top             =   990
         Width           =   2955
      End
      Begin VB.ComboBox cboOperator 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   3060
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Select a customer grouping"
         Top             =   990
         Width           =   1440
      End
      Begin VB.TextBox txtArgument 
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
         Left            =   4500
         MultiLine       =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number. Hit ENTER to fetch."
         Top             =   975
         Width           =   1710
      End
      Begin VB.ComboBox cboArgument 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   4485
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Select a customer grouping"
         Top             =   990
         Width           =   1725
      End
      Begin VB.CommandButton cmdAddRule 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Add rule"
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
         Left            =   6390
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Click to find all customers mwith an address containing . . ."
         Top             =   1005
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Start new customer list"
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
      Left            =   1935
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Click to find all customers mwith an address containing . . ."
      Top             =   3930
      UseMaskColor    =   -1  'True
      Width           =   2355
   End
   Begin VB.CommandButton cmdOutlook_Export 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export to Outlook"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4515
      Picture         =   "frmBrowseCustomersEx.frx":0399
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7725
      Width           =   1860
   End
   Begin VB.CommandButton cmdDeselectAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Deselect all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8145
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdSelectAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select all"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7830
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdAddSelected 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add selected from ruleset to current list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   10950
      Visible         =   0   'False
      Width           =   3885
   End
   Begin VB.CommandButton cmdManage 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&View current list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   10650
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmdLists 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select current list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10350
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   8505
      Picture         =   "frmBrowseCustomersEx.frx":0723
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid CustGrid 
      Height          =   3390
      Left            =   105
      OleObjectBlob   =   "frmBrowseCustomersEx.frx":0AAD
      TabIndex        =   0
      Top             =   4290
      Width           =   7860
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   6495
      TabIndex        =   29
      Top             =   7740
      Width           =   1650
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   150
      TabIndex        =   22
      Top             =   3960
      Width           =   1725
   End
   Begin VB.Label lblDefaultListName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   345
      Left            =   4305
      TabIndex        =   3
      Top             =   3930
      Width           =   2325
   End
End
Attribute VB_Name = "frmBrowseCustomersEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oS As z_Search
Dim c As Collection
Dim o As Collection
Dim mLeft As Long
Dim mRowHeight As Long
Dim mColumnSpacing As Long
Dim mDescription As String
Dim mTopMargin As Long
Dim mPrintWidth As Long

Dim XR As XArrayDB
Const cbase = 1

Dim cCust As c_C_Customer
Dim dispCust As d_C_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim oCust As a_Customer
Dim blnNoRecordsReturned As Boolean
Dim XA As New XArrayDB
'#If H_CENTRAL <> 1 Then
Dim ofrm As frmCustomerPreview
'#End If
Dim ofrmLoy As frmLoyaltyPreview
Dim CustomerTypes_tl As z_TextList
Dim InterestGroups_tl As z_TextList
Dim Stores_tl As z_TextList

Dim RuleSet_tl As z_TextList
Dim arDir(1 To 6) As Integer

Private Sub cboCriterion_Click()
    ReloadCriteriaSpecs
End Sub

Private Sub cmdAddRule_Click()
Dim strArg As String

    If cboArgument.Visible = True Then
        strArg = cboArgument
    Else
        strArg = Trim(txtArgument)
    End If
    oS.AddRule cboCriterion, cboOperator, strArg
    ReloadRuleSetList
End Sub
Private Sub ReloadRuleSetList()
Dim oRule As rule

Dim i As Long
Dim lstItem As ListItem
    i = 0
    lvw.ListItems.Clear
    For Each oRule In oS.Rules
        'strRule = strRule & IIf(Len(strRule) > 0, vbCrLf, "") & oRule.Criterion & " " & oRule.Operator & " " & oRule.Argument
        Set lstItem = lvw.ListItems.Add
        With lstItem
            .Text = oRule.Criterion
            .SubItems(1) = oRule.Operator
            .SubItems(2) = oRule.Argument
            .SubItems(3) = oRule.ID
        End With
    Next
    cmdRemoveRule.Enabled = (lvw.ListItems.Count > 0)
End Sub

Private Sub cmdAddSelected_Click()
    On Error GoTo errHandler
Dim i As Long
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If lngDefaultListID = 0 Then
        MsgBox "You must select a customer list first.", , "Can't do this"
    Else
        For i = 1 To XA.UpperBound(1)
            If XA(i, 1) = True Then
                strSQL = "INSERT INTO tLISTITEM (LISTITEM_LIST_ID,LISTITEM_TP_ID) VALUES (" & lngDefaultListID & "," & CLng(XA(i, 7)) & ")"
                oPC.COShort.Execute strSQL

            End If
        Next i
    End If

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217873 Then Resume Next
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdAddSelected_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClearRuleSet_Click()
    oS.ClearRules
    ReloadRuleSetList
    Me.txtRulesetName = ""
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Confirm delete ruleset: " & cboRuleSets & "?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    oS.DeleteRuleSet Val(RuleSet_tl.Key(cboRuleSets))
    MsgBox "Ruleset: " & cboRuleSets & "deleted.", vbInformation, "Status"
    ReloadRulesetCombo
End Sub



Private Sub cmdExportToFile_Click()
Dim oSQL As New z_SQL
Dim strFilename As String

    Screen.MousePointer = vbHourglass
    strFilename = "\PBKS\ExportFiles\" & Format(Now, "YYYYMMDD-HHNN") & ".txt"
    oSQL.ExportCustomerList strFilename
    Screen.MousePointer = vbDefault
    MsgBox "The exported file is saved as " & strFilename
End Sub

Private Sub cmdFind1_Click()
    If lvw.ListItems.Count < 1 Then
        MsgBox "Please create at least one rule first.", vbInformation, "Can't do this"
        Exit Sub
    End If
    txtSQL = oS.BuildList()
    Find
    lblResults.Caption = CStr(XA.UpperBound(1)) & " rows"
End Sub



Private Sub cmdFix_Click()
    Screen.MousePointer = vbHourglass
    oS.FixAddresses
    Screen.MousePointer = vbDefault
    MsgBox "Done"
End Sub

Private Sub cmdImportList_Click()
Dim strFilename As String
Dim oImp As New z_Import


On Error GoTo 0
    cd1.InitDir = oPC.SharedFolderRoot
    cd1.DefaultExt = ".csv"
    cd1.DialogTitle = "locate List of account numberd in .csv format"
    On Error GoTo ERR_CANCELLED
    cd1.ShowOpen
    strFilename = cd1.FileName
    oImp.ImportCSVAcnoList strFilename
    On Error GoTo 0
    Find
    lblResults.Caption = CStr(XA.UpperBound(1)) & " rows"
    

ERR_CANCELLED:
    Exit Sub
    
End Sub

Private Sub cmdLabels_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL
Dim res As Long

    mDescription = GetSetting("CENTRAL", "LABELS", "LABELNAME", "")

    res = oSQL.RunGetRecordset("SELECT * FROM tMAILLABEL WHERE ML_DESCRIPTION = '" & mDescription & "'", enText, Array(), "", rs)
    If Not rs.EOF Then
        mLeft = FNN(rs.Fields(2))
        mRowHeight = FNN(rs.Fields(3))
        mColumnSpacing = FNN(rs.Fields(4))
        mTopMargin = FNN(rs.Fields(5))
        mPrintWidth = FNN(rs.Fields(6))
    Else
        MsgBox "There is no default mail label, click 'Set labels' to design one"
        
    End If
    
    
    cCust.PrintLabels mLeft, mRowHeight, mColumnSpacing, mTopMargin, mPrintWidth

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomersEx.cmdLabels_Click"
End Sub

Private Sub cmdLists_Click()
    On Error GoTo errHandler
Dim frm As New frmLists
    frm.Show vbModal
    If lngDefaultListID > 0 Then
        lblDefaultListName.Caption = strDefaultListName
    Else
        lblDefaultListName.Caption = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdLists_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdLoadRuleSet_Click()
    oS.LoadRuleSet cboRuleSets
    ReloadRuleSetList
    txtRulesetName = cboRuleSets
End Sub

Private Sub cmdManage_Click()
    On Error GoTo errHandler
Dim frm As New frmListsManage
    frm.Show vbModal
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdManage_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub cmdRemoveRule_Click()
    If MsgBox("Remove ruleset: " & lvw.SelectedItem.Text & " " & lvw.SelectedItem.SubItems(1) & " " & lvw.SelectedItem.SubItems(2) & "?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    oS.removeRule lvw.SelectedItem.Index
    ReloadRuleSetList
End Sub

'Private Sub cmdSaveLabelsettings_Click()
'    If IsNumeric(txtLeft) And IsNumeric(txtRowHeight) And IsNumeric(txtColumnSpacing) Then
'        SaveSetting "CENTRAL", "Labels", "Left", txtLeft
'        SaveSetting "CENTRAL", "Labels", "RowHeight", txtRowHeight
'        SaveSetting "CENTRAL", "Labels", "ColumnSpacing", txtColumnSpacing
'
'    Else
'        MsgBox "Invalid value among label settings (not numeric)"
'    End If
'End Sub
'Private Sub cmdLoadLabelSettings_Click()
'    mLeft = GetSetting("CENTRAL", "Labels", "Left", 71)
'    mRowHeight = GetSetting("CENTRAL", "Labels", "RowHeight", 71)
'    mColumnSpacing = GetSetting("CENTRAL", "Labels", "ColumnSpacing", 71)
'    txtLeft = CStr(mLeft)
'    txtRowHeight = CStr(mRowHeight)
'    txtColumnSpacing = CStr(mColumnSpacing)
'End Sub
Private Sub cmdSaveRule_Click()
    oS.SaveRuleSet txtRulesetName
    ReloadRulesetCombo
    MsgBox "Rules set: " & txtRulesetName & "has been saved", vbInformation, "Status"
End Sub

Private Sub cmdSelectAll_Click()
Dim i As Integer
    For i = 1 To XA.UpperBound(1)
        XA(i, 1) = True
    Next
    Me.CustGrid.Refresh
End Sub
Private Sub cmdDeselectAll_Click()
Dim i As Integer
    For i = 1 To XA.UpperBound(1)
        XA(i, 1) = False
    Next
    Me.CustGrid.Refresh
End Sub


Private Sub cmdOutlook_Export_Click()
Dim frm As New frmOutlookExport
    frm.Component cCust
    frm.Show 'vbModal

End Sub

Private Sub cmdSendToBranches_Click()
Dim f As New frmStoreSelection
    f.Component "Customer records to be updated on . . .", "FromCustBrowse"
    f.Show vbModal
    Unload f
End Sub

Private Sub cmdSetlabels_Click()
Dim f As New frmLabelDesign

    
    f.Show vbModal
    mLeft = f.LabelLeft
    mRowHeight = f.LabelRowHeight
    mColumnSpacing = f.LabelColumnSpacing
    Unload f
    
    
'    Label_1.Visible = Not Me.Label_1.Visible
'    Label_2.Visible = Not Me.Label_2.Visible
'    Label_3.Visible = Not Me.Label_3.Visible
'    txtLeft.Visible = Not txtLeft.Visible
'    txtRowHeight.Visible = Not txtRowHeight.Visible
'    txtColumnSpacing.Visible = Not txtColumnSpacing.Visible
'    txtsaveas.Visible = Not txtsaveas.Visible
'    cmdSaveLabelsettings.Visible = Not cmdSaveLabelsettings.Visible
'    cmdLoadLabelSettings.Visible = Not cmdLoadLabelSettings.Visible
End Sub

Private Sub cmdStart_Click()
    oS.StartSearch
'    ReloadRuleSetList
    XA.Clear
    XA.ReDim 1, 0, 1, 7
    CustGrid.ReBind
    CustGrid.Refresh
    Me.lblResults.Caption = ""
End Sub



Private Sub Command1_Click()
    If lvw.ListItems.Count < 1 Then
        MsgBox "Please create at least one rule first.", vbInformation, "Can't do this"
        Exit Sub
    End If
    oS.BuildList
    Find
    lblResults.Caption = CStr(XA.UpperBound(1)) & " rows"

End Sub

Private Sub CustGrid_LostFocus()
    CustGrid.Update
End Sub


Private Sub CustGrid_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = 13 Then
        CustGrid_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CustGrid_DblClick()
    On Error GoTo errHandler
Dim lngID As Long
Dim blnEdit As Boolean
    If IsNull(CustGrid.Bookmark) Then Exit Sub
    lngID = Val(XA(CustGrid.Bookmark, 7))
    Set oCust = Nothing
    Set oCust = New a_Customer
    oCust.Load lngID
        Set ofrm = New frmCustomerPreview
        ofrm.Component oCust    ', False
        ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Find()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    Set cCust = Nothing
    Set cCust = New c_C_Customer
    cCust.LoadEx
    LoadArray
    CustGrid.ReBind
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Find"
End Sub

Private Sub ReloadCriteriaSpecs()
    Set o = oS.GetOperatorList(cboCriterion)
    LoadComboboxColl cboOperator, o
    Select Case cboCriterion
    Case "Customer type"
        txtArgument.Visible = False
        cboArgument.Visible = True
        LoadCombo cboArgument, CustomerTypes_tl
       ' cboArgument = CustomerTypes_tl.Item("0")
    Case "Customer group"
        txtArgument.Visible = False
        cboArgument.Visible = True
        LoadCombo cboArgument, InterestGroups_tl
       ' cboArgument = InterestGroups_tl.Item("0")
    Case "Town"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "Province"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "Post code"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "ISP"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "Name"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "Phone"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "ACno"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "Originating store"
        txtArgument.Visible = False
        cboArgument.Visible = True
        LoadCombo cboArgument, Stores_tl
    Case "Originating store set"
        txtArgument.Visible = True
        cboArgument.Visible = False
    Case "Email address"
        txtArgument.Visible = True
        cboArgument.Visible = False
    End Select
    
End Sub

Private Sub Form_Activate()
    SetMenu
End Sub

Private Sub Form_Click()
    txtArgument.Width = 1620
    txtArgument.Height = 375
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub

Private Sub Form_Load()
On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Top = 400
        Left = 200
        Width = 9840
        Height = 9800
    End If
    SetMenu
    Set CustomerTypes_tl = New z_TextList
    CustomerTypes_tl.Load ltCustomerTypeActive
    
    Set InterestGroups_tl = New z_TextList
    InterestGroups_tl.Load ltInterestGroupActive
    
    Set Stores_tl = New z_TextList
    Stores_tl.Load ltStores
    
    ReloadRulesetCombo
    
    Set oS = New z_Search
    oS.StartSearch
    Set c = oS.GetCriterionList
    LoadComboboxColl cboCriterion, c
    ReloadCriteriaSpecs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub ReloadRulesetCombo()
    Set RuleSet_tl = New z_TextList
    RuleSet_tl.Load ltRuleSet
    LoadCombo cboRuleSets, RuleSet_tl

End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    CustGrid.Width = Me.Width - (CustGrid.Left + 1900)
    cmdClose.Left = Me.Width - 1800
    cmdExportToFile.Left = Me.Width - 1800
    Me.cmdLabels.Left = Me.Width - 1800
    lngDiff = CustGrid.Height
    CustGrid.Height = Me.Height - (CustGrid.Top + 1500)
    lngDiff = CustGrid.Height - lngDiff
    cmdOutlook_Export.Top = cmdOutlook_Export.Top + lngDiff
    cmdClose.Top = cmdClose.Top + lngDiff
    cmdLabels.Top = cmdLabels.Top + lngDiff
    cmdExportToFile.Top = cmdExportToFile.Top + lngDiff
    cmdDeselectAll.Top = cmdDeselectAll.Top + lngDiff
    cmdSelectAll.Top = cmdSelectAll.Top + lngDiff
    lblResults.Top = lblResults.Top + lngDiff
    cmdSetlabels.Top = cmdSetlabels.Top + lngDiff
    cmdImportList.Top = cmdImportList.Top + lngDiff
    cmdSendToBranches.Top = cmdSendToBranches.Top + lngDiff
    
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oCust = Nothing
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_C_Customer
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 7
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 2) = objItem.Fullname2
            XA.Value(lngIndex, 3) = objItem.AcNo
            XA.Value(lngIndex, 4) = objItem.Cellf
            XA.Value(lngIndex, 5) = objItem.SalesQty
            XA.Value(lngIndex, 6) = objItem.SalesValueF
            XA.Value(lngIndex, 7) = objItem.ID
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    CustGrid.Array = XA
    Me.cmdOutlook_Export.Enabled = (cCust.Count > 0)
    Me.cmdLabels.Enabled = (cCust.Count > 0)
    Me.cmdExportToFile.Enabled = (cCust.Count > 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.LoadArray"
End Sub
Public Sub AddToList()
    On Error GoTo errHandler
Dim i As Long
Dim strSQL As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If lngDefaultListID = 0 Then
        MsgBox "You must select a customer list first.", , "Can't do this"
    Else
        For i = 1 To CustGrid.SelBookmarks.Count
            strSQL = "INSERT INTO tLISTITEM (LISTITEM_LIST_ID,LISTITEM_TP_ID) VALUES (" & lngDefaultListID & "," & CLng(XA(CustGrid.SelBookmarks(i - 1), 7)) & ")"
            oPC.COShort.Execute strSQL
        Next i
    End If

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.AddToList"
End Sub

Public Sub RemoveFromList()
    On Error GoTo errHandler
    MsgBox "remove"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.RemoveFromList"
End Sub
Private Sub CustGrid_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

If ColIndex = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
'    If Direction = 0 Then
'        Direction = 1
'    Else
'        Direction = 0
'    End If
    If arDir(ColIndex + 1) = 1 Then
        arDir(ColIndex + 1) = 0
    Else
        arDir(ColIndex + 1) = 1
    End If
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, arDir(ColIndex + 1), GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    CustGrid.Refresh
    Screen.MousePointer = vbDefault

End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 3, 4
            GetRowType = XTYPE_STRING
        Case 5, 6
            GetRowType = XTYPE_CURRENCY
    End Select
End Function


Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub
Private Sub SetMenu()

    Forms(0).mnuSaveColumnWidths.Enabled = True
    
End Sub

Private Sub txtArgument_dblCLick()
    txtArgument.Width = GetMin(txtArgument.Width * 1.5, 3500)
    txtArgument.Height = GetMin(txtArgument.Height * 2, 2600)
    txtArgument.ZOrder 0
End Sub

Private Sub txtArgument_LostFocus()
    txtArgument.Width = 1620
    txtArgument.Height = 375
End Sub


'Private Sub txtLeft_Change()
'    If IsNumeric(txtLeft) Then mLeft = CLng(txtLeft)
'End Sub
'Private Sub txtLeft_Validate(Cancel As Boolean)
'    Cancel = Not IsNumeric(txtLeft)
'End Sub
'
'Private Sub txtRowHeight_Change()
'    If IsNumeric(txtRowHeight) Then mRowHeight = CLng(txtRowHeight)
'End Sub
'Private Sub txtRowHeight_Validate(Cancel As Boolean)
'    Cancel = Not IsNumeric(txtRowHeight)
'End Sub
'Private Sub txtColumnSpacing_Change()
'    If IsNumeric(txtColumnSpacing) Then mColumnSpacing = CLng(txtColumnSpacing)
'End Sub
'Private Sub txtColumnSpacing_Validate(Cancel As Boolean)
'    Cancel = Not IsNumeric(txtColumnSpacing)
'End Sub

Private Sub txtRulesetName_Change()
    cmdSaveRule.Enabled = (Len(txtRulesetName) > 5)
End Sub

Private Sub cmdSimpleSearch_Click()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim OpenResult As Integer
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    oPC.COShort.Execute "INSERT INTO tS SELECT DISTINCT TPID FROM vSearchCustomers WHERE TPID = " & CStr(lngTPID)
    Find
    lblResults.Caption = CStr(XA.UpperBound(1)) & " rows"
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomersEx.cmdSimpleSearch_Click"
End Sub

Public Sub mnuAlertHistory()
    On Error GoTo errHandler
Dim f As New frmAlertHistory
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim strCustname As String
Dim strCustAcno As String
Dim lngTPID As Long
    
    If CustGrid.SelBookmarks.Count < 1 Then
        MsgBox "Select a customer first.", vbOKOnly + vbExclamation, "Can't do this"
        Exit Sub
    End If
    If CustGrid.SelBookmarks.Count > 1 Then
        MsgBox "You can only read messages for one customer.", vbOKOnly + vbExclamation, "Can't do this"
        Exit Sub
    End If

    lngTPID = CLng(XA(CustGrid.SelBookmarks(0), 7))
    strCustname = CStr(XA(CustGrid.SelBookmarks(0), 2))
    strCustAcno = CStr(XA(CustGrid.SelBookmarks(0), 3))
    
    If lngTPID = 0 Then Exit Sub
    
    f.Component strCustAcno
    f.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.mnuAlertHistory"
End Sub

'Private Sub CustGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    On Error GoTo errHandler
'   If Button = 2 Then   ' Check if right mouse button
'                        ' was clicked.
'
'      PopupMenu Forms(0).mnuCustomerBrowseContext ' Display the File menu as a
'                        ' pop-up menu.
'   End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseCustomers.CustGrid_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
'End Sub

Private Sub CustGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
    '  PopupMenu Forms(0).mnuBrowseCustomerPopup   ' Display the File menu as a
      PopupMenu Forms(0).mnuCustomerBrowseContext   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseCustomers.CustGrid_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub


