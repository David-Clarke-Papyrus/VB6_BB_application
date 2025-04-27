VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B4B5B73C-172E-47B1-BFC2-C6F740957D01}#1.0#0"; "VB Control Manager.ocx"
Begin VB.Form frmDebtors 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Debtors"
   ClientHeight    =   11145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21840
   FillColor       =   &H00FCF2EB&
   Icon            =   "frmDebtors.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11145
   ScaleWidth      =   21840
   WindowState     =   2  'Maximized
   Begin VBControlManager.ControlManager CM 
      Height          =   10590
      Left            =   -150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -60
      Width           =   20325
      _ExtentX        =   35851
      _ExtentY        =   18680
      ActiveColor     =   12582912
      BackColor       =   9525832
      Size            =   4
      TitleBar_CloseVisible=   0   'False
      Begin VB.Frame frStatement 
         BackColor       =   &H00F7EDE8&
         Caption         =   "Statements"
         ForeColor       =   &H8000000D&
         Height          =   10410
         Left            =   14520
         TabIndex        =   26
         Top             =   180
         Width           =   5793
         Begin TabDlg.SSTab StatementTab 
            Height          =   9780
            Left            =   240
            TabIndex        =   30
            Top             =   405
            Width           =   5430
            _ExtentX        =   9578
            _ExtentY        =   17251
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   15326426
            ForeColor       =   -2147483646
            TabCaption(0)   =   "Statement"
            TabPicture(0)   =   "frmDebtors.frx":038A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label40"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label1"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "dtpStatementDate"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "dtpStatement"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "arvStatement"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "cmdLoadStatement"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cmdToExcel"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "cmdToPDF"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Statements (all)"
            TabPicture(1)   =   "frmDebtors.frx":03A6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Label2"
            Tab(1).Control(1)=   "Label3"
            Tab(1).Control(2)=   "dtpStatementDateAll"
            Tab(1).Control(3)=   "dtpStatementsAll"
            Tab(1).Control(4)=   "arvStatementAll"
            Tab(1).Control(5)=   "Command1"
            Tab(1).ControlCount=   6
            Begin VB.CommandButton cmdToPDF 
               BackColor       =   &H00E7E6D8&
               Caption         =   "PDF"
               Height          =   360
               Left            =   2640
               Style           =   1  'Graphical
               TabIndex        =   44
               Top             =   900
               Width           =   1380
            End
            Begin VB.CommandButton cmdToExcel 
               BackColor       =   &H00E7E6D8&
               Caption         =   "Spreadsheet"
               Height          =   360
               Left            =   4095
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   900
               Width           =   1380
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00E7E6D8&
               Caption         =   "&Show all statements"
               Height          =   390
               Left            =   -74175
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   900
               Width           =   1725
            End
            Begin VB.CommandButton cmdLoadStatement 
               BackColor       =   &H00E7E6D8&
               Caption         =   "&Show statement"
               Height          =   390
               Left            =   825
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   870
               Width           =   1725
            End
            Begin DDActiveReportsViewer2Ctl.ARViewer2 arvStatement 
               Height          =   3615
               Left            =   135
               TabIndex        =   32
               Top             =   1350
               Width           =   4740
               _ExtentX        =   8361
               _ExtentY        =   6376
               SectionData     =   "frmDebtors.frx":03C2
            End
            Begin MSComCtl2.DTPicker dtpStatement 
               Height          =   315
               Left            =   810
               TabIndex        =   33
               Top             =   465
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               Format          =   59113473
               CurrentDate     =   39980
            End
            Begin MSComCtl2.DTPicker dtpStatementDate 
               Height          =   345
               Left            =   3690
               TabIndex        =   35
               Top             =   435
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   609
               _Version        =   393216
               Format          =   59113473
               CurrentDate     =   39980
            End
            Begin DDActiveReportsViewer2Ctl.ARViewer2 arvStatementAll 
               Height          =   3615
               Left            =   -74865
               TabIndex        =   37
               Top             =   1365
               Width           =   4740
               _ExtentX        =   8361
               _ExtentY        =   6376
               SectionData     =   "frmDebtors.frx":03FE
            End
            Begin MSComCtl2.DTPicker dtpStatementsAll 
               Height          =   315
               Left            =   -74190
               TabIndex        =   38
               Top             =   480
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               Format          =   59113473
               CurrentDate     =   39980
            End
            Begin MSComCtl2.DTPicker dtpStatementDateAll 
               Height          =   345
               Left            =   -71310
               TabIndex        =   40
               Top             =   450
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   609
               _Version        =   393216
               Format          =   59113473
               CurrentDate     =   39980
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Statement date"
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   -72600
               TabIndex        =   41
               Top             =   495
               Width           =   1185
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Since"
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   -74850
               TabIndex        =   39
               Top             =   525
               Width           =   600
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Statement date"
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   2400
               TabIndex        =   36
               Top             =   480
               Width           =   1185
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Since"
               ForeColor       =   &H8000000D&
               Height          =   240
               Left            =   150
               TabIndex        =   34
               Top             =   510
               Width           =   600
            End
         End
      End
      Begin VB.Frame fr2 
         BackColor       =   &H00F7EDE8&
         Height          =   5655
         Left            =   15
         TabIndex        =   25
         Top             =   4920
         Width           =   14528
         Begin TabDlg.SSTab dbLedgerTab 
            Height          =   4875
            Left            =   465
            TabIndex        =   45
            Top             =   420
            Width           =   12450
            _ExtentX        =   21960
            _ExtentY        =   8599
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            BackColor       =   -2147483629
            ForeColor       =   -2147483646
            TabCaption(0)   =   "Ledger"
            TabPicture(0)   =   "frmDebtors.frx":043A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "LedgerGrid"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdMatching"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "cbSince"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Ledger - matching"
            TabPicture(1)   =   "frmDebtors.frx":0456
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "LedgerMGrid"
            Tab(1).ControlCount=   1
            Begin VB.CommandButton cbSince 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Last week"
               Height          =   450
               Left            =   165
               Style           =   1  'Graphical
               TabIndex        =   47
               TabStop         =   0   'False
               Top             =   570
               Width           =   6375
            End
            Begin VB.CommandButton cmdMatching 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Matching"
               Height          =   465
               Left            =   6615
               Style           =   1  'Graphical
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   570
               Width           =   870
            End
            Begin TrueOleDBGrid60.TDBGrid LedgerGrid 
               Height          =   3240
               Left            =   615
               OleObjectBlob   =   "frmDebtors.frx":0472
               TabIndex        =   48
               Top             =   1200
               Width           =   8925
            End
            Begin TrueOleDBGrid60.TDBGrid LedgerMGrid 
               Height          =   3255
               Left            =   -74820
               OleObjectBlob   =   "frmDebtors.frx":4AD9
               TabIndex        =   49
               Top             =   465
               Width           =   11970
            End
         End
      End
      Begin VB.Frame fr1 
         BackColor       =   &H00F7EDE8&
         Height          =   4635
         Left            =   0
         TabIndex        =   1
         Top             =   180
         Width           =   14528
         Begin VB.CommandButton cmdRemittance 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7E6D8&
            Caption         =   "New remittance"
            Height          =   465
            Left            =   10410
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   3180
            Width           =   1995
         End
         Begin VB.CommandButton cmdRecalcAgeing 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7E6D8&
            Caption         =   "Recalculate"
            Height          =   465
            Left            =   6975
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   3180
            Width           =   960
         End
         Begin VB.TextBox txtArg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00915A48&
            Height          =   375
            Left            =   1830
            TabIndex        =   23
            Text            =   "<Debtor_by_name_or_A/C_no>"
            Top             =   195
            Width           =   5265
         End
         Begin VB.Frame frBalances 
            BackColor       =   &H00F9F2EE&
            ForeColor       =   &H8000000D&
            Height          =   1200
            Left            =   15
            TabIndex        =   2
            Top             =   3855
            Width           =   7485
            Begin VB.TextBox txtCurBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   360
               Width           =   960
            End
            Begin VB.TextBox txt30Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   3255
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txt60Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   4290
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txt90Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   5325
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txt120PlusBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00DBFAFB&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   6345
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   360
               Width           =   990
            End
            Begin VB.TextBox txtBalance 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
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
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   360
               Width           =   1110
            End
            Begin VB.TextBox txtBF120PlusBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   6345
               Locked          =   -1  'True
               TabIndex        =   8
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBF90Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   5325
               Locked          =   -1  'True
               TabIndex        =   7
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBF60Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   4290
               Locked          =   -1  'True
               TabIndex        =   6
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBF30Bal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   3255
               Locked          =   -1  'True
               TabIndex        =   5
               Top             =   750
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.TextBox txtBFCurBal 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   2250
               Locked          =   -1  'True
               TabIndex        =   4
               Top             =   750
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.TextBox txtBFBalance 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00ECEAEA&
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
               Height          =   315
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   750
               Visible         =   0   'False
               Width           =   1110
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   1140
               TabIndex        =   22
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "This month"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   2250
               TabIndex        =   21
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "30 days"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   3240
               TabIndex        =   20
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "60 days"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   4260
               TabIndex        =   19
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "90 days"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   5310
               TabIndex        =   18
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "120+ days"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   255
               Left            =   6270
               TabIndex        =   17
               Top             =   150
               Width           =   975
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Current"
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
               Left            =   330
               TabIndex        =   16
               Top             =   390
               Width           =   675
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Month start"
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
               Left            =   60
               TabIndex        =   15
               Top             =   780
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin TrueOleDBGrid60.TDBGrid Grid 
            Height          =   2535
            Left            =   720
            OleObjectBlob   =   "frmDebtors.frx":9D49
            TabIndex        =   24
            Top             =   570
            Width           =   7095
         End
         Begin TrueOleDBGrid60.TDBGrid RemittanceGrid 
            Height          =   2535
            Left            =   7935
            OleObjectBlob   =   "frmDebtors.frx":D7CA
            TabIndex        =   29
            Top             =   570
            Width           =   4470
         End
      End
   End
End
Attribute VB_Name = "frmDebtors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmRemittancesBrowse As frmBrowseDBJNLs
Dim frmCashBookMaintenance As frmCashBookMaintenance
Dim frmRemittancePreview As frmCRemittancePreview
Dim frmCustJnl As frmCustJnl
Dim frmCashBook As frmCashBook
Dim cCust As c_Customer
Dim arStatement As arStatement_b
Dim arStatementAll As arStatement_All

Dim dteDate1 As Date
Dim dteDate2 As Date
Dim cJNL As c_JNL
Dim dCN As d_JNL
Dim XA As New XArrayDB
Dim enSince As enumSince
Dim rs As New ADODB.Recordset
Dim oSQL As New z_SQL
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
Dim strCustomerName As String
Dim oTRs As c_DebtorsTransPerTP
Dim oREMs As c_CustRemittances
Dim XB As New XArrayDB
Dim XR As New XArrayDB
Dim XRM As New XArrayDB
Dim flgLoading As Boolean
Dim oCust As a_Customer
Dim Res As Boolean
Dim oSM As z_StockManager
Private Sub cbSince_Click()
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
End Sub

Private Sub CC_OnDimValueClick(ByVal AxisSection As CCubeX2.IAxisSection, ByVal Level As Long)
    Set oCust = Nothing
    Set oCust = New a_Customer
    Res = oCust.Load(0, AxisSection.getValue(Level))
    LoadTransactions AxisSection.getValue(Level)
    LoadLedger
End Sub


Private Sub CM_SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    resize
End Sub

Private Sub cmdLoadStatement_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL

    If oCust Is Nothing Then Exit Sub
    oSQL.RunProc "[AgeInvoices]", Array(lngTPID), ""
   
    Set rs = New ADODB.Recordset
    oSQL.RunGetRecordset "SELECT * FROM vOpenItemAll_1 WHERE TPID = " & CStr(oCust.ID) & " AND (BALANCE <> 0 OR dbDocDate > '" & ReverseDate(Me.dtpStatement) & "') ORDER BY AGE,dbDocDate DESC,crDocDate DESC ", enText, "", "", rs
    Set arStatement = New arStatement_b
    arStatement.component rs, oPC.Configuration.DefaultCompany, oCust.NameAndCode(100), oPC.Configuration.DefaultCompany.StreetAddress, oPC.Configuration.DefaultCompany.BankDetails, oPC.Configuration.DefaultCompany.VatNumber
    arvStatement.Zoom = 75
    arvStatement.ReportSource = arStatement

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.cmdLoadStatement_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMatching_Click()
Dim frm As New frmPaymentMatch

    Screen.MousePointer = vbHourglass
    frm.component oCust.ID, oCust.NameAndCode(50), Top, Left
    Screen.MousePointer = vbDefault
    frm.Show

End Sub

Private Sub LoadDebtorsStatement()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    oSQL.RunProc "[AgeInvoices]", Array(oCust.ID), ""
    oCust.Reload
    LoadTransactions
    LoadLedger
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadDebtorsStatement"
End Sub

Private Sub cmdRecalcAgeing_Click()
Dim oSQL As New z_SQL

    oSQL.RunProc "[AgeInvoices]", Array(oCust.ID), ""
    oCust.Reload
    LoadTransactions oCust.ID
    LoadLedger
    
End Sub

Private Sub cmdRemittance_Click()
    If oCust Is Nothing Then
        MsgBox "Select a debtor first.", vbInformation + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Set frmCustomerRemittance = New frmCustomerRemittance
    If oCust.ParentCustomerID > 0 Then
        If MsgBox("This customer appears to have accounts paid by " & oCust.ParentCustomerName & "." & vbCrLf & "Continue with remittance?", vbQuestion + vbOKCancel, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    frmCustomerRemittance.component lngTPID, strCustomerName, 0, Date, "", 0
    frmCustomerRemittance.Show

End Sub

Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    arStatement.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & IIf(Right(oPC.LocalFolder, 1) = "\", "", "\") & "TEMP\" & "Statement_" & Replace(Replace(strCustomerName, "/", "-"), " ", "_") & "_" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fs)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(arStatement.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    arStatement.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & IIf(Right(oPC.LocalFolder, 1) = "\", "", "\") & "TEMP\" & "Statement_" & Replace(Replace(strCustomerName, "/", "-"), " ", "_") & "_" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fs)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(arStatement.Pages)
    OpenFileWithApplication fn, enExcel
End Sub


Private Sub Command1_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL


    Set rs = New ADODB.Recordset
    oSQL.RunGetRecordset "SELECT *,Fullname + '(' + Acno + ')' fullnameAcno FROM vOpenItemAll_1 WHERE (BALANCE <> 0 OR dbDocDate > '" & ReverseDate(Me.dtpStatement) & "') ORDER BY fullnameAcno,AGE,dbDocDate DESC,crDocDate DESC ", enText, "", "", rs
    Set arStatementAll = New arStatement_All
    arvStatementAll.ReportSource = arStatementAll
    arStatementAll.component rs, oPC.Configuration.DefaultCompany, "TESTING", oPC.Configuration.DefaultCompany.StreetAddress, oPC.Configuration.DefaultCompany.BankDetails, oPC.Configuration.DefaultCompany.VatNumber
    arvStatementAll.Zoom = 75

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.Command1_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub dbLedgerTab_Click(PreviousTab As Integer)
    resize
End Sub


Private Sub Form_Load()
    flgLoading = True
    Me.Grid.Top = 750
    Me.LedgerGrid.Top = 750
    Me.cbSince.Top = 270
    Me.Top = 200
    Me.Left = 50
    Me.Width = 6600
    
    Me.Height = 4000
    enSince = OptionLoop(enSince, 5)
    
    cbSince.Caption = TranslateSince(CInt(enSince))
    SetDateArgs
    SetGridLayout Me.Grid, Me.Name & Grid.Name
    SetGridLayout Me.LedgerGrid, Me.Name & LedgerGrid.Name
    SetFormSize Me
    SetCM Me, CM
    StatementTab.Tab = 0
    Me.WindowState = vbNormal
    flgLoading = False
    
End Sub
Private Sub LoadBrowse()
    Find
    LoadArray
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.ReDim 1, cCust.Count, 1, 6
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Fullname2
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
            XA.Value(lngIndex, 4) = objItem.Balance
            XA.Value(lngIndex, 5) = objItem.ID
            XA.Value(lngIndex, 6) = objItem.Blocked
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadArray"
End Sub

Private Sub Grid_DblClick()
Dim lngID As Long
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL


    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oCust = Nothing
    Set oCust = New a_Customer
    lngTPID = FNN(XA(Grid.Bookmark, 5))
    strCustomerName = FNS(XA(Grid.Bookmark, 1))
    oCust.Load lngTPID
    LoadTransactions FNN(XA(Grid.Bookmark, 5))
    LoadLedger
    
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oSQL.RunGetRecordset "SELECT * FROM vOpenItemAll_1 WHERE TPID = " & CStr(oCust.ID) & " AND (BALANCE <> 0 OR dbDocDate > '" & ReverseDate(Me.dtpStatement) & "') ", enText, "", "", rs
    LoadLedgerMGrid rs
    
    LoadRemittances FNN(XA(Grid.Bookmark, 5))
    LoadRemittanceGrid
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadLedgerMGrid(rs As ADODB.Recordset)
Dim i  As Integer
Dim j As Integer
    On Error GoTo errHandler
    If rs.EOF And rs.BOF Then
        XRM.Clear
        XRM.ReDim 1, 0, 1, 8
        LedgerMGrid.ReBind
        LedgerMGrid.Caption = strCustomerName
        LedgerMGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"
       Exit Sub
    End If
    
    For i = 1 To LedgerMGrid.Columns.Count
        LedgerMGrid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "LedgerMGrid", CStr(i), LedgerMGrid.Columns(i - 1).Width)
    Next
    
    XRM.Clear
    LedgerMGrid.Refresh
    i = 1
    j = 1
    Do While Not rs.EOF
            XRM.ReDim 1, j, 1, 15
            XRM.Value(j, 1) = FNS(rs.Fields("dbDoc"))
            XRM.Value(j, 2) = FNS(rs.Fields("dbDocType"))
            XRM.Value(j, 3) = FNS(rs.Fields("dbDocDate"))
            XRM.Value(j, 4) = FNS(rs.Fields("dbAmt"))
            XRM.Value(j, 5) = FNS(rs.Fields("crDoc"))
            XRM.Value(j, 6) = FNS(rs.Fields("crDocType"))
            XRM.Value(j, 7) = FNS(rs.Fields("crDocDate"))
            XRM.Value(j, 8) = FNS(rs.Fields("crAmt"))
            XRM.Value(j, 9) = IIf(FNS(rs.Fields("dbDoc")) = "unalloc", "unalloc", "")
            rs.MoveNext
        '    XRM.Value(j, 7) = oREMs.Item(j).TRID
            j = j + 1
    Loop
    If j > 1 Then   'only if there are contents to display
        XRM.QuickSort 1, XRM.UpperBound(1), 9, XORDER_DESCEND, XTYPE_STRING, 3, XORDER_DESCEND, XTYPE_DATE, 7, XORDER_DESCEND, XTYPE_DATE, 6, XORDER_ASCEND, XTYPE_STRING
        LedgerMGrid.Array = XRM
        LedgerMGrid.ReBind
        LedgerMGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadLedgerMGrid"
End Sub

Private Sub LoadLedger(Optional pAcno As String)
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double

    Me.txtCurBal = oCust.BalanceCurF
    Me.txt30Bal = oCust.Balance30F
    Me.txt60Bal = oCust.Balance60F
    Me.txt90Bal = oCust.Balance90F
    Me.txt120PlusBal = oCust.Balance120F
    Me.txtBalance = oCust.BalanceF
    
    If oTRs.Count = 0 Then
        XB.Clear
        XB.ReDim 1, 0, 1, 8
        LedgerGrid.ReBind
        LedgerGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"
        Exit Sub
    End If
    
    For i = 1 To LedgerGrid.Columns.Count
        LedgerGrid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "LedgerGrid", CStr(i), LedgerGrid.Columns(i - 1).Width)
    Next
    
    XB.Clear
    LedgerGrid.ReBind
   ' Set oSQL = New z_SQL
   ' oSM.RecalculateTPBalance oCust.ID
    i = 1
    j = 1
    Do While i <= oTRs.Count
        If oTRs.Item(i).DocType <> "BF" Then
            XB.ReDim 1, j, 1, 8
            XB.Value(j, 1) = oTRs.Item(i).DOCCode
            XB.Value(j, 2) = oTRs.Item(i).DocType
            XB.Value(j, 3) = oTRs.Item(i).DocDateF
            XB.Value(j, 4) = oTRs.Item(i).DebitF
            XB.Value(j, 5) = oTRs.Item(i).CreditF
            XB.Value(j, 6) = oTRs.Item(i).Memo
            XB.Value(j, 7) = oTRs.Item(i).DOCID
            XB.Value(j, 7) = oTRs.Item(i).DOCCaptureDate
 '           dblBal = dblBal + oTRs.Item(i).Debit
 '           dblBal = dblBal - oTRs.Item(i).Credit
            j = j + 1
        End If
        i = i + 1
    Loop
    XB.QuickSort 1, XB.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE, 1, XORDER_DESCEND, XTYPE_STRING
    LedgerGrid.Array = XB
    LedgerGrid.ReBind
    LedgerGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadLedger"
End Sub
Private Sub LoadRemittanceGrid()
Dim i  As Integer
Dim j As Integer
    On Error GoTo errHandler
    If oREMs.Count = 0 Then
        XR.Clear
        XR.ReDim 1, 0, 1, 8
        RemittanceGrid.ReBind
        RemittanceGrid.Caption = strCustomerName
        RemittanceGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"
       Exit Sub
    End If
    
    For i = 1 To RemittanceGrid.Columns.Count
        RemittanceGrid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "RemittanceGrid", CStr(i), RemittanceGrid.Columns(i - 1).Width)
    Next
    
    XR.Clear
    RemittanceGrid.ReBind
    i = 1
    j = 1
    Do While j <= oREMs.Count
            XR.ReDim 1, j, 1, 8
            XR.Value(j, 1) = oREMs.Item(j).DocumentCode
            XR.Value(j, 2) = oREMs.Item(j).DocumentNominalDateF
            XR.Value(j, 3) = oREMs.Item(j).DocumentReference
            XR.Value(j, 8) = oREMs.Item(j).DocumentNumberDate
            XR.Value(j, 7) = oREMs.Item(j).TRID
'            XR.Value(j, 6) = oREMs.Item(j).Memo
            j = j + 1
    Loop
    XR.QuickSort 1, XR.UpperBound(1), 8, XORDER_DESCEND, XTYPE_DATE
    RemittanceGrid.Array = XR
    RemittanceGrid.ReBind
    RemittanceGrid.Caption = oCust.Fullname & " (" & oCust.AcNo & ")"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadRemittanceGrid"
End Sub

Private Sub LoadTransactions(Optional CustID As Long, Optional pAcno As String)
    On Error GoTo errHandler
    Set oTRs = Nothing
    Set oTRs = New c_DebtorsTransPerTP
    oTRs.Load CustID, CDate("2000-01-01"), False, pAcno
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadTransactions"
End Sub
Private Sub LoadRemittances(Optional CustID As Long, Optional pAcno As String)
    On Error GoTo errHandler
    Set oREMs = Nothing
    Set oREMs = New c_CustRemittances
    oREMs.Load CustID, CDate("2000-01-01"), False, pAcno
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LoadRemittances"

End Sub

Private Sub Find()
    On Error GoTo errHandler
Dim bNotFound As Boolean
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
Dim byear As Boolean
Dim yr As String
Dim mth As String
Dim strDate1 As String
Dim strDate2 As String
Dim lngCount As Long

    bNotFound = False
    If Left(txtArg, 3) = "yr=" Then byear = True
    
    If txtArg > " " And Not (byear) Then
        'Search for Reference
        Set cJNL = Nothing
        Set cJNL = New c_JNL
        cJNL.Load bNotFound, 0, "", txtArg, dteDate1, dteDate2
        If bNotFound Then
            'Search for customer by AcJNLO
            Set cJNL = Nothing
            Set cJNL = New c_JNL
            SetDateArgs
            cJNL.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
            If bNotFound Then
               Set frm = New frmBrowseCustomers2
               frm.component txtArg, lngCount
               If lngCount > 1 Then
                    frm.Show vbModal
                    lngTPID = frm.CustomerID
                    Unload frm
                ElseIf lngCount = 1 Then
                    lngTPID = frm.CustomerID
                    Unload frm
                End If
               If lngTPID > 0 Then
                    Set cJNL = Nothing
                    Set cJNL = New c_JNL
                    SetDateArgs
                    cJNL.Load bNotFound, lngTPID, "", "", dteDate1, dteDate2
               End If
            End If
        Else
            enSince = 1
            cbSince.Caption = TranslateSince(1)
        End If
    Else
        Set cJNL = Nothing
        Set cJNL = New c_JNL
        If byear Then
            yr = Mid(txtArg, 4, 4)
            mth = Mid(txtArg, 9, 2)
            If mth > "" Then
                strDate1 = yr & "-" & mth & "-01"
                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
            Else
                strDate1 = yr & "-01-01"
                strDate2 = yr & "-12-31"
            End If
            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
                SetDateArgs
            Else
                dteDate1 = CDate(strDate1)
                dteDate2 = CDate(strDate2)
            End If
        Else
            SetDateArgs
        End If
        cJNL.Load bNotFound, 0, "", "", dteDate1, dteDate2
    End If

EXIT_Handler:
    mSetfocus Grid
    MousePointer = vbDefault
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.Find"
End Sub


Private Sub Grid_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowTypeXA(ColIndex), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    Grid.Refresh
    Screen.MousePointer = vbDefault

End Sub
Private Function GetRowTypeXA(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 0, 1, 2
            GetRowTypeXA = XTYPE_STRING
        Case 3
            GetRowTypeXA = XTYPE_CURRENCY
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.GetRowTypeXA(ColIndex)", ColIndex
End Function

Private Sub LedgerGrid_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler

Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex)
    LedgerGrid.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.LedgerGrid_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 2
            GetRowType = XTYPE_DATE
        Case 3, 4
            GetRowType = XTYPE_CURRENCY
        Case Else
            GetRowType = XTYPE_STRING
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.GetRowType(ColIndex)", ColIndex
End Function


Private Sub RemittanceGrid_DblClick()
Dim frm As New frmCRemittancePreview

    frm.component FNN(XR(RemittanceGrid.Bookmark, 7)), strCustomerName, 0
    frm.Show
    
End Sub

Private Sub StatementTab_Click(PreviousTab As Integer)
    If StatementTab.Tab = 0 Then
        arvStatement.Visible = True
        arvStatementAll.Visible = False
        arvStatement.Width = NonNegative_Lng(Me.frStatement.Width - 350)
        arvStatement.Height = NonNegative_Lng(Me.frStatement.Height - 1400)
    Else
        arvStatement.Visible = False
        arvStatementAll.Visible = True
        arvStatementAll.Width = NonNegative_Lng(Me.frStatement.Width - 350)
        arvStatementAll.Height = NonNegative_Lng(Me.frStatement.Height - 1400)
    End If

End Sub


Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    
    If KeyAscii = 13 Then  ' The ENTER key.
       HandleResults
        If cCust.Count > 1 Then
            On Error Resume Next
            Grid.SetFocus
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub SetDateArgs()
    On Error GoTo errHandler
    Select Case enSince
    Case enAny
        dteDate1 = CDate("1995-01-01")
        dteDate2 = DateAdd("d", 1, Date)
    Case enWeek
        dteDate1 = DateAdd("d", -7, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enMonth
        dteDate1 = DateAdd("m", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enQuarter
        dteDate1 = DateAdd("q", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    Case enYear
        dteDate1 = DateAdd("yyyy", -1, Date)
        dteDate2 = DateAdd("d", 1, Date)
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.SetDateArgs"
End Sub
'Private Sub pBuildMenus()
'
'    With Me.SmartMenuXP.MenuItems
''        ' Root > File...
'        .Add 0, "keyFile", , "&File"
'        .Add "keyFile", "keyExit", , "E&xit", , vbAltMask, vbKeyQ
''
'''Root>view
'        .Add 0, "keyView", , "View"
'        .Add "keyView", "KeyRemittances", , "Browse &remittances"
'        .Add "keyView", "KeyJournals", , "Browse &journals"
'        .Add "keyView", "KeyLedgerView", , "Browse customer &ledger"
'        .Add "keyView", "KeyStatementView", , "&Statement"
'
'''Root>Actions
'        .Add 0, "keyActions", , "Actions"
'        .Add "keyActions", "KeyCaptureRemittance", , "&Capture remittance"
'        .Add "keyActions", "KeyMatch", , "&Matching"
'        .Add "keyActions", "KeyProduceStatements", , "&Produce statements"
'        .Add "keyActions", "KeyOpenCashBook", , "C&ash book"
'        .Add "keyActions", "KeyTemplate", , "Cash book &template"
'
'
'    End With
'
'    SmartMenuXP.Font.Name = "Ms Sans Serif"
'    SmartMenuXP.BackColor = &HF7EDE8
'    SmartMenuXP.Font.Size = 9
'
'End Sub
'
'Private Sub SmartMenuXP_Click(ByVal ID As Long)
'    With SmartMenuXP.MenuItems
'
'        Select Case .key(ID)
'            Case "keyExit"
'                    Unload Me
'            Case "KeyRemittances"
'                Set frmRemittancesBrowse = New frmBrowseDBJNLs
'                frmRemittancesBrowse.Show
'            Case "KeyCaptureRemittance"
'                Set frm = New frmBrowseCustomers2
'                frm.Show vbModal
'                lngTPID = frm.CustomerID
'                strCustomerName = frm.CustomerName
'                Unload frm
'                If lngTPID = 0 Then Exit Sub
'                Set frmCustomerRemittance = New frmCustomerRemittance
'                frmCustomerRemittance.component lngTPID, strCustomerName, Me.hwnd, Date, "", 0
'                frmCustomerRemittance.Show
'                Set oTRs = Nothing
'                Set oTRs = New c_DebtorsTransPerTP
'                oTRs.Load lngTPID, CDate("2000-01-01")
'                LedgerGrid.Caption = strCustomerName
'            Case "KeyOpenCashBook"
'                Set frmCashBook = New frmCashBook
'                frmCashBook.Show
'            Case "KeyTemplate"
'                Set frmCashBookMaintenance = New frmCashBookMaintenance
'                frmCashBookMaintenance.Show
'        End Select
'
'    End With
'
'End Sub

Private Sub Form_Resize()
    resize
End Sub
Private Sub resize()
    txtArg.Top = fr1.Top + 10
    txtArg.Left = fr1.Left + 50
    
    Grid.Left = fr1.Left + 50
    Grid.Top = fr1.Top + 470
    Grid.Width = NonNegative_Lng(fr1.Width * 0.65)
    Grid.Height = NonNegative_Lng(fr1.Height - 1800)
    
    RemittanceGrid.Left = fr1.Left + Grid.Width + 300
    RemittanceGrid.Top = fr1.Top + 470
    RemittanceGrid.Width = NonNegative_Lng(fr1.Width * 0.3)
    RemittanceGrid.Height = NonNegative_Lng(fr1.Height - 1800)
    Me.cmdRemittance.Top = fr1.Top + RemittanceGrid.Height + 550
    cmdRemittance.Left = NonNegative_Lng(RemittanceGrid.Left + RemittanceGrid.Width - cmdRemittance.Width)
    
    frBalances.Left = fr1.Left + 50
    frBalances.Top = NonNegative_Lng(fr1.Height - 1000)
    frBalances.Height = 900
    
    
    CM.Size = 80
    cmdRecalcAgeing.Top = frBalances.Top + 420
    cmdRecalcAgeing.Left = frBalances.Left + 7500

    dbLedgerTab.Top = NonNegative_Lng(fr2.Top - fr1.Height - 300)
    dbLedgerTab.Left = fr2.Left + 80
    dbLedgerTab.Height = NonNegative_Lng(fr2.Height - 400)
    dbLedgerTab.Width = NonNegative_Lng(fr2.Width - 250)
    
    If dbLedgerTab.Tab = 0 Then
        LedgerMGrid.Visible = False
        LedgerGrid.Visible = True
        cbSince.Visible = True
        cmdMatching.Visible = True
        
        cbSince.Top = NonNegative_Lng(dbLedgerTab.Top + 250)
        cbSince.Left = dbLedgerTab.Left + 50
        cmdMatching.Top = cbSince.Top
        LedgerGrid.Left = fr2.Left + 100
        LedgerGrid.Top = NonNegative_Lng(dbLedgerTab.Top + 800)
        LedgerGrid.Width = NonNegative_Lng(dbLedgerTab.Width - 260)
        LedgerGrid.Height = NonNegative_Lng(dbLedgerTab.Height - 1000)
    Else
        LedgerMGrid.Visible = True
        LedgerGrid.Visible = False
        cbSince.Visible = False
        cmdMatching.Visible = False
        LedgerMGrid.Left = fr2.Left + 100
        LedgerMGrid.Top = NonNegative_Lng(dbLedgerTab.Top + 250)
        LedgerMGrid.Width = NonNegative_Lng(dbLedgerTab.Width - 260)
        LedgerMGrid.Height = NonNegative_Lng(dbLedgerTab.Height - 1000)
    End If
    
    
    StatementTab.Width = NonNegative_Lng(Me.frStatement.Width - 450)
    StatementTab.Height = NonNegative_Lng(Me.frStatement.Height - 650)
    
    If StatementTab.Tab = 0 Then
        arvStatement.Visible = True
        arvStatementAll.Visible = False
        arvStatement.Width = NonNegative_Lng(Me.frStatement.Width - 700)
        arvStatement.Height = NonNegative_Lng(StatementTab.Height - 1500)
    Else
        arvStatement.Visible = False
        arvStatementAll.Visible = True
        arvStatementAll.Width = NonNegative_Lng(frStatement.Width - 700)
        arvStatementAll.Height = NonNegative_Lng(StatementTab.Height - 1500)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.Grid, Me.Name & Grid.Name
    SaveLayout Me.LedgerGrid, Me.Name & LedgerGrid.Name
    SaveLayout Me.LedgerMGrid, Me.Name & LedgerMGrid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
    SaveSplits Me.Name, Me.CM
End Sub
Private Sub txtArg_GotFocus()
    AutoSelect txtArg
End Sub

Private Sub Grid_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      PopupMenu Forms(0).mnuDebtorPopup   ' Display the File menu as a
                        ' pop-up menu.
   End If
End Sub


Private Sub HandleResults(Optional plngCount As Long)
    On Error GoTo errHandler
    If txtArg = "" Then Exit Sub
    Set cCust = Nothing
    Set cCust = New c_Customer
    Screen.MousePointer = vbHourglass
    
    cCust.LoadEasy Replace(txtArg, "'", "''"), False ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    plngCount = cCust.Count
    LoadArray
    Grid.ReBind
    Grid.Enabled = True
    If cCust.Count = 1 Then
        LoadTransactions FNN(XA(Grid.Bookmark, 5))
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors.HandleResults(plngCount)", plngCount
End Sub

Public Sub mnuNewRemittance()
    lngTPID = FNN(XA(Grid.Bookmark, 5))
    If lngTPID > 0 Then
        Set frmCustomerRemittance = New frmCustomerRemittance
        frmCustomerRemittance.component lngTPID, FNS(XA(Grid.Bookmark, 1)), Me.hwnd, Date, "", 0
        frmCustomerRemittance.Show
    End If
End Sub
