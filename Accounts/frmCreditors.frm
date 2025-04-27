VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B4B5B73C-172E-47B1-BFC2-C6F740957D01}#1.0#0"; "VB Control Manager.ocx"
Begin VB.Form frmCreditors 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Creditors"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19155
   FillColor       =   &H00FCF2EB&
   Icon            =   "frmCreditors.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   19155
   WindowState     =   2  'Maximized
   Begin VBControlManager.ControlManager CM 
      Height          =   10530
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   18574
      BackColor       =   9525832
      Size            =   4
      TitleBar_CloseVisible=   0   'False
      Begin VB.Frame fr3 
         BackColor       =   &H00F7EDE8&
         Height          =   10350
         Left            =   8010
         TabIndex        =   28
         Top             =   180
         Width           =   7995
         Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
            Height          =   9285
            Left            =   330
            TabIndex        =   42
            Top             =   690
            Width           =   6930
            _ExtentX        =   12224
            _ExtentY        =   16378
            SectionData     =   "frmCreditors.frx":038A
         End
      End
      Begin VB.Frame fr2 
         BackColor       =   &H00F7EDE8&
         Height          =   6795
         Left            =   0
         TabIndex        =   27
         Top             =   3720
         Width           =   7965
         Begin TabDlg.SSTab LedgerTab 
            Height          =   4965
            Left            =   165
            TabIndex        =   29
            Top             =   315
            Width           =   7620
            _ExtentX        =   13441
            _ExtentY        =   8758
            _Version        =   393216
            Tab             =   1
            TabHeight       =   520
            BackColor       =   16053473
            ForeColor       =   -2147483646
            TabCaption(0)   =   "Ledger"
            TabPicture(0)   =   "frmCreditors.frx":03C6
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "LedgerGrid"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cbSince"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "cmdShowStatement"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Remittance preparation"
            TabPicture(1)   =   "frmCreditors.frx":03E2
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "lblPaymentGrid"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "lblDueIn"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "Label1"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "PaymentsGrid"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "cmdPrepare"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "cmdGeneratePayments"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).Control(6)=   "txtTotal"
            Tab(1).Control(6).Enabled=   0   'False
            Tab(1).Control(7)=   "txtDueDays"
            Tab(1).Control(7).Enabled=   0   'False
            Tab(1).Control(8)=   "chkSelectedCreditorOnly"
            Tab(1).Control(8).Enabled=   0   'False
            Tab(1).Control(9)=   "cmdPrintPaymentOrder"
            Tab(1).Control(9).Enabled=   0   'False
            Tab(1).ControlCount=   10
            TabCaption(2)   =   "Payment confirmation"
            TabPicture(2)   =   "frmCreditors.frx":03FE
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label2"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).Control(1)=   "Label4"
            Tab(2).Control(1).Enabled=   0   'False
            Tab(2).Control(2)=   "lblCreditorname"
            Tab(2).Control(2).Enabled=   0   'False
            Tab(2).Control(3)=   "ConfirmationGrid"
            Tab(2).Control(3).Enabled=   0   'False
            Tab(2).Control(4)=   "txtTotalPayments"
            Tab(2).Control(4).Enabled=   0   'False
            Tab(2).Control(5)=   "cmdCreatePaymentRecords"
            Tab(2).Control(5).Enabled=   0   'False
            Tab(2).Control(6)=   "Text1"
            Tab(2).Control(6).Enabled=   0   'False
            Tab(2).ControlCount=   7
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               ForeColor       =   &H8000000D&
               Height          =   300
               Left            =   -69825
               TabIndex        =   47
               Text            =   "1"
               Top             =   690
               Width           =   1305
            End
            Begin VB.CommandButton cmdCreatePaymentRecords 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Create payment transactions"
               Enabled         =   0   'False
               Height          =   450
               Left            =   -70080
               Style           =   1  'Graphical
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   4380
               Width           =   2565
            End
            Begin VB.TextBox txtTotalPayments 
               Alignment       =   2  'Center
               ForeColor       =   &H8000000D&
               Height          =   300
               Left            =   -71850
               TabIndex        =   44
               Text            =   "1"
               Top             =   690
               Width           =   615
            End
            Begin VB.CommandButton cmdShowStatement 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Statement"
               Height          =   315
               Left            =   -69150
               Style           =   1  'Graphical
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   4545
               Width           =   1620
            End
            Begin VB.CommandButton cmdPrintPaymentOrder 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Payment order"
               Height          =   345
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   4515
               Width           =   2115
            End
            Begin VB.CheckBox chkSelectedCreditorOnly 
               Caption         =   "Selected creditor"
               Enabled         =   0   'False
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   1785
               TabIndex        =   40
               Top             =   390
               Width           =   1575
            End
            Begin VB.TextBox txtDueDays 
               Alignment       =   2  'Center
               ForeColor       =   &H8000000D&
               Height          =   300
               Left            =   675
               TabIndex        =   38
               Text            =   "1"
               Top             =   390
               Width           =   390
            End
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3345
               TabIndex        =   35
               Text            =   "txtTotal"
               Top             =   4485
               Width           =   1845
            End
            Begin VB.CommandButton cmdGeneratePayments 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Update actions"
               Height          =   345
               Left            =   5415
               Style           =   1  'Graphical
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   4500
               Width           =   2115
            End
            Begin VB.CommandButton cmdPrepare 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Prepare remittances"
               Height          =   300
               Left            =   3420
               Style           =   1  'Graphical
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   390
               Width           =   1545
            End
            Begin VB.CommandButton cbSince 
               Appearance      =   0  'Flat
               BackColor       =   &H00E7E6D8&
               Caption         =   "Last week"
               Height          =   450
               Left            =   -74955
               Style           =   1  'Graphical
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   405
               Width           =   2025
            End
            Begin TrueOleDBGrid60.TDBGrid LedgerGrid 
               Height          =   3240
               Left            =   -74940
               OleObjectBlob   =   "frmCreditors.frx":041A
               TabIndex        =   30
               Top             =   945
               Width           =   7845
            End
            Begin TrueOleDBGrid60.TDBGrid PaymentsGrid 
               Height          =   3240
               Left            =   -1710
               OleObjectBlob   =   "frmCreditors.frx":6845
               TabIndex        =   31
               Top             =   915
               Width           =   9090
            End
            Begin TrueOleDBGrid60.TDBGrid ConfirmationGrid 
               Height          =   2805
               Left            =   -74940
               OleObjectBlob   =   "frmCreditors.frx":D852
               TabIndex        =   49
               Top             =   1095
               Width           =   7410
            End
            Begin VB.Label lblCreditorname 
               Alignment       =   2  'Center
               BackColor       =   &H00F7EDE8&
               BorderStyle     =   1  'Fixed Single
               Height          =   300
               Left            =   -74460
               TabIndex        =   50
               Top             =   345
               Width           =   6390
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Total value"
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   -70830
               TabIndex        =   48
               Top             =   690
               Width           =   885
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Total number of payments made"
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   -74490
               TabIndex        =   45
               Top             =   690
               Width           =   2595
            End
            Begin VB.Label Label1 
               Caption         =   "days"
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   1140
               TabIndex        =   39
               Top             =   435
               Width           =   405
            End
            Begin VB.Label lblDueIn 
               Caption         =   "Due in "
               ForeColor       =   &H8000000D&
               Height          =   285
               Left            =   135
               TabIndex        =   37
               Top             =   435
               Width           =   585
            End
            Begin VB.Label lblPaymentGrid 
               BackColor       =   &H00CDFAFA&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   5130
               TabIndex        =   36
               Top             =   375
               Width           =   2310
            End
         End
      End
      Begin VB.Frame fr1 
         BackColor       =   &H00F7EDE8&
         Height          =   3345
         Left            =   -15
         TabIndex        =   1
         Top             =   180
         Width           =   7950
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
            Left            =   720
            TabIndex        =   25
            Text            =   "<Creditor_by_name_or_A/C_no>"
            Top             =   180
            Width           =   5265
         End
         Begin VB.Frame frBalances 
            BackColor       =   &H00F9F2EE&
            ForeColor       =   &H8000000D&
            Height          =   1200
            Left            =   15
            TabIndex        =   4
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
               TabIndex        =   16
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
               TabIndex        =   15
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
               TabIndex        =   14
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
               TabIndex        =   13
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
               TabIndex        =   12
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
               TabIndex        =   11
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
               TabIndex        =   10
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
               TabIndex        =   9
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
               TabIndex        =   8
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
               TabIndex        =   7
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
               TabIndex        =   6
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
               TabIndex        =   5
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
               TabIndex        =   24
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               TabIndex        =   19
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
               TabIndex        =   18
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
               TabIndex        =   17
               Top             =   780
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdMatchPayments 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Match payments"
            Height          =   450
            Left            =   7620
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3870
            Width           =   870
         End
         Begin VB.CommandButton cmdInvoiceAge 
            Appearance      =   0  'Flat
            BackColor       =   &H00E7E6D8&
            Caption         =   "Last week"
            Height          =   450
            Left            =   6015
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   165
            Width           =   1320
         End
         Begin TrueOleDBGrid60.TDBGrid Grid 
            Height          =   2265
            Left            =   705
            OleObjectBlob   =   "frmCreditors.frx":124AB
            TabIndex        =   26
            Top             =   765
            Width           =   7095
         End
      End
   End
End
Attribute VB_Name = "frmCreditors"
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
Dim cVendors As c_Supplier

Dim dteDate1 As Date
Dim dteDate2 As Date
Dim cInvoices As c_Invoices
Dim XA As New XArrayDB
Dim enSince As enumSince
Dim rs As New ADODB.Recordset
Dim oSQL As New z_SQL
Dim frm As frmBrowseCustomers2
Dim lngTPID As Long
Dim strCustomerName As String
Dim oTRs As c_CreditorsTransactionsPerTP
Dim oREMs As c_RemittancesInPreparation
Dim XB As New XArrayDB
Dim XC As New XArrayDB
Dim flgLoading As Boolean
Dim oVendor As a_Supplier
Dim Res As Boolean
Dim oSM As z_StockManager
Dim dblTotal As Double

Private Sub cbSince_Click()
    enSince = OptionLoop(enSince, 5)
    cbSince.Caption = TranslateSince(CInt(enSince))
End Sub

Private Sub CM_SplitterMoveEnd(ByVal IdSplitter As Long, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    resize
End Sub

Private Sub cmdCreatePaymentRecords_Click()
Dim oSQL As New z_SQL
    
    oSQL.CreateCreditorsPayments
    
End Sub

Private Sub cmdGeneratePayments_Click()
    GeneratePaymentsOrder
End Sub

Private Sub cmdPrepare_Click()
Dim oSQL As New z_SQL
Dim lngDueDays As Long

   ' oSQL.PrepareRemittances
    If ConvertToLng(txtDueDays, lngDueDays) = False Then
        MsgBox "Invalid due days value", vbInformation + vbOKOnly, "Invalid due days"
        Exit Sub
    End If
    If oVendor Is Nothing Then
        LoadRemittancesInPreparation lngDueDays, chkSelectedCreditorOnly = 1, 0
    Else
        LoadRemittancesInPreparation lngDueDays, chkSelectedCreditorOnly = 1, oVendor.ID
    End If
    LoadRemittancesInPreparationGrid
    RefreshTotal
End Sub

Private Sub cmdPrintPaymentOrder_Click()
Dim arPO As arPaymentOrder
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL

    Set rs = oSQL.GetPaymentOrder

    Set arPO = Nothing
    Set arPO = New arPaymentOrder

    arPO.Visible = False

    Set arv.ReportSource = arPO
    arv.Zoom = 75
    arPO.component rs

End Sub

Private Sub cmdShowStatement_Click()
Dim arS As New arStatement_b
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL

    Set rs = oSQL.GetPaymentOrder

    Set arS = Nothing
    Set arS = New arPaymentOrder

    arS.Visible = False

    Set arv.ReportSource = arS
    arv.Zoom = 75

    arS.component rs, oPC.Configuration.DefaultCompany, "Creditors statement", "Details", "Bank details", "VAT number"

End Sub

Private Sub Form_Load()
    flgLoading = True
    Me.Grid.Top = 750
    Me.LedgerGrid.Top = 750
    Me.cbSince.Top = 270
    Me.Top = 200
    Me.Left = 50
    Me.Width = 6600
   ' PaymentsGrids.Visible = True
    Me.Height = 4000
    enSince = OptionLoop(enSince, 5)
    
    cbSince.Caption = TranslateSince(CInt(enSince))
    SetDateArgs
    
    SetGridLayout Me.Grid, Me.Name & Grid.Name
    SetGridLayout Me.LedgerGrid, Me.Name & LedgerGrid.Name
    SetGridLayout Me.PaymentsGrid, Me.Name & PaymentsGrid.Name
    
    SetFormSize Me
    SetCM Me, CM
    
    Me.WindowState = vbNormal
    flgLoading = False
    
End Sub
Private Sub LoadBrowse()
'    Find
'    LoadArray
End Sub
Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Supplier
Dim itmList As ListItem
Dim lngIndex As Long
Dim i As Integer
    XA.ReDim 1, cVendors.Count, 1, 6
    For lngIndex = 1 To cVendors.Count
        With objItem
            Set objItem = cVendors.Item(lngIndex)
            XA.Value(lngIndex, 1) = objItem.Name
            XA.Value(lngIndex, 2) = objItem.AcNo
            XA.Value(lngIndex, 3) = objItem.Phone
           ' XA.Value(lngIndex, 4) = objItem.Balance
            XA.Value(lngIndex, 5) = objItem.ID
          '  XA.Value(lngIndex, 6) = objItem.Blocked
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    Grid.Array = XA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.LoadArray"
End Sub

Private Sub Grid_DblClick()
Dim lngID As Long
Dim bNotFound As Boolean

    If IsNull(Grid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Set oVendor = Nothing
    Set oVendor = New a_Supplier
    oVendor.Load FNN(XA(Grid.Bookmark, 5))
    strCustomerName = oVendor.Name & " (" & oVendor.AcNo & ")"
    lblPaymentGrid.Caption = strCustomerName
    
    LoadTransactions FNN(XA(Grid.Bookmark, 5))
    LoadLedger
    
    PaymentsGrid.ReBind
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadTransactions(Optional CustID As Long, Optional pAcno As String)
    On Error GoTo errHandler
    Set oTRs = Nothing
    Set oTRs = New c_CreditorsTransactionsPerTP
    oTRs.Load CustID, CDate("2000-01-01"), False, pAcno
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.LoadTransactions"
End Sub
Private Sub LoadRemittancesInPreparation(Optional DueDays As Long, Optional bSelected As Boolean, Optional lngTPID As Long)
    On Error GoTo errHandler
    Set oREMs = Nothing
    Set oREMs = New c_RemittancesInPreparation
    oREMs.Load DueDays, bSelected, lngTPID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.LoadTransactions"
End Sub


Private Sub LoadLedgerGrid(Optional pAcno As String)
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.LoadLedgerGrid"
End Sub



Private Sub lblPaymentGrid_Change()
    chkSelectedCreditorOnly.Enabled = (lblPaymentGrid.Caption > "")
End Sub

Private Sub PaymentsGrid_ButtonClick(ByVal ColIndex As Integer)
    If ColIndex <> 12 Then Exit Sub
    Select Case UCase(XC(PaymentsGrid.Bookmark, 13))
    Case "PAY"
        XC(PaymentsGrid.Bookmark, 11) = XC(PaymentsGrid.Bookmark, 17)
        XC(PaymentsGrid.Bookmark, 12) = "pay"
        XC(PaymentsGrid.Bookmark, 13) = "HOLD"
    Case "HOLD"
        XC(PaymentsGrid.Bookmark, 11) = ""
        XC(PaymentsGrid.Bookmark, 12) = ""
        XC(PaymentsGrid.Bookmark, 13) = "PAY"
    End Select
    PaymentsGrid.Refresh
    RefreshTotal
End Sub

Private Sub txtArg_GotFocus()
    AutoSelect txtArg
End Sub

Private Sub txtArg_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Trim(txtArg) = "" Or Trim(txtArg) = "*" Then Exit Sub
    
    If KeyAscii = 13 Then  ' The ENTER key.
       HandleResults
        If cVendors.Count > 1 Then
            On Error Resume Next
            Grid.SetFocus
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.txtArg_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
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
    ErrorIn "frmCreditors.SetDateArgs"
End Sub
Private Sub LedgerTab_Click(PreviousTab As Integer)
    resize
End Sub

Private Sub Form_Resize()
    resize
End Sub
Private Sub resize()
    Me.txtArg.Top = fr1.Top + 10
    Me.txtArg.Left = fr1.Left + 50
    
    Me.Grid.Left = fr1.Left + 50
    Me.Grid.Top = fr1.Top + 470
    Me.Grid.Width = NonNegative_Lng(fr1.Width - 400)
    Me.Grid.Height = NonNegative_Lng(fr1.Height - 1800)
    
    
    frBalances.Left = fr1.Left + 50
    frBalances.Top = NonNegative_Lng(fr1.Height - 1000)
    frBalances.Height = 900
    
 '   CM.MarginRight = 10
 '   CM.MarginLeft = 10
    CM.Size = 80
    
    Me.LedgerTab.Left = fr2.Left + 50
    Me.LedgerTab.Top = NonNegative_Lng(fr2.Top - fr1.Height - 150)
    Me.LedgerTab.Width = NonNegative_Lng(fr2.Width - 200)
    Me.LedgerTab.Height = NonNegative_Lng(fr2.Height - 600)
    arv.Left = 50
    arv.Width = NonNegative_Lng(fr3.Width - 100)
    arv.Height = fr3.Height - 1000
    
    If LedgerTab.Tab = 0 Then
        LedgerGrid.Visible = True
        cbSince.Top = NonNegative_Lng(fr2.Top - fr1.Height + 20)
        cbSince.Left = fr2.Left + 50
        cbSince.Width = NonNegative_Lng(Me.LedgerTab.Width - 150)
    
        LedgerGrid.Left = Me.LedgerTab.Left + 50
        LedgerGrid.Top = NonNegative_Lng(fr2.Top - fr1.Height + 500)
        LedgerGrid.Width = NonNegative_Lng(Me.LedgerTab.Width - 150)
        LedgerGrid.Height = NonNegative_Lng(Me.LedgerTab.Height - 950)
    ElseIf LedgerTab.Tab = 1 Then
        PaymentsGrid.Visible = True
        cmdPrepare.Left = Me.LedgerTab.Left + 3500
        PaymentsGrid.Left = Me.LedgerTab.Left + 50
        PaymentsGrid.Top = fr2.Top - fr1.Height + 330
        PaymentsGrid.Width = NonNegative_Lng(Me.LedgerTab.Width - 170)
        PaymentsGrid.Height = NonNegative_Lng(Me.LedgerTab.Height - 1250)
        cmdGeneratePayments.Top = NonNegative_Lng(Me.LedgerTab.Height - 450)
        cmdPrepare.Top = NonNegative_Lng(Me.PaymentsGrid.Top - 400)
        cmdPrintPaymentOrder.Top = cmdGeneratePayments.Top
        cmdPrintPaymentOrder.Left = 200
        Me.txtTotal.Top = Me.cmdGeneratePayments.Top
    Else
        Me.lblCreditorname.Caption = strCustomerName
        Me.PaymentsGrid.Visible = False
        Me.LedgerGrid.Visible = False
        ConfirmationGrid.Top = fr2.Top - fr1.Height + 650
        ConfirmationGrid.Left = Me.LedgerTab.Left + 50
        ConfirmationGrid.Width = NonNegative_Lng(Me.LedgerTab.Width - 170)
        ConfirmationGrid.Height = NonNegative_Lng(Me.LedgerTab.Height - 1250)
        cmdCreatePaymentRecords.Top = fr2.Top - fr1.Height + 200
        cmdCreatePaymentRecords.Left = LedgerTab.Left + 7050
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.Grid, Me.Name & Grid.Name
    SaveLayout Me.LedgerGrid, Me.Name & LedgerGrid.Name
    SaveLayout Me.PaymentsGrid, Me.Name & PaymentsGrid.Name
    SaveFormSize Me.Name, Me.Height, Me.Width
    SaveSplits Me.Name, Me.CM
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        LoadBrowse
    End If
End Sub


Private Sub HandleResults(Optional plngCount As Long)
    On Error GoTo errHandler
    If txtArg = "" Then Exit Sub
    Set cVendors = Nothing
    Set cVendors = New c_Supplier
    Screen.MousePointer = vbHourglass
    
    cVendors.LoadEasy Replace(txtArg, "'", "''"), False ', txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    plngCount = cVendors.Count
    LoadArray
    Grid.ReBind
    Grid.Enabled = True
    If cVendors.Count = 1 Then
        LoadTransactions FNN(XA(Grid.Bookmark, 5))
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.HandleResults(plngCount)", plngCount
End Sub

Private Sub LoadLedger(Optional pAcno As String)
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double
    
    If oTRs.Count = 0 Then
        XB.Clear
        LedgerGrid.ReBind
        Exit Sub
    End If
    
    For i = 1 To LedgerGrid.Columns.Count
        LedgerGrid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "LedgerGrid", CStr(i), LedgerGrid.Columns(i - 1).Width)
    Next
    
    XB.Clear
    LedgerGrid.ReBind
   ' Set oSQL = New z_SQL
   ' oSM.RecalculateTPBalance oVendor.ID
    i = 1
    j = 1
    Do While i <= oTRs.Count
        If oTRs.Item(i).DocType <> "BF" Then
            XB.ReDim 1, j, 1, 18
            XB.Value(j, 1) = oTRs.Item(i).DOCCode
            XB.Value(j, 2) = oTRs.Item(i).DocType
            XB.Value(j, 3) = oTRs.Item(i).CreditorDocDateF
            XB.Value(j, 4) = oTRs.Item(i).TRProcessingDateF
            XB.Value(j, 5) = oTRs.Item(i).DueDateF
            XB.Value(j, 6) = oTRs.Item(i).PayableAmountF
            XB.Value(j, 7) = oTRs.Item(i).PayableAfterSettDiscF
            XB.Value(j, 8) = oTRs.Item(i).SettlementDueDateF
            XB.Value(j, 9) = oTRs.Item(i).ClaimValueF
    '        XB.Value(j, 9) = oTRs.Item(i).CreditF
    '        XB.Value(j, 10) = oTRs.Item(i).Memo
'   '         XB.Value(j, 7) = oTRs.Item(i).DOCID
'            XB.Value(j, 7) = oTRs.Item(i).DOCCaptureDate
 '           dblBal = dblBal + oTRs.Item(i).Debit
 '           dblBal = dblBal - oTRs.Item(i).Credit
            XB.Value(j, 15) = oTRs.Item(i).TRID
            j = j + 1
        End If
        i = i + 1
    Loop
    XB.QuickSort 1, XB.UpperBound(1), 15, XORDER_DESCEND, XTYPE_LONG, 2, XORDER_DESCEND, XTYPE_STRING
    LedgerGrid.Array = XB
    LedgerGrid.ReBind
    LedgerGrid.Caption = oVendor.Name & " (" & oVendor.AcNo & ")"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.LoadLedger"
End Sub

Private Sub LoadRemittancesInPreparationGrid()
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double
    
    If oREMs.Count = 0 Then
        XC.Clear
        PaymentsGrid.ReBind
        Exit Sub
    End If
    
    For i = 1 To PaymentsGrid.Columns.Count
        PaymentsGrid.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "PaymentsGrid", CStr(i), PaymentsGrid.Columns(i - 1).Width)
    Next
    
    XC.Clear
    PaymentsGrid.ReBind
    i = 1
    j = 1
    Do While i <= oREMs.Count
        If oREMs.Item(i).DocType <> "BF" Then
            XC.ReDim 1, j, 1, 18
            XC.Value(j, 1) = oREMs.Item(i).SupplierInvoiceCode
            XC.Value(j, 2) = oREMs.Item(i).CreditorDocDateF
            XC.Value(j, 3) = oREMs.Item(i).DOCCode
            XC.Value(j, 4) = oREMs.Item(i).DocDateF
            XC.Value(j, 5) = oREMs.Item(i).DueDateF
            XC.Value(j, 6) = oREMs.Item(i).PayableAmountF
            
            XC.Value(j, 7) = oREMs.Item(i).EffectivePayableAfterSettDiscF
            XC.Value(j, 8) = oREMs.Item(i).SettlementDueDateF
            
            
            XC.Value(j, 9) = oREMs.Item(i).ClaimValueF
            XC.Value(j, 10) = CStr(oREMs.Item(i).OwingF)
            XC.Value(j, 11) = oREMs.Item(i).TempRemittanceF
            XC.Value(j, 12) = oREMs.Item(i).TempRemittanceStatus
            If UCase(XC.Value(j, 12)) = "" Then
                XC.Value(j, 13) = "PAY"
            ElseIf UCase(XC.Value(j, 12)) = "PAY" Then
                XC.Value(j, 13) = "HOLD"
            End If
            XC.Value(j, 15) = oREMs.Item(i).TRID
            XC.Value(j, 16) = oREMs.Item(i).REMID
            XC.Value(j, 17) = oREMs.Item(i).Owing
            j = j + 1
        End If
        i = i + 1
    Loop
    XC.QuickSort 1, XC.UpperBound(1), 15, XORDER_DESCEND, XTYPE_LONG, 2, XORDER_DESCEND, XTYPE_STRING
    PaymentsGrid.Array = XC
    PaymentsGrid.ReBind
 '   PaymentsGrid.Caption = oVendor.Name & " (" & oVendor.AcNo & ")"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreditors.LoadRemittancesInPreparationGrid"
End Sub

Private Sub RefreshTotal()
Dim i As Integer

    dblTotal = 0
    For i = 1 To XC.UpperBound(1)
        dblTotal = dblTotal + IIf(IsNumeric(StripToNumerics(XC.Value(i, 11))) = True, StripToNumerics(XC.Value(i, 11)), 0)
    Next i
    txtTotal = Format(dblTotal, "###,##0.00")
End Sub

Private Sub GeneratePaymentsOrder()
Dim i As Integer

    For i = 1 To XC.UpperBound(1)
        If FNS(XC.Value(i, 12)) > "" Then
            oSQL.UpdateRemittanceOrderStatus FNN(XC.Value(i, 16)), FNN(XC.Value(i, 15)), FNDBL(XC.Value(i, 11)), "", FNS(XC.Value(i, 12))
        End If
    Next

End Sub
