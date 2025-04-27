VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCustomerPreview 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Customer"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   9930
   Begin VB.TextBox txtContactPhone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   3765
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1725
   End
   Begin VB.TextBox txtContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   3765
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5970
      Width           =   1725
   End
   Begin VB.CheckBox chkBlock 
      BackColor       =   &H00F7EDE8&
      Height          =   375
      Left            =   8295
      TabIndex        =   21
      Top             =   510
      Width           =   315
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   75
      TabIndex        =   18
      Top             =   990
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14013889
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Terms"
      TabPicture(0)   =   "frmCustomerPreview.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtNotes"
      Tab(0).Control(1)=   "txtOurAcnoWithClient"
      Tab(0).Control(2)=   "chkSepInvs"
      Tab(0).Control(3)=   "txtParent"
      Tab(0).Control(4)=   "chkOneLinePerInvoice"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frVAT"
      Tab(0).Control(6)=   "chkUsesQuoted"
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(8)=   "txtTerms"
      Tab(0).Control(9)=   "txtCreditLimit"
      Tab(0).Control(10)=   "txtDefaultDiscount"
      Tab(0).Control(11)=   "chkCompleteOrder"
      Tab(0).Control(12)=   "Label7"
      Tab(0).Control(13)=   "Label24"
      Tab(0).Control(14)=   "Label39"
      Tab(0).Control(15)=   "Label38"
      Tab(0).Control(16)=   "Label34"
      Tab(0).Control(17)=   "Label26"
      Tab(0).Control(18)=   "Label15"
      Tab(0).Control(19)=   "Label13"
      Tab(0).Control(20)=   "Label10"
      Tab(0).Control(21)=   "Label41"
      Tab(0).Control(22)=   "lblRep"
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Addresses"
      TabPicture(1)   =   "frmCustomerPreview.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label27"
      Tab(1).Control(1)=   "G1"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "txtSAN"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Templates"
      TabPicture(2)   =   "frmCustomerPreview.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ledger"
      TabPicture(3)   =   "frmCustomerPreview.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblRecords"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label37"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "GD"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "DTPicker"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "frBalances"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdPrintList"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdMonthBack"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdNewPayment"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdAllocate"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Statement"
      TabPicture(4)   =   "frmCustomerPreview.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdLoadStatement"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdStatementPDF"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdStatementToExcel"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "GO"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "arStatementViewer"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "dtpStatement"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Label40"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      Begin VB.CommandButton cmdLoadStatement 
         BackColor       =   &H00F7EDE8&
         Caption         =   "&Load"
         Height          =   300
         Left            =   -69765
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   435
         Width           =   870
      End
      Begin VB.CommandButton cmdStatementPDF 
         BackColor       =   &H00F7EDE8&
         Caption         =   "PDF"
         Height          =   270
         Left            =   -67305
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   570
         Width           =   720
      End
      Begin VB.CommandButton cmdStatementToExcel 
         BackColor       =   &H00F7EDE8&
         Caption         =   "Excel"
         Height          =   270
         Left            =   -66555
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   570
         Width           =   720
      End
      Begin VB.CommandButton cmdAllocate 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Allocate"
         Height          =   300
         Left            =   615
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   4470
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.CommandButton cmdNewPayment 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&new payment"
         Height          =   315
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4455
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdMonthBack 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   5085
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   4425
         Width           =   870
      End
      Begin VB.CommandButton cmdPrintList 
         BackColor       =   &H00C4BCA4&
         Cancel          =   -1  'True
         Caption         =   "&Print list"
         Height          =   300
         Left            =   6795
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Frame frBalances 
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   1140
         TabIndex        =   77
         Top             =   480
         Width           =   7485
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
            TabIndex        =   89
            Top             =   750
            Visible         =   0   'False
            Width           =   1110
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
            TabIndex        =   88
            Top             =   750
            Visible         =   0   'False
            Width           =   960
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
            Top             =   750
            Visible         =   0   'False
            Width           =   990
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
            TabIndex        =   84
            Top             =   750
            Visible         =   0   'False
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
            TabIndex        =   83
            Top             =   360
            Width           =   1110
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
            TabIndex        =   82
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
            TabIndex        =   81
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
            TabIndex        =   80
            Top             =   360
            Width           =   990
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
            TabIndex        =   79
            Top             =   360
            Width           =   990
         End
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
            TabIndex        =   78
            Top             =   360
            Width           =   960
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
            TabIndex        =   97
            Top             =   780
            Visible         =   0   'False
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
            TabIndex        =   96
            Top             =   390
            Width           =   675
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   91
            Top             =   150
            Width           =   975
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
            TabIndex        =   90
            Top             =   150
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Templates"
         ForeColor       =   &H8000000D&
         Height          =   3870
         Left            =   -74775
         TabIndex        =   64
         Top             =   585
         Width           =   4305
         Begin VB.TextBox txtCreditNoteTemplatename 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   2112
            Width           =   2085
         End
         Begin VB.TextBox txtQuotationTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   660
            Width           =   2085
         End
         Begin VB.TextBox txtInvoiceTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   2475
            Width           =   2085
         End
         Begin VB.TextBox txtApproReturnTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1749
            Width           =   2085
         End
         Begin VB.TextBox txtApproTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   66
            Top             =   1386
            Width           =   2085
         End
         Begin VB.TextBox txtSalesOrderTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   65
            Top             =   1023
            Width           =   2085
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit notes"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   76
            Top             =   2127
            Width           =   1320
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Quotations"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   75
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Invoices"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   74
            Top             =   2490
            Width           =   1320
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Appro returns"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   73
            Top             =   1764
            Width           =   1320
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Appros"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   72
            Top             =   1401
            Width           =   1320
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Sales orders"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   71
            Top             =   1038
            Width           =   1320
         End
      End
      Begin VB.TextBox txtSAN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   -72525
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   360
         Width           =   2025
      End
      Begin VB.Frame Frame4 
         Caption         =   "Document delivery method"
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
         Height          =   780
         Left            =   -72285
         TabIndex        =   57
         Top             =   3930
         Width           =   3870
         Begin VB.OptionButton optFaxManual 
            Caption         =   "Print and then fax"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   165
            TabIndex        =   60
            Top             =   285
            Value           =   -1  'True
            Width           =   1860
         End
         Begin VB.OptionButton optEmail 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   2925
            TabIndex        =   59
            Top             =   285
            Width           =   825
         End
         Begin VB.OptionButton optEDI 
            Caption         =   "E.D.I."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   1890
            TabIndex        =   58
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   705
         Left            =   -74850
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   3960
         Width           =   5565
      End
      Begin VB.TextBox txtOurAcnoWithClient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   -71175
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text1"
         Top             =   3375
         Width           =   1620
      End
      Begin VB.CheckBox chkSepInvs 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -71010
         TabIndex        =   44
         Top             =   2715
         Width           =   315
      End
      Begin VB.TextBox txtParent 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73275
         TabIndex        =   43
         Top             =   735
         Width           =   3555
      End
      Begin VB.CheckBox chkOneLinePerInvoice 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71010
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2055
         Width           =   315
      End
      Begin VB.Frame frVAT 
         Caption         =   "V.A.T."
         ForeColor       =   &H8000000D&
         Height          =   1665
         Left            =   -68940
         TabIndex        =   38
         Top             =   645
         Width           =   2565
         Begin VB.CheckBox chkVATable 
            Caption         =   "Pays V.A.T"
            Enabled         =   0   'False
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   285
            TabIndex        =   41
            Top             =   225
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtVATNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   300
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   690
            Width           =   1860
         End
         Begin VB.CheckBox chkShowVAT 
            Caption         =   "Show VAT deducted if not VATable."
            Enabled         =   0   'False
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   300
            TabIndex        =   39
            Top             =   1095
            Value           =   1  'Checked
            Width           =   2160
         End
      End
      Begin VB.CheckBox chkUsesQuoted 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -71010
         TabIndex        =   37
         Top             =   2385
         Width           =   315
      End
      Begin VB.Frame Frame3 
         Caption         =   "Accounting"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1485
         Left            =   -68940
         TabIndex        =   33
         Top             =   2565
         Width           =   2565
         Begin VB.OptionButton optOI 
            Caption         =   "Open item"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   405
            Left            =   285
            TabIndex        =   35
            Top             =   1635
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtAccountingACCNUM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   345
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Accounting package A/c no."
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   -45
            TabIndex        =   36
            Top             =   270
            Width           =   2370
         End
      End
      Begin VB.TextBox txtTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   1680
         Width           =   1875
      End
      Begin VB.TextBox txtCreditLimit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1365
         Width           =   1875
      End
      Begin VB.TextBox txtDefaultDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   -73290
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   1050
         Width           =   1875
      End
      Begin VB.CheckBox chkCompleteOrder 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -71010
         TabIndex        =   29
         Top             =   3045
         Width           =   315
      End
      Begin TrueOleDBGrid60.TDBGrid GO 
         Height          =   2535
         Left            =   -64920
         OleObjectBlob   =   "frmCustomerPreview.frx":008C
         TabIndex        =   27
         Top             =   1320
         Width           =   8445
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         DragIcon        =   "frmCustomerPreview.frx":5673
         Height          =   3045
         Left            =   -74325
         OleObjectBlob   =   "frmCustomerPreview.frx":5AB5
         TabIndex        =   62
         Top             =   810
         Width           =   8295
      End
      Begin MSComCtl2.DTPicker DTPicker 
         Height          =   315
         Left            =   3630
         TabIndex        =   102
         Top             =   4440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70844417
         CurrentDate     =   39980
      End
      Begin TrueOleDBGrid60.TDBGrid GD 
         Height          =   2985
         Left            =   570
         OleObjectBlob   =   "frmCustomerPreview.frx":8EC8
         TabIndex        =   103
         Top             =   1395
         Width           =   8445
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arStatementViewer 
         Height          =   3855
         Left            =   -74655
         TabIndex        =   107
         Top             =   840
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6800
         SectionData     =   "frmCustomerPreview.frx":DB17
      End
      Begin MSComCtl2.DTPicker dtpStatement 
         Height          =   315
         Left            =   -71190
         TabIndex        =   109
         Top             =   435
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70844417
         CurrentDate     =   39980
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -71910
         TabIndex        =   110
         Top             =   465
         Width           =   600
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2880
         TabIndex        =   104
         Top             =   4125
         Width           =   600
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "SAN number (for EDI)"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74160
         TabIndex        =   63
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   -74850
         TabIndex        =   56
         Top             =   3735
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Our vendor no. with client"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73350
         TabIndex        =   54
         Top             =   3450
         Width           =   2145
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Generate separate invoices for separate orders"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -74790
         TabIndex        =   52
         Top             =   2775
         Width           =   3510
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Uses quoted price on invoice"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -74790
         TabIndex        =   51
         Top             =   2430
         Width           =   3510
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent customer"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74595
         TabIndex        =   50
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "One line per invoice"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -74790
         TabIndex        =   49
         Top             =   2115
         Width           =   3510
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74865
         TabIndex        =   48
         Top             =   1725
         Width           =   1380
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Credit limit"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74865
         TabIndex        =   47
         Top             =   1410
         Width           =   1380
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Default discount"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74865
         TabIndex        =   46
         Top             =   1110
         Width           =   1380
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Only invoice when order is complete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -74790
         TabIndex        =   45
         Top             =   3120
         Width           =   3510
      End
      Begin VB.Label lblRep 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   -74805
         TabIndex        =   20
         Top             =   4050
         Width           =   3270
      End
      Begin VB.Label lblRecords 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   4170
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00F7EDE8&
      Caption         =   "&Delete"
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
      Left            =   6645
      Picture         =   "frmCustomerPreview.frx":DB53
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5910
      Width           =   1000
   End
   Begin VB.PictureBox picNoGO 
      Height          =   420
      Left            =   1245
      Picture         =   "frmCustomerPreview.frx":DEDD
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   16
      Top             =   -120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picOver 
      Height          =   420
      Left            =   1365
      Picture         =   "frmCustomerPreview.frx":E31F
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   15
      Top             =   -165
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox PicDrop 
      Height          =   420
      Left            =   675
      Picture         =   "frmCustomerPreview.frx":E761
      ScaleHeight     =   360
      ScaleWidth      =   450
      TabIndex        =   14
      Top             =   -105
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtRecordLastChanged 
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
      Height          =   270
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6330
      Width           =   1125
   End
   Begin VB.TextBox txtRecordAdded 
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
      Height          =   270
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5970
      Width           =   1140
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00F7EDE8&
      Caption         =   "&Close"
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
      Left            =   8685
      Picture         =   "frmCustomerPreview.frx":EBA3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5910
      Width           =   1000
   End
   Begin VB.TextBox txtInitials 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   180
      Width           =   1020
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5475
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   180
      Width           =   585
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00F7EDE8&
      Caption         =   "&Edit"
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
      Left            =   7665
      Picture         =   "frmCustomerPreview.frx":EF2D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5910
      Width           =   1000
   End
   Begin VB.TextBox txtAcno 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6105
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   1020
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1005
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   3825
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1005
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   3825
   End
   Begin VB.Label lblTemporary 
      BackStyle       =   0  'Transparent
      Caption         =   "* Temporary *"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   7680
      TabIndex        =   28
      Top             =   165
      Width           =   1305
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2970
      TabIndex        =   26
      Top             =   6360
      Width           =   750
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2970
      TabIndex        =   24
      Top             =   6000
      Width           =   750
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Blocked"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7245
      TabIndex        =   22
      Top             =   570
      Width           =   915
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   3750
      X2              =   1275
      Y1              =   15
      Y2              =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record last changed: "
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
      Left            =   -75
      TabIndex        =   13
      Top             =   6330
      Width           =   1800
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record added: "
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
      Left            =   405
      TabIndex        =   12
      Top             =   5985
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A/c. Num."
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5190
      TabIndex        =   5
      Top             =   630
      Width           =   930
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   345
      TabIndex        =   4
      Top             =   240
      Width           =   585
   End
End
Attribute VB_Name = "frmCustomerPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCust As a_Customer
Dim frmCP As frmCustomer
Dim XA As New XArrayDB
Dim XB As New XArrayDB
Dim XO As New XArrayDB
Dim vRowBookmark As Variant
Dim oTRs As c_DebtorsTransPerTP
Dim bBusiness As Boolean
Dim lngMonthsBack As Long
Dim arStatement As arStatement_b

Public Sub component(pCust As a_Customer)
    On Error GoTo errHandler
    Set oCust = pCust
    Me.Caption = "Customer: " & oCust.Name
    bBusiness = (oCust.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID)
    frVAT.Visible = bBusiness
    txtTitle.Visible = Not bBusiness
    txtInitials.Visible = Not bBusiness
    ExpandCaption
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.component(pCust)", pCust
End Sub
Public Sub Component2(PID As Long)
    On Error GoTo errHandler

    Set oCust = New a_Customer
    oCust.Load PID
    Me.Caption = "Customer master preview: " & oCust.Name
    bBusiness = (oCust.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID)
    frVAT.Visible = bBusiness
    txtTitle.Visible = Not bBusiness
    txtInitials.Visible = Not bBusiness
    ExpandCaption
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.Component2(PID)", PID
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
 '   SaveLayout Me.G1, Me.Name & "G1"
    SaveLayout Me.GD, Me.Name & "GD"
    SaveLayout Me.GO, Me.Name & "GO"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.mnuSaveLayout"
End Sub

Private Sub ExpandCaption()
    On Error GoTo errHandler
    If oCust.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID Then
        Me.Caption = Me.Caption & " (business)"
    ElseIf oCust.CustomerTypeID = oPC.Configuration.BookClubCustomerTypeID Then
        Me.Caption = Me.Caption & " (book club)"
    ElseIf oCust.CustomerTypeID = oPC.Configuration.PrivateCustomerTypeID Then
        Me.Caption = Me.Caption & " (private)"
    End If
    If oCust.CanBeDeleted Then
        Me.Caption = Me.Caption & " (temporary)"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.ExpandCaption"
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdDelete_Click()
'    On Error GoTo errHandler
'Dim XA As XArrayDB
'Dim XB As XArrayDB
''Dim frm As frmTPActivity
'Dim frm1 As frmTPOldDocs
'Dim oDPTP As c_DocsPerTP
'Dim lngResult As Long
'Dim oSM As z_StockManager
'
'    If MsgBox("Confirm you want to delete " & oCust.Fullname, vbQuestion + vbYesNo, "Confirm") = vbNo Then
'        Exit Sub
'    End If
'    Set XA = New XArrayDB
'    Set XB = New XArrayDB
'    If oCust.OKForDeletion(XA, XB, oDPTP) Then
'        If XA.UpperBound(1) > 0 Then
'            Set frm1 = New frmTPOldDocs
'            frm1.ComponentXA XA, oCust.Fullname, "There are documents belonging to this customer, but they are dated prior to the last stock take and will be deleted if the customer is deleted."
'            frm1.Show vbModal
'            If Not frm1.ToDelete Then
'                Unload frm
'                Exit Sub
'            End If
'            Unload frm
'        End If
'        Set oSM = New z_StockManager
'        oSM.DeleteUnusedPTs
'        oCust.BeginEdit
'        oCust.DeleteCustomer
'        oCust.ApplyEdit lngResult
'        MsgBox "Customer deleted! Form will close.", vbInformation, "Action complete"
'        Set oSM = Nothing
'        Unload Me
'    Else
'        MsgBox "There are associated documents which may not be deleted yet. You cannot delete this customer." & vbCrLf & "Use the 'Related documents button to see details.", , "Can't delete"
'     '   Set frm = New frmTPActivity
'     '   frm.Component oDPTP, oCust.Fullname
'     '   frm.Show vbModal
'     '   Unload frm
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCustomerPreview.cmdDelete_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean

    If frmCP Is Nothing Then
        Set frmCP = New frmCustomer
    End If
    blnEdit = True
    oCust.BeginEdit
    frmCP.component oCust ', lngID
    frmCP.Show
    
EXIT_Handler:
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdLoadStatement_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL

 '   If SecurityControl(enSECURITY_ISSUPERVISOR, , "Loading statements", "You need to have supervisor status to load statements.") = False Then Exit Sub

    Set rs = New ADODB.Recordset
    oSQL.RunGetRecordset "SELECT * FROM vOpenItemAll_1 WHERE TPID = " & CStr(oCust.ID) & " AND (BALANCE <> 0 OR dbDocDate > '" & ReverseDate(Me.dtpStatement) & "') ORDER BY AGE,dbDocDate,crDocDate ", enText, "", "", rs
    'arStatement.Run
    Set arStatement = New arStatement_b
    arStatementViewer.ReportSource = arStatement
    arStatement.component rs, oPC.Configuration.DefaultCompany, oCust.NameAndCode(100), oPC.Configuration.DefaultCompany.StreetAddress, oPC.Configuration.DefaultCompany.BankDetails, oPC.Configuration.DefaultCompany.VatNumber
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdLoadStatement_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdMonthBack_Click()
    On Error GoTo errHandler
    LoadDebtorsStatement
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdMonthBack_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdNewPayment_Click()
'    On Error GoTo errHandler
'Dim frm As frmCustPmt
'Dim oSQL As New z_SQL
'    Set frm = New frmCustPmt
'    frm.Component2 oCust.ID, oCust.NameAndCode(40), XO.Value(GO.Bookmark, 9)
'    frm.Show
'    Screen.MousePointer = vbHourglass
'    oSQL.RunProc "MatchPaymentsAuto", Array(oCust.ID), ""
'    LoadDebtorsStatement
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCustomerPreview.cmdNewPayment_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cmdPrintStatement_Click()
    On Error GoTo errHandler
Dim oFSO As New FileSystemObject
Dim oXML As zXML
Dim sFN As String
Dim sPDF As String
Dim f As frmStatementControl
Dim oSQL As New z_SQL
    Set f = New frmStatementControl
    f.component True
    f.Show vbModal


    Screen.MousePointer = vbHourglass
    sFN = "STATEMENT-" & oCust.AcNo & "-" & Format(Date, "YYYYMMDD") & ".XML"
    oSQL.RunProc "StatementperCust_XML", Array(oCust.ID, oPC.SharedFolderRoot & "\TEMP\" & sFN), ""


    sPDF = "STATEMENT-" & oCust.AcNo & "-" & Format(Date, "YYYYMMDD") & ".PDF"

    If oFSO.FileExists(oPC.SharedFolderRoot & "\Statements\" & sFN) Then
        oFSO.DeleteFile (oPC.SharedFolderRoot & "\Statements\" & sFN)
    End If
    
    Set oXML = New zXML
    oXML.PrintXML oPC.SharedFolderRoot & "\TEMP\" & sFN, oPC.SharedFolderRoot & "\TEMP", _
                    oPC.SharedFolderRoot & "\Templates\", _
                    oPC.LocalFolder & "\Executables", _
                    False, _
                    oPC.SharedFolderRoot & "\TEMP\" & sPDF
    Set oXML = Nothing
    
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdPrintStatement_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdPrintList_Click()
    On Error GoTo errHandler
    GD.PrintInfo.PrintPreview
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdPrintList_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStatementToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    arStatement.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "ST" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fs)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(arStatement.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdStatementToExcel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStatementPDF_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    arStatement.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "ST" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fs)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(arStatement.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdStatementPDF_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Top = 50
        Left = 50
        Width = 10000
        Height = 7245
    End If
    LoadControls
    If oPC.RunsAccountsTF Then
        Me.SSTab1.Tab = 1
        LoadDebtorsStatement
    Else
        Me.SSTab1.Tab = 0
    End If
    Select Case oCust.DispatchMethod
    Case "E"
        optEDI = True
    Case "M"
        optEmail = True
    Case "P"
        optFaxManual = True
    End Select
    frBalances.Visible = True
    GD.Visible = True
    GO.Visible = False
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Public Sub LoadControls()
    On Error GoTo errHandler
Dim i As Integer

    Me.DTPicker.Value = DateAdd("m", -6, Date)
    txtName = oCust.Name
    txtPhone = oCust.Phone
    txtTitle = oCust.Title
    txtParent = oCust.ParentCustomerName
    txtInitials = oCust.Initials
 '   txtPaysVAT = IIf(oCust.VATable, "YES", "NO")
    txtVATNumber = oCust.VatNumber
  '  txtCustomerType = oCust.CustomerTypesALL_tl.Item(oCust.CustomerTypeID)
    txtRecordAdded = oCust.DateRecordAddedF
    txtRecordLastChanged = oCust.DateRecordLastChangedF
    txtAcno = oCust.AcNo
    txtSAN = oCust.SAN
    txtDefaultDiscount = oCust.DefaultDiscountF & " discount"
    txtNotes = oCust.Note
    txtOurAcnoWithClient = oCust.OurACnoWithClient
    txtAccountingACCNUM = oCust.AccAcno
    txtContact = oCust.ContactPerson
    txtContactPhone = oCust.ContactpersonPhoneF
    
    txtSalesOrderTemplateName = oCust.SalesOrderTemplateName
    txtApproTemplateName = oCust.ApproTemplateName
    txtApproReturnTemplateName = oCust.ApproReturnTemplateName
    txtQuotationTemplateName = oCust.QuotationTemplateName
    txtInvoiceTemplateName = oCust.InvoiceTemplateName
    txtCreditNoteTemplatename = oCust.CreditNoteTemplateName
    
    lblTemporary.Visible = oCust.CanBeDeleted
    lblRep.Caption = "Sales rep: " & IIf(oCust.RepName > "", oCust.RepName, "<NONE>")
    Me.txtCreditLimit = oCust.CreditLimitF
    Me.txtTerms = oCust.TermsF
    Me.txt120PlusBal = oCust.Balance120F
    Me.txt30Bal = oCust.Balance30F
    Me.txt60Bal = oCust.Balance60F
    Me.txt90Bal = oCust.Balance90F
    Me.txtCurBal = oCust.BalanceCurF
    Me.txtBalance = oCust.BalanceF
    Me.chkBlock = IIf(oCust.Blocked = True, 1, 0)
    Me.chkUsesQuoted = IIf(oCust.UseQuotedPrice = True, 1, 0)
 '   cmdAlerts.Visible = Not oPC.IncludeSupplierFeatures
    Me.chkVATable = IIf(oCust.VATable, 1, 0)
    Me.chkShowVAT = IIf(oCust.ShowVAT, 1, 0)
    Me.chkSepInvs = IIf(oCust.GenerateSeparateInvoicesForSeparateOrders, 1, 0)
    Me.chkOneLinePerInvoice = IIf(oCust.OneLinePerInvoice = True, 1, 0)
    Me.chkCompleteOrder = IIf(oCust.CompleteOrder = True, 1, 0)
    LoadArray
'    LoadTPIGs
'    LoadTPCCs
   ' Me.SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.LoadControls"
End Sub

'Private Sub RefreshBalances()
'    oCust.RefreshBalances
'    txt120PlusBal = oCust.Balance120F
'    txt30Bal = oCust.Balance30F
'    txt60Bal = oCust.Balance60F
'    txt90Bal = oCust.Balance90F
'    txtCurBal = oCust.BalanceCurF
'    txtBalance = oCust.BalanceF
'End Sub

'Private Sub LoadTPIGs()
'    On Error GoTo errHandler
'Dim oTPIG As a_IG
'    With Me.lbIG
'        .Clear
'        For Each oTPIG In oCust.InterestGroups
'            .AddItem oTPIG.Description
'        Next
'        If .ListCount > 0 Then .ListIndex = 0
'    End With
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCustomerPreview.LoadTPIGs"
'End Sub
'Private Sub LoadTPCCs()
'    On Error GoTo errHandler
'Dim oTPCC As a_IG
'    With Me.lbCC
'        .Clear
'        For Each oTPCC In oCust.CustomerTypes
'            .AddItem oTPCC.Description   ', oTPIG.Key
'        Next
'        If .ListCount > 0 Then .ListIndex = 0
'    End With
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCustomerPreview.LoadTPCCs"
'End Sub


Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    SSTab1.Width = NonNegative_Lng(Me.Width - 370)
    lngDiff = SSTab1.Height
    SSTab1.Height = NonNegative_Lng(Me.Height - 2400)
    lngDiff = (SSTab1.Height - lngDiff)
    arStatementViewer.Height = NonNegative_Lng(Me.Height - 3000)
    arStatementViewer.Width = NonNegative_Lng(Me.Width - 1000)
    GD.Height = NonNegative_Lng(Me.Height - 4100)
    cmdDelete.Top = cmdDelete.Top + lngDiff
    cmdEdit.Top = cmdEdit.Top + lngDiff
    cmdClose.Top = cmdClose.Top + lngDiff
    cmdPrintList.Top = cmdPrintList.Top + lngDiff
    cmdMonthBack.Top = cmdMonthBack.Top + lngDiff
    cmdNewPayment.Top = cmdNewPayment.Top + lngDiff
    cmdAllocate.Top = cmdAllocate.Top + lngDiff
  '  cmdAllPurchases.Top = cmdAllPurchases.Top + lngDiff
    Label37.Top = Label37.Top + lngDiff
    DTPicker.Top = DTPicker.Top + lngDiff
 '   cmdMatchPayments.Top = cmdMatchPayments.Top + lngDiff
    Label8.Top = Label8.Top + lngDiff
    Label5.Top = Label5.Top + lngDiff
    Label35.Top = Label35.Top + lngDiff
    Label36.Top = Label36.Top + lngDiff
    txtRecordAdded.Top = txtRecordAdded.Top + lngDiff
    txtRecordLastChanged.Top = txtRecordLastChanged.Top + lngDiff
    txtContact.Top = txtContact.Top + lngDiff
    txtContactPhone.Top = txtContactPhone.Top + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

'Private Sub G1_DblClick()
'    On Error GoTo errHandler
'Dim frm As frmAddressPreview
'Dim lngID As Long
'    If IsNull(G1.Bookmark) Then Exit Sub
'    Set frm = New frmAddressPreview
'    lngID = Val(XA(G1.Bookmark, 5))
'    frm.component oCust.Addresses.Item(lngID)
'    frm.Show vbModal
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCustomerPreview.G1_DblClick", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub LoadArray()
    On Error GoTo errHandler
Dim objItem As d_Customer
Dim lngIndex As Long
Dim i As Integer

    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "G1", CStr(i), G1.Columns(i - 1).Width)
    Next

    XA.ReDim 1, oCust.Addresses.Count, 1, 6
    For lngIndex = 1 To oCust.Addresses.Count
        With objItem
            XA.Value(lngIndex, 1) = lngIndex
            XA.Value(lngIndex, 2) = oCust.Addresses(lngIndex).AddressMailing
            XA.Value(lngIndex, 3) = CreateRoleString(oCust.Addresses(lngIndex))
            XA.Value(lngIndex, 4) = oCust.Addresses(lngIndex).GetsCatalogue
            XA.Value(lngIndex, 5) = oCust.Addresses(lngIndex).key
            XA.Value(lngIndex, 6) = oCust.Addresses(lngIndex).ForMailing
        End With
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    If XA.UpperBound(1) > 1 Then
        Me.lblRecords = XA.UpperBound(1) & " addresses"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.LoadArray"
End Sub

Private Function CreateRoleString(pAddress As a_Address) As String
    On Error GoTo errHandler
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & ",", "")
    str = str & IIf(pAddress.DelTo = True, "Del" & ",", "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & ",", "")
    str = str & IIf(pAddress.Appro = True, "Appro" & ",", "")
    CreateRoleString = str
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.CreateRoleString(pAddress)", pAddress
End Function
Private Function CreateRoleString2(pAddress As a_Address) As String
    On Error GoTo errHandler
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & vbCrLf, "")
    str = str & IIf(pAddress.DelTo = True, "Del" & vbCrLf, "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & vbCrLf, "")
    str = str & IIf(pAddress.Appro = True, "Appro" & vbCrLf, "")
    CreateRoleString2 = str
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.CreateRoleString2(pAddress)", pAddress
End Function


Private Sub G1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
' If the button is up and we get MouseMove, that means
' we exited the form and tried to drop elsewhere.
' Reset the drag upon returning.
    If Button = 0 Then ResetDragDrop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.G1_MouseMove(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If Bookmark < 1 Then Exit Sub
    If XA(Bookmark, 6) = True Then
        RowStyle.BackColor = RGB(282, 274, 180)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub ResetDragDrop()
    On Error GoTo errHandler
' Turn off drag-and-drop by resetting the highlight and data
' control caption.
    If G1.MarqueeStyle = dbgSolidCellBorder Then Exit Sub
    G1.MarqueeStyle = dbgSolidCellBorder
    G1.MarqueeStyle = dbgSolidCellBorder
'    SB1.SimpleText = "Drag an address"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.ResetDragDrop"
End Sub
Private Sub G1_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
    On Error GoTo errHandler
' Set the current cell to the one being dragged
    G1.Col = ColIndex
    G1.Bookmark = RowBookmark
    vRowBookmark = RowBookmark
    ' Set up drag operation, such as creating visual effects by
    ' highlighting the cell or row being dragged.
            ' Highlight the phone number cell to indicate data
            ' from the cell is being dragged.
            G1.MarqueeStyle = dbgHighlightRow
'            SB1.SimpleText = "Dragging an address . . ."
    ' Use VB manual drag support (put TDBGrid1 into drag mode)
    G1.Drag vbBeginDrag
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.G1_DragCell(SplitIndex,RowBookmark,ColIndex)", Array(SplitIndex, _
         RowBookmark, ColIndex), EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
    On Error GoTo errHandler
' DragOver provides different visual feedback as we are
' dragging a row, or just the phone number.

    Dim dragFrom As String
    Dim overCol As Integer
    Dim overRow As Long
    
    
    Select Case State
        Case vbEnter
            G1.MarqueeStyle = dbgHighlightRow
            G1.DragIcon = picOver.Picture
        Case vbLeave
            G1.MarqueeStyle = dbgHighlightRow
            G1.DragIcon = picNoGO.Picture
        Case vbOver
            overRow = G1.RowContaining(Y)
            Debug.Print overRow
            If overRow >= 0 Then G1.Row = overRow
'            If vRowBookmark = overRow Then
'                G1.DragIcon = picOver.Picture
'            Else
'                G1.DragIcon = PicDrop.Picture
'            End If
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.G1_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub G1_DragDrop(Source As Control, x As Single, Y As Single)
    On Error GoTo errHandler
    Dim overRow As Long
        MsgBox "Merging address no: " & vRowBookmark & " Into: " & G1.Bookmark
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.G1_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_DblClick()
    On Error GoTo errHandler
    If Not IsNull(oCust.BillTOAddress) Then
        On Error Resume Next
    
        Clipboard.Clear
        Clipboard.SetText oCust.BillTOAddress.AddressMailing
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.Form_DblClick", , EA_NORERAISE
    HandleError
End Sub

'Private Sub PrepareColumns()
'    tvwTR.Nodes.Clear
'    tvwTR.LevelDefs.Clear
'    tvwTR.NodeDragDrop = 0
'    tvwTR.Nodes.Clear
'    tvwTR.MultiSelect = 0
'    tvwTR.VertSpacing = 0
'    tvwTR.LevelDefs.Add "Root"    '0
'    tvwTR.LevelDefs.Add "Branch1" '1
'    tvwTR.LevelDefs(0).Font.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs.Add , , "Code"
'    tvwTR.LevelDefs(1).ColumnDefs.Add , , "Date"
'    tvwTR.LevelDefs(1).ColumnDefs.Add , , "Debit"
'    tvwTR.LevelDefs(1).ColumnDefs.Add , , "Credit"
'    tvwTR.LevelDefs(1).ColumnDefs.Add , , "Balance"
'    tvwTR.LevelDefs(1).ColumnDefs.Add , , ""
'    tvwTR.LevelDefs(1).Font.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(0).CaptionBackColor = tvwTR.DefBackColor   'RGB(255, 255, 255)
'    tvwTR.LevelDefs(1).ColumnDefs(0).CaptionFont3D = 0
'    tvwTR.LevelDefs(1).ColumnDefs(0).CaptionBorderStyle = 4 'gtBorderStyleSingle
'    tvwTR.LevelDefs(1).ColumnDefs(0).CaptionFont.Bold = False
'    tvwTR.LevelDefs(1).ColumnDefs(0).CaptionFont.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(0).Width = 1100
'    tvwTR.LevelDefs(1).ColumnDefs(1).CaptionBackColor = tvwTR.DefBackColor
'    tvwTR.LevelDefs(1).ColumnDefs(1).CaptionFont3D = 0
'    tvwTR.LevelDefs(1).ColumnDefs(1).CaptionBorderStyle = 4 'gtBorderStyleSingle
'    tvwTR.LevelDefs(1).ColumnDefs(1).CaptionFont.Bold = False
'    tvwTR.LevelDefs(1).ColumnDefs(1).CaptionFont.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(1).Width = 1100
'    tvwTR.LevelDefs(1).ColumnDefs(2).CaptionBackColor = tvwTR.DefBackColor
'    tvwTR.LevelDefs(1).ColumnDefs(2).CaptionFont3D = 0
'    tvwTR.LevelDefs(1).ColumnDefs(2).CaptionBorderStyle = 4 'gtBorderStyleSingle
'    tvwTR.LevelDefs(1).ColumnDefs(2).CaptionFont.Bold = False
'    tvwTR.LevelDefs(1).ColumnDefs(2).CaptionFont.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(2).Width = 1100
'    tvwTR.LevelDefs(1).ColumnDefs(3).CaptionBackColor = tvwTR.DefBackColor
'    tvwTR.LevelDefs(1).ColumnDefs(3).CaptionFont3D = 0
'    tvwTR.LevelDefs(1).ColumnDefs(3).CaptionBorderStyle = 4 'gtBorderStyleSingle
'    tvwTR.LevelDefs(1).ColumnDefs(3).CaptionFont.Bold = False
'    tvwTR.LevelDefs(1).ColumnDefs(3).CaptionFont.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(3).Width = 1000
'    tvwTR.LevelDefs(1).ColumnDefs(4).CaptionBackColor = tvwTR.DefBackColor
'    tvwTR.LevelDefs(1).ColumnDefs(4).CaptionFont3D = 0
'    tvwTR.LevelDefs(1).ColumnDefs(4).CaptionBorderStyle = 4 'gtBorderStyleSingle
'    tvwTR.LevelDefs(1).ColumnDefs(4).CaptionFont.Bold = False
'    tvwTR.LevelDefs(1).ColumnDefs(4).CaptionFont.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(4).Width = 1100
'    tvwTR.LevelDefs(1).ColumnDefs(5).CaptionBackColor = tvwTR.DefBackColor
'    tvwTR.LevelDefs(1).ColumnDefs(5).CaptionFont3D = 0
'    tvwTR.LevelDefs(1).ColumnDefs(5).CaptionBorderStyle = 4 'gtBorderStyleSingle
'    tvwTR.LevelDefs(1).ColumnDefs(5).CaptionFont.Bold = False
'    tvwTR.LevelDefs(1).ColumnDefs(5).CaptionFont.Size = 9
'    tvwTR.LevelDefs(1).ColumnDefs(5).Width = 1100
'End Sub
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
    ErrorIn "frmCustomerPreview.LoadDebtorsStatement"
End Sub
Private Sub LoadTransactions()
    On Error GoTo errHandler
    Set oTRs = Nothing
    Set oTRs = New c_DebtorsTransPerTP
    oTRs.Load oCust.ID, Me.DTPicker.Value
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.LoadTransactions"
End Sub




Private Sub GD_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
      
'    If Button = 2 Then   ' Check if right mouse button
'        PopupMenu Forms(0).mnuCustomerPreviewPopup   ' Display the File menu as a
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.GD_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    If SSTab1.Tab = 1 Then
        LoadDebtorsStatement
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.SSTab1_Click(PreviousTab)", PreviousTab, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadLedger()
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double

    If oTRs.Count = 0 Then
        XB.Clear
        GD.ReBind
        Exit Sub
    End If
    
    For i = 1 To GD.Columns.Count
        GD.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "GD", CStr(i), GD.Columns(i - 1).Width)
    Next
    
    XB.Clear
    GD.ReBind
    
    Me.txtCurBal = oCust.BalanceCurF
    Me.txt30Bal = oCust.Balance30F
    Me.txt60Bal = oCust.Balance60F
    Me.txt90Bal = oCust.Balance90F
    Me.txt120PlusBal = oCust.Balance120F
    Me.txtBalance = oCust.BalanceF
    
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
            XB.Value(j, 6) = oTRs.Item(i).MEMO
            XB.Value(j, 7) = oTRs.Item(i).DOCID
            XB.Value(j, 7) = oTRs.Item(i).DOCCaptureDate
 '           dblBal = dblBal + oTRs.Item(i).Debit
 '           dblBal = dblBal - oTRs.Item(i).Credit
            j = j + 1
        End If
        i = i + 1
    Loop
    XB.QuickSort 1, XB.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
    GD.Array = XB
    GD.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.LoadLedger"
End Sub
Private Sub LoadLedgerOpenItem()
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
Dim dblBal As Double

    If oTRs.Count = 0 Then Exit Sub
    For i = 1 To GO.Columns.Count
        GO.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "GO", CStr(i), GO.Columns(i - 1).Width)
    Next
    
    XO.Clear
    GO.ReBind
    j = 1
    Do While j <= oTRs.Count
            XO.ReDim 1, j, 1, 10
            XO.Value(j, 1) = oTRs.Item(j).dbDoc
            XO.Value(j, 2) = oTRs.Item(j).dbDate
            XO.Value(j, 3) = oTRs.Item(j).dbDocType
            XO.Value(j, 4) = oTRs.Item(j).dbAmt
            XO.Value(j, 5) = oTRs.Item(j).crDoc
            XO.Value(j, 6) = oTRs.Item(j).crDate
            XO.Value(j, 7) = oTRs.Item(j).crDocType
            XO.Value(j, 8) = oTRs.Item(j).crAmt
            XO.Value(j, 9) = oTRs.Item(j).DOCID
            j = j + 1
    Loop
    GO.Array = XO
    GO.ReBind
    GO.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.LoadLedgerOpenItem"
End Sub

'Public Sub mnuPayInvoice()
'    On Error GoTo errHandler
'Dim T As New XArrayDB
'Dim f As New frmCustPmt
'Dim i As Integer
'
'    For i = 1 To GD.SelBookmarks.Count
'        T.ReDim 0, i, 1, 2
'        T(i, 1) = XB.Value(GD.SelBookmarks(i - 1), 1)
'        T(i, 2) = XB.Value(GD.SelBookmarks(i - 1), 7)
'    Next i
'
'
'        f.component oCust.ID, oCust.NameAndCode(50), T
'        f.Show 'vbModal
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCustomerPreview.mnuPayInvoice"
'End Sub


