VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmCustomerPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
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
      TabIndex        =   42
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
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5970
      Width           =   1725
   End
   Begin VB.CommandButton cmdHistory 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&History"
      Height          =   420
      Left            =   7875
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   60
      Width           =   840
   End
   Begin VB.CheckBox chkBlock 
      BackColor       =   &H00D3D3CB&
      Height          =   375
      Left            =   8295
      TabIndex        =   30
      Top             =   510
      Width           =   315
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   90
      TabIndex        =   18
      Top             =   1035
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14013889
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Classification"
      TabPicture(0)   =   "frmCustomerPreview.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtOurAcnoWithClient"
      Tab(0).Control(1)=   "frVAT"
      Tab(0).Control(2)=   "txtCustomerType"
      Tab(0).Control(3)=   "lbIG"
      Tab(0).Control(4)=   "lbCC"
      Tab(0).Control(5)=   "txtNotes"
      Tab(0).Control(6)=   "Label24"
      Tab(0).Control(7)=   "lblRep"
      Tab(0).Control(8)=   "lblTemporary"
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(12)=   "Label7"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Terms"
      TabPicture(1)   =   "frmCustomerPreview.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label39"
      Tab(1).Control(1)=   "Label38"
      Tab(1).Control(2)=   "Label34"
      Tab(1).Control(3)=   "Label26"
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(5)=   "Label13"
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(7)=   "chkSepInvs"
      Tab(1).Control(8)=   "txtParent"
      Tab(1).Control(9)=   "chkOneLinePerInvoice"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(11)=   "chkUsesQuoted"
      Tab(1).Control(12)=   "Frame3"
      Tab(1).Control(13)=   "txtTerms"
      Tab(1).Control(14)=   "txtCreditLimit"
      Tab(1).Control(15)=   "txtDefaultDiscount"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Addresses"
      TabPicture(2)   =   "frmCustomerPreview.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label27"
      Tab(2).Control(1)=   "G1"
      Tab(2).Control(2)=   "txtSAN"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame4"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Templates"
      TabPicture(3)   =   "frmCustomerPreview.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblRecords"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Transactions"
      TabPicture(4)   =   "frmCustomerPreview.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label37"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "GO"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "GD"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "DTPicker1"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdAllocate"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdNewPayment"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdMonthBack"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdMatchPayments"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Command1"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "cmdAllPurchases"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "cmdAllOrders"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "cmdCurrentOrders"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "cmdPrintList"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "frBalances"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "Statement"
      TabPicture(5)   =   "frmCustomerPreview.frx":008C
      Tab(5).ControlEnabled=   0   'False
<<<<<<< .working
      Tab(5).Control(0)=   "cmdStatementToExcel"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdStatementPDF"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "arStatementViewer"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "cmdLoadStatement"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "dtpStatement"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label40"
      Tab(5).Control(5).Enabled=   0   'False
=======
      Tab(5).Control(0)=   "Label40"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "dtpStatement"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdLoadStatement"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "arStatementViewer"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cmdStatementPDF"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cmdStatementToExcel"
>>>>>>> .merge-right.r247
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).ControlCount=   6
      Begin VB.CommandButton cmdStatementToExcel 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Excel"
         Height          =   270
         Left            =   -66540
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   525
         Width           =   720
      End
      Begin VB.CommandButton cmdStatementPDF 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   270
         Left            =   -67290
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   540
         Width           =   720
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arStatementViewer 
         Height          =   3855
         Left            =   -74640
         TabIndex        =   124
         Top             =   810
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6800
         SectionData     =   "frmCustomerPreview.frx":00A8
      End
      Begin VB.CommandButton cmdLoadStatement 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Load"
         Height          =   300
         Left            =   -69735
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   405
         Width           =   870
      End
      Begin VB.Frame Frame5 
         Caption         =   "Templates"
         ForeColor       =   &H8000000D&
         Height          =   3870
         Left            =   -74790
         TabIndex        =   107
         Top             =   405
         Width           =   6315
         Begin VB.TextBox txtSalesOrderTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   113
            Top             =   1023
            Width           =   2085
         End
         Begin VB.TextBox txtApproTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   112
            Top             =   1386
            Width           =   2085
         End
         Begin VB.TextBox txtApproReturnTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   111
            Top             =   1749
            Width           =   2085
         End
         Begin VB.TextBox txtInvoiceTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   110
            Top             =   2475
            Width           =   2085
         End
         Begin VB.TextBox txtQuotationTemplateName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   109
            Top             =   660
            Width           =   2085
         End
         Begin VB.TextBox txtCreditNoteTemplatename 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   300
            Left            =   1695
            Locked          =   -1  'True
            TabIndex        =   108
            Top             =   2112
            Width           =   2085
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            Caption         =   "Sales orders"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   119
            Top             =   1038
            Width           =   1320
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "Appros"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   118
            Top             =   1401
            Width           =   1320
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Appro returns"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   117
            Top             =   1764
            Width           =   1320
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            Caption         =   "Invoices"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   116
            Top             =   2490
            Width           =   1320
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Quotations"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   115
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit notes"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   114
            Top             =   2127
            Width           =   1320
         End
      End
      Begin VB.Frame frBalances 
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   210
         TabIndex        =   84
         Top             =   405
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
            TabIndex        =   96
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
            TabIndex        =   95
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
            TabIndex        =   94
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
            TabIndex        =   93
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
            TabIndex        =   92
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
            TabIndex        =   91
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   88
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
            TabIndex        =   87
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
            TabIndex        =   86
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
            TabIndex        =   85
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
            TabIndex        =   104
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
            TabIndex        =   103
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
            TabIndex        =   102
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
            TabIndex        =   101
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
            TabIndex        =   100
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
            TabIndex        =   99
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
            TabIndex        =   98
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
            TabIndex        =   97
            Top             =   780
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdPrintList 
         BackColor       =   &H00C4BCA4&
         Cancel          =   -1  'True
         Caption         =   "&Print list"
         Height          =   300
         Left            =   6435
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   4425
         Width           =   1455
      End
      Begin VB.CommandButton cmdCurrentOrders 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Current orders"
         Enabled         =   0   'False
         Height          =   450
         Left            =   8715
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   495
         Width           =   870
      End
      Begin VB.CommandButton cmdAllOrders 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&All orders"
         Enabled         =   0   'False
         Height          =   450
         Left            =   8715
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1335
         Width           =   870
      End
      Begin VB.CommandButton cmdAllPurchases 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&All purchases"
         Height          =   450
         Left            =   8715
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2190
         Width           =   870
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&All returns"
         Enabled         =   0   'False
         Height          =   450
         Left            =   8715
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3060
         Width           =   870
      End
      Begin VB.CommandButton cmdMatchPayments 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Match payments"
         Enabled         =   0   'False
         Height          =   450
         Left            =   8655
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   4320
         Width           =   870
      End
      Begin VB.CommandButton cmdMonthBack 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   4410
         Width           =   870
      End
      Begin VB.CommandButton cmdNewPayment 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&new payment"
         Height          =   315
         Left            =   1245
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   4440
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton cmdAllocate 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Allocate"
         Height          =   300
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   4455
         Visible         =   0   'False
         Width           =   930
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
         Left            =   -74715
         TabIndex        =   68
         Top             =   4035
         Width           =   3870
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
            TabIndex        =   71
            Top             =   285
            Width           =   855
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
            TabIndex        =   70
            Top             =   285
            Width           =   1320
         End
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
            TabIndex        =   69
            Top             =   285
            Value           =   -1  'True
            Width           =   1860
         End
      End
      Begin VB.TextBox txtSAN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   -73020
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   495
         Width           =   2025
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
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   1125
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
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   1665
         Width           =   1875
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
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   2235
         Width           =   1875
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
         Left            =   -69480
         TabIndex        =   52
         Top             =   2190
         Width           =   2775
         Begin VB.TextBox txtAccountingACCNUM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   600
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   840
            Width           =   1785
         End
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
            TabIndex        =   54
            Top             =   1635
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.OptionButton optBF 
            Caption         =   "Balance forward"
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
            TabIndex        =   53
            Top             =   1305
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Accounting package A/c no."
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   210
            TabIndex        =   56
            Top             =   480
            Width           =   2370
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
         Left            =   -71865
         TabIndex        =   51
         Top             =   3105
         Width           =   285
      End
      Begin VB.Frame Frame1 
         Caption         =   "V.A.T."
         ForeColor       =   &H8000000D&
         Height          =   1665
         Left            =   -69495
         TabIndex        =   47
         Top             =   450
         Width           =   2565
         Begin VB.CheckBox chkShowVAT 
            Caption         =   "Show VAT deducted if not VATable."
            Enabled         =   0   'False
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   300
            TabIndex        =   50
            Top             =   1095
            Value           =   1  'Checked
            Width           =   2160
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   300
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   690
            Width           =   1860
         End
         Begin VB.CheckBox chkVATable 
            Caption         =   "Pays V.A.T"
            Enabled         =   0   'False
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   285
            TabIndex        =   48
            Top             =   225
            Value           =   1  'Checked
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkOneLinePerInvoice 
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71775
         TabIndex        =   46
         TabStop         =   0   'False
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
         Left            =   -73335
         TabIndex        =   45
         Top             =   585
         Visible         =   0   'False
         Width           =   2190
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
         Left            =   -71895
         TabIndex        =   44
         Top             =   3525
         Width           =   315
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
         Left            =   -68625
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   3705
         Width           =   1875
      End
      Begin VB.Frame frVAT 
         Caption         =   "V.A.T."
         ForeColor       =   &H8000000D&
         Height          =   1875
         Left            =   -71160
         TabIndex        =   32
         Top             =   1230
         Visible         =   0   'False
         Width           =   2505
         Begin VB.TextBox txtVATNumber 
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
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1320
            Width           =   2025
         End
         Begin VB.TextBox txtPaysVAT 
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
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   600
            Width           =   900
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "VAT number"
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
            Left            =   210
            TabIndex        =   36
            Top             =   1050
            Width           =   1515
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Pays VAT"
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
            Left            =   180
            TabIndex        =   35
            Top             =   330
            Width           =   855
         End
      End
      Begin VB.TextBox txtCustomerType 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   -74820
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   630
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.ListBox lbIG 
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
         Height          =   615
         Left            =   -74820
         TabIndex        =   24
         Top             =   2430
         Width           =   2970
      End
      Begin VB.ListBox lbCC 
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
         Height          =   615
         Left            =   -74820
         TabIndex        =   22
         Top             =   1320
         Width           =   2970
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
         Left            =   -74820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   3555
         Width           =   5565
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         DragIcon        =   "frmCustomerPreview.frx":00E4
         Height          =   3045
         Left            =   -74700
         OleObjectBlob   =   "frmCustomerPreview.frx":0526
         TabIndex        =   72
         Top             =   855
         Width           =   8295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3285
         TabIndex        =   75
         Top             =   4425
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
<<<<<<< .working
         Format          =   85131265
=======
         Format          =   62193665
>>>>>>> .merge-right.r247
         CurrentDate     =   39980
      End
      Begin TrueOleDBGrid60.TDBGrid GD 
         Height          =   2925
         Left            =   195
         OleObjectBlob   =   "frmCustomerPreview.frx":3939
         TabIndex        =   105
         Top             =   1380
         Width           =   8445
      End
      Begin TrueOleDBGrid60.TDBGrid GO 
         Height          =   2535
         Left            =   10080
         OleObjectBlob   =   "frmCustomerPreview.frx":8588
         TabIndex        =   106
         Top             =   1620
         Width           =   8445
      End
      Begin MSComCtl2.DTPicker dtpStatement 
         Height          =   315
         Left            =   -71175
         TabIndex        =   122
         Top             =   405
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
<<<<<<< .working
         Format          =   85131265
=======
         Format          =   62193665
>>>>>>> .merge-right.r247
         CurrentDate     =   39980
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -71835
         TabIndex        =   123
         Top             =   450
         Width           =   600
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2625
         TabIndex        =   120
         Top             =   4440
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
         Left            =   -74670
         TabIndex        =   73
         Top             =   525
         Width           =   1695
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
         Left            =   -74925
         TabIndex        =   66
         Top             =   1185
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
         Left            =   -74925
         TabIndex        =   65
         Top             =   1710
         Width           =   1380
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
         Left            =   -74925
         TabIndex        =   64
         Top             =   2280
         Width           =   1380
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
         Height          =   255
         Left            =   -73710
         TabIndex        =   63
         Top             =   2805
         Width           =   1500
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
         Left            =   -74655
         TabIndex        =   62
         Top             =   630
         Visible         =   0   'False
         Width           =   1275
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
         Height          =   255
         Left            =   -74145
         TabIndex        =   61
         Top             =   3195
         Width           =   1980
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
         Left            =   -74580
         TabIndex        =   60
         Top             =   3525
         Width           =   2385
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
         Left            =   -68745
         TabIndex        =   38
         Top             =   3420
         Width           =   2145
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
         Left            =   -74820
         TabIndex        =   29
         Top             =   4350
         Width           =   3270
      End
      Begin VB.Label lblTemporary 
         BackStyle       =   0  'Transparent
         Caption         =   "* Temporary *"
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
         Left            =   -71700
         TabIndex        =   28
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer type"
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
         Left            =   -74820
         TabIndex        =   27
         Top             =   390
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Interest groups"
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
         Left            =   -74820
         TabIndex        =   25
         Top             =   2160
         Width           =   1380
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer classification"
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
         Left            =   -74820
         TabIndex        =   23
         Top             =   1050
         Width           =   2295
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
         Left            =   -74820
         TabIndex        =   21
         Top             =   3330
         Width           =   2295
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
         Left            =   -74520
         TabIndex        =   19
         Top             =   4470
         Width           =   2475
      End
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
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
      Left            =   6660
      Picture         =   "frmCustomerPreview.frx":DB6F
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5910
      Width           =   1000
   End
   Begin VB.PictureBox picNoGO 
      Height          =   420
      Left            =   1245
      Picture         =   "frmCustomerPreview.frx":DEF9
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
      Picture         =   "frmCustomerPreview.frx":E33B
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
      Picture         =   "frmCustomerPreview.frx":E77D
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
      BackColor       =   &H00C4BCA4&
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
      Picture         =   "frmCustomerPreview.frx":EBBF
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
      BackColor       =   &H00C4BCA4&
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
      Picture         =   "frmCustomerPreview.frx":EF49
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
      TabIndex        =   43
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
      TabIndex        =   41
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
      TabIndex        =   31
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

Public Sub Component(pCust As a_Customer)
    Set oCust = pCust
    Me.Caption = "Customer: " & oCust.Name
    bBusiness = (oCust.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID)
    frVAT.Visible = bBusiness
    txtTitle.Visible = Not bBusiness
    txtInitials.Visible = Not bBusiness
    ExpandCaption
    Me.cmdMatchPayments.Enabled = True
End Sub
Public Sub Component2(pID As Long)

    Set oCust = New a_Customer
    oCust.Load pID
    Me.Caption = "Customer master preview: " & oCust.Name
    bBusiness = (oCust.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID)
    frVAT.Visible = bBusiness
    txtTitle.Visible = Not bBusiness
    txtInitials.Visible = Not bBusiness
    ExpandCaption
    Me.cmdMatchPayments.Enabled = True
End Sub
Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name & "G1"
    SaveLayout Me.GD, Me.Name & "GD"
    SaveLayout Me.GO, Me.Name & "GO"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuSaveLayout"
End Sub

Private Sub ExpandCaption()
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
End Sub

Private Sub cmdAllocate_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    Screen.MousePointer = vbHourglass
    oSQL.RunProc "MatchPaymentsAuto", Array(oCust.ID), ""
    LoadDebtorsStatement
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdAllocate_Click"
End Sub

Private Sub cmdAllPurchases_Click()
Dim frm As frmCustPurch
Dim oCP As c_SalesPerCustomer

    Set frm = New frmCustPurch
    Set oCP = New c_SalesPerCustomer
    oCP.Load oCust.ID
    frm.Component oCP, oCust.FullName
    frm.Show vbModal
    Set oCP = Nothing

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim frm As frmTPActivity
Dim frm1 As frmTPOldDocs
Dim oDPTP As c_DocsPerTP
Dim lngResult As Long
Dim oSM As z_StockManager

    If MsgBox("Confirm you want to delete " & oCust.FullName, vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Set XA = New XArrayDB
    Set XB = New XArrayDB
    If oCust.OKForDeletion(XA, XB, oDPTP) Then
        If XA.UpperBound(1) > 0 Then
            Set frm1 = New frmTPOldDocs
            frm1.ComponentXA XA, oCust.FullName, "There are documents belonging to this customer, but they are dated prior to the last stock take and will be deleted if the customer is deleted."
            frm1.Show vbModal
            If Not frm1.ToDelete Then
                Unload frm
                Exit Sub
            End If
            Unload frm
        End If
        Set oSM = New z_StockManager
        oSM.DeleteUnusedPTs
        oCust.BeginEdit
        oCust.DeleteCustomer
        oCust.ApplyEdit lngResult
        MsgBox "Customer deleted! Form will close.", vbInformation, "Action complete"
        Set oSM = Nothing
        Unload Me
    Else
        MsgBox "There are associated documents which may not be deleted yet. You cannot delete this customer." & vbCrLf & "Use the 'Related documents button to see details.", , "Can't delete"
     '   Set frm = New frmTPActivity
     '   frm.Component oDPTP, oCust.Fullname
     '   frm.Show vbModal
     '   Unload frm
    End If
End Sub

Private Sub cmdEdit_Click()
Dim blnEdit As Boolean

    On Error GoTo ERR_Handler
    If frmCP Is Nothing Then
        Set frmCP = New frmCustomer
    End If
    blnEdit = True
    oCust.BeginEdit
    frmCP.Component oCust ', lngID
    frmCP.Show
    
EXIT_Handler:
    Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub

'Private Sub cmdMergeAddresses_Click()
'    fMerge.Visible = True
'End Sub



Private Sub cmdHistory_Click()
Dim f As New frmHistory
    
    f.Component oCust.ID
    f.Show
    
End Sub

Private Sub cmdLoadStatement_Click()
Dim rs As ADODB.Recordset
Dim oSQL As New z_SQL

    If SecurityControl(enSECURITY_ISSUPERVISOR, , "Loading statements", "You need to have supervisor status to load statements.") = False Then Exit Sub

    Set rs = New ADODB.Recordset
    oSQL.RunGetRecordset "SELECT * FROM vOpenItemAll_1 WHERE TPID = " & CStr(oCust.ID) & " AND (BALANCE <> 0 OR dbDocDate > '" & ReverseDate(Me.dtpStatement) & "') ORDER BY AGE,dbDocDate,crDocDate ", enText, "", "", rs
    'arStatement.Run
    Set arStatement = New arStatement_b
    arStatementViewer.ReportSource = arStatement
    arStatement.Component rs, oPC.Configuration.DefaultCompany, oCust.NameAndCode(100), oPC.Configuration.DefaultCompany.StreetAddress, oPC.Configuration.DefaultCompany.BankDetails, oPC.Configuration.DefaultCompany.VATNumber
End Sub

Private Sub cmdMatchPayments_Click()
    On Error GoTo errHandler
Dim frm As New frmPaymentMatch

    frm.Component oCust.ID, oCust.NameAndCode(50)
    frm.Show vbModal
    Screen.MousePointer = vbHourglass
    oCust.Reload
    LoadDebtorsStatement
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPreview.cmdMatchPayments_Click"
End Sub

Private Sub cmdMonthBack_Click()
    LoadDebtorsStatement
End Sub

Private Sub cmdNewPayment_Click()
Dim frm As frmCustPmt
Dim oSQL As New z_SQL
    Set frm = New frmCustPmt
    frm.Component2 oCust.ID, oCust.NameAndCode(40), XO.Value(GO.Bookmark, 9)
    frm.Show
    Screen.MousePointer = vbHourglass
    oSQL.RunProc "MatchPaymentsAuto", Array(oCust.ID), ""
    LoadDebtorsStatement
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintStatement_Click()
Dim oFSO As New FileSystemObject
Dim oXML As zXML
Dim sFN As String
Dim sPDF As String
Dim f As frmStatementControl
Dim oSQL As New z_SQL
    Set f = New frmStatementControl
    f.Component True
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

End Sub



Private Sub cmdPrintList_Click()
    GD.PrintInfo.PrintPreview
End Sub

Private Sub cmdStatementToExcel_Click()
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
    OpenFileWithApplication fn
End Sub

Private Sub cmdStatementPDF_Click()
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
    OpenFileWithApplication fn
End Sub

Private Sub Form_Load()
    If Me.WindowState <> 2 Then
        top = 50
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
    
End Sub


Public Sub LoadControls()
Dim i As Integer

    Me.DTPicker1.Value = DateAdd("m", -6, Date)
    txtName = oCust.Name
    txtPhone = oCust.Phone
    txtTitle = oCust.Title
    txtParent = oCust.ParentCustomerName
    txtInitials = oCust.Initials
    txtPaysVAT = IIf(oCust.VATable, "YES", "NO")
    txtVATNumber = oCust.VATNumber
    txtCustomerType = oCust.CustomerTypesALL_tl.Item(oCust.CustomerTypeID)
    txtRecordAdded = oCust.DateRecordAddedF
    txtRecordLastChanged = oCust.DateRecordLastChangedF
    txtAcno = oCust.AcNo
    txtSAN = oCust.SAN
    txtDefaultDiscount = oCust.DefaultDiscountF & " discount"
    txtNotes = oCust.Note
    txtOurAcnoWithClient = oCust.ouracnowithclient
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
    lblRep.Caption = "Sales rep: " & IIf(oCust.Repname > "", oCust.Repname, "<NONE>")
    Me.txtCreditLimit = oCust.CreditLimitF
    Me.txtTerms = oCust.TermsF
    Me.txt120PlusBal = oCust.Balance120F
    Me.txt30Bal = oCust.Balance30F
    Me.txt60Bal = oCust.Balance60F
    Me.txt90Bal = oCust.Balance90F
    Me.txtCurBal = oCust.BalanceCurF
    Me.txtBalance = oCust.BalanceF
    Me.chkBlock = IIf(oCust.bLocked = True, 1, 0)
    Me.chkOneLinePerInvoice = IIf(oCust.OneLinePerInvoice = True, 1, 0)
    Me.chkUsesQuoted = IIf(oCust.UseQuotedPrice = True, 1, 0)
'    If oCust.PaymentStyle = "B" Then
'        optBF = True
'    ElseIf oCust.PaymentStyle = "O" Then
'        optOI = True
'    End If
    Me.chkVATable = IIf(oCust.VATable, 1, 0)
    Me.chkShowVAT = IIf(oCust.ShowVAT, 1, 0)
    Me.chkSepInvs = IIf(oCust.GenerateSeparateInvoicesForSeparateOrders, 1, 0)
    Me.chkOneLinePerInvoice = IIf(oCust.OneLinePerInvoice, 1, 0)
    LoadArray
    LoadTPIGs
    LoadTPCCs
   ' Me.SSTab1.Tab = 0
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

Private Sub LoadTPIGs()
Dim oTPIG As a_IG
    With Me.lbIG
        .Clear
        For Each oTPIG In oCust.InterestGroups
            .AddItem oTPIG.Description
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub
Private Sub LoadTPCCs()
    On Error GoTo errHandler
Dim oTPCC As a_IG
    With Me.lbCC
        .Clear
        For Each oTPCC In oCust.CustomerTypes
            .AddItem oTPCC.Description   ', oTPIG.Key
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLoyaltyPreview.LoadTPCCs"
End Sub


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    SSTab1.Width = Me.Width - 370
    lngDiff = SSTab1.Height
    SSTab1.Height = Me.Height - 2400
    lngDiff = SSTab1.Height - lngDiff
    arStatementViewer.Height = Me.Height - 3000
    arStatementViewer.Width = Me.Width - 1000
    GD.Height = Me.Height - 4100
    cmdDelete.top = cmdDelete.top + lngDiff
    cmdEdit.top = cmdEdit.top + lngDiff
    cmdClose.top = cmdClose.top + lngDiff
    cmdPrintList.top = cmdPrintList.top + lngDiff
    cmdMonthBack.top = cmdMonthBack.top + lngDiff
    cmdNewPayment.top = cmdNewPayment.top + lngDiff
    cmdAllocate.top = cmdAllocate.top + lngDiff
    cmdAllPurchases.top = cmdAllPurchases.top + lngDiff
    Label37.top = Label37.top + lngDiff
    DTPicker1.top = DTPicker1.top + lngDiff
    cmdMatchPayments.top = cmdMatchPayments.top + lngDiff
    Label8.top = Label8.top + lngDiff
    Label5.top = Label5.top + lngDiff
    Label35.top = Label35.top + lngDiff
    Label36.top = Label36.top + lngDiff
    txtRecordAdded.top = txtRecordAdded.top + lngDiff
    txtRecordLastChanged.top = txtRecordLastChanged.top + lngDiff
    txtContact.top = txtContact.top + lngDiff
    txtContactPhone.top = txtContactPhone.top + lngDiff

End Sub

Private Sub G1_DblClick()
Dim frm As frmAddressPreview
Dim lngID As Long
    If IsNull(G1.Bookmark) Then Exit Sub
    Set frm = New frmAddressPreview
    lngID = val(XA(G1.Bookmark, 5))
    frm.Component oCust.Addresses.Item(lngID)
    frm.Show vbModal
End Sub

Private Sub LoadArray()
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
End Sub

Private Function CreateRoleString(pAddress As a_Address) As String
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & ",", "")
    str = str & IIf(pAddress.DelTo = True, "Del" & ",", "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & ",", "")
    str = str & IIf(pAddress.Appro = True, "Appro" & ",", "")
    CreateRoleString = str
End Function
Private Function CreateRoleString2(pAddress As a_Address) As String
Dim str As String
    str = ""
    str = str & IIf(pAddress.BillTo = True, "Bill" & vbCrLf, "")
    str = str & IIf(pAddress.DelTo = True, "Del" & vbCrLf, "")
    str = str & IIf(pAddress.OrderTo = True, "Order" & vbCrLf, "")
    str = str & IIf(pAddress.Appro = True, "Appro" & vbCrLf, "")
    CreateRoleString2 = str
End Function


Private Sub G1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
' If the button is up and we get MouseMove, that means
' we exited the form and tried to drop elsewhere.
' Reset the drag upon returning.
    If Button = 0 Then ResetDragDrop
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Bookmark < 1 Then Exit Sub
    If XA(Bookmark, 6) = True Then
        RowStyle.BackColor = RGB(282, 274, 180)
    End If
End Sub
Private Sub ResetDragDrop()
' Turn off drag-and-drop by resetting the highlight and data
' control caption.
    If G1.MarqueeStyle = dbgSolidCellBorder Then Exit Sub
    G1.MarqueeStyle = dbgSolidCellBorder
    G1.MarqueeStyle = dbgSolidCellBorder
'    SB1.SimpleText = "Drag an address"
End Sub
Private Sub G1_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
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
End Sub
Private Sub G1_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
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
End Sub

Private Sub G1_DragDrop(Source As Control, x As Single, Y As Single)
    Dim overRow As Long
        MsgBox "Merging address no: " & vRowBookmark & " Into: " & G1.Bookmark
End Sub

Private Sub Form_DblClick()
    If Not IsNull(oCust.billtoaddress) Then
    On Error Resume Next
    
        Clipboard.Clear
        Clipboard.SetText oCust.billtoaddress.AddressMailing
    End If
    
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
Dim oSQL As New z_SQL

    oSQL.RunProc "[AgeInvoices]", Array(oCust.ID), ""
    oCust.Reload
    LoadTransactions
    LoadStatement
    
End Sub
Private Sub LoadTransactions()
    Set oTRs = Nothing
    Set oTRs = New c_DebtorsTransPerTP
    oTRs.Load oCust.ID, Me.DTPicker1.Value
End Sub




Private Sub GD_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
      
'    If Button = 2 Then   ' Check if right mouse button
'        PopupMenu Forms(0).mnuCustomerPreviewPopup   ' Display the File menu as a
'    End If
    Exit Sub
errHandler:
    ErrorIn "frmCustomerPreview.GD_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
        LoadDebtorsStatement
    End If
End Sub

Private Sub LoadStatement()
Dim i As Long
Dim j As Long
Dim dblBal As Double

    If oTRs.Count = 0 Then Exit Sub
    
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
        If oTRs.Item(i).doctype <> "BF" Then
            XB.ReDim 1, j, 1, 8
            XB.Value(j, 1) = oTRs.Item(i).DocCode
            XB.Value(j, 2) = oTRs.Item(i).doctype
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
    XB.QuickSort 1, XB.UpperBound(1), 3, XORDER_DESCEND, XTYPE_DATE
    GD.Array = XB
    GD.ReBind
End Sub
Private Sub LoadStatementOpenItem()
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
End Sub

Public Sub mnuPayInvoice()
Dim T As New XArrayDB
Dim f As New frmCustPmt
Dim i As Integer

    For i = 1 To GD.SelBookmarks.Count
        T.ReDim 0, i, 1, 2
        T(i, 1) = XB.Value(GD.SelBookmarks(i - 1), 1)
        T(i, 2) = XB.Value(GD.SelBookmarks(i - 1), 7)
    Next i
    
    
        f.Component oCust.ID, oCust.NameAndCode(50), T
        f.Show 'vbModal
End Sub


