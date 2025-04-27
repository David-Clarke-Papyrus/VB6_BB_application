VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmCustomer 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   9150
   Begin VB.ComboBox cboSalesRep 
      Height          =   315
      Left            =   1095
      Style           =   2  'Dropdown List
      TabIndex        =   87
      Top             =   5055
      Width           =   2910
   End
   Begin VB.TextBox txtContactPhone 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1110
      MaxLength       =   25
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   5730
      Width           =   1725
   End
   Begin VB.TextBox txtContact 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1110
      MaxLength       =   25
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1725
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4140
      Left            =   195
      TabIndex        =   25
      Top             =   870
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   7303
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   13882315
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Classification"
      TabPicture(0)   =   "frmCustomer2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "frCustomerClassification"
      Tab(0).Control(3)=   "frInterestGroup"
      Tab(0).Control(4)=   "txtNote"
      Tab(0).Control(5)=   "chkTemporary"
      Tab(0).Control(6)=   "txtOurAcnoWithClient"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Terms"
      TabPicture(1)   =   "frmCustomer2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtDefaultDiscount"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cboTerms"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtCreditLimit"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdUnlockPrices"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkUseQuotedPrice"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "frVAT"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkBlock"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkOneLinePerInvoice"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdKeep"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtParent"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkSepInvs"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "chkCompleteOrder"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Addresses"
      TabPicture(2)   =   "frmCustomer2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblRecords"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "cbOrderTo"
      Tab(2).Control(3)=   "cbDelTo"
      Tab(2).Control(4)=   "cbAppro"
      Tab(2).Control(5)=   "cbBillTo"
      Tab(2).Control(6)=   "G1"
      Tab(2).Control(7)=   "cmdAdd"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdRemove"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdEdit"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Frame5"
      Tab(2).Control(11)=   "txtSAN"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Templates"
      TabPicture(3)   =   "frmCustomer2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox chkCompleteOrder 
         Alignment       =   1  'Right Justify
         Caption         =   "Only invoice when order is complete"
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
         Left            =   1140
         TabIndex        =   86
         Top             =   3345
         Width           =   3135
      End
      Begin VB.CheckBox chkSepInvs 
         Alignment       =   1  'Right Justify
         Caption         =   "Generate separate invoices for separate orders"
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
         Left            =   390
         TabIndex        =   85
         Top             =   2985
         Width           =   3885
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
         Left            =   2895
         TabIndex        =   79
         Top             =   705
         Width           =   2280
      End
      Begin VB.CommandButton cmdKeep 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5190
         Picture         =   "frmCustomer2.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   690
         Width           =   360
      End
      Begin VB.Frame Frame2 
         Caption         =   "Templates"
         ForeColor       =   &H8000000D&
         Height          =   3015
         Left            =   -74730
         TabIndex        =   65
         Top             =   540
         Width           =   6315
         Begin VB.TextBox txtCreditNoteTemplatename 
            Height          =   300
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   76
            Top             =   2112
            Width           =   2085
         End
         Begin VB.TextBox txtQuotationTemplateName 
            Height          =   300
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   74
            Top             =   660
            Width           =   2085
         End
         Begin VB.TextBox txtInvoiceTemplateName 
            Height          =   300
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   72
            Top             =   2460
            Width           =   2085
         End
         Begin VB.TextBox txtApproReturnTemplateName 
            Height          =   300
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   70
            Top             =   1749
            Width           =   2085
         End
         Begin VB.TextBox txtApproTemplateName 
            Height          =   300
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   68
            Top             =   1386
            Width           =   2085
         End
         Begin VB.TextBox txtSalesOrderTemplateName 
            Height          =   300
            Left            =   1695
            MaxLength       =   12
            TabIndex        =   66
            Top             =   1023
            Width           =   2085
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit notes"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   77
            Top             =   2127
            Width           =   1320
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Quotations"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   75
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Invoices"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   73
            Top             =   2490
            Width           =   1320
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Appro returns"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   71
            Top             =   1764
            Width           =   1320
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Appros"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   69
            Top             =   1401
            Width           =   1320
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Sales orders"
            ForeColor       =   &H8000000D&
            Height          =   270
            Left            =   225
            TabIndex        =   67
            Top             =   1038
            Width           =   1320
         End
      End
      Begin VB.TextBox txtSAN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   -73215
         MaxLength       =   25
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   405
         Width           =   2025
      End
      Begin VB.CheckBox chkOneLinePerInvoice 
         Height          =   375
         Left            =   4065
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2295
         Width           =   315
      End
      Begin VB.CheckBox chkBlock 
         Height          =   375
         Left            =   1695
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   2310
         Width           =   315
      End
      Begin VB.TextBox txtOurAcnoWithClient 
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
         Left            =   -68880
         TabIndex        =   53
         Top             =   3045
         Width           =   1860
      End
      Begin VB.Frame frVAT 
         Caption         =   "V.A.T."
         ForeColor       =   &H8000000D&
         Height          =   1755
         Left            =   5850
         TabIndex        =   50
         Top             =   510
         Width           =   2565
         Begin VB.CheckBox chkShowVAT 
            Caption         =   "Show VAT deducted if not VATable."
            Enabled         =   0   'False
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   300
            TabIndex        =   60
            Top             =   1215
            Value           =   1  'Checked
            Width           =   2160
         End
         Begin VB.CheckBox chkVATable 
            Caption         =   "Pays V.A.T"
            ForeColor       =   &H8000000D&
            Height          =   450
            Left            =   285
            TabIndex        =   52
            Top             =   225
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtVATNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   300
            TabIndex        =   51
            Top             =   690
            Width           =   1860
         End
      End
      Begin VB.CheckBox chkUseQuotedPrice 
         Alignment       =   1  'Right Justify
         Caption         =   "Use quoted price on invoices"
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
         Left            =   1620
         TabIndex        =   12
         Top             =   2655
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         Caption         =   "Document delivery method"
         ForeColor       =   &H8000000D&
         Height          =   765
         Left            =   -74865
         TabIndex        =   49
         Top             =   3240
         Width           =   7095
         Begin VB.PictureBox Picture 
            Height          =   375
            Left            =   195
            ScaleHeight     =   315
            ScaleWidth      =   6705
            TabIndex        =   89
            Top             =   270
            Width           =   6765
            Begin VB.OptionButton optEDI 
               Caption         =   "E.D.I."
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   2400
               TabIndex        =   92
               TabStop         =   0   'False
               Top             =   0
               Width           =   1485
            End
            Begin VB.OptionButton optEmail 
               Caption         =   "Email"
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   4575
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   0
               Width           =   1800
            End
            Begin VB.OptionButton optFaxManual 
               Caption         =   "Print and then fax"
               ForeColor       =   &H8000000D&
               Height          =   315
               Left            =   0
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   1890
            End
         End
      End
      Begin VB.CommandButton cmdUnlockPrices 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Unlock limits and discount"
         Height          =   555
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Accounting"
         ForeColor       =   &H8000000D&
         Height          =   1410
         Left            =   5850
         TabIndex        =   45
         Top             =   2430
         Width           =   2565
         Begin VB.TextBox txtAccountingACCNUM 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   435
            TabIndex        =   56
            Top             =   705
            Width           =   1785
         End
         Begin VB.OptionButton optBF 
            Caption         =   "Balance forward"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   375
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1260
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.OptionButton optOI 
            Caption         =   "Open item"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   375
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   1560
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Accounting package A/c no."
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   105
            TabIndex        =   57
            Top             =   390
            Width           =   2370
         End
      End
      Begin VB.TextBox txtCreditLimit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4140
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1380
      End
      Begin VB.ComboBox cboTerms 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmCustomer2.frx":03FA
         Left            =   1695
         List            =   "frmCustomer2.frx":0411
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1920
         Width           =   1755
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
         Height          =   405
         Left            =   -69885
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2595
         Width           =   870
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -70815
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2595
         Width           =   945
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -71730
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2595
         Width           =   930
      End
      Begin VB.CheckBox chkTemporary 
         Alignment       =   1  'Right Justify
         Caption         =   "Temporary customer"
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
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -68880
         TabIndex        =   8
         Top             =   4110
         Width           =   2295
      End
      Begin VB.TextBox txtDefaultDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1710
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1530
         Width           =   720
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
         Height          =   675
         Left            =   -74790
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3060
         Width           =   5385
      End
      Begin VB.Frame frInterestGroup 
         Caption         =   "Interest group"
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
         Height          =   2025
         Left            =   -70650
         TabIndex        =   30
         Top             =   600
         Width           =   4080
         Begin VB.ListBox lbIG 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   135
            TabIndex        =   33
            Top             =   795
            Width           =   2460
         End
         Begin VB.ComboBox cboIG 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   375
            Width           =   2505
         End
         Begin VB.CommandButton cmdAddIG 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Add &group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   330
            Width           =   1305
         End
         Begin VB.CommandButton cmdRemoveIG 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Remove"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2625
            Style           =   1  'Graphical
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1380
            Width           =   1050
         End
      End
      Begin VB.Frame frCustomerClassification 
         Caption         =   "Customer classification"
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
         Height          =   2025
         Left            =   -74790
         TabIndex        =   26
         Top             =   600
         Width           =   4020
         Begin VB.ListBox lbCC 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   135
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   795
            Width           =   2430
         End
         Begin VB.ComboBox cboCC 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   375
            Width           =   2475
         End
         Begin VB.CommandButton cmdAddCC 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Add &group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   345
            Width           =   1305
         End
         Begin VB.CommandButton cmdRemoveCC 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Remove"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1410
            Width           =   1050
         End
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         DragIcon        =   "frmCustomer2.frx":0445
         Height          =   1695
         Left            =   -74865
         OleObjectBlob   =   "frmCustomer2.frx":0887
         TabIndex        =   14
         Top             =   780
         Width           =   8325
      End
      Begin CoolButtonControl.CoolButton cbBillTo 
         Height          =   300
         Left            =   -72510
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2595
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
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
         Caption         =   "Bill"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbAppro 
         Height          =   300
         Left            =   -73275
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2595
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
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
         Caption         =   "Appro"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbDelTo 
         Height          =   300
         Left            =   -74040
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2595
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
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
         Caption         =   "Deliver"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbOrderTo 
         Height          =   300
         Left            =   -74820
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2595
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   529
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
         Caption         =   "Order"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent customer (for accounting system)"
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
         Left            =   2790
         TabIndex        =   80
         Top             =   450
         Width           =   3180
      End
      Begin VB.Label Label12 
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
         Left            =   -74865
         TabIndex        =   64
         Top             =   450
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "One line per invoice"
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
         Left            =   2205
         TabIndex        =   62
         Top             =   2370
         Width           =   1500
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
         Left            =   585
         TabIndex        =   59
         Top             =   2385
         Width           =   915
      End
      Begin VB.Label lblRecords 
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -74745
         TabIndex        =   55
         Top             =   4365
         Width           =   2895
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Our vendor no. with client"
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
         Height          =   300
         Left            =   -69015
         TabIndex        =   54
         Top             =   2805
         Width           =   2100
      End
      Begin VB.Label Label11 
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
         Left            =   225
         TabIndex        =   44
         Top             =   1935
         Width           =   1395
      End
      Begin VB.Label Label10 
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
         Left            =   2565
         TabIndex        =   43
         Top             =   1530
         Width           =   1395
      End
      Begin VB.Label Label9 
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
         Left            =   210
         TabIndex        =   35
         Top             =   1530
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
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
         Left            =   -74790
         TabIndex        =   34
         Top             =   2760
         Width           =   465
      End
   End
   Begin VB.TextBox txtMobile 
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
      Height          =   345
      Left            =   5598
      TabIndex        =   3
      Top             =   345
      Width           =   1830
   End
   Begin VB.TextBox txtAcno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   7470
      TabIndex        =   4
      Top             =   345
      Width           =   1395
   End
   Begin VB.TextBox txtFN 
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
      Height          =   345
      Left            =   2561
      TabIndex        =   1
      Top             =   345
      Width           =   1935
   End
   Begin VB.TextBox txtTitle 
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
      Height          =   345
      Left            =   4537
      TabIndex        =   2
      Top             =   345
      Width           =   1020
   End
   Begin VB.TextBox txtPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   6045
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdDuplicates 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Check for duplicates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6705
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
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
      Left            =   7020
      Picture         =   "frmCustomer2.frx":3C9A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5520
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   7950
      Picture         =   "frmCustomer2.frx":4024
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Width           =   915
   End
   Begin VB.TextBox txtName 
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
      Height          =   345
      Left            =   180
      TabIndex        =   0
      Top             =   345
      Width           =   2340
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales rep:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   210
      TabIndex        =   88
      Top             =   5100
      Width           =   765
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
      Left            =   315
      TabIndex        =   84
      Top             =   5760
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
      Left            =   315
      TabIndex        =   82
      Top             =   5445
      Width           =   750
   End
   Begin VB.Label lblMobile 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Height          =   240
      Left            =   6315
      TabIndex        =   24
      Top             =   105
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A/c. Num."
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
      Left            =   7530
      TabIndex        =   23
      Top             =   75
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Default phone or email"
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
      Height          =   300
      Left            =   210
      TabIndex        =   22
      Top             =   6075
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label lblFirstname 
      BackStyle       =   0  'Transparent
      Caption         =   "First name (if person)"
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
      Left            =   2610
      TabIndex        =   21
      Top             =   105
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Title (if person)"
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
      Height          =   240
      Left            =   4485
      TabIndex        =   20
      Top             =   105
      Width           =   1290
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2460
      X2              =   300
      Y1              =   210
      Y2              =   870
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   1020
      Left            =   4095
      TabIndex        =   18
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
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
      Left            =   210
      TabIndex        =   17
      Top             =   105
      Width           =   975
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCust As a_Customer
Attribute oCust.VB_VarHelpID = -1
Dim flgLoading As Boolean
Private colClassErrors As Collection
Dim XA As New XArrayDB
Dim strEMail As String
Dim bAlternativeCustomerSelected As Boolean
Dim tlRep As z_TextList
Dim tlDoc As z_TextList
Dim struct_OldTerms As OldCustomerDiscounts
Dim strPriceChangeReason As String
Dim lngSMIDPriceChange As Long
Dim bPriceChange As Boolean


Public Property Get EMail() As String
    EMail = strEMail
End Property

Private Sub cboSalesRep_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetRepID tlRep.Key(cboSalesRep)
    oCust.Repname = cboSalesRep
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cboSalesRep_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkBlock_LostFocus()
Dim bCancelled As Boolean
Dim bIsSUpervisor As Boolean
    If SecurityControl(enSECURITY_BLOCK_DEBTORS, bCancelled, "Enter your signature", "You do not have permission to set debtors' blocking (or your signature is invalid)", bIsSUpervisor) = True Then
        oCust.SetBlocked (chkBlock = 1)
    End If

End Sub

Private Sub chkBlock_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkBlock_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub chkOneLinePerInvoice_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.OneLinePerInvoice = (chkOneLinePerInvoice = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkOneLinePerInvoice_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkShowVAT_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.ShowVAT = (chkShowVAT = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkShowVAT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkUseQuotedPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oCust.SetUseQuotedPrice (chkUseQuotedPrice = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkUseQuotedPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub chkCompleteOrder_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oCust.SetOrderCOmplete (chkCompleteOrder = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkCompleteOrder_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddCC_Click()
    On Error GoTo errHandler
Dim oCC As New a_IG
    If flgLoading Then Exit Sub
    If cboCC = "" Then Exit Sub
    Set oCC = oCust.CustomerTypes.Add
    oCC.TPID = oCust.ID
    oCC.IGID = oCust.CustomerTypesALL_tl.Key(cboCC)
    oCC.Description = cboCC
    oCC.ApplyEdit
    oCC.BeginEdit
    cboCC.RemoveItem cboCC.ListIndex
    If cboCC.ListCount > 0 Then
        cboCC.ListIndex = 0
    Else
        cboCC.ListIndex = -1
    End If
    LoadTPCCs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdAddCC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDefaultAddress_MouseEnter()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdDefaultAddress_MouseEnter", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdKeep_Click()
    On Error GoTo errHandler
Dim frmS As frmBrowseCustomers2
    Set frmS = New frmBrowseCustomers2
    frmS.Show vbModal
    txtParent = frmS.CustomerName & " " & frmS.Accnum
    oCust.ParentCustomerID = frmS.CustomerID
    'oSUpp. = frmS.SupplierID
    Unload frmS
    Set frmS = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdKeep_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveCC_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbCC = "" Then Exit Sub
    oCust.CustomerTypes.Remove oCust.CustomerTypes.Key(Me.lbCC)
    cboCC.AddItem Me.lbCC
    cboCC.ListIndex = 0
    LoadTPCCs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdRemoveCC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkVatable_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.VATable = (chkVATable = 1)
    Me.chkShowVAT.Enabled = (chkVATable = 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.chkVatable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo errHandler
Dim frm As frmAddress
Dim oAdd As a_Address
    If flgLoading Then Exit Sub
    If oCust.Addresses.Count > 0 Then
        If oCust.Addresses(oCust.Addresses.Count).Addressee = "" Then
            MsgBox "There is already an incomplete address for this customer. Delete it or use it rather than adding a new address", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
    Set frm = New frmAddress
    Set oAdd = oCust.Addresses.Add
    oAdd.BeginEdit
    oAdd.SetAddressee oCust.Title & IIf(oCust.Title > "", " ", "") & oCust.Initials & IIf(oCust.Initials > "", " ", "") & oCust.Name
    frm.component oAdd
    frm.Show vbModal
    LoadArray
    LoadIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdAddIG_Click()
    On Error GoTo errHandler
Dim oIG As a_IG
    If flgLoading Then Exit Sub
    If cboIG = "" Then Exit Sub
    Set oIG = oCust.InterestGroups.Add
 '   oIG.BeginEdit
    oIG.TPID = oCust.ID
    oIG.IGID = oCust.InterestGroupsActive_tl.Key(cboIG)
    oIG.Description = cboIG
    oIG.ApplyEdit
    oIG.BeginEdit
    cboIG.RemoveItem cboIG.ListIndex
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdAddIG_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdRemoveIG_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lbIG = "" Then Exit Sub
    oCust.InterestGroups.Remove oCust.InterestGroups.Key(Me.lbIG)
    cboIG.AddItem Me.lbIG
    cboIG.ListIndex = 0
    LoadTPIGs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdRemoveIG_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub Component2(PID As Long)
    On Error GoTo errHandler
    Set oCust = New a_Customer
    oCust.Load PID
    Me.Caption = "Customer master edit: " & oCust.Name
    CustomizeForm
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.Component2(PID)", PID
End Sub


Public Sub component(pCust As a_Customer)
    On Error GoTo errHandler
    Set oCust = pCust
  '  oCust.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.component(pCust)", pCust
End Sub
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.EnableOK(pOK)", pOK
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    flgLoading = True
    oCust.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim frmP As frmCustomerPreview
Dim oSM As z_StockManager
Dim errRepeat As Integer
    If flgLoading Then Exit Sub
    
    oCust.GenerateSeparateInvoicesForSeparateOrders = (Me.chkSepInvs.Value = 1)
    If oPC.GetProperty("SecureTF") = "TRUE" Then
        If bPriceChange = True Then  'Write an audit record
            
            Set oSM = New z_StockManager
            If struct_OldTerms.Blocked <> oCust.Blocked Then
                oSM.InsertAuditRecord "BL", strPriceChangeReason, IIf(struct_OldTerms.Blocked, 1, 0), IIf(oCust.Blocked, 1, 0), lngSMIDPriceChange, "", oCust.CustomerID
            End If
            If struct_OldTerms.CreditLimit <> oCust.CreditLimit Then
                oSM.InsertAuditRecord "CL", strPriceChangeReason, Format(struct_OldTerms.CreditLimit, "###,##0"), Format(oCust.CreditLimit, "###,##0"), lngSMIDPriceChange, "", oCust.CustomerID
            End If
            If struct_OldTerms.Discount <> oCust.DefaultDiscount Then
                oSM.InsertAuditRecord "CD", strPriceChangeReason, PBKSPercentF(struct_OldTerms.Discount), oCust.DefaultDiscountF, lngSMIDPriceChange, "", oCust.CustomerID
            End If
            If struct_OldTerms.Terms <> oCust.Terms Then
                oSM.InsertAuditRecord "TM", strPriceChangeReason, Format(struct_OldTerms.Terms, "###,##0"), Format(oCust.Terms, "###,##0"), lngSMIDPriceChange, "", oCust.CustomerID
            End If
        End If
    End If
    
    
    
    If oCust.IsNew Then
        bAlternativeCustomerSelected = False
        oCust.LookforDuplicates
        If bAlternativeCustomerSelected Then
            oCust.CancelEdit
            Unload Me
            Exit Sub
        End If
    End If
    If oCust.CustomerIndexClashes = True Then
        MsgBox "This account number has already been used for another customer. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If

    oCust.ApplyEdit lngResult
    If lngResult = 0 Then
        Set frmP = New frmCustomerPreview
        frmP.component oCust
        frmP.LoadControls
        frmP.Show

        Unload Me
    ElseIf lngResult = 22 Then
        MsgBox "You are trying to save a customer with duplicate values." & vbCrLf & "These are likely to be in the Acc No. field or in the address description fields.", , "Can't save"
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmCustomer: cmdOK_Click"  'unknown source
        If errRepeat < 5 Then
            Resume
        Else
            LogSaveToFile "Access violation in frmCustomer: cmdOK_Click after 5 re-attempts"
            MsgBox "Memory error trying to save form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdOK_Click", , EA_NORERAISE, , "Line number", Array(Erl())
    HandleError
End Sub


Private Sub cmdUnlockPrices_Click()
    On Error GoTo errHandler
Dim bIsSUpervisor As Boolean
Dim bCancelled As Boolean
Dim strName As String
Dim frm As New frmAudit_ProductPrices

    If SecurityControl(enSECURITY_CUSTDISCOUNT_AUTH, bCancelled, "Enter your signature", "You do not have permission to unlock the price fields (or your signature is invalid)", bIsSUpervisor, strName, lngSMIDPriceChange) = True Then
        bPriceChange = True
        'Set old prices
        struct_OldTerms.CreditLimit = oCust.CreditLimit
        struct_OldTerms.Discount = oCust.DefaultDiscount
        struct_OldTerms.Terms = oCust.Terms
        struct_OldTerms.Blocked = oCust.Blocked
        frm.lstreasons.Visible = False
        frm.cmdOK.Enabled = False
        frm.txtReason.TOP = 1000
        frm.txtReason.Height = 2100
        frm.Show vbModal
        strPriceChangeReason = frm.Reason
        If Not frm.Cancelled Then
            Me.txtDefaultDiscount.Locked = False
            Me.txtDefaultDiscount.BackColor = &H80000005
            Me.txtCreditLimit.Locked = False
            Me.txtCreditLimit.BackColor = &H80000005
            Me.cboTerms.Locked = False
            Me.cboTerms.BackColor = &H80000005
          '  Me.chkBlock.Enabled = True
        End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdUnlockPrices_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    flgLoading = True
    If Me.WindowState <> 2 Then
        TOP = 0
        Left = 50
        Height = 7150
        Width = 9250
    End If
    CustomizeForm
    Me.Caption = "Customer master edit: " & oCust.Name
    
    txtName = oCust.Name
    txtFN = oCust.Initials
    txtAcno = oCust.AcNo
    txtSAN = oCust.SAN
    txtTitle = oCust.Title
    txtNote = oCust.Note
    txtMobile = oCust.MOBILE
    txtParent = oCust.ParentCustomerName
    txtSalesOrderTemplateName = oCust.SalesOrderTemplateName
    txtApproTemplateName = oCust.ApproTemplateName
    txtApproReturnTemplateName = oCust.ApproReturnTemplateName
    txtQuotationTemplateName = oCust.QuotationTemplateName
    txtInvoiceTemplateName = oCust.InvoiceTemplateName
    txtCreditNoteTemplatename = oCust.CreditNoteTemplateName
    txtDefaultDiscount = oCust.DefaultDiscountF
    chkVATable = IIf(oCust.VATable, 1, 0)
    chkShowVAT = IIf(oCust.ShowVAT, 1, 0)
    Me.chkOneLinePerInvoice = IIf(oCust.OneLinePerInvoice, 1, 0)
    txtVATNumber = oCust.VatNumber
    chkShowVAT.Enabled = (chkVATable = 0)
    txtOurAcnoWithClient = oCust.OurACnoWithClient
    txtAccountingACCNUM = oCust.AccAcno
    txtContact = oCust.ContactPerson
    txtContactPhone = oCust.ContactpersonPhoneF
    
    txtDefaultDiscount = oCust.DefaultDiscountF
    chkTemporary = IIf(oCust.CanBeDeleted, 1, 0)
    chkBlock = IIf(oCust.Blocked = True, 1, 0)
    Me.chkSepInvs = IIf(oCust.GenerateSeparateInvoicesForSeparateOrders, 1, 0)
    
    Me.chkUseQuotedPrice = IIf(oCust.UseQuotedPrice = True, 1, 0)
    Me.chkCompleteOrder = IIf(oCust.CompleteOrder = True, 1, 0)
    
    If oPC.GetProperty("SecureTF") = "TRUE" Then
        cmdUnlockPrices.Visible = True
        Me.txtDefaultDiscount.Locked = True
        Me.txtDefaultDiscount.BackColor = &H80000018
        Me.txtCreditLimit.Locked = True
        Me.txtCreditLimit.BackColor = &H80000018
        Me.cboTerms.Locked = True
        Me.cboTerms.BackColor = &H80000018
    Else
        cmdUnlockPrices.Visible = False
        Me.txtDefaultDiscount.Locked = False
        Me.txtDefaultDiscount.BackColor = &H80000005
        Me.txtCreditLimit.Locked = False
        Me.txtCreditLimit.BackColor = &H80000005
        Me.cboTerms.Locked = False
        Me.cboTerms.BackColor = &H80000005
    End If
    
    Set tlRep = New z_TextList
    tlRep.Load ltSalesRep, , "<None>"
    LoadCombo cboSalesRep, tlRep
    If oCust.Repname > "" Then
        On Error Resume Next
        cboSalesRep.text = oCust.Repname
        If Err Then
            cboSalesRep.text = "<None>"
            Err.Clear
        End If
        On Error GoTo errHandler
    End If
    Me.txtCreditLimit = oCust.CreditLimitF
    For i = 0 To cboTerms.ListCount - 1
        If cboTerms.ItemData(i) = oCust.Terms Then
            cboTerms.ListIndex = i
        End If
    Next

    'LoadCombo oPC.Configuration.DocumentControls., oPC.Configuration
    
    If oCust.PaymentStyle = "B" Then
        optBF = True
    ElseIf oCust.PaymentStyle = "O" Then
        optOI = True
    End If
    LoadArray
    LoadIGs
    LoadTPIGs
'    cboCC.Text = oCust.CustomerTypesActive_tl.Item(oCust.CustomerTypeID)
    LoadCCs
    LoadTPCCs
    RestrictInterestGroups
    RestrictCustomerTypes
    oCust.GetStatus
    Select Case oCust.DispatchMethod
    Case "E"
        optEDI = True
    Case "M"
        optEmail = True
    Case "P"
        optFaxManual = True
    End Select
    
    Me.SSTab1.Tab = 0
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub CustomizeForm()
    On Error GoTo errHandler
    If oCust.CustomerTypeID = oPC.Configuration.BusinessCustomerTypeID Then
        Me.Caption = Me.Caption & " (business)"
     '   frVAT.Visible = True
        txtFN.Visible = False
        txtTitle.Visible = False
        lblFirstname.Visible = False
        lblTitle.Visible = False
        lblMobile.Visible = False
        txtMobile.Visible = False
        txtName.Width = txtName.Width * 2
    ElseIf oCust.CustomerTypeID = oPC.Configuration.BookClubCustomerTypeID Then
        Me.Caption = Me.Caption & " (book club)"
    '    frVAT.Visible = False
        txtFN.Visible = False
        txtTitle.Visible = False
        lblFirstname.Visible = False
        lblTitle.Visible = False
        lblMobile.Visible = False
        txtMobile.Visible = False
    ElseIf oCust.CustomerTypeID = oPC.Configuration.PrivateCustomerTypeID Then
        Me.Caption = Me.Caption & " (private)"
     '   frVAT.Visible = False
        txtFN.Visible = True
        txtTitle.Visible = True
        lblFirstname.Visible = True
        lblTitle.Visible = True
        lblMobile.Visible = True
        txtMobile.Visible = True
        
    End If
    If oCust.CustomerTypes.IsALoyaltyMember = True Then
        Me.frCustomerClassification.Enabled = False
        Me.frInterestGroup.Enabled = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.CustomizeForm"
End Sub

Private Sub LoadCCs()
    On Error GoTo errHandler
    LoadCombo Me.cboCC, oCust.CustomerTypesActive_tl
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadCCs"
End Sub
Private Sub RestrictInterestGroups()
    On Error GoTo errHandler
Dim oTPIG As a_IG
Dim i As Integer

    For Each oTPIG In oCust.InterestGroups
        For i = cboIG.ListCount To 1 Step -1
            cboIG.ListIndex = i - 1
            If oTPIG.Description = cboIG Then
                cboIG.RemoveItem cboIG.ListIndex
            End If
        Next
    Next
    If cboIG.ListCount > 0 Then
        cboIG.ListIndex = 0
    Else
        cboIG.ListIndex = -1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.RestrictInterestGroups"
End Sub
Private Sub RestrictCustomerTypes()
    On Error GoTo errHandler
Dim oTPIG As a_IG
Dim i As Integer

    For Each oTPIG In oCust.CustomerTypes
        For i = cboCC.ListCount To 1 Step -1
            cboCC.ListIndex = i - 1
            If oTPIG.Description = cboCC Then
                cboCC.RemoveItem cboCC.ListIndex
            End If
        Next
    Next
    If cboCC.ListCount > 0 Then
        cboCC.ListIndex = 0
    Else
        cboCC.ListIndex = -1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.RestrictCustomerTypes"
End Sub

Private Sub LoadIGs()
    On Error GoTo errHandler
    LoadCombo Me.cboIG, oCust.InterestGroupsActive_tl
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadIGs"
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
    ErrorIn "frmCustomer.LoadTPCCs"
End Sub
Private Sub LoadTPIGs()
    On Error GoTo errHandler
Dim oTPIG As a_IG
    With Me.lbIG
        .Clear
        For Each oTPIG In oCust.InterestGroups
            .AddItem oCust.InterestGroupsAll_tl.Item(CStr(oTPIG.IGID))
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadTPIGs"
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oCust = Nothing
    Set colClassErrors = Nothing
    Set XA = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub oCust_Valid(strMsg As String)
    On Error GoTo errHandler
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oCust_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
    HandleError
End Sub

Private Function ShowDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
Dim frm As frmDuplicateCustomers
Dim tmpCust As a_Customer
    
    Set frm = New frmDuplicateCustomers
    frm.component Me.txtName, pDuplicates
    frm.Show vbModal
    If frm.SelectedCustomer > 0 Then
        Set Forms(0).frmMainCustomerPreview = Nothing
        Set Forms(0).frmMainCustomerPreview = New frmCustomerPreview
        Set tmpCust = New a_Customer
        tmpCust.Load frm.SelectedCustomer
        Forms(0).frmMainCustomerPreview.component tmpCust
        Unload frm
        bAlternativeCustomerSelected = True
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.ShowDuplicates(pDuplicates)", pDuplicates
End Function

Private Sub optBF_Click()
    On Error GoTo errHandler
    If optBF = True Then
        oCust.SetPaymentStyle "B"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.optBF_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub oSupp_Valid(strMsg As String)
    On Error GoTo errHandler
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.oSupp_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub


Private Sub optOI_Click()
    On Error GoTo errHandler
    If optOI = True Then
        oCust.SetPaymentStyle "O"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.optOI_Click", , EA_NORERAISE
    HandleError
End Sub







Private Sub txtCreditLimit_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If txtCreditLimit <> oCust.CreditLimitF Then
        Cancel = Not oCust.SetCreditLimit(Trim(txtCreditLimit))
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtCreditLimit_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub




Private Sub txtDefaultDiscount_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtDefaultDiscount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtDefaultDiscount_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtDefaultDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If txtDefaultDiscount <> oCust.DefaultDiscountF Then
        Cancel = Not oCust.SetDefaultDiscount(Trim(txtDefaultDiscount))
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtDefaultDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub cboTerms_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oCust.SetTerms cboTerms.ItemData(cboTerms.ListIndex)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cboTerms_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub txtMobile_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtMobile = oCust.MOBILE
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtMobile_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtMobile_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetCell txtMobile
    If Err Then
      Beep
      intPos = txtMobile.SelStart
      txtMobile = oCust.MOBILE
      txtMobile.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtMobile_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtMobile_Validate(Cancel As Boolean)
        On Error Resume Next
    If txtMobile = "" Then Exit Sub
    txtMobile = PhoneFormat(txtMobile, "")
    Cancel = Not oCust.SetCell(txtMobile)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtMobile_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtContact_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetContactPerson (txtContact)
    If Err Then
      Beep
      intPos = txtContact.SelStart
      txtContact = oCust.Title
      txtContact.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtContact_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtContact_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetContactPerson(txtContact)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtContact_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtContactPhone_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetContactpersonPhone txtContactPhone
    If Err Then
      Beep
      intPos = txtContactPhone.SelStart
      txtContactPhone = oCust.ContactpersonPhoneF
      txtContactPhone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtContactPhone_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtContactPhone_Validate(Cancel As Boolean)
        On Error Resume Next
    If txtContactPhone = "" Then Exit Sub
    txtContactPhone = PhoneFormat(txtContactPhone, "")
    Cancel = Not oCust.SetContactpersonPhone(txtContactPhone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtContactPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtOurAcnoWithClient_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetOurACnoWithClient txtOurAcnoWithClient
    If Err Then
      Beep
      intPos = txtOurAcnoWithClient.SelStart
      txtOurAcnoWithClient = oCust.OurACnoWithClient
      txtOurAcnoWithClient.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtOurAcnoWithClient_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtOurAcnoWithClient_LostFocus()
    On Error GoTo errHandler
    txtOurAcnoWithClient = oCust.OurACnoWithClient
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtOurAcnoWithClient_LostFocus", , EA_NORERAISE
    HandleError
End Sub





'Private Sub txtBusPhone_LostFocus()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    txtBusphone = oAdd.BusPhone
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAddress.txtBusPhone_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub txtPhone_LostFocus()
    On Error GoTo errHandler
    txtPhone = oCust.Phone
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtPhone_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
        On Error Resume Next
    Cancel = Not oCust.SetPhone(txtPhone)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPhone_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetPhone (txtPhone)
    If Err Then
      Beep
      intPos = txtPhone.SelStart
      txtPhone = oCust.Phone
      txtPhone.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
    txtName = oCust.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetName (txtName)
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oCust.Name
      txtName.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtName_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    oCust.SetName txtName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtName_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtQuotationTemplateName_Change()
        On Error Resume Next
    oCust.SetQuotationTemplateName txtQuotationTemplateName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtQuotationTemplateName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSalesOrderTemplateName_Change()
        On Error Resume Next
    oCust.SetSalesOrderTemplateName txtSalesOrderTemplateName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtSalesOrderTemplateName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtApproTemplateName_Change()
        On Error Resume Next
    oCust.SetApproTemplateName txtApproTemplateName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtApproTemplateName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtApproReturnTemplateName_Change()
        On Error Resume Next
    oCust.SetApproReturnTemplateName txtApproReturnTemplateName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtApproReturnTemplateName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCreditNoteTemplatename_Change()
        On Error Resume Next
    oCust.SetCreditNoteTemplateName txtCreditNoteTemplatename
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtCreditNoteTemplatename_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtInvoiceTemplateName_Change()
        On Error Resume Next
    oCust.SetInvoiceTemplateName txtInvoiceTemplateName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtInvoiceTemplateName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSAN_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetSAN (txtSAN)
    If Err Then
      Beep
      intPos = txtSAN.SelStart
      txtSAN = oCust.SAN
      txtSAN.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtSAN_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSAN_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetSAN(txtSAN)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtSAN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtVATNumber_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetVATNumber (txtVATNumber)
    If Err Then
      Beep
      intPos = txtVATNumber.SelStart
      txtVATNumber = oCust.VatNumber
      txtVATNumber.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtVATNumber_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVATNumber_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetVATNumber(txtVATNumber)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtVATNumber_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtAcno_LostFocus()
    On Error GoTo errHandler
    txtAcno = oCust.AcNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAcno_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetAcNO (txtAcno)
    If Err Then
      Beep
      intPos = txtAcno.SelStart
      txtAcno = oCust.AcNo
      txtAcno.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAcno_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Validate(Cancel As Boolean)
        On Error Resume Next

    If flgLoading Then Exit Sub
    oCust.SetAcNO txtAcno
    If oCust.CustomerIndexClashes = True Then
        MsgBox "This account number has already been used for another customer. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAcno_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'txtAccountingACCNUM
Private Sub txtAccountingACCNUM_LostFocus()
    On Error GoTo errHandler
    txtAccountingACCNUM = oCust.AccAcno
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAccountingACCNUM_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAccountingACCNUM_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetAccAcNo (txtAccountingACCNUM)
    If Err Then
      Beep
      intPos = txtAccountingACCNUM.SelStart
      txtAccountingACCNUM = oCust.AccAcno
      txtAccountingACCNUM.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAccountingACCNUM_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAccountingACCNUM_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetAccAcNo(txtAccountingACCNUM)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtAccountingACCNUM_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtFN_LostFocus()
    On Error GoTo errHandler
    txtFN = oCust.Initials
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtFN_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetInitials (txtFN)
    If Err Then
      Beep
      intPos = txtFN.SelStart
      txtFN = oCust.Initials
      txtFN.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtFN_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFN_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetInitials(txtFN)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtFN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetNote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    txtNote = oCust.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
        On Error Resume Next
Dim intPos As Integer
  '  txtNote = HandleTextWithBites(txtNote)
    oCust.SetNote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oCust.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    txtTitle = oCust.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
        On Error Resume Next
Dim intPos As Integer
    oCust.SetTitle (txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oCust.Title
      txtTitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
        On Error Resume Next
    If flgLoading Then Exit Sub
    Cancel = Not oCust.SetTitle(txtTitle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub LoadArray()
    On Error GoTo errHandler
'Dim objItem As d_Customer
Dim lngIndex As Long
    XA.ReDim 1, oCust.Addresses.Count, 1, 6
    For lngIndex = 1 To oCust.Addresses.Count
        XA.Value(lngIndex, 1) = lngIndex
        XA.Value(lngIndex, 2) = oCust.Addresses(lngIndex).AddressMailing
        If XA.Value(lngIndex, 2) = "" Then
            XA.Value(lngIndex, 2) = "<Double-click to edit>"
        End If
        XA.Value(lngIndex, 3) = CreateRoleString(oCust.Addresses(lngIndex))
        XA.Value(lngIndex, 4) = oCust.Addresses(lngIndex).GetsCatalogue
        XA.Value(lngIndex, 5) = oCust.Addresses(lngIndex).Key
        XA.Value(lngIndex, 6) = oCust.Addresses(lngIndex).ForMailing
    Next
    XA.QuickSort 1, XA.UpperBound(1), 1, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind
  '  G1.Refresh
    If XA.UpperBound(1) > 1 Then
        Me.lblRecords = XA.UpperBound(1) & " addresses"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.LoadArray"
End Sub

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmAddress
Dim lngID As Long
Dim oAdd As a_Address
Dim Key As String


    Set frm = New frmAddress
    Key = XA(G1.Bookmark, 5)
    If Key > "" Then
        Set oAdd = oCust.Addresses.Item(Key)
        If oCust.Initials > "" Then
            oAdd.SetAddressee oCust.Title & " " & oCust.Initials & " " & oCust.Name
        End If
        frm.component oAdd
        frm.Show vbModal
        LoadArray
    End If
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmCustomer: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmCustomer: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdRemove_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    oCust.Addresses.Remove XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbAppro_Click()
    On Error GoTo errHandler
    oCust.SetApproAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbAppro_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbBillTo_Click()
    On Error GoTo errHandler
    oCust.SetBillToAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbBillTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbDelTo_Click()
    On Error GoTo errHandler
    oCust.SetDelToAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbOrderTo_Click()
    On Error GoTo errHandler
    oCust.SetOrderToAddressidx XA(G1.Bookmark, 5)
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cbOrderTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmAddress
Dim oAdd As a_Address
    If flgLoading Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    Set frm = New frmAddress
    Set oAdd = oCust.Addresses.Item(XA(G1.Bookmark, 5))
    If oAdd.Addressee = "" Then oAdd.SetAddressee oCust.Title & " " & oCust.Initials & " " & oCust.Name
    frm.component oAdd
    frm.Show vbModal
    LoadArray
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.cmdEdit_Click", , EA_NORERAISE
    HandleError
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
    ErrorIn "frmCustomer.CreateRoleString(pAddress)", pAddress
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
    ErrorIn "frmCustomer.CreateRoleString2(pAddress)", pAddress
End Function


Private Sub optEDI_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetDispatchMethod "E"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.optEDI_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optEmail_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetDispatchMethod "M"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.optEmail_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optFaxManual_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oCust.SetDispatchMethod "P"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomer.optFaxManual_Click", , EA_NORERAISE
    HandleError
End Sub
