VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSupplier 
   BackColor       =   &H00F7EDE8&
   Caption         =   "Supplier"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   ControlBox      =   0   'False
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   10260
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
      Left            =   8955
      Picture         =   "frmSupplier.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4125
      Width           =   1000
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
      Left            =   7935
      Picture         =   "frmSupplier.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4125
      Width           =   1000
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
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   450
      Width           =   4275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3165
      Left            =   165
      TabIndex        =   6
      Top             =   870
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   5583
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   600
      BackColor       =   16248296
      ForeColor       =   10485760
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1. Ordering"
      TabPicture(0)   =   "frmSupplier.frx":0A9E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(2)=   "lvwDeals"
      Tab(0).Control(3)=   "cmdAddDeal"
      Tab(0).Control(4)=   "cmdRemoveDeal"
      Tab(0).Control(5)=   "cmdEditDeal"
      Tab(0).Control(6)=   "cboCurr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(8)=   "cboDispatchMode"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&2. Addresses"
      TabPicture(1)   =   "frmSupplier.frx":0ABA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbOrderTo"
      Tab(1).Control(1)=   "cbDelTo"
      Tab(1).Control(2)=   "cmdApproAddress"
      Tab(1).Control(3)=   "cbBillTo"
      Tab(1).Control(4)=   "lvwAddresses"
      Tab(1).Control(5)=   "cmdAdd"
      Tab(1).Control(6)=   "cmdRemove"
      Tab(1).Control(7)=   "cmdEdit"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "&3. Miscellaneous"
      TabPicture(2)   =   "frmSupplier.frx":0AD6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(3)=   "Label6"
      Tab(2).Control(4)=   "Label4"
      Tab(2).Control(5)=   "Label8"
      Tab(2).Control(6)=   "Label5"
      Tab(2).Control(7)=   "lblCurrencyConversionRate"
      Tab(2).Control(8)=   "txtParent"
      Tab(2).Control(9)=   "cmdKeep"
      Tab(2).Control(10)=   "txtReturnEndMonths"
      Tab(2).Control(11)=   "txtReturnStartMonths"
      Tab(2).Control(12)=   "chkActive"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtDefaultDeliverydays"
      Tab(2).Control(14)=   "chkVatable"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtRecordAdded"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtRecordLastChanged"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "chkDistributorOnly"
      Tab(2).Control(18)=   "txtCurrencyConversionRate"
      Tab(2).ControlCount=   19
      TabCaption(3)   =   "&4. Note"
      TabPicture(3)   =   "frmSupplier.frx":0AF2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtNote"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Terms"
      TabPicture(4)   =   "frmSupplier.frx":0B0E
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label16"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label17"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label18(0)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label18(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Label11"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cboTerms"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cboSettlementDiscountTerms"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cboSettlementDiscountTermsType"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "txtSettlementDiscount"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "cboTermType"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      Begin VB.ComboBox cboTermType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmSupplier.frx":0B2A
         Left            =   2040
         List            =   "frmSupplier.frx":0B31
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   1125
         Width           =   1755
      End
      Begin VB.TextBox txtSettlementDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   7350
         TabIndex        =   57
         Top             =   705
         Width           =   690
      End
      Begin VB.ComboBox cboSettlementDiscountTermsType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmSupplier.frx":0B3A
         Left            =   7320
         List            =   "frmSupplier.frx":0B44
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1590
         Width           =   1755
      End
      Begin VB.ComboBox cboSettlementDiscountTerms 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmSupplier.frx":0B6B
         Left            =   7335
         List            =   "frmSupplier.frx":0B6D
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1155
         Width           =   1755
      End
      Begin VB.ComboBox cboTerms 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "frmSupplier.frx":0B6F
         Left            =   2040
         List            =   "frmSupplier.frx":0B86
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   690
         Width           =   1755
      End
      Begin VB.ComboBox cboDispatchMode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68430
         Style           =   2  'Dropdown List
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2475
         Width           =   1860
      End
      Begin VB.TextBox txtCurrencyConversionRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71505
         MaxLength       =   8
         TabIndex        =   46
         Top             =   2175
         Width           =   540
      End
      Begin VB.CheckBox chkDistributorOnly 
         Alignment       =   1  'Right Justify
         Caption         =   "We receive from but do not order from this supplier."
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   -69675
         TabIndex        =   45
         Top             =   930
         Width           =   4065
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -74835
         MultiLine       =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   525
         Width           =   7560
      End
      Begin VB.TextBox txtRecordLastChanged 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   285
         Left            =   -67620
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1680
      End
      Begin VB.TextBox txtRecordAdded 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   285
         Left            =   -67620
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2370
         Width           =   1680
      End
      Begin VB.CheckBox chkVatable 
         Caption         =   "Charges V.A.T."
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74730
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2535
         Width           =   2550
      End
      Begin VB.TextBox txtDefaultDeliverydays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -73275
         TabIndex        =   3
         Top             =   585
         Width           =   690
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Inactive"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74730
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2520
      End
      Begin VB.TextBox txtReturnStartMonths 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -70650
         TabIndex        =   5
         Top             =   1740
         Width           =   435
      End
      Begin VB.TextBox txtReturnEndMonths 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   -71850
         TabIndex        =   4
         Top             =   1740
         Width           =   435
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
         Left            =   -65955
         Picture         =   "frmSupplier.frx":0BBA
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   570
         Width           =   360
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
         Left            =   -68250
         TabIndex        =   31
         Top             =   585
         Width           =   2280
      End
      Begin VB.Frame Frame1 
         Caption         =   "Document delivery method"
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
         Height          =   1455
         Left            =   -70050
         TabIndex        =   24
         Top             =   930
         Width           =   3420
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
            Left            =   555
            TabIndex        =   27
            Top             =   630
            Width           =   1785
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
            Left            =   555
            TabIndex        =   26
            Top             =   990
            Width           =   1560
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
            Left            =   555
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   300
            Value           =   -1  'True
            Width           =   1650
         End
      End
      Begin VB.ComboBox cboCurr 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68475
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   510
         Width           =   1860
      End
      Begin VB.CommandButton cmdEditDeal 
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
         Left            =   -72870
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2415
         Width           =   930
      End
      Begin VB.CommandButton cmdRemoveDeal 
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
         Left            =   -73815
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2415
         Width           =   930
      End
      Begin VB.CommandButton cmdAddDeal 
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
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2415
         Width           =   930
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
         Left            =   -66540
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2445
         Width           =   930
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
         Left            =   -67500
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2445
         Width           =   930
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
         Left            =   -68445
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2445
         Width           =   930
      End
      Begin MSComctlLib.ListView lvwAddresses 
         Height          =   1845
         Left            =   -74730
         TabIndex        =   14
         Top             =   525
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3254
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Address type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Phone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Purpose"
            Object.Width           =   3598
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDeals 
         Height          =   1845
         Left            =   -74730
         TabIndex        =   18
         Top             =   525
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   3254
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Discount"
            Object.Width           =   1605
         EndProperty
      End
      Begin CoolButtonControl.CoolButton cbBillTo 
         Height          =   300
         Left            =   -73950
         TabIndex        =   20
         Top             =   2460
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
      Begin CoolButtonControl.CoolButton cmdApproAddress 
         Height          =   300
         Left            =   -74730
         TabIndex        =   21
         Top             =   2460
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
         Left            =   -72375
         TabIndex        =   22
         Top             =   2460
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
         Left            =   -73155
         TabIndex        =   23
         Top             =   2460
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
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Settlement discount"
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
         Left            =   5460
         TabIndex        =   58
         Top             =   735
         Width           =   1680
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Settlement discount terms type"
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
         Index           =   1
         Left            =   4035
         TabIndex        =   56
         Top             =   1620
         Width           =   3105
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Settlement discount terms"
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
         Index           =   0
         Left            =   4035
         TabIndex        =   54
         Top             =   1185
         Width           =   3105
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Term type"
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
         Left            =   510
         TabIndex        =   52
         Top             =   1155
         Width           =   1320
      End
      Begin VB.Label Label16 
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
         Left            =   525
         TabIndex        =   51
         Top             =   705
         Width           =   1320
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usual dispatch method"
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
         Left            =   -70125
         TabIndex        =   49
         Top             =   2520
         Width           =   1650
      End
      Begin VB.Label lblCurrencyConversionRate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rate for converting supplier's currency to local"
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
         Left            =   -74880
         TabIndex        =   47
         Top             =   2220
         Width           =   3285
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Record last changed: "
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
         Left            =   -69465
         TabIndex        =   43
         Top             =   2715
         Width           =   1800
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Record added: "
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
         Left            =   -68970
         TabIndex        =   42
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usual lead time (days)"
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
         Left            =   -74910
         TabIndex        =   41
         Top             =   615
         Width           =   1545
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock can be returned when delivered <="
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
         Left            =   -74880
         TabIndex        =   40
         Top             =   1785
         Width           =   3090
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "and >="
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
         Left            =   -71310
         TabIndex        =   39
         Top             =   1785
         Width           =   585
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "months prior (e.g. <= 12 and >= 6 months prior)"
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
         Left            =   -70095
         TabIndex        =   38
         Top             =   1785
         Width           =   3645
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent supplier"
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
         Left            =   -69375
         TabIndex        =   37
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Order in this currency"
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
         Left            =   -70065
         TabIndex        =   19
         Top             =   540
         Width           =   1620
      End
   End
   Begin VB.TextBox txtAcno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7800
      TabIndex        =   1
      Top             =   105
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   5430
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   6705
      X2              =   495
      Y1              =   30
      Y2              =   915
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   180
      TabIndex        =   10
      Top             =   4470
      Width           =   3780
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
      Left            =   840
      TabIndex        =   9
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc. Num."
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
      Left            =   6990
      TabIndex        =   8
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Default phone: "
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
      Left            =   -15
      TabIndex        =   7
      Top             =   480
      Width           =   1365
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oSupp As a_Supplier
Attribute oSupp.VB_VarHelpID = -1
Private colClassErrors As Collection
Dim flgLoading As Boolean

Private Sub cboCurr_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oSupp.setDefaultCurrency oPC.Configuration.Currencies.FindByDescription(cboCurr)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cboCurr_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadCurrs()
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim oItem As ListItem
Dim i As Integer
    For Each oCurr In oPC.Configuration.Currencies
        Me.cboCurr.AddItem oCurr.Description
    Next
    cboCurr = oSupp.DefaultCurrency.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.LoadCurrs"
End Sub
Private Sub LoadDispatchModes()
    On Error GoTo errHandler
    LoadCombo Me.cboDispatchMode, oSupp.DispatchModes
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.LoadDispatchModes"
End Sub

Private Sub cboDispatchMode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oSupp.SetDispatchModeID cboDispatchMode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cboDispatchMode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub cboSettlementDiscountTerms_Validate(Cancel As Boolean)
    oSupp.SetSettlementTerms cboSettlementDiscountTerms.ItemData(cboSettlementDiscountTerms.ListIndex)
End Sub

Private Sub cboSettlementDiscountTermsType_Validate(Cancel As Boolean)
    oSupp.SetSettlementTermsType cboSettlementDiscountTermsType.ItemData(cboSettlementDiscountTermsType.ListIndex)
End Sub

Private Sub cboTerms_Validate(Cancel As Boolean)
    oSupp.SetTerms cboTerms.ItemData(cboTerms.ListIndex)
End Sub


Private Sub cboTermType_Validate(Cancel As Boolean)
    oSupp.SetTermsType cboTermType.ItemData(cboTermType.ListIndex)
End Sub

Private Sub chkActive_Click()
    On Error GoTo errHandler
    oSupp.UseStatus = IIf(chkActive = 1, 1, 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.chkActive_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkDistributorOnly_Click()
    On Error GoTo errHandler
    oSupp.DoNotOrderFrom = (chkDistributorOnly = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.chkDistributorOnly_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkVatable_Click()
    On Error GoTo errHandler
    oSupp.VATable = (chkVATable = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.chkVatable_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo errHandler
Dim frm As frmAddress
Dim oAdd As a_Address
    Set frm = New frmAddress
    Set oAdd = oSupp.Addresses.Add
    frm.component oAdd
    frm.Show vbModal
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdAddDeal_Click()
    On Error GoTo errHandler
Dim frm As frmDeal
Dim oDeal As a_Deal
    Set frm = New frmDeal
    Set oDeal = oSupp.Deals.Add
    frm.component oDeal
    frm.Show vbModal
    LoadDeals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdAddDeal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdKeep_Click()
    On Error GoTo errHandler
Dim frmS As frmBrowseSUppliers2
    Set frmS = New frmBrowseSUppliers2
    frmS.Show vbModal
    txtParent = frmS.SupplierName & " " & frmS.Accnum
    oSupp.ParentSupplierID = frmS.SupplierID
    'oSUpp. = frmS.SupplierID
    Unload frmS
    Set frmS = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdKeep_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo errHandler
    If lvwAddresses.SelectedItem Is Nothing Then Exit Sub
    oSupp.Addresses.Remove Val(lvwAddresses.SelectedItem.key)
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdRemoveDeal_Click()
    On Error GoTo errHandler
Dim ocOPS As New c_OrdersPerSupplier
Dim bRecsreturned As Boolean
    If lvwDeals.SelectedItem Is Nothing Then Exit Sub
    ocOPS.Load oSupp.ID, oSupp.Deals(lvwDeals.SelectedItem.key).ID
    If ocOPS.Count > 0 Then
        MsgBox "There are order lines with this deal. You cannot delete it.", vbInformation, "Action denied"
        Exit Sub
    End If
    Set ocOPS = Nothing
    oSupp.Deals.Remove lvwDeals.SelectedItem.key
    LoadDeals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdRemoveDeal_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub lvwAddresses_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.lvwAddresses_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwAddresses_DblClick()
    On Error GoTo errHandler
Dim frm As frmAddress
    If lvwAddresses.SelectedItem.Index > 0 Then
        Set frm = New frmAddress
        frm.component oSupp.Addresses.Item((lvwAddresses.SelectedItem.key))
        frm.Show vbModal
        LoadAddresses
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.lvwAddresses_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDeals_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.lvwDeals_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDeals_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.lvwDeals_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDeals_DblClick()
    On Error GoTo errHandler
Dim frm As frmDeal
    If lvwDeals.SelectedItem.Index > 0 Then
        Set frm = New frmDeal
        frm.component oSupp.Deals.Item(Val(lvwDeals.SelectedItem.key))
        frm.Show vbModal
        LoadDeals
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.lvwDeals_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDel_Click()
    On Error GoTo errHandler
Dim ocOPS As New c_OrdersPerSupplier
Dim bRecsreturned As Boolean
    ocOPS.Load oSupp.ID
    If ocOPS.Count > 0 Then
        MsgBox "There are orders stored for this supplier. You cannot delete it.", vbInformation, "Action denied"
        Exit Sub
    End If
    Set ocOPS = Nothing
    Me.LinCancel.Visible = True
    oSupp.DeleteSupplier
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.mnuDel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oSupp_BillToADdressChanged()
    On Error GoTo errHandler
    txtPhone = oSupp.billtoaddress.Phone
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.oSupp_BillToADdressChanged", , EA_NORERAISE
    HandleError
End Sub
Public Sub component(pSupp As a_Supplier)
    On Error GoTo errHandler
    Set oSupp = pSupp
    oSupp.BeginEdit
    Me.Caption = "Supplier master edit: " & oSupp.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.component(pSupp)", pSupp
End Sub
Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.EnableOK(pOK)", pOK
End Sub


Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oSupp.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdApproAddress_Click()
    On Error GoTo errHandler
    oSupp.SetApproAddressidx (lvwAddresses.SelectedItem.key)
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdApproAddress_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbBillTo_Click()
    On Error GoTo errHandler
    oSupp.SetBillToAddressidx (lvwAddresses.SelectedItem.key)
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cbBillTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbDelTo_Click()
    On Error GoTo errHandler
    oSupp.SetDelToAddressidx (lvwAddresses.SelectedItem.key)
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cbDelTo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cbOrderTo_Click()
    On Error GoTo errHandler
    oSupp.SetOrderToAddressidx (lvwAddresses.SelectedItem.key)
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cbOrderTo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmAddress
    If lvwAddresses.SelectedItem Is Nothing Then Exit Sub
    Set frm = New frmAddress
    frm.component oSupp.Addresses.Item(lvwAddresses.SelectedItem.key)
    frm.Show vbModal
    LoadAddresses
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEditDeal_Click()
    On Error GoTo errHandler
Dim frm As frmDeal
    If lvwDeals.SelectedItem Is Nothing Then Exit Sub
    Set frm = New frmDeal
    frm.component oSupp.Deals.Item(Val(lvwDeals.SelectedItem.key))
    frm.Show vbModal
    LoadDeals
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdEditDeal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long


    If oSupp.SupplierIndexClashes = True Then
        MsgBox "This account number has already been used for another supplier. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If

    oSupp.ApplyEdit lngResult
    If lngResult = 0 Then
        Unload Me
    ElseIf lngResult = 22 Then
        MsgBox "You are trying to save a supplier with duplicate values." & vbCrLf & "These are likely to be in the Acc No. field or in the address description fields.", , "Can't save"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCurrencyConversionRate_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtCurrencyConversionRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtCurrencyConversionRate_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCurrencyConversionRate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
   ' If CInt(txtCurrencyConversionRate) <> oSupp.ConversionToLocalFactor Then
        Cancel = Not oSupp.SetConversionToLocalFactor(Trim(txtCurrencyConversionRate))
  '  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtCurrencyConversionRate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
Dim i As Integer

    On Error GoTo errHandler
    flgLoading = True
    If Me.WindowState <> 2 Then
        Top = 150
        Left = 50
        Height = 5300
        Width = 10300
    End If
    lblCurrencyConversionRate.Visible = oPC.SupplierBasedCurrencyConversion
    txtCurrencyConversionRate.Visible = oPC.SupplierBasedCurrencyConversion
    txtCurrencyConversionRate = oSupp.ConversionToLocalFactorF
    txtName = oSupp.Name
    txtAcno = oSupp.AcNo
    txtNote = oSupp.Note
    txtRecordAdded = oSupp.DateRecordAddedF
    txtRecordLastChanged = oSupp.DateRecordLastChangedF
    txtDefaultDeliverydays = oSupp.DefaultETA
    txtParent = oSupp.ParentSupplierName
    chkVATable = IIf(oSupp.VATable, 1, 0)
    chkDistributorOnly = IIf(oSupp.DoNotOrderFrom, 1, 0)
    chkActive = IIf(oSupp.UseStatus = 0, 0, 1)
    Select Case oSupp.DispatchMethod
    Case "E"
        optEDI = True
    Case "M"
        optEmail = True
    Case "P"
        optFaxManual = True
    End Select
    optEDI.Enabled = oPC.EDIEnabled
    txtReturnStartMonths = oSupp.ReturnStartMonths
    txtReturnEndMonths = oSupp.ReturnEndMonths
'    Me.txtEDINumber = oSupp.GFXNumber
    Me.SSTab1.Tab = 0
    If Not oSupp.billtoaddress Is Nothing Then
        txtPhone = oSupp.billtoaddress.PhoneandFax
    End If
'    Me.txtEDIType = oSupp.EDIType
'    txtPOFTPAddress = oSupp.PO_FTPAddress
'    txtPOFTPUser = oSupp.PO_FTPUser
'    txtPOFTPPassword = oSupp.PO_FTPPassword
'    txtPOFTPFolder = oSupp.PO_FTPFolder
'    txtINVFTPAddress = oSupp.INV_FTPAddress
'    txtINVFTPUser = oSupp.INV_FTPUser
'    txtINVFTPPassword = oSupp.INV_FTPPassword
'    txtINVFTPFolder = oSupp.INV_FTPFolder
    
    LoadAddresses
    LoadDeals
    SetLvw
    LoadCurrs
    LoadDispatchModes
    flgLoading = False
    
    cboTermType.List(0) = "end of month"
    cboTermType.ItemData(0) = 1
    cboTermType.List(1) = "date of invoice"
    cboTermType.ItemData(1) = 2
    
    
    For i = 0 To 29
        cboSettlementDiscountTerms.List(i) = i + 1
        cboSettlementDiscountTerms.ItemData(i) = i + 1
    Next
  
    cboSettlementDiscountTermsType.List(0) = "date of invoice"
    cboSettlementDiscountTermsType.ItemData(0) = 1
    cboSettlementDiscountTermsType.List(1) = "receipt of goods"
    cboSettlementDiscountTermsType.ItemData(1) = 2
  
    cboTerms = oSupp.TermsF
    If oSupp.TermsTypeF > "" Then
        cboTermType = oSupp.TermsTypeF
    End If
    txtSettlementDiscount = oSupp.SettlementDiscountF
    If oSupp.SettlementTerms > 0 Then
        cboSettlementDiscountTerms = CStr(oSupp.SettlementTerms)
    End If
    If oSupp.SettlementTermsTypeF > "" Then
        cboSettlementDiscountTermsType = oSupp.SettlementTermsTypeF
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadAddresses()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String
Dim lngSelectedItemKey As Long
    If Not lvwAddresses.SelectedItem Is Nothing Then
        lngSelectedItemKey = lvwAddresses.SelectedItem.Index
    End If
    'lvwAddresses.SelectedItem
    lvwAddresses.ListItems.Clear
    For i = 1 To oSupp.Addresses.Count
        Set objItm = Me.lvwAddresses.ListItems.Add
        With objItm
            .key = oSupp.Addresses(i).key  'i & "k" 'oSupp.Addresses(i).ID & "K"
            .Text = oSupp.Addresses(i).Addressee
            .SubItems(1) = oSupp.Addresses(i).Line1
            .SubItems(2) = oSupp.Addresses(i).Phone
            .SubItems(3) = IIf(oSupp.Addresses(i).Appro, "App", "") & IIf(oSupp.Addresses(i).BillTo, " Bill", "") & IIf(oSupp.Addresses(i).DelTo, " Del", "") & IIf(oSupp.Addresses(i).OrderTo, " Order", "")                  'IIf(oCust.BillToADdressIdx = i, "Default", "")
        End With
    Next i
    If lngSelectedItemKey > 0 Then lvwAddresses.ListItems.Item(lngSelectedItemKey).Selected = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.LoadAddresses"
End Sub
Private Sub LoadDeals()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwDeals.ListItems.Clear
    For i = 1 To oSupp.Deals.Count
        Set objItm = Me.lvwDeals.ListItems.Add
        With objItm
            .key = oSupp.Deals(i).key 'i & "k" 'oSupp.Addresses(i).ID & "K"
            .Text = oSupp.Deals(i).Description
            .SubItems(1) = oSupp.Deals(i).DiscountF
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.LoadDeals"
End Sub
Private Sub lvwAddresses_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.lvwAddresses_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub optEDI_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oSupp.SetDispatchMethod "E"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.optEDI_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optEmail_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oSupp.SetDispatchMethod "M"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.optEmail_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optFaxManual_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oSupp.SetDispatchMethod "P"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.optFaxManual_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oSupp_Valid(strMsg As String)
    On Error GoTo errHandler
    EnableOK (strMsg = "")
    lblErrors.Caption = strMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.oSupp_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub


Private Sub txtDefaultDeliverydays_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oSupp.SetDefaultETA(txtDefaultDeliverydays)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtDefaultDeliverydays_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtEDIType_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.setEDITYPE (Me.txtEDIType)
'    If Err Then
'      Beep
'      intPos = txtEDIType.SelStart
'      txtEDIType = oSupp.EDIType
'      txtEDIType.SelStart = intPos - 1
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtEDIType_Change", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtEDIType_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.setEDITYPE(txtEDIType)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtEDIType_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub

Private Sub txtName_LostFocus()
    On Error GoTo errHandler
    txtName = oSupp.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtName_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oSupp.SetName (txtName)
    If Err Then
      Beep
      intPos = txtName.SelStart
      txtName = oSupp.Name
      txtName.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtName_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtName_Validate(Cancel As Boolean)
    On Error GoTo errHandler
   ' Cancel = Not oSupp.SetName(txtName)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtName_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'Private Sub txtEDINumber_LostFocus()
'    On Error GoTo errHandler
'    txtEDINumber = oSupp.GFXNumber
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtEDINumber_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtEDINumber_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.setGFXNumber (txtEDINumber)
'    If Err Then
'      Beep
'      intPos = txtEDINumber.SelStart
'      txtEDINumber = oSupp.GFXNumber
'      txtEDINumber.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtEDINumber_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtEDINumber_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.setGFXNumber(txtEDINumber)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtEDINumber_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'
Private Sub txtAcno_LostFocus()
    On Error GoTo errHandler
    txtAcno = oSupp.AcNo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtAcno_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oSupp.SetAcNO (txtAcno)
    
    If Err Then
      Beep
      intPos = txtAcno.SelStart
      txtAcno = oSupp.AcNo
      txtAcno.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtAcno_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAcno_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oSupp.SetAcNO txtAcno
    If oSupp.SupplierIndexClashes = True Then
        MsgBox "This account number has already been used for another supplier. This record cannot be saved.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtAcno_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'Private Sub txtPhone_Change()
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetPhone txtPhone
'    If Err Then
'      Beep
'      intPos = txtPhone.SelStart
'      txtPhone = oSupp.Phone
'      txtPhone.SelStart = intPos - 1
'    End If
'End Sub
'Private Sub txtPhone_Validate(Cancel As Boolean)
'    Cancel = Not oSupp.SetPhone(txtPhone)
'End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oSupp.setnote(txtNote)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    txtNote = oSupp.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    oSupp.setnote (txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oSupp.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetLvw()
    On Error GoTo errHandler
Dim style As Long
Dim hHeader As Long
   
'  'get the handle to the listview header
'   hHeader = SendMessage(lvwAddresses.hwnd, LVM_GETHEADER, 0, ByVal 0&)
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
'      Call SetWindowPos(lvwAddresses.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If
'
'   hHeader = SendMessage(lvwDeals.hwnd, LVM_GETHEADER, 0, ByVal 0&)
'   style = GetWindowLong(hHeader, GWL_STYLE)
'   style = style Xor HDS_BUTTONS
'   If style Then
'      Call SetWindowLong(hHeader, GWL_STYLE, style)
'      Call SetWindowPos(lvwDeals.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.SetLvw"
End Sub



Private Sub txtReturnStartMonths_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oSupp.SetReturnStartMonths(txtReturnStartMonths)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtReturnStartMonths_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtReturnEndMonths_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oSupp.SetReturnEndMonths(txtReturnEndMonths)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplier.txtReturnEndMonths_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtPOFTPAddress_LostFocus()
'    On Error GoTo errHandler
'    txtPOFTPAddress = oSupp.PO_FTPAddress
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPAddress_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPAddress_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetPO_FTPAddress (txtPOFTPAddress)
'    If Err Then
'        Beep
'        intPos = txtPOFTPAddress.SelStart
'        txtPOFTPAddress = oSupp.PO_FTPAddress
'        txtPOFTPAddress.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPAddress_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPAddress_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetPO_FTPAddress(txtPOFTPAddress)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPAddress_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtPOFTPUser_LostFocus()
'    On Error GoTo errHandler
'    txtPOFTPUser = oSupp.PO_FTPUser
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPUser_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPUser_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetPO_FTPUser (txtPOFTPUser)
'    If Err Then
'        Beep
'        intPos = txtPOFTPUser.SelStart
'        txtPOFTPUser = oSupp.PO_FTPUser
'        txtPOFTPUser.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPUser_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPUser_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetPO_FTPUser(txtPOFTPUser)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPUser_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPPassword_LostFocus()
'    On Error GoTo errHandler
'    txtPOFTPPassword = oSupp.PO_FTPPassword
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPPassword_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPPassword_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetPO_FTPPassword (Me.txtPOFTPPassword)
'    If Err Then
'        Beep
'        intPos = txtPOFTPPassword.SelStart
'        txtPOFTPPassword = oSupp.PO_FTPPassword
'        txtPOFTPPassword.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPPassword_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPPassword_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetPO_FTPPassword(txtPOFTPPassword)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPPassword_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtPOFTPFolder_LostFocus()
'    On Error GoTo errHandler
'    txtPOFTPFolder = oSupp.PO_FTPFolder
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPFolder_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtPOFTPFolder_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetPO_FTPFolder (Me.txtPOFTPFolder)
'    If Err Then
'        Beep
'        intPos = txtPOFTPFolder.SelStart
'        txtPOFTPFolder = oSupp.PO_FTPFolder
'        txtPOFTPFolder.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtPOFTPFolder_Change", , EA_NORERAISE
'    HandleError
'End Sub
'
'
'
'
'Private Sub txtINVFTPAddress_LostFocus()
'    On Error GoTo errHandler
'    txtINVFTPAddress = oSupp.INV_FTPAddress
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPAddress_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPAddress_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetINV_FTPAddress (txtINVFTPAddress)
'    If Err Then
'        Beep
'        intPos = txtINVFTPAddress.SelStart
'        txtINVFTPAddress = oSupp.INV_FTPAddress
'        txtINVFTPAddress.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPAddress_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPAddress_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetINV_FTPAddress(txtINVFTPAddress)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPAddress_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtINVFTPUser_LostFocus()
'    On Error GoTo errHandler
'    txtINVFTPUser = oSupp.INV_FTPUser
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPUser_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPUser_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetINV_FTPUser (txtINVFTPUser)
'    If Err Then
'        Beep
'        intPos = txtINVFTPUser.SelStart
'        txtINVFTPUser = oSupp.INV_FTPUser
'        txtINVFTPUser.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPUser_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPUser_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetINV_FTPUser(txtINVFTPUser)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPUser_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtINVFTPPassword_LostFocus()
'    On Error GoTo errHandler
'    txtINVFTPPassword = oSupp.INV_FTPPassword
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPPassword_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPPassword_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetINV_FTPPassword (txtINVFTPPassword)
'    If Err Then
'        Beep
'        intPos = txtINVFTPPassword.SelStart
'        txtINVFTPPassword = oSupp.INV_FTPPassword
'        txtINVFTPPassword.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPPassword_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPPassword_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetINV_FTPPassword(txtINVFTPPassword)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPPassword_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub txtINVFTPFolder_LostFocus()
'    On Error GoTo errHandler
'    txtINVFTPFolder = oSupp.INV_FTPFolder
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPFolder_LostFocus", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPFolder_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'    On Error Resume Next
'    oSupp.SetINV_FTPFolder (txtINVFTPFolder)
'    If Err Then
'        Beep
'        intPos = txtINVFTPFolder.SelStart
'        txtINVFTPFolder = oSupp.INV_FTPFolder
'        txtINVFTPFolder.SelStart = intPos - 1
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPFolder_Change", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub txtINVFTPFolder_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    Cancel = Not oSupp.SetINV_FTPFolder(txtINVFTPFolder)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplier.txtINVFTPFolder_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub
'

Private Sub txtSettlementDiscount_Change()

End Sub

Private Sub txtSettlementDiscount_Validate(Cancel As Boolean)
Dim bOK As Boolean
    
    bOK = oSupp.SetSettlementDiscount(txtSettlementDiscount)
    Cancel = Not bOK
End Sub
