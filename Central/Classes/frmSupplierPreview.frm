VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupplierPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Supplier"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   ControlBox      =   0   'False
   Icon            =   "frmSupplierPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   10335
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
      Left            =   8070
      Picture         =   "frmSupplierPreview.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4095
      Width           =   1000
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
      Left            =   9090
      Picture         =   "frmSupplierPreview.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4095
      Width           =   1000
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
      Left            =   7065
      Picture         =   "frmSupplierPreview.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4095
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   1005
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   13882315
      ForeColor       =   8388608
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
      TabPicture(0)   =   "frmSupplierPreview.frx":0E28
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtDispatchMode"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "txtCurrency"
      Tab(0).Control(3)=   "lvwDeals"
      Tab(0).Control(4)=   "Label14"
      Tab(0).Control(5)=   "Label10"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&2. Addresses"
      TabPicture(1)   =   "frmSupplierPreview.frx":0E44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwAddresses"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&3. Recent orders"
      TabPicture(2)   =   "frmSupplierPreview.frx":0E60
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwOrders"
      Tab(2).Control(1)=   "cmdShowOrders"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Miscellaneous"
      TabPicture(3)   =   "frmSupplierPreview.frx":0E7C
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label8"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label5"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label4"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label7"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label6"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label11"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label12"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblWeReceiveFrom"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblPaysVAT"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lblInactive"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label13"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "lblCurrencyConversionRate"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label15"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "txtRecordAdded"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtRecordLastChanged"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "txtDefaultDeliverydays"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "chkActive"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "chkVATable"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "txtReturnEndMonths"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "txtReturnStartMonths"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "txtEDINumber"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "txtParent"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "txtEDIType"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "txtCurrencyConversionRate"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "cmdRecalc"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "txtFTPAddress"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).ControlCount=   27
      TabCaption(4)   =   "Note"
      TabPicture(4)   =   "frmSupplierPreview.frx":0E98
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtNote"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtFTPAddress 
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
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2520
      End
      Begin VB.TextBox txtDispatchMode 
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
         Left            =   -69330
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdRecalc 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Recalc all product prices"
         Height          =   300
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1845
         Width           =   2220
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
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   43
         Top             =   1830
         Width           =   540
      End
      Begin VB.TextBox txtEDIType 
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
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   885
         Width           =   705
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
         Left            =   7200
         TabIndex        =   33
         Top             =   585
         Width           =   2520
      End
      Begin VB.Frame Frame1 
         Caption         =   "Document delivery method"
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
         Height          =   1515
         Left            =   -71070
         TabIndex        =   29
         Top             =   915
         Width           =   2385
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
            TabIndex        =   32
            Top             =   375
            Value           =   -1  'True
            Width           =   1665
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
            TabIndex        =   31
            Top             =   990
            Width           =   930
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
            Left            =   555
            TabIndex        =   30
            Top             =   683
            Width           =   960
         End
      End
      Begin VB.TextBox txtEDINumber 
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
         Left            =   3810
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   585
         Width           =   1875
      End
      Begin VB.TextBox txtReturnStartMonths 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   4665
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1500
         Width           =   435
      End
      Begin VB.TextBox txtReturnEndMonths 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   3330
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1500
         Width           =   435
      End
      Begin VB.CheckBox chkVATable 
         Caption         =   "Pays V.A.T."
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
         Height          =   465
         Left            =   3405
         TabIndex        =   21
         Top             =   2145
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txtNote 
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
         Height          =   2340
         Left            =   -74820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   420
         Width           =   6420
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Inactive"
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
         Height          =   360
         Left            =   3420
         TabIndex        =   19
         Top             =   2535
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txtDefaultDeliverydays 
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   585
         Width           =   690
      End
      Begin VB.TextBox txtCurrency 
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
         Left            =   -69330
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   1335
      End
      Begin VB.TextBox txtRecordLastChanged 
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
         Height          =   285
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2535
         Width           =   1815
      End
      Begin VB.TextBox txtRecordAdded 
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
         Height          =   285
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1815
      End
      Begin VB.CommandButton cmdShowOrders 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Showorders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   390
         Width           =   2040
      End
      Begin MSComctlLib.ListView lvwOrders 
         CausesValidation=   0   'False
         Height          =   1920
         Left            =   -74790
         TabIndex        =   7
         Top             =   750
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   3387
         SortKey         =   4
         View            =   3
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date sold"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Title"
            Object.Width           =   7303
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   " "
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lvwAddresses 
         Height          =   2190
         Left            =   -74730
         TabIndex        =   13
         Top             =   465
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   3863
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Address"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Phone"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fax"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Purpose"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Mail by"
            Object.Width           =   1781
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDeals 
         Height          =   2190
         Left            =   -74760
         TabIndex        =   14
         Top             =   480
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   3863
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
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Discount"
            Object.Width           =   2824
         EndProperty
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FTP address"
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
         Left            =   5955
         TabIndex        =   49
         Top             =   1350
         Width           =   1155
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
         Left            =   -71025
         TabIndex        =   47
         Top             =   2550
         Width           =   1620
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
         Left            =   225
         TabIndex        =   44
         Top             =   1890
         Width           =   3285
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EDI Type"
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
         Left            =   2550
         TabIndex        =   42
         Top             =   945
         Width           =   1155
      End
      Begin VB.Label lblInactive 
         Caption         =   "Inactive"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   240
         TabIndex        =   40
         Top             =   2580
         Width           =   2535
      End
      Begin VB.Label lblPaysVAT 
         Caption         =   "Charges V.A.T."
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   240
         TabIndex        =   39
         Top             =   2250
         Width           =   2385
      End
      Begin VB.Label lblWeReceiveFrom 
         Alignment       =   1  'Right Justify
         Caption         =   "**We receive from but do not order from this supplier.**"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   5370
         TabIndex        =   38
         Top             =   975
         Width           =   4080
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
         Left            =   6045
         TabIndex        =   34
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EDI Number"
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
         TabIndex        =   28
         Top             =   630
         Width           =   1155
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
         Left            =   225
         TabIndex        =   26
         Top             =   1545
         Width           =   3075
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
         Left            =   3900
         TabIndex        =   25
         Top             =   1545
         Width           =   600
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "months prior"
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
         Left            =   5175
         TabIndex        =   24
         Top             =   1545
         Width           =   1080
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
         Left            =   180
         TabIndex        =   18
         Top             =   630
         Width           =   1620
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
         Left            =   -71010
         TabIndex        =   15
         Top             =   570
         Width           =   1605
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
         Left            =   5790
         TabIndex        =   12
         Top             =   2565
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
         Left            =   6285
         TabIndex        =   11
         Top             =   2250
         Width           =   1305
      End
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
      Height          =   330
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   945
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
      Top             =   525
      Width           =   4980
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
      Width           =   4980
   End
   Begin VB.Line LinCancel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2100
      X2              =   -60
      Y1              =   60
      Y2              =   720
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
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier code"
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
      Left            =   6135
      TabIndex        =   4
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   60
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmSupplierPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSupp As a_Supplier
Dim frmSP As frmSupplier

Public Sub component(pSupp As a_Supplier)
    On Error GoTo errHandler
    Set oSupp = pSupp
    Me.Caption = "Supplier master preview: " & oSupp.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.component(pSupp)", pSupp
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim ocOPS As New c_OrdersPerSupplier
Dim bRecsreturned As Boolean
Dim iRes As Long

    If MsgBox("Confirm you wish to delete supplier: " & oSupp.NameAndCode(35), vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    MsgBox "Warning - code skipped here"
'    ocOPS.Load oSupp.ID
'    If ocOPS.Count > 0 Then
'        MsgBox "There are orders stored for this supplier. You cannot delete it.", vbInformation, "Action denied"
'        Exit Sub
'    End If
    Set ocOPS = Nothing
    Me.LinCancel.Visible = True
    oSupp.BeginEdit
    oSupp.DeleteSupplier
    oSupp.ApplyEdit iRes
    MsgBox "Supplier has been deleted. Form will close", vbInformation + vbOKOnly, "status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean

    If frmSP Is Nothing Then
        Set frmSP = New frmSupplier
    End If
    blnEdit = True
    frmSP.component oSupp ', lngID
    frmSP.Show
    
EXIT_Handler:
    Unload Me
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRecalc_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    oSQL.RunProc "UpdateLocalPricesOfDistributor", Array(oSupp.ID), ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.cmdRecalc_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdShowOrders_Click()
    On Error GoTo errHandler
Dim oSPS As c_OrdersPerSupplier
    Screen.MousePointer = vbHourglass
    Set oSPS = New c_OrdersPerSupplier
    oSPS.Load oSupp.ID
    LoadOrders oSPS
    Set oSPS = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.cmdShowOrders_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.top = 50
        Me.left = 50
        Me.Height = 5300
        Me.Width = 10300
    End If
'    lblCurrencyConversionRate.Visible = oPC.SupplierBasedCurrencyConversion
'    txtCurrencyConversionRate.Visible = oPC.SupplierBasedCurrencyConversion
'    cmdRecalc.Visible = oPC.SupplierBasedCurrencyConversion
'    txtCurrencyConversionRate = oSupp.ConversionToLocalFactorF
    
    txtName = oSupp.Name
    txtRecordAdded = oSupp.DateRecordAddedF
    txtRecordLastChanged = oSupp.DateRecordLastChangedF
    txtAcno = oSupp.AcNo
    txtCurrency = oSupp.DefaultCurrency.Description
    txtDefaultDeliverydays = oSupp.DefaultETA
    txtReturnStartMonths = oSupp.ReturnStartMonths
    txtReturnEndMonths = oSupp.ReturnEndMonths
    txtNote = oSupp.Note
    Me.txtParent = oSupp.ParentSupplierName
    txtFTPAddress = oSupp.FTPAddress
    txtEDINumber = oSupp.GFXNumber
    txtEDIType = oSupp.EDIType
    txtDispatchMode = oSupp.DispatchMode
  '  chkVATable = IIf(oSupp.Vatable, 1, 0)
  '  chkActive = IIf(oSupp.UseStatus = 0, 0, 1)
    lblWeReceiveFrom.Visible = oSupp.DoNotOrderFrom
    If oSupp.UseStatus = 0 Then
        lblInactive.Caption = "Active"
    Else
        lblInactive.Caption = "Inactive"
    End If
    If oSupp.VATable Then
        lblPaysVAT.Caption = "Charges V.A.T."
    Else
        lblPaysVAT.Caption = "Does not charge V.A.T."
    End If
    Select Case oSupp.DispatchMethod
    Case "E"
        optEDI = True
    Case "M"
        optEmail = True
    Case "P"
        optFaxManual = True
    End Select
    LoadAddresses
    LoadDeals
   ' SetLvw
    Me.SSTab1.Tab = 0
    If Not oSupp.Addresses.DefaultAddress Is Nothing Then
        txtPhone = oSupp.Addresses.DefaultAddress.PhoneandFax
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadAddresses()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwAddresses.ListItems.Clear
    For i = 1 To oSupp.Addresses.Count
        Set objItm = Me.lvwAddresses.ListItems.Add
        With objItm
            .Key = oSupp.Addresses(i).Key
            .Text = oSupp.Addresses(i).Addressee
            .SubItems(1) = oSupp.Addresses(i).Line1
            .SubItems(2) = oSupp.Addresses(i).Phone
            .SubItems(3) = oSupp.Addresses(i).Fax
            .SubItems(4) = IIf(oSupp.Addresses(i).Appro, "App", "") & IIf(oSupp.Addresses(i).BillTo, " Bill", "") & IIf(oSupp.Addresses(i).DelTo, " Del", "") & IIf(oSupp.Addresses(i).OrderTo, " Order", "")                  'IIf(oCust.DefaultAddressIdx = i, "Default", "")
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.LoadAddresses"
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
            .Key = oSupp.Deals.Item(i).Key
            .Text = oSupp.Deals(i).Description
            .SubItems(1) = oSupp.Deals(i).DiscountF
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.LoadDeals"
End Sub

Private Sub LoadOrders(oCP As c_OrdersPerSupplier)
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwOrders.ListItems.Clear
    For i = 1 To oCP.Count
        Set objItm = Me.lvwOrders.ListItems.Add
        With objItm
            .Text = oCP(i).dateOfOrder
            .SubItems(1) = oCP(i).code
            .SubItems(2) = oCP(i).Title
            .SubItems(3) = oCP(i).Price
            .SubItems(4) = oCP(i).dateOfOrderForSort
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.LoadOrders(oCP)", oCP
End Sub


Private Sub lvwAddresses_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwAddresses_AfterLabelEdit(Cancel,NewString)", Array(Cancel, _
         NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwAddresses_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwAddresses_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwAddresses_DblClick()
    On Error GoTo errHandler
Dim frm As frmAddressPreview
    If lvwAddresses.SelectedItem.Index > 0 Then
        Set frm = New frmAddressPreview
        frm.component oSupp.Addresses.Item(lvwAddresses.SelectedItem.Key)
        frm.Show vbModal
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwAddresses_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwPurchases_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwPurchases_AfterLabelEdit(Cancel,NewString)", Array(Cancel, _
         NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwPurchases_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwPurchases_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'Private Sub SetLvw()
'    On Error GoTo errHandler
'Dim style As Long
'Dim hHeader As Long
'
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
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSupplierPreview.SetLvw"
'End Sub

Private Sub lvwDeals_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwDeals_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDeals_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.lvwDeals_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub Form_DblClick()
    On Error GoTo errHandler

    If Not IsNull(oSupp.billtoaddress) Then
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText oSupp.billtoaddress.AddressMailing
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSupplierPreview.Form_DblClick", , EA_NORERAISE
    HandleError
End Sub

