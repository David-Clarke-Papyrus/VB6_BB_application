VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProductAQ 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Stock"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   ControlBox      =   0   'False
   Icon            =   "frmProductObs.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleMode       =   0  'User
   ScaleWidth      =   11714.04
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C4BCA4&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   8100
      Picture         =   "frmProductObs.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4665
      Width           =   1050
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   5505
      Width           =   4350
   End
   Begin VB.CommandButton cmdNewCode 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&New code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   45
      TabIndex        =   19
      Top             =   2385
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1. Stock"
      TabPicture(0)   =   "frmProductObs.frx":0614
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(6)=   "Label21"
      Tab(0).Control(7)=   "Label31"
      Tab(0).Control(8)=   "Label32"
      Tab(0).Control(9)=   "txtOnHand"
      Tab(0).Control(10)=   "txtReserved"
      Tab(0).Control(11)=   "txtPreOwned"
      Tab(0).Control(12)=   "txtRRP"
      Tab(0).Control(13)=   "txtSP"
      Tab(0).Control(14)=   "txtCost"
      Tab(0).Control(15)=   "txtSpecial"
      Tab(0).Control(16)=   "txtUKPrice"
      Tab(0).Control(17)=   "txtUSPrice"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "&2. Pre-owned copies"
      TabPicture(1)   =   "frmProductObs.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwCopies"
      Tab(1).Control(1)=   "cmdRemove"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAddCopy"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "&3. Details"
      TabPicture(2)   =   "frmProductObs.frx":064C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label20"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label24"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label25"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtVAT"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdSetDefault"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkNonStock"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chkObsolete"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboCatHead"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cboCategory"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtBinding"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtBIC"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtSeriesTitle"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "&4. Notes"
      TabPicture(3)   =   "frmProductObs.frx":0668
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtComment"
      Tab(3).Control(1)=   "txtSummary"
      Tab(3).Control(2)=   "txtDescription"
      Tab(3).Control(3)=   "txtNote"
      Tab(3).Control(4)=   "Label30"
      Tab(3).Control(5)=   "Label29"
      Tab(3).Control(6)=   "Label28"
      Tab(3).Control(7)=   "Label27"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "&5. Statistics"
      TabPicture(4)   =   "frmProductObs.frx":0684
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).Control(1)=   "txtLastDeliveredPrice"
      Tab(4).Control(2)=   "txtLastOrderedPrice"
      Tab(4).Control(3)=   "txtLastOrderedQty"
      Tab(4).Control(4)=   "txtLastDeliveredQty"
      Tab(4).Control(5)=   "txtLastOrdered"
      Tab(4).Control(6)=   "txtLastDelivered"
      Tab(4).Control(7)=   "txtDateKastModified"
      Tab(4).Control(8)=   "txtDateAdded"
      Tab(4).Control(9)=   "Label23"
      Tab(4).Control(10)=   "Label22"
      Tab(4).Control(11)=   "Label12"
      Tab(4).Control(12)=   "Label11"
      Tab(4).ControlCount=   13
      Begin VB.Frame Frame1 
         Caption         =   "Aged stock"
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
         Height          =   915
         Left            =   -74745
         TabIndex        =   80
         Top             =   1845
         Width           =   5220
         Begin VB.TextBox txt18Plus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   375
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   85
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   375
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   375
            Left            =   2415
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   375
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txtAgedDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   375
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   330
            Width           =   1170
         End
      End
      Begin VB.TextBox txtSeriesTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   1380
         TabIndex        =   79
         Top             =   465
         Width           =   5925
      End
      Begin VB.TextBox txtBIC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   2040
         TabIndex        =   78
         Top             =   1440
         Width           =   1380
      End
      Begin VB.TextBox txtBinding 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   2025
         TabIndex        =   77
         Top             =   1935
         Width           =   1395
      End
      Begin VB.TextBox txtUSPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -68865
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   1035
         Width           =   1380
      End
      Begin VB.TextBox txtUKPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -68865
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   525
         Width           =   1380
      End
      Begin VB.TextBox txtComment 
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
         Height          =   555
         Left            =   -73725
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   1725
         Width           =   6525
      End
      Begin VB.TextBox txtSummary 
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
         Height          =   555
         Left            =   -73725
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   450
         Width           =   6525
      End
      Begin VB.TextBox txtDescription 
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
         Height          =   555
         Left            =   -73725
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   1080
         Width           =   6525
      End
      Begin VB.ComboBox cboCategory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5910
         TabIndex        =   64
         Top             =   990
         Width           =   1395
      End
      Begin VB.TextBox txtLastDeliveredPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -68055
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   570
         Width           =   960
      End
      Begin VB.TextBox txtLastOrderedPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -68055
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   1080
         Width           =   960
      End
      Begin VB.TextBox txtLastOrderedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -68670
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtLastDeliveredQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -68670
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   570
         Width           =   555
      End
      Begin VB.TextBox txtLastOrdered 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -70110
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1080
         Width           =   1380
      End
      Begin VB.TextBox txtLastDelivered 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -69810
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   570
         Width           =   1380
      End
      Begin VB.TextBox txtDateKastModified 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1050
         Width           =   1815
      End
      Begin VB.TextBox txtDateAdded 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtSpecial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   -71595
         TabIndex        =   48
         Top             =   2040
         Width           =   1380
      End
      Begin VB.TextBox txtCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   -71595
         TabIndex        =   45
         Top             =   1525
         Width           =   1380
      End
      Begin VB.TextBox txtSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   -71595
         TabIndex        =   43
         Top             =   1010
         Width           =   1380
      End
      Begin VB.ComboBox cboCatHead 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2025
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   990
         Width           =   2520
      End
      Begin VB.CheckBox chkObsolete 
         Caption         =   "Obsolete"
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
         Height          =   480
         Left            =   6060
         TabIndex        =   40
         Top             =   2355
         Width           =   1245
      End
      Begin VB.TextBox txtRRP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   -71595
         TabIndex        =   38
         Top             =   495
         Width           =   1380
      End
      Begin VB.TextBox txtPreOwned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -73815
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1515
         Width           =   1380
      End
      Begin VB.TextBox txtReserved 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -73815
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1005
         Width           =   1380
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   375
         Left            =   -73815
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   495
         Width           =   1380
      End
      Begin VB.CommandButton cmdAddCopy 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&New copy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2325
         Width           =   1050
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C4BCA4&
         Cancel          =   -1  'True
         Caption         =   "&Remove selected"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -73665
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2325
         Width           =   1830
      End
      Begin VB.CheckBox chkNonStock 
         Caption         =   "Non-stock"
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
         Height          =   480
         Left            =   6060
         TabIndex        =   28
         Top             =   1905
         Width           =   1245
      End
      Begin VB.CommandButton cmdSetDefault 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Default V.A.T. rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2400
         Width           =   1755
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   1380
         TabIndex        =   25
         Top             =   2400
         Width           =   1380
      End
      Begin VB.TextBox txtNote 
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
         Height          =   555
         Left            =   -73725
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   2370
         Width           =   6525
      End
      Begin MSComctlLib.ListView lvwCopies 
         Height          =   1845
         Left            =   -74820
         TabIndex        =   31
         Top             =   465
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   3254
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Serial"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Price"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Purchased"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   5362
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Sold"
            Object.Width           =   2189
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "catalogues"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "U.S. Price"
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
         Height          =   285
         Left            =   -70275
         TabIndex        =   76
         Top             =   1095
         Width           =   1290
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "U.K. Price"
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
         Height          =   285
         Left            =   -70275
         TabIndex        =   74
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment"
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   72
         Top             =   1740
         Width           =   1035
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Summary"
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   70
         Top             =   465
         Width           =   1035
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   68
         Top             =   1095
         Width           =   1035
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Note"
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
         Height          =   285
         Left            =   -74670
         TabIndex        =   66
         Top             =   2385
         Width           =   825
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
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
         Height          =   285
         Left            =   4695
         TabIndex        =   65
         Top             =   1035
         Width           =   1080
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "B.I.C"
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
         Height          =   285
         Left            =   795
         TabIndex        =   63
         Top             =   1470
         Width           =   1080
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Series title"
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
         Height          =   285
         Left            =   165
         TabIndex        =   62
         Top             =   525
         Width           =   1080
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Last ordered"
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
         Height          =   285
         Left            =   -71640
         TabIndex        =   57
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Last delivered"
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
         Height          =   285
         Left            =   -71625
         TabIndex        =   55
         Top             =   615
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Last modified"
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
         Height          =   285
         Left            =   -74925
         TabIndex        =   53
         Top             =   1095
         Width           =   1245
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Added"
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
         Height          =   285
         Left            =   -74970
         TabIndex        =   52
         Top             =   630
         Width           =   1290
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Special"
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
         Height          =   285
         Left            =   -72390
         TabIndex        =   49
         Top             =   2100
         Width           =   750
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Binding"
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
         Height          =   285
         Left            =   840
         TabIndex        =   47
         Top             =   1980
         Width           =   1080
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost"
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
         Height          =   285
         Left            =   -72435
         TabIndex        =   46
         Top             =   1575
         Width           =   750
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "S.P."
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
         Height          =   285
         Left            =   -72390
         TabIndex        =   44
         Top             =   1065
         Width           =   750
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Catalogue heading"
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
         Height          =   285
         Left            =   135
         TabIndex        =   42
         Top             =   1035
         Width           =   1755
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "R.R.P."
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
         Height          =   285
         Left            =   -72390
         TabIndex        =   39
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Pre-owned"
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
         Height          =   285
         Left            =   -74940
         TabIndex        =   37
         Top             =   1575
         Width           =   1050
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Reserved"
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
         Height          =   285
         Left            =   -74940
         TabIndex        =   35
         Top             =   1065
         Width           =   1050
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "On hand"
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
         Height          =   285
         Left            =   -74940
         TabIndex        =   33
         Top             =   555
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "V.A.T. Rate"
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
         Height          =   285
         Left            =   150
         TabIndex        =   26
         Top             =   2430
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
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
      Height          =   690
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3420
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   8085
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2700
      Width           =   1050
   End
   Begin VB.TextBox txtPubPlace 
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
      Height          =   360
      Left            =   6420
      TabIndex        =   9
      Top             =   1905
      Width           =   2520
   End
   Begin VB.TextBox txtPubDate 
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
      Height          =   360
      Left            =   6420
      TabIndex        =   8
      Top             =   1485
      Width           =   2520
   End
   Begin VB.TextBox txtEdition 
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
      Height          =   360
      Left            =   6420
      TabIndex        =   6
      Top             =   645
      Width           =   2520
   End
   Begin VB.TextBox txtPublisher 
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
      Height          =   360
      Left            =   6420
      TabIndex        =   7
      Top             =   1065
      Width           =   2520
   End
   Begin VB.TextBox txtSubtitle 
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
      Height          =   585
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1290
      Width           =   3900
   End
   Begin VB.TextBox txtAuthor 
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
      Height          =   360
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   3915
   End
   Begin VB.TextBox txtTitle 
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
      Height          =   570
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   675
      Width           =   3900
   End
   Begin VB.TextBox txtEAN 
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
      Height          =   360
      Left            =   4380
      TabIndex        =   2
      Top             =   75
      Width           =   2790
   End
   Begin VB.TextBox txtCode 
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
      Height          =   360
      Left            =   750
      TabIndex        =   0
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Publication place"
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
      Height          =   255
      Left            =   4785
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Publication date"
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
      Height          =   255
      Left            =   4890
      TabIndex        =   20
      Top             =   1515
      Width           =   1470
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edition"
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
      Height          =   255
      Left            =   5715
      TabIndex        =   18
      Top             =   705
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Publisher"
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
      Height          =   255
      Left            =   5385
      TabIndex        =   17
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Author"
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
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   1965
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subtitle"
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
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   1305
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Title"
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
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   705
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "EAN"
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
      Height          =   255
      Left            =   3645
      TabIndex        =   13
      Top             =   135
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code"
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
      Height          =   255
      Left            =   30
      TabIndex        =   12
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmProductAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean
Dim tlCatHead As z_TextList
Dim mCancel As Boolean

Sub Component(pProduct As a_Product)
        Set oProd = pProduct
        oProd.beginEdit
        oProd.NonStock = False
        Set tlCatHead = New z_TextList
        tlCatHead.Load ltCatalogueHeadings
End Sub


Private Sub cboCatHead_Click()
    oProd.setCatalogueheadingID tlCatHead.Key(cboCatHead)
End Sub

Private Sub cmdAddCopy_Click()
Dim frm As frmCopy
Dim oCopy As a_Copy
    Set frm = New frmCopy
    Set oCopy = oProd.Copies.Add
    frm.Component oCopy
    frm.Show vbModal
    Set oCopy = Nothing
    Set frm = Nothing
    LoadCopies
End Sub


Private Sub cmdDelete_Click()

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


End Sub

Private Sub cmdRemove_Click()
Dim oCopy As a_Copy
    Set oCopy = oProd.Copies(lvwCopies.SelectedItem.Key)
    oCopy.beginEdit
    oCopy.Delete
    oCopy.ApplyEdit
    lvwCopies.SelectedItem.Text = lvwCopies.SelectedItem.Text & "(DEL)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If oProd.IsEditing Then oProd.CancelEdit
End Sub

Private Sub oProd_Valid(IsValid As Boolean, strMsg As String)
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = IsValid
    Me.cmdAddCopy.Enabled = IsValid
    Me.cmdRemove.Enabled = IsValid
End Sub
Private Sub oProd_CodeToBeGenerated()
    Me.txtEAN = ""
    Me.txtEAN.Enabled = False
End Sub
Private Sub cmdCancel_Click()
    oProd.CancelEdit
    Unload Me
End Sub

Private Sub cmdNewCode_Click()
    Me.txtCode = "#"
    oProd.SetCode "#"
End Sub

Private Sub cmdOK_Click()
Dim lngResult As Long
    oProd.ApplyEdit lngResult
    If lngResult = 99 Then
        MsgBox "Invalid values - check that the code is has not been already used", , "Save failed"
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Left = 10
    Top = 10
    Width = 9400
    Height = 6550
    LoadControls
    Me.cmdNewCode.Enabled = oProd.IsNew
End Sub
Private Sub LoadControls()
    flgLoading = True
    txtSeriesTitle = oProd.SeriesTitle
    txtSP = oProd.SP
    txtCode = oProd.Code
    txtEAN = oProd.EAN
    txttitle = oProd.Title
    txtsubtitle = oProd.SubTitle
    txtauthor = oProd.Author
    txtEdition = oProd.Edition
    txtpublisher = oProd.Publisher
    txtPubPlace = oProd.PublicationPlace
    Me.txtPubDate = oProd.PublicationDate
    Me.txtBinding = oProd.BindingCode
    Me.txtComment = oProd.Comment
    Me.txtDescription = oProd.Description
    Me.txtNote = oProd.Note
    Me.txtSummary = oProd.Summary
    Me.txtLastDeliveredQty = oProd.QtyLastDelivered
    Me.txtLastDeliveredPrice = oProd.PriceLastDeliveredF
    Me.txtLastOrderedQty = oProd.QtyLastOrdered
    Me.txtLastOrderedPrice = oProd.PriceLastOrderedF
    Me.txtRRP = oProd.RRPF
    Me.txtCost = oProd.CostF
    LoadCombo cboCatHead, tlCatHead
    LoadCopies
    flgLoading = False
End Sub
Private Sub LoadCopies()
Dim objItm As ListItem
Dim i, j As Integer
Dim tmp As String
Dim strCatalogues As String

    lvwCopies.ListItems.Clear
    For i = 1 To oProd.Copies.Count
        strCatalogues = ""
        For j = 1 To oProd.Copies(i).CatalogueEntries.Count
            strCatalogues = strCatalogues & oProd.Copies(i).CatalogueEntries(j).Serial
            If j < oProd.Copies(i).CatalogueEntries.Count Then strCatalogues = strCatalogues & ", "
        Next j
        Set objItm = Me.lvwCopies.ListItems.Add
        With objItm
            .Key = oProd.Copies(i).Key
            .Text = oProd.Copies(i).Serial
            .SubItems(1) = oProd.Copies(i).PriceF
            .SubItems(2) = oProd.Copies(i).PurchaseDateF
            .SubItems(3) = oProd.Copies(i).Comment
            .SubItems(4) = oProd.Copies(i).SoldDateF
            .SubItems(5) = strCatalogues
         '   .SubItems(3) = IIf(oCust.Addresses(i).ID = oCust.DefaultAddress.ID, "Default", "")
        End With
    Next i

End Sub
Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvwCopies_DblClick()
Dim frm As frmCopy
Dim oCopy As a_Copy
    Set oCopy = oProd.Copies(lvwCopies.SelectedItem.Key)
    Set frm = New frmCopy
    
    frm.Component oCopy
    frm.Show

End Sub


Private Sub txtCode_Validate(Cancel As Boolean)
    oProd.SetCode txtCode
End Sub


Private Sub txtSubtitle_LostFocus()
    If flgLoading Then Exit Sub
    txtsubtitle = oProd.SubTitle
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtSubtitle_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetSubTitle(txtsubtitle)
    If Err Then
      Beep
      intPos = txtsubtitle.SelStart
      txtsubtitle = oProd.SubTitle
      txtsubtitle.SelStart = intPos - 1
    End If
End Sub


Private Sub txtTitle_LostFocus()
    If flgLoading Then Exit Sub
    txttitle = oProd.Title
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtTitle_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetTitle(txttitle)
    If Err Then
      Beep
      intPos = txttitle.SelStart
      txttitle = oProd.Title
      txttitle.SelStart = intPos - 1
    End If
End Sub
Private Sub txtAuthor_LostFocus()
    If flgLoading Then Exit Sub
    txtauthor = oProd.Author
End Sub
Private Sub txtAuthor_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtAuthor_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetAuthor(txtauthor)
    If Err Then
      Beep
      intPos = txtauthor.SelStart
      txtauthor = oProd.Author
      txtauthor.SelStart = intPos - 1
    End If
End Sub
Private Sub txtPublisher_LostFocus()
    If flgLoading Then Exit Sub
    txtpublisher = oProd.Publisher
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtPublisher_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetPublisher(txtpublisher)
    If Err Then
      Beep
      intPos = txtpublisher.SelStart
      txtpublisher = oProd.Publisher
      txtpublisher.SelStart = intPos - 1
    End If
End Sub
Private Sub txtPubDate_LostFocus()
    If flgLoading Then Exit Sub
    txtPubDate = oProd.PublicationDate
End Sub
Private Sub txtPubDate_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtPubDate_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetPublicationDate(txtPubDate)
    If Err Then
      Beep
      intPos = txtPubDate.SelStart
      txtPubDate = oProd.PublicationDate
      txtPubDate.SelStart = intPos - 1
    End If
End Sub
Private Sub txtPubPlace_LostFocus()
    If flgLoading Then Exit Sub
    txtPubPlace = oProd.PublicationPlace
End Sub
Private Sub txtPubPlace_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtPubPlace_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetPublicationPlace(txtPubPlace)
    If Err Then
      Beep
      intPos = txtPubPlace.SelStart
      txtPubPlace = oProd.PublicationPlace
      txtPubPlace.SelStart = intPos - 1
    End If
End Sub
Private Sub txtEdition_LostFocus()
    If flgLoading Then Exit Sub
    txtEdition = oProd.Edition
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtEdition_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetEdition(txtEdition)
    If Err Then
      Beep
      intPos = txtEdition.SelStart
      txtEdition = oProd.Edition
      txtEdition.SelStart = intPos - 1
    End If
End Sub
Private Sub txtEAN_LostFocus()
    If flgLoading Then Exit Sub
    txtEAN = oProd.EAN
End Sub
Private Sub txtEAN_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtEAN_Change()
Dim intPos As Integer
    If flgLoading Then Exit Sub
    On Error Resume Next
    mCancel = Not oProd.SetEAN(txtEAN)
    If Err Then
      Beep
      intPos = txtEAN.SelStart
      txtEAN = oProd.EAN
      txtEAN.SelStart = intPos - 1
    End If
End Sub
Private Sub txtNote_LostFocus()
    If flgLoading Then Exit Sub
    txtNote = oProd.Note
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtNote_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetNote(txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oProd.Note
      txtNote.SelStart = intPos - 1
    End If
End Sub

Private Sub txtSummary_LostFocus()
    If flgLoading Then Exit Sub
    txtSummary = oProd.Summary
End Sub
Private Sub txtSummary_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtSummary_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetSummary(txtSummary)
    If Err Then
      Beep
      intPos = txtSummary.SelStart
      txtSummary = oProd.Summary
      txtSummary.SelStart = intPos - 1
    End If
End Sub

Private Sub txtDescription_LostFocus()
    If flgLoading Then Exit Sub
    txtDescription = oProd.Description
End Sub
Private Sub txtDescription_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtDescription_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetDescription(txtDescription)
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oProd.Description
      txtDescription.SelStart = intPos - 1
    End If
End Sub
Private Sub txtComment_LostFocus()
    If flgLoading Then Exit Sub
    txtComment = oProd.Comment
End Sub
Private Sub txtComment_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtComment_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetComment(txtComment)
    If Err Then
      Beep
      intPos = txtComment.SelStart
      txtComment = oProd.Comment
      txtComment.SelStart = intPos - 1
    End If
End Sub
Private Sub txtSP_LostFocus()
    If flgLoading Then Exit Sub
    txtSP = oProd.SPF
End Sub
Private Sub txtSP_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtSP_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetSP(txtSP)
    If Err Then
      Beep
      intPos = txtSP.SelStart
      txtSP = oProd.SP
      txtSP.SelStart = intPos - 1
    End If
End Sub
Private Sub txtRRP_LostFocus()
    If flgLoading Then Exit Sub
    txtRRP = oProd.RRPF
End Sub
Private Sub txtRRP_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtRRP_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetRRP(txtRRP)
    If Err Then
      Beep
      intPos = txtRRP.SelStart
      txtRRP = oProd.RRP
      txtRRP.SelStart = intPos - 1
    End If
End Sub
Private Sub txtCost_LostFocus()
    If flgLoading Then Exit Sub
    txtCost = oProd.CostF
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    Cancel = mCancel
End Sub
Private Sub txtCost_Change()
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetCost(txtCost)
    If Err Then
      Beep
      intPos = txtCost.SelStart
      txtCost = oProd.Cost
      txtCost.SelStart = intPos - 1
    End If
End Sub

