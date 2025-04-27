VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProductPreviewAQ 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Details"
   ClientHeight    =   6435
   ClientLeft      =   1200
   ClientTop       =   1230
   ClientWidth     =   9825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   9825
   Begin VB.TextBox txtEAN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   7635
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   300
      Width           =   1950
   End
   Begin VB.TextBox txtEdition 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   360
      Left            =   6945
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   990
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   8610
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5445
      Width           =   1110
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1605
      Left            =   0
      TabIndex        =   18
      Top             =   4470
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   2831
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   10485760
      TabCaption(0)   =   "&1. Stock"
      TabPicture(0)   =   "frmProductPreviewAQ.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label32"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label31"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label15"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtUSPrice"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUKPrice"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtSpecial"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCost"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSP"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtRRP"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPreOwned"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtReserved"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtOnHand"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "&2. Pre-owned copies"
      TabPicture(1)   =   "frmProductPreviewAQ.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&3. Details"
      TabPicture(2)   =   "frmProductPreviewAQ.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtVAT"
      Tab(2).Control(1)=   "cmdSetDefault"
      Tab(2).Control(2)=   "chkNonStock"
      Tab(2).Control(3)=   "chkObsolete"
      Tab(2).Control(4)=   "cboCatHead"
      Tab(2).Control(5)=   "cboCategory"
      Tab(2).Control(6)=   "txtBinding"
      Tab(2).Control(7)=   "txtBIC"
      Tab(2).Control(8)=   "txtSeriesTitle"
      Tab(2).Control(9)=   "Label10"
      Tab(2).Control(10)=   "Label17"
      Tab(2).Control(11)=   "Label20"
      Tab(2).Control(12)=   "Label24"
      Tab(2).Control(13)=   "Label25"
      Tab(2).Control(14)=   "Label26"
      Tab(2).ControlCount=   15
      TabCaption(3)   =   "&4. Notes"
      TabPicture(3)   =   "frmProductPreviewAQ.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label30"
      Tab(3).Control(1)=   "Label29"
      Tab(3).Control(2)=   "Label28"
      Tab(3).Control(3)=   "Label27"
      Tab(3).Control(4)=   "txtComment"
      Tab(3).Control(5)=   "txtSummary"
      Tab(3).Control(6)=   "txtDescription"
      Tab(3).Control(7)=   "txtNote"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "&5. Statistics"
      TabPicture(4)   =   "frmProductPreviewAQ.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame2"
      Tab(4).Control(1)=   "txtDateAdded"
      Tab(4).Control(2)=   "txtDateLastModified"
      Tab(4).Control(3)=   "txtLastDelivered"
      Tab(4).Control(4)=   "txtLastOrdered"
      Tab(4).Control(5)=   "txtLastDeliveredQty"
      Tab(4).Control(6)=   "txtLastOrderedQty"
      Tab(4).Control(7)=   "txtLastOrderedPrice"
      Tab(4).Control(8)=   "txtLastDeliveredPrice"
      Tab(4).Control(9)=   "Label3"
      Tab(4).Control(10)=   "Label12"
      Tab(4).Control(11)=   "Label22"
      Tab(4).Control(12)=   "Label23"
      Tab(4).ControlCount=   13
      TabCaption(5)   =   "&6. Wants"
      TabPicture(5)   =   "frmProductPreviewAQ.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "lvwWants"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame2 
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
         Left            =   -74640
         TabIndex        =   76
         Top             =   1920
         Width           =   5220
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   78
            Top             =   330
            Width           =   825
         End
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
            TabIndex        =   77
            Top             =   330
            Width           =   825
         End
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
         Left            =   -73170
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   645
         Width           =   1815
      End
      Begin VB.TextBox txtDateLastModified 
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
         Left            =   -73170
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1155
         Width           =   1815
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
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   675
         Width           =   1380
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
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1185
         Width           =   1380
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
         Left            =   -68280
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   675
         Width           =   555
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
         Left            =   -68280
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1185
         Width           =   555
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
         Left            =   -67665
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1185
         Width           =   960
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
         Left            =   -67665
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   675
         Width           =   960
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -73815
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   2310
         Width           =   7230
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -73815
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   1095
         Width           =   7230
      End
      Begin VB.TextBox txtSummary 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -73815
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   495
         Width           =   7230
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -73815
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   1710
         Width           =   7230
      End
      Begin VB.TextBox txtVAT 
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
         Left            =   -73620
         TabIndex        =   47
         Top             =   2490
         Width           =   1380
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
         Left            =   -72210
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2490
         Width           =   1755
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
         Left            =   -67935
         TabIndex        =   45
         Top             =   1995
         Width           =   1245
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
         Left            =   -67935
         TabIndex        =   44
         Top             =   2445
         Width           =   1245
      End
      Begin VB.ComboBox cboCatHead 
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
         Height          =   360
         Left            =   -72975
         TabIndex        =   43
         Top             =   1140
         Width           =   2520
      End
      Begin VB.ComboBox cboCategory 
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
         Height          =   360
         Left            =   -69090
         TabIndex        =   42
         Top             =   1140
         Width           =   2520
      End
      Begin VB.TextBox txtBinding 
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
         Left            =   -73620
         TabIndex        =   41
         Top             =   2070
         Width           =   1380
      End
      Begin VB.TextBox txtBIC 
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
         Left            =   -72960
         TabIndex        =   40
         Top             =   1590
         Width           =   1380
      End
      Begin VB.TextBox txtSeriesTitle 
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
         Left            =   -73620
         TabIndex        =   39
         Top             =   615
         Width           =   7020
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
         Left            =   -73755
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   615
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
         Left            =   -73755
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1125
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
         Left            =   -73755
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1635
         Width           =   1380
      End
      Begin VB.TextBox txtRRP 
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
         Left            =   -71445
         TabIndex        =   26
         Top             =   615
         Width           =   1380
      End
      Begin VB.TextBox txtSP 
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
         Left            =   -71445
         TabIndex        =   25
         Top             =   1125
         Width           =   1380
      End
      Begin VB.TextBox txtCost 
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
         Left            =   -71445
         TabIndex        =   24
         Top             =   1650
         Width           =   1380
      End
      Begin VB.TextBox txtSpecial 
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
         Left            =   -71445
         TabIndex        =   23
         Top             =   2160
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
         Left            =   -68610
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   630
         Width           =   1380
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
         Left            =   -68610
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1140
         Width           =   1380
      End
      Begin MSComctlLib.ListView lvwWants 
         Height          =   2460
         Left            =   90
         TabIndex        =   20
         Top             =   540
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   4339
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14155263
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2152
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   3951
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Note"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label3 
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
         Left            =   -74580
         TabIndex        =   75
         Top             =   735
         Width           =   1290
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
         Left            =   -74670
         TabIndex        =   74
         Top             =   1200
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
         Left            =   -71235
         TabIndex        =   73
         Top             =   720
         Width           =   1395
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
         Left            =   -71250
         TabIndex        =   72
         Top             =   1215
         Width           =   1395
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
         Left            =   -74760
         TabIndex        =   61
         Top             =   2325
         Width           =   825
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
         Left            =   -74970
         TabIndex        =   60
         Top             =   1110
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
         Left            =   -74970
         TabIndex        =   59
         Top             =   510
         Width           =   1035
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
         Left            =   -74970
         TabIndex        =   58
         Top             =   1725
         Width           =   1035
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
         Left            =   -74850
         TabIndex        =   53
         Top             =   2520
         Width           =   1080
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
         Left            =   -74865
         TabIndex        =   52
         Top             =   1185
         Width           =   1755
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
         Left            =   -74850
         TabIndex        =   51
         Top             =   2085
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
         Left            =   -74835
         TabIndex        =   50
         Top             =   675
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
         Left            =   -74205
         TabIndex        =   49
         Top             =   1620
         Width           =   1080
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
         Left            =   -70305
         TabIndex        =   48
         Top             =   1185
         Width           =   1080
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
         Left            =   -74880
         TabIndex        =   38
         Top             =   675
         Width           =   1065
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
         Left            =   -74880
         TabIndex        =   37
         Top             =   1185
         Width           =   1065
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
         Left            =   -74880
         TabIndex        =   36
         Top             =   1695
         Width           =   1065
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
         Left            =   -72285
         TabIndex        =   35
         Top             =   660
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
         Left            =   -72285
         TabIndex        =   34
         Top             =   1185
         Width           =   750
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
         Left            =   -72285
         TabIndex        =   33
         Top             =   1695
         Width           =   750
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
         Left            =   -72285
         TabIndex        =   32
         Top             =   2220
         Width           =   750
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
         Left            =   -70020
         TabIndex        =   31
         Top             =   690
         Width           =   1290
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
         Left            =   -70020
         TabIndex        =   30
         Top             =   1200
         Width           =   1290
      End
   End
   Begin VB.TextBox txtPubDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2325
      Width           =   2655
   End
   Begin VB.TextBox txtPubPlace 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1875
      Width           =   2655
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      Height          =   825
      Left            =   8610
      Picture         =   "frmProductPreviewAQ.frx":00A8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3150
      Width           =   1095
   End
   Begin VB.TextBox txtpublisher 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1425
      Width           =   2655
   End
   Begin VB.TextBox txtsubtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H00800000&
      Height          =   660
      Left            =   735
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1620
      Width           =   4215
   End
   Begin VB.TextBox txttitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   735
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   990
      Width           =   4215
   End
   Begin VB.TextBox txtauthor 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2355
      Width           =   3975
   End
   Begin VB.TextBox txtCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   255
      Width           =   1800
   End
   Begin VB.CommandButton cmdsearchisbn 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   225
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find By ISBN"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   735
      TabIndex        =   0
      Top             =   30
      Width           =   3255
      Begin VB.TextBox txtisbnsearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   135
         TabIndex        =   1
         Top             =   285
         Width           =   1995
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EAN"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7125
      TabIndex        =   83
      Top             =   345
      Width           =   585
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edition"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6240
      TabIndex        =   63
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pub. date"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5940
      TabIndex        =   17
      Top             =   2355
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pub. place"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5970
      TabIndex        =   15
      Top             =   1905
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Publisher"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6075
      TabIndex        =   11
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Subtitle"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   -45
      TabIndex        =   9
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Title"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   -45
      TabIndex        =   7
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Author"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   105
      TabIndex        =   5
      Top             =   2370
      Width           =   570
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Code"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   330
      Width           =   495
   End
End
Attribute VB_Name = "frmProductPreviewAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private roProd As a_Product
Private lngID As Long
Private lslist As ListItem

Private Sub cmdbrowse_Click()
    Unload Me
End Sub

Public Sub Component(objcomponent As a_Product)

    Set roProd = objcomponent

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
Dim frm As frmProductAQ
Dim oProd As a_Product
    Set frm = New frmProductAQ
 '   Set oProd = New a_Product
'    oProd.Load roProd.ID, 0
    frm.Component roProd
  '  roProd.BeginEdit
    frm.Show
    
    Unload Me
    Set oProd = Nothing
    Set frm = Nothing
    Exit Sub
ERRH:
    MsgBox Error

End Sub

Private Sub cmdsearchisbn_Click()
    lvwCopies.ListItems.Clear
    Me.lvwWants.ListItems.Clear
    Set roProd = Nothing
    Set roProd = New a_Product
    With roProd
    .Load "", 0, txtisbnsearch
       
    txtAuthor = .Author
    txtCode = .Code
    txtSubtitle = .SubTitle
    txtTitle = .Title
    txtPublisher = .Publisher
        
    End With
    
'    For Each oINVOICEDisplay In roProd.OutstandingINVOICE
'
'   Set lslist = _
'    lsinvdetails.ListItems.Add
'    With lslist
'        .Key = Format$(oINVOICEDisplay.TransactionID) & "k"
'        .Text = oINVOICEDisplay.TPName
'        .SubItems(1) = oINVOICEDisplay.InvoiceNumber
'   '     .SubItems(2) = oINVOICEDisplay.TDateFormatted
'
'    End With
'
'    Next
'
'   For Each oOSSODisplay In roProd.OutstandingSOrder
'
'   Set lslist = _
'        lssodetails.ListItems.Add
'        With lslist
'            .Key = Format$(oOSSODisplay.TID) & "k"
'            .Text = oOSSODisplay.Code & " " & oOSSODisplay.TDateFormatted & " " & oOSSODisplay.SOLFulfilled
'            .SubItems(1) = oOSSODisplay.QtyFirm & "/" & oOSSODisplay.QtySS
'            .SubItems(2) = oOSSODisplay.ReceivedSoFar
'
'        End With
'
'    Next
'
'    For Each oOSCODisplay In roProd.OutstandingCOrder
'
'    Set lslist = _
'        lscodetails.ListItems.Add
'        With lslist
'            .Key = Format$(oOSCODisplay.TID) & "k"
'            .Text = oOSCODisplay.Code & " " & oOSCODisplay.TDateFormatted & " " & oOSCODisplay.COLFulFilled
'            .SubItems(1) = oOSCODisplay.COLQty
'            .SubItems(2) = oOSCODisplay.COLCollected
'        End With
'    Next
'
'    For Each oAPPRODisplay In roProd.OutstandingAPPRO
'
'    Set lslist = _
'        lsapprodetails.ListItems.Add
'        With lslist
'            .Key = Format$(oAPPRODisplay.TID) & "k"
'            .Text = oAPPRODisplay.Code & " " & oAPPRODisplay.TDateFormatted
'            .SubItems(1) = oAPPRODisplay.APPQty
'            .SubItems(2) = oAPPRODisplay.APPReturned
'        End With
'    Next
'
End Sub

Private Sub Form_Load()
Dim oOSSODisplay As d_SOLine
Dim oOSCODisplay As d_OSCOrder
Dim oAPPRODisplay As d_OSAPPRO
Dim oINVOICEDisplay As d_Invoice
    Top = 10
    Left = 50
    Width = 10000
    Height = 7000
    txtAuthor = roProd.Author
    txtSubtitle = roProd.SubTitle
    txtSeriesTitle = roProd.SeriesTitle
    txtSP = roProd.SP
    txtCode = roProd.Code
    txtEAN = roProd.EAN
    txtTitle = roProd.Title
    txtSubtitle = roProd.SubTitle
    txtEdition = roProd.Edition
    txtPublisher = roProd.Publisher
    txtPubPlace = roProd.PublicationPlace
    txtPubDate = roProd.PublicationDate
    txtBinding = roProd.BindingCode
    txtComment = roProd.Comment
    txtDescription = roProd.Description
    txtNote = roProd.Note
    txtSummary = roProd.Summary
    txtLastDeliveredQty = roProd.QtyLastDeliveredF
    txtLastDeliveredPrice = roProd.PriceLastDeliveredF
    txtLastOrderedQty = roProd.QtyLastOrderedF
    txtLastOrderedPrice = roProd.PriceLastOrderedF
    Me.txtLastDelivered = roProd.DateLastDeliveredF
    Me.txtLastOrdered = roProd.DateLastOrderedF
    Me.txt12 = roProd.Stock12
    Me.txt18 = roProd.Stock18
    Me.txt18Plus = roProd.Stock18Plus
    Me.txt6 = roProd.Stock6
    Me.txtAgedDate = roProd.StockAgedDate
    txtRRP = roProd.RRPF
    Me.txtDateAdded = roProd.DaterecordAddedF
    Me.txtDateLastModified = roProd.DateLastModifiedF
    Me.txtVAT = roProd.VATRateF
    Me.txtBIC = roProd.BIC
    Me.chkNonStock = IIf(roProd.NonStock, 1, 0)
    Me.chkObsolete = IIf(roProd.obsolete, 1, 0)
    
    LoadCopies
    LoadWants
    Me.SSTab1.Tab = 0
'   For Each oINVOICEDisplay In roProd.OutstandingINVOICE
'
'   Set lslist = _
'    lsinvdetails.ListItems.Add
'    With lslist
'        .Key = Format$(oINVOICEDisplay.TransactionID) & "k"
'        .Text = oINVOICEDisplay.TPName
'        .SubItems(1) = oINVOICEDisplay.InvoiceNumber
'        .SubItems(2) = oINVOICEDisplay.TDateFormatted
'
'    End With
'
'    Next
'
'   For Each oOSSODisplay In roProd.OutstandingSOrder
'
'   Set lslist = _
'        lssodetails.ListItems.Add
'        With lslist
'            .Key = Format$(oOSSODisplay.TID) & "k"
'            .Text = oOSSODisplay.Code & " " & oOSSODisplay.TDateFormatted & " " & oOSSODisplay.SOLFulfilled
'            .SubItems(1) = oOSSODisplay.QtyFirm & "/" & oOSSODisplay.QtySS
'            .SubItems(2) = oOSSODisplay.ReceivedSoFar
'
'        End With
'
'    Next
    
    
End Sub
Private Sub LoadCopies()
Dim objItm As ListItem
Dim i, j As Integer
Dim tmp As String
Dim strCatalogues As String


    lvwCopies.ListItems.Clear
    For i = 1 To roProd.Copies.Count
        strCatalogues = ""
        For j = 1 To roProd.Copies(i).CatalogueEntries.Count
            strCatalogues = strCatalogues & roProd.Copies(i).CatalogueEntries(j).Serial
            If j < roProd.Copies(i).CatalogueEntries.Count Then strCatalogues = strCatalogues & ", "
        Next j
        Set objItm = Me.lvwCopies.ListItems.Add
        With objItm
            .Key = roProd.Copies(i).Key
            .Text = roProd.Copies(i).Serial
            .SubItems(1) = roProd.Copies(i).PriceF
            .SubItems(2) = roProd.Copies(i).PurchaseDateF
            .SubItems(3) = roProd.Copies(i).Comment
            .SubItems(4) = roProd.Copies(i).SoldDateF
            .SubItems(5) = strCatalogues
         '   .SubItems(3) = IIf(oCust.Addresses(i).ID = oCust.DefaultAddress.ID, "Default", "")
        End With
    Next i

End Sub
Private Sub LoadWants()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String
Dim strCatalogues As String


    lvwWants.ListItems.Clear
    For i = 1 To roProd.Wants.Count
        Set objItm = Me.lvwWants.ListItems.Add
        With objItm
            .Key = roProd.Wants(i).ID & "k"
            .Text = roProd.Wants(i).ReqDateF
            .SubItems(1) = roProd.Wants(i).CustomerName
            .SubItems(2) = roProd.Wants(i).Note
        End With
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set roProd = Nothing

End Sub

'Private Sub lsapprodetails_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub
'
'Private Sub lscodetails_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub
'
'Private Sub lsinvdetails_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub
'
'Private Sub lssodetails_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub

'Private Sub lssodetails_DblClick()
'
'    If lssodetails.ListItems.Count = 0 Then Exit Sub
'    Dim lngprod As Long
'    Set oOSSOrder = New d_ROOSSOrder
'    lngprod = oOSSOrder.Load(Val(lssodetails.SelectedItem.Key))
'    'frmSODetails.Component oOSSOrder
'    'frmSODetails.Show vbModal
'    Set oOSSOrder = Nothing
'
'End Sub
'
'Private Sub lsinvdetails_DblClick()
'
'Dim frmInvoice As frmInvoicePreview
'Dim strprod As String
'    If lsinvdetails.ListItems.Count = 0 Then Exit Sub
'    Set oInvoice = New a_Invoice_P
'    Set frmInvoice = New frmInvoicePreview
'    strprod = oInvoice.Fetch(Val(lsinvdetails.SelectedItem.Key))
'    frmInvoice.component Val(lsinvdetails.SelectedItem.Key)
'    frmInvoice.Show
'    Set oInvoice = Nothing
'    Set frmInvoice = Nothing
'
'End Sub
'
'Private Sub lscodetails_DblClick()
'
'Dim frmCOrder As frmCOrderPreview
'Dim strCO As String
'    If lscodetails.ListItems.Count = 0 Then Exit Sub
'    Set oCOrder = New a_COrder
'    Set frmCOrder = New frmCOrderPreview
'    oCOrder.Load Val(lscodetails.SelectedItem.Key)
'    frmCOrder.component Val(lscodetails.SelectedItem.Key)
'    frmCOrder.Show
'    Set oCOrder = Nothing
'    Set frmCOrder = Nothing
'
'End Sub
'
Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvwCopies_DblClick()
Dim frm As frmCopyPreview
Dim oCopy As a_Copy
    Set oCopy = roProd.Copies(lvwCopies.SelectedItem.Key)
    Set frm = New frmCopyPreview
    frm.Component oCopy
    frm.Show
    
End Sub

Private Sub lvwWants_BeforeLabelEdit(Cancel As Integer)

End Sub
