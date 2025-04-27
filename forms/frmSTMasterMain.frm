VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmMain 
   Caption         =   "Import and manage stocktakes"
   ClientHeight    =   6390
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   11880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.txt"
      DialogTitle     =   "Locate scanner files"
      MaxFileSize     =   30000
   End
   Begin VB.Frame fr7 
      Height          =   285
      Left            =   15
      TabIndex        =   19
      Top             =   5655
      Width           =   4260
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   150
         Left            =   60
         TabIndex        =   20
         Top             =   105
         Visible         =   0   'False
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   265
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5520
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   9737
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      Tab             =   4
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   12632256
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Step 1: New or append"
      TabPicture(0)   =   "frmSTMasterMain.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Step 2: Import"
      TabPicture(1)   =   "frmSTMasterMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(1)=   "G1"
      Tab(1).Control(2)=   "Label30"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Step 3: Review"
      TabPicture(2)   =   "frmSTMasterMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "lvwReview"
      Tab(2).Control(2)=   "cboFilename"
      Tab(2).Control(3)=   "cmdGo"
      Tab(2).Control(4)=   "cmdDeleteSelectedRow"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Step 4: Review (2)"
      TabPicture(3)   =   "frmSTMasterMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "txtNegAdj"
      Tab(3).Control(2)=   "cmdClearNegQtys"
      Tab(3).Control(3)=   "Label31"
      Tab(3).Control(4)=   "Label7"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Step 5: Correct and build"
      TabPicture(4)   =   "frmSTMasterMain.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "lbl1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label9"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtNote"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdBuild(1)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtDateTime"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Frame5"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Step&6 Finalize"
      TabPicture(5)   =   "frmSTMasterMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdFinalize"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "cmdProvisional"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Step &7 Print Summary"
      TabPicture(6)   =   "frmSTMasterMain.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label11"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label12"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Label13"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Label14"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Label15"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Label32"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "Frame4"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "txtTotalProducts"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "txtTotalItems"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "txtValueOfStockRetail"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "txtValueOfStockCost"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "txtAvgDisc"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "cmdPrint"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).ControlCount=   13
      Begin VB.Frame Frame5 
         Caption         =   "Validate"
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
         Height          =   1830
         Left            =   270
         TabIndex        =   78
         Top             =   1485
         Width           =   6375
         Begin VB.ComboBox cboCheck 
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
            ForeColor       =   &H8000000D&
            Height          =   360
            ItemData        =   "frmSTMasterMain.frx":00C4
            Left            =   915
            List            =   "frmSTMasterMain.frx":00D4
            TabIndex        =   81
            Top             =   480
            Width           =   4290
         End
         Begin VB.TextBox txtCheck 
            Alignment       =   2  'Center
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
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   5325
            TabIndex        =   80
            Top             =   450
            Width           =   855
         End
         Begin VB.CommandButton cmdReport 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Report"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3585
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   1155
            Width           =   2610
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Type of check"
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
            Height          =   225
            Left            =   945
            TabIndex        =   83
            Top             =   195
            Width           =   1365
         End
         Begin VB.Label Label21 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Note: Counts only apply where an adjustment has been made to the existing quantity."
            ForeColor       =   &H8000000D&
            Height          =   795
            Left            =   75
            TabIndex        =   82
            Top             =   945
            Width           =   2865
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Correction"
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
         Height          =   2565
         Left            =   -71595
         TabIndex        =   61
         Top             =   2235
         Width           =   6420
         Begin VB.CommandButton cmdFetch 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Fetch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3585
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   255
            Width           =   1035
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   1980
            Width           =   1035
         End
         Begin VB.TextBox txtDiff 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00BCF9FC&
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
            Height          =   375
            Left            =   5715
            TabIndex        =   68
            Top             =   975
            Width           =   555
         End
         Begin VB.TextBox txtCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00BCF9FC&
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
            Height          =   375
            Left            =   5130
            TabIndex        =   67
            Top             =   975
            Width           =   555
         End
         Begin VB.TextBox txtCode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00BCF9FC&
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
            Height          =   375
            Left            =   135
            TabIndex        =   66
            Top             =   975
            Width           =   1395
         End
         Begin VB.TextBox txtCorrection 
            Alignment       =   2  'Center
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
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   2685
            TabIndex        =   65
            Top             =   2010
            Width           =   855
         End
         Begin VB.TextBox txtBal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00BCF9FC&
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
            Height          =   375
            Left            =   4545
            TabIndex        =   64
            Top             =   975
            Width           =   555
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00BCF9FC&
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
            Height          =   375
            Left            =   1590
            TabIndex        =   63
            Top             =   975
            Width           =   2925
         End
         Begin VB.TextBox txtID 
            Alignment       =   2  'Center
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
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   2175
            TabIndex        =   62
            Top             =   285
            Width           =   1275
         End
         Begin VB.Label Label29 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Diff"
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
            Height          =   225
            Left            =   5700
            TabIndex        =   77
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label28 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Count"
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
            Height          =   225
            Left            =   5115
            TabIndex        =   76
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label27 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   1575
            TabIndex        =   75
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label26 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
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
            Height          =   225
            Left            =   195
            TabIndex        =   74
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label25 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Corrected count"
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
            Height          =   225
            Left            =   1170
            TabIndex        =   73
            Top             =   2055
            Width           =   1410
         End
         Begin VB.Label Label24 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Calc."
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
            Height          =   225
            Left            =   4530
            TabIndex        =   72
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "ID of row to change"
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
            Height          =   225
            Left            =   240
            TabIndex        =   71
            Top             =   315
            Width           =   1740
         End
      End
      Begin VB.CommandButton cmdProvisional 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Provisional adjustments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -74325
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   2250
         Width           =   2835
      End
      Begin VB.TextBox txtNegAdj 
         Alignment       =   2  'Center
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -64830
         TabIndex        =   58
         Top             =   1155
         Width           =   1710
      End
      Begin VB.CommandButton cmdClearNegQtys 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Clear negative O.H.quantities"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -73965
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1110
         Width           =   4065
      End
      Begin VB.TextBox txtDateTime 
         Alignment       =   2  'Center
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
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   8925
         TabIndex        =   53
         Top             =   3225
         Width           =   2835
      End
      Begin VB.CommandButton cmdBuild 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Build"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Index           =   1
         Left            =   8910
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3885
         Width           =   2850
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   720
         Left            =   8925
         TabIndex        =   51
         Top             =   1515
         Width           =   2805
      End
      Begin VB.CommandButton cmdFinalize 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Finalize"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   -65265
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2250
         Width           =   2610
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Print"
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
         Left            =   -71295
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3870
         Width           =   1665
      End
      Begin VB.TextBox txtAvgDisc 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   390
         Left            =   -71265
         TabIndex        =   43
         Top             =   3150
         Width           =   1665
      End
      Begin VB.TextBox txtValueOfStockCost 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   390
         Left            =   -71265
         TabIndex        =   42
         Top             =   2610
         Width           =   1665
      End
      Begin VB.TextBox txtValueOfStockRetail 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   390
         Left            =   -71265
         TabIndex        =   41
         Top             =   2070
         Width           =   1665
      End
      Begin VB.TextBox txtTotalItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   390
         Left            =   -71265
         TabIndex        =   40
         Top             =   1530
         Width           =   1665
      End
      Begin VB.TextBox txtTotalProducts 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   390
         Left            =   -71265
         TabIndex        =   39
         Top             =   990
         Width           =   1665
      End
      Begin VB.Frame Frame4 
         Caption         =   "Review other stocktakes"
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
         Height          =   4530
         Left            =   -68280
         TabIndex        =   26
         Top             =   570
         Width           =   4860
         Begin VB.CommandButton cmdPrintO 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Print"
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
            Left            =   2715
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   3885
            Width           =   1665
         End
         Begin VB.TextBox txtTotalProductsO 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   390
            Left            =   2700
            TabIndex        =   32
            Top             =   1230
            Width           =   1665
         End
         Begin VB.TextBox txtTotalItemsO 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   390
            Left            =   2700
            TabIndex        =   31
            Top             =   1770
            Width           =   1665
         End
         Begin VB.TextBox txtValueOfStockRetailO 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   390
            Left            =   2700
            TabIndex        =   30
            Top             =   2310
            Width           =   1665
         End
         Begin VB.TextBox txtValueOfStockCostO 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   390
            Left            =   2700
            TabIndex        =   29
            Top             =   2850
            Width           =   1665
         End
         Begin VB.TextBox txtAvgDiscO 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   390
            Left            =   2700
            TabIndex        =   28
            Top             =   3390
            Width           =   1665
         End
         Begin VB.ComboBox cboOtherstocktakes 
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
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   825
            TabIndex        =   27
            Top             =   480
            Width           =   3510
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total products"
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
            Left            =   120
            TabIndex        =   38
            Top             =   1290
            Width           =   2490
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Total items"
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
            TabIndex        =   37
            Top             =   1830
            Width           =   2490
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Value of stock (retail)"
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
            TabIndex        =   36
            Top             =   2370
            Width           =   2490
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Value of stock (cost)"
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
            TabIndex        =   35
            Top             =   2910
            Width           =   2490
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Average discount"
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
            TabIndex        =   34
            Top             =   3450
            Width           =   2490
         End
      End
      Begin VB.CommandButton cmdDeleteSelectedRow 
         BackColor       =   &H00D8D9C4&
         Caption         =   "&Delete selected row"
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
         Left            =   -66075
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4725
         Width           =   1965
      End
      Begin VB.CommandButton cmdGo 
         BackColor       =   &H00D8D9C4&
         Caption         =   "&Go"
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
         Left            =   -68085
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   810
         Width           =   840
      End
      Begin VB.ComboBox cboFilename 
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
         ForeColor       =   &H8000000D&
         Height          =   360
         ItemData        =   "frmSTMasterMain.frx":0128
         Left            =   -73695
         List            =   "frmSTMasterMain.frx":012A
         TabIndex        =   11
         Top             =   810
         Width           =   5310
      End
      Begin MSComctlLib.ListView lvwReview 
         Height          =   3300
         Left            =   -74850
         TabIndex        =   10
         Top             =   1395
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   5821
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Title"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Delivered price"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   " Import scanned files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1770
         Index           =   1
         Left            =   -74790
         TabIndex        =   7
         Top             =   540
         Width           =   7365
         Begin VB.CommandButton cmdImportSimple 
            BackColor       =   &H00C4BCA4&
            Caption         =   "Import"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   870
            Index           =   0
            Left            =   4755
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   555
            Width           =   2220
         End
         Begin VB.Label lblFiles 
            BackStyle       =   0  'Transparent
            Height          =   330
            Left            =   3510
            TabIndex        =   22
            Top             =   240
            Width           =   3675
         End
         Begin VB.Label Label2 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Import the text files containing the count data. You may select multiple files. This step may be repeated."
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
            Height          =   1005
            Left            =   195
            TabIndex        =   9
            Top             =   450
            Width           =   2955
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Choose existing or new stocktake"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4005
         Index           =   0
         Left            =   -74655
         TabIndex        =   2
         Top             =   915
         Width           =   10530
         Begin VB.ComboBox cboOperator 
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
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   345
            TabIndex        =   85
            Top             =   2385
            Width           =   2850
         End
         Begin VB.CommandButton cmdDeleteSA 
            BackColor       =   &H00D8D9C4&
            Caption         =   "&Delete selected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   7380
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2295
            Width           =   2085
         End
         Begin VB.ListBox lstStocktakes 
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            Height          =   810
            Left            =   6900
            TabIndex        =   14
            Top             =   1335
            Width           =   2925
         End
         Begin VB.ComboBox cboSAs 
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
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   345
            TabIndex        =   4
            Text            =   "cboSAs"
            Top             =   3345
            Width           =   3555
         End
         Begin VB.CommandButton cmdNewSA 
            BackColor       =   &H00D8D9C4&
            Caption         =   "Ne&w stock adjustment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   6975
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   3255
            Width           =   2955
         End
         Begin VB.Frame Frame3 
            Caption         =   "Unissued stocktakes"
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
            Height          =   2730
            Left            =   6720
            TabIndex        =   16
            Top             =   300
            Width           =   3285
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "You may delete any of these before starting a new one."
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
               Height          =   555
               Left            =   180
               TabIndex        =   17
               Top             =   435
               Visible         =   0   'False
               Width           =   2910
            End
         End
         Begin VB.Label lblCheckList 
            BackColor       =   &H80000000&
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
            ForeColor       =   &H8000000D&
            Height          =   1215
            Left            =   390
            TabIndex        =   6
            Top             =   510
            Width           =   5805
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Operator"
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
            Height          =   225
            Left            =   345
            TabIndex        =   86
            Top             =   2130
            Width           =   885
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Unissued stocktakes"
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
            Left            =   345
            TabIndex        =   5
            Top             =   2985
            Visible         =   0   'False
            Width           =   3525
         End
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         Height          =   2640
         Left            =   -74775
         OleObjectBlob   =   "frmSTMasterMain.frx":012C
         TabIndex        =   24
         Top             =   2700
         Width           =   10650
      End
      Begin VB.Label Label32 
         Caption         =   "For the Cost of inventory report in respect of the adjustments and the final list of adjustments, use the Reports application"
         ForeColor       =   &H8000000D&
         Height          =   570
         Left            =   -74100
         TabIndex        =   84
         Top             =   4785
         Width           =   4740
      End
      Begin VB.Label Label31 
         Caption         =   "Using this this date for the adjustment transaction           ( it must be before the cut-off date of the stock-take)"
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
         Height          =   570
         Left            =   -69840
         TabIndex        =   59
         Top             =   1080
         Width           =   4830
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Stocktake date and time"
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
         Height          =   225
         Left            =   8925
         TabIndex        =   55
         Top             =   2970
         Width           =   2310
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
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
         Height          =   225
         Left            =   8925
         TabIndex        =   54
         Top             =   1290
         Width           =   885
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Average discount"
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
         Left            =   -73830
         TabIndex        =   49
         Top             =   3210
         Width           =   2490
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Value of stock (cost)"
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
         Left            =   -73830
         TabIndex        =   48
         Top             =   2670
         Width           =   2490
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Value of stock (retail)"
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
         Left            =   -73830
         TabIndex        =   47
         Top             =   2130
         Width           =   2490
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Total items"
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
         Left            =   -73830
         TabIndex        =   46
         Top             =   1590
         Width           =   2490
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Total products"
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
         Left            =   -73830
         TabIndex        =   45
         Top             =   1050
         Width           =   2490
      End
      Begin VB.Label Label30 
         Caption         =   "Invalid scanned codes"
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
         Height          =   300
         Left            =   -74715
         TabIndex        =   25
         Top             =   2430
         Width           =   2910
      End
      Begin VB.Label lbl1 
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
         Height          =   3165
         Left            =   255
         TabIndex        =   21
         Top             =   1155
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000000&
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
         ForeColor       =   &H8000000D&
         Height          =   1275
         Left            =   -66210
         TabIndex        =   18
         Top             =   4170
         Width           =   4020
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "File name"
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
         Height          =   225
         Left            =   -74805
         TabIndex        =   12
         Top             =   870
         Width           =   885
      End
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   6000
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   10583
            MinWidth        =   10584
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   9945
      TabIndex        =   56
      Top             =   5670
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTemplate As String
Dim strFilename As String
Dim flgShowWORD As Boolean
Dim iCurrentPage As Integer
Dim tlSA As z_TextList
Dim tlSAO As z_TextList
Dim lngSAID As Long
Dim WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1
Dim tlFilename As z_TextList
Dim tlStaff As z_TextList
Dim dteDateTime As Date
Dim oSAO As a_Stktke
Dim bLoading As Boolean
Dim iStageNumber As Integer
Dim bViewOnly As Boolean
Dim strTitle As String
Dim strSQL As String
Private Sub cboOtherstocktakes_Click()
    On Error GoTo errHandler
    Set oSAO = New a_Stktke
    oSAO.Load tlSAO.Key(Me.cboOtherstocktakes)
    Me.txtTotalItemsO = oSAO.TotalItems
    Me.txtTotalProductsO = oSAO.TotalProducts
    Me.txtValueOfStockCostO = oSAO.ValueOfStockCostF
    Me.txtValueOfStockRetailO = oSAO.ValueOfStockRetailF
    Me.txtAvgDiscO = oSAO.AvgDiscountF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cboOtherstocktakes_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cboSAs_Click()
    On Error GoTo errHandler
    If bLoading Then Exit Sub
    
    If tlSA.Key(cboSAs.Text) > 0 Then
        lngSAID = tlSA.Key(cboSAs.Text)
    End If

    LoadSA lngSAID
    oSA.BeginEdit
    
    Me.txtDateTime = oSA.CutoffDate
    LoadListOfImports
    Me.Frame5.Enabled = True
    Me.Frame6.Enabled = True
    txtNegAdj = Format(DateAdd("h", -1, Now), "dd/mm/yyyy HH:NN")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cboSAs_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadSA(pSAID As Long)
    Set oSA = Nothing
    Set oSA = New a_Stktke
    oSA.Load pSAID
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim strSQL As String
    Set tlSA = New z_TextList
    tlSA.Load ltStockTake, "IN PROCESS"
    
    If tlSA.Count > 0 Then
'        LoadSA MostRecentIssuedSAID
'        bViewOnly = True
'        SSTab1.Tab = 3
'        lblCode.Caption = oSA.Code
'    Else
        LoadSA tlSA.Key(tlSA.ItemByOrdinalIndex(1))
        oSA.BeginEdit
        Frame5.Enabled = True
        Frame6.Enabled = True
        txtNegAdj = Format(DateAdd("h", -1, Now), "dd/mm/yyyy HH:NN")
        LoadListOfImports
    End If
    
    LoadCombo cboSAs, tlSA
    LoadListbox Me.lstStocktakes, tlSA
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadControls"
End Sub

Private Function MostRecentIssuedSAID() As Long
Dim lngSAID As Long
Dim rs As ADODB.Recordset

    lngSAID = 0
    Set rs = oPC.CO.Execute("SELECT MAX(STKTKE_ID) FROM tSTKTKE JOIN tTR ON STKTKE_ID = TR_ID WHERE TR_STATUS in (3,4)")
    If rs.State <> 0 Then
        If Not rs.EOF Then
            lngSAID = CLng(rs.Fields(0))
        End If
    End If
    MostRecentIssuedSAID = lngSAID
End Function
Private Sub cmdFT_Click()
    On Error GoTo errHandler
Dim fs As New Scripting.FileSystemObject
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim dteStart, dteEnd As Date
    
    Screen.MousePointer = vbHourglass
    dteStart = Now()
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdBuild_Click(Index As Integer)
    On Error GoTo errHandler
    If MsgBox("Ensure there are no users connected to the database before continuing. Press Cancel to skip the build for now.", vbInformation + vbOKCancel, "Warning") = vbCancel Then Exit Sub
    oSA.SalesPersonID = tlStaff.Key(cboOperator)
    Me.SB1.Panels(2).Text = "Building stock-take . . ."
    Screen.MousePointer = vbHourglass
    
    oSA.CreateStockAdjustment txtNote, dteDateTime
    
    Me.Frame5.Enabled = True
    Me.Frame6.Enabled = True
    Me.SB1.Visible = False
    Me.fr7.Visible = False
    Me.SSTab1.Tab = 3
    Screen.MousePointer = vbDefault
    MsgBox "Build complete", vbOKOnly, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdBuild_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClearNegQtys_Click()
Dim dteAdj As Date
    Screen.MousePointer = vbHourglass
    dteAdj = CDate(txtNegAdj)
    oSA.ClearNegativeQtys dteAdj
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDeleteSelectedRow_Click()
    oPC.CO.Execute "DELETE FROM STOCKTAKE_WORK1 WHERE ID = " & val(lvwReview.SelectedItem.Key)
    cmdGo_Click
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim strSQL
Dim rs As ADODB.Recordset

    strSQL = "SELECT tSTKTKEL.*,P_Code,P_Title,P_QtyONHand FROM tSTKTKEL INNER JOIN tProduct on tSTKTKEL.STKTKEL_P_ID = tproduct.P_ID WHERE STKTKEL_ID = " & CLng(txtID)
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.CO, adOpenKeyset
    If rs.EOF Then
        rs.Close
        txtCode = ""
        txtTitle = ""
        txtBal = ""
        txtCount = ""
        txtDiff = ""
        Me.cmdSave.Enabled = False
        Exit Sub
    End If
        txtCode = FNS(rs.Fields("P_Code"))
        txtTitle = FNS(rs.Fields("P_Title"))
        txtBal = FNN(rs.Fields("P_QtyOnHand"))
        txtCount = FNN(rs.Fields("STKTKEL_Qty"))
        txtDiff = FNN(rs.Fields("STKTKEL_Difference"))
    rs.Close
    Set rs = Nothing
    Me.cmdSave.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdFinalize_Click()
    On Error GoTo errHandler
    If MsgBox("Confirm you wish to finalize the stock-take. Changes will not be possible after this.", vbInformation + vbYesNo, "Confirm") = vbYes Then
        Me.SB1.Panels(2).Text = "Finalizing . . ."
        Screen.MousePointer = vbHourglass
        oSA.Finalize
        
        Me.txtTotalItems = oSA.TotalItems
        Me.txtTotalProducts = oSA.TotalProducts
        Me.txtValueOfStockCost = oSA.ValueOfStockCostF
        Me.txtValueOfStockRetail = oSA.ValueOfStockRetailF
        Me.txtAvgDisc = oSA.AvgDiscountF
        Me.SB1.Panels(1).Text = oSA.Code
        Me.SB1.Panels(2).Text = "Done"
        Screen.MousePointer = vbDefault
        MsgBox "Stock take is finalized. You can print the results.", vbInformation, "Status"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFinalize_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim itmList As ListItem
    strSQL = "SELECT * from vScanSTImportFile WHERE FN='" & cboFilename & "'"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.CO
    lvwReview.ListItems.Clear
    Do While Not rs.EOF
        Set itmList = lvwReview.ListItems.Add
        itmList.Key = FNS(rs.Fields("ID")) & "k"
        itmList.Text = FNS(rs.Fields("CODE"))
        itmList.SubItems(1) = FNS(rs.Fields("P_Title"))
        itmList.SubItems(2) = Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        itmList.SubItems(3) = Format(FNN(rs.Fields("P_LastPriceDelivered")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdImportSimple_Click(Index As Integer)
    On Error GoTo errHandler

Dim iresult As Integer
Dim fs As New Scripting.FileSystemObject
Dim lngBadRecords As Long
Dim dteEffectiveDate As Date
Dim lngLastSAID As Long
Dim rs As ADODB.Recordset
Dim lngFilecount As Long
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\STOCKTKE") Then
        fs.CreateFolder oPC.SharedFolderRoot & "\STOCKTKE"
    End If
    CD1.InitDir = oPC.SharedFolderRoot & "\STOCKTKE"
    CD1.FLAGS = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer
    CD1.CancelError = True
    CD1.Filter = "Text Files (*.txt)|*.txt"
    CD1.ShowOpen
    If CD1.FileName = "" Then
        MsgBox "You must specify an existing file name!", vbInformation, "Invalid filename"
    Else
        strFilename = CD1.FileName
        SB1.Panels(2) = "Importing from: " & strFilename
    End If

    Me.PB1.Visible = True
    Screen.MousePointer = vbHourglass
    Me.Refresh
    
    oSA.ImportSimple strFilename, lngBadRecords, lngFilecount 'creates a stocktake and imports the data into STOCKTAKE_WORKC consolidated by filename
    oSA.PrepareMissingData
    
    Me.SB1.Panels(2).Text = ""
    lngLastSAID = oSA.TransactionID
    LoadListOfImports
    Me.lblFiles.Caption = lngFilecount & " files imported"
    Me.PB1.Visible = False
    Screen.MousePointer = vbDefault
    MsgBox "Import complete: " & CStr(lngFilecount) & " files imported", vbOKOnly, "Status"
    Me.SSTab1.Tab = 2
    Me.Refresh
EXIT_Handler:
    Me.PB1.Visible = False
    Exit Sub
errHandler:
    ErrPreserve
    If Err = 32755 Then
        GoTo EXIT_Handler
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdImportSimple_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadListOfImports()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
    cboFilename.Clear
    Set rs = oSA.Filenames
    If rs.State > 0 Then   'the recordset is open i.e. rows have been found
        Do While Not rs.EOF And Not rs.BOF
            cboFilename.AddItem rs.Fields(1)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadListOfImports"
End Sub

Private Sub cmdNewSA_Click()
    On Error GoTo errHandler
    If MsgBox("It is strongly recommended that you backup the database immediately before you start this procedure." & vbCrLf & "Select CANCEL if you wish to stop now.", vbOKCancel + vbExclamation, "Important") = vbCancel Then
        Exit Sub
    End If
    Set oSA = New a_Stktke
    oSA.BeginEdit
    oSA.SetStatus stInProcess
    oSA.Zeroising = False
    Me.lvwReview.ListItems.Clear
    SB1.Panels(1).Text = "Code:" & oSA.Code & ", Date:" & oSA.CaptureDate & ", Status:" & oSA.Status
    Me.SSTab1.Tab = 1
    oSA.PrepareTempFiles
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdNewSA_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim ar As New arSummary
    arSummary.txtTitle2 = "Stock take at " & oPC.Configuration.DefaultStore.Description
    arSummary.txtTitle = "Summary of stocktake " & oSA.Code & " dated " & Format(oSA.CutoffDate, "General Date")
    arSummary.txtAvgDiscount = oSA.AvgDiscountF
    arSummary.txtCostValue = oSA.ValueOfStockCostF
    arSummary.txtQtyItem = oSA.TotalItems
    arSummary.txtQTYProduct = oSA.TotalProducts
    arSummary.txtRetailValue = oSA.ValueOfStockRetailF
    arSummary.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintAdjustmentsFinal_Click()
    PrintAdjustMentReportFinal
End Sub

Private Sub cmdPrintO_Click()
    On Error GoTo errHandler
Dim ar As New arSummary
    arSummary.txtTitle = "Summary of stocktake " & Trim(oSAO.Code) & " dated " & oSAO.CaptureDate
    arSummary.txtAvgDiscount = oSAO.AvgDiscountF
    arSummary.txtCostValue = oSAO.ValueOfStockCostF
    arSummary.txtQtyItem = oSAO.TotalItems
    arSummary.txtQTYProduct = oSAO.TotalProducts
    arSummary.txtRetailValue = oSAO.ValueOfStockRetailF
    arSummary.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdPrintO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdProvisional_Click()

    PrintAdjustMentReport
End Sub

Private Sub cmdReport_Click()
    On Error GoTo errHandler
Dim ar As arValidation
Dim ar3 As arMissing_1
Dim rs As ADODB.Recordset
Dim strTitle As String
Dim tmpDouble As Double
Dim tmpNumber As Long
Dim tmpCurrency As Currency

  '  Set rs = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Me.SB1.Panels(2).Text = "Generating report . . ."
    Select Case Me.cboCheck
    Case "Qty on hand greater than"
        If Not ConvertToLng(txtCheck, tmpNumber) Then
            MsgBox "Invalid value in criterion box"
            Exit Sub
        End If
    
        strSQL = "SELECT tProduct.* FROM tProduct JOIN STOCKTAKE_WORKC ON P_ID = PID Where P_QtyOnHand > " & tmpNumber & " ORDER BY P_TITLE"
        strTitle = "Qty on hand greater than " & txtCheck
        PrintValidation_C
    Case "Qty negative"
        Set ar = New arValidation
        ar.Printer.Orientation = ddOPortrait
        strTitle = "Qty negative "
        strSQL = "SELECT tProduct.* FROM tProduct JOIN STOCKTAKE_WORKC ON P_ID = PID  Where P_QTYONHAND < 0   ORDER BY P_TITLE"
        Set rs = New ADODB.Recordset
        rs.Open strSQL, oPC.CO, adOpenKeyset
        ar.Component rs, strTitle
        ar.Caption = strTitle
        ar.Show
        Set rs = Nothing
        Set ar = Nothing
    Case "Adjustment greater than (+ve or -ve)"
        PrintAdjustMentReport
        
    Case "Count greater than"
        Set ar = New arValidation
        ar.Printer.Orientation = ddOPortrait
        If Not ConvertToLng(txtCheck, tmpNumber) Then
            MsgBox "Invalid value in criterion box"
            Exit Sub
        End If
        strTitle = "Count greater than " & txtCheck
        strSQL = "SELECT CNT,tProduct.* FROM tProduct  JOIN STOCKTAKE_WORKC ON P_ID = PID Where CNT > " & tmpNumber & " ORDER BY P_TITLE"
        Set rs = New ADODB.Recordset
        rs.Open strSQL, oPC.CO, adOpenKeyset
        ar.Caption = strTitle
        ar.Component rs, strTitle
        ar.Show
        Set rs = Nothing
        Set ar = Nothing
    Case "Price greater than"
        If Not ConvertToCurr(txtCheck, tmpCurrency) Then
            MsgBox "Invalid value in criterion box"
            Exit Sub
        Else
  '          Me.txtCheck = Format(tmpCurrency, "Currency")
        End If
        strTitle = "Price greater than " & txtCheck
        strSQL = "SELECT tProduct.* FROM tProduct   JOIN STOCKTAKE_WORKC ON P_ID = PID Where P_SP > " & CLng(tmpCurrency) * oPC.Configuration.DefaultCurrency.Divisor & " ORDER BY P_TITLE"
        PrintValidation_C
    Case "Discount greater than"
        Set ar = New arValidation
        ar.Printer.Orientation = ddOPortrait
        If Not ConvertToDBL(txtCheck, tmpDouble) Then
            MsgBox "Invalid value in criterion box"
            Exit Sub
        Else
            Me.txtCheck = Format(tmpDouble, "Percent")
            If MsgBox("Confirm you want to list all counted products where the most recent discount is greater than " & txtCheck & " percent", vbQuestion + vbOKCancel, "Confirm request") = vbCancel Then
                Exit Sub
            End If
        End If
        strTitle = "Difference between R.R.P. and cost is greater than " & tmpDouble & " percent"
        strSQL = "SELECT tSTKTKEL.*, tProduct.* FROM tProduct INNER JOIN tSTKTKEL on tSTKTKEL.STKTKEL_P_ID = tproduct.P_ID WHERE (((P_SP - P_Cost) / CDBL(P_SP + .1) > " & tmpDouble & ") OR ((P_SP - P_Cost) / CDBL(P_SP + .1) < 0 ))AND P_Cost > 0 AND P_SP > 0 AND STKTKEL_Qty > 0 AND STKTKEL_TR_ID = " & oSA.TransactionID & " ORDER BY P_TITLE"
        rs.Open strSQL, oPC.CO, adOpenKeyset
        ar.Component rs, strTitle
        ar.Caption = strTitle
        ar.Show
        Set rs = Nothing
        Set ar = Nothing
    End Select
    If strSQL = "" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    Me.SB1.Panels(2).Text = ""
    cmdSave.Enabled = False
    Exit Sub
    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdReport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim strSQL
Dim rs As ADODB.Recordset
    If Not IsNumeric(txtCorrection) Then Exit Sub
    strSQL = "SELECT tSTKTKEL.* FROM tSTKTKEL WHERE tSTKTKEL.STKTKEL_ID = " & CLng(txtID)
    Set rs = New ADODB.Recordset
    rs.Open strSQL, oPC.CO, adOpenKeyset, adLockOptimistic
    rs.Fields("STKTKEL_Qty") = CLng(Me.txtCorrection)
    rs.Fields("STKTKEL_Difference") = CLng(txtCorrection) - CLng(txtBal)
    rs.Update
    rs.Close
    Set rs = Nothing
    cmdSave.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdSave_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim strSQL As String

    bViewOnly = False
    bLoading = True
    lblCheckList.Caption = "Check List " _
                & "1. You have taken a backup of the database immediately before continuing with the next step" & vbCrLf _
                & "2. You have cancelled all old purchase orders " & vbCrLf _
                & "           ( use the data management option on the Console application" & vbCrLf _
                & "3. You have cancelled all old purchase orders " & vbCrLf _
                & "           ( use the data management option on the Console application" & vbCrLf _
                & "4. You have run a dayend update procedure after the most recent capture operation" & vbCrLf _
                & "           ( e.g. Delivery, Transfer, Invoice etc.)."

    LoadControls
    Me.SSTab1.Tab = 0
    LoadStaff
    loadOtherSA
    bLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadStaff()
    On Error GoTo errHandler
    Set tlStaff = New z_TextList
    tlStaff.Load ltStaff
    LoadCombo cboOperator, tlStaff
    cboOperator.ListIndex = 0
    oSA.SalesPersonID = tlStaff.Key(cboOperator)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadStaff"
End Sub
Private Sub loadOtherSA()
    On Error GoTo errHandler
Dim strSQL As String
    Set tlSAO = New z_TextList
    tlSAO.Load ltStockTake, "ISSUED"
    LoadCombo Me.cboOtherstocktakes, tlSAO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.loadOtherSA"
End Sub
Private Sub cmdDeleteSA_Click()
    On Error GoTo errHandler
Dim oSA As a_Stktke
    Set oSA = New a_Stktke
    Screen.MousePointer = vbHourglass
    oSA.Load tlSA.Key(lstStocktakes.List(lstStocktakes.ListIndex))
    oSA.BeginEdit
    oSA.Delete
    oSA.ApplyEdit
    LoadControls
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdDeleteSA_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlSA = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSettings_Click()
    On Error GoTo errHandler
Dim frm As New frmSettings
    frm.Component strTemplate, flgShowWORD
    frm.Show vbModal
    strTemplate = frm.Template
    flgShowWORD = frm.ShowWORD
    SB1.Panels(1).Text = "Template: " & strTemplate & IIf(flgShowWORD, " (visible)", "(Background)")
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSettings_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuViewExisiting_Click()
    On Error GoTo errHandler
Dim oSA As a_Stktke
    Set oSA = New a_Stktke
    'oSA.Load
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuViewExisiting_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub oSA_MaxImportRows(i As Long)
    On Error GoTo errHandler
    Me.PB1.Max = i
    Me.PB1.Min = 0
    Me.PB1.Value = 0
    Me.PB1.Visible = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSA_MaxImportRows(i)", i, EA_NORERAISE
    HandleError
End Sub

Private Sub oSA_LineCOuntChange(i As Long)
    On Error GoTo errHandler
    Me.PB1.Value = i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSA_LineCOuntChange(i)", i, EA_NORERAISE
    HandleError
End Sub
Private Sub oSA_FinishedImporting()
    On Error GoTo errHandler
    Me.SB1.Panels(2).Text = "Consolidating duplicate codes. . ."
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSA_FinishedImporting", , EA_NORERAISE
    HandleError
End Sub
Private Sub oSA_BuildingTA()
    On Error GoTo errHandler
    Me.SB1.Panels(2).Text = "Building transaction. . ."
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSA_BuildingTA", , EA_NORERAISE
    HandleError
End Sub
Private Sub oSA_Zeroising()
    On Error GoTo errHandler
    Me.SB1.Panels(2).Text = "Zeroising titles absent from stock-take. . ."
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSA_Zeroising", , EA_NORERAISE
    HandleError
End Sub

Private Sub oSA_ImportFile(pName As String)
    On Error GoTo errHandler
    Me.SB1.Panels(2).Text = "Importing . . . " & pName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSA_ImportFile(pName)", pName, EA_NORERAISE
    HandleError
End Sub

Private Sub SSTab1_click(Previoustab As Integer)
    If bViewOnly = True Then
        If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Or SSTab1.Tab = 2 Or SSTab1.Tab = 4 Or SSTab1.Tab = 5 Then
            SSTab1.Tab = 3
        End If
    End If
End Sub

Private Sub txtDateTime_Change()
    On Error GoTo errHandler
    If IsDate(txtDateTime) Then
        dteDateTime = CDate(txtDateTime)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtDateTime_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtSAID_Change()
    On Error GoTo errHandler
    Me.cmdFetch.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtSAID_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtID_Change()
    cmdSave.Enabled = False
End Sub


Private Sub txtNegAdj_Validate(Cancel As Boolean)
    Cancel = Not IsDate(txtNegAdj)
End Sub
Private Sub PrintAdjustMentReport()
Dim arB As arValidation_B
Dim tmpNumber As Long
Dim rs As ADODB.Recordset

        Set arB = New arValidation_B
        arB.Printer.Orientation = ddOPortrait
        strSQL = "SELECT tSTKTKEL.STKTKEL_ID,STKTKEL_P_ID,STKTKEL_QTY,ISNULL(STKTKEL_Difference,0) as STKTKEL_Difference,tProduct.* FROM tProduct INNER JOIN tSTKTKEL on tSTKTKEL.STKTKEL_P_ID = tproduct.P_ID Where dbo.GetMod(STKTKEL_Difference) > " & tmpNumber & " AND STKTKEL_TR_ID = " & oSA.TransactionID & " ORDER BY P_TITLE"
        Set rs = New ADODB.Recordset
        rs.Open strSQL, oPC.CO, adOpenKeyset
        strTitle = "Provisional adjustments"
        arB.Caption = strTitle
        arB.Component rs, strTitle
        arB.Show
        Set rs = Nothing
        Set arB = Nothing
End Sub
Private Sub PrintAdjustMentReportFinal()
Dim arD As arValidation_D
Dim tmpNumber As Long
Dim rs As ADODB.Recordset

        Set arD = New arValidation_D
        arD.Printer.Orientation = ddOPortrait
        strSQL = "SELECT tSTKTKEL.STKTKEL_ID,STKTKEL_P_ID,STKTKEL_QTY,STKTKE_CUTOFFDATE,ISNULL(STKTKEL_Difference,0) as STKTKEL_Difference,tProduct.* FROM tProduct INNER JOIN tSTKTKEL on tSTKTKEL.STKTKEL_P_ID = tproduct.P_ID INNER JOIN tSTKTKE ON STKTKEL_TR_ID = STKTKE_ID Where dbo.GetMod(STKTKEL_Difference) > " & tmpNumber & " AND STKTKEL_TR_ID = " & oSA.TransactionID & " ORDER BY P_TITLE"
        Set rs = New ADODB.Recordset
        rs.Open strSQL, oPC.CO, adOpenKeyset
        strTitle = "Final adjustments for stock-take with cutoff: " & rs.Fields("STKTKE_CUTOFFDATE")
        arD.Caption = strTitle
        arD.Component rs, strTitle
        arD.Show
        Set rs = Nothing
        Set arD = Nothing
End Sub

Private Sub PrintValidation_C()
Dim arC As arValidation_C
Dim rs As ADODB.Recordset
        
        Set arC = New arValidation_C
        arC.Printer.Orientation = ddOPortrait
        Set rs = New ADODB.Recordset
        rs.Open strSQL, oPC.CO, adOpenKeyset
        arC.Caption = strTitle
        arC.Component rs, strTitle
        arC.Show
        Set rs = Nothing
        Set arC = Nothing

End Sub
