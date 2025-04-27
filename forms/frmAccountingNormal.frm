VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmMainNormal 
   Caption         =   "Documents for posting to accounting"
   ClientHeight    =   7290
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12645
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTestPastel 
      BackColor       =   &H00CED0BF&
      Caption         =   "Test Pastel connection"
      Height          =   390
      Left            =   420
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6195
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   2490
      TabIndex        =   62
      Text            =   "Text2"
      Top             =   6180
      Width           =   9975
   End
   Begin VB.CommandButton cmdTx 
      BackColor       =   &H00CED0BF&
      Caption         =   "Transmission controls"
      Height          =   465
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6510
      Visible         =   0   'False
      Width           =   1785
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   21
      Top             =   7005
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10874
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   10874
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00CED0BF&
      Caption         =   "Save column widths"
      Height          =   285
      Left            =   10365
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3930
      Visible         =   0   'False
      Width           =   1620
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   10663
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Stage 1. Unposted documents "
      TabPicture(0)   =   "frmAccountingNormal.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdAudit"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdFilterstore"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "G"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCount"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Stage 2. Pre-posting checks"
      TabPicture(1)   =   "frmAccountingNormal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "lblCOunt3"
      Tab(1).Control(5)=   "G3"
      Tab(1).Control(6)=   "txtCS"
      Tab(1).Control(7)=   "cmdFetch2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Text1"
      Tab(1).Control(9)=   "Frame4"
      Tab(1).Control(10)=   "cmdEmailProblems"
      Tab(1).Control(11)=   "cmdUntick_2"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Stage 3. Final review and post documents"
      TabPicture(2)   =   "frmAccountingNormal.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblCOunt4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "G4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdPost"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdShowReady"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdFetch3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Stage 4. Review completed postings"
      TabPicture(3)   =   "frmAccountingNormal.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdUnsent"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(2)=   "G2"
      Tab(3).Control(3)=   "lblCount2"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdUnsent 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Un-sent"
         Height          =   615
         Left            =   -63510
         Picture         =   "frmAccountingNormal.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   540
         Width           =   765
      End
      Begin VB.CommandButton cmdAudit 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Audit transactions"
         Height          =   420
         Left            =   -65565
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   5235
         Width           =   1455
      End
      Begin VB.CommandButton cmdUntick_2 
         BackColor       =   &H00CED0BF&
         Caption         =   "Untick all"
         Height          =   285
         Left            =   -66690
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdFetch3 
         BackColor       =   &H00CED0BF&
         Enabled         =   0   'False
         Height          =   555
         Left            =   10065
         MaskColor       =   &H00D3D3CB&
         Picture         =   "frmAccountingNormal.frx":03FA
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   690
         Width           =   705
      End
      Begin VB.CommandButton cmdEmailProblems 
         BackColor       =   &H00CED0BF&
         Caption         =   "Email problems to source branches"
         Height          =   420
         Left            =   -67335
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5505
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   -71160
         TabIndex        =   36
         Top             =   450
         Width           =   4380
         Begin VB.OptionButton optAll_2 
            Caption         =   "All"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   3615
            TabIndex        =   41
            Top             =   150
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optCN_2 
            Caption         =   "CN"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   2805
            TabIndex        =   40
            Top             =   150
            Width           =   570
         End
         Begin VB.OptionButton optINV_2 
            Caption         =   "Inv."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   1995
            TabIndex        =   39
            Top             =   150
            Width           =   615
         End
         Begin VB.OptionButton optRET_2 
            Caption         =   "Retn."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   1050
            TabIndex        =   38
            Top             =   150
            Width           =   690
         End
         Begin VB.OptionButton optSI_2 
            Caption         =   "Supp.Inv"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   45
            TabIndex        =   37
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   -73410
         TabIndex        =   35
         Top             =   465
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.CommandButton cmdFetch2 
         BackColor       =   &H00D3D3CB&
         Enabled         =   0   'False
         Height          =   570
         Left            =   -63690
         MaskColor       =   &H00D3D3CB&
         Picture         =   "frmAccountingNormal.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   540
         Width           =   705
      End
      Begin VB.CommandButton cmdShowReady 
         BackColor       =   &H00CED0BF&
         Caption         =   "Send to Pastel"
         Height          =   390
         Left            =   20235
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   9960
         Width           =   1620
      End
      Begin VB.Frame Frame3 
         Caption         =   "Show posting log"
         ForeColor       =   &H8000000D&
         Height          =   975
         Left            =   -74220
         TabIndex        =   26
         Top             =   435
         Width           =   10635
         Begin VB.OptionButton optRange4 
            Caption         =   "Range"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   4755
            TabIndex        =   55
            Top             =   330
            Width           =   810
         End
         Begin VB.OptionButton optG2Today 
            Caption         =   "Today"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton optG2Yesterday 
            Caption         =   "Yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Width           =   1050
         End
         Begin VB.CommandButton cmdFetch4 
            BackColor       =   &H00D3D3CB&
            Height          =   435
            Left            =   8325
            Picture         =   "frmAccountingNormal.frx":0B0E
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   315
            Width           =   495
         End
         Begin VB.OptionButton optG2DBY 
            Caption         =   "Day before yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2355
            TabIndex        =   27
            Top             =   375
            Width           =   2280
         End
         Begin MSComCtl2.DTPicker DPRecFrom4 
            Height          =   300
            Left            =   6735
            TabIndex        =   56
            Top             =   210
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39627
         End
         Begin MSComCtl2.DTPicker DPRecTo4 
            Height          =   300
            Left            =   6735
            TabIndex        =   57
            Top             =   555
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39627
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Between date"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5550
            TabIndex        =   59
            Top             =   255
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "and date (incl)"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5550
            TabIndex        =   58
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.TextBox txtCS 
         Height          =   465
         Left            =   -73665
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   7515
         Width           =   5880
      End
      Begin VB.CommandButton cmdPost 
         BackColor       =   &H00CED0BF&
         Caption         =   "Send for posting"
         Height          =   390
         Left            =   10065
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   5085
         Width           =   1620
      End
      Begin VB.Frame Frame1 
         Caption         =   "Show documents received"
         ForeColor       =   &H8000000D&
         Height          =   945
         Left            =   -74775
         TabIndex        =   8
         Top             =   480
         Width           =   9105
         Begin VB.CommandButton cmdReloadintray 
            BackColor       =   &H00CED0BF&
            Caption         =   "Reload in-tray"
            Height          =   315
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   585
            Width           =   1200
         End
         Begin VB.OptionButton optRange 
            Caption         =   "Range"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   4770
            TabIndex        =   24
            Top             =   330
            Width           =   810
         End
         Begin VB.OptionButton optDBY 
            Caption         =   "Day before yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   2145
            TabIndex        =   19
            Top             =   315
            Width           =   1800
         End
         Begin VB.OptionButton opttoday 
            Caption         =   "Today"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   315
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton optYesterday 
            Caption         =   "Yesterday"
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1020
            TabIndex        =   10
            Top             =   315
            Width           =   1050
         End
         Begin VB.CommandButton cmdFetch1 
            BackColor       =   &H00D3D3CB&
            Height          =   540
            Left            =   8235
            Picture         =   "frmAccountingNormal.frx":0E98
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   240
            Width           =   660
         End
         Begin MSComCtl2.DTPicker DPRecFrom 
            Height          =   300
            Left            =   6750
            TabIndex        =   12
            Top             =   210
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39627
         End
         Begin MSComCtl2.DTPicker DPRecTO 
            Height          =   300
            Left            =   6750
            TabIndex        =   22
            Top             =   555
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   39627
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "and date (incl)"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5565
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Between date"
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   5565
            TabIndex        =   13
            Top             =   255
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   -74205
         TabIndex        =   2
         Top             =   1395
         Width           =   4965
         Begin VB.OptionButton optALL 
            Caption         =   "All"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   3615
            TabIndex        =   7
            Top             =   150
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton optCN 
            Caption         =   "CN"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   2805
            TabIndex        =   6
            Top             =   150
            Width           =   570
         End
         Begin VB.OptionButton optINV 
            Caption         =   "Inv."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   1995
            TabIndex        =   5
            Top             =   150
            Width           =   615
         End
         Begin VB.OptionButton optRET 
            Caption         =   "Retn."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   1050
            TabIndex        =   4
            Top             =   150
            Width           =   690
         End
         Begin VB.OptionButton optSI 
            Caption         =   "Supp.Inv"
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   45
            TabIndex        =   3
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdFilterstore 
         BackColor       =   &H00D3D3CB&
         Height          =   360
         Left            =   -69210
         MaskColor       =   &H00D3D3CB&
         Picture         =   "frmAccountingNormal.frx":1222
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1530
         Width           =   480
      End
      Begin TrueOleDBGrid60.TDBGrid G 
         Height          =   3660
         Left            =   -74790
         OleObjectBlob   =   "frmAccountingNormal.frx":15AC
         TabIndex        =   14
         Top             =   1995
         Width           =   9165
      End
      Begin TrueOleDBGrid60.TDBGrid G2 
         Height          =   4065
         Left            =   -74205
         OleObjectBlob   =   "frmAccountingNormal.frx":5CE6
         TabIndex        =   31
         Top             =   1395
         Width           =   10650
      End
      Begin TrueOleDBGrid60.TDBGrid G4 
         Height          =   4785
         Left            =   390
         OleObjectBlob   =   "frmAccountingNormal.frx":A565
         TabIndex        =   33
         Top             =   690
         Width           =   9615
      End
      Begin TrueOleDBGrid60.TDBGrid G3 
         Height          =   4305
         Left            =   -74730
         OleObjectBlob   =   "frmAccountingNormal.frx":F040
         TabIndex        =   42
         Top             =   1140
         Width           =   11745
      End
      Begin VB.Label lblCount2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74115
         TabIndex        =   52
         Top             =   5550
         Width           =   2445
      End
      Begin VB.Label lblCOunt4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   405
         TabIndex        =   51
         Top             =   5520
         Width           =   2445
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         Caption         =   "All un-posted documents"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   1140
         TabIndex        =   50
         Top             =   465
         Width           =   2850
      End
      Begin VB.Label lblCOunt3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74745
         TabIndex        =   49
         Top             =   5460
         Width           =   2445
      End
      Begin VB.Label Label7 
         Caption         =   "Documents to post"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   -74415
         TabIndex        =   45
         Top             =   900
         Width           =   5625
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -71745
         TabIndex        =   44
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Show only store"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74625
         TabIndex        =   43
         Top             =   525
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Day reflects most recently transmitted date (in the case of re-transmissions)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -64350
         TabIndex        =   17
         Top             =   6375
         Width           =   5910
      End
      Begin VB.Label lblCount 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74790
         TabIndex        =   16
         Top             =   5685
         Width           =   2445
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Filter"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1605
         Width           =   360
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSQL 
         Caption         =   "SQL"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug on"
      End
      Begin VB.Menu mnuColumnsSave 
         Caption         =   "Save column widths"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMainNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strBy As String
Dim dteSince As Date
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim XC As XArrayDB
Dim XD As XArrayDB
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As ADODB.Recordset
Dim dteSelected As Date
Dim bAllOK As Boolean
Dim currentPastelPeriod As String

Private Sub cmd_Click()
    On Error GoTo ErrHandler

    
    bAllOK = ValidateNewRows
    
    If bAllOK Then
        InsertToPastelLedger
    End If
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmd_Click"
    HandleError
End Sub

Private Sub cmdAudit_Click()
    On Error GoTo ErrHandler
Dim f As New frmMissingTransactions

    f.Show vbModal
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdAudit_Click"
    HandleError
End Sub
Private Sub cmdEmailProblems_Click()
    On Error GoTo ErrHandler
Dim rsBranch As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim oXML As New z_XML
Dim oEmail As New z_HOEmail
Dim res As Boolean
Dim dummy As String

    rsBranch.CursorLocation = adUseClient
    rs.CursorLocation = adUseClient
'For each branch
    rsBranch.Open "SELECT STORE_CODE,STORE_EMAIL,STORE_NAME,STORE_CONTACT FROM tStore ORDER BY STORE_CODE", oPC.COShort, adOpenStatic, adLockOptimistic
    Do While Not rsBranch.EOF
    'Find all items with acno problems separated by source branch
        rs.Open "SELECT P_TPNAME,P_DateReceivedAtHO,P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_AMT,P_REF FROM tInTRay WHERE P_INVALIDDOC <>0 AND P_SRC = '" & FNS(rsBranch.Fields(0)) & "' ORDER BY dbo.fnStripNonnumericChars(P_DESCR)", oPC.COShort, adOpenStatic, adLockOptimistic
    'Produce an XML document
        If rs.RecordCount > 0 Then
            oXML.GenerateXMLBranchReport rs, FNS(rsBranch.Fields("STORE_CODE")), FNS(rsBranch.Fields("STORE_NAME")), FNS(rsBranch.Fields("STORE_CONTACT")), FNS(rsBranch.Fields("STORE_EMAIL"))
            oXML.CreateFiles "BCE", dummy
            oEmail.PrepareSendMail
            res = oEmail.SendOneMessage("Head office Pastel invalid account numbers", "Please examine the attached document and correct the account numbers for the suppliers.", oXML.PDF_Filename, Format(Date, "dd-mm-yyyy"), CStr(rsBranch.Fields(1)), "", oPC.EMAIL_SenderName, oPC.EMAIL_EmailFrom)
            
        End If
'Produce an HTML and a spreadsheet
        rs.Close
        rsBranch.MoveNext
    Loop
'Send to each branch
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdEmailProblems_Click"
    HandleError
End Sub



'Private Sub cmdFetch3_Click()
'    On Error GoTo errHandler
'
''Transfer selected rows from tIntray to tJournals
'    oPC.COShort.Execute "TransferFromIntray"
'    cmdUnposted_Click
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdFetch3_Click"
'End Sub

Private Sub cmdFetch3_Click()
    On Error GoTo ErrHandler
    If Not XD Is Nothing Then
        XD.Clear
        G4.ReBind
    End If
    
    Set XD = New XArrayDB
    If Not rs4 Is Nothing Then
        If rs4.State <> 0 Then rs4.Close
    End If
    Set rs4 = New ADODB.Recordset
    rs4.CursorLocation = adUseClient
    
    rs4.Open "Select P_SRC,P_DATE,P_JOURNALTYPE,P_TPNAME,P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT,P_ACTION,ID FROM tINTRAY " _
             & " WHERE ISNULL(P_POSTEDDATE,'19000101') < '20000101' And ISNULL(P_INVALIDDOC,0) = 0 And ISNULL(P_ACTION,0) = 1", oPC.COShort, adOpenDynamic, adLockOptimistic
    
    If rs4.EOF Then
        Screen.MousePointer = vbDefault
        XD.Clear
        XD.ReDim 1, 0, 1, 10
        G4.ReBind
        G4.Refresh
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
    
    LoadGrid4
    XD.QuickSort XD.LowerBound(1), XD.UpperBound(1), 1, 0, XTYPE_STRING
    G4.Array = XD
    G4.ReBind
    G4.Refresh
    lblCOunt4.Caption = CStr(rs4.RecordCount) & " records"
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFetch3_Click"
    HandleError
End Sub







Private Sub cmdPost_Click()
    On Error GoTo ErrHandler
Dim res As Boolean

    If MsgBox("Current Pastel period: " & FindPastelPeriod(Date) & ". If this is wrong please click on CANCEL", vbInformation + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oPC.COShort.Execute "TransferFromIntray"
    res = InsertToPastelLedger
    If Not XD Is Nothing Then
        XD.Clear
        G4.ReBind
    End If
    
    If res Then MarkJournalsExistingNowInPastel
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdPost_Click"
    HandleError
End Sub
Private Sub MarkJournalsExistingNowInPastel()
    On Error GoTo ErrHandler
Dim rsJournals As New ADODB.Recordset
Dim rsPas As New ADODB.Recordset
Dim s As String
Dim cnnPas As New ADODB.Connection

    rsJournals.Open "SELECT * FROM tJournals WHERE P_POSTEDDATE IS NULL", oPC.COShort, adOpenDynamic, adLockOptimistic
    rsPas.CursorLocation = adUseClient
    cnnPas.ConnectionString = oPC.PastelConnectionstring
    cnnPas.CommandTimeout = 0
    cnnPas.Open
    Do While Not rsJournals.EOF
        s = "SELECT * FROM LedgerTransactions WHERE ACCNUMBER = '" & FNS(rsJournals.Fields("P_ACNO")) & "' AND DESCRIPTION LIKE '" & rsJournals.Fields("P_DESCR") & "%'"
        MsgBox "Commendstring= " & s
        rsPas.Open s, cnnPas
        If rsPas.RecordCount > 0 Then
             rsJournals.Fields("P_POSTEDDATE") = Now()
             rsJournals.Update
        End If
        rsPas.Close
        rsJournals.MoveNext
    Loop
    Set rsPas = Nothing
    cnnPas.Close
    
    rsJournals.Close
    Set rsJournals = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MarkJournalsExistingNowInPastel"
End Sub
Private Sub cmdTestPastel_Click()
    On Error GoTo ErrHandler
Dim cnnPas As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim rsPas As ADODB.Recordset

    cnnPas.ConnectionString = Text2
    cnnPas.Open Text2
    Set rsPas = New ADODB.Recordset
    rsPas.CursorLocation = adUseClient
    rsPas.Open "Select Count(*) FROM LedgerTransactions", cnnPas
    MsgBox "Count of ledger transactions = " & CStr(rsPas.Fields(0))
    rsPas.Close
'
    MsgBox "Connected"
    
    MsgBox "Current period = " & FindPastelPeriod(Date)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdTestPastel_Click"
    HandleError
End Sub
Private Function ValidateNewRows()
    On Error GoTo ErrHandler
Dim cnnPas As New ADODB.Connection
Dim rsPas As New ADODB.Recordset
Dim strMissingAcnos As String
Dim bOK As Boolean
Dim strPos As String
 
    'first load the recordset
    If Not rs3 Is Nothing Then
        If rs3.State = 1 Then rs3.Close
    End If
    Set rs3 = Nothing
    Set rs3 = New ADODB.Recordset
    rs3.CursorLocation = adUseClient
            strPos = "0"
    
    rs3.Open "Select P_SRC,P_DATE,P_JOURNALTYPE,P_TPNAME,P_ACNO,P_DESCR,P_REF,P_AMT,P_UNPOSTED,P_INVALIDDOC,ID,P_ACTION FROM tInTray " _
             & " WHERE ISNULL(P_POSTEDDATE,'19000101') < '20000101' ORDER BY P_JOURNALTYPE", oPC.COShort, adOpenDynamic, adLockOptimistic
    
    If rs3.EOF Then
        Screen.MousePointer = vbDefault
        XC.Clear
        XC.ReDim 1, 0, 1, 10
        G3.ReBind
        MsgBox "No records found", vbInformation, "Status"
        Exit Function
    End If
   
   
'Checking that account numbers for suppliers exist
    bOK = True
            strPos = "1"
    rsPas.CursorLocation = adUseClient
  '  MsgBox "Connection in ValidateNewRows = " & oPC.PastelConnectionstring
    cnnPas.ConnectionString = oPC.PastelConnectionstring
    cnnPas.Open
    strMissingAcnos = ""
    rs3.MoveFirst
            strPos = "2"
    Do While Not rs3.EOF
        If Left(FNS(rs3.Fields("P_JournalType")), 2) = "CR" Then
            rsPas.Open "SELECT * FROM SupplierMaster WHERE SupplCode = '" & FNS(rs3.Fields("P_ACNO")) & "'", cnnPas
        Else
            If Left(FNS(rs3.Fields("P_JournalType")), 2) = "DB" Then
                rsPas.Open "SELECT * FROM CustomerMaster WHERE CustomerCode = '" & FNS(rs3.Fields("P_ACNO")) & "'", cnnPas
            End If
        End If
'MsgBox "Acno = " & FNS(rs3.Fields("P_ACNO"))
        If rsPas.RecordCount = 0 Or FNS(rs3.Fields("P_ACNO")) = "" Then
            bOK = False
            strPos = "3"
            rs3.Fields("P_INVALIDDOC") = 1
          '  MsgBox "Update pos"
            rs3.Update
            strMissingAcnos = strMissingAcnos & IIf(strMissingAcnos > "", ", ", "") & IIf(FNS(rs3.Fields("P_ACNO")) = "", "<missing>", FNS(rs3.Fields("P_ACNO")))
            strPos = "4"
        Else
            rs3.Fields("P_INVALIDDOC") = 0
            rs3.Update
        End If
        rsPas.Close
        rs3.MoveNext
    Loop
    rs3.Close
    Set rs3 = Nothing
  '  MsgBox "Records missing: " & strMissingAcnos
    cnnPas.Close
    Set cnnPas = Nothing
    
    ValidateNewRows = bOK
    
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ValidateNewRows"
End Function
Private Function InsertToPastelLedger() As Boolean
    On Error GoTo ErrHandler
Dim cnnPas As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim rsSQL As New ADODB.Recordset
Dim SQL As String
    InsertToPastelLedger = False
    
    cnnPas.ConnectionString = oPC.PastelConnectionstring
    cnnPas.Open
    cnnPas.BeginTrans
    rsSQL.CursorLocation = adUseClient
    rsSQL.Open "SELECT SQLString_query,SQLString_Type,SQLString_SequenceNo FROM tAccounting_SQL", oPC.COShort, adOpenStatic
    
    'Post suppliers invoices here
    rs.CursorLocation = adUseClient
    rs.Open "SELECT tJournals.*,STORE_CostOfSalesAccount,STORE_CODE FROM tJournals JOIN tStore ON P_SRC = STORE_CODE WHERE P_JOURNALTYPE = 'CRIV' AND ISNULL(P_POSTEDDATE,'1950-01-01') < '1952-01-01' ORDER BY P_ACNO ", oPC.COShort, adOpenStatic
    Do While Not rs.EOF
        rsSQL.Filter = "SQLString_TYPE = 'PURCHASES' and SQLString_SequenceNo = 1"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("P_ACNO")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS(rs.Fields("P_AMT")))
        SQL = Replace(SQL, "**P_TAXAMT**", FNS(rs.Fields("P_TAXAMT")))
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", CStr(CDbl(rs.Fields("P_TAXAMT"))))    ''  **ThisCurrTaxAmt**
        SQL = Replace(SQL, "**P_TAXTYPE**", FNN(rs.Fields("P_TAXTYPE")))
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(rs.Fields("STORE_CostOfSalesAccount")), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
    'Then post to contra account: If PERIODIC then Supplier control account. IF PERPETUAL then Inventory Control Account
        rsSQL.Filter = "SQLString_TYPE = 'PURCHASES' and SQLString_SequenceNo = 2"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("STORE_CostOfSalesAccount")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS((CDbl(rs.Fields("P_AMT")) - CDbl(rs.Fields("P_TAXAMT")))) * -1)
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
    'Then post to VAT
        rsSQL.Filter = "SQLString_TYPE = 'PURCHASES' and SQLString_SequenceNo = 3"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(oPC.VATAccount))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", "ZContras")
        SQL = Replace(SQL, "**P_AMT**", FNS(CStr(CDbl(rs.Fields("P_TAXAMT")) * -1)))
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_REF")) & " - Tax entry")
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
        rs.MoveNext
    Loop
    rs.Close
    
'Post returns to suppliers here
    rs.Open "SELECT tJournals.*,STORE_CostOfSalesAccount,STORE_CODE FROM tJournals JOIN tStore ON P_SRC = STORE_CODE WHERE P_JOURNALTYPE = 'CRRT' AND ISNULL(P_POSTEDDATE,'1950-01-01') < '1952-01-01' ORDER BY P_ACNO ", oPC.COShort, adOpenStatic
    Do While Not rs.EOF
        rsSQL.Filter = "SQLString_TYPE = 'ReturnsToSuppliers' and SQLString_SequenceNo = 1"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("P_ACNO")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS(rs.Fields("P_AMT")))
        SQL = Replace(SQL, "**P_TAXAMT**", FNS(rs.Fields("P_TAXAMT")))
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", CStr(CDbl(rs.Fields("P_TAXAMT"))))    ''  **ThisCurrTaxAmt**
        SQL = Replace(SQL, "**P_TAXTYPE**", FNN(rs.Fields("P_TAXTYPE")))
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(rs.Fields("STORE_CostOfSalesAccount")), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
    'Then post to contra account: If PERIODIC then Supplier control account. IF PERPETUAL then Inventory Control Account
        rsSQL.Filter = "SQLString_TYPE = 'ReturnsToSuppliers' and SQLString_SequenceNo = 2"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("STORE_CostOfSalesAccount")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS((CDbl(rs.Fields("P_AMT")) - CDbl(rs.Fields("P_TAXAMT")))) * -1)
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
    'Then post to VAT
        rsSQL.Filter = "SQLString_TYPE = 'ReturnsToSuppliers' and SQLString_SequenceNo = 3"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(oPC.VATAccount))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", "ZContras")
        SQL = Replace(SQL, "**P_AMT**", FNS(CStr(CDbl(rs.Fields("P_TAXAMT")) * -1)))
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_REF")) & " - Tax entry")
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
        rs.MoveNext
    Loop
    rs.Close
    
  'Post customer invoices here
      rs.CursorLocation = adUseClient
    rs.Open "SELECT tJournals.*,STORE_CostOfSalesAccount,STORE_CODE FROM tJournals JOIN tStore ON P_SRC = STORE_CODE WHERE P_JOURNALTYPE = 'DBIV' AND ISNULL(P_POSTEDDATE,'1950-01-01') < '1952-01-01' ORDER BY P_ACNO ", oPC.COShort, adOpenStatic
    Do While Not rs.EOF
        rsSQL.Filter = "SQLString_TYPE = 'SALES' and SQLString_SequenceNo = 1"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("P_ACNO")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS(rs.Fields("P_AMT")))
        SQL = Replace(SQL, "**P_TAXAMT**", FNS(rs.Fields("P_TAXAMT")))
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", CStr(CDbl(rs.Fields("P_TAXAMT"))))    ''  **ThisCurrTaxAmt**
        SQL = Replace(SQL, "**P_TAXTYPE**", FNN(rs.Fields("P_TAXTYPE")))
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(rs.Fields("STORE_CostOfSalesAccount")), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
    'Then post to contra account: If PERIODIC then Cost of sales account. IF PERPETUAL then Inventory Control Account
        rsSQL.Filter = "SQLString_TYPE = 'SALES' and SQLString_SequenceNo = 2"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("STORE_CostOfSalesAccount")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS((CDbl(rs.Fields("P_AMT")) - CDbl(rs.Fields("P_TAXAMT")))) * -1)
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
    'Then post to VAT
        rsSQL.Filter = "SQLString_TYPE = 'SALES' and SQLString_SequenceNo = 3"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(oPC.VATAccount))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", "ZContras")
        SQL = Replace(SQL, "**P_AMT**", FNS(CStr(CDbl(rs.Fields("P_TAXAMT")) * -1)))
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_REF")) & " - Tax entry")
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
        rs.MoveNext
    Loop
    rs.Close

'Post customer returns(Our credit notes to customers) here
    rs.CursorLocation = adUseClient
    rs.Open "SELECT tJournals.*,STORE_CostOfSalesAccount,STORE_CODE FROM tJournals JOIN tStore ON P_SRC = STORE_CODE WHERE P_JOURNALTYPE = 'DBRT' AND ISNULL(P_POSTEDDATE,'1950-01-01') < '1952-01-01' ORDER BY P_ACNO ", oPC.COShort, adOpenStatic
    Do While Not rs.EOF
        rsSQL.Filter = "SQLString_TYPE = 'SALESRETURNS' and SQLString_SequenceNo = 1"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("P_ACNO")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS(rs.Fields("P_AMT")))
        SQL = Replace(SQL, "**P_TAXAMT**", FNS(rs.Fields("P_TAXAMT")))
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", CStr(CDbl(rs.Fields("P_TAXAMT"))))    ''  **ThisCurrTaxAmt**
        SQL = Replace(SQL, "**P_TAXTYPE**", FNN(rs.Fields("P_TAXTYPE")))
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(rs.Fields("STORE_CostOfSalesAccount")), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
    'Then post to contra account: If PERIODIC then Supplier control account. IF PERPETUAL then Inventory Control Account
        rsSQL.Filter = "SQLString_TYPE = 'SALESRETURNS' and SQLString_SequenceNo = 2"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(rs.Fields("STORE_CostOfSalesAccount")))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", Right(FNS(rs.Fields("P_REF")), 7) & FNS(rs.Fields("STORE_CODE")))
        SQL = Replace(SQL, "**P_AMT**", FNS((CDbl(rs.Fields("P_AMT")) - CDbl(rs.Fields("P_TAXAMT")))) * -1)
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_DESCR")))
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
    'Then post to VAT
        rsSQL.Filter = "SQLString_TYPE = 'SALESRETURNS' and SQLString_SequenceNo = 3"
        SQL = rsSQL.Fields(0)
        SQL = Replace(SQL, "**ACNO**", FNS(oPC.VATAccount))
        SQL = Replace(SQL, "**P_DATE**", Format(FND(rs.Fields("P_DATE")), "YYYY-MM-DD"))
        SQL = Replace(SQL, "**P_REF**", "ZContras")
        SQL = Replace(SQL, "**P_AMT**", FNS(CStr(CDbl(rs.Fields("P_TAXAMT")) * -1)))
        SQL = Replace(SQL, "**P_TAXAMT**", "0")
        SQL = Replace(SQL, "**ThisCurrTaxAmt**", "0")
        SQL = Replace(SQL, "**P_TAXTYPE**", "0")
        SQL = Replace(SQL, "**P_DESCR**", FNS(rs.Fields("P_REF")) & " - Tax entry")
        SQL = Replace(SQL, "**CONTRA**", IIf(oPC.Periodic_or_Perpetual = "PERIODIC", FNS(oPC.SupplierControlAccount), oPC.InventoryControlAccount))
        SQL = Replace(SQL, "**currentPastelPeriod**", "1" & CStr(FindPastelPeriod(FND(rs.Fields("P_DATE")))))
        If mnuDebug.Checked Then MsgBox SQL
        cnnPas.Execute SQL
        
        rs.MoveNext
    Loop
    rs.Close

    
    Set rs = Nothing
    rsSQL.Close
    Set rsSQL = Nothing
    cnnPas.CommitTrans
    cnnPas.Close
    Set cnnPas = Nothing
    
    InsertToPastelLedger = True
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.InsertToPastelLedger", , , cnnPas
End Function
Private Function FindPastelPeriod(pDate As Date) As String
    On Error GoTo ErrHandler
Dim cnnPas As New ADODB.Connection
Dim rsPasPeriods As New ADODB.Recordset
Dim dte As Date

    If rsPasPeriods.State = 0 Then
        rsPasPeriods.CursorLocation = adUseClient
    
        cnnPas.ConnectionString = oPC.PastelConnectionstring
        cnnPas.Open
       
        rsPasPeriods.Open "SELECT PERSTARTTHIS01,PERSTARTTHIS02,PERSTARTTHIS03,PERSTARTTHIS04,PERSTARTTHIS05,PERSTARTTHIS06,PERSTARTTHIS07,PERSTARTTHIS08,PERSTARTTHIS09,PERSTARTTHIS10,PERSTARTTHIS11,PERSTARTTHIS12 FROM LedgerParameters", cnnPas, adOpenStatic
    End If
    
    dte = pDate
 
    If dte > CDate(rsPasPeriods.Fields(11)) Then
        FindPastelPeriod = "12"
    Else
    If dte > CDate(rsPasPeriods.Fields(10)) Then
        FindPastelPeriod = "11"
    Else
    If dte > CDate(rsPasPeriods.Fields(9)) Then
        FindPastelPeriod = "10"
    Else
    If dte > CDate(rsPasPeriods.Fields(8)) Then
        FindPastelPeriod = "09"
    Else
    If dte > CDate(rsPasPeriods.Fields(7)) Then
        FindPastelPeriod = "8"
    Else
    If dte > CDate(rsPasPeriods.Fields(6)) Then
        FindPastelPeriod = "07"
    Else
    If dte > CDate(rsPasPeriods.Fields(5)) Then
        FindPastelPeriod = "06"
    Else
    If dte > CDate(rsPasPeriods.Fields(4)) Then
        FindPastelPeriod = "05"
    Else
    If dte > CDate(rsPasPeriods.Fields(3)) Then
        FindPastelPeriod = "04"
    Else
    If dte > CDate(rsPasPeriods.Fields(2)) Then
        FindPastelPeriod = "03"
    Else
    If dte > CDate(rsPasPeriods.Fields(1)) Then
        FindPastelPeriod = "02"
    Else
    If dte > CDate(rsPasPeriods.Fields(0)) Then
        FindPastelPeriod = "01"
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    End If
    Set rsPasPeriods.ActiveConnection = Nothing
    cnnPas.Close

    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FindPastelPeriod"
End Function
Private Sub cmdFetch4_Click()
    On Error GoTo ErrHandler
Dim strPos As String

    Screen.MousePointer = vbHourglass
    If Not XB Is Nothing Then
        XB.Clear
        G2.ReBind
    End If
    Set rs2 = New ADODB.Recordset
  '  rs2.CursorLocation = adUseClient
strPos = "2"
    If optG2Today = True Then
strPos = "2.1"
            rs2.Open "Select P_POSTEDDATE,P_SRC,P_DATE,P_GLOBALTRID,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME, " _
                & " P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT,P_DateReceivedAtHO FROM tJournals " _
                & " WHERE dbo.dte(P_POSTEDDATE) = dbo.startOfDay(Getdate())" _
                & " ORDER BY P_POSTEDDATE,P_DESCR ", oPC.COShort
    Else
    If Me.optG2Yesterday = True Then
strPos = "2.2"
            rs2.Open "Select P_POSTEDDATE,P_SRC,P_DATE,P_GLOBALTRID,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME, " _
                & " P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT,P_DateReceivedAtHO FROM tJournals " _
                & " WHERE dbo.dte(P_POSTEDDATE) = dbo.startOfDay(DATEADD(d,-1,Getdate()))" _
                & " ORDER BY P_POSTEDDATE,P_DESCR ", oPC.COShort
    Else
    If optG2DBY = True Then
strPos = "2.3"
            rs2.Open "Select P_POSTEDDATE,P_SRC,P_DATE,P_GLOBALTRID,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME, " _
                & " P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT,P_DateReceivedAtHO FROM tJournals " _
                & " WHERE dbo.dte(P_POSTEDDATE) = dbo.startOfDay(DATEADD(d,-2,Getdate()))" _
                & " ORDER BY P_POSTEDDATE,P_DESCR ", oPC.COShort
    Else
strPos = "2.4"
            rs2.Open "Select P_POSTEDDATE,P_SRC,P_DATE,P_GLOBALTRID,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME, " _
                & " P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT,P_DateReceivedAtHO FROM tJournals " _
                & " WHERE P_POSTEDDATE >= dbo.startOfDay('" & ReverseDate(Me.DPRecFrom4) & "') AND " _
                & " P_POSTEDDATE <= dbo.EndOfDay('" & ReverseDate(Me.DPRecTo4) & "')" _
                & " ORDER BY P_POSTEDDATE,P_DESCR ", oPC.COShort
    End If
    End If
    End If
strPos = "3"
    If rs2.EOF Then
        Screen.MousePointer = vbDefault
        If XB Is Nothing Then Set XB = New XArrayDB
        XB.Clear
        XB.ReDim 1, 0, 1, 10
        G2.ReBind
        G2.Refresh
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
strPos = "3"
    LoadGrid2
    XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), 1, 0, XTYPE_STRING
    G2.Array = XB
    G2.ReBind
    G2.Refresh
strPos = "5"
    lblCount2.Caption = CStr(rs2.RecordCount) & " records"
    rs2.Close
    Set rs2 = Nothing
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFetch4_Click"
    HandleError
End Sub


Private Sub cmdFetch1_Click()
    On Error GoTo ErrHandler
    Set XA = New XArrayDB
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    If opttoday = True Then
            rs.Open "Select P_SRC,P_DATE,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME,P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT FROM tInTray " _
                & " WHERE dbo.dte(P_DateReceivedAtHO) = dbo.startOfDay(Getdate())", oPC.COShort
    Else
    If Me.optYesterday = True Then
            rs.Open "Select P_SRC,P_DATE,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME,P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT FROM tInTray " _
                & " WHERE dbo.dte(P_DateReceivedAtHO) = dbo.startOfDay(DATEADD(d,-1,Getdate()))", oPC.COShort
    Else
    If optDBY = True Then
            rs.Open "Select P_SRC,P_DATE,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME,P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT FROM tInTray " _
                & " WHERE dbo.dte(P_DateReceivedAtHO) = dbo.startOfDay(DATEADD(d,-2,Getdate()))", oPC.COShort
    Else
      '  MsgBox " WHERE dbo.dte(P_DateReceivedAtHO) >= dbo.startOfDay('" & ReverseDate(Me.DPRecFrom) & "' AND dbo.dte(P_DateReceivedAtHO) <= dbo.EndOfDay('" & ReverseDate(Me.DPRecTO) & "')"
            rs.Open "Select P_SRC,P_DATE,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME,P_ACNO,dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT FROM tInTray " _
                & " WHERE P_DateReceivedAtHO >= dbo.startOfDay('" & ReverseDate(Me.DPRecFrom) & "') AND P_DateReceivedAtHO <= dbo.EndOfDay('" & ReverseDate(Me.DPRecTO) & "')", oPC.COShort
    End If
    End If
    End If
    If rs.EOF Then
        Screen.MousePointer = vbDefault
        XA.Clear
        XA.ReDim 1, 0, 1, 10
        G.ReBind
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
    LoadGrid
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 1, 0, XTYPE_STRING
    G.Array = XA
    G.ReBind
    G.Refresh
    If optRET.Value = True Then
        optRET_Click
    Else
        If optCN.Value = True Then
            optCN_Click
        Else
            If optSI.Value = True Then
                optSI_Click
            Else
                If optINV.Value = True Then
                    optINV_Click
                End If
            End If
        End If
    End If
    lblCount.Caption = CStr(rs.RecordCount) & " records"
    Me.cmdFetch2.Enabled = True
    If Not XC Is Nothing Then
        XC.Clear
        G3.ReBind
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFetch1_Click"
    HandleError
End Sub

Private Sub cmdFilterstore_Click()
    If rs Is Nothing Then Exit Sub

    On Error GoTo ErrHandler
'    If txtStore > "" Then
'        rs.Filter = " P_SRC = '" & txtStore & "'"
'    Else
      '  rs.Filter = ""
'    End If
        rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFilterstore_Click"
    HandleError
End Sub



Private Sub cmdTx_Click()
    On Error GoTo ErrHandler
Dim frm As New frmTransmissionControl

    frm.Show vbModal

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdTx_Click"
    HandleError
End Sub

Private Sub cmdUnsent_Click()
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
'
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'
'        rs.Open "Select TR_ID,TR_CODE,TP_NAME,TP_ACNO,TR_TYPE,TR_CAPTUREDATE,TR_PROCESSINGDATE,TR_DATETOPASTEL FROM tTR JOIN tTP ON TR_TP_ID = TP_ID " _
'            & " WHERE ((TR_TYPE IN (4,8) AND TR_STATUS IN (3,4)) OR (TR_TYPE in (3,11) AND TR_STATUS =4))  AND ISNULL(TR_DATETOPASTEL,0) < '2000-01-01'", oPC.COShort
'
'    If rs.EOF Then
'        Screen.MousePointer = vbDefault
'        MsgBox "No records found", vbInformation, "Status"
'        Exit Sub
'    End If
'
'    LoadGrid
'    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 1, 0, XTYPE_STRING
'    G.Array = XA
'    G.ReBind
'    G.Refresh
'    lblCount = CStr(rs.RecordCount) & " records"
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdUnsent_Click"
    HandleError
End Sub

Private Sub cmdUntick_2_Click()
    On Error GoTo ErrHandler
Dim i As Integer

    For i = 1 To XC.UpperBound(1)
        XC(i, 9) = 0
    Next
    oPC.COShort.Execute "UPDATE tINTRAY SET P_ACTION = 0 "
    G3.ReBind
    If Not XD Is Nothing Then
        XD.Clear
        G4.ReBind
    End If
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdUntick_2_Click"
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrHandler
    mnuSaveLayout
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Command1_Click"
    HandleError
End Sub

Private Sub cmdFetch2_Click()
    On Error GoTo ErrHandler
    Set XC = New XArrayDB
    
    Screen.MousePointer = vbHourglass
    
    'Mark rows that have account numbers that don't exist on Pastel
    If MsgBox("Run Pastel validation?", vbYesNo) = vbYes Then
        ValidateNewRows
    End If
    
    'Mark rows that have already been posted
    oPC.COShort.Execute "MarkTrayItemsAlreadyPosted"
    
    'Mark rows that are too old
    oPC.COShort.Execute "MarkTrayItemsTooOld"
    
    'Mark lines action default
    oPC.COShort.Execute "MarkDefaultAction"
    
    Set rs3 = New ADODB.Recordset
    rs3.CursorLocation = adUseClient
    
    rs3.Open "Select P_SRC,P_DATE,dbo.fnExplainJournalTypes(P_JOURNALTYPE) as P_JOURNALTYPE,P_TPNAME,P_ACNO, " _
            & " dbo.fnStripNonnumericChars(P_DESCR) as P_DESCR,P_REF,P_AMT,P_UNPOSTED,dbo.InterpretInvalidCode(P_INVALIDDOC,P_UNPOSTED)  P_INVALIDDOC,ID,P_ACTION FROM tInTray " _
            & " WHERE ISNULL(P_POSTEDDATE,'19000101') < '20000101'", oPC.COShort, adOpenDynamic, adLockOptimistic
    
    If rs3.EOF Then
        Screen.MousePointer = vbDefault
        XC.Clear
        XC.ReDim 1, 0, 1, 10
        G3.ReBind
        MsgBox "No records found", vbInformation, "Status"
        Exit Sub
    End If
    
        
    'Show all tIntray rows with statuses
    LoadGrid3
    lblCOunt3.Caption = CStr(rs3.RecordCount) & " records"
    If optRET_2.Value = True Then
        optRET_2_Click
    Else
        If optCN_2.Value = True Then
            optCN_2_Click
        Else
            If optSI_2.Value = True Then
                optSI_2_Click
            Else
                If optINV_2.Value = True Then
                    optINV_2_Click
                End If
            End If
        End If
    End If
        
    XC.QuickSort XC.LowerBound(1), XC.UpperBound(1), 1, 0, XTYPE_STRING, 2, 0, XTYPE_STRING, 4, 0, XTYPE_DATE
    G3.Array = XC
    G3.ReBind
    G3.Refresh
    Me.cmdFetch3.Enabled = True
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFetch2_Click"
    HandleError
End Sub

Private Sub cmdReloadintray_Click()
    On Error GoTo ErrHandler
    oPC.COShort.Execute "EXEC dbo.ExportTRsToPastel_INDEPENDENT"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdReloadintray_Click"
    HandleError
End Sub
Private Sub DPRecFrom_GotFocus()
    On Error GoTo ErrHandler
    optRange.Value = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DPRecFrom_GotFocus"
    HandleError
End Sub

Private Sub DPRecTo_GotFocus()
    On Error GoTo ErrHandler
    optRange.Value = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DPRecTo_GotFocus"
    HandleError
End Sub

Private Sub DPRecFrom4_GotFocus()
    On Error GoTo ErrHandler
    optRange4.Value = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DPRecFrom4_GotFocus"
    HandleError
End Sub


Private Sub DPRecTo4_GotFocus()
    On Error GoTo ErrHandler
    optRange4.Value = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DPRecTo4_GotFocus"
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
Dim i As Integer

    SSTab1.Tab = 0
    For i = 1 To G.Columns.Count
        G.Columns(i - 1).Width = GetSetting(App.EXEName, Me.Name & "A", CStr(i), G.Columns(i - 1).Width)
    Next
    For i = 1 To G2.Columns.Count
        G2.Columns(i - 1).Width = GetSetting(App.EXEName, Me.Name & "B", CStr(i), G2.Columns(i - 1).Width)
    Next
    For i = 1 To G3.Columns.Count
        G3.Columns(i - 1).Width = GetSetting(App.EXEName, Me.Name & "C", CStr(i), G3.Columns(i - 1).Width)
    Next
    For i = 1 To G4.Columns.Count
        G4.Columns(i - 1).Width = GetSetting(App.EXEName, Me.Name & "D", CStr(i), G4.Columns(i - 1).Width)
    Next
    Me.StatusBar1.Panels(1).Text = oPC.PastelConnectionstring
    Me.StatusBar1.Panels(2).Text = oPC.servername
    
    Text2 = oPC.PastelConnectionstring
    DPRecTo4 = Date
    DPRecFrom4 = DateAdd("d", -3, Date)
    DPRecTO = Date
    DPRecFrom = DateAdd("d", -3, Date)
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load"
    HandleError
End Sub

Private Sub G_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
Static Direction As Variant

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType_G(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G_HeadClick(ColIndex)", ColIndex
    HandleError
End Sub
Private Function GetRowType_G(ColIndex As Integer) As Variant
    On Error GoTo ErrHandler
    Select Case ColIndex
        Case 1, 2, 5, 6
            GetRowType_G = XTYPE_STRING
        Case 4
            GetRowType_G = XTYPE_DATE
        Case 3, 7
            GetRowType_G = XTYPE_NUMBER
    End Select
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetRowType_G(ColIndex)", ColIndex
    HandleError
End Function
Private Sub G3_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
Static Direction As Variant
    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XC.QuickSort XC.LowerBound(1), XC.UpperBound(1), ColIndex + 1, Direction, GetRowType_G3(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G3.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G3_HeadClick(ColIndex)", ColIndex
    HandleError
End Sub
Private Function GetRowType_G3(ColIndex As Integer) As Variant
    On Error GoTo ErrHandler
    Select Case ColIndex
        Case 1, 2, 5, 6, 8
            GetRowType_G3 = XTYPE_STRING
        Case 4
            GetRowType_G3 = XTYPE_DATE
        Case 3, 7, 9
            GetRowType_G3 = XTYPE_NUMBER
    End Select
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetRowType_G3(ColIndex)", ColIndex
    HandleError
End Function
Private Sub G4_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
Static Direction As Variant
    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XD.QuickSort XD.LowerBound(1), XD.UpperBound(1), ColIndex + 1, Direction, GetRowType_G4(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G4.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G4_HeadClick(ColIndex)", ColIndex
    HandleError
End Sub
Private Function GetRowType_G4(ColIndex As Integer) As Variant
    On Error GoTo ErrHandler
    Select Case ColIndex
        Case 1, 2, 5, 6
            GetRowType_G4 = XTYPE_STRING
        Case 4
            GetRowType_G4 = XTYPE_DATE
        Case 3, 7
            GetRowType_G4 = XTYPE_NUMBER
    End Select
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetRowType_G4(ColIndex)", ColIndex
    HandleError
End Function
Private Sub G2_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
Static Direction As Variant
    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XB.QuickSort XB.LowerBound(1), XB.UpperBound(1), ColIndex + 1, Direction, GetRowType_G2(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G2.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G2_HeadClick(ColIndex)", ColIndex
    HandleError
End Sub
Private Function GetRowType_G2(ColIndex As Integer) As Variant
    On Error GoTo ErrHandler
    Select Case ColIndex
        Case 4, 5, 6
            GetRowType_G2 = XTYPE_STRING
        Case 1, 2, 3
            GetRowType_G2 = XTYPE_DATE
        Case 2, 7
            GetRowType_G2 = XTYPE_NUMBER
    End Select
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetRowType_G2(ColIndex)", ColIndex
    HandleError
End Function

Private Sub LoadGrid()
    On Error GoTo ErrHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XA = Nothing
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
'    lngArrayRows = rs.RecordCount
'    XA.ReDim 1, lngArrayRows, 1, 10
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
                XA.ReDim 1, lngIndex, 1, 10
                XA.Value(lngIndex, 1) = FNS(rs.Fields("P_SRC"))
                XA.Value(lngIndex, 2) = FNS(rs.Fields("P_JOURNALTYPE"))
                XA.Value(lngIndex, 3) = FNS(rs.Fields("P_DESCR"))
                XA.Value(lngIndex, 4) = FNS(rs.Fields("P_DATE"))
                XA.Value(lngIndex, 5) = FNS(rs.Fields("P_ACNO")) & " - " & FNS(rs.Fields("P_TPName"))
                XA.Value(lngIndex, 6) = FNS(rs.Fields("P_REF"))
                XA.Value(lngIndex, 7) = Format(FNDBL(rs.Fields("P_AMT")), "###,##0.00")
                lngIndex = lngIndex + 1
                rs.MoveNext
        Loop
    End If
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadGrid"
End Sub

Private Sub LoadGrid2()
    On Error GoTo ErrHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XB = Nothing
    Set XB = New XArrayDB
    XB.Clear
    iRecs = i
    lngIndex = 1
    If Not rs2.EOF Then
        rs2.MoveFirst
        Do While Not rs2.EOF
                XB.ReDim 1, lngIndex, 1, 10
                XB.Value(lngIndex, 1) = Format(FNS(rs2.Fields("P_DATE")), "dd-mm-yyyy")
                XB.Value(lngIndex, 2) = FNS(rs2.Fields("P_DATERECEIVEDATHO"))
                XB.Value(lngIndex, 3) = IIf(FND(rs2.Fields("P_POSTEDDATE")) < "2000-01-01", "", FND(rs2.Fields("P_POSTEDDATE")))
                XB.Value(lngIndex, 4) = FNS(rs2.Fields("P_JOURNALTYPE"))
                XB.Value(lngIndex, 5) = FNS(rs2.Fields("P_REF"))
                XB.Value(lngIndex, 6) = FNS(rs2.Fields("P_ACNO")) & " - " & FNS(rs2.Fields("P_TPName"))
                XB.Value(lngIndex, 7) = Format(FNN(rs2.Fields("P_AMT")), "##0.00;(##0.00)")
                XB.Value(lngIndex, 10) = FNS(rs2.Fields("P_GLOBALTRID"))
                lngIndex = lngIndex + 1
                rs2.MoveNext
        Loop
    End If
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadGrid2"
End Sub
Private Sub LoadGrid3()
    On Error GoTo ErrHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XC = Nothing
    Set XC = New XArrayDB
    XC.Clear
    iRecs = i
    lngIndex = 1
'    lngArrayRows = rs3.RecordCount
'    XC.ReDim 1, lngArrayRows, 1, 13
    If Not rs3.EOF Then
        rs3.MoveFirst
        Do While Not rs3.EOF
                XC.ReDim 1, lngIndex, 1, 13
                XC.Value(lngIndex, 1) = FNS(rs3.Fields("P_SRC"))
                XC.Value(lngIndex, 2) = FNS(rs3.Fields("P_JOURNALTYPE"))
                XC.Value(lngIndex, 3) = FNS(rs3.Fields("P_DESCR"))
                XC.Value(lngIndex, 4) = FNS(rs3.Fields("P_DATE"))
                XC.Value(lngIndex, 5) = FNS(rs3.Fields("P_ACNO")) & " - " & FNS(rs3.Fields("P_TPName"))
                XC.Value(lngIndex, 6) = FNS(rs3.Fields("P_REF"))
                XC.Value(lngIndex, 7) = Format(FNDBL(rs3.Fields("P_AMT")), "###,##0.00")
                XC.Value(lngIndex, 8) = FNS(rs3.Fields("P_INVALIDDOC")) '& "/" & IIf(FNN(rs3.Fields("P_UNPOSTED")) = False, "ready for posting", "")
                XC.Value(lngIndex, 9) = IIf(FNN(rs3.Fields("P_ACTION")) = True, 1, 0)
                XC.Value(lngIndex, 13) = FNS(rs3.Fields("ID"))
                lngIndex = lngIndex + 1
                rs3.MoveNext
        Loop
    End If
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadGrid3"
End Sub
Private Sub LoadGrid4()
    On Error GoTo ErrHandler
Dim lngIndex As Long
Dim tmp As String
Dim i As Integer
Dim iRecs As Long
Dim lngArrayRows As Long

    i = 0
    Set XD = Nothing
    Set XD = New XArrayDB
    XD.Clear
    iRecs = i
    lngIndex = 1
'    lngArrayRows = rs4.RecordCount
'    XD.ReDim 1, lngArrayRows, 1, 13
    If Not rs4.EOF Then
        rs4.MoveFirst
        Do While Not rs4.EOF
                XD.ReDim 1, lngIndex, 1, 13
                XD.Value(lngIndex, 1) = FNS(rs4.Fields("P_SRC"))
                XD.Value(lngIndex, 2) = FNS(rs4.Fields("P_JOURNALTYPE"))
                XD.Value(lngIndex, 3) = FNS(rs4.Fields("P_DESCR"))
                XD.Value(lngIndex, 4) = FNS(rs4.Fields("P_DATE"))
                XD.Value(lngIndex, 5) = FNS(rs4.Fields("P_ACNO")) & " - " & FNS(rs4.Fields("P_TPName"))
                XD.Value(lngIndex, 6) = FNS(rs4.Fields("P_REF"))
                XD.Value(lngIndex, 7) = Format(FNDBL(rs4.Fields("P_AMT")), "###,##0.00")
                XD.Value(lngIndex, 8) = ""
                XD.Value(lngIndex, 9) = IIf(FNN(rs4.Fields("P_ACTION")) = True, 1, 0)
                XD.Value(lngIndex, 13) = FNS(rs4.Fields("ID"))
                lngIndex = lngIndex + 1
                rs4.MoveNext
        Loop
    End If
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadGrid4"
End Sub

Private Function TranslateDocType(i As Integer) As String
    On Error GoTo ErrHandler
    Select Case i
        Case 3
            TranslateDocType = "INV"
        Case 8
            TranslateDocType = "CN"
        Case 4
            TranslateDocType = "PUR"
        Case 11
            TranslateDocType = "RET"
    End Select
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TranslateDocType(i)", i
End Function
Private Sub DPREC_GotFocus()
    On Error GoTo ErrHandler
    Me.optDBY = False
    Me.optYesterday = False
    Me.opttoday = False
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DPREC_GotFocus"
    HandleError
End Sub


Private Sub G3_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
    If ColIndex = 8 Then
        oPC.COShort.Execute "UPDATE tIntray SET P_ACTION = " & IIf(G3.Text = -1, CStr(1), CStr(0)) & " WHERE ID = '" & XC(G3.Bookmark, 13) & "'"
    End If
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G3_AfterColUpdate(ColIndex)", ColIndex
    HandleError
End Sub

Private Sub G3_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo ErrHandler
    If ColIndex = 8 Then
        If XC(G3.Bookmark, 8) <> "" Then
            Cancel = 1
        End If
    End If
        
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G3_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, Cancel)
    HandleError
End Sub



Private Sub mnuBranches_Click()
    On Error GoTo ErrHandler
Dim f As New frmBranch

    f.Show vbModal
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBranches_Click"
    HandleError
End Sub

Private Sub mnuColumnsSave_Click()
mnuSaveLayout
End Sub

Private Sub mnuExit_Click()
    On Error GoTo ErrHandler
    Unload Me
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click"
    HandleError
End Sub

Private Sub mnuSQL_Click()
Dim f As New frmSQL
    f.Show
End Sub


Private Sub optAll_2_Click()
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If rs Is Nothing Then Exit Sub
    If optAll_2 Then
        rs3.Filter = ""
        LoadGrid3
        G3.Array = XC
        G3.ReBind
        lblCount = CStr(rs3.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMainNormal.optAll_2_Click"
End Sub

Private Sub optALL_Click()
    If rs Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If optALL Then
        rs.Filter = ""
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.optALL_Click"
    HandleError
End Sub



Private Sub Option5_Click()

End Sub

Private Sub optCN_2_Click()
    On Error GoTo ErrHandler
    If rs3 Is Nothing Then Exit Sub
     Screen.MousePointer = vbHourglass
   If optCN_2 Then
        rs3.Filter = ""
        rs3.Filter = "P_JOURNALTYPE = 'Rtn.from Cust.'"
        rs3.Requery
        LoadGrid3
        G3.Array = XC
        G3.ReBind
        lblCount = CStr(rs3.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMainNormal.optCN_2_Click"
End Sub

Private Sub optINV_2_Click()
    On Error GoTo ErrHandler
    If rs3 Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    If optINV_2 Then
        rs3.Filter = ""
        rs3.Filter = "P_JOURNALTYPE = 'Cust.inv'"
        rs3.Requery
        LoadGrid3
        G3.Array = XC
        G3.ReBind
        lblCount = CStr(rs3.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMainNormal.optINV_2_Click"
End Sub

Private Sub optRET_2_Click()
    On Error GoTo ErrHandler
    If rs3 Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    If optRET_2 Then
        rs3.Filter = ""
        rs3.Filter = "P_JOURNALTYPE = 'Rtn to Supp'"
        rs3.Requery
        LoadGrid3
        G3.Array = XC
        G3.ReBind
        lblCount = CStr(rs3.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMainNormal.optRET_2_Click"
End Sub

Private Sub optRET_Click()
    If rs Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If optRET Then
        rs.Filter = ""
        rs.Filter = "P_JOURNALTYPE = 'Rtn to Supp'"
        rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.optRET_Click"
    HandleError
End Sub

Private Sub optSI_2_Click()
    On Error GoTo ErrHandler
    If rs3 Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    If optSI_2 Then
        rs3.Filter = ""
        rs3.Filter = "P_JOURNALTYPE = 'Supp.inv.'"
        rs3.Requery
        LoadGrid3
        G3.Array = XC
        G3.ReBind
        lblCount = CStr(rs3.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub

    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMainNormal.optSI_2_Click"
End Sub

Private Sub optSI_Click()
    If rs Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If optSI Then
        rs.Filter = ""
        rs.Filter = "P_JOURNALTYPE = 'Supp.inv.'"
        rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.optSI_Click"
    HandleError
End Sub
Private Sub optINV_Click()
    If rs Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If optINV Then
        rs.Filter = "P_JOURNALTYPE = 'Cust.inv'"
        'rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.optINV_Click"
    HandleError
End Sub

Private Sub optCN_Click()
    If rs Is Nothing Then Exit Sub
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    If optCN Then
        rs.Filter = "P_JOURNALTYPE = 'RtnfromCust.'"
       ' rs.Requery
        LoadGrid
        G.Array = XA
        G.ReBind
        lblCount = CStr(rs.RecordCount) & " records"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.optCN_Click"
    HandleError
End Sub
Private Sub G_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
Dim i As Integer
    i = ColIndex + 1
    If i = 7 Then   'checkbox
        oPC.COShort.Execute "UPDATE tTR SET TR_DATETOPASTEL = NULL WHERE TR_ID =  " & XA(G.Bookmark, 10)
        XA(G.Bookmark, 6) = ""
        G.RefetchRow
    End If
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G_ButtonClick(ColIndex)", ColIndex
    HandleError
End Sub


Public Sub mnuSaveLayout()
    On Error GoTo ErrHandler
    SaveLayout Me.G, Me.Name & "A"
    SaveLayout Me.G2, Me.Name & "B"
    SaveLayout Me.G3, Me.Name & "C"
    SaveLayout Me.G4, Me.Name & "D"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSaveLayout"
End Sub



Private Sub txtStore_DblClick()
    On Error GoTo ErrHandler
   ' txtStore = ""
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtStore_DblClick"
    HandleError
End Sub

