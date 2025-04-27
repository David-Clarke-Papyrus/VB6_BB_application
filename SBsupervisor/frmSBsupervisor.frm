VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSBsupervisor 
   Caption         =   "Service broker supervisor"
   ClientHeight    =   12960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   22530
   LinkTopic       =   "Form1"
   ScaleHeight     =   12960
   ScaleWidth      =   22530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword_T 
      Height          =   330
      Left            =   12975
      TabIndex        =   51
      Top             =   330
      Width           =   480
   End
   Begin VB.TextBox txtPassword_I 
      Height          =   330
      Left            =   3435
      TabIndex        =   50
      Top             =   330
      Width           =   480
   End
   Begin VB.ComboBox cboStores 
      Height          =   315
      Left            =   7350
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   330
      Width           =   2490
   End
   Begin VB.PictureBox TargetPicD 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   14520
      Picture         =   "frmSBsupervisor.frx":0000
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   48
      Top             =   345
      Width           =   285
   End
   Begin VB.PictureBox InitPicD 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4905
      Picture         =   "frmSBsupervisor.frx":038A
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   47
      Top             =   360
      Width           =   285
   End
   Begin VB.PictureBox TargetPic 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   14505
      Picture         =   "frmSBsupervisor.frx":0714
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   46
      Top             =   360
      Width           =   285
   End
   Begin VB.PictureBox INITPic 
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   4905
      Picture         =   "frmSBsupervisor.frx":0A9E
      ScaleHeight     =   270
      ScaleWidth      =   285
      TabIndex        =   45
      Top             =   390
      Width           =   285
   End
   Begin VB.CheckBox chkBranchMainOn 
      Caption         =   "Debug on"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   5280
      TabIndex        =   44
      Top             =   390
      Width           =   1515
   End
   Begin VB.CheckBox chkBranchDebugOn 
      Caption         =   "Debug on"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   14910
      TabIndex        =   43
      Top             =   360
      Width           =   1515
   End
   Begin VB.TextBox txt_T_Instance 
      Height          =   330
      Left            =   9885
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   330
      Width           =   3045
   End
   Begin VB.CommandButton cmd_T_Connect 
      Caption         =   "Connect"
      Height          =   345
      Left            =   13500
      TabIndex        =   16
      Top             =   330
      Width           =   945
   End
   Begin VB.CheckBox chk_I_Central 
      Caption         =   "Central"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2910
      TabIndex        =   14
      Top             =   60
      Width           =   795
   End
   Begin VB.CheckBox chk_I_HO 
      Caption         =   "HO"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2310
      TabIndex        =   13
      Top             =   60
      Width           =   1545
   End
   Begin VB.CommandButton cmd_I_Connect 
      Caption         =   "Connect"
      Height          =   345
      Left            =   4035
      TabIndex        =   12
      Top             =   345
      Width           =   780
   End
   Begin VB.TextBox txt_I_Instance 
      Height          =   330
      Left            =   180
      TabIndex        =   11
      Top             =   330
      Width           =   3150
   End
   Begin VB.CheckBox chk_I_HUB 
      Caption         =   "HUB"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   3855
      TabIndex        =   10
      Top             =   60
      Width           =   645
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   12030
      Left            =   150
      TabIndex        =   0
      Top             =   735
      Width           =   22080
      _ExtentX        =   38947
      _ExtentY        =   21220
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   970
      BackColor       =   -2147483638
      ForeColor       =   -2147483646
      TabCaption(0)   =   "Logs and Errors"
      TabPicture(0)   =   "frmSBsupervisor.frx":0E28
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DC_T_SQLSVR"
      Tab(0).Control(1)=   "DCL"
      Tab(0).Control(2)=   "DC_I_SQLSVR"
      Tab(0).Control(3)=   "cmdDisconnect"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Testing round-trip question and response"
      TabPicture(1)   =   "frmSBsupervisor.frx":0E44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frInitiator"
      Tab(1).Control(1)=   "frTarget"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Transmission queue"
      TabPicture(2)   =   "frmSBsupervisor.frx":0E60
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Settings on databases"
      TabPicture(3)   =   "frmSBsupervisor.frx":0E7C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "frTab4_I"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Branch"
         Height          =   2895
         Left            =   -74805
         TabIndex        =   54
         Top             =   4620
         Width           =   19500
         Begin VB.CommandButton cmdEnableServiceBRoker_T 
            Caption         =   "Enable service broker"
            Height          =   510
            Left            =   645
            TabIndex        =   56
            Top             =   1065
            Width           =   2205
         End
         Begin VB.CommandButton cmdEnableAdhoc_T 
            Caption         =   "Enable ad-hoc queries,xp_CmdShell and OLE Automation"
            Height          =   510
            Left            =   645
            TabIndex        =   55
            Top             =   405
            Width           =   2865
         End
         Begin VB.Label Label 
            Caption         =   $"frmSBsupervisor.frx":0E98
            Height          =   645
            Left            =   2940
            TabIndex        =   57
            Top             =   1035
            Width           =   4470
         End
      End
      Begin VB.Frame frTab4_I 
         Caption         =   "Non-branch"
         Height          =   2895
         Left            =   -74805
         TabIndex        =   52
         Top             =   780
         Width           =   19500
         Begin VB.CommandButton cmdEnableAdhoc_I 
            Caption         =   "Enable ad-hoc queries"
            Height          =   510
            Left            =   645
            TabIndex        =   53
            Top             =   405
            Width           =   2205
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "branch"
         Height          =   6015
         Left            =   105
         TabIndex        =   28
         Top             =   5865
         Width           =   21675
         Begin VB.CommandButton cmdDeleteTargetSessionConversations 
            Caption         =   "Clear all rows in _SessionConversations (necessary if this store's IP address changes) "
            Height          =   810
            Left            =   11505
            TabIndex        =   59
            Top             =   255
            Width           =   2730
         End
         Begin VB.CommandButton cmdCLear 
            Caption         =   "Clear all conversations in sys.transmissionqueue"
            Height          =   525
            Left            =   14610
            TabIndex        =   58
            Top             =   240
            Width           =   2730
         End
         Begin VB.CommandButton cmdDeleteSell2 
            Caption         =   "Delete selected with cleanup"
            Height          =   345
            Left            =   14610
            TabIndex        =   42
            Top             =   1350
            Width           =   2730
         End
         Begin VB.CommandButton cmdStartRQ 
            BackColor       =   &H00C4BCA4&
            Height          =   330
            Left            =   7440
            Picture         =   "frmSBsupervisor.frx":0F37
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   705
            Width           =   360
         End
         Begin VB.CommandButton cmdStopRQ 
            BackColor       =   &H00C4BCA4&
            Height          =   330
            Left            =   6990
            Picture         =   "frmSBsupervisor.frx":12C1
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   720
            Width           =   360
         End
         Begin VB.CommandButton cmdLoadRemoteQueues 
            Height          =   330
            Left            =   480
            Picture         =   "frmSBsupervisor.frx":164B
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   255
            Width           =   645
         End
         Begin VB.CommandButton cmd_Target_TxQ 
            Height          =   330
            Left            =   20160
            Picture         =   "frmSBsupervisor.frx":19D5
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   1365
            Width           =   1230
         End
         Begin TrueOleDBGrid60.TDBGrid GTarget_TxQ 
            Bindings        =   "frmSBsupervisor.frx":1D5F
            Height          =   4080
            Left            =   60
            OleObjectBlob   =   "frmSBsupervisor.frx":1D71
            TabIndex        =   30
            Top             =   1725
            Width           =   21375
         End
         Begin MSAdodcLib.Adodc DC_T_Txq 
            Height          =   330
            Left            =   20040
            Top             =   255
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc DC_R_Q 
            Height          =   330
            Left            =   0
            Top             =   465
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin TrueOleDBGrid60.TDBGrid GRQ 
            Bindings        =   "frmSBsupervisor.frx":6A1D
            Height          =   1530
            Left            =   1155
            OleObjectBlob   =   "frmSBsupervisor.frx":6A2F
            TabIndex        =   34
            Top             =   135
            Width           =   5595
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "non-branch"
         Height          =   5220
         Left            =   90
         TabIndex        =   25
         Top             =   615
         Width           =   21675
         Begin VB.CommandButton cmdDeleteAllINITSessionConversations 
            Caption         =   "Clear all rows in _SessionConversations (necessary if this store's IP address changes)"
            Height          =   810
            Left            =   11400
            TabIndex        =   61
            Top             =   285
            Width           =   2730
         End
         Begin VB.CommandButton cmdClearINIT 
            Caption         =   "Clear all conversations in sys.transmissionqueue"
            Height          =   525
            Left            =   14505
            TabIndex        =   60
            Top             =   285
            Width           =   2730
         End
         Begin VB.CommandButton cmdDeleteSel1 
            Caption         =   "Delete selected with cleanup"
            Height          =   345
            Left            =   14550
            TabIndex        =   41
            Top             =   1335
            Width           =   2730
         End
         Begin VB.CommandButton cmdLoadLocalQueues 
            Height          =   330
            Left            =   540
            Picture         =   "frmSBsupervisor.frx":A233
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   255
            Width           =   645
         End
         Begin VB.CommandButton cmdStopLQ 
            BackColor       =   &H00C4BCA4&
            Height          =   330
            Left            =   7065
            Picture         =   "frmSBsupervisor.frx":A5BD
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   720
            Width           =   360
         End
         Begin VB.CommandButton cmdStartLQ 
            BackColor       =   &H00C4BCA4&
            Height          =   330
            Left            =   7515
            Picture         =   "frmSBsupervisor.frx":A947
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   720
            Width           =   360
         End
         Begin VB.CommandButton cmd_Init_TxQ 
            Height          =   330
            Left            =   20235
            Picture         =   "frmSBsupervisor.frx":ACD1
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   1335
            Width           =   1230
         End
         Begin TrueOleDBGrid60.TDBGrid GInit_TxQ 
            Bindings        =   "frmSBsupervisor.frx":B05B
            Height          =   3390
            Left            =   90
            OleObjectBlob   =   "frmSBsupervisor.frx":B06D
            TabIndex        =   27
            Top             =   1710
            Width           =   21435
         End
         Begin MSAdodcLib.Adodc DC_I_Txq 
            Height          =   330
            Left            =   19905
            Top             =   690
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSAdodcLib.Adodc DC_L_Q 
            Height          =   330
            Left            =   60
            Top             =   465
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin TrueOleDBGrid60.TDBGrid GLQ 
            Bindings        =   "frmSBsupervisor.frx":FD17
            Height          =   1530
            Left            =   1215
            OleObjectBlob   =   "frmSBsupervisor.frx":FD29
            TabIndex        =   40
            Top             =   135
            Width           =   5595
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Branch"
         Height          =   5715
         Left            =   -74865
         TabIndex        =   22
         Top             =   5685
         Width           =   21555
         Begin VB.CommandButton cmdRestartB 
            Caption         =   "Re-start error log"
            Height          =   300
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   225
            Width           =   1320
         End
         Begin VB.CommandButton cmdRefreshTarget_SQLSVR 
            Height          =   300
            Left            =   135
            Picture         =   "frmSBsupervisor.frx":1352D
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   255
            Width           =   1200
         End
         Begin TrueOleDBGrid60.TDBGrid GTarget_SQLSVR 
            Bindings        =   "frmSBsupervisor.frx":138B7
            Height          =   4875
            Left            =   60
            OleObjectBlob   =   "frmSBsupervisor.frx":138C9
            TabIndex        =   24
            Top             =   660
            Width           =   21405
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Non branch"
         Height          =   5010
         Left            =   -74910
         TabIndex        =   19
         Top             =   630
         Width           =   21585
         Begin VB.CommandButton cmdRestartNB 
            Caption         =   "Re-start error log"
            Height          =   300
            Left            =   10080
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   1320
         End
         Begin VB.CommandButton cmdRefreshINIT_SQLSVR 
            Height          =   300
            Left            =   105
            Picture         =   "frmSBsupervisor.frx":16970
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   225
            Width           =   1320
         End
         Begin TrueOleDBGrid60.TDBGrid GINIT_SQLSVR 
            Bindings        =   "frmSBsupervisor.frx":16CFA
            Height          =   4350
            Left            =   45
            OleObjectBlob   =   "frmSBsupervisor.frx":16D0C
            TabIndex        =   21
            Top             =   570
            Width           =   21285
         End
      End
      Begin VB.Frame frTarget 
         Caption         =   "Branch"
         Height          =   5790
         Left            =   -74895
         TabIndex        =   3
         Top             =   6105
         Width           =   21795
         Begin VB.CommandButton cmdRefreshTarget_SB 
            Height          =   300
            Left            =   120
            Picture         =   "frmSBsupervisor.frx":19B51
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   255
            Width           =   810
         End
         Begin VB.CommandButton cmdClearTargetLog 
            Caption         =   "Clear _tSBLog"
            Height          =   300
            Left            =   10095
            TabIndex        =   8
            Top             =   285
            Width           =   1140
         End
         Begin MSAdodcLib.Adodc DC_T_SB 
            Height          =   330
            Left            =   3660
            Top             =   195
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin TrueOleDBGrid60.TDBGrid GTarget_SB 
            Bindings        =   "frmSBsupervisor.frx":19EDB
            Height          =   5055
            Left            =   120
            OleObjectBlob   =   "frmSBsupervisor.frx":19EED
            TabIndex        =   7
            Top             =   630
            Width           =   21480
         End
      End
      Begin VB.Frame frInitiator 
         Caption         =   "Non branch"
         Height          =   5490
         Left            =   -74910
         TabIndex        =   2
         Top             =   630
         Width           =   21780
         Begin VB.CommandButton cmdRefreshINIT_SB 
            Height          =   300
            Left            =   120
            Picture         =   "frmSBsupervisor.frx":1D320
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   255
            Width           =   1110
         End
         Begin VB.CommandButton cmdClearLog 
            Caption         =   "Clear _tSBLog"
            Height          =   300
            Left            =   9945
            TabIndex        =   4
            Top             =   255
            Width           =   1260
         End
         Begin TrueOleDBGrid60.TDBGrid GINIT_SB 
            Bindings        =   "frmSBsupervisor.frx":1D6AA
            Height          =   4785
            Left            =   120
            OleObjectBlob   =   "frmSBsupervisor.frx":1D6BC
            TabIndex        =   6
            Top             =   600
            Width           =   21525
         End
         Begin MSAdodcLib.Adodc DC_I_SB 
            Height          =   330
            Left            =   3690
            Top             =   195
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
            CursorOptions   =   0
            CacheSize       =   50
            MaxRecords      =   0
            BOFAction       =   0
            EOFAction       =   0
            ConnectStringType=   1
            Appearance      =   1
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   "Adodc1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         Height          =   345
         Left            =   -54375
         TabIndex        =   1
         Top             =   11490
         Width           =   1080
      End
      Begin MSAdodcLib.Adodc DC_I_SQLSVR 
         Height          =   405
         Left            =   -68970
         Top             =   5355
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc DCL 
         Height          =   405
         Left            =   -74835
         Top             =   11520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc DC_T_SQLSVR 
         Height          =   405
         Left            =   -67275
         Top             =   5325
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Connect to branch"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   9840
      TabIndex        =   18
      Top             =   60
      Width           =   1470
   End
   Begin VB.Label Label2 
      Caption         =   "Connect to non branch"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   180
      TabIndex        =   15
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmSBsupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnINIT As ADODB.Connection
Dim cnINITM As ADODB.Connection
Dim cnTarget As ADODB.Connection
Dim cnTargetM As ADODB.Connection
Dim strMainConnectionString As String
Dim rs As New ADODB.Recordset
Dim rsl As New ADODB.Recordset
Dim strCommandFilePath As String
Dim fs As New FileSystemObject
Dim oTF As z_TextFileSimple
Dim db As String
Dim dbbr As String
Dim InitiatorType As String
Dim TargetType As String
Dim rsRQ As ADODB.Recordset
Dim INITConnected As Boolean
Dim TargetConnected As Boolean
Dim rsStores As New ADODB.Recordset
Dim flgLoading As Boolean

Private Sub cboStores_Change()
    rsStores.Find "STORE_NAME = " & cboStores
    txt_T_Instance = FNS(rsStores.Fields(1))
End Sub

Private Sub cboStores_Click()
    If rsStores Is Nothing Then Exit Sub
    If rsStores.State = 0 Then Exit Sub
    If rsStores.EOF And rsStores.BOF Then Exit Sub
    If flgLoading Then Exit Sub
    rsStores.MoveFirst
    rsStores.Find "STORE_NAME = '" & cboStores & "'"
    txt_T_Instance = FNS(rsStores.Fields(1)) & "\PBKSINSTANCE2"
End Sub

Private Sub chkBranchDebugOn_Click()
    SetDebugStatus cnTarget, IIf(chkBranchDebugOn, True, False)

End Sub

Private Sub chkBranchMainOn_Click()
    
    SetDebugStatus cnINIT, IIf(Me.chkBranchMainOn, True, False)
End Sub

Private Sub cmd_I_Connect_Click()
    On Error GoTo ErrHandler
Dim x
    If chk_I_HO = 1 Then
        db = "PBKSHO"
    ElseIf chk_I_Central = 1 Then
        db = "PBKSC"
    Else
        db = "HUB"
    End If
    strMainConnectionString = "Provider=SQLNCLI;Persist Security Info=False;Data Source=" & txt_I_Instance & ";Initial Catalog=" & db & ";User Id=sa;Password=" & txtPassword_I & ";Connect Timeout=10"
    Set cnINIT = New ADODB.Connection
    cnINIT.Open strMainConnectionString
    cnINIT.CommandTimeout = 240
    INITConnected = True
    SetConnectionIcons "INIT", True
    
    strMainConnectionString = "Provider=SQLNCLI;Persist Security Info=False;Data Source=" & txt_I_Instance & ";Initial Catalog=" & "master" & ";User Id=sa;Password=" & txtPassword_I & ";Connect Timeout=10"
    Set cnINITM = New ADODB.Connection
    cnINITM.Open strMainConnectionString
    cnINITM.CommandTimeout = 240
    
    
    SaveSetting "PBKS", "Supervisor", "InitiatorConnectionString", txt_I_Instance
    If Me.chk_I_HUB = 1 Then
        InitiatorType = "H"
    Else
        If Me.chk_I_HO = 1 Then
            InitiatorType = "A"
        Else
            If Me.chk_I_Central = 1 Then
                InitiatorType = "C"
            End If
        End If
    End If
    SaveSetting "PBKS", "Supervisor", "InitiatorType", InitiatorType
    
    chkBranchMainOn = IIf(GetDebugStatus(cnINIT), 1, 0)
    LoadBranchStoreCombo
    
    Exit Sub
ErrHandler:
    If Left(Error, 12) = "Login failed" Then
        MsgBox "Cannot connect to server", vbInformation, "Can't connect"
        Err.Clear
        Exit Sub
    End If
    ErrPreserve
    If InStr(1, Error, "Error Locating Server") > 0 Then
        MsgBox "Cannot connect to server", vbInformation, "Can't connect"
        Err.Clear
        Exit Sub
    End If
    ErrPreserve
    If Error = "Timeout expired" Or InStr(1, Error, "SQL Network Interfaces: Error Locating Server/Instance Specified") > 0 Then
        MsgBox "The attempt to connect has timed out, possibly your instance name is wrong.", vbInformation + vbOKOnly, "Can't connect to " & txt_I_Instance
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.cmd_I_Connect_Click", , EA_NORERAISE
End Sub

Private Sub cmd_Init_TxQ_Click()
    Screen.MousePointer = vbHourglass
    LoadTxQ cnINIT, GInit_TxQ, DC_I_Txq
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_T_Connect_Click()
    On Error GoTo ErrHandler
    dbbr = "PBKS"

    strMainConnectionString = "Provider=SQLNCLI;Persist Security Info=False;Data Source=" & txt_T_Instance & ";Initial Catalog=" & dbbr & ";User Id=sa;Password=" & txtPassword_T & ";Connect Timeout=10"
    Set cnTarget = New ADODB.Connection
    cnTarget.CommandTimeout = 240
    cnTarget.Open strMainConnectionString
    TargetConnected = True
    SetConnectionIcons "TARGET", True

    SaveSetting "PBKS", "Supervisor", "TargetConnectionString", txt_T_Instance
    SaveSetting "PBKS", "Supervisor", "TargetType", ""
    chkBranchDebugOn.Value = IIf(GetDebugStatus(cnTarget), 1, 0)
    Exit Sub
ErrHandler:
    ErrPreserve
    If Left(Error, 12) = "Login failed" Then
        MsgBox "Cannot connect to server" & vbCrLf & strMainConnectionString, vbInformation, "Can't connect"
        Err.Clear
        Exit Sub
    End If
    If InStr(1, Error, "Error Locating Server") > 0 Then
        MsgBox "Cannot connect to server" & vbCrLf & strMainConnectionString, vbInformation, "Can't connect"
        Err.Clear
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.cmd_T_Connect_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdCLear_Click()
Dim s As String
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter


    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnTarget
    cmd.CommandText = "_Clearqueue"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    On Error Resume Next
    cmd.Execute
    Set cmd = Nothing

End Sub

Private Sub cmdClearLog_Click()
    cnINIT.Execute "TRUNCATE TABLE _tSBLog"
End Sub

Private Sub cmdClearTargetLog_Click()
    cnTarget.Execute "TRUNCATE TABLE _tSBLog"

End Sub



Private Sub DeleteConversation(pHandle As String, Cnn As ADODB.Connection)
Dim s As String
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter

    s = "END CONVERSATION '" & pHandle & "' WITH CLEANUP"
    If MsgBox("Delete the conversation: " & s & "?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = Cnn
    cmd.CommandText = "_EndConversationWithCleanup"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    Set par = cmd.CreateParameter("@HANDLE", adGUID, adParamInput, , pHandle)
    cmd.Parameters.Append par
    On Error Resume Next
    cmd.Execute
    Set cmd = Nothing
    
End Sub
Private Sub Delete_SessionConversations(Cnn As ADODB.Connection)
Dim s As String
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter

    If MsgBox("Delete all rows in _SessionConversations (innocuous action)? ", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    On Error Resume Next
    Cnn.Execute "TRUNCATE TABLE _SessionConversations"
    
End Sub


Private Sub cmdDeleteAllINITSessionConversations_Click()
    Delete_SessionConversations cnINIT

End Sub

Private Sub cmdDeleteSel1_Click()
Dim i As Integer
On Error Resume Next
  '  DeleteConversation GInit_TxQ.Columns(7), cnINIT
    For i = GInit_TxQ.SelBookmarks.Count To 1 Step -1
        GInit_TxQ.Bookmark = GInit_TxQ.SelBookmarks(i - 1)
        DeleteConversation GInit_TxQ.Columns(7), cnINIT
    Next
End Sub

Private Sub cmdDeleteSell2_Click()
Dim i As Integer

    For i = GTarget_TxQ.SelBookmarks.Count To 1 Step -1
        GTarget_TxQ.Bookmark = GTarget_TxQ.SelBookmarks(i - 1)
        DeleteConversation GTarget_TxQ.Columns(7), cnTarget
    Next

End Sub

Private Sub cmdEnableAdhoc_T_Click()
    EnableAdhocQueries True
    EnableOLEAutomation True
    EnablexpCmdShell True
End Sub

Private Sub cmdEnableServiceBRoker_T_Click()
    EnableServiceBroker True

End Sub

Private Sub cmdLoadLocalQueues_Click()
    LoadQueues cnINIT, GLQ, DC_L_Q

End Sub
Private Sub cmdRefreshINIT_SQLSVR_Click()
    Screen.MousePointer = vbHourglass
    LoadSQLSVRLog False, cnINIT, GINIT_SQLSVR, DC_I_SQLSVR
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdRefreshTarget_SQLSVR_Click()
    Screen.MousePointer = vbHourglass
    LoadSQLSVRLog True, cnTarget, GTarget_SQLSVR, DC_T_SQLSVR
    Screen.MousePointer = vbDefault
End Sub
Private Sub cmdRefreshINIT_SB_Click()
    
    Screen.MousePointer = vbHourglass
    LoadSBMsg cnINIT, GINIT_SB, DC_I_SB
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRefreshSQLLog_Click()

End Sub

Private Sub cmdRefreshTarget_SB_Click()
    Screen.MousePointer = vbHourglass
    LoadSBMsg cnTarget, GTarget_SB, DC_T_SB
    Screen.MousePointer = vbDefault
End Sub
Private Sub LoadSBMsg(pcn As ADODB.Connection, pG As TDBGrid, pDC As Adodc)
    On Error GoTo ErrHandler

    If pcn Is Nothing Then
        MsgBox "You must connect first"
        Exit Sub
    End If

    If pcn.State = 0 Then
        MsgBox "You must connect first"
        Exit Sub
    End If
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    If pcn.State = 0 Then
        MsgBox "You must first open the connection"
        Exit Sub
    End If
    rs.Open "SELECT SBL_DATE,SBL_MSG,SBL_PROC,Left(CAST(SBL_XMLDATA as VARCHAR(MAX)),500) as XMLdata FROM _tSBLOG WHERE SBL_DATE > DATEADD(d,-10,Getdate()) ORDER BY SBL_SEQ DESC", pcn, adOpenKeyset, adLockOptimistic
    Set pDC.Recordset = rs
    pG.DataSource = pDC

    pG.Refresh
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.LoadSBMsg(pcn,pG,pDC)", Array(pcn, pG, pDC)
End Sub
Private Sub cmd_Target_TxQ_Click()
    Screen.MousePointer = vbHourglass
    LoadTxQ cnTarget, GTarget_TxQ, DC_T_Txq
    Screen.MousePointer = vbDefault

End Sub
Private Sub LoadTxQ(pcn As ADODB.Connection, pG As TDBGrid, pDC As Adodc)
    On Error GoTo ErrHandler
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    
    rs.Open "SELECT conversation_handle,enqueue_time,from_service_name, to_service_name,message_type_name,is_conversation_error,is_end_of_dialog,transmission_status From sys.transmission_queue order by enqueue_time DESC", pcn, adOpenKeyset, adLockOptimistic
    Set pDC.Recordset = rs
    pG.DataSource = pDC
    pG.ReBind
    pG.Refresh
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.LoadTxQ(pcn,pG,pDC)", Array(pcn, pG, pDC)
End Sub

Private Sub cmdLoadRemoteQueues_Click()
    
    LoadQueues cnTarget, GRQ, DC_R_Q
    
End Sub
Private Sub LoadQueues(pcn As ADODB.Connection, pG As TDBGrid, pDC As Adodc)
Dim i As Integer
    On Error GoTo ErrHandler
    i = 0
retry:
    Set rsRQ = Nothing
    Set rsRQ = New ADODB.Recordset
    rsRQ.Open "SELECT Name,is_activation_enabled,is_receive_enabled,is_enqueue_enabled From sys.service_queues WHERE is_ms_shipped = 0 ORDER BY NAME", pcn, adOpenKeyset, adLockOptimistic
    Set pDC.Recordset = rsRQ
    pG.DataSource = pDC
    pG.ReBind
    pG.Refresh

    Exit Sub
ErrHandler:
    If i = 0 Then
        i = i + 1
        cmd_T_Connect_Click
        GoTo retry
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.LoadQueues(pcn,pG,pDC)", Array(pcn, pG, pDC)
End Sub

Private Sub cmdSOH_Click()

End Sub


Private Sub cmdRestartB_Click()
    RecycleErrorLog True  ', cnINIT, GINIT_SQLSVR, DC_I_SQLSVR

End Sub

Private Sub cmdStartLQ_Click()
    startQ GLQ.Text, cnINIT
End Sub

Private Sub cmdStartRQ_Click()
    startQ GRQ.Text, cnTarget
End Sub

Private Sub cmdStopLQ_Click()
    stopQ GLQ.Text, cnINIT
End Sub

Private Sub cmdStopRQ_Click()
    stopQ GRQ.Text, cnTarget

End Sub
Private Sub stopQ(s As String, Cnn As ADODB.Connection)
Dim cmd As New ADODB.Command
Dim res As Recordset
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = Cnn
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = OFF;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing

End Sub
Private Sub startQ(s As String, Cnn As ADODB.Connection)
Dim cmd As New ADODB.Command
Dim res As Recordset
    cmd.ActiveConnection = Cnn
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = ON;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing

End Sub

Private Sub cmdEnableAdhoc_I_Click()
    EnableAdhocQueries False
End Sub

Private Sub cmdDeleteTargetSessionConversations_Click()
    Delete_SessionConversations cnTarget
End Sub

Private Sub cmdClearINIT_Click()
Dim s As String
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter


    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cnINIT
    cmd.CommandText = "_Clearqueue"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    On Error Resume Next
    cmd.Execute
    Set cmd = Nothing

End Sub

Private Sub Form_Load()
Dim T As String

    flgLoading = True
   strSharedServerFolder = "\\PBKS_S"
   Me.txt_I_Instance = GetSetting("PBKS", "Supervisor", "InitiatorConnectionString", "")
   InitiatorType = GetSetting("PBKS", "Supervisor", "InitiatorType", "")

   Me.txt_T_Instance = GetSetting("PBKS", "Supervisor", "TargetConnectionString", "")
   TargetType = GetSetting("PBKS", "Supervisor", "TargetType", "")
   If InitiatorType = "H" Then
        Me.chk_I_HUB = 1
    Else
        If InitiatorType = "A" Then
            Me.chk_I_HO = 1
        Else
            If InitiatorType = "C" Then
                chk_I_Central = 1
            End If
        End If
    End If
    INITConnected = False
    TargetConnected = False
    SetConnectionIcons "INIT", False
    SetConnectionIcons "TARGET", False
     flgLoading = False
   
End Sub


Private Sub ExecuteScript(isBranch As Boolean)
Dim strCommand As String
Dim res As Boolean
Dim svrName As String

    If isBranch Then
        db = "PBKS"
        svrName = txt_T_Instance
        strCommand = "SQLCMD -Usa -P" & txtPassword_T & " -S" & svrName & " -d" & db & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    Else
        db = "PBKSC"
        svrName = txt_I_Instance
        strCommand = "SQLCMD -Usa -P" & txtPassword_I & " -S" & svrName & " -d" & db & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    End If
    If fs.FileExists(strCommandFilePath) Then
        res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 40000, True)
    End If
    
    
End Sub
Private Sub cmdRestartNB_Click()
    RecycleErrorLog False  ', cnINIT, GINIT_SQLSVR, DC_I_SQLSVR

End Sub

Private Sub RecycleErrorLog(isBranch As Boolean)
Dim fs As New FileSystemObject
    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        
            oTF.WriteToTextFile "USE [Master]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "EXEC sp_cycle_errorlog ;"
            oTF.WriteToTextFile "GO"
           
        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
    
    If fs.FileExists(strCommandFilePath) Then
        ExecuteScript isBranch
    Else
        MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
    End If
End Sub
Private Sub EnableAdhocQueries(isBranch As Boolean)
Dim fs As New FileSystemObject
    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
            If isBranch Then
                oTF.WriteToTextFile "USE PBKS"
            Else
                oTF.WriteToTextFile "USE PBKSC"
            End If
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "sp_configure 'show advanced options', 1"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "reconfigure"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "sp_configure 'Ad Hoc Distributed Queries', 1"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "reconfigure"
            oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
    
    If fs.FileExists(strCommandFilePath) Then
        ExecuteScript isBranch
    Else
        MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
    End If
End Sub
Private Sub EnableOLEAutomation(isBranch As Boolean)
Dim fs As New FileSystemObject
    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
            If isBranch Then
                oTF.WriteToTextFile "USE PBKS"
            Else
                oTF.WriteToTextFile "USE PBKSC"
            End If
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "sp_configure 'show advanced options', 1"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "reconfigure"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "sp_configure 'Ole Automation Procedures', 1"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "reconfigure"
            oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
    
    If fs.FileExists(strCommandFilePath) Then
        ExecuteScript isBranch
    Else
        MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
    End If
End Sub
Private Sub EnablexpCmdShell(isBranch As Boolean)
Dim fs As New FileSystemObject
    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
            If isBranch Then
                oTF.WriteToTextFile "USE PBKS"
            Else
                oTF.WriteToTextFile "USE PBKSC"
            End If
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "sp_configure 'show advanced options', 1"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "reconfigure"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "sp_configure 'xp_cmdshell', 1"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "reconfigure"
            oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
    
    If fs.FileExists(strCommandFilePath) Then
        ExecuteScript isBranch
    Else
        MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
    End If
End Sub

Private Sub EnableServiceBroker(isBranch As Boolean)
Dim fs As New FileSystemObject
    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
            If isBranch Then
                oTF.WriteToTextFile "USE PBKS"
            Else
                oTF.WriteToTextFile "USE PBKSC"
            End If
            oTF.WriteToTextFile "use master"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "ALTER DATABASE ServiceBrokerTest"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "SET ENABLE_BROKER;"
            oTF.WriteToTextFile "GO"
        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If
    
    If fs.FileExists(strCommandFilePath) Then
        ExecuteScript isBranch
    Else
        MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
    End If

End Sub
Private Sub LoadSQLSVRLog(isBranch As Boolean, pcn As ADODB.Connection, pG As TDBGrid, pDC As Adodc)
    
    ReadErrorLog isBranch
    
    Set rsl = Nothing
    Set rsl = New ADODB.Recordset
    
    rsl.Open "SELECT LogDate,ProcessInfo,vchMessage FROM tERRLOG  ORDER by LogDate DESC", pcn, adOpenKeyset, adLockOptimistic
    Set pDC.Recordset = rsl
    pG.DataSource = pDC
    pG.ReBind
    pG.Refresh

End Sub
Private Sub ReadErrorLog(isBranch As Boolean)
    
        PrepareScript
        If fs.FileExists(strCommandFilePath) Then
            ExecuteScript isBranch
        Else
            MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
        End If

End Sub
Private Sub PrepareScript()
    strCommandFilePath = "c:\PBKS\PrepareServiceBrokerScript.SQL"
    If fs.FolderExists(fs.GetParentFolderName(strCommandFilePath)) Then
        Set oTF = New z_TextFileSimple
        oTF.OpenTextFile strCommandFilePath
        
        oTF.WriteToTextFile "TRUNCATE TABLE tERRLOG"
        oTF.WriteToTextFile "GO"
        oTF.WriteToTextFile "INSERT INTO tERRLOG EXEC master.dbo.xp_readerrorlog"
        oTF.WriteToTextFile "GO"
        
    
        oTF.CloseTextFile
        Set oTF = Nothing
    Else
        MsgBox strCommandFilePath & " does not exist"
    End If

End Sub

Private Function GetDebugStatus(cn As ADODB.Connection) As Boolean
    On Error GoTo ErrHandler
Dim rs As New ADODB.Recordset
    rs.Open "SELECT CF_DEBUG FROM tConfiguration", cn, adOpenStatic
    GetDebugStatus = IIf(rs.Fields(0) = "TRUE", True, False)
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.GetDebugStatus(cn)", cn
End Function
Private Sub SetDebugStatus(cn As ADODB.Connection, val As Boolean)
    On Error GoTo ErrHandler
Dim rs As New ADODB.Recordset
    If cn Is Nothing Then Exit Sub
    If cn.State = 0 Then Exit Sub
    On Error Resume Next
    rs.Open "SELECT CF_DEBUG FROM tConfiguration", cn, , adLockOptimistic
    If val = True Then
        rs.Fields(0) = "TRUE"
    Else
        rs.Fields(0) = "FALSE"
    End If
    rs.Update
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSBsupervisor.SetDebugStatus(cn,val)", Array(cn, val)
End Sub
Private Sub SetConnectionIcons(Init_Target As String, bConnected As Boolean)
    If UCase(Init_Target) = "INIT" Then
        INITPic.Visible = bConnected
        InitPicD.Visible = Not bConnected
        cmdRefreshINIT_SQLSVR.Enabled = bConnected
        cmdRestartNB.Enabled = bConnected
        cmdRefreshINIT_SB.Enabled = bConnected
        cmdClearLog.Enabled = bConnected
        cmdLoadLocalQueues.Enabled = bConnected
        cmdStopLQ.Enabled = bConnected
        cmdStartLQ.Enabled = bConnected
        cmd_Init_TxQ.Enabled = bConnected
        cmdDeleteSel1.Enabled = bConnected
        chkBranchMainOn.Enabled = bConnected
    ElseIf UCase(Init_Target) = "TARGET" Then
        TargetPic.Visible = bConnected
        TargetPicD.Visible = Not bConnected
        cmdRefreshTarget_SQLSVR.Enabled = bConnected
        cmdRestartB.Enabled = bConnected
        cmdRefreshTarget_SB.Enabled = bConnected
        cmdClearTargetLog.Enabled = bConnected
        cmdLoadRemoteQueues.Enabled = bConnected
        cmdStopRQ.Enabled = bConnected
        cmdStartRQ.Enabled = bConnected
        cmdDeleteSell2.Enabled = bConnected
        cmd_Target_TxQ.Enabled = bConnected
        chkBranchDebugOn.Enabled = bConnected
    End If

End Sub

Private Sub LoadBranchStoreCombo()
    On Error Resume Next
    rsStores.Open "SELECT Store_Name,Store_VPN_Address  FROM tstore WHERE UPPER(STORE_SYSTEMTYPE) = 'PAPYRUS'", cnINIT, adOpenStatic
    With cboStores
        .Clear
        Do While Not rsStores.EOF
            .AddItem rsStores.Fields("Store_Name")
            rsStores.MoveNext
        Loop
        If .ListCount > 0 Then .ListIndex = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Turning debug mode off before closing.", vbInformation + vbOKOnly, "Notice"
    SetDebugStatus cnTarget, False
    SetDebugStatus cnINIT, False
End Sub
