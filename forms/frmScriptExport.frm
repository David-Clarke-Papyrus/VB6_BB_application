VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Papyrus II support"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9165
   Icon            =   "frmScriptExport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   7995
      Top             =   105
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6435
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   11351
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   670
      BackColor       =   14737632
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User support      "
      TabPicture(0)   =   "frmScriptExport.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtResults"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "SSTab2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Database tuning and analysis    "
      TabPicture(1)   =   "frmScriptExport.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDBName"
      Tab(1).Control(1)=   "cmdConnect"
      Tab(1).Control(2)=   "cmdDumpTriggers"
      Tab(1).Control(3)=   "cmdShrink"
      Tab(1).Control(4)=   "cmdTableStats"
      Tab(1).Control(5)=   "cmdRebuildIndexes"
      Tab(1).Control(6)=   "cmdEXport"
      Tab(1).Control(7)=   "cmdTables"
      Tab(1).Control(8)=   "chkAutoshrink"
      Tab(1).Control(9)=   "G1"
      Tab(1).ControlCount=   10
      Begin TabDlg.SSTab SSTab2 
         Height          =   2895
         Left            =   270
         TabIndex        =   12
         Top             =   690
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         Tab             =   1
         TabsPerRow      =   5
         TabHeight       =   520
         BackColor       =   13489106
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
         TabCaption(0)   =   "General"
         TabPicture(0)   =   "frmScriptExport.frx":03C2
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "lblThree"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblOne"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdSend"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdFetch"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "chkIncludeExportFiles"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Extra"
         TabPicture(1)   =   "frmScriptExport.frx":03DE
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "cmdUpdateFromScript"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Command4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Command3"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdINI"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Command5"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Command6"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cmdInstallServices"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Script execution"
         TabPicture(2)   =   "frmScriptExport.frx":03FA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtServername"
         Tab(2).Control(1)=   "txtDatabasename"
         Tab(2).Control(2)=   "cmdRunScript"
         Tab(2).Control(3)=   "txtScript"
         Tab(2).Control(4)=   "Label8"
         Tab(2).Control(5)=   "Label7"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Linked servers"
         TabPicture(3)   =   "frmScriptExport.frx":0416
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "cmdDropLinkedServers"
         Tab(3).Control(1)=   "cmdAddlinkedServer"
         Tab(3).Control(2)=   "txtLocalName"
         Tab(3).Control(3)=   "txtLinkedServer"
         Tab(3).Control(4)=   "G"
         Tab(3).Control(5)=   "Adodc1"
         Tab(3).Control(6)=   "Label3"
         Tab(3).Control(7)=   "Label2"
         Tab(3).ControlCount=   8
         TabCaption(4)   =   "Properties"
         TabPicture(4)   =   "frmScriptExport.frx":0432
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "cmdAddProperty"
         Tab(4).Control(1)=   "txtPropDescription"
         Tab(4).Control(2)=   "txtPropName"
         Tab(4).Control(3)=   "txtPropValue"
         Tab(4).Control(4)=   "Label6"
         Tab(4).Control(5)=   "Label5"
         Tab(4).Control(6)=   "Label4"
         Tab(4).ControlCount=   7
         Begin VB.CommandButton cmdInstallServices 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Install services"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   3915
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   510
            Width           =   3510
         End
         Begin VB.TextBox txtServername 
            Height          =   375
            Left            =   -71475
            TabIndex        =   46
            Top             =   2295
            Width           =   1155
         End
         Begin VB.TextBox txtDatabasename 
            Height          =   375
            Left            =   -68460
            TabIndex        =   44
            Top             =   2295
            Width           =   1155
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Execute UPDATE_DATA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5085
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   $"frmScriptExport.frx":044E
            Top             =   1995
            Width           =   2370
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Execute UPDATESPOS.SQL in the PBKS\Downloads folder."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3930
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1200
            Width           =   3510
         End
         Begin VB.CheckBox chkIncludeExportFiles 
            Caption         =   "Include report files - usually - No)"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   -68100
            TabIndex        =   41
            Top             =   1290
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdAddProperty 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Add property"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -70050
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1890
            Width           =   1170
         End
         Begin VB.TextBox txtPropDescription 
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
            Height          =   360
            Left            =   -72570
            TabIndex        =   38
            Text            =   "Text1"
            Top             =   2190
            Width           =   2325
         End
         Begin VB.TextBox txtPropName 
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
            Height          =   360
            Left            =   -72570
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   1380
            Width           =   2325
         End
         Begin VB.TextBox txtPropValue 
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
            Height          =   360
            Left            =   -72570
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   1770
            Width           =   2325
         End
         Begin VB.CommandButton cmdINI 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Load .INI to properties"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   2010
            Width           =   1050
         End
         Begin VB.CommandButton cmdDropLinkedServers 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Drop all linked servers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   -68760
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2010
            Width           =   1530
         End
         Begin VB.CommandButton cmdAddlinkedServer 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Add linked server"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   -70470
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   630
            Width           =   1170
         End
         Begin VB.TextBox txtLocalName 
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
            Height          =   360
            Left            =   -72960
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   1020
            Width           =   2325
         End
         Begin VB.TextBox txtLinkedServer 
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
            Height          =   360
            Left            =   -72960
            TabIndex        =   26
            Text            =   "Text1"
            Top             =   630
            Width           =   2325
         End
         Begin VB.CommandButton cmdRunScript 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Execute script"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -74880
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2160
            Width           =   1755
         End
         Begin VB.TextBox txtScript 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H8000000D&
            Height          =   1695
            Left            =   -74910
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   420
            Width           =   7740
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Get some more quotes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2010
            Width           =   3510
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Send database"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   345
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   510
            Width           =   3510
         End
         Begin VB.CommandButton cmdUpdateFromScript 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Execute UPDATES.SQL in the PBKS\Downloads folder."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1185
            Width           =   3510
         End
         Begin VB.Frame Frame1 
            Caption         =   "Support dial-in"
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
            Height          =   990
            Left            =   -74910
            TabIndex        =   17
            Top             =   1740
            Width           =   7545
            Begin VB.CommandButton Command2 
               BackColor       =   &H00D3D3CB&
               Caption         =   "Disconnect"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   5325
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   300
               Width           =   1305
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00D3D3CB&
               Caption         =   "Connect"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   3855
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   315
               Width           =   1305
            End
            Begin VB.Label Label1 
               Caption         =   "Connect to internet for remote support."
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
               Height          =   270
               Left            =   345
               TabIndex        =   20
               Top             =   405
               Width           =   3420
            End
         End
         Begin VB.CommandButton cmdFetch 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Fetch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -70050
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   600
            Width           =   1845
         End
         Begin VB.CommandButton cmdSend 
            BackColor       =   &H00D3D3CB&
            Caption         =   "Send "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   -70050
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1215
            Width           =   1815
         End
         Begin TrueOleDBGrid60.TDBGrid G 
            Bindings        =   "frmScriptExport.frx":04D7
            Height          =   1245
            Left            =   -74850
            OleObjectBlob   =   "frmScriptExport.frx":04EC
            TabIndex        =   32
            Top             =   1560
            Width           =   5355
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   405
            Left            =   -68790
            Top             =   570
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   714
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   2
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
         Begin VB.Label Label8 
            Caption         =   "Server name"
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
            Left            =   -72735
            TabIndex        =   47
            Top             =   2340
            Width           =   1245
         End
         Begin VB.Label Label7 
            Caption         =   "Database name"
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
            Left            =   -69975
            TabIndex        =   45
            Top             =   2325
            Width           =   1470
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Description"
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
            Left            =   -74460
            TabIndex        =   39
            Top             =   2220
            Width           =   1680
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
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
            Left            =   -74460
            TabIndex        =   37
            Top             =   1440
            Width           =   1680
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Value"
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
            Left            =   -74460
            TabIndex        =   36
            Top             =   1830
            Width           =   1680
         End
         Begin VB.Label Label3 
            Caption         =   "Local name"
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
            Left            =   -74850
            TabIndex        =   30
            Top             =   990
            Width           =   1920
         End
         Begin VB.Label Label2 
            Caption         =   "Linked server name"
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
            Left            =   -74850
            TabIndex        =   29
            Top             =   690
            Width           =   1920
         End
         Begin VB.Label lblOne 
            Caption         =   "Fetch the latest update files from the support site."
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
            Left            =   -74460
            TabIndex        =   16
            Top             =   705
            Width           =   4695
         End
         Begin VB.Label lblThree 
            Caption         =   "Send database script to support."
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
            Left            =   -73230
            TabIndex        =   15
            Top             =   1350
            Width           =   3000
         End
      End
      Begin VB.TextBox txtDBName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   450
         Left            =   -68340
         TabIndex        =   11
         Text            =   "PBKS"
         Top             =   510
         Width           =   1620
      End
      Begin VB.TextBox txtResults 
         BackColor       =   &H00E3F9FD&
         ForeColor       =   &H8000000D&
         Height          =   2655
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3645
         Width           =   7920
      End
      Begin VB.CommandButton cmdConnect 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Connect to database (before any other action on this tab)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -72450
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   300
         Width           =   3525
      End
      Begin VB.CommandButton cmdDumpTriggers 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Save Trigger scripts to TRIGGERS.TXT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74805
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   945
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdShrink 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Shrink now"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74805
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2730
         Width           =   2130
      End
      Begin VB.CommandButton cmdTableStats 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Table statistics"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3330
         Width           =   2145
      End
      Begin VB.CommandButton cmdRebuildIndexes 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Rebuild indexes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3930
         Width           =   2145
      End
      Begin VB.CommandButton cmdEXport 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Export script for tables,views etc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74805
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1545
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdTables 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Export script for tables only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2145
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CheckBox chkAutoshrink 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Auto-shrink"
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
         Height          =   405
         Left            =   -74610
         TabIndex        =   1
         Top             =   330
         Width           =   1635
      End
      Begin TrueOleDBGrid60.TDBGrid G1 
         Height          =   4110
         Left            =   -72480
         OleObjectBlob   =   "frmScriptExport.frx":3F82
         TabIndex        =   8
         Top             =   1320
         Width           =   5760
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iFilenum1 As Integer
Dim iFilenum2 As Integer
Dim strLocalRootFolder As String
Dim strFolderOut As String
Dim strLocalPath As String
Dim strServerMachine As String
Dim bInternetDialup As Boolean
Dim strConnectionName As String
Dim strDownloadFolder As String
Dim strUsername As String
Dim strPWD As String
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2
Dim rs As ADODB.Recordset
Dim strBackupFolder As String
Dim sSQL As String

Dim FTPAddress As String
Dim FTPFolder As String

Dim FTPUsername As String
Dim FTPPassword As String
Private strClientCode As String
Private Type TableStats
DataSpaceUsed As String
End Type
Dim strServerName As String
Dim strPOSServerName As String

Dim strPassword As String
Dim X1 As New XArrayDB
Dim strServerMachineName As String
Dim ADOConn As New ADODB.Connection

Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetClassName Lib "USER32" _
    Alias "GetClassNameA" (ByVal hWnd&, _
    ByVal lpClassName$, ByVal nMaxCount&) As Long

Dim sPage As String
Dim rsProperty As New ADODB.Recordset
Dim arCommandLine() As String
Dim mDatabaseName As String
Dim mSystemPrefix As String  ' ie PBKS or PS
Dim mINIFile As String
Dim strPCName As String

Sub Initializesettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strPos As String

    strPos = "0"
    strPCName = Trim(Me.NameOfPC)
    strLocalRootFolder = "C:\" & mSystemPrefix
    mINIFile = strLocalRootFolder & "\" & mSystemPrefix & "WS.INI"
    
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\" & mSystemPrefix & "_S"
        strServerMachine = GetIniKeyValue(mINIFile, "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strServerMachineSharedFolder = "\\" & strServerMachine & "\" & mSystemPrefix & "_S"
    Else
        strServerMachine = GetIniKeyValue(strLocalRootFolder & "\" & mSystemPrefix & "WS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strServerMachineSharedFolder = "C:\" & mSystemPrefix
    End If
    strPos = "1"
    
    strServerName = GetIniKeyValue(mINIFile, "NETWORK", "MAINSQLSERVER", strPCName)
    strPOSServerName = GetIniKeyValue(mINIFile, "NETWORK", "POSSQLSERVER", strPCName)
    strPassword = GetIniKeyValue(mINIFile, "NETWORK", "PASSWORD", "")
    
    strBackupFolder = strLocalRootFolder & "\BU\"
    strPos = "2"
    
    LoadProperties
    strPos = "3"
    
    If fs.FileExists(strServerMachineSharedFolder & "\" & mSystemPrefix & ".INI") Then
        strFETCHLOGSFROM = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FETCHLOGSFROM", "")
        FTPAddress = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPADDRESS", "")
        FTPFolder = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPFOLDER", "")
        FTPUsername = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPUSERNAME", "bt000SA1")
        FTPPassword = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPPASSWORD", "1beach")
        bInternetDialup = (GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP", "TRUE") = "TRUE")
        strConnectionName = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME", "")
    End If
    strFolderOut = strServerMachineSharedFolder & "\FilesForExport"
    strDownloadFolder = strServerMachineSharedFolder & "\DownloadFolder"
   ' MsgBox strDownloadFolder

    On Error Resume Next
    FTPAddress = GetProperty("FTPADDRESS")  'GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPADDRESS", "")
    FTPFolder = GetProperty("FTPFOLDER")   'FTPFolder = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPFOLDER", "")
    FTPUsername = GetProperty("FTPUsername")   'GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPUSERNAME", "bt000SA1")
    FTPPassword = GetProperty("FTPPassword")      'GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FTPPASSWORD", "1beach")
    bInternetDialup = GetProperty("INTERNETDIALUP") = "TRUE"         '(GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP", "TRUE") = "TRUE")
    strConnectionName = GetProperty("CONNECTIONNAME")     'GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME", "")
    strFETCHLOGSFROM = GetProperty("FETCHLOGSFROM")     'GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME", "")
    On Error GoTo errHandler
    
    strPos = "5"
    Set ADOConn = New ADODB.Connection
    ADOConn.Provider = "sqloledb"
    ADOConn.ConnectionTimeout = 10
    sSQL = "Data Source=" & strServerName & ";Initial Catalog=" & mDatabaseName & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
    ADOConn.Open sSQL
    strPos = "6"
    If mSystemPrefix = "PBKS" Then
        Set rs = New ADODB.Recordset
        rs.Open "SELECT dbo.tStore.STORE_Code FROM dbo.tConfiguration INNER JOIN dbo.tStore ON dbo.tConfiguration.CF_DefaultStoreID = dbo.tStore.STORE_ID", ADOConn, adOpenStatic
        strPos = "7"
        If Not rs.EOF And Not rs.BOF Then
            strClientCode = FNS(rs.Fields(0))
        Else
            strClientCode = "UNK"
        End If
        rs.Close
        Set rs = Nothing
    End If
    strPos = "9"
    ADOConn.Close
    strPos = "10"
    On Error Resume Next
    If Not fs.FolderExists(strFolderOut) Then
        strPos = "9.01"
        fs.CreateFolder strFolderOut
        strPos = "9.1"
    End If
    If Not fs.FolderExists(strDownloadFolder) Then
        strPos = "9.02"
        fs.CreateFolder strDownloadFolder
        strPos = "9.2"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Initializesettings", , , , "strPos", Array(strPos)
    HandleError
End Sub
Private Sub Connect(Optional pname As String)
    On Error GoTo errHandler
Dim strPos As String

    Set oSQLServer = New SQLDMO.SQLServer
    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
    strPos = "1"
    With oSQLServer
        .LoginSecure = False
        .AutoReConnect = False
        .Connect strServerName, "sa", strPassword
    End With
    strPos = "2"
    If pname > "" Then
        Set oDatabase = oSQLServer.Databases(pname)
    Else
        Set oDatabase = oSQLServer.Databases("PBKS")
    End If
    strPos = "3"
    If ADOConn.State <> adStateOpen Then
        ADOConn.Provider = "sqloledb"
        ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword
    End If
    strPos = "4"
    LoadTriggers
    strServerMachineName = GetIniKeyValue(strLocalPath & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Connect", , , , strPos, Array(strPos, strServerName)
End Sub
Private Function Disconnect()
    On Error GoTo errHandler
    On Error Resume Next
    oSQLServer.Disconnect
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Disconnect"
End Function
Public Property Get Clientcode() As String
    Clientcode = strClientCode
End Property
Public Property Get NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NameOfPC"
End Property

Public Sub ExportScript()
    On Error GoTo errHandler
Dim s As String
Dim Flag As SQLDMO_SCRIPT_TYPE
Dim oTable As SQLDMO.Table
Dim oStoredProc As SQLDMO.StoredProcedure2
Dim oView As SQLDMO.View2
Dim oUser As SQLDMO.User
Dim oUDF As SQLDMO.UserDefinedFunction
Dim oDBRole As SQLDMO.DatabaseRole2

    Screen.MousePointer = vbHourglass
    Set oDatabase = oSQLServer.Databases("PBKS")
    s = ""
  For Each oStoredProc In oDatabase.StoredProcedures
   ' Debug.Print oStoredProc.Name
    s = s & oStoredProc.Script
  Next
  For Each oView In oDatabase.Views
    s = s & oView.Script
  Next
  For Each oUser In oDatabase.Users
    s = s & oUser.Script
  Next

  Flag = SQLDMOScript_Default Or SQLDMOScript_Indexes Or SQLDMOScript_DRI_AllConstraints Or SQLDMOScript_Triggers Or SQLDMOScript_DRI_ForeignKeys
  For Each oTable In oDatabase.Tables
    If Not oTable.SystemObject Then
      s = s & oTable.Script(Flag)
    End If
  Next
    iFilenum2 = FreeFile
Dim fs As New FileSystemObject
    fs.DeleteFile strFolderOut & "\DBScript.SQL"
    Open strFolderOut & "\DBScript.SQL" For Output As #iFilenum2
    Print #iFilenum2, s
    Close #iFilenum2
    Screen.MousePointer = vbDefault
    MsgBox "Done"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.ExportScript"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ExportScript"
End Sub


Private Sub chkAutoshrink_Click()
    On Error GoTo errHandler
    oDatabase.DBOption.AutoShrink = (chkAutoshrink = 1)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.chkAutoshrink_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.chkAutoshrink_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkAutoshrink_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oDatabase.DBOption.AutoShrink = (chkAutoshrink = 1)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.chkAutoshrink_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.chkAutoshrink_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub







Private Sub cmdAddlinkedServer_Click()
Dim strPath As String
Dim strCommand As String
Dim strScript As String
Dim oTF As New z_TextFile
Dim strOutput As String
Dim Res As Boolean

    lg "Adding linked server . . ."
    strScript = "EXEC sp_addlinkedserver @server='" & Me.txtLocalName & "' ,@provider ='SQLOLEDB',@datasrc = '" & Me.txtLinkedServer & "' , @srvproduct=''"
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & strServerName & " -dmaster -Q """ & strScript & """ -o" & strServerMachineSharedFolder
        'ShellandWait strCommand & "\OSQL_LOG.TXT", 100
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand & "\OSQL_LOG.TXT")
    
    strScript = "EXEC sp_addlinkedsrvlogin '" & txtLocalName & "' , 'false', 'sa', 'sa', ''"
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & strServerName & " -dmaster -Q """ & strScript & """ -o" & strServerMachineSharedFolder
    
        'ShellandWait strCommand & "\OSQL_LOG.TXT", 100
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand & "\OSQL_LOG.TXT")
    
    oTF.OpenTextFileToRead strServerMachineSharedFolder & "\OSQL_LOG.TXT"
    strOutput = oTF.ReadWholeFilewithBreaks
    oTF.CloseTextFile

    lg strOutput

End Sub



Private Sub cmdDropLinkedServers_Click()
Dim strPath As String
Dim strCommand As String
Dim strScript As String
Dim oTF As New z_TextFile
Dim strOutput As String
Dim Res As Boolean

    strScript = "exec sp_dropserver '" & txtLocalName & "','droplogins'"
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & strServerName & " -dmaster -Q """ & strScript & """ -o" & strServerMachineSharedFolder

        lg "Dropping linked server . . ."
        'ShellandWait strCommand & "\OSQL_LOG.TXT", 100
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand & "\OSQL_LOG.TXT")
        
    oTF.OpenTextFileToRead strServerMachineSharedFolder & "\OSQL_LOG.TXT"
    strOutput = oTF.ReadWholeFilewithBreaks
    oTF.CloseTextFile
    lg strOutput


End Sub
Private Sub cmdConnect_Click()
    On Error GoTo errHandler
    If txtDBName > "" Then
        Connect txtDBName
    Else
        Connect
    End If
    chkAutoshrink = IIf(oDatabase.DBOption.AutoShrink, 1, 0)
    cmdDumpTriggers.Enabled = True
    cmdEXport.Enabled = True
    cmdTables.Enabled = True
    cmdShrink.Enabled = True
    cmdTableStats.Enabled = True
    cmdRebuildIndexes.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdConnect_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdEXport_Click()
    On Error GoTo errHandler
    ExportScript
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdEXport_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdExtractSchema_Click()
Dim strCommand  As String
Dim strCommandFilename As String
Dim fs As New FileSystemObject
Dim Res

    strCommandFilename = strLocalRootFolder & "\SchemaCommand.SDP"
    strCommand = App.Path & "\Executables\SQLSCHEMAv4.EXE " & strCommandFilename
    
    If fs.FileExists(strCommandFilename) Then
        'ShellandWait strCommand
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand)
        
    End If

End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim OpenHndl As Long
'Dim pWNDW As Long
Dim lThreadId  As Long
Dim lProcessId As Long

    If MsgBox("Ensure all Papyrus II applications (including the print server) are closed on the server before clicking OK." & vbCrLf _
    & "You can click Cancel to leave this procedure.", vbOKCancel + vbInformation, "Warning") = vbCancel Then
        Exit Sub
    End If
    
    'Force ending of all remnant type processes belonging to PBKS
    
    
    OpenHndl = FindWindow(vbNullString, "PBKS Application")
    If OpenHndl <> 0 Then
        Call SendMessage(OpenHndl, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKS Console")
    If OpenHndl <> 0 Then
        Call SendMessage(OpenHndl, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKS Reports")
    If OpenHndl <> 0 Then
        Call SendMessage(OpenHndl, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKSUI.EXE")
    If OpenHndl <> 0 Then
        Call SendMessage(OpenHndl, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "CONSOLE.EXE")
    If OpenHndl <> 0 Then
        Call SendMessage(OpenHndl, WM_CLOSE, 0&, 0&)
    End If
    
    FetchFiles
    
    lg vbCrLf & "Running Update_Data SQL script. . ."
    ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=45", "sa", ""
    ADOConn.CommandTimeout = 300
    ADOConn.Execute "Execute UPDATE_DATA"
    ADOConn.Close
   
    
    
    MsgBox "The Fetch operation has finished", vbInformation + vbOKOnly, "Status"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdINI_Click()
    On Error GoTo errHandler
    If MsgBox("You want to load the values in the .INI file into the properties table?.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
'    MsgBox "server name = " & strServerName
' '   ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=;Network Library=dbmssocn;Connect Timeout=45", "sa", ""
'    MsgBox "Pos 1"
    InsertProperty "BACKUP", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "FOLDERS", "BACKUP", "REM")
    InsertProperty "BACKUPCOMPRESSION", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "BACKUPCOMPRESSION ", "TRUE")
    InsertProperty "BACKUPMEDIUM", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "BACKUPMEDIUM ", "DISK")
    InsertProperty "CDTYPE", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "CDTYPE ", "RO")
    InsertProperty "LABELPRINTER", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "LABELPRINTER ", "OKI")
    InsertProperty "BOOKFINDROOT", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "BOOKDATA", "BOOKFINDROOT", "c:\Bookfind")
    InsertProperty "BOOKFINDFACET", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "BOOKDATA", "BOOKFINDFACET", "WEBK")
    InsertProperty "MAXBROWSE", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "MAXBROWSE", "1000")
    InsertProperty "EDIENABLED", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "PRINTING", "EDIENABLED", "FALSE")
    InsertProperty "DEFAULTAREACODE", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "LOCAL", "DEFAULTAREACODE", "")
    InsertProperty "INTERNETDIALUP", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP ", "TRUE")
    InsertProperty "CONNECTIONNAME", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME ", "")
    InsertProperty "TIMERINTERVAL", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "FRONTDESK", "TIMERINTERVAL", "3000")
    InsertProperty "SendsCR", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "FRONTDESK", "SENDSCR", False)
    InsertProperty "COMPORTSETTINGS", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "FRONTDESK", "COMPORTSETTINGS", "9600,n,8,1")
    InsertProperty "COMPORTNumber", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "FRONTDESK", "COMPORTNumber", "1")
    InsertProperty "POSACTIVE", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "POS", "POSACTIVE", "FALSE")
    InsertProperty "TRANSFERCALC", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "TRANSFERCALC", "VATDISC")
    InsertProperty "SUPPLERINVOICETOLERANCE", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "SUPPLERINVOICETOLERANCE", ".005")
    InsertProperty "ROUNDPRICETO", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "ROUNDPRICETO", "0")
    InsertProperty "VOUCHERREPORTTOGETHER", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "VOUCHERREPORTTOGETHER", "")
    InsertProperty "ENABLEBOOKCLUBRETURN", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "ENABLEBOOKCLUBRETURN", "FALSE")
    InsertProperty "PRINTSERVERMACHINE", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "NETWORK", "PRINTSERVERMACHINE", "")
    InsertProperty "ISSUEBOOKCLUBRETURNDOCS", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "POS", "ISSUEBOOKCLUBRETURNDOCS", "FALSE")
    InsertProperty "ALLOWANTIQUARIANSEARCH", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "ALLOWANTIQUARIANSEARCH", "1")
    InsertProperty "SHOWQUOTES", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "SHOWQUOTES", "TRUE")
    InsertProperty "SHOWALLAPPROS", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "SHOWALLAPPROS", "TRUE")
    InsertProperty "EXPORTTOPASTELENABLED", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "PASTEL", "EXPORTTOPASTELENABLED", "FALSE")
    InsertProperty "CONTRA_ACCOUNT_INV", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "PASTEL", "CONTRA_ACCOUNT_INV", "")
    InsertProperty "CONTRA_ACCOUNT_SINV", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "PASTEL", "CONTRA_ACCOUNT_SINV", "")
    InsertProperty "INVTOTALSEQ", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "INVTOTALSEQ", "EVR")
    InsertProperty "HIDELOCALSKUONINV", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "HIDELOCALSKUONINV", "FALSE")
    InsertProperty "SetSupplierIDFROMPO", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "SetSupplierIDFROMPO", "FALSE")
    InsertProperty "AllowInvoiceDateOverride", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "AllowInvoiceDateOverride", "FALSE")
    InsertProperty "BOOKFINDISBN13ENABLED", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "BOOKDATA", "BOOKFINDISBN13ENABLED", "FALSE")
    
    InsertProperty "FETCHLOGSFROM", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "FETCHLOGSFROM", "FALSE")
    InsertProperty "FTPADDRESS", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "FTPADDRESS", "FALSE")
    InsertProperty "FTPFOLDER", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "FTPFOLDER", "FALSE")
    InsertProperty "FTPUSERNAME", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "FTPUSERNAME", "FALSE")
    InsertProperty "FTPPASSWORD", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "FTPPASSWORD", "FALSE")
    InsertProperty "INTERNETDIALUP", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP", "FALSE")
    InsertProperty "CONNECTIONNAME", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME", "FALSE")
    InsertProperty "FETCHLOGSFROM", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "FETCHLOGSFROM", "")
    InsertProperty "ADMINISTRATOREMAIL", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "ADMINISTRATOREMAIL", "")
    
    InsertProperty "ADMINISTRATOREMAIL", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "SUPPORT", "ADMINISTRATOREMAIL", "")
    
    InsertProperty "CENTRALFTPADDRESS", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "CENTRAL", "FTPADDRESS", "")
    InsertProperty "CENTRALFTPUSERNAME", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "CENTRAL", "FTPUSERNAME", "")
    InsertProperty "CENTRALFTPPASSWORD", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "CENTRAL", "FTPPASSWORD", "")
    InsertProperty "CENTRALFTPFOLDER", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "CENTRAL", "FTPFOLDER", "")
    InsertProperty "DELAYINSECONDS", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "CENTRAL", "DELAYINSECONDS", "1000")
    
    InsertProperty "SMTPServer", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "SMTPServer", "")
    InsertProperty "SMTP_Username", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "SMTP_Username", "")
    InsertProperty "SMTP_Password", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "SMTP_Password", "")
    InsertProperty "EmailFrom", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "EmailFrom", "")
    InsertProperty "Subject", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "Subject", "")
    InsertProperty "SenderName", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "SenderName", "")
    InsertProperty "TestMode", GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "EMAIL", "TestMode", "TRUE")
    
    strFETCHLOGSFROM = GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "FETCHLOGSFROM", "")

    MsgBox "Done", vbInformation, "Status"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdINI_Click"
End Sub
Private Sub InsertProperty(pPropertyName As String, pVal As String, Optional pDescription As String)
    On Error Resume Next
    ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=45", "sa", ""
    
    If pDescription > "" Then
        ADOConn.Execute "INSERT INTO tPROPERTY (PropertyKey,PropertyValue,PropertyDescription) VALUES ('" & pPropertyName & "','" & pVal & "','" & pDescription & "')"
    Else
        ADOConn.Execute "INSERT INTO tPROPERTY (PropertyKey,PropertyValue) VALUES ('" & pPropertyName & "','" & pVal & "')"
    End If
    
    ADOConn.Close
End Sub

Private Sub cmdAddProperty_Click()
    On Error Resume Next
    InsertProperty Trim(txtPropName), Trim(txtPropValue), Trim(txtPropDescription)
    If Err Then
        MsgBox "The property could not be added, perhaps it already exists.", vbInformation, "Status"
    End If
End Sub


Private Sub cmdInstallServices_Click()
Dim fs As New FileSystemObject

    Screen.MousePointer = vbHourglass
    If Not fs.FileExists("C:\PBKS\Services\SRVANY.EXE") Then
        MsgBox "The file C:\PBKS\Services\SRVANY.EXE does not exist and the dispatcher and/or the POS server services could not be set up."
        Exit Sub
    End If

'Sets up POS Server as a service
    If GetProperty("POSACTIVE") = "TRUE" And fs.FileExists("C:\PBKS\Services\SRVANY.EXE") Then
        SetupServicePOS
    End If

'Sets up PBKS_Dispatch as a service
    If fs.FileExists("C:\PBKS\Services\SRVANY.EXE") Then
        SetupServiceDispatch
    End If
    Screen.MousePointer = vbDefault
    MsgBox "Services added", vbInformation + vbOKOnly, "Status"
    
End Sub



Private Sub cmdRebuildIndexes_Click()
    On Error GoTo errHandler
Dim oTable As SQLDMO.Table
    For Each oTable In oDatabase.Tables
        If Not oTable.SystemObject Then oTable.RebuildIndexes
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdRebuildIndexes_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRunScript_Click()
    On Error GoTo errHandler
Dim strCommand As String
Dim oTF As New z_TextFile
Dim strOutput As String
Dim Res As Boolean

    Screen.MousePointer = vbHourglass
    
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & txtServername & " -d" & txtDatabasename & " -Q """ & Trim(txtScript) & """ -o" & strServerMachineSharedFolder
    
    lg "Executing script . . ."
        
    'ShellandWait strCommand & "\OSQL_LOG.TXT", 100
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand & "\OSQL_LOG.TXT")

    oTF.OpenTextFileToRead strServerMachineSharedFolder & "\OSQL_LOG.TXT"
    strOutput = oTF.ReadWholeFilewithBreaks

    lg strOutput
    Screen.MousePointer = vbDefault
    MsgBox "Execution attempt complete", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdRunScript_Click"
End Sub

Private Sub cmdShrink_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    oDatabase.Shrink 10, SQLDMOShrink_Default
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdShrink_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTables_Click()
    On Error GoTo errHandler
Dim s As String
Dim Flag As SQLDMO_SCRIPT_TYPE
Dim oTable As SQLDMO.Table
Dim oStoredProc As SQLDMO.StoredProcedure2
Dim oView As SQLDMO.View2
Dim oUser As SQLDMO.User
Dim oUDF As SQLDMO.UserDefinedFunction
Dim oDBRole As SQLDMO.DatabaseRole2
Dim srtrs As ADODB.Recordset
Dim sTmp As String
Dim objDMO  As z_SQLDMO

    Set objDMO = New z_SQLDMO
    objDMO.Component oDatabase
    Screen.MousePointer = vbHourglass
    objDMO.CreateTableScript strFolderOut
    Screen.MousePointer = vbDefault
    MsgBox "Done"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.cmdTables_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdTables_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTableStats_Click()
    On Error GoTo errHandler
Dim oTable As SQLDMO.Table
Dim oStats As TableStats
Dim rs As ADODB.Recordset
Dim frm As New frmTableSTats

    Set rs = New ADODB.Recordset
    rs.Fields.Append "Name", adVarChar, 40
    rs.Fields.Append "DataSpaceUsed", adVarChar, 30
    rs.Fields.Append "IndexSpaceUsed", adVarChar, 30
    rs.Fields.Append "Rows", adVarChar, 30
    rs.Open
    For Each oTable In oDatabase.Tables
        rs.AddNew
        rs.Fields("Name") = oTable.Name
        rs.Fields("DataSpaceUsed") = oTable.DataSpaceUsed
        rs.Fields("IndexSpaceUsed") = oTable.IndexSpaceUsed
        rs.Fields("Rows") = oTable.Rows
        rs.Update
    Next
    frm.Component rs
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdTableStats_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSend_Click()
    On Error GoTo errHandler
Dim objDMO As New z_SQLDMO
Dim strCommandFilename As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim RetVal
Dim Res As Boolean
Dim fils
Dim f As File

    Screen.MousePointer = vbHourglass
    lg "Creating schema . . . "
    DoEvents
    
    Set fils = fs.GetFolder(strServerMachineSharedFolder & "\FilesForExport").files
    For Each f In fils
        If UCase(Right(f.Name, 4)) = ".XML" Then
            f.Delete True
        End If
    Next
    
    strCommand = strServerMachineSharedFolder & "\Executables\SQLSCHEMAV4.EXE SQLSCHEMA_" & mSystemPrefix & ".SDP"
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    If fs.FileExists(strServerMachineSharedFolder & "\Executables\SQLSCHEMA_PBKSFD.SDP") Then
        MsgWaitObj 5000
        strCommand = strServerMachineSharedFolder & "\Executables\SQLSCHEMAV4.EXE SQLSCHEMA_PBKSFD.SDP"
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    Else
        If GetProperty("POSACTIVE") = "TRUE" Then
            MsgBox "No schema for POS database is produced"
        End If
    End If
    
    lg "Transmitting to support . . . " & FTPAddress & " : " & FTPFolder
    ManageTransmit False ' (chkIncludeExportFiles = 1)
    Screen.MousePointer = vbDefault
    lg "Finished "
    MsgBox "Files sent to support", , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdSend_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdUpdateFromScript_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    HandleScript
    Screen.MousePointer = vbDefault
    MsgBox "Script has run", vbInformation, "status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdUpdateFromScript_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUpdatePOSFromScript_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    HandleScriptPOS
    Screen.MousePointer = vbDefault
    MsgBox "Script has run", vbInformation, "status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdUpdatePOSFromScript_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub CommandButton1_Click()
    On Error GoTo errHandler
MsgBox "Hello"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CommandButton1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
'Dim clsIPHost As New IPAddrsHostName.clsIPAddrsHostName
Dim lngResult As Long
Dim i As Long
Dim s As String
Dim nPort As Integer
Dim sServer As String
Dim objSWbemLocator
Dim objSWbemServices
Dim colSWbemObjectSet
Dim Obj
Dim strOSVersion As String

    lg "Connecting . . . "
    Set fINET = New wininet
    If bInternetDialup = True Then
        lngResult = fINET.StartDUN(0, strConnectionName, True)
    End If
    strOSVersion = OSVersion
    If strOSVersion = "Windows 2000" Or strOSVersion = "Windows XP" Then
        Set objSWbemLocator = New WbemScripting.SWbemLocator
        Set objSWbemServices = objSWbemLocator.ConnectServer(sServer, "root\cimv2", "", "")
        Dim col As Collection
        
        
        Set colSWbemObjectSet = objSWbemServices.ExecQuery("Select * From Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
        lg ""
        lg "Network address"
        lg "-----------------"
        For Each Obj In colSWbemObjectSet
            For i = 0 To UBound(Obj.IPAddress)
                If InStr(1, Obj.Description(i), "PPP") > 1 Then
                    s = s & Obj.Description(i) & ":" & Obj.IPAddress(i) & vbCrLf
                End If
            Next
        Next
        lg s
        lg "-----------------"
        
        Set objSWbemServices = Nothing
        Set colSWbemObjectSet = Nothing
        Set objSWbemLocator = Nothing
        Set col = Nothing
    End If


End Sub
'Private Sub wsTCP_OnDataArrival(ByVal bytesTotal As Long)
'  Dim sBuffer As String
'  Dim i, j As Integer
'
'    wsTCP.GetData sBuffer
' ' txtSource = txtSource & sBuffer
'    wsTCP.CloseWinsock
'Set wsTCP = Nothing
'    i = InStr(1, sBuffer, "IP Address properties of your Internet Connection") + Len("IP Address properties of your Internet Connection")
'    j = InStr(1, sBuffer, "-->")
'
'lg "IP address for support person -->  " & Mid(sBuffer, i, j - i)
'End Sub
'Private Sub wsTCP_OnConnect()
'  wsTCP.SendData "GET /" & sPage & " HTTP/1.0" & vbCrLf & vbCrLf
'End Sub

Private Sub Command2_Click()
    fINET.Hangup
End Sub

Private Sub Command3_Click()
    FetchQuotes
End Sub

Private Sub Command4_Click()
    On Error GoTo errHandler
Dim FTP1 As New FTPClass
Dim fs As New FileSystemObject
Dim Zip
Dim Res
Dim cmd As ADODB.Command

    Screen.MousePointer = vbHourglass
    lg "Backing up database . . . "
    DoEvents
    
    ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=45", "sa", ""
    ADOConn.Execute "dbo.RenumberQuotes"
    
    
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = ADOConn
    If fs.FileExists(strBackupFolder & "\PBKS.BAK") Then fs.DeleteFile strBackupFolder & "\PBKS.BAK", True

    cmd.CommandType = adCmdText
    cmd.CommandText = "BACKUP DATABASE " & "PBKS" & " to disk = '" & strBackupFolder & "\PBKS.BAK' WITH INIT, NAME = 'Full Backup of PBKSDATA'"
    cmd.Execute

    Set cmd = Nothing
    ADOConn.Close
    
    
    If Not fs.FolderExists(strBackupFolder) Then
        fs.CreateFolder strServerMachineSharedFolder & "\BU"
    End If
    
    Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
    If fs.FileExists(strBackupFolder & "\DB.ZIP") Then
        fs.DeleteFile strBackupFolder & "\DB.ZIP", True
    End If
    
    lg "Zipping database . . . "
    
    
    Zip.CreateZip strBackupFolder & "\DB.ZIP", ""
    Zip.PreservePaths = False
    Zip.AddFile strBackupFolder & "\PBKS.BAK", ""
    If Zip.LastError <> 0 Then
        MsgBox "Zipping errors file was not successful. Contact support person"
    End If
    Zip.Close
    Set Zip = Nothing
    
    lg "Sending database to " & FTPAddress & " . . . "
    
    
    Res = FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True)
    If Res = True Then
        Res = FTP1.SetCurrentFolder(FTPFolder & "/DB")
       ' Res = FTP1.SetCurrentFolder("public_ftp/LOGS")
        If Res = False Then
            lg "Cannot set current folder " & FTPFolder & "/DB"
            Exit Sub
        End If
    '''''''''''''''''''
    'SEND File=======================================================
        Res = FTP1.PutFile(strBackupFolder & "\DB.ZIP", strClientCode & "_DB.ZIP", True)
        If Res = False Then
            lg "Cannot put file " & strBackupFolder & "\DB.ZIP"""
            Exit Sub
        End If
    '''''''
    'Close FTP connection============================================
        FTP1.CloseFTP
    Else
        lg "Cannot open FTP site: " & FTPAddress & ",   " & FTPUsername & ",   " & FTPPassword
    End If
    
    Screen.MousePointer = vbDefault
    lg "Database sent to " & FTPAddress & " successfully"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Command4_Click"
End Sub

Private Sub Command5_Click()
    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    HandleScriptPOS
    Screen.MousePointer = vbDefault
    MsgBox "POS Script has run", vbInformation, "status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Command5_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command6_Click()
    Screen.MousePointer = vbHourglass
    ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=45", "sa", ""
    ADOConn.CommandTimeout = 300
    
    ADOConn.Execute "Execute UPDATE_DATA"
    
    ADOConn.Close
    Screen.MousePointer = vbDefault
    MsgBox "UPDATE_DATA stored procedure has run", vbInformation, "status"

End Sub

'Private Sub Command3_Click()
'Dim i As Integer
'
'For i = 1 To 100
'    lg "This is a test"
'Next i
'    lg "This is a test - last line"
'
'End Sub

Private Sub Form_Load()
Dim OpenHndl As Long
Dim pWNDW As Long
Dim lThreadId  As Long
Dim lProcessId As Long
    
    On Error GoTo errHandler
    
    arCommandLine = Split(Command(), " ")
    If UBound(arCommandLine) >= 0 Then
        mDatabaseName = arCommandLine(0)
    Else
        mDatabaseName = "PBKS"
    End If
    If UBound(arCommandLine) >= 1 Then
        mSystemPrefix = arCommandLine(1)
    Else
        mSystemPrefix = "PBKS"
    End If
    
    cmdDumpTriggers.Enabled = False
    cmdEXport.Enabled = False
    cmdTables.Enabled = False
    cmdShrink.Enabled = False
    cmdTableStats.Enabled = False
    cmdRebuildIndexes.Enabled = False
    Me.SSTab1.Tab = 0
    Initializesettings
    
    OpenHndl = FindWindow(vbNullString, "PBKS Application")
    If OpenHndl <> 0 Then
        Call SendMessage(pWNDW, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKSUI")
    If OpenHndl <> 0 Then
        Call SendMessage(pWNDW, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKSDLL")
    If OpenHndl <> 0 Then
        Call SendMessage(pWNDW, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKS Console")
    If OpenHndl <> 0 Then
        Call SendMessage(OpenHndl, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKS Reports")
    If OpenHndl <> 0 Then
        Call SendMessage(pWNDW, WM_CLOSE, 0&, 0&)
    End If
    OpenHndl = FindWindow(vbNullString, "PBKS POS Server")
    If OpenHndl <> 0 Then
        Call SendMessage(pWNDW, WM_CLOSE, 0&, 0&)
    End If
    Me.SSTab1.Tab = 0
    Me.SSTab2.Tab = 0
    

    txtServername = strServerName
    txtDatabasename = "PBKS"

'    Me.Adodc2.CommandType = adCmdText
'    Me.Adodc2.RecordSource = "Select PROPT_ID,PROPT_DESCRIPTION FROM tPropertyType ORDER BY PROPT_DESCRIPTION"
'    Me.Adodc2.ConnectionString = ADOConn.ConnectionString
'    List1.Refresh
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If Not oSQLServer Is Nothing Then
        oSQLServer.Disconnect
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub LoadTriggers()
    On Error GoTo errHandler
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim oT As SQLDMO.Table
    Screen.MousePointer = vbHourglass

    k = 0
    For i = 1 To oDatabase.Tables.Count
        Set oT = oDatabase.Tables(i)
        For j = 1 To oT.Triggers.Count
            k = k + 1
            X1.ReDim 1, k, 1, 5
            X1(k, 1) = oT.Name
            X1(k, 2) = oT.Triggers(j).Name
            X1(k, 3) = oT.Triggers(j).Enabled
            X1(k, 4) = j
            X1(k, 5) = i
        Next j
    Next i
    G1.Array = X1
    G1.ReBind
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadTriggers"
End Sub
Private Sub cmdDumpTriggers_Click()
    On Error GoTo errHandler

Dim str As String
Dim fs As New FileSystemObject
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim oT As SQLDMO.Table
Dim objDMO  As z_SQLDMO

    Set objDMO = New z_SQLDMO
    objDMO.Component oDatabase
    
    Screen.MousePointer = vbHourglass
    
    objDMO.CreateTriggerScript strFolderOut

    Screen.MousePointer = vbDefault
    MsgBox "Done"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdDumpTriggers_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    oDatabase.Tables(X1(G1.Bookmark, 5)).Triggers(X1(G1.Bookmark, 4)).Enabled = G1.Text
    LoadTriggers
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.G1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Public Sub ManageTransmit(bIncludeScripts As Boolean)
    On Error GoTo errHandler
Dim lngResult As Long
Dim Res As Boolean
Dim Zip
Dim FTP1 As New FTPClass
Dim fs As New FileSystemObject
Dim arLogs() As String
Dim i As Integer
Dim strPos As String

    Screen.MousePointer = vbHourglass
strPos = "pos 1"
    Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
    If fs.FileExists(strServerMachineSharedFolder & "\BU\LOG.ZIP") Then
        fs.DeleteFile strServerMachineSharedFolder & "\BU\LOG.ZIP", True
    End If
strPos = "pos 2"
    
    Zip.CreateZip strServerMachineSharedFolder & "\BU\LOG.ZIP", ""
strPos = "pos 3"
    
    Zip.ProcessSubfolders = True
    Zip.BasePath = ""
strPos = "pos 4"
    
    Zip.PreservePaths = False
    Zip.AddFile strServerMachineSharedFolder & "\Errors.*", ""
    Zip.AddFile strServerMachineSharedFolder & "\Printers\*.*", ""
    Zip.AddFile strServerMachineSharedFolder & "\Sendlog*.*", ""
    Zip.AddFile strServerMachineSharedFolder & "\Fetchlog*.*", ""
    Zip.AddFile strServerMachineSharedFolder & "\DB.*", ""
    Zip.AddFile strServerMachineSharedFolder & "\emaillog.*", ""
strPos = "pos 5"
    
    Zip.AddFile strServerMachineSharedFolder & "\FilesForExport\*.XML", ""
    Zip.AddFile strServerMachineSharedFolder & "\TEMPLATES\*.XSL", ""
    Zip.AddFile strServerMachineSharedFolder & "\TEMPLATES\*.XSLT", ""
strPos = "pos 6"
    
    arLogs = Split(strFETCHLOGSFROM, ",")
strPos = "pos 7"
    
    For i = 0 To UBound(arLogs)
        If arLogs(i) > "" Then
            If fs.FileExists(arLogs(i) & "\POSErrors.txt") Then
strPos = "pos 7" & arLogs(i)
                Zip.AddFile arLogs(i) & "\POSErrors.txt", ""
            End If
        End If
    Next i
strPos = "pos 8"
    
    If bIncludeScripts Then
        If fs.FolderExists(strFolderOut) Then
strPos = "pos 9" & strFolderOut
            Zip.AddFile strFolderOut & "\*.*", ""
        End If
    End If
strPos = "pos 10"
    
    If Zip.LastError <> 0 Then
        MsgBox "Zipping errors file was not successful. Contact support person"
    End If
strPos = "pos 11"
    
    Zip.Close
    Set Zip = Nothing
strPos = "pos 12"
    
''''''''''''''''''''''''
    Set fINET = New wininet
    If bInternetDialup = True Then
        lngResult = fINET.StartDUN(0, strConnectionName, True)
    End If
    
''OPEN FTP Connection===========================================
    Res = FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True)
strPos = "pos 13"
    If Res = True Then
        If FTPFolder > "" Then
            If Left(FTPFolder, 1) <> "/" Then FTPFolder = "/" & FTPFolder
        End If
        Res = FTP1.SetCurrentFolder("/LOGS" & IIf(FTPFolder > "", FTPFolder, ""))
       ' Res = FTP1.SetCurrentFolder("public_ftp/LOGS")
        If Res = False Then
            lg "Cannot set current folder " & "/LOGS" & IIf(FTPFolder > "", FTPFolder, "")
            Exit Sub
        End If
    '''''''''''''''''''
strPos = "pos 14"
    
    'SEND File=======================================================
        Res = FTP1.PutFile(strServerMachineSharedFolder & "\BU\LOG.ZIP", strClientCode & "_LOG.ZIP", True)
        If Res = False Then
            lg "Cannot put file " & strServerMachineSharedFolder & "\BU\LOG.ZIP"
            Exit Sub
        End If
    '''''''
strPos = "pos 15"
    
    'Close FTP connection============================================
        FTP1.CloseFTP
    Else
        lg "Cannot open FTP site: " & FTPAddress & ",   " & FTPUsername & ",   " & FTPPassword
    End If
'    If fINET.IsNetConnectOnline Then
'Close Internet connection=======================================
    fINET.Hangup
''''''''''''''''''''''''''''
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ManageTransmit(bIncludeScripts)", bIncludeScripts, , , "strPos", Array(strPos)
End Sub
Public Sub FetchQuotes()
    On Error GoTo errHandler
Dim lngResult As Long
Dim FTP1 As New FTPClass
Dim fs As New FileSystemObject
Dim f, fc, fol
Dim FTPFile As FTPFileClass
Dim Res As Boolean
Dim Zip
Dim strSQL As String
Dim oT As New z_TextFile
Dim ln As String
Dim ar() As String

    Set fINET = New wininet
 '   If Not fINET.IsNetConnectOnline Then
    If bInternetDialup = True Then
        lngResult = fINET.StartDUN(0, strConnectionName, True)
    End If
       ' Check lngResult = 0, ERR_DUNALREADYOPEN, "Cannot open connection,perhaps it is already open"
    lg "Opening FTP connection . . ."
''OPEN FTP Connection===========================================
    Check FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True), EXC_GENERAL, "Opening FTP"
        lg "Fetching files in Common. . ." & FTPFolder & "/QUOTES"
    Check FTP1.SetCurrentFolder(FTPFolder & "/QUOTES"), EXC_GENERAL, "setting FTP folder"
    For Each FTPFile In FTP1.files
        lg ". . . " & FTPFile.FileName
       Check FTP1.GetFile(FTPFile.FileName, strDownloadFolder & "\" & FTPFile.FileName, True), EXC_GENERAL, "Getting FTP file"
    Next
'unzip all zip files
    lg vbCrLf & "Unzipping. . ."
    Set fol = fs.GetFolder(strDownloadFolder)
    Set fc = fol.files
    For Each f In fc
        If UCase(Left(f.Name, 4)) = "QUOT" And UCase(Right(f.Name, 4)) = ".ZIP" Then
            Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
            Zip.OpenZip (f.Path)
            Zip.BasePath = strDownloadFolder
            Zip.ExtractFile ("*.*")
            Zip.Close
            Set Zip = Nothing
            f.Delete True
        End If
    Next
    
    lg "Updating database . . ."
    For Each f In fc
        If UCase(Left(f.Name, 4)) = "QUOT" Then
        
        'ADOConn.Execute "DBCC CHECKIDENT (tQuote, RESEED, 0)"  '(this resets the ID to start at 1)
'            strSQL = "BULK INSERT PBKS.dbo.tQuote From '" & F.Path & "'" & " WITH (FIELDTERMINATOR = '|', ROWTERMINATOR = '\n',  MAXERRORS =500 )"
            ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=45", "sa", ""
'            ADOConn.Execute strSQL
'            ADOConn.Execute "UPDATE tQUOTE SET FILENAME = '" & F.Name & "' WHERE FILENAME IS NULL"
            
            oT.OpenTextFileToRead f.Path
            Do While Not oT.IsEOF
                ln = oT.ReadLinefromTextFile
                ar = Split(ln, "|")
                On Error Resume Next
                ADOConn.Execute "Insert into tQuote (WORDS,FILENAME) VALUES ('" & ar(1) & "','" & f.Name & "')"
                On Error GoTo errHandler
            Loop
            oT.CloseTextFile
            
            
            
            ADOConn.Close
            f.Delete True
        End If
        ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=45", "sa", ""
        ADOConn.Execute "dbo.RenumberQuotes"
        ADOConn.Close
        
    Next
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FetchQuotes"
    Exit Sub
    Resume
End Sub
Public Sub FetchFiles()
    On Error GoTo errHandler
Dim lngResult As Long
Dim Res As Boolean
Dim Zip
Dim FTP1 As New FTPClass
Dim fs As New FileSystemObject
Dim f, fc, fol
Dim FTPFile As FTPFileClass
Dim strBUFolder As String
Dim cmd As ADODB.Command
Dim strPos As String

    If strServerName = strPCName Then
       lg "Backing up the database . . ."
       strPos = "1"
       
       '=======
       Connect
       '=======
       
       If fs.FolderExists(strLocalRootFolder & "\BU") Then
           strBUFolder = strLocalRootFolder & "\BU\"
       Else
           strBUFolder = strLocalRootFolder & "\"
       End If
       strPos = "2"
       Set cmd = New ADODB.Command
       cmd.CommandTimeout = 0
       Set cmd.ActiveConnection = ADOConn
       oDatabase.Shrink 10, SQLDMOShrink_Default
       If fs.FileExists(strBUFolder & "PBKS.BAK") Then fs.DeleteFile strBUFolder & "PBKS.BAK", True
       If fs.FileExists(strBUFolder & "PBKSMASTER.BAK") Then fs.DeleteFile strBUFolder & "PBKSMASTER.BAK", True
    
       strPos = "3"
    '   Timer1.Enabled = True
       cmd.CommandType = adCmdText
       cmd.CommandText = "BACKUP DATABASE PBKS to disk = '" & strBUFolder & "PBKS.BAK' WITH INIT, NAME = 'Full Backup of PBKSDATA'"
       cmd.Execute
       cmd.CommandText = "BACKUP DATABASE MASTER to disk = '" & strBUFolder & "PBKSMASTER.BAK' WITH INIT, NAME = 'Full Backup of PBKSMASTER'"
       cmd.Execute
    '   Timer1.Enabled = False
       Set cmd = Nothing
       
       If Not (fs.FileExists(strLocalRootFolder & "\BU\" & "PBKS.BAK") And fs.FileExists(strLocalRootFolder & "\BU\" & "PBKSMASTER.BAK")) Then
           MsgBox "Backup was not successful.Contact support person"
           Exit Sub
       End If
    End If
    strPos = "4"
    lg vbCrLf & "Connecting . . ."
    Set fINET = New wininet
 '   If Not fINET.IsNetConnectOnline Then
 'MsgBox "Fetrch files: bInternetDialup = " & bInternetDialup
    If bInternetDialup = True Then
        lngResult = fINET.StartDUN(0, strConnectionName, True)
    End If
       ' Check lngResult = 0, ERR_DUNALREADYOPEN, "Cannot open connection,perhaps it is already open"
    lg "Opening FTP connection . . ."
''OPEN FTP Connection===========================================
    Check FTP1.OpenFTP(FTPAddress, FTPUsername, FTPPassword, True), EXC_GENERAL, "Opening FTP"
'''''''''''''''''''

    strPos = "5"
'Clear all old files in receiving folder
    Set fol = fs.GetFolder(strDownloadFolder)
    If fol.files.Count > 0 Then
        fs.DeleteFile strDownloadFolder & "\*.*", True
    End If
'Fetch Files=======================================================
'Look for files in the 'Common' folder and in the folder named after the client (as in PBKS.INI)
'Fetch all these into the download folder on the client's server machine
    lg "Fetching files in /COMMON. . ." & IIf(FTPFolder > "", FTPFolder, "")
    If FTPFolder > "" Then
        If Left(FTPFolder, 1) <> "/" Then FTPFolder = "/" & FTPFolder
    End If
    Check FTP1.SetCurrentFolder("/COMMON" & IIf(FTPFolder > "", FTPFolder, "")), EXC_GENERAL, "setting FTP folder"
    For Each FTPFile In FTP1.files
        lg ". . . " & FTPFile.FileName
       Check FTP1.GetFile(FTPFile.FileName, strDownloadFolder & "\" & FTPFile.FileName, True), EXC_GENERAL, "Getting FTP file"
    Next
    strPos = "6"
'unzip all zip files
    lg vbCrLf & "Unzipping. . ."
    Set fol = fs.GetFolder(strDownloadFolder)
    Set fc = fol.files
    For Each f In fc
        If UCase(Right(f.Name, 4)) = ".ZIP" Then
            Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
            Zip.OpenZip (f.Path)
            Zip.BasePath = strDownloadFolder
            Zip.ExtractFile ("*.*")
            Zip.Close
            Set Zip = Nothing
            f.Delete True
        End If
    Next
    strPos = "7"
'Distribute them according to their type as follows
'.EXE .DLL go to folder:Patches;  .DOT go to folder:templates;  .SQL stay in download folder; .NOT stay in download folder
'.XSL or .XSLT got to TEMPLATES folder
'Display progress in test box and when downloading is complete display contents of .NOT file if it exists
    lg "Moving and registering files . . ."
    
    'copies to folders and registers DLLs
    HandleDownload
    
    strPos = "8"
    'Runs SQL scripts if they exist
    lg "Updating database . . ."
    HandleScript
    strPos = "9"
    
    
    
Dim oTF As z_TextFile
    Set oTF = New z_TextFile
    oTF.OpenTextFile strServerMachineSharedFolder & "\UpdateLog.txt"
    oTF.WriteToTextFile Trim(txtResults)
    oTF.CloseTextFile
    
    Check FTP1.SetCurrentFolder("/COMMON" & FTPFolder & "/LOGS"), EXC_GENERAL, "Error setting FTP folder"
'''''''''''''''''''
'SEND File=======================================================
    If fs.FileExists(strServerMachineSharedFolder & "\UpdateLog.txt") Then
        If FTP1.FileExists("UpdateLog.txt") Then
            FTP1.DeleteFile ("UpdateLog.txt")
        End If
        Res = FTP1.PutFile(strServerMachineSharedFolder & "\UpdateLog.txt", "UpdateLog.txt", True)
        Check Res, EXC_GENERAL, "Error transmitting FTP file"
    End If
    If fs.FileExists(strServerMachineSharedFolder & "\OSQL_LOG.txt") Then
        If FTP1.FileExists("OSQL_LOG.txt") Then
            FTP1.DeleteFile ("OSQL_LOG.txt")
        End If
        Res = FTP1.PutFile(strServerMachineSharedFolder & "\OSQL_LOG.txt", "OSQL_LOG.txt", True)
        Check Res, EXC_GENERAL, "Error transmitting FTP file"
    End If
        
    If fs.FileExists(strServerMachineSharedFolder & "\OSQL2_LOG.txt") Then
        If FTP1.FileExists("OSQL2_LOG.txt") Then
            FTP1.DeleteFile ("OSQL2_LOG.txt")
        End If
        Res = FTP1.PutFile(strServerMachineSharedFolder & "\OSQL2_LOG.txt", "OSQL2_LOG.txt", True)
        Check Res, EXC_GENERAL, "Error transmitting FTP file"
    End If
    
    strPos = "10"
'Close FTP connection============================================
    lg "Closing connection. . ."
    On Error Resume Next
    FTP1.CloseFTP
'    If fINET.IsNetConnectOnline Then
'Close Internet connection=======================================
 '   fINET.Hangup
''''''''''''''''''''''''''''
    Set FTP1 = Nothing
    Disconnect
    Exit Sub
errHandler:
'Resume
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.FetchFiles", , , , "Error position = ", Array(strPos)
End Sub






Private Sub lg(pText As String)
    On Error GoTo errHandler
    txtResults = txtResults & vbCrLf & pText
    txtResults.Refresh
    txtResults.SelStart = Len(txtResults) - 1
    txtResults.SelLength = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.lg(pText)", pText
End Sub
Private Sub ClearLog()
    On Error GoTo errHandler
    txtResults = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ClearLog"
End Sub
Private Sub cmdExtract_Click()
    On Error GoTo errHandler
    HandleDownload
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.cmdExtract_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub HandleDownload()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim f, fc, fol
Dim lngReturn As Long
Dim strErrMsg As String
Dim strNewName As String
Dim strErrPos As String


'Get names of all DLLs on the DownloadFolder shared folder on the server
    Set fol = fs.GetFolder(strDownloadFolder).files
    strErrPos = "1"
'Unregister all files of the same names as the downloaded files in the PBKS\Executables folder on the workstation and rename or delete then
    For Each f In fol
        If fs.FileExists(strServerMachineSharedFolder & "\Executables\" & f.Name) Then
            If UCase(Right(f.Name, 4)) = ".DLL" Or UCase(Right(f.Name, 4)) = ".OCX" Then
                If Not UnregisterComEx(strServerMachineSharedFolder & "\Executables\" & f.Name, lngReturn, strErrMsg) Then
                    lg "Cannot unregister " & strServerMachineSharedFolder & "\Executables\" & f.Name & "Procedure halted without completing." & vbCrLf & "Error message is: " & strErrMsg
                    Exit Sub
                End If
            End If
            strNewName = strServerMachineSharedFolder & "\Executables\" & "o" & f.Name
            If fs.FileExists(strNewName) Then
                fs.DeleteFile strNewName, True
            End If
            Name strServerMachineSharedFolder & "\Executables\" & f.Name As strNewName
        End If
    Next
    strErrPos = "2"
'Copy all the DLLs and EXEs on the DownloadFolder shared folder to the PBKS\Executables folder on the workstation
    Set fol = fs.GetFolder(strDownloadFolder).files
    strErrPos = "3"
    For Each f In fol
        If UCase(Right(f.Name, 4)) = ".DLL" Or UCase(Right(f.Name, 4)) = ".OCX" Or UCase(Right(f.Name, 4)) = ".EXE" Then
            fs.CopyFile strDownloadFolder & "\" & f.Name, strServerMachineSharedFolder & "\Executables\" & f.Name, True
        End If
        If UCase(Right(f.Name, 4)) = ".XSL" Or UCase(Right(f.Name, 5)) = ".XSLT" Then
            fs.CopyFile strDownloadFolder & "\" & f.Name, strServerMachineSharedFolder & "\TEMPLATES\" & f.Name, True
        End If
        If UCase(Right(f.Name, 4)) = ".BRE" Or UCase(Right(f.Name, 5)) = ".BDO" Then
            fs.CopyFile strDownloadFolder & "\" & f.Name, strServerMachineSharedFolder & "\ARIA\" & f.Name, True
        End If

    Next
    
    'We must delete any old stuff in Patches
    Set fol = fs.GetFolder(strServerMachineSharedFolder & "\Patches").files
    For Each f In fol
        lg "Deleting file . . ." & strServerMachineSharedFolder & "\Patches\" & f.Name
        f.Delete True
    Next
    strErrPos = "4"
    If fs.FileExists(strDownloadFolder & "\POS.EXE") Then
        fs.CopyFile strDownloadFolder & "\POS.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    If fs.FileExists(strDownloadFolder & "\PBKSDLL.DLL") Then
        fs.CopyFile strDownloadFolder & "\PBKSDLL.DLL", strServerMachineSharedFolder & "\Patches\", True
    End If
    strErrPos = "5"
    If fs.FileExists(strDownloadFolder & "\PBKSUI.EXE") Then
        fs.CopyFile strDownloadFolder & "\PBKSUI.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    strErrPos = "6"
    If fs.FileExists(strDownloadFolder & "\FRONTDESK.EXE") Then
        fs.CopyFile strDownloadFolder & "\FRONTDESK.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    strErrPos = "7"
    If fs.FileExists(strDownloadFolder & "\FRONTDESKDLL.DLL") Then
        fs.CopyFile strDownloadFolder & "\FRONTDESKDLL.DLL", strServerMachineSharedFolder & "\Patches\", True
    End If
    If fs.FileExists(strDownloadFolder & "\PBKSREPUI.EXE") Then
        fs.CopyFile strDownloadFolder & "\PBKSREPUI.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    If fs.FileExists(strDownloadFolder & "\PBKSREPDLL.DLL") Then
        fs.CopyFile strDownloadFolder & "\PBKSREPDLL.DLL", strServerMachineSharedFolder & "\Patches\", True
    End If
    If fs.FileExists(strDownloadFolder & "\POSPROPERTYMANAGER.EXE") Then
        fs.CopyFile strDownloadFolder & "\POSPROPERTYMANAGER.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    If fs.FileExists(strDownloadFolder & "\STMANUAL.EXE") Then
        fs.CopyFile strDownloadFolder & "\STMANUAL.EXE", strServerMachineSharedFolder & "\Patches\", True
    End If
    If fs.FileExists(strDownloadFolder & "\STMANUALDLL.DLL") Then
        fs.CopyFile strDownloadFolder & "\STMANUALDLL.DLL", strServerMachineSharedFolder & "\Patches\", True
    End If
    strErrPos = "8"
    strErrPos = "9"
'Register all DLLs on the workstation PBKS\Executables folder
'    Set fol = fs.GetFolder(strServerMachineSharedFolder & "\Executables").Files
    Set fol = fs.GetFolder(strDownloadFolder).files
    For Each f In fol
        If UCase(Right(f.Name, 4)) = ".DLL" Or UCase(Right(f.Name, 4)) = ".OCX" Then
            If Not RegisterComEx(strServerMachineSharedFolder & "\Executables\" & f.Name, lngReturn, strErrMsg) Then
                lg "Cannot register " & strServerMachineSharedFolder & "\Executables\" & f.Name & " Procedure halted without completing." & vbCrLf & "Error message is: " & strErrMsg
            Else
                lg "Registering file . . ." & strServerMachineSharedFolder & "\Executables\" & f.Name
            End If
        End If
    Next
    strErrPos = "10"
    Exit Sub
errHandler:
    ErrPreserve
    If strErrPos > "4" And strErrPos < "9" Then
        lg "Cannot copy a file to \Patches or PBKS.INI to \PBKS_S." & vbCrLf & "Position is: " & strErrPos
        Resume Next
    End If
        
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleDownload", , , , "Position", Array(strErrPos)
End Sub

Private Sub HandleScript()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strPath As String
Dim strCommand As String
Dim Res As Boolean
    strPath = strDownloadFolder & "\UPDATES.SQL"
    
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & strServerName & " -dPBKS -i" & strPath & " -o" & strServerMachineSharedFolder
  '  MsgBox strPath
    If fs.FileExists(strPath) Then
        lg "Updating using UPDATES.SQL executing . . ." & vbCrLf & strCommand & "\OSQL_LOG.TXT"
        'Res = ShellandWait(strCommand & "\OSQL_LOG.TXT", 30000)
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand & "\OSQL_LOG.TXT")
    End If
    
    Exit Sub
   
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleScript", , , , "strCommand  res", Array(strCommand, Res)
End Sub
Private Sub HandleScriptPOS()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strPath As String
Dim strCommand As String
Dim Res
    strPath = strDownloadFolder & "\UPDATESPOS.SQL"
    
    strCommand = "OSQL.EXE -Usa -P" & strPassword & " -S" & strPOSServerName & " -dPBKSFD -i" & strPath & " -o" & strServerMachineSharedFolder
    If fs.FileExists(strPath) Then
        lg "Updating using UPDATESPOS.SQL executing . . ."
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand & "\OSQLPOS_LOG.TXT")
    End If
    Exit Sub
   
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.HandleScriptPOS", , , , "strCommand  res", Array(strCommand, Res)
End Sub
Private Sub HandleScriptUpdateData()

End Sub
Private Function fRunningInIde() As Boolean
    On Error GoTo errHandler
Dim sClassName As String
Dim nStrLen    As Long

    '
    ' See if we're running in the IDE.
    '
    sClassName = String$(260, vbNullChar)
    nStrLen = GetClassName(Me.hWnd, sClassName, Len(sClassName))
    If nStrLen Then sClassName = Left$(sClassName, nStrLen)
    
    fRunningInIde = (sClassName = "ThunderFormDC")
   ' MsgBox sClassName & "    " & fRunningInIde
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.fRunningInIde"
End Function
Public Function GetProperty(pKey As String) As String
    On Error GoTo errHandler
    rsProperty.MoveFirst
    rsProperty.Find "PropertyKey = '" & pKey & "'"
    If rsProperty.EOF Then
        GetProperty = ""
        Exit Function
    End If
    If rsProperty.Fields.Count > 0 Then GetProperty = Trim(FNS(rsProperty.Fields(1)))
    Exit Function
errHandler:
    ErrPreserve
    If Err = 3021 Then
        MsgBox "Missing property: " & pKey
        Exit Function
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.GetProperty(pKey)", pKey
End Function
Public Function LoadProperties() As Boolean
    On Error GoTo errHandler
Dim strPos As String
Dim sSQL As String

strPos = "Pos 1"
    ADOConn.Provider = "sqloledb"
    sSQL = "Data Source=" & strServerName & ";Initial Catalog=" & mDatabaseName & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
    
    ADOConn.Open sSQL
strPos = "Pos 2"
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.Open sSQL, ADOConn, adOpenKeyset, adLockOptimistic
strPos = "Pos 3"
    Set rsProperty.ActiveConnection = Nothing
    
    ADOConn.Close
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.LoadProperties", , , , "Connection String", Array(sSQL, strPos)
End Function

