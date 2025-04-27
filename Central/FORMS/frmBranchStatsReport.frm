VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBranchStatsReport 
   Caption         =   "Branch loyalty customer status report"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   11295
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Refresh"
      Height          =   390
      Left            =   8625
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   135
      UseMaskColor    =   -1  'True
      Width           =   2580
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Get updated stats from branches"
      Height          =   390
      Left            =   165
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2565
      UseMaskColor    =   -1  'True
      Width           =   2580
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   9885
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmBranchStatsReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2565
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmBranchStatsReport.frx":038A
      Height          =   1950
      Left            =   150
      OleObjectBlob   =   "frmBranchStatsReport.frx":039F
      TabIndex        =   0
      Top             =   555
      Width           =   11040
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   3540
      Top             =   2805
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
   Begin VB.Label lblAddCnt 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   2370
      TabIndex        =   3
      Top             =   195
      Width           =   1230
   End
   Begin VB.Label lblTPCnt 
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1155
      TabIndex        =   2
      Top             =   195
      Width           =   1155
   End
End
Attribute VB_Name = "frmBranchStatsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TPCnt As Long
Dim AddCnt As Long

Dim oSQL As New z_SQL

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Dim f As New frmStoreSelection

    f.Component "Branch statistics request", "BS"
    f.Show vbModal
    MsgBox "Wait a while (could be a minute or more) and then click 'Refresh'.", vbInformation, "Status"
    Unload f
End Sub

Private Sub Form_Load()

    oSQL.CountingLocalLoyaltyCustomerRecords TPCnt, AddCnt
    Me.lblTPCnt.Caption = "Cust: " & CStr(TPCnt)
    Me.lblAddCnt.Caption = "Add: " & CStr(AddCnt)
    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select * FROM tLoyaltyCustStats"
    Me.Adodc1.ConnectionString = oPC.ConnectionString
    G.DataSource = Me.Adodc1
    Me.Width = 11040
    Me.Height = 3630

End Sub
