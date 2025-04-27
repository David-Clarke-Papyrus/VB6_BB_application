VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmSQL 
   Caption         =   "Dynamic SQL code for accounting"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   10035
      TabIndex        =   2
      Top             =   5775
      Width           =   285
   End
   Begin VB.TextBox txtSQL 
      ForeColor       =   &H8000000D&
      Height          =   2385
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3735
      Width           =   9810
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmSQL.frx":0000
      Height          =   2670
      Left            =   165
      OleObjectBlob   =   "frmSQL.frx":0015
      TabIndex        =   1
      Top             =   420
      Width           =   9765
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   8700
      Top             =   0
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
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()
    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select * FROM tACCOUNTING_SQL ORDER BY SQLString_Type, SQLString_SequenceNo"
    Me.Adodc1.ConnectionString = oPC.PapyrusConnectionstring
    G.DataSource = Me.Adodc1

End Sub

Private Sub Form_Unload(Cancel As Integer)
    G.Update
End Sub



Private Sub G_DblClick()
    Me.txtSQL = G.Text
End Sub
Private Sub Command1_Click()
    G.Col = 2
    G.Text = txtSQL
End Sub

