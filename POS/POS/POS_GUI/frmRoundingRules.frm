VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRoundingRules 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Rounding rules"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F2E0D9&
      Caption         =   "OK"
      Height          =   405
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2070
      Width           =   1125
   End
   Begin MSAdodcLib.Adodc DC1 
      Height          =   330
      Left            =   510
      Top             =   2835
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
      CommandType     =   8
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
   Begin VB.CommandButton cmdLoadDefaults 
      BackColor       =   &H00F2E0D9&
      Caption         =   "Load defaults"
      Height          =   405
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2070
      Width           =   1725
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   1830
      Left            =   90
      OleObjectBlob   =   "frmRoundingRules.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   4545
   End
End
Attribute VB_Name = "frmRoundingRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XA As New XArrayDB

Sub GetRules()

    oPC.OpenLocalDatabase
    DC1.CommandType = adCmdText
    DC1.RecordSource = "SELECT RR_LowerBound,RR_UpperBound,RR_RoundTo FROM tROUNDINGRULE"
    DC1.ConnectionString = oPC.DBLocalConn.ConnectionString
    G.DataSource = Me.DC1
    

End Sub

Private Sub cmdLoadDefaults_Click()
    oPC.OpenLocalDatabase
    oPC.DBLocalConn.Execute "EXEC LoadRoundingRules"
    oPC.CloseLocalDatabase
    DC1.Refresh
    'G.ReBind
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
    GetRules
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oPC.CloseLocalDatabase

End Sub
