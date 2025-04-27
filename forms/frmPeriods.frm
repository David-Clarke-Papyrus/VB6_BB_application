VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmPeriods 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Periods"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form2"
   ScaleHeight     =   3435
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmPeriods.frx":0000
      Height          =   2835
      Left            =   150
      OleObjectBlob   =   "frmPeriods.frx":0015
      TabIndex        =   0
      Top             =   120
      Width           =   5355
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   90
      Top             =   2970
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
End
Attribute VB_Name = "frmPeriods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select * FROM tPERIOD ORDER BY PER_DATE"
    Me.Adodc1.ConnectionString = oPC.ConnectionString
    G.DataSource = Me.Adodc1

End Sub
Private Sub Form_Unload(Cancel As Integer)
    G.Update
End Sub

