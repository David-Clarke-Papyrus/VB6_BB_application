VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmServiceBrokerMonitor 
   Caption         =   "Service broker monitor"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkWithCleanup 
      Caption         =   "with cleanup"
      Height          =   270
      Left            =   3480
      TabIndex        =   3
      ToolTipText     =   "Use this only if the other side of the conversation is no longer available"
      Top             =   3480
      Width           =   1290
   End
   Begin VB.CommandButton cmdDeleteConversation 
      Caption         =   "delete conversation for selected row"
      Height          =   330
      Left            =   225
      TabIndex        =   2
      Top             =   3435
      Width           =   3075
   End
   Begin MSAdodcLib.Adodc DC1 
      Height          =   555
      Left            =   5460
      Top             =   3480
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   979
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
   Begin VB.CommandButton cmdTransmissionQueue 
      Caption         =   "Transmission queue"
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   60
      Width           =   1770
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmSBMonitor.frx":0000
      Height          =   2970
      Left            =   240
      OleObjectBlob   =   "frmSBMonitor.frx":0012
      TabIndex        =   0
      Top             =   435
      Width           =   11445
   End
End
Attribute VB_Name = "frmServiceBrokerMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDeleteConversation_Click()
    On Error GoTo errHandler
Dim sConversationHandle As String
Dim oSQL As New z_SQL
Dim sPos As String

    sConversationHandle = DC1.Recordset.Fields(0)
    If MsgBox("You want to delete a conversation for the selected row? - conversation handle : " & sConversationHandle, vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    sPos = "Pos 1"
    Screen.MousePointer = vbHourglass

    oSQL.RemoveSBConversation sConversationHandle
    sPos = "Pos 2"
   ' cmdTransmissionQueue_Click
    sPos = "Pos 3"
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmServiceBrokerMonitor.cmdDeleteConversation_Click", , , , "sPOS", Array(sPos)
    HandleError
End Sub

Private Sub cmdTransmissionQueue_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
Dim s As String


    Set res = New ADODB.Recordset
    Me.DC1.CommandType = adCmdText
    Me.DC1.RecordSource = "SELECT * frOM sys.Transmission_Queue ORDER BY enqueue_Time"
    Me.DC1.ConnectionString = oPC.ConnectionString
    G.DataSource = Me.DC1
   
'    res.CursorLocation = adUseClient
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = cn
'    cmd.CommandText = "SELECT * frOM sys.Transmission_Queue ORDER BY enqueue_Time"
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 360
'    Set res = cmd.Execute
'    Set DC1.Recordset = res
'    G.DataSource = DC1
'    G.ReBind
'    G.Refresh
'    Set cmd = Nothing
    

End Sub
