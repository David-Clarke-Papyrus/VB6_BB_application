VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmPropertyTypes 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Property types"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmPropertyTypes.frx":0000
      Height          =   2160
      Left            =   375
      OleObjectBlob   =   "frmPropertyTypes.frx":0015
      TabIndex        =   0
      Top             =   285
      Width           =   3780
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   450
      Top             =   2550
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
Attribute VB_Name = "frmPropertyTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConn As New z_POSConnection

Private Sub Form_Load()


    oConn.dbConnect Forms(0).strLocalServername, Forms(0).strPassword

    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select PROPT_ID,PROPT_DESCRIPTION FROM tPropertyType ORDER BY PROPT_DESCRIPTION"
    Me.Adodc1.ConnectionString = oConn.ConnectionString
    G.DataSource = Me.Adodc1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    G.Update
End Sub

Private Sub G_BeforeDelete(Cancel As Integer)
Dim lngID As Long
Dim iCnt As Integer
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
    lngID = Adodc1.Recordset.Fields(0)
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "ISPropertyTypeIDinUse"
    cmd.CommandType = adCmdStoredProc
'''-------------------------------
''    OpenResult = oPC.OpenDBSHort
'''-------------------------------
    Set par = cmd.CreateParameter("@ID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngID
    Set par = cmd.CreateParameter("@Cnt", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oConn.DBConn
    cmd.Execute
    
    If (cmd.Parameters(1) > 0) Then
        Cancel = True
        MsgBox "You cannot delete this property type as a property still belongs to it.", vbInformation + vbOKOnly, "Cannot do this"
    End If
    Set cmd = Nothing
    Set par = Nothing
'''---------------------------------------------------
''    If OpenResult = 0 Then oPC.DisconnectDBShort
'''---------------------------------------------------
End Sub

Private Sub G_Error(ByVal DataError As Integer, Response As Integer)
Response = 0
End Sub

