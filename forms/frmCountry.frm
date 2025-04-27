VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmCountry 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Countries"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4410
   Begin TrueOleDBGrid60.TDBGrid G 
      Bindings        =   "frmCountry.frx":0000
      Height          =   2160
      Left            =   240
      OleObjectBlob   =   "frmCountry.frx":0015
      TabIndex        =   0
      Top             =   210
      Width           =   3780
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   315
      Top             =   2475
      Visible         =   0   'False
      Width           =   2025
      _ExtentX        =   3572
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
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset


Private Sub Form_Load()
    On Error GoTo errHandler

    Me.Adodc1.CommandType = adCmdText
    Me.Adodc1.RecordSource = "Select CTR_ID,CTR_Name FROM tCountry ORDER BY CTR_NAME"
    Me.Adodc1.ConnectionString = oPC.ConnectionString
    G.DataSource = Me.Adodc1
    Me.Width = 4350
    Me.Height = 3800
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCountry.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    G.Update
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCountry.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G_BeforeDelete(Cancel As Integer)
    On Error GoTo errHandler
Dim lngID As Long
Dim iCnt As Integer
Dim cmd As New ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
    lngID = Adodc1.Recordset.fields(0)
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "ISCountryIDinUse"
    cmd.CommandType = adCmdStoredProc
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set par = cmd.CreateParameter("@CNTRYID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngID
    Set par = cmd.CreateParameter("@Cnt", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    If (cmd.Parameters(1) > 0) Then
        Cancel = True
        MsgBox "You cannot delete this country as an address is using it.", vbInformation + vbOKOnly, "Cannot do this"
    End If
    Set cmd = Nothing
    Set par = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCountry.G_BeforeDelete(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G_Error(ByVal DataError As Integer, Response As Integer)
    On Error GoTo errHandler
Response = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCountry.G_Error(DataError,Response)", Array(DataError, Response), EA_NORERAISE
    HandleError
End Sub
