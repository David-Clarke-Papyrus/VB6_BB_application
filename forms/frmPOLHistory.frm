VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmPOLHistory 
   BackColor       =   &H00D3D3CB&
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5970
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOLHistory.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print the invoice"
      Top             =   2460
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   2175
      Left            =   180
      OleObjectBlob   =   "frmPOLHistory.frx":038A
      TabIndex        =   0
      Top             =   180
      Width           =   6795
   End
End
Attribute VB_Name = "frmPOLHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mPOLID As Long
Dim rs As ADODB.Recordset
Dim XA As New XArrayDB

Public Sub component(POLID As Long)
    On Error GoTo errHandler
    mPOLID = POLID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLHistory.component(POLID)", POLID
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Set rs = Nothing
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLHistory.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim par As ADODB.Parameter
Dim cmd As ADODB.Command
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetPOLHistory"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@POLID", adInteger)
    cmd.Parameters.Append par
    par.Value = mPOLID
    
    cmd.ActiveConnection = oPC.COShort
    Set rs = cmd.execute
    
    LoadGrid
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLHistory.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Integer

    XA.Clear
    If ((rs.eof) And (rs.BOF)) Then Exit Sub
    Do While Not rs.eof
        i = i + 1
        XA.ReDim 1, i, 1, 8
        XA.Value(i, 1) = Trim(rs.fields(5))
        XA.Value(i, 2) = Trim(rs.fields(2))
        XA.Value(i, 3) = Trim(rs.fields(3))
        XA.Value(i, 4) = Format(Trim(rs.fields(4)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        XA.Value(i, 5) = PBKSPercentF(CDbl(Trim(rs.fields(6))))
        XA.Value(i, 6) = Format(CDate(Trim(rs.fields(7))), "dd/mm/yyyy HH:NN")
        rs.MoveNext
    Loop
    G.Array = XA
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOLHistory.LoadGrid"
End Sub
