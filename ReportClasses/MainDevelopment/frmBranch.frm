VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmBranch 
   Caption         =   "Edit branches"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid60.TDBGrid G 
      Height          =   1860
      Left            =   210
      OleObjectBlob   =   "frmBranch.frx":0000
      TabIndex        =   0
      Top             =   555
      Width           =   8760
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim x As New XArrayDB



Private Sub Form_Load()
    LoadGrid
End Sub

Private Sub LoadGrid()
Dim i As Integer

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tBRANCH", oPC.COShort, adOpenDynamic, adLockOptimistic
'    X.ReDim 1, rs.RecordCount, 1, 5
'    i = 1
'    Do While Not rs.EOF
'        X(i, 1) = FNS(rs.Fields("BR_CODE"))
'        X(i, 2) = FNS(rs.Fields("BR_Name"))
'        X(i, 3) = FNS(rs.Fields("BR_Contact"))
'        X(i, 4) = FNS(rs.Fields("BR_Email"))
'    Loop
    Set G.DataSource = rs
End Sub
