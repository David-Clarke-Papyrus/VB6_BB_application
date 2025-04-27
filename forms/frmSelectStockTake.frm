VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmSelectStocktake 
   Caption         =   "Re run movements since stocktake"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkResetCosts 
      Alignment       =   1  'Right Justify
      Caption         =   "Reset any costs to zero before re-running"
      ForeColor       =   &H8000000D&
      Height          =   420
      Left            =   1380
      TabIndex        =   4
      Top             =   2955
      Width           =   1965
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   645
      Left            =   225
      Picture         =   "frmSelectStockTake.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2955
      Width           =   840
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   3345
      Picture         =   "frmSelectStockTake.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2925
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2340
      Left            =   225
      OleObjectBlob   =   "frmSelectStockTake.frx":0714
      TabIndex        =   0
      Top             =   480
      Width           =   4140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-run movements since stocktake dated . . ."
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   300
      TabIndex        =   3
      Top             =   225
      Width           =   3240
   End
End
Attribute VB_Name = "frmSelectStocktake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As New XArrayDB
Dim rs As New ADODB.Recordset


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    On Error GoTo errHandler
Dim lngStkTkeID As Long
Dim oSQL As z_SQL
Dim Res As Long

    If IsNull(G1.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    lngStkTkeID = FNN(x(G1.Bookmark, 2))
    If lngStkTkeID > 0 Then
        Set oSQL = New z_SQL
        Res = oSQL.ReRunTransactions(lngStkTkeID, Now(), (chkResetCosts = 1))
        If Res <> 0 Then
            Err.Raise vbObjectError + 101, "RerunTransactions", "Stored procedure failed"
        End If
    End If
    Screen.MousePointer = vbDefault
    Unload Me
    MsgBox "Re-run of transactions is complete.", vbInformation + vbOKOnly, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSelectStocktake.cmdGo_Click"
End Sub

Private Sub Form_Load()
    LoadArray
    G1.Array = x
    G1.ReBind
    G1.Refresh
    
End Sub

Private Sub LoadArray()
Dim i As Integer

    oPC.OpenDBSHort
    rs.CursorLocation = adUseClient
    rs.Open "SELECT STKTKE_CUTOFFDATE,TR_ID FROM tSTKTKE JOIN tTR ON STKTKE_ID = TR_ID WHERE TR_STATUS IN (3,4) ORDER BY STKTKE_CUTOFFDATE DESC", oPC.COShort
    x.ReDim 1, rs.RecordCount, 1, 3
    i = 0
    Do While Not rs.EOF
        i = i + 1
        x(i, 1) = Format(FND(rs.Fields("STKTKE_CUTOFFDATE")), "DD-MM-YYYY HH:NN")
        x(i, 2) = FNN(rs.Fields("TR_ID"))
        rs.MoveNext
    Loop
End Sub
