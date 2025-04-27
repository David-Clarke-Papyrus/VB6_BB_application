VERSION 5.00
Begin VB.Form frmSTAT 
   BackColor       =   &H00D3D3CB&
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   14265
   Begin VB.CommandButton cmdGrid2toSpreadsheet 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Exp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11970
      Picture         =   "frmSTAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13005
      Picture         =   "frmSTAT.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5625
      Width           =   1000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transactions per day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   6720
      Width           =   2205
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Snapshot view"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   255
      TabIndex        =   1
      Top             =   6840
      Width           =   1455
   End
End
Attribute VB_Name = "frmSTAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim rs As ADODB.Recordset
'Dim OpenResult As Integer
'Dim mdteFrom As Date
'Dim mdteTo As Date
'Dim X As New XArrayDB
'Dim XX As New XArrayDB
'
'Public Sub Component(pdteFrom As Date, pdteTo As Date)
'    mdteFrom = pdteFrom
'    mdteTo = pdteTo
'End Sub
'
'Private Sub cmdClose_Click()
'Unload Me
'End Sub
'
'Private Sub cmdDefaultLayout_Click()
'SetDefaultWidths
'End Sub
'
'
'Private Sub cmdSaveLayout_Click()
'    SaveLayout G, Me.Name & "G"
'    SaveLayout GG, Me.Name & "GG"
'End Sub
'
'
'
'Private Sub LoadGrids()
'    On Error GoTo errHandler
'Dim i As Integer
'Dim iRow As Integer
'Dim iCol As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    rs.Open "Select * FROM  tSTAT WHERE STAT_DATE BETWEEN '" & ReverseDate(mdteFrom) & " ' AND '" & ReverseDate(mdteTo) & "' ORDER BY STAT_DATE DESC", oPC.COShort, adOpenStatic, adLockReadOnly
'    X.Clear
'    X.ReDim 1, rs.RecordCount, 1, 14
'    iRow = 1
'    If Not rs.EOF Then
'        Do While Not rs.EOF
'                X(iRow, 1) = Trim(rs.Fields("STAT_DATE"))
'                X(iRow, 2) = Trim(rs.Fields("STAT_VOS_Retail"))
'                X(iRow, 3) = Trim(rs.Fields("STAT_VOS_Cost"))
'                X(iRow, 4) = Trim(rs.Fields("STAT_OnHand_Qtyproducts"))
'                X(iRow, 5) = Trim(rs.Fields("STAT_OnHand_QtyItems"))
'                X(iRow, 6) = Trim(rs.Fields("STAT_OOS_QtyItems"))
'                X(iRow, 7) = Trim(rs.Fields("STAT_OOS_Value_Retail"))
'                X(iRow, 8) = Trim(rs.Fields("STAT_OOS_Value_Cost"))
'                X(iRow, 9) = Trim(rs.Fields("STAT_COOS_QtyItems"))
'                X(iRow, 10) = Trim(rs.Fields("STAT_COOS_Value_Retail"))
'                X(iRow, 11) = Trim(rs.Fields("STAT_COOS_Value_Cost"))
'                X(iRow, 12) = Trim(rs.Fields("STAT_Appros_QtyItems"))
'                X(iRow, 13) = Trim(rs.Fields("STAT_Appros_Value_Retail"))
'                X(iRow, 14) = Trim(rs.Fields("STAT_Appros_Value_Cost"))
'
'            iRow = iRow + 1
'            rs.MoveNext
'        Loop
'        rs.MoveFirst
'        G.Array = X
'    End If
'    XX.Clear
'    XX.ReDim 1, rs.RecordCount, 1, 22
'    iRow = 1
'    If Not rs.EOF Then
'        Do While Not rs.EOF
'                XX(iRow, 1) = Trim(rs.Fields("STAT_DATE"))
'                XX(iRow, 2) = Trim(rs.Fields("STAT_DEL_QtyItems"))
'                XX(iRow, 3) = Trim(rs.Fields("STAT_DEL_Value_Retail"))
'                XX(iRow, 4) = Trim(rs.Fields("STAT_DEL_Value_Cost"))
'                XX(iRow, 5) = Trim(rs.Fields("STAT_INV_QtyItems"))
'                XX(iRow, 6) = Trim(rs.Fields("STAT_INV_Value_Retail"))
'                XX(iRow, 7) = Trim(rs.Fields("STAT_INV_Value_Cost"))
'                XX(iRow, 8) = Trim(rs.Fields("STAT_CS_QtyItems"))
'                XX(iRow, 9) = Trim(rs.Fields("STAT_CS_Value_Retail"))
'                XX(iRow, 10) = Trim(rs.Fields("STAT_CS_Value_Cost"))
'                XX(iRow, 11) = Trim(rs.Fields("STAT_PO_QtyItems"))
'                XX(iRow, 12) = Trim(rs.Fields("STAT_PO_Value_Retail"))
'                XX(iRow, 13) = Trim(rs.Fields("STAT_PO_Value_Cost"))
'                XX(iRow, 14) = Trim(rs.Fields("STAT_CO_QtyItems"))
'                XX(iRow, 15) = Trim(rs.Fields("STAT_CO_Value_Retail"))
'                XX(iRow, 16) = Trim(rs.Fields("STAT_CO_Value_Cost"))
'                XX(iRow, 17) = Trim(rs.Fields("STAT_TFRIN_QtyItems"))
'                XX(iRow, 18) = Trim(rs.Fields("STAT_TFRIN_Value_Retail"))
'                XX(iRow, 19) = Trim(rs.Fields("STAT_TFRIN_Value_Cost"))
'                XX(iRow, 20) = Trim(rs.Fields("STAT_TFROUT_QtyItems"))
'                XX(iRow, 21) = Trim(rs.Fields("STAT_TFROUT_Value_Retail"))
'                XX(iRow, 22) = Trim(rs.Fields("STAT_TFROUT_Value_Cost"))
'
'            iRow = iRow + 1
'            rs.MoveNext
'        Loop
'        GG.Array = XX
'    End If
'    If Me.WindowState <> 2 Then
'        Me.Height = 6000
'        Me.Width = 10000
'        Me.left = 500
'        Me.top = 1000
'    End If
'    For i = 1 To G.Columns.Count
'        G.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "G", CStr(i), 500)
'    Next
'    For i = 1 To GG.Columns.Count
'        GG.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "GG", CStr(i), 500)
'    Next
'    G.ExtendRightColumn = False
'    G.Width = 9500
'    GG.ExtendRightColumn = False
'    GG.Width = 9500
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSTAT.Form_Load"
'    HandleError
'End Sub
'Private Sub SetDefaultWidths()
'Dim i As Integer
'    For i = 1 To G.Columns.Count
'        G.Columns(i - 1).Width = 500
'    Next
'    For i = 1 To GG.Columns.Count
'        GG.Columns(i - 1).Width = 500
'    Next
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    rs.Close
'    Set rs = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'End Sub
Private Sub cmdGrid2toSpreadsheet_Click()

End Sub
