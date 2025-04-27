VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeliveries 
   Caption         =   "Deliveries"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvw 
      Height          =   4470
      Left            =   30
      TabIndex        =   0
      Top             =   45
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   7885
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDeliveries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strPID As String
Public Sub Component(pPID As String)
    strPID = pPID
    FetchDeliveriesByTitle
    LoadListView
    
End Sub
Private Function FetchDeliveriesByTitle() As ADODB.Recordset
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetDeliveriesByPID"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@PID", adVarChar, adParamInput, 50, strPID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oCnn.Connection
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    Set FetchDeliveriesByTitle = rs
    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
    
End Function
Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    If rs Is Nothing Then Exit Sub
    For i = 1 To rs.RecordCount
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Text = rs.Fields(1)
            .SubItems(1) = Format(FNCURR(rs.Fields("D_PRICE")) / 100, "currency")
            .SubItems(2) = FNS(rs.Fields("D_SRC"))
        End With
        rs.MoveNext
    Next i
End Sub

Private Sub Form_Load()
    Me.Left = 200
    Me.Top = 1000
End Sub


Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    lvw.SortKey = ColumnHeader.Position - 1
    lvw.SortOrder = Direction
    Exit Sub

End Sub
