VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRawData 
   Caption         =   "Raw data"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00CED0BF&
      Caption         =   "Fetch"
      Height          =   570
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   990
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   1245
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   135
      Width           =   1860
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4470
      Left            =   285
      TabIndex        =   0
      Top             =   870
      Width           =   8955
      _ExtentX        =   15796
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EAN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date added"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Supplier"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   405
      TabIndex        =   2
      Top             =   165
      Width           =   720
   End
End
Attribute VB_Name = "frmRawData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strPID As String
Dim frmD As New frmDeliveries

Private Sub cmdGo_Click()
    Set rs = FetchByTitle(txtTitle)
    LoadListView
End Sub

Private Function FetchByTitle(pTitle As String) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetByTitle"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Description", adVarChar, adParamInput, 200, pTitle)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@ErrorMsg", adVarChar, adParamOutput, 200)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oCnn.Connection
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    'Set rs = cmd.Execute
    'iReturn = Trim(cmd.Parameters(1))
  '  If iReturn > "" Then
  '      MsgBox "error in FetchByTitle"
  '  Else
        Set FetchByTitle = rs
  '  End If
    Set cmd = Nothing
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    CreateNewInvoice = iReturn
    
End Function
Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To rs.RecordCount
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Key = rs.Fields(0)
            .Text = rs.Fields(1)
            .SubItems(1) = rs.Fields("P_DESCRIPTION")
            .SubItems(2) = Format(rs.Fields("P_DateAdded"), "yyyy-mm-dd Hh:Nn")
            .SubItems(3) = rs.Fields("S_NAME")
        End With
        rs.MoveNext
    Next i
End Sub

Private Sub Form_Load()
    Me.Left = 200
    Me.Top = 1000
    Me.Width = 10000
    Me.Height = 6000
End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    
End Sub
'Private Function GetRowType(ColIndex As Integer) As Variant
'    On Error GoTo errHandler
'    Select Case ColIndex
'        Case 1, 2, 3
'            GetRowType = XTYPE_STRING
'        Case Else
'            GetRowType = XTYPE_INTEGER
'    End Select
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.GetRowType(ColIndex)", ColIndex
'End Function

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    lvw.SortKey = ColumnHeader.Position - 1
    lvw.SortOrder = Direction
    '    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
 '   G1.Refresh
    Exit Sub

End Sub

Private Sub lvw_DblClick()
    strPID = lvw.SelectedItem.Key
    If frmD Is Nothing Then
        Set frmD = New frmDeliveries
        frmD.Component strPID
        frmD.Show
    End If
    frmD.Component strPID
    frmD.Show
    

    
End Sub
