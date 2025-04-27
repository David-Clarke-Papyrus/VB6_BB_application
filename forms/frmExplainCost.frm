VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplainCost 
   BackColor       =   &H00EBEBEB&
   Caption         =   "Cost history"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Close"
      Height          =   315
      Left            =   4665
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1425
      Width           =   525
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1260
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   2223
      SortKey         =   8
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "date"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Orig. qty"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Orig cost"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "New qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "New price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Extra ch. alloc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "New avg cost"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "SortedDate"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmExplainCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub component(pRs As ADODB.Recordset)
    On Error GoTo errHandler
    Set rs = pRs
    If rs.RecordCount > 0 Then
        LoadList
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.component(pRS)", pRs
End Sub

Private Sub LoadList()
    On Error GoTo errHandler
    
Dim lstItem As ListItem
Dim i As Long



    lvw.ListItems.Clear
    rs.MoveFirst
    Do While Not rs.eof
        Set lstItem = lvw.ListItems.Add
        lstItem.text = FNS(rs.fields(0))
      '  lstItem.key = FNS(rs.Fields(0))
        lstItem.SubItems(1) = FNS(rs.fields(1))
        lstItem.SubItems(2) = Format(FNDBL(rs.fields(2)), "###,##0")
        lstItem.SubItems(3) = Format(FNDBL(rs.fields(3)), "###,##0.00")
        lstItem.SubItems(4) = Format(FNDBL(rs.fields(4)), "###,##0")
        lstItem.SubItems(5) = Format(FNDBL(rs.fields(5)), "###,##0.00")
        lstItem.SubItems(6) = Format(FNDBL(rs.fields(6)), "###,##0.00")
        lstItem.SubItems(7) = Format(FNDBL(rs.fields(7)), "###,##0.00")
        lstItem.SubItems(8) = ReverseDateTime(FND(rs.fields(0)))
        rs.MoveNext
    Loop
    If lvw.ListItems.Count > 1 Then
        Set lvw.SelectedItem = lvw.ListItems(lvw.ListItems.Count)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.LoadList"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    SetLvwLayout Me.lvw, Me.Name
    SetFormSize Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    lvw.Width = NonNegative_Lng(Me.Width - (lvw.Left + 200))
    lvw.Height = NonNegative_Lng(Me.Height - (lvw.TOP + 820))
    cmdclose.TOP = Me.lvw.Height + 150
    cmdclose.Left = lvw.Width / 2

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    SaveLayoutLvw lvw, Me.Name, Me.Height, Me.Width

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

