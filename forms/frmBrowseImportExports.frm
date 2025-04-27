VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowseImportExports 
   Caption         =   "Browse imports and exports"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      CausesValidation=   0   'False
      Default         =   -1  'True
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
      Left            =   120
      Picture         =   "frmBrowseImportExports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2625
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2310
      Left            =   90
      TabIndex        =   0
      Top             =   285
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4075
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Import/Export"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmBrowseImportExports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As adodb.Recordset

Public Sub Component(pRS As adodb.Recordset)
    On Error GoTo errHandler
    Set rs = pRS
    If rs.RecordCount > 0 Then
        LoadList
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmExplainCost.component(pRS)", pRS
End Sub

Private Sub LoadList()
    On Error GoTo errHandler
    
Dim lstItem As ListItem
Dim i As Long



    lvw.ListItems.Clear
    rs.MoveFirst
    Do While Not rs.EOF
        Set lstItem = lvw.ListItems.Add
        lstItem.Text = FNS(rs.Fields(1))
        lstItem.SubItems(1) = FNS(rs.Fields(2))
        lstItem.SubItems(2) = FNN(rs.Fields(0))
        rs.MoveNext
    Loop

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
    lvw.Height = NonNegative_Lng(Me.Height - (lvw.Top + 1320))
    cmdClose.Top = Me.lvw.Height + 350
    cmdClose.Left = lvw.Left

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



Private Sub lvw_DblClick()
Dim f As frmImportExportDetails
Dim rs As adodb.Recordset

    Set rs = New adodb.Recordset
    rs.Open "SELECT * FROM vBrowseAccountingDebtorsExport WHERE FKEY = " & CStr(lvw.SelectedItem.SubItems(2)) & " ORDER BY SignedDate ", oPC.CO, adOpenStatic
    
    Set f = New frmImportExportDetails
    f.Component "Data exported on " & CStr(lvw.SelectedItem.Text), rs
    f.Show

End Sub
