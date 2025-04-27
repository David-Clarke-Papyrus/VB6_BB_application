VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMergeProducts 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge products"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   11250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFindKeep 
      Height          =   345
      Left            =   8475
      Picture         =   "frmMergeProducts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   570
      Width           =   375
   End
   Begin VB.CommandButton cmdFindLose 
      Height          =   345
      Left            =   3990
      Picture         =   "frmMergeProducts.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   615
      Width           =   375
   End
   Begin MSComctlLib.ListView lvwLose 
      Height          =   1245
      Left            =   150
      TabIndex        =   6
      Top             =   1155
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   2196
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ISBN-13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "S.P."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "PID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdMerge 
      BackColor       =   &H00C4BCA4&
      Caption         =   "MERGE"
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
      Left            =   5100
      Picture         =   "frmMergeProducts.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2835
      Width           =   1000
   End
   Begin VB.TextBox txtLose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1755
      TabIndex        =   0
      Top             =   585
      Width           =   2190
   End
   Begin VB.TextBox txtKeep 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   555
      Width           =   2190
   End
   Begin MSComctlLib.ListView lvwKeep 
      Height          =   1245
      Left            =   5610
      TabIndex        =   7
      Top             =   1155
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   2196
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ISBN-13"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "S.P."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "PID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EAN or ISBN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   5010
      TabIndex        =   9
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label ISBN 
      BackStyle       =   0  'Transparent
      Caption         =   "EAN or ISBN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   540
      TabIndex        =   8
      Top             =   630
      Width           =   1155
   End
   Begin VB.Label lblTo 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1230
      Left            =   6705
      TabIndex        =   5
      Top             =   2430
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFrom 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1170
      Left            =   945
      TabIndex        =   4
      Top             =   2430
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this product . . .             into . . .                       this product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   1920
      TabIndex        =   3
      Top             =   90
      Width           =   10005
   End
End
Attribute VB_Name = "frmMergeProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strLosePID As String
Dim strKeepPID As String
Dim bLoseOK As Boolean
Dim bKeepOK As Boolean

Private Sub cmdFindKeep_Click()
    On Error GoTo errHandler
Dim frm As New frmQuickProductFind
Dim strCode As String
    strCode = txtKeep
    frm.component strCode
    
    frm.Show vbModal
    If frm.QtyQuickFound = 0 Then
        MsgBox "Nothing found", vbInformation, "Status"
    End If
    If frm.Cancelled = False Then
        If frm.EAN > "" Then txtKeep = frm.EAN
    End If
    txtKeep.SetFocus
    Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.cmdFindKeep_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFindLose_Click()
    On Error GoTo errHandler
Dim frm As New frmQuickProductFind
Dim strCode As String
    strCode = txtLose
    frm.component strCode
    frm.Show vbModal
    If frm.QtyQuickFound = 0 Then
        MsgBox "Nothing found", vbInformation, "Status"
    End If
    If frm.Cancelled = False Then
        If frm.EAN > "" Then txtLose = frm.EAN
    End If
    txtLose.SetFocus
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.cmdFindLose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMerge_Click()
    On Error GoTo errHandler
Dim oSQL As z_SQL
Dim lngResult As Long
Dim OpenResult As Integer

    Screen.MousePointer = vbHourglass
    If strLosePID > "" And strKeepPID > "" Then
        If strLosePID = strKeepPID Then
            MsgBox "You have selected the same item to keep and to replace.", vbOKOnly + vbInformation, "Can't do this"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Set oSQL = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        lngResult = oSQL.RunProc("dbo.MergeProducts", Array(FNS(strKeepPID), FNS(strLosePID), oPC.Configuration.DefaultStoreID, gSTAFFID), "")
        Set oSQL = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        Me.txtKeep = ""
        Me.txtLose = ""
        Me.lvwKeep.ListItems.Clear
        Me.lvwLose.ListItems.Clear
        MsgBox "The Merge has completed", , "Status"
    Else
        MsgBox "One or other of the product codes is invalid"
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.cmdMerge_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub lvwKeep_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To lvwKeep.ListItems.Count
        If lvwKeep.ListItems(i).Selected = True Then
            strKeepPID = lvwKeep.ListItems(i).SubItems(5)
        End If
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.lvwKeep_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLose_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To lvwLose.ListItems.Count
        If lvwLose.ListItems(i).Selected = True Then
            strLosePID = lvwLose.ListItems(i).SubItems(5)
        End If
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.lvwLose_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtLose_LostFocus()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim moProd As a_Product
Dim itm As ListItem
Dim str As String
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lvwLose.ListItems.Clear
    If FNS(txtLose) > "" Then
        str = Replace(txtLose, "'", "''")
        Set rs = oPC.COShort.execute("EXEC dbo.FINDALLProducts '" & FNS(Left(str, 20)) & "'")
        If Not rs.eof Then
            strLosePID = FNS(rs.fields(5))
            Do While Not rs.eof
                Set itm = lvwLose.ListItems.Add
                itm.text = FNS(rs.fields(0))
                itm.SubItems(1) = FNS(rs.fields(1))
                itm.SubItems(2) = FNS(rs.fields(2))
                itm.SubItems(3) = FNS(rs.fields(3))
                itm.SubItems(4) = Format(FNN(rs.fields(4)) / oPC.Configuration.DefaultCurrency.Divisor, "###,##0.00")
                itm.SubItems(5) = FNS(rs.fields(5))
                rs.MoveNext
            Loop
            lvwLose.ListItems(1).Selected = True
            bLoseOK = True
        Else
            strLosePID = ""
            bLoseOK = False
        End If
    End If
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    cmdMerge.Enabled = bLoseOK And bKeepOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.txtLose_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtKeep_LostFocus()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim moProd As a_Product
Dim itm As ListItem
'    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lvwKeep.ListItems.Clear
    If FNS(txtKeep) > "" Then
        Set rs = oPC.COShort.execute("EXEC dbo.FINDALLProducts '" & FNS(Left(txtKeep, 20)) & "'")
        If Not rs.eof Then
            strKeepPID = FNS(rs.fields(5))
            Do While Not rs.eof
                Set itm = lvwKeep.ListItems.Add
                itm.text = FNS(rs.fields(0))
                itm.SubItems(1) = FNS(rs.fields(1))
                itm.SubItems(2) = FNS(rs.fields(2))
                itm.SubItems(3) = FNS(rs.fields(3))
                itm.SubItems(4) = Format(FNN(rs.fields(4)) / oPC.Configuration.DefaultCurrency.Divisor, "###,##0.00")
                itm.SubItems(5) = FNS(rs.fields(5))
                rs.MoveNext
            Loop
            lvwKeep.ListItems(1).Selected = True
            bKeepOK = True
        Else
            strKeepPID = ""
            bKeepOK = False
        End If
    End If
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    cmdMerge.Enabled = bLoseOK And bKeepOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMergeProducts.txtKeep_LostFocus", , EA_NORERAISE
    HandleError
End Sub

