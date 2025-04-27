VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMergeProducts 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Merge products"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   11250
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwLose 
      Height          =   1245
      Left            =   150
      TabIndex        =   6
      Top             =   1020
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
      Height          =   390
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3780
      Width           =   1695
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
      Left            =   750
      TabIndex        =   0
      Top             =   450
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
      Top             =   420
      Width           =   2190
   End
   Begin MSComctlLib.ListView lvwKeep 
      Height          =   1245
      Left            =   5610
      TabIndex        =   7
      Top             =   1020
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
      Left            =   5640
      TabIndex        =   5
      Top             =   2310
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
      Left            =   150
      TabIndex        =   4
      Top             =   2280
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Merge this product . . .                                           into . . .              this product"
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
      Left            =   825
      TabIndex        =   3
      Top             =   135
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

Private Sub cmdMerge_Click()
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
        lngResult = oSQL.RunProc("dbo.MergeProducts", Array(FNS(strKeepPID), FNS(strLosePID), oPC.Configuration.DefaultStoreID), "")
        Set oSQL = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        
        MsgBox "The Merge has completed", , "Status"
        Unload Me
    Else
        MsgBox "One or other of the product codes is invalid"
    End If
    Screen.MousePointer = vbDefault
End Sub





Private Sub lvwKeep_Validate(Cancel As Boolean)
Dim i As Integer
    For i = 1 To lvwKeep.ListItems.Count
        If lvwKeep.ListItems(i).Selected = True Then
            strKeepPID = lvwKeep.ListItems(i).SubItems(5)
        End If
    Next
End Sub

Private Sub lvwLose_Validate(Cancel As Boolean)
Dim i As Integer
    For i = 1 To lvwLose.ListItems.Count
        If lvwLose.ListItems(i).Selected = True Then
            strLosePID = lvwLose.ListItems(i).SubItems(5)
        End If
    Next
End Sub

Private Sub txtLose_LostFocus()
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim moProd As a_Product
Dim itm As ListItem

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lvwLose.ListItems.Clear
    If FNS(txtLose) > "" Then
        Set rs = oPC.COSHORT.Execute("EXEC dbo.FINDALLProducts " & FNS(txtLose))
        If Not rs.EOF Then
            strLosePID = FNS(rs.Fields(5))
            Do While Not rs.EOF
                Set itm = lvwLose.ListItems.Add
                itm.Text = FNS(rs.Fields(0))
                itm.SubItems(1) = FNS(rs.Fields(1))
                itm.SubItems(2) = FNS(rs.Fields(2))
                itm.SubItems(3) = FNS(rs.Fields(3))
                itm.SubItems(4) = Format(FNN(rs.Fields(4)) / 100, "###,##0.00")
                itm.SubItems(5) = FNS(rs.Fields(5))
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
End Sub
Private Sub txtKeep_LostFocus()
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
        Set rs = oPC.COSHORT.Execute("EXEC dbo.FINDALLProducts " & FNS(txtKeep))
        If Not rs.EOF Then
            strKeepPID = FNS(rs.Fields(5))
            Do While Not rs.EOF
                Set itm = lvwKeep.ListItems.Add
                itm.Text = FNS(rs.Fields(0))
                itm.SubItems(1) = FNS(rs.Fields(1))
                itm.SubItems(2) = FNS(rs.Fields(2))
                itm.SubItems(3) = FNS(rs.Fields(3))
                itm.SubItems(4) = Format(FNN(rs.Fields(4)) / 100, "###,##0.00")
                itm.SubItems(5) = FNS(rs.Fields(5))
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
End Sub

