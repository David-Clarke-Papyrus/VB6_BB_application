VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScratch 
   Caption         =   "Papyrus clipboard"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearRefs 
      BackColor       =   &H00F2E0D9&
      Caption         =   "Clear refs"
      Height          =   555
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3615
      Width           =   525
   End
   Begin VB.CommandButton cmdExcelExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Excel"
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
      Left            =   105
      Picture         =   "frmScratch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3585
      Width           =   1000
   End
   Begin VB.CommandButton cmdDeselectAll 
      BackColor       =   &H00F2E0D9&
      Caption         =   "£"
      BeginProperty Font 
         Name            =   "Wingdings 2"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3615
      Width           =   390
   End
   Begin VB.CommandButton cmdSelectAll 
      BackColor       =   &H00F2E0D9&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3615
      Width           =   390
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   630
      Left            =   6030
      Picture         =   "frmScratch.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3585
      Width           =   975
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3465
      Left            =   75
      TabIndex        =   1
      Top             =   105
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   6112
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty SOR"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Discount"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ref"
         Object.Width           =   3598
      EndProperty
   End
   Begin VB.CommandButton cmdNewOrd 
      BackColor       =   &H00C4BCA4&
      Caption         =   "OK"
      Height          =   630
      Left            =   7065
      Picture         =   "frmScratch.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3585
      Width           =   975
   End
   Begin VB.Label lblCount 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   315
      Left            =   3330
      TabIndex        =   6
      Top             =   3630
      Width           =   2655
   End
End
Attribute VB_Name = "frmScratch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub component(pRs As ADODB.Recordset)
    On Error GoTo errHandler
Dim i As Integer
Dim tmp As String
Dim objItm As ListItem
    Set rs = pRs
    rs.MoveFirst
    lvw.ListItems.Clear
    For i = 1 To rs.RecordCount
        Set objItm = lvw.ListItems.Add
        With objItm
            .Key = rs.fields(0)
            .text = rs.fields(8)
            .SubItems(1) = rs.fields(11)
            .SubItems(2) = IIf(oPC.AllowsSSInvoicing, rs.fields(4), rs.fields(3))
            If oPC.AllowsSSInvoicing = False Then
                lvw.ColumnHeaders(4).Width = 0
            End If
            .SubItems(3) = rs.fields(5)
            .SubItems(4) = Format(CDbl(rs.fields(6)) / 100, "###,##0.00")
            .SubItems(5) = FormatPercent(CDbl(rs.fields(7)) / 100, 2)
            .SubItems(6) = FNS(rs.fields("REF"))
            .Checked = True
            rs.MoveNext
        End With
    Next i
    lblCount.Caption = CStr(rs.RecordCount) & " records"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.component(pRS)", pRs
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("This will cancel the selections you have made, please confirm.", vbInformation + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    Else
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClearRefs_Click()
    On Error GoTo errHandler
Dim i As Integer
    
    For i = 1 To lvw.ListItems.Count
            lvw.ListItems(6).text = ""
    Next i


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.cmdClearRefs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDeselectAll_Click()
    On Error GoTo errHandler
Dim i As Integer
    
    For i = 1 To lvw.ListItems.Count
            lvw.ListItems(i).Checked = False
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.cmdDeselectAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExcelExport_Click()
On Error GoTo errHandler
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arClipboard_ForExcel
Dim i As Long
Dim strExecutable As String

    If rs Is Nothing Then Exit Sub
    If Not rs.BOF Then rs.MoveFirst
    If rs.eof Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    rpt.component rs, "Papyrus clipboard " & Format(Now(), "dd/mm/yyyy Hh:Nn")
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "\TEMP"
    End If
    sFile = oPC.LocalFolder & "\TEMP\PapyrusClipboard.XLS"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    xls.FileName = sFile
    If rpt.Pages.Count > 0 Then
        xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
          If strExecutable = "" Then
              MsgBox "There is no application set on this computer to open the file: " & sFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
          Else
            F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus, 10000
          End If
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.cmdExcelExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNewOrd_Click()
    On Error GoTo errHandler
    If MsgBox("This will keep just the rows you have ticked in the clipboard, please confirm.", vbInformation + vbOKCancel, "Warning") = vbCancel Then
        Exit Sub
    Else
        UpdateClipboard
        Unload Me
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.cmdNewOrd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectAll_Click()
    On Error GoTo errHandler
Dim i As Integer
    
    For i = 1 To lvw.ListItems.Count
            lvw.ListItems(i).Checked = True
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.cmdSelectAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.lvw_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandler
    Item.Selected = Not Item.Selected
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.lvw_ItemCheck(Item)", Item, EA_NORERAISE
    HandleError
End Sub

Private Sub UpdateClipboard()
    On Error GoTo errHandler
Dim i As Integer
    rs.MoveLast
    Do While Not rs.BOF
        If lvw.ListItems.Item(CStr(rs.fields(0))).Checked = False Then
            rs.Delete
        End If
        rs.MovePrevious
    Loop
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmScratch.UpdateClipboard"
End Sub
