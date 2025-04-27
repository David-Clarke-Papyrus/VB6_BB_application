VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCathead 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Catalogue headings"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5025
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Height          =   525
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   915
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
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
      Height          =   4785
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   9360
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3975
         Width           =   885
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1095
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3975
         Width           =   885
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   195
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3975
         Width           =   885
      End
      Begin MSComctlLib.ListView lvwCatHead 
         Height          =   3540
         Left            =   195
         TabIndex        =   2
         Top             =   405
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   6244
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Description"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Sort tag"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Belongs to . . ."
            Object.Width           =   3598
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCathead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCathead As a_cathead
Dim flgLoading As Boolean
Dim bSetOK As Boolean
Dim chexCathead As chex_Cathead

Private Sub oCathead_Valid(pErrors As String, pValid As Boolean)
    On Error GoTo errHandler
    Me.cmdOK.Enabled = pValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.oCathead_Valid(pErrors,pValid)", Array(pErrors, pValid), EA_NORERAISE
    HandleError
End Sub



Private Sub LoadCatheads()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwCatHead.ListItems.Clear
    For i = 1 To chexCathead.Count
        Set objItm = Me.lvwCatHead.ListItems.Add
        With objItm
            .Key = chexCathead(i).Key
            .text = chexCathead(i).Description & IIf(chexCathead(i).IsDeleted, "<Deleted>", "")
            .SubItems(1) = chexCathead(i).SortTag
            .SubItems(2) = chexCathead(i).ParentHeading
        End With
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.LoadCatheads"
End Sub
Private Sub cmdAdd_Click()
    On Error GoTo errHandler
'Dim oCathead As a_cathead
'Dim frm As frmEditCatHead
'
'    Set oCathead = chexCathead.Add
'    Set frm = New frmEditCatHead
'    frm.component oCathead
'    frm.Show vbModal
'    LoadCatheads
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim oCathead As a_cathead
Dim strError As String
    If Not lvwCatHead.SelectedItem Is Nothing Then
        Set oCathead = chexCathead(lvwCatHead.SelectedItem.Key)
        oCathead.BeginEdit
        oCathead.Delete
        oCathead.ApplyEdit
        LoadCatheads
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
'Dim oCathead As a_cathead
'Dim frm As frmEditCatHead
'    If Not lvwCatHead.SelectedItem Is Nothing Then
'        Set oCathead = chexCathead(lvwCatHead.SelectedItem.Key)
'        Set frm = New frmEditCatHead
'        frm.component oCathead
'        frm.Show vbModal
'        LoadCatheads
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If chexCathead.IsEditing Then chexCathead.CancelEdit
    Set chexCathead = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCatHead_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.lvwCatHead_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCatHead_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.lvwCatHead_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.mnuClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadSAs()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.LoadSAs"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    If Me.WindowState <> 2 Then
        Me.Width = 9900
        Me.Height = 6200
        Me.TOP = 100
        Me.Left = 100
    End If
    Set chexCathead = New chex_Cathead
    chexCathead.BeginEdit
    LoadControls
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    GetCatheads
    LoadCatheads
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.LoadControls"
End Sub
Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim strError As String
    chexCathead.ApplyEdit strError
    If strError > "" Then
        MsgBox strError, , "Can't save"
        chexCathead.BeginEdit
    Else
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    chexCathead.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub EnableOK(flgValid As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = flgValid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.EnableOK(flgValid)", flgValid
End Sub


Private Sub GetCatheads()
    On Error GoTo errHandler
    chexCathead.CancelEdit
    Set chexCathead = Nothing
    Set chexCathead = New chex_Cathead
    chexCathead.Load
    chexCathead.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.GetCatheads"
End Sub

Private Sub lvwCatHead_DblClick()
    On Error GoTo errHandler
    cmdEdit_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCathead.lvwCatHead_DblClick", , EA_NORERAISE
    HandleError
End Sub
