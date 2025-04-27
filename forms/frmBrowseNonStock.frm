VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBrowseServiceItem 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse service items"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   ControlBox      =   0   'False
   Icon            =   "frmBrowseNonStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   6015
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   4875
      Picture         =   "frmBrowseNonStock.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3225
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Current service items"
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
      Height          =   3060
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   5805
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1260
         Picture         =   "frmBrowseNonStock.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2340
         Width           =   1000
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   255
         Picture         =   "frmBrowseNonStock.frx":0C9E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2340
         Width           =   1000
      End
      Begin MSComctlLib.ListView lvwServiceItem 
         Height          =   1905
         Left            =   240
         TabIndex        =   1
         Top             =   390
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   3360
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   6068
         EndProperty
      End
   End
End
Attribute VB_Name = "frmBrowseServiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cNS As c_ServiceItem

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim oProd As a_Product
    Set oProd = New a_Product
    If lvwServiceItem.ListItems.Count = 0 Then Exit Sub
    
    oProd.Load Me.lvwServiceItem.SelectedItem.Key, 0
    oProd.BeginEdit
    oProd.Delete
    oProd.ApplyEdit
    Set oProd = Nothing
    GetRecs
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim oProd As a_Product
Dim frm As frmServiceItem
    If lvwServiceItem.ListItems.Count = 0 Then Exit Sub
    Set frm = New frmServiceItem
    Set oProd = New a_Product
    oProd.Load Me.lvwServiceItem.SelectedItem.Key, 0
    frm.component oProd
    frm.Show
    Set oProd = Nothing
    Set frm = Nothing
    GetRecs
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.SetMenu"
End Sub
Private Sub UnsetMenu()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.UnsetMenu"
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    If Me.WindowState <> 2 Then
        TOP = 50
        Left = 50
        Height = 4400
        Width = 6400
    End If
    GetRecs
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i, j As Integer

    Me.lvwServiceItem.ListItems.Clear
    For i = 1 To cNS.Count
        Set objItm = Me.lvwServiceItem.ListItems.Add
        With objItm
            .Key = cNS(i).PID
            .text = cNS(i).code
            .SubItems(1) = cNS(i).Description
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.LoadListView"
End Sub
Private Sub GetRecs()
    On Error GoTo errHandler
    Set cNS = Nothing
    Set cNS = New c_ServiceItem
    cNS.Load
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.GetRecs"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwServiceItem_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.lvwServiceItem_AfterLabelEdit(Cancel,NewString)", Array(Cancel, _
         NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwServiceItem_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.lvwServiceItem_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwServiceItem_DblClick()
    On Error GoTo errHandler
Dim oProd As a_Product
Dim frm As frmServiceItem
    If lvwServiceItem.ListItems.Count = 0 Then Exit Sub
    Set frm = New frmServiceItem
    Set oProd = New a_Product
    oProd.Load Me.lvwServiceItem.SelectedItem.Key, 0
    frm.component oProd
    frm.Show
    Set oProd = Nothing
    Set frm = Nothing
    GetRecs
    LoadListView

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseServiceItem.lvwServiceItem_DblClick", , EA_NORERAISE
    HandleError
End Sub
