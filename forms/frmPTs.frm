VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPTs 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product types"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   5220
   Begin VB.CommandButton cmdDefault 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Set as default"
      Height          =   345
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1710
      Visible         =   0   'False
      Width           =   1470
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
      Left            =   3360
      Picture         =   "frmPTs.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3900
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4170
      Left            =   255
      TabIndex        =   2
      Top             =   345
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   7355
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.CommandButton cmdADD 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Add"
      Default         =   -1  'True
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
      Left            =   3420
      Picture         =   "frmPTs.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   375
      Width           =   1000
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
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
      Left            =   3420
      Picture         =   "frmPTs.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   990
      Width           =   1000
   End
End
Attribute VB_Name = "frmPTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oPT As a_PT
Attribute oPT.VB_VarHelpID = -1
Dim tlProductTypes As z_TextList


Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdDefault_Click()
' Dim pors As String
'    oPC.Configuration.BeginEdit
'   ' oPC.Configuration.DefaultPT = lvw.SelectedItem.key
'    oPC.Configuration.ApplyEdit pors
'    LoadList
'
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    oPC.Configuration.RefreshProductTypes
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Lvw_DblClick()
    On Error GoTo errHandler
Dim frm As frmPT
    If lvw.SelectedItem.Index < 1 Then Exit Sub
    Set oPT = New a_PT
    oPT.Load tlProductTypes.Key(lvw.SelectedItem.text)
    If oPT.PTID > 0 Then
    
        Set frm = New frmPT
        frm.component oPT
        frm.Show vbModal
        Set oPT = Nothing
        Set frm = Nothing
        LoadListView
    Else
        MsgBox "You cannot edit this item.", vbInformation, "Status"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.Lvw_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub oPT_DBActionSTatus(Status As Integer)
    On Error GoTo errHandler
    Select Case Status
    Case 22
        MsgBox "Addition of product type failed, it would have created a duplicate value"
    End Select
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.oPT_DBActionSTatus(Status)", Status, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdAdd_Click()
    On Error GoTo errHandler
Dim frm As frmPT
    Set oPT = New a_PT
    Set frm = New frmPT
    frm.component oPT
    frm.Show vbModal
    Set oPT = Nothing
    Set frm = Nothing
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.cmdAdd_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oPT_Inuse()
    On Error GoTo errHandler
    MsgBox "This product type is being used. You should merge it with another product type " & vbCrLf _
    & "rather than deleting it." & vbCrLf _
    & "See under Tools/Utilities/Merge two product types"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.oPT_Inuse", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmPT
    Set oPT = New a_PT
    oPT.Load tlProductTypes.Key(lvw.SelectedItem.text)
'    oPT.Load lvwLines.SelectedItem.Key
    If oPT.PTID > 0 Then
    
        Set frm = New frmPT
        frm.component oPT
        frm.Show vbModal
        Set oPT = Nothing
        Set frm = Nothing
        LoadListView
    Else
        MsgBox "You cannot edit this item.", vbInformation, "Status"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        TOP = 500
        Left = 250
        Width = 4800
        Height = 5400
    End If
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltProductType
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
    Set tlProductTypes = Nothing
    Set tlProductTypes = New z_TextList
    tlProductTypes.Load ltProductType
    LoadList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.LoadListView"
End Sub
Private Sub LoadList()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To tlProductTypes.Count
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .text = tlProductTypes.ItemByOrdinalIndex(i)
           ' .Bold = tlProductTypes.f4ByOrdinalIndex(i) = "True"
           ' .key = tlProductTypes.key(i)
        End With
    Next i
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPTs.LoadList"
End Sub

