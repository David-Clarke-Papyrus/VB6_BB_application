VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClientList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Client List"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstClient 
      Height          =   2055
      Left            =   195
      TabIndex        =   2
      Top             =   180
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Client Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3638
      TabIndex        =   1
      Top             =   2340
      Width           =   1665
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1088
      TabIndex        =   0
      Top             =   2340
      Width           =   1665
   End
End
Attribute VB_Name = "frmClientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oPSvs As Z_PollingServices

Public Sub Component(pPSvs As Z_PollingServices)
    On Error GoTo errHandler
    Set oPSvs = pPSvs
    LoadList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClientList.Component(pPSvs)", pPSvs
End Sub

Private Sub LoadList()
    On Error GoTo errHandler
Dim lst As ListItem
Dim i As Integer

    With Me.lstClient
        .ListItems.Clear
        .ColumnHeaders(1).Width = 1800
        .ColumnHeaders(2).Width = .Width - 1800 - 80
        
        For i = 0 To oPSvs.ClientListCount
            Set lst = .ListItems.Add
            lst.Text = oPSvs.ClientName(i)
            lst.SubItems(1) = oPSvs.ClientPathName(i)
        Next i
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClientList.LoadList"
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim i As Integer
    With lstClient
        If MsgBox("Are you sure you want to delete client:" & vbLf & .SelectedItem.Text & _
            vbLf & vbLf & "No undelete available!", vbYesNo + vbExclamation, _
            "Delete Client?") = vbYes Then
            oPSvs.DeleteClient .SelectedItem.Text
            LoadList
        End If
                
    End With
    Me.cmdDelete.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClientList.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClientList.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lstClient_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandler
    Me.cmdDelete.Enabled = Item.Selected
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmClientList.lstClient_ItemClick(Item)", Item, EA_NORERAISE
    HandleError
End Sub
