VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSUPPDEAL 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Deals"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelect 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select"
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
      Left            =   1470
      Picture         =   "frmSUPPDEAL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvwDeals 
      Height          =   1845
      Left            =   75
      TabIndex        =   1
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   3254
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
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
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Discount"
         Object.Width           =   1605
      EndProperty
   End
End
Attribute VB_Name = "frmSUPPDEAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tlDeals As z_TextList
Dim oSupp As a_Supplier
Dim oDeal As a_Deal
Public Sub component(pSupp As a_Supplier)
    On Error GoTo errHandler
    Set oSupp = pSupp
    LoadDeals
    Me.Caption = oSupp.NameAndCode(35)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSUPPDEAL.component(pSupp)", pSupp
End Sub
Private Sub cmdSelect_Click()
    On Error GoTo errHandler
    If Not (lvwDeals.SelectedItem Is Nothing) Then
        Set oDeal = oSupp.Deals(val(lvwDeals.SelectedItem.Key))
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSUPPDEAL.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadDeals()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwDeals.ListItems.Clear
    For i = 1 To oSupp.Deals.Count
        Set objItm = Me.lvwDeals.ListItems.Add
        With objItm
            .Key = i & "k" 'oSupp.Addresses(i).ID & "K"
            .text = oSupp.Deals(i).Description
            .SubItems(1) = oSupp.Deals(i).DiscountF
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSUPPDEAL.LoadDeals"
End Sub

Public Property Get SelectedDeal() As a_Deal
    Set SelectedDeal = oDeal
End Property


Private Sub lvwDeals_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSUPPDEAL.lvwDeals_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDeals_DblClick()
    On Error GoTo errHandler
    Set oDeal = oSupp.Deals(val(lvwDeals.SelectedItem.Key))
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSUPPDEAL.lvwDeals_DblClick", , EA_NORERAISE
    HandleError
End Sub
