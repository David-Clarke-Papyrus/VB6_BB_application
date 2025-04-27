VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBICTree 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Search by BIC code"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TreeView T 
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9975
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Go"
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
      Left            =   6090
      Picture         =   "BICTree.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5115
      Width           =   1000
   End
End
Attribute VB_Name = "frmBICTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSelectedBICCode As String

Public Sub FillBICTree()
    On Error GoTo errHandler
Dim objItem As d_BICCode
Dim nodX As Node
Dim lngIndex As Long
Dim RootIndex
Dim lngColour As Long
Dim ar() As Long

    If oPC.Configuration.BICs Is Null Then
        MsgBox "No BIC codes have been loaded."
    End If
    If oPC.Configuration.BICs.Count < 1 Then
        MsgBox "No BIC codes have been loaded."
    End If
    
    T.Nodes.Clear
    T.Nodes.Add , , "0", "BIC codes"
    RootIndex = T.Nodes("0").Index
    ReDim ar(0 To 10)
    ar(0) = "0"
    T.Nodes(RootIndex).Expanded = True
    If oPC.Configuration.BICs.Count = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    For lngIndex = 1 To oPC.Configuration.BICs.Count
        Set objItem = oPC.Configuration.BICs.Item(lngIndex)
        ar(objItem.Level) = lngIndex
        T.Nodes.Add CStr(ar(objItem.Level - 1)), gtRelationshipChild, CStr(lngIndex), objItem.code & ": " & objItem.Description
    Next lngIndex
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICTree.FillBICTree"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICTree.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    FillBICTree
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICTree.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub GT_Click()
    On Error GoTo errHandler
'MsgBox "Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICTree.GT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub GT_ClickNode(Node As Object, SubItem As Object)
    On Error GoTo errHandler
Dim ar() As String
    ar = Split(Node.text, ":")
    strSelectedBICCode = ar(0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBICTree.GT_ClickNode(Node,SubItem)", Array(Node, SubItem), EA_NORERAISE
    HandleError
End Sub
Public Property Get SelectedCode() As String
    SelectedCode = strSelectedBICCode
End Property
