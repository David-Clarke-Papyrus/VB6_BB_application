VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPossibleDuplicateCOLS 
   Caption         =   "Other sales order lines with identical reference"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDoNotContinue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2085
      Picture         =   "frmPossibleDuplicateCOLS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1000
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Continue"
      Height          =   615
      Left            =   4170
      Picture         =   "frmPossibleDuplicateCOLS.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvwLines 
      Height          =   1530
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   2699
      View            =   3
      SortOrder       =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Document"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Signed by"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Product"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Qty"
         Object.Width           =   353
      EndProperty
   End
End
Attribute VB_Name = "frmPossibleDuplicateCOLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim bOK As Boolean
Public Property Get OKToContinue() As Boolean
    OKToContinue = bOK
End Property
Public Sub component(pRs As ADODB.Recordset)
    On Error GoTo errHandler
    bOK = False
    Set rs = pRs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPossibleDuplicateCOLS.component(pRS)", pRs
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim li As ListItem
    
    Do While Not rs.eof
        Set li = lvwLines.ListItems.Add
        li.text = FNS(rs.fields("DocCode"))
        li.SubItems(1) = FNS(rs.fields("DocDate"))
        li.SubItems(2) = FNS(rs.fields("Signedby"))
        li.SubItems(3) = FNS(rs.fields("ProductCode"))
        li.SubItems(4) = FNS(rs.fields("ProductDescription"))
        li.SubItems(5) = FNS(rs.fields("Qty"))
        li.Checked = True
        rs.MoveNext
    Loop
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPossibleDuplicateCOLS.LoadGrid"
End Sub

Private Sub cmdContinue_Click()
    bOK = True
    Me.Hide
End Sub

Private Sub cmdDoNotContinue_Click()
    bOK = False
    Me.Hide
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    lvwLines.Checkboxes = False
    lvwLines.Width = 8900
    lvwLines.ColumnHeaders(1).Width = 1400
    lvwLines.ColumnHeaders(2).Width = 1200
    lvwLines.ColumnHeaders(3).Width = 1000
    lvwLines.ColumnHeaders(4).Width = 1500
    lvwLines.ColumnHeaders(5).Width = 3000
    lvwLines.ColumnHeaders(6).Width = 700
    LoadGrid
    Me.Width = 9400
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPossibleDuplicateCOLS.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPossibleDuplicateCOLS.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
