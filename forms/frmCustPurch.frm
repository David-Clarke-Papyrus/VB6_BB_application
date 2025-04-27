VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustPurch 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales to "
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvw 
      CausesValidation=   0   'False
      Height          =   4200
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7408
      SortKey         =   4
      View            =   3
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date sold"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   7303
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   " "
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmCustPurch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSPC As c_SalesPerCustomer

Public Sub component(pSPC As c_SalesPerCustomer, pFullName As String)
    On Error GoTo errHandler
    Set oSPC = pSPC
    Me.Caption = "Sales to : " & pFullName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPurch.component(pSPC,pFullName)", Array(pSPC, pFullName)
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    LoadPurchases

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPurch.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadPurchases()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To oSPC.Count
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .text = oSPC(i).dateOfSaleF
            .SubItems(1) = oSPC(i).code
            .SubItems(2) = oSPC(i).Title
            .SubItems(3) = oSPC(i).Price
            .SubItems(4) = oSPC(i).dateOfSaleForSort
        End With
    Next i

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPurch.LoadPurchases"
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
    Me.lvw.Height = NonNegative_Lng(Me.Height - 800)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPurch.Form_Resize", , EA_NORERAISE
    HandleError
End Sub


