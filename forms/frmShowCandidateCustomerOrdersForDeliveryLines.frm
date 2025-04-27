VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShowCandidateCustomerOrdersForDeliveryLines 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Candidate customer orders"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   10875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Close"
      Height          =   315
      Left            =   9900
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   75
      Width           =   825
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1485
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   10710
      _ExtentX        =   18891
      _ExtentY        =   2619
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Order"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contact details"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty owed"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Note"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Customers waiting for this item"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   8460
   End
End
Attribute VB_Name = "frmShowCandidateCustomerOrdersForDeliveryLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset

Public Sub component(pRs As ADODB.Recordset)
    Set rs = pRs
    LoadListView
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    lvw.Width = NonNegative_Lng(Me.Width - 250)
    lvw.Height = NonNegative_Lng(Me.Height - 1000)
End Sub

Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    Do While Not rs.eof
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .text = FNS(rs.fields("CustomerName"))
            .SubItems(1) = FNS(rs.fields("DocumentCode"))
            .SubItems(2) = FND(rs.fields("DocumentDate"))
            .SubItems(3) = FNS(rs.fields("ContactDetails"))
            .SubItems(4) = FNN(rs.fields("Qty"))
            .SubItems(5) = FNS(rs.fields("Note"))

            
        End With
        rs.MoveNext
    Loop
    
End Sub
