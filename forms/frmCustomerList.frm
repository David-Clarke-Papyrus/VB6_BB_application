VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerList 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Customer from List"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2700
      Width           =   975
   End
   Begin MSComctlLib.ListView lstCustomer 
      Height          =   2415
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   33023
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmCustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CustomerID As Long

Public Sub Component(rs As ADODB.Recordset)
Dim lst As ListItem
    
    rs.MoveFirst
    With lstCustomer
        Do While Not rs.EOF
            
            Set lst = .ListItems.Add()
            lst.Tag = NZ(rs!Customer_ID)
            lst.Text = NZS(rs!C_Name)
            lst.SubItems(1) = NZS(rs!C_Address)
            lst.SubItems(2) = NZS(rs!C_Phone)
            
            rs.MoveNext
        Loop
        rs.MoveFirst
        'Set rs = Nothing
    
        If .ListItems.Count > 1 Then
            Me.cmdOK.Enabled = False
            .SelectedItem.Selected = False
        End If
    End With
    
End Sub

Private Sub cmdOK_Click()
    CustomerID = Val(Me.lstCustomer.SelectedItem.Tag)
    Me.Hide
End Sub

Private Sub Form_Load()
    With Me.lstCustomer
        .ColumnHeaders(1).Width = 1440
        .ColumnHeaders(2).Width = 2040
        .ColumnHeaders(3).Width = 960
    End With
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Debug.Print "Width of " & ColumnHeader.Text & " = " & ColumnHeader.Width
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.cmdOK.Enabled = True
    CustomerID = Val(Item.Tag)
End Sub
