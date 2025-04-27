VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWant 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Want"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   45
      Width           =   4950
      Begin VB.TextBox txtTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   1185
         TabIndex        =   10
         Top             =   255
         Width           =   870
      End
      Begin VB.TextBox txtAccNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1185
         TabIndex        =   9
         Top             =   675
         Width           =   1500
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
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
         Height          =   720
         Left            =   3195
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   1560
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1185
         TabIndex        =   7
         Top             =   1095
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Name"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   210
         TabIndex        =   13
         Top             =   285
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acc.num."
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   210
         TabIndex        =   12
         Top             =   690
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Phone"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   420
         TabIndex        =   11
         Top             =   1110
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5505
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4125
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5505
      Width           =   930
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   990
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4620
      Width           =   4335
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   990
      TabIndex        =   0
      Top             =   4080
      Width           =   2715
   End
   Begin MSComctlLib.ListView lvwCustomers 
      Height          =   1365
      Left            =   345
      TabIndex        =   14
      Top             =   1740
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2408
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Acc. num."
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000C0&
      Height          =   705
      Left            =   300
      TabIndex        =   16
      Top             =   5385
      Width           =   2520
   End
   Begin VB.Label lblCustomer 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   975
      TabIndex        =   15
      Top             =   3570
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Note"
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
      Height          =   300
      Left            =   165
      TabIndex        =   3
      Top             =   4665
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Date"
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
      Height          =   300
      Left            =   165
      TabIndex        =   1
      Top             =   4125
      Width           =   690
   End
End
Attribute VB_Name = "frmWant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oWant As a_Want
Attribute oWant.VB_VarHelpID = -1
Dim flgLoading As Boolean
Dim tlHeadings As z_TextList
Dim bCancel As Boolean
Dim cCust As c_Customer
Dim dispCust As d_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim oCust As a_Customer
Dim blnNoRecordsReturned As Boolean

Public Sub component(pWant As a_Want)
    Set oWant = pWant
    oWant.BeginEdit
End Sub

Private Sub cmdCancel_Click()
    oWant.CancelEdit
    Unload Me
End Sub

Private Sub cmdOK_Click()
    oWant.ApplyEdit
    Unload Me
End Sub
Private Sub EnableOK(pStatus As Boolean)
    cmdOK.Enabled = pStatus
End Sub

Private Sub Form_Load()
    Set cCust = New c_Customer
    Me.Top = 1800
    Me.Left = 50
    Me.Width = 6000
    Me.Height = 6800
    LoadControls

End Sub

Private Sub lvwCustomers_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub oWant_Valid(errors As String, pValid As Boolean)
    EnableOK pValid
    Me.lblErrors = errors
End Sub

Private Sub cmdFind_Click()

    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set cCust = Nothing
    Set cCust = New c_Customer
    MousePointer = vbHourglass
    Me.lvwCustomers.ListItems.Clear
    
    
    cCust.Load Me.txtTP, txtPhone, Me.txtAccnum 'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    
    LoadListView

EXIT_Handler:
    MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub cmdFind_LostFocus()
    LoadControls
End Sub

Private Sub Form_Terminate()
    Set cCust = Nothing
End Sub

Private Sub lvwCustomers_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub

Private Sub lvwCustomers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    
    lvwCustomers.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lvwCustomers.Sorted = True
    If lvwCustomers.SortOrder = lvwAscending Then
        lvwCustomers.SortOrder = lvwDescending
    Else
        lvwCustomers.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwCustomers_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
    lngID = val(lvwCustomers.SelectedItem.Key)
    oWant.SetTPID lngID
    oWant.CustomerName = cCust.Item(lvwCustomers.SelectedItem.Key).Fullname
    Me.lblCustomer.Caption = cCust.Item(lvwCustomers.SelectedItem.Key).Fullname
End Sub

Private Sub lvwCustomers_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then lvwCustomers_DblClick
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set cCust = Nothing
End Sub


Private Sub txtAccNum_Change()
    strACCNum = txtAccnum
End Sub

Private Sub txtAccNum_Validate(Cancel As Boolean)
    If txtAccnum > "" Then
        cmdFind.Enabled = True
    End If
End Sub

Private Sub LoadControls()
    flgLoading = True
    Me.txtNote = oWant.Note
    Me.txtDate = oWant.reqdate
    Me.lblCustomer = oWant.CustomerName
    flgLoading = False
End Sub

Private Sub LoadListView()
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwCustomers.ListItems.Clear
    For i = 1 To cCust.Count
        Set objItm = Me.lvwCustomers.ListItems.Add
        With objItm
            .Key = cCust(i).ID & "K"
            .Text = cCust(i).FullName2 '& (IIf(Len(Trim(cCust(i).Name)) <= 1, "", "(" & Trim(cCust(i).Phone) & ")"))
            .SubItems(1) = cCust(i).AcNo
            .SubItems(2) = cCust(i).Phone
        End With
    Next i
End Sub



Private Sub txtDate_GotFocus()
    AutoSelect txtDate
End Sub

Private Sub txtDate_LostFocus()
    If flgLoading Then Exit Sub
    txtDate = oWant.ReqDateF
End Sub
Private Sub txtDate_Validate(Cancel As Boolean)
    Cancel = Not oWant.SetRequestdate(txtDate)
End Sub

Private Sub txtNote_Change()
    oWant.Note = txtNote
End Sub
