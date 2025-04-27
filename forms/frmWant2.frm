VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWant 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Want"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   225
      TabIndex        =   8
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   60
      Width           =   4920
      Begin VB.TextBox txtArg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   945
         TabIndex        =   10
         ToolTipText     =   "Enter A/C number, name, start of name, telephone number or end of telephone number and click FIND."
         Top             =   315
         Width           =   1500
      End
      Begin VB.CommandButton CMDfIND 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2565
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Looks for A/C number first, then name, then phone number"
         Top             =   285
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Look for"
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
         Height          =   270
         Left            =   75
         TabIndex        =   11
         Top             =   330
         Width           =   780
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
      TabIndex        =   3
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
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Height          =   330
      Left            =   1065
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4170
      Width           =   4155
   End
   Begin MSComctlLib.ListView lvwCustomers 
      Height          =   1155
      Left            =   225
      TabIndex        =   4
      Top             =   1260
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2037
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
         Name            =   "Arial Narrow"
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
   Begin MSComctlLib.ListView lvwOrder 
      Height          =   1155
      Left            =   225
      TabIndex        =   7
      Top             =   2910
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2037
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
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   3176
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   1589
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer wants these titles"
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
      Height          =   225
      Left            =   240
      TabIndex        =   12
      Top             =   2670
      Width           =   3015
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   705
      Left            =   285
      TabIndex        =   6
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
      Height          =   300
      Left            =   330
      TabIndex        =   5
      Top             =   2940
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   240
      TabIndex        =   1
      Top             =   4215
      Width           =   690
   End
End
Attribute VB_Name = "frmWant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim WithEvents oWant As a_Want
Dim flgLoading As Boolean
Dim tlHeadings As z_TextList
Dim bCancel As Boolean
Dim cCust As c_Customer
Dim dispCust As d_Customer
Dim lngTPID As Long
Dim strACCNum As String
Dim oCust As a_Customer
Dim oCO As a_CO
Dim oCOL As a_COL
Dim strPID As String
Dim blnNoRecordsReturned As Boolean
Dim oProduct As a_Product

Public Sub component(pProduct As a_Product)
    On Error GoTo errHandler
    Set oProduct = pProduct
    strPID = oProduct.PID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.component(pProduct)", pProduct
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
   ' oWant.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim strResult As String
    If MsgBox("Confirm: You are adding " & oProduct.Title & " to the wants list for : " & cCust.Item(lvwCustomers.SelectedItem.Index).Fullname2, vbInformation + vbYesNo, "Confirmation") = vbYes Then
        oCOL.SetWantDate Format(Date, "yyyy-mm-dd")
        oCOL.ApplyEdit
        oCO.SetStatus stISSUED
        
        oCO.ApplyEdit strResult
        If strResult > "" Then
            MsgBox "Want has not been saved for the following reason: ", , "Problem" & strResult
        End If
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub EnableOK(pStatus As Boolean)
    On Error GoTo errHandler
    cmdOK.Enabled = pStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.EnableOK(pStatus)", pStatus
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    Set cCust = New c_Customer
    If Me.WindowState <> 2 Then
        Me.TOP = 1800
        Me.Left = 50
        Me.Width = 6000
        Me.Height = 6800
    End If
    LoadControls
    EnableOK False
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Label5_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.Label5_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCustomers_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.lvwCustomers_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub oWant_Valid(errors As String, pValid As Boolean)
    On Error GoTo errHandler
    EnableOK pValid
    Me.lblErrors = errors
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.oWant_Valid(errors,pValid)", Array(errors, pValid), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind_Click()
    On Error GoTo errHandler

    blnNoRecordsReturned = False
    
    Set cCust = Nothing
    Set cCust = New c_Customer
    MousePointer = vbHourglass
    Me.lvwCustomers.ListItems.Clear
    
    
    cCust.LoadEasy txtArg, False ', txtPhone, Me.txtAccnum  'blnNoRecordsReturned, lngTPID, strInvoiceNum, dteDate1, dteDate2
    
'    If blnNoRecordsReturned Then
'        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
'        GoTo EXIT_Handler
'    End If
    
    LoadListView

EXIT_Handler:
    MousePointer = vbDefault
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.cmdFind_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFind_LostFocus()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.cmdFind_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCustomers_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.lvwCustomers_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCustomers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.lvwCustomers_ColumnClick(ColumnHeader)", ColumnHeader, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCustomers_DblClick()
    On Error GoTo errHandler
Dim blnEdit As Boolean
    If lvwCustomers.SelectedItem.Index < 1 Then Exit Sub
    lngTPID = val(lvwCustomers.SelectedItem.Key)
  '  oWant.SetTPID lngID
   ' oWant.CustomerName = cCust.Item(lvwCustomers.SelectedItem.Key).Fullname
 '   Me.lblCustomer.Caption = cCust.Item(lvwCustomers.SelectedItem.Key).FullIdentification
    lblCustomer.Caption = lvwCustomers.SelectedItem.text
    Set oCO = New a_CO
    If oCO.LoadWantsForTP(lngTPID) Then 'an order is found
        LoadCO
        oCO.BeginEdit
    Else
        Set oCO = Nothing
        Set oCO = New a_CO
        oCO.BeginEdit
        oCO.SetCustomer lngTPID
        oCO.OrderType = enWant
        lvwOrder.ListItems.Clear
    End If
    Set oCOL = oCO.COLines.Add
    oCOL.BeginEdit
    oCOL.SetLineProduct strPID
    EnableOK True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.lvwCustomers_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCustomers_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then lvwCustomers_DblClick
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.lvwCustomers_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set cCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


'Private Sub txtAccNum_Change()
'    strACCNum = txtAccnum
'End Sub
'
'Private Sub txtAccNum_Validate(Cancel As Boolean)
'    If txtAccnum > "" Then
'        CMDfIND.Enabled = True
'    End If
'End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
 '   Me.txtNote = oWant.Note
   ' Me.txtDate = oWant.reqdate
 '   Me.lblCustomer = oWant.CustomerName
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.LoadControls"
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwCustomers.ListItems.Clear
    For i = 1 To cCust.Count
        Set objItm = Me.lvwCustomers.ListItems.Add
        With objItm
            .Key = cCust(i).ID & "K"
            .text = cCust(i).Fullname2 '& (IIf(Len(Trim(cCust(i).Name)) <= 1, "", "(" & Trim(cCust(i).Phone) & ")"))
            .SubItems(1) = cCust(i).AcNo
            .SubItems(2) = cCust(i).Phone
        End With
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.LoadListView"
End Sub
Private Sub LoadCO()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    Me.lvwOrder.ListItems.Clear
    For i = 1 To oCO.COLines.Count
        Set objItm = Me.lvwOrder.ListItems.Add
        With objItm
            .Key = oCO.COLines(i).Key
            .text = oCO.COLines(i).CodeF
            .SubItems(1) = oCO.COLines(i).TitleAuthorPublisher
            .SubItems(2) = oCO.COLines(i).PriceF
        End With
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.LoadCO"
End Sub

Private Sub txtNote_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lngTPID = 0 Then Exit Sub
    oCOL.Note = txtNote
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmWant.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub
