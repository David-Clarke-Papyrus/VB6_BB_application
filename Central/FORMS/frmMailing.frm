VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmMailing 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse mailing addresses"
   ClientHeight    =   7530
   ClientLeft      =   750
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "frmMailing.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7080
      Picture         =   "frmMailing.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6390
      Width           =   930
   End
   Begin VB.CheckBox chkCataloguePrinting 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Final printing for catalogue"
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
      Height          =   465
      Left            =   1665
      TabIndex        =   18
      Top             =   6570
      Width           =   2640
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B4FFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6225
      TabIndex        =   16
      Top             =   6315
      Width           =   315
   End
   Begin VB.CommandButton cmdLabels 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print labels"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   15
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   6570
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find addresses meeting these criteria"
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
      Height          =   1980
      Left            =   60
      TabIndex        =   2
      ToolTipText     =   "Select any one criteria.  If using dates, a selection between dates is catered for"
      Top             =   105
      Width           =   7950
      Begin VB.ComboBox cboIG3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   1500
         Width           =   1725
      End
      Begin VB.ComboBox cboIG2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   1065
         Width           =   1725
      End
      Begin VB.ComboBox cboIG1 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   630
         Width           =   1725
      End
      Begin VB.ComboBox cboCT 
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
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   135
         TabIndex        =   7
         Text            =   "Combo1"
         ToolTipText     =   "Select a customer type or allow any customer to be found"
         Top             =   660
         Width           =   1725
      End
      Begin CoolButtonControl.CoolButton cbCatalogue 
         Height          =   465
         Left            =   4320
         TabIndex        =   4
         ToolTipText     =   "Click to change option"
         Top             =   345
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Catalogue: no restrictions"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6885
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Click to find all customers matching the retrictions selected."
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   750
      End
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
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Customer name starts like this"
         Top             =   1515
         Width           =   1695
      End
      Begin CoolButtonControl.CoolButton cbMailType 
         Height          =   465
         Left            =   4320
         TabIndex        =   5
         ToolTipText     =   "Click to change option"
         Top             =   1365
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Surface or Airmail"
         Style           =   1
         BackStyle       =   0
      End
      Begin CoolButtonControl.CoolButton cbOverseas 
         Height          =   465
         Left            =   4320
         TabIndex        =   6
         ToolTipText     =   "Click to change option"
         Top             =   855
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   820
         BackColor       =   14737632
         ForeColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Local and overseas"
         Style           =   1
         BackStyle       =   0
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest groups"
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
         Left            =   2295
         TabIndex        =   12
         Top             =   375
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer type"
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
         Left            =   150
         TabIndex        =   8
         Top             =   405
         Width           =   1605
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name like . . ."
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
         Left            =   150
         TabIndex        =   3
         Top             =   1245
         Width           =   1680
      End
   End
   Begin TrueOleDBGrid60.TDBGrid CustGrid 
      Height          =   4095
      Left            =   60
      OleObjectBlob   =   "frmMailing.frx":01F5
      TabIndex        =   9
      Top             =   2220
      Width           =   7935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gets catalogue"
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
      Left            =   4875
      TabIndex        =   17
      Top             =   6360
      Width           =   1260
   End
   Begin VB.Label lblReccFound 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   45
      TabIndex        =   15
      Top             =   6300
      Width           =   2790
   End
End
Attribute VB_Name = "frmMailing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCust As c_C_Customer
Dim oCust As a_Customer
Dim dispCust As d_C_Customer
Dim flgLoading As Boolean
Dim CustomerTypes_tl As z_TextList
Dim InterestGroups_tl As z_TextList
Dim lngTPID As Long
Dim strACCNum As String
Dim blnNoRecordsReturned As Boolean
Dim enArg_Cat As enCatalogue
Dim enArg_Overseas As enOverseas
Dim enArg_MailType As enMailType
#If H_CENTRAL = 1 Then
Dim ofrm As frmLoyaltyPreview
#Else
Dim ofrm As frmCustomerPreview
#End If
Dim XA As New XArrayDB


Private Sub cbCatalogue_Click()
    enArg_Cat = OptionLoop(enArg_Cat, 3)
    Select Case enArg_Cat
    Case enGetsCatalogueYes
        cbCatalogue.Caption = "Gets catalogue - YES"
    Case enGetsCatalogueNo
        cbCatalogue.Caption = "Gets catalogue - NO"
    Case enGetsCatalogueEither
        cbCatalogue.Caption = "Catalogue: no restriction"
    End Select
End Sub

Private Sub cboIG1_click()
    If cboIG1.ListIndex > 0 Then
        cboIG2.Enabled = True
    Else
        If cboIG2.Enabled Then
            cboIG2.ListIndex = 0
            cboIG2.Enabled = False
        End If
        If cboIG3.Enabled Then
            cboIG3.ListIndex = 0
            cboIG3.Enabled = False
        End If
    End If
End Sub
Private Sub cboIG2_click()
    If cboIG2.ListIndex > 0 Then
        cboIG3.Enabled = True
    Else
        If cboIG3.Enabled Then
            cboIG3.ListIndex = 0
            cboIG3.Enabled = False
        End If
    End If
End Sub

Private Sub cbOverseas_Click()
    enArg_Overseas = OptionLoop(enArg_Overseas, 3)
    Select Case enArg_Overseas
    Case enOverseasYes
        cbOverseas.Caption = "Overseas"
    Case enOverseasNo
        cbOverseas.Caption = "Local"
    Case enOverseasEither
        cbOverseas.Caption = "Local and overseas"
    End Select
End Sub
Private Sub cbMailType_Click()
    enArg_MailType = OptionLoop(enArg_MailType, 4)
    Select Case enArg_MailType
    Case enAirmail
        cbMailType.Caption = "Airmail"
    Case enSurfaceMail
        cbMailType.Caption = "Surface mail"
    Case enEitherMail
        cbMailType.Caption = "Surface or Airmail"
    Case enAll
        cbMailType.Caption = "Mailing and non-mailing"
    End Select
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim bRecsFound As Boolean
    On Error GoTo ERR_Handler
    blnNoRecordsReturned = False
    
    Set cCust = Nothing
    Set cCust = New c_C_Customer
    MousePointer = vbHourglass
    
    cCust.LoadForMailing bRecsFound, enArg_Cat, enArg_Overseas, enArg_MailType, Me.txtArg, _
            CustomerTypes_tl.Key(cboCT), InterestGroups_tl.Key(cboIG1), InterestGroups_tl.Key(cboIG2), InterestGroups_tl.Key(cboIG3)
    
    If blnNoRecordsReturned Then
        MsgBox "No records found", vbOKOnly + vbInformation, "Papyrus Invoicing Information"
        GoTo EXIT_Handler
    End If
    
    LoadArray
    CustGrid.ReBind
    Me.lblReccFound.Caption = CStr(cCust.Count) & " records"
'    LoadListView
'
EXIT_Handler:
    MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub cmdFind_LostFocus()
 '   LoadControls
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdLabels_Click()
'Dim frm As frmCatNo
'Dim strNo As String
'
'    If Me.chkCataloguePrinting Then
'        Set frm = New frmCatNo
'        frm.Show vbModal
'        If frm.Cancelled = False Then
'            strNo = frm.CatNo
'            cCust.markup strNo
'        End If
'        Unload frm
'    End If
    cCust.PrintLabels
End Sub

Private Sub CustGrid_DblClick()
Dim lngID As Long
Dim blnEdit As Boolean
#If H_CENTRAL = 1 Then
    Set ofrm = New frmLoyaltyPreview
#Else
    Set ofrm = New frmCustomerPreview
#End If
    lngID = Val(XA(CustGrid.Bookmark, 5))
    Set oCust = Nothing
    Set oCust = New a_Customer
    oCust.Load lngID
    ofrm.Component oCust    ', False
    ofrm.Show
End Sub



'Private Sub CustGrid_DragDrop(Source As Control, X As Single, Y As Single)
'
'End Sub
'
'Private Sub CustGrid_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'
'End Sub

Private Sub Form_Terminate()
    Set oCust = Nothing
    Set cCust = Nothing
End Sub

'Private Sub lvwCustomers_AfterLabelEdit(Cancel As Integer, NewString As String)
'Cancel = True
'End Sub
'
'Private Sub lvwCustomers_BeforeLabelEdit(Cancel As Integer)
'Cancel = True
'End Sub

'Private Sub lvwCustomers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    ' When a ColumnHeader object is clicked, the ListView control is
'    ' sorted by the subitems of that column.
'    ' Set the SortKey to the Index of the ColumnHeader - 1
'
'    lvwCustomers.SortKey = ColumnHeader.Index - 1
'    ' Set Sorted to True to sort the list.
'    lvwCustomers.Sorted = True
'    If lvwCustomers.SortOrder = lvwAscending Then
'        lvwCustomers.SortOrder = lvwDescending
'    Else
'        lvwCustomers.SortOrder = lvwAscending
'    End If
'End Sub

'Private Sub lvwCustomers_DblClick()
'Dim lngID As Long
'Dim blnEdit As Boolean
'    Set ofrm = New frmCustomerPreview
'    lngID = val(lvwCustomers.SelectedItem.Key)
'    Set oCust = Nothing
'    Set oCust = New a_Customer
'    oCust.Load lngID
'    ofrm.Component oCust    ', False
'    ofrm.Show
'End Sub

'Private Sub lvwCustomers_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then lvwCustomers_DblClick
'End Sub

Private Sub Form_Load()
    Set cCust = New c_C_Customer
    Set CustomerTypes_tl = New z_TextList
    Set InterestGroups_tl = New z_TextList
    CustomerTypes_tl.Load ltCustomerTypeActive
    InterestGroups_tl.Load ltInterestGroupActive, , "<ANY>"
    LoadCombo cboIG1, InterestGroups_tl
    LoadCombo cboIG2, InterestGroups_tl
    LoadCombo cboIG3, InterestGroups_tl
    Me.top = 0
    Me.left = 130
    Me.Width = 8200
    Me.Height = 7550
    LoadControls
    enArg_MailType = enEitherMail
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cCust = Nothing
    Set ofrm = Nothing
    Set CustomerTypes_tl = Nothing
End Sub


'Private Sub txtAccNum_Change()
'    strACCNum = txtAccnum
'End Sub
'
'Private Sub txtAccNum_Validate(Cancel As Boolean)
'    If txtAccnum > "" Then
'        cmdFind.Enabled = True
'    End If
'End Sub

Private Sub LoadControls()
    flgLoading = True
    LoadCombo cboCT, CustomerTypes_tl
    cboCT = CustomerTypes_tl.Item("0")
    txtArg = ""
    lngTPID = 0
    flgLoading = False
End Sub

'Private Sub LoadListView()
'Dim objItm As ListItem
'Dim i As Integer
'Dim tmp As String
'
'    lvwCustomers.ListItems.Clear
'    For i = 1 To cCust.Count
'        Set objItm = Me.lvwCustomers.ListItems.Add
'        With objItm
'            .Key = cCust(i).ID & "K"
'            .Text = cCust(i).Name '& (IIf(Len(Trim(cCust(i).Name)) <= 1, "", "(" & Trim(cCust(i).Phone) & ")"))
'           ' .SubItems(1) = cCust(i).TPName
'            .SubItems(1) = cCust(i).AcNo
'            .SubItems(2) = cCust(i).Phone
'            .SubItems(3) = cCust(i).CustomerTypeDescription
''            If cCust(i).Status = "VOID" Then
''                objItm.ForeColor = CL_DARKBLUE
''            ElseIf cCust(i).Status = "IN PROCESS" Then
''                objItm.ForeColor = vbRed
''            End If
'        End With
'    Next i
'End Sub
'Private Sub SetLvw()
'Dim style As Long
'Dim hHeader As Long
'   hHeader = SendMessage(lvwCustomers.hwnd, LVM_GETHEADER, 0, ByVal 0&)
'   style = GetWindowLong(hHeader, GWL_STYLE)
'   style = style Xor HDS_BUTTONS
'   If style Then
'      Call SetWindowLong(hHeader, GWL_STYLE, style)
'      Call SetWindowPos(lvwCustomers.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If
'End Sub
Private Sub LoadArray()
Dim objItem As d_C_Customer
Dim itmList As ListItem
Dim lngIndex As Long
    XA.ReDim 1, cCust.Count, 1, 6
    For lngIndex = 1 To cCust.Count
        With objItem
            Set objItem = cCust.Item(lngIndex)
'            Set itmList = lvwSO.ListItems.Add(Key:=Format$(objItem.TID) & " K")
            XA.Value(lngIndex, 1) = lngIndex
            XA.Value(lngIndex, 2) = objItem.Fullname2
            XA.Value(lngIndex, 3) = objItem.ListAddress
      '      XA.Value(lngIndex, 3) = objItem.Phone
      '      XA.Value(lngIndex, 4) = objItem.CustomerTypeDescription
            XA.Value(lngIndex, 5) = objItem.ID
            XA.Value(lngIndex, 6) = objItem.GetsCatalogue
        End With
    Next
 '   XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    CustGrid.Array = XA
    cmdLabels.Enabled = cCust.Count > 0
End Sub

Private Sub CustGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XA(Bookmark, 6) = True Then
        RowStyle.BackColor = RGB(282, 274, 180)
    End If
End Sub

Private Sub lblRecords_Click()

End Sub

