VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmORREQ 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Order request"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8340
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      ForeColor       =   &H00714942&
      Height          =   300
      Left            =   5520
      TabIndex        =   17
      Text            =   "1"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdDeleteSelected 
      Caption         =   "X"
      Height          =   210
      Left            =   7815
      TabIndex        =   33
      Top             =   3315
      Width           =   240
   End
   Begin VB.CommandButton cmdAdd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      Picture         =   "frmORREQ2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2880
      Width           =   300
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1980
      Left            =   135
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3540
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   7424322
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Deposit per copy"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   776
      EndProperty
   End
   Begin VB.CommandButton cmdClearCust 
      Caption         =   "X"
      Height          =   210
      Left            =   5085
      TabIndex        =   30
      Top             =   510
      Width           =   240
   End
   Begin VB.TextBox txtAcno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00CDFAFA&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   3585
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   28
      ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
      Top             =   135
      Width           =   1740
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "New customer's details"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   1905
      Left            =   105
      TabIndex        =   21
      Top             =   630
      Width           =   5250
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2790
         MaxLength       =   100
         TabIndex        =   3
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   480
         Width           =   2325
      End
      Begin VB.TextBox txtInitials 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   870
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   480
         Width           =   1740
      End
      Begin VB.TextBox txtPhone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   285
         MaxLength       =   45
         TabIndex        =   4
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   1005
         Width           =   1740
      End
      Begin VB.TextBox txtTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   285
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   285
         MaxLength       =   100
         TabIndex        =   5
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   1500
         Width           =   1740
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   780
         Left            =   2235
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1005
         Width           =   2880
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   60
         TabIndex        =   27
         Top             =   810
         Width           =   705
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname or initials"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   900
         TabIndex        =   26
         Top             =   270
         Width           =   1770
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   2820
         TabIndex        =   25
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   285
         TabIndex        =   24
         Top             =   270
         Width           =   705
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   45
         TabIndex        =   23
         Top             =   1305
         Width           =   705
      End
      Begin VB.Label lblAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   2325
         TabIndex        =   22
         Top             =   795
         Width           =   1470
      End
   End
   Begin VB.CommandButton cmdFindCustomer 
      BackColor       =   &H00DACDCD&
      Caption         =   "Find customer"
      Height          =   375
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   2445
   End
   Begin VB.TextBox txtDep1 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00714942&
      Height          =   300
      Left            =   3720
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdSelectItem1 
      Height          =   315
      Left            =   2025
      Picture         =   "frmORREQ2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2925
      Width           =   345
   End
   Begin VB.TextBox txtItem1 
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   105
      TabIndex        =   14
      Top             =   2895
      Width           =   1830
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   750
      IMEMode         =   3  'DISABLE
      Left            =   210
      MaxLength       =   350
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   6135
      Width           =   4410
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6405
      Width           =   1260
   End
   Begin VB.TextBox txtDeposit 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   6330
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      Top             =   5595
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
      Caption         =   "&OK"
      Height          =   465
      Left            =   6735
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6405
      Width           =   1260
   End
   Begin VB.Label txtSPRO 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      ForeColor       =   &H8000000C&
      Height          =   300
      Left            =   2640
      TabIndex        =   36
      Top             =   2900
      Width           =   975
   End
   Begin VB.Label lblMaxLines 
      BackStyle       =   0  'Transparent
      Caption         =   "------You can store up to 35 lines here-----"
      ForeColor       =   &H00915A48&
      Height          =   285
      Left            =   450
      TabIndex        =   35
      Top             =   5535
      Width           =   4320
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00714942&
      Height          =   270
      Left            =   6975
      TabIndex        =   34
      Top             =   3300
      Width           =   705
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   240
      Left            =   6360
      TabIndex        =   32
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ac no."
      ForeColor       =   &H00714942&
      Height          =   270
      Left            =   3060
      TabIndex        =   29
      Top             =   150
      Width           =   555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   255
      Left            =   450
      TabIndex        =   20
      Top             =   2655
      Width           =   1245
   End
   Begin VB.Label lblDep 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit per copy (e.g. 20.00)      Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   210
      Left            =   3000
      TabIndex        =   18
      Top             =   2655
      Width           =   3930
   End
   Begin VB.Label lblItem1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   855
      Left            =   5640
      TabIndex        =   15
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   405
      TabIndex        =   13
      Top             =   5910
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   330
      Left            =   5430
      TabIndex        =   12
      Top             =   5685
      Width           =   1005
   End
End
Attribute VB_Name = "frmORREQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strCustomer As String
Dim strItem As String
Dim strDeposit As String
Dim lngTPID As Long
Dim oCust As a_Customer
Dim dblTotalDeposit As Double
Dim dblDeposit As Double

Private xMLDoc As ujXML

Private Sub cmdAdd_Click()
    On Error GoTo errHandler

      Dim lstItem As ListItem
10        If Not IsISBN13(txtItem1) And Not IsISBN10(txtItem1) Then Exit Sub
20        If lvw.ListItems.Count >= 30 Then
30            MsgBox "You can not store more than 30 lines in an order request." & vbCrLf & "Please start another order request.", vbInformation + vbOKOnly, "Can't add another row"
40            Exit Sub
50        End If
60        Set lstItem = lvw.ListItems.Add
70        lstItem.Text = txtItem1
80        lstItem.SubItems(1) = lblItem1.Caption
90        lstItem.SubItems(2) = lblPrice.Caption
100       lstItem.SubItems(3) = Format(CDbl(IIf(txtDep1 = "", "0", txtDep1)), "###,##0.00")
          lstItem.SubItems(4) = Format(CDbl(IIf(txtDep1 = "", "0", txtQty)), "###,##0")
110       RecalculateDeposit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORREQ.cmdAdd_Click"
End Sub
Private Sub RecalculateDeposit()
Dim i As Integer

    dblTotalDeposit = 0
    For i = 1 To lvw.ListItems.Count
        dblTotalDeposit = dblTotalDeposit + (CDbl(lvw.ListItems(i).SubItems(3)) * CDbl(lvw.ListItems(i).SubItems(4)))
    Next
    Me.txtDeposit = Format(dblTotalDeposit, "###,##0.00")
End Sub
Private Sub cmdCancel_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdClearCust_Click()
    lngTPID = 0
    Set oCust = Nothing
    Me.txtPhone = ""
    Me.txtName = ""
    Me.txtInitials = ""
    Me.txtTitle = ""
    Me.txtPhone = ""
    Me.txtAcno = ""
    Me.txtEmail.Visible = True
    Me.txtAddress.Visible = True
    Me.lblAddress.Visible = True
    Me.lblEmail.Visible = True

End Sub

Private Sub cmdDeleteSelected_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    lvw.ListItems.Remove (lvw.SelectedItem.Index)
    RecalculateDeposit
End Sub

Private Sub cmdFindCustomer_Click()
Dim frmC As frmBrowseCustomers2

    Set frmC = New frmBrowseCustomers2
    frmC.Show vbModal
    If frmC.CustomerID = 0 Then
        lngTPID = 0
        Set oCust = Nothing
        Me.txtPhone = ""
        Me.txtName = ""
        Me.txtInitials = ""
        Me.txtTitle = ""
        Me.txtPhone = ""
        Me.txtAcno = ""
        Me.txtEmail.Visible = True
        Me.txtAddress.Visible = True
        Me.lblAddress.Visible = True
        Me.lblEmail.Visible = True
        Frame1.Caption = "New customer's details"
        Frame1.Enabled = True
        Unload frmC
        Set frmC = Nothing
        Exit Sub
    End If
    lngTPID = frmC.CustomerID
    Set oCust = frmC.SelectedCustomer
    Unload frmC
    Me.txtPhone = oCust.AcNo
    Me.txtName = oCust.Name
    Me.txtInitials = oCust.Initials
    Me.txtTitle = oCust.title
    Me.txtPhone = oCust.Phone
    Me.txtAcno = oCust.AcNo
    Me.txtEmail.Visible = False
    Me.txtAddress.Visible = False
    Me.lblAddress.Visible = False
    Me.lblEmail.Visible = False
'    Me.txtEmail = oCust.Addresses.DefaultAddress.EMail
'    Me.txtAddress = oCust.Addresses.DefaultAddress.AddressMailing
    lngTPID = frmC.CustomerID
    Frame1.Caption = "Existing customer"
    Frame1.Enabled = False
End Sub

Private Sub cmdOK_Click()
'    If dblTotalDeposit < oPC.DefaultDeposit Then
'        If MsgBox("You are accepting an unusually low deposit. Do you wish to continue?", vbQuestion + vbYesNo, "Please check deposit value") = vbNo Then
'            Exit Sub
'        End If
'    End If
    Me.Hide
End Sub



Private Sub cmdSelectItem1_Click()
Dim f As New frmQuickProductFind

    If IsISBN13(Trim(txtItem1)) Or IsISBN10(Trim(txtItem1)) Then
        f.component Trim(txtItem1)
    Else
        f.component "/" & txtItem1
    End If
    f.Show vbModal
    txtItem1 = f.EAN
    If IsNumeric(f.Price) Then
        txtSPRO = "(s.p.:" & f.PriceF & ")"
        If (oPC.DefaultDeposit <> 100) Then
            txtDep1 = Format(IIf(f.Price > 0, RoundUp(CDbl(f.Price * (oPC.DefaultDeposit / 100))), 9999.99), "###,##0.00")
        Else
            txtDep1 = Format(f.Price, "###,##0.00")
        End If
   '     dblDeposit = IIf(f.Price > 0, f.Price, oPC.DefaultDeposit)
        Me.lblPrice.Caption = f.Price
    End If
    Me.lblItem1 = f.Description
    Unload f
    
End Sub

Private Sub Form_Load()
Dim arType() As String
Dim i As Integer

    
  '  txtDeposit = CStr(oPC.DefaultDeposit)
    
    SetlvwLayout Me.lvw, Me.Name
    SetFormSize Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oCust = Nothing

    SaveLayoutlvw Me.lvw, Me.Name, Me.Height, Me.Width
End Sub

Private Sub lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

'Private Sub txtDeposit_Validate(Cancel As Boolean)
'    Cancel = (Not (IsNumeric(strDeposit)))
'    CheckOKStatus
'End Sub
'Private Sub txtDeposit_Change()
'    strDeposit = txtDeposit
'End Sub
Public Property Get DepositF() As String
    DepositF = strDeposit
End Property

Public Property Get Deposit() As Long
    Deposit = CLng(dblTotalDeposit * oPC.CurrencyDivisor)
End Property


Private Sub txtDeposit_GotFocus()
    AutoSelect Controls("txtDeposit")
End Sub


Private Sub txtDep1_GotFocus()
    AutoSelect Controls("txtDep1")
End Sub

Private Sub txtDep1_LostFocus()
  '  txtDep1.Text = Format(CDbl(txtDep1), "###,##0.00")
End Sub
Private Sub txtDep1_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtDep1) And txtDep1 <> ""
End Sub


Private Sub txtItem_Change()
   strItem = txtItem
End Sub
'Public Property Get Item() As String
'   Item = Replace(strItem, vbTab, "")
'End Property
'Private Sub txtItem_Validate(Cancel As Boolean)
'  '  Cancel = (Not (Len(txtItem) > 6))
'    CheckOKStatus
'End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

'Private Sub CheckOKStatus()
'    If Len(Me.txtItem) > 6 And Len(txtCustomer) > 10 Then
'        Me.cmdOK.Enabled = True
'    End If
'End Sub

Public Function GetDetailsXML() As String
    On Error GoTo errHandler

Dim i As Integer

'MsgBox "in GetDetailsXML"
    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "OR_1"
            .chCreate "MessageType"
                .elText = "ORDER REQUEST"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "CustomerID"
                   .elText = CStr(lngTPID)
            .elCreateSibling "CustomerAcno"
                   .elText = stripTab(txtAcno)
            .elCreateSibling "CustomerTitle"
                   .elText = stripTab(txtTitle)
            .elCreateSibling "CustomerInitials"
                   .elText = stripTab(txtInitials)
            .elCreateSibling "CustomerName"
                   .elText = stripTab(txtName)
            .elCreateSibling "CustomerPhone"
                   .elText = stripTab(txtPhone)
            .elCreateSibling "CustomerEmail"
                   .elText = stripTab(txtEmail)
            .elCreateSibling "CustomerAddress"
                   .elText = stripTab(txtAddress)
            .elCreateSibling "Notes"
                .elText = stripTab(txtItem)
            .elCreateSibling "Deposit"
                .elText = str(stripTab(dblTotalDeposit))   'note Str uses decimal point always, whereas CStr uses the local setting which may be a comma
            .elCreateSibling "ItemList"
            For i = 1 To lvw.ListItems.Count
                    .chCreate "Item"
                    .chCreate "EAN", True
                        .elText = stripTab(lvw.ListItems(i).Text)
                    .elCreateSibling "PR", True
                        .elText = stripTab(lvw.ListItems(i).SubItems(2))
                    .elCreateSibling "DESCR", True
                        .elText = stripTab(lvw.ListItems(i).SubItems(1))
                    .elCreateSibling "DEP", True
                        .elText = stripTab(lvw.ListItems(i).SubItems(3))
                    .elCreateSibling "QTY", True
                        .elText = stripTab(lvw.ListItems(i).SubItems(4))
                    .navUP
                    .navUP
            Next
            .navUP
    End With
    GetDetailsXML = xMLDoc.docXML
    
        Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORREQ.GetDetailsXML"

End Function
Public Function GetDetailsForSlip() As String
    On Error GoTo errHandler
Dim i As Integer
Dim s As String
Dim s2 As String

             s = FNS(Me.txtAcno) & "~" & FNS(Me.txtTitle) & "~" & FNS(Me.txtInitials) & "~" & FNS(Me.txtName) _
             & "~" & FNS(Me.txtPhone) & "~" & FNS(Me.txtEmail) & "~" & FNS(Me.txtAddress) _
             & "~" & FNS(Me.txtItem) & "~"
             
            For i = 1 To lvw.ListItems.Count
                    s2 = lvw.ListItems(i).Text & "^^" _
                    & Me.lvw.ListItems(i).SubItems(1)
                s = s & "|" & s2
            Next
    GetDetailsForSlip = s
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORREQ.GetDetailsForSlip"
End Function

Private Sub txtTitle_GotFocus()
    Me.txtEmail.Visible = True
    Me.txtAddress.Visible = True
    Me.lblAddress.Visible = True
    Me.lblEmail.Visible = True
End Sub
