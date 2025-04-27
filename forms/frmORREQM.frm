VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmORREQ 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Order request"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   735
      Width           =   1260
   End
   Begin VB.CommandButton cmdDeleteSelected 
      Caption         =   "X"
      Height          =   210
      Left            =   7770
      TabIndex        =   32
      Top             =   3300
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
      Left            =   3360
      Picture         =   "frmORREQM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2895
      Width           =   300
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1980
      Left            =   135
      TabIndex        =   30
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
      NumItems        =   4
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
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Deposit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClearCust 
      Caption         =   "X"
      Height          =   210
      Left            =   5085
      TabIndex        =   29
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
      TabIndex        =   27
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
      TabIndex        =   20
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
         Left            =   270
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
         ScrollBars      =   2  'Vertical
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   270
         Width           =   1770
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00714942&
         Height          =   270
         Left            =   2820
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
      Height          =   345
      Left            =   2445
      TabIndex        =   17
      Top             =   2895
      Width           =   855
   End
   Begin VB.CommandButton cmdSelectItem1 
      Height          =   315
      Left            =   2025
      Picture         =   "frmORREQM.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2925
      Width           =   345
   End
   Begin VB.TextBox txtItem1 
      ForeColor       =   &H00714942&
      Height          =   345
      Left            =   120
      TabIndex        =   15
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
      Left            =   195
      MaxLength       =   350
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   5850
      Width           =   4035
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00DACDCD&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6255
      Width           =   1260
   End
   Begin VB.TextBox txtDeposit 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   6840
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      Top             =   5805
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00DACDCD&
      Caption         =   "&OK"
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
      Height          =   465
      Left            =   6825
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6255
      Width           =   1260
   End
   Begin VB.Label lblLocked 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Actioned"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   465
      Left            =   5625
      TabIndex        =   37
      Top             =   1860
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit taken"
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
      Height          =   270
      Left            =   4350
      TabIndex        =   36
      Top             =   5580
      Width           =   1245
   End
   Begin VB.Label lblDepositTaken 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   285
      Left            =   4350
      TabIndex        =   35
      Top             =   5835
      Width           =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      ForeColor       =   &H00714942&
      Height          =   270
      Left            =   6930
      TabIndex        =   33
      Top             =   3270
      Width           =   705
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   240
      Left            =   3735
      TabIndex        =   31
      Top             =   2955
      Width           =   1035
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ac no."
      ForeColor       =   &H00714942&
      Height          =   270
      Left            =   3060
      TabIndex        =   28
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
      TabIndex        =   19
      Top             =   2655
      Width           =   1245
   End
   Begin VB.Label lblDep 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit (e.g. 45.99)"
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
      Left            =   2370
      TabIndex        =   18
      Top             =   2655
      Width           =   2250
   End
   Begin VB.Label lblItem1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   420
      Left            =   4305
      TabIndex        =   16
      Top             =   2820
      Width           =   3135
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
      Left            =   390
      TabIndex        =   14
      Top             =   5610
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Total deposit"
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
      Height          =   270
      Left            =   6870
      TabIndex        =   13
      Top             =   5580
      Width           =   1125
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
Dim dblCtrlDeposit As Double
Dim dblDeposit As Double
Dim SortedDoc As ujXML
Dim Res As Boolean
Dim rs As ADODB.Recordset
Dim dteExchangeDate As Date
Dim guid_OrderRequest As String
Private xMLDoc As ujXML
Dim Blocked As Boolean

Public Sub component(strXML As String, sDate As String, OR_GUID As String, pLocked As Boolean, bOK As Boolean)
56710 On Error GoTo errHandler
           
56720   guid_OrderRequest = OR_GUID
56730     If Left(strXML, 1) <> "<" Then Exit Sub
56740     bOK = ValidateFile(strXML)
56750     If Not bOK Then Exit Sub
56760     LoadFromXML strXML
56770     dteExchangeDate = CDate(sDate)
56780     lblDepositTaken.Caption = Format(dblTotalDeposit, "###,##0.00")
56790     Blocked = pLocked
56800     lblLocked.Visible = Blocked
56810     cmdOK.Enabled = Not Blocked
56820     Me.cmdAdd.Enabled = Not Blocked
56830     Me.cmdClearCust.Enabled = Not Blocked
56840     Me.cmdDeleteSelected.Enabled = Not Blocked
56850     Me.cmdFindCustomer.Enabled = Not Blocked
56860     Me.cmdAdd.Enabled = Not Blocked
56870     Me.cmdSelectItem1.Enabled = Not Blocked
56880     Exit Sub
errHandler:
56890     If ErrMustStop Then Debug.Assert False: Resume
56900     ErrorIn "frmORREQ.component", , EA_NORERAISE
56910     HandleError
End Sub
Private Sub LoadFromXML(pXML As String)
      Dim strAcno As String
      Dim strTitle As String
      Dim strInitials As String
      Dim strCustname As String

56920     On Error GoTo errHandler

56930         Set xMLDoc = New ujXML
56940         xMLDoc.docLoadXML pXML
56950         xMLDoc.navTop
              
56960         xMLDoc.navLocate "CustomerID"
56970         lngTPID = FNN(xMLDoc.Element.Text)
              
56980         xMLDoc.navLocate "CustomerAcno"
56990         txtAcno = xMLDoc.Element.Text
              
57000         xMLDoc.navLocate "CustomerTitle"
57010         txtTitle = xMLDoc.Element.Text
              
57020         xMLDoc.navLocate "CustomerName"
57030         txtName = xMLDoc.Element.Text
              
57040         xMLDoc.navLocate "CustomerInitials"
57050         txtInitials = xMLDoc.Element.Text
              
57060         xMLDoc.navLocate "CustomerPhone"
57070         txtPhone = xMLDoc.Element.Text
              
57080         xMLDoc.navLocate "CustomerEmail"
57090         txtEmail = xMLDoc.Element.Text
              
57100         xMLDoc.navLocate "CustomerAddress"
57110         txtAddress = Replace(xMLDoc.Element.Text, Chr(10), vbCrLf)
              
57120         xMLDoc.navLocate "Notes"
57130         txtItem.Text = Replace(xMLDoc.Element.Text, Chr(10), vbCrLf)
              
57140         xMLDoc.navLocate "Deposit"
57150         dblCtrlDeposit = CDbl(xMLDoc.Element.Text)
              
              
57160         Res = xMLDoc.navLocate("ItemList")
57170         Set SortedDoc = xMLDoc.docCreateViewer(True)
57180         SortedDoc.navTop
57190         If SortedDoc.chCount > 0 Then
57200             SortedDoc.elForEachElem Me
57210         End If
              
57220         If lngTPID > 0 Then
57230             Frame1.Caption = "Existing customer"
57240             Frame1.Enabled = False
57250         Else
57260             Frame1.Caption = "New customer's details"
57270             Frame1.Enabled = True
57280         End If
57290     Exit Sub
errHandler:
57300     If ErrMustStop Then Debug.Assert False: Resume
57310     ErrorIn "frmORREQ.LoadFromXML", , EA_NORERAISE
57320     HandleError
End Sub
Public Sub ProcessElement(ByVal xObj As ujXML, ByVal NavAction As XENUM_ITER_NAV, ByRef Param As Variant, ByRef SkipChildren As Boolean)
      Dim s As String
      Dim sEAN As String
      Dim sDep As String
      Dim sDescr As String
      Dim sPrice As String

57330     If IsMissing(Param) Then Param = ""
57340     If Param = "PRINT" Then
57350         If xObj.Element.nodeName = "Item" Then
57360             If NavAction <> XNAV_TO_PARENT Then
57370             xObj.navFirstChild
57380                 sEAN = xObj.Element.Text
57390                 Res = xObj.navNext
57400                 sPrice = xObj.Element.Text
57410                 Res = xObj.navNext
57420                 sDescr = xObj.Element.Text
57430                 Res = xObj.navNext
57440                 sDep = xObj.Element.Text
                      
57450                 rs.AddNew '("EAN", "Descr", "Price", "Dep"),(sEAN,sDescr,sPrice,sDep)
57460                 rs.fields("EAN") = sEAN
57470                 rs.fields("Descr") = sDescr
57480                 rs.fields("Price") = sPrice
57490                 rs.fields("Dep") = sDep
                      
57500                 xObj.navUP
57510             End If
57520         End If
57530     Else
57540         If xObj.Element.nodeName = "Item" Then
57550             If NavAction <> XNAV_TO_PARENT Then
57560             xObj.navFirstChild
57570                 sEAN = xObj.Element.Text
57580                 Res = xObj.navNext
57590                 sPrice = xObj.Element.Text
57600                 Res = xObj.navNext
57610                 sDescr = xObj.Element.Text
57620                 Res = xObj.navNext
57630                 sDep = xObj.Element.Text
57640                 If sEAN > "" Then
57650                     AddOrderline sEAN, sDescr, sPrice, sDep
57660                 End If
57670                 xObj.navUP
57680             End If
57690         End If
57700     End If
End Sub

Private Sub AddOrderline(sEAN As String, sDescr As String, sPrice As String, sDep As String)
      Dim lstItem As ListItem
57710     Set lstItem = lvw.ListItems.Add
57720     lstItem.Text = sEAN
57730     lstItem.SubItems(1) = sDescr
57740     lstItem.SubItems(2) = sPrice
57750     If IsNumeric(sDep) Then
57760         lstItem.SubItems(3) = Format(CDbl(sDep), "###,##0.00")
57770     End If
57780     RecalculateDeposit
57790     cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And (Not Blocked)
End Sub
Private Sub RecalculateDeposit()
      Dim i As Integer

57800     dblTotalDeposit = 0
57810     For i = 1 To lvw.ListItems.Count
57820         dblTotalDeposit = dblTotalDeposit + CDbl(StripToNumerics(IIf(lvw.ListItems(i).SubItems(3) = "", "0", lvw.ListItems(i).SubItems(3))))
57830     Next
57840     Me.txtDeposit = Format(dblTotalDeposit, "###,##0.00")
End Sub

Private Sub cmdAdd_Click()
      Dim lstItem As ListItem
57850     If Not IsISBN13(txtItem1) And Not IsISBN10(txtItem1) Then Exit Sub
57860     Set lstItem = lvw.ListItems.Add
57870     lstItem.Text = txtItem1
57880     lstItem.SubItems(1) = lblItem1.Caption
57890     lstItem.SubItems(2) = lblPrice.Caption
57900     If IsNumeric(txtDep1) Then
57910         lstItem.SubItems(3) = Format(CDbl(txtDep1), "###,##0.00")
57920     Else
57930         lstItem.SubItems(3) = ""
57940     End If
          
57950     RecalculateDeposit

57960     cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And (Not Blocked)
End Sub

Private Sub cmdCancel_Click()
57970     bCancelled = True
57980     Me.Hide
End Sub

Private Sub cmdClearCust_Click()
57990     lngTPID = 0
58000     Set oCust = Nothing
58010     Me.txtPhone = ""
58020     Me.txtName = ""
58030     Me.txtInitials = ""
58040     Me.txtTitle = ""
58050     Me.txtPhone = ""
58060     Me.txtAcno = ""
58070     Me.txtEmail.Visible = True
58080     Me.txtAddress.Visible = True
58090     Me.lblAddress.Visible = True
58100     Me.lblEmail.Visible = True

End Sub

Private Sub cmdDeleteSelected_Click()
58110     If lvw.SelectedItem Is Nothing Then
58120         Exit Sub
58130     End If
58140     lvw.ListItems.Remove (lvw.SelectedItem.Index)
58150     RecalculateDeposit
58160     cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And lvw.ListItems.Count > 0 And (Not Blocked)
End Sub

Private Sub cmdFindCustomer_Click()
      Dim frmC As frmBrowseCustomers2

58170     Set frmC = New frmBrowseCustomers2
58180     frmC.Show vbModal
58190     If frmC.CustomerID = 0 Then
58200         lngTPID = 0
58210         Set oCust = Nothing
58220         Me.txtPhone = ""
58230         Me.txtName = ""
58240         Me.txtInitials = ""
58250         Me.txtTitle = ""
58260         Me.txtPhone = ""
58270         Me.txtAcno = ""
58280         Me.txtEmail.Visible = True
58290         Me.txtAddress.Visible = True
58300         Me.lblAddress.Visible = True
58310         Me.lblEmail.Visible = True
58320         Frame1.Caption = "New customer's details"
58330         Frame1.Enabled = True
58340         Unload frmC
58350         Set frmC = Nothing
58360         Exit Sub
58370     End If
58380     lngTPID = frmC.CustomerID
58390     Set oCust = frmC.SelectedCustomer
58400     Unload frmC
58410     Me.txtPhone = oCust.AcNo
58420     Me.txtName = oCust.Name
58430     Me.txtInitials = oCust.Initials
58440     Me.txtTitle = oCust.Title
58450     Me.txtPhone = oCust.Phone
58460     Me.txtAcno = oCust.AcNo
58470     Me.txtEmail.Visible = False
58480     Me.txtAddress.Visible = False
58490     Me.lblAddress.Visible = False
58500     Me.lblEmail.Visible = False
      '    Me.txtEmail = oCust.Addresses.DefaultAddress.EMail
      '    Me.txtAddress = oCust.Addresses.DefaultAddress.AddressMailing
58510     lngTPID = frmC.CustomerID
58520     Frame1.Caption = "Existing customer"
58530     Frame1.Enabled = False
End Sub

Private Sub cmdOK_Click()
         ' MsgBox "SaveXML"
          
58540     PlaceOrder
        ' Me.Hide
End Sub



Private Sub cmdPrint_Click()
58550     On Error GoTo errHandler
      Dim ar As New arOrderRequest

58560     ar.fCustomer = txtAcno & vbCrLf & txtTitle & " " & txtInitials & " " & txtName & vbCrLf & txtPhone & vbCrLf & txtEmail & vbCrLf & txtAddress
58570     ar.fNote = strItem
58580     ar.fHeader = "ORDER REQUEST " & Format(dteExchangeDate, "DD/MM/YYYY HH:NN")
58590     ar.fFooter = "Printed " & Format(Now(), "DD/MM/YYYY HH:NN")
58600     Set rs = New ADODB.Recordset
58610     rs.fields.Append "EAN", adVarChar, 20
58620     rs.fields.Append "Descr", adVarChar, 500
58630     rs.fields.Append "Price", adVarChar, 20
58640     rs.fields.Append "Dep", adVarChar, 20
58650     rs.open , , adOpenDynamic, adLockOptimistic
          
58660     Res = xMLDoc.navLocate("ItemList")
58670     Set SortedDoc = xMLDoc.docCreateViewer(True)
58680     SortedDoc.navTop
58690     SortedDoc.elForEachElem Me, "PRINT"
58700     If Not rs.eof Then
58710         rs.MoveFirst
58720     End If
58730     ar.component rs
58740     ar.Caption = "Printing selected order request"
58750     ar.Show vbModal
          
58760     Exit Sub
errHandler:
58770     If ErrMustStop Then Debug.Assert False: Resume
58780     ErrorIn "frmORREQ.cmdPrint_Click", , EA_NORERAISE
58790     HandleError
End Sub

Private Sub cmdSelectItem1_Click()
      Dim f As New frmQuickProductFindOR

58800     If Me.txtItem1 > "" Then
58810         f.component txtItem1
58820     Else
58830         f.component "/"
58840     End If
58850     f.Show vbModal
58860     txtItem1 = f.EAN
58870     Me.txtDep1 = f.Price
58880     dblDeposit = FNDBL(f.Price)
58890     Me.lblPrice.Caption = f.Price
58900     Me.lblItem1 = f.Description
58910     Unload f
          
End Sub

Private Sub Form_Load()
      Dim arType() As String
      Dim i As Integer
58920     SetLvwLayout Me.lvw, Me.Name
58930     SetFormSize Me
58940     lblLocked.Visible = Blocked
End Sub

Private Sub Form_Unload(Cancel As Integer)
58950     SaveLayoutLvw Me.lvw, Me.Name, Me.Height, Me.Width
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
58960 Cancel = True
End Sub

Public Property Get DepositF() As String
58970     DepositF = strDeposit
End Property

Public Property Get Deposit() As Long
58980     Deposit = CLng(strDeposit)
End Property

Private Sub txtDep1_GotFocus()
58990     AutoSelect Controls("txtDeposit")
End Sub

Private Sub txtDep1_LostFocus()
   ' If txtDep1 > "" Then txtDep1 = txtDep1  'txtDep1.Text = Format(CDbl(txtDep1), "###,##0.00")
End Sub


Private Sub txtDep1_Validate(Cancel As Boolean)
59000     Cancel = Not IsNumeric(txtDep1) And txtDep1 <> ""
End Sub

Private Sub txtItem_Change()
59010    strItem = txtItem
End Sub
Public Property Get Cancelled() As Boolean
59020     Cancelled = bCancelled
End Property

Public Function GetDetailsXML() As String
59030 On Error GoTo errHandler

          Dim i As Integer
59040 Set xMLDoc = New ujXML
          
59050 With xMLDoc
59060     .docProgID = "MSXML2.DOMDocument"
59070     .docInit "OR_1"
59080         .chCreate "MessageType"
59090             .elText = "ORDER REQUEST"
59100         .elCreateSibling "MessageCreationDate"
59110           .elText = Format(Now(), "yyyymmddHHNN")
59120       .elCreateSibling "CustomerID"
59130              .elText = CStr(lngTPID)
59140       .elCreateSibling "CustomerAcno"
59150              .elText = Me.txtAcno
59160       .elCreateSibling "CustomerTitle"
59170              .elText = Me.txtTitle
59180       .elCreateSibling "CustomerInitials"
59190              .elText = Me.txtInitials
59200       .elCreateSibling "CustomerName"
59210              .elText = Me.txtName
59220        .elCreateSibling "CustomerPhone"
59230                 .elText = Me.txtPhone
59240          .elCreateSibling "CustomerEmail"
59250              .elText = Me.txtEmail
59260       .elCreateSibling "CustomerAddress"
59270              .elText = Me.txtAddress
59280       .elCreateSibling "Notes"
59290           .elText = Me.txtItem
59300       .elCreateSibling "Deposit"
59310           .elText = CStr(dblTotalDeposit)
59320       .elCreateSibling "ItemList"
59330       For i = 1 To lvw.ListItems.Count
59340               .chCreate "Item"
59350               .chCreate "EAN", True
59360                   .elText = lvw.ListItems(i).Text
59370               .elCreateSibling "PR", True
59380                   .elText = lvw.ListItems(i).SubItems(2)
59390               .elCreateSibling "DESCR", True
59400                   .elText = lvw.ListItems(i).SubItems(1)
59410               .elCreateSibling "DEP", True
59420                   .elText = lvw.ListItems(i).SubItems(3)
                    
59430                  .navUP
59440               .navUP
59450       Next
59460       .navUP
59470   End With

59480  GetDetailsXML = xMLDoc.docXML
59490 Exit Function
errHandler:
59500     If ErrMustStop Then Debug.Assert False: Resume
59510     ErrorIn "frmORReq.GetDetailsXML"
End Function

Private Sub txtTitle_GotFocus()
59520     Me.txtEmail.Visible = True
59530     Me.txtAddress.Visible = True
59540     Me.lblAddress.Visible = True
59550     Me.lblEmail.Visible = True
End Sub



'=============
Private Sub PlaceOrder()
59560     On Error GoTo errHandler
      Dim i As Integer
      Dim oCOL As a_COL
      Dim lngResult As Long
      Dim bFound As Boolean
      Dim bProductToOrder As Boolean
      Dim bZeroDeposit As Boolean
      Dim bIssue As Boolean
      Dim bAlternativeCustomerSelected As Boolean
      Dim oCust As a_Customer
      Dim oCO As a_CO
59570     bIssue = False
59580     If oPC.Configuration.SignTransactions = True Then
59590         If SecurityControl(enSECURITY_CO_SIGN, , "Sign this order", DOCAPPROVAL, , , gSTAFFID) = False Then
59600                Exit Sub
59610         End If
59620     End If
59630     Screen.MousePointer = vbHourglass
59640     If Not lngTPID > 0 Then 'A customer has not been specified
              
59650         Set oCust = New a_Customer
59660         oCust.BeginEdit
59670         oCust.InitializeNewCustomer enPrivate
59680         oCust.SetPhone txtPhone
59690         oCust.SetName txtName
59700         oCust.SetInitials Me.txtInitials
59710         oCust.SetTitle txtTitle
59720         oCust.SetCustomerTypeCasual
59730         oCust.SetControl txtPhone
59740         oCust.SetAccAcNo oPC.GetProperty("DefaultAccountingAccno")
59750         oCust.Addresses.FindByDescription("Default").SetEmail txtEmail
59760         oCust.Addresses.FindByDescription("Default").SetAddress Me.txtAddress
59770         If InStr(1, txtEmail, "@") > 0 Then
59780             oCust.SetDispatchMethod "M"
59790         End If
59800         If oCust.IsValid = False Then
59810             oCust.CancelEdit
59820             Set oCust = Nothing
59830             MsgBox "The customer cannot be saved, it is invalid. Please check, correct and try again.", vbInformation + vbOKOnly, "Can't do this"
59840             Exit Sub
59850         End If
59860         bAlternativeCustomerSelected = False
59870         oCust.LookforDuplicates
59880         If Not bAlternativeCustomerSelected Then
59890             oCust.ApplyEdit lngResult
59900         End If
59910         lngTPID = oCust.ID
59920     End If
          
59930     Set oCust = Nothing
59940     Set oCO = New a_CO
59950     oCO.BeginEdit
59960     oCO.SetCustomer lngTPID
59970     oCO.OrderType = enNormalCO
59980     oCO.SetMemo FNS(txtItem)
59990     oCO.StaffID = gSTAFFID
60000     oCO.ORGUID = guid_OrderRequest
60010     For i = 1 To lvw.ListItems.Count
60020         Set oCOL = oCO.COLines.Add
60030         oCOL.BeginEdit
60040         oCOL.SetLineProduct , lvw.ListItems(i).Text
60050         oCOL.SetQty 1
60060         oCOL.SetDeposit CDbl(val(StripToNumerics(lvw.ListItems(i).SubItems(3)))) * oPC.Configuration.DefaultCurrency.Divisor
60070         oCOL.SetRef ""
60080         oCOL.DepositStatus = "P"
60090         If DateDiff("d", Date, oCOL.ETA) <= 1 Then
60100             oCOL.SetETA "2w"
60110         End If
60120         oCOL.ApplyEdit
60130     Next
60140     oCO.SetStatus stISSUED
60150     oCO.Post
          
60160     Screen.MousePointer = vbDefault
60170     MsgBox "Order placed", , "Status"
          
      Dim f As New frmCOPreview
60180     f.component oCO.TRID, False
60190     Me.Hide
60200     f.Show
60210     Exit Sub
errHandler:
60220     If ErrMustStop Then Debug.Assert False: Resume
60230     ErrorIn "frmORReq.PlaceOrder"
End Sub

