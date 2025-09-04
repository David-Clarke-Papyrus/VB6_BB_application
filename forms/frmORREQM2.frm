VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
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
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      ForeColor       =   &H00714942&
      Height          =   300
      Left            =   5520
      TabIndex        =   38
      Text            =   "1"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Print"
      Height          =   465
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   120
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
      Left            =   6600
      Picture         =   "frmORREQM2.frx":0000
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
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   600
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
         Caption         =   "Surname"
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
      Height          =   300
      Left            =   3720
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdSelectItem1 
      Height          =   315
      Left            =   2025
      Picture         =   "frmORREQM2.frx":038A
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
      Height          =   465
      Left            =   5190
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6375
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
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   6720
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
      Height          =   465
      Left            =   6705
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6375
      Width           =   1260
   End
   Begin VB.Label lblMatchStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00714942&
      BackStyle       =   0  'Transparent
      Caption         =   "Match"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00714942&
      Height          =   270
      Left            =   5730
      TabIndex        =   40
      Top             =   5820
      Width           =   930
   End
   Begin VB.Label txtSPRO 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      ForeColor       =   &H8000000C&
      Height          =   300
      Left            =   2640
      TabIndex        =   39
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblLocked 
      BackStyle       =   0  'Transparent
      Caption         =   "Actioned"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   5520
      TabIndex        =   37
      Top             =   2280
      Width           =   1095
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
      Left            =   4305
      TabIndex        =   36
      Top             =   5550
      Width           =   1245
   End
   Begin VB.Label lblDepositTaken 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00714942&
      Height          =   285
      Left            =   4305
      TabIndex        =   35
      Top             =   5805
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
      Left            =   5520
      TabIndex        =   31
      Top             =   1920
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
      Width           =   3570
   End
   Begin VB.Label lblItem1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00714942&
      Height          =   780
      Left            =   5520
      TabIndex        =   16
      Top             =   960
      Width           =   2535
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
      Caption         =   "Deposit accepted"
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
      Left            =   6390
      TabIndex        =   13
      Top             =   5580
      Width           =   1725
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
Dim dblQty As Double
Dim dblDeposit As Double
Dim SortedDoc As ujXML
Dim Res As Boolean
Dim rs As ADODB.Recordset
Dim dteExchangeDate As Date
Dim guid_OrderRequest As String
Private xMLDoc As ujXML
Dim Blocked As Boolean

Public Sub component(strXML As String, sDate As String, OR_GUID As String, pLocked As Boolean, bOK As Boolean)
On Error GoTo errHandler
     
  guid_OrderRequest = OR_GUID
    If Left(strXML, 1) <> "<" Then Exit Sub
    bOK = ValidateFile(strXML)
    If Not bOK Then Exit Sub
    LoadFromXML strXML
    dteExchangeDate = CDate(sDate)
    lblDepositTaken.Caption = Format(dblCtrlDeposit, "###,##0.00")
    Blocked = pLocked
    lblLocked.Visible = Blocked
    cmdOK.Enabled = Not Blocked
    cmdCancel.Caption = IIf(Blocked, "Close", "Cancel")
    Me.cmdAdd.Enabled = Not Blocked
    Me.cmdClearCust.Enabled = Not Blocked
    Me.cmdDeleteSelected.Enabled = Not Blocked
    Me.cmdFindCustomer.Enabled = Not Blocked
    Me.cmdAdd.Enabled = Not Blocked
    Me.cmdSelectItem1.Enabled = Not Blocked
    RecalculateDeposit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORREQ.component", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadFromXML(pXML As String)
Dim strAcno As String
Dim strTitle As String
Dim strInitials As String
Dim strCustname As String
Dim Res As Boolean
    On Error GoTo errHandler

        Set xMLDoc = New ujXML
        xMLDoc.docLoadXML pXML
        xMLDoc.navTop
        
        xMLDoc.navLocate "CustomerID"
        lngTPID = FNN(xMLDoc.Element.text)
        
        xMLDoc.navLocate "CustomerAcno"
        txtAcno = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerTitle"
        txtTitle = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerName"
        txtName = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerInitials"
        txtInitials = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerPhone"
        txtPhone = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerEmail"
        txtEmail = xMLDoc.Element.text
        
        xMLDoc.navLocate "CustomerAddress"
        txtAddress = Replace(xMLDoc.Element.text, Chr(10), vbCrLf)
        
        xMLDoc.navLocate "Notes"
        txtItem.text = Replace(xMLDoc.Element.text, Chr(10), vbCrLf)
        
        xMLDoc.navLocate "Deposit"
        dblCtrlDeposit = val(xMLDoc.Element.text)
        
        If Not xMLDoc.navLocate("Qty") Then
            Res = xMLDoc.navLocate("QTY")
        End If
        If Res Then   ' some earlier ORs will have XMLDocs with no such node as Qty. Qty therefore must defulat to 1(one).
            dblQty = val(xMLDoc.Element.text)
        Else
            dblQty = 1
        End If
        
        
        Res = xMLDoc.navLocate("ItemList")
        Set SortedDoc = xMLDoc.docCreateViewer(True)
        SortedDoc.navTop
        If SortedDoc.chCount > 0 Then
            SortedDoc.elForEachElem Me
        End If
        
        If lngTPID > 0 Then
            Frame1.Caption = "Existing customer"
            Frame1.Enabled = False
        Else
            Frame1.Caption = "New customer's details"
            Frame1.Enabled = True
        End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORREQ.LoadFromXML", , EA_NORERAISE
    HandleError
End Sub
Public Sub ProcessElement(ByVal xObj As ujXML, ByVal NavAction As XENUM_ITER_NAV, ByRef Param As Variant, ByRef SkipChildren As Boolean)
Dim s As String
Dim sEAN As String
Dim sDep As String
Dim sDescr As String
Dim sPrice As String
Dim sQty As String
    If IsMissing(Param) Then Param = ""
    If Param = "PRINT" Then
        If xObj.Element.nodeName = "Item" Then
            If NavAction <> XNAV_TO_PARENT Then
            xObj.navFirstChild
                sEAN = xObj.Element.text
                Res = xObj.navNext
                sPrice = xObj.Element.text
                Res = xObj.navNext
                sDescr = xObj.Element.text
                Res = xObj.navNext
                sDep = xObj.Element.text
                
                rs.AddNew '("EAN", "Descr", "Price", "Dep"),(sEAN,sDescr,sPrice,sDep)
                rs.Fields("EAN") = sEAN
                rs.Fields("Descr") = sDescr
                rs.Fields("Price") = sPrice
                rs.Fields("Dep") = sDep
                
                xObj.navUP
            End If
        End If
    Else
        If xObj.Element.nodeName = "Item" Then
            If NavAction <> XNAV_TO_PARENT Then
            xObj.navFirstChild
                sEAN = xObj.Element.text
                Res = xObj.navNext
                sPrice = xObj.Element.text
                Res = xObj.navNext
                sDescr = xObj.Element.text
                Res = xObj.navNext
                sDep = xObj.Element.text
                Res = xObj.navNext
                If UCase(xObj.Element.nodeName) = UCase("Qty") Then
                    sQty = xObj.Element.text
                Else
                    sQty = 1
                End If
                If sEAN > "" Then
                    AddOrderline sEAN, sDescr, sPrice, sDep, sQty
                End If
                xObj.navUP
            End If
        End If
    End If
End Sub

Private Sub AddOrderline(sEAN As String, sDescr As String, sPrice As String, sDep As String, sQty As String)
Dim lstItem As ListItem
    Set lstItem = lvw.ListItems.Add
    lstItem.text = sEAN
    lstItem.SubItems(1) = sDescr
    lstItem.SubItems(2) = sPrice
    lstItem.SubItems(3) = sQty
    If IsNumeric(sDep) Then
        lstItem.SubItems(3) = Format(CDbl(sDep), "###,##0.00")
    End If
     If IsNumeric(sQty) Then
        lstItem.SubItems(4) = Format(CDbl(sQty), "###,##0")
    End If
    'we can't recalculate the deposit as it has already been taken
    'RecalculateDeposit
    cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And (Not Blocked)
End Sub
Private Sub RecalculateDeposit()
Dim i As Integer

    dblTotalDeposit = 0
    For i = 1 To lvw.ListItems.Count
        dblTotalDeposit = dblTotalDeposit + ((val(StripToNumerics(IIf(lvw.ListItems(i).SubItems(3) = "", "0", lvw.ListItems(i).SubItems(3))))) * (val(StripToNumerics(IIf(lvw.ListItems(i).SubItems(4) = "", "0", lvw.ListItems(i).SubItems(4))))))
    Next
    Me.txtDeposit = Format(dblTotalDeposit, "###,##0.00")
    
    cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And (Not Blocked)
    lblMatchStatus.Caption = IIf((dblCtrlDeposit = dblTotalDeposit), "match", "mismatch")
    lblMatchStatus.ForeColor = IIf((dblCtrlDeposit = dblTotalDeposit), vbBlack, vbRed)

End Sub

Private Sub cmdAdd_Click()
Dim lstItem As ListItem
    If Not IsISBN13(txtItem1) And Not IsISBN10(txtItem1) Then Exit Sub
    Set lstItem = lvw.ListItems.Add
    lstItem.text = txtItem1
    lstItem.SubItems(1) = lblItem1.Caption
    lstItem.SubItems(2) = lblPrice.Caption
    lstItem.SubItems(3) = Format(CDbl(IIf(txtDep1 = "", "0", txtDep1)), "###,##0.00")
    lstItem.SubItems(4) = Format(CDbl(IIf(txtDep1 = "", "0", txtQty)), "###,##0")
    If IsNumeric(txtDep1) Then
        lstItem.SubItems(3) = Format(CDbl(txtDep1), "###,##0.00")
    Else
        lstItem.SubItems(3) = ""
    End If
    
    RecalculateDeposit

'    cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And (Not Blocked)
'    lblMatchStatus.Caption = IIf((dblCtrlDeposit = dblTotalDeposit) And (Not Blocked), "match", "mismatch")
'    lblMatchStatus.ForeColor = IIf((dblCtrlDeposit = dblTotalDeposit) And (Not Blocked), vbBlack, vbRed)

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
    Me.txtQty = 0
    Me.txtEmail.Visible = True
    Me.txtAddress.Visible = True
    Me.lblAddress.Visible = True
    Me.lblEmail.Visible = True

End Sub

Private Sub cmdDeleteSelected_Click()
    If lvw.SelectedItem Is Nothing Then
        Exit Sub
    End If
    lvw.ListItems.Remove (lvw.SelectedItem.Index)
    RecalculateDeposit
    cmdOK.Enabled = (dblCtrlDeposit = dblTotalDeposit) And lvw.ListItems.Count > 0 And (Not Blocked)
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
    Me.txtQty = 0
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
    Me.txtTitle = oCust.Title
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
   ' MsgBox "SaveXML"
    
    PlaceOrder
  ' Me.Hide
End Sub



Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim ar As New arOrderRequest

    ar.fCustomer = txtAcno & vbCrLf & txtTitle & " " & txtInitials & " " & txtName & vbCrLf & txtPhone & vbCrLf & txtEmail & vbCrLf & txtAddress
    ar.fNote = strItem
    ar.fHeader = "ORDER REQUEST " & Format(dteExchangeDate, "DD/MM/YYYY HH:NN")
    ar.fFooter = "Printed " & Format(Now(), "DD/MM/YYYY HH:NN")
    Set rs = New ADODB.Recordset
    rs.Fields.Append "EAN", adVarChar, 20
    rs.Fields.Append "Descr", adVarChar, 500
    rs.Fields.Append "Price", adVarChar, 20
    rs.Fields.Append "Dep", adVarChar, 20
    rs.Fields.Append "Qty", adVarChar, 20
    rs.Open , , adOpenDynamic, adLockOptimistic
    
    Res = xMLDoc.navLocate("ItemList")
    Set SortedDoc = xMLDoc.docCreateViewer(True)
    SortedDoc.navTop
    SortedDoc.elForEachElem Me, "PRINT"
    If Not rs.eof Then
        rs.MoveFirst
    End If
    ar.component rs
    ar.Caption = "Printing selected order request"
    ar.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORREQ.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectItem1_Click()
Dim f As New frmQuickProductFindOR

    If Me.txtItem1 > "" Then
        f.component txtItem1
    Else
        f.component "/"
    End If
    f.Show vbModal
    txtItem1 = f.EAN
    Me.txtDep1 = f.Price
    dblDeposit = FNDBL(f.Price)
'    iQty = CInt(f.Qty)
    Me.lblPrice.Caption = Format(CDbl(f.Price), "###,##0.00")
    Me.lblItem1 = f.Description
    Unload f
    
End Sub

Private Sub Form_Load()
Dim arType() As String
Dim i As Integer
    SetLvwLayout Me.lvw, Me.Name
    SetFormSize Me
    lblLocked.Visible = Blocked
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLayoutLvw Me.lvw, Me.Name, Me.Height, Me.Width
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Public Property Get DepositF() As String
    DepositF = strDeposit
End Property

Public Property Get Deposit() As Long
    Deposit = CLng(strDeposit)
End Property

Private Sub txtDep1_GotFocus()
    AutoSelect Controls("txtDeposit")
End Sub

Private Sub txtDep1_LostFocus()
   ' If txtDep1 > "" Then txtDep1 = txtDep1  'txtDep1.Text = Format(CDbl(txtDep1), "###,##0.00")
End Sub


Private Sub txtDep1_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric(txtDep1) And txtDep1 <> ""
End Sub

Private Sub txtItem_Change()
   strItem = txtItem
End Sub
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Public Function GetDetailsXML() As String
On Error GoTo errHandler

    Dim i As Integer
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
             .elText = Me.txtAcno
      .elCreateSibling "CustomerTitle"
             .elText = Me.txtTitle
      .elCreateSibling "CustomerInitials"
             .elText = Me.txtInitials
      .elCreateSibling "CustomerName"
             .elText = Me.txtName
       .elCreateSibling "CustomerPhone"
                .elText = Me.txtPhone
         .elCreateSibling "CustomerEmail"
             .elText = Me.txtEmail
      .elCreateSibling "CustomerAddress"
             .elText = Me.txtAddress
      .elCreateSibling "Notes"
          .elText = Me.txtItem
      .elCreateSibling "Deposit"
      RecalculateDeposit
          .elText = CStr(dblTotalDeposit)
      .elCreateSibling "ItemList"
      For i = 1 To lvw.ListItems.Count
              .chCreate "Item"
              .chCreate "EAN", True
                  .elText = lvw.ListItems(i).text
              .elCreateSibling "PR", True
                  .elText = lvw.ListItems(i).SubItems(2)
              .elCreateSibling "DESCR", True
                  .elText = lvw.ListItems(i).SubItems(1)
              .elCreateSibling "DEP", True
                  .elText = lvw.ListItems(i).SubItems(3)
              .elCreateSibling "QTY", True
                  .elText = lvw.ListItems(i).SubItems(4)
              
                 .navUP
              .navUP
      Next
      .navUP
  End With

 GetDetailsXML = xMLDoc.docXML
Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORReq.GetDetailsXML"
End Function

Private Sub txtTitle_GotFocus()
    Me.txtEmail.Visible = True
    Me.txtAddress.Visible = True
    Me.lblAddress.Visible = True
    Me.lblEmail.Visible = True
End Sub



'=============
Private Sub PlaceOrder()
    On Error GoTo errHandler
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
Dim f As New frmCOPreview

    bIssue = False
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_CO_SIGN, , "Sign this order", DOCAPPROVAL, , , gSTAFFID) = False Then
               Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    If Not lngTPID > 0 Then 'A customer has not been specified
        
        Set oCust = New a_Customer
        oCust.BeginEdit
        oCust.InitializeNewCustomer enPrivate
        oCust.SetPhone txtPhone
        oCust.SetName txtName
        oCust.SetInitials Me.txtInitials
        oCust.SetTitle txtTitle
        oCust.SetCustomerTypeCasual
        oCust.SetControl txtPhone
        oCust.SetAccAcNo oPC.GetProperty("DefaultAccountingAccno")
        oCust.Addresses.FindByDescription("Default").SetEmail txtEmail
        oCust.Addresses.FindByDescription("Default").SetAddress Me.txtAddress
        If InStr(1, txtEmail, "@") > 0 Then
            oCust.SetDispatchMethod "M"
        End If
        If oCust.IsValid = False Then
            oCust.CancelEdit
            Set oCust = Nothing
            MsgBox "The customer cannot be saved, it is invalid. Please check, correct and try again.", vbInformation + vbOKOnly, "Can't do this"
            Exit Sub
        End If
        bAlternativeCustomerSelected = False
        oCust.LookforDuplicates
        If Not bAlternativeCustomerSelected Then
            oCust.ApplyEdit lngResult
        End If
        lngTPID = oCust.ID
    End If
    
    Set oCust = Nothing
    Set oCO = New a_CO
    oCO.BeginEdit
    If oCO.SetCustomer(lngTPID) Then
        oCO.OrderType = enNormalCO
        oCO.SetMemo FNS(txtItem)
        oCO.StaffID = gSTAFFID
        oCO.ORGUID = guid_OrderRequest
        For i = 1 To lvw.ListItems.Count
            Set oCOL = oCO.COLines.Add
            oCOL.BeginEdit
            oCOL.SetLineProduct , lvw.ListItems(i).text
            oCOL.SetQty CInt(val(StripToNumerics(lvw.ListItems(i).SubItems(4))))
            oCOL.SetDeposit CDbl(val(StripToNumerics(lvw.ListItems(i).SubItems(3)))) * CDbl(val(StripToNumerics(lvw.ListItems(i).SubItems(4)))) * oPC.Configuration.DefaultCurrency.Divisor
            oCOL.SetRef ""
            oCOL.DepositStatus = "P"
            If DateDiff("d", Date, oCOL.ETA) <= 1 Then
                oCOL.SetETA "2w"
            End If
            oCOL.ApplyEdit
        Next
        oCO.SetStatus stISSUED
        oCO.Post
        
        Screen.MousePointer = vbDefault
        MsgBox "Order placed", , "Status"
            
        f.component oCO.TRID, False
        Me.Hide
        f.Show
    Else
        oCO.CancelEdit
        MsgBox "Order not placed, customer not recognized.", , "Status"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmORReq.PlaceOrder"
End Sub

