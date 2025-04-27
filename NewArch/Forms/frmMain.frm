VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Receive stock"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15045
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":038A
   ScaleHeight     =   10800
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvw 
      Height          =   2610
      Left            =   315
      TabIndex        =   7
      Top             =   5655
      Visible         =   0   'False
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   4604
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account No."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Supplier name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   255
      TabIndex        =   0
      Text            =   "txtInput"
      Top             =   9450
      Width           =   4665
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   2880
      Left            =   225
      TabIndex        =   14
      Top             =   2220
      Width           =   14130
      _ExtentX        =   24924
      _ExtentY        =   5080
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Disc."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Additional charges (e.g. freight and insurance)"
      ForeColor       =   &H00915A48&
      Height          =   225
      Left            =   240
      TabIndex        =   20
      Top             =   1860
      Width           =   4080
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total quantity items"
      ForeColor       =   &H00915A48&
      Height          =   225
      Left            =   2355
      TabIndex        =   19
      Top             =   1530
      Width           =   1965
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total invoice value"
      ForeColor       =   &H00915A48&
      Height          =   225
      Left            =   2355
      TabIndex        =   18
      Top             =   1200
      Width           =   1965
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice date"
      ForeColor       =   &H00915A48&
      Height          =   225
      Left            =   2355
      TabIndex        =   17
      Top             =   870
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice ref."
      ForeColor       =   &H00915A48&
      Height          =   225
      Left            =   2355
      TabIndex        =   16
      Top             =   540
      Width           =   1965
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
      ForeColor       =   &H00915A48&
      Height          =   225
      Left            =   2355
      TabIndex        =   15
      Top             =   210
      Width           =   1965
   End
   Begin VB.Label lblSupplierInvoiceAdditionalCharges 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4485
      TabIndex        =   13
      Top             =   1830
      Width           =   1845
   End
   Begin VB.Label lblSupplierInvoiceQty 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4485
      TabIndex        =   12
      Top             =   1500
      Width           =   1845
   End
   Begin VB.Label lblSupplierInvoiceValue 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4485
      TabIndex        =   11
      Top             =   1170
      Width           =   1845
   End
   Begin VB.Label lblSupplierInvoiceDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4485
      TabIndex        =   10
      Top             =   840
      Width           =   1845
   End
   Begin VB.Label lblSupplierInvoiceRef 
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4485
      TabIndex        =   9
      Top             =   510
      Width           =   7515
   End
   Begin VB.Label lblMessages 
      BackColor       =   &H80000009&
      Height          =   810
      Left            =   10725
      TabIndex        =   8
      Top             =   9105
      Width           =   4170
   End
   Begin VB.Label lblSupplierName 
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4485
      TabIndex        =   6
      Top             =   180
      Width           =   7515
   End
   Begin VB.Label lblState 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   10410
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   -30
      TabIndex        =   4
      Top             =   10095
      Width           =   360
   End
   Begin VB.Label SB 
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
      ForeColor       =   &H8000000D&
      Height          =   720
      Left            =   225
      TabIndex        =   3
      Top             =   9960
      Width           =   14775
   End
   Begin VB.Label lblPrompt 
      Height          =   345
      Left            =   225
      TabIndex        =   2
      Top             =   8595
      Width           =   4665
   End
   Begin VB.Label lblInput 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   435
      TabIndex        =   1
      Top             =   9135
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim itest As Integer
Dim res As Boolean

Dim bShiftDown As Boolean

Dim WithEvents zStateManager As z_StateManagement
Attribute zStateManager.VB_VarHelpID = -1

Public enNewState As enState

Dim lstItem As ListItem
Dim iRow As Integer
Dim strRaw As String
Dim strPrefix As String
Dim strSuffix As String


'EVENTS ---------------------------------------
Public Sub oGRN_TotalChange(lngTotal As String, lngTotalForeign As String, strQtyTotal As String)
End Sub

Public Sub zStateManager_Messages(msg As String)
    lblMessages.Caption = msg
End Sub
Public Sub zStateManager_PresentGRN()
Dim i As Integer
    With zStateManager.GRN
        lblSupplierName.Caption = .Supplier.NameAndCode(100)
        lblSupplierInvoiceRef.Caption = .SupplierInvoiceRef
        lblSupplierInvoiceDate.Caption = .SupplierInvoiceDateF
        lblSupplierInvoiceQty.Caption = .BatchQtyTotalF
        lblSupplierInvoiceValue.Caption = .BatchTotalF
        lblSupplierInvoiceAdditionalCharges.Caption = .BatchTotalExtrasF
        
        lvwDetails.ListItems.Clear
        For i = 1 To .DeliveryLines.Count
            Set lstItem = lvwDetails.ListItems.Add(, CStr(i) & "k")
            lstItem.Text = val(lstItem.Key)
            lstItem.SubItems(1) = .DeliveryLines(i).CodeF
            lstItem.SubItems(2) = .DeliveryLines(i).Title
            lstItem.SubItems(3) = .DeliveryLines(i).PriceF(False)
            lstItem.SubItems(4) = .DeliveryLines(i).QtyFirmF
            lstItem.SubItems(5) = .DeliveryLines(i).DiscountF
            lstItem.SubItems(6) = .DeliveryLines(i).PLessDiscExtF(False)
            lstItem.SubItems(9) = .DeliveryLines(i).EAN
        Next i
    End With
End Sub
Public Sub zStateManager_PresentInvoiceDate(strDate As String)
    lblSupplierInvoiceDate.Caption = strDate
End Sub
Public Sub zStateManager_PresentSupplierName(strName As String)
    lblSupplierName.Caption = strName
End Sub
Public Sub zStateManager_PresentInvoiceRef(strRef As String)
    lblSupplierInvoiceRef.Caption = strRef
End Sub
Public Sub zStateManager_PresentInvoiceQuantity(strQty As String)
    lblSupplierInvoiceQty.Caption = strQty
End Sub
'=====
Public Sub zStateManager_PresentInvoiceLinePrice(strPrice As String)
    lblSupplierInvoiceQty.Caption = strPrice
End Sub
Public Sub zStateManager_PresentInvoiceLineQty(strQty As String)
    lblSupplierInvoiceQty.Caption = strQty
End Sub
Public Sub zStateManager_PresentInvoiceLineDiscount(strDiscount As String)
    lblSupplierInvoiceQty.Caption = strDiscount
End Sub
            
          
'-----------------------------------------------
Public Sub zStateManager_ShowBrowsedGRNs(BrowsedGRNs As c_DELs)
Dim lstItem As ListItem
Dim i As Long

    lvw.ColumnHeaders.Clear
    lvw.ColumnHeaders.Add 1, , "No.", 500, 0
    lvw.ColumnHeaders.Add 2, , "Started", 1000, 0
    lvw.ColumnHeaders.Add 3, , "Code", 1200, 0
    lvw.ColumnHeaders.Add 4, , "Supplier", 1200, 0
    lvw.ListItems.Clear
    For i = 1 To BrowsedGRNs.Count
        Set lstItem = lvw.ListItems.Add
        With BrowsedGRNs.Item(i)
            lstItem.Text = CStr(i)
            lstItem.SubItems(1) = .DocDateF
            lstItem.SubItems(2) = .Ref & .StaffNameB
            lstItem.SubItems(3) = .TPName
            lstItem.Key = CStr(.TRID) & CStr("k")
        End With

    Next i
    lvw.Visible = True
    setInputBox "", "", "", True
    lblInput.Caption = "Select document"
    Stat "Enter row number of required document."
    AutoSelect txtInput
    txtInput.BackColor = RGB(230, 250, 210)
    ClearColours
    
End Sub

'----------------------------------------------
Private Sub LoadSupplierSelectList()
Dim lstItem As ListItem
Dim i As Long
    lvw.ColumnHeaders.Clear
    lvw.ColumnHeaders.Add 1, , "Acno"
    lvw.ColumnHeaders.Add 2, , "Name"
    lvw.ColumnHeaders.Add 3, , "sss"
    lvw.ListItems.Clear
    For i = 1 To zStateManager.SupplierCollection.Count
        Set lstItem = lvw.ListItems.Add
        With zStateManager.SupplierCollection.Item(i)
            lstItem.Text = CStr(i)
            lstItem.SubItems(1) = .AcNo
            lstItem.SubItems(2) = .Name
            lstItem.Key = CStr(.ID) & CStr("k")
        End With

    Next i
    lvw.Visible = True
End Sub


Public Sub Start()
    iRow = 0
    Set zStateManager = New z_StateManagement
    zStateManager.Start
    PrepareForm
End Sub
Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    
    
    bShiftDown = (Shift = 1)
    If KeyCode = 13 Then
        strRaw = UCase(Trim$(txtInput))
        enNewState = eState_Null
        If LenB(strRaw) = 0 Then
            enNewState = eState_Invalid
            Exit Sub
        End If
        If SeparateInput(strRaw, strPrefix, strSuffix) = False Then 'We cannot process this input
            enNewState = zStateManager.PresentState
            Exit Sub
        End If
        
        Select Case zStateManager.PresentState
        Case eState_SuppliersFound
            enNewState = zStateManager.GetNewState(val(lvw.ListItems(CLng(strRaw)).Key))
        Case eState_Browse
            enNewState = zStateManager.GetNewState(val(lvw.ListItems(CLng(strRaw)).Key))
        Case eState_LineIdentifier
            If UCase(strPrefix) = "L" Then
                enNewState = zStateManager.GetNewState(val(lvwDetails.ListItems(CLng(strSuffix)).SubItems(9)), strPrefix, strSuffix)
            Else
                enNewState = zStateManager.GetNewState(strRaw, strPrefix, strSuffix)
            End If
        Case Else
            enNewState = zStateManager.GetNewState(txtInput)
        End Select
       
        If zStateManager.PresentState = eState_End Then
            If MsgBox("You want to close the application. Confirm.", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
                enNewState = eState_End  'not right fix later
            End If
            res = zStateManager.SaveAndClose
            If res Then
                Unload Me
            Else
                MsgBox "Unable to close GRN", vbInformation + vbOKOnly, "Can't close application"
            End If
            Exit Sub
        End If
        
        PrepareForm
        
    ElseIf KeyCode = 40 And zStateManager.PresentState = eState_SuppliersFound Then   'Keydown
        If lvw.GetFirstVisible.Index <= lvw.ListItems.Count - 3 Then
            lvw.ListItems(lvw.GetFirstVisible.Index + 3).EnsureVisible
        End If
    
    ElseIf KeyCode = 38 Then   ' Key up
        If lvw.GetFirstVisible.Index <> 1 Then
            lvw.ListItems(lvw.GetFirstVisible.Index - 1).EnsureVisible
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.txtInput_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE, , "input value", Array(txtInput)
    HandleError
End Sub


Private Sub PrepareForm()
    On Error GoTo errHandler
    txtInput.BackColor = vbWhite
    Select Case zStateManager.PresentState
        Case eState_Start
            setInputBox "", "", "", True
            lblInput.Caption = "Identify supplier"
            Stat "Enter account number or the first few letters of the start of the supplier name. Use '.' to browse."
            AutoSelect txtInput
            txtInput.BackColor = RGB(230, 250, 210)
            ClearColours
            lblSupplierName.BackColor = vbGreen
        
        Case eState_Browse
            setInputBox "", "", "", True
            lblInput.Caption = "Select document"
            Stat "Enter row number of required document."
            AutoSelect txtInput
            txtInput.BackColor = RGB(230, 250, 210)
            ClearColours
        
        Case eState_SuppliersFound
            LoadSupplierSelectList
            setInputBox "", "", "", True
            lblInput.Caption = "More than one supplier found"
            Stat "Select the supplier from the list by typing the identifying number"
            AutoSelect txtInput
            lblSupplierName.BackColor = vbGreen
        
        Case eState_SupplierInvoiceRef
            lvw.ListItems.Clear
            lvw.Visible = False
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                setInputBox zStateManager.GRN.SupplierInvoiceRef, "", "", True
            End If
            lblInput.Caption = "Supplier invoice reference"
            Stat "Enter the supplier's invoice reference"
            AutoSelect txtInput
            ClearColours
            lblSupplierInvoiceRef.BackColor = vbGreen
            
        Case eState_SupplierInvoiceDate
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                setInputBox zStateManager.GRN.SupplierInvoiceDateF, "", "", True
            End If
            lblInput.Caption = "Supplier invoice date"
            Stat "Enter the date of the supplier's invoice in format ddmmyyyy"
            AutoSelect txtInput
            ClearColours
            lblSupplierInvoiceDate.BackColor = vbGreen
        
        Case eState_SupplierInvoiceValue
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                setInputBox zStateManager.GRN.BatchTotal, "", "", True
            End If
            lblInput.Caption = "Supplier invoice value"
            Stat "Enter the value of the supplier's invoice in its own currency e.g. 2100.22"
            AutoSelect txtInput
            ClearColours
            lblSupplierInvoiceValue.BackColor = vbGreen
      
        Case eState_SupplierInvoiceQuantity
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                setInputBox zStateManager.GRN.BatchQtyTotal, "", "", True
            End If
            lblInput.Caption = "Supplier invoice qty"
            Stat "Enter the number of items on supplier's invoice."
            AutoSelect txtInput
            ClearColours
            lblSupplierInvoiceQty.BackColor = vbGreen
        
        Case eState_SupplierInvoiceAdditionalCharges
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                setInputBox zStateManager.GRN.BatchTotalExtras, "", "", True
            End If
            lblInput.Caption = "Supplier invoice additional charges"
            Stat "Enter the value of the extra charges on the supplier's invoice"
            AutoSelect txtInput
            ClearColours
            lblSupplierInvoiceAdditionalCharges.BackColor = vbGreen
        
        Case eState_LineIdentifier
            If Not zStateManager.GRNL Is Nothing Then
                If zStateManager.GRNL.IsEditing Then zStateManager.GRNL.ApplyEdit
            End If
            Me.lblSupplierInvoiceAdditionalCharges.Caption = zStateManager.GRN.BatchTotalExtrasF
            If Not zStateManager.GRNL Is Nothing Then
                lstItem.SubItems(4) = zStateManager.GRNL.QtyFirmF
            End If
            
            setInputBox "", "", "", True
            lblInput.Caption = "Line item code"
            Stat "Enter the line item code e.g. ISBN number or EAN number to add a new line or 'Ln where n is the line number you want to edit."
            AutoSelect txtInput
            ClearColours
            
        Case eState_LinePrice
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                If Not zStateManager.IsLineEditing Then
                    iRow = iRow + 1
                    Set lstItem = lvwDetails.ListItems.Add(, CStr(iRow) & "k")
                    lstItem.Text = lstItem.Key
                    lstItem.SubItems(1) = zStateManager.GRNL.CodeF
                    lstItem.SubItems(2) = zStateManager.GRNL.Title
                    setInputBox "", "", "", True
                Else
                    setInputBox zStateManager.GRNL.Price(False), "", "", True
                End If
            End If
            lblInput.Caption = "Line item price"
            Stat "Enter the line item price as shown on the invoice e.g. 149.99"
            AutoSelect txtInput
      
        Case eState_LineQuantity
            If zStateManager.GRN Is Nothing Then
                setInputBox "", "", "", True
            Else
                setInputBox zStateManager.GRNL.QtyFirm, "", "", True
            End If
       '     lstItem.SubItems(3) = zStateManager.GRNL.PriceF(False)
            
            setInputBox "", "", "", True
            lblInput.Caption = "Line quantity"
            Stat "Enter the line quantity"
            AutoSelect txtInput
      
        Case eState_LinePrice
       '     lstItem.SubItems(3) = zStateManager.GRNL.PriceF(False)
            
            setInputBox "", "", "", True
            lblInput.Caption = "Line quantity"
            Stat "Enter the line quantity"
            AutoSelect txtInput
      
    End Select
    Me.lblState.Caption = InterpretState
    
    Exit Sub
errHandler:    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PrepareForm"
End Sub
Private Function InterpretState() As String
    On Error GoTo errHandler
    Select Case zStateManager.PresentState
    Case 0
        InterpretState = "eStart"
    Case 1
        InterpretState = "eSale"
    Case 2
        InterpretState = "eTitle"
    Case 3
        InterpretState = "eQty"
    Case 4
        InterpretState = "eDiscount"
    Case 5
        InterpretState = "ePrice"
    Case 6
        InterpretState = "elogin"
    Case 7
        InterpretState = "ePaymentAmt"
    Case 8
        InterpretState = "eConfirmation"
    Case 9
        InterpretState = "eSearchCustomer"
    Case 20
        InterpretState = "eXTerminate"
    Case 21
        InterpretState = "eZTerminate"
    Case 22
        InterpretState = "eRebuildIndexes"
    Case 23
        InterpretState = "eHelp"
    Case 24
        InterpretState = "ecancelsale"
    Case 25
        InterpretState = "eCashRefund"
    Case 26
        InterpretState = "ePriceCashRefund"
    Case 27
        InterpretState = "eQtyCashRefund"
    Case 28
        InterpretState = "eDiscountCashRefund"
    Case 29
        InterpretState = "eConfirmationCashrefund"
    Case 30
        InterpretState = "eVoid"
    Case 31
        InterpretState = "eReviewExchanges"
    Case 32
        InterpretState = "eShowExchange"
    Case 33
        InterpretState = "eOPenDrawer"
    Case 34
        InterpretState = "eStatus"
    Case 35
        InterpretState = "eState_Null"
    Case 36
        InterpretState = "ePrevious"
    Case 37
        InterpretState = "eDelete"
    Case 38
        InterpretState = "eDeletePayment"
    Case 39
        InterpretState = "eShowvoucherType"
    Case 40
        InterpretState = "eOperatorsReport"
    Case 41
        InterpretState = "eCreditNote"
    Case 42
        InterpretState = "ePriceCreditNote"
    Case 43
        InterpretState = "eDiscountCreditNote"
    Case 44
        InterpretState = "eQtyCreditNote"
    Case 45
        InterpretState = "eRefundDeposit"
    Case 46
        InterpretState = "eConfirmationRefundDeposit"
    Case 47
        InterpretState = "eSearchCustomerfordepositRefund"
    Case 48
        InterpretState = "eRefundType_Cash"
    Case 49
        InterpretState = "eRefundType_Creditcard"
    Case 50
        InterpretState = "eSearchCustomerforAppro"
    Case 51
        InterpretState = "eAppro"
    Case 52
        InterpretState = "ePriceAppro"
    Case 53
        InterpretState = "eDiscountAppro"
    Case 54
        InterpretState = "eQtyAppro"
    Case 55
        InterpretState = "eConfirmationAppro"
    Case 56
        InterpretState = "eApproReturn"
    Case 57
        InterpretState = "eSearchCustomerforApproReturn"
    Case 58
        InterpretState = "ePettyCash"
    Case 59
        InterpretState = "ePettyCashAmt"
    Case 60
        InterpretState = "ePettyCashConfirmation"
    Case 61
        InterpretState = "ePettyCashReason"
    Case 62
        InterpretState = "ePettyCashCredit"
    Case 63
        InterpretState = "ePettyCashCreditConfirmation"
    Case 64
        InterpretState = "ePettyCashCreditAmt"
    Case 65
        InterpretState = "eSearchCustomerfordeposit"
    Case 66
        InterpretState = "eDiscountDeposit"
    Case 67
        InterpretState = "eSelectDepositLineRef"
    Case 68
        InterpretState = "eSelectDepositLine"
    Case 69
        InterpretState = "eSelectDepositLineForRefund"
    Case 70
        InterpretState = "ePriceDeposit"
    Case 71
        InterpretState = "eQtyDeposit"
    Case 72
        InterpretState = "eInvoice"
    Case 73
        InterpretState = "eInvoiceno"
    Case 74
        InterpretState = "eInvoiceMode"
    Case 75
        InterpretState = "eConfirmationInvoiceCollection"
    Case 76
        InterpretState = "eConfirmationDepositRefund"
    Case 77
        InterpretState = "eConfirmationDeposit"
    Case 78
        InterpretState = "eConfirmationCreditNote"
    Case 89
        InterpretState = "eCollect"
    Case 90
        InterpretState = "ePaymentType_Cash"
    Case 91
        InterpretState = "ePaymentType_Cheque"
    Case 92
        InterpretState = "ePaymentType_CreditCard"
    Case 93
        InterpretState = "ePaymentType_CreditVoucher"
    Case 94
        InterpretState = "ePaymentType_CreditVoucherRef"
    Case 95
        InterpretState = "ePaymentType_voucher"
    Case 96
        InterpretState = "ePaymentType_ChequeRef"
    Case 97
        InterpretState = "ePaymentType_CreditVoucherRef"
    Case 98
        InterpretState = "ePaymentType_voucherRef"
    Case 99
        InterpretState = "ePaymentType_RedeemDeposit"
    Case 100
        InterpretState = "eRefundType_CreditVoucher"
    End Select

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.InterpretState"
End Function



Private Sub Stat(msg As String)
    On Error GoTo errHandler
    SB.Caption = msg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Stat(msg)", msg
End Sub

Private Sub SetTip(pMsg As String)
    On Error GoTo errHandler
    lblInput.Caption = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SetTip(pMsg)", pMsg
End Sub
Private Sub setInputBox(pText As String, pPasswordChar As String, pChange As String, bAutoSelect As Boolean)
    On Error GoTo errHandler
    txtInput = pText
    txtInput.PasswordChar = pPasswordChar
    If bAutoSelect Then
        AutoSelect txtInput
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.setInputBox(pText,pPasswordChar,pChange,bAutoSelect)", Array(pText, _
         pPasswordChar, pChange, bAutoSelect)
End Sub

Public Sub AutoSelect(CTR As Control)
    On Error GoTo errHandler
    With CTR
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.AutoSelect(CTR)", CTR
End Sub

Private Function SeparateInput(pRaw As String, pPrefix As String, pSuffix As String) As Boolean
    On Error GoTo errHandler
Dim i As Integer
Dim iMax As Integer
Dim c As String
Dim bAlpha As Boolean

    If IsDate(pRaw) Then
        SeparateInput = True
        Exit Function
    End If
    If IsNumeric(pRaw) Then
        SeparateInput = True
        Exit Function
    End If

    pPrefix = ""
    pSuffix = ""
    SeparateInput = True
    iMax = Len(pRaw)
    If InStr(1, pRaw, ",") > 0 Then  'there are commas in the string meaning a multiple selection
        For i = 1 To iMax
            c = Mid(pRaw, i, 1)
            If Not (IsNumeric(c) Or c = ",") Then
                SeparateInput = False
                Exit Function
            End If
        Next
        pSuffix = pRaw
    Else
        bAlpha = True
        If iMax > 9 Then
            SeparateInput = True
            pPrefix = pRaw
            pSuffix = pRaw
            Exit Function
        End If
        If Left(pRaw, 1) = "#" Then
            SeparateInput = True
            pPrefix = pRaw
            pSuffix = pRaw
            Exit Function
        End If
        If (zStateManager.PresentState <> eState_Start) Or UCase(Left(pRaw, 1)) = "D" Or UCase(Left(pRaw, 1)) = "V" Then
            bAlpha = Not IsNumeric(Left(pRaw, 1))
            i = 1
            If bAlpha Then
                Do While Not IsNumeric(Mid(pRaw, i, 1)) And i <= iMax
                    c = Mid(pRaw, i, 1)
                    pPrefix = pPrefix & c
                    i = i + 1
                Loop
                Do While IsNumeric(Mid(pRaw, i, 1)) And i <= iMax
                    c = Mid(pRaw, i, 1)
                    pSuffix = pSuffix & c
                    i = i + 1
                Loop
            Else
                Do While IsNumeric(Mid(pRaw, i, 1)) And i <= iMax
                    c = Mid(pRaw, i, 1)
                    pPrefix = pPrefix & c
                    i = i + 1
                Loop
                Do While Not IsNumeric(Mid(pRaw, i, 1)) And i <= iMax
                    c = Mid(pRaw, i, 1)
                    pSuffix = pSuffix & c
                    i = i + 1
                Loop
            End If
        Else
            SeparateInput = True
            pPrefix = pRaw
            pSuffix = pRaw
        
        End If
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SeparateInput(pRaw,pPrefix,pSuffix)", Array(pRaw, pPrefix, pSuffix)
End Function

Private Sub ClearColours()
    lblSupplierName.BackColor = vbWhite
    lblSupplierInvoiceRef.BackColor = vbWhite
    lblSupplierInvoiceDate.BackColor = vbWhite
    lblSupplierInvoiceValue.BackColor = vbWhite
    lblSupplierInvoiceQty.BackColor = vbWhite
    lblSupplierInvoiceAdditionalCharges.BackColor = vbWhite
End Sub
