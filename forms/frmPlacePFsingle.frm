VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmPlaceTransaction 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Place customer pro-forma invoice"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   9765
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New customer's details"
      ForeColor       =   &H8000000D&
      Height          =   1500
      Left            =   135
      TabIndex        =   14
      Top             =   1065
      Width           =   9345
      Begin VB.TextBox txtEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2085
         TabIndex        =   4
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   1005
         Width           =   3030
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   915
         Left            =   5355
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   435
         Width           =   3240
      End
      Begin VB.TextBox txtTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   285
         TabIndex        =   0
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   450
         Width           =   435
      End
      Begin VB.TextBox txtPhone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   285
         TabIndex        =   3
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   1005
         Width           =   1725
      End
      Begin VB.TextBox txtInitials 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   885
         TabIndex        =   1
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   450
         Width           =   1740
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2790
         TabIndex        =   2
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   450
         Width           =   2325
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2190
         TabIndex        =   20
         Top             =   795
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   5355
         TabIndex        =   19
         Top             =   195
         Width           =   1470
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   285
         TabIndex        =   18
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2820
         TabIndex        =   17
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Firstname or initials"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   900
         TabIndex        =   16
         Top             =   240
         Width           =   1770
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   90
         TabIndex        =   15
         Top             =   795
         Width           =   705
      End
   End
   Begin VB.TextBox txtSPInstr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   600
      Left            =   225
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4590
      Width           =   5655
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Search for existing customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   135
      Width           =   2925
   End
   Begin VB.CommandButton cmdPlaceTransaction 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Create pro-forma invoice"
      Enabled         =   0   'False
      Height          =   825
      Left            =   6375
      Picture         =   "frmPlacePFsingle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4425
      Width           =   1470
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   8070
      Picture         =   "frmPlacePFsingle.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4635
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   1440
      Left            =   150
      OleObjectBlob   =   "frmPlacePFsingle.frx":0714
      TabIndex        =   6
      Top             =   2760
      Width           =   9390
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "S&pecial instructions"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   135
      TabIndex        =   13
      Top             =   4350
      Width           =   1890
   End
   Begin VB.Label lblFound 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Found"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   3120
      TabIndex        =   12
      Top             =   255
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "or create new customer record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   -60
      TabIndex        =   11
      Top             =   705
      Width           =   3420
   End
End
Attribute VB_Name = "frmPlaceTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XA As XArrayDB
Dim tlBookclubs As z_TextList
Dim bOrderPlaced As Boolean
Dim lngTPID As Long
Dim strType As String
Dim WithEvents oCust As a_Customer
Attribute oCust.VB_VarHelpID = -1
Dim oTran As Object
Dim oLine As Object

Dim bAlternativeCustomerSelected As Boolean

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.SetMenu"
End Sub


Private Sub cmdPlaceTransaction_Click()
    On Error GoTo errHandler
    PlaceTransaction strType
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.cmdPlaceTransaction_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
    G1.Columns(4).Width = 0
    SaveLayout Me.G1, Me.Name

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.G1_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.oCust_PossibleDuplicates(pDuplicates)", pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
    HandleError
End Sub


Private Function ShowDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
Dim frm As frmDuplicateCustomers
Dim tmpCust As a_Customer
    
    Set frm = New frmDuplicateCustomers
    frm.component Me.txtName, pDuplicates
    Screen.MousePointer = vbDefault
    frm.Show vbModal
    If frm.SelectedCustomer > 0 Then
        Set tmpCust = New a_Customer
        oCust.CancelEdit
        oCust.Load frm.SelectedCustomer
        Unload frm
        bAlternativeCustomerSelected = True
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.ShowDuplicates(pDuplicates)", pDuplicates
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.ShowDuplicates(pDuplicates)", pDuplicates
End Function

Public Sub component(p As XArrayDB, pType As String)
    On Error GoTo errHandler
Dim i As Long
Dim j As Long
    
    Set XA = Nothing
    Set XA = New XArrayDB
    XA.ReDim 1, p.UpperBound(1), 1, p.UpperBound(2)
    For i = 1 To p.UpperBound(1)
        For j = 1 To p.UpperBound(2)
            XA(i, j) = p(i, j)
        Next j
    Next
   ' Set XA = p
    strType = pType
    If UCase(strType) = "INVOICE" Then
        Me.Caption = "Place customer invoice"
        cmdPlaceTransaction.Caption = "&Create invoice"
    ElseIf UCase(strType) = "PF" Then
        Me.Caption = "Place pro-forma invoice"
        cmdPlaceTransaction.Caption = "&Create pro-forma invoice"
    ElseIf UCase(strType) = "QUOTATION" Then
        Me.Caption = "Place pro-forma invoice"
        cmdPlaceTransaction.Caption = "&Create quotation"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.component(p,pType)", Array(p, pType)
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If Not bOrderPlaced Then
        If MsgBox("You have not taken an action. Do you wish to close this form?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub frmPlaceTransaction_Click()
    On Error GoTo errHandler
    G1.Update
        PlaceTransaction strType
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.frmPlaceTransaction_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub PlaceOnReserve()
    On Error GoTo errHandler
Dim i As Integer
Dim lngResult As Long
Dim bProductToOrder As Boolean
Dim OpenResult As Integer

    For i = 1 To XA.UpperBound(1)
        If XA(i, 4) > 0 Then
            bProductToOrder = True
        End If
    Next
    If bProductToOrder = False Then
        MsgBox "There are no non-zero quantities on this order. No action will be taken!", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If


    If MsgBox("You are placing on reserve for " & txtInitials & " " & txtName & "?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    If Not lngTPID > 0 Then 'A customer has not been specified
        
        Set oCust = New a_Customer
        oCust.BeginEdit
        oCust.SetPhone txtPhone
        oCust.SetName txtName
        oCust.SetInitials Me.txtInitials
        oCust.SetTitle txtTitle
        oCust.SetCustomerTypeCasual
        oCust.SetControl txtPhone
        oCust.ApplyEdit lngResult
        If lngResult = 22 Then  'There is already a customer with the same search phone value"
            MsgBox "The customer record has not been saved as there is already a record for this customer."
            Set oCust = Nothing
            Exit Sub
        End If
        lngTPID = oCust.ID
    End If


'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    For i = 1 To XA.UpperBound(1)
        oPC.COShort.execute "UPDATE tPRODUCT SET P_QTYRESERVED = P_QTYRESERVED + " & CLng(XA(i, 4)) & " WHERE P_ID = '" & XA(i, 7) & "'"
        oPC.COShort.execute "UPDATE tStoreP SET STP_QTYRESERVED = STP_QTYRESERVED + " & CLng(XA(i, 4)) & " WHERE STP_P_ID = '" & XA(i, 7) & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        oPC.COShort.execute "INSERT INTO tRM (RM_P_ID,RM_TP_ID,RM_QTY,RM_DATE,RM_TYPE) VALUES ('" & XA(i, 7) & "'," & lngTPID & "," & CLng(XA(i, 4)) & ",{d '" & Format(Date, "yyyy-mm-dd") & "'},'I')"
    Next
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    MsgBox "Placed on reserve", , "Status"
    Unload Me

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.PlaceOnReserve"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.PlaceOnReserve"
End Sub
Private Sub PlaceTransaction(strType As String)
    On Error GoTo errHandler
Dim i As Integer
Dim lngResult As Long
Dim bFound As Boolean
Dim bProductToInvoice As Boolean
Dim bZeroDeposit As Boolean
Dim bIssue As Boolean
Dim sDescription As String
Dim f As Form

    If strType = "PF" Then
        sDescription = "Pro-forma invoice"
        Set f = New frmInvoicePreview
    ElseIf strType = "INVOICE" Then
        sDescription = "Invoice"
        Set f = New frmInvoicePreview
    ElseIf strType = "QUOTATION" Then
        sDescription = "Quotation"
        Set f = New frmQuotationPreview
    End If
    bIssue = False
    bProductToInvoice = False
    bZeroDeposit = False
    For i = 1 To XA.UpperBound(1)
        If XA(i, 4) > 0 Then
            bProductToInvoice = True
        End If
        If XA(i, 5) < 200 Then
            bZeroDeposit = True
        End If
    Next
    If bProductToInvoice = False Then
        MsgBox "There are no non-zero quantities on this document. No action will be taken!", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    If MsgBox("You are placing " & sDescription & " for " & txtInitials & " " & txtName & "?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    If (oPC.GetProperty("IssueQuickPFs") = "TRUE" And strType = "PF") Or (oPC.GetProperty("IssueQuickInvoices") = "TRUE" And strType <> "PF") Then
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_CO_SIGN, , "Sign this " & sDescription, DOCAPPROVAL, , , gSTAFFID) = False Then
                Exit Sub
            Else
                bIssue = True
            End If
        Else
            If oTran.Status = stInProcess Then
                If MsgBox("Issue this " & sDescription & "?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                    Exit Sub
                Else
                    bIssue = True
                End If
            End If
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
        oCust.SetNote txtAddress
        oCust.SetAccAcNo oPC.GetProperty("DefaultAccountingAccno")
        oCust.Addresses.FindByDescription("Default").SetAddress Me.txtAddress
        oCust.Addresses.FindByDescription("Default").SetEmail txtEmail
        If InStr(1, txtEmail, "@") > 0 Then
            oCust.SetDispatchMethod "M"
        End If
        
        bAlternativeCustomerSelected = False
        oCust.LookforDuplicates
        If Not bAlternativeCustomerSelected Then
            oCust.ApplyEdit lngResult
        End If
        lngTPID = oCust.ID
    End If
    
    
    Set oTran = Nothing
    If strType = "PF" Then
        Set oTran = New a_Invoice
    ElseIf strType = "INVOICE" Then
        Set oTran = New a_Invoice
    ElseIf strType = "QUOTATION" Then
        Set oTran = New a_QU
    End If

    oTran.BeginEdit
    oTran.SetCustomer lngTPID
   ' oTran.setMemo FNS(txtAddress)
    oTran.VATable = True
    oTran.StaffID = gSTAFFID
    oTran.SetMemo FNS(txtSPInstr)

    For i = 1 To XA.UpperBound(1)
        If strType = "PF" Then
            Set oLine = oTran.InvoiceLines.Add
        ElseIf strType = "INVOICE" Then
            Set oLine = oTran.InvoiceLines.Add
        ElseIf strType = "QUOTATION" Then
            Set oLine = oTran.QuoteLines.Add
        End If
        oLine.BeginEdit
        oLine.PID = XA(i, 8)
       oLine.Price = XA(i, 9)
        oLine.VATRate = oPC.Configuration.VATRate
      '  oLine.code = XA(i, 7)
        oLine.SetQty XA(i, 4)
      '  oLine.SetDeposit XA(i, 5)
        oLine.SetRef XA(i, 6)
        oLine.ApplyEdit
    Next
    If bIssue Then
        oTran.SetStatus stCOMPLETE
    Else
        oTran.SetStatus stInProcess
    End If
    If strType = "PF" Then oTran.SetProforma
    oTran.Post
    
    Screen.MousePointer = vbDefault
    MsgBox sDescription & " issued", , "Status"
    f.component oTran.QuoteID
    f.Show
    Unload Me
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.cmdPlaceTransaction_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.PlaceTransaction(strType)", strType
End Sub
'Private Sub PlaceOrder()
'    On Error GoTo errHandler
'Dim i As Integer
'Dim oCOL As a_COL
'Dim lngResult As Long
'Dim bFound As Boolean
'Dim bProductToOrder As Boolean
'Dim bZeroDeposit As Boolean
'
'    bProductToOrder = False
'    bZeroDeposit = False
'    For i = 1 To XA.UpperBound(1)
'        If XA(i, 4) > 0 Then
'            bProductToOrder = True
'        End If
'        If XA(i, 5) < 200 Then
'            bZeroDeposit = True
'        End If
'    Next
'    If bProductToOrder = False Then
'        MsgBox "There are no non-zero quantities on this order. No action will be taken!", vbCritical + vbOKOnly, "Warning"
'        Exit Sub
'    End If
'    If MsgBox("You are placing an order for " & txtInitials & " " & txtName & "?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
'        Exit Sub
'    End If
'    If bZeroDeposit = True Then
'        If MsgBox("You are placing an order with a small deposit. Confirm?", vbOKCancel + vbQuestion, "Warning") = vbCancel Then
'            Exit Sub
'        End If
'    End If
'    Screen.MousePointer = vbHourglass
'    If Not lngTPID > 0 Then 'A customer has not been specified
'
'        Set oCust = New a_Customer
'        oCust.BeginEdit
'        oCust.InitializeNewCustomer enPrivate
'        oCust.SetPhone txtPhone
'        oCust.SetName txtName
'        oCust.SetInitials Me.txtInitials
'        oCust.SetTitle txtTitle
'        oCust.SetCustomerTypeCasual
'        oCust.SetControl txtPhone
'
'        bAlternativeCustomerSelected = False
'        oCust.LookforDuplicates
'        If Not bAlternativeCustomerSelected Then
'            oCust.ApplyEdit lngResult
'        End If
'        lngTPID = oCust.ID
'    End If
'
'
'    Set oCust = Nothing
'    Set oCO = New a_CO
'    oCO.BeginEdit
'    oCO.SetCustomer lngTPID
'    oCO.OrderType = enNormalCO
'    oCO.setMemo FNS(txtSPInstr)
'    For i = 1 To XA.UpperBound(1)
'        Set oCOL = oCO.COLines.Add
'        oCOL.BeginEdit
'        oCOL.SetLineProduct XA(i, 7)
'        oCOL.SetQty XA(i, 4)
'        oCOL.SetDeposit XA(i, 5)
'        oCOL.SetRef XA(i, 6)
'        oCOL.DepositStatus = "P"
'        If DateDiff("d", Date, oCOL.ETA) <= 1 Then
'            oCOL.SetETA "2w"
'        End If
'        oCOL.ApplyEdit
'    Next
'    oCO.SetStatus stISSUED
'    oCO.Post
'
'    Screen.MousePointer = vbDefault
'    MsgBox "Order placed", , "Status"
'    Unload Me
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.PlaceOrder"
'End Sub

Private Sub cmdSearch_Click()
    On Error GoTo errHandler
Dim frmC As frmBrowseCustomers2

    Set frmC = New frmBrowseCustomers2
    frmC.Show vbModal
    If frmC.CustomerID = 0 Then
        txtName = ""
        txtPhone = ""
        Unload frmC
        Set frmC = Nothing
        lblFound.Visible = False
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
    Me.txtEmail = oCust.Addresses.DefaultAddress.EMail
    Me.txtAddress = oCust.Addresses.DefaultAddress.AddressMailing
    Me.G1.ReBind
    lngTPID = frmC.CustomerID
    Set frmC = Nothing
    lblFound.Visible = (lngTPID <> 0)
    cmdPlaceTransaction.Enabled = (lngTPID <> 0) And XA.Count(1) > 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    cmdPlaceTransaction.Enabled = XA.Count(1) > 0 And lngTPID > 0 And lblFound.Visible = True
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    G1.Array = XA
    G1.ReBind
    SetGridLayout Me.G1, Me.Name
    Me.cmdPlaceTransaction.Caption = "Place pro-forma"
    Me.Caption = "Quick pro-forma"
    Me.Height = 6200
    Me.Width = 10000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.Form_Load", , EA_NORERAISE
    HandleError
End Sub
'Private Sub cboBC_SelectionChanged()
'    On Error GoTo errHandler
'    lngTPID = tlBookclubs.key(cboBC.Items.CellCaption(cboBC.Items.SelectedItem, 0))
'    lblFound.Visible = (lngTPID <> 0)
'    txtName = cboBC.Items.CellCaption(cboBC.Items.SelectedItem, 0)
'    frmPlaceTransaction.Enabled = (lngTPID <> 0) And XA.Count(1) > 0
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.cboBC_SelectionChanged", , EA_NORERAISE
'    HandleError
'End Sub

'Sub SetupCboBC()
'    On Error GoTo errHandler
'    cboBC.BeginUpdate
'    cboBC.WidthList = 330
'    cboBC.HeightList = 142
'    cboBC.AllowSizeGrip = True
'    cboBC.AutoDropDown = True
'
'    cboBC.Columns.Add "Name"
'    cboBC.Columns.Add "Phone"
'
'    cboBC.Columns(0).Width = 190
'    cboBC.Columns(1).Width = 120
'    cboBC.BackColorLock = Me.BackColor
'    cboBC.EndUpdate
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.SetupCboBC"
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlBookclubs = Nothing
    Set XA = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Label2_Click()
    On Error GoTo errHandler

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.Label2_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.Label2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Text1_Change()
    On Error GoTo errHandler

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.Text1_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.Text1_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_AfterDelete()
    On Error GoTo errHandler
    If XA.Count(1) = 0 Then Me.cmdPlaceTransaction.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.G1_AfterDelete", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.G1_AfterDelete", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
    G1.InsertMode = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.G1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.G1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim lngTmp As Long
    If Not ConvertToLng(G1.text, lngTmp) Then
        Cancel = True
    ElseIf CLng(G1.text) < 0 Then
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.G1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.G1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub G1_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_LostFocus()
    On Error GoTo errHandler
    G1.Update
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.G1_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.G1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_Change()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To XA.UpperBound(1)
        XA(i, 6) = Left(txtName, 10)
    Next
    G1.ReBind
    cmdPlaceTransaction.Enabled = Len(txtPhone) > 5 And Len(txtName) > 1 And XA.Count(1) > 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.txtName_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.txtName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Change()
    On Error GoTo errHandler
    cmdPlaceTransaction.Enabled = Len(txtPhone) > 5 And Len(txtName) > 1 And XA.Count(1) > 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.txtPhone_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.txtPhone_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim bFound As Boolean
Dim oC As New c_Customer
Dim strMsg As String
Dim dCustomer As d_Customer
    If Len(txtPhone) < 6 Then Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceTransaction.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceTransaction.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

