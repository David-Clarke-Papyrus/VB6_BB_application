VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmPlaceCO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Place customer order"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New customer's details"
      ForeColor       =   &H8000000D&
      Height          =   1905
      Left            =   4260
      TabIndex        =   17
      Top             =   1005
      Width           =   5250
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
         Top             =   465
         Width           =   435
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
      Begin VB.TextBox txtInitials 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   885
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "Enter product code, reference A/C/ no. or start of customer name. Hit ENTER to fetch."
         Top             =   480
         Width           =   1740
      End
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
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "&Address"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2325
         TabIndex        =   0
         Top             =   795
         Width           =   1470
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   45
         TabIndex        =   23
         Top             =   1305
         Width           =   705
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   285
         TabIndex        =   21
         Top             =   270
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
         TabIndex        =   20
         Top             =   270
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
         TabIndex        =   19
         Top             =   270
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
         Left            =   60
         TabIndex        =   18
         Top             =   810
         Width           =   705
      End
   End
   Begin VB.TextBox txtSPInstr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000D&
      Height          =   600
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4695
      Width           =   5655
   End
   Begin VB.CommandButton cmdClearBC 
      BackColor       =   &H00C4BCA4&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1035
      Width           =   300
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
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   165
      Width           =   2925
   End
   Begin VB.CommandButton cmdPlaceOrder 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Place order"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6390
      Picture         =   "frmPlaceCO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4635
      Width           =   1470
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   8070
      Picture         =   "frmPlaceCO.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4635
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   1440
      Left            =   150
      OleObjectBlob   =   "frmPlaceCO.frx":0714
      TabIndex        =   22
      Top             =   2955
      Width           =   9390
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboBC 
      Height          =   315
      Left            =   135
      OleObjectBlob   =   "frmPlaceCO.frx":535F
      TabIndex        =   24
      Top             =   270
      Width           =   3165
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "or"
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
      Left            =   3180
      TabIndex        =   16
      Top             =   210
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "S&pecial instructions"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   165
      TabIndex        =   15
      Top             =   4485
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
      Left            =   7380
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   3780
      X2              =   3780
      Y1              =   540
      Y2              =   2235
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
      Left            =   5175
      TabIndex        =   12
      Top             =   750
      Width           =   3420
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Book club"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   390
      TabIndex        =   11
      Top             =   615
      Width           =   1350
   End
End
Attribute VB_Name = "frmPlaceCO"
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
Dim oCO As a_CO
Dim bAlternativeCustomerSelected As Boolean
Dim oSQL As New z_SQL

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.mnuSaveLayout"
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
    ErrorIn "frmPlaceCO.SetMenu"
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.G1_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub oCust_PossibleDuplicates(pDuplicates As c_Customer)
    On Error GoTo errHandler
    ShowDuplicates pDuplicates


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.oCust_PossibleDuplicates(pDuplicates)", pDuplicates
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.oCust_PossibleDuplicates(pDuplicates)", pDuplicates, EA_NORERAISE
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
 '       Set Forms(0).frmMainCustomerPreview = Nothing
 '       Set Forms(0).frmMainCustomerPreview = New frmCustomerPreview
        Set tmpCust = New a_Customer
        oCust.CancelEdit
        oCust.Load frm.SelectedCustomer
 '       oCust.BeginEdit
 '       Forms(0).frmMainCustomerPreview.Component tmpCust
        Unload frm
        bAlternativeCustomerSelected = True
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.ShowDuplicates(pDuplicates)", pDuplicates
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.ShowDuplicates(pDuplicates)", pDuplicates
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
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.Component(p,pType)", Array(p, pType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.component(p,pType)", Array(p, pType)
End Sub


Private Sub cmdClearBC_Click()
    On Error GoTo errHandler
    If cboBC.Items.ItemCount = 0 Then
        Exit Sub
    End If
    lngTPID = 0
    cboBC.Items.SelectItem(cboBC.Items(0)) = False
    lngTPID = 0
    lblFound.Visible = (lngTPID <> 0)
    cmdPlaceOrder.Enabled = (lngTPID <> 0) And XA.Count(1) > 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.cmdClearBC_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.cmdClearBC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    If Not bOrderPlaced Then
        If MsgBox("You have not taken an action. Do you wish to close this form?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.cmdClose_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPlaceOrder_Click()
    On Error GoTo errHandler
    G1.Update
    
    If strType = "ORDER" Then
        PlaceOrder
    Else
        PlaceOnReserve
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.cmdPlaceOrder_Click", , EA_NORERAISE
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
    
        If oSQL Is Nothing Then Set oSQL = New z_SQL
        oSQL.RunSQL "UPDATE tPRODUCT SET P_QTYRESERVED = P_QTYRESERVED + " & CLng(XA(i, 4)) & " WHERE P_ID = '" & XA(i, 7) & "'"
        oSQL.RunSQL "UPDATE tStoreP SET STP_QTYRESERVED = STP_QTYRESERVED + " & CLng(XA(i, 4)) & " WHERE STP_P_ID = '" & XA(i, 7) & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
        oSQL.RunSQL "INSERT INTO tRM (RM_P_ID,RM_TP_ID,RM_QTY,RM_DATE,RM_TYPE) VALUES ('" & XA(i, 7) & "'," & lngTPID & "," & CLng(XA(i, 4)) & ",{d '" & Format(Date, "yyyy-mm-dd") & "'},'I')"
    Next
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    MsgBox "Placed on reserve", , "Status"
    Unload Me

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.PlaceOnReserve"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.PlaceOnReserve"
End Sub

Private Sub PlaceOrder()
    On Error GoTo errHandler
Dim i As Integer
Dim oCOL As a_COL
Dim lngResult As Long
Dim bFound As Boolean
Dim bProductToOrder As Boolean
Dim bZeroDeposit As Boolean
Dim bIssue As Boolean
Dim errRepeat As Integer

    errRepeat = 0

    bIssue = False

    bProductToOrder = False
    bZeroDeposit = False
    For i = 1 To XA.UpperBound(1)
        If XA(i, 4) > 0 Then
            bProductToOrder = True
        End If
        If XA(i, 5) < 200 Then
            bZeroDeposit = True
        End If
    Next
    If bProductToOrder = False Then
        MsgBox "There are no non-zero quantities on this order. No action will be taken!", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If
    If MsgBox("You are placing an order for " & txtInitials & " " & txtName & "?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    If bZeroDeposit = True Then
        If MsgBox("You are placing an order with a small deposit. Confirm?", vbOKCancel + vbQuestion, "Warning") = vbCancel Then
            Exit Sub
        End If
    End If
    If (oPC.getProperty("IssueQuickCOs") = "TRUE") Then
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_CO_SIGN, , "Sign this document", DOCAPPROVAL, , , gSTAFFID) = False Then
                Exit Sub
            Else
                bIssue = True
            End If
        Else
            If oCO.STATUS = stInProcess Then
                If MsgBox("Issue this document?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
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
        oCust.SetAccAcNo oPC.getProperty("DefaultAccountingAccno")
        oCust.Addresses.FindByDescription("Default").SetEmail txtEmail
        oCust.Addresses.FindByDescription("Default").SetAddress Me.txtAddress
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
    
    
    Set oCust = Nothing
    Set oCO = New a_CO
    oCO.BeginEdit
    oCO.SetCustomer lngTPID
    oCO.OrderType = enNormalCO
    oCO.SetMemo FNS(txtSPInstr)
    oCO.StaffID = gSTAFFID
 '   Set oCO.BillTOAddress = oCO.Customer.BillTOAddress
 '   Set oCO.DelToAddress = oCO.Customer.DelToAddress
    For i = 1 To XA.UpperBound(1)
        Set oCOL = oCO.COLines.Add
        oCOL.BeginEdit
        oCOL.SetLineProduct XA(i, 8)
        oCOL.SetQty XA(i, 4)
        oCOL.SetDeposit XA(i, 5)
        oCOL.SetRef XA(i, 6)
        oCOL.DepositStatus = "P"
        If DateDiff("d", Date, oCOL.ETA) <= 1 Then
            oCOL.SetETA "2w"
        End If
        oCOL.ApplyEdit
    Next
    If bIssue Then
        oCO.SetStatus stISSUED
    Else
        oCO.SetStatus stInProcess
    End If
    oCO.Post
    
    Screen.MousePointer = vbDefault
    MsgBox "Order placed", , "Status"
Dim f As New frmCOPreview
    f.component oCO.TRID, False
    Unload Me
    f.Show
    Exit Sub
errHandler:
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmPlaceCOs: PlaceOrder"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            MsgBox "Memory error trying to load form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load Place customer order form."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.PlaceOrder"
End Sub

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
    If oCust.Addresses.Count > 0 Then
        txtEmail = oCust.Addresses.DefaultAddress.EMail
        txtAddress = oCust.Addresses.DefaultAddress.AddressMailing
    End If
    lngTPID = frmC.CustomerID
    Set frmC = Nothing
    lblFound.Visible = (lngTPID <> 0)
    cmdPlaceOrder.Enabled = (lngTPID <> 0) And XA.Count(1) > 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    cmdPlaceOrder.Enabled = XA.Count(1) > 0 And lngTPID > 0 And lblFound.Visible = True
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.Form_Activate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    G1.Array = XA
    G1.ReBind
    Set tlBookclubs = New z_TextList
    tlBookclubs.Load ltBookclub
    SetupCboBC
    LoadComboEx cboBC.Items, tlBookclubs
    If strType = "ORDER" Then
        Me.cmdPlaceOrder.Caption = "Place order"
        Me.Caption = "Quick order"
    Else
        Me.cmdPlaceOrder.Caption = "Put on reserve"
        Me.Caption = "On Reserve"
    End If
    SetGridLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboBC_SelectionChanged()
    On Error GoTo errHandler
    lngTPID = tlBookclubs.key(cboBC.Items.CellCaption(cboBC.Items.SelectedItem, 0))
    lblFound.Visible = (lngTPID <> 0)
    txtName = cboBC.Items.CellCaption(cboBC.Items.SelectedItem, 0)
    cmdPlaceOrder.Enabled = (lngTPID <> 0) And XA.Count(1) > 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.cboBC_SelectionChanged", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.cboBC_SelectionChanged", , EA_NORERAISE
    HandleError
End Sub

Sub SetupCboBC()
    On Error GoTo errHandler
    cboBC.BeginUpdate
    cboBC.WidthList = 330
    cboBC.HeightList = 142
    cboBC.AllowSizeGrip = True
    cboBC.AutoDropDown = True
    
    cboBC.Columns.Add "Name"
    cboBC.Columns.Add "Phone"
    
    cboBC.Columns(0).Width = 190
    cboBC.Columns(1).Width = 120
    cboBC.BackColorLock = Me.BackColor
    cboBC.EndUpdate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.SetupCboBC"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.SetupCboBC"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set tlBookclubs = Nothing
    Set XA = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Label2_Click()
    On Error GoTo errHandler

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.Label2_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.Label2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Text1_Change()
    On Error GoTo errHandler

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.Text1_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.Text1_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_AfterDelete()
    On Error GoTo errHandler
    If XA.Count(1) = 0 Then Me.cmdPlaceOrder.Enabled = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.G1_AfterDelete", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.G1_AfterDelete", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
    G1.InsertMode = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.G1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.G1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim lngTmp As Long
    
    If ColIndex = 3 Then
        If IsNumeric(G1.Text) Then
            If Not ConvertToLng(G1.Text, lngTmp) Then
                Cancel = True
            ElseIf CLng(G1.Text) < 0 Then
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.G1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub G1_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.G1_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_LostFocus()
    On Error GoTo errHandler
    G1.Update
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.G1_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.G1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtName_Change()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To XA.UpperBound(1)
        XA(i, 6) = Left(txtName, 10)
    Next
    G1.ReBind
    cmdPlaceOrder.Enabled = Len(txtPhone) > 5 And Len(txtName) > 1 And XA.Count(1) > 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.txtName_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.txtName_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPhone_Change()
    On Error GoTo errHandler
    cmdPlaceOrder.Enabled = Len(txtPhone) > 5 And Len(txtName) > 1 And XA.Count(1) > 0
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPlaceCO.txtPhone_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.txtPhone_Change", , EA_NORERAISE
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
'    ErrorIn "frmPlaceCO.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPlaceCO.txtPhone_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

