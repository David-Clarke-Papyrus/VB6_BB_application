VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "COOLBU~1.OCX"
Begin VB.Form frmCOPreview 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D3D3CB&
   Caption         =   "Order"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11340
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmCOPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   270
      TabIndex        =   21
      Top             =   5640
      Width           =   1590
      Begin VB.Label lblOS 
         Caption         =   "OS"
         Height          =   300
         Left            =   435
         TabIndex        =   25
         Top             =   15
         Width           =   345
      End
      Begin VB.Label lblFUL 
         Caption         =   "FUL"
         Height          =   300
         Left            =   765
         TabIndex        =   24
         Top             =   15
         Width           =   345
      End
      Begin VB.Label lblCAN 
         Caption         =   "CAN"
         Height          =   300
         Left            =   1140
         TabIndex        =   23
         Top             =   15
         Width           =   345
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Key:"
         Height          =   300
         Left            =   30
         TabIndex        =   22
         Top             =   15
         Width           =   345
      End
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1065
      Left            =   255
      TabIndex        =   20
      Top             =   750
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   1879
      BackColor       =   13882315
      ForeColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   1
      BackStyle       =   0
   End
   Begin VB.TextBox txtIssued 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
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
      Height          =   1140
      Left            =   2910
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1995
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCOPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Print the invoice"
      Top             =   4860
      Width           =   855
   End
   Begin VB.TextBox txtCurrency 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00706034&
      Height          =   255
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   420
      Width           =   1635
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3825
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   4440
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
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
      Height          =   720
      Left            =   1125
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCOPreview.frx":284D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   90
      Width           =   1545
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   255
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCOPreview.frx":2997
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   60
      Width           =   1545
   End
   Begin VB.TextBox txtDocCOde 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   375
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3000
      Left            =   240
      OleObjectBlob   =   "frmCOPreview.frx":2CA1
      TabIndex        =   26
      Top             =   1815
      Width           =   10815
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1095
      X2              =   2565
      Y1              =   0
      Y2              =   825
   End
   Begin VB.Label txtTPFax 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   330
      TabIndex        =   17
      Top             =   1455
      Width           =   2865
   End
   Begin VB.Label txtTPPhone 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   330
      TabIndex        =   16
      Top             =   1110
      Width           =   2865
   End
   Begin VB.Label txtTPName 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   330
      TabIndex        =   15
      Top             =   810
      Width           =   2865
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Goods to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7845
      TabIndex        =   13
      Top             =   780
      Width           =   1050
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Bill to:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5085
      TabIndex        =   12
      Top             =   780
      Width           =   660
   End
   Begin VB.Label lblBillToAddress 
      BackColor       =   &H00D3D3CB&
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
      Height          =   945
      Left            =   5865
      TabIndex        =   11
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label lblDelToAddress 
      BackColor       =   &H00D3D3CB&
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
      Height          =   945
      Left            =   9015
      TabIndex        =   10
      Top             =   780
      Width           =   2055
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   5505
      TabIndex        =   8
      Top             =   4935
      Width           =   3495
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   9090
      TabIndex        =   6
      Top             =   4935
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   705
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Invoice No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1365
   End
End
Attribute VB_Name = "frmCOPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCO As c_COs
Dim oCO As a_CO
Dim dblTotal As Double
Dim XA As XArrayDB
Dim flgLoading As Boolean

Public Sub Component(pID As Long)
Dim lngID As Long
    lngID = pID
    Set oCO = New a_CO
    oCO.Load lngID, True
    If oCO.OrderType = enWant Then
        Me.Caption = "Wants for " & oCO.Customer.Fullname & oCO.StaffNameB
    ElseIf oCO.OrderType = enNormalCO Then
       ' Me.Caption = "Order from " & oCO.Customer.Fullname & oCO.StaffNameB
        Me.Caption = "Order from " & oCO.Customer.Fullname & oCO.StaffNameB & IIf(oCO.OrderRef > "", "  (ref:" & oCO.OrderRef & ")", "")
    End If
    flgLoading = True
    LoadControls
    SetMenu
    lblOS.BackColor = COLOR_PALEYELLOW
    lblFUL.BackColor = COLOR_FULFILLED
    lblCAN.BackColor = COLOR_CANCELLED
    flgLoading = False
End Sub
Public Sub ComponentObject(pInvoice As a_CO)
    Set oCO = pInvoice
    Me.Caption = "Order from " & oCO.TPName
    flgLoading = True
    LoadControls
    flgLoading = False
End Sub

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub

Private Sub SetMenu()
    Forms(0).mnuVoid.Enabled = (oCO.statusF = "IN PROCESS" And oCO.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oCO.statusF = "ISSUED")
    Forms(0).mnuCancelLine.Enabled = (oCO.statusF = "ISSUED" And oCO.IsNew = False)
    Forms(0).mnuFulfil.Enabled = (oCO.statusF = "ISSUED") 'And oCO.CanCancel = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    
    ofrm.Component oCO.Memo
    ofrm.Show vbModal
    oSM.setMemo ofrm.Memo, oCO.TRID
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = "Note: " & ofrm.Memo
    oSM.setMemo ofrm.Memo, oCO.TRID
    oCO.setMemo ofrm.Memo
    
    Unload ofrm

    Set ofrm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCO.mnuMemo"
End Sub


Public Sub mnuFulfilLine()
    On Error GoTo errHandler
Dim oCOL As a_COL
    Set oCOL = oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 15)))
'    If oCOL.Fulfilled <> "OS" Then
'        MsgBox "This line is not outstanding and cannot be marked fulfilled.", vbExclamation + vbOKOnly, "Can't do this"
'        Exit Sub
'    Else
'        If oCOL.QtyDispatched = 0 Then
'            If MsgBox("This line has not been received at all and should be marked cancelled." & vbCrLf & "Do you want to mark the line cancelled?", vbQuestion + vbYesNo, "Can't do this") = vbNo Then
'               Exit Sub
'            End If
'            Screen.MousePointer = vbHourglass
'            oCOL.CancelLine
'            RefreshData
'            Screen.MousePointer = vbDefault
'        Else
            If MsgBox("Do you wish to mark the selected line as fulfilled?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            Screen.MousePointer = vbHourglass
            oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 15))).FulfilLine
            RefreshData
            Screen.MousePointer = vbDefault
    '    End If
  '  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuFulfilLine"
End Sub
Public Sub mnuCancelLine()
Dim oP As a_Product
Dim str As String
Dim oCOL As a_COL
    Set oCOL = oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 15)))
    If oCOL.QtyDispatched > 0 Then
        MsgBox "This line is partially fulfilled and can only be marked as fulfilled.", vbExclamation + vbOKOnly, "Can't do this"
    Else
        If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        Set oP = New a_Product
        oCO.COLines.FindLineByID(val(XA(G1.Bookmark, 15))).CancelLine
        RefreshData
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub Form_Activate()
    SetMenu
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub

Private Sub LoadControls()
Dim dblVAT As Double
Dim dblConversionRate As Double
Dim strCurrencyFormat As String
Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strAddress As String
Dim strTotalCaption As String
Dim strTotalValues As String
    On Error GoTo ERR_Handler
    
        With oCO
            Me.txtDate = .DocDateF
            If DateDiff("d", .DocDate, .issDate) > 1 Then
                Me.txtIssued = "Issued: " & .IssDateF
            Else
                txtIssued = ""
            End If
            Me.txtDocCOde = .DocCode
            Me.txtStatus = .statusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If oPC.getProperty("CanEditCOs") = "TRUE" Then
                If .Status = stInProcess Or .Status = stISSUED Or .OrderType = enWant Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            Else
                If .Status = stInProcess Or .OrderType = enWant Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            End If
            Me.txtDocCOde = .DocCode
            Me.txtTPName = .Customer.Fullname & IIf(Len(.TPAccNum) > 0, " (" & .TPAccNum & ")", "")
            txtTPPhone = .Customer.Phone
            If Not (.Customer.BillTOAddress Is Nothing) Then
                Me.txtTPFax = "Fax: " & .Customer.BillTOAddress.Fax
            End If
            Me.txtTPMemo = IIf(Len(.Memo) > 0, "Note:  " & Trim$(.Memo), "")
            txtTPMemo.Visible = (txtTPMemo > "")
            If .BillToAddressID > 0 Then
                If Not .BillTOAddress Is Nothing Then strAddress = .BillTOAddress.AddressMailing
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .GoodsAddressID > 0 Then
                If Not .DelToAddress Is Nothing Then strAddress = .DelToAddress.AddressMailing
            End If
            Me.lblDelToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            .CalculateTotal
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
        End With
   '     LoadListView
        LoadGrid
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    
Resume
End Sub


Private Sub cbCust_Click()
Dim frm As New frmCustomerPreview
    If flgLoading Then Exit Sub
    frm.Component oCO.Customer
    frm.Show
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim frm As frmPrintingOptions_CO
'
    Set frm = New frmPrintingOptions_CO
    frm.ComponentObject oCO
    frm.Show vbModal
    
EXIT_HANDLER:
 '   Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
Resume
End Sub
Private Sub cmdEdit_Click()
Dim blnEdit As Boolean
Dim frm As frmCO
Dim bCancel As Boolean

    On Error GoTo ERR_Handler
    WaitMsg "Loading . . .", True, Me
    Set frm = New frmCO
    blnEdit = True
    frm.Component bCancel, oCO
    frm.Show
    WaitMsg "", False, Me

EXIT_HANDLER:
    Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    Resume
End Sub

Private Sub LoadGrid()
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long

    Set XA = New XArrayDB
    XA.Clear
    On Error GoTo ERR_Handler
    G1.Columns(9).Width = 0
    If oCO.OrderType = enWant Then
        G1.Columns(2).Width = 4000
        G1.Columns(3).Width = 1500
        G1.Columns(4).Width = 3500
        G1.Columns(5).Width = 0
        G1.Columns(6).Width = 0
        G1.Columns(7).Width = 0
        G1.Columns(8).Width = 0
        G1.Columns(9).Width = 0
    End If
    XA.ReDim 1, oCO.COLines.Count, 1, 16
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    For i = 1 To oCO.COLines.Count
             XA(i, 1) = oCO.COLines(i).CodeF
             XA(i, 2) = oCO.COLines(i).TitleAuthorPublisher
             If oCO.OrderType = enWant Then
                 XA(i, 3) = oCO.COLines(i).WantDateF
                 G1.Columns(2).Caption = "Date of want"
                 G1.Columns(3).Caption = "Note"
                 XA(i, 4) = oCO.COLines(i).Note
             Else
                 XA(i, 3) = oCO.COLines(i).Ref
                 G1.Columns(2).Caption = "Ref"
                 XA(i, 4) = oCO.COLines(i).QtyOrdered_QtyDispatched
            End If
             If oCO.COLines(i).Deposit > 0 Then
                 XA(i, 5) = oCO.COLines(i).DepositF
                 XA(i, 6) = oCO.COLines(i).DepositStatus
             Else
                 XA(i, 5) = " "
                 XA(i, 6) = " "
             End If
             XA(i, 7) = oCO.COLines(i).PriceF
             XA(i, 8) = oCO.COLines(i).DiscountF
             XA(i, 9) = oCO.COLines(i).ExtensionF
             XA(i, 11) = oCO.COLines(i).Fulfilled
             XA(i, 12) = oCO.COLines(i).Key
             XA(i, 13) = oCO.COLines(i).code
             XA(i, 14) = oCO.COLines(i).POLID
             XA(i, 15) = oCO.COLines(i).COLineID
             XA(i, 16) = oCO.COLines(i).Ean
            
            If oCO.COLines(i).Note > "" Or oCO.COLines(i).lastaction > "" Then
                XA(i, 10) = "Note:  " & oCO.COLines(i).Note & " Last action:" & oCO.COLines(i).lastactionAndDate
                G1.Columns(9).Width = 4000
            End If
    Next i
    
    G1.Array = XA
    G1.ReBind

    
EXIT_HANDLER:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_HANDLER
    Resume
End Sub


Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKey4 Then
        If MsgBox("Confirm close?", vbOKCancel, "Close form") = vbOK Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    flgLoading = True
    
    Me.top = 50
    Me.left = 50
    Me.Height = 6500
    Me.Width = 11500
    flgLoading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
    Set oCO = Nothing
   
End Sub



Private Sub Label5_DblClick()
Dim frm As frmCustomerPreview
    Set frm = New frmCustomerPreview
    frm.Component oCO.Customer
    frm.Show
End Sub


Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
   ' str = FNS(XA.Value(G1.Bookmark, 13))
    str = IIf(FNS(XA.Value(G1.Bookmark, 16)) > "", FNS(XA.Value(G1.Bookmark, 16)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.G1_Click", , EA_NORERAISE
    HandleError

End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 12))
    If str = "" Then Exit Sub
    Forms(0).mnuCancelLine.Enabled = oCO.COLines(str).QtyDispatched = 0
   ' str = FNS(XA.Value(G1.Bookmark, 13))
    str = IIf(FNS(XA.Value(G1.Bookmark, 16)) > "", FNS(XA.Value(G1.Bookmark, 16)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
End Sub

Private Sub G1_SelChange(Cancel As Integer)
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 12))
    If str = "" Then Exit Sub
    Forms(0).mnuCancelLine.Enabled = oCO.COLines(str).QtyDispatched = 0
    str = FNS(XA.Value(G1.Bookmark, 13))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
End Sub
Public Sub mnuCancel()
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelCO oCO
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Private Sub G1_DblClick()
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 12))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oCO.COLines(str).pID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.Component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuFileExit_Click()
    Me.Hide
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 3, 10
            GetRowType = XTYPE_STRING
        Case 4, 5, 6, 7, 8, 9
            GetRowType = XTYPE_INTEGER
    End Select
End Function


Public Sub mnuVoid()
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oCO.VoidDocument
    RefreshData
End Sub
Public Sub RefreshData()
    oCO.ReLoad
    LoadControls
End Sub

Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    
    If XA(Bookmark, 11) = "CAN" Then
        RowStyle.BackColor = COLOR_CANCELLED
    ElseIf XA(Bookmark, 11) = "FUL" Then
        RowStyle.BackColor = COLOR_FULFILLED
    ElseIf XA(Bookmark, 14) > 0 Then
        RowStyle.BackColor = COLOR_PALEYELLOW
    End If
        
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuShowOLHistGrp   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPOPreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub

Public Sub ShowPreviousOLVersions()
Dim frm As frmCOLHistory
Dim COLID As Long

    COLID = val(XA.Value(G1.Bookmark, 15))
    Set frm = New frmCOLHistory
    frm.Component COLID
    frm.Show vbModal
End Sub
