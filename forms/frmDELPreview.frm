VERSION 5.00
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDELPreview 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Delivery preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11970
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmDELPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
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
      Picture         =   "frmDELPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtCurrency 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   9780
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   570
      Width           =   1305
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
      Picture         =   "frmDELPreview.frx":284D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtDate 
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
      TabIndex        =   5
      Top             =   210
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
      Picture         =   "frmDELPreview.frx":2997
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   285
      Left            =   9420
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1680
   End
   Begin VB.TextBox txtInvoiceNum 
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
      Height          =   330
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   210
      Width           =   1545
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   6800
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   14416635
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title / Author / Publisher"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "for Ref."
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "P.O."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Firm"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "SS"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Disc."
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   2118
      EndProperty
   End
   Begin CoolButtonControl.CoolButton cbSupp 
      Height          =   795
      Left            =   3825
      TabIndex        =   10
      Top             =   75
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1402
      BackColor       =   14737632
      ForeColor       =   14737632
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
   Begin VB.Label lblSI 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   390
      TabIndex        =   2
      Top             =   540
      Width           =   2970
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1110
      X2              =   2580
      Y1              =   0
      Y2              =   825
   End
   Begin VB.Label txtTPFax 
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
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6945
      TabIndex        =   14
      Top             =   540
      Width           =   2265
   End
   Begin VB.Label txtTPPhone 
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
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6945
      TabIndex        =   13
      Top             =   225
      Width           =   2265
   End
   Begin VB.Label txtTPName 
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
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   3945
      TabIndex        =   12
      Top             =   240
      Width           =   2910
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2985
      TabIndex        =   8
      Top             =   4845
      Width           =   2325
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7845
      TabIndex        =   7
      Top             =   4890
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   765
      Left            =   285
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3450
   End
End
Attribute VB_Name = "frmDELPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cDEL As c_DELs
Dim oDel As a_Delivery
Dim dblTotal As Double

Private Sub SetMenu()
    Forms(0).mnuVoid.Enabled = (oDel.statusF = "IN PROCESS" And oDel.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oDel.statusF = "ISSUED")
    Forms(0).mnuCancelLine.Enabled = (oDel.statusF = "ISSUED" And oDel.IsNew = False)
End Sub
Private Sub Form_Activate()
    SetMenu
End Sub
Private Sub UnsetMenu()
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuInvNote.Enabled = False
    Forms(0).mnuGenDisc.Enabled = False
    Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuAdjust.Enabled = False
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub

Public Sub component(pID As Long)
Dim lngID As Long
    lngID = pID
    Set oDel = New a_Delivery
    oDel.Load lngID
  '  oDel.CalculateTotals
    LoadControls
    SetMenu
End Sub
Public Sub ComponentObject(pDelivery As a_Delivery)
    Set oDel = pDelivery
    Me.Caption = "Delivery from " & oDel.TPName
    LoadControls
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
    
        With oDel
            Me.txtDate = .DOCDate
            Me.txtStatus = .statusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            cmdEdit.Enabled = .Status = stInProcess
            Me.txtInvoiceNum = .DocCode
            Me.txtTPName = .Supplier.NameAndCode(24)
            Me.txtTPPhone = "Phone: " & .Supplier.OrderToAddress.Phone
            Me.txtTPFax = "Fax: " & .Supplier.OrderToAddress.Fax
            Me.txtCurrency = oDel.CaptureCurrency.Description
            Me.lblTotalValues = oDel.TotalLessDiscExtF(oDel.isFOreignCurrency)
            lblSI.Caption = .SupplierInvoiceRef & " : " & .SupplierInvoiceDateF
        End With
        LoadListView
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
Resume
End Sub

Private Sub cmdPreview_Click()
'Dim frm As frmPreview_
'    oDEL.PrintInvoice_Display True
End Sub

Private Sub cbSupp_Click()
Dim frm As New frmSupplier
    frm.component oDel.Supplier
    frm.Show
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim frm As frmPrintingOptions_DEL
'
    Set frm = New frmPrintingOptions_DEL
    frm.ComponentObject oDel
    frm.Show vbModal
    
EXIT_Handler:
 '   Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
Resume
End Sub
Private Sub cmdEdit_Click()
Dim blnEdit As Boolean
Dim frm As frmdel
Dim bCancel As Boolean
    On Error GoTo ERR_Handler
    Set frm = New frmdel
    blnEdit = True
    frm.component bCancel, , oDel
    frm.Show

EXIT_Handler:
    Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

Private Sub LoadListView()
Dim lstItem As ListItem
Dim i As Integer
Dim currDeposit As Currency
Dim currPrice As Currency
Dim dblVAT As Double
Dim strSummaryDescription As String
Dim strSummary As String
Dim lngTotal As Long
Dim lngDepositTotal As Long
Dim tmp
    On Error GoTo ERR_Handler
                    
    lvw.ListItems.Clear
    For i = 1 To oDel.DeliveryLines.Count
            Set lstItem = lvw.ListItems.Add
            With lstItem
                .Key = oDel.DeliveryLines(i).DELLID & "k"
                .Text = oDel.DeliveryLines(i).CodeF
                .SubItems(1) = oDel.DeliveryLines(i).Title
                .SubItems(2) = oDel.DeliveryLines(i).Ref
                .SubItems(3) = oDel.DeliveryLines(i).POCode
                tmp = oDel.DeliveryLines(i).POLQtyFirm
                .SubItems(4) = oDel.DeliveryLines(i).QtyFirm & IIf(tmp > 0, "(", "") & IIf(tmp > 0, tmp, "") & IIf(tmp > 0, ")", "")
                tmp = oDel.DeliveryLines(i).POLQtySS
                .SubItems(5) = oDel.DeliveryLines(i).QtySS & IIf(tmp > 0, "(", "") & IIf(tmp > 0, tmp, "") & IIf(tmp > 0, ")", "")
                tmp = oDel.DeliveryLines(i).POLDiscount
                .SubItems(7) = oDel.DeliveryLines(i).DiscountF & IIf(tmp > 0, "(", "") & IIf(tmp > 0, oDel.DeliveryLines(i).POLDiscountF, "") & IIf(tmp > 0, ")", "")
                tmp = oDel.DeliveryLines(i).POLPrice
                .SubItems(6) = oDel.DeliveryLines(i).PriceF(oDel.isFOreignCurrency) & IIf(tmp > 0, "(", "") & IIf(tmp > 0, oDel.DeliveryLines(i).POLPriceF(oDel.isFOreignCurrency), "") & IIf(tmp > 0, ")", "")
                .SubItems(8) = oDel.DeliveryLines(i).PLessDiscExtF(oDel.isFOreignCurrency)
            
            End With
    Next i
    

    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub
'Private Sub LoadSummary(pPostage As Currency, pVAT As Double, pConversionRate As Double, pCurrFormat As String, curTotalValue As Currency, curTotalDeposits As Currency)
'Dim currPrice As Currency
'Dim strDiscount As String
'Dim dblVAT As Double
'Dim strSummaryDescription As String
'Dim strSummary As String
'    dblVAT = (curTotalValue / (1 + pVAT)) * pVAT
'    If pPostage = 0 And curTotalDeposits = 0 And oDEL.VATAble Then
'        strSummaryDescription = "(Includes VAT of " & Format(dblVAT, pCurrFormat) & ")          Total: "
'        strSummary = Format(curTotalValue, pCurrFormat)
'    ElseIf oDEL.VATAble Then
'        strSummaryDescription = "Subtotal:"
'        strSummary = Format(curTotalValue, pCurrFormat)
'        If curTotalDeposits <> 0 Then
'            Me.lblDeposits = "(Deposits paid : " & Format(curTotalDeposits, pCurrFormat) & ")                          "
'        End If
'        If pPostage <> 0 Then
'            strSummaryDescription = strSummaryDescription & vbCrLf & "Plus Postage && handling:"
'            strSummary = strSummary & vbCrLf & Format(pPostage, pCurrFormat)
'        End If
'        strSummaryDescription = strSummaryDescription & vbCrLf & "(Includes VAT of " & Format(dblVAT, pCurrFormat) & ")          Total: "
'        strSummary = strSummary & vbCrLf & Format((curTotalValue + pPostage), pCurrFormat)
'    ElseIf Not oDEL.VATAble Then
'        strSummaryDescription = "Subtotal:"
'        strSummary = Format(curTotalValue, pCurrFormat)
'        strSummaryDescription = strSummaryDescription & vbCrLf & "Less VAT"
'        strSummary = strSummary & vbCrLf & Format(dblVAT, pCurrFormat)
'        If curTotalDeposits <> 0 Then
'            Me.lblDeposits = "(Deposits paid : " & Format(curTotalDeposits, pCurrFormat) & ")                          "
'        End If
'        If pPostage <> 0 Then
'            strSummaryDescription = strSummaryDescription & vbCrLf & "Postage && handling:"
'            strSummary = strSummary & vbCrLf & Format(pPostage, pCurrFormat)
'        End If
'        strSummaryDescription = strSummaryDescription & vbCrLf & "(Total excluding VAT " & Format((curTotalValue - dblVAT) + pPostage, pCurrFormat) & ")"
'    End If
'    Me.lblDescription = strSummaryDescription
'    Me.lblSummary = strSummary
'
'End Sub

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
    
    Me.top = 50
    Me.left = 50
    Me.Height = 6500
    Me.Width = 11500
    SetLvw
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
    Set oDel = Nothing
End Sub



Private Sub Label5_DblClick()
Dim frm As frmSupplierPreview
    Set frm = New frmSupplierPreview
    frm.component oDel.Supplier
    frm.Show
End Sub

Private Sub Lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub
Private Sub lvw_Click()
    Clipboard.SetText oDel.DeliveryLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).Code
End Sub


Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub Lvw_DblClick()
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim oP As a_Product
    If lvw.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oDel.DeliveryLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).pID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
End Sub

'Private Sub mnuFileEdit_Click()
'Dim ofrm As New frmInvoice
'Dim blnEdit As Boolean
'
'    On Error GoTo ERR_Handler
'
'    If mnuFileEdit.Caption = "&Edit" Then
'        blnEdit = True
'        ofrm.Component oDEL ', lngID
'        ofrm.Show vbModal
'    Else
'        blnEdit = False
'        ofrm.PrintInvoice
'    End If
'
'EXIT_Handler:
'    Unload ofrm
'    Set ofrm = Nothing
'    Unload Me
'    Exit Sub
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'End Sub
'
Private Sub SetLvw()
Dim style As Long
Dim hHeader As Long
   
  'get the handle to the listview header
   hHeader = SendMessage(lvw.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   
  'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)
   
  'modify the style by toggling the HDS_BUTTONS style
   style = style Xor HDS_BUTTONS
   
  'set the new style and redraw the listview
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvw.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If


End Sub

Public Sub mnuCancel()
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelDEL oDel
    RefreshData
    Screen.MousePointer = vbDefault
End Sub
Public Sub RefreshData()
    oDel.Reload
    LoadControls
End Sub


Public Sub mnuVoid()
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oDel.VoidDocument
    RefreshData
End Sub


