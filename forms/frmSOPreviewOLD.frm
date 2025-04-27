VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmSOPreview
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase order preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmSOPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCurrencyMsg 
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
      Height          =   375
      Left            =   3000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5115
      Width           =   3075
   End
   Begin VB.CommandButton cmdClose 
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
      Picture         =   "frmSOPreview.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin VB.TextBox txtDeliverTo 
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
      Height          =   750
      Left            =   7785
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   15
      Width           =   1980
   End
   Begin VB.TextBox txtTPMemo 
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
      Height          =   480
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5685
      Width           =   3735
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
      Picture         =   "frmSOPreview.frx":0635
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
      Top             =   165
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
      Picture         =   "frmSOPreview.frx":077F
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
      Height          =   315
      Left            =   9930
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   30
      Width           =   1155
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
      Height          =   345
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   165
      Width           =   1545
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3975
      Left            =   225
      TabIndex        =   0
      Top             =   840
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7011
      SortKey         =   1
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483635
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ref"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Firm"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "SS"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Rec."
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Disc."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "E.T.A."
         Object.Width           =   2540
      EndProperty
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   675
      Left            =   3855
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   30
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   1191
      BackColor       =   14737632
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
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1215
      X2              =   2685
      Y1              =   0
      Y2              =   825
   End
   Begin VB.Label Label3 
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
      Height          =   285
      Left            =   9315
      TabIndex        =   16
      Top             =   75
      Width           =   2055
   End
   Begin VB.Label txtPhone 
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
      Height          =   210
      Left            =   4365
      TabIndex        =   14
      Top             =   465
      Width           =   3105
   End
   Begin VB.Label txtSuppname 
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
      Height          =   285
      Left            =   4365
      TabIndex        =   13
      Top             =   60
      Width           =   3105
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   3930
      TabIndex        =   12
      Top             =   60
      Width           =   270
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   6390
      TabIndex        =   9
      Top             =   4935
      Width           =   2610
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   9090
      TabIndex        =   7
      Top             =   4920
      Width           =   1845
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   540
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   75
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
      TabIndex        =   2
      Top             =   240
      Width           =   1365
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuactions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit document"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print document"
      End
      Begin VB.Menu mnuVoid 
         Caption         =   "&Void document"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel document"
      End
      Begin VB.Menu mnuCancelLine 
         Caption         =   "&Cancel selected line"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmSOPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cPO As c_POs
Dim oPO As a_PO
Dim dblTotal As Double

Dim lngID As Long

Public Sub Component(pID As Long)
    lngID = pID
    Set oPO = New a_PO
    oPO.Load lngID, True
    Me.Caption = "Purchase order to " & oPO.TPName
    LoadControls
    SetMenu
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
    
        With oPO
            Me.txtDate = .DocDate
            Me.txtStatus = .statusF
            CancelLine.Visible = (.status = stCANCELLED Or .status = stVOID)
            If .status = stInProcess Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtInvoiceNum = .DocCode
            Me.txtSuppname = .Supplier.NameAndCode(20)
            Me.txtPhone = .Supplier.OrderToAddress.PhoneandFax
            Me.txtTPMemo = IIf(Len(.TPMemo) > 0, "Note:  " & Trim$(.TPMemo), "")
            txtDeliverTo = .DeliverToAddress
            .DisplayTotals strTotalCaption, strTotalValues, oPO.ISForeignCurrency
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
        End With
        LoadListView
        'txtCurrencyMsg = oPO.
        SetLvw
        Me.cmdClose.SetFocus
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
Resume
End Sub



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
   oPO.PrintPO_Display
End Sub

Private Sub cmdPrint_Click()
Dim frm As frmPrintingOptions_PO
'
    Set frm = New frmPrintingOptions_PO
    frm.ComponentObject oPO
    frm.Show vbModal
    
EXIT_Handler:
 '   Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
End Sub
Private Sub cmdEdit_Click()
Dim blnEdit As Boolean
Dim frm As frmPO
Dim bCancel As Boolean
    On Error GoTo ERR_Handler
  '  If frmInvoiceAQ Is Nothing Then
        Set frm = New frmPO
  '  End If
    blnEdit = True
    frm.Show
    frm.Component bCancel, oPO ', lngID

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

    On Error GoTo ERR_Handler
    lvw.ListItems.Clear
    For i = 1 To oPO.POLines.Count
        Set lstItem = lvw.ListItems.Add
        With lstItem
            .Key = oPO.POLines(i).POLID & "k"
            .Text = oPO.POLines(i).ProductCodeF
            .SubItems(1) = oPO.POLines(i).TitleAuthor
            .SubItems(2) = oPO.POLines(i).Ref
            .SubItems(3) = oPO.POLines(i).QtyFirm
            .SubItems(4) = oPO.POLines(i).Qtyseesafe
            .SubItems(5) = oPO.POLines(i).QtyReceivedSoFar
            If oPO.ISForeignCurrency Then
                .SubItems(6) = oPO.POLines(i).PriceF_Foreign
                .SubItems(8) = oPO.POLines(i).ExtensionF_Foreign
            Else
                .SubItems(6) = oPO.POLines(i).PriceF
                .SubItems(8) = oPO.POLines(i).ExtensionF
            End If
            .SubItems(7) = oPO.POLines(i).DiscountF
            .SubItems(9) = oPO.POLines(i).ETAF
        If oPO.POLines(i).Fulfilled = "CAN" Then
            lstItem.ForeColor = vbRed
            .ListSubItems(1).ForeColor = vbRed
            .ListSubItems(2).ForeColor = vbRed
            .SubItems(1) = "***CANCELLED***" & oPO.POLines(i).TitleAuthor
        End If
        End With
        If oPO.POLines(i).Note > "" Then
            Set lstItem = lvw.ListItems.Add
            lstItem.SubItems(1) = "Note:  " & oPO.POLines(i).Note
        End If
    Next i
   ' If lvw.ListItems.Count > 0 Then lvw.ListItems(1).Selected = True

    
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
'    If pPostage = 0 And curTotalDeposits = 0 And oPO.VATAble Then
'        strSummaryDescription = "(Includes VAT of " & Format(dblVAT, pCurrFormat) & ")          Total: "
'        strSummary = Format(curTotalValue, pCurrFormat)
'    ElseIf oPO.VATAble Then
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
'    ElseIf Not oPO.VATAble Then
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

Private Sub Form_Activate()
    mnuCancelLine.Enabled = False
End Sub

Private Sub Form_Load()
    
    Me.top = 50
    Me.left = 50
    Me.Height = 6500
    Me.Width = 11500

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oPO = Nothing
    Me.Hide
End Sub



Private Sub cbTP_Click()
Dim frm As frmSupplierPreview
    Set frm = New frmSupplierPreview
    frm.Component oPO.Supplier
    frm.Show
End Sub

Private Sub Lvw_AfterLabelEdit(Cancel As Integer, NewString As String)
Cancel = True
End Sub


Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvw_Click()
    mnuCancelLine.Enabled = oPO.POLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).QtyReceivedSoFar = 0
End Sub

Private Sub Lvw_DblClick()
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim oP As a_Product
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oPO.POLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).pID, 0
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

'Private Sub mnuFileEdit_Click()
'Dim ofrm As New frmInvoice
'Dim blnEdit As Boolean
'
'    On Error GoTo ERR_Handler
'
'    If mnuFileEdit.Caption = "&Edit" Then
'        blnEdit = True
'        ofrm.Component oPO ', lngID
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
Private Sub mnuFileExit_Click()
    Me.Hide
End Sub

Private Sub SetLvw()
Dim style As Long
Dim hHeader As Long
   hHeader = SendMessage(lvw.hwnd, LVM_GETHEADER, 0, ByVal 0&)
   style = GetWindowLong(hHeader, GWL_STYLE)
   style = style Xor HDS_BUTTONS
   If style Then
      Call SetWindowLong(hHeader, GWL_STYLE, style)
      Call SetWindowPos(lvw.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
   End If
End Sub
Private Sub SetMenu()
    mnuVoid.Enabled = (oPO.statusF = "IN PROCESS")
    mnuCancel.Enabled = (oPO.statusF = "ISSUED") And oPO.CanCancel = True
End Sub

Private Sub mnuCancel_Click()
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelPO oPO
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuCancelLine_Click()
Dim oP As a_Product
    If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oPO.POLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).CancelLine
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuEdit_Click()
    cmdEdit_Click
End Sub

Private Sub mnuPrint_Click()
    cmdPrint_Click
End Sub

Public Sub RefreshData()
    oPO.Reload
    LoadControls
End Sub

Private Sub mnuvoid_Click()
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oPO.VoidDocument
End Sub
