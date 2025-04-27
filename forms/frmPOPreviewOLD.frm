VERSION 5.00
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmPOPreview 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Purchase order preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11400
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmPOPreviewOLD.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSaveLayout 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Save layout"
      Height          =   285
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5640
      Width           =   975
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
      Left            =   9750
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   465
      Width           =   1305
   End
   Begin VB.TextBox txtCurrencyRates 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00706034&
      Height          =   555
      Left            =   2970
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   4875
      Width           =   3210
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
      Left            =   1980
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPOPreviewOLD.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Print the invoice"
      Top             =   4890
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
      TabIndex        =   9
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
      Left            =   1305
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5670
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
      Picture         =   "frmPOPreviewOLD.frx":0635
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print the invoice"
      Top             =   4890
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
      TabIndex        =   4
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
      Picture         =   "frmPOPreviewOLD.frx":077F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print the invoice"
      Top             =   4890
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
      TabIndex        =   2
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
      TabIndex        =   0
      Top             =   165
      Width           =   1545
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   675
      Left            =   3855
      TabIndex        =   10
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
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3945
      Left            =   225
      OleObjectBlob   =   "frmPOPreviewOLD.frx":08C9
      TabIndex        =   18
      Top             =   795
      Width           =   10725
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
      TabIndex        =   15
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   1
      Top             =   240
      Width           =   1365
   End
End
Attribute VB_Name = "frmPOPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cPO As c_POs
Dim WithEvents oPO As a_PO
Attribute oPO.VB_VarHelpID = -1
Dim dblTotal As Double
Dim XA As XArrayDB
Dim lngID As Long

Private Sub Form_Activate()
    SetMenu
End Sub
Private Sub cmdSaveLayout_Click()
    SaveLayout Me.G1, Me.Name
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

Private Sub SetMenu()
    Forms(0).mnuVoid.Enabled = (oPO.statusF = "IN PROCESS")
    Forms(0).mnuCancel.Enabled = (oPO.statusF = "ISSUED") And oPO.CanCancel = True
    Forms(0).mnuCancelLine.Enabled = (oPO.statusF = "ISSUED")
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuInvNote.Enabled = False
    Forms(0).mnuGenDisc.Enabled = False
    Forms(0).mnuInvAdd.Enabled = False
End Sub


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
    Screen.MousePointer = vbHourglass
    
    With oPO
        Me.txtDate = .DOCDate
        Me.txtStatus = .statusF
        CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
        If .Status = stInProcess Then
            cmdEdit.Enabled = True
        Else
            cmdEdit.Enabled = False
        End If
        Me.txtInvoiceNum = .DocCode
        Me.txtSuppname = .Supplier.NameAndCode(20)
        Me.txtPhone = .Supplier.OrderToAddress.PhoneandFax
        Me.txtTPMemo = IIf(Len(.TPMemo) > 0, "Note:  " & Trim$(.TPMemo), "")
        txtDeliverTo = .DeliverToAddress
        .DisplayTotals strTotalCaption, strTotalValues, oPO.isFOreignCurrency
        lblTotalCaption.Caption = strTotalCaption
        lblTotalValues.Caption = strTotalValues
        txtCurrency = oPO.CaptureCurrency.Description
    End With
    LoadGrid
    Screen.MousePointer = vbDefault
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
   oPO.PrintPO_Display (oPO.isFOreignCurrency)
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

'Private Sub LoadListView()
'Dim lstItem As ListItem
'Dim i As Integer
'Dim currDeposit As Currency
'Dim currPrice As Currency
'Dim dblVAT As Double
'Dim strSummaryDescription As String
'Dim strSummary As String
'Dim lngTotal As Long
'Dim lngDepositTotal As Long
'
'    On Error GoTo ERR_Handler
'    lvw.ListItems.Clear
'    For i = 1 To oPO.polines.Count
'        Set lstItem = lvw.ListItems.Add
'        With lstItem
'            .Key = oPO.polines(i).POLID & "k"
'            .Text = oPO.polines(i).ProductCodeF
'            .SubItems(1) = oPO.polines(i).TitleAuthor
'            .SubItems(2) = oPO.polines(i).Ref
'            .SubItems(3) = oPO.polines(i).QtyFirm
'            .SubItems(4) = oPO.polines(i).Qtyseesafe
'            .SubItems(5) = oPO.polines(i).QtyReceivedSoFar
'                .SubItems(6) = oPO.polines(i).PriceF(oPO.isFOreignCurrency)
'                .SubItems(8) = oPO.polines(i).PLessDiscExtF(oPO.isFOreignCurrency)
'            .SubItems(7) = oPO.polines(i).DiscountF
'            .SubItems(9) = oPO.polines(i).ETAF
'        If oPO.polines(i).fulfilled = "CAN" Then
'            lstItem.ForeColor = vbRed
'            .ListSubItems(1).ForeColor = vbRed
'            .ListSubItems(2).ForeColor = vbRed
'            .SubItems(1) = "***CANCELLED***" & oPO.polines(i).TitleAuthor
'        End If
'        End With
'        If oPO.polines(i).Note > "" Or oPO.polines(i).lastaction > "" Then
'            Set lstItem = lvw.ListItems.Add
'            lstItem.SubItems(1) = "Note:  " & oPO.polines(i).Note & " Last action:" & oPO.polines(i).lastactionAndDate
'        End If
'    Next i
'   ' If lvw.ListItems.Count > 0 Then lvw.ListItems(1).Selected = True
'
'
'EXIT_Handler:
'    Exit Sub
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'End Sub
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
Dim tmp
    Set XA = New XArrayDB
    XA.Clear
    On Error GoTo ERR_Handler
    XA.ReDim 1, oPO.polines.Count, 1, 15
    For i = 1 To oPO.polines.Count
    
    
        With oPO.polines(i)
                XA(i, 14) = .POLID & "k"
                XA(i, 15) = .fulfilled
                XA(i, 1) = .ProductCodeF
                XA(i, 2) = .TitleAuthor
                XA(i, 3) = .Ref
                XA(i, 4) = .QtyFirm
                XA(i, 5) = .Qtyseesafe
                XA(i, 6) = .QtyReceivedSoFar
                XA(i, 7) = .PriceF(oPO.isFOreignCurrency)
                XA(i, 9) = .PLessDiscExtF(oPO.isFOreignCurrency)
                XA(i, 8) = .DiscountF
                XA(i, 10) = .ETAF
                XA(i, 11) = "Note:  " & .Note & " Last action:" & .lastactionAndDate
                XA(i, 12) = .ProductCode
                XA(i, 13) = .pID
        End With
    Next i
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", "frmPOPreview", CStr(i), G1.Columns(i - 1).Width)
    Next
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind

    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub


Private Sub Form_Load()
    
    Me.top = 50
    Me.left = 50
    Me.Height = 6500
    Me.Width = 11500

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnsetMenu
    Set oPO = Nothing
End Sub



Private Sub cbTP_Click()
Dim frm As frmSupplierPreview
    Set frm = New frmSupplierPreview
    frm.Component oPO.Supplier
    frm.Show
End Sub

Private Sub G1_Click()
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 12))
    If str = "" Then Exit Sub
    Clipboard.SetText str
End Sub
Private Sub G1_DblClick()
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim oP As a_Product
Dim str As String
    str = FNS(XA.Value(G1.Bookmark, 13))
    If str = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load str, 0 'oDel.DeliveryLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).pID, 0
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


Public Sub mnuCancel()
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelPO oPO
    RefreshData
    Screen.MousePointer = vbDefault
End Sub

Public Sub mnuCancelLine()
Dim oP As a_Product
    If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oPO.polines.FindLineByID(val(XA(G1.Bookmark, 14))).CancelLine
    RefreshData
    Screen.MousePointer = vbDefault
End Sub


Public Sub mnuVoid()
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oPO.VoidDocument
    RefreshData
End Sub
Public Sub RefreshData()
    oPO.Reload
    LoadControls
End Sub

Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XA(Bookmark, 15) = "CAN" Then
        RowStyle.BackColor = &HC0C0C0
    End If
'    If XA(Bookmark, 15) = "IN PROCESS" Then
'        RowStyle.BackColor = &H80FF80
'    End If
'    If XA(Bookmark, 15) = "COMPLETE" Then
'        RowStyle.BackColor = &HFFFFC0
'    End If
End Sub

Private Sub G1_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 3, 4
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
End Function
