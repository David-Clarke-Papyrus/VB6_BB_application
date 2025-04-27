VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmDELPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Goods received note preview"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   13545
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmDELPreviewB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   13545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAdditionalcharges 
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
      Height          =   315
      Left            =   6900
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5475
      Width           =   2115
   End
   Begin VB.CommandButton cmdCustALloc 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Customer fulfilments"
      Height          =   405
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5355
      Width           =   1755
   End
   Begin VB.CommandButton cmdLabels 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Print product labels"
      Height          =   405
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5355
      Width           =   1755
   End
   Begin VB.TextBox txtTPMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   1080
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4710
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
      Height          =   600
      Left            =   1905
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDELPreviewB.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close the G.R.N."
      Top             =   4725
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
      Height          =   315
      Left            =   9345
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   1560
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
      Height          =   600
      Left            =   1035
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDELPreviewB.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print or preview"
      Top             =   4725
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
      Height          =   600
      Left            =   150
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmDELPreviewB.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print the invoice"
      Top             =   4725
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
      Height          =   315
      Left            =   9390
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   1515
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
      TabIndex        =   0
      Top             =   210
      Width           =   1545
   End
   Begin CoolButtonControl.CoolButton cbSupp 
      Height          =   870
      Left            =   3825
      TabIndex        =   8
      Top             =   60
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1535
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
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3585
      Left            =   120
      OleObjectBlob   =   "frmDELPreviewB.frx":3240
      TabIndex        =   13
      Top             =   1065
      Width           =   12135
   End
   Begin VB.Label lblSI 
      Alignment       =   1  'Right Justify
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
      Left            =   675
      TabIndex        =   1
      Top             =   540
      Width           =   2970
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1110
      X2              =   2775
      Y1              =   0
      Y2              =   990
   End
   Begin VB.Label txtTPFax 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   12
      Top             =   540
      Width           =   2265
   End
   Begin VB.Label txtTPPhone 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   11
      Top             =   225
      Width           =   2265
   End
   Begin VB.Label txtTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Height          =   510
      Left            =   3945
      TabIndex        =   10
      Top             =   240
      Width           =   2910
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   9345
      TabIndex        =   6
      Top             =   4830
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
Dim XA As XArrayDB
Dim bMemoExpanded As Boolean

Dim PrintCommandButtonCTRLDown As Boolean
Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
End Sub
Private Sub cmdPrint_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim ShiftTest As Integer
   PrintCommandButtonCTRLDown = False
   ShiftTest = Shift And 7
   Select Case ShiftTest
      Case 1 ' or vbShiftMask
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True
      End Select
End Sub

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error Resume Next
    Forms(0).mnuVoid.Enabled = (oDel.Status = stInProcess And oDel.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oDel.Status = stISSUED Or oDel.Status = stCOMPLETE)
    Forms(0).mnuCancelLine.Enabled = False ' FOR NOW . . . ((oDel.Status = stISSUED Or oDel.Status = stCOMPLETE) And oDel.IsNew = False)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuDelact.Enabled = True
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.SetMenu"
End Sub

'Public Sub PrintSupplierClaims()
'    If IsNull(Grid.Bookmark) Then Exit Sub
'    Screen.MousePointer = vbHourglass
'    Set frm = New frmSCPreview
'    frm.Component oDel.TRID, FNS(XA.Value(Grid.Bookmark, 5)), FNS(XA.Value(Grid.Bookmark, 4)), FNS(XA.Value(Grid.Bookmark, 1)), _
'        FNS(XA.Value(Grid.Bookmark, 10)), FNN(XA.Value(Grid.Bookmark, 8))
'    frm.Show
'    Screen.MousePointer = vbDefault
'
'End Sub

Public Sub CustomerAllocations()
    On Error GoTo errHandler
Dim cCOLALLOC As chex_COLAllocation
Dim frmAlloc As frmCOLAllocation_FromDel
    Set cCOLALLOC = Nothing
    Set cCOLALLOC = New chex_COLAllocation
    cCOLALLOC.Load oDel.TRID, True
    If cCOLALLOC.Count > 0 Then
        Set frmAlloc = New frmCOLAllocation_FromDel
        frmAlloc.component cCOLALLOC, "DELIVERY", True
        frmAlloc.Show
    Else
        MsgBox "There were no allocations for this GRN", , "Information"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.CustomerAllocations"
End Sub








Private Sub cmdCustALloc_Click()
    On Error GoTo errHandler
Dim oSQL As z_SQL
Dim rs As ADODB.Recordset
Dim rpt As arCOLSFulfilled

        Set oSQL = New z_SQL
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        
        oSQL.GetDynamicRecordset_Improved "SELECT * FROM vDeliveryAllocations WHERE TR_ID = " & CStr(oDel.TRID), enText, Array(), "", rs
        Set rpt = New arCOLSFulfilled
        rpt.Printer.Orientation = ddOPortrait
        rpt.component rs
        rpt.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cmdCustALloc_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
    lngID = PID
    Set oDel = New a_Delivery
    oDel.Load lngID
    Me.Caption = "Goods received (preview) from " & oDel.TPNAME & oDel.StaffNameB
    If DateDiff("d", oDel.DOCDate, oDel.ProcessingDate) > 1 Then
        Me.Caption = Me.Caption & " Issued: " & oDel.ProcessingDateF
    End If
    LoadControls
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pDelivery As a_Delivery)
    On Error GoTo errHandler
    Set oDel = pDelivery
    oDel.CalculateTotals
    Me.Caption = "Goods received (preview) from " & oDel.TPNAME & oDel.StaffNameB
    If DateDiff("d", oDel.DOCDate, oDel.ProcessingDate) > 1 Then
        Me.Caption = Me.Caption & " Issued: " & oDel.ProcessingDateF
    End If
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.ComponentObject(pDelivery)", pDelivery
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim dblVAT As Double
Dim dblConversionRate As Double
Dim strCurrencyFormat As String
Dim curTotalDeposits As Currency
Dim curTotalValue As Currency
Dim strAddress As String
Dim strTotalCaption As String
Dim strTotalValues As String
    
        With oDel
            Me.txtDate = .DocDateF
            If DateDiff("d", .DOCDate, .IssDate) > 1 Then
                lblSI.Caption = "Issued: " & .IssDateF
            Else
                lblSI.Caption = ""
            End If
            Me.txtStatus = .StatusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            cmdEdit.Enabled = .Status = stInProcess
            cmdLabels.Enabled = (.Status = stCOMPLETE Or .Status = stISSUED)
            Me.txtInvoiceNum = .DOCCode
            Me.txtTPName = .Supplier.NameAndCode(24)
            If Not .Supplier.OrderToAddress Is Nothing Then
                Me.txtTPPhone = "Phone: " & .Supplier.OrderToAddress.Phone
                Me.txtTPFax = "Fax: " & .Supplier.OrderToAddress.Fax
            End If
            Me.txtCurrency = oDel.CaptureCurrency.Description & IIf(oDel.CaptureCurrency Is oPC.Configuration.DefaultCurrency, "", "(" & oDel.CurrencyConversionInverseRate & ")")
            Me.txtAdditionalcharges = IIf(oDel.BatchTotalExtras > 0, "Extra charges: " & oDel.BatchTotalExtrasF, "")
            Me.lblTotalValues = oDel.TotalLessDiscExtF(oDel.ISForeignCurrency)
            Me.Caption = "Goods received (preview) from " & oDel.TPNAME & oDel.StaffNameB & "   (" & .SupplierInvoiceRef & " : " & .SupplierInvoiceDateF & ")"
            Me.txtTPMemo = IIf(Len(.Memo) > 0, "Note:  " & Trim$(.Memo), "")
            txtTPMemo.Visible = (txtTPMemo > "")
        End With
        LoadGrid
        
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.LoadControls"
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
'Dim frm As frmPreview_
'    oDEL.PrintInvoice_Display True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbSupp_Click()
    On Error GoTo errHandler
Dim frm As New frmSupplierPreview
    frm.component oDel.Supplier
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cbSupp_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_DEL
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oDel.DeliveryLines.SortLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oDel.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oDel.constDOCCODE).QtyCopies
        End If
       If oDel.ExportToXML(enView, , , qtyLinesToPrint, True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
    
        Set frm = New frmPrintingOptions_DEL
        frm.ComponentObject oDel
        frm.Show vbModal
    End If
EXIT_Handler:
 '   Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frmBB As frmdelBB
Dim frmDelStyle2 As frmdel_Style2

Dim bCancel As Boolean
Dim strPreviousStatusBarCaption As String

    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).text
    Forms(0).SB1.Panels(2).text = "LOADING . . ."
    blnEdit = True
    
    If oPC.UniqueProducts Then
        Set frmDelStyle2 = New frmdel_Style2
        frmDelStyle2.component False, , oDel
        frmDelStyle2.Show
    Else
        Set frmBB = New frmdelBB
        frmBB.component False, , oDel
        frmBB.Show
    End If
    Unload Me
    
    Forms(0).SB1.Panels(2).text = strPreviousStatusBarCaption

EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
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
    XA.ReDim 1, oDel.DeliveryLines.Count, 1, 20
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", "frmDelPreview", CStr(i), G1.Columns(i - 1).Width)
    Next
    For i = 1 To oDel.DeliveryLines.Count
        With oDel.DeliveryLines(i)
                XA(i, 15) = .PID
                XA(i, 1) = .CodeF
                XA(i, 2) = .Title
                XA(i, 3) = .Ref
                XA(i, 4) = .POCode
                tmp = .POLQtyFirm
                XA(i, 5) = .QtyFirm & IIf(tmp > 0, "(", "") & IIf(tmp > 0, tmp, "") & IIf(tmp > 0, ")", "")
                tmp = .POLQtySS
                XA(i, 6) = .QtySS & IIf(tmp > 0, "(", "") & IIf(tmp > 0, tmp, "") & IIf(tmp > 0, ")", "")
                If .ReasonID > "" Then
                XA(i, 7) = .QtyShort & IIf(.ReasonID > 0, " (" & oPC.Configuration.ReturnReasons.f3(.ReasonID) & ")", "")
                Else
                    XA(i, 7) = ""
                End If
                tmp = .POLDiscount
                XA(i, 8) = .PriceF(oDel.ISForeignCurrency) & IIf(tmp > 0, "(", "") & IIf(tmp > 0, .POLPriceF(oDel.ISForeignCurrency), "") & IIf(tmp > 0, ")", "")
                XA(i, 9) = .PriceSellF
                XA(i, 10) = oPC.Configuration.Multibuys.ItemByF4(.MBCode)
                XA(i, 11) = .DiscountF & IIf(tmp > 0, "(", "") & IIf(tmp > 0, .POLDiscountF, "") & IIf(tmp > 0, ")", "")
                tmp = .POLPrice
                XA(i, 12) = .PLessDiscExtF(oDel.ISForeignCurrency)
                XA(i, 13) = .Note
                XA(i, 14) = .code
                XA(i, 16) = .DELLID
                XA(i, 17) = .EAN
                XA(i, 18) = .DELLID
            
        End With
    Next i
    XA.QuickSort 1, XA.UpperBound(1), 16, XORDER_ASCEND, XTYPE_NUMBER
    G1.Array = XA
    G1.ReBind
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.LoadGrid"
End Sub


Private Sub Command1_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If KeyCode = vbKey4 Then
        If MsgBox("Confirm close?", vbOKCancel, "Close form") = vbOK Then
            Unload Me
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    If Me.WindowState <> 2 Then
'        Me.TOP = 50
'        Me.Left = 50
'        Me.Height = 6500
'        Me.Width = 11500
        SetFormSize Me
    End If
    
    cmdCustALloc.Visible = Not oPC.IncludeSupplierFeatures  ' this is a retail environment and customer orders are held back at counter, not invoiced immediately

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error Resume Next
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 400))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1720))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
    cmdLabels.TOP = cmdLabels.TOP + lngDiff
    cmdCustALloc.TOP = cmdCustALloc.TOP + lngDiff
    txtAdditionalcharges.TOP = txtAdditionalcharges.TOP + lngDiff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oDel = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub Label5_DblClick()
    On Error GoTo errHandler
Dim frm As frmSupplierPreview
    Set frm = New frmSupplierPreview
    frm.component oDel.Supplier
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.Label5_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
  '  str = FNS(XA.Value(G1.Bookmark, 12))
    str = IIf(FNS(XA.Value(G1.Bookmark, 17)) > "", FNS(XA.Value(G1.Bookmark, 17)), FNS(XA.Value(G1.Bookmark, 14)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      PopupMenu Forms(0).mnuDelivery   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub


Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 17)) > "", FNS(XA.Value(G1.Bookmark, 17)), FNS(XA.Value(G1.Bookmark, 14)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frmA As frmProductPrevAQ
Dim frm As frmProductPrev
Dim frmU As frmProductSinglePreview
Dim oP As a_Product
Dim str As String
    str = FNS(XA.Value(G1.Bookmark, 15))
    If str = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load str, 0 'oDel.DeliveryLines.FindLineByID(val(Me.lvw.SelectedItem.Key)).pID, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.component oP
        frmA.Show
    Else
        If oPC.UniqueProducts Then
            Set frmU = New frmProductSinglePreview
            frmU.component oP
            frmU.Show
        Else
            Set frm = New frmProductPrev
            frm.component oP
            frm.Show
        End If
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmDELPreview: G1_DblCLick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmDELPreview: G1_DblCLick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub



Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager

    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_GRN_SIGN, , "Cancel this GRN", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oDel.Status = stInProcess Then
            If MsgBox("Do you want to cancel this GRN?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If

    'If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelDEL oDel
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.mnuCancel"
End Sub
'Public Sub mnuCancelLine()
'    If MsgBox("Do you want to cancel this document line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
'    Screen.MousePointer = vbHourglass
'    oSM.CancelDELL FNN(XA.Value(G1.Bookmark, 14))
'    RefreshData
'    Screen.MousePointer = vbDefault
'End Sub

Public Sub RefreshData()
    On Error GoTo errHandler
    oDel.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.RefreshData"
End Sub


Public Sub mnuVoid()
    On Error GoTo errHandler
    
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_GRN_SIGN, , "Void this GRN", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If oDel.Status = stInProcess Then
            If MsgBox("Do you want to void this GRN?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    
  '  If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oDel.VoidDocument
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.mnuVoid"
End Sub


Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If XA Is Nothing Then Exit Sub
    On Error Resume Next
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.GetRowType(ColIndex)", ColIndex
End Function

'Public Sub PrintSupplierClaim()
'    On Error GoTo errHandler
'    If oDel.HasSupplierClaim Then
'        oDel.PrintSupplierClaim
'    Else
'        MsgBox "There is no discrepancy information to report.", vbInformation, "Can't do this"""
'    End If
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmDELPreview.PrintSupplierClaim"
'End Sub

Public Sub PrintLabels()
    On Error GoTo errHandler
Dim frm As frmPrintLabels
    Set frm = New frmPrintLabels
    frm.component "D", oDel
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.PrintLabels"
End Sub

Private Sub G1_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    MsgBox "Column = " & ColIndex
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.G1_ButtonClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLabels_Click()
    On Error GoTo errHandler
    PrintLabels
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.cmdLabels_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler

    txtTPMemo = HandleTextWithBites(txtTPMemo)

'Dim strArg As String
'Dim iStart As Integer
'Dim iEnd As Integer
'Dim oU As New z_UTIL
'Dim strResult As String
'Dim f As frmFindTextBite
'
'    iStart = 0
'    iEnd = 0
'    iStart = InStr(1, txtTPMemo, "?") + 1
'    If iStart = 0 Then Exit Sub
'    strResult = ""
'    iEnd = InStr(iStart, txtTPMemo, "?")
'    If iStart > 0 And iEnd > iStart Then
'        strArg = Trim(Mid(txtTPMemo, iStart, iEnd - iStart))
'        strResult = oU.GetTextBite(strArg)
'        If strResult > "" Then
'            txtTPMemo = Replace(txtTPMemo, "?" & strArg & "?", strResult)
'        End If
'    Else
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.txtTPMemo_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DblClick()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.TOP = txtTPMemo.TOP + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    Else
        bMemoExpanded = True
        txtTPMemo.Height = txtTPMemo.Height + 800
        txtTPMemo.Width = txtTPMemo.Width + 800
        txtTPMemo.TOP = txtTPMemo.TOP - 800
        txtTPMemo.ZOrder 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.txtTPMemo_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_LostFocus()
    On Error GoTo errHandler
    If bMemoExpanded Then
        txtTPMemo.Height = txtTPMemo.Height - 800
        txtTPMemo.Width = txtTPMemo.Width - 800
        txtTPMemo.TOP = txtTPMemo.TOP + 800
        bMemoExpanded = False
        txtTPMemo.ZOrder 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.txtTPMemo_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)

'    If InStr(1, txtTPMemo, Chr(14)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oDel.TRID
    'oDel.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragOver(Source As Control, x As Single, _
    Y As Single, State As Integer)
    On Error GoTo errHandler
    Dim picdocument As PictureBox
        ' Optionally move the cursor position so
        ' the user can see where the drop would happen.
        txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
        txtTPMemo.SelLength = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_DragDrop(Source As Control, x As Single, _
    Y As Single)
    On Error GoTo errHandler
    txtTPMemo.SelStart = TextBoxCursorPos(txtTPMemo, x, Y)
    txtTPMemo.SelLength = 0
    txtTPMemo.SelText = Source
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oDel.TRID
    oDel.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub

Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_DeliveryLine
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.Open
    For Each oLine In oDel.DeliveryLines
 '       If Not oLine.Product.IsServiceItem Then
        rs.AddNew
        rs.Fields("GUID") = CreateGUID
        rs.Fields("PID") = oLine.PID
        rs.Fields("Qty") = oLine.QtyFirm
        rs.Fields("QtyFirm") = oLine.QtyFirm + oLine.QtySS
        rs.Fields("QtySS") = oLine.QtySS
        rs.Fields("Price") = oLine.Price(oDel.ISForeignCurrency)
        rs.Fields("DISCOUNTRATE") = oLine.Discount
        rs.Fields("CODEF") = oLine.CodeF
        rs.Fields("EANF") = oLine.EAN
        rs.Fields("TITLE") = oLine.Title
        rs.Fields("VATRATE") = oLine.VATRate
        rs.Fields("REF") = oLine.Ref
        rs.Fields("ETA") = CDate(0)
        rs.Update
  '      End If
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        On Error Resume Next
        fs.CreateFolder (oPC.SharedFolderRoot & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
            rs.Save oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.mnuCopyLines"
End Sub

Public Sub mnuOpenClaim()
    On Error GoTo errHandler
Dim frm As New frmSupplierRetFromDelivery
Dim olDELL As a_DeliveryLine
Dim olDEL As a_Delivery
Dim sClaim As String
Dim oSQL As z_SQL
Dim ClaimTRID As Long

    Set olDEL = New a_Delivery
    olDEL.Load oDel.TRID
    olDEL.BeginEdit
        Set olDELL = olDEL.DeliveryLines.FindLineByID(XA(G1.Bookmark, 16))
    
        frm.component olDELL.Discount, olDELL.Price(oDel.ISForeignCurrency), olDELL.QtyShort, olDELL.ReasonID, _
            olDELL.CorrectedDiscount, olDELL.CorrectedPrice(olDEL.ISForeignCurrency)
        
        frm.Show vbModal
        If Not frm.IsCancelled Then
            If Not adjustQtysReceived(olDELL, frm.QtyClaim) Then
                MsgBox "You have claimed more items than you received or this claim has already been processed - check qty on hand.", vbOKOnly + vbInformation, "Can't do this"
                olDEL.CancelEdit
                Set olDEL = Nothing
                Exit Sub
            End If
            Set oSQL = New z_SQL
            oSQL.CreateSupplierClaim olDEL.TPID, ClaimTRID
            sClaim = ""
            olDELL.ReasonID = frm.Reasons
            olDELL.SetQtyShort frm.QtyClaim
            olDELL.ClaimID = ClaimTRID
            If frm.QtyClaim > 0 Then
                sClaim = IIf(frm.QtyClaim > 0, "+", "-") & (frm.QtyClaim)
            End If
            If frm.CorrectedDiscount <> olDELL.Discount Then
                olDELL.SetCorrectedDiscount frm.CorrectedDiscount
                sClaim = sClaim & IIf(sClaim > "", ", ", "") & frm.CorrectedDiscountF
            End If
            If frm.CorrectedPrice <> olDELL.Price(olDEL.ISForeignCurrency) Then
                olDELL.SetCorrectedPrice frm.CorrectedPrice
                sClaim = sClaim & IIf(sClaim > "", ", ", "") & CStr(frm.CorrectedPrice)
            End If
            sClaim = "CL:" & sClaim
        End If
        olDELL.Note = sClaim
        
        olDEL.ApplyEdit
        Set olDEL = Nothing
        Unload frm

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.mnuOpenClaim"
End Sub
Private Function adjustQtysReceived(oDL As a_DeliveryLine, lngQtyClaimed As Long) As Boolean
Dim lngQtyRem As Long
Dim lngDiff As Long
    lngDiff = oDL.QtyTotal - (oDL.QtyFirm + oDL.QtySS)
    If lngDiff > 0 Then
        oDL.SetQtySS oDL.QtySS + lngDiff
    End If
    adjustQtysReceived = True
    If oDL.QtyFirm > 0 Then
        If oDL.QtyFirm >= lngQtyClaimed Then
            oDL.SetQtyFirm (oDL.QtyFirm - lngQtyClaimed)
            Exit Function
        Else
            lngQtyRem = lngQtyClaimed - oDL.QtyFirm
            oDL.SetQtyFirm 0
        End If
    End If
    If oDL.QtySS > 0 Then
        oDL.SetQtySS oDL.QtySS - lngQtyRem
        If oDL.QtySS < 0 Then
            adjustQtysReceived = False
        End If
    Else
        adjustQtysReceived = False
    End If
        
        
End Function


Public Sub PrintSupplierClaim()
    On Error GoTo errHandler
Dim frm As frmBrowseSupplierClaims

    Screen.MousePointer = vbHourglass
    Set frm = New frmBrowseSupplierClaims
    frm.component oDel.DOCCode
  '  frm.Component oDel.TRID, 0, 0, oDel.Supplier.NameAndCode(60), oDel.Supplier.ClaimNeedsApproval, oDel.Supplier.ID
    frm.Show
    Screen.MousePointer = vbDefault
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDELPreview.PrintSupplierClaim"
End Sub
