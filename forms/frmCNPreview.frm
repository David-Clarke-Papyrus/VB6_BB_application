VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmCNPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Credit note preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11430
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmCNPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11430
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
      Height          =   615
      Left            =   2310
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCNPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Close the credit note"
      Top             =   4875
      Width           =   1000
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
      Left            =   9435
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   420
      Width           =   1635
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
      Height          =   1320
      Left            =   3375
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4830
      Width           =   3135
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
      TabIndex        =   8
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
      Height          =   615
      Left            =   1275
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCNPreview.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print or preview"
      Top             =   4875
      Width           =   1000
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
      Left            =   2085
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
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
      Height          =   615
      Left            =   255
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCNPreview.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   1000
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
      Top             =   240
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   2640
      Left            =   300
      OleObjectBlob   =   "frmCNPreview.frx":3240
      TabIndex        =   17
      Top             =   2160
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1110
      Left            =   240
      TabIndex        =   19
      Top             =   915
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   1958
      BackColor       =   -2147483638
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
      TabIndex        =   18
      Top             =   555
      Width           =   2970
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   375
      TabIndex        =   16
      Top             =   1050
      Width           =   2730
   End
   Begin VB.Label lblTPPhone 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   375
      TabIndex        =   15
      Top             =   1365
      Width           =   2730
   End
   Begin VB.Label lblTPFax 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   375
      TabIndex        =   14
      Top             =   1680
      Width           =   2730
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
      TabIndex        =   12
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
      Left            =   6525
      TabIndex        =   10
      Top             =   4920
      Width           =   2490
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   255
      TabIndex        =   7
      Top             =   5505
      Width           =   450
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
      Height          =   810
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   60
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
Attribute VB_Name = "frmCNPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cCN As c_CNs
Dim oCN As a_CN
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
        ' Print "You pressed the SHIFT key."
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True

      End Select
End Sub

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.mnuSaveLayout"
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oCN.StatusF = "IN PROCESS" And oCN.IsNew = False)
    Forms(0).mnuCancel.Enabled = False  '(oCN.statusF = "ISSUED")
    Forms(0).mnuCancelLine.Enabled = False '(oCN.statusF = "ISSUED" And oCN.IsNew = False)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.SetMenu"
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
Dim i As Integer

    lngID = PID
    Set oCN = New a_CN
    oCN.Load lngID, True
    Me.Caption = "Credit note (preview) for " & oCN.Customer.NameAndCode(25) & oCN.StaffNameB
    LoadControls
    SetMenu

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pInvoice As a_CN)
    On Error GoTo errHandler
    Set oCN = pInvoice
    Me.Caption = "Credit note (preview) for " & oCN.Customer.NameAndCode(25) & oCN.StaffNameB
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.ComponentObject(pInvoice)", pInvoice
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
    
        With oCN
            Me.txtDate = .DocDateF
            If DateDiff("d", .DOCDate, .IssDate) > 1 Then
                lblSI.Caption = "Issued: " & .IssDateF
            Else
                lblSI.Caption = ""
            End If
            Me.txtInvoiceNum = .DOCCode
            Me.txtStatus = .StatusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If .Status = stInProcess Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtInvoiceNum = .DOCCode
            Me.lblTPName = .TPNAME & IIf(Len(.TPACCNum) > 0, " (" & .TPACCNum & ")", "")
            If Not .Customer.BillTOAddress Is Nothing Then
                Me.lblTPPhone = .Customer.BillTOAddress.Phone
                Me.lblTPFax = .Customer.BillTOAddress.Fax
            End If
            Me.txtTPMemo = IIf(Len(.TPMemo) > 0, .TPMemo, "")
            txtTPMemo.Visible = (txtTPMemo > "")
            If .BillToAddressID > 0 Then
                If Not .BillTOAddress Is Nothing Then
                    strAddress = .BillTOAddress.AddressMailing
                End If
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .GoodsAddressID > 0 Then
                If Not .DelToAddress Is Nothing Then
                    strAddress = .DelToAddress.AddressMailing
                End If
            End If
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
        End With
        LoadGrid
        'LoadListView
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.LoadControls"
End Sub



Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    
    If oCN.Customer.ID > 0 Then
        frm.component oCN.Customer
        frm.Show
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_CN

Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oCN.CNLines.SortLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oCN.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oCN.constDOCCODE).QtyCopies
        End If

       If oCN.ExportToXML(qtyLinesToPrint, , enView, , True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Set frm = New frmPrintingOptions_CN
        frm.component oCN
        frm.Show vbModal
        LoadGrid
    End If
EXIT_Handler:
 '   Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frm As frmCN
    WaitMsg "Loading . . .", True, Me
    Set frm = New frmCN
    blnEdit = True
    frm.component , oCN
    frm.Show
    WaitMsg "", False, Me

EXIT_Handler:
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.cmdEdit_Click", , EA_NORERAISE
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

    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oCN.CNLines.Count, 1, 13
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
  '  G1.Columns(2).Width = 333
    For i = 1 To oCN.CNLines.Count
            XA(i, 11) = oCN.CNLines(i).Key
            XA(i, 12) = oCN.CNLines(i).ProductCode
            XA(i, 13) = oCN.CNLines(i).EAN
            XA(i, 1) = oCN.CNLines(i).ProductCodeF
            XA(i, 2) = oCN.CNLines(i).TitleAuthorPublisher
            XA(i, 3) = oCN.CNLines(i).INVLineCode
            XA(i, 4) = oCN.CNLines(i).QtyComboF
            XA(i, 5) = oCN.CNLines(i).PriceF(False)
            XA(i, 6) = oCN.CNLines(i).DiscountPercentF
            XA(i, 7) = oCN.CNLines(i).PLessDiscExtF(False)
            XA(i, 10) = oCN.CNLines(i).PID
            If oCN.CNLines(i).Note > "" Then
                XA(i, 8) = "Note:  " & oCN.CNLines(i).Note
                G1.Columns(7).Width = 4000
            Else
                G1.Columns(7).Caption = ""
                XA(i, 8) = ""
                G1.Columns(7).Width = 0
            End If
            
    Next i
    
    G1.Array = XA
    G1.ReBind

    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.LoadGrid"
End Sub

'Private Sub LoadSummary(pPostage As Currency, pVAT As Double, pConversionRate As Double, pCurrFormat As String, curTotalValue As Currency, curTotalDeposits As Currency)
'Dim currPrice As Currency
'Dim strDiscount As String
'Dim dblVAT As Double
'Dim strSummaryDescription As String
'Dim strSummary As String
'    dblVAT = (curTotalValue / (1 + pVAT)) * pVAT
'    If pPostage = 0 And curTotalDeposits = 0 And oCN.VATAble Then
'        strSummaryDescription = "(Includes VAT of " & Format(dblVAT, pCurrFormat) & ")          Total: "
'        strSummary = Format(curTotalValue, pCurrFormat)
'    ElseIf oCN.VATAble Then
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
'    ElseIf Not oCN.VATAble Then
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
    ErrorIn "frmCNPreview.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Height = 6500
        Me.Width = 11500
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 550))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1700))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
 '   lblDescription.Top = lblDescription.Top + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oCN = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    'str = FNS(XA.Value(G1.Bookmark, 12))
    str = IIf(FNS(XA.Value(G1.Bookmark, 13)) > "", FNS(XA.Value(G1.Bookmark, 13)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 13)) > "", FNS(XA.Value(G1.Bookmark, 13)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
 '   Forms(0).mnuCancelLine.Enabled = oCO.COLines(str).QtyDispatched = 0
    str = FNS(XA.Value(G1.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oCN.CNLines(str).PID, 0
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
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmCNPreview: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmCNPreview: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.GetRowType(ColIndex)", ColIndex
End Function


Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub



Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelCreditNote oCN
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.mnuCancel"
End Sub


Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oCN.VoidDocument
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oCN.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.RefreshData"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.txtTPMemo_Change", , EA_NORERAISE
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
    ErrorIn "frmCNPreview.txtTPMemo_DblClick", , EA_NORERAISE
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
    ErrorIn "frmCNPreview.txtTPMemo_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)

'    If InStr(1, txtTPMemo, Chr(13)) > 0 Then
'        If MsgBox("There are multiple lines in the memo you are saving.", vbExclamation + vbOKCancel, "Warning") = vbCancel Then
'            Cancel = True
'            Exit Sub
'        End If
'    End If
Dim oSM As New z_StockManager
    oSM.SetMemo txtTPMemo, oCN.TRID
    oCN.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmCNPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
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
    oSM.SetMemo txtTPMemo, oCN.TRID
    oCN.SetMemo txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCNPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub



