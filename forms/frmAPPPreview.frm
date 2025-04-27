VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmAPPPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appro preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11430
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmAPPPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdInvoice 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8970
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Generate invoice from this appro"
      Top             =   435
      Width           =   960
   End
   Begin CoolButtonControl.CoolButton cb1 
      Height          =   915
      Left            =   3765
      TabIndex        =   19
      Top             =   15
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   1614
      BackColor       =   13882315
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
      Left            =   255
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   540
      Width           =   3390
   End
   Begin VB.CommandButton cmdSlips 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Print slips"
      Height          =   375
      Left            =   255
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5625
      Width           =   1560
   End
   Begin VB.CommandButton cmdReturn 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9945
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Generate return from this appro"
      Top             =   435
      Width           =   960
   End
   Begin VB.TextBox txtBillTo 
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
      Height          =   810
      Left            =   7155
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   15
      Width           =   1680
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
      Picture         =   "frmAPPPreview.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Close the appro"
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
      Height          =   315
      Left            =   9435
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   45
      Width           =   1455
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
      Height          =   1140
      Left            =   2895
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4755
      Visible         =   0   'False
      Width           =   3135
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
      Picture         =   "frmAPPPreview.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Print or preview"
      Top             =   4875
      Width           =   855
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
      Picture         =   "frmAPPPreview.frx":085E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
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
      Height          =   300
      Left            =   2100
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   210
      Width           =   1545
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3750
      Left            =   210
      OleObjectBlob   =   "frmAPPPreview.frx":0BE8
      TabIndex        =   16
      Top             =   1005
      Width           =   10725
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1140
      X2              =   2760
      Y1              =   0
      Y2              =   915
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Left            =   3870
      TabIndex        =   14
      Top             =   45
      Width           =   270
   End
   Begin VB.Label txtName 
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
      Left            =   4305
      TabIndex        =   13
      Top             =   90
      Width           =   2580
   End
   Begin VB.Label txtPhone 
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
      Left            =   4305
      TabIndex        =   12
      Top             =   510
      Width           =   2580
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Enabled         =   0   'False
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
      Height          =   1335
      Left            =   6405
      TabIndex        =   8
      Top             =   4860
      Width           =   2235
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
      Height          =   315
      Left            =   4890
      TabIndex        =   6
      Top             =   4845
      Width           =   360
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Height          =   1335
      Left            =   8760
      TabIndex        =   5
      Top             =   4860
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   780
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   120
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
Attribute VB_Name = "frmAPPPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cCN As c_CNs
Dim oAPP As a_APP
Dim dblTotal As Double
Dim XA As XArrayDB
Dim bMemoExpanded As Boolean
Dim PrintCommandButtonCTRLDown As Boolean

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
Private Sub CB1_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
   ' If flgLoading Then Exit Sub
    frm.component oAPP.Customer
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.CB1_Click", , EA_NORERAISE
    HandleError
End Sub



Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oAPP.Status = stInProcess And oAPP.IsNew = False)
    Forms(0).mnuCancel.Enabled = False '(oAPP.Status = stISSUED)
    Forms(0).mnuCancelLine.Enabled = False  '(oAPP.Status = stInProcess)   ' And oAPP.IsNew = False)
    Forms(0).mnuDelLine.Enabled = False   'oAPP.Status = stInProcess
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Forms(0).mnuPastelinestoNEW = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.SetMenu"
End Sub

Private Sub cmdReturn_Click()
    On Error GoTo errHandler
Dim frm As New frmAppro_AUTORETURN
    frm.component oAPP
    frm.Show vbModal
    RefreshData
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cmdReturn_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdReturn_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdInvoice_Click()
    On Error GoTo errHandler
Dim frm As New frmAppro_AUTOINV
    frm.component oAPP
    frm.Show vbModal
    RefreshData
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cmdInvoice_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdInvoice_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSlips_Click()
    On Error GoTo errHandler
    oAPP.PrintSlips
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cmdSlips_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdSlips_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CoolButton1_MouseEnter()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.CoolButton1_MouseEnter", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.Form_Activate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.Form_Deactivate", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
    lngID = PID
    Set oAPP = New a_APP
    oAPP.Load lngID, True
    Me.Caption = "Appro (Preview) for " & oAPP.TPNAME & oAPP.StaffNameB & "           FROM: " & oPC.Configuration.Companies.FindCompanyByID(oAPP.COMPID).CompanyName
    LoadControls
   ' cmdInvoice.Visible = Not oPC.POSActive
   ' cmdReturn.Visible = Not oPC.POSActive
    cmdInvoice.Visible = True
    cmdReturn.Visible = True
   
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.Component(pID)", PID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pAPP As a_APP)
    On Error GoTo errHandler
    Set oAPP = pAPP
    Me.Caption = "Appro (preview) for " & oAPP.TPNAME & oAPP.StaffNameB
    LoadControls
    cmdInvoice.Visible = Not oPC.POSActive
    cmdReturn.Visible = Not oPC.POSActive
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.ComponentObject(pAPP)", pAPP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.ComponentObject(pAPP)", pAPP
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
    
        With oAPP
            Me.txtDate = .DocDateF
            If DateDiff("d", .DOCDate, .IssDate) > 1 Then
                Me.txtIssued = "Issued: " & .IssDateF
            Else
                txtIssued = ""
            End If
            Me.txtInvoiceNum = .DOCCode
            Me.txtStatus = .StatusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
        '    Me.txtComp = "From: " & .BillingCompany.CompanyName
            If .Status = stInProcess Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtInvoiceNum = .DOCCode
            Me.txtName = .Customer.NameAndCode(20)
            If Not .Customer.BillTOAddress Is Nothing Then
                Me.txtPhone = .Customer.BillTOAddress.PhoneandFax
            End If
            txtTPMemo = IIf(Len(.Memo) > 0, .Memo, "")
            txtTPMemo.Visible = (txtTPMemo > "")
            If .APPROTOID > 0 Then
                If Not .ApproToAddress Is Nothing Then
                    strAddress = .ApproToAddress.AddressMailing
                End If
            End If
            txtBillTo = IIf(strAddress > "", strAddress, "unknown")
            .CalculateTotal
            
        End With
        LoadGrid
        cmdReturn.Enabled = (oAPP.Status = stISSUED)
        cmdInvoice.Enabled = (oAPP.Status = stISSUED)
        cmdSlips.Enabled = (oAPP.Status = stISSUED)
        lblTotalValues.Width = 2400
        lblTotalValues.Left = 8500
        
        lblTotalValues.Caption = oAPP.TotalNetF & vbCrLf & "Includes V.A.T. : " & oAPP.TotalVATF
'        If oPC.POSActive Then
'            cmdReturn.Enabled = oAPP.Customer.IsBookClub And oAPP.Status = stISSUED
'        End If
        
EXIT_Handler:
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.LoadControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.LoadControls"
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
'Dim frm As frmPreview_
'    oAPP.PrintInvoice_Display True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cmdPreview_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As frmCustomerPreview
    Set frm = New frmCustomerPreview
    frm.component oAPP.Customer
    frm.Show
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cbTP_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cmdClose_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_APP
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer

    
    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oAPP.ApproLines.SortLines enSequence, True
    
        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oAPP.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oAPP.constDOCCODE).QtyCopies
        End If

       If oAPP.ExportToXML(enView, , , qtyLinesToPrint, True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Set frm = New frmPrintingOptions_APP
        frm.ComponentObject oAPP
        frm.Show vbModal
    End If
EXIT_Handler:

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frm As frmAPP
    WaitMsg "Loading . . .", True, Me
    Set frm = New frmAPP
    blnEdit = True
    frm.component oAPP
    frm.Show
    WaitMsg "", False, Me

EXIT_Handler:
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.cmdEdit_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.cmdEdit_Click", , EA_NORERAISE
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
    SetGridLayout Me.G1, Me.Name
    G1.Columns(8).Width = 0 'set note width to zero by default
    XA.ReDim 1, oAPP.ApproLines.Count, 1, 13
    For i = 1 To oAPP.ApproLines.Count
        With oAPP.ApproLines(i)
            XA(i, 11) = .Key
            XA(i, 1) = .CodeF
            XA(i, 2) = .Title
            XA(i, 3) = .Ref
            XA(i, 4) = .Returns & " " & .Invoices
            XA(i, 5) = .Qty & "(" & .QtyReturned & ")"
            XA(i, 6) = .PriceF
            XA(i, 7) = .DiscountF
            XA(i, 8) = .ExtensionNetF
      '      XA(i, 9) = .GetSTatus
            XA(i, 10) = .code
            XA(i, 12) = .Qty - .QtyReturned
            XA(i, 13) = .EAN
            
            If .Note > "" Then
                XA(i, 9) = "Note:  " & .Note
                G1.Columns(8).Width = 4000
            End If
        End With
    Next i
    XA.QuickSort 1, XA.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
    G1.Array = XA
    G1.ReBind

    
EXIT_Handler:
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.LoadGrid"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Height = 6500
        Me.Width = 11500
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 400))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1620))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    cmdSlips.TOP = cmdSlips.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oAPP = Nothing
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
 '   str = FNS(XA.Value(G1.Bookmark, 10))
    str = IIf(FNS(XA.Value(G1.Bookmark, 13)) > "", FNS(XA.Value(G1.Bookmark, 13)), FNS(XA.Value(G1.Bookmark, 10)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
'    str = FNS(XA.Value(G1.Bookmark, 10))
    str = IIf(FNS(XA.Value(G1.Bookmark, 13)) > "", FNS(XA.Value(G1.Bookmark, 13)), FNS(XA.Value(G1.Bookmark, 10)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 10))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
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
    oP.Load oAPP.ApproLines(str).PID, 0
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
        LogSaveToFile "Access violation in frmAPPPreview: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmAPPPreview: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelAPP oAPP
    RefreshData
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.mnuCancel"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.mnuCancel"
End Sub
Private Sub RefreshData()
    On Error GoTo errHandler
    oAPP.Reload
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.RefreshData"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.RefreshData"
End Sub


Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oAPP.VoidDocument
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.mnuVoid"
End Sub
Private Sub G1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_INTEGER
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.GetRowType(ColIndex)", ColIndex
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.GetRowType(ColIndex)", ColIndex
End Function
'Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
'
''    If XA(Bookmark, 10) = "CAN" Then
''        RowStyle.BackColor = RGB(232, 174, 180)
''    End If
'
'End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    
    If XA(Bookmark, 12) <= 0 Then
        RowStyle.BackColor = &HFFFFC0  'RGB(232, 174, 180)
    End If
        
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPPPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    ofrm.component oAPP.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oAPP.TRID
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = ofrm.Memo
    oAPP.Memo = ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAPP.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.mnuMemo"
End Sub

Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_APPL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In oAPP.ApproLines
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        rs.fields("Qty") = oLine.Qty
        rs.fields("QtyFirm") = oLine.Qty
        rs.fields("QtySS") = 0
        rs.fields("Price") = oLine.Price
        rs.fields("DISCOUNTRATE") = oLine.Discount
        rs.fields("CODEF") = oLine.CodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oLine.VATRate
        rs.fields("REF") = oLine.Ref
        rs.Update
    Next
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
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
    ErrorIn "frmAPPPreview.mnuCopyLines"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngID As Long

    Set rs = oPC.LinesClipboard
    If rs.BOF And rs.eof Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & oAPP.DOCCode, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    rs.MoveFirst
    Do While Not rs.eof
        oAPP.PasteLine FNS(rs.fields("PID")), FNN(rs.fields("QTY")), FNN(rs.fields("PRICE")), FNDBL(rs.fields("DISCOUNTRATE")), FNDBL(rs.fields("VATRATE")), FNS(rs.fields("REF"))
        rs.MoveNext
    Loop
    
    lngID = oAPP.TRID
    Set oAPP = Nothing
    Set oAPP = New a_APP
    oAPP.Load lngID, True
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.mnuPastelines"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.txtTPMemo_Change", , EA_NORERAISE
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
    ErrorIn "frmAPPPreview.txtTPMemo_DblClick", , EA_NORERAISE
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
    ErrorIn "frmAPPPreview.txtTPMemo_LostFocus", , EA_NORERAISE
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
    oSM.SetMemo txtTPMemo, oAPP.TRID
    oAPP.Memo = txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmAPPPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
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
    oSM.SetMemo txtTPMemo, oAPP.TRID
    oAPP.Memo = txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub




