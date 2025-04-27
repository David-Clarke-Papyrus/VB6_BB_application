VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmQuotationPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Quotation"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11565
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmQuotationPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frHeader 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Header"
      ForeColor       =   &H8000000D&
      Height          =   2430
      Left            =   1185
      TabIndex        =   19
      Top             =   1530
      Visible         =   0   'False
      Width           =   7320
      Begin VB.TextBox txtForAttn 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   990
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1725
         Width           =   3240
      End
      Begin VB.TextBox txtTPMemo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   1185
         Left            =   990
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   450
         Width           =   5925
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Memo"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   150
         TabIndex        =   23
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For attn."
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   165
         TabIndex        =   22
         Top             =   1785
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdMemo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4875
      Width           =   255
   End
   Begin VB.CommandButton cmdCopyContents 
      BackColor       =   &H00C4BCA4&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   330
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
      Left            =   2055
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmQuotationPreview.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close the form"
      Top             =   4875
      Width           =   855
   End
   Begin VB.CommandButton cmdUP 
      BackColor       =   &H00C4BCA4&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4140
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      BackColor       =   &H00C4BCA4&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4470
      Width           =   330
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
      Left            =   1185
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmQuotationPreview.frx":2B2C
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Left            =   315
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmQuotationPreview.frx":2EB6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print the invoice"
      Top             =   4875
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3210
      Left            =   60
      OleObjectBlob   =   "frmQuotationPreview.frx":3240
      TabIndex        =   7
      Top             =   1530
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbDelto 
      Height          =   1425
      Left            =   6615
      TabIndex        =   10
      Top             =   30
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   2514
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
   Begin CoolButtonControl.CoolButton cbBillTo 
      Height          =   1425
      Left            =   4065
      TabIndex        =   11
      Top             =   30
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   2514
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
   Begin CoolButtonControl.CoolButton cbCust 
      Height          =   1425
      Left            =   45
      TabIndex        =   12
      Top             =   30
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   2514
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
      ShowFocusRect   =   -1  'True
      BackStyle       =   0
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   6900
      TabIndex        =   18
      Top             =   60
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   4230
      TabIndex        =   17
      Top             =   75
      Width           =   660
   End
   Begin VB.Label lblBillToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   4230
      TabIndex        =   16
      Top             =   330
      Width           =   2055
   End
   Begin VB.Label lblDelToAddress 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   6810
      TabIndex        =   15
      Top             =   330
      Width           =   2055
   End
   Begin VB.Label lblTPName 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   165
      Width           =   3600
      WordWrap        =   -1  'True
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   10020
      X2              =   11490
      Y1              =   0
      Y2              =   825
   End
   Begin VB.Label txtStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   390
      Left            =   9225
      TabIndex        =   13
      Top             =   300
      Width           =   1770
   End
   Begin VB.Label lblTotalCaption 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   6225
      TabIndex        =   3
      Top             =   4860
      Width           =   2775
   End
   Begin VB.Label lblTotalValues 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   4860
      Width           =   1845
   End
End
Attribute VB_Name = "frmQuotationPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cQU As c_QU
Dim oQU As a_QU
Dim dblTotal As Double
Dim XA As XArrayDB
Dim oSM As z_StockManager
Dim mbShowMemo As Boolean
Dim frmQ As frmQuotation

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
    SaveLayout Me.G1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuSaveLayout"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oQU.Status = stInProcess And oQU.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oQU.Status = stISSUED)
    Forms(0).mnuCancelLine.Enabled = False  '(oQU.Status = stISSUED)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuCreateCreditNote.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Forms(0).mnuHeader.Enabled = True
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Forms(0).mnuPastelinestoNEW = True
    If oPC.EmailQuote And (oQU.Status = stCOMPLETE Or oQU.Status = stISSUED) Then
        If Not oQU.Customer.BillTOAddress Is Nothing Then
            If (oQU.Customer.DispatchMethod = "M" And oQU.Customer.BillTOAddress.EMail > "") Then
                Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForQuoteEmail
                Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForQuoteEmail
            Else
                Forms(0).mnuEmail.Enabled = False
                Forms(0).mnuOutlook.Enabled = False
            End If
        Else
            Forms(0).mnuEmail.Enabled = False
            Forms(0).mnuOutlook.Enabled = False
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.SetMenu"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.SetMenu"
End Sub

Private Sub cbHeader_Click()
    On Error GoTo errHandler
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cbHeader_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cbHeader_Click", , EA_NORERAISE
    HandleError
End Sub
'Public Sub mnuHeader()
'    Header
'End Sub
'Private Sub Header()
'Dim frm As New frmHeader
'Dim strRef As String
'Dim strMemo As String
'
'    frm.Component False, "Request for quote reference", "Date", oQU.OrderNumber, oQU.OrderDateF, oQU.Memo
'    frm.Show vbModal
'    Unload frm
'
'End Sub

Private Sub cmdMemo_Click()
    On Error GoTo errHandler
    ShowMemo Not mbShowMemo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cmdMemo_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cmdMemo_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowMemo(bON As Boolean)
    On Error GoTo errHandler
        mbShowMemo = bON
        frHeader.Visible = bON
        If bON Then txtTPMemo.SetFocus
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.ShowMemo(bOn)", bOn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.ShowMemo(bOn)", bON
End Sub


Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.Form_Activate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.Form_Deactivate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
Dim strLabel As String

    lngID = PID
    Set oQU = New a_QU
    oQU.Load lngID, True
    Me.Caption = "Quotation" & "  " & oQU.DOCCode & "    " & oQU.DOCDate & " "
    If DateDiff("d", oQU.DOCDate, oQU.CaptureDate) > 1 Then
        Me.Caption = Me.Caption & " Issued: " & oQU.CaptureDateF
    End If
    Me.Caption = Me.Caption & "   " & oQU.StaffNameB
    
    If oQU.SalesRepName > "" Then
        Caption = Me.Caption & "  (Rep: " & oQU.SalesRepName & ")"
    End If
    If oPC.Configuration.Companies.Count > 1 Then
        Caption = oPC.Configuration.Companies.FindCompanyByID(oQU.COMPID).CompanyName & ": " & Me.Caption
    End If
    LoadControls
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.Component(pID)", PID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pDoc As a_QU)
    On Error GoTo errHandler
Dim strLabel As String
    Set oQU = pDoc
    Caption = "Quotation " & oQU.DOCCode & "    " & oQU.DOCDate & " "
    Me.Caption = Me.Caption & "   " & oQU.StaffNameB
    If oQU.SalesRepName > "" Then
        Me.Caption = Me.Caption & "  (Rep: " & oQU.SalesRepName & ")"
    End If
    If oPC.Configuration.Companies.Count > 1 Then
        Caption = oPC.Configuration.Companies.FindCompanyByID(oQU.COMPID).CompanyName & ": " & Me.Caption
    End If
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.ComponentObject(pDoc)", pDoc
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.ComponentObject(pDoc)", pDoc
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
    
        With oQU
            If (.Status = stInProcess) Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtStatus.Caption = .StatusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If oPC.GetProperty("CanEditQUs") = "TRUE" Then
                If .Status = stInProcess Or .Status = stISSUED Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            Else
                If .Status = stInProcess Then
                    cmdEdit.Enabled = True
                Else
                    cmdEdit.Enabled = False
                End If
            End If
            Me.txtTPMemo = IIf(Len(.Memo) > 0, .Memo, "")
            Me.txtForAttn = .ForAttn
            lblTPName.Caption = .Customer.Fullname & IIf(Len(.TPACCNum) > 0, " (" & .TPACCNum & ")", "")
            Me.txtTPMemo = IIf(Len(.Memo) > 0, FNS(.Memo), "")
            If .BillToAddressID > 0 Then
                If Not .BillTOAddress Is Nothing Then
                    strAddress = .BillTOAddress.AddressMailing
                End If
            End If
            Me.lblBillToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            If .DelToAddressID > 0 Then
                If Not .DelToAddress Is Nothing Then
                    strAddress = .DelToAddress.AddressMailing
                End If
            End If
            Me.lblDelToAddress.Caption = IIf(strAddress > "", strAddress, "unknown")
            dblConversionRate = .CurrencyFactor
            If .CurrencyFormat > "" Then
                strCurrencyFormat = .CurrencyFormat
            Else
                strCurrencyFormat = "Currency"
            End If
            .DisplayTotals strTotalCaption, strTotalValues, False
            lblTotalCaption.Caption = strTotalCaption
            lblTotalValues.Caption = strTotalValues
        End With
        LoadGrid
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_HANDLER
'Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.LoadControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.LoadControls"
End Sub


Private Sub cbCust_Click()
    On Error GoTo errHandler
Dim frm As New frmCustomerPreview
    frm.component oQU.Customer
    frm.Show
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cbCust_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cbCust_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cmdClose_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_QU
Dim i As Long

Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        oQU.QuoteLines.SortLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oQU.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oQU.constDOCCODE).QtyCopies
        End If

       If oQU.ExportToXML(True, Dummy, False, enView, qtyLinesToPrint, , , True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else

        Set frm = New frmPrintingOptions_QU
        frm.ComponentObject oQU
        frm.Show vbModal
        LoadGrid
    End If
EXIT_Handler:

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim strPreviousStatusBarCaption As String
    strPreviousStatusBarCaption = Forms(0).SB1.Panels(2).text
    Forms(0).SB1.Panels(2).text = "LOADING . . ."
    Set frmQ = New frmQuotation
    blnEdit = True
    frmQ.component , oQU
    frmQ.Show
    Unload Me
    Forms(0).SB1.Panels(2).text = strPreviousStatusBarCaption

EXIT_Handler:
   ' Unload Me
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_HANDLER
'    Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cmdEdit_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUP_Click()
    On Error GoTo errHandler
Dim i As Long
    If G1.Bookmark > 1 Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oQU.BeginEdit
        oQU.QuoteLines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark - 1, 11))
        oQU.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cmdUP_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cmdUP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdDown_Click()
    On Error GoTo errHandler
Dim i As Long
    If G1.Bookmark < XA.UpperBound(1) Then
        Screen.MousePointer = vbHourglass
        i = G1.Bookmark
        oQU.BeginEdit
        oQU.QuoteLines.swap FNS(XA.Value(G1.Bookmark, 11)), FNS(XA.Value(G1.Bookmark + 1, 11))
        oQU.ApplyEdit
        LoadGrid
        Screen.MousePointer = vbDefault
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.cmdDown_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.cmdDown_Click", , EA_NORERAISE
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
Dim oSM As New z_StockManager

    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oQU.QuoteLines.Count, 1, 19
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    G1.Columns(7).Width = 1
    For i = 1 To oQU.QuoteLines.Count
            XA(i, 11) = oQU.QuoteLines.Item(i).Key
            XA(i, 12) = oQU.QuoteLines.Item(i).code
            XA(i, 15) = oQU.QuoteLines.Item(i).PID
          '  XA(i, 16) = IIf(oQU.QuoteLines(i).SubstitutesAvailable, "Y", "N")
            XA(i, 17) = oQU.QuoteLines.Item(i).QuoteLineID
            XA(i, 18) = oQU.QuoteLines.Item(i).COLID
            XA(i, 19) = oQU.QuoteLines.Item(i).EAN
            If oQU.QuoteLines.Item(i).CodeF = "" Then
                XA(i, 1) = FormatISBN13(oQU.QuoteLines.Item(i).code)
                'XA(i, 1) = oQU.DocLines(i).code
            Else
                XA(i, 1) = oQU.QuoteLines.Item(i).CodeF
            End If
            XA(i, 2) = oQU.QuoteLines.Item(i).TitleAuthorPublisher
                XA(i, 3) = oQU.QuoteLines.Item(i).Qty
            XA(i, 4) = oQU.QuoteLines.Item(i).PriceF(False) & IIf(oQU.QuoteLines.Item(i).VATRate <> oPC.Configuration.VATRate, "v", "")
            XA(i, 5) = oQU.QuoteLines.Item(i).DiscountPercentF
            XA(i, 6) = oQU.QuoteLines.Item(i).Ref
            XA(i, 7) = oQU.QuoteLines.Item(i).ExtF(False)
            XA(i, 8) = oQU.QuoteLines.Item(i).Note
            XA(i, 10) = oQU.QuoteLines.Item(i).Sequence
            If oQU.QuoteLines.Item(i).Note > "" Then
                If oQU.QuoteLines.Item(i).Note = "Substitute" Then
                    XA(i, 9) = "Note:  " & oQU.QuoteLines.Item(i).Note & "  (Operator: right-mouse click for substitution options!)"
                Else
                XA(i, 9) = "Note:  " & oQU.QuoteLines.Item(i).Note
                End If
                G1.Columns(7).Width = 4000
            End If
            XA(i, 14) = oQU.QuoteLines.Item(i).Qty
    Next i
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), 10, 0, GetRowType(10)
    
    G1.Array = XA
    G1.ReBind

    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_HANDLER
'    Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.LoadGrid"
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.TOP = 50
        Me.Left = 50
        Me.Height = 6500
        Me.Width = 11600
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.Form_Load", , EA_NORERAISE
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
  '  cmdToCO.Top = cmdToCO.Top + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff
    cmdDown.TOP = cmdDown.TOP + lngDiff
    cmdUP.TOP = cmdUP.TOP + lngDiff
    cmdDown.Left = NonNegative_Lng(Me.Width - 540)
    cmdUP.Left = NonNegative_Lng(Me.Width - 540)
    cmdCopyContents.Left = NonNegative_Lng(Me.Width - 540)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.Form_Resize"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    If frmQ Is Nothing Then
        If oQU.IsEditing Then oQU.CancelEdit
    End If
    Set oQU = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.Form_Unload(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub G1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuInvoicePreview   ' Display the File menu as a
                        ' pop-up menu.
   End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.G1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub InsertSubstitutes()
    On Error GoTo errHandler
Dim frm As frmInsertSubstitute
Dim oIL As a_InvoiceLine
Dim str As String
Dim lngQty As Long

    If FNS(XA.Value(G1.Bookmark, 16)) <> "Y" Then
        MsgBox "There are no substitutes available for this item.", vbOKOnly + vbInformation, "Status"
        Exit Sub
    End If
    Set frm = New frmInsertSubstitute
    str = FNS(XA.Value(G1.Bookmark, 15))
    lngQty = FNN(XA.Value(G1.Bookmark, 3))
   
    frm.component oQU.Customer.NameAndCode(50), lngQty, XA.Value(G1.Bookmark, 15), XA.Value(G1.Bookmark, 18), XA.Value(G1.Bookmark, 17), oQU.QuoteID, "Q"
    frm.Show vbModal
    Unload frm
    Unload Me
    MsgBox "Substitutions have been made.", vbOKOnly, "Status"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.InsertSubstitutes"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.InsertSubstitutes"
End Sub
Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler

    If FNN(XA(Bookmark, 13)) > 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmInvoicePreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle), EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub G1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 19)) > "", FNS(XA.Value(G1.Bookmark, 19)), FNS(XA.Value(G1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    
   Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_SelChange(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    
    str = FNS(XA.Value(G1.Bookmark, 11))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load oQU.QuoteLines(str).PID, 0
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_DblClick"
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmQuotationPreview: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmQuotationPreview: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_DblClick", , EA_NORERAISE
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
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    G1.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.G1_HeadClick(ColIndex)", ColIndex
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 7, 9
            GetRowType = XTYPE_STRING
        Case 3, 4, 6, 5, 8
            GetRowType = XTYPE_INTEGER
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.GetRowType(ColIndex)", ColIndex
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.GetRowType(ColIndex)", ColIndex
End Function

'

Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oQU.CancelDocument
    RefreshData
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuCancel"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuCancel"
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    oQU.VoidDocument
    RefreshData
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oQU.Reload
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.RefreshData"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.RefreshData"
End Sub

Public Sub mnuEmail()
    On Error GoTo errHandler
Dim Res As Boolean
Dim lQU As a_QU
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If oQU.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lQU = New a_QU
        lQU.Load oQU.QuoteID, True
        Res = lQU.ExportToXML(True, strFilename, False, enMail, 1, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oQU.Customer.DispatchMethod = "E" Then
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdTransmit_Click", , EA_NORERAISE
'    HandleError
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuEmail"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuEmail"
End Sub

Public Sub mnuOutlook()
    On Error GoTo errHandler
Dim ol As Object
Dim olns As Object
Dim oMI As Object
Dim mfol As Object
Dim fol As Object
Dim fold As Object
Dim Res As Boolean
Dim lQU As a_QU
Dim pAttachmentfilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String
Dim tmp As String
Dim fs As New FileSystemObject
Dim PapyrusDraftsFolder As String
Dim OutlookParentFolder As String

    If oQU.Customer.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lQU = New a_QU
        lQU.Load oQU.QuoteID, True
        Res = lQU.ExportToXML(True, pAttachmentfilename, False, enMail, 1, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf oQU.Customer.DispatchMethod = "E" Then
    End If

    Set ol = CreateObject("Outlook.Application")
    Set olns = ol.GetNamespace("MAPI")
    
    OutlookParentFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERMAIN", "")
    PapyrusDraftsFolder = GetIniKeyValue(oPC.LocalFolder & "\PBKSWS.INI", "NETWORK", "OUTLOOKFOLDERSUB", "")
    
    If PapyrusDraftsFolder > "" Then
        Set fol = olns.Folders(OutlookParentFolder)
        Set fold = fol.Folders
        fold.Add PapyrusDraftsFolder
        Set mfol = fold(PapyrusDraftsFolder)
    End If
    
  '  Set mfol = olns.GetDefaultFolder(olFolderOutbox)
    Set oMI = ol.CreateItem(0)
    If pAttachmentfilename > "" Then
        tmp = fs.GetBaseName(pAttachmentfilename)
        strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_") - 1)
    Else
        strReference = ""
    End If
    
    With oMI
        If oPC.TestMode Then
            .To = oPC.EmailFrom
        Else
            .To = oQU.BillTOAddress.EMail
        End If
        .Subject = "Quotation: " & strReference
        .BodyFormat = 2   'HTML format
        .Body = ""
        .HTMLBody = FNS(strWholeMessage)
        
        .Attachments.Add (pAttachmentfilename)
        .ReadReceiptRequested = True
        .Close (0)  'save and close
        If PapyrusDraftsFolder > "" Then .Move mfol
    End With
    Set oMI = Nothing
    Set olns = Nothing
    Set ol = Nothing
    Set oSM = New z_StockManager
    oSM.LogTransmission oQU.QuoteID, "Sent to Outlook: " & Format(Date, "dd/mm/yyyy")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotePreview.mnuOutlook"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuOutlook"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuOutlook"
End Sub
Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_QUL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In oQU.QuoteLines
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        rs.fields("Qty") = oLine.Qty
        rs.fields("QtyFirm") = oLine.Qty
        rs.fields("QtySS") = 0
        rs.fields("Price") = oLine.Price
        rs.fields("DISCOUNTRATE") = oLine.DiscountPercent
        rs.fields("CODEF") = oLine.CodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oLine.VATRate
        rs.fields("REF") = oLine.Ref
        rs.fields("EXTRACHARGEPID") = oLine.ExtraPID
        rs.fields("EXTRACHARGEVALUE") = oLine.ExtraCharge
        rs.fields("FCPrice") = oLine.ForeignPrice
        rs.fields("FCFactor") = oLine.FCFactor
        rs.fields("FCID") = oLine.FCID
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

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuCopyLines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuCopyLines"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long

    
    Set rs = oPC.LinesClipboard
    If rs.BOF And rs.eof Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to quotation " & oQU.DOCCodeF, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    rs.MoveFirst
    Do While Not rs.eof
        oQU.PasteLine FNS(rs.fields("PID")), FNN(rs.fields("QTY")), FNN(rs.fields("PRICE")), FNDBL(rs.fields("DISCOUNTRATE")), FNDBL(rs.fields("VATRATE")), _
                    FNS(rs.fields("REF")), FNS(rs.fields("EXTRACHARGEPID")), FNN(rs.fields("EXTRACHARGEVALUE")), _
                    FNN(rs.fields("FCPRICE")), FNDBL(rs.fields("FCFACTOR")), FNN(rs.fields("FCID"))
        rs.MoveNext
    Loop
    
    lngQUID = oQU.QuoteID
    Set oQU = Nothing
    Set oQU = New a_QU
    oQU.Load lngQUID, True
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuPastelines"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuPastelines"
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    
    ofrm.component oQU.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oQU.QuoteID
    
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = "Note: " & ofrm.Memo
    oSM.SetMemo ofrm.Memo, oQU.QuoteID
    oQU.SetMemo ofrm.Memo
    
    Unload ofrm

    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuMemo"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.mnuMemo"
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.txtTPMemo_Change"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.txtTPMemo_Change", , EA_NORERAISE
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
    oSM.SetMemo txtTPMemo, oQU.QuoteID
    oQU.SetMemo txtTPMemo
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.txtTPMemo_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtForAttn_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.SetForAttnQU txtForAttn, oQU.QuoteID
    oQU.SetForAttn txtForAttn
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmQuotationPreview.txtForAttn_Validate(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmQuotationPreview.txtForAttn_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

