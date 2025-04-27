VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Object = "{BE9AD7B4-2F12-4067-96E1-FBB7FB01D8EA}#9.0#0"; "CoolButton.ocx"
Begin VB.Form frmAPPRPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Appro return preview"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11430
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "frmAPPRPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
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
      Left            =   2115
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   210
      Width           =   1545
   End
   Begin VB.TextBox txtDocCode 
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
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   240
      Width           =   1545
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
      Left            =   270
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   540
      Width           =   3390
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
      Left            =   3255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4695
      Visible         =   0   'False
      Width           =   2805
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
      Left            =   7740
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   30
      Width           =   1830
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
      Height          =   615
      Left            =   2220
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAPPRPreview.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close the appro return"
      Top             =   4725
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
      Height          =   315
      Left            =   9735
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   1455
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
      Left            =   1230
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAPPRPreview.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Print or preview"
      Top             =   4725
      Width           =   1000
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
      Left            =   195
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAPPRPreview.frx":085E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Print the invoice"
      Top             =   4725
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid G1 
      Height          =   3675
      Left            =   210
      OleObjectBlob   =   "frmAPPRPreview.frx":0BE8
      TabIndex        =   11
      Top             =   990
      Width           =   10725
   End
   Begin CoolButtonControl.CoolButton cbTP 
      Height          =   855
      Left            =   3780
      TabIndex        =   16
      Top             =   30
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   1508
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
      Left            =   375
      TabIndex        =   15
      Top             =   240
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   780
      Left            =   210
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   1155
      X2              =   2775
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
      TabIndex        =   9
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
      Height          =   285
      Left            =   4305
      TabIndex        =   8
      Top             =   45
      Width           =   3105
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
      Height          =   210
      Left            =   4335
      TabIndex        =   7
      Top             =   450
      Width           =   3105
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
      Left            =   6420
      TabIndex        =   3
      Top             =   4920
      Width           =   2580
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
      Left            =   9075
      TabIndex        =   2
      Top             =   4935
      Width           =   1845
   End
End
Attribute VB_Name = "frmAPPRPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cCN As c_CNs
Dim oAPPR As a_APPR
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
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True
      End Select
End Sub
Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuSaveLayout"
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (oAPPR.StatusF = "IN PROCESS" And oAPPR.IsNew = False)
    Forms(0).mnuCancel.Enabled = (oAPPR.StatusF = "ISSUED")
    Forms(0).mnuCancelLine.Enabled = (oAPPR.StatusF = "ISSUED" And oAPPR.IsNew = False)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.SetMenu"
End Sub


Public Sub component(PID As Long)
    On Error GoTo errHandler
Dim lngID As Long
    lngID = PID
    Set oAPPR = New a_APPR
    oAPPR.Load lngID, True
    Me.Caption = "Appro return from " & oAPPR.Customer.NameAndCode(40) & oAPPR.StaffNameB
    LoadControls
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.component(PID)", PID
End Sub
Public Sub ComponentObject(pAPP As a_APPR)
    On Error GoTo errHandler
    Set oAPPR = pAPP
    Me.Caption = "Appro return for " & oAPPR.Customer.NameAndCode(40) & oAPPR.StaffNameB
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.ComponentObject(pAPP)", pAPP
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
    
        With oAPPR
            Me.txtDate = .DocDateF
            If DateDiff("d", .DOCDate, .IssDate) > 1 Then
                Me.txtIssued = "Issued: " & .IssDateF
            Else
                txtIssued = ""
            End If
            Me.txtDocCode = .DOCCode
            Me.txtStatus = .StatusF
            CancelLine.Visible = (.Status = stCANCELLED Or .Status = stVOID)
            If .Status = stInProcess Then
                cmdEdit.Enabled = True
            Else
                cmdEdit.Enabled = False
            End If
            Me.txtName = .Customer.Name
            If Not .Customer.BillTOAddress Is Nothing Then
                Me.txtPhone = .Customer.BillTOAddress.PhoneandFax
                Me.txtTPMemo = IIf(Len(.Memo) > 0, "Note:  " & Trim$(.Memo), "")
            End If
            txtTPMemo.Visible = (txtTPMemo > "")
        End With
        LoadGrid
        
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.LoadControls"
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo errHandler
'Dim frm As frmPreview_
'    oAPPR.PrintInvoice_Display True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.cmdPreview_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cbTP_Click()
    On Error GoTo errHandler
Dim frm As frmCustomerPreview
    Set frm = New frmCustomerPreview
    frm.component oAPPR.Customer
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.cbTP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_APPR

'Dim oDOC As a_DocumentControl
'Dim qtyLinesToPrint As Integer
'
'
'    If PrintCommandButtonCTRLDown Then
'        PrintCommandButtonCTRLDown = False
'
'        Screen.MousePointer = vbHourglass
'        oAPPR.APPRLines.SortLines enSequence, True
'
'        Set oDOC = oPC.Configuration.DocumentControls.FindDC(oAPP.constDOCCODE)
'        If oDOC Is Nothing Then
'            qtyLinesToPrint = 1
'        Else
'            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(oAPPR.constDOCCODE).QtyCopies
'        End If
'
'       If oAPPR.ExportToXML(enView, , , qtyLinesToPrint) = False Then
'           Screen.MousePointer = vbDefault
'           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
'       End If
'       Screen.MousePointer = vbDefault
'    Else
'

    Set frm = New frmPrintingOptions_APPR
    frm.ComponentObject oAPPR
    frm.Show vbModal
    
EXIT_Handler:

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frm As frmAPPR
    WaitMsg "Loading . . .", True, Me
    Set frm = New frmAPPR
    blnEdit = True
    frm.component oAPPR
    frm.Show
    WaitMsg "", False, Me

EXIT_Handler:
    Unload Me
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Integer
    
    Set XA = New XArrayDB
    XA.Clear
    For i = 1 To G1.Columns.Count
        G1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), G1.Columns(i - 1).Width)
    Next
    G1.Columns(5).Width = 0
    XA.ReDim 1, oAPPR.APPRLines.Count, 1, 10
    For i = 1 To oAPPR.APPRLines.Count
        With oAPPR.APPRLines(i)
            XA(i, 7) = .Key
            XA(i, 8) = .code
            XA(i, 9) = .EAN
            XA(i, 10) = .Fulfilled
            XA(i, 1) = .CodeF
            XA(i, 2) = .Title
            XA(i, 3) = .Qty
            XA(i, 4) = .QtyIssued & "(" & .QtyReturned & ")"
            XA(i, 5) = .ApproCode & "(" & .ApproDateF & ")"
            If oAPPR.APPRLines(i).Note > "" Then
                XA(i, 6) = "Note:  " & oAPPR.APPRLines(i).Note
                G1.Columns(5).Width = 4000
            End If
        End With
    Next i
    
    G1.Array = XA
    G1.ReBind

    
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.LoadGrid"
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    G1.Width = NonNegative_Lng(Me.Width - (G1.Left + 400))
    lngDiff = G1.Height
    G1.Height = NonNegative_Lng(Me.Height - (G1.TOP + 1220))
    lngDiff = (G1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    lblTotalCaption.TOP = lblTotalCaption.TOP + lngDiff
    lblTotalValues.TOP = lblTotalValues.TOP + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub G1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    
    If XA(Bookmark, 10) = "CAN" Then
        RowStyle.BackColor = COLOR_CANCELLED
    ElseIf XA(Bookmark, 10) = "FUL" Then
        RowStyle.BackColor = COLOR_FULFILLED
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.G1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
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
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    G1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.G1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 5
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.GetRowType(ColIndex)", ColIndex
End Function

Private Sub G1_DblClick()
    On Error GoTo errHandler
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(G1.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    str = FNS(XA.Value(G1.Bookmark, 7))
    If str = "" Then Exit Sub
    Set oP = New a_Product
   ' oP.Load oAPP.ApproLines(str).pID, 0
    oP.Load oAPPR.APPRLines(val(str)).PID, 0
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
        LogSaveToFile "Access violation in frmAPPRPreview: G1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmAPPRPreview: G1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.G1_DblClick", , EA_NORERAISE
    HandleError
End Sub
'Private Sub lvwLines_DblClick()
'Dim frm As frmProductPrev
'Dim frmA As frmProductPrevAQ
'Dim oP As a_Product
'    Screen.MousePointer = vbHourglass
'    Set oP = New a_Product
'    oP.Load oAPPR.APPRLines.FindLineByID(val(Me.lvwLines.SelectedItem.Key)).pID, 0
'    If oPC.Configuration.AntiquarianYN Then
'        Set frmA = New frmProductPrevAQ
'        frmA.Component oP
'        frmA.Show
'    Else
'        Set frm = New frmProductPrev
'        frm.Component oP
'        frm.Show
'    End If
'    Screen.MousePointer = vbDefault
'End Sub

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
    ErrorIn "frmAPPRPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oAPPR = Nothing
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub lvwLines_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.lvwLines_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub G1_Click()
    On Error GoTo errHandler
Dim str As String

    If IsNull(G1.Bookmark) Then Exit Sub
    str = IIf(FNS(XA.Value(G1.Bookmark, 9)) > "", FNS(XA.Value(G1.Bookmark, 9)), FNS(XA.Value(G1.Bookmark, 8)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.G1_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub lvwLines_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.lvwLines_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub mnuFileExit_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuFileExit_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub SetLvw()
'Dim style As Long
'Dim hHeader As Long
'
'  'get the handle to the listview header
'   hHeader = SendMessage(lvwLines.hwnd, LVM_GETHEADER, 0, ByVal 0&)
'
'  'get the current style attributes for the header
'   style = GetWindowLong(hHeader, GWL_STYLE)
'
'  'modify the style by toggling the HDS_BUTTONS style
'   style = style Xor HDS_BUTTONS
'
'  'set the new style and redraw the listview
'   If style Then
'      Call SetWindowLong(hHeader, GWL_STYLE, style)
'      Call SetWindowPos(lvwLines.hwnd, Me.hwnd, 0, 0, 0, 0, SWP_FLAGS)
'   End If
'
'
'End Sub


Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelAPPR oAPPR
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuCancel"
End Sub

Private Sub mnuClose_Click()
    On Error GoTo errHandler
    cmdClose_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuEdit_Click()
    On Error GoTo errHandler
    cmdEdit_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPrint_Click()
    On Error GoTo errHandler
    cmdPrint_Click
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oAPPR.VoidDocument
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuVoid"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    oAPPR.Reload
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.RefreshData"
End Sub

Public Sub mnuCancelLine()
    On Error GoTo errHandler
Dim oP As a_Product
Dim str As String
    If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    str = FNS(XA.Value(G1.Bookmark, 7))
    oAPPR.APPRLines(str).CancelLine
    RefreshData
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuCancelLine"
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    ofrm.component oAPPR.Memo
    ofrm.Show vbModal
    oSM.SetMemo ofrm.Memo, oAPPR.TRID
    txtTPMemo.Visible = (ofrm.Memo > "")
    txtTPMemo = ofrm.Memo
    oAPPR.Memo = ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.mnuMemo"
End Sub

Private Sub G1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String
On Error Resume Next
    If LastRow = "" Then Exit Sub
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(G1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
    
On Error GoTo errHandler
    str = IIf(FNS(XA.Value(G1.Bookmark, 9)) > "", FNS(XA.Value(G1.Bookmark, 9)), FNS(XA.Value(G1.Bookmark, 8)))
    If str = "" Then Exit Sub
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.G1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub txtTPMemo_Change()
    On Error GoTo errHandler
    txtTPMemo = HandleTextWithBites(txtTPMemo)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.txtTPMemo_Change", , EA_NORERAISE
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
    ErrorIn "frmAPPRPreview.txtTPMemo_DblClick", , EA_NORERAISE
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
    ErrorIn "frmAPPRPreview.txtTPMemo_LostFocus", , EA_NORERAISE
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
    oSM.SetMemo txtTPMemo, oAPPR.TRID
    oAPPR.Memo = txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.txtTPMemo_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmAPPRPreview.txtTPMemo_DragOver(Source,x,Y,State)", Array(Source, x, Y, State), _
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
    oSM.SetMemo txtTPMemo, oAPPR.TRID
    oAPPR.Memo = txtTPMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAPPRPreview.txtTPMemo_DragDrop(Source,x,Y)", Array(Source, x, Y), EA_NORERAISE
    HandleError
End Sub



