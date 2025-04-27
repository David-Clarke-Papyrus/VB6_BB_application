VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmReturnPreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Products for return"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   11130
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
      Left            =   180
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmReturnPreview.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Print the invoice"
      Top             =   4650
      Width           =   1000
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
      Left            =   1215
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmReturnPreview.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print or preview"
      Top             =   4635
      Width           =   1000
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
      Height          =   615
      Left            =   2235
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmReturnPreview.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Close the return"
      Top             =   4650
      Width           =   1000
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
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   5
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
      Height          =   345
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   1545
   End
   Begin VB.CommandButton cmdReverse 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Set status back to 'Requested'"
      Height          =   285
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5745
      Visible         =   0   'False
      Width           =   2295
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
      Height          =   1245
      Left            =   3300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   2685
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
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   4065
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3420
      Left            =   180
      OleObjectBlob   =   "frmReturnPreview.frx":0A9E
      TabIndex        =   0
      Top             =   1155
      Width           =   10860
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   810
      Left            =   75
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
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1365
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   855
      X2              =   2520
      Y1              =   0
      Y2              =   1065
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
      Left            =   420
      TabIndex        =   6
      Top             =   615
      Width           =   2970
   End
End
Attribute VB_Name = "frmReturnPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cRL As c_RL
Dim ooR As a_R
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim rs As ADODB.Recordset
Dim lngBaooRows As Long
Dim strType As String
Dim dteSince As Date
Dim lngRID As Long
Dim strSupplierName As String

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

Private Sub cmdPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.mnuSaveLayout"
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.mnuSaveLayout"
End Sub

Private Sub cmdReverse_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.ReverseReturnToRequested ooR.TRID
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.cmdReverse_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.cmdReverse_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Form_Activate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Form_Deactivate"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (ooR.Status <> stCOMPLETE)
    Forms(0).mnuCancel.Enabled = (ooR.Status = stCOMPLETE)
    Forms(0).mnuCancelLine.Enabled = (ooR.Status = stCOMPLETE)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    If (ooR.Status = stISSUED Or ooR.Status = stCOMPLETE) Then
        If Not ooR.Supplier.OrderToAddress Is Nothing Then
            If (oPC.EDIEnabled And ooR.Supplier.GFXNumber > "" And ooR.Supplier.DispatchMethod = "E") Then
                Forms(0).mnuEmail.Enabled = False
                Forms(0).mnuOutlook.Enabled = False
                Forms(0).mnuEDI.Enabled = oPC.EDIEnabled
            Else
                If (oPC.EmailPO And ooR.Supplier.DispatchMethod = "M" And ooR.Supplier.OrderToAddress.EMail > "") Then
                    Forms(0).mnuEmail.Enabled = Not oPC.UsesOutlookForPOEmail
                    Forms(0).mnuOutlook.Enabled = oPC.UsesOutlookForPOEmail
                    Forms(0).mnuEDI.Enabled = False
                Else
                    Forms(0).mnuEmail.Enabled = False
                    Forms(0).mnuOutlook.Enabled = False
                    Forms(0).mnuEDI.Enabled = False
                End If
            End If
        Else
                Forms(0).mnuEmail.Enabled = False
                Forms(0).mnuOutlook.Enabled = False
        End If
    Else
        Forms(0).mnuEmail.Enabled = False
        Forms(0).mnuOutlook.Enabled = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.SetMenu"
End Sub
Public Sub mnuEmail()
    On Error GoTo errHandler
Dim Res As Boolean
Dim lOR As a_R
Dim strFilename As String
Dim strDestinationEmail As String
Dim strWholeMessage As String
Dim strReference As String

    If ooR.Supplier.DispatchMethod = "M" Then
        Screen.MousePointer = vbHourglass
        Set lOR = New a_PO
        lOR.Load ooR.TRID
        Res = lOR.ExportToXML(ooR.ISForeignCurrency, strFilename, False, enMail, 0, strDestinationEmail, strWholeMessage)
        Screen.MousePointer = vbDefault
    ElseIf ooR.Supplier.DispatchMethod = "E" Then
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.mnuEmail"
End Sub
Public Sub ComponentObject(pOOR As a_R)
    On Error GoTo errHandler
    Set ooR = pOOR
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.ComponentObject(pOOR)", pOOR
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.ComponentObject(pOOR)", pOOR
End Sub
Public Sub component(pRID As Long)
    On Error GoTo errHandler

    lngRID = pRID
    Set ooR = New a_R
    ooR.Load lngRID
    Me.Caption = "Return (preview) to " & ooR.TPNAME & ooR.StaffNameB
    strSupplierName = ooR.TPNAME
    SetMenu
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Component(pRID)", pRID
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.component(pRID)", pRID
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim dODPO As d_POLine
Dim dteTMP As Date

    Set rs = New ADODB.Recordset
    lngArrayRows = ooR.RLines.Count
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, lngArrayRows, 1, 13
    Grid1.Columns(5).Width = 0
    lngIndex = 1
    Do While lngIndex <= ooR.RLines.Count
            XA.Value(lngIndex, 1) = ooR.RLines(lngIndex).CodeF
            XA.Value(lngIndex, 2) = ooR.RLines(lngIndex).Title
            XA.Value(lngIndex, 3) = ooR.RLines(lngIndex).QtyRequested & "," & ooR.RLines(lngIndex).QtyApproved & "," & ooR.RLines(lngIndex).QtyReturned
            XA.Value(lngIndex, 4) = ooR.RLines(lngIndex).DiscountF
            XA.Value(lngIndex, 5) = ooR.RLines(lngIndex).SINVRef ' & "  " & cRL(lngIndex).SupplierInvoiceDate
            XA.Value(lngIndex, 7) = ooR.RLines(lngIndex).PID
            XA.Value(lngIndex, 8) = ooR.RLines(lngIndex).Status
            XA.Value(lngIndex, 9) = ooR.RLines(lngIndex).ID
            If ooR.RLines(lngIndex).Note > "" Then
                XA(lngIndex, 6) = "Note:  " & ooR.RLines(lngIndex).Note
                Grid1.Columns(5).Width = 4000
            End If
            XA.Value(lngIndex, 13) = ooR.RLines(lngIndex).EAN
            lngIndex = lngIndex + 1
    Loop
    XA.QuickSort 1, lngArrayRows, 11, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("You want to close this form. Your changes are saved and will be available when next you open it and choose 'Use existing order slate'", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.cmdCancel_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.cmdClose_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim blnEdit As Boolean
Dim frm As frmReturn
Dim bCancel As Boolean
    Set frm = New frmReturn
    blnEdit = True
    frm.Show
    frm.component bCancel, ooR

EXIT_Handler:
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.cmdEdit_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_R
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        ooR.RLines.SortLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(ooR.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(ooR.constDOCCODE).QtyCopies
        End If

       If ooR.ExportToXML(ooR.ISForeignCurrency, Dummy, True, enView, qtyLinesToPrint) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbHourglass
        Set frm = New frmPrintingOptions_R
        frm.ComponentObject ooR
        Screen.MousePointer = vbDefault
        frm.Show vbModal
    End If
EXIT_Handler:
 '   Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.cmdPrint_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    Grid1.Width = NonNegative_Lng(Me.Width - (Grid1.Left + 400))
    lngDiff = Grid1.Height
    Grid1.Height = NonNegative_Lng(Me.Height - (Grid1.TOP + 1800))
    lngDiff = (Grid1.Height - lngDiff)
    cmdEdit.TOP = cmdEdit.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    txtTPMemo.TOP = txtTPMemo.TOP + lngDiff
    cmdReverse.TOP = cmdReverse.TOP + lngDiff

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    'rs.Close
    Set rs = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Form_Unload(Cancel)", Cancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product
    Screen.MousePointer = vbHourglass
    strPID = XA.Value(Grid1.Bookmark, 7)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.component oProd
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmReturnPreview: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmReturnPreview: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
   If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If col = 3 Or col = 4 Then
        CellStyle.BackColor = vbWhite
    End If
 '  If Col = 3 And (XA.Value(Bookmark, 3) <> Col) Then
 '       CellStyle.BackColor = vbCyan
 '   End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Grid1_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", _
'         Array(Condition, Split, Bookmark, Col, CellStyle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", _
         Array(Condition, Split, Bookmark, col, CellStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim lngSupplierID As Long
Dim lngDEALID As Long
    If XA(Bookmark, 8) = "CAN" Then
        RowStyle.BackColor = &HC0C0C0
        RowStyle.Font.Strikethrough = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim strTmp As String
Dim bTmp As Boolean
Dim f1 As String
Dim f2 As String
Dim f3 As String
Dim lngTmp As Long

    If Not ConvertToLng(Grid1.text, lngTmp) Then
        Cancel = True
        Exit Sub
    End If
    XA.Value(Grid1.Bookmark, ColIndex + 1) = FNN(Grid1.text)
    If (XA.Value(Grid1.Bookmark, 5) < 0) Or (XA.Value(Grid1.Bookmark, 5) > XA.Value(Grid1.Bookmark, 4)) Then
        Grid1.text = OldValue
        Cancel = True
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
'         OldValue, Cancel)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
    
    Grid1.Refresh
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Grid1_HeadClick(ColIndex)", ColIndex
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 3, 4
            GetRowType = XTYPE_NUMBER
        Case Else
            GetRowType = XTYPE_STRING
    End Select
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.GetRowType(ColIndex)", ColIndex
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.Width = 11500
        Me.Height = 6500
        Me.Left = 100
        Me.TOP = 100
    End If
    Me.Grid1.TOP = 1000
    Me.Grid1.Height = 3400
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Form_Load"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
    Next
    LoadGrid
    Me.Caption = "Return (preview) no. " & ooR.DOCCode & ooR.StaffNameB & "      To:" & ooR.TPNAME
        If ooR.Status = stInProcess Then
            cmdEdit.Enabled = True
        Else
            cmdEdit.Enabled = False
        End If
        Me.txtDate = ooR.DocDateF
        If DateDiff("d", ooR.DOCDate, ooR.IssDate) > 1 Then
            lblSI.Caption = "Issued: " & ooR.IssDateF
        Else
            lblSI.Caption = ""
        End If
        Me.txtInvoiceNum = ooR.DOCCode

    txtStatus = ooR.StatusF
  '  lblSupplier.Caption = ooR.TPName & "  " & ooR.DocCode

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.LoadControls"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.LoadControls"
End Sub

Private Sub ValidateRow(pOKAtPresent As Boolean)
    On Error GoTo errHandler
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.ValidateRow(pOKAtPresent)", pOKAtPresent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.ValidateRow(pOKAtPresent)", pOKAtPresent
End Sub
Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    ooR.VoidDocument
    RefreshData
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.mnuVoid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.mnuVoid"
End Sub
Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelR ooR.TRID
    RefreshData
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.mnuCancel"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.mnuCancel"
End Sub
Public Sub mnuCancelLine()
    On Error GoTo errHandler
Dim oP As a_Product
    If MsgBox("Do you wish to cancel the selected line?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    ooR.RLines.FindLineByID(val(XA(Grid1.Bookmark, 9))).CancelLine
    RefreshData
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.mnuCancelLine"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.mnuCancelLine"
End Sub
Public Sub RefreshData()
    On Error GoTo errHandler
    ooR.Reload
    LoadControls
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.RefreshData"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.RefreshData"
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String
On Error Resume Next
    If LastRow = "" Then Exit Sub
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
On Error GoTo errHandler
    str = IIf(FNS(XA.Value(Grid1.Bookmark, 13)) > "", FNS(XA.Value(Grid1.Bookmark, 13)), FNS(XA.Value(Grid1.Bookmark, 1)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_Click()
Dim str As String

On Error Resume Next
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
On Error GoTo errHandler

    str = IIf(FNS(XA.Value(Grid1.Bookmark, 13)) > "", FNS(XA.Value(Grid1.Bookmark, 13)), FNS(XA.Value(Grid1.Bookmark, 12)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturnPreview.Grid1_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuCopyLines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_RL
Dim fs As New FileSystemObject

    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
    rs.open
    For Each oLine In ooR.RLines
 '       If Not oLine.Product.IsServiceItem Then
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        If ooR.Status = stISSUED Then
            rs.fields("Qty") = oLine.QtyRequested
        ElseIf ooR.Status = stCOMPLETE Then
            rs.fields("Qty") = oLine.QtyReturned
        End If
'        If ooR.ISForeignCurrency Then
'            rs.Fields("Price") = oLine.ForeignPrice
'        Else
            rs.fields("Price") = oLine.Price(ooR.ISForeignCurrency)
'        End If
        rs.fields("DISCOUNTRATE") = oLine.Discount
        rs.fields("CODEF") = oLine.CodeF
        rs.fields("EANF") = oLine.EAN
        rs.fields("EAN") = oLine.EAN
        rs.fields("TITLE") = oLine.Title
        rs.fields("VATRATE") = oPC.Configuration.VATRate
        rs.fields("REF") = oLine.SINVRef
        rs.fields("ETA") = CDate(0)
        rs.Update
  '      End If
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
    ErrorIn "frmInvoicePreview.mnuCopyLines"
End Sub

