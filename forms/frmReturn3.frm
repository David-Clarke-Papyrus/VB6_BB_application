VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmReturn3 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Products for return"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11130
   Begin VB.CommandButton cmdRefused 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print refused list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5250
      Width           =   2295
   End
   Begin VB.CommandButton cmdSaveLayout 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Save layout"
      Height          =   315
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5835
      Visible         =   0   'False
      Width           =   975
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
      Left            =   7845
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmReturn3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Print the invoice"
      Top             =   5280
      Width           =   1000
   End
   Begin VB.CommandButton cmdPSPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6795
      Picture         =   "frmReturn3.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5250
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
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   255
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
      Left            =   165
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   255
      Width           =   1545
   End
   Begin VB.CommandButton cmdFixMissingInvoiceRefs 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Fix missing invoice references"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6375
      Visible         =   0   'False
      Width           =   3735
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
      Left            =   7215
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   3750
   End
   Begin VB.CommandButton cmdApprovalRequest 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print approval request"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6435
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton cmdreprint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print return slip"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6600
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdPrintPickList 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print picking list"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5250
      Width           =   2295
   End
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Remove from stock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      Picture         =   "frmReturn3.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2100
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   4110
      Left            =   60
      OleObjectBlob   =   "frmReturn3.frx":0A9E
      TabIndex        =   0
      Top             =   1080
      Width           =   10920
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      ForeColor       =   &H8000000D&
      Height          =   885
      Left            =   5250
      TabIndex        =   13
      Top             =   6000
      Width           =   3720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   810
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   3495
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
      Left            =   435
      TabIndex        =   10
      Top             =   630
      Width           =   2970
   End
   Begin VB.Line CancelLine 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   870
      X2              =   2535
      Y1              =   15
      Y2              =   1080
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
      Left            =   135
      TabIndex        =   9
      Top             =   255
      Width           =   1365
   End
End
Attribute VB_Name = "frmReturn3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim cRL As c_RL
'Dim dR As d_R
Dim ar As a_R
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim strAutoorManType As String
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim rs As ADODB.Recordset
Dim lngBadRows As Long
Dim strType As String
Dim dteSince As Date
Dim lngRID As Long
Dim strSupplierName As String

Dim PrintCommandButtonCTRLDown As Boolean

Private Sub cmdPSPrint_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim ShiftTest As Integer
   PrintCommandButtonCTRLDown = False
   ShiftTest = Shift And 7
   Select Case ShiftTest
      Case 1 ' or vbShiftMask
      Case 2 ' or vbCtrlMask
         PrintCommandButtonCTRLDown = True
      End Select
End Sub

Private Sub cmdPSPrint_KeyUp(KeyCode As Integer, Shift As Integer)
        PrintCommandButtonCTRLDown = False
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (ar.Status <> stCOMPLETE)
    Forms(0).mnuCancel.Enabled = (ar.Status = stCOMPLETE)
    Forms(0).mnuCancelLine.Enabled = False '(ar.Status = stCOMPLETE)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyLines.Enabled = True
    Forms(0).mnuPastelines.Enabled = True
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    If (ar.Status = stISSUED Or ar.Status = stCOMPLETE) Then
        If Not ar.Supplier.OrderToAddress Is Nothing Then
            If (oPC.EDIEnabled And ar.Supplier.GFXNumber > "" And ar.Supplier.DispatchMethod = "E") Then
                Forms(0).mnuEmail.Enabled = False
                Forms(0).mnuOutlook.Enabled = False
                Forms(0).mnuEDI.Enabled = oPC.EDIEnabled
            Else
                If (oPC.EmailPO And ar.Supplier.DispatchMethod = "M" And ar.Supplier.OrderToAddress.EMail > "") Then
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
    ErrorIn "frmReturn3.SetMenu"
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.mnuSaveLayout"
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefused_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
Dim oReport As New arReturnRefusedList

    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "Requested", adInteger
    rs.fields.Append "Approved", adInteger
    rs.fields.Append "Returned", adInteger
    rs.fields.Append "Title", adVarChar, 50
    rs.fields.Append "Code", adVarChar, 25
    rs.fields.Append "Section", adVarChar, 25
    rs.fields.Append "Refs", adVarChar, 250
    rs.open
    For i = 1 To lngArrayRows
        If FNN(XA.Value(i, 4)) > FNN(XA.Value(i, 5)) Or FNN(XA.Value(i, 7)) > 0 Then
            rs.AddNew
                rs.fields("RLID") = FNN(XA.Value(i, 9))
                rs.fields("Requested") = FNN(XA.Value(i, 4))
                rs.fields("Approved") = FNN(XA.Value(i, 5))
                rs.fields("Returned") = FNN(XA.Value(i, 6))
                rs.fields("Title") = Left(XA.Value(i, 2), 50)
                rs.fields("Code") = Left(FNS(XA.Value(i, 1)), 25)
                rs.fields("Section") = Left(FNS(XA.Value(i, 11)), 25)
                rs.fields("Refs") = Left(FNS(XA.Value(i, 12)), 250)
            rs.Update
        End If
    Next
    If rs.RecordCount > 0 Then rs.MoveFirst
    rs.Sort = "Title"
        oReport.component rs, strSupplierName, ar.DOCCode
    oReport.Width = 13000
    oReport.Height = 6000
    oReport.documentName = "Test"
    oReport.Show


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdRefused_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    oSM.CancelR ar.TRID
    Screen.MousePointer = vbDefault
    MsgBox "Document cancelled", vbInformation, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.mnuCancel"
End Sub
Public Sub mnuVoid()
    On Error GoTo errHandler
    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    ar.VoidDocument
    MsgBox "Document voided", vbInformation, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.mnuVoid"
End Sub

Private Sub SetFormControls()
    On Error GoTo errHandler
    strSupplierName = ar.Supplier.NameAndCode(30)
    Me.cmdreprint.Enabled = (ar.StatusF = "STOCK RETURNED")
    Me.cmdApprovalRequest.Enabled = (ar.StatusF = "STOCK RETURNED" Or ar.StatusF = "AUTHORIZATION REQUESTED")
    Me.cmdIssue.Enabled = Not (ar.StatusF = "STOCK RETURNED" Or ar.StatusF = "VOID")
    Me.cmdApprovalRequest.Enabled = Not (ar.StatusF = "STOCK RETURNED")
    Me.Caption = ar.DOCCode & "  Return (preview) to " & strSupplierName & ar.StaffNameB & IIf(ar.ApprovalRef > "", " Authorization No. " & ar.ApprovalRef, "") & "  " & IIf(ar.ApprovalTermDateF > "", " Approval termination:  " & ar.ApprovalTermDateF, "")
    Me.cmdreprint.Enabled = (ar.Status > 3)
    Me.cmdPSPrint.Enabled = True 'ar.StatusF <> "VOID" And ar.StatusF <> "CANCELLED"
    Me.cmdPrintPickList.Enabled = ar.StatusF <> "VOID" And ar.StatusF <> "CANCELLED"
   ' Me.Caption = "Products for Return no. " & ar.DocCode & ar.StaffNameB & "     To:" & ar.TPNAME
    txtStatus = ar.StatusF
    Me.txtDate = ar.DocDateF
    Me.txtInvoiceNum = ar.DOCCode
    If ar.StatusF = "STOCK RETURNED" Then
        Grid1.Columns(4).AllowFocus = False
        Grid1.Columns(5).AllowFocus = False
        Grid1.Columns(6).AllowFocus = False
    ElseIf ar.StatusF = "AUTHORIZATION REQUESTED" Then
        Grid1.Columns(6).AllowFocus = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.SetFormControls"
End Sub
Public Sub component(pR As d_R)
    On Error GoTo errHandler
    Set ar = New a_R
    ar.Load pR.TRID
    lngRID = ar.TRID
    strAutoorManType = ar.RType
    SetFormControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.component(pR)", pR
End Sub
Public Sub Component2(pTRID As Long)
    On Error GoTo errHandler
    Set ar = New a_R
    ar.Load pTRID
    lngRID = ar.TRID
    strAutoorManType = ar.RType
    SetFormControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Component2(pTRID)", pTRID
End Sub

Private Sub Form_Initialize()
    PrintCommandButtonCTRLDown = False
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.Width = 11500
        Me.Height = 6500
        Me.Left = 100
        Me.TOP = 100
    End If
    LoadGrid
    lblNote.Caption = ar.Memo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Public Sub ManageReturnRejection()
    On Error GoTo errHandler
Dim frm As frmReturnRejection
Dim lngRL As Long
Dim lngQtyReturned As Long
Dim lngQtyRejected As Long
Dim strTitle As String
Dim strNote As String

    If FNN(Grid1.Bookmark) > 0 Then
        lngRL = FNN(XA.Value(Grid1.Bookmark, 9))
        strTitle = FNS(XA.Value(Grid1.Bookmark, 2))
        lngQtyReturned = FNN(XA.Value(Grid1.Bookmark, 6))
        lngQtyRejected = FNN(XA.Value(Grid1.Bookmark, 7))
        strNote = FNS(XA.Value(Grid1.Bookmark, 3))
        Set frm = New frmReturnRejection
        frm.component strTitle, lngRL, lngQtyRejected, lngQtyReturned, strNote
        frm.Show vbModal
        Set ar = Nothing
        Set ar = New a_R
        ar.Load lngRID
        SetFormControls
        LoadGrid
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.ManageReturnRejection"
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim dteTMP As Date
Dim i As Integer

    Set rs = New ADODB.Recordset
    lngArrayRows = ar.RLines.Count
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, lngArrayRows, 1, 18
    lngIndex = 1
    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", "frmReturn3", CStr(i), Grid1.Columns(i - 1).Width)
    Next
    Do While lngIndex <= ar.RLines.Count
            XA.Value(lngIndex, 1) = ar.RLines(lngIndex).CodeF
            If strAutoorManType = "E" Then
                XA.Value(lngIndex, 2) = ar.RLines(lngIndex).Title & "  (" & ar.RLines(lngIndex).SINVRef & ")"
            Else
                XA.Value(lngIndex, 2) = ar.RLines(lngIndex).Title
            End If
            XA.Value(lngIndex, 3) = ar.RLines(lngIndex).Note 'make it the same initially
            XA.Value(lngIndex, 4) = ar.RLines(lngIndex).QtyRequested 'make it the same initially
            XA.Value(lngIndex, 5) = ar.RLines(lngIndex).QtyApproved 'make it the same initially
            XA.Value(lngIndex, 6) = ar.RLines(lngIndex).QtyReturned
            XA.Value(lngIndex, 7) = ar.RLines(lngIndex).QtyRejected
            
            XA.Value(lngIndex, 8) = ar.RLines(lngIndex).PID
            XA.Value(lngIndex, 9) = ar.RLines(lngIndex).ID
            XA.Value(lngIndex, 10) = ar.RLines(lngIndex).Pubcode
            XA.Value(lngIndex, 11) = ar.RLines(lngIndex).Sections
            XA.Value(lngIndex, 12) = ar.RLines(lngIndex).SINVRef
            XA.Value(lngIndex, 13) = ar.RLines(lngIndex).PLessDiscExtF(ar.ISForeignCurrency)
            XA.Value(lngIndex, 14) = ar.RLines(lngIndex).DELLID
            XA.Value(lngIndex, 15) = ar.RLines(lngIndex).PriceF(ar.ISForeignCurrency)
            XA.Value(lngIndex, 16) = ar.RLines(lngIndex).DiscountF
            XA.Value(lngIndex, 17) = ar.RLines(lngIndex).EAN
            XA.Value(lngIndex, 18) = ar.RLines(lngIndex).MainAuthor
            lngIndex = lngIndex + 1
    Loop
    XA.QuickSort 1, lngArrayRows, 2, XORDER_ASCEND, XTYPE_STRING    'Not by sections here surely? 11, XORDER_ASCEND, XTYPE_STRING,
    Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("You want to close this form. Your changes are saved and will be available when next you open it and choose 'Use existing order slate'", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdApprovalRequest_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
Dim oReport As New arReturnsRequest

    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "DELLID", adInteger
    rs.fields.Append "PubCode", adVarWChar, 10
    rs.fields.Append "Requested", adInteger
    rs.fields.Append "Title", adVarWChar, 150
    rs.fields.Append "Code", adVarWChar, 15
    rs.fields.Append "SupplierInvoiceRef", adVarWChar, 200
    rs.open
    For i = 1 To lngArrayRows
        rs.AddNew
            rs.fields("RLID") = XA.Value(i, 9)
            rs.fields("PubCode") = XA.Value(i, 10)
            rs.fields("Requested") = XA.Value(i, 4)
            rs.fields("Title") = XA.Value(i, 2)
            rs.fields("Code") = XA.Value(i, 1)
            rs.fields("DELLID") = XA.Value(i, 14)
            rs.fields("SupplierInvoiceRef") = XA.Value(i, 12)
        rs.Update
    Next
    rs.Sort = "TITLE"
    If rs.RecordCount > 0 Then rs.MoveFirst
    
'    If Not ar Is Nothing Then
'        oReport.Component rs, strSupplierName, ar.DocCode
'    Else
        oReport.component rs, strSupplierName, ar.DOCCode
'    End If
    oReport.Show
    
   
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdApprovalRequest_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim frm As frmApproval
Dim mApprovalRef As String
Dim mApprovalTermDate As Date
Dim lngStockReturnedTotal As Long
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long

    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_RETFIN_SIGN, , "Confirm you have reviewed the data and wish to return stock.", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
            If MsgBox("Confirm you have reviewed the data and wish to return stock.", vbYesNo + vbQuestion, "Confirm") = vbNo Then
                Exit Sub
            End If
    End If
    If ar.RType = "P" Or ar.RType = "S" Then
        Set frm = New frmApproval
        frm.Show vbModal
        If frm.IsCancelled Then
            Unload frm
            Exit Sub
        Else
            mApprovalRef = frm.ApprovalRef
            mApprovalTermDate = frm.ApprovalDate
            ar.ApprovalRef = mApprovalRef
            Unload frm
        End If
    End If
     
    WaitMsg "Issuing return  . . .", True, Me
    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "PID", adGUID
    rs.fields.Append "Approved", adInteger
    rs.fields.Append "Counted", adInteger
    rs.fields.Append "Returned", adInteger
    rs.fields.Append "System", adInteger
    rs.fields.Append "DELLID", adInteger
    rs.fields.Append "Price", adVarWChar, 15
    rs.fields.Append "Discount", adVarWChar, 15
 '   rs.Fields("Discount").NumericScale = 2
    rs.open
    lngStockReturnedTotal = 0
    For i = 1 To lngArrayRows
 '       If XA.Value(i, 7) > 0 Then
            rs.AddNew
                rs.fields("RLID") = XA.Value(i, 9)
                rs.fields("PID") = FNS(XA.Value(i, 8))
                rs.fields("Approved") = XA.Value(i, 5)
              '  rs.Fields("Counted") = XA.Value(i, 3)
                rs.fields("Returned") = XA.Value(i, 6)
                lngStockReturnedTotal = lngStockReturnedTotal + CLng(rs.fields("Returned"))
                rs.fields("DELLID") = XA.Value(i, 14)
                rs.fields("Price") = XA.Value(i, 15)
                rs.fields("Discount") = XA.Value(i, 16)
            rs.Update
  '      End If
    Next
    If lngStockReturnedTotal <= 0 Then
        If MsgBox("There are no items to return." & vbCrLf & "Do you want to continue?", vbExclamation + vbYesNo, "Warning") = vbNo Then
            GoTo EXIT_Handler
        End If
    End If
    Set oSM = New z_StockManager
    
    If Not oSM.GenerateReturns_Step2(lngRID, rs, ar.StaffID, mApprovalRef, mApprovalTermDate) Then
        MsgBox "The update operation failed, contact your support person", vbExclamation, "Warning"
        GoTo EXIT_Handler
    End If
'    'reload to get the value using the actual qty returned
    Set ar = Nothing
    Set ar = New a_R
    ar.Load lngRID
    SetFormControls


EXIT_Handler:

    WaitMsg "", False, Me
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrintPickList_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
Dim oReport As New arReturnPickList

    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "Counted", adInteger
    rs.fields.Append "Approved", adInteger
    rs.fields.Append "Returned", adInteger
    rs.fields.Append "Systemcalculated", adInteger
    rs.fields.Append "Title", adVarChar, 50
    rs.fields.Append "Code", adVarChar, 20
    rs.fields.Append "Section", adVarChar, 25
    rs.fields.Append "MainAuthor", adVarChar, 25
    rs.open
    For i = 1 To lngArrayRows
        rs.AddNew
            rs.fields("RLID") = FNN(XA.Value(i, 9))
            rs.fields("Approved") = FNN(XA.Value(i, 5))
            rs.fields("Returned") = FNN(XA.Value(i, 6))
            rs.fields("Title") = Left(FNS(XA.Value(i, 2)), 50)
            rs.fields("Code") = FNS(XA.Value(i, 1))
            rs.fields("Section") = Left(FNS(XA.Value(i, 11)), 25)
            rs.fields("MainAuthor") = Left(FNS(XA.Value(i, 18)), 25)
        rs.Update
    Next
    If rs.RecordCount > 0 Then rs.MoveFirst
    rs.Sort = "Section,MainAuthor,Title"
'    If Not ar Is Nothing Then
        oReport.component rs, strSupplierName, ar.DOCCode
'    Else
'        oReport.Component rs, strSupplierName, dR.DocCode
'    End If
    oReport.Width = 13000
    oReport.Height = 6000
    oReport.documentName = "Test"
    oReport.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdPrintPickList_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub PrintReturnSlip()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
Dim oReport As New arReturnWithStock

    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "PubCode", adVarWChar, 10
    rs.fields.Append "ReTurned", adInteger
    rs.fields.Append "Value", adVarWChar, 15
    rs.fields.Append "Title", adVarWChar, 150
    rs.fields.Append "Code", adVarWChar, 15
    rs.fields.Append "SupplierInvoiceRef", adVarWChar, 100
    rs.fields.Append "Price", adVarWChar, 10
    rs.fields.Append "Discount", adVarWChar, 10
    rs.open
    For i = 1 To lngArrayRows
        If FNN(XA.Value(i, 7)) > 0 Then
            rs.AddNew
                rs.fields("RLID") = XA.Value(i, 9)
                rs.fields("PubCode") = XA.Value(i, 10)
                rs.fields("Returned") = XA.Value(i, 6)
                rs.fields("Value") = XA.Value(i, 13)
                rs.fields("Title") = Left(XA.Value(i, 2), rs.fields("Title").DefinedSize)
                rs.fields("Code") = XA.Value(i, 1)
                rs.fields("SupplierInvoiceRef") = XA.Value(i, 12)
                rs.fields("Discount") = XA.Value(i, 16)
                rs.fields("Price") = XA.Value(i, 15)
            rs.Update
        End If
    Next
    If rs.RecordCount > 0 Then rs.MoveFirst
    rs.Sort = "Title"
'    If Not ar Is Nothing Then
        oReport.component rs, strSupplierName, ar.DOCCode, ar.ApprovalRef, ar.TotalPayableF(False)
'    Else
'        oReport.Component rs, strSupplierName, dR.DocCode, dR.ApprovalRef
'    End If
    oReport.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.PrintReturnSlip"
End Sub
Private Sub cmdreprint_Click()
    On Error GoTo errHandler
    PrintReturnSlip
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdreprint_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    'rs.Close
    Set rs = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub




Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product
    Screen.MousePointer = vbHourglass
    strPID = XA.Value(Grid1.Bookmark, 8)
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
        LogSaveToFile "Access violation in frmReturn3: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmReturn3: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If col = 3 Or col = 5 Or col = 6 Then
        CellStyle.BackColor = vbWhite
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", _
         Array(Condition, Split, Bookmark, col, CellStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim lngSupplierID As Long
Dim lngDEALID As Long

'    If XA.Value(Bookmark, 4) <> XA.Value(Bookmark, 5) Then
'        RowStyle.BackColor = vbCyan
'    ElseIf XA.Value(Bookmark, 5) <> XA.Value(Bookmark, 6) Then
'        RowStyle.BackColor = &HC0FFC0 'greenish
'    Else
'        RowStyle.BackColor = &HDBFAFB 'Pale yellow
'    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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

    If Not ConvertToLng(Grid1.text, lngTmp) And Grid1.col <> 2 Then
        Cancel = True
        Exit Sub
    End If
    If ColIndex <> 2 Then
        XA.Value(Grid1.Bookmark, ColIndex + 1) = FNN(Grid1.text)
        If CLng(XA.Value(Grid1.Bookmark, 6)) > CLng(XA.Value(Grid1.Bookmark, 5)) Then
            'Grid1.Text = OldValue
            Cancel = True
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If ColIndex = 4 Then  'Approved column
        oSM.UpdateReturnLineApproved FNN(XA.Value(Grid1.Bookmark, 9)), FNN(XA.Value(Grid1.Bookmark, 5))
    ElseIf ColIndex = 5 Then  'Returning column
        oSM.UpdateReturnLineReturning FNN(XA.Value(Grid1.Bookmark, 9)), FNN(XA.Value(Grid1.Bookmark, 6))
    ElseIf ColIndex = 2 Then  'Notes column
        oSM.UpdateReturnLineNote FNN(XA.Value(Grid1.Bookmark, 9)), FNS(Grid1.text)
        ar.RLines(Grid1.Bookmark).Note = FNS(Grid1.text)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2
            GetRowType = XTYPE_STRING
        Case Else
            GetRowType = XTYPE_NUMBER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.GetRowType(ColIndex)", ColIndex
End Function


Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
'    Grid1.ReBind
    Grid1.Update
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub ValidateRow()
    On Error GoTo errHandler
 '   If CLng(XA.Value(Grid1.Bookmark, 6)) > CLng(XA.Value(Grid1.Bookmark, 3)) Then
 '       XA.Value(Grid1.Bookmark, 7) = XA.Value(Grid1.Bookmark, 4)
 '   End If
    
    If CLng(XA.Value(Grid1.Bookmark, 6)) > CLng(XA.Value(Grid1.Bookmark, 5)) Then
        XA.Value(Grid1.Bookmark, 6) = GetMin(XA.Value(Grid1.Bookmark, 5), XA.Value(Grid1.Bookmark, 4))
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.ValidateRow"
End Sub
Private Sub Grid1_BeforeUpdate(Cancel As Integer)
    On Error GoTo errHandler
    ValidateRow
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_BeforeUpdate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_AfterUpdate()
    On Error GoTo errHandler
'    ValidateRow
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_AfterUpdate", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPSPrint_Click()
    On Error GoTo errHandler
Dim frm As frmPrintingOptions_R
'
Dim oDOC As a_DocumentControl
Dim qtyLinesToPrint As Integer
Dim Dummy As String

    If PrintCommandButtonCTRLDown Then
        PrintCommandButtonCTRLDown = False

        Screen.MousePointer = vbHourglass
        ar.RLines.SortLines enSequence, True

        Set oDOC = oPC.Configuration.DocumentControls.FindDC(ar.constDOCCODE)
        If oDOC Is Nothing Then
            qtyLinesToPrint = 1
        Else
            qtyLinesToPrint = oPC.Configuration.DocumentControls.FindDC(ar.constDOCCODE).QtyCopies
        End If

       If ar.ExportToXML(ar.ISForeignCurrency, Dummy, True, enView, qtyLinesToPrint, , , , True) = False Then
           Screen.MousePointer = vbDefault
           MsgBox "Printing has failed, probably the document type has not been configured for this workstation.", vbInformation + vbOKOnly, "Can't do this"
       End If
       Screen.MousePointer = vbDefault
    Else
        Screen.MousePointer = vbHourglass
        Set frm = New frmPrintingOptions_R
        frm.ComponentObject ar
        Screen.MousePointer = vbDefault
        frm.Show vbModal
    End If
    
EXIT_Handler:
 '   Unload Me
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdPSPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
    If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuReturnPopup   ' Display the File menu as a
                        ' pop-up menu.
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturns3.Grid1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, x, Y), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSaveLayout_Click()
    On Error GoTo errHandler
    SaveLayout Me.Grid1, Me.Name
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPOPreview.cmdSaveLayout_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuMemo()
    On Error GoTo errHandler
Dim ofrm As New frmNote
Dim oSM As New z_StockManager
    ofrm.component ar.Memo
    ofrm.Show vbModal
    oSM.UpdateReturnMemo ar.TRID, ofrm.Memo
    lblNote.Caption = ofrm.Memo
    ar.Memo = ofrm.Memo
    Unload ofrm
    Set ofrm = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPO.mnuMemo"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.mnuMemo"
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String

On Error Resume Next
    If LastRow = "" Then Exit Sub
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
On Error GoTo errHandler
    str = IIf(FNS(XA.Value(Grid1.Bookmark, 17)) > "", FNS(XA.Value(Grid1.Bookmark, 17)), FNS(XA.Value(Grid1.Bookmark, 1)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_Click()
Dim str As String
On Error Resume Next
    If XA.UpperBound(1) = 0 Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
On Error GoTo errHandler
    str = IIf(FNS(XA.Value(Grid1.Bookmark, 17)) > "", FNS(XA.Value(Grid1.Bookmark, 17)), FNS(XA.Value(Grid1.Bookmark, 1)))
    If str = "" Then Exit Sub
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturnPreview.Grid1_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.Grid1_Click", , EA_NORERAISE
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
    For Each oLine In ar.RLines
 '       If Not oLine.Product.IsServiceItem Then
        rs.AddNew
        rs.fields("GUID") = CreateGUID
        rs.fields("PID") = oLine.PID
        If ar.Status = stISSUED Then
            rs.fields("Qty") = oLine.QtyRequested
            rs.fields("QtyFirm") = oLine.QtyRequested
            rs.fields("QtySS") = 0
        ElseIf ar.Status = stCOMPLETE Then
            rs.fields("Qty") = oLine.QtyReturned
        End If
'        If ooR.ISForeignCurrency Then
'            rs.Fields("Price") = oLine.ForeignPrice
'        Else
            rs.fields("Price") = oLine.Price(ar.ISForeignCurrency)
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
    ErrorIn "frmReturn3.mnuCopyLines"
End Sub

Public Sub mnuPastelines()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oLine As a_RL
Dim s As String

    Set rs = oPC.LinesClipboard
    If rs.State = 0 Then Exit Sub
    If MsgBox("Confirm you are adding " & CStr(rs.RecordCount) & " lines to document " & ar.DOCCode, vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
   ' rs.Open
    If rs.BOF And rs.eof Then Exit Sub
    rs.MoveFirst
    Do While Not rs.eof
        Set oLine = ar.RLines.Add
        oLine.BeginEdit
        oLine.PID = rs.fields("PID")
        oLine.SINVRef = FNS(rs.fields("REF"))
        If ar.Status = stISSUED Then
            oLine.QtyRequested = FNDBL(rs.fields("Qty"))
            oLine.QtyReturned = 0
        ElseIf ar.Status = stCOMPLETE Then
            oLine.QtyReturned = FNDBL(rs.fields("Qty"))
        End If
        oLine.Price(ar.ISForeignCurrency) = FNDBL(rs.fields("Price"))
        oLine.Discount = FNDBL(rs.fields("DISCOUNTRATE"))
        oLine.CodeF = FNS(rs.fields("CODEF"))
        oLine.EAN = FNS(rs.fields("EAN"))
        oLine.Title = FNS(rs.fields("TITLE"))
        oLine.ApplyEdit
        rs.MoveNext
    Loop
    rs.Close
    ar.ApplyEdit
    ar.BeginEdit
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn3.mnuPastelines"
End Sub

