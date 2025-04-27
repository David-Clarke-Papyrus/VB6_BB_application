VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmReturn2 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Products for return"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   11490
   Begin VB.CommandButton cmdIssue 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issue return request"
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
      Left            =   8865
      Picture         =   "frmReturn2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5190
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Print list"
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
      Left            =   180
      Picture         =   "frmReturn2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5205
      Width           =   1320
   End
   Begin VB.CommandButton cmdreprint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Reprint approval request"
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
      Height          =   615
      Left            =   1605
      Picture         =   "frmReturn2.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5220
      Width           =   2295
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
      Left            =   7815
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmReturn2.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Print the invoice"
      Top             =   5190
      Width           =   1000
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
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6270
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
      Left            =   6765
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   4530
      Left            =   180
      OleObjectBlob   =   "frmReturn2.frx":0E28
      TabIndex        =   0
      Top             =   600
      Width           =   10860
   End
   Begin VB.Label lblSupplier 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   195
      TabIndex        =   2
      Top             =   105
      Width           =   5265
   End
End
Attribute VB_Name = "frmReturn2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cRL As c_RL
Attribute cRL.VB_VarHelpID = -1
Dim dR As d_R
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim rs As ADODB.Recordset
Dim lngBadRows As Long
Dim strType As String
Dim dteSince As Date
Dim lngRID As Long
Dim strSupplierName As String
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
    ErrorIn "frmReturn2.mnuSaveLayout"
End Sub

Public Sub component(pR As d_R, pSupplierName As String)
    On Error GoTo errHandler

    lngRID = pR.TRID
    Set dR = pR
    Me.cmdIssue.Enabled = (dR.StatusF = "IN PROCESS")
    Me.cmdreprint.Enabled = (dR.StatusF = "ISSUED")
    Me.Caption = dR.DOCCode & "  Return to " & pSupplierName & dR.StaffNameB & IIf(dR.ApprovalRef > "", "      Authorization No. " & dR.ApprovalRef, "")
    strSupplierName = pSupplierName
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.Component(pR,pSUpplierName)", Array(pR, pSUpplierName)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.component(pR,pSUpplierName)", Array(pR, pSupplierName)
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (dR.Status <> stCOMPLETE)
    Forms(0).mnuCancel.Enabled = (dR.Status = stCOMPLETE)
    Forms(0).mnuCancelLine.Enabled = False '(ar.Status = stCOMPLETE)
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = True
    Forms(0).mnuSalesComm.Enabled = False
    'Forms(0).mnuInvAdd.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = True
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.SetMenu"
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub


Public Sub mnuCancel()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    If MsgBox("Do you want to cancel this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    oSM.CancelR dR.TRID
    Screen.MousePointer = vbDefault
    MsgBox "Document cancelled", vbInformation, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.mnuCancel"
End Sub
Public Sub mnuVoid()
    On Error GoTo errHandler
Dim OpenResult As Integer

    If MsgBox("Do you want to void this document?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.execute "UPDATE tTR SET TR_STATUS = 1 WHERE TR_ID = " & dR.TRID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    MsgBox "Document voided", vbInformation, "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.mnuVoid"
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
Dim i As Integer
    
    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", "frmReturn2", CStr(i), Grid1.Columns(i - 1).Width)
    Next
    Set rs = New ADODB.Recordset
    lngArrayRows = cRL.Count
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, lngArrayRows, 1, 15
    lngIndex = 1
    Do While lngIndex <= cRL.Count
            XA.Value(lngIndex, 1) = cRL(lngIndex).code
            If dR.RType = "E" Then
                XA.Value(lngIndex, 2) = cRL(lngIndex).ProductDescription & cRL(lngIndex).code
            Else
                XA.Value(lngIndex, 2) = cRL(lngIndex).ProductDescription & cRL(lngIndex).SupplierInvoiceRef
            End If
            XA.Value(lngIndex, 3) = cRL(lngIndex).QtySystem
            XA.Value(lngIndex, 4) = cRL(lngIndex).QtySystem 'make it the same initially
            XA.Value(lngIndex, 5) = cRL(lngIndex).QtyRequested 'make it the same initially
            XA.Value(lngIndex, 6) = cRL(lngIndex).QtyApproved
            XA.Value(lngIndex, 7) = cRL(lngIndex).QtyReturned
            XA.Value(lngIndex, 8) = cRL(lngIndex).PID
            XA.Value(lngIndex, 9) = cRL(lngIndex).RLID
            XA.Value(lngIndex, 10) = cRL(lngIndex).Pubcode
            XA.Value(lngIndex, 11) = cRL(lngIndex).Section
            XA.Value(lngIndex, 12) = cRL(lngIndex).SupplierInvoiceRef
            XA.Value(lngIndex, 13) = cRL(lngIndex).SupplierInvoiceDate
            XA.Value(lngIndex, 14) = cRL(lngIndex).ProductDescription
         '   XA.Value(lngIndex, 15) = cRL(lngIndex).SupplierInvoiceDate
    '        Debug.Print cRL(lngIndex).SupplierInvoiceRef
            lngIndex = lngIndex + 1
           ' rs.MoveNext
    Loop
    XA.QuickSort 1, lngArrayRows, 11, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("You want to close this form. Your changes are saved and will be available when next you open it and choose 'Use existing order slate'", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdIssue_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
    
    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_RETREQ_SIGN, , "Confirm you have reviewed the data and now wish to create a request.", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If MsgBox("Confirm you have reviewed the data and now wish to create a request.", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
            Exit Sub
        End If
    End If
    
    WaitMsg "Issuing return  . . .", True, Me
    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "PID", adGUID
    rs.fields.Append "System", adInteger
    rs.fields.Append "Counted", adInteger
    rs.fields.Append "Requested", adInteger
    rs.fields.Append "SupplierInvoiceRef", adVarChar, 100
    rs.fields.Append "SupplierInvoiceDate", adDate
    rs.open
    For i = 1 To lngArrayRows
        rs.AddNew
            rs.fields("RLID") = XA.Value(i, 9)
            rs.fields("PID") = FNS(XA.Value(i, 8))
            rs.fields("System") = XA.Value(i, 3)
            rs.fields("Counted") = XA.Value(i, 4)
            rs.fields("Requested") = FNS(XA.Value(i, 5))
            rs.fields("SupplierInvoiceRef") = FNS(XA.Value(i, 12))
            rs.fields("SupplierInvoiceDate") = XA.Value(i, 13)
        rs.Update
    Next
    Set oSM = New z_StockManager
    If Not oSM.GenerateReturns_Step1(lngRID, rs, gSTAFFID) Then
        MsgBox "The update operation failed, contact your support person"
        Exit Sub
    End If
    PrintRequest
    cmdIssue.Enabled = False
    WaitMsg "", False, Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.cmdIssue_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.cmdIssue_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
Dim oReport As New arReturns1a
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "Counted", adInteger
    rs.fields.Append "Approved", adInteger
    rs.fields.Append "Systemcalculated", adInteger
    rs.fields.Append "Title", adVarChar, 50
    rs.fields.Append "Code", adVarChar, 20
    rs.fields.Append "Section", adVarChar, 25
    rs.fields.Append "SuppInv", adVarChar, 100
    rs.fields.Append "SuppInvDate", adVarChar, 100
    rs.open
    For i = 1 To lngArrayRows
        rs.AddNew
            rs.fields("RLID") = XA.Value(i, 9)
            rs.fields("Counted") = XA.Value(i, 4)
            rs.fields("Approved") = XA.Value(i, 7)
            rs.fields("Systemcalculated") = XA.Value(i, 3)
            rs.fields("Title") = Left(XA.Value(i, 14), 50)
            rs.fields("Code") = FNS(XA.Value(i, 1))
            rs.fields("Section") = FNS(XA.Value(i, 11))
            rs.fields("SuppInv") = FNS(XA.Value(i, 12))
            rs.fields("SuppInvDate") = FND(XA.Value(i, 13))
        rs.Update
    Next
    If rs.RecordCount > 0 Then rs.MoveFirst
    rs.Sort = "Title"
    oReport.component rs, strSupplierName, dR.DOCCode
    oReport.Width = 13000
    oReport.Height = 6000
    oReport.documentName = "Test"
    oReport.Show
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.cmdPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub PrintRequest()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oSM As z_StockManager
Dim i As Long
Dim oReport As New arReturnsRequest
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    rs.fields.Append "RLID", adInteger
    rs.fields.Append "PubCode", adVarWChar, 10
    rs.fields.Append "Requested", adInteger
    rs.fields.Append "Title", adVarWChar, 50
    rs.fields.Append "Code", adVarWChar, 20
    rs.fields.Append "SupplierInvoiceRef", adVarWChar, 100
    rs.open
    For i = 1 To lngArrayRows
        If XA.Value(i, 5) <> "0" Then
            rs.AddNew
                rs.fields("RLID") = XA.Value(i, 9)
                rs.fields("PubCode") = XA.Value(i, 10)
                rs.fields("Requested") = XA.Value(i, 5)
                rs.fields("Title") = Left(XA.Value(i, 2), rs.fields("Title").DefinedSize)
                rs.fields("Code") = FNS(XA.Value(i, 1))
                rs.fields("SupplierInvoiceRef") = FNS(XA.Value(i, 12))
            rs.Update
        End If
    Next
    If rs.RecordCount > 0 Then rs.MoveFirst
    rs.Sort = "Title"
    oReport.component rs, strSupplierName, dR.DOCCode
    oReport.Width = 13000
    oReport.Height = 6000
    oReport.documentName = "Test"
    oReport.Show
    Screen.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.PrintRequest"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.PrintRequest"
End Sub

Private Sub cmdreprint_Click()
    On Error GoTo errHandler
  PrintRequest
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.cmdreprint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.cmdreprint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    'rs.Close
    Set rs = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.UpdateReturnLineRequested FNN(XA.Value(Grid1.Bookmark, 9)), FNN(XA.Value(Grid1.Bookmark, ColIndex + 1))

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
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
        LogSaveToFile "Access violation in frmReturn2: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmReturn2: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_DblClick", , EA_NORERAISE
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
'    ErrorIn "frmReturn2.Grid1_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", _
'         Array(Condition, Split, Bookmark, Col, CellStyle), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", _
         Array(Condition, Split, Bookmark, col, CellStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim lngSupplierID As Long
Dim lngDEALID As Long

    If XA.Value(Bookmark, 3) <> XA.Value(Bookmark, 4) Then
        RowStyle.BackColor = &HC0FFC0 'RGB(100, 174, 20)
    Else
        RowStyle.BackColor = &HDBFAFB 'RGB(100, 174, 20)
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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
'    ErrorIn "frmReturn2.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
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
'    ErrorIn "frmReturn2.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.GetRowType(ColIndex)", ColIndex
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.GetRowType(ColIndex)", ColIndex
End Function

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.Width = 11500
        Me.Height = 6500
        Me.Left = 100
        Me.TOP = 100
    End If
    Set cRL = New c_RL
    cRL.Load lngRID
    LoadGrid
    Me.cmdreprint.Enabled = (dR.Status > 0)
    Me.Caption = "Products for Return no. " & dR.DOCCode & dR.StaffNameB
    txtStatus = dR.StatusF
    lblSupplier.Caption = dR.TPNAME
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
'    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.Grid1_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub ValidateRow(pOKAtPresent As Boolean)
    On Error GoTo errHandler
'    If IsNull(Grid1.Bookmark) Then Exit Sub
'    If (XA.Value(Grid1.Bookmark, 4) > 0) And (XA.Value(Grid1.Bookmark, 5) <= XA.Value(Grid1.Bookmark, 4)) Then
'        lngBadRows = lngBadRows + 1
'        XA.Value(Grid1.Bookmark, 22) = "X"
'    Else
'        If Not pOKAtPresent Then
'            lngBadRows = lngBadRows - 1
'        End If
'        XA.Value(Grid1.Bookmark, 22) = ""
'    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.ValidateRow(pOKAtPresent)", pOKAtPresent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.ValidateRow(pOKAtPresent)", pOKAtPresent
End Sub
Private Sub cmdFixMissingInvoiceRefs_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim oS As New a_Supplier
    oS.Load dR.TPID
    oSM.FixMissingInvoiceRefsOnReturn lngRID, DateAdd("m", oS.ReturnStartMonths * -1, Date), DateAdd("m", oS.ReturnEndMonths * -1, Date)
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmReturn2.cmdFixMissingInvoiceRefs_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmReturn2.cmdFixMissingInvoiceRefs_Click", , EA_NORERAISE
    HandleError
End Sub
