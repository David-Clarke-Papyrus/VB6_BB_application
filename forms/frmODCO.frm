VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmODCO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer order line reconciliation"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   11745
   Begin VB.CommandButton cmdExcelExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Exp"
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
      Left            =   1140
      Picture         =   "frmODCO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5235
      Width           =   1000
   End
   Begin VB.TextBox txtPreviousReports 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   3885
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   5220
      Width           =   5415
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print list"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5235
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrintPrev 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print previous actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5580
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   9450
      Picture         =   "frmODCO.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5220
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   10470
      Picture         =   "frmODCO.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Updates the actions and presents option to print customer advice report"
      Top             =   5220
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3285
      Left            =   150
      OleObjectBlob   =   "frmODCO.frx":0A9E
      TabIndex        =   0
      Top             =   345
      Width           =   11340
   End
   Begin TrueOleDBGrid60.TDBGrid GRID2 
      Height          =   1170
      Left            =   165
      OleObjectBlob   =   "frmODCO.frx":7F9C
      TabIndex        =   5
      Top             =   3945
      Width           =   11325
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   9390
      TabIndex        =   9
      Top             =   -30
      Width           =   360
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Actions"
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
      Height          =   240
      Left            =   2145
      TabIndex        =   8
      Top             =   5205
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase orders"
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
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   3690
      Width           =   1905
   End
End
Attribute VB_Name = "frmODCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRS As ADODB.Recordset
Dim cODCO As c_COLOD
Attribute cODCO.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim XC As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Public Sub mnuSaveLayout()
    On Error GoTo ErrHandler
    SaveLayout Me.Grid1, Me.Name
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
End Sub

Public Sub Component2(PID As String)

End Sub

Private Sub Form_Activate()
    SetMenu
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub
Public Sub Component(pODCO As c_COLOD, dteSince As Date, strOperatorName As String, strCustomers As String)
    On Error GoTo ErrHandler
Dim strSQL As String
    Set cODCO = pODCO
    Me.Caption = "Customer orders due prior to " & Format(dteSince, "dd/mm/yyyy") & IIf(LenB(strCustomers) > 0, " (" & strCustomers & ")", "") & IIf(LenB(strOperatorName) > 0, " (" & strOperatorName & ")", "")
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Component(pODCO)", pODCO
End Sub

Private Sub LoadGrid()
    On Error GoTo ErrHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim tmp As String
Dim qtyRecs As Long
Dim lngAwaiting As Long
Dim lngAllocation As Long
Dim lngAvailableToAllocate As Long
Dim i As Integer
Dim dODCO As d_COLine

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = cODCO.Count
    XA.ReDim 1, lngArrayRows, 1, 22
    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
    Next
    
    For Each dODCO In cODCO
            XA.Value(lngIndex, 1) = dODCO.CustName 'dODCO.Titleshort(15) & "  (" & dODCO.QtyOnHand & ")" '& "(" & dODCO.QtyJustReceived & ")"
            XA.Value(lngIndex, 2) = dODCO.DocCode
            XA.Value(lngIndex, 3) = dODCO.Ref
            XA.Value(lngIndex, 4) = dODCO.code
            XA.Value(lngIndex, 5) = dODCO.Title
            
            XA.Value(lngIndex, 6) = dODCO.DocDateF
            XA.Value(lngIndex, 7) = dODCO.ETAF
            
            
            
            XA.Value(lngIndex, 8) = dODCO.qty
            XA.Value(lngIndex, 9) = dODCO.qty - dODCO.QtyDispatched
            XA.Value(lngIndex, 10) = dODCO.QtyOnHand
            XA.Value(lngIndex, 11) = ""
             XA.Value(lngIndex, 12) = ""
            XA.Value(lngIndex, 13) = ""
           
            
            
            XA.Value(lngIndex, 12) = dODCO.COLID
            XA.Value(lngIndex, 15) = dODCO.previousCustomerReports
            XA.Value(lngIndex, 16) = dODCO.PID
            XA.Value(lngIndex, 17) = dODCO.LastActionDateF
            XA.Value(lngIndex, 18) = dODCO.lastaction
            lngIndex = lngIndex + 1
    Next
    XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 4, XORDER_ASCEND, XTYPE_DATE
    Grid1.Array = XA
    Grid1.ReBind
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrHandler
    If MsgBox("Note: Any entries in the last column will be lost?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    Unload Me
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    Grid1.Width = Me.Width - (Grid1.Left + 400)
    lngDiff = Grid1.Height
    Grid1.Height = Me.Height - (Grid1.Top + 2900)
    lngDiff = Grid1.Height - lngDiff
    GRID2.Top = GRID2.Top + lngDiff
    Label1.Top = Label1.Top + lngDiff
    Label2.Top = Label2.Top + lngDiff
    cmdPrint.Top = cmdPrint.Top + lngDiff
    cmdExcelExport.Top = cmdExcelExport.Top + lngDiff
    cmdPrintPrev.Top = cmdPrintPrev.Top + lngDiff
    cmdCancel.Top = cmdCancel.Top + lngDiff
    cmdOK.Top = cmdOK.Top + lngDiff
    txtPreviousReports.Top = txtPreviousReports.Top + lngDiff
End Sub

Private Sub lblHelp_Click()
    On Error GoTo ErrHandler
Dim strHelp As String

    strHelp = "Action has two parts:" & vbCrLf & vbCrLf _
                & "1: " & vbCrLf _
                & "   '1W' diarize 1 week hence" & vbCrLf _
                & "   '1M' diarize 1 month hence" & vbCrLf _
                & "   '2W' diarize 2 weeks hence" & vbCrLf _
                & "   '2M' diarize 2 months hence" & vbCrLf _
                & "   '3W' diarize 3 weeks hence" & vbCrLf _
                & "   '3M' diarize 3 months hence" & vbCrLf & vbCrLf _
                & "2: (Report to customer)" & vbCrLf _
                & "   Type report message and use dictionary codes" & vbCrLf _
                & "   if desired by enclosing in square brackets." & vbCrLf & vbCrLf _
                & "   e.g. 3WSorry John[a][3w] will diarize for three weeks hence and " & vbCrLf _
                & "         on then report will say  'Sorry John we are chasing the supplier for delivery Delivery expected in three weeks'" & vbCrLf _
                & "         NOTE: This depends on the dictionary codes being set to the default values." & vbCrLf
    MsgBox strHelp, vbOKOnly, "How to enter action codes"
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdHelp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
Dim rs As ADODB.Recordset
Dim i As Long
Dim oSM As New z_StockManager
Dim XB As New XArrayDB

    Set rs = New ADODB.Recordset
    rs.Fields.Append "COLID", adInteger
    rs.Fields.Append "ACT1", adChar, 2
    rs.Fields.Append "ACT2", adVarChar, 120
    rs.Open
    For i = 1 To lngArrayRows
        If (XA.Value(i, 11) > "" Or XA.Value(i, 14) > "") Then
            rs.AddNew
            rs.Fields("COLID") = XA.Value(i, 12)
            rs.Fields("ACT1") = XA.Value(i, 14)
            rs.Fields("ACT2") = XA.Value(i, 11)
            rs.Update
        End If
    Next i
    oSM.ActionODCO rs
    
    oSM.saveCustomerOrderStatusReport XA, oPC.WorkstationName
    oSM.PrintCustomerOrderStatusReport oPC.WorkstationName
        
    Unload Me
    
    Set oSM = Nothing
    Set rs = Nothing
    
    Exit Sub
ErrHandler:
    ErrPreserve
    If Err = 521 Then
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrHandler
Dim rpt As New arODCO
    If XA.UpperBound(1) = 0 Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    rpt.Component XA
    rpt.Show vbModal
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdExcelExport_Click()
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arODCO_ForExcel
Dim i As Long
Dim strExecutable As String

    If XA.UpperBound(1) = 0 Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    rpt.Component XA
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "\TEMP"
    End If
    sFile = oPC.LocalFolder & "\TEMP\OS_CustOrders.XLS"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    xls.FileName = sFile
    If rpt.Pages.Count > 0 Then
        xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
            F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbNormalFocus
    End If

End Sub

Private Sub cmdPrintPrev_Click()
    On Error GoTo ErrHandler
Dim frm As New frmPrintPreviousActions
    frm.Show vbModal
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdPrintPrev_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdReset_Click()
'Dim i As Integer
'    If MsgBox("You want to clear all entries in the last column?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
'    For i = 1 To lngArrayRows
'        XA.Value(i, 8) = ""
'    Next i
'    Grid1.ReBind
'End Sub


Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo ErrHandler
Dim i As Integer
Dim iMaxRows As Integer
    If IsNull(Grid1.Bookmark) Then Exit Sub
    Set XC = Nothing
    Set XC = New XArrayDB
    iMaxRows = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs.Count
    XC.ReDim 1, iMaxRows, 1, 10
    For i = 1 To iMaxRows
        XC.Value(i, 1) = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs(i).supplier
        XC.Value(i, 2) = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs(i).DocCode
        XC.Value(i, 3) = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs(i).DocDateF
        XC.Value(i, 4) = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs(i).qtyTotal
        XC.Value(i, 5) = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs(i).previousactions
        XC.Value(i, 6) = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").POLs(i).TRID
    Next
    XC.QuickSort 1, iMaxRows, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_DATE
    GRID2.Array = XC
    GRID2.ReBind
    Me.txtPreviousReports = cODCO(XA.Value(Grid1.Bookmark, 12) & "k").previousCustomerReports
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo ErrHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product

    strPID = XA.Value(Grid1.Bookmark, 18)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.Component oProd
        frm.Show
    End If
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo ErrHandler
Dim strStatus As String
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle), _
         EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo ErrHandler
Dim strTmp As String
Dim bTmp As Boolean
Dim f1 As String
Dim f2 As String

    strTmp = ConvertCOLActionCodes(Grid1.Text, bTmp, f1, f2)
    Grid1.Text = strTmp & " " & f2
    XA.Value(Grid1.Bookmark, 14) = f1
    XA.Value(Grid1.Bookmark, 15) = f2
    
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub MarkRowsValid(pOK As Integer, pKey As String)
    On Error GoTo ErrHandler
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.MarkRowsValid(pOK,pKey)", Array(pOK, pKey)
End Sub
Private Sub Form_Load()
    On Error GoTo ErrHandler
    If Me.WindowState <> 2 Then
        Me.Width = 12000
        Me.Height = 6500
        Me.Left = 100
        Me.Top = 100
    End If
    LoadGrid
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_LostFocus()
    On Error GoTo ErrHandler
    Grid1.Update
    Grid1.ReBind
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub GRID2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo ErrHandler
Cancel = True
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.GRID2_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub GRID2_DblClick()
Dim lngTRID As Long
Dim frm As frmPOPreview

    Set frm = New frmPOPreview
    frm.Component FNN(XC.Value(GRID2.Bookmark, 6))
    frm.Show
    
End Sub
Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo ErrHandler
Static Direction As Variant
    
    If XA.UpperBound(1) = 0 Then Exit Sub
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
 '   If ColIndex = 0 Then ColIndex = 11
    
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    Grid1.Refresh
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.XA_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2, 3, 4, 5
            GetRowType = XTYPE_STRING
        Case 6
            GetRowType = XTYPE_DATE
    End Select
End Function

