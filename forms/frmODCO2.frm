VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCXB"
Begin VB.Form frmODCO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer order line reconciliation"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   13845
   Begin VB.CommandButton cmdTickSelected 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Tick selected"
      Height          =   315
      Left            =   9645
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2580
      Width           =   1305
   End
   Begin VB.CommandButton cmdUnTickSelected 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Un-tick selected"
      Height          =   315
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2580
      Width           =   1305
   End
   Begin VB.TextBox txtCustomMessage 
      Height          =   1305
      Left            =   8220
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4305
      Width           =   2805
   End
   Begin VB.CommandButton cmdExcelExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Excel"
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
      Left            =   135
      Picture         =   "frmODCO2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4305
      Width           =   1000
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
      Left            =   1170
      Picture         =   "frmODCO2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4305
      Width           =   840
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
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5310
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
      Left            =   11250
      Picture         =   "frmODCO2.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4995
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
      Left            =   12300
      Picture         =   "frmODCO2.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Updates the actions and presents option to print customer advice report"
      Top             =   4995
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frmODCO2.frx":0E28
      TabIndex        =   0
      Top             =   270
      Width           =   13515
   End
   Begin TrueOleDBGrid60.TDBGrid Grid2 
      Height          =   1095
      Left            =   135
      OleObjectBlob   =   "frmODCO2.frx":9C92
      TabIndex        =   5
      Top             =   2940
      Width           =   13170
   End
   Begin TrueOleDBGrid60.TDBGrid G3 
      Height          =   1305
      Left            =   2115
      OleObjectBlob   =   "frmODCO2.frx":F8E8
      TabIndex        =   10
      Top             =   4305
      Width           =   5880
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "General message to customer"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8325
      TabIndex        =   15
      Top             =   4080
      Width           =   2475
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Previous reports"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   2265
      TabIndex        =   11
      Top             =   4080
      Width           =   2475
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
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   12330
      TabIndex        =   8
      Top             =   -60
      Visible         =   0   'False
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
      TabIndex        =   7
      Top             =   4350
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
      Left            =   240
      TabIndex        =   6
      Top             =   2715
      Width           =   1905
   End
End
Attribute VB_Name = "frmODCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COLS As ADODB.Recordset
Dim POLS As ADODB.Recordset
Dim COLActs As ADODB.Recordset
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim XC As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim lngPaid As Long

Private Sub cmdTickSelected_Click()
    On Error GoTo errHandler
Dim i As Integer

    For i = 0 To (Grid1.SelBookmarks.Count - 1)
        Grid1.Bookmark = Grid1.SelBookmarks(i)
        Grid1.Columns(14) = 1
    Next
    Grid1.Update
    Grid1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdTickSelected_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUnTickSelected_Click()
    On Error GoTo errHandler
Dim i As Integer

    For i = 0 To Grid1.SelBookmarks.Count - 1
        Grid1.Bookmark = Grid1.SelBookmarks(i)
        Grid1.Columns(14) = 0
    Next
    Grid1.Update
    Grid1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdUnTickSelected_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.Grid1, Me.Name, Me.Height, Me.Width
    SaveLayout Me.Grid2, Me.Name & "2", Me.Height, Me.Width
    SaveLayout Me.G3, Me.Name & "3", Me.Height, Me.Width
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.mnuSaveLayout"
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = False
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.SetMenu"
End Sub

Public Sub Component3(pCOLs As ADODB.Recordset, pPOLs As ADODB.Recordset, pCOLActs As ADODB.Recordset, dteSince As Date, strOperatorName As String, strCustomers As String, pChangedSince As Date)
    On Error GoTo errHandler
    Set COLS = pCOLs
    Set POLS = pPOLs
    Set COLActs = pCOLActs
    If dteSince > 0 Then
        Me.Caption = "Customer orders due prior to " & Format(dteSince, "dd/mm/yyyy") & IIf(LenB(strCustomers) > 0, " (" & strCustomers & ")", "") & IIf(LenB(strOperatorName) > 0, " (" & strOperatorName & ")", "")
    Else
        If pChangedSince > 0 Then
            Me.Caption = "Customer orders where product status or ETA altered since " & Format(pChangedSince, "dd/mm/yyyy") & IIf(LenB(strCustomers) > 0, " " & strCustomers & " ", "") & IIf(LenB(strOperatorName) > 0, " " & strOperatorName & " ", "")
        Else
            Me.Caption = "All customer orders " & IIf(LenB(strCustomers) > 0, " (" & strCustomers & ")", "") & IIf(LenB(strOperatorName) > 0, " (" & strOperatorName & ")", "")
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Component3(pCOLs,pPOLs,pCOLActs,dteSince,strOperatorName,strCustomers," & _
        "pChangedSince)", Array(pCOLs, pPOLs, pCOLActs, dteSince, strOperatorName, strCustomers, _
         pChangedSince)
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub LoadGrid()
    On Error GoTo errHandler
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
    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
    Next
    For i = 1 To Grid2.Columns.Count
        Grid2.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "2", CStr(i), Grid2.Columns(i - 1).Width)
    Next
    For i = 1 To G3.Columns.Count
        G3.Columns(i - 1).Width = GetSetting("PBKS", Me.Name & "3", CStr(i), G3.Columns(i - 1).Width)
    Next
    XA.Clear
    XA.ReDim 1, COLS.RecordCount, 1, 23
    For lngIndex = 1 To COLS.RecordCount
        XA.Value(lngIndex, 1) = FNS(COLS.fields("TP_NAME"))
        XA.Value(lngIndex, 2) = FNS(COLS.fields("TRCODE"))
        XA.Value(lngIndex, 3) = FNS(COLS.fields("COL_REF"))
        XA.Value(lngIndex, 4) = FNS(COLS.fields("CODEF"))
        XA.Value(lngIndex, 5) = FNS(COLS.fields("P_TITLE"))
        XA.Value(lngIndex, 6) = FND(COLS.fields("TRDATE"))
        XA.Value(lngIndex, 7) = FND(COLS.fields("COL_ETA"))
        XA.Value(lngIndex, 9) = FNN(COLS.fields("COL_QTY"))
        XA.Value(lngIndex, 8) = FNS(COLS.fields("P_Status"))
        XA.Value(lngIndex, 10) = FNN(COLS.fields("COL_QTYDispatched"))
        XA.Value(lngIndex, 11) = FNN(COLS.fields("COL_QTY")) - FNN(COLS.fields("COL_QTYDispatched"))
        XA.Value(lngIndex, 12) = FNN(COLS.fields("P_QtyOnHand"))
        XA.Value(lngIndex, 13) = 0  'cancel
        XA.Value(lngIndex, 14) = ""  'NewETA
        XA.Value(lngIndex, 15) = ""  'message
        XA.Value(lngIndex, 16) = ""  'Select
        XA.Value(lngIndex, 17) = ""  'message
        XA.Value(lngIndex, 18) = FNN(COLS.fields("COL_ID"))
        XA.Value(lngIndex, 19) = FNS(COLS.fields("P_ID"))
        XA.Value(lngIndex, 20) = FNS(COLS.fields("TR_ID"))
        COLS.MoveNext
    Next
    XA.QuickSort 1, COLS.RecordCount, 1, XORDER_ASCEND, XTYPE_STRING, 2, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.LoadGrid"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.LoadGrid"
End Sub
Private Sub MergeArrays()
    On Error GoTo errHandler
Dim i As Integer
Dim idxb As Integer

    For i = 1 To XA.UpperBound(1)
        idxb = 0
        POLS.Filter = "PID = '" & FNS(XA(i, 19)) & "'"
        POLS.Sort = "LastReminderDate DESC"
        If POLS.RecordCount > 0 Then
            If FNS(POLS.fields("LastSupplierMessage")) > "" Then
                XA.Value(i, 22) = FNS(POLS.fields("LastSupplierMessage"))
            End If
        End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.MergeArrays"
End Sub
'Private Function FindinXB(pPID As String) As Integer
'Dim i As Integer
'
'    For i = 1 To XB.UpperBound(1)
'        If FNS(XB.Value(i, 10)) = pPID Then
'            FindinXB = i
'            Exit For
'        End If
'    Next
'
'End Function
Private Sub ReLoadGrid2()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    If IsNull(POLS) Then Exit Sub
    If POLS.RecordCount < 1 Then
        XB.Clear
        Grid2.Array = XB
        Grid2.ReBind
        Exit Sub
    End If
    XB.Clear
    XB.ReDim 1, POLS.RecordCount, 1, 15
    POLS.MoveFirst
    For lngIndex = 1 To POLS.RecordCount
        XB.Value(lngIndex, 1) = FNS(POLS.fields("SupplierName"))
    '    If FND(POLs.Fields("DiarizeDate")) > "2000-01-01" Then
            XB.Value(lngIndex, 2) = FNS(POLS.fields("TRCODE"))
    '    Else
    '        XB.Value(lngIndex, 2) = "n/a"
   '     End If
        XB.Value(lngIndex, 3) = FND(POLS.fields("TRDATE"))
        XB.Value(lngIndex, 4) = FNN(POLS.fields("QTYOS"))
        XB.Value(lngIndex, 5) = FNS(POLS.fields("POLSTATUS"))
        If FND(POLS.fields("LastReminderDate")) > "2000-01-01" Then
            XB.Value(lngIndex, 6) = Format(FND(POLS.fields("LastReminderDate")), "DD/mm/yyyy")
        Else
            XB.Value(lngIndex, 6) = ""
        End If
        XB.Value(lngIndex, 7) = FNS(POLS.fields("LastSupplierMessage"))
        
        XB.Value(lngIndex, 8) = FNS(POLS.fields("P_STATUS"))
        XB.Value(lngIndex, 11) = FNS(POLS.fields("TRID"))
        XB.Value(lngIndex, 10) = FNS(POLS.fields("PID"))
        
'        If FNS(POLS.Fields("PSC_OLDSTATUS")) <> FNS(POLS.Fields("PSC_NEWSTATUS")) Then
'            XB.Value(lngIndex, 11) = FNS(POLS.Fields("PSC_OLDSTATUS")) & "/" & FNS(POLS.Fields("PSC_ONEWSTATUS"))
'        End If
'        If FND(POLS.Fields("PSC_OldEta")) <> FND(POLS.Fields("PSC_NewETA")) Then
'            XB.Value(lngIndex, 11) = FND(POLS.Fields("PSC_OldEta")) & "/" & FND(POLS.Fields("PSC_NewETA"))
'        End If
'        XB.Value(lngIndex, 12) = FNS(POLS.Fields("PSC_SupplierMessage"))
'        XB.Value(lngIndex, 13) = FND(POLS.Fields("PSC_DATE"))
        
        
        POLS.MoveNext
    Next
    XB.QuickSort 1, POLS.RecordCount, 1, XORDER_DESCEND, XTYPE_DATE
    Grid2.Array = XB
    Grid2.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.ReLoadGrid2"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.ReLoadGrid2"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("You are choosing to close without taking action?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then Exit Sub
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    cmdPrint.Visible = True
p 1
    Grid1.Width = NonNegative_Lng(Me.Width - (Grid1.Left + 400))
    cmdTickSelected.Left = NonNegative_Lng(Me.Width - 3200)
p 2
    cmdUnTickSelected.Left = NonNegative_Lng(Me.Width - 1850)
    lngDiff = Grid1.Height
    
p 3
    Grid1.Height = NonNegative_Lng(Me.Height - (Grid1.TOP + 3800))
    lngDiff = (Grid1.Height - lngDiff)
    Grid2.TOP = Grid2.TOP + lngDiff
    Grid2.Width = NonNegative_Lng(Me.Width - (Grid2.Left + 400))
    
p 4
    Label1.TOP = Label1.TOP + lngDiff
    Label2.TOP = Label2.TOP + lngDiff
    Label3.TOP = Label3.TOP + lngDiff
    
p 5
    cmdTickSelected.TOP = cmdTickSelected.TOP + lngDiff
    cmdUnTickSelected.TOP = cmdUnTickSelected.TOP + lngDiff
    Me.txtCustomMessage.TOP = txtCustomMessage.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    
p 6
    cmdExcelExport.TOP = cmdExcelExport.TOP + lngDiff
    cmdPrintPrev.TOP = cmdPrintPrev.TOP + lngDiff
    cmdCancel.TOP = cmdCancel.TOP + lngDiff
    cmdOK.TOP = cmdOK.TOP + lngDiff
    G3.TOP = G3.TOP + lngDiff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Form_Resize", , EA_NORERAISE, , "strErrPos", Array(strErrPos)
    HandleError
End Sub



Private Sub Form_Unload(Cancel As Integer)
    SaveLayout Me.Grid1, Me.Name, Me.Height, Me.Width
    SaveLayout Me.Grid2, Me.Name & "2", Me.Height, Me.Width
    SaveLayout Me.G3, Me.Name & "3", Me.Height, Me.Width

End Sub

Private Sub Grid1_ButtonClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If ColIndex = 10 Then
        If XA(Grid1.Bookmark, 13) = 1 Then
            XA(Grid1.Bookmark, 13) = 0
        Else
            XA(Grid1.Bookmark, 13) = 1
        End If
    End If
    Grid1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_ButtonClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim strStatus As String
    If XA(Bookmark, 13) = -1 Then
        RowStyle.BackColor = COLOUR_FULFIL_MORETHANONHAND
    Else
        RowStyle.BackColor = COLOR_PALEYELLOW
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle), _
'         EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, RowStyle), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
    If IsNull(Grid1.Bookmark) Then Exit Sub
    POLS.Filter = "PID = '" & FNS(XA(Grid1.Bookmark, 19)) & "'"
    ReLoadGrid2
    COLActs.Filter = "COLID = " & FNS(XA(Grid1.Bookmark, 18))
    ReLoadGrid3
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim i As Long
Dim oSM As New z_StockManager
Dim XB As New XArrayDB
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim oSQL As New z_SQL
Dim f As New frmTrackingActions
Dim Res As Boolean

    If oPC.Configuration.SignTransactions = True Then
        If SecurityControl(enSECURITY_CO_SIGN, , "Sign this action", DOCAPPROVAL, , , gSTAFFID) = False Then
               Exit Sub
        End If
    End If
    Screen.MousePointer = vbHourglass
    
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "doc_COL_ACTION"
            .chCreate "MessageType"
                .elText = "COL_ACTION"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "WORKSTATION"
                .elText = oPC.WorkstationName
            .elCreateSibling "CUSTOMERMESSAGE", True
                If txtCustomMessage > "" Then
                    .elText = txtCustomMessage
                Else
                    If XA.UpperBound(1) > 1 Then
                       .elText = "These items are overdue. We are following up with suppliers. More information will follow."
                    Else
                        .elText = "This item is overdue. We are following up with suppliers. More information will follow."
                    End If
                End If
            .elCreateSibling "DetailLines", True
            For i = 1 To XA.UpperBound(1)
                Grid1.Bookmark = i
                If Grid1.Columns(15) <> "" Then
                    .chCreate "ITEM"
                    .chCreate "COLID"
                        .elText = XA.Value(i, 18)
                    .elCreateSibling "CANCEL", True
                        .elText = IIf(FNB(XA.Value(i, 13)), 1, 0)
                    .elCreateSibling "DETAILEDMESSAGE", True
                        .elText = FNS(XA.Value(i, 15))
                    .elCreateSibling "NEWETA", True
                        .elText = FNS(XA.Value(i, 14))
                    .navUP
                    .navUP
                 End If
            Next i

         XMLArgs = .docXML
  
    End With
    
    If XMLArgs > "" Then
        Res = oSM.ActionODCOL(XMLArgs, lngPaid)
        If Res = False Then
            MsgBox "There are too many orders to report on in one batch. Restrict your selection and re-find the outstanding orders.", vbInformation, "Can't do this (too many overdue orders in one batch)"
            Unload Me
            Set oSM = Nothing
            Set rs = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = vbDefault

    If Forms(0).frmTRacking Is Nothing Then
        Set Forms(0).frmTRacking = New frmTrackingActions
    End If
    Forms(0).frmTRacking.component "", ""
    Forms(0).frmTRacking.Show
    Unload Me
    Set oSM = Nothing
    Set rs = Nothing
    
'errHandler:
'    ErrPreserve
'    If err = 521 Then
'        Resume Next
'    End If
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.cmdOK_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim rpt As New arODCO
    If XA.UpperBound(1) = 0 Then
        MsgBox "There are no lines to print.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    rpt.component XA
    rpt.Show
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.cmdPrint_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdExcelExport_Click()
    On Error GoTo errHandler
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
    MergeArrays
    rpt.component XA, Me.Caption
    rpt.Run False
    If Not fs.FolderExists(oPC.LocalFolder & "TEMP") Then
        fs.CreateFolder oPC.LocalFolder & "TEMP"
    End If
    sFile = oPC.LocalFolder & "TEMP\OS_CustOrders"
    i = 0
    Do Until fs.FileExists(sFile & ".XLS") = False
        i = i + 1
        sFile = sFile & "_" & CStr(i)
    Loop
        
        
        
    sFile = sFile & ".XLS"
    xls.FileName = sFile
    If rpt.Pages.Count > 0 Then
        xls.Export rpt.Pages
    End If
    Screen.MousePointer = vbDefault
    If MsgBox("Spreadsheet file saved in: " & sFile & vbCrLf & "Do you want to open it?", vbQuestion + vbYesNo, "Export complete") = vbYes Then
            strExecutable = GetPDFExecutable(sFile)
          If strExecutable = "" Then
              MsgBox "There is no application set on this computer to open the file: " & sFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
          Else
             F_7_AB_1_ShellAndWaitSimple strExecutable & " " & sFile, vbHide, 1000
          End If
    End If

'errHandler:
'    ErrPreserve
'    If err = 70 Then
'        MsgBox "It looks like a previously generated document is open in Excel. Please close before continuing.", vbInformation + vbOKOnly, "Can't do this"
'        Exit Sub
'    End If
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.cmdExcelExport_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdExcelExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrintPrev_Click()
    On Error GoTo errHandler
Dim frm As New frmPrintPreviousActions
    frm.Show vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.cmdPrintPrev_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.cmdPrintPrev_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product
    If IsNull(Grid1.Bookmark) Then Exit Sub
    strPID = XA.Value(Grid1.Bookmark, 19)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.component oProd
        frm.Show
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmODCO: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmODCO: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If ColIndex = 10 Then
                XA(Grid1.Bookmark, 13) = Trim(Grid1.text)
    End If
    If ColIndex = 11 Or ColIndex = 12 Or ColIndex = 13 Then
                XA(Grid1.Bookmark, 16) = -1  'FNN(Trim(Grid1.Text))
    End If
    Grid1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    
'    If ColIndex = 10 Then
'        If Trim(Grid1.Text) > "" Then
'            If Not IsNumeric(Trim(Grid1.Text)) Then
'                Cancel = True
'                Grid1.Columns(ColIndex).Value = OldValue
'                Beep
'            End If
'        End If
'    End If
'    If ColIndex = 11 Then
'        If Trim(Grid1.Text) > "" Then
'            If (Not IsNumeric(Trim(Grid1.Text))) Then
'                Cancel = True
'                Grid1.Columns(ColIndex).Value = OldValue
'                Beep
'            End If
'        End If
'    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub MarkRowsValid(pOK As Integer, pKey As String)
    On Error GoTo errHandler
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.MarkRowsValid(pOK,pKey)", Array(pOK, pKey)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.MarkRowsValid(pOK,pKey)", Array(pOK, pKey)
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    SetMenu
    If Me.WindowState <> 2 Then
        Me.Width = 12000
        Me.Height = 6015
        Me.Left = 100
        Me.TOP = 100
    End If
    Set XA = New XArrayDB
    Set XB = New XArrayDB
    Set XC = New XArrayDB
    
    SetGridLayout Me.Grid1, Me.Name
    SetGridLayout Me.Grid2, Me.Name
    SetFormSize Me

    LoadGrid
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.Form_Load", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
    Grid1.Update
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.Grid1_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub GRID2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.GRID2_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.GRID2_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub GRID2_DblClick()
    On Error GoTo errHandler
Dim lngTRID As Long
Dim frm As frmPOPreview
    If IsNull(Grid2.Bookmark) Then Exit Sub
    Set frm = New frmPOPreview
    frm.component FNN(XB.Value(Grid2.Bookmark, 11))
    frm.Show
    
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmODCO: GRID2_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmODCO: GRID2_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.GRID2_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.XA_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.Grid1_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4, 5
            GetRowType = XTYPE_STRING
        Case 6
            GetRowType = XTYPE_DATE
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.GetRowType(ColIndex)", ColIndex
End Function



Private Sub ReLoadGrid3()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String


    If IsNull(COLActs) Then Exit Sub
    If COLActs.RecordCount < 1 Then Exit Sub
    XC.Clear
    XC.ReDim 1, COLActs.RecordCount, 1, 4
    COLActs.MoveFirst
    For lngIndex = 1 To COLActs.RecordCount
        
        If FND(COLActs.fields("CustomerReportDate")) > "2000-01-01" Then
            XC.Value(lngIndex, 1) = Format(FND(COLActs.fields("CustomerReportDate")), "dd/mm/yyyy")
        Else
            XC.Value(lngIndex, 1) = ""
        End If
        XC.Value(lngIndex, 2) = FNS(COLActs.fields("CustomerReport")) & "/" & FNS(COLActs.fields("AvailabilityStatus")) & "/" & FNS(COLActs.fields("COLAction"))
        COLActs.MoveNext
    Next
    XC.QuickSort 1, COLActs.RecordCount, 1, XORDER_DESCEND, XTYPE_DATE
    G3.Array = XC
    G3.ReBind

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmODCO.ReLoadGrid3"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.ReLoadGrid3"
End Sub

Private Sub txtCustomMessage_Change()
    On Error GoTo errHandler
    txtCustomMessage = HandleTextWithBites(txtCustomMessage)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODCO.txtCustomMessage_Change", , EA_NORERAISE
    HandleError
End Sub
