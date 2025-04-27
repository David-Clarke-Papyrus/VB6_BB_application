VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmODPO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Purchase order line reconciliation"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11250
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   11250
   Begin VB.TextBox lblPastActions 
      Height          =   885
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmODPO.frx":0000
      Top             =   4290
      Width           =   9015
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
      Left            =   9945
      Picture         =   "frmODPO.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5250
      Width           =   1000
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
      Left            =   8940
      Picture         =   "frmODPO.frx":0390
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5250
      Width           =   1000
   End
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
      Left            =   1335
      Picture         =   "frmODPO.frx":071A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5250
      Width           =   1000
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00E6E7CB&
      Caption         =   "&Next"
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
      Height          =   450
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5295
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H00E6E7CB&
      Cancel          =   -1  'True
      Caption         =   "&Prev"
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
      Height          =   450
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5295
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print list"
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
      Left            =   120
      Picture         =   "frmODPO.frx":0AA4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5250
      Width           =   1200
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Reset"
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
      Left            =   5280
      Picture         =   "frmODPO.frx":0E2E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "This removes all actions from the actions column"
      Top             =   5250
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   3825
      Left            =   120
      OleObjectBlob   =   "frmODPO.frx":11B8
      TabIndex        =   0
      Top             =   330
      Width           =   10815
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
      Left            =   9060
      TabIndex        =   5
      Top             =   -60
      Width           =   360
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Previous actions"
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
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   4365
      Width           =   1680
   End
   Begin VB.Label lblPastActionsobs 
      BackColor       =   &H00DBFAFB&
      Height          =   345
      Left            =   6630
      TabIndex        =   3
      Top             =   5475
      Visible         =   0   'False
      Width           =   2865
   End
End
Attribute VB_Name = "frmODPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRS As ADODB.Recordset
Dim cODPO As c_POLsOS
Attribute cODPO.VB_VarHelpID = -1
Dim XA As XArrayDB
Dim XB As XArrayDB
Dim iRecs As Integer
Dim lngArrayRows As Long
Dim bEOF As Boolean
Dim bBOF As Boolean
Dim bActioned As Boolean
Dim bActionTaken As Boolean
Dim flgLoading As Boolean

Public Sub mnuSaveLayout()
    On Error GoTo errHandler
    SaveLayout Me.Grid1, Me.Name
    Exit Sub
errHandler:
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

Private Sub cmdExportExcel_Click()

End Sub

Private Sub Form_Activate()
    SetMenu
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub
Public Sub Component(pODPO As c_POLsOS, dteSince As Date, strOperatorName As String, Optional pCust As String, Optional pSUpplierName As String)
    On Error GoTo errHandler
Dim strSQL As String
    Set cODPO = pODPO
    bActionTaken = False
    bActioned = False
    If dteSince = 0 Then
        Me.Caption = "Purchase orders for " & IIf(LenB(strOperatorName) > 0, " (" & strOperatorName & ")", "") & IIf(pCust > "", " Customer: " & pCust, "")
    Else
        Me.Caption = "Purchase orders due prior to " & Format(dteSince, "dd/mm/yyyy") & IIf(LenB(strOperatorName) > 0 And strOperatorName <> "<All>", " (" & strOperatorName & ")", "") & IIf(pCust > "", " Customer: " & pCust, "") & IIf(pSUpplierName > "", " Supplier: " & pSUpplierName, "")
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Component(pODPO)", pODPO
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
Dim dODPO As d_POLine

    i = 0
    Set XA = New XArrayDB
    XA.Clear
    iRecs = i
    lngIndex = 1
    lngArrayRows = cODPO.Count
    XA.ReDim 1, lngArrayRows, 1, 17
    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
    Next
    
    For Each dODPO In cODPO
            XA.Value(lngIndex, 1) = dODPO.supplier
            XA.Value(lngIndex, 2) = dODPO.DocCode
            XA.Value(lngIndex, 3) = dODPO.code
            XA.Value(lngIndex, 4) = dODPO.Title ' & ": " & dODPO.code
            XA.Value(lngIndex, 5) = dODPO.DocDateF
            XA.Value(lngIndex, 6) = dODPO.QtyFirm & "/" & dODPO.QtySS
            XA.Value(lngIndex, 7) = dODPO.ReceivedSoFar
            XA.Value(lngIndex, 8) = dODPO.QtyOS
            XA.Value(lngIndex, 9) = "No action"
            
            XA.Value(lngIndex, 10) = ""
           
            XA.Value(lngIndex, 11) = dODPO.previousactions
            
            XA.Value(lngIndex, 12) = dODPO.POLID
            XA.Value(lngIndex, 13) = "N"
            XA.Value(lngIndex, 14) = "N"
            XA.Value(lngIndex, 15) = ""      'dODPO.Title
            XA.Value(lngIndex, 16) = dODPO.pID
      '      XA.Value(lngIndex, 16) = "N"
            lngIndex = lngIndex + 1
    Next
    XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind
  '  Grid1.Splits(0).Columns(0).Merge = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.LoadGrid"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
'    If cODPO.IsEditing Then cODPO.CancelEdit
'    Set cODPO = Nothing
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdHelp_Click()
    On Error GoTo errHandler
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdHelp_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdOK_Click()  'done
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim i As Long
Dim oSM As New z_StockManager
Dim bReminders As Boolean
Dim arReminders As arPOReminder
Dim frm As frmPrintRemindersheet
Dim OpenResult As Integer

    Screen.MousePointer = vbHourglass
    bActioned = True
    bActionTaken = False
    bReminders = False
    Set rs = New ADODB.Recordset
    rs.Fields.Append "POLID", adInteger
    rs.Fields.Append "ACT1", adChar, 1
    rs.Fields.Append "ACT2", adChar, 1
    rs.Fields.Append "ACT3", adChar, 2
    rs.Fields.Append "PID", adVarChar, 40
    rs.Fields.Append "Note", adVarChar, 150
    rs.Open
    For i = 1 To lngArrayRows
        If XA.Value(i, 13) <> "N" Or XA.Value(i, 14) <> "N" Or XA.Value(i, 15) <> "" Or XA.Value(i, 10) <> "" Then
            rs.AddNew
            rs.Fields("POLID") = XA.Value(i, 12)
            rs.Fields("ACT1") = XA.Value(i, 13)
            rs.Fields("ACT2") = XA.Value(i, 14)
            rs.Fields("Note") = XA.Value(i, 10)
            rs.Fields("PID") = XA.Value(i, 16)
            If XA.Value(i, 14) = "R" Then
                bReminders = True
            End If
            rs.Fields("ACT3") = XA.Value(i, 15)
            rs.Update
        End If
    Next i
    If Not (rs.eof And rs.BOF) Then
        oSM.ActionODPO rs
    End If

    If bReminders Then
        Screen.MousePointer = vbDefault
        Set frm = New frmPrintRemindersheet
        frm.Show vbModal
        oSM.PrintPurchaseOrderReminderReport oPC.WorkstationName, frm.chkPagePerSupplier = 1
        Unload frm
        Screen.MousePointer = vbDefault
    End If
    
    cmdOK.Enabled = False
    MsgBox "Done", , "Status"
    Unload Me
    Set oSM = Nothing
    Set rs = Nothing
    
   Exit Sub
errHandler:
    ErrPreserve
    Screen.MousePointer = vbDefault
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim rpt As New arODPO
    rpt.Component XA
    rpt.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdExcelExport_Click()
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean
Dim fs As New FileSystemObject
Dim rpt As New arODPO_ForExcel
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
    sFile = oPC.LocalFolder & "\TEMP\OS_PurchaseOrders.XLS"
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

Private Sub cmdReset_Click() 'done
    On Error GoTo errHandler
Dim i As Integer
    If MsgBox("You want to clear all entries in the Action column?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    For i = 1 To lngArrayRows
        XA.Value(i, 9) = "No action"
    Next i
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.cmdReset_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    Grid1.Width = Me.Width - (Grid1.Left + 400)
    lngDiff = Grid1.Height
    Grid1.Height = Me.Height - (Grid1.Top + 2320)
    lngDiff = Grid1.Height - lngDiff
    cmdPrint.Top = cmdPrint.Top + lngDiff
    cmdExcelExport.Top = cmdExcelExport.Top + lngDiff
    cmdReset.Top = cmdReset.Top + lngDiff
    cmdCancel.Top = cmdCancel.Top + lngDiff
    cmdOK.Top = cmdOK.Top + lngDiff
    
    Label1.Top = Label1.Top + lngDiff
    lblPastActions.Top = lblPastActions.Top + lngDiff

End Sub

Private Sub Grid1_AfterUpdate()
    On Error GoTo errHandler
'    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_AfterUpdate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer) 'done
    On Error GoTo errHandler
If IsNull(Grid1.Bookmark) Then Exit Sub
Debug.Print Grid1.Bookmark & "  :  " & XA.Value(Grid1.Bookmark, 11)
    lblPastActions.Text = XA.Value(Grid1.Bookmark, 11)
    bActionTaken = True
    Grid1.SelStart = 1
    Grid1.SelLength = 99
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick() 'done
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product

    strPID = XA.Value(Grid1.Bookmark, 16)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.Component oProd
        frm.Show
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer) 'done
    On Error GoTo errHandler
Dim strTmp As String
Dim bTmp As Boolean
Dim f1 As String
Dim f2 As String
Dim f3 As String
Dim strNote As String
    If ColIndex = 8 Then
    strTmp = ConvertPOLActionCodes(Grid1.Text, bTmp, f1, f2, f3)
    Cancel = bTmp
    If Not Cancel Then
        Grid1.Text = strTmp
        cmdCancel.Caption = "&Cancel"
        XA.Value(Grid1.Bookmark, 13) = f1
        XA.Value(Grid1.Bookmark, 14) = f2
        XA.Value(Grid1.Bookmark, 15) = f3
    End If
    ElseIf ColIndex = 9 Then
        XA.Value(Grid1.Bookmark, 10) = Grid1.Text
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Me.Width = 12000
        Me.Height = 6500
        Me.Left = 100
        Me.Top = 100
    End If
    Me.Caption = "Track overdue purchase orders"
    SetGridLayout Me.Grid1, Me.Name
    
    LoadGrid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
 '   Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub






Private Sub Grid1_SelChange(Cancel As Integer)
    On Error GoTo errHandler
  ' Grid1.SelStart = 1
  '  Grid1.SelLength = 99
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid1_SelChange(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lblHelp_Click()
Dim strHelp As String
    strHelp = "Action has three parts:" & vbCrLf & vbCrLf _
                & "1: " & vbCrLf _
                & "   'N' no change in book status" & vbCrLf _
                & "   'O' book is out of print" & vbCrLf _
                & "   'R' book is being reprinted" & vbCrLf & vbCrLf _
                & "2: " & vbCrLf _
                & "   'N' no action to be taken" & vbCrLf _
                & "   'C' cancel order line" & vbCrLf _
                & "   'R' print reminder" & vbCrLf & vbCrLf _
                & "3: " & vbCrLf _
                & "   '1W' diarize 1 week hence" & vbCrLf _
                & "   '1M' diarize 1 month hence" & vbCrLf _
                & "   '2W' diarize 2 weeks hence" & vbCrLf _
                & "   '2M' diarize 2 months hence" & vbCrLf _
                & "   '3W' diarize 3 weeks hence" & vbCrLf _
                & "   '3M' diarize 3 months hence" & vbCrLf & vbCrLf _
                & "   e.g. NR3W means 'No change in book status," & vbCrLf _
                & "         print reminder," & vbCrLf _
                & "         diarize three weeks hence" & vbCrLf _
                & "   e.g. RN3M means 'book is being reprinted," & vbCrLf _
                & "         no action to be taken," & vbCrLf _
                & "         diarize three months hence."
    MsgBox strHelp, vbOKOnly, "How to enter action codes"

End Sub
Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant
    If flgLoading Then Exit Sub

    Screen.MousePointer = vbHourglass
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
  '  If ColIndex = 2 Then
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
  '  Else
  '      XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1), 3, XORDER_DESCEND, XTYPE_DATE  'XTYPE_INTEGER
  ' End If
    
    Grid1.Refresh
    Screen.MousePointer = vbDefault

    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmODPO.Grid_HeadClick(ColIndex)", ColIndex
End Sub

'Public Sub mnuSaveLayout()
'    On Error GoTo errHandler
'    SaveLayout Me.Grid, Me.Name
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
'End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    Select Case ColIndex
        Case 1, 2
            GetRowType = XTYPE_STRING
        Case 5
            GetRowType = XTYPE_DATE
    End Select
End Function
