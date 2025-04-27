VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmREORDER_CO 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Reorder products for customer's orders"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   FillColor       =   &H00FFC0FF&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   11280
   Begin VB.CommandButton cmdExp 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Exp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5250
      Width           =   420
   End
   Begin VB.ComboBox cboStaffMember 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1530
      TabIndex        =   8
      Top             =   0
      Width           =   2835
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   285
      Left            =   5310
      TabIndex        =   7
      Top             =   5460
      Visible         =   0   'False
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5250
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   570
      Left            =   8790
      Picture         =   "frmREORDER_CO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5250
      Width           =   1035
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Use existing reorder slate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1725
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5235
      Width           =   1590
   End
   Begin VB.CommandButton cmdRecalc 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Prepare new reorder slate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5220
      Width           =   1590
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
      Height          =   375
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5265
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Generate orders"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5235
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid Grid1 
      Height          =   4590
      Left            =   75
      OleObjectBlob   =   "frmREORDER_CO.frx":00AB
      TabIndex        =   0
      Top             =   420
      Width           =   11205
   End
   Begin VB.Label lblStaffMember 
      Caption         =   "Staff member"
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
      Height          =   270
      Left            =   360
      TabIndex        =   9
      Top             =   45
      Width           =   1200
   End
End
Attribute VB_Name = "frmREORDER_CO"
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
Dim rs As ADODB.Recordset
Dim lngBadRows As Long
Dim strType As String
Dim dteSince As Date
Dim WithEvents oPOG As z_GenerateTRs
Attribute oPOG.VB_VarHelpID = -1
Dim bFilterPrint As Boolean
Dim lngStaffID As Long
Dim tlOperators As z_TextList
Dim flgLoading As Boolean
Dim OpenResult As Integer

Public Sub Component(pType As String)
Dim lngLastStaffID As Long

    On Error GoTo errHandler
    flgLoading = True
    strType = pType
    If strType = "CUST" Then
        Me.Caption = "Place purchase orders from customer requests"
        lblStaffMember.Visible = False
        cboStaffMember.Visible = False
        If oPC.Configuration.ReorderPerCOL = True Then
            lblStaffMember.Visible = True
            cboStaffMember.Visible = True
        End If
    Else
        lblStaffMember.Visible = False
        cboStaffMember.Visible = False
        Me.Caption = "Place purchase orders from sales and transfers-out"
    End If
    If Not oPC.Configuration.ReorderPerCOL Then
        Grid1.Splits(0).Columns(17).Visible = False
        Grid1.Splits(0).Columns(18).Visible = False
    End If
    
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Component(pType)", pType
End Sub
Public Sub Component2(pDate As Date)
    On Error GoTo errHandler
    dteSince = pDate
    lblStaffMember.Visible = False
    cboStaffMember.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Component2(pDate)", pDate
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
    lngArrayRows = 0
    rs.CursorLocation = adUseClient

    If strType = "CUST" Then
        If oPC.Configuration.ReorderPerCOL Then
            rs.Open "SELECT * FROM dbo.tREORDERCUSTByCOL WHERE STATUS <> 'X' AND STAFFID = " & lngStaffID & " AND WSNAME = '" & oPC.NameOfPC & "'", oPC.CO, adOpenKeyset, adLockOptimistic
        Else
            rs.Open "SELECT * FROM dbo.tREORDERGENERAL WHERE STATUS <> 'X'" & " AND WSNAME = '" & oPC.NameOfPC & "'", oPC.CO, adOpenKeyset, adLockOptimistic
        End If
    Else
        rs.Open "SELECT * FROM dbo.tREORDERGENERAL WHERE STATUS <> 'X'" & " AND WSNAME = '" & oPC.NameOfPC & "'", oPC.CO, adOpenKeyset, adLockOptimistic
    End If
    If Not rs.eof Then
        Do While Not rs.eof
            lngArrayRows = lngArrayRows + 1
            rs.MoveNext
        Loop
        rs.MoveFirst
    End If

    lngBadRows = 0
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, lngArrayRows, 1, 27
    lngIndex = 1
    Do While Not rs.eof
            XA.Value(lngIndex, 1) = FNS(rs.Fields("PRCODE"))
            XA.Value(lngIndex, 2) = FNS(rs.Fields("DESCRIP"))
         '   If XA.Value(lngIndex, 2) Like "Applied*" Then MsgBox "here"
            XA.Value(lngIndex, 3) = Format(FNN(rs.Fields("PRICE")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
            XA.Value(lngIndex, 4) = FNN(rs.Fields("ONHAND"))
            XA.Value(lngIndex, 5) = FNS(rs.Fields("QTYCO"))
            XA.Value(lngIndex, 6) = FNN(rs.Fields("QTYPO"))
            XA.Value(lngIndex, 7) = FNN(rs.Fields("QTYAPP"))
            
            
            XA.Value(lngIndex, 8) = FNS(rs.Fields("LASTSIXWEEKS"))
            XA.Value(lngIndex, 9) = FNS(rs.Fields("LASTSIXMONTHS"))
            XA.Value(lngIndex, 10) = FNN(rs.Fields("TOTALSOLD"))
            XA.Value(lngIndex, 11) = FNN(rs.Fields("QtyFirm"))
            XA.Value(lngIndex, 12) = FNN(rs.Fields("QtySS"))
            dteTMP = FND(rs.Fields("LASTORDEREDDATE"))
            XA.Value(lngIndex, 13) = IIf(dteTMP > CDate(0), Format(dteTMP, "dd/mm/yyyy"), "")
            dteTMP = FND(rs.Fields("LASTRECEIVEDDATE"))
            XA.Value(lngIndex, 14) = IIf(dteTMP > CDate(0), Format(dteTMP, "dd/mm/yyyy"), "")
            
            XA.Value(lngIndex, 15) = FNS(rs.Fields("PT"))
            
            XA.Value(lngIndex, 16) = FNS(rs.Fields("PUBLISHER"))
            XA.Value(lngIndex, 17) = FNS(rs.Fields("LASTSUPPLIERNAME"))
            XA.Value(lngIndex, 18) = FNS(rs.Fields("LASTDEALNAME"))
            If strType = "CUST" Then
                If oPC.Configuration.ReorderPerCOL Then
                    XA.Value(lngIndex, 19) = FNS(rs.Fields("CODate"))
                End If
                XA.Value(lngIndex, 20) = FNS(rs.Fields("Ref"))
            Else
            End If
            XA.Value(lngIndex, 21) = FNS(rs.Fields("PID"))
            XA.Value(lngIndex, 22) = FNN(rs.Fields("LASTSUPPLIERID"))
            XA.Value(lngIndex, 23) = FNN(rs.Fields("LASTDEALID"))
            XA.Value(lngIndex, 26) = FNS(rs.Fields("COLID"))
            XA.Value(lngIndex, 27) = FNS(rs.Fields("PRICE"))
            XA.Value(lngIndex, 24) = rs.Bookmark
            If (((XA.Value(lngIndex, 11) > 0) Or (XA.Value(lngIndex, 12) > 0)) And ((XA.Value(lngIndex, 23) = 0) Or (XA.Value(lngIndex, 22) = 0))) Then
                XA.Value(lngIndex, 25) = "X"
                lngBadRows = lngBadRows + 1
            End If
            lngIndex = lngIndex + 1
            rs.MoveNext
    Loop
    XA.QuickSort 1, lngArrayRows, 2, XORDER_ASCEND, XTYPE_STRING  ', 4, XORDER_ASCEND, XTYPE_DATE
    Grid1.Array = XA
    Grid1.ReBind
    cmdGenerate.Enabled = (lngBadRows = 0)
    If Not rs.eof Then rs.MoveFirst
    mSetfocus Grid1
'''---------------------------------------------------
''    If OpenResult = 0 Then oPC.DisconnectDBShort
'''---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.LoadGrid"
End Sub



Private Sub cboStaffMember_Change()
    If flgLoading Then Exit Sub
    lngStaffID = oPC.Configuration.Staff.FindStaffByName(cboStaffMember).ID
    LoadGrid
End Sub

Private Sub cboStaffMember_Click()
    If flgLoading Then Exit Sub
    lngStaffID = oPC.Configuration.Staff.FindStaffByName(cboStaffMember).ID
    LoadGrid
End Sub

Private Sub cboStaffMember_LostFocus()
    lngStaffID = oPC.Configuration.Staff.FindStaffByName(cboStaffMember).ID
    SaveSetting "PBKS", "USERS", "REORDER", CStr(lngStaffID)
    LoadGrid
End Sub

Private Sub cboStaffMember_Validate(Cancel As Boolean)
    lngStaffID = oPC.Configuration.Staff.FindStaffByName(cboStaffMember).ID
    SaveSetting "PBKS", "USERS", "REORDER", CStr(lngStaffID)
    LoadGrid
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
    ErrorIn "frmREORDER_CO.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExp_Click()
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim ar As arReorderSlate_ForExcel
Dim arS As arReorderSlateSummary_ForExcel
Dim frm As frmPrintReordersheet
Dim sFile As String
Dim fs As New FileSystemObject

    Set frm = New frmPrintReordersheet
    frm.Show vbModal
    If rs.eof And rs.BOF Then
        Unload frm
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    rs.Sort = frm.Sequence
    rs.MoveFirst
    sFile = oPC.LocalFolder & "\TEMP\ReorderfromSales.XLS"
    If fs.FileExists(sFile) Then
        fs.DeleteFile sFile, True
    End If
    xls.FileName = sFile
    If frm.chkSummary = 0 Then
        Set ar = New arReorderSlate_ForExcel
        ar.Component rs, frm.OrderedOnly
        ar.Printer.Orientation = ddOLandscape
        ar.Run False
        If ar.Pages.Count > 0 Then
            xls.Export ar.Pages
        End If
    Else
        Set arS = New arReorderSlateSummary_ForExcel
        arS.Component rs, frm.OrderedOnly
        arS.Printer.Orientation = ddOLandscape
        arS.Run False
        If arS.Pages.Count > 0 Then
            xls.Export arS.Pages
        End If
    End If
    Screen.MousePointer = vbDefault
    MsgBox "Spreadsheet file saved in: " & sFile, vbInformation, "Export complete"
    Unload frm
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo errHandler
Dim oUtil As New z_UTIL
Dim OpenResult As Integer

    Screen.MousePointer = vbHourglass
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    
    If strType = "CUST" And oPC.Configuration.ReorderPerCOL Then
        If Not oUtil.TableExists("tREORDERCUSTByCOL") Then
            MsgBox "There is no existing Reorder slate", vbInformation, "Can't do this"
        Else
            LoadGrid
        End If
    Else
        If Not oUtil.TableExists("tREORDERGENERAL") Then
            MsgBox "There is no existing Reorder slate", vbInformation, "Can't do this"
        Else
            LoadGrid
        End If
    End If
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
    '--------------
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    Screen.MousePointer = vbDefault
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.cmdLoad_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdGenerate_Click()
    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
Dim i As Long

    If oPC.Configuration.Signtransactions = True Then
        If SecurityControl(enSECURITY_PO_SIGN, , "Close this form and generate purchase orders?", DOCAPPROVAL) = False Then
               Exit Sub
        End If
    Else
        If MsgBox("You want to close this form and generate the purchase orders.", vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If

    Screen.MousePointer = vbHourglass
    oPC.CO.Execute "EXEC DropTableIfExists tCREATEPOS_TEMP,''"
    oPC.CO.Execute "CREATE TABLE tCREATEPOs_TEMP (PID CHAR(40),CODE CHAR(13),COLID INTEGER,REF CHAR(25),TPID INTEGER,DLID INTEGER,QTYFIRM INTEGER,QTYSS INTEGER,PRICE INTEGER)"
    For i = 1 To lngArrayRows
        If XA.Value(i, 11) > 0 Or XA.Value(i, 12) > 0 Then
            oPC.CO.Execute "INSERT INTO tCREATEPOS_TEMP (PID,CODE,COLID,REF,TPID,DLID,QTYFIRM,QTYSS,PRICE) VALUES ('" & XA.Value(i, 21) & "','" & XA.Value(i, 1) & "'," & XA.Value(i, 26) & ",'" & XA.Value(i, 20) & "'," & XA.Value(i, 22) & "," & XA.Value(i, 23) & "," & XA.Value(i, 11) & "," & XA.Value(i, 12) & "," & XA.Value(i, 27) & ")"
        End If
    Next i
    PB1.Visible = True
    Set oPOG = New z_GenerateTRs
    oPOG.GeneratePOs gSTAFFID, strType
    PB1.Visible = False
    Set oPOG = Nothing
   ' Set rs = Nothing
    If strType = "CUST" Then
        oPC.CO.Execute "DELETE FROM tREORDERCUSTByCOL WHERE STAFFID = " & lngStaffID & " AND WSNAME = '" & oPC.NameOfPC & "'"
    Else
        oPC.CO.Execute "DELETE FROM tREORDERGENERAL WHERE WSNAME = '" & oPC.NameOfPC & "'"
    End If
    PB1.Visible = False
    Screen.MousePointer = vbDefault
    MsgBox "Generation of orders complete.", , "Status"
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.cmdGenerate_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Grid1_HeadClick(ByVal ColIndex As Integer)
Static Direction As Variant
  '  If ColIndex > 1 Then Exit Sub
    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
        XA.QuickSort XA.LowerBound(1), XA.UpperBound(1), ColIndex + 1, Direction, XTYPE_STRING ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    Grid1.Refresh
End Sub

Private Sub oPOG_Progress(pCount As Long)
    On Error GoTo errHandler
    PB1.Value = pCount
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.oPOG_Progress(pCount)", pCount, EA_NORERAISE
    HandleError
End Sub
Private Sub oPOG_PBMax(pMax As Long)
    On Error GoTo errHandler
    PB1.Min = 0
    PB1.Max = pMax
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.oPOG_PBMax(pMax)", pMax, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdPrint_Click()
    On Error GoTo errHandler
Dim ar As arReorderSlate
Dim arS As arReorderSlateSummary
Dim frm As frmPrintReordersheet

    Set frm = New frmPrintReordersheet
    frm.Show vbModal
    If rs.eof And rs.BOF Then
        Unload frm
        Exit Sub
    End If
    rs.Sort = frm.Sequence
    rs.MoveFirst
    If frm.chkSummary = 0 Then
        Set ar = New arReorderSlate
        ar.Component rs, frm.OrderedOnly
        ar.Printer.Orientation = ddOLandscape
        ar.Show vbModal
    Else
        Set arS = New arReorderSlateSummary
        arS.Component rs, frm.OrderedOnly
        arS.Printer.Orientation = ddOPortrait
        arS.Show vbModal
    End If
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_Click()
    On Error GoTo errHandler
    If IsNull(Grid1.Bookmark) Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(FNS(XA(Grid1.Bookmark, 1)), ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuReorder   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_MouseDown(Button,Shift,X,Y)", Array(Button, Shift, X, Y), _
         EA_NORERAISE
    HandleError
End Sub
Public Sub RemoveFromList()
    On Error GoTo errHandler
Dim lngCOLID As Long
Dim oSM As New z_StockManager
Dim strPID As String
    If MsgBox("Remove this title from the reorder list?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    lngCOLID = XA.Value(Grid1.Bookmark, 26)
    strPID = XA.Value(Grid1.Bookmark, 21)
    If strType = "CUST" Then
        oPC.CO.Execute "UPDATE tREORDER1 SET Status = 'X' WHERE PID = '" & strPID & "'"
        oSM.MarkCOLsActionedForProduct strPID
    Else
        oPC.CO.Execute "UPDATE tREORDER2 SET Status = 'X' WHERE PID = '" & strPID & "'"
        oSM.MarkProductObsolete strPID
    End If
    
    LoadGrid
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.RemoveFromList"
End Sub
Private Sub cmdRecalc_Click()
    On Error GoTo errHandler
Dim errLoop As ADODB.Error
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim frm As frmSalesSince
Dim OpenResult As Integer

    If MsgBox("This action will erase the current reorder list and prepare a new one, continue?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    If strType <> "CUST" Then
        Set frm = New frmSalesSince
        frm.Component Me
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Exit Sub
        End If
        Unload frm
    End If
    Screen.MousePointer = vbHourglass
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 300
    Set prm = cmd.CreateParameter("@WSNAME", adVarChar, adParamInput, 50, oPC.NameOfPC)
    cmd.Parameters.Append prm
    
    If strType = "CUST" Then
        If oPC.Configuration.ReorderPerCOL Then
            cmd.CommandText = "REORDERCUST_ByCOL"
            Set prm = cmd.CreateParameter("@StaffID", adInteger, adParamInput, , lngStaffID)
            cmd.Parameters.Append prm
        Else
            cmd.CommandText = "REORDERCUST"
            Set prm = cmd.CreateParameter("@pDate", adDate, adParamInput, , dteSince)
            cmd.Parameters.Append prm
        End If
    Else
        cmd.CommandText = "REORDERSALES"
        Set prm = cmd.CreateParameter("@pDate", adDate, adParamInput, , dteSince)
        cmd.Parameters.Append prm
    End If
    
    
    cmd.Execute
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
    '--------------
    LoadGrid
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.cmdRecalc_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdReset_Click()
    On Error GoTo errHandler
Dim i As Long
    If MsgBox("You want to reset all the firm and seesafe quantities to zero?", vbYesNo + vbQuestion, "Confirm") = vbNo Then
        Exit Sub
    End If
    For i = 1 To XA.UpperBound(1)
        XA.Value(i, 11) = 0
        XA.Value(i, 12) = 0
    Next
    Grid1.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.cmdReset_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub SetDeal()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseSUppliers2
Dim frm2 As frmSUPPDEAL
Dim oSupp As a_Supplier
Dim oDeal As a_Deal

    If IsNull(Grid1.Bookmark) Then Exit Sub
    If (FNN(XA.Value(Grid1.Bookmark, 22) <> 0) And FNN(XA.Value(Grid1.Bookmark, 23) <> 0)) Then
        If MsgBox("You are wanting to change the supplier or deal where they are already noted?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If

    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    If lngTPID = 0 Then
        Exit Sub
    Else
        Set frm2 = New frmSUPPDEAL
        Set oSupp = New a_Supplier
        oSupp.Load frm.SupplierID
        frm2.Component oSupp
        frm2.Show vbModal
    End If
    Set oDeal = frm2.SelectedDeal
    Unload frm
    Unload frm2
    If oDeal Is Nothing Then
        MsgBox "There is no deal selected for this supplier. You must select a deal before an order can be generated for this product", vbOKOnly + vbInformation, "Warning"
        Exit Sub
    End If
    XA.Value(Grid1.Bookmark, 22) = oSupp.ID
    XA.Value(Grid1.Bookmark, 23) = oDeal.ID
    XA.Value(Grid1.Bookmark, 17) = oSupp.NameAndCode(25)
    XA.Value(Grid1.Bookmark, 18) = oDeal.Description & " (" & oDeal.DiscountF & ")"
    rs.Bookmark = XA.Value(Grid1.Bookmark, 24)
    rs.Fields("LASTSUPPLIERID") = oSupp.ID
    rs.Fields("LASTDEALID") = oDeal.ID
    rs.Fields("LASTSUPPLIERNAME") = oSupp.NameAndCode(35)
    rs.Fields("LASTDEALNAME") = left(oDeal.Description, 26) & " (" & oDeal.DiscountF & ")"
    rs.Update
    Grid1.Refresh
    Set oSupp = Nothing
    Set oDeal = Nothing
    ValidateRow (XA.Value(Grid1.Bookmark, 25) <> "X")
    cmdGenerate.Enabled = (lngBadRows = 0)
    mSetfocus Grid1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.SetDeal"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set rs = Nothing
'---------------------------------------------------
    oPC.CloseDB
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim strPID As String
Dim frm As frmProductPrev
Dim oProd As a_Product
    Screen.MousePointer = vbHourglass
    strPID = XA.Value(Grid1.Bookmark, 21)
    If strPID > "" Then
        Set oProd = New a_Product
        oProd.Load strPID, 0
        Set frm = New frmProductPrev
        frm.Component oProd
        Screen.MousePointer = vbDefault
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_SelChange(Cancel As Integer)
Dim str As String
    If IsNull(Grid1.Bookmark) Then Exit Sub
    str = FNS(XA.Value(Grid1.Bookmark, 1))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
Dim lngSupplierID As Long
Dim lngDEALID As Long

    lngSupplierID = FNN(XA(Bookmark, 22))
    lngDEALID = FNN(XA(Bookmark, 23))
    If lngSupplierID = 0 Or lngDEALID = 0 Then
        RowStyle.BackColor = RGB(232, 174, 180)
    End If
    If (lngSupplierID = 0 Or lngDEALID = 0) And (FNN(XA.Value(Bookmark, 11)) > 0 Or FNN(XA.Value(Bookmark, 12)) > 0) Then
        RowStyle.BackColor = RGB(232, 174, 220)
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
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

    If rs.eof And rs.BOF Then Exit Sub
    If Not ConvertToLng(Grid1.Text, lngTmp) Then
        Cancel = True
        Exit Sub
    End If
    XA.Value(Grid1.Bookmark, ColIndex + 1) = FNN(Grid1.Text)
    rs.Bookmark = XA.Value(Grid1.Bookmark, 24)
    Select Case ColIndex
    Case 10
            rs.Fields("QTYFIRM") = FNN(Grid1.Text)
            ValidateRow (XA.Value(Grid1.Bookmark, 25) <> "X")
    Case 11
            rs.Fields("QTYSS") = FNN(Grid1.Text)
            ValidateRow (XA.Value(Grid1.Bookmark, 25) <> "X")
    End Select
    rs.Update
    cmdGenerate.Enabled = (lngBadRows = 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If rs.eof And rs.BOF Then Exit Sub
    If IsNull(Grid1.Bookmark) Then Exit Sub
    rs.Bookmark = XA.Value(Grid1.Bookmark, 24)
    If ColIndex = 10 Then
        rs.Fields("QTYFIRM") = FNN(XA.Value(Grid1.Bookmark, 11))
        rs.Update
    ElseIf ColIndex = 11 Then
        rs.Fields("QTYSS") = FNN(XA.Value(Grid1.Bookmark, 12))
        rs.Update
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
Dim lngLastStaffID As Long
    On Error GoTo errHandler
'-------------------------------
    OpenResult = oPC.OpenDBSHort
    oPC.OpenDB ""
'-------------------------------
    flgLoading = True
    Set tlOperators = New z_TextList
    tlOperators.Load ltStaff
    If oPC.Configuration.ReorderPerCOL Then
        LoadCombo cboStaffMember, tlOperators
        lngLastStaffID = CLng(GetSetting("PBKS", "USERS", "REORDER", "0"))
        If lngLastStaffID > 0 Then
            If Not oPC.Configuration.Staff.FindStaffByID(lngLastStaffID) Is Nothing Then
                cboStaffMember = oPC.Configuration.Staff.FindStaffByID(lngLastStaffID).StaffName
                lngStaffID = lngLastStaffID
            Else
                lngStaffID = tlOperators.Key(cboStaffMember)
            End If
        Else
            lngStaffID = tlOperators.Key(cboStaffMember)
        End If
    Else
        cboStaffMember.Visible = False
        lblStaffMember.Visible = False
    End If
    Me.Width = 11500
    Me.Height = 6500
    Me.left = 100
    Me.top = 100
    bFilterPrint = False
    flgLoading = False
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
    '--------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_LostFocus()
    On Error GoTo errHandler
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.Grid1_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub ValidateRow(pOKAtPresent As Boolean)
    On Error GoTo errHandler
    If IsNull(Grid1.Bookmark) Then Exit Sub
    If ((XA.Value(Grid1.Bookmark, 11) > 0) Or (XA.Value(Grid1.Bookmark, 12) > 0)) And ((XA.Value(Grid1.Bookmark, 22) = 0) Or (XA.Value(Grid1.Bookmark, 23) = 0)) Then
        lngBadRows = lngBadRows + 1
        XA.Value(Grid1.Bookmark, 25) = "X"
    Else
        If Not pOKAtPresent Then
            lngBadRows = lngBadRows - 1
        End If
        XA.Value(Grid1.Bookmark, 25) = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmREORDER_CO.ValidateRow(pOKAtPresent)", pOKAtPresent
End Sub
Public Sub ShowSalesPatterns()
Dim frmSales As frmSalesCH
Dim oProduct As a_Product
Dim strPID As String
    Screen.MousePointer = vbHourglass
    Set oProduct = New a_Product
    strPID = XA.Value(Grid1.Bookmark, 21)
    If strPID = "" Then Exit Sub

    oProduct.Load strPID, 0
    If oProduct.pID = "" Then Exit Sub
    Set frmSales = New frmSalesCH
    frmSales.Component oProduct
    frmSales.Show
    Screen.MousePointer = vbDefault
    Set frmSales = Nothing
End Sub

