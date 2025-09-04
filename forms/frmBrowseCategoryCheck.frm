VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmCategoryChecks 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse category checks"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   17880
   Begin VB.CommandButton cmdOperatorSignoff 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Operator sign-off"
      Height          =   570
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8685
      Width           =   2175
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Issue"
      Height          =   570
      Left            =   2400
      Picture         =   "frmBrowseCategoryCheck.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8685
      Width           =   1290
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   135
      TabIndex        =   1
      Top             =   405
      Width           =   17430
      _ExtentX        =   30745
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   13882315
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Category check list"
      TabPicture(0)   =   "frmBrowseCategoryCheck.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "arViewer"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdToPDF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdToExcel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Category check corrections"
      TabPicture(1)   =   "frmBrowseCategoryCheck.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid1"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdToExcel 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Spreadsheet"
         Height          =   360
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1380
      End
      Begin VB.CommandButton cmdToPDF 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1380
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
         Height          =   6150
         Left            =   165
         TabIndex        =   2
         Top             =   735
         Width           =   16935
         _ExtentX        =   29871
         _ExtentY        =   10848
         SectionData     =   "frmBrowseCategoryCheck.frx":03C2
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   7440
         Left            =   -74850
         OleObjectBlob   =   "frmBrowseCategoryCheck.frx":03FE
         TabIndex        =   3
         Top             =   420
         Width           =   16950
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   16575
      Picture         =   "frmBrowseCategoryCheck.frx":5CF4
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8670
      Width           =   1000
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   270
      Left            =   14805
      TabIndex        =   6
      Top             =   30
      Width           =   2655
   End
End
Attribute VB_Name = "frmCategoryChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oSM As New z_StockManager
Dim oSQL As z_SQL
Dim rpt As arCategoryCheck
Dim lngCatChkID As Long
Dim strHeading As String
Dim XA As New XArrayDB
Dim lngStaffID As Long
Dim lngSupervisorID As Long
Dim dteUpdatedDate As Date
Dim bLoaded As Boolean
Dim Status As Long

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuVoid.Enabled = (Status = 2) Or (Status = 3)
    Forms(0).mnuCancel.Enabled = False
    Forms(0).mnuCancelLine.Enabled = False
    Forms(0).mnuCancelINactive.Enabled = False
    Forms(0).mnuFulfil.Enabled = False
    Forms(0).mnuDelLine.Enabled = False
    Forms(0).mnuMemo.Enabled = False
    Forms(0).mnuSalesComm.Enabled = False
    Forms(0).mnuCopyDoc.Enabled = False
    Forms(0).mnuSaveColumnWidths.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.SetMenu"
End Sub
Public Sub mnuVoid()
    On Error GoTo errHandler
 '   If Not ((((FNN(rs.fields("CATCHK_Status"))) = 2) Or (((FNN(rs.fields("CATCHK_Status"))) = 3)))) Then Exit Sub
    If (FNN(rs.Fields("CATCHK_Status")) = 4) Then Exit Sub
        If oPC.Configuration.SignTransactions = True Then
            If SecurityControl(enSECURITY_ISOPERATOR, , "Void this category check", DOCAPPROVAL) = False Then
                   Exit Sub
            End If

        End If
        
    If oSQL Is Nothing Then Set oSQL = New z_SQL
    oSQL.RunSQL "UPDATE tCATCHK SET CATCHK_Status = 1 WHERE CATCHK_ID = " & CStr(lngCatChkID)
    Me.lblStatus = "Voided"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOPreview.mnuVoid"
End Sub
Public Sub component(pCATCHKID As Long, pLabel As String, SignedOffName As String, SignedOffBy As Long, pEmpty As Boolean)
    On Error GoTo errHandler
    lngCatChkID = pCATCHKID
      Set oSQL = New z_SQL
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oSQL.CategoryCheck rs, lngCatChkID
    If rs.eof Then
      pEmpty = True
      Exit Sub
    End If
    Status = FNN(rs.Fields("CATCHK_Status"))

    strHeading = pLabel
    Me.Caption = strHeading
  '  Me.cmdUpdate.Enabled = (SignedOffBy = 0 and )
 '  Me.cmdOperatorSignoff.Enabled = (SignedOffBy = 0)
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.component(pCATCHKID,pLabel,SignedOffName,SignedOffBy)", _
         Array(pCATCHKID, pLabel, SignedOffName, SignedOffBy)
End Sub

Private Sub cmdOperatorSignoff_Click()
    On Error GoTo errHandler
Dim i As Long
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim bCancelled As Boolean
Dim bIsSUpervisor As Boolean
Dim strName As String
Dim lngSMID As Long
Dim Strguid As String

    Me.Grid1.Update
    If SecurityControl(enSECURITY_ISOPERATOR, bCancelled, "Enter your signature", "You do not have permission to sign-off the capture of a category check (or your signature is invalid)", bIsSUpervisor, strName, lngSMID) = True Then
        If oSQL Is Nothing Then Set oSQL = New z_SQL
        oSQL.RunSQL "UPDATE tCATCHK SET CATCHK_STAFF_ID = " & CStr(lngSMID) & " WHERE CATCHK_ID = " & CStr(lngCatChkID)
        lngStaffID = lngSMID
    End If
            Set xMLDoc = New ujXML
            With xMLDoc
                .docProgID = "MSXML2.DOMDocument"
                .docInit "doc_CatChk"
                    .chCreate "MessageType"
                        .elText = "doc_CatChk"
                    .elCreateSibling "MessageCreationDate"
                        .elText = Format(Now(), "yyyymmddHHNN")
                    .elCreateSibling "CatCHkID"
                        .elText = CStr(lngCatChkID)
                    .elCreateSibling "StaffID"
                        .elText = CStr(lngStaffID)
                    .elCreateSibling "SupervisorID"
                        .elText = CStr(lngSupervisorID)
                    .elCreateSibling "Status"
                        .elText = CStr("3")
                    .elCreateSibling "DetailLines", True
                    For i = 1 To XA.UpperBound(1)
                    If (FNN(XA(i, 6)) <> FNN(XA(i, 4))) And (Not IsEmpty(XA(i, 6))) Then
                            .chCreate "ITEM"
                            .chCreate "PID"
                                .elText = FNS(XA(i, 15))
                            .elCreateSibling "CATCHKLID", True
                                .elText = CStr(FNN(XA(i, 13)))
                            .elCreateSibling "COUNT", True
                                .elText = CStr((XA(i, 14)))
                            .elCreateSibling "DIFF", True
                                .elText = CStr(FNN(XA(i, 4)) - FNN(XA(i, 14)))
                            .navUP
                            .navUP
                        End If
                    Next i
        
                 XMLArgs = .docXML
          
            End With
            
            If XMLArgs > "" Then
                oSM.InsertScript Strguid, XMLArgs
                If Strguid > "" Then
                    oSM.UpdateCategoryCheck Strguid
                End If
            End If
        Me.lblStatus = "In process"
        
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.cmdOperatorSignoff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo errHandler
Dim i As Long
Dim xMLDoc As ujXML
Dim XMLArgs As String
Dim bCancelled As Boolean
Dim bIsSUpervisor As Boolean
Dim strName As String
Dim lngSMID As Long
Dim Strguid As String

    If SecurityControl(enSECURITY_STKADJ_SIGN, bCancelled, "Enter your signature", "You do not have permission to update a category check (or your signature is invalid)", bIsSUpervisor, strName, lngSMID) = True Then
            lngSupervisorID = lngSMID
            Set xMLDoc = New ujXML
            With xMLDoc
                .docProgID = "MSXML2.DOMDocument"
                .docInit "doc_CatChk"
                    .chCreate "MessageType"
                        .elText = "doc_CatChk"
                    .elCreateSibling "MessageCreationDate"
                        .elText = Format(Now(), "yyyymmddHHNN")
                    .elCreateSibling "CatCHkID"
                        .elText = CStr(lngCatChkID)
                    .elCreateSibling "StaffID"
                        .elText = CStr(lngStaffID)
                    .elCreateSibling "SupervisorID"
                        .elText = CStr(lngSupervisorID)
                    .elCreateSibling "Status"
                        .elText = CStr("4")
                    .elCreateSibling "DetailLines", True
                    For i = 1 To XA.UpperBound(1)
                    If (FNN(XA(i, 14)) <> FNN(XA(i, 4))) And (Not IsEmpty(XA(i, 14))) Then
                            .chCreate "ITEM"
                            .chCreate "PID"
                                .elText = FNS(XA(i, 15))
                            .elCreateSibling "CATCHKLID", True
                                .elText = CStr(FNN(XA(i, 13)))
                            .elCreateSibling "COUNT", True
                                .elText = CStr(FNN(XA(i, 14)))
                            .elCreateSibling "DIFF", True
                                .elText = CStr(FNN(XA(i, 4)) - FNN(XA(i, 14)))
                            .navUP
                            .navUP
                        End If
                    Next i
        
                 XMLArgs = .docXML
          
            End With
            
            If XMLArgs > "" Then
                oSM.InsertScript Strguid, XMLArgs
                If Strguid > "" Then
                    oSM.UpdateCategoryCheck Strguid
                End If
            End If
    End If
    Me.lblStatus = "Issued"
    Me.cmdOperatorSignoff.Enabled = False
    Me.cmdUpdate.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.cmdUpdate_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
    bLoaded = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Me.Left = 300
    Me.TOP = 500
    SetMenu
    SSTab1.Tab = 0
    SetFormSize Me
    Set rpt = Nothing
    Set oSQL = New z_SQL
        
    LoadGrid

    If rs.eof = False Then
        lngStaffID = FNN(rs.Fields("CATCHK_STAFF_ID"))
        lngSupervisorID = FNN(rs.Fields("CATCHK_Supervisor_ID"))
        dteUpdatedDate = FND(rs.Fields("CATCHK_UPDATEDDATE"))
        ShowStatus (FNN(rs.Fields("CATCHK_Status")))
        Set rpt = New arCategoryCheck
        rpt.Printer.Orientation = ddOLandscape
        rpt.PageSettings.PaperSize = 9
        rpt.PageSettings.LeftMargin = 700
        rpt.PageSettings.RightMargin = 0
        rpt.PageSettings.TopMargin = 600
        rpt.PageSettings.BottomMargin = 700
        rpt.component strHeading, rs
    
        Me.arViewer.ReportSource = rpt

    Else
        MsgBox "No records", vbInformation, "Status"
    End If
    cmdUpdate.Enabled = ((FNN(rs.Fields("CATCHK_Status"))) = 3)
    cmdOperatorSignoff.Enabled = ((FNN(rs.Fields("CATCHK_Status"))) = 2)
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub ShowStatus(StatusCode As Long)
    On Error GoTo errHandler
    If StatusCode = 2 Then
        Me.lblStatus.Caption = "In process"
    Else
        If StatusCode = 3 Then
            Me.lblStatus.Caption = "Captured"
        Else
            If StatusCode = 4 Then
                Me.lblStatus.Caption = "Issued"
            Else
                If StatusCode = 1 Then
                    Me.lblStatus.Caption = "Voided"
                Else
                    Me.lblStatus.Caption = "Unknown"
                End If
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.ShowStatus(StatusCode)", StatusCode
End Sub
Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
Dim lngDiffH As Long
    SSTab1.Width = NonNegative_Lng(Me.Width - 400)
    lngDiff = SSTab1.Height
    SSTab1.Height = NonNegative_Lng(Me.Height - (SSTab1.TOP + 1200))
    lngDiff = (SSTab1.Height - lngDiff)
    arViewer.Width = NonNegative_Lng(SSTab1.Width - 500)
    arViewer.Height = NonNegative_Lng(SSTab1.Height - 500)
    Grid1.Height = NonNegative_Lng(SSTab1.Height - 1200)
    Grid1.Width = NonNegative_Lng(SSTab1.Width - 400)
    Me.cmdOperatorSignoff.Left = NonNegative_Lng(SSTab1.Left + 1000)
    Me.cmdUpdate.Left = NonNegative_Lng(SSTab1.Left + 3100)
    Me.cmdOperatorSignoff.TOP = NonNegative_Lng(Me.Height - 1150)
    Me.cmdUpdate.TOP = NonNegative_Lng(Me.Height - 1150)
    Me.lblStatus.Left = NonNegative_Lng(Me.Width - 3000)
    Me.cmdClose.TOP = NonNegative_Lng(Me.Height - 1150)
    Me.cmdClose.Left = NonNegative_Lng(Me.Width - 1280)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Form_Resize", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdToPDF_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    rpt.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "CategoryChecks" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.cmdToPDF_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    rpt.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "CategoryChecks" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.cmdToExcel_Click", , EA_NORERAISE
    HandleError
End Sub
Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Nothing, Me.Name, Me.Height, Me.Width
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.mnuSaveLayout"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Form_Unload(Cancel)", Cancel, EA_NORERAISE
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
Dim dODPO As d_POLine
Dim lngArrayRows As Long
    
    rs.MoveFirst
    
    If bLoaded Then Exit Sub
    i = 0
    Set XA = New XArrayDB
    XA.Clear
    lngIndex = 1
    lngArrayRows = rs.RecordCount
    XA.ReDim 1, lngArrayRows, 1, 15
    For i = 1 To Grid1.Columns.Count
        If i <> 9 And i <> 10 And i <> 11 Then
            Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
        End If
    Next
    rs.MoveFirst
    Do While Not rs.eof
            XA.Value(lngIndex, 1) = FNS(rs.Fields("EAN"))
            XA.Value(lngIndex, 2) = FNS(rs.Fields("ProductDescription"))
            XA.Value(lngIndex, 3) = FNS(rs.Fields("Author"))
            XA.Value(lngIndex, 4) = FNS(rs.Fields("CATCHKL_SystemQty"))
            If FNS(rs.Fields("P_QtyOnHand")) <> FNS(rs.Fields("CATCHKL_SystemQty")) Then
                XA.Value(lngIndex, 5) = FNS(rs.Fields("P_QtyOnHand"))
            End If
            If FNS(rs.Fields("CATCHKL_Counted")) <> FNS(rs.Fields("CATCHKL_SystemQty")) Then
                XA.Value(lngIndex, 6) = FNS(rs.Fields("CATCHKL_Counted"))
            End If
            XA.Value(lngIndex, 13) = FNS(rs.Fields("CATCHKL_ID"))
            XA.Value(lngIndex, 14) = FNS(rs.Fields("CATCHKL_Counted"))
            XA.Value(lngIndex, 15) = FNS(rs.Fields("P_ID"))
            rs.MoveNext
            lngIndex = lngIndex + 1
    Loop
   ' XA.QuickSort 1, lngArrayRows, 1, XORDER_ASCEND, XTYPE_STRING, 5, XORDER_ASCEND, XTYPE_DATE, 3, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XA
    Grid1.ReBind
    bLoaded = True
    rs.MoveFirst
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.LoadGrid"
End Sub


Private Sub Grid1_DblClick()
    On Error GoTo errHandler
Dim frmMM As frmMovements
Dim oProd As New a_Product
Dim x As Long
Dim Y As Long

    If IsNull(Grid1.Bookmark) Then Exit Sub
    oProd.Load XA(Grid1.Bookmark, 15), 0
    oProd.ReloadRecentMovements
    Set frmMM = New frmMovements
   ' frmMM.Component oProd, Me.top + 200, Me.Left + 1000
    If PointsToMe(Me.hWnd, x, Y) Then
        frmMM.component oProd, x, Y
    Else
        frmMM.component oProd, 0, 0
    End If
    frmMM.Show

    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmCategoryChecks: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmCategoryChecks: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    If SSTab1.Tab = 1 Then
  '      LoadGrid
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.SSTab1_Click(PreviousTab)", PreviousTab, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_AfterColUpdate(ByVal ColIndex As Integer)
    On Error GoTo errHandler
    If ColIndex = 5 Then
        If Grid1.text <> "" Then
            If (Not IsNumeric(Grid1.text)) Then Exit Sub
            XA(Grid1.Bookmark, 14) = FNN(Grid1.text)
        Else
             XA(Grid1.Bookmark, 14) = ""
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Grid1_AfterColUpdate(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeUpdate(Cancel As Integer)
    On Error GoTo errHandler
    If Grid1.text = "" Then Exit Sub
    If (Not IsNumeric(Grid1.text)) Then
        Cancel = 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCategoryChecks.Grid1_BeforeUpdate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_Error(ByVal DataError As Integer, Response As Integer)
    MsgBox "Invalid value in cell, values must be numeric or cell must be empty.", vbOKOnly + vbInformation, "Error"
    Response = 0
End Sub


