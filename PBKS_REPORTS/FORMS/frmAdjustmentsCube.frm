VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{801C12A5-BE41-41CD-AE48-C666E77F2F02}#2.0#0"; "CCubeX20.ocx"
Begin VB.Form frmAdjustmentsCube 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock adjustments "
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11805
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
   Begin CCubeX2.ContourCubeX cc 
      Height          =   4950
      Left            =   195
      TabIndex        =   0
      Top             =   765
      Width           =   11220
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   13882315
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   ""
      InactiveDimAreaBkColor=   14215660
      AutoSize        =   0   'False
      UnusedDataAreaColor=   -2147483643
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmAdjustmentsCube.frx":0000
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Picture         =   "frmAdjustmentsCube.frx":0468
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   75
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10380
      Picture         =   "frmAdjustmentsCube.frx":07F2
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   165
      Picture         =   "frmAdjustmentsCube.frx":0B7C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5850
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpSince 
      Height          =   375
      Left            =   1290
      TabIndex        =   4
      Top             =   210
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      Format          =   126418945
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpUntil 
      Height          =   375
      Left            =   3345
      TabIndex        =   6
      Top             =   210
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      Format          =   126418945
      CurrentDate     =   37421
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "until"
      ForeColor       =   &H00915A48&
      Height          =   300
      Left            =   2985
      TabIndex        =   7
      Top             =   300
      Width           =   300
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "Since"
      ForeColor       =   &H00915A48&
      Height          =   300
      Left            =   585
      TabIndex        =   5
      Top             =   300
      Width           =   450
   End
End
Attribute VB_Name = "frmAdjustmentsCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTRID As Long
Dim lngCashBookLineID As Long
Dim lngAmt As Long
Dim dteDate As Date
Dim strReason As String
Dim bAmt As Boolean
Dim bDate As Boolean
Dim bReason As Boolean
Dim strCustomerName As String
Dim strInvoices As String
Dim lngInvoiceID As Long
Dim XA As New XArrayDB
Dim x As New XArrayDB
Dim rs As New ADODB.Recordset
Dim tlChildCustomers As z_TextList
Dim flgLoading As Boolean
Dim bDirty As Boolean

Private Sub cc_KeyUp(ByVal KeyCode As Long, ByVal Shift As Long)
  Dim Col As Long
  Dim Row As Long
  With cc
'   Press "C" key for sorting current column
    If KeyCode = vbKeyC And .VAxis.Dims.Count > 0 Then
      With .CurrentCell
         cc.SortGridByFact xda_vertical, .Col, .Row
      End With
'   Press "R" key for sorting current row
    ElseIf KeyCode = vbKeyR And .HAxis.Dims.Count > 0 Then
      With .CurrentCell
         cc.SortGridByFact xda_horizontal, .Col, .Row
      End With
'   Press "A" key for abandon sorting
    ElseIf KeyCode = vbKeyA Then
      .CancelFactSorting
    End If
  End With
End Sub

Private Sub cmdClose_Click()
Dim bInProcess As Boolean

    Unload Me
End Sub

Private Sub cmdOpen_Click()
Dim i As Integer
Dim oSQL As New z_SQL
    Screen.MousePointer = vbHourglass
On Error Resume Next

rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

      '  MsgBox "Select * FROM ahv_AllStockAdjustments WHERE dte >= '" & Format(dtpSince, "yyyy-mm-dd") & "' and dte <= '" & Format(dtpUntil, "yyyy-mm-dd") & "'"
        oSQL.GetDynamicRecordset_Improved "Select * FROM ahv_AllStockAdjustments WHERE dte >= '" & Format(dtpSince, "yyyy-mm-dd") & "' and dte <= '" & Format(dtpUntil, "yyyy-mm-dd") & "'", enText, Array(), "", rs
        Preparecube
        LoadContourcubeLayout oPC.LocalFolder & "Templates\AdjustmentsLayout.txt", cc
       
        Screen.MousePointer = vbDefault

End Sub

Private Sub cmdExport_Click()
Dim fs As New FileSystemObject
Dim F As String
    F = oPC.SharedFolderRoot & "\HTML\Adjustments.xls"
   ' f = "c:\TEST.xls"
    If fs.FileExists(F) Then fs.DeleteFile F, True
    If fs.FileExists(F) Then
        MsgBox "The file '" & F & "' cannot be cleared, so it cannot be re-generated." & vbCrLf & "Check that no application (e.g. Excel or Openoffice) is holding it open.", vbInformation + vbOKOnly, "Can't export file"
    Else
        If cc.RowCount > 0 Then
            cc.ReportToFile F, "", xolaprpt_XLS
            OpenFileWithApplication F, enExcel
        End If
    End If
End Sub





Private Sub Form_Activate()
Dim oFSo As New FileSystemObject

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    flgLoading = True
    
    bAmt = False
    bDate = False
    If Me.WindowState <> 2 Then
        Left = 70
        top = 70
        Width = 3990
        Height = 4620
    End If
' Allocate space for 300 rows, 4 columns
 '   XA.ReDim 0, 299, 0, 7

    Dim Row As Long, Col As Integer
    Me.dtpSince = DateAdd("m", -6, Date)
    Me.dtpUntil = DateAdd("d", 1, Date)
    
    SetFormSize Me
   ' Preparecube
    cc.Header.text = "Press C to sort current column, R to sort current row, or A to cancel sorting"
    cc.Header.Visible = True
    bDirty = False
    flgLoading = False
    
    

    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Form_Load", , EA_NORERAISE
    HandleError
End Sub



Private Sub Preparecube()
    On Error GoTo errHandler
Dim oFSo As New FileSystemObject
Dim oTLS As New z_TextListSimple
Dim Fact As IViewFact
    
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then
        CloseCube
        MsgBox "No records", , "Status"
        Exit Sub
    End If
    rs.MoveFirst
    If rs.EOF Then
        CloseCube
        MsgBox "No records", , "Status"
    End If
    
    If Not rs.EOF Then
        
        CloseCube
        With cc.Cube
            .Dims.Add("AdjustmentType", "AdjustmentType", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("ProdCombo", "ProdCombo", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("dte", "dte", , xda_vertical).MoveTo xda_vertical
            .Dims.Add("DocCode", "DocCode", , xda_vertical).MoveTo xda_vertical
            .BaseFacts.Add "QtyPre", "QtyPre"
            .Facts.Add "QtyPre", "QtyPre", xfaa_SUM
            .BaseFacts.Add "QtyDifference", "QtyDifference"
            .Facts.Add "QtyDifference", "QtyDifference", xfaa_SUM
            .BaseFacts.Add "Cost", "Cost"
            .Facts.Add "Cost", "Cost", xfaa_SUM
            .BaseFacts.Add "SP", "SP"
            .Facts.Add "SP", "SP", xfaa_SUM
            cc.Facts(0).Appearance.Format = "###,##0;-###,##0"
            cc.Facts(0).Caption = "Qty pre count."
            cc.Facts(1).Appearance.Format = "###,##0;-###,##0"
            cc.Facts(1).Caption = "Adjustment"
            cc.Facts(2).Appearance.Format = "###,##0.00;(###,##0.00)"
            cc.Facts(2).Caption = "Cost"
            cc.Facts(3).Appearance.Format = "###,##0.00;(###,##0.00)"
            cc.Facts(3).Caption = "Price"
            cc.NoGrandTotals = False
           ' CC.Dims(0).NoTotals = True
           ' CC.Dims(1).NoTotals = True
            cc.TitleSettings.text = "Adjustments summary"
            cc.VAxis.DrillDownLevel = 0
            For Each Fact In cc.Facts
              Fact.Visible = True
            Next
            Set rs.ActiveConnection = Nothing
            Screen.MousePointer = vbHourglass
            If Not rs.EOF Then
                .Open rs
                cc.Active = True
                cc.Visible = cc.Active
            Else
                cc.Active = False
                cc.Visible = cc.Active
            End If
        End With
        Me.Refresh
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformance.Preparecube"
    HandleError
End Sub

'Private Sub cmdFind1_Click()
'    On Error GoTo errHandler
'    Screen.MousePointer = vbHourglass
'    Find
'    Grid.ReBind
'    Grid.Bookmark = 1
'
'    Screen.MousePointer = vbDefault
'    Exit Sub
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmCRemittancePreview.cmdFind1_Click", , EA_NORERAISE
'    HandleError
'End Sub
'Private Sub Find()
'    On Error GoTo errHandler
'Dim bNotFound As Boolean
'Dim frm As frmBrowseCustomers2
'Dim lngTRID As Long
'Dim byear As Boolean
'Dim yr As String
'Dim mth As String
'Dim strDate1 As String
'Dim strDate2 As String
'Dim lngCount As Long
'
'    bNotFound = False
'    If Left(txtArg, 3) = "yr=" Then byear = True
'
'    If txtArg > " " And Not (byear) Then
'        'Search for Reference
'        Set cJNL = Nothing
'        Set cJNL = New c_JNL
'        cJNL.Load bNotFound, 0, "", txtArg, dteDate1, dteDate2
'        If bNotFound Then
'            'Search for customer by AcJNLO
'            Set cJNL = Nothing
'            Set cJNL = New c_JNL
'            SetDateArgs
'            cJNL.Load bNotFound, 0, txtArg, "", dteDate1, dteDate2
'            If bNotFound Then
'               Set frm = New frmBrowseCustomers2
'               frm.component txtArg, lngCount
'               If lngCount > 1 Then
'                    frm.Show vbModal
'                    lngTRID = frm.CustomerID
'                    Unload frm
'                ElseIf lngCount = 1 Then
'                    lngTRID = frm.CustomerID
'                    Unload frm
'                End If
'               If lngTRID > 0 Then
'                    Set cJNL = Nothing
'                    Set cJNL = New c_JNL
'                    SetDateArgs
'                    cJNL.Load bNotFound, lngTRID, "", "", dteDate1, dteDate2
'               End If
'            End If
'        Else
'            enSince = 1
'            cbSince.Caption = TranslateSince(1)
'        End If
'    Else
'        If byear Then
'            yr = Mid(txtArg, 4, 4)
'            mth = Mid(txtArg, 9, 2)
'            If mth > "" Then
'                strDate1 = yr & "-" & mth & "-01"
'                strDate2 = yr & "-" & mth & "-" & LastDayOfMonth(yr & "-" & mth & "-01")
'            Else
'                strDate1 = yr & "-01-01"
'                strDate2 = yr & "-12-31"
'            End If
'            If Not (IsDate(strDate1) And IsDate(strDate2)) Then
'                SetDateArgs
'            Else
'                dteDate1 = CDate(strDate1)
'                dteDate2 = CDate(strDate2)
'            End If
'        Else
'            SetDateArgs
'        End If
'        cJNL.Load bNotFound, 0, "", "", dteDate1, dteDate2
'    End If
'
'EXIT_Handler:
'    mSetfocus Grid
'    MousePointer = vbDefault
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmBrowseDBJNLs.Find"
'End Sub

Private Sub CloseCube()
    On Error GoTo errHandler
 With cc
   .Active = False
   .Cube.Dims.Clear
   .Cube.Facts.Clear
   .Cube.BaseFacts.Clear
 End With
' CheckEnabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CloseCube"
End Sub
Private Sub AfterOpen()
    On Error GoTo errHandler
 cc.Visible = cc.Active
 CheckVisible
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.AfterOpen"
End Sub

Private Sub CheckVisible()
    On Error GoTo errHandler
    cc.Visible = True 'cc.Active
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesPT.CheckVisible"
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    If flgLoading Then Exit Sub
    
    cc.Width = NonNegative_Lng(Me.Width - 700)
    cc.Height = NonNegative_Lng(Me.Height - 2100)

    cmdClose.top = NonNegative_Lng(Me.Height - 1300)
    cmdClose.Left = NonNegative_Lng(Me.Width - 1900)
    cmdExport.top = cmdClose.top
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustPmt.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cc.Dims.Count > 0 Then
        SaveContourCubeLayout CStr(oPC.LocalFolder & "Templates\AdjustmentsLayout.txt"), Me.cc
    End If

    SaveFormSize Me.Name, Me.Height, Me.Width
    
End Sub

