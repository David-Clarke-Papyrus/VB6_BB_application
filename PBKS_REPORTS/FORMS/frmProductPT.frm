VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmProductPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Documents per supplier"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin CCubeX4.ContourCubeX CC 
      Height          =   4440
      Left            =   105
      TabIndex        =   14
      Top             =   1170
      Width           =   11190
      Active          =   0   'False
      Transposed      =   0   'False
      NULLValueString =   ""
      Descending      =   0   'False
      NoTotals        =   0   'False
      NoGrandTotals   =   0   'False
      Caption         =   ""
      BackColor       =   16645369
      Enabled         =   -1  'True
      Alive           =   0   'False
      BorderStyle     =   0
      AllowDimOutside =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   "Totals"
      InactiveDimAreaBkColor=   15854051
      AutoSize        =   0   'False
      UnusedDataAreaColor=   16645369
      MousePointer    =   0
      Object.Visible         =   -1  'True
      InfoURL         =   "http://www.contourcomponents.com/contourcube_user_guide.htm"
      UseThemes       =   0   'False
      WordWrap        =   -1  'True
      FlatStyle       =   0
      FactsVAlignment =   0
      UnusedTreeAreaColor=   16645369
      DimLevelGradient=   14007466
      TreeLineColor   =   14007466
      DimLevelGradientStep=   20
      AllowDimVertical=   -1  'True
      AllowDimHorizontal=   -1  'True
      DrawOptions     =   2
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmProductPT.frx":0000
   End
   Begin VB.ComboBox cboStore 
      Height          =   315
      Left            =   1095
      TabIndex        =   8
      Text            =   "cboStore"
      Top             =   615
      Width           =   2685
   End
   Begin VB.CheckBox chkExVAT 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Values Ex V.A.T."
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5010
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CheckBox chkLDP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Use last delivered cost (not weighted average)"
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   5010
      TabIndex        =   11
      Top             =   330
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   10575
      Picture         =   "frmProductPT.frx":21D2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   105
      Width           =   1000
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   660
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
      Left            =   60
      Picture         =   "frmProductPT.frx":255C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   9540
      Picture         =   "frmProductPT.frx":28E6
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   1000
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Print"
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
      Left            =   1095
      Picture         =   "frmProductPT.frx":2C70
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1170
      TabIndex        =   3
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   77856769
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3345
      TabIndex        =   4
      Top             =   105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   77856769
      CurrentDate     =   37421
   End
   Begin VB.Label lblNote1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   930
      Width           =   11220
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Stores"
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
      Height          =   345
      Left            =   -450
      TabIndex        =   9
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Height          =   270
      Left            =   2700
      TabIndex        =   6
      Top             =   150
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "between"
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
      Height          =   270
      Left            =   195
      TabIndex        =   5
      Top             =   165
      Width           =   840
   End
End
Attribute VB_Name = "frmProductPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dte1 As Date
Dim dte2 As Date
Dim bOSOnly As Boolean
Dim rs As ADODB.Recordset
Dim bInclVAT As Boolean
Dim strCostOptions As String
Dim tlStores As New z_TextList

Dim lngStoreID As Long
Dim strStoreName As String

Public Sub Component(pRs As ADODB.Recordset, pType As String, Optional pInclVAT As Boolean = True, Optional pCostOptions As String = "")
    Set rs = pRs
    If rs.EOF Then Exit Sub
    bInclVAT = pInclVAT
    strCostOptions = pCostOptions
'    If bInclVAT Then
'        If strCostOptions = "LPD" Then
'            Me.lblHeading.Caption = "NOTE: All values include VAT, 'Cost' represents last delivered cost(if available)"
'        Else
'            Me.lblHeading.Caption = "NOTE: All values include VAT,'Cost' represents weighted average cost"
'        End If
'    Else
'        If strCostOptions = "LPD" Then
'            Me.lblHeading.Caption = "NOTE: All values exclude VAT, 'Cost' represents last delivered cost(if available)"
'        Else
'            Me.lblHeading.Caption = "NOTE: All values exclude VAT, 'Cost' represents weighted average cost"
'        End If
'    End If
    If UCase(pType) = "SUPPLIER" Then
        Me.Caption = "Documents per supplier"
    ElseIf UCase(pType) = "STORE" Then
        Me.Caption = "Documents per store"
    End If
End Sub
Private Sub cmdClose_Click()
    On Error GoTo errHandler
    CC.Active = False
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim SQL As String

    CC.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    WaitMsg "Loading the pivot table . . . ", True, Me
    CC.DataSourceType = xcdt_Recordset
    CC.Open rs
   
    CC.Active = True
    WaitMsg "", False, Me
    Me.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExport_Click()
    On Error GoTo errHandler
Dim Res As Boolean
Dim fs As New FileSystemObject
Dim strExecutable As String

    If fs.FileExists(oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html") Then
        On Error Resume Next
        fs.DeleteFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", True
        If fs.FileExists(oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html") Then
            MsgBox "It looks like a document is already open, Papyrus cannot create a new version of the document until it is closed.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
    CC.ReportToFile oPC.SharedFolderRoot & "\Temp\TransferDocuments.xls", "", 1
'    cc.ReportToFile "c:\PBKS\Temp\TransferDocuments.xls", "", 1
'    OpenFileWithApplication "c:\PBKS\Temp\TransferDocuments.xls", enExcel
    OpenFileWithApplication oPC.SharedFolderRoot & "\Temp\TransferDocuments.xls", enExcel
'
'    strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
'    If strExecutable = "" Then
'        MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder." & vbCrLf & "Report will not open now but is saved in " & oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", vbInformation, "Status"
'        Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'    Shell strExecutable & " " & oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", vbNormalFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPT.cmdExport_Click"
End Sub

Private Sub cmdOK_Click()
Dim strSQL As String

    If lngStoreID > 0 Then
        strSQL = "SELECT * FROM vTFR_General WHERE  TR_CaptureDate between '" & ReverseDate(dtpFrom) & "' AND '" & ReverseDate(dtpTo) & "' AND TR_TP_ID = " & lngStoreID
    Else
        strSQL = "SELECT * FROM vTFR_General WHERE  TR_CaptureDate between '" & ReverseDate(dtpFrom) & "' AND '" & ReverseDate(dtpTo) & "'"
    End If
    Set rs = New ADODB.Recordset
    DoEvents
    rs.Open strSQL, oPC.CO
    Screen.MousePointer = vbDefault
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
    Else
        LoadPT
    End If
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    CC.PrintCube (xprf_NoPreview)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    Set tlStores = New z_TextList
    tlStores.Load ltStores, , "<ANY>"
    LoadCombo cboStore, tlStores
    lngStoreID = tlStores.KeyByOrdinalIndex(1)
    dtpFrom.Value = FirstOfMonth(DateAdd("m", -1, Date))
    dtpTo.Value = EndOfDay(Date)
End Sub

Private Sub LoadPT()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
    CC.Cube.Dims.Clear
    CC.Cube.Facts.Clear
    CC.Cube.BaseFacts.Clear

    If rs Is Nothing Then Exit Sub
    CC.Cube.Dims.Add "StoreName", "StoreName", xoft_String, xda_vertical
    CC.Cube.Dims.Add "Document", "Document", xoft_String, xda_vertical
    CC.Cube.Dims.Add "Description", "Description", xoft_String, xda_vertical
    CC.Cube.Dims.Add "DocumentType", "DocumentType", xoft_String, xda_horizontal
    CC.Cube.Dims.Add "Mth", "Mth", xoft_String, xda_horizontal
  '  CC.Cube.Dims.Add "PT_Code", "PT_Code", xoft_String, xda_outside
    
    CC.Cube.BaseFacts.Add "Qty", "Qty"
    CC.Cube.Facts.Add "Qty", "Qty", xfaa_SUM
    
    If bInclVAT Then
       ' CC.AddFact "Val", "ValNet", xfaa_SUM, "Value"
        CC.Cube.BaseFacts.Add "ValNet", "Val"
        CC.Cube.Facts.Add "ValNet", "ValNet", xfaa_SUM
        CC.Facts("ValNet").Visible = True
        If strCostOptions = "LPD" Then
 '           CC.AddFact "COST", "LPDIncVat", xfaa_SUM, "Cost"
            CC.Cube.BaseFacts.Add "COST", "LPDIncVat"
            CC.Cube.Facts.Add "COST", "COST", xfaa_SUM
            CC.Facts("COST").Visible = True
        Else
  '          CC.AddFact "COST", "Cost", xfaa_SUM, "Cost"
            CC.Cube.BaseFacts.Add "COST", "Cost"
            CC.Cube.Facts.Add "COST", "Cost", xfaa_SUM
            CC.Facts("COST").Visible = True
        End If
    Else
      '  CC.Cube.AddFact "Val", "ValNetExVAT", xfaa_SUM, "Value"
        CC.Cube.BaseFacts.Add "ValNetExVAT", "Val"
        CC.Cube.Facts.Add "ValNetExVAT", "ValNetExVAT", xfaa_SUM
        CC.Facts("ValNetExVAT").Visible = True
        If strCostOptions = "LPD" Then
        '    CC.AddFact "COST", "LPDExVat", xfaa_SUM, "Cost"
            CC.Cube.BaseFacts.Add "LPDExVat", "cost"
            CC.Cube.Facts.Add "LPDExVat", "LPDExVat", xfaa_SUM
            CC.Facts("LPDExVat").Visible = True
        Else
        '    CC.AddFact "COST", "costExVat", xfaa_SUM, "Cost"
            CC.Cube.BaseFacts.Add "costExVat", "cost"
            CC.Cube.Facts.Add "costExVat", "costExVat", xfaa_SUM
            CC.Facts("costExVat").Visible = True
        End If
    End If
    
'    CC.DimFlags("Description") = xfNoTotals + xfNoGrandTotals
'    CC.DimFlags("DocumentType") = xfNoTotals + xfNoGrandTotals
'    CC.DimFlags("Mth") = xfNoTotals + xfNoGrandTotals
'
'
'
'    CC.FieldFormat("Qty") = "######0"
'    If UCase(oPC.GetProperty("HideSeparatorsInCubes")) = "TRUE" Then
'        CC.FieldFormat("CQty") = "#####00"
'        CC.FieldFormat("Val") = "#####00"
'        CC.FieldFormat("COST") = "#####00"
'    Else
'        CC.FieldFormat("CQty") = "###,##0"
'        CC.FieldFormat("Val") = "####,##0.00"
'        CC.FieldFormat("COST") = "####,##0.00"
'    End If
'  '  cc.FieldFormat("CQty") = "##,##0"
'    CC.HDrillDownLevel = 1
'    CC.VDrillDownLevel = 1
'
'    CC.Active = False
'
    DoEvents
    Screen.MousePointer = vbHourglass
    CC.Cube.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        CC.Cube.Open rs, True
        CC.Cube.Active = True
    Else
        MsgBox "No records", , "Status"
    End If
   
    Me.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.LoadPT", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    CC.Width = Me.Width - (CC.Left + 400)
    lngDiff = CC.Height
    CC.Height = Me.Height - (CC.top + 1220)
    lngDiff = CC.Height - lngDiff
    cmdPrint.top = cmdPrint.top + lngDiff
    cmdExport.top = cmdExport.top + lngDiff
    cmdClose.top = cmdClose.top + lngDiff

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Sub cmdAll_Click()
    If cboStore.ListCount = 0 Then Exit Sub
    lngStoreID = 0
    cboStore.ListIndex = 0
End Sub

'Private Sub Form_Initialize()
'Dim lngStoreID As Long
'Dim lngPTID As Long
'Dim strSQL As String
'Dim rs As adodb.Recordset
'
'    Me.SB1.Panels(1).Text = "Loading . . . "
'    If lngStoreID > 0 Then
'        strSQL = "SELECT * FROM vTFR_General WHERE  TR_CaptureDate between '" & ReverseDate(dte1) & "' AND '" & ReverseDate(dte2) & "' AND TR_TP_ID = " & lngStoreID
'    Else
'        strSQL = "SELECT * FROM vTFR_General WHERE  TR_CaptureDate between '" & ReverseDate(dte1) & "' AND '" & ReverseDate(dte2) & "'"
'    End If
'    Set rs = New adodb.Recordset
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    rs.Open strSQL, oPC.CO
'    Screen.MousePointer = vbDefault
'    If rs.EOF Then
'        rs.Close
'        Set rs = Nothing
'        Me.SB1.Panels(1).Text = ""
'        GoTo EXIT_Handler
'    End If
'
'    Set frmR = New frmProductPT
'    frmR.Component rs, "STORE"
'    Me.SB1.Panels(1).Text = ""
'    frmR.Show 'vbModal
'    Set rs = Nothing
'
'
'End Sub


'Private Sub LoadStores()
'Dim vntItem As Variant
'    On Error GoTo errHandler
'Dim i As Integer
'Dim ar() As String
'    If tlStores.Count = 0 Then Exit Sub
'    cboStore.BeginUpdate
'    ReDim ar(0 To 1, tlStores.Count)
'    cboStore.Items.RemoveAllItems
'    i = 1
'        For Each vntItem In tlStores
'            ar(0, i) = tlStores.Item(i + 1)
'            ar(1, i) = tlStores.Key(i + 1)
'            i = i + 1
'        Next
'    cboStore.PutItems ar
'    cboStore.EndUpdate
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmSProductPT.LoadStores"
'End Sub
'

Property Get StoreID() As Long
    StoreID = lngStoreID
End Property
Property Get StartDate() As Date
    StartDate = CDate(dtpFrom.Value)
End Property
Property Get EndDate() As Date
    EndDate = CDate(dtpTo.Value)
End Property
Property Get StoreName() As String
    StoreName = strStoreName
End Property



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

