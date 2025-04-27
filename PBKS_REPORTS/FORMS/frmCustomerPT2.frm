VERSION 5.00
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmCustomerPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer performance"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCustomerPT2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   Begin CCubeX4.ContourCubeX CC 
      Height          =   4995
      Left            =   105
      TabIndex        =   4
      Top             =   465
      Width           =   11550
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
      AllowDimOutside =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      TotalsString    =   "Totals"
      InactiveDimAreaBkColor=   13160660
      AutoSize        =   0   'False
      UnusedDataAreaColor=   16777215
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
      DrawOptions     =   7
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmCustomerPT2.frx":038A
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
      Picture         =   "frmCustomerPT2.frx":23B6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5670
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
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
      Left            =   10380
      Picture         =   "frmCustomerPT2.frx":2740
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5670
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
      Left            =   75
      Picture         =   "frmCustomerPT2.frx":2ACA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5655
      Width           =   1000
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   285
      TabIndex        =   0
      Top             =   120
      Width           =   11220
   End
End
Attribute VB_Name = "frmCustomerPT"
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

Public Sub Component(pRs As ADODB.Recordset, pType As String, Optional pInclVAT As Boolean = True, Optional pCostOptions As String = "")
    Set rs = pRs
    If rs.EOF Then Exit Sub
    bInclVAT = pInclVAT
    strCostOptions = pCostOptions
    If bInclVAT Then
        If strCostOptions = "LPD" Then
            Me.lblHeading.Caption = "NOTE: All values include VAT, 'Cost' represents last delivered cost(if available)"
        Else
            Me.lblHeading.Caption = "NOTE: All values include VAT,'Cost' represents weighted average cost"
        End If
    Else
        If strCostOptions = "LPD" Then
            Me.lblHeading.Caption = "NOTE: All values exclude VAT, 'Cost' represents last delivered cost(if available)"
        Else
            Me.lblHeading.Caption = "NOTE: All values exclude VAT, 'Cost' represents weighted average cost"
        End If
    End If
    rs.MoveFirst
    If UCase(pType) = "CUSTOMER" Then
        Me.Caption = "Documents per customer"
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
Dim Res As Boolean
Dim fs As New FileSystemObject
Dim strExecutable As String

    On Error GoTo errHandler
    If fs.FileExists(oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html") Then
        On Error Resume Next
        fs.DeleteFile oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html", True
        If fs.FileExists(oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html") Then
            MsgBox "It looks like a document is already open, Papyrus cannot create a new version of the document until it is closed.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
    CC.ExportToFile oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html", oPC.SharedFolderRoot & "\HTML\CustomerCharts.html"
    OpenFileWithApplication oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html", enExcel
'    strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
'    If strExecutable = "" Then
'        MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder." & vbCrLf & "Report will not open now but is saved in " & oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", vbInformation, "Status"
'        Exit Sub
'    End If
'    Screen.MousePointer = vbHourglass
'    Shell strExecutable & " " & oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html", vbNormalFocus
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
   
    CC.PrintCube True
  '  MsgBox "Printed", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPT.cmdPrint_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
TOP = 400
Left = 20
Width = 11900
Height = 6800

    CC.Cube.Dims.Clear
    CC.Cube.Facts.Clear
    CC.Cube.BaseFacts.Clear
    CC.Cube.Dims.Add "Customer", "CustName", xoft_String, xda_vertical
    CC.Cube.Dims.Add "Document", "DOCCode", xoft_String, xda_vertical
    CC.Cube.Dims.Add "Description", "Descr", xoft_String, xda_vertical
    CC.Cube.Dims.Add "Document type", "SaleType", xoft_String, xda_horizontal
    CC.Cube.Dims.Add "Month", "Mth", xoft_String, xda_outside
    CC.Cube.Dims.Add "Product type", "PT_Code", xoft_String, xda_outside
   ' CC.Cube.Dims.Add "Supplier", "Supplier", xda_outside, 1
    
    CC.Cube.BaseFacts.Add "Qty", "Qty"
    CC.Cube.Facts.Add "Qty", "Qty", xfaa_SUM
    CC.Facts("Qty").Visible = True
    CC.Facts("Qty").Visible = True
    
        CC.Cube.BaseFacts.Add "Val", "Val"
        CC.Cube.Facts.Add "Val", "Val", xfaa_SUM
        CC.Cube.BaseFacts.Add "ValExVat", "ValExVat"
        CC.Cube.Facts.Add "ValExVat", "ValExVat", xfaa_SUM
        CC.Cube.BaseFacts.Add "costExVat", "costExVat"
        CC.Cube.Facts.Add "CostExVat", "costExVat", xfaa_SUM
    CC.Facts("Val").Visible = True
    CC.Facts("ValExVat").Visible = True
    CC.Facts("CostExVat").Visible = True
 
'    CC.AddFact "Val", "Val", xfaa_SUM, "Value"
    
'    CC.DimFlags("Descr") = xfNoTotals + xfNoGrandTotals
'    CC.DimFlags("SaleType") = xfNoTotals + xfNoGrandTotals
'    CC.DimFlags("Mth") = xfNoTotals + xfNoGrandTotals
''    CC.AddFormula "Turn", "DL_Value", "testTurn"
'
'    CC.FieldFormat("Qty") = "######0"
'    If UCase(oPC.GetProperty("HideSeparatorsInCubes")) = "TRUE" Then
'    MsgBox "Hiding separators"
'        CC.FieldFormat("Val") = "#####00.00"
'        CC.FieldFormat("COST") = "#####00.00"
'    Else
'        CC.FieldFormat("Val") = "###,#00.00"
'        CC.FieldFormat("COST") = "###,#00.00"
'    End If
'    CC.HDrillDownLevel = 1
'    CC.VDrillDownLevel = 1
'    CC.Active = False
    
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
    ErrorIn "frmPT.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    CC.Width = Me.Width - (CC.Left + 400)
    lngDiff = CC.Height
    CC.Height = Me.Height - (CC.TOP + 1220)
    lngDiff = CC.Height - lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdExport.TOP = cmdExport.TOP + lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff

End Sub
