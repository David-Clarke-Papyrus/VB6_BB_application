VERSION 5.00
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
Begin VB.Form frmCustomerPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Customer performance"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frmCustomerPT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11880
   Begin CCubeX4.ContourCubeX CC 
      Height          =   5445
      Left            =   210
      TabIndex        =   4
      Top             =   150
      Width           =   11385
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
      DrawOptions     =   7
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmCustomerPT.frx":038A
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
      Picture         =   "frmCustomerPT.frx":23B4
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Picture         =   "frmCustomerPT.frx":273E
      Style           =   1  'Graphical
      TabIndex        =   1
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
      Picture         =   "frmCustomerPT.frx":2AC8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5670
      Width           =   1000
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   285
      TabIndex        =   3
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
    cc.Active = False
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

    cc.Active = False
    If rs.State <> 0 Then
        rs.Close
    End If
    WaitMsg "Loading the pivot table . . . ", True, Me
    cc.DataSourceType = xcdt_Recordset
    cc.Open rs
   
    cc.Active = True
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
    cc.ExportToFile oPC.SharedFolderRoot & "\HTML\CustomerDocuments.html", oPC.SharedFolderRoot & "\HTML\CustomerCharts.html"
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
   
    cc.PrintCube True
  '  MsgBox "Printed", vbInformation, "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerPT.cmdPrint_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple
top = 400
Left = 20
Width = 11900
Height = 6800

    cc.Cube.Dims.Clear
    cc.Cube.Facts.Clear
    cc.Cube.BaseFacts.Clear
    cc.Cube.Dims.Add "Customer", "CustName", xoft_String, xda_vertical
    cc.Cube.Dims.Add "Document", "DOCCode", xoft_String, xda_vertical
    cc.Cube.Dims.Add "Description", "Descr", xoft_String, xda_vertical
    cc.Cube.Dims.Add "Document type", "SaleType", xoft_String, xda_horizontal
    cc.Cube.Dims.Add "Month", "Mth", xoft_String, xda_outside
    cc.Cube.Dims.Add "Product type", "PT_Code", xoft_String, xda_outside
   ' CC.Cube.Dims.Add "Supplier", "Supplier", xda_outside, 1
    
    cc.Cube.BaseFacts.Add "Qty", "Qty"
    cc.Cube.Facts.Add "Qty", "Qty", xfaa_SUM
    cc.Facts("Qty").Visible = True
    cc.Facts("Qty").Visible = True
    
        cc.Cube.BaseFacts.Add "Val", "Val"
        cc.Cube.Facts.Add "Val", "Val", xfaa_SUM
     '   CC.Cube.BaseFacts.Add "ValExVat", "ValExVat"
     '   CC.Cube.Facts.Add "ValExVat", "ValExVat", xfaa_SUM
        cc.Cube.BaseFacts.Add "costExVat", "costExVat"
        cc.Cube.Facts.Add "CostExVat", "costExVat", xfaa_SUM
    cc.Facts("Val").Visible = True
  '  CC.Facts("ValExVat").Visible = True
    cc.Facts("CostExVat").Visible = True
 
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
    cc.Cube.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        cc.Cube.Open rs, True
        cc.Cube.Active = True
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
    cc.Width = Me.Width - (cc.Left + 400)
    lngDiff = cc.Height
    cc.Height = Me.Height - (cc.top + 1220)
    lngDiff = cc.Height - lngDiff
    cmdPrint.top = cmdPrint.top + lngDiff
    cmdExport.top = cmdExport.top + lngDiff
    cmdClose.top = cmdClose.top + lngDiff

End Sub

