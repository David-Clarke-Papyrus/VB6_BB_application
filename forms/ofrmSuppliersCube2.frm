VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{CCA2C66D-33FD-11D5-8D72-005004532BDF}#1.3#0"; "CCubeX.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSuppliersCube 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Suppliers cube"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin CCubeX.ContourCubeX CC 
      Height          =   5370
      Left            =   105
      TabIndex        =   13
      Top             =   1770
      Width           =   11550
      BackColor       =   13882315
      Enabled         =   -1  'True
      MainAxis        =   0
      DataSourceType  =   0
      ConnectionString=   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=PT;Data Source=PAPYRUS-94TNP9S"
      SQL             =   ""
      PreGrouping     =   -1  'True
      Active          =   0   'False
      HDrillDownLevel =   -1
      VDrillDownLevel =   1
      Transposed      =   0   'False
      SuppressZeroRows=   0   'False
      SuppressZeroCols=   0   'False
      ViewFlags       =   0
      BorderStyle     =   1
      AllowInactiveDimArea=   -1  'True
      AllowFilter     =   -1  'True
      AllowExpand     =   -1  'True
      AllowPivot      =   -1  'True
      AllowTitle      =   0   'False
      ShowAsPercent   =   0
      TotalsString    =   ""
      CubeTitle       =   "Documents per supplier per period"
      TitleAlign      =   0
      TitleBkColor    =   13160660
      DimBkColor      =   13160660
      DimTitleBkColor =   6956042
      DimTitleInactiveBkColor=   8421504
      DimFilterBkColor=   13160660
      InactiveDimAreaBkColor=   13160660
      HeadingBkColor  =   13160660
      DataGridColor   =   13160660
      DataBkColor     =   16777215
      TotalBkColor    =   14679807
      GrandTotalBkColor=   14679807
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DimFilterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeadingFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DataFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TotalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GrandTotalFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
      Object.Visible         =   -1  'True
      MousePointer    =   0
      TitleForeColor  =   0
      DimForeColor    =   0
      DimTitleForeColor=   16777215
      DimFilterForeColor=   0
      HeadingForeColor=   0
      DataForeColor   =   0
      TotalForeColor  =   -2147483640
      GrandTotalForeColor=   -2147483640
      UnusedDataAreaColor=   -2147483643
      MainAxisDim     =   ""
      DimTitleDragBkColor=   32768
      FactsCaption    =   "Facts"
      ShowFactsBitmap =   -1  'True
      ADOCursorLocation=   2
      AutoRefreshView =   0   'False
      FPErrString     =   "FPErr"
      NULLValueString =   ""
      NonExistentValueString=   ""
      DefaultFactFormat=   "###,###,###,###,###,##0.00"
      AllowFactFilter =   -1  'True
      VERSION_NO      =   2
      FIELDS_SETTINGS =   $"frmSuppliersCube2.frx":0000
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   315
      Left            =   7275
      OleObjectBlob   =   "frmSuppliersCube2.frx":2004
      TabIndex        =   17
      Top             =   75
      Visible         =   0   'False
      Width           =   1320
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
      Picture         =   "frmSuppliersCube2.frx":33AE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7185
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
      Picture         =   "frmSuppliersCube2.frx":3738
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7185
      Width           =   1000
   End
   Begin VB.CheckBox chkLDP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Use last delivered cost (not weighted average)"
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   5295
      TabIndex        =   12
      Top             =   555
      Width           =   2415
   End
   Begin VB.CheckBox chkExVAT 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Values Ex V.A.T."
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5295
      TabIndex        =   11
      Top             =   225
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   10455
      Picture         =   "frmSuppliersCube2.frx":3AC2
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   150
      Width           =   1000
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
      Left            =   9390.001
      Picture         =   "frmSuppliersCube2.frx":3E4C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   150
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
      Left            =   4245
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   690
      Width           =   660
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   390
      Left            =   1665
      TabIndex        =   4
      Top             =   690
      Width           =   2550
   End
   Begin VB.CommandButton cmdSupp 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Select &supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   675
      Width           =   1440
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   225
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
      Format          =   16580609
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3555
      TabIndex        =   1
      Top             =   225
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
      Format          =   16580609
      CurrentDate     =   37421
   End
   Begin VB.Label lblNote1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   285
      TabIndex        =   5
      Top             =   1155
      Width           =   11220
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   270
      TabIndex        =   16
      Top             =   1470
      Width           =   11220
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
      Left            =   405
      TabIndex        =   8
      Top             =   255
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
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
      Left            =   7035
      TabIndex        =   6
      Top             =   1065
      Visible         =   0   'False
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
      Left            =   2910
      TabIndex        =   2
      Top             =   255
      Width           =   435
   End
End
Attribute VB_Name = "frmSuppliersCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim strSupplierName As String
Dim lngPTID As Long
Dim strPT As String
Dim bCancelled As Boolean
Dim bInclVAT As Boolean
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim bPOETA As Boolean

Public Sub Component(pPOETA As Boolean)
    bPOETA = pPOETA
End Sub


Private Sub SetupPT()
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add "Seesafe"
    cboProductType.Columns(0).Width = 190
    cboProductType.Columns(1).Width = 0
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate
End Sub




Private Sub chkExVAT_Click()
  '  bIncludeVAT
End Sub

Private Sub cmdAll_Click()
    strSupplierName = "<ALL>"
    lngTPID = 0
    txtSupplier = strSupplierName
End Sub


Private Sub cmdOK_Click()
Dim oSQL As New z_SQL
   ' Me.lblNote1.Caption = ""
    Me.lblHeading.Caption = ""
    
    If Not bPOETA Then
        If lngTPID > 0 Then
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_1 WHERE  TR_CaptureDate between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "') AND TR_TP_ID = " & lngTPID
        Else
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_1 WHERE  TR_CaptureDate between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "')"
        End If
        Me.lblNote1.Caption = "Dates used are our delivery received dates."
    Else
'        If lngTPID > 0 Then
'            strSQL = "SELECT *,Qty*-1 as CQTY FROM vPOLS_WithETA WHERE  POL_ETA between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "') AND TR_TP_ID = " & lngTPID
'        Else
'            strSQL = "SELECT *,Qty*-1 as CQTY FROM vPOLS_WithETA WHERE  POL_ETA between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "')"
'        End If
        If lngTPID > 0 Then
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_1 WHERE  Dte2 between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "') AND TR_TP_ID = " & lngTPID
        Else
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_1 WHERE  Dte2 between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "')"
        End If
        lblNote1.Caption = "Dates used are supplier invoice dates."

    End If
    Set rs = New ADODB.Recordset
    oSQL.GetDynamicRecordset_Improved strSQL, enText, Array(), "", rs
  
    Preparecube
    If chkLDP = 1 Then
        Me.lblHeading.Caption = "NOTE: 'RetailValue' : Retail price pre-disc (inc), 'Linecost': Retail price post-disc (excl), 'InvtryCost':Last del. cost(if available) (excl)"
    Else
        Me.lblHeading.Caption = "NOTE: 'RetailValue' : Retail price pre-disc (inc), 'Linecost': Retail price post-disc (excl), 'InvtryCost':Avg cost (excl)"
    End If
    
End Sub


Private Sub cmdSupp_Click()
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    strSupplierName = frm.SupplierName
    txtSupplier = strSupplierName
    Unload frm
    If lngTPID = 0 Then Exit Sub

End Sub

Private Sub Form_Load()
Dim ar() As String
    cboProductType.BeginUpdate
'    oPC.Configuration.ProductTypes.CollectionAsArray ar
'    cboProductType.PutItems ar
'    cboProductType.EndUpdate
    bInclVAT = False

    If bPOETA Then
        Me.Caption = "Purchase orders by expected delivery date"
        lblNote1 = "Purchase orders are selected by the line item expected delivery date"
    End If
    SetupPT
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub


Private Sub cboProductType_SelectionChanged()
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    strPT = cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0)
End Sub

'Property Get SupplierID() As Long
'    SupplierID = lngTPID
'End Property
'Property Get PTID() As Long
'    PTID = lngPTID
'End Property
'Property Get StartDate() As Date
'    StartDate = CDate(dtpFrom.Value)
'End Property
'Property Get EndDate() As Date
'    EndDate = CDate(dtpTo.Value)
'End Property
'Property Get SupplierName() As String
'    SupplierName = strSupplierName
'End Property
'Property Get PTName() As String
'    PTName = strPT
'End Property
'Public Property Get Cancelled() As Boolean
'    Cancelled = bCancelled
'End Property
'



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

'Private Sub LoadPivot()
'    On Error GoTo errHandler
'Dim rs As New ADODB.Recordset
'Dim SQL As String
'
'    CC.Active = False
'    If rs.State <> 0 Then
'        rs.Close
'    End If
'    WaitMsg "Loading the pivot table . . . ", True, Me
'    CC.DataSourceType = xcdt_Recordset
'    CC.Open rs
'
'    CC.Active = True
'    WaitMsg "", False, Me
'    Me.Refresh
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmPT.cmdFetch_Click", , EA_NORERAISE
'    HandleError
'End Sub

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
    CC.ExportToFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", xet_html
    strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
    If strExecutable = "" Then
        MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder." & vbCrLf & "Report will not open now but is saved in " & oPC.SharedFolderRoot & "\HTML\SupplierCharts.html", vbInformation, "Status"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Shell strExecutable & " " & oPC.SharedFolderRoot & "\HTML\SupplierDocuments.html", vbNormalFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSuppliersCube.cmdExport_Click"
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    CC.PrintCube True, False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Preparecube()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple

    If rs Is Nothing Then Exit Sub
    CC.Active = False
    CC.ClearFields
    CC.AddDimension "SupplierName", "Supplier", xda_vertical, 1
    CC.AddDimension "Document", "Document", xda_vertical, 2
    CC.AddDimension "Description", "Description", xda_vertical, 3
    CC.AddDimension "DocumentType", "Document type", xda_horizontal, 1
    CC.AddDimension "Mth", "Month", xda_horizontal, 2
    CC.AddDimension "PT_Code", "Product type", xda_outside, 1
    CC.AddFact "CQty", "QTY", xfaa_SUM, "Qty"
    
    CC.AddFact "Val", "Val", xfaa_SUM, "RetailValue"
    If bInclVAT Then
        CC.AddFact "ValNet", "ValNet", xfaa_SUM, "LineCost"
        If chkLDP = 1 Then
            CC.AddFact "COST", "LPDExVat", xfaa_SUM, "InvtryCost"
        Else
            CC.AddFact "COST", "Cost", xfaa_SUM, "InvtryCost"
        End If
    Else
        CC.AddFact "ValNet", "ValNetExVAT", xfaa_SUM, "LineCost"
        If chkLDP = 1 Then
            CC.AddFact "COST", "LPDExVat", xfaa_SUM, "InvtryCost"
        Else
            CC.AddFact "COST", "costExVat", xfaa_SUM, "InvtryCost"
        End If
    End If
    
    CC.DimFlags("Description") = xfNoTotals + xfNoGrandTotals
    CC.DimFlags("DocumentType") = xfNoTotals + xfNoGrandTotals
    CC.DimFlags("Mth") = xfNoTotals + xfNoGrandTotals
    CC.FieldFormat("CQty") = "##,##0"
    CC.HDrillDownLevel = 1
    CC.VDrillDownLevel = 1
    
    DoEvents
    Screen.MousePointer = vbHourglass
    CC.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        CC.Open rs
        CC.Active = True
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
    CC.Width = Me.Width - (CC.left + 400)
    lngDiff = CC.Height
    CC.Height = Me.Height - (CC.top + 1220)
    lngDiff = CC.Height - lngDiff
    cmdPrint.top = cmdPrint.top + lngDiff
    cmdExport.top = cmdExport.top + lngDiff
  '  cmdClose.top = cmdClose.top + lngDiff

End Sub
