VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7A5C485E-4ACE-4C72-B64D-46119DEDD852}#4.0#0"; "CCubeX40.ocx"
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
   Begin CCubeX4.ContourCubeX CC 
      Height          =   5745
      Left            =   105
      TabIndex        =   13
      Top             =   1395
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
      BorderStyle     =   0
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
      DrawOptions     =   2
      ConnectionString=   ""
      DataSourceType  =   0
      VERSION_NO      =   2
      CCubeXMetadata  =   $"frmSuppliersCube2.frx":0000
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   315
      Left            =   7275
      OleObjectBlob   =   "frmSuppliersCube2.frx":21D5
      TabIndex        =   17
      Top             =   165
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
      Picture         =   "frmSuppliersCube2.frx":357F
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
      Left            =   60
      Picture         =   "frmSuppliersCube2.frx":3909
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
      Picture         =   "frmSuppliersCube2.frx":3C93
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
      Picture         =   "frmSuppliersCube2.frx":401D
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
      Format          =   77856769
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
      Format          =   77856769
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
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_1 WHERE  dte between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "') AND TR_TP_ID = " & lngTPID
        Else
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_1 WHERE  dte between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "')"
        End If
    Else
        If lngTPID > 0 Then
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_2 WHERE  Dte2 between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "') AND TR_TP_ID = " & lngTPID
        Else
            strSQL = "SELECT *,Qty*-1 as CQTY FROM zME_2 WHERE  Dte2 between dbo.StartofDay('" & ReverseDate(dtpFrom) & "') AND dbo.endofDay('" & ReverseDate(dtpTo) & "')"
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
    cc.Active = False
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
    Dim strCommand As String
    Dim strfIN As String
    Dim strfOut As String

    If fs.FileExists(oPC.SharedFolderRoot & "\HTML\SupplierDocuments.xls") Then
        On Error Resume Next
        fs.DeleteFile oPC.SharedFolderRoot & "\HTML\SupplierDocuments.xls", True
        If fs.FileExists(oPC.SharedFolderRoot & "\HTML\SupplierDocuments.xls") Then
            MsgBox "It looks like a document is already open, Papyrus cannot create a new version of the document until it is closed.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
    cc.ReportToFile oPC.SharedFolderRoot & "\Temp\SupplierDocuments.xls", "", 1
'    cc.ReportToFile "c:\Temp\SupplierDocuments.xls", "", ccubex4.xolaprpt_XLS
'    OpenFileWithApplication "c:\Temp\SupplierDocuments.xls", enExcel
 
    OpenFileWithApplication oPC.SharedFolderRoot & "\Temp\SupplierDocuments.xls", enExcel
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSuppliersCube.cmdExport_Click"
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    cc.PrintCube
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmPT.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Preparecube()
    On Error GoTo errHandler
Dim oTLS As New z_TextListSimple

    cc.Cube.Dims.Clear
    cc.Cube.Facts.Clear
    cc.Cube.BaseFacts.Clear
    
    If rs Is Nothing Then Exit Sub
    cc.Cube.Dims.Add "Supplier", "SupplierName", xoft_String, xda_vertical
    cc.Cube.Dims.Add "Document", "Document", xoft_String, xda_vertical
    cc.Cube.Dims.Add "Description", "Description", xoft_String, xda_vertical
    cc.Cube.Dims.Add "Document type", "DocumentType", xoft_String, xda_horizontal
    cc.Cube.Dims.Add "Mth", "Mth", xoft_String, xda_horizontal
 '   CC.Cube.Dims.Add "Product type", "PT_Code", xoft_String, xda_outside
    
    cc.Cube.BaseFacts.Add "Qty", "Qty"
    cc.Cube.Facts.Add "Qty", "Qty", xfaa_SUM
    cc.Facts("Qty").Visible = True

    cc.Cube.BaseFacts.Add "Val", "Val"
    cc.Cube.Facts.Add "Val", "Val", xfaa_SUM
    cc.Facts("Val").Visible = True
'    CC.Facts("Val1").Caption = "Value"
    
    cc.Cube.BaseFacts.Add "Cost", "LineCostExVat"
    cc.Cube.Facts.Add "Cost", "Cost", xfaa_SUM   'here the second parameter must match the first of the 'base fact"
    cc.Facts("Cost").Visible = True
 '   CC.Facts("Cost").Caption = "Cost"
    
'
    DoEvents
    Screen.MousePointer = vbHourglass
    cc.Cube.DataSourceType = xcdt_Recordset
    If Not rs.EOF Then
        cc.Cube.Open rs, True
        cc.Active = True
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
  '  cmdClose.top = cmdClose.top + lngDiff

End Sub
