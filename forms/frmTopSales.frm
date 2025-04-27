VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTopSaless 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Top sales"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   15600
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboSection 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmTopSales.frx":0000
      Left            =   8280
      List            =   "frmTopSales.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   720
      Width           =   3240
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00915A48&
      Height          =   315
      Left            =   1005
      TabIndex        =   18
      Text            =   "100"
      Top             =   735
      Width           =   720
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   11295
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   255
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   12750
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   255
      Width           =   1380
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
      Left            =   10185
      Picture         =   "frmTopSales.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   15
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
      Left            =   9195
      Picture         =   "frmTopSales.frx":038E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   15
      Width           =   1000
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      Height          =   1365
      Left            =   12960
      TabIndex        =   9
      Top             =   2715
      Visible         =   0   'False
      Width           =   1665
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   12
         Top             =   525
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   135
         TabIndex        =   11
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optCSV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&CSV"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   150
         TabIndex        =   10
         Top             =   870
         Width           =   1065
      End
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
      Left            =   8355
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   180
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
      Height          =   390
      Left            =   5775
      TabIndex        =   4
      Top             =   165
      Width           =   2550
   End
   Begin VB.CommandButton cmdSupp 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&per &supplier"
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
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   165
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1020
      TabIndex        =   0
      Top             =   210
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
      Format          =   191299585
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   195
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
      Format          =   191299585
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arTopSalesViewer 
      Height          =   6615
      Left            =   165
      TabIndex        =   15
      Top             =   1200
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   11668
      SectionData     =   "frmTopSales.frx":0718
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   315
      Left            =   3705
      OleObjectBlob   =   "frmTopSales.frx":0754
      TabIndex        =   22
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      TabIndex        =   21
      Top             =   750
      Width           =   1170
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "items"
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
      Left            =   1815
      TabIndex        =   20
      Top             =   765
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
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
      Left            =   525
      TabIndex        =   19
      Top             =   765
      Width           =   465
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
      Left            =   150
      TabIndex        =   8
      Top             =   270
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
      Left            =   2415
      TabIndex        =   6
      Top             =   750
      Width           =   1170
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
      Left            =   2445
      TabIndex        =   2
      Top             =   255
      Width           =   435
   End
End
Attribute VB_Name = "frmTopSaless"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPID As Long
Dim strSupplierName As String
Dim lngPTID As Long
Dim strPT As String
Dim rpt As arTopSales
Dim oRPT As z_reports
Dim lngcatID As Long
Dim strCat As String

Private Sub cboSection_Click()
    On Error GoTo errHandler
    lngcatID = oPC.Configuration.Sections.Key(cboSection)
    strCat = cboSection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTopSales.cboSection_Click", , EA_NORERAISE
    HandleError
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


Private Sub cmdAll_Click()
    strSupplierName = "<ALL>"
    lngTPID = 0
    txtSupplier = strSupplierName
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim blnNoRecsReturned As Boolean
Dim strErrMsg As String
Dim lngTPID As Long
Dim lngPTID As Long
Dim dte1 As Date
Dim dte2 As Date
Dim strSupplierName As String
Dim strPTName As String
Dim rs As ADODB.Recordset
Dim strCaption As String

    Set oRPT = New z_reports
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    strErrMsg = oRPT.TopSales(rs, Me.dtpFrom, Me.dtpTo, Me.SupplierID, Me.PTID, lngcatID, CLng(txtQty))
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No records returned", vbOKOnly, "Papyrus II Reports"
        GoTo EXIT_Handler
    End If
    
    strCaption = "Top sales for the period " & Format(StartDate, "dd/mm/yyyy") & " to " & Format(EndDate, "dd/mm/yyyy")
    If SupplierName > "" Then
        strCaption = strCaption & "  for supplier: " & SupplierName
    End If
    If PTName > "" Then
        strCaption = strCaption & "  for product type: " & PTName
    End If
    If strCat > "" Then
        strCaption = strCaption & "  for main category: " & strCat
    End If
    
    Set rpt = Nothing
    Set rpt = New arTopSales
   ' rpt.Visible = False
    rpt.Component rs, strCaption
    
    arTopSalesViewer.ReportSource = rpt
    
    Screen.MousePointer = vbDefault
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTopSaless.cmdOK_Click"
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

Private Sub Form_Initialize()
Dim ar() As String
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate
    lngcatID = 0

End Sub

Private Sub Form_Load()
    SetupPT
    LoadCombo cboSection, oPC.Configuration.Sections
    cboSection = "<ALL>"
    
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub


Private Sub cboProductType_SelectionChanged()
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    strPT = cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0)
End Sub

Property Get SupplierID() As Long
    SupplierID = lngTPID
End Property
Property Get PTID() As Long
    PTID = lngPTID
End Property
Property Get StartDate() As Date
    StartDate = CDate(dtpFrom.Value)
End Property
Property Get EndDate() As Date
    EndDate = CDate(dtpTo.Value)
End Property
Property Get SupplierName() As String
    SupplierName = strSupplierName
End Property
Property Get PTName() As String
    PTName = strPT
End Property
Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If rpt Is Nothing Then Exit Sub
    rpt.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "TopSales" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If rpt Is Nothing Then Exit Sub
    rpt.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "TopSales" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arTopSalesViewer.Width = Me.Width - 600
    lngDiff = arTopSalesViewer.Height
    arTopSalesViewer.Height = Me.Height - 1500
    lngDiff = arTopSalesViewer.Height - lngDiff
    cmdToExcel.Left = arTopSalesViewer.Left + arTopSalesViewer.Width - cmdToExcel.Width
    cmdToPDF.Left = arTopSalesViewer.Left + arTopSalesViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    Cancel = Not (IsNumeric(txtQty))
End Sub
