VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTradingPerformanceReport 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Trading performance"
   ClientHeight    =   10755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17955
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10755
   ScaleWidth      =   17955
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615.001
      Left            =   120
      TabIndex        =   1
      Top             =   255
      Width           =   16845
      _ExtentX        =   29713
      _ExtentY        =   16960
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "By supplier"
      TabPicture(0)   =   "frmTradingPerformanceReport.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSupplier"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "arViewerSupplier"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSelectSuppliers"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdAllSuppliers"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdToPDF_Supplier"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdToExcel_Supplier"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOKSuppliers"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frSelection"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "By category"
      TabPicture(1)   =   "frmTradingPerformanceReport.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblCategory"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "arViewerCategory"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdOKCategory"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboSection"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdToExcel_Category"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdToPDF_Category"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Summary"
      TabPicture(2)   =   "frmTradingPerformanceReport.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdToPDF_Summary"
      Tab(2).Control(1)=   "cmdToExcel_Summary"
      Tab(2).Control(2)=   "cmdOKSummary"
      Tab(2).Control(3)=   "ARViewerSummary"
      Tab(2).ControlCount=   4
      Begin VB.Frame frSelection 
         Caption         =   "Selection"
         Height          =   570
         Left            =   -68895.01
         TabIndex        =   20
         Top             =   375
         Width           =   3540
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            Height          =   270
            Left            =   2640
            TabIndex        =   23
            Top             =   225
            Width           =   585
         End
         Begin VB.OptionButton optTop25 
            Caption         =   "Top 25"
            Height          =   270
            Left            =   405
            TabIndex        =   22
            Top             =   225
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optTop50 
            Caption         =   "Top 50"
            Height          =   270
            Left            =   1515
            TabIndex        =   21
            Top             =   225
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdToPDF_Summary 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   -61200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmdToExcel_Summary 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Spreadsheet"
         Height          =   360
         Left            =   -59790
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmdToPDF_Category 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   13800
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmdToExcel_Category 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Spreadsheet"
         Height          =   360
         Left            =   15210
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   540
         Width           =   1380
      End
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
         ItemData        =   "frmTradingPerformanceReport.frx":0054
         Left            =   1590
         List            =   "frmTradingPerformanceReport.frx":0056
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   525
         Width           =   3240
      End
      Begin VB.CommandButton cmdOKSuppliers 
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
         Height          =   570
         Left            =   -65205
         Picture         =   "frmTradingPerformanceReport.frx":0058
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   375
         Width           =   1000
      End
      Begin VB.CommandButton cmdOKCategory 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   6120
         Picture         =   "frmTradingPerformanceReport.frx":03E2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   375
         Width           =   1000
      End
      Begin VB.CommandButton cmdOKSummary 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   -68715.01
         Picture         =   "frmTradingPerformanceReport.frx":076C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   390
         Width           =   1000
      End
      Begin VB.CommandButton cmdToExcel_Supplier 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Spreadsheet"
         Height          =   360
         Left            =   -59775
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   540
         Width           =   1380
      End
      Begin VB.CommandButton cmdToPDF_Supplier 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   -61230
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   555
         Width           =   1380
      End
      Begin VB.CommandButton cmdAllSuppliers 
         BackColor       =   &H00C4BCA4&
         Caption         =   "All"
         Height          =   450
         Left            =   -69615.01
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   585
      End
      Begin VB.CommandButton cmdSelectSuppliers 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Select supplier"
         Height          =   435
         Left            =   -74790.01
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   435
         Width           =   1470
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewerSupplier 
         Height          =   8370.001
         Left            =   -74805.01
         TabIndex        =   2
         Top             =   960
         Width           =   16470
         _ExtentX        =   29051
         _ExtentY        =   14764
         SectionData     =   "frmTradingPerformanceReport.frx":0AF6
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewerCategory 
         Height          =   8370.001
         Left            =   195
         TabIndex        =   3
         Top             =   960
         Width           =   16470
         _ExtentX        =   29051
         _ExtentY        =   14764
         SectionData     =   "frmTradingPerformanceReport.frx":0B32
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARViewerSummary 
         Height          =   8370.001
         Left            =   -74805.01
         TabIndex        =   9
         Top             =   960
         Width           =   16470
         _ExtentX        =   29051
         _ExtentY        =   14764
         SectionData     =   "frmTradingPerformanceReport.frx":0B6E
      End
      Begin VB.Label lblCategory 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   225
         TabIndex        =   14
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label lblSupplier 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   -73290.01
         TabIndex        =   6
         Top             =   480
         Width           =   3660
      End
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
      Left            =   17025
      Picture         =   "frmTradingPerformanceReport.frx":0BAA
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   870
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   720
      TabIndex        =   19
      Top             =   30
      Width           =   6270
   End
End
Attribute VB_Name = "frmTradingPerformanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngTPSID As Long
Dim lngTPCID As Long
Dim strTPName As String
Dim strCustomerOrSupplier As String
Dim oRPT As z_reports
Dim bNP As Boolean
Dim rsSupplier As ADODB.Recordset
Dim rsCategory As ADODB.Recordset
Dim rsSummary As ADODB.Recordset
Dim arSupplier As New arPerformance
Dim arCategory As New arPerformance
Dim arSummary As New arPerformance
Dim tlCat As New z_TextList

Private Sub cboSection_Click()
    On Error GoTo errHandler
    lngTPCID = oPC.Configuration.Sections.Key(cboSection)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cboSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAllCategories_Click()
    On Error GoTo errHandler
    lngTPCID = 0
    lngTPSID = 0
    lblCategory = "<ALL>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdAllCategories_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAllSuppliers_Click()
    On Error GoTo errHandler
    lngTPCID = 0
    lngTPSID = 0
    lblSupplier = "<ALL>"
    frSelection.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdAllSuppliers_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOKSuppliers_Click()
    On Error GoTo errHandler
    Dim lngTop As Long
    
    Screen.MousePointer = vbHourglass
    Set oRPT = New z_reports
    If lngTPSID = 0 Then
        If optTop50 = True Then
            lngTop = 50
        ElseIf optTop25 = True Then
            lngTop = 25
        Else
            lngTop = 0
        End If
    Else
        lngTop = 0
    End If
    Set rsSupplier = oRPT.GetPerformanceData_Supplier(lngTPSID, lngTop) 'rs ', Me.dtpFrom, Me.dtpTo, lngTPCID, lngTPSID, bNP
    Set oRPT = Nothing
    
    Set arSupplier = New arPerformance
    arSupplier.PageSettings.Orientation = ddOLandscape
    arSupplier.Component "Supplier", "", rsSupplier
    
    arViewerSupplier.ReportSource = arSupplier
    Screen.MousePointer = vbDefault

EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    ErrorIn "frmPeriodDialogue.cmdOK_Click"
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdOKSuppliers_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOKCategory_Click()
    On Error GoTo errHandler
    
    Screen.MousePointer = vbHourglass
    Set oRPT = New z_reports
    Set rsCategory = oRPT.GetPerformanceData_Category(lngTPCID)  'rs ', Me.dtpFrom, Me.dtpTo, lngTPCID, lngTPSID, bNP
    Set oRPT = Nothing
    
    Set arCategory = New arPerformance
    arCategory.PageSettings.Orientation = ddOLandscape
    arCategory.Component "Category", "", rsCategory
    
    arViewerCategory.ReportSource = arCategory
    Screen.MousePointer = vbDefault

EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    ErrorIn "frmPeriodDialogue.cmdOK_Click"
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdOKCategory_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOKSummary_Click()
    On Error GoTo errHandler
    
    Screen.MousePointer = vbHourglass
    Set oRPT = New z_reports
    Set rsSummary = oRPT.GetPerformanceData_Summary  'rs ', Me.dtpFrom, Me.dtpTo, lngTPCID, lngTPSID, bNP
    Set oRPT = Nothing
    
    Set arSummary = New arPerformance
    arSummary.PageSettings.Orientation = ddOLandscape
    arSummary.Component "Summary", "", rsSummary
    
    ARViewerSummary.ReportSource = arSummary
    Screen.MousePointer = vbDefault

EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdOKSummary_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get TPCID() As Long
    TPCID = lngTPCID
End Property
Public Property Get TPSID() As Long
    TPSID = lngTPSID
End Property

Private Sub cmdSelectCategory_Click()
    On Error GoTo errHandler
Dim frmS As frmBrowseSUppliers2
        Set frmS = New frmBrowseSUppliers2
        frmS.Show vbModal
        lngTPSID = frmS.SupplierID
        strTPName = frmS.SupplierName
        Unload frmS
        lblSupplier.Caption = strTPName
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdSelectCategory_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelectSuppliers_Click()
    On Error GoTo errHandler
Dim frmS As frmBrowseSUppliers2
        Set frmS = New frmBrowseSUppliers2
        frmS.Show vbModal
        lngTPSID = frmS.SupplierID
        strTPName = frmS.SupplierName
        Unload frmS
        lblSupplier.Caption = strTPName
        frSelection.Enabled = False

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdSelectSuppliers_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdToPDF_Supplier_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If arSupplier Is Nothing Then Exit Sub
    If arSupplier.Pages.Count = 0 Then Exit Sub
    arSupplier.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValueBySupplier" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(arSupplier.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdToPDF_Supplier_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Supplier_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If arSupplier Is Nothing Then Exit Sub
    If arSupplier.Pages.Count = 0 Then Exit Sub
    arSupplier.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValueBySupplier" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(arSupplier.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdToExcel_Supplier_Click", , EA_NORERAISE
    HandleError
End Sub
'''''''''''''''''''''
Private Sub cmdToPDF_Category_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If arCategory Is Nothing Then Exit Sub
    If arCategory.Pages.Count = 0 Then Exit Sub
    arCategory.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValueByCategory" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(arCategory.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdToPDF_Category_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Category_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If arCategory Is Nothing Then Exit Sub
    If arCategory.Pages.Count = 0 Then Exit Sub
    arCategory.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValueByCategory" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(arCategory.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdToExcel_Category_Click", , EA_NORERAISE
    HandleError
End Sub
'''''''''''''''''''
Private Sub cmdToPDF_Summary_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If arSummary Is Nothing Then Exit Sub
    If arSummary.Pages.Count = 0 Then Exit Sub
    arSummary.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValueSummary" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(arSummary.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdToPDF_Summary_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Summary_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If arSummary Is Nothing Then Exit Sub
    If arSummary.Pages.Count = 0 Then Exit Sub
    arSummary.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValueSummary" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(arSummary.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.cmdToExcel_Summary_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
'    tlCat.Load ltSectionsActive, , "<any>"
    LoadCombo cboSection, oPC.Configuration.Sections
    Me.SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.Form_Load", , EA_NORERAISE
    HandleError
End Sub


'Public Property Get NP() As Boolean
'    NP = (chkNP = 1)
'End Property


Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    SSTab1.Width = NonNegative_Lng(Me.Width - 1300)
    arViewerSupplier.Width = NonNegative_Lng(SSTab1.Width - 400)
    arViewerCategory.Width = NonNegative_Lng(SSTab1.Width - 400)
    ARViewerSummary.Width = NonNegative_Lng(SSTab1.Width - 400)
    cmdClose.left = SSTab1.Width + 200
    
    lngDiff = arViewerSupplier.Height
    SSTab1.Height = NonNegative_Lng(Me.Height - 900)
    
    arViewerSupplier.Height = NonNegative_Lng(SSTab1.Height - 1100)
    arViewerCategory.Height = NonNegative_Lng(SSTab1.Height - 1100)
    ARViewerSummary.Height = NonNegative_Lng(SSTab1.Height - 1100)
    
    lngDiff = arViewerSupplier.Height - lngDiff
    cmdToExcel_Supplier.left = NonNegative_Lng(SSTab1.Width - 1605)
    cmdToPDF_Supplier.left = NonNegative_Lng(SSTab1.Width - 3020)
    
    cmdToExcel_Category.left = NonNegative_Lng(SSTab1.Width - 1605)
    cmdToPDF_Category.left = NonNegative_Lng(SSTab1.Width - 3020)

    cmdToExcel_Summary.left = NonNegative_Lng(SSTab1.Width - 1605)
    cmdToPDF_Summary.left = NonNegative_Lng(SSTab1.Width - 3020)
    Select Case SSTab1.Tab
    Case 0
        cmdToExcel_Category.Visible = False
        cmdToPDF_Category.Visible = False
        cmdToExcel_Summary.Visible = False
        cmdToPDF_Summary.Visible = False
    Case 1
        cmdToExcel_Supplier.Visible = False
        cmdToPDF_Supplier.Visible = False
        cmdToExcel_Summary.Visible = False
        cmdToPDF_Summary.Visible = False
    Case 2
        cmdToExcel_Category.Visible = False
        cmdToPDF_Category.Visible = False
        cmdToExcel_Supplier.Visible = False
        cmdToPDF_Supplier.Visible = False
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    Select Case SSTab1.Tab
    Case 0
        cmdToExcel_Supplier.Visible = True
        cmdToPDF_Supplier.Visible = True
        cmdToExcel_Category.Visible = False
        cmdToPDF_Category.Visible = False
        cmdToExcel_Summary.Visible = False
        cmdToPDF_Summary.Visible = False
    Case 1
        cmdToExcel_Category.Visible = True
        cmdToPDF_Category.Visible = True
        cmdToExcel_Supplier.Visible = False
        cmdToPDF_Supplier.Visible = False
        cmdToExcel_Summary.Visible = False
        cmdToPDF_Summary.Visible = False
    Case 2
        cmdToExcel_Summary.Visible = True
        cmdToPDF_Summary.Visible = True
        cmdToExcel_Category.Visible = False
        cmdToPDF_Category.Visible = False
        cmdToExcel_Supplier.Visible = False
        cmdToPDF_Supplier.Visible = False
    End Select

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTradingPerformanceReport.SSTab1_Click(PreviousTab)", PreviousTab, EA_NORERAISE
    HandleError
End Sub

