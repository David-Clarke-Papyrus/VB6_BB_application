VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesByPeriod 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Period selection"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14265
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   14265
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkIsSummary 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Summarize"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   6480
      TabIndex        =   18
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox cboSection 
      Appearance      =   0  'Flat
      Height          =   345
      ItemData        =   "frmSalesByPeriod.frx":0000
      Left            =   6435
      List            =   "frmSalesByPeriod.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   315
      Width           =   2790
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   75
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   12585
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   435
      Width           =   1380
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
      Left            =   9570
      Picture         =   "frmSalesByPeriod.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   75
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   10605
      Picture         =   "frmSalesByPeriod.frx":038E
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   75
      Width           =   1000
   End
   Begin VB.CheckBox chkNP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New page per section"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   75
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   11775
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   1665
      Begin VB.OptionButton optCSV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&CSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   870
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   135
         TabIndex        =   6
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2355
      TabIndex        =   1
      Top             =   90
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   114163713
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4545
      TabIndex        =   2
      Top             =   75
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   114163713
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   6495
      Left            =   30
      TabIndex        =   11
      Top             =   990
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   11456
      SectionData     =   "frmSalesByPeriod.frx":0718
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
      Left            =   6450
      TabIndex        =   17
      Top             =   30
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "later date ( end of day)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   240
      Left            =   4590
      TabIndex        =   15
      Top             =   435
      Width           =   1710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "earlier date (start of day)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   240
      Left            =   2295
      TabIndex        =   14
      Top             =   435
      Width           =   1800
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4140
      TabIndex        =   3
      Top             =   135
      Width           =   555
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select period between"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSalesByPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim bCancel As Boolean
Public enPrevPrintCSV As enumReportPresentation
Dim oRPT As z_reports
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rpt As Object    'arSalesByPeriod
Dim mType As String
Dim mIsSummary As Boolean
Dim lngcatID As Long
Dim strCat As String


Public Sub Component(pMsg As String, typ As String)
    Me.Caption = pMsg
    lblDescription.Caption = pMsg
    mType = typ
        
End Sub
Public Sub Component2(Optional pShowchkNP As Boolean)
    Me.chkNP.Visible = pShowchkNP
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
10        On Error GoTo errHandler
          
Dim dteLimitToView As Date
Dim oSQL As z_SQL
Dim strCaption As String

    If oPC.BlindCashup = True Then
        Set oSQL = New z_SQL
        dteLimitToView = oSQL.GetDateOfEarliestUnSignedSession
        If Me.dtpTo >= StartOfDay(dteLimitToView) Then
            MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
20        Screen.MousePointer = vbHourglass
          strCaption = lblDescription.Caption
30        Set rpt = Nothing
40        Set oRPT = New z_reports
50        Set rs = New ADODB.Recordset
60        rs.CursorLocation = adUseClient
70        Select Case mType
          Case "M", "W"
80            oRPT.SalesByPeriod rs, Me.dtpFrom, dtpTo, mType
90            Set rpt = New arSalesByPeriod
100       Case "PT"
110           Set rs2 = New ADODB.Recordset
120           rs2.CursorLocation = adUseClient
130           oRPT.SalesByPTByDate rs, dtpFrom, dtpTo
140           Set rpt = New arSalesSummaryByPT
150       Case "CAT"
160           Set rs = New ADODB.Recordset
170           rs.CursorLocation = adUseClient
                If chkIsSummary = 1 Then
                    oRPT.SalesBySectionByDate rs, dtpFrom, dtpTo, lngcatID, True
                    Set rpt = New arSalesSummaryBySectionSumm
                Else
180                 oRPT.SalesBySectionByDate rs, dtpFrom, dtpTo, lngcatID, False
                    Set rpt = New arSalesSummaryBySection
                End If
190
                If strCat > "" Then
                    strCaption = strCaption & "  for main category: " & strCat
                End If
200       End Select
210       rpt.Component strCaption, rs, Me.dtpFrom, Me.dtpTo
          
220       Me.arViewer.ReportSource = rpt
          
230       Screen.MousePointer = vbDefault
          
EXIT_Handler:
240       Me.MousePointer = vbDefault
250       Exit Sub
errHandler:
260       ErrorIn "frmPeriodDialogue.cmdOK_Click"
End Sub



Private Sub Form_Initialize()
    lngcatID = 0

End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    LoadCombo cboSection, oPC.Configuration.Sections
    cboSection = "<ALL>"
    
    
    dtpFrom.Value = DateAdd("m", -6, Date)
    dtpTo.Value = Date
End Sub
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

Public Property Get FromDate() As Date
    FromDate = dteFrom
End Property
Public Property Get ToDate() As Date
    ToDate = dteTo
End Property
Public Property Get CancelReport() As Boolean
    CancelReport = bCancel
End Property
Public Property Get Preview() As Boolean
    Preview = bPreview
End Property
Public Property Get NP() As Boolean
    NP = (chkNP = 1)
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
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
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
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arViewer.Width = Me.Width - 600
    lngDiff = arViewer.Height
    arViewer.Height = Me.Height - 1800
    lngDiff = arViewer.Height - lngDiff
    cmdToExcel.Left = arViewer.Left + arViewer.Width - cmdToExcel.Width
    cmdToPDF.Left = arViewer.Left + arViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10

End Sub
