VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesBySupplier 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales by supplier"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15015
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
   ScaleHeight     =   8460
   ScaleWidth      =   15015
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1200
      Width           =   555
   End
   Begin VB.CheckBox chkNP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New page per section"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   300
      TabIndex        =   13
      Top             =   585
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdSelectTPS 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Select supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1185
      Width           =   1470
   End
   Begin VB.Frame frCost 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   8895
      TabIndex        =   9
      Top             =   30
      Visible         =   0   'False
      Width           =   3780
      Begin VB.OptionButton optLDC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses last delivered cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   390
         Left            =   255
         TabIndex        =   11
         Top             =   315
         Width           =   3135
      End
      Begin VB.OptionButton optWeighted 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses weighted average cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   10
         Top             =   645
         Value           =   -1  'True
         Width           =   3360
      End
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11295
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1275
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1275
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
      Left            =   6720
      Picture         =   "frmSalesBySupplier.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   135
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   7755
      Picture         =   "frmSalesBySupplier.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   135
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2580
      TabIndex        =   1
      Top             =   135
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   193593345
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   193593345
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arSalesBySupplierViewer 
      Height          =   6600
      Left            =   90
      TabIndex        =   6
      Top             =   1680
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   11642
      SectionData     =   "frmSalesBySupplier.frx":0714
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
      Left            =   4845
      TabIndex        =   17
      Top             =   480
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
      Left            =   2550
      TabIndex        =   16
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label lblSupplier 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<ALL>"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1830
      TabIndex        =   14
      Top             =   1245
      Width           =   4290
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4380
      TabIndex        =   3
      Top             =   180
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
      Height          =   255
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSalesBySupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar As arSalesSummary1
Dim rs As ADODB.Recordset
Dim oRPT As New z_reports
Dim lngTPSID As Long
Dim lngTPCID As Long
Dim bNP As Boolean
Dim strTPName As String

Private Sub cmdAll_Click()
    lngTPSID = 0
    lblSupplier.Caption = "<ALL>"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim dteLimitToView As Date
Dim oSQL As z_SQL
    On Error GoTo errHandler
    If oPC.BlindCashup = True Then
        Set oSQL = New z_SQL
        dteLimitToView = oSQL.GetDateOfEarliestUnSignedSession
        If Me.dtpTo >= StartOfDay(dteLimitToView) Then
            MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
'    If oPC.BlindCashup Then
'        Set oSQL = New z_SQL
'        dteLastAllowableSessionDate = oSQL.GetDateOfMostRecentFullySignedOffDailySession(Me.dtpFrom)
'        If Me.dtpTo > dteLastAllowableSessionDate Then
'            MsgBox "There are unsigned cash ups prior to your selected end date. These cannot be included in the report. Select an earlier end date.", vbInformation, "Can't do this"
'            Exit Sub
'        End If
'    End If
    oRPT.SalesSummary1 rs, Me.dtpFrom.Value, Me.dtpTo.Value, lngTPCID, lngTPSID, bNP
    Me.MousePointer = vbDefault
    
    arSalesBySupplierViewer.ReportSource = Nothing
    
    Set ar = Nothing
    Set ar = New arSalesSummary1
    ar.Visible = False
    If bNP Then ar.Sections(4).NewPage = ddNPAfter
    ar.Component "Sales summary by supplier", rs, Me.dtpFrom.Value, Me.dtpTo.Value
    arSalesBySupplierViewer.ReportSource = ar
    
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesBySupplier.cmdOK_Click"
End Sub

Private Sub cmdSelectTPS_Click()
Dim frmS As frmBrowseSUppliers2
        Set frmS = New frmBrowseSUppliers2
        frmS.Show vbModal
        lngTPSID = frmS.SupplierID
        strTPName = frmS.SupplierName
        Unload frmS
    lblSupplier.Caption = strTPName
    If lngTPSID = 0 Then Exit Sub

End Sub

Private Sub Form_Load()
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
End Sub



Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enExcel
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arSalesBySupplierViewer.Width = Me.Width - 600
    lngDiff = arSalesBySupplierViewer.Height
    arSalesBySupplierViewer.Height = Me.Height - 2700
    lngDiff = arSalesBySupplierViewer.Height - lngDiff
    cmdToExcel.left = arSalesBySupplierViewer.left + arSalesBySupplierViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arSalesBySupplierViewer.left + arSalesBySupplierViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

