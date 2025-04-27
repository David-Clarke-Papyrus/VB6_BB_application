VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesDetails 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales details"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
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
   ScaleWidth      =   12960
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkIncludeAppros 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Include appros"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   615
      TabIndex        =   9
      Top             =   555
      Width           =   1785
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   10950
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   330
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   330
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
      Left            =   6975
      Picture         =   "frmSalesDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   105
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   8010
      Picture         =   "frmSalesDetails.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   105
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2595
      TabIndex        =   1
      Top             =   135
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   193658883
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
      Format          =   193658883
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arSalesDetailViewer 
      Height          =   7305
      Left            =   105
      TabIndex        =   6
      Top             =   945
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   12885
      SectionData     =   "frmSalesDetails.frx":0714
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
      TabIndex        =   11
      Top             =   495
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
      TabIndex        =   10
      Top             =   495
      Width           =   1800
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
Attribute VB_Name = "frmSalesDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim bCancel As Boolean
Dim ar As arSalesDetail
Dim rs As ADODB.Recordset
Dim oRPT As New z_reports
Public enPrevPrintCSV As enumReportPresentation

Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property

Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
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
    Set rs = New ADODB.Recordset
    oRPT.SalesDetail rs, dtpFrom, dtpTo, IIf(Me.chkIncludeAppros = 1, True, False)
    
    Set ar = Nothing
    Set ar = New arSalesDetail
    ar.Visible = False
    
    Set arSalesDetailViewer.ReportSource = ar
    ar.Component rs, dtpFrom, dtpTo
    SaveSetting "PBKS", "ReportsSettings", "IncludeAppros_1", CStr(Me.chkIncludeAppros)
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesDetails.cmdOK_Click"
End Sub

Private Sub Form_Load()
    
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
    Me.chkIncludeAppros = CLng(GetSetting("PBKS", "ReportsSettings", "IncludeAppros_1", "0"))
    
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
    If TryToDeleteFile(fn) = False Then
        Exit Sub
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
    If TryToDeleteFile(fn) = False Then
        Exit Sub
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enExcel
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arSalesDetailViewer.Width = Me.Width - 600
    lngDiff = arSalesDetailViewer.Height
    arSalesDetailViewer.Height = Me.Height - 1500
    lngDiff = arSalesDetailViewer.Height - lngDiff
    cmdToExcel.left = arSalesDetailViewer.left + arSalesDetailViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arSalesDetailViewer.left + arSalesDetailViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

