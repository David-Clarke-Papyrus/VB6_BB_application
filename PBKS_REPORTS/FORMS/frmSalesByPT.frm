VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesByPT 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales by product type"
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
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   345
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   11145
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   345
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
      Left            =   6615
      Picture         =   "frmSalesByPT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   90
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   7680
      Picture         =   "frmSalesByPT.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   1000
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
         TabIndex        =   7
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
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   61079555
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   75
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   61079555
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   6615
      Left            =   45
      TabIndex        =   10
      Top             =   870
      Width           =   13980
      _ExtentX        =   24659
      _ExtentY        =   11668
      SectionData     =   "frmSalesByPT.frx":0714
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
Attribute VB_Name = "frmSalesByPT"
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
Dim rpt As New arSalesByPeriod
Dim mPeriod As String


Public Sub Component(pMsg As String, Period As String)
    lblDescription.Caption = pMsg
    mPeriod = Period
        
End Sub
'Public Sub Component2(Optional pShowchkNP As Boolean)
'    Me.chkNP.Visible = pShowchkNP
'End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    Screen.MousePointer = vbHourglass
    
    Set oRPT = New z_reports
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    oRPT.SalesByPeriod rs, dtpFrom, dtpTo, mPeriod
    Set rpt = Nothing
    Set rpt = New arSalesByPeriod
    rpt.Component lblDescription.Caption, rs, Me.dtpFrom, Me.dtpTo
    
    Me.arViewer.ReportSource = rpt
    
    Screen.MousePointer = vbDefault
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmPeriodDialogue.cmdOK_Click"
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
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
'Public Property Get NP() As Boolean
'    NP = (chkNP = 1)
'End Property

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
    cmdToExcel.left = arViewer.left + arViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arViewer.left + arViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10

End Sub
