VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesBySection 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales by category"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13095
   ControlBox      =   0   'False
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
   ScaleHeight     =   8445
   ScaleWidth      =   13095
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkNP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New page per section"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   2325
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9765
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   405
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11220
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   405
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   7305
      Left            =   180
      TabIndex        =   2
      Top             =   825
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   12885
      SectionData     =   "frmSalesBySection.frx":0000
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
      Left            =   6960
      Picture         =   "frmSalesBySection.frx":003C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   615
      Left            =   8130
      Picture         =   "frmSalesBySection.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2325
      TabIndex        =   6
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   58589187
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4530
      TabIndex        =   7
      Top             =   0
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   58589187
      CurrentDate     =   37421
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
      Left            =   15
      TabIndex        =   9
      Top             =   45
      Width           =   2160
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4110
      TabIndex        =   8
      Top             =   60
      Width           =   555
   End
End
Attribute VB_Name = "frmSalesBySection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim bUseLDP As Boolean
Dim ar As arSalesBySectionByDate
Dim oRPT As New z_reports
Dim rs As ADODB.Recordset

Public Property Get UseLDP() As Boolean
    UseLDP = bUseLDP
End Property


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    Set oRPT = New z_reports
    Set rs = New ADODB.Recordset
    oRPT.SalesBySectionByDate rs, dtpFrom, dtpTo
    Set ar = Nothing
    Set ar = New arSalesBySectionByDate
    arViewer.ReportSource = ar
    ar.Component rs, "Sales by section from " & Format(dtpFrom, "dd/mm/yyyy") & " to " & Format(dtpTo, "dd/mm/yyyy"), Format(dtpTo, "dd/mm/yyyy"), "From " & Format(dtpFrom, "dd/mm/yyyy") & " to " & Format(dtpTo, "dd/mm/yyyy")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOI.cmdOK_Click"
End Sub

Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    ar.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fs)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(ar.Pages)
    OpenFileWithApplication fn
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    ar.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fs)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn
End Sub

Private Sub Form_Load()
    dtpFrom.Value = FirstOfMonth(Date)
    dtpTo.Value = LastOfMonth(Date)
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arViewer.Width = Me.Width - 600
    lngDiff = arViewer.Height
    arViewer.Height = Me.Height - 1500
    lngDiff = arViewer.Height - lngDiff
    cmdToExcel.left = arViewer.left + arViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arViewer.left + arViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub
