VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDebtors_all 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Debtors transactions all"
   ClientHeight    =   10485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
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
   ScaleHeight     =   10485
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   10605
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   570
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   570
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Style"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   4410
      TabIndex        =   11
      Top             =   45
      Width           =   3900
      Begin VB.OptionButton optLDC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses last delivered cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   13
         Top             =   315
         Width           =   3555
      End
      Begin VB.OptionButton optWeighted 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses weighted average cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   12
         Top             =   645
         Value           =   -1  'True
         Width           =   3555
      End
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
      Left            =   8385
      Picture         =   "frmDebtors_All.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   165
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   9420
      Picture         =   "frmDebtors_All.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   165
      Width           =   1000
   End
   Begin VB.CheckBox chkNP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New page per section"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   165
      TabIndex        =   7
      Top             =   885
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
      Left            =   13860
      TabIndex        =   4
      Top             =   375
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
      Left            =   180
      TabIndex        =   1
      Top             =   495
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   193658883
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2385
      TabIndex        =   2
      Top             =   480
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   193658883
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arDebtorsViewer 
      Height          =   7980
      Left            =   165
      TabIndex        =   14
      Top             =   1320
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   14076
      SectionData     =   "frmDebtors_All.frx":0714
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   1965
      TabIndex        =   3
      Top             =   540
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
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDebtors_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim strCostSetting As String
Dim oRPT As New z_reports
Dim rs As ADODB.Recordset
Dim mStyle As String
Dim rpt As arSalesByCustomer1

Public Sub Component(Style As String)
    mStyle = Style
    Select Case mStyle
    Case ""
        Me.Caption = "Debtors' transactions - all"
    Case "CREDIT RETURNS"
        Me.Caption = "Debtors' transactions - credit returns"
    Case "INVOICE SALES"
        Me.Caption = "Debtors' transactions - invoice sales"
    End Select
End Sub
Public Property Get CostSetting() As String
    CostSetting = strCostSetting
End Property

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    strCostSetting = IIf(Me.optWeighted = True, "W", "LPD")
    SaveSetting "PBKS", "Reports", "Debtors_all", strCostSetting
    
   ' Screen.MousePointer = vbHourglass
    
    Set rs = New ADODB.Recordset
    oRPT.SalesByCustomer rs, dtpFrom, dtpTo, mStyle, IIf(Me.optWeighted = True, "W", "LPD")
    Set rpt = Nothing
    Set rpt = New arSalesByCustomer1
    rpt.PageSettings.Orientation = ddOLandscape
    rpt.Component rs, dtpFrom, dtpTo, strCostSetting
    
    arDebtorsViewer.ReportSource = rpt
   
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors_all.cmdOK_Click"
End Sub

Private Sub Form_Load()
    
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
    
    strCostSetting = GetSetting("PBKS", "Reports", "DebtorsAll", "W")
    If strCostSetting = "W" Then
        Me.optWeighted = True
    Else
        Me.optLDC = True
    End If
    
End Sub


Public Property Get FromDate() As Date
    FromDate = dteFrom
End Property
Public Property Get ToDate() As Date
    ToDate = dteTo
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
    fn = oPC.LocalFolder & "\TEMP\" & "DebtorsTrans" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

Private Sub cmdToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If rpt Is Nothing Then Exit Sub
    rpt.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "DebtorsTrans" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDebtors_all.cmdToExcel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arDebtorsViewer.Width = Me.Width - 600
    lngDiff = arDebtorsViewer.Height
    arDebtorsViewer.Height = Me.Height - 1500
    lngDiff = arDebtorsViewer.Height - lngDiff
    cmdToExcel.left = arDebtorsViewer.left + arDebtorsViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arDebtorsViewer.left + arDebtorsViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

