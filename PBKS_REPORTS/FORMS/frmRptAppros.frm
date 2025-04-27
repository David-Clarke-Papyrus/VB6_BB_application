VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAppros 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Outstanding Appros"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
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
   ScaleHeight     =   9180
   ScaleWidth      =   13215
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   10020
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1215
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1215
      Width           =   1380
   End
   Begin VB.CommandButton cmdSelectTPS 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Select customer"
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1515
      Width           =   1470
   End
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
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1545
      Width           =   555
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Style"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   5550
      TabIndex        =   10
      Top             =   60
      Width           =   5175
      Begin VB.OptionButton optLDC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses last delivered cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   12
         Top             =   315
         Width           =   4785
      End
      Begin VB.OptionButton optWeighted 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses weighted average cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   11
         Top             =   645
         Value           =   -1  'True
         Width           =   4785
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   11880
      Picture         =   "frmRptAppros.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   210
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
      Left            =   10845
      Picture         =   "frmRptAppros.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   210
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Appros issued"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1350
      Left            =   150
      TabIndex        =   6
      Top             =   60
      Width           =   5400
      Begin VB.OptionButton optBetween 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Between"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   135
         TabIndex        =   2
         Top             =   810
         Width           =   1035
      End
      Begin VB.OptionButton optPriorTo 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Prior to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   315
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker dtpApproPriorTo 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   315
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   193658881
         CurrentDate     =   37421
      End
      Begin MSComCtl2.DTPicker dtpApproDate1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   193658881
         CurrentDate     =   37421
      End
      Begin MSComCtl2.DTPicker dtpApproDate2 
         Height          =   375
         Left            =   3135
         TabIndex        =   4
         Top             =   810
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   193658881
         CurrentDate     =   37421
      End
      Begin VB.Label Label1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "(inclusive)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   4515
         TabIndex        =   8
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label19 
         BackColor       =   &H00D3D3CB&
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   2640
         TabIndex        =   7
         Top             =   855
         Width           =   555
      End
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arApprosViewer 
      Height          =   6765
      Left            =   90
      TabIndex        =   17
      Top             =   1980
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   11933
      SectionData     =   "frmRptAppros.frx":0714
   End
   Begin VB.Label lblCustomer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<ALL>"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   1590
      TabIndex        =   18
      Top             =   1590
      Width           =   4290
   End
End
Attribute VB_Name = "frmAppros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim WithEvents oRpts As z_reports
Attribute oRpts.VB_VarHelpID = -1
Dim oTxtList As z_TextList
Dim strTPName As String
Dim lngTPID As Long
Dim enPrevPrintCSV As enumReportPresentation
Dim strCostSetting As String
Dim ar As arAppro
Dim rs As ADODB.Recordset

Public Property Get ReportPresentation() As enumReportPresentation
    On Error GoTo errHandler
    ReportPresentation = enPrevPrintCSV
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.ReportPresentation"
End Property

Public Sub Component(pType As String)
    On Error GoTo errHandler
    strType = pType
    If pType = "ALL" Then
        Me.Caption = "All appros issued (whether returned or not)"
    ElseIf pType = "OS" Then
        Me.Caption = "Outstanding appros"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.Component(pType)", pType
End Sub

Private Sub cmdAll_Click()
    On Error GoTo errHandler
    lngTPID = 0
    lblCustomer.Caption = "<ALL>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.cmdAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim Date1 As Date
Dim Date2 As Date
Dim blnPrint As Boolean
Dim blnNoRecordsReturned As Boolean
Dim strErrMsg As String
Dim strTitle As String
Dim strFooter As String
Dim strCustname As String
    
    strCostSetting = IIf(Me.optWeighted = True, "W", "LDC")
    SaveSetting "PBKS", "Reports", "OSApprosCost", strCostSetting
    If Me.optPriorTo.Value = True Then
        Date1 = dtpApproPriorTo.Value
    ElseIf optBetween.Value = True Then
        Date1 = dtpApproDate1.Value
        Date2 = dtpApproDate2.Value
    End If
    
    If strType = "ALL" Then
        strErrMsg = oRpts.ApprosIssued(rs, lngTPID, Date1, Date2, strTPName, strTitle, IIf(optWeighted = True, "W", "LD"))
    Else
        strErrMsg = oRpts.ApprosOS(rs, lngTPID, Date1, Date2, strTPName, strTitle, IIf(optWeighted = True, "W", "LD"))
                
    End If
    Me.MousePointer = vbDefault
    arApprosViewer.ReportSource = Nothing
    

    Set ar = Nothing
    Set ar = New arAppro
    ar.Visible = False
    ar.Component rs, strTitle, "FOOTOER", optWeighted = False
    arApprosViewer.ReportSource = ar

EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdSelectCustomer_Click()
'    On Error GoTo errHandler
'Dim frm As frmBrowseCustomers2
'    Set frm = New frmBrowseCustomers2
'    frm.Show vbModal
'    lngTPID = frm.CustomerID
'    strCustomerName = left(frm.CustomerName, 40) & IIf(frm.Accnum > "", " (" & frm.Accnum & ")", "")
'    Me.lblCustomer.Caption = strCustomerName
'    Unload frm
'    If lngTPID > 0 Then Me.chkApproAll = 0
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmAppros.cmdSelectCustomer_Click"
'End Sub

Private Sub cmdSelectTPS_Click()
    On Error GoTo errHandler
Dim frmC As frmBrowseCustomers2
        Set frmC = New frmBrowseCustomers2
        frmC.Show vbModal
        lngTPID = frmC.CustomerID
        strTPName = frmC.CustomerName
        Unload frmC
    lblCustomer.Caption = strTPName
    If lngTPID = 0 Then Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.cmdSelectTPS_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    
    lngTPID = 0
    Set oRpts = New z_reports
    Set oTxtList = New z_TextList
    optPriorTo.Value = True
    dtpApproPriorTo.Value = DateAdd("m", -1, Date)
    dtpApproDate1.Value = DateAdd("m", -2, Date)
    dtpApproDate2.Value = DateAdd("m", -1, Date)
    strCostSetting = GetSetting("PBKS", "Reports", "OSApprosCost", "W")
    If strCostSetting = "W" Then
        Me.optWeighted = True
    Else
        Me.optLDC = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    Set oRpts = Nothing
    Set oTxtList = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToPDF_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "Appros" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.cmdToPDF_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "Appros" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAppros.cmdToExcel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arApprosViewer.Width = Me.Width - 600
    lngDiff = arApprosViewer.Height
    arApprosViewer.Height = Me.Height - 2700
    lngDiff = arApprosViewer.Height - lngDiff
    cmdToExcel.left = arApprosViewer.left + arApprosViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arApprosViewer.left + arApprosViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

