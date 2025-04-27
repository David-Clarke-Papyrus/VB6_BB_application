VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgedStock 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Period selection"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
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
   ScaleHeight     =   7440
   ScaleWidth      =   12000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1485
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   8745
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1485
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   5235
      Left            =   135
      TabIndex        =   14
      Top             =   1905
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9234
      SectionData     =   "frmPeriodSupplierDialogue2.frx":0000
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
      Height          =   435
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   90
      Width           =   660
   End
   Begin VB.Frame frCost 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Cost "
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   135
      TabIndex        =   10
      Top             =   750
      Width           =   3900
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
      Begin VB.OptionButton optLDC 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Uses last delivered cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   255
         TabIndex        =   11
         Top             =   315
         Width           =   3555
      End
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
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   1470
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
      Left            =   4155
      Picture         =   "frmPeriodSupplierDialogue2.frx":003C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1230
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   5235
      Picture         =   "frmPeriodSupplierDialogue2.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1230
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   120
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
      Left            =   2340
      TabIndex        =   2
      Top             =   105
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   193658883
      CurrentDate     =   37421
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "later date"
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
      Left            =   2340
      TabIndex        =   9
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "earlier date"
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
      Left            =   105
      TabIndex        =   8
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label txtTP 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<ALL>"
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   6630
      TabIndex        =   7
      Top             =   75
      Width           =   3450
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   1920
      TabIndex        =   3
      Top             =   165
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
      Left            =   4410
      TabIndex        =   0
      Top             =   -885
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAgedStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim bCancel As Boolean
Dim enPrevPrintCSV As enumReportPresentation
Dim lngTPSID As Long
Dim lngTPCID As Long
Dim strTPName As String
Dim strCustomerOrSupplier As String
Dim rs As ADODB.Recordset
Dim ar As New arAgedStock

Private Sub cmdAll_Click()
    txtTP = "<ALL>"
    lngTPSID = 0
End Sub

Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim oRPT As New z_reports
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    oRPT.AgedStock rs, Me.dtpFrom, Me.dtpTo, Me.optLDC, lngTPSID
    
    Set ar = New arAgedStock
    arViewer.ReportSource = ar
    ar.Component rs, Me.dtpFrom, Me.dtpTo, optLDC
   ' Set oRPT = Nothing
    

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmAgedStock.cmdOK_Click"
End Sub

Public Property Get TPCID() As Long
    TPCID = lngTPCID
End Property
Public Property Get TPSID() As Long
    TPSID = lngTPSID
End Property

Private Sub cmdSelectTPS_Click()
Dim frmS As frmBrowseSUppliers2
        Set frmS = New frmBrowseSUppliers2
        frmS.Show vbModal
        lngTPSID = frmS.SupplierID
        strTPName = frmS.SupplierName
        Unload frmS
    txtTP = strTPName
    If lngTPSID = 0 Then Exit Sub

End Sub
Private Sub cmdSelectTPC_Click()
Dim frmC As frmBrowseCustomers2
        Set frmC = New frmBrowseCustomers2
        frmC.Show vbModal
        lngTPCID = frmC.CustomerID
        strTPName = frmC.CustomerName
        Unload frmC
    txtTP = strTPName
    If lngTPCID = 0 Then Exit Sub

End Sub

Private Sub Form_Load()

    
    dtpFrom.Value = FirstOfMonth(DateAdd("M", -6, Date))
    dtpTo.Value = LastOfMonth(DateAdd("M", -4, Date))
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
    arViewer.Width = Me.Width - 600
    lngDiff = arViewer.Height
    arViewer.Height = Me.Height - 2600
    lngDiff = arViewer.Height - lngDiff
    cmdToExcel.left = arViewer.left + arViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arViewer.left + arViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

