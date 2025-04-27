VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesSummaryPerProduct 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Sales summary per product"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14730
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
   ScaleHeight     =   9885
   ScaleWidth      =   14730
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C4BCA4&
      Caption         =   "All"
      Height          =   450
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1185
      Width           =   585
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   13125
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1230
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   11730
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1230
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arSalesSummary2Viewer 
      Height          =   8040
      Left            =   270
      TabIndex        =   13
      Top             =   1680
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   14182
      SectionData     =   "frmSalesSummaryPerProduct.frx":0000
   End
   Begin VB.Frame frCost 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Costing"
      ForeColor       =   &H8000000D&
      Height          =   1065
      Left            =   6690
      TabIndex        =   10
      Top             =   30
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
   Begin VB.CommandButton cmdSelectTPC 
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
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1470
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
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
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
      Left            =   11385
      Picture         =   "frmSalesSummaryPerProduct.frx":003C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   12420
      Picture         =   "frmSalesSummaryPerProduct.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1000
   End
   Begin VB.CheckBox chkNP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "New page per section"
      ForeColor       =   &H8000000D&
      Height          =   405
      Left            =   4425
      TabIndex        =   4
      Top             =   465
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   255
      TabIndex        =   1
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   193658881
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   405
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      Format          =   193658881
      CurrentDate     =   37421
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
      Left            =   2505
      TabIndex        =   18
      Top             =   780
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
      Left            =   210
      TabIndex        =   17
      Top             =   780
      Width           =   1800
   End
   Begin VB.Label txtTP 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Left            =   1755
      TabIndex        =   8
      Top             =   1230
      Width           =   3660
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   2040
      TabIndex        =   3
      Top             =   465
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
      Top             =   75
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSalesSummaryPerProduct"
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
Dim rs As ADODB.Recordset
Dim ar As New arSalesSummary2

Public Sub Component(pMsg As String, pOneOrTwoDates As Integer, pDefaultDate1 As Date, pDefaultDate2 As Date, pCustomerOrSupplier As String, Optional pShowchkNP As Boolean)
    strCustomerOrSupplier = pCustomerOrSupplier
    cmdSelectTPS.Visible = (strCustomerOrSupplier = "S" Or strCustomerOrSupplier = "B")
    cmdSelectTPC.Visible = (strCustomerOrSupplier = "C" Or strCustomerOrSupplier = "B")
    lblDescription.Caption = pMsg
    Me.chkNP.Visible = pShowchkNP
    Me.dtpFrom.Value = pDefaultDate1
    If pDefaultDate2 <> CDate(0) Then dtpTo.Value = pDefaultDate2
    If pOneOrTwoDates = 1 Then
        dtpTo.Visible = False
        lblAnd.Visible = False
    End If
        
End Sub
Public Sub Component2(Optional pShowchkNP As Boolean)
    Me.chkNP.Visible = pShowchkNP
End Sub

Private Sub cmdAll_Click()
    lngTPCID = 0
    lngTPSID = 0
    txtTP = "<ALL>"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim dteLimitToView As Date
Dim oSQL As z_SQL
    On Error GoTo errHandler
    
    lngTPCID = Me.TPCID
    lngTPSID = Me.TPSID
    bNP = Me.NP
    
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
    Screen.MousePointer = vbHourglass
    Set oRPT = New z_reports
    oRPT.SalesSummary3 rs, Me.dtpFrom, Me.dtpTo, lngTPCID, lngTPSID, bNP
    Set oRPT = Nothing
    
    Set ar = New arSalesSummary2
    ar.Component "", rs, Me.dtpFrom, Me.dtpTo
    If bNP Then ar.Sections(4).NewPage = ddNPAfter
    arSalesSummary2Viewer.ReportSource = ar
    Screen.MousePointer = vbDefault

EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmPeriodDialogue.cmdOK_Click"
End Sub

'Private Sub cmdSelectTP_Click()
'Dim frm As frmBrowseTPs2
'    Set frm = New frmBrowseTPs2
'    frm.Show vbModal
'    lngTPID = frm.TPID
'    strTPName = frm.TPName
'    txtTP.Caption = strTPName
'    Unload frm
'    If lngTPID = 0 Then Exit Sub
'
'
'End Sub
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

Private Sub Form_Load()
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
        txtTP = "<ALL>"

End Sub


Public Property Get NP() As Boolean
    NP = (chkNP = 1)
End Property


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arSalesSummary2Viewer.Width = Me.Width - 600
    lngDiff = arSalesSummary2Viewer.Height
    arSalesSummary2Viewer.Height = Me.Height - 1800
    lngDiff = arSalesSummary2Viewer.Height - lngDiff
    cmdToExcel.left = arSalesSummary2Viewer.left + arSalesSummary2Viewer.Width - cmdToExcel.Width
    cmdToPDF.left = arSalesSummary2Viewer.left + arSalesSummary2Viewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub
