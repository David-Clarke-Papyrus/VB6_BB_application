VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmCOI 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Value of inventory"
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
   Begin VB.Frame frmStyle 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Style"
      ForeColor       =   &H8000000D&
      Height          =   525
      Left            =   210
      TabIndex        =   10
      Top             =   60
      Visible         =   0   'False
      Width           =   2640
      Begin VB.OptionButton optWithStockTurn 
         BackColor       =   &H00D3D3CB&
         Caption         =   "with S/T"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1410
         TabIndex        =   12
         Top             =   225
         Width           =   1095
      End
      Begin VB.OptionButton optStd 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Default"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   225
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Breakdown by"
      ForeColor       =   &H8000000D&
      Height          =   690
      Left            =   3030
      TabIndex        =   6
      Top             =   60
      Width           =   4155
      Begin VB.OptionButton optPublisher 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Publisher"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   2850
         TabIndex        =   9
         Top             =   330
         Width           =   1110
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Category"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1575
         TabIndex        =   8
         Top             =   330
         Width           =   1485
      End
      Begin VB.OptionButton optPT 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Product type"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   90
         TabIndex        =   7
         Top             =   330
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arCOIViewer 
      Height          =   6975
      Left            =   180
      TabIndex        =   3
      Top             =   1155
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   12303
      SectionData     =   "frmCOI.frx":0000
   End
   Begin VB.CheckBox chkLDP 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Use last delivered cost (not weighted average)"
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   195
      TabIndex        =   2
      Top             =   630
      Width           =   2685
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
      Left            =   7245
      Picture         =   "frmCOI.frx":003C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   615
      Left            =   8415
      Picture         =   "frmCOI.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1000
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Note that negative qty on hand figures are adjusted to 0 for the purposes of this report."
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
      Height          =   240
      Left            =   2970
      TabIndex        =   13
      Top             =   900
      Width           =   8580
   End
End
Attribute VB_Name = "frmCOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim bUseLDP As Boolean
Dim ar As Object
Dim oRPT As New z_reports
Dim rs As ADODB.Recordset

Public Property Get UseLDP() As Boolean
    UseLDP = bUseLDP
End Property


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim ByWhat As String

    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    If Me.optPT = True Then
        ByWhat = "PT"
    ElseIf Me.optCategory = True Then
        ByWhat = "C"
    ElseIf Me.optPublisher = True Then
        ByWhat = "P"
    End If
    Set rs = New ADODB.Recordset
    oRPT.COI rs, True, Me.UseLDP, ByWhat
    
'    If optWithStockTurn = 1 Then
'        Set ar = New arCOI_WithTurn
'    Else
        Set ar = New arCOI
'    End If
    
    arCOIViewer.ReportSource = ar
    ar.Component rs, "All prices and costs are Ex VAT", bUseLDP, ByWhat
    Screen.MousePointer = vbDefault
EXIT_Handler:
    Screen.MousePointer = vbDefault
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCOI.cmdOK_Click"
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
    arCOIViewer.Width = Me.Width - 600
    lngDiff = arCOIViewer.Height
    arCOIViewer.Height = Me.Height - 1500
    lngDiff = arCOIViewer.Height - lngDiff
    cmdToExcel.left = arCOIViewer.left + arCOIViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arCOIViewer.left + arCOIViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub
