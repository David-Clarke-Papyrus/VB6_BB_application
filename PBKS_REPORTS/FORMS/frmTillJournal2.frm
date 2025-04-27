VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTillJournal 
   Caption         =   "Till journal"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   13080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   615
      Left            =   8115
      Picture         =   "frmTillJournal2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   150
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
      Left            =   6915
      Picture         =   "frmTillJournal2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   150
      Width           =   1000
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11205
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   405
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   405
      Width           =   1380
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2310
      TabIndex        =   4
      Top             =   165
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   95223811
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4515
      TabIndex        =   5
      Top             =   165
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   95223811
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arV 
      Height          =   7305
      Left            =   165
      TabIndex        =   8
      Top             =   825
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   12885
      SectionData     =   "frmTillJournal2.frx":0714
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4095
      TabIndex        =   7
      Top             =   225
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
      Left            =   0
      TabIndex        =   6
      Top             =   210
      Width           =   2160
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTillJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim bUseLDP As Boolean
Dim ar As arTillJournal
Dim oRPT As New z_reports
Dim rsLostScans As ADODB.Recordset


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim MSG As String
Dim dteLimitToView As Date
Dim oSQL As z_SQL
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

    Set oRPT = New z_reports
    Set rsLostScans = New ADODB.Recordset
    oRPT.TillJournal rsLostScans, dtpFrom, dtpTo
    Set ar = Nothing
    Set ar = New arTillJournal
    
    arV.ReportSource = ar
    MSG = "Till journal from " & Format(dtpFrom, "dd/mm/yyyy") & " to " & Format(dtpTo, "dd/mm/yyyy")
    ar.Component rsLostScans, MSG, dtpFrom, dtpTo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTIllJournal.cmdOK_Click"
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

Private Sub Form_Load()
    dtpFrom.Value = FirstOfMonth(Date)
    dtpTo.Value = LastOfMonth(Date)
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arV.Width = Me.Width - 600
    lngDiff = arV.Height
    arV.Height = Me.Height - 1500
    lngDiff = arV.Height - lngDiff
    cmdToExcel.left = arV.left + arV.Width - cmdToExcel.Width
    cmdToPDF.left = arV.left + arV.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

