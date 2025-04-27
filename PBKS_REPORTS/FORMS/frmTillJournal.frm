VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTrackingNegativeStock 
   Caption         =   "Negative stock"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   13080
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   615
      Left            =   8115
      Picture         =   "frmTillJournal.frx":0000
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
      Left            =   6945
      Picture         =   "frmTillJournal.frx":038A
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
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   62521347
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4515
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   62521347
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arvNeg 
      Height          =   7305
      Left            =   165
      TabIndex        =   8
      Top             =   825
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   12885
      SectionData     =   "frmTillJournal.frx":0714
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4095
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
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
      Top             =   45
      Visible         =   0   'False
      Width           =   2160
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTrackingNegativeStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim bUseLDP As Boolean
Dim ar As arNegQtyOnHand
Dim oRPT As New z_reports
Dim rsNeg As ADODB.Recordset

Public Property Get UseLDP() As Boolean
    UseLDP = bUseLDP
End Property


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    
    Set oRPT = New z_reports
    Set rsNeg = New ADODB.Recordset
    oRPT.QtyOnHandNegative rsNeg ', dtpFrom, dtpTo
    Set ar = Nothing
    Set ar = New arNegQtyOnHand
    arvNeg.ReportSource = ar
    ar.Component rsNeg, ""
 '   ar.Show
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
    arvNeg.Width = Me.Width - 600
    lngDiff = arvNeg.Height
    arvNeg.Height = Me.Height - 1500
    lngDiff = arvNeg.Height - lngDiff
    cmdToExcel.left = arvNeg.left + arvNeg.Width - cmdToExcel.Width
    cmdToPDF.left = arvNeg.left + arvNeg.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

