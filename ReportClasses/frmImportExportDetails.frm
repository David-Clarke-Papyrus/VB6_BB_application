VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmImportExportDetails 
   Caption         =   "Import/export details"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18645
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   18645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   195
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   1725
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   195
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arvImportExportDetails 
      Height          =   8055
      Left            =   330
      TabIndex        =   0
      Top             =   660
      Width           =   17985
      _ExtentX        =   31724
      _ExtentY        =   14208
      SectionData     =   "frmImportExportDetails.frx":0000
   End
End
Attribute VB_Name = "frmImportExportDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arIED As arImportExportDetails

Public Sub Component(pHEading As String, pRS As adodb.Recordset)
    Set arIED = New arImportExportDetails
    arIED.PageSettings.Orientation = ddOLandscape
    arIED.Component pHEading, pRS
    
    arvImportExportDetails.ReportSource = arIED
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdToPDF_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If arIED Is Nothing Then Exit Sub
    If arIED.Pages.Count = 0 Then Exit Sub
    arIED.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "ImportExportDetails" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.Filename = fn
    Call pdfExpt.Export(arIED.Pages)
    OpenFileWithApplication fn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmImportExportDetails.cmdToPDF_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If arIED Is Nothing Then Exit Sub
    If arIED.Pages.Count = 0 Then Exit Sub
    arIED.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "IMportExportDetails" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.Filename = fn
    Call ExcelExpt.Export(arIED.Pages)
    OpenFileWithApplication fn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmIMportExportDetails.cmdToExcel__Click", , EA_NORERAISE
    HandleError
End Sub
'''''''''''''''''''''

Private Sub Form_Resize()
    Me.arvImportExportDetails.Width = Me.Width - 800
    Me.arvImportExportDetails.Height = Me.Height - 1300
End Sub
