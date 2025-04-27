VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmLastStockTakeList 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Last stock take list"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12960
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
   ScaleHeight     =   8460
   ScaleWidth      =   12960
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   10950
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   330
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   330
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
      Left            =   90
      Picture         =   "frmLastStockTakeList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   1125
      Picture         =   "frmLastStockTakeList.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   1000
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arVWR 
      Height          =   7305
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   12885
      SectionData     =   "frmLastStockTakeList.frx":0714
   End
End
Attribute VB_Name = "frmLastStockTakeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim bCancel As Boolean
Dim rs As ADODB.Recordset
Dim oRPT As New z_reports
Dim ar As New arStockTakeList


Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset

    
    oRPT.LastStockTakeList False, rs
    
    Set ar = Nothing
    Set ar = New arStockTakeList
    ar.Visible = False
    
    Set arVWR.ReportSource = ar
    ar.Component rs, False
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmLastStockTakeList.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get CancelReport() As Boolean
    CancelReport = bCancel
End Property
Public Property Get Preview() As Boolean
    Preview = bPreview
End Property

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
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
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
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enExcel
End Sub

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arVWR.Width = Me.Width - 600
    lngDiff = arVWR.Height
    arVWR.Height = Me.Height - 1500
    lngDiff = arVWR.Height - lngDiff
    cmdToExcel.left = arVWR.left + arVWR.Width - cmdToExcel.Width
    cmdToPDF.left = arVWR.left + arVWR.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

