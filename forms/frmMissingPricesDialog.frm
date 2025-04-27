VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmMissingPricesDialog 
   Caption         =   "Missing prices"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16890
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   16890
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSalesInLastMonths 
      Alignment       =   2  'Center
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   3840
      TabIndex        =   15
      Text            =   "0"
      Top             =   135
      Width           =   285
   End
   Begin VB.CheckBox chkRRP 
      Caption         =   "RR price"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   600
      Width           =   1215
   End
   Begin VB.CheckBox chkSP 
      Caption         =   "Selling price"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   360
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkCostPrice 
      Caption         =   "Cost price"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   345
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   15135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   345
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arMissingViewer 
      Height          =   6705
      Left            =   255
      TabIndex        =   8
      Top             =   900
      Width           =   16275
      _ExtentX        =   28707
      _ExtentY        =   11827
      SectionData     =   "frmMissingPricesDialog.frx":0000
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00F2E0D9&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   600
      Left            =   10800
      Picture         =   "frmMissingPricesDialog.frx":003C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00F2E0D9&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   600
      Left            =   9810
      Picture         =   "frmMissingPricesDialog.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   135
      Width           =   945
   End
   Begin VB.TextBox txtMinimumPrice 
      Height          =   345
      Left            =   6435
      TabIndex        =   3
      Text            =   "0"
      Top             =   270
      Width           =   795
   End
   Begin VB.CheckBox chkAnyQtyOH 
      Alignment       =   1  'Right Justify
      Caption         =   "Any quantity on hand"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   585
      TabIndex        =   2
      Top             =   540
      Width           =   1815
   End
   Begin VB.TextBox txtQtyOnHand 
      Height          =   345
      Left            =   1605
      TabIndex        =   0
      Text            =   "0"
      Top             =   180
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "months"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4230
      TabIndex        =   16
      Top             =   165
      Width           =   585
   End
   Begin VB.Label Label4 
      Caption         =   "Sales in the last"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   2580
      TabIndex        =   14
      Top             =   165
      Width           =   1170
   End
   Begin VB.Label Label3 
      Caption         =   "cents"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   7275
      TabIndex        =   7
      Top             =   300
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Look for prices <"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   5100
      TabIndex        =   4
      Top             =   315
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Quantity on hand >"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   165
      TabIndex        =   1
      Top             =   225
      Width           =   1485
   End
End
Attribute VB_Name = "frmMissingPricesDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim mQtyOH As Long
Dim mMinPrice As Long
Dim mCancelled As Boolean
Dim ar As arMissingPrices

Public Property Get IsCancelled() As Boolean
    IsCancelled = mCancelled
End Property
Public Property Get QtyOH() As Long
    QtyOH = mQtyOH
End Property
Public Property Get MinimumPrice() As Long
    MinimumPrice = mMinPrice
End Property

Private Sub chkAnyQtyOH_Click()
    If chkAnyQtyOH = 1 Then
        Me.txtQtyOnHand = ""
        mQtyOH = -9999999
    Else
        Me.txtQtyOnHand = "0"
        mQtyOH = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    'mCancelled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim oRPT As New z_reports
Dim lngXMonths As Long
    Screen.MousePointer = vbHourglass
    If IsNumeric(txtSalesInLastMonths) Then
        lngXMonths = CLng(txtSalesInLastMonths)
    Else
        lngXMonths = 0
    End If
    oRPT.MissingPrices mQtyOH, mMinPrice, (Me.chkCostPrice.Value = 1), (Me.chkSP.Value = 1), (chkRRP.Value = 1), lngXMonths, rs
    Set ar = New arMissingPrices
    ar.Component rs
    arMissingViewer.ReportSource = ar
    Screen.MousePointer = vbDefault
    

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
    arMissingViewer.Width = Me.Width - 600
    lngDiff = arMissingViewer.Height
    arMissingViewer.Height = Me.Height - 1800
    lngDiff = arMissingViewer.Height - lngDiff
    cmdToExcel.left = arMissingViewer.left + arMissingViewer.Width - cmdToExcel.Width
    cmdToPDF.left = arMissingViewer.left + arMissingViewer.Width - cmdToExcel.Width - cmdToPDF.Width - 10

End Sub

Private Sub txtMinimumPrice_Change()
    If IsNumeric(txtMinimumPrice) Then
        mMinPrice = CLng(txtMinimumPrice)
    End If
End Sub

Private Sub txtMinimumPrice_Validate(Cancel As Boolean)
    If Not IsNumeric(txtMinimumPrice) Then Cancel = True
End Sub

Private Sub txtQtyOnHand_Change()
    If txtQtyOnHand = "" Then
        mMinPrice = -9999999
    Else
        mMinPrice = CLng(txtQtyOnHand)
    End If
End Sub

Private Sub txtQtyOnHand_Validate(Cancel As Boolean)
    If Not IsNumeric(txtQtyOnHand) Then Cancel = True
End Sub
