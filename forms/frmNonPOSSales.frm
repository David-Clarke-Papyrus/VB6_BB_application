VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNonPOSSales 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Date selection"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   14205
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   345
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   345
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   6135
      Left            =   120
      TabIndex        =   7
      Top             =   915
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   10821
      SectionData     =   "frmNonPOSSales.frx":0000
   End
   Begin VB.CommandButton cmdGrid1toSpreadsheet 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Send to Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
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
      Left            =   6990
      Picture         =   "frmNonPOSSales.frx":003C
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   105
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   5970
      Picture         =   "frmNonPOSSales.frx":03C6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   105
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2445
      TabIndex        =   0
      Top             =   120
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   221839361
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4470
      TabIndex        =   1
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   221839361
      CurrentDate     =   37421
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   3885
      TabIndex        =   3
      Top             =   150
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Non POS sales between"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   2145
   End
End
Attribute VB_Name = "frmNonPOSSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim bCancelled As Boolean
Dim mdteFrom As Date
Dim mdteTo As Date
Dim x As New XArrayDB
Dim XX As New XArrayDB
Dim rpt As arNonPOSTRansactions


Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
Dim dteFrom As Date
Dim dteTo As Date

    Screen.MousePointer = vbHourglass
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
    rs.open "SELECT * FROM vBackofficeTrading WHERE DTE BETWEEN dbo.startofday('" & ReverseDate(Me.dtpFrom) & "') AND dbo.endofday('" & ReverseDate(Me.dtpTo) & "')  ORDER BY TYP DESC,CODE ASC", oPC.COShort, adOpenForwardOnly
    
    Set rpt = New arNonPOSTRansactions
    rpt.component rs, "Invoices and credit notes issued from  between " & Format(dtpFrom, "DD/MM/YYYY") & " and " & Format(dtpTo, "DD/MM/YYYY")
    Me.arViewer.ReportSource = rpt
   ' Set rs = Nothing
'---------------------------------------------------
   ' If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNonPOSSales.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToPDF_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    rpt.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enPDF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNonPOSSales.cmdToPDF_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    rpt.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNonPOSSales.cmdToExcel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Me.dtpFrom = Date
    Me.dtpTo = Date
    Me.Height = 9000
    Me.Width = 13000
    Me.TOP = 500
    Me.Left = 500
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNonPOSSales.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    arViewer.Width = NonNegative_Lng(Me.Width - 480)
    arViewer.Height = NonNegative_Lng(Me.Height - 1540)
    

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmNonPOSSales.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

