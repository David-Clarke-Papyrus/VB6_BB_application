VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Begin VB.Form frmCreateCategoryCheck 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Create category checks"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   14640
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   8910
      Picture         =   "frmCreateCategoryCheck.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   150
      Width           =   1000
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   8850
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   915
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   10275
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   915
      Width           =   1380
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arViewer 
      Height          =   4890
      Left            =   135
      TabIndex        =   4
      Top             =   1425
      Width           =   14400
      _ExtentX        =   25400
      _ExtentY        =   8625
      SectionData     =   "frmCreateCategoryCheck.frx":038A
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   8445
      Begin VB.CheckBox chkFilter 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Show where selection is the only category"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   855
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.ComboBox cboMB 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3945
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   465
         Width           =   2070
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   210
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   465
         Width           =   3240
      End
      Begin VB.CommandButton cmdMasterList 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Create category check"
         Height          =   435
         Left            =   6285
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   405
         Width           =   1965
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Multibuy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   3870
         TabIndex        =   9
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   390
         TabIndex        =   3
         Top             =   195
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmCreateCategoryCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oSM As New z_StockManager
Dim oSQL As z_SQL
Dim rpt As arCategoryCheck
Dim lngCatChkID As Long
Dim sMB As String
Dim tlMB As New z_TextList
Dim tlCat As New z_TextList

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreateCategoryCheck.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMasterList_Click()
    On Error GoTo errHandler
    If Not SecurityControlforSupervisor Then
        Exit Sub
    End If
    lngCatChkID = oSM.GenerateCategoryCheck(tlCat.Key(cboSection), gSTAFFID, FNB(Me.chkFilter = 1), FNN(tlMB.Key(cboMB)))
    sMB = tlMB.Key(cboMB)
    
    Set rpt = Nothing
    Set oSQL = New z_SQL
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    oSQL.CategoryCheck rs, lngCatChkID
    Set rpt = New arCategoryCheck
    If tlCat.Key(cboSection) > 0 Then
        rpt.component "Category Check: " & cboSection, rs
    ElseIf FNN(tlMB.Key(cboMB)) > 0 Then
        rpt.component "Category Check: Multibuys: " & cboMB, rs
    End If
    
    Me.arViewer.ReportSource = rpt
    rpt.Printer.PaperSize = 9
    rpt.Printer.Orientation = ddOLandscape
    rpt.PageSettings.LeftMargin = 700
    rpt.PageSettings.RightMargin = 0
    Screen.MousePointer = vbDefault
   
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreateCategoryCheck.cmdMasterList_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    rpt.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "CategoryChecks" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel

End Sub

Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    rpt.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "CategoryChecks" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
    If fs.FileExists(fn) Then
        fs.DeleteFile (fn)
    End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enPDF

End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    tlMB.Load ltMultibuys, , "<any>"
    tlCat.Load ltSectionsActive, , "<any>"
    LoadCombo cboSection, tlCat
    LoadCombo cboMB, tlMB
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreateCategoryCheck.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
Dim lngDiffH As Long
  '  SSTab1.Width = Me.Width - 500
  '  lngDiff = SSTab1.Height
  '  SSTab1.Height = Me.Height - (SSTab1.top + 800)
  '  lngDiff = SSTab1.Height - lngDiff
    arViewer.Width = NonNegative_Lng(Me.Width - 500)
    arViewer.Height = NonNegative_Lng(Me.Height - 2500)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCreateCategoryCheck.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Label1_Click()
    cboSection.text = "<any>"
End Sub

Private Sub Label2_Click()
    cboMB.text = "<any>"
End Sub
