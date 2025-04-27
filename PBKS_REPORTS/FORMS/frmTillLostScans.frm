VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTillLostScans 
   Caption         =   "Lost scans"
   ClientHeight    =   8475.001
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475.001
   ScaleWidth      =   13125
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   630
      Left            =   8160
      Picture         =   "frmTillLostScans.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   405
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
      Left            =   7140
      Picture         =   "frmTillLostScans.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   420
      Width           =   1000
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   660
      Width           =   1380
   End
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   9930.001
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   660
      Width           =   1380
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2385
      TabIndex        =   4
      Top             =   420
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   103153667
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4590
      TabIndex        =   5
      Top             =   420
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   103153667
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arV 
      Height          =   6780
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   11959
      SectionData     =   "frmTillLostScans.frx":0714
   End
   Begin MSComctlLib.ListView lvwTills 
      Height          =   855
      Left            =   90
      TabIndex        =   9
      Top             =   405
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1508
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Station"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Workstations"
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
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   150
      Width           =   1470
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4170
      TabIndex        =   7
      Top             =   480
      Width           =   555
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select period between"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3150
      TabIndex        =   6
      Top             =   120
      Width           =   2160
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmTillLostScans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strType As String
Dim bUseLDP As Boolean
Dim ar As arTillLostScans
Dim oRPT As New z_reports
Dim rsLostScans As ADODB.Recordset


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
10        On Error GoTo errHandler
      Dim msg As String
      Dim oSQL As z_SQL
      Dim Status As Integer
20        Set oRPT = New z_reports
30        Set rsLostScans = Nothing
40        Set rsLostScans = New ADODB.Recordset
50        rsLostScans.CursorLocation = adUseClient
60        Screen.MousePointer = vbHourglass
70        oRPT.TillLostScans rsLostScans, dtpFrom, dtpTo, lvwTills.SelectedItem.Text, Status
80        Screen.MousePointer = vbDefault
90        If Status = False Then
100           MsgBox "The till point " & lvwTills.SelectedItem.Text & " cannot be found. It may be switched off or offline. " & vbCrLf & "Correct and re-try.", vbInformation + vbOKOnly, "Cannot connect to till"
110           Exit Sub
120       End If
130       Set ar = Nothing
140       Set ar = New arTillLostScans
          
150       arV.ReportSource = ar
160       msg = "Till lost scans from " & Format(dtpFrom, "dd/mm/yyyy") & " to " & Format(dtpTo, "dd/mm/yyyy")
170       ar.Component rsLostScans, msg, dtpFrom, dtpTo
180       Exit Sub
errHandler:
190       If ErrMustStop Then Debug.Assert False: Resume
200       ErrorIn "frmTIllJournal.cmdOK_Click"
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
    LoadlistviewTills GetWorkstations
    lvwTills.ListItems(1).Selected = True

    dtpFrom.Value = FirstOfMonth(Date)
    dtpTo.Value = LastOfMonth(Date)
End Sub
Private Sub LoadlistviewTills(rs As ADODB.Recordset)
Dim lstItem As ListItem
Dim i As Integer

    Do While Not rs.EOF
        Set lstItem = lvwTills.ListItems.Add
        With lstItem
            .Text = FNS(rs.Fields(2))
            .Key = FNS(rs.Fields(1))
        End With
        rs.MoveNext
    Loop
    
End Sub
Private Function GetWorkstations() As ADODB.Recordset
Dim OpenResult As Integer
Dim rs As New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tPOSCLIENT", oPC.COShort, adOpenUnspecified, adLockUnspecified
    Set rs.ActiveConnection = Nothing
    Set GetWorkstations = rs
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function

Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arV.Width = Me.Width - 600
    lngDiff = arV.Height
    arV.Height = Me.Height - 1500
    lngDiff = arV.Height - lngDiff
    cmdToExcel.Left = arV.Left + arV.Width - cmdToExcel.Width
    cmdToPDF.Left = arV.Left + arV.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

