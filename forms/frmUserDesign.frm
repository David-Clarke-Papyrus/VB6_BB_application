VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUserDesign 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Custom reports"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   4485
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Output to"
      ForeColor       =   &H8000000D&
      Height          =   2085
      Left            =   330
      TabIndex        =   3
      Top             =   210
      Width           =   3615
      Begin VB.CommandButton cmdExecute 
         BackColor       =   &H00CDCFAD&
         Caption         =   "Execute report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1170
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "HTML file"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2040
         TabIndex        =   8
         Top             =   690
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton optHTMLfile 
         BackColor       =   &H00D3D3CB&
         Caption         =   "HTML file"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   2040
         TabIndex        =   7
         Top             =   330
         Width           =   1095
      End
      Begin VB.OptionButton optXMLFile 
         BackColor       =   &H00D3D3CB&
         Caption         =   "XML file"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   480
         TabIndex        =   6
         Top             =   1050
         Width           =   1095
      End
      Begin VB.OptionButton optDisplay 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Display"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   480
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optCSVFile 
         BackColor       =   &H00D3D3CB&
         Caption         =   "CSV file"
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   480
         TabIndex        =   4
         Top             =   690
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00CDCFAD&
      Caption         =   "Export CSV data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4245
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00CDCFAD&
      Caption         =   "New  report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2805
      Width           =   1995
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00CDCFAD&
      Caption         =   "Edit report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3525
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog ReportDialog 
      Left            =   2460
      Top             =   2055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Find report"
      Filter          =   "Business Reports Files (*.bre)|*.bre"
   End
End
Attribute VB_Name = "frmUserDesign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim brReportManager As AriacomDll.brReportManager 'Ariacom.brReportManager
Dim reportID As String
Dim strFilename As String
Dim res

Private Sub cmdCreate_Click()
Dim reportTitle As String
On Error GoTo errHandler


    strFilename = ""
    reportID = brReportManager.CreateReport
    cmdEdit_Click

    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.cmdCreate_Click(Index)", EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    top = 500
    left = 300
    Width = 4500
    Height = 5500
    Screen.MousePointer = vbHourglass
    
    Set brReportManager = New AriacomDll.brReportManager
    brReportManager.LoadBusinessDomainFromFile oPC.LocalFolder & "\Aria\PBKS.BDO", "su", ""

    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Function Fixname(pName As String)
    On Error GoTo errHandler
Dim i As Integer
Dim strOut As String
Dim c As String

    strOut = ""
    For i = 1 To Len(pName) - 4
        c = Mid(pName, i, 1)
        If c = "_" Then c = " "
        strOut = strOut & c
    Next
    Fixname = strOut
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.Fixname(pName)", pName
End Function

Private Sub cmdLoad_Click()
    On Error GoTo errHandler

    ReportDialog.DefaultExt = ".bre"
    ReportDialog.DialogTitle = "Find report file"
    ReportDialog.InitDir = oPC.SharedFolderRoot & "\ARIA"
    ReportDialog.CancelError = False
    ReportDialog.ShowOpen
    If ReportDialog.FileName = "" Then Exit Sub

    Screen.MousePointer = vbHourglass
    strFilename = ReportDialog.FileName
    reportID = brReportManager.LoadReportFromFile(ReportDialog.FileName)
    Screen.MousePointer = vbDefault
    cmdEdit_Click

    

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.cmdLoad_Click(Index)", EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExecute_Click()
Dim res
Dim reportID
Dim ret
Dim sFile As String
Dim fs As FileSystemObject
Dim strExecutable As String

    On Error GoTo errHandler
    ReportDialog.DefaultExt = ".bre"
    ReportDialog.DialogTitle = "Find report file"
    ReportDialog.InitDir = oPC.SharedFolderRoot & "\ARIA"
    ReportDialog.CancelError = True
    On Error Resume Next
    ReportDialog.ShowOpen
    If Err = 32755 Then
        Exit Sub
    End If
    If ReportDialog.FileName = "" Then Exit Sub
    
    On Error GoTo errHandler
    reportID = brReportManager.LoadReportFromFile(ReportDialog.FileName)
    reportID = left(reportID, Len(reportID) - 1)
    If Me.optDisplay = True Then
        brReportManager.ExecuteAndViewReport reportID, True
    ElseIf Me.optCSVFile = True Then
        ret = brReportManager.PromptRestrictions(reportID)
        If ret = False Then Exit Sub
        Set fs = New FileSystemObject
        sFile = oPC.SharedFolderRoot & "\ARIA\Reportfiles\" & fs.GetBaseName(ReportDialog.FileName) & Format(Now, "yyyymmyyHHnn") & ".CSV"

        brReportManager.ExecuteReportToOutput reportID, "CSV_Output"
       ' strExecutable = GetPDFExecutable(sfile) & " " & sfile
       ' Screen.MousePointer = vbDefault
        MsgBox "Spreadsheet file saved as: " & oPC.SharedFolderRoot & "\ARIA\Reportfiles\" & fs.GetBaseName(ReportDialog.FileName) & " . . .", vbInformation, "Report finished"
        'Shell strExecutable, vbNormalFocus
    ElseIf Me.optHTMLfile = True Then
        ret = brReportManager.PromptRestrictions(reportID)
        If ret = False Then Exit Sub
        brReportManager.ExecuteReportToOutput reportID, "HTML_Output", True
        MsgBox "Report complete", , "Status"
    ElseIf Me.optXMLFile = True Then
        ret = brReportManager.PromptRestrictions(reportID)
        If ret = False Then Exit Sub
        brReportManager.ExecuteReportToOutput reportID, "XML_Output", True
        MsgBox "Report complete", , "Status"
    End If
    
   ' SetForegroundWindow Me.hwnd
   ' Me.Visible = True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.cmdExecute_Click(Index)", EA_NORERAISE
    HandleError
End Sub
Private Sub cmdExport_Click()
Dim lngResult As Long
Dim rptString
Dim oTF As New z_TextFile
Dim iStart, iEnd As Long
Dim strReplaceable, strReplacing As String
Dim strStartTag, strEndTag As String
Dim strPos As String

    On Error GoTo errHandler
    
    Set fINET = New wininet
    If oPC.InternetDialup = True Then
        lngResult = fINET.StartDUN(0, oPC.Connectionname, True)
    End If
    

    
    strStartTag = "<OutputDest>"
    strEndTag = "</OutputDest>"
    strReplacing = oPC.Configuration.DefaultStore.Code & "_LOY_" & Format(Date, "DDMMYYYY")
    
    ReportDialog.DefaultExt = ".bre"
    ReportDialog.DialogTitle = "Select report file"
    ReportDialog.InitDir = oPC.SharedFolderRoot & "\ARIA"
    ReportDialog.CancelError = False
    ReportDialog.ShowOpen
    If ReportDialog.FileName = "" Then Exit Sub
    Screen.MousePointer = vbHourglass

    oTF.OpenTextFileToRead ReportDialog.FileName
    rptString = oTF.ReadWholeFile
    iStart = InStr(1, rptString, strStartTag)
    iEnd = InStr(1, rptString, strEndTag)
    strReplaceable = Mid(rptString, iStart + Len(strStartTag), iEnd - iStart - Len(strStartTag))
    rptString = Replace(rptString, strReplaceable, strReplacing, 1)

    reportID = brReportManager.LoadReport(rptString)
    brReportManager.SetReportTitle reportID, oPC.Configuration.DefaultStore.Code & "_Loyalty_" & Format(Date, "ddmmyyyy")
    
'MsgBox "Pos 1"
    lngResult = brReportManager.PromptRestrictions(reportID)
    
'MsgBox "Pos 2"
    
    brReportManager.ExecuteReportToOutput reportID, "Delim", 0
    
 '  MsgBox "Pos 5"
    
    fINET.Hangup
    
    Screen.MousePointer = vbDefault
    SetForegroundWindow Me.hwnd
    Me.Visible = True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.cmdExecute_Click(Index)", EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
    
        'Set start step of the editor
        ' 0 Choose Report Type
        ' 1 Choose Elements
        ' 2 Choose Options
        ' 3 Choose Output And Schedule
        Dim startStep As Long
        startStep = 1
        
        'Set visible buttons (bit mask)
        ' 0 &H1 Cancel
        ' 1 &H2 Done
        ' 2 &H4 View Result
        ' 3 &H8 Load From File
        ' 4 &H10 Save To File
        ' 5 &H20 Save Without Execute
        ' 6 &H40 Save And Execute
        Dim visibleButtons As Long
        visibleButtons = 0
        visibleButtons = visibleButtons Xor &H1 'Cancel
        visibleButtons = visibleButtons Xor &H4 'View Result
        visibleButtons = visibleButtons Xor &H20 'Save Without Execute
        visibleButtons = visibleButtons Xor &H40 'Save And Execute
        
        Dim Result As Long
     '   reportID = brReportManager.LoadReportFromFile(oPC.SharedFolderRoot & "\ARIA\" & Replace(lv.SelectedItem, " ", "_") & ".bre")
        Result = brReportManager.EditReport(reportID, startStep, visibleButtons)
        If Err.Number <> 0 Then
         MsgBox Err.Description & Chr(13) & Chr(10) & brReportManager.LastError
        End If
        
        'Result indicates which button was chosen to exit the dialog
        ' 0 for Cancel (no modification on the report)
        ' 1 for Done (report is saved)
        ' 5 for Done without Execute
        ' 6 for Done with Execute
        If Result = 4 Then
            cmdSave_Click
        ElseIf Result = 5 Then
            cmdSave_Click
        ElseIf Result = 6 Then
            cmdExecute_Click
        End If

    SetForegroundWindow Me.hwnd
    Me.Visible = True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHandler
        'Get report title
        Dim reportTitle As String
       ' reportTitle = brReportManager.GetReportTitle(reportID)
    
        'Get report file name to save
On Error GoTo DialogCancel
        ReportDialog.FileName = strFilename
        ReportDialog.ShowSave
On Error GoTo errHandler
        'Save the report definition to file
        brReportManager.SaveReportToFile Trim(reportID), ReportDialog.FileName
        strFilename = ""
    Exit Sub
    
DialogCancel: 'User pressed the Cancel button
    Exit Sub

errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUserDesign.cmdSave_Click", , EA_NORERAISE
End Sub

