VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{E281C260-6F27-11D1-8AF0-00A0C98CD92B}#2.0#0"; "ardespro2.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUDR 
   BackColor       =   &H00D5D5C1&
   Caption         =   "User designed reports"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8505
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Delete"
      Height          =   300
      Left            =   1395
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   390
      Width           =   825
   End
   Begin VB.CommandButton cmdSaveAs 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Save as"
      Enabled         =   0   'False
      Height          =   465
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   765
   End
   Begin VB.CommandButton cmdXML 
      BackColor       =   &H00D5D5C1&
      Caption         =   "XML"
      Height          =   465
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   900
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Export"
      Enabled         =   0   'False
      Height          =   345
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   150
      Width           =   1245
   End
   Begin VB.CommandButton cmdImport 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Import"
      Height          =   345
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   150
      Width           =   1245
   End
   Begin VB.CommandButton cmdData 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Data"
      Enabled         =   0   'False
      Height          =   465
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   75
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   465
      Left            =   10050
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   765
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00D5D5C1&
      Caption         =   "New report"
      Height          =   465
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   1245
   End
   Begin VB.CommandButton cmdLoadReport 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Load Report"
      Height          =   465
      Left            =   4725
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   1185
   End
   Begin VB.ComboBox cboReports 
      Height          =   315
      Left            =   1410
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   60
      Width           =   3240
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6690
      Left            =   90
      TabIndex        =   0
      Top             =   810
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   11800
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14013889
      TabCaption(0)   =   "Designer"
      TabPicture(0)   =   "frmUDR.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "arD"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Preview"
      TabPicture(1)   =   "frmUDR.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "arv"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdToExcel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdToPDF"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdToPDF 
         BackColor       =   &H00D5D5C1&
         Caption         =   "PDF"
         Height          =   360
         Left            =   9750
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   315
         Width           =   1380
      End
      Begin VB.CommandButton cmdToExcel 
         BackColor       =   &H00D5D5C1&
         Caption         =   "Spreadsheet"
         Height          =   360
         Left            =   11205
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   315
         Width           =   1380
      End
      Begin DDActiveReportsDesignerCtl.ARDesigner arD 
         Height          =   5205
         Left            =   -74820
         TabIndex        =   1
         Top             =   420
         Width           =   12525
         _ExtentX        =   22093
         _ExtentY        =   9181
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
         Height          =   5025
         Left            =   330
         TabIndex        =   14
         Top             =   780
         Width           =   12360
         _ExtentX        =   21802
         _ExtentY        =   8864
         SectionData     =   "frmUDR.frx":0038
      End
   End
End
Attribute VB_Name = "frmUDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tl As New z_TextListSimple
Dim rpt As DDActiveReports2.ActiveReport
Dim sec As DDActiveReports2.Section
Dim ctl As DDActiveReports2.DataControl
Dim strSQL As String
Dim frm As frmConfigureDataForReport
Attribute frm.VB_VarHelpID = -1
Dim brpt() As Byte
Dim oMD As New z_ReportMetadata
Dim fs As New FileSystemObject
Dim oTF As z_TextFile
Dim ReportState As Integer
Const rsEmpty = 1
Const rsEditing = 2
Dim strReportName As String

Private Sub cboReports_Click()
    On Error GoTo errHandler
    If arD.IsDirty Then
        If MsgBox(oMD.Report_name & " has not been saved. Continue?", vbYesNo + vbInformation, "Warning") = vbNo Then
            Exit Sub
        End If
    End If

    arD.Visible = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cboReports_Click"
    HandleError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    If cboReports.text = "" Then Exit Sub
    If MsgBox("You are wanting to delete report: " & cboReports.text & ". Confirm.", vbYesNo + vbQuestion, "Warning") = vbNo Then
        Exit Sub
    End If
    oSQL.DeleteReport Me.cboReports.text
    LoadListOfReports
    Me.cboReports.text = ""
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdDelete_Click"
    HandleError
End Sub

Private Sub cmdExport_Click()
    On Error GoTo errHandler
Dim FileName As String
Dim fs As New FileSystemObject
Dim oSQL As New z_SQL

'    If MsgBox("The report will be saved before exporting. Continue?", vbYesNo + vbQuestion, "Warning") = vbNo Then
'        Exit Sub
'    End If
'    InsertLayoutXML
'    oSQL.SaveReport oMD.Report_name, oMD.Metadata_XML

    If Not fs.FolderExists(oPC.SharedFolderRoot & "\Ad-hoc reports specifications") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\Ad-hoc reports specifications")
    End If
    CD1.DefaultExt = "XML"
    CD1.DialogTitle = "Save report to external file"
    CD1.InitDir = oPC.SharedFolderRoot & "\Ad-hoc reports specifications"
    CD1.FileName = Me.cboReports
    CD1.ShowSave
    FileName = CD1.FileName
    If FileName = "" Then
        MsgBox "Report not exported", vbInformation + vbOKOnly, "Warning"
        Exit Sub
    End If
    Set oTF = New z_TextFile
    oTF.OpenTextFile FileName
    oTF.WriteToTextFile_NoLineTerminator oMD.Metadata_XML
    oTF.CloseTextFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdExport_Click"
    HandleError
End Sub
Private Sub cmdImport_Click()
    On Error GoTo errHandler
Dim FileName As String
Dim s As String
Dim fs As New FileSystemObject
Dim oSQL As New z_SQL

    If arD.IsDirty Then
        If MsgBox(oMD.Report_name & " has not been saved. Continue?", vbYesNo + vbInformation, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    CD1.DefaultExt = ".XML"
    CD1.DialogTitle = "Load report from external file"
    CD1.InitDir = oPC.SharedFolderRoot & "\Ad-hoc reports specifications"
    CD1.FLAGS = &H1000
'    CD1.CancelError
    CD1.ShowOpen
 '   MsgBox "Pos 5"
    If CD1.FileName = "" Then Exit Sub
    FileName = CD1.FileName
    Set oTF = New z_TextFile
    s = oTF.ReadFileBinary(FileName)
    
    Set oMD = Nothing
    Set oMD = New z_ReportMetadata
    oMD.LoadMetadataToXML s
    
    Me.cboReports = oMD.Report_name
    Me.Caption = Me.Caption & "   report: " & oMD.Report_name
    
    rpt.LoadLayout StringToByteArray(oMD.Layout_fromXML, False, True)
    InsertLayoutXML
    oSQL.SaveReport oMD.Report_name, oMD.Metadata_XML
    
    LoadListOfReports
    
    Caption = "User designed reports   (" & oMD.Report_name & ")"
    strSQL = oMD.GetSQL
    ctl.ConnectionString = oPC.CO.ConnectionString
    ctl.Password = oPC.Password
    ctl.Source = strSQL
    arD.LoadFromObject rpt
    
    oTF.CloseTextFile
    
    cmdData.Enabled = True
    cmdSave.Enabled = True
    cmdSaveAs.Enabled = True
    cmdExport.Enabled = True
    arD.Visible = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdImport_Click"
    HandleError
End Sub


Private Sub cmdSave_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim rs As ADODB.Recordset
Dim s As String
 
    If oSQL.ReportExists(oMD.Report_name) Then
        If MsgBox("This report already exists and will be updated.Please confirm.", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            cmdSaveAs_Click
            Exit Sub
        End If
    End If
    InsertLayoutXML
    oSQL.SaveReport oMD.Report_name, oMD.Metadata_XML
    LoadListOfReports
    Me.cboReports.text = oMD.Report_name
    Set oMD = New z_ReportMetadata

    oSQL.LoadReport Me.cboReports, s
    oMD.LoadMetadataToXML s
    rpt.LoadLayout StringToByteArray(oMD.Layout_fromXML, False, True)
    arD.Visible = True
   
    Me.Caption = "User designed reports   (" & oMD.Report_name & ")"
    strSQL = oMD.GetSQL
    ctl.Source = strSQL
    arD.LoadFromObject rpt
    cmdData.Enabled = True
    cmdSave.Enabled = True
    cmdSaveAs.Enabled = True
    cmdExport.Enabled = True
    MsgBox "Report is saved", , "Status"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdSave_Click"
    HandleError
End Sub

Private Sub cmdSaveAs_Click()
    On Error GoTo errHandler
Dim F As New frmReportName
Dim oSQL As New z_SQL

top:
    F.txtReportname = Me.cboReports
    F.Show vbModal
    
    oMD.Report_name = F.Reportname
    Unload F
    Caption = "User designed reports   (" & oMD.Report_name & ")"
    If oSQL.ReportExists(oMD.Report_name) Then
        If MsgBox("This report already exists and will be updated.Please confirm.", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            GoTo top
        End If
    End If
    InsertLayoutXML
    oSQL.SaveReport oMD.Report_name, oMD.Metadata_XML
    LoadListOfReports
    cboReports.text = oMD.Report_name
    arD.Visible = True
    MsgBox "Report is saved", , "Status"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdSaveAs_Click"
    HandleError
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo errHandler
    Set rpt = New DDActiveReports2.ActiveReport
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdLoad_Click"
    HandleError
End Sub

Private Sub cmdConfiguration_Click()
    On Error GoTo errHandler

    frm.Show vbModal
    Unload frm
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdConfiguration_Click"
    HandleError
End Sub

Private Sub cmdLoadReport_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim rs As ADODB.Recordset
Dim s As String



    If Me.cboReports = "" Then Exit Sub
    Set oMD = New z_ReportMetadata

  '  oSQL.LoadReport Me.cboReports, oMD.Report_view, oMD.Metadata_XML, brpt
    oSQL.LoadReport Me.cboReports, s
    oMD.LoadMetadataToXML s
    oMD.ConnectionString = oPC.ConnectionString
    rpt.LoadLayout StringToByteArray(oMD.Layout_fromXML, False, True)
    oSQL.SaveReport rpt.documentName, oMD.Metadata_XML
    arD.Visible = True
    arD.ToolbarsVisible = ddTBAlignment + ddTBExplorer + ddTBFields + ddTBFormat + ddTBMenu + ddTBPropertyToolbox + ddTBStandard + ddTBToolBox
    Me.Caption = "User designed reports   (" & oMD.Report_name & ")"
    strSQL = oMD.GetSQL
   
    ctl.Source = strSQL
    ctl.Password = oPC.Password
    arD.LoadFromObject rpt
    cmdData.Enabled = True
    cmdSave.Enabled = True
    cmdSaveAs.Enabled = True
    cmdExport.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdLoadReport_Click"
    HandleError
End Sub

Private Sub cmdNew_Click()
    On Error GoTo errHandler
Dim F As New frmReportName

    F.Show vbModal
    
    Set oMD = New z_ReportMetadata
    oMD.InitializeMetaDataXML
    
    oMD.Report_name = F.Reportname
    Unload F
    Me.Caption = "User designed reports   (" & oMD.Report_name & ")"
    If oMD.Report_name = "" Then
        Exit Sub
    End If
    
    Set frm = New frmConfigureDataForReport
    frm.Component oMD
    frm.Show vbModal
    strSQL = oMD.GetSQL
    Set rpt = New DDActiveReports2.ActiveReport
    
    Set sec = rpt.Sections("Detail")
    Set ctl = sec.Controls.Add("DDActiveReports2.DataControl")
        ctl.Name = "ADOData"
        ctl.ConnectionString = oPC.ConnectionString
        ctl.Source = strSQL
        ctl.Left = 4 * 1440
        ctl.top = 0
    arD.Visible = True
    arD.NewLayout
    arD.LoadFromObject rpt
    Unload frm
    cmdData.Enabled = True
    cmdSave.Enabled = True
    cmdSaveAs.Enabled = True
    cmdExport.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdNew_Click"
    HandleError
End Sub
Private Sub cmdData_Click()
    On Error GoTo errHandler

    Set frm = New frmConfigureDataForReport
    frm.Component oMD
    frm.Show vbModal
    strSQL = oMD.GetSQL
    ctl.Source = strSQL
    arD.LoadFromObject rpt
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdData_Click"
    HandleError
End Sub


Private Sub cmdToPDF_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If rpt Is Nothing Then Exit Sub
    rpt.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & oMD.Report_name & "_" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enPDF
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdToPDF_Click"
    HandleError
End Sub

Private Sub cmdToExcel_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If rpt Is Nothing Then Exit Sub
    rpt.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "TEMP\" & StripToAlphanumeric(oMD.Report_name) & "_" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(rpt.Pages)
    OpenFileWithApplication fn, enExcel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdToExcel_Click"
    HandleError
End Sub

Private Sub cmdXML_Click()
    On Error GoTo errHandler
rpt.SaveLayout "c:\PBKS\TESTREPORT.rpx", ddSOFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.cmdXML_Click"
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler

    LoadListOfReports
    Set rpt = New DDActiveReports2.ActiveReport
    
    Set sec = rpt.Sections("Detail")
    Set ctl = sec.Controls.Add("DDActiveReports2.DataControl")
        ctl.Name = "ADOData"
        ctl.ConnectionString = oPC.ConnectionString
        ctl.Source = strSQL
        ctl.Left = 4 * 1440
        ctl.top = 0
    'Set active Tab to the designer
    SSTab1.Tab = 0
'    Set rpt = New ActiveReport  'Activate all the toolbars
    arD.LoadFromObject rpt
    arD.ToolbarsAccessible = ddTBToolBox + ddTBAlignment + ddTBExplorer + ddTBFields + ddTBFormat + ddTBMenu + ddTBPropertyToolbox + ddTBStandard
    arD.ToolbarsVisible = ddTBToolBox + ddTBAlignment + ddTBExplorer + ddTBFields + ddTBFormat + ddTBMenu + ddTBPropertyToolbox + ddTBStandard
    cmdXML.Visible = (oPC.servername = "PBKS-SVR")
    arV.Left = 180
    arV.Width = 12360
  '  arD.left = 180
    arD.Width = Me.Width - 1000
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.Form_Load"
    HandleError
End Sub

Private Sub LoadListOfReports()
    On Error GoTo errHandler
    tl.Load sltReportList
    LoadComboFromTextListSimple Me.cboReports, tl
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.LoadListOfReports"
End Sub


Private Sub Form_Resize()
    On Error GoTo errHandler
Dim lngDiff As Long
    
'    arV.left = SSTab1.left + 180
'    arD.left = SSTab1.left + 180
'    arV.Width = SSTab1.Width - 180
'    arD.Width = SSTab1.Width - 180
'
    SSTab1.Width = NonNegative_Lng(Me.Width - 400)
    lngDiff = SSTab1.Height
    SSTab1.Height = NonNegative_Lng(Me.Height - 1200)
    lngDiff = SSTab1.Height - lngDiff
  '  arv.top = SSTab1.top + 3800
    arV.Width = NonNegative_Lng(Me.Width - 900)
    lngDiff = arV.Height
    arV.Height = NonNegative_Lng(Me.Height - 2100)
    lngDiff = arV.Height - lngDiff
    
    arD.Width = NonNegative_Lng(Me.Width - 900)
    lngDiff = arD.Height
    arD.Height = NonNegative_Lng(Me.Height - 2000)
    lngDiff = arD.Height - lngDiff
    cmdToExcel.Left = NonNegative_Lng(arV.Width - 1500)
    cmdToPDF.Left = NonNegative_Lng(arV.Width - 3500)
    If SSTab1.Tab = 0 Then
        cmdToExcel.Visible = False
        cmdToPDF.Visible = False
    Else
        cmdToExcel.Visible = True
        cmdToPDF.Visible = True
    
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.Form_Resize"
    HandleError
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo errHandler
    Select Case PreviousTab
    Case Is = 0
        prepPreview
        cmdToExcel.Visible = True
        cmdToPDF.Visible = True
    Case Is = 1
        prepDesigner
        cmdToExcel.Visible = False
        cmdToPDF.Visible = False
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.SSTab1_Click(PreviousTab)", PreviousTab
    HandleError
End Sub

Private Sub prepPreview()
    On Error GoTo errHandler
'Must be used to writes the designer's layout   'to the report so it can be previewed
    arD.SaveToObject rpt
    
    
    
    rpt.restart 'Run the new report
    rpt.ResetScripts
   ' rpt.AddNamedItem "vbo", vbo
   ' rpt.AddCode ScriptCode()
    rpt.ScriptDebuggerEnabled = True
    'rpt.Run False   'Add the report to the veiwer
    Set arV.ReportSource = rpt
'errHndl:
'    MsgBox "Error Previewing the Report: " & Err.Number & " " & Err.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.prepPreview"
End Sub
Private Sub prepDesigner()
    On Error GoTo errHandler
    If Not arV.ReportSource Is Nothing Then
        arV.ReportSource.Cancel
        Set arV.ReportSource = Nothing
    End If
'errHndl:    MsgBox "Error in Design Preview: " & Err.Number & " " & Err.Description
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.prepDesigner"
End Sub

Private Function InsertLayoutXML() As String
    On Error GoTo errHandler
    brpt = Me.arD.Report.SaveLayout("", ddSOByteArray)
    oMD.AppendLayout ByteArrayToString(brpt)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmUDR.InsertLayoutXML"
End Function

