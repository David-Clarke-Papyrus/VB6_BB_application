VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Step_6 
   BackColor       =   &H00E8E8DD&
   Caption         =   "Step 6 -  Finalize"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E8E8DD&
      Caption         =   "Stock take build"
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
      Height          =   4185
      Left            =   180
      TabIndex        =   2
      Top             =   225
      Width           =   6540
      Begin VB.Frame Frame2 
         BackColor       =   &H00E8E8DD&
         Caption         =   "Partial stock take option"
         ForeColor       =   &H8000000D&
         Height          =   960
         Left            =   450
         TabIndex        =   7
         Top             =   330
         Width           =   5685
         Begin VB.TextBox txtPartialDescription 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1170
            TabIndex        =   9
            Top             =   525
            Width           =   3855
         End
         Begin VB.CheckBox chkPartial 
            BackColor       =   &H00E8E8DD&
            Caption         =   "This is a partial stock-take- no quantities will be zeroed."
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   600
            TabIndex        =   8
            ToolTipText     =   "Use this option only if you are counting just part of the total stock."
            Top             =   210
            Width           =   4365
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   270
            TabIndex        =   10
            Top             =   555
            Width           =   885
         End
      End
      Begin VB.CommandButton cmdProvisional 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Discrepancy report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1455
         Width           =   2835
      End
      Begin VB.CommandButton cmdBuild 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Finalize"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3150
         Width           =   2850
      End
      Begin VB.TextBox txtDateTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   1935
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2595
         Width           =   2835
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Stocktake date and time (e.g. 22-08-2010 10:30 PM)"
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
         Height          =   300
         Left            =   795
         TabIndex        =   5
         Top             =   2220
         Width           =   4860
      End
   End
   Begin VB.CommandButton cmdPrev_to_5 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4545
      Width           =   840
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00D8D9C4&
      Caption         =   "&Close"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5865
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4515
      Width           =   840
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   4935
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Step_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents oSA As a_Stktke
Attribute oSA.VB_VarHelpID = -1
Dim strSql As String
Dim strFilename As String
Dim strTitle As String
Dim dteDateTime As Date
Dim xls As New ActiveReportsExcelExport.ARExportExcel

Public Sub Component(pSA As a_Stktke)
    Set oSA = pSA
    txtDateTime = Format(mSTDateTime, "dd-mm-yyyy HH:NN AM/PM")
End Sub



Private Sub cmdNext_To_5_Click()

End Sub

Private Sub chkPartial_Click()
    txtPartialDescription.Enabled = (chkPartial = 1)

End Sub

Private Sub cmdClose_Click()
    
    Unload Me
End Sub

'Private Sub cmdNext_To_8_Click()
'    Set frm8 = New frm_Step_8
'    frm8.Component oSA
'    frm8.Show
'    Unload Me
'End Sub

Private Sub cmdPrev_to_5_Click()
    Set frm5 = New frm_Step_5
    frm5.Component oSA
    frm5.Show
    Unload Me
End Sub

Private Sub cmdProvisional_Click()
'    Screen.MousePointer = vbHourglass
'    PrintAdjustMentReport
'    Screen.MousePointer = vbDefault

Dim f As New frmReportRepresentation
Dim bExVat As Boolean
Dim enPresentation As enumReportPresentation

    f.Show vbModal
    enPresentation = f.ReportPresentation
    bExVat = f.ExVAT
    Unload f

    Screen.MousePointer = vbHourglass
    DiscrepancyReport "ALL", bExVat, chkPartial = 1, enPresentation
    Screen.MousePointer = vbDefault


End Sub
Private Sub DiscrepancyReport(pType As String, bExVat As Boolean, bPartial As Boolean, Optional pReportPresentation As enumReportPresentation)
Dim arB As arValidation_B
Dim tmpNumber As Long
Dim rs As adodb.Recordset
Dim strSql As String
Dim strTitle As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

        Set arB = New arValidation_B
        arB.Printer.Orientation = ddOLandscape
        If bExVat Then
            If pType = "ALL" Then
                If bPartial Then
                    strSql = "Select * FROM vDiscrepancyExVAT_Provisional WHERE NOT PID is NULL ORDER BY SEC,P_TITLE"
                Else
                    strSql = "Select * FROM vDiscrepancyExVAT_Provisional ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (All values Ex VAT)"
            ElseIf pType = "POS" Then
                If bPartial Then
                    strSql = "Select * FROM vDiscrepancyExVAT_Provisional WHERE NOT PID is NULL AND DIFF < 0 ORDER BY SEC,P_TITLE"
                Else
                    strSql = "Select * FROM vDiscrepancyExVAT_Provisional WHERE DIFF < 0 ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment up)  (All values Ex VAT)"
            ElseIf pType = "NEG" Then
                If bPartial Then
                    strSql = "Select * FROM vDiscrepancyExVAT_Provisional WHERE NOT PID is NULL AND  DIFF > 0  ORDER BY SEC,P_TITLE"
                Else
                    strSql = "Select * FROM vDiscrepancyExVAT_Provisional WHERE DIFF > 0  ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment down) (All values Ex VAT)"
            End If
        Else
            If pType = "ALL" Then
                If bPartial Then
                    strSql = "Select * FROM vDiscrepancy_Provisional WHERE NOT PID is NULL ORDER BY SEC,P_TITLE"
                Else
                    strSql = "Select * FROM vDiscrepancy_Provisional ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (All values Incl VAT)"
            ElseIf pType = "POS" Then
                If bPartial Then
                    strSql = "Select * FROM vDiscrepancy_Provisional WHERE NOT PID is NULL AND DIFF < 0 ORDER BY SEC,P_TITLE"
                Else
                    strSql = "Select * FROM vDiscrepancy_Provisional WHERE DIFF < 0 ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment up) (All values Incl VAT)"
            ElseIf pType = "NEG" Then
                If bPartial Then
                    strSql = "Select * FROM vDiscrepancy_Provisional WHERE NOT PID is NULL AND  DIFF > 0  ORDER BY SEC,P_TITLE"
                Else
                    strSql = "Select * FROM vDiscrepancy_Provisional WHERE DIFF > 0  ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment down) (All values Incl VAT)"
            End If
        End If
        Set rs = New adodb.Recordset
        rs.CursorLocation = adUseClient
        rs.Open strSql, oPC.COshort, adOpenStatic
        Set rs.ActiveConnection = Nothing
      '  arB.Component rs, strTitle
        Screen.MousePointer = vbDefault
        arB.Component rs, strTitle
        arB.Top = 1000
        arB.Left = 400
        arB.Width = 12000
        arB.Height = 6000
      '  arB.Show
       PresentReport arB, pReportPresentation, strTitle, 1000, 400, 12000, 6000, "L"
        Set rs = Nothing
        Set arB = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Sub
Public Sub PresentReport(rpt As Object, pReportPresentation As enumReportPresentation, pReportName As String, Optional T As Long, Optional L As Long, Optional w As Long, Optional H As Long, Optional pOrientation As String)
Dim strExecutable As String
Dim i As Integer
Dim fs As New FileSystemObject
Dim sFile As String

    If Not IsMissing(pOrientation) Then
        If pOrientation = "L" Then
            rpt.Printer.Orientation = ddOLandscape
        Else
            rpt.Printer.Orientation = ddOPortrait
        End If
    End If
    If pReportPresentation = enPreview And Not rpt.WindowState = 2 Then
        If T > 0 Then rpt.Top = T
        If L > 0 Then rpt.Left = L
        If w > 0 Then rpt.Width = w
        If H > 0 Then rpt.Height = H
        
        rpt.Show
    ElseIf pReportPresentation = enPrintOut Then
        rpt.PrintReport True
    ElseIf pReportPresentation = enCSV Then
        rpt.Run False
        If Not fs.FolderExists(oPC.LocalFolder & "TEMP") Then
            fs.CreateFolder oPC.LocalFolder & "TEMP"
        End If
        sFile = oPC.LocalFolder & "TEMP\" & pReportName & ".XLS"
        i = 0
        Do Until fs.FileExists(sFile) = False
            i = i + 1
            sFile = oPC.LocalFolder & "TEMP\" & pReportName & "_" & CStr(i) & ".XLS"
        Loop
        
        
        
        xls.FileName = sFile
        If rpt.Pages.Count > 0 Then
            xls.Export rpt.Pages
        End If
        strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
        If strExecutable = "" Then
            MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder or Excel not installed." & vbCrLf & "Report will not open now but is saved in " & sFile, vbInformation, "Status"
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Shell strExecutable & " " & """" & sFile & """", vbNormalFocus
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub PrintAdjustMentReport()
Dim arB As arValidation_B
Dim tmpNumber As Long
Dim rs As adodb.Recordset
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        Set arB = New arValidation_B
        arB.Printer.Orientation = ddOLandscape
        strSql = "Select * FROM vDiscrepancy ORDER BY SEC,P_TITLE"

        Set rs = New adodb.Recordset
        rs.Open strSql, oPC.COshort, adOpenStatic
        strTitle = "Provisional adjustments"
        arB.Caption = strTitle
        arB.Component rs, strTitle
        arB.Left = 400
        arB.Top = 1000
        arB.Width = 12000
        arB.Height = 6000
        arB.Show vbModal
        Set rs = Nothing
        Set arB = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Sub


Private Sub Form_Load()
    Me.cmdBuild.Enabled = (UCase(oSA.Status) = "IN PROCESS")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = 1 Then
        If MsgBox("Do you want to close the stock-take application?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Private Sub cmdBuild_Click()
Dim lngTimeout As Long

    If chkPartial = False Then
        If MsgBox("WARNING: This stocktake is a count of the whole store and items not counted will have their quantity on hand values set to zero." & vbCrLf & "Click CANCEL to skip finalization procedure,", vbInformation + vbOKCancel, "Warning") = vbCancel Then
            MsgBox "The finalize procedure has been skipped.", vbInformation, "Status"
            Exit Sub
        End If
    End If
    If dteDateTime = CDate(0) Then
        MsgBox "You have not specified a cut-off date. You cannot build the stock-take", vbExclamation, "Cannot do this"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    SB.Panels(1).Text = "Finalizing stock take. This may take some time."
    DoEvents
    oSA.CreateStockAdjustment dteDateTime, chkPartial
    
    Me.cmdPrev_to_5.Enabled = False
    cmdBuild.Enabled = False
    
    
'-NOW ALL DONE in 'CreateStockAdjustment'
'''    lngTimeout = oPC.COshort.CommandTimeout
'''    oPC.COshort.CommandTimeout = 0
'''    'update Qty onhand after stock take so any movements dated after the stock take are applied.
'''    oPC.COshort.Execute "UPDATE tPRODUCT SET P_QtyOnHand = ISNULL(P_QtyLastStockTake,0) + ISNULL(Qty,0) FROM tPRODUCT Left JOIN zRecentMM   on PID = P_ID WHERE NOT (P_PRODUCTTYPE IN ( 'M','N'))"
'''    oPC.COshort.Execute "UPDATE tPRODUCT SET P_CostLastStockTake = CAST(P_Cost as MONEY)  WHERE NOT (P_PRODUCTTYPE IN ( 'M','N'))"
'''    oPC.COshort.CommandTimeout = lngTimeout
    
    Me.cmdClose.Enabled = True
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Screen.MousePointer = vbDefault
    MsgBox "Finalize procedure complete", vbOKOnly, "Status"
    Exit Sub
End Sub

Private Sub txtDateTime_Change()
    cmdBuild.Enabled = IsDate(txtDateTime)
    If (IsDate(txtDateTime)) Then
        dteDateTime = CDate(txtDateTime)
    End If
End Sub

