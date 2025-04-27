VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTransmissionControl 
   Caption         =   "Transmission control"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780.001
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   9780.001
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame8 
      Caption         =   "SOH_RS_Q"
      Height          =   1050
      Left            =   3750
      TabIndex        =   42
      Top             =   6495
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_SOH_RS_Q 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_SOH_RS_Q 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_SOH_RS_Q 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_SOH_RS_Q 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   46
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Invocation_RS_Q"
      Height          =   1050
      Left            =   3735
      TabIndex        =   37
      Top             =   5370
      Width           =   3525
      Begin VB.CommandButton chkGetStatus_Invocation_RS_Q 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdStartQ_Invocation_RS_Q 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_Invocation_RS_Q 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblQStatus_Invocation_RS_Q 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   41
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Alert consumer_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   32
      Top             =   6510
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_AlertConsumerQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_AlertConsumerQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":11B2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_AlertConsumerQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_AlertConsumerQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   36
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "AlertLoad consumer_Q"
      Height          =   1050
      Left            =   3720
      TabIndex        =   27
      Top             =   4185
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_AlertLoadQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1830
         Picture         =   "frmTransmissionControlCENTRAL.frx":153C
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_AlertLoadQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":18C6
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_AlertLoadQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_AlertLoadQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   75
         TabIndex        =   31
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "SOHConsumer_Q"
      Height          =   1050
      Left            =   3735
      TabIndex        =   22
      Top             =   2985
      Width           =   3525
      Begin VB.CommandButton chkGetStatus_SOHQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdStartQ_SOHQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":1C50
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_SOHQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":1FDA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblQStatus_SOHQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   26
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MasterLoyaltySource_Q"
      Height          =   1050
      Left            =   90
      TabIndex        =   16
      Top             =   5370
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_MLSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":2364
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_MLSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":26EE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_MLSQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_MLSQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   20
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdClearDebug 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear _tSBLog  table"
      Height          =   495
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2190
      Width           =   1215
   End
   Begin VB.CommandButton cmdRecycle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Recycle ERRORLOG"
      Height          =   435
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4515
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefreshTimer 
      Height          =   300
      Left            =   7665
      Picture         =   "frmTransmissionControlCENTRAL.frx":2A78
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   45
      Width           =   810
   End
   Begin VB.Frame Frame2 
      Caption         =   "SalesConsumer_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   8
      Top             =   4185
      Width           =   3525
      Begin VB.CommandButton chkGetStatus_SCQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdStartQ_SCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":2E02
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_SCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":318C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblQStatus_SCQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   12
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LoyaltyConsumer_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   3
      Top             =   2985
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_LCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControlCENTRAL.frx":3516
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_LCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControlCENTRAL.frx":38A0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_LCQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_LCQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   5
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdOpenLog 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Open log"
      Height          =   420
      Left            =   7875
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4020
      Width           =   1800
   End
   Begin VB.CommandButton cmdClearQ 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear queue"
      Height          =   555
      Left            =   8895.001
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3405
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   8685.001
      Picture         =   "frmTransmissionControlCENTRAL.frx":3C2A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5685
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7965
      Top             =   1005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   2655
      Left            =   120
      TabIndex        =   21
      Top             =   45
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   4683
      SortKey         =   1
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Procedure"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTransmissionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilename As String
Dim rs2 As ADODB.Recordset


Private Sub chkGetStatus_AlertConsumerQ_Click()
Dim res As Integer
    res = QEnabled("ALERTLOAD_CONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_AlertConsumerQ.Caption = "ALERTLOAD_CONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_AlertConsumerQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_AlertConsumerQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_AlertConsumerQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_AlertLoadQ_Click()
Dim res As Integer
    res = QEnabled("ALERTLOAD_CONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_AlertLoadQ.Caption = "ALERTLOAD_CONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_AlertLoadQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_AlertLoadQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_AlertLoadQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_Invocation_RS_Q_Click()
Dim res As Integer
    res = QEnabled("INVOCATION_RS_Q")

    If res = 999 Then
        Me.lblQStatus_Invocation_RS_Q.Caption = "INVOCATION_RS_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_Invocation_RS_Q.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_Invocation_RS_Q.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_Invocation_RS_Q.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_LCQ_Click()
Dim res As Integer
    res = QEnabled("LOYALTYCONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_LCQ.Caption = "LOYALTYCONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_LCQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_LCQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_LCQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_SOH_RS_Q_Click()
Dim res As Integer
    res = QEnabled("SOH_RS_Q")

    If res = 999 Then
        Me.lblQStatus_SOH_RS_Q.Caption = "SOH_RS_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_SOH_RS_Q.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_SOH_RS_Q.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_SOH_RS_Q.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_SOHQ_Click()
Dim res As Integer
    res = QEnabled("SOHCONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_SOHQ.Caption = "SOHCONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_SOHQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_SOHQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_SOHQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_MLSQ_Click()
Dim res As Integer
    res = QEnabled("MASTERLOYALTYSOURCE_Q")

    If res = 999 Then
        Me.lblQStatus_MLSQ.Caption = "MASTERLOYALTYSOURCE_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_MLSQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_MLSQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_MLSQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_SCQ_Click()
Dim res As Integer
    res = QEnabled("SALESCONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_SCQ.Caption = "SALESCONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_SCQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_SCQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_SCQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub



Private Sub cmdClearDebug_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "DELETE FROM _tSBLog"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oPC.DisconnectDBShort

End Sub

Private Sub cmdClearQ_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "_ClearQueue"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oPC.DisconnectDBShort
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdOpenLog_Click()
'Dim frm As New frmFile
'    frm.Show vbModal
    cmdFindLogFile_Click
    Shell "NOTEPAD.EXE '" & strFilename & "'", vbNormalFocus
End Sub
Private Sub cmdFindLogFile_Click()
Dim fs As New FileSystemObject

    strFilename = GetSetting("PBKSC", "SB", "LOGFILEPATH", "")
    If fs.GetBaseName(strFilename) <> "ERRORLOG" Then
        cd1.DialogTitle = "Open SQL Server log file"
        cd1.DefaultExt = ""
        cd1.InitDir = "c:\Program files\Microsoft SQL SERVER"
        cd1.ShowOpen
        strFilename = cd1.FileName
        SaveSetting "PBKSC", "SB", "LOGFILEPATH", strFilename
    End If
    
End Sub



Private Function QEnabled(s As String) As Integer
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As ADODB.Recordset
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT IS_RECEIVE_ENABLED FROM sys.service_queues WHERE name = '" & s & "'"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 50
    Set res = cmd.Execute
    If Not res.State = 0 Then
        If Not res.EOF Then
            QEnabled = CLng(res.Fields(0))
        End If
    Else
        QEnabled = 999
    End If
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Set cmd = Nothing
End Function

Private Sub cmdRefreshTimer_Click()
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
Dim s As String

    OpenResult = oPC.OpenDBSHort
    Set rs2 = Nothing
    Set rs2 = New ADODB.Recordset
    rs2.CursorLocation = adUseClient
    rs2.Open "SELECT TOP 150 * FROM _tSBLog Order By SBL_DATE DESC", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    
    LoadListView
    If OpenResult = 0 Then oPC.DisconnectDBShort
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTransmissionControl.cmdRefreshTimer_Click"
End Sub
Private Sub LoadListView()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvw.ListItems.Clear
    For i = 1 To rs2.RecordCount
        Set objItm = Me.lvw.ListItems.Add
        With objItm
            .Text = Format(FND(rs2.Fields("SBL_DATE")), "yyyy-mm-dd Hh:Nn")
            .SubItems(1) = FNS(rs2.Fields("SBL_Msg"))
            .SubItems(2) = FNS(rs2.Fields("SBL_PROC"))
        End With
        rs2.MoveNext
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTransmissionControl.LoadListView"
End Sub

Private Sub cmdStartQ_AlertLoadQ_Click()
    startQ "ALERTLOAD_CONSUMER_Q"

End Sub

Private Sub cmdStartQ_Invocation_RS_Q_Click()
    startQ "INVOCATION_RS_Q"

End Sub

Private Sub cmdStartQ_LCQ_Click()
    startQ "LOYALTYCONSUMER_Q"
End Sub

Private Sub cmdStartQ_MLSQ_Click()
    startQ "MASTERLOYALTYSOURCE_Q"

End Sub

Private Sub cmdStartQ_SOH_RS_Q_Click()
    startQ "SOH_RS_Q"

End Sub

Private Sub cmdStartQ_SOHQ_Click()
    startQ "SOHCONSUMER_Q"

End Sub

Private Sub cmdStopQ_AlertLoadQ_Click()
    stopQ "ALERTLOAD_CONSUMER_Q"

End Sub

Private Sub cmdStopQ_Invocation_RS_Q_Click()
    stopQ "INVOCATION_RS_Q"

End Sub

Private Sub cmdStopQ_LCQ_Click()
    stopQ "LOYALTYCONSUMER_Q"
End Sub
Private Sub cmdStartQ_SCQ_Click()
    startQ "SALESCONSUMER_Q"
End Sub

Private Sub cmdStopQ_MLSQ_Click()
    stopQ "MASTERLOYALTYSOURCE_Q"

End Sub

Private Sub cmdStopQ_SCQ_Click()
    stopQ "SALESCONSUMER_Q"
End Sub
Private Sub stopQ(s As String)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = OFF;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Set cmd = Nothing

End Sub
Private Sub startQ(s As String)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = ON;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Set cmd = Nothing

End Sub
Private Sub cmdRecycle_Click()
    RecycleErrorLog
End Sub

Private Sub RecycleErrorLog()
Dim OpenResult As Integer
Dim strCommandFilePath As String
Dim oTF As New z_TextFile
Dim fs As New FileSystemObject

    OpenResult = oPC.OpenDBSHort
        If OpenResult = 0 Then
            strCommandFilePath = "\\" & oPC.NameOfPC & "\PBKS_S\RecycleErrorLog.SQL"
            Set oTF = New z_TextFile
            oTF.OpenTextFile strCommandFilePath
            
            oTF.WriteToTextFile "USE [Master]"
            oTF.WriteToTextFile "GO"
            oTF.WriteToTextFile "EXEC sp_cycle_errorlog ;"
    
            oTF.WriteToTextFile "GO"
            oTF.CloseTextFile
            Set oTF = Nothing
            If fs.FileExists(strCommandFilePath) Then
                ExecuteScript strCommandFilePath
            Else
                MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
            End If
            oPC.DisconnectDBShort
            MsgBox "ERRORLOG recycled"
        Else
            MsgBox "Cannot open database. Script has not run"
        End If

End Sub
Private Sub ExecuteScript(strCommandFilePath)
Dim strCommand As String
Dim res As Boolean
Dim fs As New FileSystemObject
    
    strCommand = "SQLCMD -Usa -P" & oPC.Password & " -S" & oPC.ServerName & " -dPBKSC -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
End Sub

Private Sub cmdStopQ_SOH_RS_Q_Click()
    stopQ "SOH_RS_Q"

End Sub

Private Sub cmdStopQ_SOHQ_Click()
    stopQ "SOHCONSUMER_Q"

End Sub

