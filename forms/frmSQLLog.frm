VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSQLLog 
   Caption         =   "Transmission control"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   5745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "MasterLoyaltyConsumer_Q"
      Height          =   1050
      Left            =   60
      TabIndex        =   29
      Top             =   5805
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton chkGetStatus_MLCQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdStartQ_MLCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmSQLLog.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_MLCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmSQLLog.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblQStatus_MLCQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   33
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdClearDebug 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear _debug"
      Height          =   330
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1155
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdRecycle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Recycle ERRORLOG"
      Enabled         =   0   'False
      Height          =   435
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2775
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefreshTimer 
      Height          =   300
      Left            =   4965
      Picture         =   "frmSQLLog.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   30
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txtTimer 
      Height          =   795
      Left            =   15
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   345
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Frame Frame3 
      Caption         =   "HubSource_Q"
      Height          =   1050
      Left            =   60
      TabIndex        =   18
      Top             =   4620
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_HSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmSQLLog.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_HSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmSQLLog.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_HSQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_HSQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   22
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SalesSource_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   13
      Top             =   3510
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton chkGetStatus_SSQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdStartQ_SSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmSQLLog.frx":11B2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_SSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmSQLLog.frx":153C
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblQStatus_SSQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   17
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LoyaltySource_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   8
      Top             =   2385
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_LSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmSQLLog.frx":18C6
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_LSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmSQLLog.frx":1C50
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_LSQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_LSQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.Frame frTimerQ 
      Caption         =   "TimerQueue"
      Height          =   1050
      Left            =   75
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_TQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmSQLLog.frx":1FDA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_TQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmSQLLog.frx":2364
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_TQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_TQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdOpenLog 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Open log"
      Height          =   420
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1800
   End
   Begin VB.CommandButton cmdClearQ 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear queue"
      Height          =   555
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Stop timer"
      Height          =   330
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Start timer"
      Height          =   330
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   4140
      Picture         =   "frmSQLLog.frx":26EE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSQLLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String
'
'Private Sub chkGetStatus_HSQ_Click()
'Dim Res As Integer
'    Res = TimerQEnabled("HUBSOURCE_Q")
'
'    If Res = 999 Then
'        Me.lblQStatus_HSQ.Caption = "HUBSOURCE_Q queue cannot be found"
'    Else
'        If Res = -1 Then
'            Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If Res = 0 Then
'                Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_HSQ.Caption = "Unknown (" & CStr(Res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_LSQ_Click()
'Dim Res As Integer
'    Res = TimerQEnabled("LOYALTYSOURCE_Q")
'
'    If Res = 999 Then
'        Me.lblQStatus_LSQ.Caption = "LOYALTYSOURCE_Q queue cannot be found"
'    Else
'        If Res = -1 Then
'            Me.lblQStatus_LSQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If Res = 0 Then
'                Me.lblQStatus_LSQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_LSQ.Caption = "Unknown (" & CStr(Res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_MLCQ_Click()
'Dim Res As Integer
'    Res = TimerQEnabled("MASTERLOYALTYCONSUMER_Q")
'
'    If Res = 999 Then
'        Me.lblQStatus_MLCQ.Caption = "MASTERLOYALTYCONSUMER_Q queue cannot be found"
'    Else
'        If Res = -1 Then
'            Me.lblQStatus_MLCQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If Res = 0 Then
'                Me.lblQStatus_MLCQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_MLCQ.Caption = "Unknown (" & CStr(Res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_SSQ_Click()
'Dim Res As Integer
'    Res = TimerQEnabled("SALESSOURCE_Q")
'
'    If Res = 999 Then
'        Me.lblQStatus_SSQ.Caption = "SALESSOURCE_Q queue cannot be found"
'    Else
'        If Res = -1 Then
'            Me.lblQStatus_SSQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If Res = 0 Then
'                Me.lblQStatus_SSQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_SSQ.Caption = "Unknown (" & CStr(Res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_TQ_Click()
'Dim Res As Integer
'    Res = TimerQEnabled("TimerQueue")
'
'    If Res = 999 Then
'        Me.lblQStatus_TQ.Caption = "Timer queue cannot be found"
'    Else
'        If Res = -1 Then
'            Me.lblQStatus_TQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If Res = 0 Then
'                Me.lblQStatus_TQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_TQ.Caption = "Unknown (" & CStr(Res) & ")"
'            End If
'        End If
'    End If
'End Sub


'Private Sub cmdClearDebug_Click()
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "DELETE FROM _tDEBUG"
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 360
'    cmd.Execute
'    Set cmd = Nothing
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'End Sub
'
'Private Sub cmdClearQ_Click()
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "_ClearQueue"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 360
'    cmd.Execute
'    Set cmd = Nothing
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'End Sub
'
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdOpenLog_Click()
'Dim frm As New frmFile
'    frm.Show vbModal
    cmdFindLogFile_Click
    Shell "NOTEPAD.EXE '" & strFileName & "'", vbNormalFocus
End Sub
Private Sub cmdFindLogFile_Click()
Dim fs As New FileSystemObject

    strFileName = GetSetting("PBKS", "SB", "LOGFILEPATH", "")
    If fs.GetBaseName(strFileName) <> "ERRORLOG" Then
        CD1.DialogTitle = "Open SQL Server log file"
        CD1.DefaultExt = ""
        CD1.InitDir = "c:\Program files\Microsoft SQL SERVER"
        CD1.ShowOpen
        strFileName = CD1.FileName
        SaveSetting "PBKS", "SB", "LOGFILEPATH", strFileName
    End If
    
End Sub

'
'Private Sub cmdRefreshTimer_Click()
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'Dim Res As Recordset
'Dim s As String
'
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "SELECT TOP 10 * frOM _tDEBUG Order By k DESC"
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 360
'    Set Res = cmd.Execute
'    s = ""
'    If Not Res.State = 0 Then
'        Do While Not Res.EOF
'            s = s & Res.Fields(1) & vbCrLf
'            Res.MoveNext
'        Loop
'        txtTimer = s
'    Else
'        txtTimer = ""
'    End If
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'    Set cmd = Nothing
'
'End Sub
'
'Private Sub cmdStart_Click()
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "_StartTimer"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 360
'    cmd.Execute
'    Set cmd = Nothing
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'
'End Sub
'
'
'
'
'Private Sub cmdStop_Click()
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "_EndTimer"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 360
'    cmd.Execute
'    Set cmd = Nothing
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'End Sub
'
'Private Function TimerQEnabled(s As String) As Integer
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'Dim Res As Recordset
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "SELECT IS_RECEIVE_ENABLED FROM sys.service_queues WHERE name = '" & s & "'"
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 360
'    Set Res = cmd.Execute
'    If Not Res.State = 0 Then
'        If Not Res.EOF Then
'            TimerQEnabled = CLng(Res.Fields(0))
'        End If
'    Else
'        TimerQEnabled = 999
'    End If
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'    Set cmd = Nothing
'End Function
'
'Private Sub cmdStartQ_TQ_Click()
'    startQ "TimerQueue"
'End Sub
'
'Private Sub cmdStopQ_MLCQ_Click()
'    stopQ "MASTERLOYALTYCONSUMER_Q"
'End Sub
'
'Private Sub cmdStopQ_TQ_Click()
'    stopQ "TimerQueue"
'End Sub
'Private Sub cmdStopQ_LSQ_Click()
'    stopQ "LOYALTYSOURCE_Q"
'End Sub
'Private Sub cmdStartQ_SSQ_Click()
'    startQ "SALESSOURCE_Q"
'End Sub
'Private Sub cmdStartQ_MLCQ_Click()
'    startQ "MASTERLOYALTYCONSUMER_Q"
'End Sub
'Private Sub cmdStopQ_SSQ_Click()
'    stopQ "SALESSOURCE_Q"
'End Sub
'Private Sub cmdStartQ_HSQ_Click()
'    startQ "HUBSOURCE_Q"
'End Sub
'Private Sub cmdStopQ_HSQ_Click()
'    stopQ "HUBSOURCE_Q"
'End Sub
'Private Sub cmdStartQ_LSQ_Click()
'    startQ "LOYALTYSOURCE_Q"
'End Sub
'Private Sub stopQ(s As String)
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'Dim Res As Recordset
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = OFF;"
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 360
'    cmd.Execute
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'    Set cmd = Nothing
'
'End Sub
'Private Sub startQ(s As String)
'Dim cmd As New ADODB.Command
'Dim OpenResult As Integer
'Dim Res As Recordset
'    OpenResult = oPC.OpenDBSHort
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = ON;"
'    cmd.CommandType = adCmdText
'    cmd.CommandTimeout = 360
'    cmd.Execute
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'    Set cmd = Nothing
'
'End Sub
'Private Sub cmdRecycle_Click()
'    RecycleErrorLog
'End Sub

'Private Sub RecycleErrorLog()
'Dim OpenResult As Integer
'Dim strCommandFilePath As String
'Dim oTF As New z_TextFile
'Dim fs As New FileSystemObject
'
'    OpenResult = oPC.OpenDBSHort
'        If OpenResult = 0 Then
'            strCommandFilePath = "\\" & oPC.NameOfPC & "\PBKS_S\RecycleErrorLog.SQL"
'            Set oTF = New z_TextFile
'            oTF.OpenTextFile strCommandFilePath
'
'            oTF.WriteToTextFile "USE [Master]"
'            oTF.WriteToTextFile "GO"
'            oTF.WriteToTextFile "EXEC sp_cycle_errorlog ;"
'
'            oTF.WriteToTextFile "GO"
'            oTF.CloseTextFile
'            Set oTF = Nothing
'            If fs.FileExists(strCommandFilePath) Then
'                ExecuteScript strCommandFilePath
'            Else
'                MsgBox "Script file: " & strCommandFilePath & " has not been created", vbOKOnly
'            End If
'            oPC.DisconnectDBShort
'            MsgBox "ERRORLOG recycled"
'        Else
'            MsgBox "Cannot open database. Script has not run"
'        End If
'
'End Sub
'Private Sub ExecuteScript(strCommandFilePath)
'Dim strCommand As String
'Dim Res As Boolean
'Dim fs As New FileSystemObject
'
'    strCommand = "SQLCMD -Usa -P" & oPC.Password & " -S" & oPC.ServerName & " -d" & oPC.DatabaseName & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
'    If fs.FileExists(strCommandFilePath) Then
'        Res = F_7_AB_1_ShellAndWaitSimple(strCommand)
'    End If
'
'
'End Sub
'
