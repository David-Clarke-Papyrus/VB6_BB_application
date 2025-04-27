VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransmissionControl 
   Caption         =   "Transmission control"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      Caption         =   "Paste consumer queue"
      Height          =   1050
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_PASQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1875
         Picture         =   "frmTransmissionControl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_PASQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControl.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_PASQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_PASQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   11
         Top             =   615
         Width           =   3315
      End
   End
   Begin VB.CommandButton cmdClearDebug 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear SB log"
      Height          =   330
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1485
   End
   Begin VB.CommandButton cmdRecycle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Recycle ERRORLOG"
      Height          =   435
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefreshTimer 
      Height          =   300
      Left            =   7305
      Picture         =   "frmTransmissionControl.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   810
   End
   Begin VB.TextBox txtTimer 
      Height          =   1395
      Left            =   15
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   345
      Width           =   8130
   End
   Begin VB.CommandButton cmdOpenLog 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Open log"
      Height          =   420
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1815
      Width           =   1800
   End
   Begin VB.CommandButton cmdClearQ 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear queue"
      Height          =   555
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3165
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   6330
      Picture         =   "frmTransmissionControl.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3075
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmTransmissionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilename As String
Dim strCommandFilePath As String

'
'Private Sub chkGetStatus_HSQ_Click()
'Dim res As Integer
'    res = TimerQEnabled("HUBSOURCE_Q")
'
'    If res = 999 Then
'        Me.lblQStatus_HSQ.Caption = "HUBSOURCE_Q queue cannot be found"
'    Else
'        If res = -1 Then
'            Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If res = 0 Then
'                Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_HSQ.Caption = "Unknown (" & CStr(res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_LSQ_Click()
'Dim res As Integer
'    res = TimerQEnabled("LOYALTYSOURCE_Q")
'
'    If res = 999 Then
'        Me.lblQStatus_LSQ.Caption = "LOYALTYSOURCE_Q queue cannot be found"
'    Else
'        If res = -1 Then
'            Me.lblQStatus_LSQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If res = 0 Then
'                Me.lblQStatus_LSQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_LSQ.Caption = "Unknown (" & CStr(res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_MLCQ_Click()
'Dim res As Integer
'    res = TimerQEnabled("MASTERLOYALTYCONSUMER_Q")
'
'    If res = 999 Then
'        Me.lblQStatus_MLCQ.Caption = "MASTERLOYALTYCONSUMER_Q queue cannot be found"
'    Else
'        If res = -1 Then
'            Me.lblQStatus_MLCQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If res = 0 Then
'                Me.lblQStatus_MLCQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_MLCQ.Caption = "Unknown (" & CStr(res) & ")"
'            End If
'        End If
'    End If
'
'End Sub

Private Sub chkGetStatus_PASQ_Click()
Dim res As Integer
    res = IsQEnabled("PASTELCONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_PASQ.Caption = "PASTELCONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_PASQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_PASQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_PASQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

'Private Sub chkGetStatus_SSQ_Click()
'Dim res As Integer
'    res = TimerQEnabled("SALESSOURCE_Q")
'
'    If res = 999 Then
'        Me.lblQStatus_SSQ.Caption = "SALESSOURCE_Q queue cannot be found"
'    Else
'        If res = -1 Then
'            Me.lblQStatus_SSQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If res = 0 Then
'                Me.lblQStatus_SSQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_SSQ.Caption = "Unknown (" & CStr(res) & ")"
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub chkGetStatus_TQ_Click()
'Dim res As Integer
'    res = TimerQEnabled("TimerQueue")
'
'    If res = 999 Then
'        Me.lblQStatus_TQ.Caption = "Timer queue cannot be found"
'    Else
'        If res = -1 Then
'            Me.lblQStatus_TQ.Caption = "IS_RECEIVE_ENABLED = true"
'        Else
'            If res = 0 Then
'                Me.lblQStatus_TQ.Caption = "IS_RECEIVE_ENABLED = false"
'            Else
'                Me.lblQStatus_TQ.Caption = "Unknown (" & CStr(res) & ")"
'            End If
'        End If
'    End If
'End Sub


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

    strFilename = GetSetting("PBKS", "SB", "LOGFILEPATH", "")
    If fs.GetBaseName(strFilename) <> "ERRORLOG" Then
        CD1.DialogTitle = "Open SQL Server log file"
        CD1.DefaultExt = ""
        CD1.InitDir = "c:\Program files\Microsoft SQL SERVER"
        CD1.ShowOpen
        strFilename = CD1.FileName
        SaveSetting "PBKS", "SB", "LOGFILEPATH", strFilename
    End If
    
End Sub


Private Sub cmdRefreshTimer_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
Dim s As String

    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT TOP 10 * frOM _tSBLog Order By SBL_DATE  DESC"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    Set res = cmd.Execute
    s = ""
    If Not res.State = 0 Then
        Do While Not res.EOF
            s = s & res.Fields(1) & vbCrLf
            res.MoveNext
        Loop
        txtTimer = s
    Else
        txtTimer = ""
    End If
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Set cmd = Nothing
    
End Sub

Private Sub cmdStart_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "_StartTimer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oPC.DisconnectDBShort
    
End Sub




Private Sub cmdStop_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "_EndTimer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oPC.DisconnectDBShort
End Sub

Private Function IsQEnabled(s As String) As Integer
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT IS_RECEIVE_ENABLED FROM sys.service_queues WHERE name = '" & s & "'"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    Set res = cmd.Execute
    If Not res.State = 0 Then
        If Not res.EOF Then
            IsQEnabled = CLng(res.Fields(0))
        End If
    Else
        IsQEnabled = 999
    End If
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Set cmd = Nothing
End Function

'Private Sub cmdStartQ_TQ_Click()
'    startQ "TimerQueue"
'End Sub

'Private Sub cmdStopQ_MLCQ_Click()
'    stopQ "MASTERLOYALTYCONSUMER_Q"
'End Sub

Private Sub cmdStopQ_PASQ_Click()
    stopQ "PASTELCONSUMER_Q"
End Sub
Private Sub cmdStartQ_PASQ_Click()
    startQ "PASTELCONSUMER_Q"

End Sub


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
    On Error GoTo errHandler
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTransmissionControl.startQ(s)", s
End Sub
Private Sub cmdRecycle_Click()
    RecycleErrorLog
End Sub

Private Sub RecycleErrorLog()
    On Error GoTo errHandler
Dim OpenResult As Integer
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTransmissionControl.RecycleErrorLog"
End Sub
Private Sub ExecuteScript(strCommandFilePath)
    On Error GoTo errHandler
Dim strCommand As String
Dim res As Boolean
Dim fs As New FileSystemObject
    
    strCommand = "SQLCMD -Usa -P" & oPC.Password & " -S" & strMainSQLServerName & " -d" & oPC.DBName & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmTransmissionControl.ExecuteScript(strCommandFilePath)", strCommandFilePath
End Sub

