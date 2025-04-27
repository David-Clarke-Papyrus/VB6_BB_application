VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransmissionControl 
   Caption         =   "Transmission control"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClearDebug 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear _debug"
      Height          =   330
      Left            =   4275
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3690
      Width           =   1485
   End
   Begin VB.CommandButton cmdRecycle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Recycle ERRORLOG"
      Height          =   435
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4620
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefreshTimer 
      Height          =   300
      Left            =   4965
      Picture         =   "frmTransmissionControl.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   30
      Width           =   810
   End
   Begin VB.TextBox txtTimer 
      Height          =   3255
      Left            =   15
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   345
      Width           =   5760
   End
   Begin VB.Frame Frame3 
      Caption         =   "HubConsumer_Q"
      Height          =   1050
      Left            =   90
      TabIndex        =   3
      Top             =   4020
      Width           =   3525
      Begin VB.CommandButton cmdStopQ_HSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_HSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControl.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_HSQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_HSQ 
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
      Left            =   3975
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4125
      Width           =   1800
   End
   Begin VB.CommandButton cmdClearQ 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear queue"
      Height          =   555
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   90
      Picture         =   "frmTransmissionControl.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5130
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
Attribute VB_Name = "frmTransmissionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String

Private Sub chkGetStatus_HSQ_Click()
Dim Res As Integer
    Res = TimerQEnabled("HUBCONSUMER_Q")

    If Res = 999 Then
        Me.lblQStatus_HSQ.Caption = "HUBCONSUMER_Q queue cannot be found"
    Else
        If Res = -1 Then
            Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If Res = 0 Then
                Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_HSQ.Caption = "Unknown (" & CStr(Res) & ")"
            End If
        End If
    End If

End Sub



Private Sub cmdClearDebug_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "DELETE FROM _tDEBUG"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oCnn.CloseDB
End Sub

Private Sub cmdClearQ_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "_ClearQueue"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oCnn.CloseDB
End Sub

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


Private Sub cmdRefreshTimer_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim Res As Recordset
Dim s As String

    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "SELECT TOP 10 * frOM _tDEBUG Order By ID DESC"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    Set Res = cmd.Execute
    s = ""
    If Not Res.State = 0 Then
        Do While Not Res.EOF
            s = s & Res.Fields(1) & vbCrLf
            Res.MoveNext
        Loop
        txtTimer = s
    Else
        txtTimer = ""
    End If
    If OpenResult = 0 Then oCnn.CloseDB
    Set cmd = Nothing
    
End Sub

Private Sub cmdStart_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "_StartTimer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oCnn.CloseDB
    
End Sub



Private Sub cmdStop_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "_EndTimer"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
    If OpenResult = 0 Then oCnn.CloseDB
End Sub

Private Function TimerQEnabled(s As String) As Integer
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim Res As Recordset
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "SELECT IS_RECEIVE_ENABLED FROM sys.service_queues WHERE name = '" & s & "'"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    Set Res = cmd.Execute
    If Not Res.State = 0 Then
        If Not Res.EOF Then
            TimerQEnabled = CLng(Res.Fields(0))
        End If
    Else
        TimerQEnabled = 999
    End If
    If OpenResult = 0 Then oCnn.CloseDB
    Set cmd = Nothing
End Function


Private Sub cmdStartQ_HSQ_Click()
    startQ "HUBCONSUMER_Q"
End Sub
Private Sub cmdStopQ_HSQ_Click()
    stopQ "HUBCONSUMER_Q"
End Sub
Private Sub stopQ(s As String)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim Res As Recordset
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = OFF;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    If OpenResult = 0 Then oCnn.CloseDB
    Set cmd = Nothing

End Sub
Private Sub startQ(s As String)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim Res As Recordset
    OpenResult = oCnn.OpenDB
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oCnn.Connection
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = ON;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    If OpenResult = 0 Then oCnn.CloseDB
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

    OpenResult = oCnn.OpenDB
        If OpenResult = 0 Then
            strCommandFilePath = "\\" & oCnn.NameOfPC & "\PBKS_S\RecycleErrorLog.SQL"
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
            oCnn.CloseDB
            MsgBox "ERRORLOG recycled"
        Else
            MsgBox "Cannot open database. Script has not run"
        End If

End Sub
Private Sub ExecuteScript(strCommandFilePath)
Dim strCommand As String
Dim Res As Boolean
Dim fs As New FileSystemObject
    
    strCommand = "SQLCMD -Usa -P" & oCnn.Password & " -S" & oCnn.ServerName & " -dHUB -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        Res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
End Sub

