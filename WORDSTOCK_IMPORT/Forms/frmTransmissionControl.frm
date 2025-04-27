VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransmissionControl 
   Caption         =   "Transmission control"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   5820
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSBMonitor 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Service broker monitor"
      Height          =   450
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2565
      Width           =   2100
   End
   Begin VB.TextBox txtSBStatus 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   1710
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3870
      Width           =   1950
   End
   Begin VB.CommandButton cmdSBToggle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Toggle service broker"
      Height          =   555
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3765
      Width           =   1545
   End
   Begin VB.CommandButton cmdClearDebug 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear _debug"
      Height          =   330
      Left            =   4290
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1155
      Width           =   1485
   End
   Begin VB.CommandButton cmdRecycle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Recycle ERRORLOG"
      Height          =   435
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2775
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefreshTimer 
      Height          =   300
      Left            =   4965
      Picture         =   "frmTransmissionControl.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   30
      Width           =   810
   End
   Begin VB.TextBox txtTimer 
      Height          =   795
      Left            =   15
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   345
      Width           =   5760
   End
   Begin VB.Frame Frame2 
      Caption         =   "SalesSource_Q"
      Height          =   1050
      Left            =   135
      TabIndex        =   5
      Top             =   1440
      Width           =   3525
      Begin VB.CommandButton chkGetStatus_SSQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         Width           =   1635
      End
      Begin VB.CommandButton cmdStartQ_SSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2985
         Picture         =   "frmTransmissionControl.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_SSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   255
         Width           =   360
      End
      Begin VB.Label lblQStatus_SSQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   9
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
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1650
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
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   4530
      Picture         =   "frmTransmissionControl.frx":0A9E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3585
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


Private Sub chkGetStatus_SSQ_Click()
Dim res As Integer
    res = TimerQEnabled("SALESSOURCE_Q")

    If res = 999 Then
        Me.lblQStatus_SSQ.Caption = "SALESSOURCE_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_SSQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_SSQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_SSQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub



Private Sub cmdClearDebug_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
  '  OpenResult = cn.Open
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = "DELETE FROM _tDEBUG"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
End Sub

Private Sub cmdClearQ_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
   ' OpenResult = cn.Open
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = "_ClearQueue"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    cmd.Execute
    Set cmd = Nothing
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
Dim res As Recordset
Dim s As String

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = "SELECT TOP 10 * frOM _tDEBUG Order By ID DESC"
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
    Set cmd = Nothing
    
End Sub

Private Sub cmdSBMonitor_Click()
Dim f As New frmServiceBrokerMonitor
    f.Show vbModal
End Sub

Private Sub cmdStartQ_SSQ_Click()
    startQ "SALESSOURCE_Q"
End Sub
Private Sub cmdStopQ_SSQ_Click()
    stopQ "SALESSOURCE_Q"
End Sub
Private Sub stopQ(s As String)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = OFF;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute

    Set cmd = Nothing

End Sub
Private Sub startQ(s As String)
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = cn
    cmd.CommandText = "ALTER QUEUE " & s & " WITH STATUS = ON;"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    cmd.Execute
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

            strCommandFilePath = "\\" & strPCName & "\PBKS_S\RecycleErrorLog.SQL"
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
            MsgBox "ERRORLOG recycled"

End Sub
Private Sub ExecuteScript(strCommandFilePath)
Dim strCommand As String
Dim res As Boolean
Dim fs As New FileSystemObject
    
    strCommand = "SQLCMD -Usa -P" & gPassword & " -S" & strServerName & " -dPBKS_WSTOCK -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
End Sub

Private Function TimerQEnabled(s As String) As Integer
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
Dim res As Recordset
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "SELECT IS_RECEIVE_ENABLED FROM sys.service_queues WHERE name = '" & s & "'"
    cmd.CommandType = adCmdText
    cmd.CommandTimeout = 360
    Set res = cmd.Execute
    If Not res.State = 0 Then
        If Not res.EOF Then
            TimerQEnabled = CLng(res.Fields(0))
        End If
    Else
        TimerQEnabled = 999
    End If
    Set cmd = Nothing
End Function

Private Function CheckBrokerEnabled() As Boolean
Dim rs As New ADODB.Recordset
Dim bEnabled As Boolean
        rs.Open "SELECT is_broker_enabled FROM master.sys.databases where name = 'PBKS_WSTOCK'", cn, adOpenKeyset
        If rs.State <> 0 Then
            If rs.EOF <> True Then
                bEnabled = rs.Fields(0)
            Else
                bEnabled = False
            End If
        Else
            bEnabled = False
        End If
        rs.Close
    CheckBrokerEnabled = bEnabled
End Function

Private Sub Form_Load()
    txtSBStatus = IIf(CheckBrokerEnabled, "Enabled", "Disabled")
End Sub
Private Sub cmdSBToggle_Click()
        cn.CommandTimeout = 30
        On Error Resume Next
        If Me.txtSBStatus = "Disabled" Then
            cn.Execute "ALTER DATABASE  PBKS_WSTOCK SET ENABLE_BROKER"
            If Err <> 0 Then
                MsgBox "The following error occurred: " & Error
            End If
        Else
            cn.Execute "ALTER DATABASE  PBKS_WSTOCK SET DISABLE_BROKER"
            If Err <> 0 Then
                MsgBox "The following error occurred: " & Error
            End If
        End If
    txtSBStatus = IIf(CheckBrokerEnabled, "Enabled", "Disabled")
    
End Sub

