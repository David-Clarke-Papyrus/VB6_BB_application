VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransmissionControl 
   Caption         =   "Transmission control"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Invocation_RQ_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   46
      Top             =   5775
      Width           =   2820
      Begin VB.CommandButton cmdStopQ_INQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_INQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_INQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_INQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   50
         Top             =   615
         Width           =   2610
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Alert_Consumer_Q"
      Height          =   1050
      Left            =   2985
      TabIndex        =   41
      Top             =   2385
      Width           =   2820
      Begin VB.CommandButton cmdStopQ_AlertQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1830
         Picture         =   "frmTransmissionControl.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_ALERTQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":0A9E
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_AlertQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_AlertQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   45
         Top             =   615
         Width           =   2610
      End
   End
   Begin VB.CommandButton cmdSBMonitor 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Service broker monitor"
      Height          =   420
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3015
      Width           =   1800
   End
   Begin VB.CommandButton cmdSBToggle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Toggle service broker"
      Height          =   495
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4590
      Width           =   1380
   End
   Begin VB.TextBox txtSBStatus 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   6540
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   5100
      Width           =   1395
   End
   Begin VB.CommandButton cmdResend 
      Caption         =   "Re-send"
      Height          =   300
      Left            =   6480
      TabIndex        =   36
      Top             =   4215
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTResend 
      Height          =   330
      Left            =   6525
      TabIndex        =   34
      Top             =   3840
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      _Version        =   393216
      Format          =   60817409
      CurrentDate     =   39696
   End
   Begin VB.Frame Frame4 
      Caption         =   "MasterLoyaltyConsumer_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   29
      Top             =   4635
      Width           =   2820
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
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":0E28
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_MLCQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":11B2
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
         Width           =   2610
      End
   End
   Begin VB.CommandButton cmdClearDebug 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear _debug"
      Height          =   330
      Left            =   870
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6765
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton cmdRecycle 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Recycle ERRORLOG"
      Height          =   360
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2640
      Width           =   1800
   End
   Begin VB.CommandButton cmdRefreshTimer 
      Height          =   300
      Left            =   7350
      Picture         =   "frmTransmissionControl.frx":153C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   30
      Width           =   810
   End
   Begin VB.TextBox txtTimer 
      Height          =   795
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   7110
      Width           =   7710
   End
   Begin VB.Frame Frame3 
      Caption         =   "HubSource_Q"
      Height          =   1050
      Left            =   3000
      TabIndex        =   18
      Top             =   4635
      Width           =   2820
      Begin VB.CommandButton cmdStopQ_HSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":18C6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_HSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":1C50
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
         Width           =   2625
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SalesSource_Q"
      Height          =   1050
      Left            =   3000
      TabIndex        =   13
      Top             =   3510
      Width           =   2820
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
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":1FDA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStopQ_SSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":2364
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
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LoyaltySource_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   8
      Top             =   3510
      Width           =   2820
      Begin VB.CommandButton cmdStopQ_LSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":26EE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_LSQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":2A78
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
         Width           =   2610
      End
   End
   Begin VB.Frame frSOHQ 
      Caption         =   "SOHSource_Q"
      Height          =   1050
      Left            =   75
      TabIndex        =   5
      Top             =   2385
      Width           =   2820
      Begin VB.CommandButton cmdStopQ_SOHQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   1845
         Picture         =   "frmTransmissionControl.frx":2E02
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   360
      End
      Begin VB.CommandButton cmdStartQ_SOHQ 
         BackColor       =   &H00C4BCA4&
         Height          =   330
         Left            =   2295
         Picture         =   "frmTransmissionControl.frx":318C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton chkGetStatus_SOHQ 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Get status"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label lblQStatus_SOHQ 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   615
         Width           =   2580
      End
   End
   Begin VB.CommandButton cmdOpenLog 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Open log"
      Height          =   330
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2295
      Width           =   1800
   End
   Begin VB.CommandButton cmdClearQ 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Clear queue"
      Height          =   555
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6630
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Stop timer"
      Height          =   330
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7005
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Start timer"
      Height          =   330
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7005
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   615
      Left            =   6930
      Picture         =   "frmTransmissionControl.frx":3516
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5655
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1890
      Left            =   60
      TabIndex        =   40
      Top             =   375
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   3334
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
   Begin VB.Label Label1 
      Caption         =   "Re-send sales data"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6555
      TabIndex        =   35
      Top             =   3615
      Width           =   1440
   End
End
Attribute VB_Name = "frmTransmissionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFileName As String
Dim rs2 As ADODB.Recordset

Private Sub chkGetStatus_AlertQ_Click()
Dim res As Integer
    res = TimerQEnabled("ALERT_CONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_AlertQ.Caption = "Timer queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_AlertQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_AlertQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_AlertQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_HSQ_Click()
Dim res As Integer
    res = TimerQEnabled("HUBSOURCE_Q")

    If res = 999 Then
        Me.lblQStatus_HSQ.Caption = "HUBSOURCE_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_HSQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_HSQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_INQ_Click()
Dim res As Integer
    res = TimerQEnabled("INVOCATION_RQ_Q")

    If res = 999 Then
        Me.lblQStatus_INQ.Caption = "INVOCATION_RQ_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_INQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_INQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_INQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_LSQ_Click()
Dim res As Integer
    res = TimerQEnabled("LOYALTYSOURCE_Q")

    If res = 999 Then
        Me.lblQStatus_LSQ.Caption = "LOYALTYSOURCE_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_LSQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_LSQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_LSQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

Private Sub chkGetStatus_MLCQ_Click()
Dim res As Integer
    res = TimerQEnabled("MASTERLOYALTYCONSUMER_Q")

    If res = 999 Then
        Me.lblQStatus_MLCQ.Caption = "MASTERLOYALTYCONSUMER_Q queue cannot be found"
    Else
        If res = -1 Then
            Me.lblQStatus_MLCQ.Caption = "IS_RECEIVE_ENABLED = true"
        Else
            If res = 0 Then
                Me.lblQStatus_MLCQ.Caption = "IS_RECEIVE_ENABLED = false"
            Else
                Me.lblQStatus_MLCQ.Caption = "Unknown (" & CStr(res) & ")"
            End If
        End If
    End If

End Sub

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

Private Sub chkGetStatus_SOHQ_Click()
Dim res As Integer
    res = TimerQEnabled("SOHSOURCE_Q")

    If res = 999 Then
        Me.lblQStatus_SOHQ.Caption = "Timer queue cannot be found"
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


'Private Sub chkGetStatus_TQ_Click()
'
'End Sub

Private Sub cmdClearDebug_Click()
Dim cmd As New ADODB.Command
Dim OpenResult As Integer
    
    OpenResult = oPC.OpenDBSHort
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "DELETE FROM _tDEBUG"
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

Private Sub cmdResend_Click()
Dim oSQL As New z_SQL
    If MsgBox("You are wanting to re-send the sales data to Central for the day: " & CStr(Format(Me.DTResend, "DD-MM-YYYY")) & ". Confirm", vbQuestion + vbOKCancel, "Confim") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    oSQL.ReSendSalesToCentral Me.DTResend
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSBMonitor_Click()
Dim f As New frmServiceBrokerMonitor
    f.Show vbModal
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




Private Sub cmdStartQ_ALERTQ_Click()
    startQ "ALERT_CONSUMER_Q"

End Sub

Private Sub cmdStartQ_INQ_Click()
    startQ "INVOCATION_RQ_Q"
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

Private Function TimerQEnabled(s As String) As Integer
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
            TimerQEnabled = CLng(res.Fields(0))
        End If
    Else
        TimerQEnabled = 999
    End If
    If OpenResult = 0 Then oPC.DisconnectDBShort
    Set cmd = Nothing
End Function

Private Sub cmdStartQ_SOHQ_Click()
    startQ "SOHSOURCE_Q"
End Sub

Private Sub cmdStopQ_AlertQ_Click()
    stopQ "ALERT_CONSUMER_Q"

End Sub

Private Sub cmdStopQ_INQ_Click()
    stopQ "INVOCATION_RQ_Q"
End Sub

Private Sub cmdStopQ_MLCQ_Click()
    stopQ "MASTERLOYALTYCONSUMER_Q"
End Sub

Private Sub cmdStopQ_SOHQ_Click()
    stopQ "TimerQueue"
End Sub
Private Sub cmdStopQ_LSQ_Click()
    stopQ "LOYALTYSOURCE_Q"
End Sub
Private Sub cmdStartQ_SSQ_Click()
    startQ "SALESSOURCE_Q"
End Sub
Private Sub cmdStartQ_MLCQ_Click()
    startQ "MASTERLOYALTYCONSUMER_Q"
End Sub
Private Sub cmdStopQ_SSQ_Click()
    stopQ "SALESSOURCE_Q"
End Sub
Private Sub cmdStartQ_HSQ_Click()
    startQ "HUBSOURCE_Q"
End Sub
Private Sub cmdStopQ_HSQ_Click()
    stopQ "HUBSOURCE_Q"
End Sub
Private Sub cmdStartQ_LSQ_Click()
    startQ "LOYALTYSOURCE_Q"
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
    
    strCommand = "SQLCMD -Usa -P" & oPC.Password & " -S" & oPC.ServerName & " -d" & oPC.DatabaseName & " -i" & strCommandFilePath & " -o" & Replace(strCommandFilePath, ".SQL", ".ERR")
    If fs.FileExists(strCommandFilePath) Then
        res = F_7_AB_1_ShellAndWaitSimple(strCommand)
    End If
    
    
End Sub

Private Function CheckBrokerEnabled() As Boolean
Dim rs As New ADODB.Recordset
Dim bEnabled As Boolean
        rs.Open "SELECT is_broker_enabled FROM master.sys.databases where name = 'PBKS'", oPC.CO, adOpenKeyset
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

Private Sub cmdSBToggle_Click()
        oPC.CO.CommandTimeout = 30
        On Error Resume Next
        If Me.txtSBStatus = "Disabled" Then
            oPC.CO.Execute "ALTER DATABASE  PBKS SET ENABLE_BROKER"
            If Err <> 0 Then
                MsgBox "The following error occurred: " & Error
            End If
        Else
            oPC.CO.Execute "ALTER DATABASE  PBKS SET DISABLE_BROKER"
            If Err <> 0 Then
                MsgBox "The following error occurred: " & Error
            End If
        End If
    txtSBStatus = IIf(CheckBrokerEnabled, "Enabled", "Disabled")
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
    txtSBStatus = IIf(CheckBrokerEnabled, "Enabled", "Disabled")
End Sub

