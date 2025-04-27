VERSION 5.00
Begin VB.Form frmDeadLetterQueue 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Undelivered sales transaction"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRetry 
      BackColor       =   &H00DACDCD&
      Cancel          =   -1  'True
      Caption         =   "&Retry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4905
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3765
      Width           =   1260
   End
   Begin VB.TextBox txtDeadLetters 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3480
      IMEMode         =   3  'DISABLE
      Left            =   225
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   180
      Width           =   5925
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00DACDCD&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3780
      Width           =   1260
   End
End
Attribute VB_Name = "frmDeadLetterQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDL As ADODB.Recordset

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdRetry_Click()
    On Error GoTo errHandler
Dim qinfo As New MSMQQueueInfo
Dim q As MSMQQueue
Dim msg As MSMQMessage
Dim NewsMsg As MSMQMessage

Dim strMachineId As String

  ' Set the format name of the dead-letter queue.
    qinfo.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\SYSTEM$;DEADLETTER"
' Open the dead-letter queue.
    Set q = qinfo.Open(Access:=MQ_RECEIVE_ACCESS, ShareMode:=MQ_DENY_NONE)
        
    ' Read the firsDLt message in the dead-letter queue.
    Set msg = q.Receive(ReceiveTimeout:=1000)
    ' Read the remaining messages in the dead-letter queue.
    Do While Not msg Is Nothing
        Set NewsMsg = New MSMQMessage
        NewsMsg.Delivery = msg.Delivery
        NewsMsg.Journal = msg.Journal
      '  NewsMsg.MaxTimeToReachQueue = msg.MaxTimeToReachQueue
        NewsMsg.Label = msg.Label
        NewsMsg.Body = msg.Body

        DispatchMessage NewsMsg
        Set NewsMsg = Nothing
        Set msg = Nothing
        Set msg = q.Receive(ReceiveTimeout:=1000)
    Loop
    
    q.Close

    Exit Sub
errHandler:
    ErrPreserve
    If Not q Is Nothing And q.IsOpen2 Then
        q.Close
    End If
    Set q = Nothing

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDeadLetterQueue.cmdRetry_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub DispatchMessage(pMsg As MSMQMessage)
Dim QI As New MSMQQueueInfo
Dim POSmsg As MSMQMessage
Dim QPOS As MSMQQueue

    Set QI = New MSMQQueueInfo
    QI.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\qpos"
    Set QPOS = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
    QI.FormatName = "DIRECT=OS:" & oPC.NameOfPC & "\Private$\qposack"

    Set POSmsg = pMsg
    POSmsg.Send QPOS


End Sub
Private Sub Form_Load()
    On Error GoTo errHandler
    Set rsDL = New ADODB.Recordset
    rsDL.Fields.Append "Head", adVarChar, 100
    rsDL.Fields.Append "Dte", adDate
    GetDeadLetterMessages
    LoadDeadLetterMessages
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDeadLetterQueue.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Sub GetDeadLetterMessages()
    On Error GoTo errHandler
Dim qinfo As New MSMQQueueInfo
Dim q As MSMQQueue
Dim msg As MSMQMessage
Dim strMachineId As String

  ' Set the format name of the dead-letter queue.
    qinfo.FormatName = "DIRECT=OS:" & oPC.NameOfPC & "\SYSTEM$;DEADLETTER"

' Open the dead-letter queue.
    Set q = qinfo.Open(Access:=MQ_RECEIVE_ACCESS, ShareMode:=MQ_DENY_NONE)
        
  ' Read the firsDLt message in the dead-letter queue.
    Set msg = q.PeekCurrent(ReceiveTimeout:=1000)
    rsDL.Open
    ' Read the remaining messages in the dead-letter queue.
    Do While Not msg Is Nothing
      rsDL.AddNew
          rsDL.Fields("Head") = msg.Label
          rsDL.Fields("Dte") = msg.SentTime
      rsDL.Update
      Set msg = q.PeekNext(ReceiveTimeout:=1000)
    Loop

  '  rsDL.Close
  '  Set rsDL = Nothing
    q.Close
    
    Exit Sub
errHandler:
    ErrPreserve
    If Not q Is Nothing And q.IsOpen2 Then
        q.Close
    End If
    If Not rsDL Is Nothing And rsDL.State = adStateOpen Then
        rsDL.Close
    End If
    Set rsDL = Nothing
    Set q = Nothing
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDeadLetterQueue.GetDeadLetterMessages"
End Sub

Sub LoadDeadLetterMessages()
    On Error GoTo errHandler
Dim strMessages As String
    If rsDL.EOF Then
        txtDeadLetters = "No transactions in queue"
    Else
        rsDL.MoveFirst
        Do While Not rsDL.EOF
            strMessages = strMessages & IIf(strMessages > "", vbCrLf, "") & rsDL.Fields(0)
            rsDL.MoveNext
        Loop
        txtDeadLetters = strMessages
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDeadLetterQueue.LoadDeadLetterMessages"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    rsDL.Close
    Set rsDL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDeadLetterQueue.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
