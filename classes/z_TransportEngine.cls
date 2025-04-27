VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_TransportEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim WithEvents timA As XTimer
Attribute timA.VB_VarHelpID = -1
Dim WithEvents timB As XTimer
Attribute timB.VB_VarHelpID = -1
Dim TickCount As Long
Dim QueueIP As Collection
Dim QueueOP As Collection
Public Type QueueOPItem
    FilePath As String * 300
    Item As String * 255
    f3 As String * 10
    F4 As String * 10
    Active As Boolean
End Type

Private Sub Class_Initialize()
    Set timA = New XTimer
    timA.Interval = 1000
    Set timB = New XTimer
    timB.Interval = 1000
End Sub

Private Sub timA_Tick()
    On Error GoTo errHandler
    timA.Enabled = False
    timB.Enabled = False
        
    EvaluateFacts
    For each
    DoAllActions
    
    timA.Enabled = True
    timB.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TransportEngine.timA_Tick"
End Sub
Private Sub timB_Tick()
    On Error GoTo errHandler
    timA.Enabled = False
    timB.Enabled = False
        
    DoAllActions
    
    timA.Enabled = True
    timB.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_TransportEngine.timB_Tick"
End Sub

Public Sub SetActionTime(iHour As Integer, iMinute As Integer)
Dim lngTicksTillAction As Long

   ' DateDiff("mm",
End Sub
