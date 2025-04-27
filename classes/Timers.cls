VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim WithEvents oXT As XTimer
Attribute oXT.VB_VarHelpID = -1
Dim TickCount As Long

Private Sub Class_Initialize()
    Set oXT = New XTimer
    oXT.Interval = 1000
End Sub

Private Sub oXT_Tick()
    On Error GoTo errHandler
    oXT.Enabled = False
        
        
    
    oXT.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsPoll.oTimer_Tick"
End Sub

Public Sub SetActionTime(iHour As Integer, iMinute As Integer)
Dim lngTicksTillAction As Long

   ' DateDiff("mm",
End Sub
