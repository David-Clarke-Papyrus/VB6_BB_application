VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_BrokenRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event BrokenRule(Errors As String)
Event RuleUnbroken(Errors As String)
Event NoBrokenRules()
Event Status(Errors As String)
Private colBroken As Collection
'Private colDescription As Collection


Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set colBroken = New Collection
'    Set colDescription = New Collection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.Class_Initialize"
    
End Sub
Public Sub BreakRule(rule As String, IsBroken As Boolean)
    On Error GoTo errHandler
    If IsBroken Then
        colBroken.Add rule, rule
    Else
        colBroken.Remove rule
    End If
':
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.BreakRule(rule,IsBroken)", Array(rule, IsBroken), EA_NORERAISE
    
End Sub
Public Sub GetStatus()
    On Error GoTo errHandler
    RaiseEvent Status(AllBrokenRules)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.GetStatus"
    
End Sub
Public Sub RuleBroken(rule As String, IsBroken As Boolean)
    On Error GoTo errHandler

    If IsBroken Then
        colBroken.Add rule, rule
   '     colDescription.Add rule, rule
        RaiseEvent BrokenRule(AllBrokenRules)
    Else
        colBroken.Remove rule
        If colBroken.Count = 0 Then
            RaiseEvent NoBrokenRules
        Else
            RaiseEvent RuleUnbroken(AllBrokenRules)
        End If
    End If
':
'MsgBox Error
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.RuleBroken(rule,IsBroken)", Array(rule, IsBroken), EA_NORERAISE
    
End Sub

Public Property Get Count() As Integer
    On Error GoTo errHandler
    Count = colBroken.Count
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.Count"
    
End Property
Public Property Get Description(rule) As String
    On Error GoTo errHandler
    Description = ""
    Description = colBroken(rule)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.Description(rule)", rule, EA_NORERAISE
    
End Property
Public Property Get AllBrokenRules() As String
    On Error GoTo errHandler
Dim strAllRules As String
Dim i As Integer
    strAllRules = ""
    For i = 1 To colBroken.Count
        If Len(strAllRules) = 0 Then
            strAllRules = colBroken(i)
        Else
            strAllRules = strAllRules & "," & colBroken(i)
        End If
    Next
    AllBrokenRules = strAllRules
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.AllBrokenRules"
    
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set colBroken = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_BrokenRules.Class_Terminate"
    
End Sub
