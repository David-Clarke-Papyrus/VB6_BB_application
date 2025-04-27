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
    Set colBroken = New Collection
'    Set colDescription = New Collection
End Sub
Public Sub BreakRule(rule As String, IsBroken As Boolean)
On Error GoTo HandleError
    If IsBroken Then
        colBroken.Add rule, rule
    Else
        colBroken.Remove rule
    End If
HandleError:
End Sub
Public Sub GetStatus()
    RaiseEvent Status(AllBrokenRules)
End Sub
Public Sub RuleBroken(rule As String, IsBroken As Boolean)
On Error GoTo HandleError

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
HandleError:
'MsgBox Error
End Sub

Public Property Get Count() As Integer
    Count = colBroken.Count
End Property
Public Property Get Description(rule) As String
On Error Resume Next
    Description = ""
    Description = colBroken(rule)
End Property
Public Property Get AllBrokenRules() As String
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
End Property

Private Sub Class_Terminate()
    Set colBroken = Nothing
End Sub
