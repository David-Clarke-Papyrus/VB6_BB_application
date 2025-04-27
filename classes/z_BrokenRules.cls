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
Attribute VB_Exposed = False
Option Explicit

Event BrokenRule(oRS As String)
Event RuleUnbroken(oRS As String)
Event NoBrokenRules()
Event ThisRuleUnBroken(por As String)
Event Status(oRS As String)
Private colBroken As Collection
'Private colDescription As Collection


Private Sub Class_Initialize()
    Set colBroken = New Collection
'    Set colDescription = New Collection
End Sub
Public Sub BreakRule(Rule As String, IsBroken As Boolean)
On Error GoTo Handleor
    If IsBroken Then
        colBroken.Add Rule, Rule
    Else
        colBroken.Remove Rule
    End If
Handleor:
End Sub
Public Sub GetStatus()
    RaiseEvent Status(AllBrokenRules)
End Sub
Public Sub RuleBroken(Rule As String, IsBroken As Boolean)
On Error GoTo Handleor

    If IsBroken Then
        colBroken.Add Rule, Rule
        RaiseEvent BrokenRule(AllBrokenRules)
    Else
        colBroken.Remove Rule
        If colBroken.Count = 0 Then
            RaiseEvent NoBrokenRules
        Else
            RaiseEvent RuleUnbroken(AllBrokenRules)
        End If
    End If
Handleor:
End Sub

Public Property Get Count() As Integer
    Count = colBroken.Count
End Property
Public Property Get Description(Rule) As String
On Error Resume Next
    Description = ""
    Description = colBroken(Rule)
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
