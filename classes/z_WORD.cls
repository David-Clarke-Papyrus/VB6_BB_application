VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PapyWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim oWD As Word.Application
Dim oDoc As Word.Document
Public Sub StartWORD()
On Error Resume Next
    Set oWD = GetObject(, "Word.Application")

    If Err <> 0 Then
        Err.Clear
        Set oWD = CreateObject("Word.Application")
    End If
    If Err Then Exit Sub
    On Error GoTo 0
    oWD.Visible = False  ' change to false when complete

End Sub
Public Sub TerminateWORD()
On Error GoTo ERR_Handler
    oWD.Quit wdDoNotSaveChanges
    Set oWD = Nothing
    If Not oWD Is Nothing Then Set oWD = Nothing
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
End Sub
Property Get WORDREF() As Object
    Set WORDREF = oWD
End Property
Function NewDoc()
    Set oDoc = oWD.Documents.Add(tmp) 'tmp)
End Function
Property Get ActiveDoc() As Word.Document
    Set ActiveDoc = oDoc
End Property

