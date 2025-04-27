VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_TextListCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colList As Collection
Private colKeys As Collection

Private Sub Class_Initialize()
  Set colList = New Collection
  Set colKeys = New Collection
End Sub

Public Function Item(ByVal Index As Variant) As String
Attribute Item.VB_UserMemId = 0
  Item = colList.Item(Index)
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colList.[_NewEnum]
End Function

Public Function Key(ByVal text As String) As Integer
  Key = colKeys.Item(text)
End Function

Public Sub Load(col As Collection)
On Error GoTo ERR_Handler
Dim i As Integer
    For i = 1 To col.Count
        colKeys.Add Format$(col.Item(i)(0)), Format$(col.Item(i)(1))
        colList.Add Format$(col.Item(i)(1)), Format$(col.Item(i)(0))
skip:
        Next
EXIT_Handler:
    Exit Sub
ERR_Handler:
    Select Case Err
    Case 457
        Resume skip
    Case Else
        Err.Raise vbObjectError + 557
    End Select
End Sub

Public Property Get Count()
    Count = colList.Count
End Property

Private Sub Class_Terminate()
    Set colList = Nothing
    Set colKeys = Nothing
End Sub
