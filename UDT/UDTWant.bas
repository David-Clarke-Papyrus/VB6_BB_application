Attribute VB_Name = "UDTWant"
Public Type WantProps
    ID As Long
    TPID As Long
    PID As String * 40
    CustomerName As String * 50
    RequestDate As Date
    Notes As String * 200
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type WantData
     buffer As String * 302
End Type

Sub lenWantProps()
Dim x As WantProps
    MsgBox LenB(x) & "        " & LenB(x) / 2
End Sub
