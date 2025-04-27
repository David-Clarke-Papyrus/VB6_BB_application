Attribute VB_Name = "UDTDeal"
Public Type DealProps
    ID As Long
    TPID As Long
    Discount As Double
    Description As String * 15
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type DealData
    buffer As String * 26
End Type

Sub lenDeal()
Dim x As DealProps
    MsgBox LenB(x) & "    " & LenB(x) / 2
End Sub

