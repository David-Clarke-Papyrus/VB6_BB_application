Attribute VB_Name = "UDTOffer"
Public Type OfferProps
    ID As Long
    TPID As Long
    Offeponse As Long
    PID As String * 40
    CustomerName As String * 50
    OfferCode As String * 15
    Serial As String * 5
    Title As String * 40
    RequestDate As Date
    OfferDate As Date
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type OfferData
     buffer As String * 168
End Type

Sub lenOfferProps()
Dim x As OfferProps
    MsgBox LenB(x) & "        " & LenB(x) / 2
End Sub

