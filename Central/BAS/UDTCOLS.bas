Attribute VB_Name = "UDTCOLS"
Public Type COLProps
    STORECODE As String * 5
    DocumentCaptureDate As Date
    DocumentCode As String * 20
    DocumentIssueDate As Date
    DocumentCapturedBy As String * 10
    CustomerName As String * 25
    CustomerAcno As String * 20
    CustomerPhone As String * 25
    OrderlineRef As String * 50
    OrderlineQty As Double
    OrderlineQtyDispatched As Date
    OrderlineQtyOutstanding As Double
    OrderlineDiscount As Double
    ProductTitle As String * 50
    ProductPublisher As String * 50
    ProductEAN As String * 25
    SellingPrice As Double
    GlobalTRID As String * 40
    COLID As Long
End Type

Public Type COLData
    buffer As String * 352
End Type
Sub lenCOLprops()
Dim X As COLProps
    MsgBox LenB(X) & "        " & LenB(X) / 2
End Sub

