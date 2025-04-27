Attribute VB_Name = "UDTCatChk"
Public Type dCatChkProps
    CATCHKID As Long
    SupplierName As String * 100
    CategoryName As String * 200
    CategoryCode As String * 10
    DOCDate As Date
    ProcessingDate As Date
    DOCCode As String * 40
    Status As Long
    OperatorID As Long
    OperatorName As String * 100
    OperatorShortName As String * 10
    SupervisorID As Long
    SupervisorShortName As String * 10
    SupervisorName As String * 10
    SignedOffByID As Long
    SignedOffShortName As String * 10
    SignedOffName As String * 10
End Type
Public Type dCatChkData
    buffer As String * 1044
End Type

Public Sub test()
Dim f As dCatChkProps
    MsgBox LenB(f)
End Sub


