Attribute VB_Name = "UDTCashSale"
Public Type CSProps
    TRID As Long
    TPID As Long
    DOCCode As String * 10
    FrontDeskComputerName As String * 250
    TextFileFullPathAndName As String * 250
    DateStarted As Date
    DateIssued As Date
    CaptureDate As Date
    TILLID As Long
    TAType As Integer
    Status As Long
    Void As Boolean
    TotalExtension As Long
    TotalVAT As Long
    TotalDiscount As Long
    TotalPayable As Long
    TotalExtensionSimple As Long
    StaffID As Long
    StaffName As String * 10
    CS_GUID_ID As String * 40
    Reportable As Integer
    
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CSData
    buffer As String * 610
End Type

Public Type CSLProps
    ID As Long
    TRID As Long
    Qty As Long
    Discount As Long
    DiscountRate As Long
    VATRate As Long
    DateTime As Date
    PID As String * 40
    EAN As String * 13
    EANFormatted As String * 16
    code As String * 10
    CodeF As String * 20
    CodeFForExport As String * 15
    Title As String * 40
    Author As String * 20
    QtyF As String * 10
    Price As Long
    EXCHANGE_GUID As String * 40
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type CSLData
    buffer As String * 247
End Type

Public Type dCSLProps
    ID As Long
    Qty As Long
    DateTime As Date
    EAN As String * 16
    code As String * 15
    Title As String * 40
    Author As String * 20
    Price As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type dCSLData
    buffer As String * 180
End Type



Public Sub testCSLProps()
Dim f As CSProps
    MsgBox LenB(f) & "    " & LenB(f) / 2
End Sub
