Attribute VB_Name = "UDTTransfer"
Public Type TFProps
    TRID As Long
    DestID As Long
    Status As Long
    DOCCode As String * 10
    InOut As String * 3
    DestinationName As String * 20
    DOCDate As Date
    CaptureDate As Date
    ExtExVAT As Long
    ExtLessDiscExVAT As Long
    DiscExVAT As Long
    TotalQtyItems As Long
    TotalCostExVAT As Long
    BatchTotal As Long
    BatchQtyTotal As Long
    StaffID As Long
    StaffName As String * 10
    Memo As String * 200
    SendersDocDate As Date
    SendersDocRef As String * 25
    Auto As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type TFData
    buffer As String * 330
End Type

Public Type TFLProps
    ID As Long
    TRID As Long
    Qty As Long
    PID As String * 40
    EAN As String * 18
    EANFormatted As String * 16
    code As String * 10
    CodeF As String * 20
    CodeFForExport As String * 18
    Title As String * 40
    Author As String * 20
    Section As String * 51
    Note As String * 100
    Binding As String * 10
    SeriesTitle As String * 50
    Price As Long
    Cost As Long
    Discount As Double
    VATRate As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type TFLData
    buffer As String * 416
End Type

Public Type dTFLProps
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
Public Type dTFLData
    buffer As String * 106
End Type



Public Sub testTFLProps()
Dim f As TFProps
    MsgBox LenB(f) / 2
End Sub

