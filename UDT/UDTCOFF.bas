Attribute VB_Name = "UDTCOFF"
Public Type COFFProps
    COFFID As Long
    COFFCOLID As Long
    COLID As Long
    COFFCOCode As String * 12
    COFFINVCode As String * 12
    COFFILID As Long
    COFFCOLLID As Long
    COFFILQTY As Long
    COFFCOLQTY As Long
    COFFTitle As String * 25
    COFFCode As String * 15
    COFFQTY As Long
    COLQty As Long
    COLQtyDispatched As Long
    CODate As Date
    COCode As String * 10
    Status As Integer
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type COFFData
    buffer As String * 102
End Type
Public Type COLAllocationProps
    ID As Long
    PID As String * 40
    COLID As Long
    DELLID As Long
    DeliveredSoFar As Long
    DeliveredSoFarSS As Long
    DeliveredSoFarFirm As Long
    QtyOnHand As Long
    QtyJustReceived As Long
    AllocatedQty As Long
    AllocatedQtySS As Long
    QuotedPrice As Long
    OrderedQty As Long
    OrderedSSQTY As Long
    QtyReserved As Long
    QtyonOrder As Long
    QtyonCO As Long
    Status As Integer
    OrderDate As Date
    Ref As String * 50
    CustomerName As String * 70
    CustomerInitials As String * 10
    CustomerTitle As String * 10
    CustomerAcno As String * 15
    OrderCode As String * 50
    DepositValue As Double
    WSLock  As String * 50
    code As String * 20
    Title As String * 25
    Phone As String * 20
    Note As String * 200
    ActionYN As Boolean
    UsesSubstitutesYN As String * 1
    ProductOH As Long
    ProductRES As Long
    CustomerBlocked As Boolean
    CreateInvoice As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type COLAllocationData
    buffer As String * 630
End Type

Sub GetCOFFLen()
Dim x As COLAllocationProps
MsgBox LenB(x) & "   " & LenB(x) / 2
End Sub

