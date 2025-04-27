Attribute VB_Name = "UDTCashups"
Public Type CashupProps
    XID As String * 50
    BranchCode As String * 20
    Tillpoint As String * 20
    OpenSessionTime As Date
    CloseSessionTime As Date
    CapturedBy As String * 50
    CapturedDate As Date
    IssuedBy As String * 50
    IssuedDate As Date
    ExplainedBy As String * 50
    ExplainedDate As Date
    OpeningFloat As Double
    ClosingFloat As Double
    Cash As Double
    Cheques As Double
    CreditCards As Double
    DebitCards As Double
    DirectDeposits As Double
    VouchersRedeemed As Double
    FloatBreakdownAtEnd As String * 200
    Explanation As String * 1000
    DiscrepancyCash As Double
    DiscrepancyCheques As Double
    DiscrepancyCards As Double
    DiscrepancyVouchers As Double
    DiscrepancyDeposits As Double
    DiscrepancyFloat As Double
    DiscrepancyTotal As Double
    STATUS As String * 20
    StatusDate As Date
    StatusSignature As String * 30
    Wages As Double
    LeavePay As Double
    SickLeave As Double
    TotalSales As Double
    COGS As Double
    Retained As Double
    Returned As Double
    GiftVouchersSold As Double
    OtherVouchersSold As Double
    BankedAfterAdjustments As Double
End Type

Public Type CashupData
    buffer As String * 1614
End Type
Sub lenECashupProps()
Dim X As CashupProps
    MsgBox LenB(X) & "        " & LenB(X) / 2
End Sub

