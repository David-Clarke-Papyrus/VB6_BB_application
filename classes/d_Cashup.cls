VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Cashup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private udtProps As CashupProps

Private Sub SetState(buffer As String)
  Dim mudtData As CashupData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
    
Public Property Get XID() As String
    XID = udtProps.XID
End Property
    
Public Property Get BranchCode() As String
    BranchCode = FNS(udtProps.BranchCode)
End Property

Public Property Get Tillpoint() As String
    Tillpoint = FNS(udtProps.Tillpoint)
End Property

Public Property Get OpenSessionTimeF() As String
    If udtProps.OpenSessionTime <> 0 Then
        OpenSessionTimeF = Format(udtProps.OpenSessionTime, "dd-mm-yy Hh:Nn")
    Else
        OpenSessionTimeF = ""
    End If
End Property

Public Property Get CloseSessionTimeF() As String
    If udtProps.CloseSessionTime <> 0 Then
        CloseSessionTimeF = Format(udtProps.CloseSessionTime, "dd-mm-yy Hh:Nn")
    Else
        CloseSessionTimeF = ""
    End If
End Property

Public Property Get TotalSales() As Double
    TotalSales = udtProps.TotalSales
End Property
Public Property Get TotalSalesF() As String
    TotalSalesF = Format(TotalSales, "###,##0.00")
End Property


Public Property Get Cash() As Double
    Cash = udtProps.Cash
End Property
Public Property Get CashF() As String
    CashF = Format(Cash, "###,##0.00")
End Property

Public Property Get Cheques() As Double
    Cheques = udtProps.Cheques
End Property
Public Property Get ChequesF() As String
    ChequesF = Format(Cheques, "###,##0.00")
End Property

Public Property Get CreditCards() As Double
    CreditCards = udtProps.CreditCards
End Property
Public Property Get CreditCardsF() As String
    CreditCardsF = Format(CreditCards, "###,##0.00")
End Property

Public Property Get DebitCards() As Double
    DebitCards = udtProps.DebitCards
End Property
Public Property Get DebitCardsF() As String
    DebitCardsF = Format(DebitCards, "###,##0.00")
End Property

Public Property Get DirectDeposits() As Double
    DirectDeposits = udtProps.DirectDeposits
End Property
Public Property Get DirectDepositsF() As String
    DirectDepositsF = Format(DirectDeposits, "###,##0.00")
End Property

Public Property Get VouchersRedeemed() As Double
    VouchersRedeemed = udtProps.VouchersRedeemed
End Property
Public Property Get VouchersRedeemedF() As String
    VouchersRedeemedF = Format(VouchersRedeemed, "###,##0.00")
End Property

Public Property Get BankedAfterAdjustments() As Double
    BankedAfterAdjustments = udtProps.BankedAfterAdjustments
End Property
Public Property Get BankedAfterAdjustmentsF() As String
    BankedAfterAdjustmentsF = Format(BankedAfterAdjustments, "###,##0.00")
End Property

Public Property Get OtherVouchersSold() As Double
    OtherVouchersSold = udtProps.OtherVouchersSold
End Property
Public Property Get OtherVouchersSoldF() As String
    OtherVouchersSoldF = Format(OtherVouchersSold, "###,##0.00")
End Property


Public Property Get GiftVouchersSold() As Double
    GiftVouchersSold = udtProps.GiftVouchersSold
End Property
Public Property Get GiftVouchersSoldF() As String
    GiftVouchersSoldF = Format(GiftVouchersSold, "###,##0.00")
End Property

Public Property Get Returned() As Double
    Returned = udtProps.Returned
End Property
Public Property Get ReturnedF() As String
    ReturnedF = Format(Returned, "###,##0.00")
End Property
Public Property Get Retained() As Double
    Retained = udtProps.Retained
End Property
Public Property Get RetainedF() As String
    RetainedF = Format(Retained, "###,##0.00")
End Property
Public Property Get COGS() As Double
    COGS = udtProps.COGS
End Property
Public Property Get COGSF() As String
    COGSF = Format(COGS, "###,##0.00")
End Property
Public Property Get SickLeave() As Double
    SickLeave = udtProps.SickLeave
End Property
Public Property Get SickLeaveF() As String
    SickLeaveF = Format(SickLeave, "###,##0.00")
End Property
Public Property Get LeavePay() As Double
    LeavePay = udtProps.LeavePay
End Property
Public Property Get LeavePayF() As String
    LeavePayF = Format(LeavePay, "###,##0.00")
End Property
Public Property Get Wages() As Double
    Wages = udtProps.Wages
End Property
Public Property Get WagesF() As String
    WagesF = Format(Wages, "###,##0.00")
End Property




Public Property Get OpeningFloat() As Double
    OpeningFloat = FNDBL(udtProps.OpeningFloat)
End Property

Public Property Get OpeningFloatF() As String
    OpeningFloatF = Format(OpeningFloat, "###,##0.00")
End Property

Public Property Get ClosingFloat() As Double
    ClosingFloat = FNDBL(udtProps.ClosingFloat)
End Property
Public Property Get ClosingFloatF() As String
    ClosingFloatF = Format(ClosingFloat, "###,##0.00")
End Property

Public Property Get DiscrepancyALL() As Double
    DiscrepancyALL = FNDBL(udtProps.DiscrepancyTotal)
End Property
Public Property Get DiscrepancyAllF() As String
    DiscrepancyAllF = Format(DiscrepancyALL, "###,##0.00")
End Property
Public Property Get STATUS() As String
    STATUS = FNS(udtProps.STATUS)
End Property
Public Property Get StatusDateF() As String
    If udtProps.StatusDate <> 0 Then
        StatusDateF = Format(udtProps.StatusDate, "dd-mm-yy Hh:Nn")
    Else
        StatusDateF = ""
    End If
End Property
Public Property Get StatusDate() As Date
    StatusDate = udtProps.StatusDate
End Property
Public Property Get StatusSignature() As String
    StatusSignature = FNS(udtProps.StatusSignature)
End Property

Private Function CalcFloatValue(s As String) As Double
Dim ar() As String
    If s = "" Then
        CalcFloatValue = 0
        Exit Function
    End If
    ar() = Split(s, ",")
    CalcFloatValue = CDbl((FNN(ar(0)) * 20000) + _
                    (FNN(ar(1)) * 10000) + _
                    (FNN(ar(2)) * 5000) + _
                    (FNN(ar(3)) * 2000) + _
                    (FNN(ar(4)) * 1000)) / 100 + _
                     CDbl((FNN(ar(5)) * 500) + _
                    (FNN(ar(6)) * 200) + _
                    (FNN(ar(7)) * 100) + _
                    (FNN(ar(8)) * 50) + _
                    (FNN(ar(9)) * 20) + _
                    (FNN(ar(10)) * 10) + _
                    (FNN(ar(11)) * 5)) / 100


End Function
Public Property Get Explanation() As String
    Explanation = FNS(udtProps.Explanation)
End Property
Public Property Get CapturedBy() As String
    CapturedBy = FNS(udtProps.CapturedBy)
End Property
Public Property Get IssuedBy() As String
    IssuedBy = FNS(udtProps.IssuedBy)
End Property
Public Property Get ExplainedBy() As String
    ExplainedBy = FNS(udtProps.ExplainedBy)
End Property
Public Property Get CapturedDateF() As String
    If udtProps.CapturedDate <> 0 Then
        CapturedDateF = Format(udtProps.CapturedDate, "dd-mm-yy Hh:Nn")
    Else
        CapturedDateF = ""
    End If
End Property
Public Property Get IssuedDateF() As String
    If udtProps.IssuedDate <> 0 Then
        IssuedDateF = Format(udtProps.IssuedDate, "dd-mm-yy Hh:Nn")
    Else
        IssuedDateF = ""
    End If
End Property
Public Property Get ExplainedDateF() As String
    If udtProps.ExplainedDate <> 0 Then
        ExplainedDateF = Format(udtProps.ExplainedDate, "dd-mm-yy Hh:Nn")
    Else
        ExplainedDateF = ""
    End If
End Property


