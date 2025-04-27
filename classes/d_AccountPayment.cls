VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_AccountPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As dPaymentProps

Public Property Get DepositID() As Long
    DepositID = mudtProps.DepositID
End Property
Public Property Let DepositID(val As Long)
    mudtProps.DepositID = val
End Property
Public Property Get DepositType() As String
    DepositType = mudtProps.DepositType
End Property
Public Property Let DepositType(val As String)
    mudtProps.DepositType = val
End Property

Public Property Get DepositAmount() As Double
    DepositAmount = mudtProps.DepositAmount
End Property
Public Property Let DepositAmount(val As Double)
    mudtProps.DepositAmount = val
End Property
Public Property Get DepositAmountF() As String
    DepositAmountF = Format(DepositAmount, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get DepositSettlementDiscount() As Double
    DepositSettlementDiscount = mudtProps.DepositSettlementDiscount
End Property
Public Property Let DepositSettlementDiscount(val As Double)
    mudtProps.DepositSettlementDiscount = val
End Property
Public Property Get DepositSettlementDiscountF() As String
    DepositSettlementDiscountF = Format(DepositSettlementDiscount, oPC.Configuration.DefaultCurrency.FormatString)
End Property



Public Property Get DepositCode() As String
    DepositCode = FNS(mudtProps.DepositCode)
End Property
Public Property Let DepositCode(val As String)
    mudtProps.DepositCode = val
End Property

Public Property Get DepositorName() As String
    DepositorName = FNS(mudtProps.DepositorName)
End Property
Public Property Let DepositorName(val As String)
    mudtProps.DepositorName = val
End Property
Public Property Get DepositorID() As Long
    DepositorID = FNS(mudtProps.DepositorID)
End Property
Public Property Let DepositorID(val As Long)
    mudtProps.DepositorID = val
End Property

Public Property Get StatusF() As String
        If mudtProps.Status = 1 Then
            StatusF = "VOID"
        ElseIf mudtProps.Status = 2 Then
            StatusF = "IN PROCESS"
        ElseIf mudtProps.Status = 3 Then
            StatusF = "ISSUED"
        ElseIf mudtProps.Status = 4 Then
            StatusF = "COMPLETE"
        ElseIf mudtProps.Status = 7 Then
            StatusF = "PROFORMA"
        ElseIf mudtProps.Status = 6 Then
            StatusF = "CANCELLED"
        End If
End Property
'Public Property Let Status(Val As Integer)
'    mudtProps.Status = Val
'End Property


Public Property Get DepositorAcNo() As String
    DepositorAcNo = mudtProps.DepositorAcNo
End Property
Public Property Let DepositorAcNo(val As String)
    mudtProps.DepositorAcNo = val
End Property

Public Property Get DepositDate() As Date
    DepositDate = mudtProps.DepositDate
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.DepositDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get DepositDateF() As String
    If mudtProps.DepositDate = CDate(0) Then
        DepositDateF = ""
    Else
        DepositDateF = Format(mudtProps.DepositDate, "dd-mm-yyyy")
    End If
End Property
Private Sub SetState(buffer As String)
  Dim mudtData As dPaymentData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

Public Property Get StaffName() As String
    If mudtProps.StaffID = 0 Then
        StaffName = ""
    Else
        StaffName = oPC.Configuration.Staff.FindStaffByID(mudtProps.StaffID).Shortname
    End If
End Property
Public Property Get StaffNameB() As String
Dim strB As String
    If mudtProps.StaffID = 0 Then
        StaffNameB = ""
        Exit Property
    End If
    strB = oPC.Configuration.Staff.FindStaffByID(mudtProps.StaffID).Shortname
    If strB > "" Then
        StaffNameB = "  (" & strB & ")"
    Else
        StaffNameB = ""
    End If
End Property


