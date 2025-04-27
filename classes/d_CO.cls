VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As dCOProps

Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Let TRID(val As Long)
    mudtProps.TRID = val
End Property
Public Property Get COLID() As Long
    COLID = mudtProps.COLID
End Property
Public Property Let COLID(val As Long)
    mudtProps.COLID = val
End Property
Public Property Get PID() As String
    PID = Trim(mudtProps.PID)
End Property
Public Property Let PID(val As String)
    mudtProps.PID = val
End Property
Public Property Get Qty() As Long
    Qty = mudtProps.Qty
End Property
Public Property Let Qty(val As Long)
    mudtProps.Qty = val
End Property
Public Property Get OrderType() As Long
    OrderType = mudtProps.OrderType
End Property
Public Property Let OrderType(val As Long)
    mudtProps.OrderType = val
End Property
Public Property Get Used() As Boolean
    Used = mudtProps.Used
End Property
Public Property Let Used(val As Boolean)
    mudtProps.Used = val
End Property

Public Property Get Price() As Long
    Price = mudtProps.Price
End Property
Public Property Let Price(val As Long)
    mudtProps.Price = val
End Property
Public Property Get Deposit() As Long
    Deposit = mudtProps.Deposit
End Property
Public Property Let Deposit(val As Long)
    mudtProps.Deposit = val
End Property
Public Property Get DepositF() As String
    DepositF = Format(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Discount() As Double
    Discount = mudtProps.Discount
End Property
Public Property Let Discount(val As Double)
    mudtProps.Discount = val
End Property
Public Property Get PriceF() As String
    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get Ref() As String
    Ref = FNS(mudtProps.Ref)
End Property
Public Property Let Ref(val As String)
    mudtProps.Ref = val
End Property


Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Let DOCCode(val As String)
    mudtProps.DOCCode = val
End Property

Public Property Get TPNAME() As String
    TPNAME = FNS(mudtProps.TPNAME)
End Property
Public Property Let TPNAME(val As String)
    mudtProps.TPNAME = val
End Property

Public Property Get StatusF() As String
    If mudtProps.OrderType = 1 Then
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
    End If
End Property
'Public Property Let Status(Val As Integer)
'    mudtProps.Status = Val
'End Property


Public Property Get TPAccNo() As String
    TPAccNo = mudtProps.TPAccNo
End Property
Public Property Let TPAccNo(val As String)
    mudtProps.TPAccNo = val
End Property

Public Property Get DOCDate() As Date
    DOCDate = mudtProps.DOCDate
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.CaptureDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get DocDateF() As String
    If mudtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd-mm-yyyy")
    End If
End Property
Public Property Get CustomerDisplay() As String
    CustomerDisplay = FNS(mudtProps.CustomerDisplay)
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As dCOData
  
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

