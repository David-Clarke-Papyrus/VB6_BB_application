VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_DEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mudtProps As dDELProps
Public Property Get Totals() As String
    Totals = Format(mudtProps.DeliveredValue / oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.CURRID).Divisor, oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.CURRID).FormatString) & " (" & CStr(mudtProps.DeliveredQty) & ")"
End Property

'Public Property Get Invoice() As String
'    Invoice = FNS(mudtProps.Invoice)
'End Property
 
Public Property Get Status() As String
    Status = mudtProps.Status
End Property

Public Property Let Status(Value As String)
    mudtProps.Status = Value
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
Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
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
Public Property Get IssDateF() As String
    If mudtProps.CaptureDate = CDate(0) Then
        IssDateF = ""
    Else
        IssDateF = Format(mudtProps.CaptureDate, "dd-mm-yyyy Hh:Nn:Ss")
    End If
End Property
 
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

Public Property Get Ref() As String
    Ref = FNS(mudtProps.DOCCode)
End Property

Public Property Let Ref(Value As String)
    mudtProps.DOCCode = Value
End Property

Public Property Get DOCDate() As Date
    DOCDate = mudtProps.DOCDate
End Property

Public Property Let DOCDate(Value As Date)
    mudtProps.DOCDate = Value
End Property

Public Property Get TPNAME() As String
    TPNAME = FNS(mudtProps.TPNAME)
End Property
Public Property Let TPNAME(Value As String)
    mudtProps.TPNAME = Value
End Property

Public Property Get TPAccNo() As String
    TPAccNo = FNS(mudtProps.TPAccNo)
End Property
Public Property Let TPAccNo(Value As String)
    mudtProps.TPAccNo = Value
End Property
Public Property Let TRID(Value As Long)
    mudtProps.TRID = Value
End Property
Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property

Public Property Get TPID() As Long
    TPID = mudtProps.TPID
End Property

Public Property Let TPID(Value As Long)
    mudtProps.TPID = Value
End Property

Private Sub SetState(buffer As String)
  Dim udtData As dDELData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

Public Property Get InvoiceDate() As Date
    InvoiceDate = mudtProps.DOCDate
End Property
Public Property Get InvoiceDateF() As String
    InvoiceDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
End Property

Public Property Let InvoiceDate(Value As Date)
    mudtProps.InvoiceDate = Value
End Property

Public Property Get InvoiceRef() As String
    InvoiceRef = FNS(mudtProps.InvoiceRef)
End Property

Public Property Let InvoiceRef(Value As String)
    mudtProps.InvoiceRef = Value
End Property

Public Property Get InvoiceValueF() As String
    InvoiceValueF = Format(mudtProps.InvoiceValue, "###,###.00")
End Property

Public Property Get InvoiceQty() As Long
    InvoiceQty = mudtProps.InvoiceQty
End Property
Public Property Get InvoiceQtyF() As String
    InvoiceQtyF = Format(mudtProps.InvoiceQty, "###,###")
End Property

Public Property Let InvoiceQty(Value As Long)
    mudtProps.InvoiceQty = Value
End Property
            
Public Property Get InvoiceShort() As Long
    InvoiceShort = mudtProps.InvoiceShort
End Property
Public Property Get InvoiceShortF() As String
    InvoiceShortF = Format(InvoiceShort, "###,###")
End Property

Public Property Let InvoiceShort(Value As Long)
    mudtProps.InvoiceShort = Value
End Property

Public Property Get DeliveredQty() As Long
    DeliveredQty = mudtProps.DeliveredQty
End Property
Public Property Get DeliveredQtyF() As String
    DeliveredQtyF = Format(DeliveredQty, "###,###")
End Property
Public Property Get DeliveredValue() As Long
    DeliveredValue = mudtProps.DeliveredValue
End Property
Public Property Get DeliveredValueF() As String
    DeliveredValueF = Format(DeliveredValue, "###,###.00")
End Property

