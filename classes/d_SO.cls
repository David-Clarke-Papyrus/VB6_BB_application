VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_PO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As dPOProps
Public Property Get Totals() As String
    Totals = Format(mudtProps.POValue, oPC.Configuration.Currencies.FindCurrencyByID(CURRID).FormatString) & " (" & CStr(mudtProps.POQty) & ")"
End Property
Public Property Get CURRID() As Long
    CURRID = mudtProps.CURRID
End Property
Public Property Get TPAccNo() As String
    TPAccNo = mudtProps.TPAccNo
End Property
Public Property Let TPAccNo(val As String)
    mudtProps.TPAccNo = val
End Property
Public Property Let Log(val As String)
    mudtProps.Log = val
End Property
Public Property Get Log() As String
    Log = FNS(mudtProps.Log)
End Property

Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Let TRID(val As Long)
    mudtProps.TRID = val
End Property
Public Property Get PIID() As Long
    PIID = mudtProps.PIID
End Property
Public Property Get PriceF() As String
    PriceF = PBKSCurrFormat(mudtProps.Price)
End Property
Public Property Get Price() As Long
    Price = mudtProps.Price
End Property
Public Property Get QtySS() As Long
    QtySS = mudtProps.QtySS
End Property
Public Property Get QtyFirm() As Long
    QtyFirm = mudtProps.QtyFirm
End Property
Public Property Get Title() As String
    Title = mudtProps.Title
End Property
Public Property Get DiscountF() As String
    DiscountF = Format(Round(mudtProps.Discount, 2), "##.##\%")
End Property
Public Property Get Discount() As Double
    Discount = mudtProps.Discount
End Property
Public Property Get POLineID() As Long
    POLineID = mudtProps.POLID
End Property
Public Property Get VATRate() As Double
    VATRate = mudtProps.VATRate
End Property
Public Property Let VATRate(val As Double)
    mudtProps.VATRate = val
End Property
Public Property Get ParentID() As Long
    ParentID = mudtProps.Parent
End Property


Public Property Get Ref() As String
    Ref = FNS(mudtProps.DOCCode)
End Property
Public Property Let Ref(val As String)
    mudtProps.DOCCode = val
End Property

Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Let DOCCode(val As String)
    mudtProps.DOCCode = val
End Property
Public Property Get DispatchMode() As String
    DispatchMode = FNS(mudtProps.DispatchMode)
End Property
Public Property Let DispatchMode(val As String)
    mudtProps.DispatchMode = val
End Property

Public Property Get TPNAME() As String
    TPNAME = FNS(mudtProps.TPNAME)
End Property
Public Property Get OrderType() As String
    OrderType = IIf(FNS(mudtProps.OrderType) = "NS", "s", "")
End Property
Public Property Let TPNAME(val As String)
    mudtProps.TPNAME = val
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



Public Property Get DOCDate() As Date
    DOCDate = mudtProps.DOCDate
End Property
Public Property Get DocDateF() As String
    If mudtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd-mm-yyyy")
    End If
End Property
Public Property Get DateForSort() As String
    DateForSort = FND(mudtProps.CaptureDate) ', "yyyy-mm-dd: Hh:Nn")
End Property
Public Property Get IssDateF() As String
    If mudtProps.CaptureDate = CDate(0) Then
        IssDateF = ""
    Else
        IssDateF = Format(mudtProps.CaptureDate, "dd-mm-yyyy Hh:Nn:Ss")
    End If
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As dPOData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get StaffName() As String
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(mudtProps.StaffID).Shortname
End Property
Public Property Get StaffNameB() As String
Dim strB As String
    If mudtProps.StaffID < 1 Then
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

