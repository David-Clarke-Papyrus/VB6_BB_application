VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_COSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mudtProps As dCOSRProps
Public Property Let Log(val As String)
    mudtProps.Log = val
End Property
Public Property Get Log() As String
    Log = FNS(mudtProps.Log)
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

'Public Property Get PriceF() As String
'    PriceF = PBKSCurrFormat(mudtProps.Price)
'End Property
'Public Property Get Price() As Long
'    Price = mudtProps.Price
'End Property
'Public Property Get Qty() As Long
'    Qty = mudtProps.Qty
'End Property
'Public Property Get Title() As String
'    Title = mudtProps.Title
'End Property
Public Property Get DOCCode() As String
    DOCCode = mudtProps.DOCCode
End Property

Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Let TRID(val As Long)
    mudtProps.TRID = val
End Property

Public Property Get Ref() As String
    Ref = FNS(mudtProps.DOCCode)
End Property
Public Property Let Ref(val As String)
    mudtProps.DOCCode = val
End Property
Public Property Get CustomerDisplay() As String
    CustomerDisplay = FNS(mudtProps.CustomerDisplay)
End Property

Public Property Get TPNAME() As String
    TPNAME = FNS(mudtProps.TPNAME)
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
    ElseIf mudtProps.Status = 6 Then
        StatusF = "CANCELLED"
    End If
End Property
Public Property Get TPAccNo() As String
    TPAccNo = mudtProps.TPAccNo
End Property
Public Property Get SMSHortname() As String
    SMSHortname = mudtProps.SM
End Property

Public Property Let TPAccNo(val As String)
    mudtProps.TPAccNo = val
End Property

Public Property Get InvoiceDate() As Date
    InvoiceDate = mudtProps.DOCDate
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.CaptureDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get TDateF() As String
    If mudtProps.DOCDate = CDate(0) Then
        TDateF = ""
    Else
        TDateF = Format(mudtProps.CaptureDate, "dd-mm-yyyy Hh:Nn")
    End If
End Property
Public Property Get tDate() As Date
    tDate = mudtProps.DOCDate
End Property
Private Sub SetState(buffer As String)
  Dim mudtData As dCOSRData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get IssDateF() As String
    If mudtProps.CaptureDate = CDate(0) Then
        IssDateF = ""
    Else
        IssDateF = Format(mudtProps.CaptureDate, "dd-mm-yyyy Hh:Nn:Ss")
    End If
End Property


