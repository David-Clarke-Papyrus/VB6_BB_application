VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mudtProps As dRProps
Public Property Get ReturnStatus() As Long
    ReturnStatus = mudtProps.ReturnStatus
End Property
 
Public Property Get Status() As Integer
    Status = FNN(mudtProps.Status)
End Property
Public Property Get StaffID() As Long
    StaffID = mudtProps.StaffID
End Property
Public Property Get StaffName() As String
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
End Property
Public Property Get StaffNameB() As String
Dim strT As String
    If mudtProps.StaffID < 1 Then
        StaffNameB = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    If strT > "" Then
        StaffNameB = "   (" & strT & ")"
    Else
        StaffNameB = ""
    End If
End Property

Public Property Let Status(Value As Integer)
    mudtProps.Status = Value
End Property
Public Property Get StatusF() As String
    If mudtProps.Status = 1 Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = 2 Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = 3 Then
        StatusF = "REQUESTED"
    ElseIf mudtProps.Status = 4 Then
        StatusF = "STOCK RETURNED"
    ElseIf mudtProps.Status = 7 Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = 6 Then
        StatusF = "CANCELLED"
    End If
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

Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Get ApprovalRef() As String
    ApprovalRef = FNS(mudtProps.ApprovalRef)
End Property
Public Property Let ApprovalRef(val As String)
    mudtProps.ApprovalRef = Trim(val)
End Property
Public Property Get RType() As String
    RType = mudtProps.RType
End Property
Public Property Get DOCDate() As Date
    DOCDate = FND(mudtProps.DOCDate)
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

'Public Property Get TPAccNo() As String
'    TPAccNo = FNS(mudtProps.TPAccNo)
'End Property
'Public Property Let TPAccNo(Value As String)
'    mudtProps.TPAccNo = Value
'End Property
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
  Dim udtData As dRData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

Public Property Let QtySystem(Value As Long)
    mudtProps.QtySystem = Value
End Property
Public Property Get QtySystem() As Long
    QtySystem = mudtProps.QtySystem
End Property
Public Property Let QtyRequested(Value As Long)
    mudtProps.QtyRequested = Value
End Property
Public Property Get QtyRequested() As Long
    QtyRequested = mudtProps.QtyRequested
End Property
Public Property Let QtyApproved(Value As Long)
    mudtProps.QtyApproved = Value
End Property
Public Property Get QtyApproved() As Long
    QtyApproved = mudtProps.QtyApproved
End Property
Public Property Let QtyReturned(Value As Long)
    mudtProps.QtyReturned = Value
End Property
Public Property Get QtyReturned() As Long
    QtyReturned = mudtProps.QtyReturned
End Property
Public Property Let QtyCounted(Value As Long)
    mudtProps.QtyCounted = Value
End Property
Public Property Get QtyCounted() As Long
    QtyCounted = mudtProps.QtyCounted
End Property

