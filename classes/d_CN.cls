VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As dCNProps

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

Public Property Get TPNAME() As String
    TPNAME = Trim(mudtProps.TPNAME)
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


Public Property Get TPAccNo() As String
    TPAccNo = mudtProps.TPAccNo
End Property
Public Property Let TPAccNo(val As String)
    mudtProps.TPAccNo = val
End Property

Public Property Get DOCDate() As Date
    DOCDate = mudtProps.CaptureDate
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
Public Property Get CustomerDisplay() As String
    CustomerDisplay = FNS(mudtProps.CustomerDisplay)
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As dInvoiceData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

