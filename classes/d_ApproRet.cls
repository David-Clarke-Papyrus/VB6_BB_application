VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_APPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mudtProps As dAPPRProps

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


Private Sub SetState(buffer As String)
  Dim mudtData As dAPPRData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
Public Property Get DocDateF() As String
    If mudtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd-mm-yyyy")
    End If
End Property

Public Function SetDocdate(val As String) As Boolean
Dim dteDate As Date
    SetDocdate = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDocdate = False
        Exit Function
    End If
    mudtProps.DOCDate = dteDate
    
End Function
    
    
Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
    
Public Property Let DOCCode(Value As String)
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
    
Public Property Get TPID() As Long
    TPID = mudtProps.TPID
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.DOCDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get IssDateF() As String
    If mudtProps.CaptureDate = CDate(0) Then
        IssDateF = ""
    Else
        IssDateF = Format(mudtProps.CaptureDate, "dd-mm-yyyy Hh:Nn:Ss")
    End If
End Property
Public Property Get CustomerDisplay() As String
    CustomerDisplay = FNS(mudtProps.CustomerDisplay)
End Property

Public Property Let TPID(Value As Long)
    mudtProps.TPID = Value
End Property

Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property

Public Property Let TRID(Value As Long)
    mudtProps.TRID = Value
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

