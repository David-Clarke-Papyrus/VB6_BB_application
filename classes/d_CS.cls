VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As CSProps

Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Let TRID(val As Long)
    mudtProps.TRID = val
End Property
Public Property Get CSGUID() As String
    CSGUID = FNS(mudtProps.CS_GUID_ID)
End Property
Public Property Get Reportable() As Integer
    Reportable = FNINT(mudtProps.Reportable)
End Property

Public Property Get StartDate() As Date
    StartDate = mudtProps.DateStarted
End Property
Public Property Get StartDateSort() As String
    StartDateSort = Format(mudtProps.DateStarted, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get StartDateF() As String
    If mudtProps.DateStarted = CDate(0) Then
        StartDateF = ""
    Else
        StartDateF = Format(mudtProps.DateStarted, "dd/mmm/yy  hh:nn")
    End If
End Property
Public Property Get EndDate() As Date
    EndDate = mudtProps.DateIssued
End Property
Public Property Get EndDateSort() As String
    EndDateSort = Format(mudtProps.DateIssued, "yyyy/mm/dd Hh:Nn:Ss")
End Property
Public Property Get EndDateF() As String
    If mudtProps.DateIssued = CDate(0) Then
        EndDateF = ""
    Else
        EndDateF = Format(mudtProps.DateIssued, "dd/mmm/yy  hh:nn")
    End If
End Property


Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Let DOCCode(val As String)
    mudtProps.DOCCode = val
End Property
Public Property Get StaffName() As String
    StaffName = FNS(mudtProps.StaffName)
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

Public Property Get DOCDate() As Date
    DOCDate = mudtProps.DateIssued
End Property
Public Property Get DocDateF() As String
    If mudtProps.DateIssued = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DateIssued, "dd-mm-yyyy")
    End If
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.DateStarted, "yyyy-mm-dd: Hh:Nn")
End Property
Private Sub SetState(buffer As String)
  Dim mudtData As CSData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

