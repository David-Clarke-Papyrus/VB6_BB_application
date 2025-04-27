VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_OSAPPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private udtProps As dAPPProps

Public Property Get TRID() As Long
    TRID = udtProps.TRID
End Property
Public Property Let TRID(Value As Long)
    udtProps.TRID = Value
End Property

Public Property Get TPNAME() As String
    TPNAME = Trim$(udtProps.TPNAME)
End Property
Public Property Let TPNAME(Value As String)
    udtProps.TPNAME = Value
End Property
Public Property Get TPAccNo() As String
    TPAccNo = Trim$(udtProps.TPAccNo)
End Property
Public Property Let TPAccNo(Value As String)
    udtProps.TPAccNo = Value
End Property

Public Property Get DOCCode() As String
    DOCCode = Trim(udtProps.DOCCode)
End Property
Public Property Let DOCCode(Value As String)
    udtProps.DOCCode = Value
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(udtProps.CaptureDate, "yyyy-mm-dd Hh Nn")
End Property
Public Property Get DOCDate() As String
    DOCDate = Trim(udtProps.DOCDate)
End Property
Public Property Get DocDateF() As String
    If udtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(udtProps.CaptureDate, "dd-mm-yyyy")
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
    udtProps.DOCDate = dteDate
End Function
Public Property Let DOCDate(Value As String)
    udtProps.DOCDate = Value
End Property
Public Property Let CaptureDate(Value As String)
    udtProps.CaptureDate = Value
End Property


Public Property Get APPQty() As Integer
    APPQty = Trim(udtProps.APPQty)
End Property
Public Property Let APPQty(Value As Integer)
    udtProps.APPQty = Value
End Property

Public Property Get APPReturned() As Integer
    APPReturned = Trim(udtProps.APPReturned)
End Property
Public Property Let APPReturned(Value As Integer)
    udtProps.APPReturned = Value
End Property

Public Property Let Status(Value As Long)
    udtProps.Status = Value
End Property
Public Property Get Status() As Long
    Status = Trim$(udtProps.Status)
End Property


Private Sub SetState(buffer As String)
  Dim mudtData As dAPPData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
