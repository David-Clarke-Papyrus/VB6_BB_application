VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_MM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private udtProps As dMMProps
    
Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property
Public Property Let PID(Value As String)
    udtProps.PID = Value
End Property

Public Property Get DOCCode() As String
    DOCCode = FNS(udtProps.DOCCode)
End Property
Public Property Let DOCCode(Value As String)
    udtProps.DOCCode = Value
End Property
Public Property Let Seq(val As Integer)
    udtProps.Seq = val
End Property
Public Property Get Seq() As Integer
    Seq = udtProps.Seq
End Property
Public Property Get Station() As String
Dim tmp As String
    Station = FNS(udtProps.Station)
End Property
Public Property Let Station(Value As String)
    udtProps.Station = Value
End Property
Public Property Get Qty() As Long
    Qty = FNN(udtProps.Qty)
End Property
Public Property Let Qty(Value As Long)
    udtProps.Qty = Value
End Property
Public Property Get DateForSort() As Date
    DateForSort = udtProps.CaptureDate
End Property

Public Property Get DOCDate() As Date
    DOCDate = udtProps.DOCDate
End Property
Public Property Get DocDateF() As String
    If udtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(udtProps.DOCDate, "dd-mm-yyyy HH:NN")
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
Public Property Let DOCDate(Value As Date)
    udtProps.DOCDate = Value
End Property
Public Property Let CaptureDate(Value As Date)
    udtProps.CaptureDate = Value
End Property

Public Property Get TRID() As Long
    TRID = udtProps.TRID
End Property
Public Property Let TRID(Value As Long)
    udtProps.TRID = Value
End Property
Public Property Get typ() As String
Dim tmp As String
    tmp = FNS(udtProps.Type)
    If tmp = "COUNTER SALES" Then
        tmp = "POS"
    ElseIf tmp = "INVOICE SALES" Then
        tmp = "INV"
    ElseIf tmp = "GOODS DELIVERY NOTE" Then
        If oPC.IncludeSupplierFeatures Then
            tmp = "GDN"
        Else
            tmp = "INV"
        End If
    ElseIf tmp = "CREDIT RETURNS" Then
        tmp = "CN"
    ElseIf tmp = "COUNTER RETURNS" Then
        tmp = "POS"
    End If
    typ = tmp
End Property
Public Property Let typ(Value As String)
    udtProps.Type = Value
End Property

Private Sub SetState(buffer As String)
  Dim udtData As dMMData
  
  udtData.buffer = buffer
  LSet udtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

