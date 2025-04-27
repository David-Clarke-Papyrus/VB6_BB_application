VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_OSSOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private udtProps As dPOProps
    
Public Property Get OrderType() As String
    OrderType = FNS(udtProps.OrderType)
End Property

Public Property Let OrderType(Value As String)
    udtProps.OrderType = Value
End Property
Public Property Get Firm() As Long
    Firm = FNN(udtProps.QtyFirm)
End Property
Public Property Let Firm(Value As Long)
    udtProps.QtyFirm = Value
End Property
    
Public Property Get SS() As Long
    SS = FNN(udtProps.QtySS)
End Property
Public Property Let SS(Value As Long)
    udtProps.QtySS = Value
End Property
Public Property Get QtyReceived() As Long
    QtyReceived = FNN(udtProps.QtyReceived)
End Property
Public Property Let QtyReceived(Value As Long)
    udtProps.QtyReceived = Value
End Property
   
    
Public Property Get Status() As Integer
    Status = FNN(udtProps.Status)
End Property

Public Property Let Status(Value As Integer)
    udtProps.Status = Value
End Property

Public Property Get Name() As String
    Name = FNS(udtProps.TPNAME)
End Property

Public Property Let Name(Value As String)
    udtProps.TPNAME = Value
End Property

Public Property Get code() As String
    code = Trim(udtProps.DOCCode)
    
End Property

Public Property Get DOCDate() As Date
    DOCDate = Format(udtProps.DOCDate, "dd/mm/yyyy")
End Property

Public Property Get DocDateF() As String
    If udtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(udtProps.CaptureDate, "dd-mm-yyyy")
    End If
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(udtProps.CaptureDate, "yyyy-mm-dd Hh:Nn:Ss")
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

Public Property Get TRID() As Long
    TRID = udtProps.TRID
End Property

Public Property Let DOCCode(Value As String)
    udtProps.DOCCode = Value
End Property
Public Property Get DOCCode() As String
    DOCCode = FNS(udtProps.DOCCode)
End Property

Public Property Let DOCDate(Value As Date)
    udtProps.DOCDate = Value
End Property
Public Property Let CaptureDate(Value As Date)
    udtProps.CaptureDate = Value
End Property

Public Property Let TRID(Value As Long)
    udtProps.TRID = Value
End Property

Private Sub SetState(buffer As String)
  Dim udtData As dPOData
  
  udtData.buffer = buffer
  LSet udtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get ETA() As Date
    ETA = udtProps.ETA
End Property
Public Property Let ETA(val As Date)
    udtProps.ETA = val
End Property

Public Property Get ETAF() As String
    If udtProps.ETA = CDate(0) Then
        ETAF = ""
    Else
        ETAF = Format(udtProps.ETA, "dd/mm/yyyy")
    End If
End Property

Public Function SetETA(val As String) As Boolean
Dim dteDate As Date
    SetETA = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetETA = False
        Exit Function
    End If
    udtProps.ETA = dteDate
   
End Function

