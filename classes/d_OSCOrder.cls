VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_OSCOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private udtProps As dCOProps
Public Property Get Ref() As String
    Ref = FNS(udtProps.Ref)
End Property
Public Property Let Ref(Value As String)
    udtProps.Ref = Value
End Property

Public Property Get DOCCode() As String
    DOCCode = FNS(udtProps.DOCCode)
End Property
Public Property Let DOCCode(Value As String)
    udtProps.DOCCode = Value
End Property
Public Property Get TPNAME() As String
    TPNAME = FNS(udtProps.TPNAME)
End Property
Public Property Let TPNAME(Value As String)
    udtProps.TPNAME = Value
End Property
Public Property Get EMail() As String
    EMail = FNS(udtProps.EMail)
End Property
Public Property Let EMail(Value As String)
    udtProps.EMail = Value
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(udtProps.CaptureDate, "yyyy-mm-dd Hh Nn")
End Property

Public Property Get DOCDate() As Date
    DOCDate = udtProps.DOCDate
End Property

Public Property Get DocDateF() As String
    If udtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(udtProps.DOCDate, "dd/mm/yyyy")
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
Public Property Get ETA() As Date
    ETA = udtProps.ETA
End Property

Public Property Get ETAF() As String
    If udtProps.ETA = CDate(0) Then
        ETAF = ""
    Else
        ETAF = Format(udtProps.ETA, "dd/mm/yyyy")
    End If
End Property
Public Property Let ETA(val As Date)
    udtProps.ETA = val
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

Public Property Get COLQty() As Integer
    COLQty = udtProps.Qty
End Property
Public Property Let COLQty(Value As Integer)
    udtProps.Qty = Value
End Property
Public Property Get COLID() As Long
    COLID = udtProps.COLID
End Property
Public Property Let COLID(Value As Long)
    udtProps.COLID = Value
End Property

Public Property Get COLCollected() As Integer
    COLCollected = udtProps.QtyDispatched
End Property
Public Property Let COLCollected(Value As Integer)
    udtProps.QtyDispatched = Value
End Property
'Public Property Get COLFulFilled() As String
'    COLFulFilled = FNS(udtProps.COLFulFilled)
'End Property

Public Property Get TRID() As Long
    TRID = Trim(udtProps.TRID)
End Property
'

Public Property Let DOCDate(Value As Date)
    udtProps.DOCDate = Value
End Property
Public Property Let CaptureDate(Value As Date)
    udtProps.CaptureDate = Value
End Property



'Public Property Let COLFulFilled(Value As String)
'    udtProps.COLFulFilled = Value
'End Property

Public Property Let TRID(Value As Long)
    udtProps.TRID = Value
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As dCOData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
