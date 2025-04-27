VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_operation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DisplayProps
    ID As Long
    TypeName As String * 35
    StartedAt As Date
    Endedat As Date
    NominalDate As Date
    Result As Integer
    OperatorName As String
    End Type

Private udtProps As DisplayProps

Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Friend Property Let ID(val As Long)
    udtProps.ID = val
End Property
Public Property Get StartedAtFormatted() As String
    StartedAtFormatted = Format(udtProps.StartedAt, "dd/mm/yyyy h:mm am/pm")
End Property
Public Property Let StartedAt(val As Date)
    If IsDate(val) Then udtProps.StartedAt = val
End Property
Public Property Get StartedAt() As Date
    StartedAt = udtProps.StartedAt
End Property
Public Property Get EndedatFormatted() As String
    EndedatFormatted = Format(udtProps.Endedat, "h:mm am/pm")
End Property
Friend Property Let Endedat(val As Date)
    If IsDate(val) Then udtProps.Endedat = val
End Property
Public Property Get NominalDateFormatted() As String
    NominalDateFormatted = Format(udtProps.NominalDate, "dd/mm/yyyy")
End Property
Public Property Let NominalDate(val As Date)
    If IsDate(val) Then udtProps.NominalDate = val
End Property
Public Property Get TypeName() As String
    TypeName = Trim(udtProps.TypeName)
End Property
Public Property Let TypeName(val As String)
    udtProps.TypeName = val
End Property
Public Property Get OperatorName() As String
    OperatorName = Trim(udtProps.OperatorName)
End Property
Public Property Let OperatorName(val As String)
    udtProps.OperatorName = val
End Property
Public Property Get ResultName() As String
    Select Case udtProps.Result
    Case 1
        ResultName = "OK"
    Case 2
        ResultName = "Failed(rollback)"
    End Select
End Property
Public Property Let Result(val As Integer)
    udtProps.Result = val
End Property

