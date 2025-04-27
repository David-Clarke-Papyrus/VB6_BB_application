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
Attribute VB_Exposed = True
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
    On Error GoTo errHandler
    ID = udtProps.ID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.ID"
    
End Property
Friend Property Let ID(Val As Long)
    On Error GoTo errHandler
    udtProps.ID = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.ID(Val)", Val
End Property
Public Property Get StartedAtFormatted() As String
    On Error GoTo errHandler
    StartedAtFormatted = Format(udtProps.StartedAt, "dd/mm/yyyy h:mm am/pm")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.StartedAtFormatted"
    
End Property
Public Property Let StartedAt(Val As Date)
    On Error GoTo errHandler
    If IsDate(Val) Then udtProps.StartedAt = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.StartedAt(Val)", Val, EA_NORERAISE
    
End Property
Public Property Get StartedAt() As Date
    On Error GoTo errHandler
    StartedAt = udtProps.StartedAt
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.StartedAt"
    
End Property
Public Property Get EndedatFormatted() As String
    On Error GoTo errHandler
    EndedatFormatted = Format(udtProps.Endedat, "h:mm am/pm")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.EndedatFormatted"
    
End Property
Friend Property Let Endedat(Val As Date)
    On Error GoTo errHandler
    If IsDate(Val) Then udtProps.Endedat = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.Endedat(Val)", Val
End Property
Public Property Get NominalDateFormatted() As String
    On Error GoTo errHandler
    NominalDateFormatted = Format(udtProps.NominalDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.NominalDateFormatted"
    
End Property
Public Property Let NominalDate(Val As Date)
    On Error GoTo errHandler
    If IsDate(Val) Then udtProps.NominalDate = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.NominalDate(Val)", Val, EA_NORERAISE
    
End Property
Public Property Get TypeName() As String
    On Error GoTo errHandler
    TypeName = Trim(udtProps.TypeName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.TypeName"
    
End Property
Public Property Let TypeName(Val As String)
    On Error GoTo errHandler
    udtProps.TypeName = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.TypeName(Val)", Val, EA_NORERAISE
    
End Property
Public Property Get OperatorName() As String
    On Error GoTo errHandler
    OperatorName = Trim(udtProps.OperatorName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.OperatorName"
    
End Property
Public Property Let OperatorName(Val As String)
    On Error GoTo errHandler
    udtProps.OperatorName = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.OperatorName(Val)", Val, EA_NORERAISE
    
End Property
Public Property Get ResultName() As String
    On Error GoTo errHandler
    Select Case udtProps.Result
    Case 1
        ResultName = "OK"
    Case 2
        ResultName = "Failed(rollback)"
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.ResultName"
    
End Property
Public Property Let Result(Val As Integer)
    On Error GoTo errHandler
    udtProps.Result = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_operation.Result(Val)", Val, EA_NORERAISE
    
End Property

