VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Stktke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
    ID As Long
    Code As String * 15
    CaptureDate As Date
    NominalDate As Date
    OperatorName As String
    End Type

Private udtProps As DisplayProps

Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Friend Property Let ID(val As Long)
    udtProps.ID = val
End Property
Public Property Get Code() As String
    Code = udtProps.Code
End Property
Public Property Let Code(val As String)
    udtProps.Code = val
End Property
Public Property Get CaptureDate() As Date
    CaptureDate = udtProps.CaptureDate
End Property
Public Property Let CaptureDate(val As Date)
    udtProps.CaptureDate = val
End Property
Public Property Get NominalDate() As Date
    NominalDate = udtProps.NominalDate
End Property
Public Property Let NominalDate(val As Date)
    udtProps.NominalDate = val
End Property
Public Property Get OperatorName() As String
    OperatorName = Trim(udtProps.OperatorName)
End Property
Public Property Let OperatorName(val As String)
    udtProps.OperatorName = val
End Property

