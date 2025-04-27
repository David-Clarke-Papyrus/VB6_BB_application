VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_ServiceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As ProductDisplayProps

Private Sub SetState(buffer As String)
  Dim mudtData As ProductDisplayData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
    
Public Property Get PID() As String
    PID = FNS(udtProps.ID)
End Property
    
Public Property Let PID(Valu As String)
    udtProps.ID = Valu
End Property
Public Property Get Description() As String
    Description = FNS(udtProps.Description)
End Property
    
Public Property Let Description(Valu As String)
    udtProps.Description = Valu
End Property
Public Property Get code() As String
    code = FNS(udtProps.code)
End Property
    
Public Property Let code(Valu As String)
    udtProps.code = Valu
End Property


