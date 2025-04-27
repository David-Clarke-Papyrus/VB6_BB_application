VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_IG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As IGProps

Private Sub SetState(buffer As String)
  Dim mudtData As IGData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
    
Public Property Get IGID() As Long
    IGID = udtProps.IGID
End Property
    
Public Property Let IGID(Value As Long)
    udtProps.IGID = Value
End Property
Public Property Get TPID() As Long
    TPID = udtProps.TPID
End Property
    
Public Property Let TPID(Value As Long)
    udtProps.TPID = Value
End Property

