VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_BICCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As BICProps

Public Property Get code() As String
    code = Trim$(mudtProps.code)
End Property
Public Property Get Description() As String
    Description = Trim$(mudtProps.Description)
End Property
Public Property Let code(Value As String)
    mudtProps.code = Value
End Property
Public Property Let Description(Value As String)
    mudtProps.Description = Value
End Property
Public Property Get ID() As Long
    ID = Trim$(mudtProps.ID)
End Property
Public Property Let ID(Value As Long)
    mudtProps.ID = Value
End Property
Public Property Get Level() As Integer
    Level = Trim$(mudtProps.Level)
End Property
Public Property Let Level(Value As Integer)
    mudtProps.Level = Value
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As BICData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub




