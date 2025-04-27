VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_operator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
    ID As Long
    Fullname As String * 15
    End Type

Private udtProps As DisplayProps

Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Friend Property Let ID(val As Long)
    udtProps.ID = val
End Property
Public Property Get Fullname() As String
    Fullname = Trim(udtProps.Fullname)
End Property
Public Property Let Fullname(val As String)
    udtProps.Fullname = val
End Property


