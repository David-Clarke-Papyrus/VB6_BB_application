VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Bookset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
  ProdSetID As Long
  Description As String * 50
  ShortDescription As String * 10
End Type

Private udtProps As DisplayProps

Public Property Get ProdSetID() As Long
  ProdSetID = udtProps.ProdSetID
End Property

Public Property Get Description() As String
  Description = udtProps.Description
End Property

Public Property Get ShortDescription() As String
  ShortDescription = udtProps.ShortDescription
End Property


Friend Property Let ProdSetID(Value As Long)
  udtProps.ProdSetID = Value
End Property


Friend Property Let Description(Value As String)
  udtProps.Description = Value
End Property


Friend Property Let ShortDescription(Value As String)
  udtProps.ShortDescription = Value
End Property
