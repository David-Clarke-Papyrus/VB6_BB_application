VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
  CustID As Long
  Name As String * 50
  Phone As String * 25
  AcNo As String * 12
End Type

Private udtProps As DisplayProps

Public Property Get CustomerID() As Long
  CustomerID = udtProps.CustID
End Property

Public Property Get Name() As String
  Name = FNS(udtProps.Name)
End Property

Public Property Get Phone() As String
  Phone = FNS(udtProps.Phone)
End Property
Public Property Get AcNo() As String
  AcNo = FNS(udtProps.AcNo)
End Property
Friend Property Let AcNo(Value As String)
  udtProps.AcNo = Value
End Property


Friend Property Let CustomerID(Value As Long)
  udtProps.CustID = Value
End Property


Friend Property Let Name(Value As String)
  udtProps.Name = Value
End Property


Friend Property Let Phone(Value As String)
  udtProps.Phone = Value
End Property
