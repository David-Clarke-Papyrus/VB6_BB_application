VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DisplayProps
    ID As Long
    CompanyName As String * 50
    CompanyCode As String * 2
    End Type

Private udtProps As DisplayProps

Public Property Get ID() As Long
    ID = udtProps.ID
End Property
Friend Property Let ID(val As Long)
    udtProps.ID = val
End Property
Public Property Get CompanyName() As String
    CompanyName = Trim(udtProps.CompanyName)
End Property
Public Property Let CompanyName(val As String)
    udtProps.CompanyName = val
End Property
Public Property Get CompanyCode() As String
    CompanyCode = Trim(udtProps.CompanyCode)
End Property
Public Property Let CompanyCode(val As String)
    udtProps.CompanyCode = val
End Property

