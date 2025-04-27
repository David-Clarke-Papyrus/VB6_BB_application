VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_APP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private udtProps As APPPOSProps

Private Sub SetState(buffer As String)
  Dim mudtData As APPPOSData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
    
Public Property Get DOCCode() As String
    DOCCode = FNS(udtProps.DOCCode)
End Property
    

Public Property Get DOCDate() As Date
    DOCDate = FND(udtProps.DOCDate)
End Property
Public Property Get DOCDateF() As String
    DOCDateF = Format(udtProps.DOCDate, "dd/mm/yyyy")
End Property

Public Property Get APPID() As Long
    APPID = FNN(udtProps.APPID)
End Property





