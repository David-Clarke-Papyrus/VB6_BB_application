VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_SCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mudtProps As dSCLProps
 
Public Property Get DocStatus() As Integer
    DocStatus = FNN(mudtProps.DocStatus)
End Property
Public Property Let DocStatus(Value As Integer)
    mudtProps.DocStatus = Value
End Property
Public Property Get DocStatusF() As String
    DocStatusF = FNS(mudtProps.DocStatusF)
'    If mudtProps.DocStatus = 2 Then
'        DocStatusF = "Open"
'    ElseIf mudtProps.DocStatus = 4 Then
'        DocStatusF = "Closed"
'    End If
End Property
Public Property Get DateForSort() As String
    DateForSort = Format(mudtProps.DOCDate, "yyyy-mm-dd Hh:Nn:Ss")
End Property
Public Property Get DocDateF() As String
    If mudtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(mudtProps.DOCDate, "dd-mm-yyyy")
    End If
End Property
Public Property Get ClaimNeedsApproval() As Boolean
    ClaimNeedsApproval = FNB(mudtProps.ClaimNeedsApproval)
End Property

Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Get DOCDate() As Date
    DOCDate = FND(mudtProps.DOCDate)
End Property

Public Property Get SupplierName() As String
    SupplierName = FNS(mudtProps.SupplierName)
End Property
Public Property Let SupplierName(Value As String)
    mudtProps.SupplierName = Value
End Property
Public Property Get SupplierAcno() As String
    SupplierAcno = FNS(mudtProps.SupplierAcno)
End Property
Public Property Let SupplierAcno(Value As String)
    mudtProps.SupplierAcno = Value
End Property

Public Property Let TRID(Value As Long)
    mudtProps.TRID = Value
End Property
Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property

Public Property Get TPID() As Long
    TPID = mudtProps.TPID
End Property

Public Property Let TPID(Value As Long)
    mudtProps.TPID = Value
End Property

Private Sub SetState(buffer As String)
  Dim udtData As dRData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get ClaimValue() As Double
    ClaimValue = mudtProps.ClaimValue
End Property
Public Property Get ClaimValueF() As String
    ClaimValueF = FNS(mudtProps.ClaimValueF)
End Property

