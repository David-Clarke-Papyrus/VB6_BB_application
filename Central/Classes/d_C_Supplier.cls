VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private udtProps As dSupplierProps

Private Sub SetState(Buffer As String)
  Dim mudtData As dSupplierData
  
  mudtData.Buffer = Buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(Buffer As String)
  SetState Buffer
End Sub
    
Public Property Get Name() As String
    Name = Trim$(udtProps.Name)
End Property
    
Public Property Let Name(Value As String)
    udtProps.Name = Value
End Property
Public Property Get Fullname() As String
    Fullname = Trim$(udtProps.Fullname)
End Property
    
Public Property Let Fullname(Value As String)
    udtProps.Fullname = Value
End Property
Public Property Get AcNo() As String
    AcNo = Trim$(udtProps.AcNo)
End Property
    
Public Property Let AcNo(Value As String)
    udtProps.AcNo = Value
End Property
Public Property Get Phone() As String
    Phone = Trim$(udtProps.Phone)
End Property
    
Public Property Let Phone(Value As String)
    udtProps.Phone = Value
End Property
    
Public Property Get ID() As Long
    ID = udtProps.ID
End Property

Public Property Let ID(Value As Long)
    udtProps.ID = Value
End Property
Public Property Get DefaultAddressID() As Long
    DefaultAddressID = udtProps.DefaultAddressID
End Property

