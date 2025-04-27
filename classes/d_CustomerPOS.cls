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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As CustomerPOSProps
Dim mKey As String

Private Sub SetState(buffer As String)
  Dim mudtData As CustomerPOSData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(val As String)
    mKey = val
End Property

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get Fullname2() As String
    Fullname2 = Name & IIf(Len(title) > 0 Or Len(Initials) > 0, ", ", "") & title & IIf(Len(title) > 0, " ", "") & Initials
End Property
    
Public Property Get Name() As String
    Name = FNS((udtProps.Name))
End Property
Public Property Let Name(Value As String)
    udtProps.Name = Value
End Property
Public Property Get FullIdentification() As String
    FullIdentification = FNS(udtProps.FullIdentification)
End Property
    
Public Property Let FullIdentification(Value As String)
    udtProps.FullIdentification = Value
End Property
Public Property Get title() As String
    title = FNS(udtProps.title)
End Property
    
Public Property Let title(Value As String)
    udtProps.title = Value
End Property
Public Property Get Initials() As String
    Initials = FNS(udtProps.Initials)
End Property
    
Public Property Let Initials(Value As String)
    udtProps.Initials = Value
End Property
Public Property Get AcNo() As String
    AcNo = FNS(udtProps.AcNo)
End Property
Public Property Let AcNo(Value As String)
    udtProps.AcNo = Value
End Property

Public Property Get CType() As String
    CType = FNS(udtProps.CType)
End Property
Public Property Let CType(Value As String)
    udtProps.CType = Value
End Property


Public Property Get Phone() As String
    Phone = FNS(udtProps.Phone)
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
Public Property Get Address() As String
    Address = FNS(udtProps.Address)
End Property
