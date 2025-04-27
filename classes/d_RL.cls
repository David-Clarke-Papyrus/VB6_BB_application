VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_RL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mudtProps As dRLProps
 

Public Property Get ProductDescription() As String
    ProductDescription = FNS(mudtProps.ProductDescription)
End Property

Public Property Let ProductDescription(Value As String)
    mudtProps.ProductDescription = Value
End Property
Public Property Get code() As String
    code = FNS(mudtProps.code)
End Property
Public Property Get EAN() As String
    EAN = FNS(mudtProps.EAN)
End Property
Public Property Get Section() As String
    Section = FNS(mudtProps.Section)
End Property
Public Property Get Note() As String
    Note = FNS(mudtProps.Note)
End Property

Public Property Get Publisher() As String
    Publisher = FNS(mudtProps.Publisher)
End Property
Public Property Get SupplierInvoiceRef() As String
    SupplierInvoiceRef = FNS(mudtProps.SupplierInvoiceRef)
End Property
Public Property Let SupplierInvoiceRef(val As String)
    mudtProps.SupplierInvoiceRef = val
End Property
Public Property Get SupplierInvoiceDate() As Date
    SupplierInvoiceDate = mudtProps.SupplierInvoiceDate
End Property
Public Property Let SupplierInvoiceDate(val As Date)
    mudtProps.SupplierInvoiceDate = val
End Property
Public Property Get SupplierInvoiceDateF() As String
    SupplierInvoiceDateF = Format(FND(mudtProps.SupplierInvoiceDate), "dd/mm/yyyy")
End Property

Public Property Get Pubcode() As String
    Pubcode = FNS(mudtProps.Pubcode)
End Property
Public Property Get RLID() As Long
    RLID = mudtProps.RLID
End Property

Public Property Let code(Value As String)
    mudtProps.code = Value
End Property

Public Property Get PID() As String
    PID = FNS(mudtProps.PID)
End Property

Public Property Let PID(Value As String)
    mudtProps.PID = Value
End Property




Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property


Public Property Let TRID(Value As Long)
    mudtProps.TRID = Value
End Property

Private Sub SetState(buffer As String)
  Dim udtData As dRLData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Let DELLID(Value As Long)
    mudtProps.DELLID = Value
End Property
Public Property Get DELLID() As Long
    DELLID = mudtProps.DELLID
End Property

Public Property Let QtySystem(Value As Long)
    mudtProps.QtySystem = Value
End Property
Public Property Get QtySystem() As Long
    QtySystem = mudtProps.QtySystem
End Property
Public Property Let QtyRequested(Value As Long)
    mudtProps.QtyRequested = Value
End Property
Public Property Get QtyRequested() As Long
    QtyRequested = mudtProps.QtyRequested
End Property
Public Property Let QtyApproved(Value As Long)
    mudtProps.QtyApproved = Value
End Property
Public Property Get QtyApproved() As Long
    QtyApproved = mudtProps.QtyApproved
End Property
Public Property Let QtyReturned(Value As Long)
    mudtProps.QtyReturned = Value
End Property
Public Property Get QtyReturned() As Long
    QtyReturned = mudtProps.QtyReturned
End Property
Public Property Let QtyCounted(Value As Long)
    mudtProps.QtyCounted = Value
End Property
Public Property Get QtyCounted() As Long
    QtyCounted = mudtProps.QtyCounted
End Property
Public Property Let Valu(Value As Long)
    mudtProps.val = Value
End Property
Public Property Get Valu() As Long
    Valu = mudtProps.val
End Property


