VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_COL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private udtProps As COLProps

Private Sub SetState(buffer As String)
  Dim mudtData As COLData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
    
    
Public Property Get BranchCode() As String
    BranchCode = FNS(udtProps.STORECODE)
End Property

Public Property Get DocumentCode() As String
    DocumentCode = FNS(udtProps.DocumentCode)
End Property

Public Property Get DocumentCaptureDateF() As String
    If udtProps.DocumentCaptureDate <> 0 Then
        DocumentCaptureDateF = Format(udtProps.DocumentCaptureDate, "dd-mm-yy Hh:Nn")
    Else
        DocumentCaptureDateF = ""
    End If
End Property

Public Property Get DocumentIssueDateF() As String
    If udtProps.DocumentIssueDate <> 0 Then
        DocumentIssueDateF = Format(udtProps.DocumentIssueDate, "dd-mm-yy Hh:Nn")
    Else
        DocumentIssueDateF = ""
    End If
End Property
Public Property Get DocumentCapturedBy() As String
    DocumentCapturedBy = FNS(udtProps.DocumentCapturedBy)
End Property
Public Property Get CustomerName() As String
    CustomerName = FNS(udtProps.CustomerName)
End Property
Public Property Get CustomerAcno() As String
    CustomerAcno = FNS(udtProps.CustomerAcno)
End Property
Public Property Get CustomerPhone() As String
    CustomerPhone = FNS(udtProps.CustomerPhone)
End Property
Public Property Get OrderlineRef() As String
    OrderlineRef = FNS(udtProps.OrderlineRef)
End Property

Public Property Get OrderlineQty() As Double
    OrderlineQty = udtProps.OrderlineQty
End Property
Public Property Get OrderlineQtyF() As String
    OrderlineQtyF = Format(OrderlineQty, "###,##0")
End Property

Public Property Get OrderlineQtyDispatched() As Double
    OrderlineQtyDispatched = udtProps.OrderlineQtyDispatched
End Property
Public Property Get OrderlineQtyDispatchedF() As String
    OrderlineQtyDispatchedF = Format(OrderlineQtyDispatched, "###,##0")
End Property

Public Property Get OrderlineQtyOutstanding() As Double
    OrderlineQtyOutstanding = udtProps.OrderlineQtyOutstanding
End Property
Public Property Get OrderlineQtyOutstandingF() As String
    OrderlineQtyOutstandingF = Format(OrderlineQtyOutstanding, "###,##0")
End Property

Public Property Get OrderlineDiscount() As Double
    OrderlineDiscount = udtProps.OrderlineDiscount
End Property
Public Property Get OrderlineDiscountF() As String
    OrderlineDiscountF = Format(OrderlineDiscount, "###,##0")
End Property

Public Property Get ProductTitle() As String
    ProductTitle = FNS(udtProps.ProductTitle)
End Property
Public Property Get ProductPublisher() As String
    ProductPublisher = FNS(udtProps.ProductPublisher)
End Property
Public Property Get ProductEAN() As String
    ProductEAN = FNS(udtProps.ProductEAN)
End Property
Public Property Get SellingPrice() As Double
    SellingPrice = FNS(udtProps.SellingPrice)
End Property
Public Property Get SellingPriceF() As String
    SellingPriceF = Format(SellingPrice, "###,##0.00")
End Property




