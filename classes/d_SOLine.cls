VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_POLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private udtProps As dPOLProps
Dim bApplied As Boolean

Public Property Let Applied(pApplied As Boolean)
    bApplied = pApplied
End Property
Public Property Get Applied() As Boolean
    Applied = bApplied
End Property
Public Property Let DocDate(Value As Date)
    udtProps.DocDate = Value
End Property
Public Property Get DocDate() As Date
    DocDate = Trim(udtProps.DocDate)
End Property
Public Property Get DOCDateF() As String
    If udtProps.DocDate = CDate(0) Then
        DOCDateF = ""
    Else
        DOCDateF = Format(udtProps.DocDate, "dd-mm-yyyy")
    End If
End Property
Public Function SetDocdate(val As String) As Boolean
Dim dteDate As Date
    SetDocdate = True
    If Trim$(val) = "" Then
        val = CDate(0)
    ElseIf Not ConvertToDate(val, dteDate) Then
        SetDocdate = False
        Exit Function
    End If
    udtProps.DocDate = dteDate
End Function

Public Property Get TRID() As Long
    TRID = Trim(udtProps.TRID)
End Property
Public Property Let TRID(Value As Long)
    udtProps.TRID = Value
End Property

'Public Property Let Code(Value As String)
'    udtProps.Code = Value
'End Property
'Public Property Get Code() As String
'    Code = Trim(udtProps.Code)
'End Property

Public Property Get QtyFirm() As Integer
    QtyFirm = Trim(udtProps.QtyFirm)
End Property
Public Property Let QtyFirm(Value As Integer)
    udtProps.QtyFirm = Value
End Property

Public Property Get QtySS() As Integer
    QtySS = Trim(udtProps.QtySS)
End Property
Public Property Let QtySS(Value As Integer)
    udtProps.QtySS = Value
End Property
Public Property Get QtyTotal() As Long
    QtyTotal = udtProps.QtyFirm + udtProps.QtySS
End Property
Public Property Get Discount() As Double
    Discount = udtProps.Discount
End Property
Public Property Let Discount(Value As Double)
    udtProps.Discount = Value
End Property
Public Property Get DiscountF() As String
    If udtProps.Discount = 0 Then
        DiscountF = ""
    Else
        DiscountF = Format(udtProps.Discount, gPercentFormatString)
    End If
End Property

Public Property Get ReceivedSoFar() As Long
    ReceivedSoFar = Trim(udtProps.RecSoFar)
End Property
Public Property Let ReceivedSoFar(Value As Long)
    udtProps.RecSoFar = Value
End Property
Public Property Get POLID() As Long
    POLID = Trim(udtProps.POLID)
End Property
Public Property Let POLID(Value As Long)
    udtProps.POLID = Value
End Property
'Public Property Get QtySS() As Long
'    QtySS = Trim(udtProps.QtySS)
'End Property
'Public Property Let QtySS(Value As Long)
'    udtProps.QtySS = Value
'End Property

'Public Property Get QtyFirm() As Long
'    QtyFirm = Trim(udtProps.QtyFirm)
'End Property
'Public Property Let QtyFirm(Value As Long)
'    udtProps.QtyFirm = Value
'End Property
Public Property Get QtyOS() As Long
    QtyOS = Trim(udtProps.QtyOutstanding)
End Property
Public Property Let QtyOS(Value As Long)
    udtProps.QtyOutstanding = Value
End Property


Public Property Get COLID() As Long
    COLID = Trim(udtProps.COLID)
End Property
Public Property Let COLID(Value As Long)
    udtProps.COLID = Value
End Property
Public Property Get SOLFulfilled() As String
    SOLFulfilled = Trim(udtProps.SOLFulfilled)
End Property
Public Property Get Supplier() As String
    Supplier = FNS(udtProps.Supplier)
End Property
Public Property Get Title() As String
    Title = FNS(udtProps.Title)
End Property
Public Property Let SOLFulfilled(Value As String)
    udtProps.SOLFulfilled = Value
End Property
Public Property Get Ref() As String
    Ref = Trim(udtProps.Ref)
End Property
Public Property Let Ref(Value As String)
    udtProps.Ref = Value
End Property
Public Property Get DocCode() As String
    DocCode = Trim(udtProps.DocCode)
End Property
Public Property Let DocCode(Value As String)
    udtProps.DocCode = Value
End Property

Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property
Public Property Let PID(Value As String)
    udtProps.PID = Value
End Property
Public Property Get Code() As String
    Code = FNS(udtProps.Code)
End Property
Public Property Let Code(Value As String)
    udtProps.Code = Value
End Property
Public Property Get EAN() As String
    EAN = FNS(udtProps.EAN)
End Property
Public Property Let EAN(Value As String)
    udtProps.EAN = Value
End Property
'Public Property Get Ref() As String
'    Ref = FNS(udtProps.Ref)
'End Property
'Public Property Let Ref(Value As String)
'    udtProps.Ref = Value
'End Property

Public Property Get POLTitle() As String
    POLTitle = Trim(udtProps.Title)
End Property
Public Property Let POLTitle(Value As String)
    udtProps.Title = Value
End Property
Public Property Get Author() As String
    Author = Trim(udtProps.Author)
End Property
Public Property Get PreviousActions() As String
    PreviousActions = FNS(udtProps.Actions)
End Property
Public Property Let Author(Value As String)
    udtProps.Author = Value
End Property

Public Property Get POLPrice() As Long
    POLPrice = Trim(udtProps.Price)
End Property
Public Property Get POLPriceF() As String
    POLPriceF = Format(udtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Let POLPrice(Value As Long)
    udtProps.Price = Value
End Property
Public Property Get POLForeignPriceF() As String
    POLForeignPriceF = Format(udtProps.ForeignPrice / udtProps.CForeignDivisor, udtProps.CFormat)
End Property
Public Property Get POLForeignPrice() As Long
    POLForeignPrice = Trim(udtProps.ForeignPrice)
End Property
Public Property Let POLForeignPrice(Value As Long)
    udtProps.ForeignPrice = Value
End Property



Private Sub SetState(buffer As String)
  Dim mudtData As dPOLData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

