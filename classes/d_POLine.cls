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
Attribute VB_Exposed = False
Option Explicit
Private udtProps As dPOLProps
Dim bApplied As Boolean

Public Property Let Applied(pApplied As Boolean)
    bApplied = pApplied
End Property
Public Property Get Applied() As Boolean
    Applied = bApplied
End Property
Public Property Let DOCDate(Value As Date)
    udtProps.DOCDate = Value
End Property
Public Property Get DOCDate() As Date
    DOCDate = Trim(udtProps.DOCDate)
End Property
Public Property Get DocDateF() As String
    If udtProps.DOCDate = CDate(0) Then
        DocDateF = ""
    Else
        DocDateF = Format(udtProps.DOCDate, "dd-mm-yyyy")
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
    udtProps.DOCDate = dteDate
End Function
Public Property Get ETA() As Date
    ETA = Trim(udtProps.ETA)
End Property
Public Property Get ETAF() As String
    If udtProps.ETA = CDate(0) Then
        ETAF = ""
    Else
        ETAF = Format(udtProps.ETA, "dd-mm-yyyy")
    End If
End Property
Public Property Get LastActionDate() As Date
    LastActionDate = Trim(udtProps.LastActionDate)
End Property
Public Property Get LastActionDateF() As String
    If udtProps.LastActionDate = CDate(0) Then
        LastActionDateF = ""
    Else
        LastActionDateF = Format(udtProps.LastActionDate, "dd-mm-yyyy")
    End If
End Property
Public Property Get LastAction() As String
    LastAction = Trim(udtProps.LastAction)
End Property

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
    ReceivedSoFar = Trim(udtProps.QtyReceivedSoFar)
End Property
Public Property Let ReceivedSoFar(Value As Long)
    udtProps.QtyReceivedSoFar = Value
End Property
Public Property Get QtyUnMatchedTmp() As Long
    QtyUnMatchedTmp = Trim(udtProps.QtyUnMatchedTmp)
End Property
Public Property Let QtyUnMatchedTmp(Value As Long)
    udtProps.QtyUnMatchedTmp = Value
End Property

Public Property Get QtySSUnMatchedTmp() As Long
    QtySSUnMatchedTmp = Trim(udtProps.QtySSUnMatchedTmp)
End Property
Public Property Let QtySSUnMatchedTmp(Value As Long)
    udtProps.QtySSUnMatchedTmp = Value
End Property

Public Property Get QtyFIRMUnMatchedTmp() As Long
    QtyFIRMUnMatchedTmp = Trim(udtProps.QtyFIRMUnMatchedTmp)
End Property
Public Property Let QtyFIRMUnMatchedTmp(Value As Long)
    udtProps.QtyFIRMUnMatchedTmp = Value
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

Public Property Get QtySSOS() As Long
    QtySSOS = Trim(udtProps.QtySSOutstanding)
End Property
Public Property Let QtySSOS(Value As Long)
    udtProps.QtySSOutstanding = Value
End Property

Public Property Get QtyFIRMOS() As Long
    QtyFIRMOS = Trim(udtProps.QtyFirmOutstanding)
End Property
Public Property Let QtyFIRMOS(Value As Long)
    udtProps.QtyFirmOutstanding = Value
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
Public Property Get SupplierID() As Long
    SupplierID = FNN(udtProps.SupplierID)
End Property
Public Property Get Actions() As String
    Actions = FNS(udtProps.Actions)
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
Public Property Get DOCCode() As String
    DOCCode = Trim(udtProps.DOCCode)
End Property
Public Property Let DOCCode(Value As String)
    udtProps.DOCCode = Value
End Property

Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property
Public Property Let PID(Value As String)
    udtProps.PID = Value
End Property
Public Property Get code() As String
    code = FNS(udtProps.code)
End Property
Public Property Let code(Value As String)
    udtProps.code = Value
End Property
Public Property Get EAN() As String
    EAN = FNS(udtProps.EAN)
End Property
Public Property Let EAN(Value As String)
    udtProps.EAN = Value
End Property
Public Property Get CodeF() As String
    CodeF = FNS(udtProps.CodeF)
End Property

Public Property Get POLTitle() As String
    POLTitle = Trim(udtProps.Title)
End Property
Public Property Let POLTitle(Value As String)
    udtProps.Title = Value
End Property
Public Property Get MainAuthor() As String
    MainAuthor = Trim(udtProps.MainAuthor)
End Property
Public Property Get PreviousActions() As String
    PreviousActions = FNS(udtProps.Actions)
End Property
Public Property Let Author(Value As String)
    udtProps.MainAuthor = Value
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

