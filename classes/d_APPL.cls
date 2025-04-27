VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_APPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As dAPPLProps

Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.DOCCode)
End Property
Public Property Get CodeF() As String
    CodeF = FNS(mudtProps.code)
End Property
Public Property Get Title() As String
    Title = FNS(mudtProps.Title)
End Property

Public Property Get TRDateF() As String
    TRDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
End Property
Public Property Get TRDateForSort() As String
    TRDateForSort = Format(mudtProps.DOCDate, "yyyy/mm/dd")
End Property

Public Property Let APPLID(Value As Long)
    mudtProps.APPLID = Value
End Property
Public Property Get APPLID() As Long
    APPLID = FNN(mudtProps.APPLID)
End Property


Public Property Let Qty(Value As Long)
    mudtProps.Qty = Value
End Property
Public Property Get Qty() As Long
    Qty = FNN(mudtProps.Qty)
End Property
Public Property Let QtyReturned(Value As Long)
    mudtProps.QtyReturned = Value
End Property
Public Property Get QtyReturned() As Long
    QtyReturned = FNN(mudtProps.QtyReturned)
End Property
Public Property Get DiscountRate() As Double
    DiscountRate = FNDBL(mudtProps.DiscountRate)
End Property


Private Sub SetState(buffer As String)
  Dim mudtData As dAPPLData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub





