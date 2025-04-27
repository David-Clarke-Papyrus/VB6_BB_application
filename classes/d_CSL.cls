VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mudtProps As dCSLProps
 
Public Property Get CodeF() As String
'Dim oProdCode As New z_ProdCode
'    oProdCode.Load FNS(mudtProps.Code)
'    CodeF = oProdCode.ISBNFormatted
    
    CodeF = FNS(mudtProps.Code)
End Property
Public Property Get ID() As Long
    ID = mudtProps.ID
End Property
Public Property Get Qty() As Long
    Qty = mudtProps.Qty
End Property

Public Property Get DateTimeForSort() As String
    DateTimeForSort = Format(mudtProps.DateTime, "yyyy-mm-dd Hh:Nn")
End Property
Public Property Get TDateF() As String
    If mudtProps.DateTime = CDate(0) Then
        TDateF = ""
    Else
        TDateF = Format(mudtProps.DateTime, "dd-mm-yyyy Hh:Nn")
    End If
End Property
 

Public Property Get EAN() As String
    EAN = FNS(mudtProps.EAN)
End Property


Public Property Get Title() As String
    Title = FNS(mudtProps.Title)
End Property
Public Property Get Price() As Long
    Price = mudtProps.Price
End Property
Public Property Get PriceF() As String
    PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property


Private Sub SetState(buffer As String)
  Dim udtData As dCSLData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub

