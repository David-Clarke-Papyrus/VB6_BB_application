VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As dSOProps

Public Property Get TransactionID() As Long
    TransactionID = mudtProps.TransactionID
End Property
Public Property Let TransactionID(Val As Long)
    mudtProps.TransactionID = Val
End Property


Public Property Get OrderNumber() As String
    OrderNumber = mudtProps.TransCode
End Property
Public Property Let OrderNumber(Val As String)
    mudtProps.TransCode = Val
End Property

Public Property Get TPName() As String
    TPName = Trim(mudtProps.TPName)
End Property
Public Property Let TPName(Val As String)
    mudtProps.TPName = Val
End Property

Public Property Get Status() As String
    If mudtProps.Status = 1 Then
        Status = "VOID"
    ElseIf mudtProps.Status = 2 Then
        Status = "IN PROCESS"
    ElseIf mudtProps.Status = 3 Then
        Status = "ISSUED"
    ElseIf mudtProps.Status = 4 Then
        Status = "COMPLETE"
    ElseIf mudtProps.Status = 7 Then
        Status = "PROFORMA"
    End If
End Property
'Public Property Let Status(Val As Integer)
'    mudtProps.Status = Val
'End Property


Public Property Get TPACCNo() As String
    TPACCNo = mudtProps.TPACCNo
End Property
Public Property Let TPACCNo(Val As String)
    mudtProps.TPACCNo = Val
End Property

Public Property Get InvoiceDate() As Date
    InvoiceDate = mudtProps.TransDate
End Property
Public Property Get TDateFormatted() As String
    If mudtProps.TransDate = CDate(0) Then
        TDateFormatted = ""
    Else
        TDateFormatted = Format(mudtProps.TransDate, "dd-mm-yyyy")
    End If
End Property
Private Sub SetState(Buffer As String)
  Dim mudtData As dSOData
  
  mudtData.Buffer = Buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(Buffer As String)
  SetState Buffer
End Sub

