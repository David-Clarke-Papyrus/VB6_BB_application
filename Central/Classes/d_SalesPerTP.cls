VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_SalesPerTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As DocsTPProps

Public Property Get DocCode() As String
    DocCode = Trim$(mudtProps.TRCODE)
End Property
Public Property Get DocDate() As Date
    DocDate = mudtProps.TRDATE
End Property
Public Property Get DocValue() As String
    DocValue = Format(mudtProps.TRValue / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString)
End Property
Public Property Get DocDateF() As String
    DocDateF = Format(mudtProps.TRDATE, "dd/mm/yyyy")
End Property
Public Property Get DocDateForSort() As String
    DocDateForSort = Format(mudtProps.TRDATE, "yyyy/mm/dd")
End Property
Public Property Get DocType() As Integer
    DocType = mudtProps.Type
End Property
Public Property Get DocTypeF() As String
    Select Case mudtProps.Type
    Case 9
        DocTypeF = "APPRO"
    Case 10
        DocTypeF = "APPRO RETURN"
    Case 3
        DocTypeF = "Invoice"
    Case 2
        DocTypeF = "WANT"
    Case 6
        If mudtProps.OrderType = 1 Then
            DocTypeF = "Order"
        Else
            DocTypeF = "Want"
        End If
    End Select
    
End Property

Public Property Get DOCID() As Long
    DOCID = mudtProps.TRID
End Property
Public Property Get DocStatus() As String
        If mudtProps.TRSTATUS = 1 Then
            DocStatus = "VOID"
        ElseIf mudtProps.TRSTATUS = 2 Then
            DocStatus = "IN PROCESS"
        ElseIf mudtProps.TRSTATUS = 3 Then
            DocStatus = "ISSUED"
        ElseIf mudtProps.TRSTATUS = 4 Then
            DocStatus = "COMPLETE"
        ElseIf mudtProps.TRSTATUS = 7 Then
            DocStatus = "PROFORMA"
        ElseIf mudtProps.TRSTATUS = 6 Then
            DocStatus = "CANCELLED"
        End If
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As DocsTPData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub



