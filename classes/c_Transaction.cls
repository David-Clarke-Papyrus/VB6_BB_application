VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_Transaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private colTransaction As Collection

Public Function Count() As Long
  Count = colTransaction.Count
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = colTransaction.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As d_Customer
Attribute Item.VB_UserMemId = 0
  Set Item = colTransaction(Index)
End Function


Public Sub Load(TradingPartnerID As Long, TransactionType As String)
  Dim objTR As d_Transaction
  Dim rs As ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim strSQL As String
  Dim TRType As Long
  
  TRType = TransactionType 'converting string to number
  
On Error GoTo ErrorHandler
  
  Set cmd.ActiveConnection = dbObj
  strSQL = "SELECT TransactionID FROM Transaction WHERE TR_TP_ID=" & TradingPartnerID & _
           " AND WHERE TR_Type =" & TRType
  Set rs = New ADODB.Recordset
  cmd.CommandText = strSQL
  rs.Open cmd, , adOpenDynamic, adLockReadOnly
  If rs.EOF Then
    MsgBox "No Transactions to Load!"
  Else
    Do While Not rs.EOF
      Set objTR = New d_Transaction
      With objTR
        .TRType = rs!TR_Type
        .TradingPartnerID = rs!TR_TP_ID
        .TRDate = rs!TR_Date
        .TRCaptureDate = rs!TR_CaptureDate
        .Code = rs!TR_Code
        .Amount = rs!TR_Amount
        .DiscountPercent = rs!TR_DiscountPercent
        .AddressID = rs!TR_Address
        .Void = rs!TR_Void
        .Printed = rs!TR_Printed
        .Memo = rs!TR_Memeo
        .Completed = rs!TR_Complete
      End With
      colTransaction.Add objTR
      Set objTR = Nothing
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  Exit Sub

ErrorHandler:
  MsgBox "Error while loading Transactions!" & Chr(10) & _
         "ERROR: " & Error
  
End Sub

Public Function GetLastTransactionID(TradingPartnerID As Long, TransactionType As String) As Long
  Dim rs As New ADODB.Recordset
  Dim cmd As New ADODB.Command
  Dim strSQL As String
  
On Error GoTo ErrorHandler
  strSQL = "SELECT TransactionID FROM Transaction WHERE TR_TP_ID=" & TradingPartnerID & _
           " AND TR_Type=" & CheckTRType(TransactionType)
  Set cmd.ActiveConnection = dbObj
  cmd.CommandText = strSQL
  With rs
    .Open cmd, , adOpenDynamic, adLockReadOnly
    .MoveLast
    GetLastTransactionID = !TransactionID
    .Close
  End With
  Exit Function
ErrorHandler:
  MsgBox "Can't GetLastTransaction!" & Chr(10) & "ERROR: " & Error
End Function



Private Sub Class_Initialize()
  Set colTransaction = New Collection
End Sub
