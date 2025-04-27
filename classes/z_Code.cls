VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Enum enCodeType
'    enCustomerOrder = 19
'    enAppro = 13
'    enApproReturn = 14
'    enCashSale = 20
'    enSupplierOrder = 16
'    enStockTake = 17
'    enOOS = 25
'    enInvoice = 21
'    enProduct = 23
'    enCReditnote = 26
'    enDelivery = 18
'    enTransfer = 22
'    enReturn = 28
'    en230 = 230
'    en240 = 240
'    en250 = 250
'    en260 = 260
'    en270 = 270
'    en280 = 280
'End Enum
Public Function GetNextCode(psequence As enCodeType, pGetandWrite As String) As Long

'    GetNextCode - returns the next code in sequence for a chosen sequence.
'    Used for automatically generating ordernumbers etc.

      On Error GoTo Err_GetNextCode
'
Dim iarg As Long
Dim iNextCode As Long
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
    
    iarg = psequence
    cnt = 0
    On Error Resume Next
    Do
        Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        rs.open "Select * FROM tControl WHERE ID = " & iarg, oPC.COShort, adOpenStatic, adLockPessimistic
        cnt = cnt + 1
        If cnt > 50 Then
            If MsgBox("Cannot allocate a transaction code for " & psequence & ". Retry?", vbOKCancel, "Table locked") = vbCancel Then
                Error 9999
                GoTo Err_GetNextCode
            Else
                cnt = 0
            End If
        End If
'        DoEvents
    Loop While (Err = 3211)

    On Error GoTo Err_GetNextCode
     iNextCode = rs![Value]
    If UCase(pGetandWrite) = "GETANDWRITE" Then
           rs![Value] = iNextCode + 1
           rs.Update
    End If
    GetNextCode = iNextCode + 1


EXIT_GetNextCode:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

Err_GetNextCode:
    Select Case Err
'
        Case Else
             MsgBox "GetNextCode:Failed"
             Resume EXIT_GetNextCode
             Exit Function
             Resume
    End Select
End Function

