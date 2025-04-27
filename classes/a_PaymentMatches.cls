VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PaymentMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub PostDebtorsPaymentOpenItem(lngInvoiceID As Long, lngPaymentID As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "PostMatchedPayment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@InvoiceID", adInteger, adParamInput, , lngInvoiceID)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PaymentID", adInteger, adParamInput, , lngPaymentID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PaymentMatches.PostDebtorsPaymentOpenItem"
End Sub

Public Sub DeletePosting(lngPaid As Long)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "DeleteMatchedPayment"
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set par = cmd.CreateParameter("@PAID", adInteger, adParamInput, , lngPaid)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing
    
 '---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PaymentMatches.DeletePosting"
End Sub
