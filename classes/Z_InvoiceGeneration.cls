VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_InvoiceGeneration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CreateFinalInvoiceFromDeliveredCustomerOrder(pCOID As Long, pStatus As enStatus)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateInvoiceFromDeliveredCustomerOrder"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@COID", adInteger, adParamInput, , pCOID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , 2)  'Create the invoice as issued
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_InvoiceGeneration.CreateFinalInvoiceFromCO(pCOID)", pCOID
End Sub

Public Sub GetAlreadyInvoicedLines(pCOID As Long, pResult As String)
'MsgBox "Not yet written"
    pResult = ""
End Sub

Public Sub GenerateInvoicesForGDNsIssuedFortCustomerWithoutCompleteRequiremen()
    On Error GoTo errHandler
'
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateInvoiceFromUninvoicedGDNs"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , 2)  'Create the invoice as issued
    cmd.Parameters.Append par
    Set par = Nothing
    
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_InvoiceGeneration.GenerateInvoicesForGDNsIssuedFortCustomerWithoutCompleteRequiremen", , EA_NORERAISE
End Sub
Public Sub GenerateGDNsFOrPreinvoicedCustomers()
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "GenerateGDNsFOrPreinvoicedCustomers"
    cmd.CommandType = adCmdStoredProc
    Set par = cmd.CreateParameter("@STATUS", adInteger, adParamInput, , 4)  'Create the invoice as issued
    cmd.Parameters.Append par
    Set par = Nothing
    
    
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
    ErrorIn "Z_InvoiceGeneration.GenerateGDNsFOrPreinvoicedCustomers", , EA_NORERAISE
End Sub

