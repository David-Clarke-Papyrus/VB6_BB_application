VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_DebtorsPerTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As DebtorsProps
Public Property Get Memo() As String
    Memo = FNS(mudtProps.Memo)
End Property
Public Property Get DOCCode() As String
    DOCCode = FNS(mudtProps.TRCODE)
End Property
Public Property Get DOCDate() As Date
    DOCDate = mudtProps.TRDATE
End Property
Public Property Get DOCCaptureDate() As Date
    DOCCaptureDate = mudtProps.TRCaptureDATE
End Property
Public Property Get DocDateF() As String
    DocDateF = Format(mudtProps.TRDATE, "dd/mm/yyyy")
End Property
Public Property Get DOCCaptureDateF() As String
    DOCCaptureDateF = Format(mudtProps.TRCaptureDATE, "dd/mm/yyyy HH:NN")
End Property
Public Property Get PayableAmount() As String
    PayableAmount = Format(mudtProps.PayableAmount, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get VATAmount() As String
    VATAmount = Format(mudtProps.VATAmount, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get CreditF() As String
    If mudtProps.Credit <= 0 Then mudtProps.Credit = mudtProps.Credit * -1
    If mudtProps.Credit = 0 Then
        CreditF = ""
    Else
        CreditF = Format(mudtProps.Credit, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get DebitF() As String
    If mudtProps.Debit = 0 Then
        DebitF = ""
    Else
        DebitF = Format(mudtProps.Debit, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get Debit() As Double
    Debit = FNDBL(mudtProps.Debit)
End Property
Public Property Get Credit() As Double
    Credit = FNDBL(mudtProps.Credit)
End Property
Public Property Get DocDateForSort() As Date
    DocDateForSort = FND(mudtProps.TRCaptureDATE)
End Property
Public Property Get DocType() As String
    DocType = FNS(mudtProps.DocType)
End Property
Public Property Get DOCID() As Long
    DOCID = mudtProps.TRID
End Property


Private Sub SetState(buffer As String)
  Dim mudtData As DebtorsData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
'Open item values
Public Property Get dbDoc() As String
        dbDoc = FNS(mudtProps.dbDoc)
End Property
Public Property Get dbDate() As String
        dbDate = FNS(mudtProps.dbDate)
End Property
Public Property Get dbDocType() As String
        dbDocType = FNS(mudtProps.dbDocType)
End Property
Public Property Get dbAmt() As String
        dbAmt = FNS(mudtProps.dbAmt)
End Property
Public Property Get dbAmtF() As String
        dbAmtF = Format(FNS(mudtProps.dbAmt), "###,##0.00")
End Property

Public Property Get crDoc() As String
        crDoc = FNS(mudtProps.crDoc)
End Property
Public Property Get crDate() As String
        crDate = FNS(mudtProps.crDate)
End Property
Public Property Get crDocType() As String
        crDocType = FNS(mudtProps.crDocType)
End Property
Public Property Get crAmt() As String
        crAmt = FNS(mudtProps.crAmt)
End Property
Public Property Get crAmtF() As String
        crAmtF = Format(FNS(mudtProps.crAmt), "###,##0.00")
End Property
Public Property Get crTotal() As String
        crTotal = FNS(mudtProps.crTotal)
End Property
Public Property Get crTotalF() As String
        crTotalF = Format(FNS(mudtProps.crTotal), "###,##0.00")
End Property
Public Property Get Balance_OI() As String
        Balance_OI = FNS(mudtProps.Balance_OI)
End Property
Public Property Get Balance_OIF() As String
        Balance_OIF = Format(FNS(mudtProps.Balance_OI), "###,##0.00")
End Property

'Brought forward balances
Public Property Get BFBalTotalF() As String
        BFBalTotalF = Format(mudtProps.BFTotal, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get BFBalCurF() As String
        BFBalCurF = Format(mudtProps.BFCur, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BFBal30F() As String
        BFBal30F = Format(mudtProps.BF30, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BFBal60F() As String
        BFBal60F = Format(mudtProps.BF60, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BFBal90F() As String
        BFBal90F = Format(mudtProps.BF90, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get BFBal120F() As String
        BFBal120F = Format(mudtProps.BF120, oPC.Configuration.DefaultCurrency.FormatString)
End Property

