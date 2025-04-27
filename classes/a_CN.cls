VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oSM As z_StockManager
Private oFSo As FileSystemObject

Dim lngLPLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLPayableLessVAT_tot As Long
Dim lngLDeposit_tot As Long
Dim lngFDeposit_tot As Long
Dim lngQty_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Private xMLDoc As ujXML
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Dim bForeign As Boolean
Event Valid(pMsg As String)
Event TotalChange(lngTotal As Long, strtotal As String, lngTotalVAT As Long, strTotalVAT As String)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Private mudtProps As CNProps
Private mcolStack As Collection

Private WithEvents mcolLines As ch_CNL
Attribute mcolLines.VB_VarHelpID = -1
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private colClassors As Collection
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
Private oCurrency_Foreign As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Const CDOCCODE = "CN"
'Type tTotal
'    Description As String
'    val As Long
'    AmtFormatted As String
'    TotalType As String
'    RunningTotal As Long
'    Sign As String
'End Type

Dim arDocTotals() As tTotal
Dim arDocTotalsF() As tTotal
Dim RunningTotal As Long
Dim jTotal As Long


Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Public Function SetMemo(val As String) As Boolean
   ' If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.TPMemo) Then
        Err.Raise 384
    End If
    If val <> mudtProps.TPMemo Then
        mudtProps.TPMemo = val
        SetDirty True
    End If
    SetMemo = True
    Exit Function
End Function
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.TPMemo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Memo"
End Property

Public Property Get COMPID() As Long
    On Error GoTo errHandler
    COMPID = mudtProps.COMPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.COMPID"
End Property
Public Property Get ForeignCurrencyID() As Long
    On Error GoTo errHandler
    ForeignCurrencyID = mudtProps.ForeignCurrencyID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ForeignCurrencyID"
End Property
Public Property Let ForeignCurrencyID(val As Long)
    On Error GoTo errHandler
Dim oCurr As a_Currency
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ForeignCurrencyID = val
    SetDirty True
    For Each oCurr In oPC.Configuration.Currencies
        If oCurr.ID = mudtProps.ForeignCurrencyID Then
            Set oCurrency_Foreign = oCurr
        End If
    Next
  '  Me.CurrencyFactor = oCurrency_Foreign.Factor
    RecalculateAllLines
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ForeignCurrencyID(val)", val
End Property
Public Property Get ForeignCurrency() As a_Currency
    Set ForeignCurrency = oCurrency_Foreign
End Property
Public Property Let CurrRate(val As Double)
    On Error GoTo errHandler
    mudtProps.CurrRate = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.CurrRate(val)", val
End Property
Public Property Let COMPID(val As Long)
    On Error GoTo errHandler
Dim oComp As a_Company
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COMPID = val
    SetDirty True
    For Each oComp In oPC.Configuration.Companies
        If oComp.ID = mudtProps.COMPID Then
            Set oBillingCompany = oComp
        End If
    Next
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.COMPID(val)", val
End Property


'-=------------------------
Public Function RecalculateAllLines()
    On Error GoTo errHandler
Dim oCNL As a_CNL
    For Each oCNL In CNLines
        oCNL.CalculateLine
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.RecalculateAllLines"
End Function
Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim oCNL As a_CNL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLDeposit_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngLPayableLessVAT_tot = 0
    lngQty_tot = 0
    mobjValid.RuleBroken "OVERFLOW", False
    'NOTE Some items on the invoice might be ServiceItem e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    For Each oCNL In mcolLines
       'If oCNL.Fulfilled <> "CAN" Then
            oCNL.CalculateLine
            lngVAT_tot = lngVAT_tot + oCNL.PLessDiscExtVAT(False)
            lngDisc_tot = lngDisc_tot + oCNL.PDiscExt(False)
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + oCNL.PLessDiscExt(False)
            lngLPayableLessVAT_tot = lngLPayableLessVAT_tot + oCNL.PLessDiscExtNonVAT(False)
            If oCNL.ServiceItem = False Then
                lngQty_tot = lngQty_tot + oCNL.Qty
            End If
    Next
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalNonVAT = lngLPayableLessVAT_tot
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngLDeposit_tot
    If (Not VATable) And ShowVAT Then
        mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngVAT_tot
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim oCNL As a_CNL
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngFPLessDiscExt_tot = 0
    lngFDeposit_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    For Each oCNL In mcolLines
      '  If oCNL.Fulfilled <> "CAN" Then
            lngVAT_tot = lngVAT_tot + oCNL.PLessDiscExtVAT(True)
            lngDisc_tot = lngDisc_tot + oCNL.PDiscExt(True)
            lngFPLessDiscExt_tot = lngFPLessDiscExt_tot + oCNL.PLessDiscExt(True)
     '   End If
    Next
    mudtProps.TotalDiscount_Foreign = lngDisc_tot
    mudtProps.TotalVAT_Foreign = lngVAT_tot
    mudtProps.TotalExtension_Foreign = lngFPLessDiscExt_tot
    mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot - lngFDeposit_tot
    If (Not Me.VATable) And ShowVAT Then
        mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot - lngVAT_tot
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'overflow
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.CalculateTotal_Foreign"
End Function
'Public Property Get TotalInvoiceDiscountAmountF() As Long
'    TotalInvoiceDiscountAmountF = Format(lngTotalInvoiceDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
Public Property Get TotalLessDiscExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalLessDiscExtF = Format(mudtProps.TotalExtension_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalLessDiscExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalLessDiscExtF(bForeign)", bForeign
End Property
Public Property Get TotalLessDiscExt(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalLessDiscExt = mudtProps.TotalExtension_Foreign
    Else
        TotalLessDiscExt = mudtProps.TotalExtension
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalLessDiscExt(bForeign)", bForeign
End Property
Public Property Get TotalVAT(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalVAT = mudtProps.TotalVAT_Foreign
    Else
        TotalVAT = mudtProps.TotalVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalVAT(bForeign)", bForeign
End Property

Public Property Get TotalVATF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalVATF = Format(mudtProps.TotalVAT_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalVATF(bForeign)", bForeign
End Property
Public Property Get TotalPayable(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalPayable = mudtProps.TotalPayable_Foreign
    Else
        TotalPayable = mudtProps.TotalPayable
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalPayable(bForeign)", bForeign
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayableF = Format(TotalPayable(bForeign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalPayableF = Format(TotalPayable(bForeign) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalPayableF(bForeign)", bForeign
End Property
Public Property Get TotalDiscount(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalDiscount = mudtProps.TotalDiscount_Foreign
    Else
        TotalDiscount = mudtProps.TotalDiscount
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalDiscount(bForeign)", bForeign
End Property
Public Property Get TotalDiscountF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalDiscountF = Format(mudtProps.TotalDiscount_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalDiscountF(bForeign)", bForeign
End Property


Public Property Get TotalNonVATF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalNonVATF = Format(mudtProps.TotalNonVAT_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalNonVATF = Format(mudtProps.TotalNonVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalNonVATF(bForeign)", bForeign
End Property

Public Property Get TotalQty() As Long
    On Error GoTo errHandler

        TotalQty = mudtProps.TotalQty

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalQty"
End Property
Public Property Get TotalQtyF() As String
    On Error GoTo errHandler

        TotalQtyF = CStr(TotalQty)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TotalQtyF"
End Property

'Public Function CalculateTotal()
'Dim oCNL As a_CNL
'Dim bLineDiscountApplied As Boolean
'    bLineDiscountApplied = False
'    lngTotal1 = 0
'    lngTotal2 = 0
'    lngTotal3 = 0
'    lngTotal4 = 0
'    lngTotalDeposit = 0
'    lngTotalVAT = 0
'    lngTotalDiscount = 0
'    For Each oCNL In mcolLines
'        lngTotalVAT = lngTotalVAT + oCNL.VATValue
'        lngTotalDiscount = lngTotalDiscount + oCNL.DiscountValue
'    '    lngTotal3 = lngTotal3 + oCNL.Extension
'        lngTotal4 = lngTotal4 + oCNL.ExtensionInclDeposit
'    '    lngTotal2 = lngTotal2 + oCNL.Extension
'    '    lngTotal1 = lngTotal1 + oCNL.ExtensionSimple
'    Next
'    mudtProps.TotalDiscount = lngTotalDiscount
'    mudtProps.TotalVAT = lngTotalVAT
'    mudtProps.TotalPayable = lngTotal4
'    If (Not VATable) And ShowVAT Then
'        mudtProps.TotalPayable = lngTotal4 - lngTotalVAT
'    End If
'End Function
'Public Property Get TotalInvoiceDiscountAmountF() As Long
'    TotalInvoiceDiscountAmountF = Format(lngTotalInvoiceDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Total_2F() As String
'        Total_2F = Format(lngTotal2 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get TotalVATF() As String
'        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get TotalPayable() As String
'        TotalPayable = mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor
'End Property
'Public Property Get TotalPayableF() As String
'        TotalPayableF = Format(TotalPayable, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get TotalDiscount() As Long
'        TotalDiscount = mudtProps.TotalDiscount
'End Property
'Public Property Get TotalDiscountF() As String
'        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get TotalDepositF() As String
'        TotalDepositF = Format(TotalDeposit, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get TotalDeposit() As Long
'        TotalDeposit = lngTotalDeposit / oPC.Configuration.DefaultCurrency.Divisor
'End Property
'Public Property Get Total1() As Double
'    Total1 = lngTotal1
'End Property
'Public Property Get Total_1F() As String
'        Total_1F = Format(lngTotal1 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Total_3F() As String
'        Total_3F = Format(lngTotal3 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'Public Property Get Total_4F() As String
'        Total_4F = Format(lngTotal4 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'
'Public Property Get TotalBalanceAfterDeposits(bForeign As Boolean) As String
'        TotalBalanceAfterDeposits = Format(TotalPayable - TotalDeposit, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
Private Sub mcolLines_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.mcolLines_CurrRowStatus(pMsg)", pMsg
End Sub
'Public Property Get ForeignCurrency() As a_Currency
'    Set ForeignCurrency = oCurrency_Foreign
'End Property
Private Sub mcolLines_RowsChange()
    On Error GoTo errHandler
    CalculateTotals
    RaiseEvent TotalChange(lngLPLessDiscExt_tot, Me.TotalLessDiscExtF(False), lngVAT_tot, Me.TotalVATF(False))
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.mcolLines_RowsChange"
End Sub
'Public Function CalculateTotal()
'Dim oIL As a_InvoiceLine
'Dim bLineDiscountApplied As Boolean
'    bLineDiscountApplied = False
'    lngTotal1 = 0
'    lngTotal2 = 0
'    lngTotal3 = 0
'    lngTotal4 = 0
'    lngTotalDeposit = 0
'    lngTotalVAT = 0
'  '  lngTotalExtensionForDiscount = 0
'    lngTotalDiscount = 0
'    'NOTE Some items on the invoice might be ServiceItem e.g. Insurance, Postage and so on. These items are non subject to any discount
'    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
'    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
'    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
'    'calculated afterwards. Remember VAT rate can vary per line.
'    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
'    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
'    iExtrasCount = 0
'    For Each oIL In mcolLines
'        If oIL.BottomOfDocument Then
'            iExtrasCount = iExtrasCount + 1
'        End If
'        lngTotalVAT = lngTotalVAT + oIL.VATAmount
'        lngTotalDiscount = lngTotalDiscount + oIL.DiscountAmount
'        lngTotal3 = lngTotal3 + oIL.Extension
'        lngTotal4 = lngTotal4 + oIL.ExtensionInclDeposit
'        If oIL.BottomOfDocument Then
'            lngTotalExtras = lngTotalExtras + oIL.Extension
'        Else
'            lngTotal2 = lngTotal2 + oIL.Extension
'            lngTotalDeposit = lngTotalDeposit + oIL.Deposit
'            lngTotal1 = lngTotal1 + oIL.ExtensionSimple
'        End If
'    Next
'    mudtProps.TotalDiscount = lngTotalDiscount
'    mudtProps.TotalVAT = lngTotalVAT
'    mudtProps.TotalPayable = lngTotal4
'    If (Not Vatable) And ShowVAT Then
'        mudtProps.TotalPayable = lngTotal4 - lngTotalVAT
'    End If
'End Function
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As CNData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.GetState"
End Function
Private Sub SetState(Buffer As String)
    On Error GoTo errHandler
  Dim udtData As CNData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetState(buffer)", Buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "CNLS", mcolLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.GetSuperState"
End Function
Public Sub SetSuperState(Buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = Buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mcolLines.SetSuperState .ReadProperty("CNLS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetSuperState(buffer)", Buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolLines.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mcolLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
 ' Debug.Print "CancelEdit " & mcolStack.Count
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.CancelEdit"
End Sub

Public Sub ApplyEdit()
    On Error GoTo errHandler
Dim objPersist As a_CN_P
Dim OpenResult As Integer
 '   CalculateTotal
    mudtProps.TPID = Me.Customer.ID
    'Set these values fromthe objects before saving -  the objects cant be saved
    If Not oBillToAddress Is Nothing Then
        mudtProps.BillToAddressID = oBillToAddress.ID
    End If
    If Not oDelToAddress Is Nothing Then
        mudtProps.DelToAddressID = oDelToAddress.ID
    End If
    
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
'''''''''''''''''''''''''''''''''
  '  oPC.COShort.BeginTrans
    
    If DOCCode = "" Then
        DOCCode = GetCNNumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
    
    If mcolStack.Count = 0 Then Err.Raise 445
    mcolLines.ApplyEdit
    Set objPersist = New a_CN_P
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.TRID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If mudtProps.IsNew Then mudtProps.CaptureDate = Now()
        If Not IsValid Then Err.Raise 445
        mcolStack.Remove mcolStack.Count
        SetSuperState objPersist.Save(GetSuperState)
        mudtProps.IsNew = False
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    mudtProps.IsDirty = False
    
   ' oPC.COShort.CommitTrans
''''''''''''''''''''''''''''''''''
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ApplyEdit"  ', , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Sub
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mcolLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IsDirty"
End Property
Public Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetDirty(pVal)", pVal
End Sub

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IsValid"
End Property
Public Function GetSTatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.GetSTatus"
End Function
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mcolLines = New ch_CNL
    mcolLines.component Me
    Set mCustomer = New a_Customer

    Set oPCL = New z_PrintClient
    mudtProps.IsNew = True
'    mudtProps.VATRate = oPC.Configuration.VATRate
'    mudtProps.Vatable = True
'    mudtProps.ShowVAT = True
 '   mudtProps.DiscountRate = 0
    mOriginalStatus = stInProcess
    Set oBillToAddress = New a_Address
    Set oDelToAddress = New a_Address
    mobjValid.BreakRule "TP", True
    LoadClassorsCollection
 '   Set oCurrency_Foreign = oPC.Configuration.DefaultCurrency
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
      
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mcolLines = Nothing
    Set mCustomer = Nothing
    Set oPCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(ors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(ors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.mobjValid_BrokenRule(ors)", ors
End Sub
Private Sub mobjValid_RuleUnbroken(ors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(ors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.mobjValid_RuleUnbroken(ors)", ors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_CN_P
Dim oADd As a_Address
Dim oComp As a_Company
Dim oCurr As a_Currency
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_CN_P
    SetSuperState objPersist.Fetch(TRID)
    mOriginalStatus = Me.STATUS
    Set oBillingCompany = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID)
    mCustomer.Load mudtProps.TPID
    Set oBillToAddress = Nothing
    Set oDelToAddress = Nothing
    '
    If BillToAddressID > 0 Then
        Set oBillToAddress = Customer.Addresses.FindByID(BillToAddressID)
    End If
        'The address might not be found if it has been removed or the customer has been merged with another
    If oBillToAddress Is Nothing Then
        If Not Customer.BillTOAddress Is Nothing Then
            Set oBillToAddress = Customer.BillTOAddress
            BillToAddressID = Customer.BillTOAddress.ID
        End If
    End If

    If GoodsAddressID > 0 Then
        Set oDelToAddress = Customer.Addresses.FindByID(GoodsAddressID)
        'The address might not be found if it has been removed or the customer has been merged with another
    End If
    If oDelToAddress Is Nothing Then
        If Not Customer.DelToAddress Is Nothing Then
            Set oDelToAddress = Customer.DelToAddress
            GoodsAddressID = Customer.DelToAddress.ID
        End If
    End If

    
'    For Each oAdd In Customer.Addresses
'        If oAdd.ID = mudtProps.BillToAddressID Then
'            Set oBillToAddress = oAdd
'        End If
'        If oAdd.ID = mudtProps.DelToAddressID Then
'            Set oDelToAddress = oAdd
'        End If
'    Next
    Set objPersist = Nothing
    mobjValid.RuleBroken "TP", False

    mcolLines_RowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub
Public Property Get Customer() As a_Customer
    On Error GoTo errHandler
    Set Customer = mCustomer
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Customer"
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Delete"
End Sub



Public Property Get TRID() As Long
    On Error GoTo errHandler
    TRID = mudtProps.TRID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TRID"
End Property
Public Property Let TRID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TRID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TRID(val)", val
End Property

Public Property Get BillToAddressID() As Long
    On Error GoTo errHandler
    BillToAddressID = mudtProps.BillToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.BillToAddressID:Get"
End Property
Public Property Let BillToAddressID(val As Long)
    On Error GoTo errHandler
    mudtProps.BillToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.BillToAddressID:Let"
End Property
Public Sub SetBillToAddress(val As a_Address)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    Set oBillToAddress = val
    mudtProps.BillToAddressID = oBillToAddress.ID
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetBillToAddress(val)", val
End Sub
Public Property Get GoodsAddressID() As Long
    On Error GoTo errHandler
    GoodsAddressID = mudtProps.DelToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.GoodsAddressID:Get"
End Property
Public Property Let GoodsAddressID(val As Long)
    On Error GoTo errHandler
    mudtProps.DelToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.GoodsAddressID:Let"
End Property
Public Sub setDelToAddress(val As a_Address)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    Set oDelToAddress = val
    mudtProps.DelToAddressID = oDelToAddress.ID
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.setDelToAddress(val)", val
End Sub
'Public Function SetGeneralDiscount(Val As String) As Boolean
'Dim dblDiscountPercent As Double
'Dim oIL As a_InvoiceLine
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetGeneralDiscount = True
'    If Trim$(Val) = "" Then
'        dblDiscountPercent = 0
'    ElseIf Not ConvertToDBL(Val, dblDiscountPercent) Then
'        SetGeneralDiscount = False
'        Exit Function
'    End If
'    If (dblDiscountPercent <= 0) Then
'        mobjValid.RuleBroken "DiscountPercent", True
'    Else
'        mobjValid.RuleBroken "DiscountPercent", False
'    End If
'    mudtProps.DiscountRate = dblDiscountPercent
'    SetDirty True
'    For Each oIL In InvoiceLines  'disctribute discount rate over all lines where non nonstaock items exists
'        If Not oIL.ServiceItem Then
'            oIL.DiscountPercent = dblDiscountPercent
'        End If
'    Next
'    mcolLines_RowsChange  'recalulates andd raises event
'    RaiseEvent reloadlist
'End Function
Public Function SetCustomer(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
        Set oBillToAddress = Me.Customer.BillTOAddress
        Set oDelToAddress = Me.Customer.DelToAddress
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetCustomer(pTPID)", pTPID
End Function
Public Function SetCustomerFromAccNum(pAccnum As String) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean

    bSuccess = Me.Customer.Load(, pAccnum)
    SetCustomerFromAccNum = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
        If Me.BillToAddressID = 0 Then
            Set oBillToAddress = Me.Customer.BillTOAddress
        End If
        If Me.GoodsAddressID = 0 Then
            Set oDelToAddress = Me.Customer.DelToAddress
        End If
    End If
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetCustomerFromAccNum(pAccnum)", pAccnum
End Function
'Public Property Get TPID() As Long
'    TPID = mudtProps.TPID
'End Property
'Public Property Let TPID(Val As Long)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.TPID = Val
'    mobjValid.RuleBroken "TP", False
'    mudtProps.IsDirty = True
'End Property
'Public Property Let ShowVAT(Val As Boolean)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.ShowVAT = Val
'    SetDirty True
'End Property
'Public Property Get ShowVAT() As Boolean
'    ShowVAT = mudtProps.ShowVAT
'End Property
Public Property Get StaffName() As String
    On Error GoTo errHandler
    If mudtProps.StaffID < 1 Then
        StaffName = ""
        Exit Property
    End If
    StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).StaffName
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.StaffName"
End Property
Public Property Get StaffNameB() As String
    On Error GoTo errHandler
Dim strT As String
    If mudtProps.StaffID < 1 Then
        StaffNameB = ""
        Exit Property
    End If
    strT = oPC.Configuration.Staff.FindStaffByID(StaffID).Shortname
    If strT > "" Then
        StaffNameB = "   (" & strT & ")"
    Else
        StaffNameB = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.StaffNameB"
End Property

Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.StaffID(val)", val
End Property

'Public Property Get CurrencyID_Foreign() As Long
'    CurrencyID_Foreign = mudtProps.CurrencyID_Foreign
'End Property
'Public Property Let CurrencyID(Val As Long)
'Dim oCurr As a_Currency
'Dim oIL As a_InvoiceLine
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.CurrencyID_Foreign = Val
'    SetDirty True
'    For Each oCurr In oPC.Configuration.Currencies
'        If oCurr.ID = mudtProps.CurrencyID_Foreign Then
'            Set oCurrency_Foreign = oCurr
'        End If
'    Next
'    Me.CurrencyFactor = oCurrency_Foreign.Factor
'    For Each oIL In InvoiceLines
'        oIL.SetForeignPriceAndDeposit
'    Next
'    CalculateTotal_Foreign
'End Property
'Public Property Get COMPID() As Long
'    COMPID = mudtProps.COMPID
'End Property
'Public Property Let COMPID(Val As Long)
'Dim oComp As a_Company
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.COMPID = Val
'    SetDirty True
'    For Each oComp In oPC.Configuration.Companies
'        If oComp.ID = mudtProps.COMPID Then
'            Set oBillingCompany = oComp
'        End If
'    Next
'End Property
Public Property Get BillingCompany() As a_Company
    Set BillingCompany = oBillingCompany
End Property
'Public Property Let BillingCompsny(val As a_Company)
'    Set oBillingCompany
'End Property
Public Property Get TPName() As String
    On Error GoTo errHandler
    TPName = FNS(mudtProps.TPName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPName"
End Property
Public Property Let TPName(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPName = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPName(val)", val
End Property

Public Property Get TPPhone() As String
    On Error GoTo errHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPPhone"
End Property
Public Property Let TPPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
    On Error GoTo errHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPFax"
End Property
Public Property Let TPFax(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPFax(val)", val
End Property

Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPACCNum(val)", val
End Property

Public Property Get TPMemo() As String
    On Error GoTo errHandler
    TPMemo = FNS(mudtProps.TPMemo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPMemo"
End Property
Public Property Let TPMemo(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPMemo = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TPMemo(val)", val
End Property

Public Property Get BusPhone() As String
    On Error GoTo errHandler
    BusPhone = FNS(mudtProps.BusPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.BusPhone"
End Property
Public Property Let BusPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BusPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.BusPhone(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DocCode"
End Property
Public Property Get DOCCodeF() As String
    On Error GoTo errHandler
    DOCCodeF = "TAX CREDIT NOTE  " & FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DOCCodeF"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DocCode(val)", val
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = mudtProps.DOCDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DocDateF"
End Property
Public Property Let DOCDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsDate(val) Then
        mudtProps.DOCDate = CDate(val)
        SetDirty True
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DocDate(val)", val
End Property

Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IssDate"
End Property
Public Property Get IssDateF() As String
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IssDateF"
End Property
Public Property Let IssDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.IssDate(val)", val
End Property

Public Property Get VATable() As Boolean
    On Error GoTo errHandler
  '  Vatable = mudtProps.Vatable
  VATable = Me.Customer.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.VATable(val)", val
End Property
Public Property Let ShowVAT(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowVAT = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ShowVAT(val)", val
End Property
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ShowVAT"
End Property

Public Property Get CNLines() As ch_CNL
  Set CNLines = mcolLines
End Property
Public Property Get DisplayCN() As d_CN
    On Error GoTo errHandler
    Set DisplayCN = New d_CN
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DisplayCN"
End Property
Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.STATUS = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.STATUS = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.STATUS = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.STATUS = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.STATUS = stPROFORMA Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.StatusF"
End Property
Public Property Get STATUS() As enStatus
    On Error GoTo errHandler
    STATUS = mudtProps.STATUS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Status"
End Property
Public Function GetCNNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim strDefaultCo As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_GetNextCNNo"
    cmd.commandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , COMPID)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@Number", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    cmd.execute
    strDefaultCo = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyCode
    GetCNNumber = strDefaultCo & "CN" & cmd.Parameters("@Number").Value
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.GetCNNumber(strTmp)", strTmp
End Function

Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property
Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property


Private Function TranslateErrors(pRawors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, iStart As Integer
    iMarker = 1
    strAllRules = ""
    If Len(pRawors) > 0 Then
        iMarker = InStr(iMarker + 1, pRawors, ",")
        If iMarker > 0 Then
            strAllRules = colClassors(Left(pRawors, iMarker - 1))
        Else
            strAllRules = colClassors(pRawors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawors, ",")
            If iMarker > 0 Then
                strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
            Else
                strRule = colClassors(Mid(pRawors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "Numeric overflow, value captured too large", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If

    mobjValid.GetSTatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function

Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
    On Error GoTo errHandler
Dim strLabels As String
Dim strValues As String
Dim i As Integer
    strLabels = ""
    strValues = ""
    If Me.VATable Then ''''''''''(1)
        strLabels = strLabels & Chr(13) & "Total"
        strValues = strValues & Chr(13) & TotalPayableF(bForeign)
        strLabels = strLabels & Chr(13) & "Includes VAT of"
        strValues = strValues & Chr(13) & TotalVATF(bForeign)
    Else  'overseas clients
        If Me.ShowVAT Then
            strLabels = strLabels & Chr(13) & "subtotal"
            strValues = strValues & Chr(13) & TotalLessDiscExtF(bForeign)
            strLabels = strLabels & Chr(13) & "Less VAT of"
            strValues = strValues & Chr(13) & TotalVATF(bForeign)
            strLabels = strLabels & Chr(13) & "Payable"
            strValues = strValues & Chr(13) & TotalPayableF(bForeign)
        Else
            strLabels = strLabels & Chr(13) & "Total"
            strValues = strValues & Chr(13) & TotalPayableF(bForeign)
        End If
    End If
    pLabels = strLabels
    pValues = strValues
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function

Public Function PrintCN(Optional pQtyCopies As Integer) As Boolean
    On Error GoTo errHandler
Dim oCNL As a_CNL
Dim bDiscountExists As Boolean
Dim strTag As String
Dim bForeign As Boolean
Dim bReferenceExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim oDC As a_DocumentControl

    bForeign = False
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If
   
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintCN = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        strTag = "CN1_"
        If oPCL.StartRequestPrinting(strTag & Me.DOCCode) = False Then
            PrintCN = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, IIf(pQtyCopies > 0, pQtyCopies, .QtyCopies)
        End With
        PrintCN = True
    End If
   
   
    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
    bQtyColumnNeeded = False
    For Each oCNL In Me.CNLines
        If oCNL.Qty > 1 Or oPC.Configuration.AntiquarianYN = False Then
            bQtyColumnNeeded = True
        End If
        If oCNL.Discount > 0 Then
            bDiscountExists = True
        End If
        If oCNL.INVLineCode > "" Then
            bReferenceExists = True
        End If
    Next
    oPCL.SendPrintingSettings oPC.Configuration.PrintingSettings
    oPCL.HideDiscountColumn Not bDiscountExists
    oPCL.HideReferenceColumn Not bReferenceExists
    oPCL.HideDepositColumn Not bDepositExists
    oPCL.HideQtyColumn Not bQtyColumnNeeded

    
    If Me.StatusF <> "ISSUED" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.LineCOunt Me.CNLines.Count
    oPCL.LineTotalCount 4
    oPCL.WriteTag "OURADDRESS", BillingCompany.StreetAddressDelimited
    oPCL.WriteTag "PHONE", IIf(Me.TPPhone > "", "Phone: " & Me.TPPhone, "")
    If Me.COMPID > 0 Then
        oPCL.WriteTag "COREGNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CoRegistrationNumber
        oPCL.WriteTag "VATNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).VatNumber
    End If
    oPCL.WriteTag "CUSTOMER", CStr(Trim(Me.Customer.Fullname)) 'CStr(Trim(Me.BillTOAddress.AddressDelimited))
    If Not Me.BillingCompany Is Nothing Then
        oPCL.WriteTag "CONAME", BillingCompany.CompanyName
    Else
        oPCL.WriteTag "CONAME", ""
    End If
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "FAX", IIf(Me.TPFax > "", "Fax: " & Me.TPFax, "")
    oPCL.WriteTag "DATE", Me.DocDateF
    oPCL.WriteTag "MEMO", PackText(TPMemo)
    oPCL.WriteTag "BILLTO", PackText(BillTOAddress.AddressMailing)
    oPCL.WriteTag "DELTO", PackText(DelToAddress.AddressMailing)
    oPCL.WriteTag "CUSTVATNUMBER", Me.Customer.VatNumber

    oPCL.WriteTag "ACCNUMBER", IIf(Me.Customer.AcNo > "", "Acc Num: " & Me.Customer.AcNo, "")
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    FillTableRows Me, bForeign, bDiscountExists
    bDiscountExists = False
    For Each oCNL In CNLines
        If oCNL.Discount > 0 Then
            bDiscountExists = True
        End If
    Next
    oPCL.HideDiscountColumn (Not bDiscountExists)
    If Me.VATable Then ''''''''''(1)
        oPCL.WriteTag "TOTAL3", "Total credited " & "|" & TotalPayableF(bForeign)
        oPCL.WriteTag "TOTAL4", "Includes VAT of " & "|" & TotalVATF(bForeign)
    Else  'overseas clients
        If Me.ShowVAT Then
            oPCL.WriteTag "TOTAL", "subtotal" & "|" & TotalLessDiscExt(bForeign)  'subtotal 3
            oPCL.WriteTag "TOTAL2", "Less VAT of " & "|" & TotalVATF(bForeign)
            oPCL.WriteTag "TOTAL3", "Total credited " & "|" & TotalPayableF(bForeign)
        Else
            oPCL.WriteTag "TOTAL3", "Total credited " & "|" & TotalPayableF(bForeign)
        End If
    End If
    oPCL.FinishRequest

EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.PrintCN"
End Function
Private Function FillTableRows(pCN As a_CN, bForeign As Boolean, pDiscountExists As Boolean)
    On Error GoTo errHandler
Dim i As Long
Dim oCNL  As a_CNL

    For i = 1 To pCN.CNLines.Count
        Set oCNL = Me.CNLines(i)
            With oCNL
                If .Discount > 0 Then pDiscountExists = True
              '  If .ServiceItem Then
              '      oPCL.CNLSend "", "", .TitleAuthorPublisher & .SpecialVATMsg, .PriceF(bForeign), .ExtensionF(bForeign), "", "", ""
              '  Else
                    oPCL.CNLSend .ProductCodeF, .QtyF, .TitleAuthorPublisher, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DiscountPercentF, .PLessDiscExtNonVATF(bForeign), .INVLineCode, .Note
              '  End If
            End With
    Next i
    FillTableRows = 0
EXIT_Handler:
'ErrHandler:
'    MsgBox error
'    GoTo EXIT_Handler
   ' Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.FillTableRows(pCN,bForeign,pDiscountExists)", Array(pCN, bForeign, pDiscountExists)
End Function
Public Function PrintCN_Display()
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim bForeign As Boolean
Dim oCNL As a_CNL
Dim strCN As String
Dim i As Long

    bForeign = (oCurrency_Foreign Is oPC.Configuration.LocalCurrency)
    strCN = "Document No. " & Me.DOCCode & vbCrLf
    For i = 1 To Me.CNLines.Count
        Set oCNL = Me.CNLines(i)
        If Not oCNL Is Nothing Then
            With oCNL
                    strCN = strCN & .ProductCodeF & "  " & .QtyF & "  " & Left(.TitleAuthorPublisher, 30) & "  " & .PriceF(bForeign) & "  " & .PLessDiscExtF(bForeign)
            End With
            strCN = strCN & vbCrLf
        End If
    Next i
    
    strCN = strCN & vbCrLf
    If Me.VATable Then ''''''''''(1)
        strCN = strCN & "Total " & TotalPayableF(bForeign) & vbCrLf
        strCN = strCN & "Includes VAT of " & TotalVATF(bForeign) & vbCrLf
    Else  'overseas clients
        If Me.ShowVAT Then
            strCN = strCN & "subtotal    " & TotalLessDiscExt(bForeign) & vbCrLf   'subtotal 3
            strCN = strCN & "Less VAT discounted of    " & TotalVATF(bForeign) & vbCrLf
            strCN = strCN & "Total credited    " & TotalPayableF(bForeign) & vbCrLf
        Else
            strCN = strCN & "Total credited   " & TotalPayableF(bForeign) & vbCrLf
        End If
    End If
    PrintCN_Display = strCN
EXIT_Handler:
'ErrHandler:
'    oPCL.FinishRequest
'    Select Case
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox error
'        GoTo EXIT_Handler
'        Resume
'    End Select
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.PrintCN_Display"
End Function

Public Function Post() As String
    On Error GoTo errHandler
Dim oSM As z_StockManager
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    Me.ApplyEdit
    Set oSM = New z_StockManager
    
    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
            Select Case STATUS
            Case stVOID
       ''         SetStatus stVOID
            Case stInProcess
                Post = "In Process"
            Case stCOMPLETE
                oSM.IssueCreditNote Me, False, False
            Case stISSUED
                oSM.IssueCreditNote Me, False, False
            Case stPROFORMA
                oSM.IssueCreditNote Me, False, True
        '        SetStatus stPROFORMA
            End Select
    Case stCOMPLETE      'Complete
            Post = "ISSUED"
    Case stPROFORMA      'Issued PROFORMA
            Select Case STATUS
            Case stVOID
                oSM.CancelCreditNote Me
        '        SetStatus stVOID
            Case stInProcess
                Post = "OR: was proforma, can't now be inprocess."
            Case stCOMPLETE
                oSM.IssueCreditNote Me, True, False
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    Set oSM = Nothing
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Post", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.STATUS = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetStatus(val)", val
End Sub

Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.VoidDocument"
End Function
Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.TRID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.Reload"
End Sub
Private Property Get StatusForPrinting() As String
    On Error GoTo errHandler
    If mudtProps.STATUS = stVOID Then
        StatusForPrinting = "VOID"
    ElseIf mudtProps.STATUS = stInProcess Then
        StatusForPrinting = "IN PROCESS"
    ElseIf mudtProps.STATUS = stISSUED Then
        StatusForPrinting = ""
    ElseIf mudtProps.STATUS = stCOMPLETE Then
        StatusForPrinting = ""
    ElseIf mudtProps.STATUS = stPROFORMA Then
        StatusForPrinting = "PROFORMA"
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusForPrinting = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.StatusForPrinting"
End Property

Public Function ExportToXML(Optional pQtyCopies As Integer, Optional pFilename As String, _
    Optional DispatchMode As enTransmitType, Optional pWholeMessage As String, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
Dim bForeign As Boolean
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strPDFFile As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim bReferenceExists As Boolean
Dim bDiscountExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim oCNL As a_CNL
Dim oDC As a_DocumentControl
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strTemplateName As String
Dim strExecutable As String
    
    If DispatchMode = enMail Then
        If oPC.UsesOutlookForINVEmail Then
            strWorkingFolder = oPC.LocalFolder & "TEMP\"
            strDispatchMethod = "Outlook"
        Else
            strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
            strDispatchMethod = "Emailing"
        End If
    ElseIf DispatchMode = enPrint Then
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Then
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
        strDispatchMethod = ""
    Else
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If
    
    
    bForeign = False
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If
    
    p 1

    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
    bQtyColumnNeeded = False
    For Each oCNL In CNLines
        If oCNL.INVLineCode > " " Then
            bReferenceExists = True
        End If
        If oCNL.Discount > 0 Then
            bDiscountExists = True
        End If
        If oCNL.Qty > 1 Or oPC.Configuration.AntiquarianYN = False Then
            bQtyColumnNeeded = True
        End If
    Next


    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        ExportToXML = False
        Exit Function
    End If
    p 2
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "CN_DOC"
            .chCreate "MessageType"
                .elText = "CREDITNOTE"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "DestinationAddress"
                If oPC.TestMode Then
                    .elText = oPC.EmailFrom
                Else
                    If Customer.DelToAddress Is Nothing Then
                        .elText = ""
                    Else
                        .elText = Customer.DelToAddress.EMail
                    End If
                End If
            .elCreateSibling "TemplateName"
                .elText = "CN_DOC"
            strTemplateName = .elText
            
            .elCreateSibling "Columns"
            If bReferenceExists And bDiscountExists Then
                    .elText = "DISC_REF"
            ElseIf bDiscountExists Then
                    .elText = "DISC"
            ElseIf bReferenceExists Then
                    .elText = "REF"
            Else
                    .elText = ""
            End If
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailInvMsg
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = DOCCodeF
            .elCreateSibling "DocDate", True
                .elText = DocDateF
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerName", True
                .elText = Customer.NameAndCode(35)
            .elCreateSibling "CustomerAcno", True
                .elText = Customer.AcNo
            .elCreateSibling "CustomerVATNumber", True
            If Customer.VatNumber > "" Then
                .elText = "Customer's VAT number: " & Customer.VatNumber
            Else
                .elText = ""
            End If
            .elCreateSibling "CustomerWithAddress", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Me.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
            .elCreateSibling "CustomerPhone", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.BillTOAddress.Phone > "", "Phone: " & Customer.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "CustomerFax", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.BillTOAddress.Fax > "", "Fax: " & Customer.BillTOAddress.Fax, "")
                End If
            .elCreateSibling "BillTo", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
            .elCreateSibling "DelTo", True
                If Me.DelToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
                
            For i = 1 To CNLines.Count
     '          If CNLines(i).BottomOfDocument = False Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    .elText = CNLines(i).ProductCodeF
                    .elCreateSibling "Title", True
                        .elText = CNLines(i).Title
                    .elCreateSibling "QtyFirm", True
                        .elText = CNLines(i).QtyF
                    .elCreateSibling "Price", True
                        .elText = CNLines(i).PriceF(bForeign)
                    .elCreateSibling "PriceExVat", True
                        .elText = CNLines(i).PriceExVatF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = CNLines(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = CNLines(i).INVLineCode
                    .elCreateSibling "Extension", True
                        .elText = CNLines(i).PLessDiscExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = CNLines(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = CNLines(i).PLessDiscExtNonVATF(bForeign)
                    .elCreateSibling "VAT", True
                        .elText = CNLines(i).PLessDiscExtVATF(bForeign)
                    .elCreateSibling "VATRate", True
                        .elText = CNLines(i).VATRateF
                        
                    .navUP
      '          End If
            Next i
            
'        SetupDocumentTotals bForeign
'
'            For i = 1 To UBound(arDocTotalsF)
'                    If arDocTotalsF(i).Description = "" Then Exit For
'                    .elCreateSibling "Totals", True
'                    .chCreate "TotalDescription"
'                        .elText = arDocTotalsF(i).Description
'                    .elCreateSibling "NumericValue", True
'                        .elText = arDocTotalsF(i).AmtFormatted
'                    .navUP
'            Next i
            
            .elCreateSibling "QtyItems", True
                .elText = Me.TotalQtyF
                
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
                .elText = IIf(Me.StaffName > "", Me.StaffName, "")
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.OrderText
                
            .elCreateSibling "TotalIncVAT", True
                .elText = Me.TotalPayableF(False)
            .elCreateSibling "TotalExVAT", True
                .elText = Me.TotalNonVATF(False)
            .elCreateSibling "TotalVAT", True
                .elText = Me.TotalVATF(False)
                
            .elCreateSibling "VATable", True
                .elText = IIf(Me.VATable, "TRUE", "FALSE")
            .elCreateSibling "ShowVAT", True
                .elText = IIf(Me.ShowVAT, "TRUE", "FALSE")
                
    End With
    p 3
'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "CN1_" & Me.DOCCode & ".xml"
    p 4, strXML
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With
    p 5
    If DispatchMode = enMail Then
        'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & fs.GetBaseName(strTemplateName) & "_HTML.xslt"
            Set fs = New FileSystemObject
            p 6, strPath
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
            p 7, strWorkingFolder & "CN_" & Me.DOCCode & ".HTML"
            If fs.FileExists(strWorkingFolder & "CN_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "CN_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "CN_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage
            oTF.CloseTextFile
    End If
p 8
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If (Me.Customer.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
p 9
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
        p 10, strPath
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
p 11
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        strFOFile = strWorkingFolder & "CN_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "CN_" & Me.DOCCode & ".PDF"
p 12, strFOFile
p 13, strPDFFile
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
'        strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
        ChDir "\PBKS\Executables\FOP\"
       ' strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
        'ShellandWait strCommand, 100
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000

        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If

    Set oSM = New z_StockManager
    Set oFSo = New FileSystemObject
    pFilename = strPDFFile
    
    If DispatchMode <> enMail And DispatchMode <> enView Then
        oSM.LogTransmission Me.TRID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    End If
    Set oSM = Nothing
    Set oFSo = Nothing
                        p 14
    ExportToXML = True
    '
'2630      If DispatchMode = enView Then
'2640          strExecutable = GetPDFExecutable(strPDFFile)
'                If strExecutable = "" Then
'                    MsgBox "There is no application set on this computer to open the file: " & strPDFFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
'                Else
'                    Shell strExecutable & " " & strPDFFile, vbNormalFocus
'                End If
'2660      End If
    If (Not fs.FileExists(strPDFFile)) And (DispatchMode = enView Or DispatchMode = enMail) Then
      MsgBox "The PDF file has not been produced: " & strPDFFile & ". Please contact support.", vbOKOnly, "Can't do this"
    Else
  If DispatchMode = enView Or DispatchMode = enPrint Then
      OpenFileWithApplication strPDFFile, enPDF, QuickPrint
  End If
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ExportToXML(bForeign)", bForeign, , , "line number", Array(Erl())
End Function

Public Function ExportToSpreadsheet(bForeign As Boolean, _
        pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim s As String
Dim s2 As String
Dim lngNumberOfLines As Long

    ExportToSpreadsheet = False
    pFilename = oPC.LocalFolder & "Temp\" & DOCCode & "_" & Format(Now(), "yyyymmddHHnn") & ".xls"
    
    oTF.OpenTextFile pFilename
    oTF.WriteToTextFile "Status" & vbTab & Me.STATUS
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCodeF
    oTF.WriteToTextFile "DocDate" & vbTab & DocDateF
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "CustomerName" & vbTab & Customer.NameAndCode(70)
    oTF.WriteToTextFile "ACNO" & vbTab & IIf(Me.Customer.AcNo > "", "Ac/no. " & Me.Customer.AcNo, "")
    oTF.WriteToTextFile "CustomerVATNumber" & vbTab & Customer.VatNumber
    oTF.WriteToTextFile_NoLineTerminator "CustomerWithAddress" & vbTab
    If Not Customer.BillTOAddress Is Nothing Then
        oTF.WriteToTextFile Replace(Customer.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
    End If
    
    oTF.WriteToTextFile_NoLineTerminator "CustomerPhone" & vbTab
    If Not Customer.BillTOAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Me.Customer.BillTOAddress.Phone > "", "Phone: " & Customer.BillTOAddress.Phone, "")
    End If
    oTF.WriteToTextFile_NoLineTerminator "CustomerFax" & vbTab
    If Not Customer.BillTOAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Customer.BillTOAddress.Fax > "", "Fax: " & Customer.BillTOAddress.Fax, "")
    End If

  If Not BillTOAddress Is Nothing Then
            s = Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
  Else
            s = ""
  End If
    oTF.WriteToTextFile "BillTo" & vbTab & s

  If Not DelToAddress Is Nothing Then
            s = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
  Else
            s = ""
  End If
    oTF.WriteToTextFile "DelTo" & vbTab & s

    s = "SKU" & vbTab & "Title" & vbTab & "Qty" & vbTab & "Price" & vbTab & "PriceExVat" & vbTab & "DiscountRate" & vbTab & "InvLineCode" & vbTab _
    & "Extension" & vbTab & "Note" & vbTab & "ExtensionExVat" & vbTab & "ExtensionVATOnly" & vbTab & "VATRate"
    
    oTF.WriteToTextFile s
               
                
  lngNumberOfLines = 0
  For i = 1 To CNLines.Count
    lngNumberOfLines = lngNumberOfLines + 1
    s = CNLines(i).ProductCodeF
    s = s & vbTab & CNLines(i).Title
    s = s & vbTab & CNLines(i).QtyF
    s = s & vbTab & CNLines(i).PriceF(bForeign)
    s = s & vbTab & CNLines(i).PriceExVatF(bForeign)
    s = s & vbTab & CNLines(i).DiscountPercentF
    s = s & vbTab & CNLines(i).INVLineCode
    s = s & vbTab & CNLines(i).PLessDiscExtF(bForeign)
    s = s & vbTab & CNLines(i).Note
    s = s & vbTab & CNLines(i).PLessDiscExtNonVATF(bForeign)
    s = s & vbTab & CNLines(i).PLessDiscExtVATF(bForeign)
    s = s & vbTab & CNLines(i).VATRateF
    oTF.WriteToTextFile s
Next

    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalPayableF(False)
    oTF.WriteToTextFile "TotalExVAT" & vbTab & TotalNonVATF(False)
    oTF.WriteToTextFile "TotalVAT" & vbTab & TotalVATF(False)
    oTF.WriteToTextFile "VATable" & vbTab & IIf(Me.VATable, "TRUE", "FALSE")
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
End Function


Public Property Get Log() As String
    On Error GoTo errHandler
    Log = Trim(mudtProps.Log)
    Exit Property
errHandler:
    ErrorIn "a_Invoice.Log"
End Property

Public Property Let Log(val As String)
    On Error GoTo errHandler
    mudtProps.Log = Trim(val)
    Exit Property
errHandler:
    ErrorIn "a_Invoice.Log(val)", val
End Property

Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .Open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "a_CN.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub


Private Sub SetupDocumentTotals(bForeign As Boolean)
    On Error GoTo errHandler
Dim i, k As Integer
Dim bBottomOfDocument As Boolean

    ReDim arDocTotals(16)
    ReDim arDocTotalsF(16)
    
    jTotal = 1
    RunningTotal = 0
    
 '   If Not Customer.VATable Then
 '   Else
 '       If Not ShowVAT Then
        addARow "Total", Me.TotalLessDiscExt(bForeign), Me.TotalLessDiscExtF(bForeign), "", "+"
 '   Else
 '       addARow "subtotal", Me.TotalLessDiscExt(bForeign), Me.TotalLessDiscExtF(bForeign), "", "+"
 '   End If
    
    If Customer.VATable Then
        addARow "(includes VAT at " & oPC.Configuration.VATRateF, TotalVAT(bForeign), TotalVATF(bForeign), "", " "
    Else
        If ShowVAT Then
            addARow "less VAT at " & oPC.Configuration.VATRateF, TotalVAT(bForeign), TotalVATF(bForeign), "GRAND", "-"
        End If
    End If
    
'    TransferToFinalTable
    k = 0
    For i = 1 To UBound(arDocTotals)
        If arDocTotals(i).Description = "" Then Exit For
        k = k + 1
        arDocTotalsF(k).Description = arDocTotals(i).Description
        arDocTotalsF(k).AmtFormatted = arDocTotals(i).AmtFormatted
        If arDocTotals(i).TotalType > "" Then
            k = k + 1
            arDocTotalsF(k).Description = IIf(arDocTotals(i).TotalType = "GRAND", "Total", "subtotal")
            arDocTotalsF(k).val = CStr(arDocTotals(i).RunningTotal)
            If bForeign Then
                arDocTotalsF(k).AmtFormatted = Format(arDocTotals(i).RunningTotal / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
            Else
                arDocTotalsF(k).AmtFormatted = Format(arDocTotals(i).RunningTotal / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
            End If
        End If
    Next i
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.SetupDocumentTotals(bForeign)", bForeign
End Sub

Private Sub addARow(pDescription, Amt As Long, AmtF As String, pTotalType As String, Optional Sign As String)
    On Error GoTo errHandler
    arDocTotals(jTotal).Description = pDescription
    arDocTotals(jTotal).val = Amt
    arDocTotals(jTotal).AmtFormatted = AmtF & IIf(Left(pDescription, 1) = "(", ")", "")
    arDocTotals(jTotal).TotalType = pTotalType
    If Sign = "-" Then
        RunningTotal = RunningTotal - Amt
    ElseIf Sign = "+" Then
        RunningTotal = RunningTotal + Amt
    End If
    arDocTotals(jTotal).RunningTotal = RunningTotal
    jTotal = jTotal + 1
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_CN.addARow(pDescription,Amt,AmtF,pTotalType)", Array(pDescription, Amt, _
         AmtF, pTotalType)
End Sub

Public Sub BuildFromInvoice(pInv As a_Invoice)
Dim oIL As a_InvoiceLine
Dim oCNL As a_CNL

    Me.SetCustomer pInv.Customer.ID
    Me.SetStatus stInProcess
    Me.ShowVAT = pInv.ShowVAT
    Me.DOCDate = Now()
    Me.COMPID = pInv.COMPID
    Me.ForeignCurrencyID = pInv.ForeignCurrency.ID
    Me.CurrRate = pInv.CurrencyFactor
    For Each oIL In pInv.InvoiceLines
        If oIL.CNLQty > 0 Then  'Marked for CN generation
            Set oCNL = Me.CNLines.Add
            oCNL.BeginEdit
            oCNL.PID = oIL.PID
            oCNL.Qty = oIL.CNLQty
            oCNL.INVLineID = oIL.InvoiceLineID
            oCNL.INVLineCode = pInv.DOCCode & ":" & oIL.Ref
            oCNL.VATRate = oIL.VATRate
            oCNL.InvPrice = oIL.Price
            oCNL.Discount = oIL.DiscountPercent
            oCNL.CopyID = oIL.PIID
            oCNL.ApplyEdit
        End If
    Next
End Sub

