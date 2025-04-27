VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim lngLPLessDiscExt_tot As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLDeposit_tot As Long
Dim lngFDeposit_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Dim lngLPayableLessVAT_tot As Long
Dim lngFPayableLessVAT_tot As Long
Dim lngTotalExtras As Long
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Dim iExtrasCount As Integer
Private xMLDoc As ujXML
Event Valid(pMsg As String)
Event TotalChange(lngTotalExt As Long, lngDeposit_tot As Long, lngVAT_tot As Long)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Event RowCount(cnt As Integer)
Private mudtProps As InvoiceProps
Private mcolStack As Collection
Private cColsOSPerCust As c_COLsPerCust
Private cAPPLsOS As c_APPLsPerTPPID
Private WithEvents mColInvoiceLines As ch_InvoiceLine
Attribute mColInvoiceLines.VB_VarHelpID = -1
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private colClassErrors As Collection
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
Private oCurrency_Foreign As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Const CDOCCODE = "IN"

Type TTotal
    Description As String
    val As Long
    AmtFormatted As String
    TotalType As String
    RunningTotal As Long
    Sign As String
End Type
Dim arDocTotals() As TTotal
Dim arDocTotalsF() As TTotal
Dim RunningTotal As Long
Dim jTotal As Long

Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
Public Property Get HasCoffs() As Boolean
    On Error GoTo errHandler
Dim bFound As Boolean
Dim oIL As a_InvoiceLine

    bFound = False
    For Each oIL In InvoiceLines
        If oIL.COFFs.Count > 0 Then
            bFound = True
        End If
    Next
    HasCoffs = bFound
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.HasCoffs"
End Property

Private Sub mColInvoiceLines_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mColInvoiceLines_CurrRowStatus(pMsg)", pMsg
End Sub
Public Property Get ForeignCurrency() As a_Currency
    Set ForeignCurrency = oCurrency_Foreign
End Property
Private Sub mcolinvoiceLines_RowsChange()
    On Error GoTo errHandler
'occurs for new rows and deletion of rows
    
    CalculateTotals
    RaiseEvent TotalChange(lngLPLessDiscExt_tot, lngLDeposit_tot, lngVAT_tot)
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mcolinvoiceLines_RowsChange"
End Sub
Private Sub mcolInvoiceLines_RowCOunt(cnt As Integer)
    RaiseEvent RowCount(cnt)
End Sub
Public Function RecalculateAllLines()
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
    For Each oIL In InvoiceLines
        oIL.CalculateLine
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.RecalculateAllLines"
End Function
Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
Dim bLineDiscountApplied As Boolean
    mobjValid.RuleBroken "OVERFLOW", False
    bLineDiscountApplied = False
    lngLPLessDiscExt_tot = 0
    lngLDeposit_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngLPayableLessVAT_tot = 0
    'NOTE Some items on the invoice might be nonstock e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non nonstock items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    iExtrasCount = 0
    lngTotalExtras = 0
    For Each oIL In mColInvoiceLines
        If (Not oIL.IsDeleted) Then
    '    If oIL.Fulfilled <> "CAN" Then
            oIL.CalculateLine
            If oIL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            lngLPayableLessVAT_tot = lngLPayableLessVAT_tot + oIL.PLessDiscExtNonVAT(False)
            lngVAT_tot = lngVAT_tot + oIL.PLessDiscExtVAT(False)
            lngDisc_tot = lngDisc_tot + oIL.PDiscExt(False)
            lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + oIL.PLessDiscExt(False)
            lngLDeposit_tot = lngLDeposit_tot + oIL.Deposit
            If oIL.BottomOfDocument Then
                lngTotalExtras = lngTotalExtras + oIL.PLessDiscExt(False)
            End If
        End If
    Next
    mudtProps.TotalExtras = lngTotalExtras
    mudtProps.DepositPaid = lngLDeposit_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalNonVAT = lngLPayableLessVAT_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalExtension = lngLPLessDiscExt_tot
    mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngLDeposit_tot
    If (Not VATable) And ShowVAT Then
        mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngVAT_tot
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        ErrClear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim oIL As a_InvoiceLine
Dim bLineDiscountApplied As Boolean
    mobjValid.RuleBroken "OVERFLOW", False
    bLineDiscountApplied = False
    lngFPLessDiscExt_tot = 0
    lngFDeposit_tot = 0
    lngVAT_tot = 0
    lngFPayableLessVAT_tot = 0
    lngDisc_tot = 0
    lngTotalExtras = 0
'    For Each oIL In InvoiceLines
'        oIL.SetForeignPriceAndDeposit
'    Next
'
    For Each oIL In mColInvoiceLines
        If (Not oIL.IsDeleted) Then
            If oIL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            lngFPayableLessVAT_tot = lngFPayableLessVAT_tot + oIL.PLessDiscExtNonVAT(True)
            lngVAT_tot = lngVAT_tot + oIL.PLessDiscExtVAT(True)
            lngDisc_tot = lngDisc_tot + oIL.PDiscExt(True)
            lngFPLessDiscExt_tot = lngFPLessDiscExt_tot + oIL.PLessDiscExt(True)
                lngFDeposit_tot = lngFDeposit_tot + oIL.Deposit_Foreign
            If oIL.BottomOfDocument Then
                lngTotalExtras = lngTotalExtras + oIL.PLessDiscExt(True)
            End If
        End If
    Next
    mudtProps.TotalExtras_Foreign = lngTotalExtras
    mudtProps.TotalDiscount_Foreign = lngDisc_tot
    mudtProps.TotalVAT_Foreign = lngVAT_tot
    mudtProps.TotalNonVAT_Foreign = lngFPayableLessVAT_tot
    mudtProps.TotalExtension_Foreign = lngFPLessDiscExt_tot
    mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot - lngFDeposit_tot
    If (Not Me.VATable) And ShowVAT Then
        mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot - lngVAT_tot
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        ErrClear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CalculateTotal_Foreign"
End Function
'Public Property Get TotalInvoiceDiscountAmountF() As Long
'    TotalInvoiceDiscountAmountF = Format(lngTotalInvoiceDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
Public Property Get TotalExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalExtF = Format((mudtProps.TotalExtension_Foreign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalExtF = Format((mudtProps.TotalExtension) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalExtWithOutExtrasF(bForeign)", bForeign
End Property
'Public Property Get TotalExtLessVATF(bForeign As Boolean) As String
'    On Error GoTo errHandler
'    If bForeign Then
'        TotalExtF = Format((mudtProps.TotalExtension_Foreign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
'    Else
'        TotalExtF = Format((mudtProps.TotalExtension) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Invoice.TotalExtWithOutExtrasF(bForeign)", bForeign
'End Property

Public Property Get TotalExtWithOutExtras(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalExtWithOutExtras = mudtProps.TotalExtension_Foreign - mudtProps.TotalExtras_Foreign
    Else
        TotalExtWithOutExtras = mudtProps.TotalExtension - mudtProps.TotalExtras
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalExtWithOutExtras(bForeign)", bForeign
End Property
Public Property Get TotalExtWithOutExtrasF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalExtWithOutExtrasF = Format((TotalExtWithOutExtras(bForeign)) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalExtWithOutExtrasF = Format((TotalExtWithOutExtras(bForeign)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalExtWithOutExtrasF(bForeign)", bForeign
End Property
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
    ErrorIn "a_Invoice.TotalLessDiscExtF(bForeign)", bForeign
End Property
Public Property Get TotalLessDiscExt(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalLessDiscExt = mudtProps.TotalExtension_Foreign
    Else
        TotalLessDiscExt = mudtProps.TotalExtension
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalLessDiscExt(bForeign)", bForeign
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
    ErrorIn "a_Invoice.TotalNonVATF(bForeign)", bForeign
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
    ErrorIn "a_Invoice.TotalVATF(bForeign)", bForeign
End Property
Public Property Get TotalPayableRaw(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalPayableRaw = mudtProps.TotalPayable_Foreign
    Else
        TotalPayableRaw = mudtProps.TotalPayable
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalPayableRaw(bForeign)", bForeign
End Property
Public Property Get TotalVATRaw(bForeign As Boolean) As Long
    On Error GoTo errHandler
    If bForeign Then
        TotalVATRaw = mudtProps.TotalVAT_Foreign
    Else
        TotalVATRaw = mudtProps.TotalVAT
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalVATRaw(bForeign)", bForeign
End Property

Public Property Get TotalPayable(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayable = mudtProps.TotalPayable_Foreign '/ oCurrency_Foreign.Divisor
    Else
        TotalPayable = mudtProps.TotalPayable ' / oPC.Configuration.DefaultCurrency.Divisor
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalPayable(bForeign)", bForeign
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayableF = Format(TotalPayable(bForeign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalPayableF = Format(TotalPayable(bForeign) / oCurrency_Foreign.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalPayableF(bForeign)", bForeign
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
    ErrorIn "a_Invoice.TotalDiscount(bForeign)", bForeign
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
    ErrorIn "a_Invoice.TotalDiscountF(bForeign)", bForeign
End Property
Public Property Get TotalDepositF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalDepositF = Format(TotalDeposit(bForeign), oCurrency_Foreign.FormatString)
    Else
        TotalDepositF = Format(TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalDepositF(bForeign)", bForeign
End Property
Public Property Get TotalDeposit(bForeign As Boolean) As Double
    On Error GoTo errHandler
    If bForeign Then
        TotalDeposit = lngFDeposit_tot / oCurrency_Foreign.Divisor
    Else
        TotalDeposit = lngLDeposit_tot / oPC.Configuration.DefaultCurrency.Divisor
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalDeposit(bForeign)", bForeign
End Property
Public Property Get InvoiceDiscountRate() As Double
    On Error GoTo errHandler
    InvoiceDiscountRate = mudtProps.DiscountRate
  '  RecalculateAllLines
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.InvoiceDiscountRate"
End Property
Public Property Get InvoiceDiscountRateF() As String
    On Error GoTo errHandler
    If mudtProps.DiscountRate = 0 Then
        InvoiceDiscountRateF = ""
    Else
        InvoiceDiscountRateF = Format(mudtProps.DiscountRate, gPercentFormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.InvoiceDiscountRateF"
End Property
'Public Property Get Total1() As Double
'    Total1 = lngTotal1
'End Property
'Public Property Get Total_1F(bForeign As Boolean) As String
'    If bForeign Then
'        Total_1F = Format(lngTotal1 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
'    Else
'        Total_1F = Format(lngTotal1 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get Total_3F(bForeign As Boolean) As String
'    If bForeign Then
'        Total_3F = Format(lngTotal3 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
'    Else
'        Total_3F = Format(lngTotal3 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get Total_4F(bForeign As Boolean) As String
'    If bForeign Then
'        Total_4F = Format(lngTotal4 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
'    Else
'        Total_4F = Format(lngTotal4 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'
Public Property Get TotalBalanceAfterDeposits(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalBalanceAfterDeposits = Format(TotalPayable(bForeign) - TotalDeposit(bForeign), oCurrency_Foreign.FormatString)
    Else
        TotalBalanceAfterDeposits = Format(TotalPayable(bForeign) - TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalBalanceAfterDeposits(bForeign)", bForeign
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As InvoiceData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As InvoiceData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "ILS", mColInvoiceLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.GetSuperState"
End Function
Public Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mColInvoiceLines.SetSuperState .ReadProperty("ILS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
  mColInvoiceLines.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  mColInvoiceLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CancelEdit"
End Sub

Public Function ApplyEdit(Optional pMsg As String) As String
    On Error GoTo errHandler
Dim strTemp As String
Dim objPersist As a_Invoice_P
    ApplyEdit = ""
    CalculateTotals
    mudtProps.TPID = Customer.ID
    mudtProps.SalesRepID = Customer.REPID
    
    'Set these values fromthe objects before saving -  the objects cant be saved
    If Not oBillToAddress Is Nothing Then
        mudtProps.BillToAddressID = oBillToAddress.ID
    End If
    If Not oDelToAddress Is Nothing Then
        mudtProps.DelToAddressID = oDelToAddress.ID
    End If
    
'''''''''''''''''''''''''''
    oPC.CO.execute "BEGIN TRANSACTION"
'    oPC.CO.execute "set implicit_transactions off"
        
    If DOCCode = "" Then
        DOCCode = GetInvoiceNumber("GETANDWRITE")
        DOCDate = Now()
    End If
    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
    
    If mcolStack.Count = 0 Then Err.Raise 445
    oPC.CO.CommandTimeout = 15
    mColInvoiceLines.ApplyEdit
    Set objPersist = New a_Invoice_P
    
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.InvoiceID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If Not IsValid Then Err.Raise 445
        If mudtProps.IsNew Then mudtProps.CaptureDate = Now()
        strTemp = objPersist.Save(GetSuperState)
    
        If UCase(Left(strTemp, 15)) <> "TIMEOUT" Then
            mcolStack.Remove mcolStack.Count
            SetSuperState strTemp
            mudtProps.IsNew = False
        Else 'save failed
            mColInvoiceLines.BeginEdit
            If Not IsMissing(pMsg) Then
                pMsg = strTemp
            End If
            ApplyEdit = "TIMEOUT"
            oPC.CO.execute "ROLLBACK TRANSACTION"

            Set objPersist = Nothing
            GoTo EXIT_Handler
        End If
    
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    SetDirty False
    oPC.CO.execute "COMMIT TRANSACTION"
''''''''''''''''''''''''''''''''''''''''
    oPC.CO.CommandTimeout = 0
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ApplyEdit", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.CO
End Function
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mColInvoiceLines.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsValid"
End Property
Public Function GetSTatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
    RaiseEvent Dirty(mudtProps.IsDirty)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.GetSTatus"
End Function
Private Sub Class_Initialize()
Dim oComp As a_Company
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mColInvoiceLines = New ch_InvoiceLine
    mColInvoiceLines.Component Me
    Set mCustomer = New a_Customer

    Set oPCL = New z_PrintClient
    mudtProps.IsNew = True
    mudtProps.VATRate = oPC.Configuration.VATRate
    mudtProps.VATable = True
    mudtProps.ShowVAT = oPC.Configuration.DiscountVATDefault
    mudtProps.COMPID = oPC.Configuration.DefaultCOMPID
    For Each oComp In oPC.Configuration.Companies
        If oComp.ID = mudtProps.COMPID Then
            Set oBillingCompany = oComp
        End If
    Next

    mudtProps.DiscountRate = 0
    mudtProps.Status = stInProcess
    mOriginalStatus = stInProcess
    Set oBillToAddress = New a_Address
    Set oDelToAddress = New a_Address
    mobjValid.BreakRule "TP", True
    LoadClassErrorsCollection
    Set oCurrency_Foreign = oPC.Configuration.DefaultCurrency
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
      
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mColInvoiceLines = Nothing
    Set mCustomer = Nothing
    Set oPCL = Nothing
    Set cColsOSPerCust = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mobjValid_BrokenRule(errors)", errors
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mobjValid_RuleUnbroken(errors)", errors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_Invoice_P
Dim oADd As a_Address
Dim oComp As a_Company
Dim oCurr As a_Currency
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_Invoice_P
    SetSuperState objPersist.Fetch(TRID)
    mOriginalStatus = Me.Status
'''''''''''
'        Set cColsOSPerCust = Nothing
'        Set cColsOSPerCust = New c_COLsPerCust
'        cColsOSPerCust.Load mudtProps.TPID
'
    
    
    mCustomer.Load mudtProps.TPID
    
    If BillToAddressID > 0 Then
        Set oBillToAddress = Customer.Addresses.FindByID(BillToAddressID)
        'The address might not be found if it has been removed or the customer has been merged with another
        If oBillToAddress Is Nothing Then
            If Not Customer.BillTOAddress Is Nothing Then
                Set oBillToAddress = Customer.BillTOAddress
                BillToAddressID = Customer.BillTOAddress.ID
            End If
        End If
    End If
    If DelToAddressID > 0 Then
        Set oDelToAddress = Customer.Addresses.FindByID(DelToAddressID)
        'The address might not be found if it has been removed or the customer has been merged with another
        If oDelToAddress Is Nothing Then
            If Not Customer.DelToAddress Is Nothing Then
                Set oDelToAddress = Customer.DelToAddress
                DelToAddressID = Customer.DelToAddress.ID
            End If
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
    For Each oComp In oPC.Configuration.Companies
        If oComp.ID = mudtProps.COMPID Then
            Set oBillingCompany = oComp
        End If
    Next
    For Each oCurr In oPC.Configuration.Currencies
        If oCurr.ID = mudtProps.CurrencyID_Foreign Then
            Set oCurrency_Foreign = oCurr
        End If
    Next
    Set objPersist = Nothing
    mobjValid.RuleBroken "TP", False
 '   Me.RecalculateAllLines
    CalculateTotals
   ' mcolinvoiceLines_RowsChange
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub
Public Property Get Customer() As a_Customer
    On Error GoTo errHandler
    Set Customer = mCustomer
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Customer"
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Delete"
End Sub



Public Property Get InvoiceID() As Long
    On Error GoTo errHandler
    InvoiceID = mudtProps.InvoiceID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.InvoiceID"
End Property
Public Property Let InvoiceID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.InvoiceID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.InvoiceID(val)", val
End Property

Public Property Get BillToAddressID() As Long
    On Error GoTo errHandler
    BillToAddressID = mudtProps.BillToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.BillToAddressID:Get"
End Property
Public Property Let BillToAddressID(val As Long)
    On Error GoTo errHandler
    mudtProps.BillToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.BillToAddressID:Let"
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
    ErrorIn "a_Invoice.SetBillToAddress(val)", val
End Sub
Public Property Get DelToAddressID() As Long
    On Error GoTo errHandler
    DelToAddressID = mudtProps.DelToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DelToAddressID:Get"
End Property
Public Property Let DelToAddressID(val As Long)
    On Error GoTo errHandler
   mudtProps.DelToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DelToAddressID:Let"
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
    ErrorIn "a_Invoice.setDelToAddress(val)", val
End Sub
Public Function SetGeneralDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim dblDiscountPercent As Double
Dim oIL As a_InvoiceLine

    If mcolStack.Count = 0 Then Err.Raise 383
    SetGeneralDiscount = True
    If Trim$(val) = "" Then
        dblDiscountPercent = 0
    ElseIf Not ConvertToDBL(val, dblDiscountPercent) Then
        SetGeneralDiscount = False
        Exit Function
    End If
    If (dblDiscountPercent <= 0) Then
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    mudtProps.DiscountRate = dblDiscountPercent
    SetDirty True
    For Each oIL In InvoiceLines  'disctribute discount rate over all lines where non nonstaock items exists
        If Not oIL.NonStock Then
            oIL.DiscountPercent = dblDiscountPercent
        End If
    Next
    mcolinvoiceLines_RowsChange  'recalulates andd raises event
    RecalculateAllLines
    RaiseEvent reloadlist
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetGeneralDiscount(val)", val
End Function
Public Function LoadCOLsOS(Optional pTPID As Long, Optional pPID As String) As c_COLsPerCust
    Set cColsOSPerCust = Nothing
    Set cColsOSPerCust = New c_COLsPerCust
    cColsOSPerCust.Load Customer.ID, pPID
End Function
Public Function LoadAPPLsOS(Optional pTPID As Long, Optional pPID As String) As c_COLsPerCust
    Set cAPPLsOS = Nothing
    Set cAPPLsOS = New c_APPLsPerTPPID
    cAPPLsOS.Load Customer.ID, pPID, ""
End Function

Public Function SetCustomer(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    If bSuccess Then
        VATable = Customer.VATable
        mobjValid.RuleBroken "TP", False
        If Not Me.Customer.BillTOAddress Is Nothing Then SetBillToAddress Me.Customer.BillTOAddress
        If Not Me.Customer.DelToAddress Is Nothing Then setDelToAddress Me.Customer.DelToAddress
      '  SetBillToAddressID Me.Customer.BillTOAddress.ID
      '  SetDelToAddressID Me.Customer.DelToAddress.ID
        SetDirty True
'        MsgBox "HERE"
'        Set cColsOSPerCust = Nothing
'        Set cColsOSPerCust = New c_COLsPerCust
'        cColsOSPerCust.Load pTPID
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetCustomer(pTPID)", pTPID
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
        If Me.DelToAddressID = 0 Then
            Set oDelToAddress = Me.Customer.DelToAddress
        End If
    End If
    Me.VATable = Customer.VATable
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetCustomerFromAccNum(pAccnum)", pAccnum
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
Public Property Let ShowVAT(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowVAT = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ShowVAT(val)", val
End Property
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ShowVAT"
End Property
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.StaffID(val)", val
End Property
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
    ErrorIn "a_Invoice.StaffName"
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
    ErrorIn "a_Invoice.StaffNameB"
End Property

Public Property Get CurrencyID_Foreign() As Long
    On Error GoTo errHandler
    CurrencyID_Foreign = mudtProps.CurrencyID_Foreign
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CurrencyID_Foreign"
End Property
Public Property Let CurrencyID(val As Long)
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim oIL As a_InvoiceLine

    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CurrencyID_Foreign = val
    SetDirty True
    For Each oCurr In oPC.Configuration.Currencies
        If oCurr.ID = mudtProps.CurrencyID_Foreign Then
            Set oCurrency_Foreign = oCurr
        End If
    Next
    Me.CurrencyFactor = oCurrency_Foreign.Factor
    RecalculateAllLines
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CurrencyID(val)", val
End Property
Public Property Get COMPID() As Long
    On Error GoTo errHandler
    COMPID = mudtProps.COMPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.COMPID"
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
    ErrorIn "a_Invoice.COMPID(val)", val
End Property
Public Property Get BillingCompany() As a_Company
    Set BillingCompany = oBillingCompany
End Property

Public Property Get CustPaid() As Boolean
    CustPaid = mudtProps.CustPaid
End Property
Public Property Let CustPaid(val As Boolean)
    mudtProps.CustPaid = val
End Property
Public Property Get CommPaid() As Boolean
    CommPaid = mudtProps.CommPaid
End Property
Public Property Let CommPaid(val As Boolean)
    mudtProps.CommPaid = val
End Property
Public Property Get SalesRepID() As Long
    SalesRepID = mudtProps.SalesRepID
End Property
Public Property Let SalesRepID(val As Long)
    mudtProps.SalesRepID = val
End Property
Public Property Get SalesRepName() As String
    SalesRepName = FNS(mudtProps.SalesRepName)
End Property
Public Property Let SalesRepName(val As String)
    mudtProps.SalesRepName = val
End Property
Public Property Get TPName() As String
    On Error GoTo errHandler
    TPName = FNS(mudtProps.TPName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPName"
End Property
Public Property Let TPName(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPName = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPName(val)", val
End Property

Public Property Get TPPhone() As String
    On Error GoTo errHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPPhone"
End Property
Public Property Let TPPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
    On Error GoTo errHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPFax"
End Property
Public Property Let TPFax(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPFax(val)", val
End Property

Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPACCNum(val)", val
End Property

'Public Property Get Memo() As String
'    Memo = FNS(mudtProps.Memo)
'End Property
'Public Property Let Memo(val As String)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Memo = val
'    SetDirty True
'End Property

Public Property Get BusPhone() As String
    On Error GoTo errHandler
    BusPhone = FNS(mudtProps.BusPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.BusPhone"
End Property
Public Property Let BusPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BusPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.BusPhone(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DocCode"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DocCode(val)", val
End Property

Public Property Get CurrencyFormat() As String
    On Error GoTo errHandler
    CurrencyFormat = FNS(mudtProps.CurrencyFormat)
    If CurrencyFormat <= "" Then CurrencyFormat = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").FormatString
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CurrencyFormat"
End Property
Public Property Let CurrencyFormat(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CurrencyFormat = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CurrencyFormat(val)", val
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = CDate(ReverseDate(mudtProps.DOCDate))  'to ensure no time included
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DocDateF"
End Property
Public Property Get TransDateF2() As String
    On Error GoTo errHandler
    TransDateF2 = Format(mudtProps.DOCDate, "dd mmmm yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TransDateF2"
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
    ErrorIn "a_Invoice.DocDate(val)", val
End Property

Public Property Get IssDate() As Date
    On Error GoTo errHandler
    IssDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IssDate"
End Property
Public Property Get IssDateF() As Date
    On Error GoTo errHandler
    IssDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IssDateF"
End Property
Public Property Let IssDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IssDate(val)", val
End Property

'Public Property Get ForeignCurrency() As a_Currency
'    Set ForeignCurrency = oCurrency_Foreign
'End Property

Public Property Get CurrencyFactor() As Double
    On Error GoTo errHandler
    If mudtProps.CurrencyFactor = 0 Then mudtProps.CurrencyFactor = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").Factor
    CurrencyFactor = mudtProps.CurrencyFactor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CurrencyFactor"
End Property
Public Property Let CurrencyFactor(val As Double)
    On Error GoTo errHandler
    mudtProps.CurrencyFactor = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CurrencyFactor(val)", val
End Property

Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATRate"
End Property
Public Property Get VATRateFormatted() As String
    On Error GoTo errHandler
    VATRate = Format(mudtProps.VATRate, "##0.00")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATRateFormatted"
End Property
Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATRate(val)", val
End Property

'Public Property Get InvoiceDate() As Date
'    InvoiceDate = mudtProps.InvoiceDate
'End Property
'Public Property Get InvoiceDateFormatted() As String
'    InvoiceDate = Format(mudtProps.InvoiceDate, "dd/mm/yyyy")
'End Property
'Public Property Let InvoiceDate(Val As Date)
'    If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.InvoiceDate = Val
'    SetDirty True
'End Property

Public Property Get VATable() As Boolean
    On Error GoTo errHandler
  '  Vatable = mudtProps.Vatable
  VATable = Me.Customer.VATable
 ' VATable = mudtProps.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VATable(val)", val
End Property


Public Property Get InvoiceLines() As ch_InvoiceLine
  Set InvoiceLines = mColInvoiceLines
End Property
Public Property Get DisplayInvoice() As d_Invoice
    On Error GoTo errHandler
    Set DisplayInvoice = New d_Invoice
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DisplayInvoice"
End Property
Private Property Get StatusForPrinting() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusForPrinting = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusForPrinting = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusForPrinting = ""
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusForPrinting = "PROFORMA"
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusForPrinting = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.StatusForPrinting"
End Property
Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.StatusF"
End Property
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Status"
End Property
Public Function SetMemo(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Memo) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Memo Then
        mudtProps.Memo = val
        SetDirty True
    End If
    SetMemo = True
    Exit Function
End Function
Public Function SetMemoDirect(val As String) As Boolean
Dim oSM As z_StockManager
    If Len(val) > Len(mudtProps.Memo) Then
        Err.Raise 384
    End If
    Set oSM = New z_StockManager
    oSM.SetMemo val, Me.InvoiceID
    
End Function
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Memo"
End Property
Public Function GetInvoiceNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim strDefaultCo As String

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.CO
    If mudtProps.Proforma Then
        cmd.CommandText = "sp_GetNextProformaNo"
    Else
        cmd.CommandText = "sp_GetNextInvNo"
    End If
    cmd.CommandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , COMPID)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@NewCode", adVarChar, adParamOutput, 15)
    cmd.Parameters.Append prm
    cmd.execute
'    strDefaultCo = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyCode
   ' GetInvoiceNumber = strDefaultCo & "I" & cmd.Parameters("@Number").Value
    GetInvoiceNumber = cmd.Parameters("@NewCode").Value
    Set cmd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.GetInvoiceNumber(strTmp)", strTmp
End Function

Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property
Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property
Public Function PrintInvoice_Display(Optional ByVal pRoundedUp As Boolean) As String
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim bForeign As Boolean
Dim oIL As a_InvoiceLine
Dim strInvoice As String
Dim i As Long

    
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If

    
    
'    bForeign = (oCurrency_Foreign Is oPC.Configuration.LocalCurrency)
    strInvoice = "Document No. " & Me.DOCCode & vbCrLf
    For i = 1 To Me.InvoiceLines.Count
        Set oIL = Me.InvoiceLines.FindLineBySeqNum(i)
        If Not oIL Is Nothing Then
            If Not oIL.IsDeleted Then
                If oIL.BottomOfDocument = False Then
                    With oIL
                        If .DiscountPercent > 0 Then bDiscountExists = True
                        If .NonStock Then
                            strInvoice = strInvoice & .TitleAuthorPublisher & "  " & .PLessDiscExtF(bForeign)
                        Else
                            strInvoice = strInvoice & .CodeF & "  " & .Ref & "  " & .QtyF & "  " & Left(.TitleAuthorPublisher, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.DiscountPercent = 0, "", .DiscountPercentF) & "  " & .PLessDiscExtF(bForeign)
                        End If
                    End With
                strInvoice = strInvoice & vbCrLf
                Else
                    iExtrasCount = iExtrasCount + 1
                End If
            End If
        End If
    Next i
    strInvoice = strInvoice & vbCrLf
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strInvoice = strInvoice & "subtotal    " & Me.TotalExtWithOutExtrasF(bForeign) & vbCrLf
        For i = 1 To InvoiceLines.Count
            If InvoiceLines(i).BottomOfDocument = True Then
                With InvoiceLines(i)
                    If .NonStock Then
                        strInvoice = strInvoice & .TitleAuthorPublisher & "    " & .PriceF(bForeign) & vbCrLf
                    End If
                End With
            End If
        Next i
    End If
    If Me.VATable Then ''''''''''(1)
        strInvoice = strInvoice & "Total " & TotalPayableF(bForeign) & vbCrLf
        strInvoice = strInvoice & "Includes VAT of " & TotalVATF(bForeign) & vbCrLf
    Else  'overseas clients
        If Me.ShowVAT Then
            strInvoice = strInvoice & "subtotal    " & TotalLessDiscExtF(bForeign) & vbCrLf   'subtotal 3
            strInvoice = strInvoice & "Less VAT of    " & TotalVATF(bForeign) & vbCrLf
            strInvoice = strInvoice & "Payable    " & TotalPayableF(bForeign) & vbCrLf
        Else
            strInvoice = strInvoice & "Total   " & TotalPayableF(bForeign) & vbCrLf
        End If
    End If
    If TotalDeposit(bForeign) > 0 Then
        strInvoice = strInvoice & "Less deposits paid: " & TotalDepositF(bForeign)
    End If
    PrintInvoice_Display = strInvoice
    
EXIT_Handler:
'ERR_Handler:
'    oPCL.FinishRequest
'    Select Case Err
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'        Resume
'    End Select
'
    Exit Function
errHandler:
    ErrPreserve
    oPCL.FinishRequest
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.PrintInvoice_Display(pRoundedUp)", pRoundedUp
End Function
Public Function PrintInvoice(bBySeqNo As Boolean, Optional ByVal pRoundedUp As Boolean, Optional pQtyCopies As Integer) As Boolean
    On Error GoTo errHandler
Dim bForeign As Boolean
Dim oIL As a_InvoiceLine
Dim oCHKIL As a_InvoiceLine
Dim bReferenceExists As Boolean
Dim bDiscountExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim strTag As String
Dim strLocalPrinter As String
Dim oDC As a_DocumentControl


 '   ExportToXML False
'Four situations obtain
'   1. In a VAT operating area
'       a) VAT is charged      (situation 1)
'       b) VAT is returned (e.g. to overseas client) (situation 2)
'       c) VAT is charged and not returned (e.g.to overseas client) (situation 3)
'   2. In an area where VAT does not apply (situation 4)

'Layout
    'Section 1: Detail lines
    '   subtotal1 (lngTotal1)
    'Section 2: possible discount
    '   subtotal2  (lngTotal2)
    'Section 3: possible extras
    '   subtotal3 (lngTotal3)
    'Section 4: possible deposits paid
    '   subtotal4 (lngTotal4)
    'Section 5: VAT payable/refundable/notnotified
    '   TotalPayable,VATAmount

    bForeign = False
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If
    
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        PrintInvoice = False
        Exit Function
    Else
        Set oPCL = New z_PrintClient
        If Me.Proforma Then
            strTag = "PR1_"
        Else
            strTag = "IN1_"
        End If
        If oPCL.StartRequestPrinting(strTag & Me.DOCCode) = False Then
            PrintInvoice = False
            Exit Function
        End If
        With oDC
            oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .Style, IIf(pQtyCopies > 0, pQtyCopies, .QtyCopies)
        End With
        PrintInvoice = True
    End If
   
    
    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
    bQtyColumnNeeded = False
    For Each oCHKIL In InvoiceLines
        If oCHKIL.Ref > " " Then
            bReferenceExists = True
        End If
        If oCHKIL.DiscountPercent > 0 Then
            bDiscountExists = True
        End If
        If oCHKIL.Deposit > 0 Then
            bDepositExists = True
        End If
        If oCHKIL.Qty > 1 Or oPC.Configuration.AntiquarianYN = False Then
            bQtyColumnNeeded = True
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
    oPCL.LineCOunt Me.InvoiceLines.Count
    oPCL.LineTotalCount 4
    oPCL.WriteTag "MEMO", PackText(Me.Memo)
    oPCL.WriteTag "OURADDRESS", PackText(oPC.Configuration.DefaultCompany.StreetAddress)
    oPCL.WriteTag "PHONE", IIf(Me.TPPhone > "", "Phone: " & Me.TPPhone, "")
    If Me.COMPID > 0 Then
        oPCL.WriteTag "COREGNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CoRegistrationNumber
        oPCL.WriteTag "VATNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).VatNumber
    End If
    oPCL.WriteTag "CUSTOMER", CStr(Trim(Me.Customer.Fullname)) '& IIf(Me.TPACCNum > "", "A(" & Me.TPACCNum & ")", "")
    oPCL.WriteTag "FAX", IIf(BillTOAddress.FaxF > "", "Fax: " & BillTOAddress.FaxF, "")
    oPCL.WriteTag "BILLTO", PackText(BillTOAddress.AddressMailing)
    oPCL.WriteTag "DELTO", PackText(DelToAddress.AddressMailing)
    oPCL.WriteTag "CUSTVATNUMBER", Me.Customer.VatNumber
    oPCL.WriteTag "DATE", Me.TransDateF2
    If Not Me.BillingCompany Is Nothing Then
        oPCL.WriteTag "CONAME", Me.BillingCompany.CompanyName
    Else
        oPCL.WriteTag "CONAME", ""
    End If
    oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
    oPCL.WriteTag "ACCNUMBER", IIf(Me.Customer.AcNo > "", "Acc Num: " & Me.Customer.AcNo, "")
    oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
    oPCL.WriteTag "BANK", PackText(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).BankDetails)
    
    iExtrasCount = FillTableRows(bBySeqNo, bForeign, bDiscountExists)
    
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        oPCL.WriteTag "TOTAL", "subtotal" & "|" & Me.TotalExtWithOutExtrasF(bForeign)
        FillExtras Me, bForeign, bDiscountExists
    End If
    If Me.VATable Then ''''''''''(1)
        oPCL.WriteTag "TOTAL", "Payable " & "|" & TotalPayableF(bForeign)
        oPCL.WriteTag "TOTAL2", "(Includes VAT of " & "|" & TotalVATF(bForeign) & ")"
    Else  'overseas clients
        If Me.ShowVAT Then
            oPCL.WriteTag "TOTAL", "subtotal" & "|" & TotalLessDiscExtF(bForeign)
            oPCL.WriteTag "TOTAL", "Less VAT of " & "|" & TotalVATF(bForeign)
            oPCL.WriteTag "TOTAL", "Payable " & "|" & TotalPayableF(bForeign)
        Else
            oPCL.WriteTag "TOTAL", "Total " & "|" & TotalPayableF(bForeign)
        End If
    End If
    If TotalDeposit(bForeign) > 0 Then
        oPCL.WriteTag "TOTAL", "Less deposits paid   " & "|" & TotalDepositF(bForeign)
        oPCL.WriteTag "TOTAL", "Balance owed =   " & "|" & TotalBalanceAfterDeposits(bForeign)
    End If
    oPCL.FinishRequest
    
EXIT_Handler:
'ERR_Handler:
'    oPCL.FinishRequest
'    Select Case Err
'    Case 5941
'        MsgBox "Bookmark missing"
'        Resume Next
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'        Resume
'    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.PrintInvoice(pRoundedUp)", pRoundedUp
End Function
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
    On Error GoTo errHandler
Dim strLabels As String
Dim strValues As String
Dim i As Integer
    strLabels = ""
    strValues = ""
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strLabels = "subtotal"
        strValues = Me.TotalExtWithOutExtrasF(bForeign)
        For i = 1 To Me.InvoiceLines.Count
            If Me.InvoiceLines(i).BottomOfDocument = True Then
                With Me.InvoiceLines(i)
                    If .NonStock Then
                        strLabels = strLabels & Chr(13) & .TitleAuthorPublisher
                        strValues = strValues & Chr(13) & .PLessDiscExtF(bForeign)
                    End If
                End With
            End If
        Next i
    End If
'    If Len(strLabels) > 0 Then
'        strLabels = strLabels & Chr(13)
'        strValues = strValues & Chr(13)
'    End If
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
    If TotalDeposit(bForeign) > 0 Then
        strLabels = strLabels & Chr(13) & "Less deposits paid"
        strValues = strValues & Chr(13) & TotalDepositF(bForeign)
        strLabels = strLabels & Chr(13) & "(balance owed ="
        strValues = strValues & Chr(13) & TotalBalanceAfterDeposits(bForeign) & ")"
    End If
    pLabels = strLabels
    pValues = strValues
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Function FillTableRows(bBySeqNo As Boolean, bForeign As Boolean, pDiscountExists As Boolean)
    On Error GoTo errHandler
Dim i As Long
Dim iExtrasCount As Integer
Dim oIL As a_InvoiceLine

    iExtrasCount = 0
    For i = 1 To Me.InvoiceLines.Count
        If bBySeqNo Then
            Set oIL = Me.InvoiceLines.FindLineBySeqNum(i)
        Else
            Set oIL = Me.InvoiceLines(i)
        End If
        If Not oIL Is Nothing Then
            If oIL.BottomOfDocument = False Then
                With oIL
                    If .DiscountPercent > 0 Then pDiscountExists = True
                    If .NonStock Then
                        oPCL.ILSend "", "", "", .Title & .SpecialVATMsg, "", "", "", .PLessDiscExtF(bForeign), "", "", "", .Note
                    Else
                        oPCL.ILSend .CodeForExport, .Ref, .QtyF, .SpecialVATMsg & .Title, .MainAuthor, .Publisher, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DepositF(bForeign), .DiscountPercentF, .PLessDiscExtNonVATF(bForeign), PackText(.Note)
                    End If
                End With
            Else
                iExtrasCount = iExtrasCount + 1
            End If
        End If
    Next i
    FillTableRows = iExtrasCount
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.FillTableRows(bBySeqNo,bForeign,pDiscountExists)", Array(bBySeqNo, bForeign, _
         pDiscountExists)
End Function
Private Function FillExtras(pInv As a_Invoice, bForeign As Boolean, pDiscountExists As Boolean) As Integer
    On Error GoTo errHandler
Dim i As Integer
Dim iCnt As Integer

    iCnt = 0
    For i = 1 To pInv.InvoiceLines.Count
        If pInv.InvoiceLines(i).BottomOfDocument = True Then
            With pInv.InvoiceLines(i)
                If .NonStock Then
                    iCnt = iCnt + 1
                    oPCL.ILSendExtra .TitleAuthorPublisher, .PLessDiscExtF(bForeign)
                End If
            End With
        End If
    Next i
    FillExtras = iCnt
EXIT_Handler:
'ERR_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
   ' Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.FillExtras(pInv,bForeign,pDiscountExists)", Array(pInv, bForeign, _
         pDiscountExists)
End Function



Private Function TranslateErrors(pRawErrors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, iStart As Integer
    iMarker = 1
    strAllRules = ""
    If Len(pRawErrors) > 0 Then
        iMarker = InStr(iMarker + 1, pRawErrors, ",")
        If iMarker > 0 Then
            strAllRules = colClassErrors(Left(pRawErrors, iMarker - 1))
        Else
            strAllRules = colClassErrors(pRawErrors)
        End If
        Do Until iMarker = 0
            iStart = iMarker + 1
            iMarker = InStr(iStart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = colClassErrors(Mid(pRawErrors, iStart, iMarker - iStart))
            Else
                strRule = colClassErrors(Mid(pRawErrors, iStart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TranslateErrors(pRawErrors)", pRawErrors
End Function
Private Sub LoadClassErrorsCollection()
    On Error GoTo errHandler
    Set colClassErrors = New Collection
    colClassErrors.Add "Missing customer", "TP"
    colClassErrors.Add "Note too short", "NOTE"
    colClassErrors.Add "Invalid discount", "DISCOUNT"
    colClassErrors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.LoadClassErrorsCollection"
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
    ErrorIn "a_Invoice.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetDirty(pVal)", pVal
End Sub


Public Function Post(Optional pNewStatus As enStatus) As String
    On Error GoTo errHandler
Dim strResult As String
Dim oSM As z_StockManager
    oPC.CO.CommandTimeout = 15
    oPC.CO.IsolationLevel = adXactSerializable
    
    oPC.CO.execute "BEGIN TRANSACTION"
    strResult = Me.ApplyEdit
    If strResult > "" Then  'probably a timeout error
    oPC.CO.execute "ROLLBACK TRANSACTION"
        Exit Function
    End If
    
    Set oSM = New z_StockManager
    If Not Me.Proforma Then
        If pNewStatus > 0 Then SetStatus pNewStatus
        Post = ""
        Select Case mOriginalStatus
        Case stVOID      'VOID
                Post = "VOID"
        Case stInProcess      'In Process
                Select Case pNewStatus
                Case stVOID
           ''         SetStatus stVOID
                Case stInProcess
                    Post = "In Process"
                Case stCOMPLETE
                    If oSM.IssueInvoice(Me, False, False, stCOMPLETE) = "TIMEOUT" Then
                        oPC.CO.execute "ROLLBACK TRANSACTION"
                        Post = "TIMEOUT"
                        GoTo EXITH
                    End If
                  '  SetStatus stCOMPLETE
                Case stPROFORMA
                    If oSM.IssueInvoice(Me, False, False, stPROFORMA) = "TIMEOUT" Then
                        oPC.CO.execute "ROLLBACK TRANSACTION"
                        Post = "TIMEOUT"
                        GoTo EXITH
                    End If
            '        SetStatus stPROFORMA
                End Select
        Case stCOMPLETE      'Complete
                Post = "ISSUED"
        Case stPROFORMA      'Issued PROFORMA
                Select Case Status
                Case stVOID
                    oSM.CancelInvoice Me
            '        SetStatus stVOID
                Case stInProcess
                    Post = "ERROR: was proforma, can't now be inprocess."
                Case stCOMPLETE
                    If oSM.IssueInvoice(Me, True, False, stCOMPLETE) = "TIMEOUT" Then
                        oPC.CO.execute "ROLLBACK TRANSACTION"
                        Post = "TIMEOUT"
                        GoTo EXITH
                    End If
            '        SetStatus stCOMPLETE
                Case stPROFORMA
                    Post = "PROFORMA"
                End Select
        End Select
    End If
    oPC.CO.execute "COMMIT TRANSACTION"
EXITH:
    oPC.CO.IsolationLevel = adXactReadCommitted
    Set oSM = Nothing
    oPC.CO.CommandTimeout = 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Post", , EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.CO
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
 '   If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetStatus(val)", val
End Sub
Public Property Get COLsOSPerCUST() As c_COLsPerCust
  Set COLsOSPerCUST = cColsOSPerCust
End Property
Public Property Get APPLsOSPerCUST() As c_APPLsPerTPPID
  Set APPLsOSPerCUST = cAPPLsOS
End Property
Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.VoidDocument"
End Function
Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.InvoiceID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.Reload"
End Sub

Public Sub SetProforma()
    mudtProps.Proforma = True
End Sub
Public Property Get Proforma() As Boolean
    Proforma = mudtProps.Proforma
End Property

Public Sub InformLocalPOSdb()
Dim DBLocalConn As ADODB.Connection
    Set DBLocalConn = New ADODB.Connection
    DBLocalConn.Provider = "sqloledb"
    DBLocalConn.Open "Data provider=sqloledb;Data Source=" & oPC.POSServername & ";Initial Catalog=PBKSFD;User Id=sa" & ";Password=" & "" & "; Network Library=dbmssocn;"
    DBLocalConn.CommandTimeout = 0
        DBLocalConn.execute "INSERT INTO tBRIDGE (INV_ID,INV_DOCCode,INV_DOCDate,INV_CustomerName,INV_Payable,INV_TPID,INV_VATVALUE) VALUES (" & Me.InvoiceID & ",'" & Me.DOCCode & "'," & Me.DOCDate & ",'" & Me.Customer.Name & "'," & Me.TotalPayableRaw(False) & "," & Me.Customer.ID & "," & Me.TotalVATRaw(False) & ")"
    DBLocalConn.Close
    Set DBLocalConn = Nothing
End Sub

Public Function ExportToXML(bBySeqNo As Boolean, Optional ByVal pRoundedUp As Boolean, Optional pQtyCopies As Integer) As Boolean
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
Dim objXSL As New MSXML2.DOMDocument30
Dim opXMLDOC As New MSXML2.DOMDocument30
Dim objXMLDOC  As New MSXML2.DOMDocument30
Dim bReferenceExists As Boolean
Dim bDiscountExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim oCHKIL As a_InvoiceLine
Dim oDC As a_DocumentControl
'Const strHead = "<?xml version=""1.0"" encoding=""UTF-8""?>" _
'            & "<PO_1 xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""C:\PBKS\Templates\PO_FO_1.xsd"">"
Dim strErrPos As String

    strErrPos = "pos 1"
    bForeign = False
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If

    strErrPos = "pos 2"
    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
    bQtyColumnNeeded = False
    For Each oCHKIL In InvoiceLines
        If oCHKIL.Ref > " " Then
            bReferenceExists = True
        End If
        If oCHKIL.DiscountPercent > 0 Then
            bDiscountExists = True
        End If
        If oCHKIL.Deposit > 0 Then
            bDepositExists = True
        End If
        If oCHKIL.Qty > 1 Or oPC.Configuration.AntiquarianYN = False Then
            bQtyColumnNeeded = True
        End If
    Next
    strErrPos = "pos 3"
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing Then
        ExportToXML = False
        Exit Function
    End If
    strErrPos = "pos 4"
    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "IN_1"
            
            .chCreate "DestinationAddress"
                If Customer.DelToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Customer.DelToAddress.Email
                End If
            .elCreateSibling "TemplateName"
            If bDiscountExists Then
                    .elText = "IN_DISC_REF_FO.XSL"
            ElseIf bReferenceExists Then
                    .elText = "IN_REF_FO.XSL"
            Else
                    .elText = "IN_FO.XSL"
            End If
    strErrPos = "pos 5"
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                .elText = oDC.PrinterName
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "IsProforma"
                .elText = IIf(Proforma, "Y", "N")
            .elCreateSibling "DocCode"
                .elText = DOCCode
            .elCreateSibling "DocDate", True
                .elText = DocDateF
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName                  'oPC.Configuration.DefaultCompany.CompanyName
    strErrPos = "pos 6"
'MsgBox "pos 6"
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerName", True
                .elText = Customer.Name
            .elCreateSibling "CustomerVATNumber", True
            If Customer.VatNumber > "" Then
                .elText = "Customer's VAT number: " & Customer.VatNumber
            Else
                .elText = ""
            End If
            .elCreateSibling "ACNO", True
                .elText = Customer.AcNo
            .elCreateSibling "BankingDetails", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(COMPID).BankDetails, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerWithAddress", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Me.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
    strErrPos = "pos 7"
'MsgBox "pos 7"
            .elCreateSibling "CustomerPhone", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.BillTOAddress.Phone > "", "Phone: " & Customer.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "CustomerFax", True
                .elText = IIf(Me.BillTOAddress.Fax > "", "Fax: " & Customer.BillTOAddress.Fax, "")
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
    strErrPos = "pos 8"
'MsgBox "pos 8"
            For i = 1 To InvoiceLines.Count
                If InvoiceLines(i).BottomOfDocument = False Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    If oPC.HideLocalSKUonINV Then
                        .elText = InvoiceLines(i).CodeForExport
                    Else
                        .elText = InvoiceLines(i).CodeF
                    End If
                    .elCreateSibling "Title", True
                        .elText = InvoiceLines(i).TitleAuthor
                    .elCreateSibling "QtyFirm", True
                        .elText = InvoiceLines(i).QtyF
                    .elCreateSibling "Price", True
                        .elText = InvoiceLines(i).PriceF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = InvoiceLines(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = InvoiceLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = InvoiceLines(i).PLessDiscExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = InvoiceLines(i).Note
                    .navUP
                End If
            Next i
    strErrPos = "pos 9"
'MsgBox "pos 9"
            For i = 1 To InvoiceLines.Count
                If InvoiceLines(i).BottomOfDocument = True Then
                    .elCreateSibling "BottomItems", True
                    .chCreate "SKU"
                        .elText = InvoiceLines(i).CodeF
                    .elCreateSibling "Title", True
                        .elText = InvoiceLines(i).TitleAuthorPublisher
                    .elCreateSibling "QtyFirm", True
                        .elText = InvoiceLines(i).QtyF
    
    
                    .elCreateSibling "Price", True
                        .elText = InvoiceLines(i).PriceF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = InvoiceLines(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = InvoiceLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = InvoiceLines(i).PLessDiscExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = InvoiceLines(i).Note
                    .navUP
                End If
            Next i
    strErrPos = "pos 10" '
'MsgBox "pos 10"
        SetupDocumentTotals bForeign
    strErrPos = "pos 11"
'MsgBox "pos 11"
            For i = 1 To UBound(arDocTotalsF)
                    If arDocTotalsF(i).Description = "" Then Exit For
                    .elCreateSibling "Totals", True
                    .chCreate "TotalDescription"
                        .elText = arDocTotalsF(i).Description
                    .elCreateSibling "NumericValue", True
                        .elText = arDocTotalsF(i).AmtFormatted
                    .navUP
            Next i
            
    strErrPos = "pos 12"
'MsgBox "pos 12"
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
    End With
    strErrPos = "pos 13"
'MsgBox "pos 13"
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\Printing\IN1_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

    strErrPos = "pos 14"
'MsgBox "pos 14"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ExportToXML(bForeign)", bForeign, , , "strErrPos", Array(strErrPos)
End Function


Private Sub SetupDocumentTotals(bForeign As Boolean)
    On Error GoTo errHandler
Dim ar() As String
Dim i, k As Integer
Dim TotalSequence As String
Dim bBottomOfDocument As Boolean

    TotalSequence = oPC.InvoiceTotalSequence
    
    ReDim arDocTotals(16)
    ReDim arDocTotalsF(16)
    ReDim ar(3)
    ar = Split(TotalSequence, ",")
    
    jTotal = 1
    RunningTotal = 0
    
    If Me.TotalPayable(bForeign) = Me.TotalExtWithOutExtras(bForeign) Then  'THer are no extras
        addARow "Total", TotalExtWithOutExtras(bForeign), TotalExtWithOutExtrasF(bForeign), "", "+"
    Else
        addARow "subtotal", TotalExtWithOutExtras(bForeign), TotalExtWithOutExtrasF(bForeign), "", "+"
    End If
    
    If Customer.VATable Then
        'Handle extras
            bBottomOfDocument = False
            For k = 1 To InvoiceLines.Count
                If InvoiceLines(k).BottomOfDocument = True Then
                    addARow InvoiceLines(k).Title, InvoiceLines(k).ForeignPrice, InvoiceLines(k).PriceF(bForeign), "", "+"
                    bBottomOfDocument = True
                End If
            Next k
            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "GRAND"
            addARow "(includes VAT at " & oPC.Configuration.VATRateF, TotalVATRaw(bForeign), TotalVATF(bForeign), "", " "
    Else
            bBottomOfDocument = False
            For k = 1 To InvoiceLines.Count
                If InvoiceLines(k).BottomOfDocument = True And InvoiceLines(k).VATRate <> 0 Then
                    addARow InvoiceLines(k).Title, InvoiceLines(k).ForeignPrice, InvoiceLines(k).PriceF(bForeign), "", "+"
                    bBottomOfDocument = True
                End If
            Next k
            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "SUB"
            If ShowVAT Then
                addARow "less VAT at " & oPC.Configuration.VATRateF, TotalVATRaw(bForeign), TotalVATF(bForeign), "SUB", "-"
            End If
          '  bBottomOfDocument = False
            For k = 1 To InvoiceLines.Count
                If InvoiceLines(k).BottomOfDocument = True And InvoiceLines(k).VATRate = 0 Then
                    addARow InvoiceLines(k).Title, InvoiceLines(k).ForeignPrice, InvoiceLines(k).PriceF(bForeign), "", "+"
                    bBottomOfDocument = True
                End If
            Next k
            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "GRAND"
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
    ErrorIn "a_Invoice.SetupDocumentTotals(bForeign)", bForeign
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
    ErrorIn "a_Invoice.addARow(pDescription,Amt,AmtF,pTotalType)", Array(pDescription, Amt, _
         AmtF, pTotalType)
End Sub

'Public Sub CopyFromInvoice(pInv As a_Invoice)
'Dim oIL As a_InvoiceLine
'Dim oNewL As a_InvoiceLine
'
'    Me.SetCustomer pInv.Customer.ID
'    Me.SetStatus stInProcess
'    Me.ShowVAT = pInv.ShowVAT
'    Me.DocDate = Now()
'    Me.COMPID = pInv.COMPID
'    For Each oIL In pInv.InvoiceLines
'        Set oNewL = Me.InvoiceLines.Add
'        oNewL.BeginEdit
'        oNewL.PID = oIL.PID
'        oNewL.Qty = oIL.Qty
'        oNewL.Price = oIL.Price
'        oNewL. = oIL.InvoiceLineID
'        oNewL.Discount = oIL.DiscountPercent
'        oNewL.CopyID = oIL.PIID
'        oNewL.ApplyEdit
'    Next
'End Sub

