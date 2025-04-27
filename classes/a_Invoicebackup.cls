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
Dim lngTotal1 As Long
Dim lngTotal2 As Long
Dim lngTotal3 As Long
Dim lngTotal4 As Long
Dim lngTotalDeposit As Long
Dim lngTotalVAT As Long
Dim lngTotalInvoiceDiscount As Long
Dim lngTotalFinal As Long
Dim lngTotalDiscount As Long
Dim lngTotalExtras As Long
Dim lngTotalExtensionInclDepositIncDiscount As Long
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Dim iExtrasCount As Integer

Event Valid(pMsg As String)
Event TotalChange(lngTotal As Long, strTotal As String, lngTotalDeposit As Long, strtotalDeposit As String, lngTotalVAT As Long, strTotalVAT As String)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Private mudtProps As InvoiceProps
Private mcolStack As Collection
Private cColsOSPerCust As c_COLsPerCust
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

Public Property Get HasCoffs() As Boolean
Dim bFound As Boolean
Dim oIL As a_InvoiceLine

    bFound = False
    For Each oIL In InvoiceLines
        If oIL.COFFs.Count > 0 Then
            bFound = True
        End If
    Next
    HasCoffs = bFound
End Property

Private Sub mColInvoiceLines_CurrRowStatus(pMsg As String)
    RaiseEvent CurrRowStatus(pMsg)
End Sub
Public Property Get ForeignCurrency() As a_Currency
    Set ForeignCurrency = oCurrency_Foreign
End Property
Private Sub mcolinvoiceLines_RowsChange()
'occurs for new rows and deletion of rows
    
    CalculateTotal
    RaiseEvent TotalChange(lngTotal2, Format(lngTotal2 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
                        lngTotalDeposit, Format(lngTotalDeposit / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString), _
                        lngTotalVAT, Format(lngTotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString))
    SetDirty True
End Sub
Public Function CalculateTotal()
Dim oIL As a_InvoiceLine
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngTotal1 = 0
    lngTotal2 = 0
    lngTotal3 = 0
    lngTotal4 = 0
    lngTotalDeposit = 0
    lngTotalVAT = 0
    lngTotalDiscount = 0
    'NOTE Some items on the invoice might be nonstock e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non nonstock items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    iExtrasCount = 0
    For Each oIL In mColInvoiceLines
        If oIL.BottomOfDocument Then
            iExtrasCount = iExtrasCount + 1
        End If
        lngTotalVAT = lngTotalVAT + oIL.VATAmount
        lngTotalDiscount = lngTotalDiscount + oIL.DiscountAmount
        lngTotal3 = lngTotal3 + oIL.Extension
        lngTotal4 = lngTotal4 + oIL.ExtensionInclDeposit
        If oIL.BottomOfDocument Then
            lngTotalExtras = lngTotalExtras + oIL.Extension
        Else
            lngTotal2 = lngTotal2 + oIL.Extension
            lngTotalDeposit = lngTotalDeposit + oIL.Deposit
            lngTotal1 = lngTotal1 + oIL.ExtensionSimple
        End If
    Next
    mudtProps.TotalDiscount = lngTotalDiscount
    mudtProps.TotalVAT = lngTotalVAT
    mudtProps.TotalPayable = lngTotal4
    If (Not VATable) And ShowVAT Then
        mudtProps.TotalPayable = lngTotal4 - lngTotalVAT
    End If
End Function
Public Function CalculateTotal_Foreign()
Dim oIL As a_InvoiceLine
Dim bLineDiscountApplied As Boolean
    bLineDiscountApplied = False
    lngTotal1 = 0
    lngTotal2 = 0
    lngTotal3 = 0
    lngTotal4 = 0
    lngTotalDeposit = 0
    lngTotalVAT = 0
    lngTotalDiscount = 0
    iExtrasCount = 0
    For Each oIL In InvoiceLines
        oIL.SetForeignPriceAndDeposit
    Next

    For Each oIL In mColInvoiceLines
        If oIL.BottomOfDocument Then
            iExtrasCount = iExtrasCount + 1
        End If
        lngTotalVAT = lngTotalVAT + oIL.VATAmount_foreign
        lngTotalDiscount = lngTotalDiscount + oIL.DiscountAmount_Foreign
        lngTotal3 = lngTotal3 + oIL.Extension_Foreign
        lngTotal4 = lngTotal4 + oIL.ExtensionInclDeposit_Foreign
        If oIL.BottomOfDocument Then
            lngTotalExtras = lngTotalExtras + oIL.Extension_Foreign
        Else
            lngTotal2 = lngTotal2 + oIL.Extension_Foreign
            lngTotalDeposit = lngTotalDeposit + oIL.Deposit_Foreign
            lngTotal1 = lngTotal1 + oIL.ExtensionSimple_foreign
        End If
    Next
    mudtProps.TotalDiscount_Foreign = lngTotalDiscount
    mudtProps.TotalVAT_Foreign = lngTotalVAT
    mudtProps.TotalPayable_Foreign = lngTotal4
    If (Not Me.VATable) And ShowVAT Then
        mudtProps.TotalPayable_Foreign = lngTotal4 - lngTotalVAT
    End If
End Function
Public Property Get TotalInvoiceDiscountAmountF() As Long
    TotalInvoiceDiscountAmountF = Format(lngTotalInvoiceDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Total_2F(bForeign As Boolean) As String
    If bForeign Then
        Total_2F = Format(lngTotal2 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        Total_2F = Format(lngTotal2 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get TotalVATF(bForeign As Boolean) As String
    If bForeign Then
        TotalVATF = Format(mudtProps.TotalVAT_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get TotalPayable(bForeign As Boolean) As String
    If bForeign Then
        TotalPayable = mudtProps.TotalPayable_Foreign / oCurrency_Foreign.Divisor
    Else
        TotalPayable = mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor
    End If
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
    If bForeign Then
        TotalPayableF = Format(TotalPayable(bForeign), oCurrency_Foreign.FormatString)
    Else
        TotalPayableF = Format(TotalPayable(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get TotalDiscount(bForeign As Boolean) As Long
    If bForeign Then
        TotalDiscount = mudtProps.TotalDiscount_Foreign
    Else
        TotalDiscount = mudtProps.TotalDiscount
    End If
End Property
Public Property Get TotalDiscountF(bForeign As Boolean) As String
    If bForeign Then
        TotalDiscountF = Format(mudtProps.TotalDiscount_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get TotalDepositF(bForeign As Boolean) As String
    If bForeign Then
        TotalDepositF = Format(TotalDeposit(bForeign), oCurrency_Foreign.FormatString)
    Else
        TotalDepositF = Format(TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get TotalDeposit(bForeign As Boolean) As Long
    If bForeign Then
        TotalDeposit = lngTotalDeposit / oCurrency_Foreign.Divisor
    Else
        TotalDeposit = lngTotalDeposit / oPC.Configuration.DefaultCurrency.Divisor
    End If
End Property
Public Property Get InvoiceDiscountRate() As Double
    InvoiceDiscountRate = mudtProps.DiscountRate
End Property
Public Property Get InvoiceDiscountRateF() As String
    If mudtProps.DiscountRate = 0 Then
        InvoiceDiscountRateF = ""
    Else
        InvoiceDiscountRateF = Format(mudtProps.DiscountRate, gPercentFormatString)
    End If
End Property
Public Property Get Total1() As Double
    Total1 = lngTotal1
End Property
Public Property Get Total_1F(bForeign As Boolean) As String
    If bForeign Then
        Total_1F = Format(lngTotal1 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        Total_1F = Format(lngTotal1 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get Total_3F(bForeign As Boolean) As String
    If bForeign Then
        Total_3F = Format(lngTotal3 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        Total_3F = Format(lngTotal3 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get Total_4F(bForeign As Boolean) As String
    If bForeign Then
        Total_4F = Format(lngTotal4 / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        Total_4F = Format(lngTotal4 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property

Public Property Get TotalBalanceAfterDeposits(bForeign As Boolean) As String
    If bForeign Then
        TotalBalanceAfterDeposits = Format(TotalPayable(bForeign) - TotalDeposit(bForeign), oCurrency_Foreign.FormatString)
    Else
        TotalBalanceAfterDeposits = Format(TotalPayable(bForeign) - TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Private Function GetState() As String
  Dim udtData As InvoiceData
  
  LSet udtData = mudtProps
  GetState = udtData.buffer
End Function
Private Sub SetState(buffer As String)
  Dim udtData As InvoiceData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Public Function GetSuperState() As String
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "ILS", mColInvoiceLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
End Function
Public Sub SetSuperState(buffer As String)
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
End Sub


Public Sub BeginEdit()
  mColInvoiceLines.BeginEdit
  mcolStack.Add GetState
End Sub

Public Sub CancelEdit()
  If mcolStack.Count = 0 Then Err.Raise 445
  mColInvoiceLines.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
End Sub

Public Sub ApplyEdit()
Dim objPersist As a_Invoice_P
    CalculateTotal
    mudtProps.TPID = Me.Customer.ID
    'Set these values fromthe objects before saving -  the objects cant be saved
    mudtProps.BillToAddressID = oBillToAddress.ID
    mudtProps.DelToAddressID = oDelToAddress.ID
        
    If DocCode = "" Then
        DocCode = GetInvoiceNumber("GETANDWRITE")
        DocDate = Date
    End If
 '   If DocDate < #1/1/1995# Then DocDate = InvoiceDate
    If CaptureDate < #1/1/1995# And mudtProps.IsNew Then CaptureDate = Now
    
    If mcolStack.Count = 0 Then Err.Raise 445
    mColInvoiceLines.ApplyEdit
    Set objPersist = New a_Invoice_P
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.InvoiceID
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
End Sub
Public Property Get IsDeleted() As Boolean
  IsDeleted = mudtProps.IsDeleted
End Property
Public Property Get IsEditing() As Boolean
  IsEditing = mcolStack.Count > 0
End Property

Public Property Get IsNew() As Boolean
  IsNew = mudtProps.IsNew
End Property

Public Property Get IsDirty() As Boolean
  IsDirty = mudtProps.IsDirty Or mColInvoiceLines.IsDirty
End Property

Public Property Get IsValid() As Boolean
    IsValid = (mobjValid.Count = 0)
End Property
Public Function GetSTatus()
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
End Function
Private Sub Class_Initialize()
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mColInvoiceLines = New ch_InvoiceLine
    mColInvoiceLines.Component Me
    Set mCustomer = New a_Customer

    Set oPCL = New z_PrintClient
    mudtProps.IsNew = True
    mudtProps.VATRate = oPC.Configuration.VATRate
    mudtProps.VATable = True
    mudtProps.ShowVAT = True
    mudtProps.DiscountRate = 0
    mOriginalStatus = stInProcess
    Set oBillToAddress = New a_Address
    Set oDelToAddress = New a_Address
    mobjValid.BreakRule "TP", True
    LoadClassErrorsCollection
    Set oCurrency_Foreign = oPC.Configuration.DefaultCurrency
End Sub

Private Sub Class_Terminate()
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
      
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mColInvoiceLines = Nothing
    Set mCustomer = Nothing
    Set oPCL = Nothing
    Set cColsOSPerCust = Nothing
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    RaiseEvent Valid(TranslateErrors(errors))
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    RaiseEvent Valid(TranslateErrors(errors))
End Sub

Private Sub mobjValid_NoBrokenRules()
    RaiseEvent Valid("")
End Sub
Private Sub mobjValid_Status(pMsg As String)
    RaiseEvent Valid(TranslateErrors(pMsg))
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
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
    For Each oADd In Customer.Addresses
        If oADd.ID = mudtProps.BillToAddressID Then
            Set oBillToAddress = oADd
        End If
        If oADd.ID = mudtProps.DelToAddressID Then
            Set oDelToAddress = oADd
        End If
    Next
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
    If Not oCurrency_Foreign Is Nothing Then
     '   bForeign = True
        CalculateTotal_Foreign
    End If

    CalculateTotal
    mcolinvoiceLines_RowsChange
End Sub
Public Property Get Customer() As a_Customer
    Set Customer = mCustomer
End Property
Public Sub Delete()
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
End Sub



Public Property Get InvoiceID() As Long
    InvoiceID = mudtProps.InvoiceID
End Property
Public Property Let InvoiceID(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.InvoiceID = val
    SetDirty True
End Property

Public Property Get BillToAddressID() As Long
    BillToAddressID = mudtProps.BillToAddressID
End Property
Public Sub SetBillToAddress(val As a_Address)
    If mcolStack.Count = 0 Then Err.Raise 383
    Set oBillToAddress = val
    mudtProps.BillToAddressID = oBillToAddress.ID
    SetDirty True
End Sub
Public Property Get DelToAddressID() As Long
    DelToAddressID = mudtProps.DelToAddressID
End Property
Public Sub setDelToAddress(val As a_Address)
    If mcolStack.Count = 0 Then Err.Raise 383
    Set oDelToAddress = val
    mudtProps.DelToAddressID = oDelToAddress.ID
    SetDirty True
End Sub
Public Function SetGeneralDiscount(val As String) As Boolean
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
        mobjValid.RuleBroken "DiscountPercent", True
    Else
        mobjValid.RuleBroken "DiscountPercent", False
    End If
    mudtProps.DiscountRate = dblDiscountPercent
    SetDirty True
    For Each oIL In InvoiceLines  'disctribute discount rate over all lines where non nonstaock items exists
        If Not oIL.NONStock Then
            oIL.DiscountPercent = dblDiscountPercent
        End If
    Next
    mcolinvoiceLines_RowsChange  'recalulates andd raises event
    RaiseEvent reloadlist
End Function
Public Function SetCustomer(pTPID As Long) As Boolean
Dim bSuccess As Boolean
    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    If bSuccess Then
        VATable = Customer.VATable
        mobjValid.RuleBroken "TP", False
        Set oBillToAddress = Me.Customer.BillTOAddress
        Set oDelToAddress = Me.Customer.DelToAddress
        SetDirty True
'        Set cColsOSPerCust = Nothing
'        Set cColsOSPerCust = New c_COLsPerCust
'        cColsOSPerCust.Load pTPID
    End If
End Function
Public Function SetCustomerFromAccNum(pAccnum As String) As Boolean
Dim bSuccess As Boolean

    bSuccess = Me.Customer.Load(, pAccnum)
    SetCustomerFromAccNum = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
        If Me.BillToAddressID = 0 Then
            Set oBillToAddress = Me.Customer.DefaultAddress
        End If
        If Me.DelToAddressID = 0 Then
            Set oDelToAddress = Me.Customer.DefaultAddress
        End If
    End If
    Me.VATable = Customer.VATable
    SetDirty True
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
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowVAT = val
    SetDirty True
End Property
Public Property Get ShowVAT() As Boolean
    ShowVAT = mudtProps.ShowVAT
End Property
Public Property Get SalesPersonID() As Long
    SalesPersonID = mudtProps.SalesPersonID
End Property
Public Property Let SalesPersonID(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SalesPersonID = val
    SetDirty True
End Property

Public Property Get CurrencyID_Foreign() As Long
    CurrencyID_Foreign = mudtProps.CurrencyID_Foreign
End Property
Public Property Let CurrencyID(val As Long)
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
    For Each oIL In InvoiceLines
        oIL.SetForeignPriceAndDeposit
    Next
    CalculateTotal_Foreign
End Property
Public Property Get COMPID() As Long
    COMPID = mudtProps.COMPID
End Property
Public Property Let COMPID(val As Long)
Dim oComp As a_Company
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.COMPID = val
    SetDirty True
    For Each oComp In oPC.Configuration.Companies
        If oComp.ID = mudtProps.COMPID Then
            Set oBillingCompany = oComp
        End If
    Next
End Property
Public Property Get BillingCompany() As a_Company
    Set BillingCompany = oBillingCompany
End Property
'Public Property Let BillingCompsny(val As a_Company)
'    Set oBillingCompany
'End Property
Public Property Get TPName() As String
    TPName = FNS(mudtProps.TPName)
End Property
Public Property Let TPName(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPName = val
    SetDirty True
End Property

Public Property Get TPPhone() As String
    TPPhone = FNS(mudtProps.TPPhone)
End Property
Public Property Let TPPhone(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
End Property

Public Property Get TPFax() As String
    TPFax = FNS(mudtProps.TPFax)
End Property
Public Property Let TPFax(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
End Property

Public Property Get TPACCNum() As String
    TPACCNum = FNS(mudtProps.TPACCNum)
End Property
Public Property Let TPACCNum(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
End Property

Public Property Get TPMemo() As String
    TPMemo = FNS(mudtProps.TPMemo)
End Property
Public Property Let TPMemo(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPMemo = val
    SetDirty True
End Property

Public Property Get BusPhone() As String
    BusPhone = FNS(mudtProps.BusPhone)
End Property
Public Property Let BusPhone(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BusPhone = val
    SetDirty True
End Property

Public Property Get DocCode() As String
    DocCode = FNS(mudtProps.DocCode)
End Property
Public Property Let DocCode(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DocCode = val
    SetDirty True
End Property

Public Property Get CurrencyFormat() As String
    CurrencyFormat = FNS(mudtProps.CurrencyFormat)
    If CurrencyFormat <= "" Then CurrencyFormat = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").FormatString
End Property
Public Property Let CurrencyFormat(val As String)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CurrencyFormat = val
    SetDirty True
End Property

Public Property Get DocDate() As Date
    DocDate = mudtProps.DocDate
End Property
Public Property Get TransDateFormatted() As String
    TransDateFormatted = Format(mudtProps.DocDate, "dd/mm/yyyy")
End Property
Public Property Let DocDate(val As Date)
    If mcolStack.Count = 0 Then Err.Raise 383
    If IsDate(val) Then
        mudtProps.DocDate = CDate(val)
        SetDirty True
    End If
End Property

Public Property Get CaptureDate() As Date
    CaptureDate = mudtProps.CaptureDate
End Property
Public Property Get CaptureDateFormatted() As Date
    CaptureDate = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
End Property
Public Property Let CaptureDate(val As Date)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
End Property

'Public Property Get ForeignCurrency() As a_Currency
'    Set ForeignCurrency = oCurrency_Foreign
'End Property

Public Property Get CurrencyFactor() As Double
    If mudtProps.CurrencyFactor = 0 Then mudtProps.CurrencyFactor = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").Factor
    CurrencyFactor = mudtProps.CurrencyFactor
End Property
Public Property Let CurrencyFactor(val As Double)
    mudtProps.CurrencyFactor = val
    SetDirty True
End Property

Public Property Get VATRate() As Double
    VATRate = mudtProps.VATRate
End Property
Public Property Get VATRateFormatted() As String
    VATRate = Format(mudtProps.VATRate, "##0.00")
End Property
Public Property Let VATRate(val As Double)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
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
  '  Vatable = mudtProps.Vatable
  VATable = Me.Customer.VATable
 ' VATable = mudtProps.VATable
End Property
Public Property Let VATable(val As Boolean)
    mudtProps.VATable = val
    SetDirty True
End Property


Public Property Get InvoiceLines() As ch_InvoiceLine
  Set InvoiceLines = mColInvoiceLines
End Property
Public Property Get DisplayInvoice() As d_Invoice
    Set DisplayInvoice = New d_Invoice
End Property
Public Property Get StatusF() As String
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
End Property
Public Property Get Status() As enStatus
    Status = mudtProps.Status
End Property
Public Function SetNote(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.TPMemo) Then
        Err.Raise 384
    End If
    If val <> mudtProps.TPMemo Then
        mudtProps.TPMemo = val
        SetDirty True
    End If
    SetNote = True
End Function
Public Property Get Note() As String
    Note = FNS(mudtProps.TPMemo)
End Property
Public Function GetInvoiceNumber(strTmp As String) As String
Dim oCode As z_Code
Dim strDefaultCo As String

    Set oCode = New z_Code
    
    GetInvoiceNumber = oCode.GetNextCode(enInvoice, strTmp)
    
    strDefaultCo = oPC.Configuration.DefaultCompany.CompanyCode
    GetInvoiceNumber = strDefaultCo & "INV" & GetInvoiceNumber
    
    Set oCode = Nothing
End Function

Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property
Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property
Public Function PrintInvoice_Display(Optional ByVal pRoundedUp As Boolean) As String
Dim bDiscountExists As Boolean
Dim bForeign As Boolean
Dim oIL As a_InvoiceLine
Dim strInvoice As String
Dim i As Long

    On Error GoTo ERR_Handler
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    Else
        bForeign = False
        CalculateTotal
    End If
    strInvoice = "INVOICENUMBER: " & Me.DocCode & vbCrLf
    For i = 1 To Me.InvoiceLines.Count
        Set oIL = Me.InvoiceLines.FindLineBySeqNum(i)
        If Not oIL Is Nothing Then
            If oIL.BottomOfDocument = False Then
                With oIL
                    If .DiscountPercent > 0 Then bDiscountExists = True
                    If .NONStock Then
                        strInvoice = strInvoice & .TitleAuthorPublisher & "  " & .ExtensionF(bForeign)
                    Else
                        strInvoice = strInvoice & .CodeF & "  " & .Ref & "  " & .QtyF & "  " & Left(.TitleAuthorPublisher, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.DiscountPercent = 0, "", .DiscountPercentF) & "  " & .ExtensionF(bForeign)
                    End If
                End With
            strInvoice = strInvoice & vbCrLf
            Else
                iExtrasCount = iExtrasCount + 1
            End If
        End If
    Next i
'   If Not bDiscountExists Then
'        CalculateTotal_Foreign
'    End If
    
    strInvoice = strInvoice & vbCrLf
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strInvoice = strInvoice & "subtotal    " & Total_2F(bForeign) & vbCrLf
        For i = 1 To InvoiceLines.Count
            If InvoiceLines(i).BottomOfDocument = True Then
                With InvoiceLines(i)
                    If .NONStock Then
                        strInvoice = strInvoice & .TitleAuthorPublisher & "    " & .ExtensionF(bForeign) & vbCrLf
                    End If
                End With
            End If
        Next i
    End If
'    If lngTotalDeposit > 0 Then
'        strInvoice = strInvoice & "subtotal    " & Total_3F(bForeign) & vbCrLf
'        strInvoice = strInvoice & "Less deposits paid    " & TotalDepositF(bForeign) & vbCrLf
'    End If
    If Me.VATable Then ''''''''''(1)
        strInvoice = strInvoice & "Total " & TotalPayableF(bForeign) & vbCrLf
        strInvoice = strInvoice & "Includes VAT of " & TotalVATF(bForeign) & vbCrLf
    Else  'overseas clients
        If Me.ShowVAT Then
            strInvoice = strInvoice & "subtotal    " & Total_4F(bForeign) & vbCrLf   'subtotal 3
            strInvoice = strInvoice & "Less VAT of    " & TotalVATF(bForeign) & vbCrLf
            strInvoice = strInvoice & "Payable    " & TotalPayableF(bForeign) & vbCrLf
        Else
            strInvoice = strInvoice & "Total   " & TotalPayableF(bForeign) & vbCrLf
        End If
    End If
    If lngTotalDeposit > 0 Then
'        strLabels = strLabels & Chr(13) & "subtotal"
'        strValues = strValues & Chr(13) & Total_3F(bForeign)
'        strLabels = strLabels & Chr(13) & "Less deposits paid   "
        strInvoice = strInvoice & "Less deposits paid: " & TotalDepositF(bForeign)
    End If
    PrintInvoice_Display = strInvoice
EXIT_Handler:
    Exit Function
ERR_Handler:
    oPCL.FinishRequest
    Select Case Err
    Case 5941
        MsgBox "Bookmark missing"
        Resume Next
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select

End Function
Public Sub PrintInvoice(Optional ByVal pRoundedUp As Boolean)
Dim bDiscountExists As Boolean
Dim bForeign As Boolean
Dim oIL As a_InvoiceLine
Dim oCHKIL As a_InvoiceLine
Dim bReferenceExists As Boolean
Dim bDepositExists As Boolean

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

    On Error GoTo ERR_Handler
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If
    Set oPCL = New z_PrintClient
    oPCL.StartRequestPrinting Me.DocCode
    If Me.StatusF <> "In PROCESS" And Me.StatusF <> "COMPLETE" Then
        oPCL.WriteTag "STATUS", Me.StatusF
    End If
    oPCL.WriteTag "INVOICENUMBER", Me.DocCode
    oPCL.WriteTag "LOGO", oPC.Configuration.DefaultCompany.LogoFilePath
    oPCL.WriteTag "OURADDRESS", oPC.Configuration.DefaultCompany.StreetAddressDelimited
    oPCL.WriteTag "PHONE", IIf(Me.TPPhone > "", "Phone: " & Me.TPPhone, "")
    oPCL.WriteTag "COREGNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CoRegistrationNumber
    oPCL.WriteTag "VATNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).VatNumber
    oPCL.WriteTag "CUSTOMER", CStr(Trim(Me.BillTOAddress.AddressDelimited))
    oPCL.WriteTag "FAX", IIf(Me.TPFax > "", "Fax: " & Me.TPFax, "")
    oPCL.WriteTag "DATE", Me.TransDateFormatted
    oPCL.WriteTag "ACCNUMBER", IIf(Me.TPACCNum > "", "Acc Num: " & Me.TPACCNum, "")
    
    iExtrasCount = FillTableRows(Me, bForeign, bDiscountExists)
    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
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
    Next
    If Not bDiscountExists Then
        oPCL.HideDiscountColumn
    End If
    If Not bReferenceExists Then
        oPCL.HideReferenceColumn
    End If
    If Not bDepositExists Then
        oPCL.HideDepositColumn
    End If
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        oPCL.WriteTag "TOTAL", "subtotal" & "|" & Me.Total_2F(bForeign)
        FillExtras Me, bForeign, bDiscountExists
    End If
    If Me.VATable Then ''''''''''(1)
        oPCL.WriteTag "TOTAL3", "Total " & "|" & TotalPayableF(bForeign)
        oPCL.WriteTag "TOTAL4", "Includes VAT of " & "|" & TotalVATF(bForeign)
    Else  'overseas clients
        If Me.ShowVAT Then
            oPCL.WriteTag "TOTAL", "subtotal" & "|" & Total_4F(bForeign)  'subtotal 3
            oPCL.WriteTag "TOTAL2", "Less VAT of " & "|" & TotalVATF(bForeign)
            oPCL.WriteTag "TOTAL3", "Payable " & "|" & TotalPayableF(bForeign)
        Else
            oPCL.WriteTag "TOTAL3", "Total " & "|" & TotalPayableF(bForeign)
        End If
    End If
    If lngTotalDeposit > 0 Then
        oPCL.WriteTag "TOTAL", "Less deposits paid   " & "|" & TotalDepositF(bForeign)
        oPCL.WriteTag "TOTAL", "Balance owed =   " & "|" & TotalBalanceAfterDeposits(bForeign)
    End If
    oPCL.FinishRequest
    
EXIT_Handler:
    Exit Sub
ERR_Handler:
    oPCL.FinishRequest
    Select Case Err
    Case 5941
        MsgBox "Bookmark missing"
        Resume Next
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select
End Sub
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
Dim strLabels As String
Dim strValues As String
Dim i As Integer
    strLabels = ""
    strValues = ""
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strLabels = "subtotal"
        strValues = Me.Total_2F(bForeign)
        For i = 1 To Me.InvoiceLines.Count
            If Me.InvoiceLines(i).BottomOfDocument = True Then
                With Me.InvoiceLines(i)
                    If .NONStock Then
                        strLabels = strLabels & Chr(13) & .TitleAuthorPublisher
                        strValues = strValues & Chr(13) & .ExtensionF(bForeign)
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
            strValues = strValues & Chr(13) & Total_4F(bForeign)
            strLabels = strLabels & Chr(13) & "Less VAT of"
            strValues = strValues & Chr(13) & TotalVATF(bForeign)
            strLabels = strLabels & Chr(13) & "Payable"
            strValues = strValues & Chr(13) & TotalPayableF(bForeign)
        Else
            strLabels = strLabels & Chr(13) & "Total"
            strValues = strValues & Chr(13) & TotalPayableF(bForeign)
        End If
    End If
    If lngTotalDeposit > 0 Then
        strLabels = strLabels & Chr(13) & "Less deposits paid"
        strValues = strValues & Chr(13) & TotalDepositF(bForeign)
        strLabels = strLabels & Chr(13) & "(balance owed ="
        strValues = strValues & Chr(13) & TotalBalanceAfterDeposits(bForeign) & ")"
    End If
    pLabels = strLabels
    pValues = strValues
End Function
Private Function FillTableRows(pInv As a_Invoice, bForeign As Boolean, pDiscountExists As Boolean)
Dim i As Long
Dim iExtrasCount As Integer
Dim oIL As a_InvoiceLine

On Error GoTo ERR_Handler
    iExtrasCount = 0
    For i = 1 To pInv.InvoiceLines.Count
        Set oIL = Me.InvoiceLines.FindLineBySeqNum(i)
        If oIL.BottomOfDocument = False Then
            With oIL
                If .DiscountPercent > 0 Then pDiscountExists = True
                If .NONStock Then
                    oPCL.ILSend "", "", "", .TitleAuthorPublisher & .SpecialVATMsg, "", .ExtensionF(bForeign), "", "", "", .Note
                Else
                    oPCL.ILSend .CodeF, .Ref, .QtyF, .TitleAuthorPublisher & .SpecialVATMsg, .PriceF(bForeign), .ExtensionF(bForeign), .DepositF(bForeign), .DiscountPercentF, .ExtensionExVATF(bForeign), .Note '"InvDate", oInv.TransDateFormatted
                End If
            End With
        Else
            iExtrasCount = iExtrasCount + 1
        End If
    Next i
    FillTableRows = iExtrasCount
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
   ' Resume
End Function
Private Function FillExtras(pInv As a_Invoice, bForeign As Boolean, pDiscountExists As Boolean) As Integer
Dim i As Integer
Dim iCnt As Integer

On Error GoTo ERR_Handler
    iCnt = 0
    For i = 1 To pInv.InvoiceLines.Count
        If pInv.InvoiceLines(i).BottomOfDocument = True Then
            With pInv.InvoiceLines(i)
                If .NONStock Then
                    iCnt = iCnt + 1
                    oPCL.ILSendExtra .TitleAuthorPublisher, .ExtensionF(bForeign)
                End If
            End With
        End If
    Next i
    FillExtras = iCnt
EXIT_Handler:
    Exit Function
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
   ' Resume
End Function



Private Function TranslateErrors(pRawErrors As String) As String
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
End Function
Private Sub LoadClassErrorsCollection()
    Set colClassErrors = New Collection
    colClassErrors.Add "Missing customer", "TP"
    colClassErrors.Add "Note too short", "NOTE"
End Sub
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
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
End Function
Private Sub SetDirty(pVal As Boolean)
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
End Sub


Public Function Post() As String
On Error GoTo ERRH
Dim oSM As z_StockManager
    On Error GoTo ERRH
    
    oPC.CO.BeginTrans
    Me.ApplyEdit
    Set oSM = New z_StockManager
    
    Post = ""
    Select Case mOriginalStatus
    Case stVOID      'VOID
            Post = "VOID"
    Case stInProcess      'In Process
            Select Case Status
            Case stVOID
       ''         SetStatus stVOID
            Case stInProcess
                Post = "In Process"
            Case stCOMPLETE
                oSM.IssueInvoice Me, False, False
              '  SetStatus stCOMPLETE
            Case stPROFORMA
                oSM.IssueInvoice Me, False, True
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
                oSM.IssueInvoice Me, True, False
        '        SetStatus stCOMPLETE
            Case stPROFORMA
                Post = "PROFORMA"
            End Select
    End Select
    oPC.CO.CommitTrans
  Exit Function
ERRH:
    MsgBox "a_Invoice:Post: " & Error
    oError.SetError Err, "a_Invoice:Post: " & Error, Now(), "", "", ""
    oPC.CO.RollbackTrans
    Post = "ERR"
    Exit Function
    Resume
End Function
Public Sub SetStatus(val As enStatus)
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Status = val
    SetDirty True
End Sub
Public Property Get COLsOSPerCUST() As c_COLsPerCust
  Set COLsOSPerCUST = cColsOSPerCust
End Property
Public Function VoidDocument()
Dim strresult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
End Function
Public Sub Reload()
Dim lngTRID As Long
    lngTRID = Me.InvoiceID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
End Sub

