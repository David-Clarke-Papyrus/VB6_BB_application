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
Attribute VB_Exposed = False
Option Explicit
Private oSM As z_StockManager
Private oFSO As FileSystemObject
Dim strCommand As String
Dim lngLPAfterDiscountExt_tot As Long
Dim lngFPAfterDiscountExt_tot As Long
Dim lngLDeposit_tot As Long
Dim lngFDeposit_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Dim lngDiscExVAT_tot As Long
Dim lngLPayableLessVAT_tot As Long
Dim lngFPayableLessVAT_tot As Long
Dim lngTotalServiceItem As Long
Dim lngTotalExtras As Long
Dim lngQty_tot As Long
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
Private colClassors As Collection
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
Private oCurrency_Foreign As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Const CDOCCODE = "IN"

Dim arDocTotals() As tTotal
Dim arDocTotalsF() As tTotal
Dim RunningTotal As Long
Dim jTotal As Long



Public Sub ConvertToNonVATInvoice()
Dim oIL As a_InvoiceLine
    For Each oIL In Me.InvoiceLines
        oIL.Price = Round(oIL.Price / ((100 + oIL.VATRate) / 100))
    Next
    Me.CalculateTotals
End Sub

Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property
'''Public Property Get HasCoffs() As Boolean
'''    On Error GoTo ErrHandler
'''Dim bFound As Boolean
'''Dim oIL As a_InvoiceLine
'''
'''    bFound = False
'''    For Each oIL In InvoiceLines
'''        If oIL.COFFs.Count > 0 Then
'''            bFound = True
'''        End If
'''    Next
'''    HasCoffs = bFound
'''    Exit Property
'''ErrHandler:
'''    If ErrMustStop Then Debug.Assert False: Resume
'''    ErrorIn "a_Invoice.HasCoffs"
'''End Property

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
    RaiseEvent TotalChange(lngLPAfterDiscountExt_tot, lngLDeposit_tot, lngVAT_tot)
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
    lngLPAfterDiscountExt_tot = 0
    lngLDeposit_tot = 0
    lngVAT_tot = 0
    lngDisc_tot = 0
    lngDiscExVAT_tot = 0
    lngLPayableLessVAT_tot = 0
    lngTotalServiceItem = 0
    iExtrasCount = 0
    'NOTE Some items on the invoice might be ServiceItem e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    iExtrasCount = 0
    lngTotalExtras = 0
    lngQty_tot = 0
    For Each oIL In mColInvoiceLines
        If (Not oIL.IsDeleted) Then
    '    If oIL.Fulfilled <> "CAN" Then
            oIL.CalculateLine
            If oIL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            If oIL.ServiceItem = False Then
                lngQty_tot = lngQty_tot + oIL.Qty
            End If
            lngLPayableLessVAT_tot = lngLPayableLessVAT_tot + oIL.PAfterDiscountExtNonVAT(False)
            lngVAT_tot = lngVAT_tot + oIL.PAfterDiscountLineValueVAT(False)
            lngDisc_tot = lngDisc_tot + oIL.PDiscExt(False)
            lngDiscExVAT_tot = lngDiscExVAT_tot + oIL.PDiscExtExVAT
            If oIL.ServiceItem = True Then
                lngTotalServiceItem = lngTotalServiceItem + oIL.PAfterDiscountExt(False)
            End If
            lngLPAfterDiscountExt_tot = lngLPAfterDiscountExt_tot + oIL.PAfterDiscountExt(False)
            lngLDeposit_tot = lngLDeposit_tot + oIL.Deposit
            If oIL.BottomOfDocument Then
                lngTotalExtras = lngTotalExtras + oIL.PAfterDiscountExt(False)
            End If
        End If
    Next
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtras = lngTotalExtras
    mudtProps.DepositPaid = lngLDeposit_tot
    mudtProps.TotalDiscount = lngDisc_tot
    mudtProps.TotalDiscountExVAT = lngDiscExVAT_tot
    mudtProps.TotalNonVAT = lngLPayableLessVAT_tot
    mudtProps.TotalVAT = lngVAT_tot
    mudtProps.TotalExtension = lngLPAfterDiscountExt_tot
    mudtProps.TotalPayable = lngLPAfterDiscountExt_tot - lngLDeposit_tot
    mudtProps.TotalServiceItem = lngTotalServiceItem

    If (Not VATable) And ShowVAT Then
        mudtProps.TotalPayable = lngLPAfterDiscountExt_tot - lngVAT_tot
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
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
    lngFPAfterDiscountExt_tot = 0
    lngFDeposit_tot = 0
    lngVAT_tot = 0
    lngFPayableLessVAT_tot = 0
    lngDisc_tot = 0
    lngTotalExtras = 0
    lngQty_tot = 0
    lngTotalServiceItem = 0
    iExtrasCount = 0
    For Each oIL In mColInvoiceLines
        If (Not oIL.IsDeleted) Then
            If oIL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            lngFPayableLessVAT_tot = lngFPayableLessVAT_tot + oIL.PAfterDiscountExtNonVAT(True)
            lngVAT_tot = lngVAT_tot + oIL.PAfterDiscountLineValueVAT(True)
            lngDisc_tot = lngDisc_tot + oIL.PDiscExt(True)
            If oIL.ServiceItem = True Then
                lngTotalServiceItem = lngTotalServiceItem + oIL.PAfterDiscountExt(True)
            End If
            
            lngFPAfterDiscountExt_tot = lngFPAfterDiscountExt_tot + oIL.PAfterDiscountExt(True)
            lngFDeposit_tot = lngFDeposit_tot + oIL.Deposit_Foreign
            
            If oIL.BottomOfDocument Then
                lngTotalExtras = lngTotalExtras + oIL.PAfterDiscountExt(True)
            End If
            lngQty_tot = lngQty_tot + oIL.Qty
        End If
    Next
    mudtProps.TotalExtras_Foreign = lngTotalExtras
    mudtProps.TotalDiscount_Foreign = lngDisc_tot
    mudtProps.TotalVAT_Foreign = lngVAT_tot
    mudtProps.TotalNonVAT_Foreign = lngFPayableLessVAT_tot
    mudtProps.TotalExtension_Foreign = lngFPAfterDiscountExt_tot
    mudtProps.TotalPayable_Foreign = lngFPAfterDiscountExt_tot - lngFDeposit_tot
    mudtProps.TotalServiceItem_Foreign = lngTotalServiceItem
    If (Not Me.VATable) And ShowVAT Then
        mudtProps.TotalPayable_Foreign = lngFPAfterDiscountExt_tot - lngVAT_tot
    End If
    Exit Function
errHandler:
    ErrPreserve
    If Err = 6 Then 'ov erflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CalculateTotal_Foreign"
End Function
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
Public Property Get Total_nonExtras_ExVAT(bForeign As Boolean)
        Total_nonExtras_ExVAT = TotalExtWithOutExtras(bForeign) - Me.TotalVATRaw(bForeign)
End Property
Public Property Get Total_nonExtras_ExVATF(bForeign As Boolean)
    If bForeign Then
        Total_nonExtras_ExVATF = Format((Total_nonExtras_ExVAT(bForeign)) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        Total_nonExtras_ExVATF = Format((Total_nonExtras_ExVAT(bForeign)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property

Public Property Get TotalAfterDiscountExtF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalAfterDiscountExtF = Format(mudtProps.TotalExtension_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalAfterDiscountExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalAfterDiscountExtF(bForeign)", bForeign
End Property
Public Property Get TotalAfterDiscountExt(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalAfterDiscountExt = mudtProps.TotalExtension_Foreign
    Else
        TotalAfterDiscountExt = mudtProps.TotalExtension
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalAfterDiscountExt(bForeign)", bForeign
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
Public Property Get TotalVAT_WithDecimal(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalVAT_WithDecimal = Format(mudtProps.TotalVAT_Foreign / oCurrency_Foreign.Divisor, "#######0.00")
    Else
        TotalVAT_WithDecimal = Format(mudtProps.TotalVAT / oCurrency_Foreign.Divisor, "#######0.00")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalVAT_WithDecimal(bForeign)", bForeign
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
Public Property Get TotalPayable_WithDecimal(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalPayable_WithDecimal = Format(mudtProps.TotalPayable_Foreign / oCurrency_Foreign.Divisor, "#######0.00")
    Else
        TotalPayable_WithDecimal = Format(mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor, "#######0.00")
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalPayable_WithDecimal(bForeign)", bForeign
End Property

Public Property Get TotalServiceItem(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalServiceItem = mudtProps.TotalServiceItem_Foreign '/ oCurrency_Foreign.Divisor
    Else
        TotalServiceItem = mudtProps.TotalServiceItem ' / oPC.Configuration.DefaultCurrency.Divisor
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalServiceItem(bForeign)", bForeign
End Property
Public Property Get TotalServiceItemF(bForeign As Boolean) As String
    On Error GoTo errHandler
    If bForeign Then
        TotalServiceItemF = Format(TotalServiceItem(bForeign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
    Else
        TotalServiceItemF = Format(TotalServiceItem(bForeign) / oCurrency_Foreign.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalServiceItemF(bForeign)", bForeign
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
Public Property Get TotalDiscountExVAT() As Long
    On Error GoTo errHandler
        TotalDiscountExVAT = mudtProps.TotalDiscountExVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalDiscount"
End Property
Public Property Get TotalDiscountExVATF() As String
    On Error GoTo errHandler
        TotalDiscountExVATF = Format(mudtProps.TotalDiscountExVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalDiscountExVATF"
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
Public Property Get DocDiscountRate() As Double
    On Error GoTo errHandler
    DocDiscountRate = mudtProps.DiscountRate
  '  RecalculateAllLines
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DocDiscountRate"
End Property
Public Property Get DocDiscountRateF() As String
    On Error GoTo errHandler
    If mudtProps.DiscountRate = 0 Then
        DocDiscountRateF = ""
    Else
        DocDiscountRateF = Format(mudtProps.DiscountRate, gPercentFormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DocDiscountRateF"
End Property
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
Dim oSM As z_StockManager
Dim OpenResult As Integer

    ApplyEdit = ""
    CalculateTotals
    mudtProps.TPID = Customer.ID
    mudtProps.SalesRepID = Customer.repid
    
    'Set these values fromthe objects before saving -  the objects cant be saved
    If Me.BillToAddressID > 0 Then
        mudtProps.BillToAddressID = Me.BillToAddressID
    Else
        If Not oBillToAddress Is Nothing Then
            mudtProps.BillToAddressID = oBillToAddress.ID
        End If
    End If
    If Me.DelToAddressID > 0 Then
        mudtProps.DelToAddressID = Me.DelToAddressID
    Else
        If Not oDelToAddress Is Nothing Then
            mudtProps.DelToAddressID = oDelToAddress.ID
        End If
    End If
    
'''''''''''''''''''''''''''
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    oPC.COShort.execute "BEGIN TRANSACTION"
    If Me.DOCDate = CDate(0) Then
        Me.DOCDate = Date
        Me.CaptureDate = Now
    End If
    If DOCCode = "" Then
        DOCCode = GetInvoiceNumber("GETANDWRITE")
  '      CaptureDate = Now()
    End If
  '  If oPC.AllowInvoiceDateOverride = False Then
  '      DOCDate = Date
  '  End If

'    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
    
    If mcolStack.Count = 0 Then Err.Raise 445
    oPC.COShort.CommandTimeout = 35
    mColInvoiceLines.ApplyEdit
    Set objPersist = New a_Invoice_P
    
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.InvoiceID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If Not IsValid Then Err.Raise 445
'        If mudtProps.IsNew Then
'            mudtProps.CaptureDate = Now()
'        End If
        strTemp = objPersist.Save(GetSuperState)
    
        If UCase(Left(strTemp, 7)) <> "TIMEOUT" Then
            mcolStack.Remove mcolStack.Count
            SetSuperState strTemp
            mudtProps.IsNew = False
        Else 'save failed
            mColInvoiceLines.BeginEdit
            If Not IsMissing(pMsg) Then
                pMsg = strTemp
            End If
            ApplyEdit = "TIMEOUT"
            Set oSM = New z_StockManager
            oSM.RollbackTransaction oPC.COShort
            Set oSM = Nothing
            Set objPersist = Nothing
        LogSaveToFile "a_Invoice:TimeoutExpired"
            GoTo EXIT_Handler
        End If
    
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    SetDirty False
    oPC.COShort.execute "COMMIT TRANSACTION"
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
    '--------------
''''''''''''''''''''''''''''''''''''''''
  '  oPC.COShort.CommandTimeout = 0
    
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    oSM.RollbackTransaction oPC.COShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ApplyEdit", , , oPC.COShort
    Exit Function
    Resume Next
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
Public Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetDirty(pVal)", pVal
End Sub
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsValid"
End Property
Public Function GetStatus()
    On Error GoTo errHandler
    If IsValid Then
        RaiseEvent Valid("")
    Else
        RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
    End If
   ' RaiseEvent Dirty(mudtProps.IsDirty)
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
    mColInvoiceLines.component Me
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
    LoadClassorsCollection
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

Private Sub mobjValid_BrokenRule(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(oRS))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.mobjValid_RuleUnbroken(ors)", oRS
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
Dim oAdd As a_Address
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
    Set oBillToAddress = Nothing
    Set oDelToAddress = Nothing
    
    mCustomer.Load mudtProps.TPID
    mudtProps.ShowVAT = Customer.ShowVAT
    
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
    
    If DelToAddressID > 0 Then
        Set oDelToAddress = Customer.Addresses.FindByID(DelToAddressID)
    End If
    'The address might not be found if it has been removed or the customer has been merged with another
    If oDelToAddress Is Nothing Then
        If Not Customer.DelToAddress Is Nothing Then
            Set oDelToAddress = Customer.DelToAddress
            DelToAddressID = Customer.DelToAddress.ID
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
    mudtProps.IsNew = False
    mudtProps.IsDirty = False
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
Public Sub SetBillToAddressImmediate(val As a_Address)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    Set oBillToAddress = val
    mudtProps.BillToAddressID = oBillToAddress.ID
    oSM.SetAddressImmediate "BILL", Me.InvoiceID, oBillToAddress.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.SetBillToAddressImmediate(val)", val
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
Public Sub setDelToAddressImmediate(val As a_Address)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    Set oDelToAddress = val
    mudtProps.DelToAddressID = oDelToAddress.ID
    oSM.SetAddressImmediate "DEL", Me.InvoiceID, oDelToAddress.ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.setDelToAddressImmediate(val)", val
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
        If Not oIL.ServiceItem Then
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
Public Function LoadCOLsOS(Optional pTPID As Long, Optional pPID As String, Optional pCOLID As Long) As c_COLsPerCust
    On Error GoTo errHandler
    Set cColsOSPerCust = Nothing
    Set cColsOSPerCust = New c_COLsPerCust
    cColsOSPerCust.Load Customer.ID, pPID, pCOLID
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.LoadCOLsOS(pTPID,pPID,pCOLID)", Array(pTPID, pPID, pCOLID)
End Function
Public Function LoadAPPLsOS(Optional pTPID As Long, Optional pPID As String) As c_COLsPerCust
    On Error GoTo errHandler
    Set cAPPLsOS = Nothing
    Set cAPPLsOS = New c_APPLsPerTPPID
    cAPPLsOS.Load Customer.ID, pPID, ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.LoadAPPLsOS(pTPID,pPID)", Array(pTPID, pPID)
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
    Me.ShowVAT = Customer.ShowVAT
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
Public Property Get NonVATDocument() As Boolean
    On Error GoTo errHandler
    NonVATDocument = mudtProps.NonVATDocument
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.NonVATDocument"
End Property
Public Property Let NonVATDocument(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.NonVATDocument = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.NonVATDocument(val)", val
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
Public Property Get Signature()
Signature = FNS(mudtProps.Signature)
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
Public Property Get TPNAME() As String
    On Error GoTo errHandler
    TPNAME = FNS(mudtProps.TPNAME)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPName"
End Property
Public Property Let TPNAME(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPNAME = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TPName(val)", val
End Property
Public Property Let IsPreDelivery(val As Boolean)
    On Error GoTo errHandler
    mudtProps.IsPreDelivery = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.IsPreDelivery(val)", val, EA_NORERAISE
End Property
Public Property Get IsPreDelivery() As Boolean
    IsPreDelivery = FNB(mudtProps.IsPreDelivery)
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
Public Property Get DOCCodeF() As String
    On Error GoTo errHandler
    Select Case mudtProps.Status
    Case 1
        DOCCodeF = "Voided " & FNS(mudtProps.DOCCode)
    Case 2
        If oPC.AllowsInvoicePicking Then
            DOCCodeF = "Picking slip in process " & FNS(mudtProps.DOCCode)
        Else
            If Me.Proforma = True Then
                DOCCodeF = "Proforma invoice in process " & FNS(mudtProps.DOCCode)
            Else
                DOCCodeF = "Invoice in process " & FNS(mudtProps.DOCCode)
            End If
        End If
    Case 3
        If oPC.AllowsInvoicePicking Then
            DOCCodeF = "Picking slip " & FNS(mudtProps.DOCCode)
        Else
            If Me.Proforma = True Then
                DOCCodeF = "Pro-forma invoice " & FNS(mudtProps.DOCCode)
            Else  'This should not occur as a tax invoice should always be status 4
                DOCCodeF = "Tax Invoice " & FNS(mudtProps.DOCCode)
            End If
        End If
    Case 4
        If Me.Proforma = True Then
            DOCCodeF = "Pro-forma invoice " & FNS(mudtProps.DOCCode)
        Else
            DOCCodeF = "Tax Invoice " & FNS(mudtProps.DOCCode)
        End If
    Case 5
        DOCCodeF = "Pro-forma invoice " & FNS(mudtProps.DOCCode)
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.DOCCodeF"
End Property
Public Property Get IsInvoice() As Boolean
    IsInvoice = True
End Property
Public Property Get IsGDN() As Boolean
    IsGDN = False
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
Public Property Get ProcessingDate() As Date
    On Error GoTo errHandler
    ProcessingDate = mudtProps.ProcessingDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ProcessingDate"
End Property
Public Property Get ProcessingDateF() As String
    On Error GoTo errHandler
    ProcessingDateF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ProcessingDateF"
End Property
Public Property Get ProcessingDateFF() As String
    On Error GoTo errHandler
    ProcessingDateFF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy Hh:Nn")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ProcessingDateFF"
End Property

Public Property Get CaptureDate() As Date
    On Error GoTo errHandler
    CaptureDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CaptureDate"
End Property
Public Property Get CaptureDateF() As String
    On Error GoTo errHandler
    CaptureDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CaptureDateF"
End Property
Public Property Let CaptureDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.CaptureDate(val)", val
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
Public Property Get QtyNonStandardVAT() As Long
Dim i As Long
Dim oIL As a_InvoiceLine

    i = 0
    For Each oIL In Me.InvoiceLines
        If oIL.VATRate <> oPC.Configuration.VATRate Then
            i = i + 1
        End If
    Next
    QtyNonStandardVAT = i
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
        
    If mudtProps.Proforma = True Then
        If mudtProps.Status = stVOID Then
            StatusForPrinting = "VOID"
        ElseIf mudtProps.Status = stInProcess Then
            StatusForPrinting = "IN PROCESS"
        ElseIf mudtProps.Status = stISSUED Then
            StatusForPrinting = ""
        ElseIf mudtProps.Status = stCOMPLETE Then
            StatusForPrinting = ""
        ElseIf mudtProps.Status = stPROFORMA Then
            StatusForPrinting = ""
        ElseIf mudtProps.Status = stCANCELLED Then
            StatusForPrinting = "CANCELLED"
        End If
    Else
        If mudtProps.Status = stVOID Then
            StatusForPrinting = "VOID"
        ElseIf mudtProps.Status = stInProcess Then
            StatusForPrinting = "IN PROCESS"
        ElseIf mudtProps.Status = stISSUED Then
            If oPC.AllowsInvoicePicking Then
                StatusForPrinting = "PICKED (Not finalized)"
            Else
                StatusForPrinting = ""
            End If
        ElseIf mudtProps.Status = stCOMPLETE Then
            StatusForPrinting = ""
        ElseIf mudtProps.Status = stPROFORMA Then
            StatusForPrinting = "PROFORMA"
        ElseIf mudtProps.Status = stCANCELLED Then
            StatusForPrinting = "CANCELLED"
        End If
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
        If oPC.AllowsInvoicePicking And Not Me.Proforma Then
            StatusF = "PICKING"
        Else
            StatusF = "ISSUED"
        End If
    ElseIf mudtProps.Status = stCOMPLETE Then
        If oPC.AllowsInvoicePicking And Not Me.Proforma Then
            StatusF = "ISSUED"
        Else
            StatusF = "COMPLETE"
        End If
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
Public Function SetForAttn(val As String) As Boolean
    If Len(val) > Len(mudtProps.ForAttn) Then
        Err.Raise 384
    End If
    If val <> mudtProps.ForAttn Then
        mudtProps.ForAttn = val
        SetDirty True
    End If
    SetForAttn = True
    Exit Function
End Function
Public Function SetMemo(val As String) As Boolean
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
Public Property Get ForAttn() As String
    On Error GoTo errHandler
    ForAttn = FNS(mudtProps.ForAttn)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ForAttn"
End Property

'Public Property Get Memo() As String
'    On Error GoTo ErrHandler
'    Memo = FNS(mudtProps.Memo)
'    Exit Property
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Invoice.Memo"
'End Property
Public Function GetInvoiceNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim strDefaultCo As String

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    If mudtProps.Proforma Then
        cmd.CommandText = "sp_GetNextProformaNo"
    Else
        If mudtProps.IsPreDelivery Then
            cmd.CommandText = "sp_GetNextPreDelivInvNo"
        Else
            cmd.CommandText = "sp_GetNextInvNo"
        End If
    End If
    cmd.CommandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , COMPID)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@NewCode", adVarChar, adParamOutput, 15)
    cmd.Parameters.Append prm
    cmd.execute
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
                        If .ServiceItem Then
                            strInvoice = strInvoice & .TitleAuthorPublisher & "  " & .PAfterDiscountExtF(bForeign)
                        Else
                            strInvoice = strInvoice & .CodeF & "  " & .Ref & "  " & .QtyF & "  " & Left(.TitleAuthorPublisher, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.DiscountPercent = 0, "", .DiscountPercentF) & "  " & .PAfterDiscountExtF(bForeign)
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
                    If .ServiceItem Then
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
            strInvoice = strInvoice & "subtotal    " & TotalAfterDiscountExtF(bForeign) & vbCrLf   'subtotal 3
            strInvoice = strInvoice & "Less VAT @ " & CStr(oPC.Configuration.VATRate) & "%    " & TotalVATF(bForeign) & vbCrLf
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
'Err_Handler:
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
            oPCL.WriteTag "TOTAL", "subtotal" & "|" & TotalAfterDiscountExtF(bForeign)
            oPCL.WriteTag "TOTAL", "Less VAT @ " & CStr(oPC.Configuration.VATRate) & "%" & "|" & TotalVATF(bForeign)
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
'Err_Handler:
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
                    If .ServiceItem Then
                        strLabels = strLabels & Chr(13) & .TitleAuthorPublisher
                        strValues = strValues & Chr(13) & .PAfterDiscountExtF(bForeign)
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
        strLabels = strLabels & Chr(13) & "Total qty items"
        strValues = strValues & Chr(13) & TotalQty
    Else  'overseas clients
        If Me.ShowVAT Then
            strLabels = strLabels & Chr(13) & "subtotal"
            strValues = strValues & Chr(13) & TotalAfterDiscountExtF(bForeign)
            strLabels = strLabels & Chr(13) & "Less VAT of"
            strValues = strValues & Chr(13) & TotalVATF(bForeign)
            strLabels = strLabels & Chr(13) & "Payable"
            strValues = strValues & Chr(13) & TotalPayableF(bForeign)
            strLabels = strLabels & Chr(13) & "Total qty items"
            strValues = strValues & Chr(13) & TotalQty
        Else
            strLabels = strLabels & Chr(13) & "Total"
            strValues = strValues & Chr(13) & TotalPayableF(bForeign)
            strLabels = strLabels & Chr(13) & "Total qty items"
            strValues = strValues & Chr(13) & TotalQty
        End If
    End If
    If TotalDeposit(bForeign) > 0 Then
        strLabels = strLabels & Chr(13) & "Less deposits paid"
        strValues = strValues & Chr(13) & TotalDepositF(bForeign)
        strLabels = strLabels & Chr(13) & "(balance owed ="
        strValues = strValues & Chr(13) & TotalBalanceAfterDeposits(bForeign) & ")"
        strLabels = strLabels & Chr(13) & "Total qty items"
        strValues = strValues & Chr(13) & TotalQty
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
                    If .ServiceItem Then
                        oPCL.ILSend "", "", "", .Title & .SpecialVATMsg, "", "", "", .PAfterDiscountExtF(bForeign), "", "", "", .Note
                    Else
                        oPCL.ILSend .CodeForExport, .Ref, .QtyF, .SpecialVATMsg & .Title, .MainAuthor, .Publisher, .PriceF(bForeign), .PAfterDiscountExtF(bForeign), .DepositF(bForeign), .DiscountPercentF, .PAfterDiscountExtNonVATF(bForeign), PackText(.Note)
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
Private Function FillExtras(pINV As a_Invoice, bForeign As Boolean, pDiscountExists As Boolean) As Integer
    On Error GoTo errHandler
Dim i As Integer
Dim iCnt As Integer

    iCnt = 0
    For i = 1 To pINV.InvoiceLines.Count
        If pINV.InvoiceLines(i).BottomOfDocument = True Then
            With pINV.InvoiceLines(i)
                If .ServiceItem Then
                    iCnt = iCnt + 1
                    oPCL.ILSendExtra .TitleAuthorPublisher, .PAfterDiscountExtF(bForeign)
                End If
            End With
        End If
    Next i
    FillExtras = iCnt
EXIT_Handler:
'Err_Handler:
'    MsgBox error
'    GoTo EXIT_Handler
   ' Resume
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.FillExtras(pInv,bForeign,pDiscountExists)", Array(pINV, bForeign, _
         pDiscountExists)
End Function



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
    ErrorIn "a_Invoice.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
    On Error GoTo errHandler
    Set colClassors = New Collection
    colClassors.Add "Missing customer", "TP"
    colClassors.Add "Note too short", "NOTE"
    colClassors.Add "Invalid discount", "DISCOUNT"
    colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
Dim s As String

    bValid = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If
    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function



Public Function Post(Optional pNewStatus As enStatus) As String
    On Error GoTo errHandler
Dim strResult As String
Dim oSM As z_StockManager
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.CommandTimeout = 0
    
    
    strResult = Me.ApplyEdit
    If strResult > "" Then  'probably a timeout or
       'Not necessary oPC.COShort.execute "ROLLBACK TRANSACTION"
        Post = strResult
        Exit Function
    End If
    
    
    Set oSM = New z_StockManager
    If pNewStatus = 0 Then
        pNewStatus = Me.Status
    End If
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
                Case stCOMPLETE, stISSUED
                    strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, pNewStatus)
                    If strResult > "" Then
                        oSM.RollbackTransaction oPC.COShort
                        Post = strResult
                        GoTo EXITH
                    End If
                  '  SetStatus stCOMPLETE
                Case stPROFORMA
                    strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, stCOMPLETE)
                    If strResult > "" Then
                        oSM.RollbackTransaction oPC.COShort
                        Post = strResult
                        GoTo EXITH
                    End If
            '        SetStatus stPROFORMA
                End Select
        Case stISSUED
                If oPC.AllowsInvoicePicking Then
                    Select Case pNewStatus
                    Case stCOMPLETE, stISSUED
                        strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, pNewStatus)
                        If strResult > "" Then
                            oSM.RollbackTransaction oPC.COShort
                            Post = strResult
                            GoTo EXITH
                        End If
                    End Select
                End If
        Case stCOMPLETE      'Complete
                Post = "ISSUED"
                If oPC.AllowsInvoicePicking Then
                    Select Case pNewStatus
                    Case stCOMPLETE, stISSUED
                        strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, pNewStatus)
                        If strResult > "" Then
                            oSM.RollbackTransaction oPC.COShort
                            Post = strResult
                            GoTo EXITH
                        End If
                    End Select
                End If
        Case stPROFORMA      'Issued PROFORMA
                Select Case Status
                Case stVOID
                    oSM.CancelInvoice Me
            '        SetStatus stVOID
                Case stInProcess
                    Post = "OR: was proforma, can't now be inprocess."
                Case stCOMPLETE
                    strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, stCOMPLETE)
                    If strResult > "" Then
                        oSM.RollbackTransaction oPC.COShort
                        Post = strResult
                        GoTo EXITH
                    End If
            '        SetStatus stCOMPLETE
                Case stPROFORMA
                    Post = "PROFORMA"
                End Select
        End Select
    End If
   ' oPC.COShort.execute "COMMIT TRANSACTION"
    oSM.CalculateInvoiceTableSummaryData Me.InvoiceID, 0
 '  oSM.AUTOGenerateReturnFromInvoice Me.InvoiceID
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXITH:
    Post = strResult
    Set oSM = Nothing
    Exit Function
errHandler:
    ErrPreserve
    Post = strResult
    
    If ErrMustStop Then Debug.Assert False: Resume
    RlsObjs oPC.COShort
    ErrorIn "a_Invoice.Post", , , oPC.COShort
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
    SetDirty True
    
End Sub
Public Property Get Proforma() As Boolean
    Proforma = mudtProps.Proforma
End Property

Public Sub InformLocalPOSdb()
Dim DBLocalConn As ADODB.Connection
    Set DBLocalConn = New ADODB.Connection
    DBLocalConn.Provider = "sqloledb"
    DBLocalConn.open "Data provider=sqloledb;Data Source=" & oPC.POSServername & ";Initial Catalog=PBKSFD;User Id=sa" & ";Password=" & "" & "; Network Library=dbmssocn;"
    DBLocalConn.CommandTimeout = 0
        DBLocalConn.execute "INSERT INTO tBRIDGE (INV_ID,INV_DOCCode,INV_DOCDate,INV_CustomerName,INV_Payable,INV_TPID,INV_VATVALUE) VALUES (" & Me.InvoiceID & ",'" & Me.DOCCode & "'," & Me.DOCDate & ",'" & Me.Customer.Name & "'," & Me.TotalPayableRaw(False) & "," & Me.Customer.ID & "," & Me.TotalVATRaw(False) & ")"
    DBLocalConn.Close
    Set DBLocalConn = Nothing
End Sub
Public Property Get SendersEmail()
    SendersEmail = FNS(mudtProps.StaffEmail)
End Property

Public Function ExportToXML(bBySeqNo As Boolean, pFilename As String, _
    Optional ByVal pRoundedUp As Boolean, _
    Optional DispatchMode As enTransmitType, Optional pQtyCopies As Integer, _
    Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, Optional QuickPrint As Boolean) As Boolean
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
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strTemplateName As String
Dim strExecutable As String
Dim bMultipleCOLReps As Boolean
Dim strTmpRep As String
Dim flgFirst As Boolean
Dim cnt As Integer
Dim lngNumberOfLines As Long

    If DispatchMode = enMail Then
        If oPC.UsesOutlookForINVEmail Then
            strWorkingFolder = oPC.LocalFolder & "TEMP\"
            strDispatchMethod = "Outlook"
        Else
            strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
            strDispatchMethod = "Emailing"
        End If
'    ElseIf DispatchMode = enPrint Then
'        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
'        strDispatchMethod = "Printing"
    ElseIf DispatchMode = enView Or DispatchMode = enPrint Then
        strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
        strDispatchMethod = ""
    Else
        strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
        strDispatchMethod = "Printing"
    End If
    
                        p 1
    bForeign = False
    If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
        bForeign = True
        CalculateTotal_Foreign
    End If
                        p 2
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
                        p 3
    Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
    If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
        ExportToXML = False
        Exit Function
    End If
                        p 4
    If pQtyCopies < 1 Then pQtyCopies = 1
    
    Set xMLDoc = New ujXML
    
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "IN_1"
            .chCreate "MessageType"
                .elText = "INVOICE"
            .elCreateSibling "MessageCreationDate"
                .elText = Format(Now(), "yyyymmddHHNN")
            .elCreateSibling "DestinationAddress"
                If oPC.TestMode Then
                    .elText = oPC.EmailFrom
                Else
                    If Customer.BillTOAddress Is Nothing Then
                        .elText = ""
                    Else
                        .elText = Customer.BillTOAddress.EMail
                    End If
                End If
            .elCreateSibling "TemplateName"
                If Me.Customer.InvoiceTemplateName > "" Then
                    .elText = Me.Customer.InvoiceTemplateName
                Else
                    .elText = "IN_DOC"
                End If
            strTemplateName = .elText
            
            .elCreateSibling "Columns"
            If bDiscountExists And bReferenceExists Then
                    .elText = "DISC_REF"
            ElseIf bDiscountExists Then
                    .elText = "DISC"
            ElseIf bReferenceExists Then
                    .elText = "REF"
            Else
                    .elText = ""
            End If
                        p 5
            .elCreateSibling "RecipientsEmail"
                If Not Me.Customer.BillTOAddress Is Nothing Then
                   .elText = Me.Customer.BillTOAddress.EMail
                    pDestinationEmailAddress = Me.Customer.BillTOAddress.EMail
                End If
            .elCreateSibling "SendersEmail"
                .elText = Me.SendersEmail
            .elCreateSibling "CopyCount"
                .elText = pQtyCopies
            .elCreateSibling "Printer"
                If Not oDC Is Nothing Then .elText = oDC.PrinterName
                '.elText = oDC.PrinterName
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailInvMsg
            .elCreateSibling "IsProforma"
                .elText = IIf(Proforma, "Y", "N")
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailInvMsg
            .elCreateSibling "DocCode"
                .elText = DOCCodeF
            .elCreateSibling "DocCodeRaw"
                .elText = DOCCode
           .elCreateSibling "StoreCode", True
                .elText = oPC.Configuration.DefaultStore.code
           .elCreateSibling "DocDate", True
                .elText = DocDateF
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName                  'oPC.Configuration.DefaultCompany.CompanyName
                                p 6
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "SenderAddressPostal", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).PostalAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerName", True
                .elText = Customer.Fullname
            .elCreateSibling "CustomerVATNumber", True
            If Customer.VatNumber > "" Then
                .elText = "Customer's VAT number: " & Customer.VatNumber
            Else
                .elText = ""
            End If
            .elCreateSibling "OurAcnowithCustomer", True
                .elText = Customer.OurACnoWithClient
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
                                p 7
            .elCreateSibling "CustomerPhone", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "")
                End If
            .elCreateSibling "CustomerFax", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
                End If
            .elCreateSibling "BillTo", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
                    If Me.ForAttn > "" Then
                        .elText = Me.ForAttn & vbCrLf & .elText
                    End If
                End If
            .elCreateSibling "DelTo", True
                If Me.DelToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
            .elCreateSibling "DispatchMode", True
                If Me.DelToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = FNS(Me.DelToAddress.Description)
                End If
                            p 8
            bMultipleCOLReps = False
            flgFirst = True
            lngNumberOfLines = 0
            For i = 1 To InvoiceLines.Count
                If InvoiceLines(i).BottomOfDocument = False Then
                    If oPC.AllowsSSInvoicing Then
                        If InvoiceLines(i).QtyFirm = 0 And InvoiceLines(i).QtySS = 0 Then
                            GoTo skip
                        End If
                    Else
                        If InvoiceLines(i).Qty = 0 Then
                            GoTo skip
                        End If
                    End If
                    lngNumberOfLines = lngNumberOfLines + 1
                    flgFirst = True
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    If oPC.HideLocalSKUonINV Then
                        .elText = InvoiceLines(i).CodeForExport
                    Else
                        .elText = InvoiceLines(i).CodeF
                    End If
                    .elCreateSibling "GDNCode", True
                        .elText = InvoiceLines(i).GDNCode
                    .elCreateSibling "Barcode", True
                        .elText = InvoiceLines(i).code
                    .elCreateSibling "Title", True
                        .elText = InvoiceLines(i).TitleAuthor
                    .elCreateSibling "Publisher", True
                        .elText = InvoiceLines(i).Publisher
                    .elCreateSibling "DistributorName", True
                        .elText = InvoiceLines(i).DistributorName
                    .elCreateSibling "DistributorAcno", True
                        .elText = InvoiceLines(i).DistributorAcno
                    If oPC.AllowsSSInvoicing Then
                        .elCreateSibling "QtyFirm", True
                            .elText = InvoiceLines(i).QtyFirmF
                        .elCreateSibling "QtySS", True
                            .elText = InvoiceLines(i).QtySSF
                    Else
                        .elCreateSibling "QtyFirm", True
                            .elText = InvoiceLines(i).QtyF
                    End If
                    .elCreateSibling "IsServiceItem", True
                        .elText = IIf(InvoiceLines(i).ServiceItem, "TRUE", "FALSE")
                    .elCreateSibling "Price", True
                        .elText = InvoiceLines(i).PriceF(bForeign)
                    .elCreateSibling "PriceExVat", True
                        .elText = InvoiceLines(i).PriceExVatF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = InvoiceLines(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = InvoiceLines(i).Ref
                    .elCreateSibling "CORep", True
                        .elText = InvoiceLines(i).CO_StaffShortname
                    If flgFirst = True Then
                        strTmpRep = .elText
                    Else
                        If .elText <> strTmpRep Then
                            bMultipleCOLReps = True
                        End If
                    End If
                    .elCreateSibling "Extension", True
                        .elText = InvoiceLines(i).PAfterDiscountExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = InvoiceLines(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = InvoiceLines(i).PAfterDiscountExtNonVATF(bForeign)
                    .elCreateSibling "VAT", True
                        .elText = InvoiceLines(i).PAfterDiscountLineValueVATF(bForeign)
                    .elCreateSibling "VATRate", True
                        .elText = InvoiceLines(i).VATRateF
                    .elCreateSibling "DiscountValue", True
                        .elText = InvoiceLines(i).PDiscExtF(bForeign)
                    .elCreateSibling "DiscountValueExVAT", True
                        .elText = InvoiceLines(i).PDiscExtExVATF(True)
                    .elCreateSibling "FCName", True
                        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
                            .elText = oPC.Configuration.Currencies.FindCurrencyByID(InvoiceLines(i).FCID).SYSNAME
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCPrice", True
                        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
                            .elText = InvoiceLines(i).ForeignPriceF
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCFactor", True
                        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
                            .elText = InvoiceLines(i).FCFactorInvF
                        Else
                            .elText = "n/a"
                        End If
                    .navUP
                    flgFirst = False
skip:
                End If
            Next i
                                p 9
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
                    .elCreateSibling "PriceExVAT", True
                        .elText = InvoiceLines(i).PriceExVatF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = InvoiceLines(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = InvoiceLines(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = InvoiceLines(i).PAfterDiscountExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = InvoiceLines(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = InvoiceLines(i).PAfterDiscountExtNonVATF(bForeign)
                    .elCreateSibling "VAT", True
                        .elText = InvoiceLines(i).PAfterDiscountLineValueVATF(bForeign)
                    .elCreateSibling "VATRate", True
                        .elText = InvoiceLines(i).VATRateF
                    .elCreateSibling "DiscountValue", True
                        .elText = InvoiceLines(i).PDiscExtF(bForeign)
                    .elCreateSibling "DiscountValueExVAT", True
                        .elText = InvoiceLines(i).PDiscExtExVATF
                    .elCreateSibling "FCName", True
                        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
                            .elText = oPC.Configuration.Currencies.FindCurrencyByID(InvoiceLines(i).FCID).Description
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCPrice", True
                        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
                            .elText = InvoiceLines(i).ForeignPriceF
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCFactor", True
                        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
                            .elText = InvoiceLines(i).FCFactorInvF
                        Else
                            .elText = "n/a"
                        End If
                    .navUP
                End If
            Next i
                                p 10
            .elCreateSibling "TotalNumberOfLines", True
                .elText = CStr(lngNumberOfLines)
            .elCreateSibling "QtyItems", True
                .elText = Me.TotalQtyF
                                p 11
            .elCreateSibling "Memo", True
                .elText = Memo
            .elCreateSibling "CompanyRegistration", True
                .elText = oPC.Configuration.DefaultCompany.CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.DefaultCompany.VatNumber
            .elCreateSibling "StaffMember", True
            If oPC.GetProperty("QuoteCOStaffNameOnInvoice") = "TRUE" Then
                If bMultipleCOLReps Then
                    .elText = Me.StaffName
                Else
                    .elText = IIf(strTmpRep > "", strTmpRep, "")
                End If
            Else
                .elText = Me.StaffName
            End If
            .elCreateSibling "OrderMessage", True
            .elText = oPC.Configuration.InvoiceText
            
            .elCreateSibling "TotalShipping", True
                .elText = Me.TotalServiceItemF(bForeign)
            .elCreateSibling "TotalDiscount", True
                .elText = Me.TotalDiscountF(bForeign)
            .elCreateSibling "TotalDiscountExVAT", True
                .elText = Me.TotalDiscountExVATF
            .elCreateSibling "TotalIncVAT", True
                .elText = Me.TotalExtF(bForeign)
            .elCreateSibling "TotalExVAT", True
                .elText = Me.TotalNonVATF(bForeign)
            .elCreateSibling "TotalVAT", True
                .elText = Me.TotalVATF(bForeign)
                
            .elCreateSibling "VATable", True
                .elText = IIf(Me.VATable, "TRUE", "FALSE")
            .elCreateSibling "ShowVAT", True
                .elText = IIf(Me.ShowVAT, "TRUE", "FALSE")
                
                
            .elCreateSibling "Totals", True
            If Me.VATable = False Then    'FOREIGN
                    If iExtrasCount > 0 Or Me.ShowVAT Then
                        .chCreate "TotalDescription"
                            .elText = "subtotal" & vbCrLf
                        .elCreateSibling "NumericValue", True
                            .elText = Me.TotalExtWithOutExtrasF(bForeign) & vbCrLf
                        .navUP
                        If Me.ShowVAT Then
                            .elCreateSibling "Totals", True
                            .chCreate "TotalDescription"
                                .elText = "less VAT at " & CStr(oPC.Configuration.VATRate) & "%" & vbCrLf
                            .elCreateSibling "NumericValue", True
                                .elText = Me.TotalVATF(bForeign) & vbCrLf
                            .navUP
                        End If
                    End If
                    If iExtrasCount > 0 Then  'Doing this here again so as not to interfere with established procedure above
                        .elCreateSibling "Totals", True
                        .chCreate "TotalDescription"
                            .elText = "subtotal" & vbCrLf
                        .elCreateSibling "NumericValue", True
                            .elText = Me.Total_nonExtras_ExVATF(bForeign) & vbCrLf
                        .navUP
                        For i = 1 To InvoiceLines.Count
                            If InvoiceLines(i).BottomOfDocument = True Then
                                .elCreateSibling "Totals", True
                                .chCreate "TotalDescription"
                                    .elText = InvoiceLines(i).TitleAuthorPublisher & vbCrLf
                                .elCreateSibling "NumericValue", True
                                    .elText = InvoiceLines(i).PriceF(bForeign) & vbCrLf
                                .navUP
                        
                            End If
                        Next i
                    End If
                        .elCreateSibling "Totals", True
                        .chCreate "TotalDescription"
                            .elText = "Total" & vbCrLf
                        .elCreateSibling "NumericValue", True
                            .elText = Me.TotalPayableF(bForeign) & vbCrLf
                        .navUP
                    
            Else
                    If iExtrasCount > 0 Then
                        .chCreate "TotalDescription"
                            .elText = "subtotal" & vbCrLf
                        .elCreateSibling "NumericValue", True
                            .elText = Me.TotalExtWithOutExtrasF(bForeign) & vbCrLf
                        .navUP
                        For i = 1 To InvoiceLines.Count
                            If InvoiceLines(i).BottomOfDocument = True Then
                                .elCreateSibling "Totals", True
                                .chCreate "TotalDescription"
                                    .elText = InvoiceLines(i).TitleAuthorPublisher & vbCrLf
                                .elCreateSibling "NumericValue", True
                                    .elText = InvoiceLines(i).PriceF(bForeign) & vbCrLf
                                .navUP

                            End If
                        Next i
                    End If
                        .elCreateSibling "Totals", True
                        .chCreate "TotalDescription"
                            .elText = "Total" & vbCrLf
                        .elCreateSibling "NumericValue", True
                            .elText = Me.TotalPayableF(bForeign) & vbCrLf
                        .navUP
                        .elCreateSibling "Totals", True
                        .chCreate "TotalDescription"
                            .elText = "(Includes VAT of " & oPC.Configuration.VATRateF & vbCrLf
                        .elCreateSibling "NumericValue", True
                            .elText = Me.TotalVATF(bForeign) & ")" & vbCrLf
                        .navUP
            End If
    End With
                                p 12

'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "IN_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        ''''''''''''''''''''''''
       'Use this if trying to print locally for a remote store (where stores resources are too small for an extremely large PDF)
        strXML = "C:\pbks\txt.xml"
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
       ''''''
       '''''
       .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

                                p 13

    If DispatchMode = enMail Then
        If oPC.EmailINVShowHTML = True Then
            'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.validateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & fs.GetBaseName(strTemplateName) & "_HTML.xslt"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
            If fs.FileExists(strWorkingFolder & "IN_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "IN_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "IN_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage
            oTF.CloseTextFile
            Set oTF = Nothing
        Else
        End If
        
    End If
    
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''USE this when producing a local PDF from a remote store'''''''''''''''''''''''''''''''''''''''''''
     '  strWorkingFolder = "C:\PBKS\PDF\"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        strFOFile = strWorkingFolder & "IN_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "IN_" & Me.DOCCode & ".PDF"
        
        
        
        

    If (Me.Customer.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
        If oPC.GetProperty("UseXalan") <> "TRUE" Then
            '====================================STANDARD METHOD -tenporarily replaces by XALAN method===========================
            Set objXSL = Nothing
            Set objXSL = New MSXML2.DOMDocument30
            objXSL.async = False
            objXSL.validateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
                            p 14
            Set opXMLDOC = New MSXML2.DOMDocument30
            opXMLDOC.async = False
            opXMLDOC.validateOnParse = False
            opXMLDOC.resolveExternals = False
            xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
    
                            p 15
            docWriteTostream strFOFile, opXMLDOC, "UNICODE"
            '====================================================================================================================
        Else
             strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
             strCommand = oPC.SharedFolderRoot & "\Executables\PBKS_XALAN.BAT" & " " & strXML & " " & strPath & " " & strFOFile & " " & oPC.SharedFolderRoot
            'strCommand = oPC.SharedFolderRoot & "\Executables\LOGBATCH.BAT " & oPC.SharedFolderRoot & "\Executables\PBKS_XALAN.BAT" & " " & strXML & " " & strPath & " " & strFOFile & " " & oPC.SharedFolderRoot & " xalan-j_2_7_1"
             F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
        End If
'====================================================================================================================
'Stage 2 Convert the .FO file to .PDF and clean up
 '       strCommand = oPC.SharedFolderRoot & "\Executables\FOP\FOP.BAT" & " " & strFOFile & " " & strPDFFile
       ' strCommand = oPC.SharedFolderRoot & "\Executables\LOGBATCH.BAT " & oPC.SharedFolderRoot & "\Executables\PBKS_FOP.BAT" & " " & strFOFile & " " & strPDFFile
                        p 155
       ' ChDrive fs.GetDrive(oPC.LocalFolder)
                        
        ChDir "\PBKS\Executables\FOP\"
        strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
     'Allow 2 minutes to produce the PDF file
      '  strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
        'so the FO file is sometimes deleted before the .PDF file is produced.
        cnt = 0
                        p 156
        Do While Not fs.FileExists(strPDFFile) And cnt < 40
            MsgWaitObj 1000
            cnt = cnt + 1
        Loop
                        p 157
        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
                        p 158
        If fs.FileExists(strPDFFile) Then
            On Error Resume Next
            SetFileProperty strPDFFile, "Title", Me.Customer.AcNo
            SetFileProperty strPDFFile, "Author", Me.Customer.Name
            On Error GoTo errHandler

        End If
    End If
                        p 16
      
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    pFilename = strPDFFile
    
    If DispatchMode <> enMail And DispatchMode <> enView Then
        oSM.LogTransmission Me.InvoiceID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    End If
    Set oSM = Nothing
    Set oFSO = Nothing
                        p 9
    oTF.CloseTextFile
    Set oTF = Nothing
    ExportToXML = True
    '
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
    ErrorIn "a_Invoice.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
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
    oTF.WriteToTextFile "Status" & vbTab & Me.Status
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCodeF
    oTF.WriteToTextFile "DocDate" & vbTab & DocDateF
    oTF.WriteToTextFile "StoreCode" & vbTab & oPC.Configuration.DefaultStore.code
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "SenderAddressPostal" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).PostalAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "CustomerName" & vbTab & Customer.Fullname
    oTF.WriteToTextFile "CustomerVATNumber" & vbTab & Customer.VatNumber
    oTF.WriteToTextFile "ACNO" & vbTab & Customer.AcNo
    oTF.WriteToTextFile "OurAcnowithCustomer" & vbTab & Customer.OurACnoWithClient
    
    oTF.WriteToTextFile_NoLineTerminator "CustomerWithAddress" & vbTab
    If Not BillTOAddress Is Nothing Then
        oTF.WriteToTextFile Replace(Me.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
    End If
    
    oTF.WriteToTextFile_NoLineTerminator "CustomerPhone" & vbTab
    If Not BillTOAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "")
    End If
    oTF.WriteToTextFile_NoLineTerminator "CustomerFax" & vbTab
    If Not BillTOAddress Is Nothing Then
        oTF.WriteToTextFile IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
    End If

                    If Not BillTOAddress Is Nothing Then
                        s = ForAttn & " " & Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
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

                    If Not DelToAddress Is Nothing Then
                        s = FNS(Me.DelToAddress.Description)
                    Else
                        s = ""
                    End If
    oTF.WriteToTextFile "DispatchMode" & vbTab & s



    s = "SKU" & vbTab & "Barcode" & vbTab & "GDNCode" & vbTab & "Title" & vbTab & "Publisher" & vbTab & "Distributor" & vbTab & "DistributorAcno" _
    & vbTab & IIf(oPC.AllowsSSInvoicing, "QtyFirm" & vbTab & "QtySS", "Qty") _
    & vbTab & "Price" & vbTab & "PriceExVat" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab & "CORep" & vbTab & "Note" & vbTab & "Extension" & vbTab & "ExtensionExVat" _
    & vbTab & "VAT" & vbTab & "VATRate" & vbTab & "DiscountValue" & vbTab & "DiscountValueExVAT" & vbTab & "FCName" & vbTab & "FCPrice" & vbTab & "FCFactor"
    
    oTF.WriteToTextFile s
               
  lngNumberOfLines = 0
  For i = 1 To InvoiceLines.Count
    lngNumberOfLines = lngNumberOfLines + 1
    If oPC.HideLocalSKUonINV Then
        s2 = InvoiceLines(i).CodeForExport
    Else
        s2 = InvoiceLines(i).CodeF
    End If
    s = s2
    s = s & vbTab & InvoiceLines(i).code
    s = s & vbTab & InvoiceLines(i).GDNCode
    s = s & vbTab & InvoiceLines(i).TitleAuthor
    s = s & vbTab & InvoiceLines(i).Publisher
    s = s & vbTab & InvoiceLines(i).DistributorName
    s = s & vbTab & InvoiceLines(i).DistributorAcno
                    If oPC.AllowsSSInvoicing Then
                            s2 = InvoiceLines(i).QtyFirmF & vbCrLf & InvoiceLines(i).QtySSF
                    Else
                            s2 = InvoiceLines(i).QtyF
                    End If
    s = s & vbTab & s2
    s = s & vbTab & InvoiceLines(i).PriceF(bForeign)
    s = s & vbTab & InvoiceLines(i).PriceExVatF(bForeign)
    s = s & vbTab & InvoiceLines(i).DiscountPercentF
    s = s & vbTab & InvoiceLines(i).Ref
    s = s & vbTab & InvoiceLines(i).CO_StaffShortname
    s = s & vbTab & InvoiceLines(i).Note
    s = s & vbTab & InvoiceLines(i).PAfterDiscountExtF(bForeign)
    s = s & vbTab & InvoiceLines(i).PAfterDiscountExtNonVATF(bForeign)
    s = s & vbTab & InvoiceLines(i).PAfterDiscountLineValueVATF(bForeign)
    s = s & vbTab & InvoiceLines(i).VATRateF
    s = s & vbTab & InvoiceLines(i).PDiscExtF(bForeign)
    s = s & vbTab & InvoiceLines(i).PDiscExtExVATF(True)
        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
            s2 = oPC.Configuration.Currencies.FindCurrencyByID(InvoiceLines(i).FCID).SYSNAME
        Else
            s2 = "n/a"
        End If
    s = s & vbTab & s2
        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
            s2 = InvoiceLines(i).ForeignPriceF
        Else
            s2 = "n/a"
        End If
    s = s & vbTab & s2
        If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
            s2 = InvoiceLines(i).FCFactorInvF
        Else
            s2 = "n/a"
        End If
    s = s & vbTab & s2
    
    oTF.WriteToTextFile s
Next

    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(lngNumberOfLines)
    oTF.WriteToTextFile "QtyItems" & vbTab & TotalQtyF
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    
    
    oTF.WriteToTextFile "TotalShipping" & vbTab & TotalServiceItemF(bForeign)
    oTF.WriteToTextFile "TotalDiscount" & vbTab & TotalDiscountF(bForeign)
    oTF.WriteToTextFile "TotalDiscountExVAT" & vbTab & TotalDiscountExVATF
    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalExtF(bForeign)
    oTF.WriteToTextFile "TotalExVAT" & vbTab & TotalNonVATF(bForeign)
    oTF.WriteToTextFile "TotalVAT" & vbTab & TotalVATF(bForeign)
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
End Function



Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument30, _
                Optional ByVal CharSet As String = "UNICODE")
    On Error GoTo errHandler
    Dim s As Object
    Set s = CreateObject("ADODB.Stream")
    With s
        If CharSet <> "" Then .CharSet = CharSet
        .open
        .WriteText obj.xml
        .SaveToFile FilePath, 2 'adSaveCreateOverWrite
        .Close
    End With
    Exit Sub
errHandler:
    ErrorIn "a_Invoice.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub
'Private Sub CreateOutlookMail(pstrXML As String)
'Dim myOutlook As Object
'Dim myMailItem As Object
'Dim xmlFile As ujXML
'Dim strXMLFile As String
'Dim strAddress As String
'Dim strSendersEmail As String
'Dim tmp As String
'Dim strReference As String
'
'    tmp = fs.GetBaseName(strXML)
'    strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_") - 1)
'    strReference = strReference & ": " & oPC.InvoiceSubject
'    Set xmlFile = New ujXML
'    xmlFile.docReadFromFile strXML, "UNICODE"
'    xmlFile.navTop
'        If xmlFile.docXslLocate("DestinationAddress") Then
'            strAddress = xmlFile.Element.Text
'        Else
'            strAddress = strEmailFrom
'        End If
'        xmlFile.navTop
'        If xmlFile.docXslLocate("SendersEmail") Then
'            strSendersEmail = xmlFile.Element.Text
'        Else
'            strSendersEmail = strEmailFrom
'        End If
'        If strSendersEmail = "" Then strSendersEmail = strEmailFrom
'        xmlFile.navTop
'        If xmlFile.docXslLocate("StaffMember") Then
'            strSender = xmlFile.Element.Text
'        Else
'            strSender = strEmailFrom
'        End If
'        If strSender = "" Then strSender = strEmailFrom
'        xmlFile.navTop
'        If strSender = "" Then strSender = strEmailFrom
'        xmlFile.navTop
'        If xmlFile.docXslLocate("DocCode") Then
'            strDocCode = xmlFile.Element.Text
'        Else
'            strDocCode = "Unknown"
'        End If
'
'    Set myOutlook = CreateObject("Outlook.Application")
'    Set myMailItem = myOutlook.CreateItem(0)
'    myMailItem.Recipients.Add strAddress
'    ' Set subject
'    myMailItem.Subject = strReference
'    myMailItem.Body = "Quick test!"
'    ' And send it!
'    myMailItem.send
'    ' Close instance
'    Set myOutlook = Nothing
'End Sub
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
Public Property Get ExchangeID() As String
    On Error GoTo errHandler
    ExchangeID = Trim(mudtProps.ExchangeID)
    Exit Property
errHandler:
    ErrorIn "a_Invoice.ExchangeID"
End Property

Public Property Get TotalQty() As Long
    On Error GoTo errHandler

        TotalQty = mudtProps.TotalQty

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalQty"
End Property
Public Function TotalQtyPerProduct(PID As String) As Long
Dim oIL As a_InvoiceLine
Dim lngTot As Long

    lngTot = 0
    For Each oIL In Me.InvoiceLines
        If oIL.PID = PID Then
            lngTot = lngTot + oIL.Qty
        End If
    Next
    TotalQtyPerProduct = lngTot
End Function

Public Property Get TotalQtyF() As String
    On Error GoTo errHandler

        TotalQtyF = CStr(TotalQty)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.TotalQtyF"
End Property

Private Sub SetupDocumentTotals(bForeign As Boolean)
    On Error GoTo errHandler
Dim i, k As Integer
Dim bBottomOfDocument As Boolean

    ReDim arDocTotals(16)
    ReDim arDocTotalsF(16)
    
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
            For k = 1 To InvoiceLines.Count
                If InvoiceLines(k).BottomOfDocument = True And InvoiceLines(k).VATRate = 0 Then
                    addARow InvoiceLines(k).Title, InvoiceLines(k).ForeignPrice, InvoiceLines(k).PriceF(bForeign), "", "+"
                    bBottomOfDocument = True
                End If
            Next k
            If bBottomOfDocument = True Or ShowVAT = True Then arDocTotals(jTotal - 1).TotalType = "GRAND"
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

Public Sub PasteLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, _
                    VATRate As Double, Ref As String, ExtraChargePID As String, ExtraChargeValue As Long, _
                    FCPrice As Long, FCFactor As Double, FCID As Long)
Dim oSM As New z_StockManager

    oSM.PasteInvoiceLine PID, QtyFirm, QtySS, Price, DiscountRate, VATRate, mudtProps.InvoiceID, Ref, ExtraChargePID, ExtraChargeValue, FCPrice, FCFactor, FCID
    
End Sub


Public Function GenerateEDIMsg() As Boolean
    On Error GoTo errHandler
Dim str As String
Dim strMsg As String
Dim dte As Date
Dim oTF As New z_TextFile
Dim strFilePath As String
Dim strFilename As String
Dim i As Integer
Dim oSM As z_StockManager
Dim oFSO As New FileSystemObject

    dte = Now()
    strFilename = Me.DOCCode '& Format(dte, "HHNN")
    
    strFilePath = oPC.SharedFolderRoot & "\EDI_Out\Invoices\" & oFSO.GetBaseName(strFilename) & ".EDI"
    
    oTF.OpenTextFile strFilePath
    strMsg = "UNA:+.? '"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    
    strMsg = "UNB+UNOA:2+" '& oPC.Configuration.GFXNumber & "+" & Me.Supplier.GFXNumber & "+" & Format(dte, "YYYYMMDD") & ":" & Format(dte, "HHNN") & "+" & strFileName & "+PASSWORD+ORDERS"
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & vbCrLf
    strMsg = "UNH++INVOIC"
    oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & vbCrLf
    strMsg = "BGM+380:::+" & Me.DOCCode & "'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    strMsg = "DTM+137:" & Format(Me.DOCDate, "YYYYMMDD") & ":102'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    
    For i = 1 To Me.InvoiceLines.Count
        strMsg = "LIN+" & Format(i, "000000") & "'"
        oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
        If InvoiceLines(i).FCID > 0 Then
            strMsg = "IMD+L+050+:::" & Replace(Left(Me.InvoiceLines(i).Title, 35), "'", "?'") & ":" & InvoiceLines(i).FCFactorInvF & "/" & InvoiceLines(i).ForeignPriceF & "'"
        Else
            strMsg = "IMD+L+050+:::" & Replace(Left(Me.InvoiceLines(i).Title, 35), "'", "?'") & "'"
        End If
        oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
        strMsg = "QTY+47:" & CStr(Me.InvoiceLines(i).Qty) & "'"
        oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
        strMsg = "MOA+203:" & CStr(Me.InvoiceLines(i).PAfterDiscountExtNonVAT_withDecimal(False)) & "'"
        oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
        strMsg = "RFF+LI:" & ClearOrderRef(CStr(Me.InvoiceLines(i).Ref)) & "'"   ' & "/" & CStr(Me.InvoiceLines(i).ForeignPriceF)
        oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
        
    Next
    strMsg = "UNS+S'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    strMsg = "CNT+2:" & CStr(Me.InvoiceLines.Count) & "'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    strMsg = "MOA+86:" & CStr(Me.TotalPayable_WithDecimal(False)) & "'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    strMsg = "ALC+C++++TX'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    strMsg = "MOA+8:" & CStr(Me.TotalVAT_WithDecimal(False)) & "'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
'
    strMsg = "UNT+56+00000000001" & "'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf

    strMsg = "UNZ+1+" & strFilename & "'"
    oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
    
'    Set oSM = New z_StockManager
'    oSM.LogTransmission Me.TRID, "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf
'    Log = "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
'    Set oSM = Nothing
    
    oTF.CloseTextFile
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.GenerateEDIMsg"
End Function
Private Function ClearOrderRef(s As String) As String
    If InStr(1, s, ",") > 0 Then
        s = Left(s, InStr(1, s, ",") - 1)
    End If
    ClearOrderRef = s
End Function
'Public Property Get Log() As String
'    On Error GoTo errHandler
'    Log = Trim(mudtProps.Log)
'    Exit Property
'errHandler:
'    ErrorIn "a_PO.Log"
'End Property
'
'Public Property Let Log(val As String)
'    On Error GoTo errHandler
'    mudtProps.Log = Trim(val)
'    Exit Property
'errHandler:
'    ErrorIn "a_PO.Log(val)", val
'End Property

