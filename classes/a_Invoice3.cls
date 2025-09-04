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
Dim lngLPAfterDiscountExt_tot As Double
Dim lngFPAfterDiscountExt_tot As Double
Dim lngLDeposit_tot As Double
Dim lngFDeposit_tot As Double
Dim lngVAT_tot As Double
Dim lngDisc_tot As Double
Dim lngDiscExVAT_tot As Double

Dim TG As Double  'Total Gross of discount
Dim TN As Double  'Total Nett Of discount
Dim TNN As Double 'Total Nett of discount, Nett of VAT
Dim d As Double  'Discount
Dim DN As Double  'Discount Nett of VAT
Dim v As Double  'VAT

Dim lngLPayableLessVAT_tot As Double
Dim lngFPayableLessVAT_tot As Double
Dim lngTotalServiceItem As Double
Dim lngTotalExtras As Double
Dim lngQty_tot As Double
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Dim iExtrasCount As Integer
Private xMLDoc As ujXML
Event Valid(pMsg As String)
Event TotalChange(lngTotalExt As Double, lngDeposit_tot As Double, lngVAT_tot As Double)
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
Public Sub ConvertToNonStandardVATInvoice(v As Double)
      Dim oIL As a_InvoiceLine
     For Each oIL In Me.InvoiceLines
        If oIL.VATRate <> 0 Then
            oIL.VATRate = v
        End If
     Next
     ApplyEdit
     BeginEdit
     Me.CalculateTotals
End Sub


Public Sub ConvertToNonVATInvoice()
      Dim oIL As a_InvoiceLine
17970     For Each oIL In Me.InvoiceLines
17980         oIL.Price = Round(oIL.Price / ((100 + oIL.VATRate) / 100))
17990     Next
18000     Me.CalculateTotals
End Sub

Public Property Get constDOCCODE() As String
18010     constDOCCODE = CDOCCODE
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
18020     On Error GoTo errHandler
18030     RaiseEvent CurrRowStatus(pMsg)
18040     Exit Sub
errHandler:
18050     If ErrMustStop Then Debug.Assert False: Resume
18060     ErrorIn "a_Invoice.mColInvoiceLines_CurrRowStatus(pMsg)", pMsg
End Sub
Public Property Get ForeignCurrency() As a_Currency
18070     Set ForeignCurrency = oCurrency_Foreign
End Property
Private Sub mcolinvoiceLines_RowsChange()
18080     On Error GoTo errHandler
      'occurs for new rows and deletion of rows
          
18090     CalculateTotals
18100     RaiseEvent TotalChange(lngLPAfterDiscountExt_tot, lngLDeposit_tot, lngVAT_tot)
18110     SetDirty True
18120     Exit Sub
errHandler:
18130     If ErrMustStop Then Debug.Assert False: Resume
18140     ErrorIn "a_Invoice.mcolinvoiceLines_RowsChange"
End Sub
Private Sub mcolInvoiceLines_RowCOunt(cnt As Integer)
18150     RaiseEvent RowCount(cnt)
End Sub
Public Function RecalculateAllLines()
18160     On Error GoTo errHandler
      Dim oIL As a_InvoiceLine
18170     For Each oIL In InvoiceLines
18180         oIL.CalculateLine
18190     Next
18200     Exit Function
errHandler:
18210     If ErrMustStop Then Debug.Assert False: Resume
18220     ErrorIn "a_Invoice.RecalculateAllLines"
End Function
Public Function CalculateTotals()
18230     On Error GoTo errHandler
18240     CalculateTotal_Local
18250     CalculateTotal_Foreign
18260     Exit Function
errHandler:
18270     If ErrMustStop Then Debug.Assert False: Resume
18280     ErrorIn "a_Invoice.CalculateTotals"
End Function
Public Function CalculateTotal_Local_X()
18290     On Error GoTo errHandler
      Dim oIL As a_InvoiceLine
      Dim bLineDiscountApplied As Boolean
18300     mobjValid.RuleBroken "OVERFLOW", False
18310     bLineDiscountApplied = False
18320     lngLPAfterDiscountExt_tot = 0
18330     lngLDeposit_tot = 0
18340     lngVAT_tot = 0
18350     lngDisc_tot = 0
18360     lngDiscExVAT_tot = 0
18370     lngLPayableLessVAT_tot = 0
18380     lngTotalServiceItem = 0
18390     iExtrasCount = 0
          'NOTE Some items on the invoice might be ServiceItem e.g. Insurance, Postage and so on. These items are non subject to any discount
          'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
          'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
          'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
          'calculated afterwards. Remember VAT rate can vary per line.
          'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
          'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
18400     iExtrasCount = 0
18410     lngTotalExtras = 0
18420     lngQty_tot = 0
18430     For Each oIL In mColInvoiceLines
18440         If (Not oIL.IsDeleted) Then
          '    If oIL.Fulfilled <> "CAN" Then
18450             oIL.CalculateLine
18460             If oIL.BottomOfDocument Then
18470                 iExtrasCount = iExtrasCount + 1
18480             End If
18490             If oIL.ServiceItem = False Then
18500                 lngQty_tot = lngQty_tot + oIL.Qty
18510             End If
18520             lngLPayableLessVAT_tot = lngLPayableLessVAT_tot + Round(oIL.PAfterDiscountExtNonVAT(False), 0)
18530             lngVAT_tot = lngVAT_tot + oIL.PAfterDiscountLineValueVAT(False)
18540             lngDisc_tot = lngDisc_tot + oIL.PDiscExt(False)
18550             lngDiscExVAT_tot = lngDiscExVAT_tot + oIL.PDiscExtExVAT
18560             If oIL.ServiceItem = True Then
18570                 lngTotalServiceItem = lngTotalServiceItem + oIL.PAfterDiscountExt(False)
18580             End If
18590             lngLPAfterDiscountExt_tot = lngLPAfterDiscountExt_tot + oIL.PAfterDiscountExt(False)
18600             lngLDeposit_tot = lngLDeposit_tot + oIL.Deposit
18610             If oIL.BottomOfDocument Then
18620                 lngTotalExtras = lngTotalExtras + oIL.PAfterDiscountExt(False)
18630             End If
18640         End If
18650     Next
18660     mudtProps.TotalQty = lngQty_tot
18670     mudtProps.TotalExtras = lngTotalExtras
18680     mudtProps.DepositPaid = lngLDeposit_tot
18690     mudtProps.TotalDiscount = lngDisc_tot
18700     mudtProps.TotalDiscountExVAT = lngDiscExVAT_tot
18710     mudtProps.TotalNonVAT = lngLPayableLessVAT_tot
18720     mudtProps.TotalVAT = lngVAT_tot
18730     mudtProps.TotalExtension = lngLPAfterDiscountExt_tot
18740     mudtProps.TotalPayable = lngLPAfterDiscountExt_tot - lngLDeposit_tot
18750     mudtProps.TotalServiceItem = lngTotalServiceItem

18760     If (Not VATable) And ShowVAT Then
18770         mudtProps.TotalPayable = lngLPAfterDiscountExt_tot - lngVAT_tot
18780     End If
18790     Exit Function
errHandler:
18800     ErrPreserve
18810     If Err = 6 Then 'ov erflow
18820         mobjValid.RuleBroken "OVERFLOW", True
18830         Clear
18840         Resume Next
18850     End If
18860     If ErrMustStop Then Debug.Assert False: Resume
18870     ErrorIn "a_Invoice.CalculateTotal_Local"
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
    TNN = 0
    lngTotalServiceItem = 0
    iExtrasCount = 0
    'NOTE Some items on the invoice might be ServiceItem e.g. Insurance, Postage and so on. These items are non subject to any discount
    'and they may mut probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    
'Copy of declarations at top for easy readability here---------------------
'Dim TG As Double  'Total Gross of discount
'Dim TN As Double  'Total Nett Of discount
'Dim TNN As Double 'Total Nett of discount, Nett of VAT
'Dim D As Double  'Discount
'Dim DN As Double  'Discount Nett of VAT
'Dim V As Double  'VAT
Dim TN_Vatable As Long 'Total Nett of discount for VATable items only
'--------------------------------------------------------------------------
    TN_Vatable = 0
    iExtrasCount = 0
    lngTotalExtras = 0
    lngQty_tot = 0
      

      TG = 0
      TN = 0
      TNN = 0
      d = 0
      DN = 0
      v = 0
      TN_Vatable = 0
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
            If oIL.VATRate > 0 Then  'assuming only VATRate is the standard one else un-VATted item
                TN_Vatable = TN_Vatable + (oIL.row_LTN)  'sum all the VATABLE items
            End If
            TNN = TNN + oIL.row_LTNN
            v = v + oIL.row_LV
            d = d + oIL.row_LDN
            DN = DN + oIL.row_LDN
            TN = TN + oIL.row_LTN
            If oIL.ServiceItem = True Then
                lngTotalServiceItem = lngTotalServiceItem + oIL.row_LTN
            End If
            lngLDeposit_tot = lngLDeposit_tot + oIL.Deposit
            If oIL.BottomOfDocument Then
                lngTotalExtras = lngTotalExtras + oIL.row_LTN
            End If
        End If
    Next
  'Recalculate VAT overall
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtras = lngTotalExtras
    mudtProps.DepositPaid = lngLDeposit_tot
    mudtProps.TotalDiscount = d
    mudtProps.TotalDiscountExVAT = DN
  'Recalculate VAT overall
    Dim FinalVAT As Double
    FinalVAT = Round(TN_Vatable - (CDbl(TN_Vatable) * 100 / (CDbl(100) + oPC.Configuration.VATRate)))
    mudtProps.TotalVAT = FinalVAT
    mudtProps.TotalNonVAT = TN - FinalVAT  'Calculate the non-Vat value by subtracting the stated Final VAT value from the total extensions
    mudtProps.TotalVATLineSummed = v
    mudtProps.VATRoundingAdjustment = v - FinalVAT
    mudtProps.TotalExtension = TN
    mudtProps.TotalPayable = mudtProps.TotalExtension - lngLDeposit_tot
    mudtProps.TotalServiceItem = lngTotalServiceItem

    
    If Customer.ShowVAT = False And Customer.VATable = False Then  'We base all calculations on an ex-VAT price
        mudtProps.TotalPayable = mudtProps.TotalExtension - lngLDeposit_tot
     End If
    If Customer.ShowVAT = True And Customer.VATable = False Then
        mudtProps.TotalPayable = mudtProps.TotalNonVAT
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
18880     On Error GoTo errHandler
      Dim oIL As a_InvoiceLine
      Dim bLineDiscountApplied As Boolean
18890     mobjValid.RuleBroken "OVERFLOW", False
18900     bLineDiscountApplied = False
18910     lngFPAfterDiscountExt_tot = 0
18920     lngFDeposit_tot = 0
18930     lngVAT_tot = 0
18940     lngFPayableLessVAT_tot = 0
18950     lngDisc_tot = 0
18960     lngTotalExtras = 0
18970     lngQty_tot = 0
18980     lngTotalServiceItem = 0
18990     iExtrasCount = 0


''=================================================
Dim TN_Vatable As Long 'Total Nett of discount for VATable items only
'--------------------------------------------------------------------------
    TN_Vatable = 0
    iExtrasCount = 0
    lngTotalExtras = 0
    lngQty_tot = 0
      
      TG = 0
      TN = 0
      TNN = 0
      d = 0
      DN = 0
      v = 0
      TN_Vatable = 0
    For Each oIL In mColInvoiceLines
        If (Not oIL.IsDeleted) Then
    '    If oIL.Fulfilled <> "CAN" Then
            oIL.CalculateLine
            'Count items to be placed at bottom of document on presentation
            If oIL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            
            'Include qty in total qty so long as it is an actual countable item ( not insurance or suchlike
            If oIL.ServiceItem = False Then
                lngQty_tot = lngQty_tot + oIL.Qty
            End If
            
            'Accumulate total vatable value
            If oIL.VATRate > 0 Then  'assuming only VATRate is the standard one else un-VATted item
                TN_Vatable = TN_Vatable + (oIL.row_LTN)  'sum all the VATABLE items
            End If
            
            'Accumulate total Nett of discount, Nett of VAT
            TNN = TNN + oIL.row_LTNN
            
            'Accumulate total VAT value
            v = v + oIL.row_LV
            
            'Accumulate total discount
            d = d + oIL.row_LD
            
            'Accumulate total discount nett of VAT
            DN = DN + oIL.row_LDN
            
            'Accumulate line total including VAT (total Nett Of discount)
            TN = TN + oIL.row_LTN
            
            
            If oIL.ServiceItem = True Then
                lngTotalServiceItem = lngTotalServiceItem + oIL.row_LTN
            End If
            lngLDeposit_tot = lngLDeposit_tot + oIL.Deposit
            If oIL.BottomOfDocument Then
                lngTotalExtras = lngTotalExtras + oIL.row_LTN
            End If
        End If
    Next
    'Set document properties
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtras_Foreign = lngTotalExtras
    mudtProps.DepositPaid = lngLDeposit_tot
    mudtProps.TotalDiscount_Foreign = d
    mudtProps.TotalDiscountExVAT = DN
  'Recalculate VAT overall
    Dim FinalVAT As Double
    FinalVAT = Round(TN_Vatable - (CDbl(TN_Vatable) * 100 / (CDbl(100) + oPC.Configuration.VATRate)))
    mudtProps.TotalVAT_Foreign = FinalVAT
    mudtProps.TotalNonVAT_Foreign = TN - FinalVAT  'Calculate the non-Vat value by subtracting the stated Final VAT value from the total extensions
    mudtProps.TotalVATLineSummed = v
    mudtProps.VATRoundingAdjustment = v - FinalVAT
    
    If Customer.ShowVAT = False And Customer.VATable = False Then  'We base all calculations on an ex-VAT price
                                                                                                 'see note in CalculateLine regarding this
        mudtProps.TotalPayable_Foreign = mudtProps.TotalExtension_Foreign - lngLDeposit_tot
    Else
        mudtProps.TotalPayable_Foreign = mudtProps.TotalNonVAT_Foreign
    End If
    mudtProps.TotalServiceItem_Foreign = lngTotalServiceItem
    mudtProps.TotalExtension_Foreign = TN
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
19380     On Error GoTo errHandler
19390     If bForeign Then
19400         TotalExtF = Format((mudtProps.TotalExtension_Foreign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
19410     Else
19420         TotalExtF = Format((mudtProps.TotalExtension) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
19430     End If
19440     Exit Property
errHandler:
19450     If ErrMustStop Then Debug.Assert False: Resume
19460     ErrorIn "a_Invoice.TotalExtWithOutExtrasF(bForeign)", bForeign
End Property

Public Property Get TotalExtWithOutExtras(bForeign As Boolean) As Long
19470     On Error GoTo errHandler
19480     If bForeign Then
19490         TotalExtWithOutExtras = mudtProps.TotalExtension_Foreign - mudtProps.TotalExtras_Foreign
19500     Else
19510         TotalExtWithOutExtras = mudtProps.TotalExtension - mudtProps.TotalExtras
19520     End If
19530     Exit Property
errHandler:
19540     If ErrMustStop Then Debug.Assert False: Resume
19550     ErrorIn "a_Invoice.TotalExtWithOutExtras(bForeign)", bForeign
End Property
Public Property Get TotalExtWithOutExtrasF(bForeign As Boolean) As String
19560     On Error GoTo errHandler
19570     If bForeign Then
19580         TotalExtWithOutExtrasF = Format((TotalExtWithOutExtras(bForeign)) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
19590     Else
19600         TotalExtWithOutExtrasF = Format((TotalExtWithOutExtras(bForeign)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
19610     End If
19620     Exit Property
errHandler:
19630     If ErrMustStop Then Debug.Assert False: Resume
19640     ErrorIn "a_Invoice.TotalExtWithOutExtrasF(bForeign)", bForeign
End Property
Public Property Get Total_nonExtras_ExVAT(bForeign As Boolean)
19650         Total_nonExtras_ExVAT = TotalExtWithOutExtras(bForeign) - Me.TotalVATRaw(bForeign)
End Property
Public Property Get Total_nonExtras_ExVATF(bForeign As Boolean)
19660     If bForeign Then
19670         Total_nonExtras_ExVATF = Format((Total_nonExtras_ExVAT(bForeign)) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
19680     Else
19690         Total_nonExtras_ExVATF = Format((Total_nonExtras_ExVAT(bForeign)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
19700     End If
End Property

Public Property Get TotalAfterDiscountExtF(bForeign As Boolean) As String
19710     On Error GoTo errHandler
19720     If bForeign Then
19730         TotalAfterDiscountExtF = Format(mudtProps.TotalExtension_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
19740     Else
19750         TotalAfterDiscountExtF = Format(mudtProps.TotalExtension / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
19760     End If
19770     Exit Property
errHandler:
19780     If ErrMustStop Then Debug.Assert False: Resume
19790     ErrorIn "a_Invoice.TotalAfterDiscountExtF(bForeign)", bForeign
End Property
Public Property Get TotalAfterDiscountExt(bForeign As Boolean) As String
19800     On Error GoTo errHandler
19810     If bForeign Then
19820         TotalAfterDiscountExt = mudtProps.TotalExtension_Foreign
19830     Else
19840         TotalAfterDiscountExt = mudtProps.TotalExtension
19850     End If
19860     Exit Property
errHandler:
19870     If ErrMustStop Then Debug.Assert False: Resume
19880     ErrorIn "a_Invoice.TotalAfterDiscountExt(bForeign)", bForeign
End Property
Public Property Get TotalNonVATF(bForeign As Boolean) As String
19890     On Error GoTo errHandler
19900     If bForeign Then
19910         TotalNonVATF = Format(mudtProps.TotalNonVAT_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
19920     Else
19930         TotalNonVATF = Format(mudtProps.TotalNonVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
19940     End If
19950     Exit Property
errHandler:
19960     If ErrMustStop Then Debug.Assert False: Resume
19970     ErrorIn "a_Invoice.TotalNonVATF(bForeign)", bForeign
End Property

Public Property Get TotalVATF(bForeign As Boolean) As String
19980     On Error GoTo errHandler
19990     If bForeign Then
20000         TotalVATF = Format(mudtProps.TotalVAT_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
20010     Else
20020         TotalVATF = Format(mudtProps.TotalVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
20030     End If
20040     Exit Property
errHandler:
20050     If ErrMustStop Then Debug.Assert False: Resume
20060     ErrorIn "a_Invoice.TotalVATF(bForeign)", bForeign
End Property
Public Property Get TotalPayableRaw(bForeign As Boolean) As Long
20070     On Error GoTo errHandler
20080     If bForeign Then
20090         TotalPayableRaw = mudtProps.TotalPayable_Foreign
20100     Else
20110         TotalPayableRaw = mudtProps.TotalPayable
20120     End If
20130     Exit Property
errHandler:
20140     If ErrMustStop Then Debug.Assert False: Resume
20150     ErrorIn "a_Invoice.TotalPayableRaw(bForeign)", bForeign
End Property
Public Property Get TotalVATRaw(bForeign As Boolean) As Long
20160     On Error GoTo errHandler
20170     If bForeign Then
20180         TotalVATRaw = mudtProps.TotalVAT_Foreign
20190     Else
20200         TotalVATRaw = mudtProps.TotalVAT
20210     End If
20220     Exit Property
errHandler:
20230     If ErrMustStop Then Debug.Assert False: Resume
20240     ErrorIn "a_Invoice.TotalVATRaw(bForeign)", bForeign
End Property
Public Property Get TotalVAT_WithDecimal(bForeign As Boolean) As String
20250     On Error GoTo errHandler
20260     If bForeign Then
20270         TotalVAT_WithDecimal = Format(mudtProps.TotalVAT_Foreign / oCurrency_Foreign.Divisor, "#######0.00")
20280     Else
20290         TotalVAT_WithDecimal = Format(mudtProps.TotalVAT / oCurrency_Foreign.Divisor, "#######0.00")
20300     End If
20310     Exit Property
errHandler:
20320     If ErrMustStop Then Debug.Assert False: Resume
20330     ErrorIn "a_Invoice.TotalVAT_WithDecimal(bForeign)", bForeign
End Property

Public Property Get TotalPayable(bForeign As Boolean) As String
20340     On Error GoTo errHandler
20350     If bForeign Then
20360         TotalPayable = mudtProps.TotalPayable_Foreign '/ oCurrency_Foreign.Divisor
20370     Else
20380         TotalPayable = mudtProps.TotalPayable ' / oPC.Configuration.DefaultCurrency.Divisor
20390     End If
20400     Exit Property
errHandler:
20410     If ErrMustStop Then Debug.Assert False: Resume
20420     ErrorIn "a_Invoice.TotalPayable(bForeign)", bForeign
End Property
Public Property Get TotalPayableF(bForeign As Boolean) As String
20430     On Error GoTo errHandler
20440     If bForeign Then
20450         TotalPayableF = Format(TotalPayable(bForeign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
20460     Else
20470         TotalPayableF = Format(TotalPayable(bForeign) / oCurrency_Foreign.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
20480     End If
20490     Exit Property
errHandler:
20500     If ErrMustStop Then Debug.Assert False: Resume
20510     ErrorIn "a_Invoice.TotalPayableF(bForeign)", bForeign
End Property
Public Property Get TotalPayable_WithDecimal(bForeign As Boolean) As String
20520     On Error GoTo errHandler
20530     If bForeign Then
20540         TotalPayable_WithDecimal = Format(mudtProps.TotalPayable_Foreign / oCurrency_Foreign.Divisor, "#######0.00")
20550     Else
20560         TotalPayable_WithDecimal = Format(mudtProps.TotalPayable / oPC.Configuration.DefaultCurrency.Divisor, "#######0.00")
20570     End If
20580     Exit Property
errHandler:
20590     If ErrMustStop Then Debug.Assert False: Resume
20600     ErrorIn "a_Invoice.TotalPayable_WithDecimal(bForeign)", bForeign
End Property

Public Property Get TotalServiceItem(bForeign As Boolean) As String
20610     On Error GoTo errHandler
20620     If bForeign Then
20630         TotalServiceItem = mudtProps.TotalServiceItem_Foreign '/ oCurrency_Foreign.Divisor
20640     Else
20650         TotalServiceItem = mudtProps.TotalServiceItem ' / oPC.Configuration.DefaultCurrency.Divisor
20660     End If
20670     Exit Property
errHandler:
20680     If ErrMustStop Then Debug.Assert False: Resume
20690     ErrorIn "a_Invoice.TotalServiceItem(bForeign)", bForeign
End Property
Public Property Get TotalServiceItemF(bForeign As Boolean) As String
20700     On Error GoTo errHandler
20710     If bForeign Then
20720         TotalServiceItemF = Format(TotalServiceItem(bForeign) / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
20730     Else
20740         TotalServiceItemF = Format(TotalServiceItem(bForeign) / oCurrency_Foreign.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
20750     End If
20760     Exit Property
errHandler:
20770     If ErrMustStop Then Debug.Assert False: Resume
20780     ErrorIn "a_Invoice.TotalServiceItemF(bForeign)", bForeign
End Property


Public Property Get TotalDiscount(bForeign As Boolean) As Long
20790     On Error GoTo errHandler
20800     If bForeign Then
20810         TotalDiscount = mudtProps.TotalDiscount_Foreign
20820     Else
20830         TotalDiscount = mudtProps.TotalDiscount
20840     End If
20850     Exit Property
errHandler:
20860     If ErrMustStop Then Debug.Assert False: Resume
20870     ErrorIn "a_Invoice.TotalDiscount(bForeign)", bForeign
End Property
Public Property Get TotalDiscountF(bForeign As Boolean) As String
20880     On Error GoTo errHandler
20890     If bForeign Then
20900         TotalDiscountF = Format(mudtProps.TotalDiscount_Foreign / oCurrency_Foreign.Divisor, oCurrency_Foreign.FormatString)
20910     Else
20920         TotalDiscountF = Format(mudtProps.TotalDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
20930     End If
20940     Exit Property
errHandler:
20950     If ErrMustStop Then Debug.Assert False: Resume
20960     ErrorIn "a_Invoice.TotalDiscountF(bForeign)", bForeign
End Property
Public Property Get TotalDiscountExVAT() As Long
20970     On Error GoTo errHandler
20980         TotalDiscountExVAT = mudtProps.TotalDiscountExVAT
20990     Exit Property
errHandler:
21000     If ErrMustStop Then Debug.Assert False: Resume
21010     ErrorIn "a_Invoice.TotalDiscount"
End Property
Public Property Get TotalDiscountExVATF() As String
21020     On Error GoTo errHandler
21030         TotalDiscountExVATF = Format(mudtProps.TotalDiscountExVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
21040     Exit Property
errHandler:
21050     If ErrMustStop Then Debug.Assert False: Resume
21060     ErrorIn "a_Invoice.TotalDiscountExVATF"
End Property

Public Property Get TotalDepositF(bForeign As Boolean) As String
21070     On Error GoTo errHandler
21080     If bForeign Then
21090         TotalDepositF = Format(TotalDeposit(bForeign), oCurrency_Foreign.FormatString)
21100     Else
21110         TotalDepositF = Format(TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
21120     End If
21130     Exit Property
errHandler:
21140     If ErrMustStop Then Debug.Assert False: Resume
21150     ErrorIn "a_Invoice.TotalDepositF(bForeign)", bForeign
End Property

Public Property Get VATRoundingAdjustmentF() As String
    VATRoundingAdjustmentF = Format(mudtProps.VATRoundingAdjustment / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get TotalDeposit(bForeign As Boolean) As Double
21160     On Error GoTo errHandler
21170     If bForeign Then
21180         TotalDeposit = lngFDeposit_tot / oCurrency_Foreign.Divisor
21190     Else
21200         TotalDeposit = lngLDeposit_tot / oPC.Configuration.DefaultCurrency.Divisor
21210     End If
21220     Exit Property
errHandler:
21230     If ErrMustStop Then Debug.Assert False: Resume
21240     ErrorIn "a_Invoice.TotalDeposit(bForeign)", bForeign
End Property
Public Property Get DocDiscountRate() As Double
21250     On Error GoTo errHandler
21260     DocDiscountRate = mudtProps.DiscountRate
        '  RecalculateAllLines
21270     Exit Property
errHandler:
21280     If ErrMustStop Then Debug.Assert False: Resume
21290     ErrorIn "a_Invoice.DocDiscountRate"
End Property
Public Property Get DocDiscountRateF() As String
21300     On Error GoTo errHandler
21310     If mudtProps.DiscountRate = 0 Then
21320         DocDiscountRateF = ""
21330     Else
21340         DocDiscountRateF = Format(mudtProps.DiscountRate, gPercentFormatString)
21350     End If
21360     Exit Property
errHandler:
21370     If ErrMustStop Then Debug.Assert False: Resume
21380     ErrorIn "a_Invoice.DocDiscountRateF"
End Property
Public Property Get TotalBalanceAfterDeposits(bForeign As Boolean) As String
21390     On Error GoTo errHandler
21400     If bForeign Then
21410         TotalBalanceAfterDeposits = Format(TotalPayable(bForeign) - TotalDeposit(bForeign), oCurrency_Foreign.FormatString)
21420     Else
21430         TotalBalanceAfterDeposits = Format(TotalPayable(bForeign) - TotalDeposit(bForeign), oPC.Configuration.DefaultCurrency.FormatString)
21440     End If
21450     Exit Property
errHandler:
21460     If ErrMustStop Then Debug.Assert False: Resume
21470     ErrorIn "a_Invoice.TotalBalanceAfterDeposits(bForeign)", bForeign
End Property
Private Function GetState() As String
21480     On Error GoTo errHandler
        Dim udtData As InvoiceData
        
21490   LSet udtData = mudtProps
21500   GetState = udtData.buffer
21510     Exit Function
errHandler:
21520     If ErrMustStop Then Debug.Assert False: Resume
21530     ErrorIn "a_Invoice.GetState"
End Function
Private Sub SetState(buffer As String)
21540     On Error GoTo errHandler
        Dim udtData As InvoiceData
        
21550   udtData.buffer = buffer
21560   LSet mudtProps = udtData
21570     Exit Sub
errHandler:
21580     If ErrMustStop Then Debug.Assert False: Resume
21590     ErrorIn "a_Invoice.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
21600     On Error GoTo errHandler
        Dim objPB As PropertyBag
        
21610   Set objPB = New PropertyBag
21620   With objPB
21630     .WriteProperty "State", GetState
21640     .WriteProperty "ILS", mColInvoiceLines.GetSuperState
21650     GetSuperState = .Contents
21660   End With
21670   Set objPB = Nothing
21680     Exit Function
errHandler:
21690     If ErrMustStop Then Debug.Assert False: Resume
21700     ErrorIn "a_Invoice.GetSuperState"
End Function
Public Sub SetSuperState(buffer As String)
21710     On Error GoTo errHandler
        Dim objPB As PropertyBag
        Dim arBuffer() As Byte
        
21720   Set objPB = New PropertyBag
21730   arBuffer = buffer
21740   With objPB
21750     .Contents = arBuffer
21760     SetState .ReadProperty("State")
21770     mColInvoiceLines.SetSuperState .ReadProperty("ILS")
21780   End With
21790   Set objPB = Nothing
21800     Exit Sub
errHandler:
21810     If ErrMustStop Then Debug.Assert False: Resume
21820     ErrorIn "a_Invoice.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
21830     On Error GoTo errHandler
21840   mColInvoiceLines.BeginEdit
21850   mcolStack.Add GetState
21860     Exit Sub
errHandler:
21870     If ErrMustStop Then Debug.Assert False: Resume
21880     ErrorIn "a_Invoice.BeginEdit"
End Sub

Public Sub CancelEdit()
21890     On Error GoTo errHandler
21900   If mcolStack.Count = 0 Then Err.Raise 445
21910   mColInvoiceLines.CancelEdit
21920   mudtProps.IsDeleted = False
21930   With mcolStack
21940     SetState .Item(.Count)
21950     .Remove .Count
21960   End With
21970     Exit Sub
errHandler:
21980     If ErrMustStop Then Debug.Assert False: Resume
21990     ErrorIn "a_Invoice.CancelEdit"
End Sub

Public Function ApplyEdit(Optional pMsg As String) As String
22000     On Error GoTo errHandler
      Dim strTemp As String
      Dim objPersist As a_Invoice_P
      Dim oSM As z_StockManager
      Dim OpenResult As Integer

22010     ApplyEdit = ""
22020     CalculateTotals
22030     mudtProps.TPID = Customer.ID
22040     mudtProps.SalesRepID = Customer.repid
          
          'Set these values fromthe objects before saving -  the objects cant be saved
22050     If Me.BillToAddressID > 0 Then
22060         mudtProps.BillToAddressID = Me.BillToAddressID
22070     Else
22080         If Not oBillToAddress Is Nothing Then
22090             mudtProps.BillToAddressID = oBillToAddress.ID
22100         End If
22110     End If
22120     If Me.DelToAddressID > 0 Then
22130         mudtProps.DelToAddressID = Me.DelToAddressID
22140     Else
22150         If Not oDelToAddress Is Nothing Then
22160             mudtProps.DelToAddressID = oDelToAddress.ID
22170         End If
22180     End If
          
      '''''''''''''''''''''''''''
          '--------------
22190     OpenResult = oPC.OpenDBSHort
          '--------------
22200     oPC.COShort.execute "BEGIN TRANSACTION"
22210     If Me.DOCDate = CDate(0) Then
22220         Me.DOCDate = Date
22230         Me.CaptureDate = Now
22240     End If
22250     If DOCCode = "" Then
22260         DOCCode = GetInvoiceNumber("GETANDWRITE")
        '      CaptureDate = Now()
22270     End If
        '  If oPC.AllowInvoiceDateOverride = False Then
        '      DOCDate = Date
        '  End If

      '    If IssDate < #1/1/1995# And mudtProps.IsNew Then IssDate = Now
          
22280     If mcolStack.Count = 0 Then Err.Raise 445
22290     oPC.COShort.CommandTimeout = 35
22300     mColInvoiceLines.ApplyEdit
22310     Set objPersist = New a_Invoice_P
          
22320     If mudtProps.IsDeleted Then
22330         objPersist.DeleteObject mudtProps.InvoiceID
22340         mcolStack.Remove mcolStack.Count
22350         mudtProps.IsNew = True
22360         mudtProps.IsDeleted = False
22370     ElseIf IsDirty Or mudtProps.IsNew Then
22380         If Not IsValid Then Err.Raise 445
      '        If mudtProps.IsNew Then
      '            mudtProps.CaptureDate = Now()
      '        End If
22390         strTemp = objPersist.Save(GetSuperState)
          
22400         If UCase(Left(strTemp, 7)) <> "TIMEOUT" Then
22410             mcolStack.Remove mcolStack.Count
22420             SetSuperState strTemp
22430             mudtProps.IsNew = False
22440         Else 'save failed
22450             mColInvoiceLines.BeginEdit
22460             If Not IsMissing(pMsg) Then
22470                 pMsg = strTemp
22480             End If
22490             ApplyEdit = "TIMEOUT"
22500             Set oSM = New z_StockManager
22510             oSM.RollbackTransaction oPC.COShort
22520             Set oSM = Nothing
22530             Set objPersist = Nothing
22540         LogSaveToFile "a_Invoice:TimeoutExpired"
22550             GoTo EXIT_Handler
22560         End If
          
22570     Else
22580         mcolStack.Remove mcolStack.Count
22590     End If
22600     Set objPersist = Nothing
22610     SetDirty False
22620     oPC.COShort.execute "COMMIT TRANSACTION"
          '--------------
22630     If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
          '--------------
      ''''''''''''''''''''''''''''''''''''''''
        '  oPC.COShort.CommandTimeout = 0
          
EXIT_Handler:
22640     Exit Function
errHandler:
22650     ErrPreserve
22660     oSM.RollbackTransaction oPC.COShort
22670     If ErrMustStop Then Debug.Assert False: Resume
22680     ErrorIn "a_Invoice.ApplyEdit", , , oPC.COShort
22690     Exit Function
22700     Resume Next
End Function
Public Property Get IsDeleted() As Boolean
22710     On Error GoTo errHandler
22720   IsDeleted = mudtProps.IsDeleted
22730     Exit Property
errHandler:
22740     If ErrMustStop Then Debug.Assert False: Resume
22750     ErrorIn "a_Invoice.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
22760     On Error GoTo errHandler
22770   IsEditing = mcolStack.Count > 0
22780     Exit Property
errHandler:
22790     If ErrMustStop Then Debug.Assert False: Resume
22800     ErrorIn "a_Invoice.IsEditing"
End Property

Public Property Get IsNew() As Boolean
22810     On Error GoTo errHandler
22820   IsNew = mudtProps.IsNew
22830     Exit Property
errHandler:
22840     If ErrMustStop Then Debug.Assert False: Resume
22850     ErrorIn "a_Invoice.IsNew"
End Property

Public Property Get IsDirty() As Boolean
22860     On Error GoTo errHandler
22870   IsDirty = mudtProps.IsDirty Or mColInvoiceLines.IsDirty
22880     Exit Property
errHandler:
22890     If ErrMustStop Then Debug.Assert False: Resume
22900     ErrorIn "a_Invoice.IsDirty"
End Property
Public Sub SetDirty(pVal As Boolean)
22910     On Error GoTo errHandler
22920     mudtProps.IsDirty = pVal
22930     RaiseEvent Dirty(pVal)
22940     Exit Sub
errHandler:
22950     If ErrMustStop Then Debug.Assert False: Resume
22960     ErrorIn "a_Invoice.SetDirty(pVal)", pVal
End Sub
Public Property Get IsValid() As Boolean
22970     On Error GoTo errHandler
22980     IsValid = (mobjValid.Count = 0)
22990     Exit Property
errHandler:
23000     If ErrMustStop Then Debug.Assert False: Resume
23010     ErrorIn "a_Invoice.IsValid"
End Property
Public Function GetStatus()
23020     On Error GoTo errHandler
23030     If IsValid Then
23040         RaiseEvent Valid("")
23050     Else
23060         RaiseEvent Valid(TranslateErrors(mobjValid.AllBrokenRules))
23070     End If
         ' RaiseEvent Dirty(mudtProps.IsDirty)
23080     Exit Function
errHandler:
23090     If ErrMustStop Then Debug.Assert False: Resume
23100     ErrorIn "a_Invoice.GetSTatus"
End Function
Private Sub Class_Initialize()
      Dim oComp As a_Company
23110     On Error GoTo errHandler
23120     Set mcolStack = New Collection
23130     Set mobjValid = New z_BrokenRules
23140     Set mColInvoiceLines = New ch_InvoiceLine
23150     mColInvoiceLines.component Me
23160     Set mCustomer = New a_Customer

23170     Set oPCL = New z_PrintClient
23180     mudtProps.IsNew = True
23190     mudtProps.VATRate = oPC.Configuration.VATRate
23200     mudtProps.VATable = True
23210     mudtProps.ShowVAT = oPC.Configuration.DiscountVATDefault
23220     mudtProps.COMPID = oPC.Configuration.DefaultCOMPID
23230     For Each oComp In oPC.Configuration.Companies
23240         If oComp.ID = mudtProps.COMPID Then
23250             Set oBillingCompany = oComp
23260         End If
23270     Next

23280     mudtProps.DiscountRate = 0
23290     mudtProps.Status = stInProcess
23300     mOriginalStatus = stInProcess
23310     Set oBillToAddress = New a_Address
23320     Set oDelToAddress = New a_Address
23330     mobjValid.BreakRule "TP", True
23340     LoadClassorsCollection
23350     Set oCurrency_Foreign = oPC.Configuration.DefaultCurrency
23360     Exit Sub
errHandler:
23370     If ErrMustStop Then Debug.Assert False: Resume
23380     ErrorIn "a_Invoice.Class_Initialize"
End Sub

Private Sub Class_Terminate()
23390     On Error GoTo errHandler
23400     If mcolStack.Count > 0 Then _
            Err.Raise vbObjectError + 1001, , "State stack is not empty"
            
23410     Set mcolStack = Nothing
23420     Set mobjValid = Nothing
23430     Set mColInvoiceLines = Nothing
23440     Set mCustomer = Nothing
23450     Set oPCL = Nothing
23460     Set cColsOSPerCust = Nothing
23470     Exit Sub
errHandler:
23480     If ErrMustStop Then Debug.Assert False: Resume
23490     ErrorIn "a_Invoice.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(oRS As String)
23500     On Error GoTo errHandler
23510     RaiseEvent Valid(TranslateErrors(oRS))
23520     Exit Sub
errHandler:
23530     If ErrMustStop Then Debug.Assert False: Resume
23540     ErrorIn "a_Invoice.mobjValid_BrokenRule(ors)", oRS
End Sub
Private Sub mobjValid_RuleUnbroken(oRS As String)
23550     On Error GoTo errHandler
23560     RaiseEvent Valid(TranslateErrors(oRS))
23570     Exit Sub
errHandler:
23580     If ErrMustStop Then Debug.Assert False: Resume
23590     ErrorIn "a_Invoice.mobjValid_RuleUnbroken(ors)", oRS
End Sub

Private Sub mobjValid_NoBrokenRules()
23600     On Error GoTo errHandler
23610     RaiseEvent Valid("")
23620     Exit Sub
errHandler:
23630     If ErrMustStop Then Debug.Assert False: Resume
23640     ErrorIn "a_Invoice.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
23650     On Error GoTo errHandler
23660     RaiseEvent Valid(TranslateErrors(pMsg))
23670     Exit Sub
errHandler:
23680     If ErrMustStop Then Debug.Assert False: Resume
23690     ErrorIn "a_Invoice.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
23700     On Error GoTo errHandler
      Dim objPersist As a_Invoice_P
      Dim oAdd As a_Address
      Dim oComp As a_Company
      Dim oCurr As a_Currency
23710     If mcolStack.Count > 0 Then Err.Raise 445
23720     If Not mudtProps.IsNew Then Err.Raise 445
          
23730     mudtProps.IsNew = False
          
23740     Set objPersist = New a_Invoice_P
23750     SetSuperState objPersist.Fetch(TRID)
23760     mOriginalStatus = Me.Status
      '''''''''''
23770     Set oBillToAddress = Nothing
23780     Set oDelToAddress = Nothing    '

          
          
23790     mCustomer.Load mudtProps.TPID
23800     mudtProps.ShowVAT = Customer.ShowVAT
          
23810     If BillToAddressID > 0 Then
23820         Set oBillToAddress = Customer.Addresses.FindByID(BillToAddressID)
              'The address might not be found if it has been removed or the customer has been merged with another
23830         If oBillToAddress Is Nothing Then
23840             If Not Customer.BillTOAddress Is Nothing Then
23850                 Set oBillToAddress = Customer.BillTOAddress
23860                 BillToAddressID = Customer.BillTOAddress.ID
23870             End If
23880         End If
23890     End If
23900     If DelToAddressID > 0 Then
23910         Set oDelToAddress = Customer.Addresses.FindByID(DelToAddressID)
              'The address might not be found if it has been removed or the customer has been merged with another
23920         If oDelToAddress Is Nothing Then
23930             If Not Customer.DelToAddress Is Nothing Then
23940                 Set oDelToAddress = Customer.DelToAddress
23950                 DelToAddressID = Customer.DelToAddress.ID
23960             End If
23970         End If
23980     End If
          
          
      '    For Each oAdd In Customer.Addresses
      '        If oAdd.ID = mudtProps.BillToAddressID Then
      '            Set oBillToAddress = oAdd
      '        End If
      '        If oAdd.ID = mudtProps.DelToAddressID Then
      '            Set oDelToAddress = oAdd
      '        End If
      '    Next
23990     For Each oComp In oPC.Configuration.Companies
24000         If oComp.ID = mudtProps.COMPID Then
24010             Set oBillingCompany = oComp
24020         End If
24030     Next
24040     For Each oCurr In oPC.Configuration.Currencies
24050         If oCurr.ID = mudtProps.CurrencyID_Foreign Then
24060             Set oCurrency_Foreign = oCurr
24070         End If
24080     Next
24090     Set objPersist = Nothing
24100     mobjValid.RuleBroken "TP", False
       '   Me.RecalculateAllLines
24110     CalculateTotals
24120     mudtProps.IsNew = False
24130     mudtProps.IsDirty = False
24140     Exit Sub
errHandler:
24150     If ErrMustStop Then Debug.Assert False: Resume
24160     ErrorIn "a_Invoice.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub
Public Property Get Customer() As a_Customer
24170     On Error GoTo errHandler
24180     Set Customer = mCustomer
24190     Exit Property
errHandler:
24200     If ErrMustStop Then Debug.Assert False: Resume
24210     ErrorIn "a_Invoice.Customer"
End Property
Public Sub Delete()
24220     On Error GoTo errHandler
24230   If mcolStack.Count = 0 Then Err.Raise 445
        
24240   mudtProps.IsDeleted = True
24250   SetDirty True
24260     Exit Sub
errHandler:
24270     If ErrMustStop Then Debug.Assert False: Resume
24280     ErrorIn "a_Invoice.Delete"
End Sub



Public Property Get InvoiceID() As Long
24290     On Error GoTo errHandler
24300     InvoiceID = mudtProps.InvoiceID
24310     Exit Property
errHandler:
24320     If ErrMustStop Then Debug.Assert False: Resume
24330     ErrorIn "a_Invoice.InvoiceID"
End Property
Public Property Get ID() As String
On Error GoTo errHandler
    ID = mudtProps.ID
Exit Property
errHandler:
        If ErrMustStop Then Debug.Assert False: Resume
        ErrorIn "a_Invoice.Id"
End Property
Public Property Let InvoiceID(val As Long)
24340     On Error GoTo errHandler
24350     If mcolStack.Count = 0 Then Err.Raise 383
24360     mudtProps.InvoiceID = val
24370     SetDirty True
24380     Exit Property
errHandler:
24390     If ErrMustStop Then Debug.Assert False: Resume
24400     ErrorIn "a_Invoice.InvoiceID(val)", val
End Property

Public Property Get BillToAddressID() As Long
24410     On Error GoTo errHandler
24420     BillToAddressID = mudtProps.BillToAddressID
24430     Exit Property
errHandler:
24440     If ErrMustStop Then Debug.Assert False: Resume
24450     ErrorIn "a_Invoice.BillToAddressID:Get"
End Property
Public Property Let BillToAddressID(val As Long)
24460     On Error GoTo errHandler
24470     mudtProps.BillToAddressID = val
24480     Exit Property
errHandler:
24490     If ErrMustStop Then Debug.Assert False: Resume
24500     ErrorIn "a_Invoice.BillToAddressID:Let"
End Property
Public Sub SetBillToAddress(val As a_Address)
24510     On Error GoTo errHandler
24520     If mcolStack.Count = 0 Then Err.Raise 383
24530     Set oBillToAddress = val
24540     mudtProps.BillToAddressID = oBillToAddress.ID
24550     SetDirty True
24560     Exit Sub
errHandler:
24570     If ErrMustStop Then Debug.Assert False: Resume
24580     ErrorIn "a_Invoice.SetBillToAddress(val)", val
End Sub
Public Sub SetBillToAddressImmediate(val As a_Address)
24590     On Error GoTo errHandler
      Dim oSM As New z_StockManager
24600     Set oBillToAddress = val
24610     mudtProps.BillToAddressID = oBillToAddress.ID
24620     oSM.SetAddressImmediate "BILL", Me.InvoiceID, oBillToAddress.ID
24630     Exit Sub
errHandler:
24640     If ErrMustStop Then Debug.Assert False: Resume
24650     ErrorIn "a_Invoice.SetBillToAddressImmediate(val)", val
End Sub

Public Property Get DelToAddressID() As Long
24660     On Error GoTo errHandler
24670     DelToAddressID = mudtProps.DelToAddressID
24680     Exit Property
errHandler:
24690     If ErrMustStop Then Debug.Assert False: Resume
24700     ErrorIn "a_Invoice.DelToAddressID:Get"
End Property
Public Property Let DelToAddressID(val As Long)
24710     On Error GoTo errHandler
24720    mudtProps.DelToAddressID = val
24730     Exit Property
errHandler:
24740     If ErrMustStop Then Debug.Assert False: Resume
24750     ErrorIn "a_Invoice.DelToAddressID:Let"
End Property
Public Sub setDelToAddress(val As a_Address)
24760     On Error GoTo errHandler
24770     If mcolStack.Count = 0 Then Err.Raise 383
24780     Set oDelToAddress = val
24790     mudtProps.DelToAddressID = oDelToAddress.ID
24800     SetDirty True
24810     Exit Sub
errHandler:
24820     If ErrMustStop Then Debug.Assert False: Resume
24830     ErrorIn "a_Invoice.setDelToAddress(val)", val
End Sub
Public Sub setDelToAddressImmediate(val As a_Address)
24840     On Error GoTo errHandler
      Dim oSM As New z_StockManager
24850     Set oDelToAddress = val
24860     mudtProps.DelToAddressID = oDelToAddress.ID
24870     oSM.SetAddressImmediate "DEL", Me.InvoiceID, oDelToAddress.ID
24880     Exit Sub
errHandler:
24890     If ErrMustStop Then Debug.Assert False: Resume
24900     ErrorIn "a_Invoice.setDelToAddressImmediate(val)", val
End Sub

Public Function SetGeneralDiscount(val As String) As Boolean
24910     On Error GoTo errHandler
      Dim dblDiscountPercent As Double
      Dim oIL As a_InvoiceLine

24920     If mcolStack.Count = 0 Then Err.Raise 383
24930     SetGeneralDiscount = True
24940     If Trim$(val) = "" Then
24950         dblDiscountPercent = 0
24960     ElseIf Not ConvertToDBL(val, dblDiscountPercent) Then
24970         SetGeneralDiscount = False
24980         Exit Function
24990     End If
25000     If (dblDiscountPercent <= 0) Then
25010         mobjValid.RuleBroken "DISCOUNT", True
25020     Else
25030         mobjValid.RuleBroken "DISCOUNT", False
25040     End If
25050     mudtProps.DiscountRate = dblDiscountPercent
25060     SetDirty True
25070     For Each oIL In InvoiceLines  'disctribute discount rate over all lines where non nonstaock items exists
25080         If Not oIL.ServiceItem Then
25090             oIL.DiscountPercent = dblDiscountPercent
25100         End If
25110     Next
25120     mcolinvoiceLines_RowsChange  'recalulates andd raises event
25130     RecalculateAllLines
25140     RaiseEvent reloadlist
25150     Exit Function
errHandler:
25160     If ErrMustStop Then Debug.Assert False: Resume
25170     ErrorIn "a_Invoice.SetGeneralDiscount(val)", val
End Function
Public Function LoadCOLsOS(Optional pTPID As Long, Optional pPID As String, Optional pCOLID As Long) As c_COLsPerCust
25180     On Error GoTo errHandler
25190     Set cColsOSPerCust = Nothing
25200     Set cColsOSPerCust = New c_COLsPerCust
25210     cColsOSPerCust.Load Customer.ID, pPID, pCOLID
25220     Exit Function
errHandler:
25230     If ErrMustStop Then Debug.Assert False: Resume
25240     ErrorIn "a_Invoice.LoadCOLsOS(pTPID,pPID,pCOLID)", Array(pTPID, pPID, pCOLID)
End Function
Public Function LoadAPPLsOS(Optional pTPID As Long, Optional pPID As String) As c_COLsPerCust
25250     On Error GoTo errHandler
25260     Set cAPPLsOS = Nothing
25270     Set cAPPLsOS = New c_APPLsPerTPPID
25280     cAPPLsOS.Load Customer.ID, pPID, ""
25290     Exit Function
errHandler:
25300     If ErrMustStop Then Debug.Assert False: Resume
25310     ErrorIn "a_Invoice.LoadAPPLsOS(pTPID,pPID)", Array(pTPID, pPID)
End Function

Public Function SetCustomer(pTPID As Long) As Boolean
25320     On Error GoTo errHandler
      Dim bSuccess As Boolean
25330     bSuccess = Customer.Load(pTPID)
25340     SetCustomer = bSuccess
25350     If bSuccess Then
25360         VATable = Customer.VATable
25370         mobjValid.RuleBroken "TP", False
25380         If Not Me.Customer.BillTOAddress Is Nothing Then SetBillToAddress Me.Customer.BillTOAddress
25390         If Not Me.Customer.DelToAddress Is Nothing Then setDelToAddress Me.Customer.DelToAddress
            '  SetBillToAddressID Me.Customer.BillTOAddress.ID
            '  SetDelToAddressID Me.Customer.DelToAddress.ID
25400         SetDirty True
      '        MsgBox "HERE"
      '        Set cColsOSPerCust = Nothing
      '        Set cColsOSPerCust = New c_COLsPerCust
      '        cColsOSPerCust.Load pTPID
25410     End If
25420     Exit Function
errHandler:
25430     If ErrMustStop Then Debug.Assert False: Resume
25440     ErrorIn "a_Invoice.SetCustomer(pTPID)", pTPID
End Function
Public Function SetCustomerFromAccNum(pAccnum As String) As Boolean
25450     On Error GoTo errHandler
      Dim bSuccess As Boolean

25460     bSuccess = Me.Customer.Load(, pAccnum)
25470     SetCustomerFromAccNum = bSuccess
25480     If bSuccess Then
25490         mobjValid.RuleBroken "TP", False
25500         If Me.BillToAddressID = 0 Then
25510             Set oBillToAddress = Me.Customer.BillTOAddress
25520         End If
25530         If Me.DelToAddressID = 0 Then
25540             Set oDelToAddress = Me.Customer.DelToAddress
25550         End If
25560     End If
25570     Me.VATable = Customer.VATable
25580     Me.ShowVAT = Customer.ShowVAT
25590     SetDirty True
25600     Exit Function
errHandler:
25610     If ErrMustStop Then Debug.Assert False: Resume
25620     ErrorIn "a_Invoice.SetCustomerFromAccNum(pAccnum)", pAccnum
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
25630     On Error GoTo errHandler
25640     If mcolStack.Count = 0 Then Err.Raise 383
25650     mudtProps.ShowVAT = val
25660     SetDirty True
25670     Exit Property
errHandler:
25680     If ErrMustStop Then Debug.Assert False: Resume
25690     ErrorIn "a_Invoice.ShowVAT(val)", val
End Property
Public Property Get NonVATDocument() As Boolean
25700     On Error GoTo errHandler
25710     NonVATDocument = mudtProps.NonVATDocument
25720     Exit Property
errHandler:
25730     If ErrMustStop Then Debug.Assert False: Resume
25740     ErrorIn "a_Invoice.NonVATDocument"
End Property
Public Property Let NonVATDocument(val As Boolean)
25750     On Error GoTo errHandler
25760     If mcolStack.Count = 0 Then Err.Raise 383
25770     mudtProps.NonVATDocument = val
25780     SetDirty True
25790     Exit Property
errHandler:
25800     If ErrMustStop Then Debug.Assert False: Resume
25810     ErrorIn "a_Invoice.NonVATDocument(val)", val
End Property
Public Property Get ShowVAT() As Boolean
25820     On Error GoTo errHandler
25830     ShowVAT = mudtProps.ShowVAT
25840     Exit Property
errHandler:
25850     If ErrMustStop Then Debug.Assert False: Resume
25860     ErrorIn "a_Invoice.ShowVAT"
End Property

Public Property Get StaffID() As Long
25870     On Error GoTo errHandler
25880     StaffID = mudtProps.StaffID
25890     Exit Property
errHandler:
25900     If ErrMustStop Then Debug.Assert False: Resume
25910     ErrorIn "a_Invoice.StaffID"
End Property
Public Property Let StaffID(val As Long)
25920     On Error GoTo errHandler
25930     If mcolStack.Count = 0 Then Err.Raise 383
25940     mudtProps.StaffID = val
25950     SetDirty True
25960     Exit Property
errHandler:
25970     If ErrMustStop Then Debug.Assert False: Resume
25980     ErrorIn "a_Invoice.StaffID(val)", val
End Property
Public Property Get StaffName() As String
25990     On Error GoTo errHandler
26000     If mudtProps.StaffID < 1 Then
26010         StaffName = ""
26020         Exit Property
26030     End If
26040     If oPC.Configuration.Staff.FindStaffByID(StaffID) Is Nothing Then
26050     StaffName = ""
26060     Exit Property
26070     End If
26080     StaffName = oPC.Configuration.Staff.FindStaffByID(StaffID).StaffName
26090     Exit Property
errHandler:
26100     If ErrMustStop Then Debug.Assert False: Resume
26110     ErrorIn "a_Invoice.StaffName"
End Property

Public Property Get StaffNameB() As String
26120     On Error GoTo errHandler
      Dim strB As String
26130     If mudtProps.StaffID < 1 Then
26140         StaffNameB = ""
26150         Exit Property
26160     End If
26170     If oPC.Configuration.Staff.FindStaffByID(mudtProps.StaffID) Is Nothing Then
26180         StaffNameB = ""
26190     Else
26200         strB = oPC.Configuration.Staff.FindStaffByID(mudtProps.StaffID).Shortname
26210         If strB > "" Then
26220             StaffNameB = "  (" & strB & ")"
26230         Else
26240             StaffNameB = ""
26250         End If
26260     End If
26270     Exit Property
errHandler:
26280     If ErrMustStop Then Debug.Assert False: Resume
26290     ErrorIn "a_Invoice.StaffNameB"
End Property
Public Property Get Signature()
26300 Signature = FNS(mudtProps.Signature)
End Property

Public Property Get CurrencyID_Foreign() As Long
26310     On Error GoTo errHandler
26320     CurrencyID_Foreign = mudtProps.CurrencyID_Foreign
26330     Exit Property
errHandler:
26340     If ErrMustStop Then Debug.Assert False: Resume
26350     ErrorIn "a_Invoice.CurrencyID_Foreign"
End Property
Public Property Let CurrencyID(val As Long)
26360     On Error GoTo errHandler
      Dim oCurr As a_Currency
      Dim oIL As a_InvoiceLine

26370     If mcolStack.Count = 0 Then Err.Raise 383
26380     mudtProps.CurrencyID_Foreign = val
26390     SetDirty True
26400     For Each oCurr In oPC.Configuration.Currencies
26410         If oCurr.ID = mudtProps.CurrencyID_Foreign Then
26420             Set oCurrency_Foreign = oCurr
26430         End If
26440     Next
26450     Me.CurrencyFactor = oCurrency_Foreign.Factor
26460     RecalculateAllLines
26470     Exit Property
errHandler:
26480     If ErrMustStop Then Debug.Assert False: Resume
26490     ErrorIn "a_Invoice.CurrencyID(val)", val
End Property
Public Property Get COMPID() As Long
26500     On Error GoTo errHandler
26510     COMPID = mudtProps.COMPID
26520     Exit Property
errHandler:
26530     If ErrMustStop Then Debug.Assert False: Resume
26540     ErrorIn "a_Invoice.COMPID"
End Property
Public Property Let COMPID(val As Long)
26550     On Error GoTo errHandler
      Dim oComp As a_Company
26560     If mcolStack.Count = 0 Then Err.Raise 383
26570     mudtProps.COMPID = val
26580     SetDirty True
26590     For Each oComp In oPC.Configuration.Companies
26600         If oComp.ID = mudtProps.COMPID Then
26610             Set oBillingCompany = oComp
26620         End If
26630     Next
26640     Exit Property
errHandler:
26650     If ErrMustStop Then Debug.Assert False: Resume
26660     ErrorIn "a_Invoice.COMPID(val)", val
End Property
Public Property Get BillingCompany() As a_Company
26670     Set BillingCompany = oBillingCompany
End Property

Public Property Get CustPaid() As Boolean
26680     CustPaid = mudtProps.CustPaid
End Property
Public Property Let CustPaid(val As Boolean)
26690     mudtProps.CustPaid = val
End Property
Public Property Get CommPaid() As Boolean
26700     CommPaid = mudtProps.CommPaid
End Property
Public Property Let CommPaid(val As Boolean)
26710     mudtProps.CommPaid = val
End Property
Public Property Get SalesRepID() As Long
26720     SalesRepID = mudtProps.SalesRepID
End Property
Public Property Let SalesRepID(val As Long)
26730     mudtProps.SalesRepID = val
End Property
Public Property Get SalesRepName() As String
26740     SalesRepName = FNS(mudtProps.SalesRepName)
End Property
Public Property Let SalesRepName(val As String)
26750     mudtProps.SalesRepName = val
End Property
Public Property Get TPNAME() As String
26760     On Error GoTo errHandler
26770     TPNAME = FNS(mudtProps.TPNAME)
26780     Exit Property
errHandler:
26790     If ErrMustStop Then Debug.Assert False: Resume
26800     ErrorIn "a_Invoice.TPName"
End Property
Public Property Let TPNAME(val As String)
26810     On Error GoTo errHandler
26820     If mcolStack.Count = 0 Then Err.Raise 383
26830     mudtProps.TPNAME = val
26840     SetDirty True
26850     Exit Property
errHandler:
26860     If ErrMustStop Then Debug.Assert False: Resume
26870     ErrorIn "a_Invoice.TPName(val)", val
End Property
Public Property Let IsPreDelivery(val As Boolean)
26880     On Error GoTo errHandler
26890     mudtProps.IsPreDelivery = val
26900     SetDirty True
26910     Exit Property
errHandler:
26920     If ErrMustStop Then Debug.Assert False: Resume
26930     ErrorIn "a_Invoice.IsPreDelivery(val)", val, EA_NORERAISE
End Property
Public Property Get IsPreDelivery() As Boolean
26940     IsPreDelivery = FNB(mudtProps.IsPreDelivery)
End Property
Public Property Get TPPhone() As String
26950     On Error GoTo errHandler
26960     TPPhone = FNS(mudtProps.TPPhone)
26970     Exit Property
errHandler:
26980     If ErrMustStop Then Debug.Assert False: Resume
26990     ErrorIn "a_Invoice.TPPhone"
End Property
Public Property Let TPPhone(val As String)
27000     On Error GoTo errHandler
27010     If mcolStack.Count = 0 Then Err.Raise 383
27020     mudtProps.TPPhone = val
27030     SetDirty True
27040     Exit Property
errHandler:
27050     If ErrMustStop Then Debug.Assert False: Resume
27060     ErrorIn "a_Invoice.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
27070     On Error GoTo errHandler
27080     TPFax = FNS(mudtProps.TPFax)
27090     Exit Property
errHandler:
27100     If ErrMustStop Then Debug.Assert False: Resume
27110     ErrorIn "a_Invoice.TPFax"
End Property
Public Property Let TPFax(val As String)
27120     On Error GoTo errHandler
27130     If mcolStack.Count = 0 Then Err.Raise 383
27140     mudtProps.TPFax = val
27150     SetDirty True
27160     Exit Property
errHandler:
27170     If ErrMustStop Then Debug.Assert False: Resume
27180     ErrorIn "a_Invoice.TPFax(val)", val
End Property

Public Property Get TPACCNum() As String
27190     On Error GoTo errHandler
27200     TPACCNum = FNS(mudtProps.TPACCNum)
27210     Exit Property
errHandler:
27220     If ErrMustStop Then Debug.Assert False: Resume
27230     ErrorIn "a_Invoice.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
27240     On Error GoTo errHandler
27250     If mcolStack.Count = 0 Then Err.Raise 383
27260     mudtProps.TPACCNum = val
27270     SetDirty True
27280     Exit Property
errHandler:
27290     If ErrMustStop Then Debug.Assert False: Resume
27300     ErrorIn "a_Invoice.TPACCNum(val)", val
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
27310     On Error GoTo errHandler
27320     BusPhone = FNS(mudtProps.BusPhone)
27330     Exit Property
errHandler:
27340     If ErrMustStop Then Debug.Assert False: Resume
27350     ErrorIn "a_Invoice.BusPhone"
End Property
Public Property Let BusPhone(val As String)
27360     On Error GoTo errHandler
27370     If mcolStack.Count = 0 Then Err.Raise 383
27380     mudtProps.BusPhone = val
27390     SetDirty True
27400     Exit Property
errHandler:
27410     If ErrMustStop Then Debug.Assert False: Resume
27420     ErrorIn "a_Invoice.BusPhone(val)", val
End Property

Public Property Get DOCCode() As String
27430     On Error GoTo errHandler
27440     DOCCode = FNS(mudtProps.DOCCode)
27450     Exit Property
errHandler:
27460     If ErrMustStop Then Debug.Assert False: Resume
27470     ErrorIn "a_Invoice.DocCode"
End Property
Public Property Get DOCCodeF() As String
27480     On Error GoTo errHandler
27490     Select Case mudtProps.Status
          Case 1
27500         DOCCodeF = "Voided " & FNS(mudtProps.DOCCode)
27510     Case 2
27520         If oPC.AllowsInvoicePicking Then
27530             DOCCodeF = "Picking slip in process " & FNS(mudtProps.DOCCode)
27540         Else
27550             If Me.Proforma = True Then
27560                 DOCCodeF = "Proforma invoice in process " & FNS(mudtProps.DOCCode)
27570             Else
27580                 DOCCodeF = "Invoice in process " & FNS(mudtProps.DOCCode)
27590             End If
27600         End If
27610     Case 3
27620         If oPC.AllowsInvoicePicking Then
27630             DOCCodeF = "Picking slip " & FNS(mudtProps.DOCCode)
27640         Else
27650             If Me.Proforma = True Then
27660                 DOCCodeF = "Pro-forma invoice " & FNS(mudtProps.DOCCode)
27670             Else  'This should not occur as a tax invoice should always be status 4
27680                 DOCCodeF = "Tax Invoice " & FNS(mudtProps.DOCCode)
27690             End If
27700         End If
27710     Case 4
27720         If Me.Proforma = True Then
27730             DOCCodeF = "Pro-forma invoice " & FNS(mudtProps.DOCCode)
27740         Else
27750             DOCCodeF = "Tax Invoice " & FNS(mudtProps.DOCCode)
27760         End If
27770     Case 5
27780         DOCCodeF = "Pro-forma invoice " & FNS(mudtProps.DOCCode)
27790     End Select
27800     Exit Property
errHandler:
27810     If ErrMustStop Then Debug.Assert False: Resume
27820     ErrorIn "a_Invoice.DOCCodeF"
End Property
Public Property Get IsInvoice() As Boolean
27830     IsInvoice = True
End Property
Public Property Get IsGDN() As Boolean
27840     IsGDN = False
End Property
Public Property Let DOCCode(val As String)
27850     On Error GoTo errHandler
27860     If mcolStack.Count = 0 Then Err.Raise 383
27870     mudtProps.DOCCode = val
27880     SetDirty True
27890     Exit Property
errHandler:
27900     If ErrMustStop Then Debug.Assert False: Resume
27910     ErrorIn "a_Invoice.DocCode(val)", val
End Property

Public Property Get CurrencyFormat() As String
27920     On Error GoTo errHandler
27930     CurrencyFormat = FNS(mudtProps.CurrencyFormat)
27940     If CurrencyFormat <= "" Then CurrencyFormat = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").FormatString
27950     Exit Property
errHandler:
27960     If ErrMustStop Then Debug.Assert False: Resume
27970     ErrorIn "a_Invoice.CurrencyFormat"
End Property
Public Property Let CurrencyFormat(val As String)
27980     On Error GoTo errHandler
27990     If mcolStack.Count = 0 Then Err.Raise 383
28000     mudtProps.CurrencyFormat = val
28010     SetDirty True
28020     Exit Property
errHandler:
28030     If ErrMustStop Then Debug.Assert False: Resume
28040     ErrorIn "a_Invoice.CurrencyFormat(val)", val
End Property

Public Property Get DOCDate() As Date
28050     On Error GoTo errHandler
28060     DOCDate = CDate(ReverseDate(mudtProps.DOCDate))  'to ensure no time included
28070     Exit Property
errHandler:
28080     If ErrMustStop Then Debug.Assert False: Resume
28090     ErrorIn "a_Invoice.DocDate"
End Property
Public Property Get DocDateF() As String
28100     On Error GoTo errHandler
28110     DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
28120     Exit Property
errHandler:
28130     If ErrMustStop Then Debug.Assert False: Resume
28140     ErrorIn "a_Invoice.DocDateF"
End Property
Public Property Get TransDateF2() As String
28150     On Error GoTo errHandler
28160     TransDateF2 = Format(mudtProps.DOCDate, "dd mmmm yyyy")
28170     Exit Property
errHandler:
28180     If ErrMustStop Then Debug.Assert False: Resume
28190     ErrorIn "a_Invoice.TransDateF2"
End Property
Public Property Let DOCDate(val As Date)
28200     On Error GoTo errHandler
28210     If mcolStack.Count = 0 Then Err.Raise 383
28220     If IsDate(val) Then
28230         mudtProps.DOCDate = CDate(val)
28240         SetDirty True
28250     End If
28260     Exit Property
errHandler:
28270     If ErrMustStop Then Debug.Assert False: Resume
28280     ErrorIn "a_Invoice.DocDate(val)", val
End Property
Public Property Get ProcessingDate() As Date
28290     On Error GoTo errHandler
28300     ProcessingDate = mudtProps.ProcessingDate
28310     Exit Property
errHandler:
28320     If ErrMustStop Then Debug.Assert False: Resume
28330     ErrorIn "a_Invoice.ProcessingDate"
End Property
Public Property Get ProcessingDateF() As String
28340     On Error GoTo errHandler
28350     ProcessingDateF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy")
28360     Exit Property
errHandler:
28370     If ErrMustStop Then Debug.Assert False: Resume
28380     ErrorIn "a_Invoice.ProcessingDateF"
End Property
Public Property Get ProcessingDateFF() As String
28390     On Error GoTo errHandler
28400     ProcessingDateFF = Format(mudtProps.ProcessingDate, "dd/mm/yyyy Hh:Nn")
28410     Exit Property
errHandler:
28420     If ErrMustStop Then Debug.Assert False: Resume
28430     ErrorIn "a_Invoice.ProcessingDateFF"
End Property

Public Property Get CaptureDate() As Date
28440     On Error GoTo errHandler
28450     CaptureDate = mudtProps.CaptureDate
28460     Exit Property
errHandler:
28470     If ErrMustStop Then Debug.Assert False: Resume
28480     ErrorIn "a_Invoice.CaptureDate"
End Property
Public Property Get CaptureDateF() As String
28490     On Error GoTo errHandler
28500     CaptureDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
28510     Exit Property
errHandler:
28520     If ErrMustStop Then Debug.Assert False: Resume
28530     ErrorIn "a_Invoice.CaptureDateF"
End Property
Public Property Let CaptureDate(val As Date)
28540     On Error GoTo errHandler
28550     If mcolStack.Count = 0 Then Err.Raise 383
28560     mudtProps.CaptureDate = val
28570     SetDirty True
28580     Exit Property
errHandler:
28590     If ErrMustStop Then Debug.Assert False: Resume
28600     ErrorIn "a_Invoice.CaptureDate(val)", val
End Property

'Public Property Get ForeignCurrency() As a_Currency
'    Set ForeignCurrency = oCurrency_Foreign
'End Property

Public Property Get CurrencyFactor() As Double
28610     On Error GoTo errHandler
28620     If mudtProps.CurrencyFactor = 0 Then mudtProps.CurrencyFactor = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").Factor
28630     CurrencyFactor = mudtProps.CurrencyFactor
28640     Exit Property
errHandler:
28650     If ErrMustStop Then Debug.Assert False: Resume
28660     ErrorIn "a_Invoice.CurrencyFactor"
End Property
Public Property Let CurrencyFactor(val As Double)
28670     On Error GoTo errHandler
28680     mudtProps.CurrencyFactor = val
28690     SetDirty True
28700     Exit Property
errHandler:
28710     If ErrMustStop Then Debug.Assert False: Resume
28720     ErrorIn "a_Invoice.CurrencyFactor(val)", val
End Property

Public Property Get VATRate() As Double
28730     On Error GoTo errHandler
28740     VATRate = mudtProps.VATRate
28750     Exit Property
errHandler:
28760     If ErrMustStop Then Debug.Assert False: Resume
28770     ErrorIn "a_Invoice.VATRate"
End Property
Public Property Get VATRateFormatted() As String
28780     On Error GoTo errHandler
28790     VATRate = Format(mudtProps.VATRate, "##0.00")
28800     Exit Property
errHandler:
28810     If ErrMustStop Then Debug.Assert False: Resume
28820     ErrorIn "a_Invoice.VATRateFormatted"
End Property
Public Property Let VATRate(val As Double)
28830     On Error GoTo errHandler
28840     If mcolStack.Count = 0 Then Err.Raise 383
28850     mudtProps.VATRate = val
28860     mudtProps.IsDirty = True
28870     Exit Property
errHandler:
28880     If ErrMustStop Then Debug.Assert False: Resume
28890     ErrorIn "a_Invoice.VATRate(val)", val
End Property
Public Property Get QtyNonStandardVAT() As Long
      Dim i As Long
      Dim oIL As a_InvoiceLine

28900     i = 0
28910     For Each oIL In Me.InvoiceLines
28920         If oIL.VATRate <> oPC.Configuration.VATRate Then
28930             i = i + 1
28940         End If
28950     Next
28960     QtyNonStandardVAT = i
End Property
Public Property Get PossibleIncorrectServiceCharge() As Boolean
      Dim oSM As New z_StockManager

28970     PossibleIncorrectServiceCharge = oSM.PossibleIncorrectServiceCharge(Me.InvoiceID)
          
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
28980     On Error GoTo errHandler
28990   VATable = Me.Customer.VATable
       ' VATable = mudtProps.VATable
29000     Exit Property
errHandler:
29010     If ErrMustStop Then Debug.Assert False: Resume
29020     ErrorIn "a_Invoice.VATable"
End Property
Public Property Let VATable(val As Boolean)
29030     On Error GoTo errHandler
29040     mudtProps.VATable = val
29050     SetDirty True
29060     Exit Property
errHandler:
29070     If ErrMustStop Then Debug.Assert False: Resume
29080     ErrorIn "a_Invoice.VATable(val)", val
End Property


Public Property Get InvoiceLines() As ch_InvoiceLine
29090   Set InvoiceLines = mColInvoiceLines
End Property
Public Property Get DisplayInvoice() As d_Invoice
29100     On Error GoTo errHandler
29110     Set DisplayInvoice = New d_Invoice
29120     Exit Property
errHandler:
29130     If ErrMustStop Then Debug.Assert False: Resume
29140     ErrorIn "a_Invoice.DisplayInvoice"
End Property
Private Property Get StatusForPrinting() As String
29150     On Error GoTo errHandler
              
29160     If mudtProps.Proforma = True Then
29170         If mudtProps.Status = stVOID Then
29180             StatusForPrinting = "VOID"
29190         ElseIf mudtProps.Status = stInProcess Then
29200             StatusForPrinting = "IN PROCESS"
29210         ElseIf mudtProps.Status = stISSUED Then
29220             StatusForPrinting = ""
29230         ElseIf mudtProps.Status = stCOMPLETE Then
29240             StatusForPrinting = ""
29250         ElseIf mudtProps.Status = stPROFORMA Then
29260             StatusForPrinting = ""
29270         ElseIf mudtProps.Status = stCANCELLED Then
29280             StatusForPrinting = "CANCELLED"
29290         End If
29300     Else
29310         If mudtProps.Status = stVOID Then
29320             StatusForPrinting = "VOID"
29330         ElseIf mudtProps.Status = stInProcess Then
29340             StatusForPrinting = "IN PROCESS"
29350         ElseIf mudtProps.Status = stISSUED Then
29360             If oPC.AllowsInvoicePicking Then
29370                 StatusForPrinting = "PICKED (Not finalized)"
29380             Else
29390                 StatusForPrinting = ""
29400             End If
29410         ElseIf mudtProps.Status = stCOMPLETE Then
29420             StatusForPrinting = ""
29430         ElseIf mudtProps.Status = stPROFORMA Then
29440             StatusForPrinting = "PROFORMA"
29450         ElseIf mudtProps.Status = stCANCELLED Then
29460             StatusForPrinting = "CANCELLED"
29470         End If
29480     End If
29490     Exit Property
errHandler:
29500     If ErrMustStop Then Debug.Assert False: Resume
29510     ErrorIn "a_Invoice.StatusForPrinting"
End Property
Public Property Get StatusF() As String
29520     On Error GoTo errHandler
29530     If mudtProps.Status = stVOID Then
29540         StatusF = "VOID"
29550     ElseIf mudtProps.Status = stInProcess Then
29560         StatusF = "IN PROCESS"
29570     ElseIf mudtProps.Status = stISSUED Then
29580         If oPC.AllowsInvoicePicking And Not Me.Proforma Then
29590             StatusF = "PICKING"
29600         Else
29610             StatusF = "ISSUED"
29620         End If
29630     ElseIf mudtProps.Status = stCOMPLETE Then
29640         If oPC.AllowsInvoicePicking And Not Me.Proforma Then
29650             StatusF = "ISSUED"
29660         Else
29670             StatusF = "COMPLETE"
29680         End If
29690     ElseIf mudtProps.Status = stPROFORMA Then
29700         StatusF = "PROFORMA"
29710     ElseIf mudtProps.Status = stCANCELLED Then
29720         StatusF = "CANCELLED"
29730     End If
29740     Exit Property
errHandler:
29750     If ErrMustStop Then Debug.Assert False: Resume
29760     ErrorIn "a_Invoice.StatusF"
End Property
Public Property Get Status() As enStatus
29770     On Error GoTo errHandler
29780     Status = mudtProps.Status
29790     Exit Property
errHandler:
29800     If ErrMustStop Then Debug.Assert False: Resume
29810     ErrorIn "a_Invoice.Status"
End Property
Public Function SetForAttn(val As String) As Boolean
29820     If Len(val) > Len(mudtProps.ForAttn) Then
29830         Err.Raise 384
29840     End If
29850     If val <> mudtProps.ForAttn Then
29860         mudtProps.ForAttn = val
29870         SetDirty True
29880     End If
29890     SetForAttn = True
29900     Exit Function
End Function
Public Function SetMemo(val As String) As Boolean
29910     If Len(val) > Len(mudtProps.Memo) Then
29920         Err.Raise 384
29930     End If
29940     If val <> mudtProps.Memo Then
29950         mudtProps.Memo = val
29960         SetDirty True
29970     End If
29980     SetMemo = True
29990     Exit Function
End Function
Public Function SetWaybill(val As String) As Boolean
30000     If Len(val) > Len(mudtProps.Waybill) Then
30010         Err.Raise 384
30020     End If
30030     If val <> mudtProps.Waybill Then
30040         mudtProps.Waybill = Left(val, 50)
30050         SetDirty True
30060     End If
30070     SetWaybill = True
30080     Exit Function
End Function
Public Function SetCourierURL(val As String) As Boolean
30090     If Len(val) > Len(mudtProps.CourierURL) Then
30100         Err.Raise 384
30110     End If
30120     If val <> mudtProps.CourierURL Then
30130         mudtProps.CourierURL = Left(val, 90)
30140         SetDirty True
30150     End If
30160     SetCourierURL = True
30170     Exit Function
End Function
Public Function SetMemoDirect(val As String) As Boolean
      Dim oSM As z_StockManager
30180     If Len(val) > Len(mudtProps.Memo) Then
30190         Err.Raise 384
30200     End If
30210     Set oSM = New z_StockManager
30220     oSM.SetMemo val, Me.InvoiceID
          
End Function
Public Property Get Memo() As String
30230     On Error GoTo errHandler
30240     Memo = FNS(mudtProps.Memo)
30250     Exit Property
errHandler:
30260     If ErrMustStop Then Debug.Assert False: Resume
30270     ErrorIn "a_Invoice.Memo"
End Property
'Public Property Get ExchangeID() As String
'    On Error GoTo errHandler
'    ExchangeID = FNS(mudtProps.ExchangeID)
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Invoice.ExchangeID"
'End Property
Public Property Get ForAttn() As String
30280     On Error GoTo errHandler
30290     ForAttn = FNS(mudtProps.ForAttn)
30300     Exit Property
errHandler:
30310     If ErrMustStop Then Debug.Assert False: Resume
30320     ErrorIn "a_Invoice.ForAttn"
End Property
Public Property Get Waybill() As String
30330     On Error GoTo errHandler
          
30340     If FNS(mudtProps.Waybill) = "" Then
30350         Waybill = "n/a"
30360     Else
30370         Waybill = FNS(mudtProps.Waybill)
30380     End If
30390     Exit Property
errHandler:
30400     If ErrMustStop Then Debug.Assert False: Resume
30410     ErrorIn "a_Invoice.Waybill"
End Property
Public Property Get CourierURL() As String
30420     On Error GoTo errHandler

30430     CourierURL = FNS(mudtProps.CourierURL)
30440     Exit Property
errHandler:
30450     If ErrMustStop Then Debug.Assert False: Resume
30460     ErrorIn "a_Invoice.CourierURL"
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
30470     On Error GoTo errHandler
      Dim cmd As ADODB.Command
      Dim prm As ADODB.Parameter
      Dim strDefaultCo As String

30480     Set cmd = New ADODB.Command
30490     cmd.ActiveConnection = oPC.COShort
30500     If mudtProps.Proforma Then
30510         cmd.CommandText = "sp_GetNextProformaNo"
30520     Else
30530         If mudtProps.IsPreDelivery Then
30540             cmd.CommandText = "sp_GetNextPreDelivInvNo"
30550         Else
30560             cmd.CommandText = "sp_GetNextInvNo"
30570         End If
30580     End If
30590     cmd.commandType = adCmdStoredProc
          
          ' Get parameter value and append parameter.
30600     Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , COMPID)
30610     cmd.Parameters.Append prm
30620     Set prm = cmd.CreateParameter("@NewCode", adVarChar, adParamOutput, 15)
30630     cmd.Parameters.Append prm
30640     cmd.execute
30650     GetInvoiceNumber = cmd.Parameters("@NewCode").Value
30660     Set cmd = Nothing
30670     Exit Function
errHandler:
30680     If ErrMustStop Then Debug.Assert False: Resume
30690     ErrorIn "a_Invoice.GetInvoiceNumber(strTmp)", strTmp
End Function

Property Get BillTOAddress() As a_Address
30700     Set BillTOAddress = oBillToAddress
End Property
Property Get DelToAddress() As a_Address
30710     Set DelToAddress = oDelToAddress
End Property
'Public Function PrintInvoice_Display(Optional ByVal pRoundedUp As Boolean) As String
'30720     On Error GoTo errHandler
'      Dim bDiscountExists As Boolean
'      Dim bForeign As Boolean
'      Dim oIL As a_InvoiceLine
'      Dim strInvoice As String
'      Dim i As Long
'
'
'30730     If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
'30740         bForeign = True
'30750         CalculateTotal_Foreign
'30760     End If
'
'
'
'      '    bForeign = (oCurrency_Foreign Is oPC.Configuration.LocalCurrency)
'30770     strInvoice = "Document No. " & Me.DOCCode & vbCrLf
'30780     For i = 1 To Me.InvoiceLines.Count
'30790         Set oIL = Me.InvoiceLines.FindLineBySeqNum(i)
'30800         If Not oIL Is Nothing Then
'30810             If Not oIL.IsDeleted Then
'30820                 If oIL.BottomOfDocument = False Then
'30830                     With oIL
'30840                         If .DiscountPercent > 0 Then bDiscountExists = True
'30850                         If .ServiceItem Then
'30860                             strInvoice = strInvoice & .TitleAuthorPublisher & "  " & .PAfterDiscountExtF(bForeign)
'30870                         Else
'30880                             strInvoice = strInvoice & .CodeF & "  " & .Ref & "  " & .QtyF & "  " & Left(.TitleAuthorPublisher, 30) & "  " & .PriceF(bForeign) & "  " & IIf(.DiscountPercent = 0, "", .DiscountPercentF) & "  " & .PAfterDiscountExtF(bForeign)
'30890                         End If
'30900                     End With
'30910                 strInvoice = strInvoice & vbCrLf
'30920                 Else
'30930                     iExtrasCount = iExtrasCount + 1
'30940                 End If
'30950             End If
'30960         End If
'30970     Next i
'30980     strInvoice = strInvoice & vbCrLf
'30990     If iExtrasCount > 0 Then  'then subtotal 2 is necessary
'31000         strInvoice = strInvoice & "subtotal    " & Me.TotalExtWithOutExtrasF(bForeign) & vbCrLf
'31010         For i = 1 To InvoiceLines.Count
'31020             If InvoiceLines(i).BottomOfDocument = True Then
'31030                 With InvoiceLines(i)
'31040                     If .ServiceItem Then
'31050                         strInvoice = strInvoice & .TitleAuthorPublisher & "    " & .PriceF(bForeign) & vbCrLf
'31060                     End If
'31070                 End With
'31080             End If
'31090         Next i
'31100     End If
'31110     If Me.VATable Then ''''''''''(1)
'31120         strInvoice = strInvoice & "Total " & TotalPayableF(bForeign) & vbCrLf
'31130         strInvoice = strInvoice & "Includes VAT of " & TotalVATF(bForeign) & vbCrLf
'31140     Else  'overseas clients
'31150         If Me.ShowVAT Then
'31160             strInvoice = strInvoice & "subtotal    " & TotalAfterDiscountExtF(bForeign) & vbCrLf   'subtotal 3
'31170             strInvoice = strInvoice & "Less VAT @ " & Format((dblVATRATE / 100), "##.00%") & "  " & TotalVATF(bForeign) & vbCrLf
'31180             strInvoice = strInvoice & "Payable    " & TotalPayableF(bForeign) & vbCrLf
'31190         Else
'31200             strInvoice = strInvoice & "Total   " & TotalPayableF(bForeign) & vbCrLf
'31210         End If
'31220     End If
'31230     If TotalDeposit(bForeign) > 0 Then
'31240         strInvoice = strInvoice & "Less deposits paid: " & TotalDepositF(bForeign)
'31250     End If
'31260     PrintInvoice_Display = strInvoice
'
'EXIT_Handler:
'      'Err_Handler:
'      '    oPCL.FinishRequest
'      '    Select Case
'      '    Case 5941
'      '        MsgBox "Bookmark missing"
'      '        Resume Next
'      '    Case Else
'      '        MsgBox error
'      '        GoTo EXIT_Handler
'      '        Resume
'      '    End Select
'      '
'31270     Exit Function
'errHandler:
'31280     ErrPreserve
'31290     oPCL.FinishRequest
'31300     If ErrMustStop Then Debug.Assert False: Resume
'31310     ErrorIn "a_Invoice.PrintInvoice_Display(pRoundedUp)", pRoundedUp
'End Function
'Public Function PrintInvoice(bBySeqNo As Boolean, Optional ByVal pRoundedUp As Boolean, Optional pQtyCopies As Integer) As Boolean
'31320     On Error GoTo errHandler
'      Dim bForeign As Boolean
'      Dim oIL As a_InvoiceLine
'      Dim oCHKIL As a_InvoiceLine
'      Dim bReferenceExists As Boolean
'      Dim bDiscountExists As Boolean
'      Dim bDepositExists As Boolean
'      Dim bQtyColumnNeeded As Boolean
'      Dim strTag As String
'      Dim strLocalPrinter As String
'      Dim oDC As a_DocumentControl
'
'
'       '   ExportToXML False
'      'Four situations obtain
'      '   1. In a VAT operating area
'      '       a) VAT is charged      (situation 1)
'      '       b) VAT is returned (e.g. to overseas client) (situation 2)
'      '       c) VAT is charged and not returned (e.g.to overseas client) (situation 3)
'      '   2. In an area where VAT does not apply (situation 4)
'
'      'Layout
'          'Section 1: Detail lines
'          '   subtotal1 (lngTotal1)
'          'Section 2: possible discount
'          '   subtotal2  (lngTotal2)
'          'Section 3: possible extras
'          '   subtotal3 (lngTotal3)
'          'Section 4: possible deposits paid
'          '   subtotal4 (lngTotal4)
'          'Section 5: VAT payable/refundable/notnotified
'          '   TotalPayable,VATAmount
'
'31330     bForeign = False
'31340     If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
'31350         bForeign = True
'31360         CalculateTotal_Foreign
'31370     End If
'
'31380     Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
'31390     If oDC Is Nothing Then
'31400         PrintInvoice = False
'31410         Exit Function
'31420     Else
'31430         Set oPCL = New z_PrintClient
'31440         If Me.Proforma Then
'31450             strTag = "PR1_"
'31460         Else
'31470             strTag = "IN1_"
'31480         End If
'31490         If oPCL.StartRequestPrinting(strTag & Me.DOCCode) = False Then
'31500             PrintInvoice = False
'31510             Exit Function
'31520         End If
'31530         With oDC
'31540             oPCL.PrinterSetup .PrinterName(), .PreviewPrint, .style, IIf(pQtyCopies > 0, pQtyCopies, .QtyCopies)
'31550         End With
'31560         PrintInvoice = True
'31570     End If
'
'
'31580     bDiscountExists = False
'31590     bReferenceExists = False
'31600     bDepositExists = False
'31610     bQtyColumnNeeded = False
'31620     For Each oCHKIL In InvoiceLines
'31630         If oCHKIL.Ref > " " Then
'31640             bReferenceExists = True
'31650         End If
'31660         If oCHKIL.DiscountPercent > 0 Then
'31670             bDiscountExists = True
'31680         End If
'31690         If oCHKIL.Deposit > 0 Then
'31700             bDepositExists = True
'31710         End If
'31720         If oCHKIL.Qty > 1 Or oPC.Configuration.AntiquarianYN = False Then
'31730             bQtyColumnNeeded = True
'31740         End If
'31750     Next
'31760     oPCL.SendPrintingSettings oPC.Configuration.PrintingSettings
'31770     oPCL.HideDiscountColumn Not bDiscountExists
'31780     oPCL.HideReferenceColumn Not bReferenceExists
'31790     oPCL.HideDepositColumn Not bDepositExists
'31800     oPCL.HideQtyColumn Not bQtyColumnNeeded
'
'31810     If Me.StatusF <> "ISSUED" And Me.StatusF <> "COMPLETE" Then
'31820         oPCL.WriteTag "STATUS", Me.StatusF
'31830     End If
'31840     oPCL.LineCOunt Me.InvoiceLines.Count
'31850     oPCL.LineTotalCount 4
'31860     oPCL.WriteTag "MEMO", PackText(Me.Memo)
'31870     oPCL.WriteTag "OURADDRESS", PackText(oPC.Configuration.DefaultCompany.StreetAddress)
'31880     oPCL.WriteTag "PHONE", IIf(Me.TPPhone > "", "Phone: " & Me.TPPhone, "")
'31890     If Me.COMPID > 0 Then
'31900         oPCL.WriteTag "COREGNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CoRegistrationNumber
'31910         oPCL.WriteTag "VATNUMBER", oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).VatNumber
'31920     End If
'31930     oPCL.WriteTag "CUSTOMER", CStr(Trim(Me.Customer.Fullname)) '& IIf(Me.TPACCNum > "", "A(" & Me.TPACCNum & ")", "")
'31940     oPCL.WriteTag "FAX", IIf(BillTOAddress.FaxF > "", "Fax: " & BillTOAddress.FaxF, "")
'31950     oPCL.WriteTag "BILLTO", PackText(BillTOAddress.AddressMailing)
'31960     oPCL.WriteTag "DELTO", PackText(DelToAddress.AddressMailing)
'31970     oPCL.WriteTag "CUSTVATNUMBER", Me.Customer.VatNumber
'31980     oPCL.WriteTag "DATE", Me.TransDateF2
'31990     If Not Me.BillingCompany Is Nothing Then
'32000         oPCL.WriteTag "CONAME", Me.BillingCompany.CompanyName
'32010     Else
'32020         oPCL.WriteTag "CONAME", ""
'32030     End If
'32040     oPCL.WriteTag "DOCUMENTCODE", Me.DOCCode
'32050     oPCL.WriteTag "ACCNUMBER", IIf(Me.Customer.AcNo > "", "Acc Num: " & Me.Customer.AcNo, "")
'32060     oPCL.WriteTag "SM", IIf(Me.StaffName > "", Me.StaffName, "")
'32070     oPCL.WriteTag "BANK", PackText(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).BankDetails)
'
'32080     iExtrasCount = FillTableRows(bBySeqNo, bForeign, bDiscountExists)
'
'32090     If iExtrasCount > 0 Then  'then subtotal 2 is necessary
'32100         oPCL.WriteTag "TOTAL", "subtotal" & "|" & Me.TotalExtWithOutExtrasF(bForeign)
'32110         FillExtras Me, bForeign, bDiscountExists
'32120     End If
'32130     If Me.VATable Then ''''''''''(1)
'32140         oPCL.WriteTag "TOTAL", "Payable " & "|" & TotalPayableF(bForeign)
'32150         oPCL.WriteTag "TOTAL2", "(Includes VAT of " & "|" & TotalVATF(bForeign) & ")"
'32160     Else  'overseas clients
'32170         If Me.ShowVAT Then
'32180             oPCL.WriteTag "TOTAL", "subtotal" & "|" & TotalAfterDiscountExtF(bForeign)
'32190             oPCL.WriteTag "TOTAL", "Less VAT @ " & Format((dblVATRATE / 100), "##.00%") & "|" & TotalVATF(bForeign)
'32200             oPCL.WriteTag "TOTAL", "Payable " & "|" & TotalPayableF(bForeign)
'32210         Else
'32220             oPCL.WriteTag "TOTAL", "Total " & "|" & TotalPayableF(bForeign)
'32230         End If
'32240     End If
'32250     If TotalDeposit(bForeign) > 0 Then
'32260         oPCL.WriteTag "TOTAL", "Less deposits paid   " & "|" & TotalDepositF(bForeign)
'32270         oPCL.WriteTag "TOTAL", "Balance owed =   " & "|" & TotalBalanceAfterDeposits(bForeign)
'32280     End If
'32290     oPCL.FinishRequest
'
'EXIT_Handler:
'      'Err_Handler:
'      '    oPCL.FinishRequest
'      '    Select Case
'      '    Case 5941
'      '        MsgBox "Bookmark missing"
'      '        Resume Next
'      '    Case Else
'      '        MsgBox error
'      '        GoTo EXIT_Handler
'      '        Resume
'      '    End Select
'32300     Exit Function
'errHandler:
'32310     If ErrMustStop Then Debug.Assert False: Resume
'32320     ErrorIn "a_Invoice.PrintInvoice(pRoundedUp)", pRoundedUp
'End Function
Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PRINTINVOICE
32330     On Error GoTo errHandler
      Dim strLabels As String
      Dim strValues As String
      Dim i As Integer
32340     strLabels = ""
32350     strValues = ""
32360     If iExtrasCount > 0 Then  'then subtotal 2 is necessary
32370         strLabels = "subtotal"
32380         strValues = Me.TotalExtWithOutExtrasF(bForeign)
32390         For i = 1 To Me.InvoiceLines.Count
32400             If Me.InvoiceLines(i).BottomOfDocument = True Then
32410                 With Me.InvoiceLines(i)
32420                     If .ServiceItem Then
32430                         strLabels = strLabels & Chr(13) & .TitleAuthorPublisher
32440                         strValues = strValues & Chr(13) & .row_LTNNF
32450                     End If
32460                 End With
32470             End If
32480         Next i
32490     End If
      '    If Len(strLabels) > 0 Then
      '        strLabels = strLabels & Chr(13)
      '        strValues = strValues & Chr(13)
      '    End If
32500     If Me.VATable Then ''''''''''(1)
32510         strLabels = strLabels & Chr(13) & "Total"
32520         strValues = strValues & Chr(13) & TotalPayableF(bForeign)
32530         strLabels = strLabels & Chr(13) & "Includes VAT of"
32540         strValues = strValues & Chr(13) & TotalVATF(bForeign)
32550         strLabels = strLabels & Chr(13) & "Total qty items"
32560         strValues = strValues & Chr(13) & TotalQty
32570     Else  'overseas clients
32580         If Me.ShowVAT Then
32590             strLabels = strLabels & Chr(13) & "subtotal"
32600             strValues = strValues & Chr(13) & TotalAfterDiscountExtF(bForeign)
32610             strLabels = strLabels & Chr(13) & "Less VAT of"
32620             strValues = strValues & Chr(13) & TotalVATF(bForeign)
32630             strLabels = strLabels & Chr(13) & "Payable"
32640             strValues = strValues & Chr(13) & TotalPayableF(bForeign)
32650             strLabels = strLabels & Chr(13) & "Total qty items"
32660             strValues = strValues & Chr(13) & TotalQty
32670         Else
32680             strLabels = strLabels & Chr(13) & "Total"
32690             strValues = strValues & Chr(13) & TotalPayableF(bForeign)
32700             strLabels = strLabels & Chr(13) & "Total qty items"
32710             strValues = strValues & Chr(13) & TotalQty
32720         End If
32730     End If
32740     If TotalDeposit(bForeign) > 0 Then
32750         strLabels = strLabels & Chr(13) & "Less deposits paid"
32760         strValues = strValues & Chr(13) & TotalDepositF(bForeign)
32770         strLabels = strLabels & Chr(13) & "(balance owed ="
32780         strValues = strValues & Chr(13) & TotalBalanceAfterDeposits(bForeign) & ")"
32790         strLabels = strLabels & Chr(13) & "Total qty items"
32800         strValues = strValues & Chr(13) & TotalQty
32810     End If
32820     pLabels = strLabels
32830     pValues = strValues
32840     Exit Function
errHandler:
32850     If ErrMustStop Then Debug.Assert False: Resume
32860     ErrorIn "a_Invoice.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
Private Function FillTableRows(bBySeqNo As Boolean, bForeign As Boolean, pDiscountExists As Boolean)
32870     On Error GoTo errHandler
      Dim i As Long
      Dim iExtrasCount As Integer
      Dim oIL As a_InvoiceLine

32880     iExtrasCount = 0
32890     For i = 1 To Me.InvoiceLines.Count
32900         If bBySeqNo Then
32910             Set oIL = Me.InvoiceLines.FindLineBySeqNum(i)
32920         Else
32930             Set oIL = Me.InvoiceLines(i)
32940         End If
32950         If Not oIL Is Nothing Then
32960             If oIL.BottomOfDocument = False Then
32970                 With oIL
32980                     If .DiscountPercent > 0 Then pDiscountExists = True
32990                     If .ServiceItem Then
33000                         oPCL.ILSend "", "", "", .Title & .SpecialVATMsg, "", "", "", .PAfterDiscountExtF(bForeign), "", "", "", .Note
33010                     Else
33020                         oPCL.ILSend .CodeForExport, .Ref, .QtyF, .SpecialVATMsg & .Title, .MainAuthor, .Publisher, .PriceF(bForeign), .PAfterDiscountExtF(bForeign), .DepositF(bForeign), .DiscountPercentF, .PAfterDiscountExtNonVATF(bForeign), PackText(.Note)
33030                     End If
33040                 End With
33050             Else
33060                 iExtrasCount = iExtrasCount + 1
33070             End If
33080         End If
33090     Next i
33100     FillTableRows = iExtrasCount
EXIT_Handler:
33110     Exit Function
errHandler:
33120     If ErrMustStop Then Debug.Assert False: Resume
33130     ErrorIn "a_Invoice.FillTableRows(bBySeqNo,bForeign,pDiscountExists)", Array(bBySeqNo, bForeign, _
               pDiscountExists)
End Function
Private Function FillExtras(pINV As a_Invoice, bForeign As Boolean, pDiscountExists As Boolean) As Integer
33140     On Error GoTo errHandler
      Dim i As Integer
      Dim iCnt As Integer

33150     iCnt = 0
33160     For i = 1 To pINV.InvoiceLines.Count
33170         If pINV.InvoiceLines(i).BottomOfDocument = True Then
33180             With pINV.InvoiceLines(i)
33190                 If .ServiceItem Then
33200                     iCnt = iCnt + 1
33210                     oPCL.ILSendExtra .TitleAuthorPublisher, .PAfterDiscountExtF(bForeign)
33220                 End If
33230             End With
33240         End If
33250     Next i
33260     FillExtras = iCnt
EXIT_Handler:
      'Err_Handler:
      '    MsgBox error
      '    GoTo EXIT_Handler
         ' Resume
33270     Exit Function
errHandler:
33280     If ErrMustStop Then Debug.Assert False: Resume
33290     ErrorIn "a_Invoice.FillExtras(pInv,bForeign,pDiscountExists)", Array(pINV, bForeign, _
               pDiscountExists)
End Function



Private Function TranslateErrors(pRawors As String) As String
33300     On Error GoTo errHandler
      Dim strOut As String
      Dim strRule, strAllRules As String
      Dim NoMoreRules As Boolean
      Dim iMarker, iStart As Integer
33310     iMarker = 1
33320     strAllRules = ""
33330     If Len(pRawors) > 0 Then
33340         iMarker = InStr(iMarker + 1, pRawors, ",")
33350         If iMarker > 0 Then
33360             strAllRules = colClassors(Left(pRawors, iMarker - 1))
33370         Else
33380             strAllRules = colClassors(pRawors)
33390         End If
33400         Do Until iMarker = 0
33410             iStart = iMarker + 1
33420             iMarker = InStr(iStart, pRawors, ",")
33430             If iMarker > 0 Then
33440                 strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
33450             Else
33460                 strRule = colClassors(Mid(pRawors, iStart))
33470             End If
                      
33480             strAllRules = strAllRules & vbCrLf & strRule
33490         Loop
33500     End If
33510     TranslateErrors = strAllRules
33520     Exit Function
errHandler:
33530     If ErrMustStop Then Debug.Assert False: Resume
33540     ErrorIn "a_Invoice.TranslateErrors(pRawors)", pRawors
End Function
Private Sub LoadClassorsCollection()
33550     On Error GoTo errHandler
33560     Set colClassors = New Collection
33570     colClassors.Add "Missing customer", "TP"
33580     colClassors.Add "Note too short", "NOTE"
33590     colClassors.Add "Invalid discount", "DISCOUNT"
33600     colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
33610     Exit Sub
errHandler:
33620     If ErrMustStop Then Debug.Assert False: Resume
33630     ErrorIn "a_Invoice.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pVal As String, pFld As String) As Boolean
33640     On Error GoTo errHandler
      Dim bValid As Boolean
      Dim s As String

33650     bValid = True
33660     If mudtProps.TPID = 0 Then
33670         mobjValid.BreakRule "TP", True
33680         If pFld = "TP" Then bValid = False
33690     Else
33700         mobjValid.BreakRule "TP", False
33710     End If
33720     mobjValid.GetStatus
33730     SetDirty bValid
33740     ValidateObject = bValid
33750     Exit Function
errHandler:
33760     If ErrMustStop Then Debug.Assert False: Resume
33770     ErrorIn "a_Invoice.ValidateObject(pVal,pFld)", Array(pVal, pFld)
End Function



Public Function Post(Optional pNewStatus As enStatus) As String
33780     On Error GoTo errHandler
      Dim strResult As String
      Dim oSM As z_StockManager
      Dim OpenResult As Integer
      '-------------------------------
33790     OpenResult = oPC.OpenDBSHort
      '-------------------------------
33800     oPC.COShort.CommandTimeout = 0
          
          
33810     strResult = Me.ApplyEdit
33820     If strResult > "" Then  'probably a timeout or
             'Not necessary oPC.COShort.execute "ROLLBACK TRANSACTION"
33830         Post = strResult
33840         Exit Function
33850     End If
          
          
33860     Set oSM = New z_StockManager
33870     If pNewStatus = 0 Then
33880         pNewStatus = Me.Status
33890     End If
33900     If Not Me.Proforma Then
33910         If pNewStatus > 0 Then SetStatus pNewStatus
33920         Post = ""
33930         Select Case mOriginalStatus
              Case stVOID      'VOID
33940                 Post = "VOID"
33950         Case stInProcess      'In Process
33960                 Select Case pNewStatus
                      Case stVOID
                 ''         SetStatus stVOID
33970                 Case stInProcess
33980                     Post = "In Process"
33990                 Case stCOMPLETE, stISSUED
34000                     strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, pNewStatus)
34010                     If strResult > "" Then
34020                         oSM.RollbackTransaction oPC.COShort
34030                         Post = strResult
34040                         GoTo EXITH
34050                     End If
                        '  SetStatus stCOMPLETE
34060                 Case stPROFORMA
34070                     strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, stCOMPLETE)
34080                     If strResult > "" Then
34090                         oSM.RollbackTransaction oPC.COShort
34100                         Post = strResult
34110                         GoTo EXITH
34120                     End If
                  '        SetStatus stPROFORMA
34130                 End Select
34140         Case stISSUED
34150                 If oPC.AllowsInvoicePicking Then
34160                     Select Case pNewStatus
                          Case stCOMPLETE, stISSUED
34170                         strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, pNewStatus)
34180                         If strResult > "" Then
34190                             oSM.RollbackTransaction oPC.COShort
34200                             Post = strResult
34210                             GoTo EXITH
34220                         End If
34230                     End Select
34240                 End If
34250         Case stCOMPLETE      'Complete
34260                 Post = "ISSUED"
34270                 If oPC.AllowsInvoicePicking Then
34280                     Select Case pNewStatus
                          Case stCOMPLETE, stISSUED
34290                         strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, pNewStatus)
34300                         If strResult > "" Then
34310                             oSM.RollbackTransaction oPC.COShort
34320                             Post = strResult
34330                             GoTo EXITH
34340                         End If
34350                     End Select
34360                 End If
34370         Case stPROFORMA      'Issued PROFORMA
34380                 Select Case Status
                      Case stVOID
34390                     oSM.CancelInvoice Me
                  '        SetStatus stVOID
34400                 Case stInProcess
34410                     Post = "OR: was proforma, can't now be inprocess."
34420                 Case stCOMPLETE
34430                     strResult = oSM.IssueInvoice(Me, Me.IsPreDelivery, stCOMPLETE)
34440                     If strResult > "" Then
34450                         oSM.RollbackTransaction oPC.COShort
34460                         Post = strResult
34470                         GoTo EXITH
34480                     End If
                  '        SetStatus stCOMPLETE
34490                 Case stPROFORMA
34500                     Post = "PROFORMA"
34510                 End Select
34520         End Select
34530     End If
         ' oPC.COShort.execute "COMMIT TRANSACTION"
34540     oSM.CalculateInvoiceTableSummaryData Me.InvoiceID, 0
       '  oSM.AUTOGenerateReturnFromInvoice Me.InvoiceID
      '---------------------------------------------------
34550     If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
EXITH:
34560     Post = strResult
34570     Set oSM = Nothing
34580     Exit Function
errHandler:
34590     ErrPreserve
34600     Post = strResult
          
34610     If ErrMustStop Then Debug.Assert False: Resume
34620     RlsObjs oPC.COShort
34630     ErrorIn "a_Invoice.Post", , , oPC.COShort
End Function
Public Sub SetStatus(val As enStatus)
34640     On Error GoTo errHandler
       '   If mcolStack.Count = 0 Then Err.Raise 383
34650     mudtProps.Status = val
34660     SetDirty True
34670     Exit Sub
errHandler:
34680     If ErrMustStop Then Debug.Assert False: Resume
34690     ErrorIn "a_Invoice.SetStatus(val)", val
End Sub
Public Property Get COLsOSPerCUST() As c_COLsPerCust
34700   Set COLsOSPerCUST = cColsOSPerCust
End Property
Public Property Get APPLsOSPerCUST() As c_APPLsPerTPPID
34710   Set APPLsOSPerCUST = cAPPLsOS
End Property
Public Function VoidDocument()
34720     On Error GoTo errHandler
      Dim strResult As String
34730     Me.BeginEdit
34740     Me.SetStatus stVOID
34750     Me.ApplyEdit
34760     Exit Function
errHandler:
34770     If ErrMustStop Then Debug.Assert False: Resume
34780     ErrorIn "a_Invoice.VoidDocument"
End Function
Public Sub Reload()
34790     On Error GoTo errHandler
      Dim lngTRID As Long
34800     lngTRID = Me.InvoiceID
34810     Class_Terminate
34820     Class_Initialize
34830     Load lngTRID, True
34840     Exit Sub
errHandler:
34850     If ErrMustStop Then Debug.Assert False: Resume
34860     ErrorIn "a_Invoice.Reload"
End Sub

Public Sub SetProforma()
34870     mudtProps.Proforma = True
34880     SetDirty True
          
End Sub
Public Property Get Proforma() As Boolean
34890     Proforma = mudtProps.Proforma
End Property

Public Sub InformLocalPOSdb()
      Dim DBLocalConn As ADODB.Connection
34900     Set DBLocalConn = New ADODB.Connection
34910     DBLocalConn.Provider = "sqloledb"
34920     DBLocalConn.Open "Data provider=sqloledb;Data Source=" & oPC.POSServername & ";Initial Catalog=PBKSFD;User Id=sa" & ";Password=" & "" & "; Network Library=dbmssocn;"
34930     DBLocalConn.CommandTimeout = 0
34940         DBLocalConn.execute "INSERT INTO tBRIDGE (INV_ID,INV_DOCCode,INV_DOCDate,INV_CustomerName,INV_Payable,INV_TPID,INV_VATVALUE) VALUES (" & Me.InvoiceID & ",'" & Me.DOCCode & "'," & Me.DOCDate & ",'" & Me.Customer.Name & "'," & Me.TotalPayableRaw(False) & "," & Me.Customer.ID & "," & Me.TotalVATRaw(False) & ")"
34950     DBLocalConn.Close
34960     Set DBLocalConn = Nothing
End Sub
Public Property Get SendersEmail()
34970     SendersEmail = FNS(mudtProps.StaffEmail)
End Property

Public Function ExportToXML(bBySeqNo As Boolean, pFilename As String, _
    Optional ByVal pRoundedUp As Boolean, _
    Optional DispatchMode As enTransmitType, Optional pQtyCopies As Integer, _
    Optional pDestinationEmailAddress As String, Optional pWholeMessage As String, _
    Optional pTemplateToUse As String, Optional QuickPrint As Boolean) As Boolean
    On Error GoTo errHandler
Dim bForeign As Boolean
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim XMLFilePath As String
Dim xslFilePath As String
Dim pdfFilePath As String
Dim FOPDebugPath As String

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

Dim zSQL As New z_SQL

 '   MsgBox "Fiddling with oPC.SharedFolderRoot"
  '  oPC.SharedFolderRoot = "\\DAVID-PC\PBKS_S"
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
    ElseIf DispatchMode = enEDI Then
        strWorkingFolder = oPC.LocalFolder & "TEMP\"
        strDispatchMethod = "EDI"
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
        .docProgID = "msxml2.DOMDocument"
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
                If pTemplateToUse > "" Then
                    .elText = fs.GetBaseName(pTemplateToUse)
                Else
                    If Me.Proforma Then
                      '    If Me.Customer.ProformaTemplateName > "" Then
                      '      .elText = Me.Customer.ProformaTemplateName
                      '    Else
                              If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\INPF_DOC.xslt") Then
                                  .elText = "INPF_DOC.xslt"
                              Else
                                  If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\IN_DOC.XSLT") Then
                                      .elText = "IN_DOC.XSLT"
                                  Else
                                      .elText = "IN_DOC_FO.XSL"
                                  End If
                              End If
                     '     End If
                    Else
                        If Me.Customer.InvoiceTemplateName > "" Then
                            .elText = Me.Customer.InvoiceTemplateName
                        ElseIf fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\IN_DOC.XSLT") Then
                            .elText = "IN_DOC.XSLT"
                        Else
                            .elText = "IN_DOC_FO.XSL"
                        End If
                    End If
                End If
            strTemplateName = .elText
       '     MsgBox "strTemplateName=" & strTemplateName
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
            .elCreateSibling "LogoPath"    'NOTE FOP-1.1 appears not to resize the logo as FOP (old version) does. We need sized logo files
'                If fs.FolderExists(oPC.SharedFolderRoot & "\Executables\FOP-1.1") Then
'                    If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGOSIZED.JPG") Then
'                          .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGOSIZED.JPG"
'                     Else
'                        If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGOSIZED.BMP") Then
'                          .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGOSIZED.BMP"
'                        Else
'                            If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG") Then
'                                .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
'                            Else
'                                 .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
'                            End If
'                        End If
'                    End If
'                Else
                      If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG") Then
                          .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                      Else
                          If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                              .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                          End If
                      End If
              '  End If
             .elText = Replace(.elText, "\", "/")
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "IsProforma"
                .elText = IIf(Proforma, "Y", "N")
'            .elCreateSibling "AccompanyingMessage"
'                .elText = IIf(Proforma, oPC.Configuration.EmailProformaMsg, oPC.Configuration.EmailInvMsg)
            .elCreateSibling "DocCode"
                .elText = DOCCodeF
            .elCreateSibling "DocCodeRaw"
                .elText = DOCCode
           .elCreateSibling "StoreCode", True
                .elText = oPC.Configuration.DefaultStore.code
           .elCreateSibling "DocDate", True
                .elText = DocDateF
           .elCreateSibling "Waybill", True
                .elText = Waybill
           .elCreateSibling "CourierURL", True
                .elText = CourierURL
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName
               p 6
'            .elCreateSibling "CDBNo", True
'                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CDBNo
'            .elCreateSibling "UniqueSupplierNo", True
'                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).UniqueSupplierNo
                
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
                    If Me.ForAttn > "" Then
                        .elText = Me.ForAttn & vbCrLf & .elText
                    End If
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
              Dim dblVATRATE As Double
              dblVATRATE = 0
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
                  If InvoiceLines(i).VATRate > dblVATRATE Then
                      dblVATRATE = InvoiceLines(i).VATRate
                  End If
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
                    .elCreateSibling "Author", True
                        .elText = InvoiceLines(i).MainAuthor
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
                    .elCreateSibling "UnitPriceInc", True
                        .elText = InvoiceLines(i).UnitPriceIncF(bForeign)
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
                        .elText = InvoiceLines(i).row_LTNF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = InvoiceLines(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = InvoiceLines(i).row_LTNNF
                    .elCreateSibling "VAT", True
                        .elText = InvoiceLines(i).row_LVF
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
'                    .elCreateSibling "Extension", True
'                        .elText = InvoiceLines(i).PAfterDiscountExtF(bForeign)
                    .elCreateSibling "Extension", True
                        .elText = InvoiceLines(i).row_LTNF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = InvoiceLines(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = InvoiceLines(i).row_LTNNF
                    .elCreateSibling "VAT", True
                        .elText = InvoiceLines(i).row_LVF
'                    .elCreateSibling "ExtensionExVAT", True
'                        .elText = InvoiceLines(i).PAfterDiscountExtNonVATF(bForeign)
'                    .elCreateSibling "VAT", True
'                        .elText = InvoiceLines(i).PAfterDiscountLineValueVATF(bForeign)
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
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CoRegistrationNumber
            .elCreateSibling "VATNumber", True
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).VatNumber
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
            .elCreateSibling "VATRoundingAdjustment", True
                .elText = Me.VATRoundingAdjustmentF()
                
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
                                .elText = "less VAT at " & Format((dblVATRATE / 100), "##.00%") & vbCrLf
                            .elCreateSibling "NumericValue", True
                                .elText = Me.TotalVATF(bForeign) & vbCrLf
                            .navUP
                        End If
                    End If
                    If iExtrasCount > 0 Then  'Doing this here again so as not to interfere with established procedure above
                        If Me.ShowVAT Then
                            .elCreateSibling "Totals", True
                            .chCreate "TotalDescription"
                                .elText = "subtotal" & vbCrLf
                            .elCreateSibling "NumericValue", True
                                .elText = Me.Total_nonExtras_ExVATF(bForeign) & vbCrLf
                            .navUP
                        End If
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
                  .elText = "(Includes VAT at " & Format((dblVATRATE / 100), "##.00%") & vbCrLf  'oPC.Configuration.VATRateF &
              .elCreateSibling "NumericValue", True
                  .elText = Me.TotalVATF(bForeign) & ")" & vbCrLf
              .navUP
            End If
    End With
                                p 12

'FINALLY PRODUCE THE .XML FILE
    strWorkingFolder = oPC.LocalFolder & "PDF\"
    XMLFilePath = strWorkingFolder & "IN_" & Me.DOCCode & ".xml"

    With xMLDoc
        If fs.FileExists(XMLFilePath) Then
            fs.DeleteFile XMLFilePath
        End If
        .docWriteToFile (XMLFilePath), False, "UNICODE", "" 'strHead
    End With
      If Not fs.FileExists(XMLFilePath) Then
          MsgBox "Cannot write XML file.(" & XMLFilePath & ")", vbOKOnly, "Can't do this"
          Exit Function
      End If
                                p 13
'WRITE THE .PDF FILE IF NECESSARY ------------------------------------------------------------------------------------
    Dim tmp As String
    If (Me.Customer.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
       xslFilePath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName
       pdfFilePath = strWorkingFolder & "IN_" & Me.DOCCode & ".PDF"
       FOPDebugPath = strWorkingFolder & "FOPDebug.txt"
        If UCase(Right(strTemplateName, 4)) = "XSLT" Then
          tmp = oPC.LocalFolder & "Executables\FOP-1.1\"
        Else
          tmp = oPC.LocalFolder & "Executables\FOP\"
        End If
        strCommand = tmp & "FOP.BAT -xml " & XMLFilePath & " -xsl " & xslFilePath & " -pdf " & pdfFilePath & " 2>" & FOPDebugPath
            LogSaveToFile "strCommand = " & strCommand
            
'Allow 2 minutes to produce the PDF file
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
        'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
        cnt = 0
        Do While Not fs.FileExists(pdfFilePath) And cnt < 40
            MsgWaitObj 1000
            cnt = cnt + 1
        Loop
    End If
                        p 158
        If fs.FileExists(pdfFilePath) Then
            On Error Resume Next
            SetFileProperty pdfFilePath, "Title", Me.Customer.AcNo
            SetFileProperty pdfFilePath, "Author", Me.Customer.Name
            On Error GoTo errHandler
        End If
                        p 16
      
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSO = New FileSystemObject
    pFilename = pdfFilePath
    
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
    If (Not fs.FileExists(pdfFilePath)) And (DispatchMode = enView Or DispatchMode = enMail) Then
            MsgBox "The PDF file has not been produced: " & pdfFilePath & ". Please contact support.", vbOKOnly, "Can't do this"
    Else
        If DispatchMode = enView Or enPrint Then
            OpenFileWithApplication pdfFilePath, enPDF, QuickPrint
        End If
    End If

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
End Function


Public Function ExportToSpreadsheet(bForeign As Boolean, _
        pFilename As String) As Boolean
40340     On Error GoTo errHandler
      Dim oTF As New z_TextFile
      Dim s As String
      Dim s2 As String
      Dim lngNumberOfLines As Long

40350     ExportToSpreadsheet = False
40360     pFilename = oPC.LocalFolder & "Temp\" & DOCCode & "_" & Format(Now(), "yyyymmddHHnn") & ".xls"
          
40370     oTF.OpenTextFile pFilename
40380     oTF.WriteToTextFile "Status" & vbTab & Me.Status
40390     oTF.WriteToTextFile "DocCode" & vbTab & DOCCodeF
40400     oTF.WriteToTextFile "DocDate" & vbTab & DocDateF
40410     oTF.WriteToTextFile "StoreCode" & vbTab & oPC.Configuration.DefaultStore.code
40420     oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName
40430     oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
40440     oTF.WriteToTextFile "SenderAddressPostal" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).PostalAddress, Chr(13) & Chr(10), Chr(10))
40450     oTF.WriteToTextFile "CustomerName" & vbTab & Customer.Fullname
40460     oTF.WriteToTextFile "CustomerVATNumber" & vbTab & Customer.VatNumber
40470     oTF.WriteToTextFile "ACNO" & vbTab & Customer.AcNo
40480     oTF.WriteToTextFile "OurAcnowithCustomer" & vbTab & Customer.OurACnoWithClient
          
40490     oTF.WriteToTextFile_NoLineTerminator "CustomerWithAddress" & vbTab
40500     If Not BillTOAddress Is Nothing Then
40510         oTF.WriteToTextFile Replace(Me.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
40520     End If
          
40530     oTF.WriteToTextFile_NoLineTerminator "CustomerPhone" & vbTab
40540     If Not BillTOAddress Is Nothing Then
40550         oTF.WriteToTextFile IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "")
40560     End If
40570     oTF.WriteToTextFile_NoLineTerminator "CustomerFax" & vbTab
40580     If Not BillTOAddress Is Nothing Then
40590         oTF.WriteToTextFile IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
40600     End If

40610                     If Not BillTOAddress Is Nothing Then
40620                         s = ForAttn & " " & Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
40630                     Else
40640                         s = ""
40650                     End If
40660     oTF.WriteToTextFile "BillTo" & vbTab & s

40670                     If Not DelToAddress Is Nothing Then
40680                         s = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
40690                     Else
40700                         s = ""
40710                     End If
40720     oTF.WriteToTextFile "DelTo" & vbTab & s

40730                     If Not DelToAddress Is Nothing Then
40740                         s = FNS(Me.DelToAddress.Description)
40750                     Else
40760                         s = ""
40770                     End If
40780     oTF.WriteToTextFile "DispatchMode" & vbTab & s



40790     s = "SKU" & vbTab & "Barcode" & vbTab & "GDNCode" & vbTab & "Title" & vbTab & "Publisher" & vbTab & "Distributor" & vbTab & "DistributorAcno" _
          & vbTab & IIf(oPC.AllowsSSInvoicing, "QtyFirm" & vbTab & "QtySS", "Qty") _
          & vbTab & "Price" & vbTab & "PriceExVat" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab & "CORep" & vbTab & "Note" & vbTab & "Extension" & vbTab & "ExtensionExVat" _
          & vbTab & "VAT" & vbTab & "VATRate" & vbTab & "DiscountValue" & vbTab & "DiscountValueExVAT" & vbTab & "FCName" & vbTab & "FCPrice" & vbTab & "FCFactor"
          
40800     oTF.WriteToTextFile s
                     
40810   lngNumberOfLines = 0
40820   For i = 1 To InvoiceLines.Count
40830     lngNumberOfLines = lngNumberOfLines + 1
40840     If oPC.HideLocalSKUonINV Then
40850         s2 = InvoiceLines(i).CodeForExport
40860     Else
40870         s2 = InvoiceLines(i).CodeF
40880     End If
40890     s = s2
40900     s = s & vbTab & InvoiceLines(i).code
40910     s = s & vbTab & InvoiceLines(i).GDNCode
40920     s = s & vbTab & InvoiceLines(i).TitleAuthor
40930     s = s & vbTab & InvoiceLines(i).Publisher
40940     s = s & vbTab & InvoiceLines(i).DistributorName
40950     s = s & vbTab & InvoiceLines(i).DistributorAcno
40960                     If oPC.AllowsSSInvoicing Then
40970                             s2 = InvoiceLines(i).QtyFirmF & vbCrLf & InvoiceLines(i).QtySSF
40980                     Else
40990                             s2 = InvoiceLines(i).QtyF
41000                     End If
41010     s = s & vbTab & s2
41020     s = s & vbTab & InvoiceLines(i).PriceF(bForeign)
41030     s = s & vbTab & InvoiceLines(i).PriceExVatF(bForeign)
41040     s = s & vbTab & InvoiceLines(i).DiscountPercentF
41050     s = s & vbTab & InvoiceLines(i).Ref
41060     s = s & vbTab & InvoiceLines(i).CO_StaffShortname
41070     s = s & vbTab & stripCRLF(InvoiceLines(i).Note)   ' change made 29/4/2024 to exclude line breaks as it causes issues for Excel interpretation
41080     s = s & vbTab & InvoiceLines(i).PAfterDiscountExtF(bForeign)
41090     s = s & vbTab & InvoiceLines(i).PAfterDiscountExtNonVATF(bForeign)
41100     s = s & vbTab & InvoiceLines(i).PAfterDiscountLineValueVATF(bForeign)
41110     s = s & vbTab & InvoiceLines(i).VATRateF
41120     s = s & vbTab & InvoiceLines(i).PDiscExtF(bForeign)
41130     s = s & vbTab & InvoiceLines(i).PDiscExtExVATF(True)
41140         If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
41150             s2 = oPC.Configuration.Currencies.FindCurrencyByID(InvoiceLines(i).FCID).SYSNAME
41160         Else
41170             s2 = "n/a"
41180         End If
41190     s = s & vbTab & s2
41200         If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
41210             s2 = InvoiceLines(i).ForeignPriceF
41220         Else
41230             s2 = "n/a"
41240         End If
41250     s = s & vbTab & s2
41260         If InvoiceLines(i).FCID > 0 And InvoiceLines(i).FCID <> oPC.Configuration.DefaultCurrencyID Then
41270             s2 = InvoiceLines(i).FCFactorInvF
41280         Else
41290             s2 = "n/a"
41300         End If
41310     s = s & vbTab & s2
          
41320     oTF.WriteToTextFile s
41330 Next

41340     oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(lngNumberOfLines)
41350     oTF.WriteToTextFile "QtyItems" & vbTab & TotalQtyF
41360     oTF.WriteToTextFile "Memo" & vbTab & Memo
          
          
41370     oTF.WriteToTextFile "TotalShipping" & vbTab & TotalServiceItemF(bForeign)
41380     oTF.WriteToTextFile "TotalDiscount" & vbTab & TotalDiscountF(bForeign)
41390     oTF.WriteToTextFile "TotalDiscountExVAT" & vbTab & TotalDiscountExVATF
41400     oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalExtF(bForeign)
41410     oTF.WriteToTextFile "TotalExVAT" & vbTab & TotalNonVATF(bForeign)
41420     oTF.WriteToTextFile "TotalVAT" & vbTab & TotalVATF(bForeign)
41430     oTF.CloseTextFile
41440     ExportToSpreadsheet = True
          
      'errHandler:
      '    If ErrMustStop Then Debug.Assert False: Resume
      '    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
41450     Exit Function
errHandler:
41460     If ErrMustStop Then Debug.Assert False: Resume
41470     ErrorIn "a_Invoice.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
End Function



Public Sub docWriteTostream(ByVal FilePath As String, obj As MSXML2.DOMDocument60, _
                Optional ByVal CharSet As String = "UNICODE")
41480     On Error GoTo errHandler
          Dim s As Object
41490     Set s = CreateObject("ADODB.Stream")
41500     With s
41510         If CharSet <> "" Then .CharSet = CharSet
41520         .Open
41530         .WriteText obj.xml
41540         .SaveToFile FilePath, 2 'adSaveCreateOverWrite
41550         .Close
41560     End With
41570     Exit Sub
errHandler:
41580     ErrorIn "a_Invoice.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub

Public Function ExportToXMLCOOS(rs As ADODB.Recordset, bBySeqNo As Boolean, pFilename As String, _
    Optional ByVal pRoundedUp As Boolean, _
    Optional DispatchMode As enTransmitType, Optional pQtyCopies As Integer, _
    Optional pDestinationEmailAddress As String, Optional pWholeMessage As String) As Boolean
41590     On Error GoTo errHandler
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

41600     If DispatchMode = enMail Then
41610         If oPC.UsesOutlookForINVEmail Then
41620             strWorkingFolder = oPC.LocalFolder & "TEMP\"
41630             strDispatchMethod = "Outlook"
41640         Else
41650             strWorkingFolder = oPC.SharedFolderRoot & "\Emails\"
41660             strDispatchMethod = "Emailing"
41670         End If
41680     ElseIf DispatchMode = enPrint Then
41690         strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
41700         strDispatchMethod = "Printing"
41710     ElseIf DispatchMode = enView Then
41720         strWorkingFolder = oPC.SharedFolderRoot & "\PDF\"
41730         strDispatchMethod = ""
41740     Else
41750         strWorkingFolder = oPC.SharedFolderRoot & "\Printing\"
41760         strDispatchMethod = "Printing"
41770     End If
         '''''' strWorkingFolder = "C:\PBKS\TEMP\"
41780                         p 1
41790     bForeign = False
41800     If Not oCurrency_Foreign Is oPC.Configuration.DefaultCurrency Then
41810         bForeign = True
41820         CalculateTotal_Foreign
41830     End If
41840                         p 2
41850                         p 3
41860     Set oDC = oPC.Configuration.DocumentControls.FindDC(CDOCCODE)
41870     If oDC Is Nothing And strDispatchMethod <> "Emailing" And strDispatchMethod <> "Outlook" Then
41880         ExportToXMLCOOS = False
41890         Exit Function
41900     End If
41910                         p 4
41920     If pQtyCopies < 1 Then pQtyCopies = 1
41930     Set xMLDoc = New ujXML
41940     With xMLDoc
41950         .docProgID = "msxml2.DOMDocument60"
41960         .docInit "CO_DOC"
41970             .chCreate "MessageType"
41980                 .elText = "INVOICECOOS"
41990             .elCreateSibling "MessageCreationDate"
42000                 .elText = Format(Now(), "yyyymmddHHNN")
42010             .elCreateSibling "DestinationAddress"
42020                 If oPC.TestMode Then
42030                     .elText = oPC.EmailFrom
42040                 Else
42050                     If Customer.BillTOAddress Is Nothing Then
42060                         .elText = ""
42070                     Else
42080                         .elText = Customer.BillTOAddress.EMail
42090                     End If
42100                 End If
42110             .elCreateSibling "TemplateName"
42120                     .elText = "INCOOS_DOC"
42130             strTemplateName = .elText
                  
42140                         p 5
42150             .elCreateSibling "RecipientsEmail"
42160                 If Not Me.Customer.BillTOAddress Is Nothing Then
42170                    .elText = Me.Customer.BillTOAddress.EMail
42180                     pDestinationEmailAddress = Me.Customer.BillTOAddress.EMail
42190                 End If
42200             .elCreateSibling "SendersEmail"
42210                 .elText = Me.SendersEmail
42220             .elCreateSibling "CopyCount"
42230                 .elText = pQtyCopies
42240             .elCreateSibling "Printer"
42250                 If Not oDC Is Nothing Then .elText = oDC.PrinterName
                      '.elText = oDC.PrinterName
42260             .elCreateSibling "LogoPath"
42270                 If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
42280                     .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
42290                 Else
42300                     .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
42310                 End If
42320             .elCreateSibling "Status"
42330                 .elText = StatusForPrinting
42340             .elCreateSibling "AccompanyingMessage"
42350                 .elText = oPC.Configuration.EmailInvMsg
42360             .elCreateSibling "IsProforma"
42370                 .elText = IIf(Proforma, "Y", "N")
42380             .elCreateSibling "AccompanyingMessage"
42390                 .elText = oPC.Configuration.EmailInvMsg
42400             .elCreateSibling "DocCode"
42410                 .elText = DOCCode
42420            .elCreateSibling "StoreCode", True
42430                 .elText = oPC.Configuration.DefaultStore.code
42440            .elCreateSibling "DocDate", True
42450                 .elText = DocDateF
42460             .elCreateSibling "Sender", True
42470                 .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName                  'oPC.Configuration.DefaultCompany.CompanyName
42480                                 p 6
42490             .elCreateSibling "SenderAddress", True
42500                 .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
42510             .elCreateSibling "SenderAddressPostal", True
42520                 .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).PostalAddress, Chr(13) & Chr(10), Chr(10))
42530             .elCreateSibling "CustomerName", True
42540                 .elText = Customer.Fullname
42550             .elCreateSibling "CustomerVATNumber", True
42560             If Customer.VatNumber > "" Then
42570                 .elText = "Customer's VAT number: " & Customer.VatNumber
42580             Else
42590                 .elText = ""
42600             End If
42610             .elCreateSibling "OurAcnowithCustomer", True
42620                 .elText = Customer.OurACnoWithClient
42630             .elCreateSibling "ACNO", True
42640                 .elText = Customer.AcNo
42650             .elCreateSibling "BankingDetails", True
42660                 .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(COMPID).BankDetails, Chr(13) & Chr(10), Chr(10))
42670             .elCreateSibling "CustomerWithAddress", True
42680                 If Me.BillTOAddress Is Nothing Then
42690                     .elText = ""
42700                 Else
42710                     .elText = Replace(Me.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
42720                 End If
42730                                 p 7
42740             .elCreateSibling "CustomerPhone", True
42750                 If Me.BillTOAddress Is Nothing Then
42760                     .elText = ""
42770                 Else
42780                     .elText = IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "")
42790                 End If
42800             .elCreateSibling "CustomerFax", True
42810                 If Me.BillTOAddress Is Nothing Then
42820                     .elText = ""
42830                 Else
42840                     .elText = IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
42850                 End If
42860             .elCreateSibling "BillTo", True
42870                 If Me.BillTOAddress Is Nothing Then
42880                     .elText = ""
42890                 Else
42900                     .elText = Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
42910                     If Me.ForAttn > "" Then
42920                         .elText = Me.ForAttn & vbCrLf & .elText
42930                     End If
42940                 End If
42950             .elCreateSibling "DelTo", True
42960                 If Me.DelToAddress Is Nothing Then
42970                     .elText = ""
42980                 Else
42990                     .elText = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
43000                 End If
43010             .elCreateSibling "DispatchMode", True
43020                 If Me.DelToAddress Is Nothing Then
43030                     .elText = ""
43040                 Else
43050                     .elText = FNS(Me.DelToAddress.Description)
43060                 End If
43070                             p 8
43080             bMultipleCOLReps = False
43090             flgFirst = True
43100             lngNumberOfLines = 0
43110             Do While Not rs.eof
43120                     lngNumberOfLines = lngNumberOfLines + 1
43130                     flgFirst = True
43140                     .elCreateSibling "DetailLine", True
43150                     .chCreate "SKU"
43160                         .elText = FNS(rs.Fields("P_EAN"))
43170                     .elCreateSibling "DocumentCode", True
43180                         .elText = FNS(rs.Fields("TRCode"))
43190                     .elCreateSibling "DocumentDate", True
43200                         .elText = Format(rs.Fields("TRDate"), "dd/mm/yyyy")
43210                     .elCreateSibling "Title", True
43220                         .elText = FNS(rs.Fields("P_Title"))
43230                     .elCreateSibling "Author", True
43240                         .elText = FNS(rs.Fields("P_MainAuthor"))
43250                     .elCreateSibling "Publisher", True
43260                         .elText = FNS(rs.Fields("P_Publisher"))
43270                     .elCreateSibling "Qty", True
43280                         .elText = CStr(FNN(rs.Fields("COL_Qty")))
43290                     .elCreateSibling "QtyDispatched", True
43300                         .elText = CStr(FNN(rs.Fields("COL_QtyDispatched")))
43310                     .elCreateSibling "ActionTaken", True
                            '  .elText = CStr(FNS(rs.Fields("COL_ActionTaken")))
43320                     .elCreateSibling "ActionDate", True
                           '   .elText = Format(FND(rs.Fields("COL_LastActionDate")), "dd/mm/yyyy")
43330                     rs.MoveNext
43340                     .navUP
43350                     flgFirst = False
skip:

43360             Loop
43370                                 p 9
43380                                 p 12

      'FINALLY PRODUCE THE .XML FILE
43390     strXML = strWorkingFolder & "IN_" & Me.DOCCode & ".xml"
43400     With xMLDoc
43410         If fs.FileExists(strXML) Then
43420             fs.DeleteFile strXML
43430         End If
             ' strXML = "C:\pbks\txt.xml"
43440         .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
43450     End With

43460                                 p 13

43470     If DispatchMode = enMail Then
43480         If oPC.EmailINVShowHTML = True Then
                  'WRITE THE .HTML FILE
43490             objXSL.async = False
43500             objXSL.ValidateOnParse = False
43510             objXSL.resolveExternals = False
43520             strPath = oPC.SharedFolderRoot & "\Templates\" & fs.GetBaseName(strTemplateName) & "_HTML.xslt"
43530             Set fs = New FileSystemObject
43540             If fs.FileExists(strPath) Then
43550                 objXSL.Load strPath
43560             End If
43570             If fs.FileExists(strWorkingFolder & "IN_" & Me.DOCCode & ".HTML") Then
43580                 fs.DeleteFile strWorkingFolder & "IN_" & Me.DOCCode & ".HTML", True
43590             End If
43600             oTF.OpenTextFileToAppend strWorkingFolder & "IN_" & Me.DOCCode & ".HTML"
43610             pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
43620             oTF.WriteToTextFile pWholeMessage
43630             oTF.CloseTextFile
43640             Set oTF = Nothing
43650         Else
43660         End If
              
43670     End If
          
          
      'WRITE THE .PDF FILE IF NECESSARY
      'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file

43680         strFOFile = strWorkingFolder & "INCOOS_" & Me.DOCCode & ".FO"
43690         strPDFFile = strWorkingFolder & "INCOOS_" & Me.DOCCode & ".PDF"

43700     If (Me.Customer.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Then
43710         If oPC.GetProperty("UseXalan") <> "TRUE" Then
                  '====================================STANDARD METHOD -tenporarily replaces by XALAN method===========================
43720             Set objXSL = Nothing
43730             Set objXSL = New MSXML2.DOMDocument60
43740             objXSL.async = False
43750             objXSL.ValidateOnParse = False
43760             objXSL.resolveExternals = False
43770             strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
                '''''  strPath = "c:\PBKS\Templates\" & strTemplateName & "_FO.XSL"
43780             Set fs = New FileSystemObject
43790             If fs.FileExists(strPath) Then
43800                 objXSL.Load strPath
43810             End If
43820                             p 14
43830             Set opXMLDOC = New MSXML2.DOMDocument60
43840             opXMLDOC.async = False
43850             opXMLDOC.ValidateOnParse = False
43860             opXMLDOC.resolveExternals = False
43870             xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
          
43880                             p 15
43890             docWriteTostream strFOFile, opXMLDOC, "UNICODE"
                  '====================================================================================================================
43900         Else
43910              strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
43920              strCommand = oPC.SharedFolderRoot & "\Executables\PBKS_XALAN.BAT" & " " & strXML & " " & strPath & " " & strFOFile & " " & oPC.SharedFolderRoot
43930              F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
43940         End If
      '====================================================================================================================
      'Stage 2 Convert the .FO file to .PDF and clean up
43950                         p 155
43960         ChDir "\PBKS\Executables\FOP\"
           'Allow 2 minutes to produce the PDF file
43970         strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
43980         F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000
              'IF IT FAILS HERE CHECK THAT JAVA IS INSTALLED
              'We do the following because the batch file executes and is not waited for by the F_7_AB_1_ShellAndWaitSimple routine
              'so the FO file is sometimes deleted before the .PDF file is produced.
43990         cnt = 0
44000                         p 156
44010         Do While Not fs.FileExists(strPDFFile) And cnt < 40
44020             MsgWaitObj 1000
44030             cnt = cnt + 1
44040         Loop
44050                         p 157
44060         If fs.FileExists(strFOFile) Then
44070             fs.DeleteFile strFOFile
44080         End If
44090                         p 158
44100         If fs.FileExists(strPDFFile) Then
44110             SetFileProperty strPDFFile, "Title", Me.Customer.AcNo
44120             SetFileProperty strPDFFile, "Author", Me.Customer.Name

44130         End If
44140     End If
44150                         p 16
            
      'Log the dispatch
44160     Set oSM = New z_StockManager
44170     Set oFSO = New FileSystemObject
44180     pFilename = strPDFFile
          
44190     If DispatchMode <> enMail And DispatchMode <> enView Then
44200         Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
44210     End If
44220     Set oSM = Nothing
44230     Set oFSO = Nothing
44240                         p 9
44250     oTF.CloseTextFile
44260     Set oTF = Nothing
44270     ExportToXMLCOOS = True
          '
44280     If (Not fs.FileExists(strPDFFile)) And (DispatchMode = enView Or DispatchMode = enMail) Then
44290             MsgBox "The PDF file has not been produced: " & strPDFFile & ". Please contact support.", vbOKOnly, "Can't do this"
44300     Else
44310         If DispatchMode = enView Then
44320             OpenFileWithApplication strPDFFile, enPDF
44330         End If
44340     End If
44350     End With
44360     Exit Function
errHandler:
44370     If ErrMustStop Then Debug.Assert False: Resume
44380     ErrorIn "a_Invoice.ExportToXMLCOOS(rs,bBySeqNo,pFilename,pRoundedUp,DispatchMode,pQtyCopies," & _
              "pDestinationEmailAddress,pWholeMessage)", Array(rs, bBySeqNo, pFilename, pRoundedUp, DispatchMode, _
               pQtyCopies, pDestinationEmailAddress, pWholeMessage), EA_NORERAISE
End Function



























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
44390     On Error GoTo errHandler
44400     Log = Trim(mudtProps.Log)
44410     Exit Property
errHandler:
44420     ErrorIn "a_Invoice.Log"
End Property

Public Property Let Log(val As String)
44430     On Error GoTo errHandler
44440     mudtProps.Log = Trim(val)
44450     Exit Property
errHandler:
44460     ErrorIn "a_Invoice.Log(val)", val
End Property
Public Property Get Exchange() As String
44470     On Error GoTo errHandler
44480     Exchange = Trim(mudtProps.Exchange)
44490     Exit Property
errHandler:
44500     ErrorIn "a_Invoice.Exchange"
End Property

Public Property Get TotalQty() As Long
44510     On Error GoTo errHandler

44520         TotalQty = mudtProps.TotalQty

44530     Exit Property
errHandler:
44540     If ErrMustStop Then Debug.Assert False: Resume
44550     ErrorIn "a_Invoice.TotalQty"
End Property
Public Function TotalQtyPerProduct(PID As String) As Long
      Dim oIL As a_InvoiceLine
      Dim lngTot As Long

44560     lngTot = 0
44570     For Each oIL In Me.InvoiceLines
44580         If oIL.PID = PID Then
44590             lngTot = lngTot + oIL.Qty
44600         End If
44610     Next
44620     TotalQtyPerProduct = lngTot
End Function

Public Property Get TotalQtyF() As String
44630     On Error GoTo errHandler

44640         TotalQtyF = CStr(TotalQty)

44650     Exit Property
errHandler:
44660     If ErrMustStop Then Debug.Assert False: Resume
44670     ErrorIn "a_Invoice.TotalQtyF"
End Property

Private Sub addARow(pDescription, Amt As Long, AmtF As String, pTotalType As String, Optional Sign As String)
45280     On Error GoTo errHandler
45290     arDocTotals(jTotal).Description = pDescription
45300     arDocTotals(jTotal).val = Amt
45310     arDocTotals(jTotal).AmtFormatted = AmtF & IIf(Left(pDescription, 1) = "(", ")", "")
45320     arDocTotals(jTotal).TotalType = pTotalType
45330     If Sign = "-" Then
45340         RunningTotal = RunningTotal - Amt
45350     ElseIf Sign = "+" Then
45360         RunningTotal = RunningTotal + Amt
45370     End If
45380     arDocTotals(jTotal).RunningTotal = RunningTotal
45390     jTotal = jTotal + 1
45400     Exit Sub
errHandler:
45410     If ErrMustStop Then Debug.Assert False: Resume
45420     ErrorIn "a_Invoice.addARow(pDescription,Amt,AmtF,pTotalType)", Array(pDescription, Amt, _
               AmtF, pTotalType)
End Sub

Public Sub PasteLine(PID As String, QtyFirm As Long, QtySS As Long, Price As Long, DiscountRate As Double, _
                    VATRate As Double, Ref As String, ExtraChargePID As String, ExtraChargeValue As Long, _
                    FCPrice As Long, FCFactor As Double, FCID As Long)
      Dim oSM As New z_StockManager

45430     oSM.PasteInvoiceLine PID, QtyFirm, QtySS, Price, DiscountRate, VATRate, mudtProps.InvoiceID, Ref, ExtraChargePID, ExtraChargeValue, FCPrice, FCFactor, FCID
          
End Sub


Public Function GenerateEDIMsg() As Boolean
45440     On Error GoTo errHandler
      Dim str As String
      Dim strMsg As String
      Dim dte As Date
      Dim oTF As New z_TextFile
      Dim strFilePath As String
      Dim strFilename As String
      Dim i As Integer
      Dim oSM As z_StockManager
      Dim oFSO As New FileSystemObject

45450     dte = Now()
45460     strFilename = Me.DOCCode '& Format(dte, "HHNN")
          
45470     strFilePath = oPC.SharedFolderRoot & "\EDI_Out\Invoices\" & oFSO.GetBaseName(strFilename) & ".EDI"
          
45480     oTF.OpenTextFile strFilePath
45490     strMsg = "UNA:+.? '"
45500     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
          
45510     strMsg = "UNB+UNOA:2+" '& oPC.Configuration.GFXNumber & "+" & Me.Supplier.GFXNumber & "+" & Format(dte, "YYYYMMDD") & ":" & Format(dte, "HHNN") & "+" & strFileName & "+PASSWORD+ORDERS"
45520     oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & vbCrLf
45530     strMsg = "UNH++INVOIC"
45540     oTF.WriteToTextFile_NoLineTerminator strMsg & "'" & vbCrLf
45550     strMsg = "BGM+380:::+" & Me.DOCCode & "'"
45560     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45570     strMsg = "DTM+137:" & Format(Me.DOCDate, "YYYYMMDD") & ":102'"
45580     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
          
45590     For i = 1 To Me.InvoiceLines.Count
45600         strMsg = "LIN+" & Format(i, "000000") & "'"
45610         oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45620         If InvoiceLines(i).FCID > 0 Then
45630             strMsg = "IMD+L+050+:::" & Replace(Left(Me.InvoiceLines(i).Title, 35), "'", "?'") & ":" & InvoiceLines(i).FCFactorInvF & "/" & InvoiceLines(i).ForeignPriceF & "'"
45640         Else
45650             strMsg = "IMD+L+050+:::" & Replace(Left(Me.InvoiceLines(i).Title, 35), "'", "?'") & "'"
45660         End If
45670         oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45680         strMsg = "QTY+47:" & CStr(Me.InvoiceLines(i).Qty) & "'"
45690         oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45700         strMsg = "MOA+203:" & CStr(Me.InvoiceLines(i).PAfterDiscountExtNonVAT_withDecimal(False)) & "'"
45710         oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45720         strMsg = "RFF+LI:" & ClearOrderRef(CStr(Me.InvoiceLines(i).Ref)) & "'"   ' & "/" & CStr(Me.InvoiceLines(i).ForeignPriceF)
45730         oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
              
45740     Next
45750     strMsg = "UNS+S'"
45760     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45770     strMsg = "CNT+2:" & CStr(Me.InvoiceLines.Count) & "'"
45780     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45790     strMsg = "MOA+86:" & CStr(Me.TotalPayable_WithDecimal(False)) & "'"
45800     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45810     strMsg = "ALC+C++++TX'"
45820     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
45830     strMsg = "MOA+8:" & CStr(Me.TotalVAT_WithDecimal(False)) & "'"
45840     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
      '
45850     strMsg = "UNT+56+00000000001" & "'"
45860     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf

45870     strMsg = "UNZ+1+" & strFilename & "'"
45880     oTF.WriteToTextFile_NoLineTerminator strMsg & vbCrLf
          
      '    Set oSM = New z_StockManager
      '    oSM.LogTransmission Me.TRID, "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf
      '    Log = "File: " & oFSO.GetFileName(oTF.FileName) & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
      '    Set oSM = Nothing
          
45890     oTF.CloseTextFile
45900     Exit Function
errHandler:
45910     If ErrMustStop Then Debug.Assert False: Resume
45920     ErrorIn "a_Invoice.GenerateEDIMsg"
End Function
Private Function ClearOrderRef(s As String) As String
45930     If InStr(1, s, ",") > 0 Then
45940         s = Left(s, InStr(1, s, ",") - 1)
45950     End If
45960     ClearOrderRef = s
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




