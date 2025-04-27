VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_InvoiceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------------------------------------
'HANDLING OUTPUT IN FOREIGN CURRENCY and capture and local display in local currency formats
'Definitions:
    'Localprice: a number representing the price in the default currency as determined in the configuration table (stored in the smallest denominated currency unit e.g. cent)
    'ForeignPrice: a number representing the price in the selected foreign currency as determined in the Invoice table (stored in the smallest denominated currency unit e.g. cent)
    'Factor: a number that when applied to the localprice yields the foreignprice (both stored in the smallest denominated currency unit e.g. cent)
    'Divisor: a number that when divided into the localprice or the foreignprice yields a number representing the price in the larger currency unit (e.g. Pounds, dollars, rands)
    'FormatString: a text string used with the Format function to represent a localprice or a foreignprice (After dividing by Divisor) as a properly formatted currency
'Each IL keeps the price and deposit (the only two actual prices, the other are percentages) in two forms,
'one as local currency (Long representing smallest denominated currency unit e.g. cent)
'and the other as foreign currency in the same representation
'The Invoice object records a currency to use for output (it may be just the local currency) and the factor at time of issue
'When the invoice is to be shown as it would appear printed, the calcTotalsOP method of the invoice object is called,
'being passed the foreign currency factor to use. All the properties showing prices anc totals formatted now use the
'foreign currency Divisor and factor and formatstring
'------------------------------------------------------------------------------------------------------------------------

Event Valid(msg As String)
'Event ExtensionChange(lngExtension As Long, strExtension As String)
Private mudtProps As ILProps
Private mcolStack As Collection
Private mobjParent As Object
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Private colClassors As Collection
Private mInvoiceParent As Object
Private WithEvents mchCOFFs As ch_COFF
Attribute mchCOFFs.VB_VarHelpID = -1

Private mKey As String
Private moProd As a_Product

Private mFPAfterDiscount As Double
Private mLPAfterDiscount As Double

Private mFPAfterDiscountExt As Double
Private mLPAfterDiscountLineValue As Double
Private mLPAfterDiscountLineValueExVAT As Double
Private mLPExt As Double
Private mFPExt As Double
Private mFPDiscExt As Double
Private mLPLineDiscountValue As Double
Private mLPDiscExVATExt As Double
Private mUnitPriceInc As Double
Private mFUnitPriceInc As Double

Private mFPAfterDiscountExtNonVAT As Double
Private mLPAfterDiscountLineValueNonVAT As Double

Private mFPAfterDiscountLineValueVAT As Double
Private mLPAfterDiscountLineValueVAT As Double

Private mFPDisc_Unit As Double
Private mLPDisc_Unit As Double

Private mFPDisc_Ext As Double
Private mLPDisc_Ext As Double
Private mExVATUnitPrice As Double
Private mSelected As Boolean

Dim LTG As Double  'LineTotal Gross of discount
Dim LTN As Double  'LineTotal Nett Of discount
Dim LTNN As Double 'LineTotal Nett of discount, Nett of VAT
Dim UPN As Double  'Unit Price Nett of discount rounded to 4 decimal places
Dim LD As Double  'LineDiscount
Dim LDN As Double  'LineDiscount Nett of VAT
Dim LV As Double  'LineVAT

Public Property Get row_LTG() As Double
    row_LTG = LTG
End Property
Public Property Get row_LTN() As Double
    row_LTN = LTN
End Property
Public Property Get row_LTNF(bForeign As Boolean, Optional BlankWhenZero As Boolean = False) As String
    On Error GoTo errHandler
    
        If LTN = 0 And BlankWhenZero Then
            row_LTNF = ""
        Else
            If bForeign Then
                row_LTNF = Format(CDbl(LTN / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
            Else
                row_LTNF = Format(CDbl(LTN / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
            End If
        End If
        
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.row_LTNF"
End Property
Public Property Get row_LTNN() As Double
    row_LTNN = LTNN
End Property
Public Property Get row_LTNNF(Optional BlankWhenZero As Boolean = False) As String
    On Error GoTo errHandler
        If LTNN = 0 And BlankWhenZero Then
            row_LTNNF = ""
        Else
            row_LTNNF = Format(CDbl(LTNN / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
        End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.row_LTNNF"
End Property
Public Property Get row_UPN() As Double
    row_UPN = UPN
End Property
Public Property Get row_LD() As Double
    row_LD = LD
End Property
Public Property Get row_LDN() As Double
    row_LDN = LDN
End Property
Public Property Get row_LV() As Double
    row_LV = LV
End Property
Public Property Get row_LVF(Optional BlankWhenZero As Boolean = False) As String
    On Error GoTo errHandler
        If LV = 0 And BlankWhenZero Then
            row_LVF = ""
        Else
            row_LVF = Format(CDbl(LV / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
        End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.row_LTNNF"
End Property
Public Property Let Selected(val As Boolean)
47900     mSelected = val
End Property
Public Property Get Selected() As Boolean
47910     Selected = mSelected
End Property

Private Sub mchCoffs_CoffQtyChanged()
47920     On Error GoTo errHandler
47930     MsgBox "Here"
47940     Exit Sub
errHandler:
47950     If ErrMustStop Then Debug.Assert False: Resume
47960     ErrorIn "a_InvoiceLine.mchCoffs_CoffQtyChanged"
End Sub
Public Property Get COFFs() As ch_COFF
47970     Set COFFs = mchCOFFs
End Property

Public Property Get Properties(pSrt As enSortField) As Variant
47980     On Error GoTo errHandler
47990     Select Case pSrt
          Case enSequence
48000         Properties = Me.Sequence
48010     Case enTitle
48020         Properties = Me.Title
48030     Case enAuthor
48040         Properties = Me.MainAuthor
48050     Case enCode
48060         Properties = Me.code
48070     Case enRef
48080         Properties = Me.Ref
48090     End Select
48100     Exit Property
errHandler:
48110     If ErrMustStop Then Debug.Assert False: Resume
48120     ErrorIn "a_InvoiceLine.Properties(pSrt)", pSrt
End Property

Public Property Get Sequence() As Long
48130     On Error GoTo errHandler
48140     Sequence = mudtProps.Sequence
48150     Exit Property
errHandler:
48160     If ErrMustStop Then Debug.Assert False: Resume
48170     ErrorIn "a_InvoiceLine.Sequence"
End Property
Public Property Let Sequence(val As Long)
48180     On Error GoTo errHandler
48190     If val <> mudtProps.Sequence Then
48200         mudtProps.Sequence = val
48210         SetDirty True
48220     End If
48230     Exit Property
errHandler:
48240     If ErrMustStop Then Debug.Assert False: Resume
48250     ErrorIn "a_InvoiceLine.Sequence(val)", val
End Property

Private Function SetForeignPriceAndDeposit()
48260     On Error GoTo errHandler
      Dim tmpPrice As Long
      Dim lngRoundTo As Long
      Dim ConvertedPrice As Long
      Dim oCurr As a_Currency
48270     mobjValid.RuleBroken "OVERFLOW", False
48280     tmpPrice = mudtProps.Price * mInvoiceParent.CurrencyFactor
48290     If Me.ServiceItem Then
48300         mudtProps.ForeignPrice = tmpPrice
48310         mudtProps.ForeignDeposit = mudtProps.Deposit * mInvoiceParent.CurrencyFactor
48320     Else
48330         lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
48340         ConvertedPrice = (RoundUp(tmpPrice, lngRoundTo))
48350         mudtProps.ForeignPrice = ConvertedPrice
48360         mudtProps.ForeignDeposit = mudtProps.Deposit / mInvoiceParent.CurrencyFactor
48370     End If
48380     Exit Function
errHandler:
48390     ErrPreserve
48400     If Err = 6 Then 'ov erflow
48410         mobjValid.RuleBroken "OVERFLOW", True
48420         Clear
48430         Resume Next
48440     End If
48450     If ErrMustStop Then Debug.Assert False: Resume
48460     ErrorIn "a_InvoiceLine.SetForeignPriceAndDeposit"
End Function
Public Function CalculateLine_X()
48470     On Error GoTo errHandler
      Dim xVATPrice As Double
      Dim bDone As Boolean
      'First in local currency
48480     mobjValid.RuleBroken "OVERFLOW", False
48490     If Not mudtProps.ServiceItem Then
48500         If oPC.AllowsSSInvoicing Then
48510             mudtProps.Qty = mudtProps.QtyFirm + mudtProps.QtySS
48520         End If
48530     End If
48540     bDone = False
48550     If Not mInvoiceParent Is Nothing Then
48560         If mInvoiceParent.Customer.ShowVAT = False And mInvoiceParent.Customer.VATable = False Then  'We base all calculations on an ex-VAT price
                'this is for where we essentially are creating an invoice that does not need to show VAT as it is going to a foreign customer
                'we however are charging the list price which shows VAT and corresponds to our published list price. ie the customer pays an inflated price as the seller does not collect and submit VAT on the sale
48570             mExVATUnitPrice = mudtProps.Price
48580             mLPAfterDiscountLineValue = Round(CDbl(mExVATUnitPrice * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountPercent) / 100)
48590             mLPLineDiscountValue = (CDbl(mudtProps.Price * mudtProps.Qty)) - mLPAfterDiscountLineValue
48600             mLPAfterDiscountLineValueVAT = 0
48610             mLPDiscExVATExt = mLPLineDiscountValue
48620             mLPAfterDiscountLineValueNonVAT = mLPAfterDiscountLineValue - mLPAfterDiscountLineValueVAT
48630             mLPExt = mudtProps.Price * mudtProps.Qty
48640             bDone = True
48650         End If
48660     End If
48670     If bDone = False Then
          'new version
Dim LT As Double  'LineTotal
Dim LV As Double  'LineVAT

48680         mLPExt = mudtProps.Price * mudtProps.Qty
48690         mLPAfterDiscountLineValue = Fix(Round((CDbl(mudtProps.Price * mudtProps.Qty) * (CDbl((100 - (mudtProps.DiscountPercent))) / 100)), 0))
48700         mUnitPriceInc = Round(mLPAfterDiscountLineValue / mudtProps.Qty, 4)
48710         mLPLineDiscountValue = mLPExt - mLPAfterDiscountLineValue
48720         mLPAfterDiscountLineValueVAT = Round(mLPAfterDiscountLineValue - (CDbl(mLPAfterDiscountLineValue) * 100 / (CDbl(100) + mudtProps.VATRate)), 2)
48730         mLPDiscExVATExt = mLPLineDiscountValue * 100# / (CDbl(100) + mudtProps.VATRate)
48740         mLPAfterDiscountLineValueNonVAT = mLPAfterDiscountLineValue - mLPAfterDiscountLineValueVAT
48750     End If
      'Then in Foreign if appropriate
48760     If mInvoiceParent Is Nothing Then Exit Function
48770     If Not mInvoiceParent.ForeignCurrency Is Nothing Then
48780         If mInvoiceParent.ForeignCurrency.ID <> oPC.Configuration.DefaultCurrency.ID Then
48790             SetForeignPriceAndDeposit
48800             mFPExt = mudtProps.ForeignPrice * mudtProps.Qty
48810             If UCase(oPC.GetProperty("CalculatePxQxD")) = "TRUE" Then
                  'updated version
48820                 mFUnitPriceInc = (CDbl(mudtProps.ForeignPrice * mudtProps.Qty) * CDbl((100 - (mudtProps.DiscountPercent))) / 100) / mudtProps.Qty
48830                 mFPAfterDiscountExt = CDbl(mudtProps.ForeignPrice * mudtProps.Qty) * CDbl((100 - (mudtProps.DiscountPercent))) / 100
48840             Else
                  'older version
48850                 mFUnitPriceInc = mudtProps.ForeignPrice * CDbl((100 - mudtProps.DiscountPercent) / 100)
48860                 mFPAfterDiscountExt = Round(CDbl(mFUnitPriceInc) * ((CDbl(100) - mudtProps.DiscountPercent) / 100))
48870             End If
48880             mFPDiscExt = (CDbl(mudtProps.ForeignPrice * mudtProps.Qty)) - mFPAfterDiscountExt
48890             mFPAfterDiscountLineValueVAT = Round(mFPAfterDiscountExt - (CDbl(mFPAfterDiscountExt) * 100 / (CDbl(100) + mudtProps.VATRate)))
48900             mFPAfterDiscountExtNonVAT = mFPAfterDiscountExt - mFPAfterDiscountLineValueVAT
48910             mFPExt = mudtProps.ForeignPrice * mudtProps.Qty
48920         End If
48930     End If
48940     Exit Function
          
errHandler:
48950     ErrPreserve
48960     If Err = 6 Then  'overflow
48970         mobjValid.RuleBroken "OVERFLOW", True
48980         Clear
48990         Resume Next
49000     End If
49010     If ErrMustStop Then Debug.Assert False: Resume
49020     ErrorIn "a_InvoiceLine.CalculateLine"
End Function
Public Sub CalculateLine()
    On Error GoTo errHandler
Dim xVATPrice As Double
Dim bDone As Boolean
'First in local currency
    mobjValid.RuleBroken "OVERFLOW", False
    If Not mudtProps.ServiceItem Then
        If oPC.AllowsSSInvoicing Then
            mudtProps.Qty = mudtProps.QtyFirm + mudtProps.QtySS
        End If
    End If
    bDone = False
    If Not mInvoiceParent Is Nothing Then
        If mInvoiceParent.Customer.ShowVAT = False And mInvoiceParent.Customer.VATable = False Then  'We base all calculations on an ex-VAT price
          'this is for where we essentially are creating an invoice that does not need to show VAT as it is going to a foreign customer
          'we however are charging the list price which shows VAT and corresponds to our published list price. ie the customer pays an inflated price
          'as the seller does not collect and submit VAT on the sale
            mExVATUnitPrice = mudtProps.Price
            mLPAfterDiscountLineValue = Round(CDbl(mExVATUnitPrice * mudtProps.Qty) * (CDbl(100) - mudtProps.DiscountPercent) / 100)
            mLPLineDiscountValue = (CDbl(mudtProps.Price * mudtProps.Qty)) - mLPAfterDiscountLineValue
            mLPAfterDiscountLineValueVAT = 0
            mLPDiscExVATExt = mLPLineDiscountValue
            mLPAfterDiscountLineValueNonVAT = mLPAfterDiscountLineValue - mLPAfterDiscountLineValueVAT
            mLPExt = mudtProps.Price * mudtProps.Qty
            bDone = True
        End If
    End If
    If bDone = False Then
        LTG = mudtProps.Price * mudtProps.Qty
        LTN = Fix(Round((CDbl(mudtProps.Price * mudtProps.Qty) * (CDbl((100 - (mudtProps.DiscountPercent))) / 100)), 0))
        UPN = Round(LTN / mudtProps.Qty, 4)
        LD = LTG - LTN
        LV = Round(LTN - (CDbl(LTN) * 100 / (CDbl(100) + mudtProps.VATRate)))
        LDN = LD * 100# / (CDbl(100) + mudtProps.VATRate)
        LTNN = LTN - LV
        
    End If
'Then in Foreign if appropriate
    If mInvoiceParent Is Nothing Then Exit Sub
    If Not mInvoiceParent.ForeignCurrency Is Nothing Then
        If mInvoiceParent.ForeignCurrency.ID <> oPC.Configuration.DefaultCurrency.ID Then
            SetForeignPriceAndDeposit
            LTG = mudtProps.ForeignPrice * mudtProps.Qty
            
            'Line total Nett of discount
            LTN = Fix(Round((CDbl(mudtProps.ForeignPrice * mudtProps.Qty) * (CDbl((100 - (mudtProps.DiscountPercent))) / 100)), 0))
            
            'Unit Price Nett of discount rounded to 4 decimal places
            UPN = Round(LTN / mudtProps.Qty, 4)
            
            'Line VAT value using line value nett of discount
            LV = Round(LTN - (CDbl(LTN) * 100 / (CDbl(100) + mudtProps.VATRate)))
            
            'Get Line discount from above
            LD = LTG - LTN
            
            'Line Discount Nett of VAT
            LDN = LD * 100# / (CDbl(100) + mudtProps.VATRate)
            
            'Line total Nett of discount, Nett of VAT
            LTNN = LTN - LV
            
        End If
    End If
    Exit Sub
    
errHandler:
    ErrPreserve
    If Err = 6 Then  'overflow
        mobjValid.RuleBroken "OVERFLOW", True
        Clear
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_InvoiceLine.CalculateLine"
End Sub
Public Property Get PAfterDiscount(val As Boolean) As Double
49030     On Error GoTo errHandler
49040     If val Then
49050         PAfterDiscount = mFPAfterDiscount
49060     Else
49070         PAfterDiscount = mLPAfterDiscount
49080     End If
49090     Exit Property
errHandler:
49100     If ErrMustStop Then Debug.Assert False: Resume
49110     ErrorIn "a_InvoiceLine.PAfterDiscount(val)", val
End Property
Public Property Get PAfterDiscountF(val As Boolean) As String
49120     On Error GoTo errHandler
49130     If val Then
49140         PAfterDiscountF = Format(CDbl(mFPAfterDiscount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
49150     Else
49160         PAfterDiscountF = Format(CDbl(mLPAfterDiscount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
49170     End If
49180     Exit Property
errHandler:
49190     If ErrMustStop Then Debug.Assert False: Resume
49200     ErrorIn "a_InvoiceLine.PAfterDiscountF(val)", val
End Property

Public Property Get PAfterDiscountExt(val As Boolean) As Double
49210     On Error GoTo errHandler
49220     If val Then
49230         PAfterDiscountExt = mFPAfterDiscountExt
49240     Else
49250         PAfterDiscountExt = mLPAfterDiscountLineValue
49260     End If
49270     Exit Property
errHandler:
49280     If ErrMustStop Then Debug.Assert False: Resume
49290     ErrorIn "a_InvoiceLine.PAfterDiscountExt(val)", val
End Property
Public Property Get UnitPriceIncF(val As Boolean) As String
49300     On Error GoTo errHandler
49310     If val Then
49320         UnitPriceIncF = Format(CDbl(mFUnitPriceInc / mInvoiceParent.ForeignCurrency.Divisor), Replace(mInvoiceParent.ForeignCurrency.FormatString, ".00", ".00##"))
49330     Else
49340         UnitPriceIncF = Format(CDbl(mUnitPriceInc / oPC.Configuration.DefaultCurrency.Divisor), Replace(oPC.Configuration.DefaultCurrency.FormatString, ".00", ".00##"))
49350     End If
49360     Exit Property
errHandler:
49370     If ErrMustStop Then Debug.Assert False: Resume
49380     ErrorIn "a_InvoiceLine.UnitPriceInc(val)", val
End Property
Public Property Get PAfterDiscountExtF(val As Boolean) As String
49390     On Error GoTo errHandler
49400     If val Then
49410         PAfterDiscountExtF = Format(CDbl(mFPAfterDiscountExt / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
49420     Else
49430         PAfterDiscountExtF = Format(CDbl(LTNN / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
49440     End If
49450     Exit Property
errHandler:
49460     If ErrMustStop Then Debug.Assert False: Resume
49470     ErrorIn "a_InvoiceLine.PAfterDiscountExtF(val)", val
End Property
'mLPDiscExVATExt
Public Property Get PDiscExtExVAT() As Double
49480     On Error GoTo errHandler
49490         PDiscExtExVAT = mLPDiscExVATExt
49500     Exit Property
errHandler:
49510     If ErrMustStop Then Debug.Assert False: Resume
49520     ErrorIn "a_InvoiceLine.PDiscExtExVAT"
End Property
Public Property Get PDiscExtExVATF(Optional BlankWhenZero As Boolean = False) As String
49530     On Error GoTo errHandler
49540         If mLPDiscExVATExt = 0 And BlankWhenZero Then
49550             PDiscExtExVATF = ""
49560         Else
49570             PDiscExtExVATF = Format(CDbl(mLPDiscExVATExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
49580         End If
49590     Exit Property
errHandler:
49600     If ErrMustStop Then Debug.Assert False: Resume
49610     ErrorIn "a_InvoiceLine.PDiscExtF"
End Property

Public Property Get PDiscExt(val As Boolean) As Double
49620     On Error GoTo errHandler
49630     If val Then
49640         PDiscExt = mFPDiscExt
49650     Else
49660         PDiscExt = mLPLineDiscountValue
49670     End If
49680     Exit Property
errHandler:
49690     If ErrMustStop Then Debug.Assert False: Resume
49700     ErrorIn "a_InvoiceLine.PDiscExt(val)", val
End Property
Public Property Get PDiscExtF(val As Boolean) As String
49710     On Error GoTo errHandler
49720     If val Then
49730         PDiscExtF = Format(CDbl(mFPDiscExt / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
49740     Else
49750         PDiscExtF = Format(CDbl(mLPLineDiscountValue / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
49760     End If
49770     Exit Property
errHandler:
49780     If ErrMustStop Then Debug.Assert False: Resume
49790     ErrorIn "a_InvoiceLine.PDiscExtF(val)", val
End Property
Public Property Get PExtF(val As Boolean) As String
49800     On Error GoTo errHandler
49810     If val Then
49820         PExtF = Format(CDbl(mFPExt / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
49830     Else
49840         PExtF = Format(CDbl(mLPExt / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
49850     End If
49860     Exit Property
errHandler:
49870     If ErrMustStop Then Debug.Assert False: Resume
49880     ErrorIn "a_InvoiceLine.PExtF(val)", val
End Property
Public Property Get PExt(val As Boolean) As Double
49890     On Error GoTo errHandler
49900     If val Then
49910         PExt = mFPExt
49920     Else
49930         PExt = mFPExt
49940     End If
49950     Exit Property
errHandler:
49960     If ErrMustStop Then Debug.Assert False: Resume
49970     ErrorIn "a_InvoiceLine.PExt(val)", val
End Property

Public Property Get PAfterDiscountLineValueVAT(val As Boolean) As Double
49980     On Error GoTo errHandler
49990     If val Then
50000         PAfterDiscountLineValueVAT = mFPAfterDiscountLineValueVAT
50010     Else
50020         PAfterDiscountLineValueVAT = mLPAfterDiscountLineValueVAT
50030     End If
50040     Exit Property
errHandler:
50050     If ErrMustStop Then Debug.Assert False: Resume
50060     ErrorIn "a_InvoiceLine.PAfterDiscountLineValueVAT(val)", val
End Property
Public Property Get PAfterDiscountLineValueVATF(val As Boolean) As String
50070     On Error GoTo errHandler
50080     If val Then
50090         PAfterDiscountLineValueVATF = Format(CDbl(mFPAfterDiscountLineValueVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
50100     Else
50110         PAfterDiscountLineValueVATF = Format(CDbl(LV / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
50120     End If
50130     Exit Property
errHandler:
50140     If ErrMustStop Then Debug.Assert False: Resume
50150     ErrorIn "a_InvoiceLine.PAfterDiscountLineValueVATF(val)", val
End Property

Public Property Get PAfterDiscountExtNonVAT(val As Boolean) As Double
50160     On Error GoTo errHandler
50170     If VATRate = 0 Then
50180         PAfterDiscountExtNonVAT = mFPAfterDiscountExtNonVAT
50190     Else
50200         PAfterDiscountExtNonVAT = mLPAfterDiscountLineValueNonVAT
50210     End If
50220     Exit Property
errHandler:
50230     If ErrMustStop Then Debug.Assert False: Resume
50240     ErrorIn "a_InvoiceLine.PAfterDiscountExtNonVAT(val)", val
End Property
Public Property Get PAfterDiscountExtNonVATF(val As Boolean) As String
50250     On Error GoTo errHandler
50260     If val Then
50270         PAfterDiscountExtNonVATF = Format(CDbl(mFPAfterDiscountExtNonVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
50280     Else
50290         PAfterDiscountExtNonVATF = Format(CDbl(mLPAfterDiscountLineValueNonVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
50300     End If
50310     Exit Property
errHandler:
50320     If ErrMustStop Then Debug.Assert False: Resume
50330     ErrorIn "a_InvoiceLine.PAfterDiscountExtNonVATF(val)", val
End Property
Public Property Get PAfterDiscountExtNonVAT_withDecimal(val As Boolean) As String
50340     On Error GoTo errHandler
50350     If val Then
50360         PAfterDiscountExtNonVAT_withDecimal = Format(mFPAfterDiscountExtNonVAT / mInvoiceParent.ForeignCurrency.Divisor, "#######0.00")
50370     Else
50380         PAfterDiscountExtNonVAT_withDecimal = Format(mLPAfterDiscountLineValueNonVAT / oPC.Configuration.DefaultCurrency.Divisor, "#######0.00")
50390     End If
50400     Exit Property
errHandler:
50410     If ErrMustStop Then Debug.Assert False: Resume
50420     ErrorIn "a_InvoiceLine.PAfterDiscountExtNonVAT_withDecimal(val)", val
End Property

Private Function GetState() As String
50430     On Error GoTo errHandler
      Dim mudtData As ILData
        
50440     LSet mudtData = mudtProps
50450     GetState = mudtData.Buffer
50460     Exit Function
errHandler:
50470     If ErrMustStop Then Debug.Assert False: Resume
50480     ErrorIn "a_InvoiceLine.GetState"
End Function
    
Private Sub SetState(Buffer As String)
50490     On Error GoTo errHandler
      Dim mudtData As ILData
        
50500     mudtData.Buffer = Buffer
50510     LSet mudtProps = mudtData
50520     Exit Sub
errHandler:
50530     If ErrMustStop Then Debug.Assert False: Resume
50540     ErrorIn "a_InvoiceLine.SetState(buffer)", Buffer
End Sub
Friend Function GetSuperState() As String
50550     On Error GoTo errHandler
      'Dim objPB As PropertyBag
      '
      '  Set objPB = New PropertyBag
      '  With objPB
      '    .WriteProperty "State", GetState
      '    .WriteProperty "COFFs", mchCOFFs.GetSuperState
      '    GetSuperState = .Contents
      '  End With
      '  Set objPB = Nothing
50560     GetSuperState = GetState

50570     Exit Function
errHandler:
50580     If ErrMustStop Then Debug.Assert False: Resume
50590     ErrorIn "a_InvoiceLine.GetSuperState"
End Function
Friend Sub SetSuperState(Buffer As String)
50600     On Error GoTo errHandler
      'Dim objPB As PropertyBag
      'Dim arBuffer() As Byte
        
50610   SetState Buffer
50620     Exit Sub
errHandler:
50630     If ErrMustStop Then Debug.Assert False: Resume
50640     ErrorIn "a_InvoiceLine.SetSuperState(buffer)", Buffer
End Sub

Public Sub ApplyEdit()
50650     On Error GoTo errHandler
50660     If mcolStack.Count = 0 Then Err.Raise 445
          'look for any other COFFs that may be applicable
50670     mchCOFFs.ApplyEdit
50680     mudtProps.SalesValue = Me.PAfterDiscountExtNonVAT(False)
50690     mudtProps.VATAmount = Me.PAfterDiscountLineValueVAT(False)
50700     If Not mobjParent Is Nothing Then
50710         mobjParent.AddDocLine Me
50720         Set mobjParent = Nothing
50730     End If
50740     mcolStack.Remove mcolStack.Count
50750     Exit Sub
errHandler:
50760     If ErrMustStop Then Debug.Assert False: Resume
50770     ErrorIn "a_InvoiceLine.ApplyEdit"
End Sub
Public Property Get Key() As String
50780     Key = mKey
End Property
Public Property Let Key(val As String)
50790     mKey = val
End Property
Public Sub BeginEdit()
50800     On Error GoTo errHandler
50810     mchCOFFs.BeginEdit
50820   mcolStack.Add GetState
50830     Exit Sub
errHandler:
50840     If ErrMustStop Then Debug.Assert False: Resume
50850     ErrorIn "a_InvoiceLine.BeginEdit"
End Sub

Public Sub CancelEdit()
50860     On Error GoTo errHandler
50870   If mcolStack.Count = 0 Then Err.Raise 445
50880     mchCOFFs.CancelEdit
        
50890   mudtProps.IsDeleted = False
50900   With mcolStack
50910     SetState .Item(.Count)
50920     .Remove .Count
50930   End With
50940     Exit Sub
errHandler:
50950     If ErrMustStop Then Debug.Assert False: Resume
50960     ErrorIn "a_InvoiceLine.CancelEdit"
End Sub

Private Sub Class_Initialize()
50970     On Error GoTo errHandler
50980     Set mcolStack = New Collection
50990     Set mobjValid = New z_BrokenRules
51000     Set mchCOFFs = New ch_COFF
51010     LoadClassorsCollection
51020     mobjValid.RuleBroken "Qty", True
51030     mobjValid.RuleBroken "Price", True
51040     mudtProps.IsNew = True
51050     Set moProd = New a_Product
51060     Exit Sub
errHandler:
51070     If ErrMustStop Then Debug.Assert False: Resume
51080     ErrorIn "a_InvoiceLine.Class_Initialize"
End Sub
Friend Sub Initialize(Parent As Object)
51090     On Error GoTo errHandler
51100     If mudtProps.IsNew Then
51110         Set mobjParent = Parent
51120     End If
51130     Exit Sub
errHandler:
51140     If ErrMustStop Then Debug.Assert False: Resume
51150     ErrorIn "a_InvoiceLine.Initialize(Parent)", Parent
End Sub
Public Sub SetParentInvoice(pParentInvoice As Object)
51160     On Error GoTo errHandler
51170     Set mInvoiceParent = pParentInvoice
51180     If Not mInvoiceParent Is Nothing Then
51190         If mInvoiceParent.DocDiscountRate > 0 Then
51200             If Not Me.ServiceItem Then
51210                 mudtProps.DiscountPercent = mInvoiceParent.DocDiscountRate
51220             End If
51230         End If
51240     End If
51250     Exit Sub
errHandler:
51260     If ErrMustStop Then Debug.Assert False: Resume
51270     ErrorIn "a_InvoiceLine.SetParentInvoice(pParentInvoice)", pParentInvoice
End Sub
Friend Property Get IsValid() As Boolean
51280     On Error GoTo errHandler
51290   IsValid = (mobjValid.Count = 0)
51300     Exit Property
errHandler:
51310     If ErrMustStop Then Debug.Assert False: Resume
51320     ErrorIn "a_InvoiceLine.IsValid"
End Property

Private Sub Class_Terminate()
51330     On Error GoTo errHandler
51340     Set moProd = Nothing
51350     Set mchCOFFs = Nothing
51360     If mobjParent Is Nothing Then
51370       If mcolStack.Count > 0 Then _
              Err.Raise vbObjectError + 1001, , "State stack is not empty"
51380     End If
51390     Set mcolStack = Nothing
51400     Exit Sub
errHandler:
51410     If ErrMustStop Then Debug.Assert False: Resume
51420     ErrorIn "a_InvoiceLine.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(ors As String)
51430     On Error GoTo errHandler
51440     RaiseEvent Valid(TranslateErrors(ors))
51450     Exit Sub
errHandler:
51460     If ErrMustStop Then Debug.Assert False: Resume
51470     ErrorIn "a_InvoiceLine.mobjValid_BrokenRule(ors)", ors
End Sub
Private Sub mobjValid_RuleUnbroken(ors As String)
51480     On Error GoTo errHandler
51490     RaiseEvent Valid(TranslateErrors(ors))
51500     Exit Sub
errHandler:
51510     If ErrMustStop Then Debug.Assert False: Resume
51520     ErrorIn "a_InvoiceLine.mobjValid_RuleUnbroken(ors)", ors
End Sub

Private Sub mobjValid_NoBrokenRules()
51530     On Error GoTo errHandler
51540     RaiseEvent Valid("")
51550     Exit Sub
errHandler:
51560     If ErrMustStop Then Debug.Assert False: Resume
51570     ErrorIn "a_InvoiceLine.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
51580     On Error GoTo errHandler
51590     RaiseEvent Valid(TranslateErrors(pMsg))
51600     Exit Sub
errHandler:
51610     If ErrMustStop Then Debug.Assert False: Resume
51620     ErrorIn "a_InvoiceLine.mobjValid_Status(pMsg)", pMsg
End Sub

'Friend Sub Load(InvoiceLineID As Long)
'    If mcolStack.Count > 0 Then Err.Raise 445
'    If Not mudtProps.IsNew Then Err.Raise 445
'
'    mudtProps.IsNew = False
'    Fetch InvoiceLineID
'End Sub

Public Sub Delete()
51630     On Error GoTo errHandler
51640   If mcolStack.Count = 0 Then Err.Raise 445
        
51650   mudtProps.IsDeleted = True
51660   mudtProps.IsDirty = True
51670     Exit Sub
errHandler:
51680     If ErrMustStop Then Debug.Assert False: Resume
51690     ErrorIn "a_InvoiceLine.Delete"
End Sub
Public Property Get BottomOfDocument() As Boolean
51700     On Error GoTo errHandler
51710     BottomOfDocument = (Right(CodeF, 2) = "_B")
51720     Exit Property
errHandler:
51730     If ErrMustStop Then Debug.Assert False: Resume
51740     ErrorIn "a_InvoiceLine.BottomOfDocument"
End Property
Public Property Get IsDeleted() As Boolean
51750     On Error GoTo errHandler
51760   IsDeleted = mudtProps.IsDeleted
51770     Exit Property
errHandler:
51780     If ErrMustStop Then Debug.Assert False: Resume
51790     ErrorIn "a_InvoiceLine.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
51800     On Error GoTo errHandler
51810   IsEditing = (mcolStack.Count > 0)
51820     Exit Property
errHandler:
51830     If ErrMustStop Then Debug.Assert False: Resume
51840     ErrorIn "a_InvoiceLine.IsEditing"
End Property

Public Property Get IsNew() As Boolean
51850     On Error GoTo errHandler
51860   IsNew = mudtProps.IsNew
51870     Exit Property
errHandler:
51880     If ErrMustStop Then Debug.Assert False: Resume
51890     ErrorIn "a_InvoiceLine.IsNew"
End Property

Public Property Get IsDirty() As Boolean
51900     On Error GoTo errHandler
51910   IsDirty = mudtProps.IsDirty Or mchCOFFs.IsDirty
51920     Exit Property
errHandler:
51930     If ErrMustStop Then Debug.Assert False: Resume
51940     ErrorIn "a_InvoiceLine.IsDirty"
End Property


Public Property Get InvoiceLineID() As Long
51950     On Error GoTo errHandler
51960     InvoiceLineID = mudtProps.InvoiceLineID
51970     Exit Property
errHandler:
51980     If ErrMustStop Then Debug.Assert False: Resume
51990     ErrorIn "a_InvoiceLine.InvoiceLineID"
End Property
Public Property Let InvoiceLineID(val As Long)
52000     On Error GoTo errHandler
52010   If mcolStack.Count = 0 Then Err.Raise 383
52020     mudtProps.InvoiceLineID = val
52030     mudtProps.IsDirty = True
52040     Exit Property
errHandler:
52050     If ErrMustStop Then Debug.Assert False: Resume
52060     ErrorIn "a_InvoiceLine.InvoiceLineID"
End Property

Public Property Let InvoiceID(val As Long)
52070     On Error GoTo errHandler
52080   If mcolStack.Count = 0 Then Err.Raise 383
52090     mudtProps.InvoiceID = val
52100     mudtProps.IsDirty = True
52110     Exit Property
errHandler:
52120     If ErrMustStop Then Debug.Assert False: Resume
52130     ErrorIn "a_InvoiceLine.InvoiceID(val)", val
End Property
Public Property Get InvoiceID() As Long
52140     On Error GoTo errHandler
52150   InvoiceID = mudtProps.InvoiceID
52160     Exit Property
errHandler:
52170     If ErrMustStop Then Debug.Assert False: Resume
52180     ErrorIn "a_InvoiceLine.InvoiceID"
End Property
Public Property Get Product() As a_Product
52190     Set Product = moProd
End Property
'Public Function SetLineProduct(Optional PID As String, Optional pCode As String) As Boolean
'    If Not moProd.IsNew Then
'        moProd.PrepareForReload
'    End If
'    With moProd
'        If .Load(PID, 0, Trim$(pCode)) <> 99 Then   'product found
'            Me.Title = .TitleAuthorPublisher
'            Me.PID = .PID
'            Me.Price = .RRP
'            If Not .DefaultCopy Is Nothing Then
'                Me.CopyID = .DefaultCopy.ID
'            End If
'            Me.ProductCode = .Code
'            Me.ProductCodeF = .Code
'            SetLineProduct = True
'        Else
'            SetLineProduct = False   'Not found
'        End If
'    End With
'
'    mobjValid.GetStatus
'End Function

Public Property Get PID() As String
52200     On Error GoTo errHandler
52210     PID = FNS(mudtProps.PID)
52220     Exit Property
errHandler:
52230     If ErrMustStop Then Debug.Assert False: Resume
52240     ErrorIn "a_InvoiceLine.PID"
End Property
Public Property Let PID(val As String)
52250     On Error GoTo errHandler
52260   If mcolStack.Count = 0 Then Err.Raise 383
52270     mudtProps.PID = val
52280     If (mudtProps.PID <= "") Then
52290         mobjValid.RuleBroken "PID", True
52300     Else
52310         mobjValid.RuleBroken "PID", False
52320     End If
52330     mudtProps.IsDirty = True
52340     Exit Property
errHandler:
52350     If ErrMustStop Then Debug.Assert False: Resume
52360     ErrorIn "a_InvoiceLine.PID(val)", val
End Property
Public Property Get PIID() As Long
52370     On Error GoTo errHandler
52380     PIID = mudtProps.PIID
52390     Exit Property
errHandler:
52400     If ErrMustStop Then Debug.Assert False: Resume
52410     ErrorIn "a_InvoiceLine.PIID"
End Property
Public Property Let PIID(val As Long)
52420     On Error GoTo errHandler
52430   If mcolStack.Count = 0 Then Err.Raise 383
52440     mudtProps.PIID = val
52450     mudtProps.IsDirty = True
52460     Exit Property
errHandler:
52470     If ErrMustStop Then Debug.Assert False: Resume
52480     ErrorIn "a_InvoiceLine.PIID(val)", val
End Property

Public Property Get Qty() As Long
52490     On Error GoTo errHandler
52500     Qty = mudtProps.Qty
52510     Exit Property
errHandler:
52520     If ErrMustStop Then Debug.Assert False: Resume
52530     ErrorIn "a_InvoiceLine.Qty"
End Property
Public Property Get QtyF() As String
52540     On Error GoTo errHandler
'52550     If ServiceItem And Right(code, 2) <> "UN" Then   'e.g. postage, insurance etc
'52560         QtyF = ""
'52570     Else
52580         QtyF = Format(Qty, "##0")
'52590     End If
52600     Exit Property
errHandler:
52610     If ErrMustStop Then Debug.Assert False: Resume
52620     ErrorIn "a_InvoiceLine.QtyF"
End Property
Public Property Let Qty(val As Long)
52630     On Error GoTo errHandler
52640   If mcolStack.Count = 0 Then Err.Raise 383
52650     If mudtProps.Qty <> val Then
52660         mudtProps.Qty = val
52670         SetDirty True
52680     End If
52690     Exit Property
errHandler:
52700     If ErrMustStop Then Debug.Assert False: Resume
52710     ErrorIn "a_InvoiceLine.Qty(val)", val
End Property
'============
Public Property Get QtySS() As Long
52720     On Error GoTo errHandler
52730     QtySS = mudtProps.QtySS
52740     Exit Property
errHandler:
52750     If ErrMustStop Then Debug.Assert False: Resume
52760     ErrorIn "a_InvoiceLine.QtySS"
End Property
Public Property Get QtySSF() As String
52770     On Error GoTo errHandler
52780     If ServiceItem Then   'e.g. postage, insurance etc
52790         QtySSF = ""
52800     Else
52810         QtySSF = Format(QtySS, "##0")
52820     End If
52830     Exit Property
errHandler:
52840     If ErrMustStop Then Debug.Assert False: Resume
52850     ErrorIn "a_InvoiceLine.QtySSF"
End Property
Public Property Let QtySS(val As Long)
52860     On Error GoTo errHandler
52870   If mcolStack.Count = 0 Then Err.Raise 383
52880     If mudtProps.QtySS <> val Then
52890         mudtProps.QtySS = val
52900         SetDirty True
52910     End If
52920     Exit Property
errHandler:
52930     If ErrMustStop Then Debug.Assert False: Resume
52940     ErrorIn "a_InvoiceLine.QtySS(val)", val
End Property
'==============
Public Property Get QtyFirm() As Long
52950     On Error GoTo errHandler
52960     QtyFirm = mudtProps.QtyFirm
52970     Exit Property
errHandler:
52980     If ErrMustStop Then Debug.Assert False: Resume
52990     ErrorIn "a_InvoiceLine.QtyFirm"
End Property
Public Property Get QtyFirmF() As String
53000     On Error GoTo errHandler
53010     If ServiceItem Then   'e.g. postage, insurance etc
53020         QtyFirmF = ""
53030     Else
53040         QtyFirmF = Format(QtyFirm, "##0")
53050     End If
53060     Exit Property
errHandler:
53070     If ErrMustStop Then Debug.Assert False: Resume
53080     ErrorIn "a_InvoiceLine.QtyFirmF"
End Property
Public Property Let QtyFirm(val As Long)
53090     On Error GoTo errHandler
53100   If mcolStack.Count = 0 Then Err.Raise 383
53110     If mudtProps.QtyFirm <> val Then
53120         mudtProps.QtyFirm = val
53130         SetDirty True
53140     End If
53150     Exit Property
errHandler:
53160     If ErrMustStop Then Debug.Assert False: Resume
53170     ErrorIn "a_InvoiceLine.QtyFirm(val)", val
End Property
'================




Public Property Get COLID() As Long
53180     On Error GoTo errHandler
53190     COLID = mudtProps.COLID
53200     Exit Property
errHandler:
53210     If ErrMustStop Then Debug.Assert False: Resume
53220     ErrorIn "a_InvoiceLine.COLID"
End Property
Public Property Let COLID(val As Long)
53230     On Error GoTo errHandler
53240   If mcolStack.Count = 0 Then Err.Raise 383
53250     If mudtProps.COLID <> val Then
53260         mudtProps.COLID = val
53270         SetDirty True
53280     End If
53290     Exit Property
errHandler:
53300     If ErrMustStop Then Debug.Assert False: Resume
53310     ErrorIn "a_InvoiceLine.COLID(val)", val
End Property

Public Property Get Publisher() As String
53320     On Error GoTo errHandler
53330     Publisher = stripCRLF(FNS(mudtProps.Publisher))
53340     Exit Property
errHandler:
53350     If ErrMustStop Then Debug.Assert False: Resume
53360     ErrorIn "a_InvoiceLine.Publisher"
End Property
Public Property Let Publisher(val As String)
53370     On Error GoTo errHandler
53380   If mcolStack.Count = 0 Then Err.Raise 383
53390     mudtProps.Publisher = val
53400     mudtProps.IsDirty = True
53410     Exit Property
errHandler:
53420     If ErrMustStop Then Debug.Assert False: Resume
53430     ErrorIn "a_InvoiceLine.Publisher(val)", val
End Property
Public Property Get DistributorName() As String
53440     On Error GoTo errHandler
53450     DistributorName = stripCRLF(FNS(mudtProps.DistributorName))
53460     Exit Property
errHandler:
53470     If ErrMustStop Then Debug.Assert False: Resume
53480     ErrorIn "a_InvoiceLine.DistributorName"
End Property
Public Property Let DistributorName(val As String)
53490     On Error GoTo errHandler
53500   If mcolStack.Count = 0 Then Err.Raise 383
53510     mudtProps.DistributorName = val
53520     mudtProps.IsDirty = True
53530     Exit Property
errHandler:
53540     If ErrMustStop Then Debug.Assert False: Resume
53550     ErrorIn "a_InvoiceLine.DistributorName(val)", val
End Property
Public Property Get DistributorAcno() As String
53560     On Error GoTo errHandler
53570     DistributorAcno = stripCRLF(FNS(mudtProps.DistributorAcno))
53580     Exit Property
errHandler:
53590     If ErrMustStop Then Debug.Assert False: Resume
53600     ErrorIn "a_InvoiceLine.DistributorAcno"
End Property
Public Property Let DistributorAcno(val As String)
53610     On Error GoTo errHandler
53620   If mcolStack.Count = 0 Then Err.Raise 383
53630     mudtProps.DistributorAcno = val
53640     mudtProps.IsDirty = True
53650     Exit Property
errHandler:
53660     If ErrMustStop Then Debug.Assert False: Resume
53670     ErrorIn "a_InvoiceLine.DistributorAcno(val)", val
End Property

'Public Property Get Description() As String
'    Description = Trim$(mudtProps.Description)
'End Property
'Public Property Let Description(Val As String)
'  If mcolStack.Count = 0 Then Err.Raise 383
'    mudtProps.Description = Val
'    mudtProps.IsDirty = True
'End Property

Public Property Get Note() As String
53680     On Error GoTo errHandler
53690     Note = FNS(mudtProps.Note)
53700     Exit Property
errHandler:
53710     If ErrMustStop Then Debug.Assert False: Resume
53720     ErrorIn "a_InvoiceLine.Note"
End Property
Public Property Let Note(val As String)
53730     On Error GoTo errHandler
53740   If mcolStack.Count = 0 Then Err.Raise 383
53750     mudtProps.Note = stripCRLF(val)
53760     mudtProps.IsDirty = True
53770     Exit Property
errHandler:
53780     If ErrMustStop Then Debug.Assert False: Resume
53790     ErrorIn "a_InvoiceLine.Note(val)", val
End Property

Public Property Get Ref() As String
53800     On Error GoTo errHandler
53810     Ref = stripCRLF(FNS(mudtProps.Ref))
53820     Exit Property
errHandler:
53830     If ErrMustStop Then Debug.Assert False: Resume
53840     ErrorIn "a_InvoiceLine.Ref"
End Property
Public Property Let Ref(val As String)
53850     On Error GoTo errHandler
53860     mudtProps.Ref = val
53870     Exit Property
errHandler:
53880     If ErrMustStop Then Debug.Assert False: Resume
53890     ErrorIn "a_InvoiceLine.Ref(val)", val
End Property

Public Property Get code() As String
10        On Error GoTo errHandler
20        code = stripCRLF(FNS(mudtProps.code))
30        Exit Property
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "a_InvoiceLine.Code"
End Property
Public Property Get EAN() As String
53950     On Error GoTo errHandler
53960     EAN = stripCRLF(FNS(mudtProps.EAN))
53970     Exit Property
errHandler:
53980     If ErrMustStop Then Debug.Assert False: Resume
53990     ErrorIn "a_InvoiceLine.EAN"
End Property
Public Property Let EAN(val As String)
53850     On Error GoTo errHandler
53860     mudtProps.EAN = val
53870     Exit Property
errHandler:
53880     If ErrMustStop Then Debug.Assert False: Resume
53890     ErrorIn "a_InvoiceLine.EAN(val)", val
End Property
Sub TestEvent()
54000     On Error GoTo errHandler
54010 RaiseEvent Valid("")
54020     Exit Sub
errHandler:
54030     If ErrMustStop Then Debug.Assert False: Resume
54040     ErrorIn "a_InvoiceLine.TestEvent"
End Sub
Public Property Let code(val As String)
54050     On Error GoTo errHandler
      Dim oSM As New z_StockManager

54060   If mcolStack.Count = 0 Then Err.Raise 383
54070     mudtProps.code = val
54080     If Left(val, 1) <> "2" Then
54090         If Len(val) = 13 And mudtProps.CodeF = "" Then
54100             mudtProps.CodeF = FormatISBN13(val)
54110         End If
54120     End If
          
54130     mudtProps.IsDirty = True
54140     Exit Property
errHandler:
54150     If ErrMustStop Then Debug.Assert False: Resume
54160     ErrorIn "a_InvoiceLine.Code(val)", val
End Property
Public Property Let CNLQty(val As Integer)
54170     mudtProps.tmpCNLQty = val
End Property
Public Property Get CNLQty() As Integer
54180     CNLQty = FNN(mudtProps.tmpCNLQty)
End Property
Public Sub SetMarker(bON As Boolean)
54190     mudtProps.Marker = bON
End Sub
Public Property Get Marker() As Boolean
54200     Marker = FNB(mudtProps.Marker)
End Property
Public Property Get CodeF() As String
54210     On Error GoTo errHandler
54220     CodeF = FNS(mudtProps.CodeF)
54230     Exit Property
errHandler:
54240     If ErrMustStop Then Debug.Assert False: Resume
54250     ErrorIn "a_InvoiceLine.CodeF"
End Property
Public Property Let CodeF(val As String)
54260     On Error GoTo errHandler
54270   If mcolStack.Count = 0 Then Err.Raise 383
54280     mudtProps.CodeF = val
54290     mudtProps.IsDirty = True
54300     Exit Property
errHandler:
54310     If ErrMustStop Then Debug.Assert False: Resume
54320     ErrorIn "a_InvoiceLine.CodeF(val)", val
End Property
Public Property Get CodeForExport() As String
54330     On Error GoTo errHandler
54340   CodeForExport = FNS(mudtProps.CodeForExport)
54350     Exit Property
errHandler:
54360     If ErrMustStop Then Debug.Assert False: Resume
54370     ErrorIn "a_InvoiceLine.CodeForExport"
End Property
Public Property Let CodeForExport(val As String)
54380     On Error GoTo errHandler
54390   If mcolStack.Count = 0 Then Err.Raise 383
54400     mudtProps.CodeForExport = val
54410     Exit Property
errHandler:
54420     If ErrMustStop Then Debug.Assert False: Resume
54430     ErrorIn "a_InvoiceLine.CodeForExport(val)", val
End Property
Public Property Get TitleAuthor()
54440     On Error GoTo errHandler
      Dim tmp As String

54450     tmp = Title
54460     If Len(tmp) > 0 Then
54470         tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
54480     End If
54490     TitleAuthor = stripCRLF(tmp)

54500     Exit Property
errHandler:
54510     If ErrMustStop Then Debug.Assert False: Resume
54520     ErrorIn "a_InvoiceLine.TitleAuthor"
End Property
Public Property Get TitleAuthorPublisher()
54530     On Error GoTo errHandler
      Dim tmp As String

54540     tmp = Title
54550     If Len(tmp) > 0 Then
54560         tmp = tmp & IIf(Len(MainAuthor) > 0, "(" & MainAuthor & ")", "")
54570     End If
54580     If Len(tmp) > 0 Then
54590         tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
54600     End If
54610     TitleAuthorPublisher = stripCRLF(tmp)

54620     Exit Property
errHandler:
54630     If ErrMustStop Then Debug.Assert False: Resume
54640     ErrorIn "a_InvoiceLine.TitleAuthorPublisher"
End Property
Public Property Get TitleAuthorPublisherWithoutBRackets()
54650     On Error GoTo errHandler
      Dim tmp As String

54660     tmp = Title
54670     If Len(tmp) > 0 Then
54680         tmp = tmp & IIf(Len(MainAuthor) > 0, " " & MainAuthor & " ", "")
54690     End If
54700     If Len(tmp) > 0 Then
54710         tmp = tmp & IIf(Len(Publisher) > 0, "(" & Publisher & ")", "")
54720     End If
54730     TitleAuthorPublisherWithoutBRackets = stripCRLF(tmp)

54740     Exit Property
errHandler:
54750     If ErrMustStop Then Debug.Assert False: Resume
54760     ErrorIn "a_InvoiceLine.TitleAuthorPublisherWithoutBRackets"
End Property
Public Property Get TitleAuthorPublisherL(pLen As Integer)
54770     On Error GoTo errHandler
54780     TitleAuthorPublisherL = stripCRLF(Left(TitleAuthorPublisher, pLen) & IIf(Len(TitleAuthorPublisher) > pLen, "...", ""))
54790     Exit Property
errHandler:
54800     If ErrMustStop Then Debug.Assert False: Resume
54810     ErrorIn "a_InvoiceLine.TitleAuthorPublisherL(pLen)", pLen
End Property
Public Property Get SpecialVATMsg() As String
54820     On Error GoTo errHandler
54830     If Me.VATRate <> oPC.Configuration.VATRate And oPC.Configuration.IsVATRegion Then
54840          SpecialVATMsg = "(" & VATRateF & " VAT) "
54850     End If
54860     Exit Property
errHandler:
54870     If ErrMustStop Then Debug.Assert False: Resume
54880     ErrorIn "a_InvoiceLine.SpecialVATMsg"
End Property
Public Property Get Title() As String
54890     On Error GoTo errHandler
54900     Title = stripCRLF(FNS(mudtProps.Title))
54910     Exit Property
errHandler:
54920     If ErrMustStop Then Debug.Assert False: Resume
54930     ErrorIn "a_InvoiceLine.Title"
End Property
Public Property Get TitleWithArticle() As String
54940     On Error GoTo errHandler
54950     TitleWithArticle = stripCRLF(FNS(mudtProps.TitleWithArticle))
54960     Exit Property
errHandler:
54970     If ErrMustStop Then Debug.Assert False: Resume
54980     ErrorIn "a_InvoiceLine.TitleWithArticle"
End Property
Public Property Let Title(val As String)
54990     On Error GoTo errHandler
55000   If mcolStack.Count = 0 Then Err.Raise 383
55010     mudtProps.Title = val
55020     mudtProps.IsDirty = True
55030     Exit Property
errHandler:
55040     If ErrMustStop Then Debug.Assert False: Resume
55050     ErrorIn "a_InvoiceLine.Title(val)", val
End Property
Public Property Get TitleF(pLen As Integer) As String
55060     On Error GoTo errHandler
55070     TitleF = Left(Title, pLen) & IIf(Len(Title) > pLen, "...", "")
55080     Exit Property
errHandler:
55090     If ErrMustStop Then Debug.Assert False: Resume
55100     ErrorIn "a_InvoiceLine.TitleF(pLen)", pLen
End Property
Public Property Get GDNCode() As String
55110     On Error GoTo errHandler
55120     GDNCode = FNS(mudtProps.GDNCode)
55130     Exit Property
errHandler:
55140     If ErrMustStop Then Debug.Assert False: Resume
55150     ErrorIn "a_InvoiceLine.GDNCode"
End Property

Public Property Get MainAuthor() As String
55160     On Error GoTo errHandler
55170     MainAuthor = stripCRLF(FNS(mudtProps.MainAuthor))
           '   tmp = tmp & IIf(Len(MainAuthor) > 0, " " & MainAuthor & " ", "")
55180     Exit Property
errHandler:
55190     If ErrMustStop Then Debug.Assert False: Resume
55200     ErrorIn "a_InvoiceLine.MainAuthor"
End Property

Public Property Get CO_StaffShortname() As String
55210     On Error GoTo errHandler
55220     CO_StaffShortname = FNS(mudtProps.CO_StaffShortname)
55230     Exit Property
errHandler:
55240     If ErrMustStop Then Debug.Assert False: Resume
55250     ErrorIn "a_InvoiceLine.CO_StaffShortname"
End Property
Public Property Let MainAuthor(val As String)
55260     On Error GoTo errHandler
55270   If mcolStack.Count = 0 Then Err.Raise 383
55280     mudtProps.MainAuthor = val
55290     mudtProps.IsDirty = True
55300     Exit Property
errHandler:
55310     If ErrMustStop Then Debug.Assert False: Resume
55320     ErrorIn "a_InvoiceLine.MainAuthor(val)", val
End Property

Public Property Get DiscountPercent() As Double
55330     On Error GoTo errHandler
55340     DiscountPercent = mudtProps.DiscountPercent
55350     Exit Property
errHandler:
55360     If ErrMustStop Then Debug.Assert False: Resume
55370     ErrorIn "a_InvoiceLine.DiscountPercent"
End Property
Public Property Let DiscountPercent(val As Double)
55380     On Error GoTo errHandler
55390   If mcolStack.Count = 0 Then Err.Raise 383
55400     If Not ServiceItem Then  'non ServiceItem items cant have discount
           '   If mInvoiceParent.DocDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
           '       mudtProps.DiscountPercent = mInvoiceParent.DocDiscountRate
           '   Else
55410             mudtProps.DiscountPercent = val
           '   End If
55420         If (val < 0 Or val > 100) Then
55430             mobjValid.RuleBroken "Discount", True
55440         Else
55450             mobjValid.RuleBroken "Discount", False
55460         End If
55470         mudtProps.IsDirty = True
        '      RaiseEvent ExtensionChange(Extension, ExtensionF)
55480     End If
55490     Exit Property
errHandler:
55500     If ErrMustStop Then Debug.Assert False: Resume
55510     ErrorIn "a_InvoiceLine.DiscountPercent(val)", val
End Property
Public Function SetDiscountPercent(val As String) As Boolean
55520     On Error GoTo errHandler
      Dim dblDiscountPercent As Double
55530     SetDiscountPercent = True
55540     If Not ServiceItem Then  'non-stock items can't have discount
55550         If Not mInvoiceParent Is Nothing Then
55560             If mInvoiceParent.DocDiscountRate > 0 Then 'we only allow discount here if no general discount applicable
55570                 mudtProps.DiscountPercent = mInvoiceParent.DocDiscountRate
55580                 Exit Function
55590             End If
55600         End If
              
55610         If mcolStack.Count = 0 Then Err.Raise 383
55620         If Trim$(val) = "" Then
55630             dblDiscountPercent = 0
55640         ElseIf Not ConvertToDBL(val, dblDiscountPercent) Then
55650             SetDiscountPercent = False
55660             Exit Function
55670         End If
55680         If (dblDiscountPercent < 0 Or dblDiscountPercent > 100) Then
55690             mobjValid.RuleBroken "Discount", True
55700         Else
55710             mobjValid.RuleBroken "Discount", False
55720         End If
55730         mudtProps.DiscountPercent = dblDiscountPercent
55740         mudtProps.IsDirty = True
55750     End If
55760     Exit Function
errHandler:
55770     If ErrMustStop Then Debug.Assert False: Resume
55780     ErrorIn "a_InvoiceLine.SetDiscountPercent(val)", val
End Function
Public Function DiscountPercentF() As String
55790     On Error GoTo errHandler
55800     If mudtProps.DiscountPercent = 0 Then
55810         DiscountPercentF = ""
55820     Else
55830         DiscountPercentF = Format(mudtProps.DiscountPercent, gPercentFormatString)
55840     End If
55850     Exit Function
errHandler:
55860     If ErrMustStop Then Debug.Assert False: Resume
55870     ErrorIn "a_InvoiceLine.DiscountPercentF"
End Function
Public Property Get VATRate() As Double
55880     On Error GoTo errHandler
55890     VATRate = mudtProps.VATRate
55900     Exit Property
errHandler:
55910     If ErrMustStop Then Debug.Assert False: Resume
55920     ErrorIn "a_InvoiceLine.VATRate"
End Property
Public Property Get VATRateF() As String
55930     On Error GoTo errHandler
55940     VATRateF = Format(mudtProps.VATRate, gPercentFormatString)
55950     Exit Property
errHandler:
55960     If ErrMustStop Then Debug.Assert False: Resume
55970     ErrorIn "a_InvoiceLine.VATRateF"
End Property

Public Property Let VATRate(val As Double)
55980     On Error GoTo errHandler
55990     If mcolStack.Count = 0 Then Err.Raise 383
56000     mudtProps.VATRate = val
56010     mudtProps.IsDirty = True
56020     Exit Property
errHandler:
56030     If ErrMustStop Then Debug.Assert False: Resume
56040     ErrorIn "a_InvoiceLine.VATRate(val)", val
End Property
Public Property Get ServiceItem() As Boolean
56050     On Error GoTo errHandler
56060     ServiceItem = mudtProps.ServiceItem
56070     Exit Property
errHandler:
56080     If ErrMustStop Then Debug.Assert False: Resume
56090     ErrorIn "a_InvoiceLine.ServiceItem"
End Property
Public Property Let ServiceItem(val As Boolean)
56100     On Error GoTo errHandler
56110     If mcolStack.Count = 0 Then Err.Raise 383
56120     mudtProps.ServiceItem = val
56130     mudtProps.IsDirty = True
56140     Exit Property
errHandler:
56150     If ErrMustStop Then Debug.Assert False: Resume
56160     ErrorIn "a_InvoiceLine.ServiceItem(val)", val
End Property
'---------------------------------
'Public Property Get VATAmount() As Long
'    VATAmount = CLng(CDbl((VATRate / (100 + VATRate))) * ExtensionInclDeposit)
'End Property
'Public Property Get VATAmount_foreign() As Long
'    VATAmount_foreign = CLng((VATRate / (100 + VATRate)) * ExtensionInclDeposit_Foreign)
'End Property
'Public Property Get VATAmountF(bForeign As Boolean) As String
'    If bForeign Then
'        VATAmountF = Format(CDbl(VATAmount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        VATAmountF = Format(CDbl(VATAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get ExtensionExVAT() As Long
'    ExtensionExVAT = (CLng((mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))) * (100 / (100 + mudtProps.VATRate))
'End Property
'Public Property Get ExtensionExVATF(bForeign As Boolean) As String
'    If bForeign Then
'        ExtensionExVATF = Format(CDbl(ExtensionExVAT / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        ExtensionExVATF = Format(CDbl(ExtensionExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get ExtensionInclDeposit() As Long
'    ExtensionInclDeposit = CLng(CLng(mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get ExtensionInclDeposit_Foreign() As Long
'    ExtensionInclDeposit_Foreign = CLng(CLng(ForeignPrice * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get Extension() As Long
'    Extension = CLng(CLng(mudtProps.Price * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get Extension_Foreign() As Long
'    Extension_Foreign = CLng(CLng(ForeignPrice * mudtProps.Qty) * ((100 - mudtProps.DiscountPercent) / 100))
'End Property
'Public Property Get ForeignPrice() As Long
'    ForeignPrice = mudtProps.ForeignPrice 'RoundUp(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor, oPC.Configuration.RoundingRules.GetRoundTo(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor)) * mInvoiceParent.ForeignCurrency.Divisor
'End Property
'
'Public Property Get ExtensionF(bForeign As Boolean) As String
'    If bForeign Then
'        ExtensionF = Format(CDbl(Extension_Foreign / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        ExtensionF = Format(CDbl(Extension / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
'Public Property Get ExtensionSimple() As Long
'    ExtensionSimple = CLng(mudtProps.Price * mudtProps.Qty)
'End Property
'Public Property Get ExtensionSimple_foreign() As Long
'    ExtensionSimple_foreign = CLng(mudtProps.ForeignPrice * mudtProps.Qty)
'End Property
'Public Property Get ExtensionSimpleF(bForeign As Boolean) As String
'    If bForeign Then
'        ExtensionSimpleF = Format(CDbl(ExtensionSimple_foreign / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        ExtensionSimpleF = Format(CDbl(ExtensionSimple / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property
Public Function PriceF(bForeign As Boolean) As String
56170     On Error GoTo errHandler
56180     If bForeign Then
56190         PriceF = Format(CDbl(mudtProps.ForeignPrice / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
56200     Else
56210         PriceF = Format(CDbl(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
56220     End If
56230     Exit Function
errHandler:
56240     If ErrMustStop Then Debug.Assert False: Resume
56250     ErrorIn "a_InvoiceLine.PriceF(bForeign)", bForeign
End Function
Public Function PriceExVatF(bForeign As Boolean) As String
56260     On Error GoTo errHandler
56270     If bForeign Then
              'surely this wont be used
56280         PriceExVatF = Format(CDbl(mudtProps.ForeignPrice * (100 / (100 + oPC.Configuration.VATRate)) / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
56290     Else
56300         PriceExVatF = Format(CDbl(Me.PriceExVAT / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
56310     End If
56320     Exit Function
errHandler:
56330     If ErrMustStop Then Debug.Assert False: Resume
56340     ErrorIn "a_InvoiceLine.PriceExVatF(bForeign)", bForeign
End Function

'Public Property Get DiscountAmount() As Long
'    DiscountAmount = CLng(ExtensionSimple - ExtensionInclDeposit)
'End Property
'Public Property Get DiscountAmount_Foreign() As Long
'    DiscountAmount_Foreign = CLng(ExtensionSimple_foreign - ExtensionInclDeposit_Foreign)
'End Property
'Public Property Get DiscountAmountF(bForeign As Boolean) As String
'    If bForeign Then
'        DiscountAmountF = Format(CDbl(DiscountAmount / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
'    Else
'        DiscountAmountF = Format(CDbl(DiscountAmount / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
'    End If
'End Property

Public Property Get Deposit_Foreign() As Long
56350     On Error GoTo errHandler
56360     Deposit_Foreign = mudtProps.ForeignDeposit
56370     Exit Property
errHandler:
56380     If ErrMustStop Then Debug.Assert False: Resume
56390     ErrorIn "a_InvoiceLine.Deposit_Foreign"
End Property
Public Property Let Deposit_Foreign(val As Long)
56400     On Error GoTo errHandler
56410   If mcolStack.Count = 0 Then Err.Raise 383
56420     mudtProps.ForeignDeposit = val
56430     mudtProps.IsDirty = True
56440     Exit Property
errHandler:
56450     If ErrMustStop Then Debug.Assert False: Resume
56460     ErrorIn "a_InvoiceLine.Deposit_Foreign(val)", val
End Property
Public Property Get Deposit() As Long
56470     On Error GoTo errHandler
56480     Deposit = mudtProps.Deposit
56490     Exit Property
errHandler:
56500     If ErrMustStop Then Debug.Assert False: Resume
56510     ErrorIn "a_InvoiceLine.Deposit"
End Property
Public Property Let Deposit(val As Long)
56520     On Error GoTo errHandler
56530   If mcolStack.Count = 0 Then Err.Raise 383
56540     If val <> mudtProps.Deposit Then
56550         mudtProps.Deposit = val
56560         SetDirty True
56570     End If
56580     Exit Property
errHandler:
56590     If ErrMustStop Then Debug.Assert False: Resume
56600     ErrorIn "a_InvoiceLine.Deposit(val)", val
End Property


Public Property Get APPLID() As Long
56610     On Error GoTo errHandler
56620     APPLID = mudtProps.APPLID
56630     Exit Property
errHandler:
56640     If ErrMustStop Then Debug.Assert False: Resume
56650     ErrorIn "a_InvoiceLine.APPLID"
End Property
Public Property Let APPLID(val As Long)
56660     On Error GoTo errHandler
56670   If mcolStack.Count = 0 Then Err.Raise 383
56680     If val <> mudtProps.APPLID Then
56690         mudtProps.APPLID = val
56700         SetDirty True
56710     End If
56720     Exit Property
errHandler:
56730     If ErrMustStop Then Debug.Assert False: Resume
56740     ErrorIn "a_InvoiceLine.APPLID(val)", val
End Property

Public Property Get APPLQTY() As Long
56750     On Error GoTo errHandler
56760     APPLQTY = mudtProps.APPLQTY
56770     Exit Property
errHandler:
56780     If ErrMustStop Then Debug.Assert False: Resume
56790     ErrorIn "a_InvoiceLine.APPLQTY"
End Property
Public Property Let APPLQTY(val As Long)
56800     On Error GoTo errHandler
56810   If mcolStack.Count = 0 Then Err.Raise 383
56820     If val <> mudtProps.APPLQTY Then
56830         mudtProps.APPLQTY = val
56840         SetDirty True
56850     End If
56860     Exit Property
errHandler:
56870     If ErrMustStop Then Debug.Assert False: Resume
56880     ErrorIn "a_InvoiceLine.APPLQTY(val)", val
End Property
Public Property Let SubstitutesAvailable(pVal As Boolean)
56890     mudtProps.SubstitutesAvailable = pVal
56900     SetDirty (True)
End Property
Public Property Get SubstitutesAvailable() As Boolean
56910     SubstitutesAvailable = FNB(mudtProps.SubstitutesAvailable)
End Property
Public Property Get CreditedQty() As Long
56920     On Error GoTo errHandler
56930     CreditedQty = mudtProps.CreditedQty
56940     Exit Property
errHandler:
56950     If ErrMustStop Then Debug.Assert False: Resume
56960     ErrorIn "a_InvoiceLine.CreditedQty"
End Property

Public Function SetDeposit(val As String) As Boolean
56970     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim lngTemp As Long
56980     lngTemp = mudtProps.Deposit
56990     bOK = SetField_strAsCurrencyToLong(mudtProps.Deposit, val, mcolStack.Count, "DEPOSIT", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
57000     If bOK And (lngTemp <> mudtProps.Deposit) Then 'only validate if conversion is OK and the new value is other than the original
57010         bOK = ValidateObject("DEPOSIT")
57020     End If
57030     SetDeposit = bOK
57040     Exit Function
errHandler:
57050     If ErrMustStop Then Debug.Assert False: Resume
57060     ErrorIn "a_InvoiceLine.SetDeposit(val)", val
End Function
Public Function DepositF(bForeign As Boolean) As String
57070     On Error GoTo errHandler
57080     If bForeign Then
57090         DepositF = Format(CDbl(mudtProps.ForeignDeposit / mInvoiceParent.ForeignCurrency.Divisor), mInvoiceParent.ForeignCurrency.FormatString)
57100     Else
57110         DepositF = Format(CDbl(mudtProps.Deposit / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
57120     End If
57130     Exit Function
errHandler:
57140     If ErrMustStop Then Debug.Assert False: Resume
57150     ErrorIn "a_InvoiceLine.DepositF(bForeign)", bForeign
End Function
Public Function SetNote(val As String) As Boolean
57160     If mcolStack.Count = 0 Then Err.Raise 383
57170     If Len(val) > Len(mudtProps.Note) Then
57180         Err.Raise 384
57190     End If
57200     If val <> mudtProps.Note Then
57210         mudtProps.Note = val
57220         SetDirty True
57230     End If
57240     SetNote = True
57250     Exit Function
End Function
Public Function SetRef(val As String) As Boolean
57260     If mcolStack.Count = 0 Then Err.Raise 383
57270     If Len(val) > Len(mudtProps.Ref) Then
57280         Err.Raise 384
57290     End If
57300     If val <> mudtProps.Ref Then
57310         mudtProps.Ref = val
57320         SetDirty True
57330     End If
57340     SetRef = True
57350     Exit Function
End Function

'------------------------------
'
'Public Property Get ForeignPrice() As Long
'    ForeignPrice = mudtProps.ForeignPrice
'End Property
Public Property Let ForeignPrice(val As Long)
57360     On Error GoTo errHandler
57370   If mcolStack.Count = 0 Then Err.Raise 383
57380     If val <> mudtProps.ForeignPrice Then
57390         mudtProps.ForeignPrice = val
57400         SetDirty True
57410     End If
57420     Exit Property
errHandler:
57430     If ErrMustStop Then Debug.Assert False: Resume
57440     ErrorIn "a_InvoiceLine.ForeignPrice(val)", val
End Property
Public Property Get ForeignPrice() As Long
57450     On Error GoTo errHandler
57460     ForeignPrice = mudtProps.ForeignPrice
57470     Exit Property
errHandler:
57480     If ErrMustStop Then Debug.Assert False: Resume
57490     ErrorIn "a_InvoiceLine.ForeignPrice"
End Property
Public Property Get PriceExVAT() As Long
57500         On Error GoTo errHandler
57510         If mInvoiceParent Is Nothing Then
57520             PriceExVAT = 0
57530             Exit Property
57540         End If
57550         If mInvoiceParent.Customer Is Nothing Then
57560             PriceExVAT = 0
57570             Exit Property
57580         End If
57590     If (mInvoiceParent.Customer.VATable = False And mInvoiceParent.Customer.ShowVAT = False) Then
57600         PriceExVAT = FNN(mudtProps.Price)
57610     Else
57620         PriceExVAT = Round(FNN(mudtProps.Price) * (100 / (100 + FNDBL(mudtProps.VATRate))))
57630     End If
          'Can't do this as some items are really zero vat - like postage etc
       '   PriceExVAT = Round(mudtProps.Price * (100 / (100 + IIf(mudtProps.VATRate = 0, oPC.Configuration.VATRate, mudtProps.VATRate))))
57640     Exit Property
errHandler:
57650     If ErrMustStop Then Debug.Assert False: Resume
57660     ErrorIn "a_InvoiceLine.PriceExVAT"
End Property
Public Property Get Price() As Long
57670     On Error GoTo errHandler
57680     Price = mudtProps.Price
57690     Exit Property
errHandler:
57700     If ErrMustStop Then Debug.Assert False: Resume
57710     ErrorIn "a_InvoiceLine.Price"
End Property
Public Property Let Price(val As Long)
57720     On Error GoTo errHandler
57730   If mcolStack.Count = 0 Then Err.Raise 383
57740     If val <> mudtProps.Price Then
57750         mudtProps.Price = val
57760         SetDirty True
57770     End If
57780     ValidateObject "PRICE"
57790     Exit Property
errHandler:
57800     If ErrMustStop Then Debug.Assert False: Resume
57810     ErrorIn "a_InvoiceLine.Price(val)", val
End Property
Public Property Get Cost() As Long
57820     On Error GoTo errHandler
57830     Cost = mudtProps.Cost
57840     Exit Property
errHandler:
57850     If ErrMustStop Then Debug.Assert False: Resume
57860     ErrorIn "a_InvoiceLine.Cost"
End Property
Public Property Let Cost(val As Long)
57870     On Error GoTo errHandler
57880   If mcolStack.Count = 0 Then Err.Raise 383
57890     If val <> mudtProps.Cost Then
57900         mudtProps.Cost = val
57910         SetDirty True
57920     End If
57930     ValidateObject "Cost"
57940     Exit Property
errHandler:
57950     If ErrMustStop Then Debug.Assert False: Resume
57960     ErrorIn "a_InvoiceLine.Cost(val)", val
End Property
Public Function CostF() As String
57970     On Error GoTo errHandler
57980         CostF = Format(CDbl(mudtProps.Cost / oPC.Configuration.DefaultCurrency.Divisor), oPC.Configuration.DefaultCurrency.FormatString)
57990     Exit Function
errHandler:
58000     If ErrMustStop Then Debug.Assert False: Resume
58010     ErrorIn "a_InvoiceLine.PCostF"
End Function

Public Property Get CodeForEditing() As String
58020     If Not IsISBN10(FNS(mudtProps.code)) And FNS(mudtProps.code) > "" Then
58030         CodeForEditing = FNS(mudtProps.code)
58040     Else
58050         If FNS(mudtProps.CodeF) > "" Then
58060             CodeForEditing = FNS(mudtProps.CodeF)
58070         Else
58080             If IsISBN13(FNS(mudtProps.EAN), True) Then
58090                 CodeForEditing = FNS(mudtProps.EAN)
58100             Else
58110                 If FNS(mudtProps.code) > "" Then
58120                     CodeForEditing = FNS(mudtProps.code)
58130                 Else
58140                     CodeForEditing = FNS(mudtProps.EAN)
58150                 End If
58160             End If
58170         End If
58180     End If
          
End Property
Public Function SetPrice(val As String) As Boolean
58190     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim lngTemp As Long
58200     lngTemp = mudtProps.Price
58210     bOK = SetField_strAsCurrencyToLong(mudtProps.Price, val, mcolStack.Count, "PRICE", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
58220     If bOK And (lngTemp <> mudtProps.Price) Then 'only validate if conversion is OK and the new value is other than the original
58230         bOK = ValidateObject("PRICE")
58240     End If
58250     SetPrice = bOK
58260     Exit Function
errHandler:
58270     If ErrMustStop Then Debug.Assert False: Resume
58280     ErrorIn "a_InvoiceLine.SetPrice(val)", val
End Function
Public Function SetQty(val As String) As Boolean
58290     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim lngTemp As Long
      val = Left(val, (Len(val) - InStr(val, "(")))  'remove the unwanted brackets in, for example 21(0)
58300     lngTemp = mudtProps.Qty
58310     bOK = SetField_LONG(mudtProps.Qty, val, "QTY", mcolStack.Count)
58320     If bOK And (lngTemp <> mudtProps.Qty) Then  'only validate if conversion is OK and the new value is other than the original
58330         bOK = ValidateObject("QTY")
58340     End If
58350     SetQty = bOK
58360     Exit Function
errHandler:
58370     If ErrMustStop Then Debug.Assert False: Resume
58380     ErrorIn "a_InvoiceLine.SetQty(val)", val
End Function

Public Function SetQtyFirm(val As String) As Boolean
58390     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim lngTemp As Long
58400     lngTemp = mudtProps.QtyFirm
58410     bOK = SetField_LONG(mudtProps.QtyFirm, val, "QTY", mcolStack.Count)
58420     If bOK And (lngTemp <> mudtProps.QtyFirm) Then  'only validate if conversion is OK and the new value is other than the original
58430         bOK = ValidateObject("QTY")
58440     End If
58450     SetQtyFirm = bOK
58460     Exit Function
errHandler:
58470     If ErrMustStop Then Debug.Assert False: Resume
58480     ErrorIn "a_InvoiceLine.SetQtyFirm(val)", val
End Function
Public Function SetQtySS(val As String) As Boolean
58490     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim lngTemp As Long
58500     lngTemp = mudtProps.QtySS
58510     bOK = SetField_LONG(mudtProps.QtySS, val, "QTY", mcolStack.Count)
58520     If bOK And (lngTemp <> mudtProps.QtySS) Then  'only validate if conversion is OK and the new value is other than the original
58530         bOK = ValidateObject("QTY")
58540     End If
58550     SetQtySS = bOK
58560     Exit Function
errHandler:
58570     If ErrMustStop Then Debug.Assert False: Resume
58580     ErrorIn "a_InvoiceLine.SetQtySS(val)", val
End Function

Private Function TranslateErrors(pRawors As String) As String
58590     On Error GoTo errHandler
      Dim strOut As String
      Dim strRule, strAllRules As String
      Dim NoMoreRules As Boolean
      Dim iMarker, iStart As Integer
58600     iMarker = 1
58610     strAllRules = ""
58620     If Len(pRawors) > 0 Then
58630         iMarker = InStr(iMarker + 1, pRawors, ",")
58640         If iMarker > 0 Then
58650             strAllRules = colClassors(Left(pRawors, iMarker - 1))
58660         Else
58670             strAllRules = colClassors(pRawors)
58680         End If
58690         Do Until iMarker = 0
58700             iStart = iMarker + 1
58710             iMarker = InStr(iStart, pRawors, ",")
58720             If iMarker > 0 Then
58730                 strRule = colClassors(Mid(pRawors, iStart, iMarker - iStart))
58740             Else
58750                 strRule = colClassors(Mid(pRawors, iStart))
58760             End If
                      
58770             strAllRules = strAllRules & vbCrLf & strRule
58780         Loop
58790     End If
58800     TranslateErrors = strAllRules
58810     Exit Function
errHandler:
58820     If ErrMustStop Then Debug.Assert False: Resume
58830     ErrorIn "a_InvoiceLine.TranslateErrors(pRawors)", pRawors
End Function

Private Sub SetDirty(pVal As Boolean)
58840     On Error GoTo errHandler
58850     If pVal Then   ' can only set to true, not to false, because other properties may have been changed
58860         mudtProps.IsDirty = True
58870     End If
58880     Exit Sub
errHandler:
58890     If ErrMustStop Then Debug.Assert False: Resume
58900     ErrorIn "a_InvoiceLine.SetDirty(pVal)", pVal
End Sub

Private Sub LoadClassorsCollection()
58910     On Error GoTo errHandler
58920     Set colClassors = New Collection
58930     colClassors.Add "Invalid price", "PRICE"
58940     colClassors.Add "Invalid quantity", "QTY"
58950     colClassors.Add "Invalid ETA", "ETA"
58960     colClassors.Add "Invalid discount", "DISCOUNT"
58970     colClassors.Add "Invalid deposit", "DEPOSIT"
58980     colClassors.Add "A numeric value is too great for calculation", "OVERFLOW"
58990     colClassors.Add "A pre-delivery invoice must match a customer order", "PREDEL"
59000     Exit Sub
errHandler:
59010     If ErrMustStop Then Debug.Assert False: Resume
59020     ErrorIn "a_InvoiceLine.LoadClassorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
59030     On Error GoTo errHandler
      Dim bValid As Boolean
59040     bValid = True
59050     If mudtProps.Price > 0 Then
59060         If pFld = "PRICE" Then bValid = False
59070         mobjValid.RuleBroken "PRICE", False
59080     Else
59090         mobjValid.RuleBroken "PRICE", True
59100     End If
59110     If mudtProps.Qty >= 0 And mudtProps.QtyFirm >= 0 And mudtProps.QtySS >= 0 Then
59120         mobjValid.RuleBroken "QTY", False
59130     Else
59140         If pFld = "QTY" Then bValid = False
59150         mobjValid.RuleBroken "QTY", True
59160     End If
59170     If (mudtProps.DiscountPercent < 0 Or mudtProps.DiscountPercent > 100) Then
59180         If pFld = "DISCOUNT" Then bValid = False
59190         mobjValid.RuleBroken "DISCOUNT", True
59200     Else
59210         mobjValid.RuleBroken "DISCOUNT", False
59220     End If
59230     If Not mInvoiceParent Is Nothing Then
59240         If mInvoiceParent.IsInvoice Then
59250             If mInvoiceParent.IsPreDelivery = True And mudtProps.COLID = 0 Then
59260                 mobjValid.RuleBroken "PREDEL", True
59270             Else
59280                 mobjValid.RuleBroken "PREDEL", False
59290             End If
59300         Else
59310                 mobjValid.RuleBroken "PREDEL", False
59320         End If
59330     End If
59340     mobjValid.GetSTatus
59350     SetDirty bValid
59360     ValidateObject = bValid
EXIT_Handler:
      'H:
      '    MsgBox "CheckOK: " & or
59370     Exit Function
errHandler:
59380     If ErrMustStop Then Debug.Assert False: Resume
59390     ErrorIn "a_InvoiceLine.ValidateObject(pFld)", pFld
End Function

Public Sub GetSTatus()
59400     On Error GoTo errHandler
59410     ValidateObject ""
59420     mobjValid.GetSTatus
59430     Exit Sub
errHandler:
59440     If ErrMustStop Then Debug.Assert False: Resume
59450     ErrorIn "a_InvoiceLine.GetSTatus"
End Sub

Public Function CoffsValid(pResult As Integer)
59460     On Error GoTo errHandler
      Dim lngTotalQtys As Long
      Dim lngTotalQtysAwaiting As Long
      Dim oCoff As a_COFF
59470     lngTotalQtys = 0
59480     For Each oCoff In COFFs
59490         lngTotalQtys = oCoff.COFFQTY + lngTotalQtys
59500         lngTotalQtysAwaiting = oCoff.COLQty - oCoff.COLQtyDispatched + lngTotalQtysAwaiting
59510     Next
59520     If lngTotalQtys = Qty Then
59530         pResult = 1    'We have allocated all we can
59540     ElseIf lngTotalQtys < Qty Then
59550         If lngTotalQtysAwaiting > lngTotalQtys Then
59560             pResult = 2  'We could allocate more
59570         Else
59580             pResult = 4 ' WE have tried to allocate more than requested
59590         End If
59600     Else
59610         pResult = 3   'We have tried to allocate more than we can
59620     End If
59630     Exit Function
errHandler:
59640     If ErrMustStop Then Debug.Assert False: Resume
59650     ErrorIn "a_InvoiceLine.CoffsValid(pResult)", pResult
End Function


Public Function SetForeignPrice(val As String) As Boolean
59660     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim lngTemp As Long
59670     lngTemp = mudtProps.ForeignPrice
59680     bOK = SetField_strAsCurrencyToLong(mudtProps.ForeignPrice, val, mcolStack.Count, "ForeignPrice", oPC.Configuration.CaptureDecimal, oPC.Configuration.DefaultCurrency.Divisor)
59690     If bOK And (lngTemp <> mudtProps.ForeignPrice) Then 'only validate if conversion is OK and the new value is other than the original
59700         bOK = ValidateObject("ForeignPrice")
59710     End If
59720     SetForeignPrice = bOK
59730     Exit Function
errHandler:
59740     If ErrMustStop Then Debug.Assert False: Resume
59750     ErrorIn "a_QUL.SetForeignPrice(val)", val
End Function
'Public Property Let ForeignPrice(val As Long)
'    On Error GoTo errHandler
'  If mcolStack.Count = 0 Then Err.Raise 383
'    If val <> mudtProps.ForeignPrice Then
'        mudtProps.ForeignPrice = val
'        SetDirty True
'    End If
'    ValidateObject "ForeignPrice"
'    Exit Property
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_QUL.ForeignPrice(val)", val
'End Property
Public Property Get ForeignPriceF() As String
59760     On Error GoTo errHandler
      Dim oFC As New a_Currency
59770     If mudtProps.FCID < 1 Then
59780         ForeignPriceF = ""
59790         Exit Sub
59800     End If
59810     Set oFC = oPC.Configuration.Currencies.FindCurrencyByID(mudtProps.FCID)
59820     ForeignPriceF = Format(CDbl(mudtProps.ForeignPrice / oFC.Divisor), oFC.FormatString)
59830     Exit Property
errHandler:
59840     If ErrMustStop Then Debug.Assert False: Resume
59850     ErrorIn "a_InvoiceLine.ForeignPriceF"
End Property

Public Function SetFCFactor(val As Double) As Boolean
59860     On Error GoTo errHandler
      Dim bOK As Boolean
      Dim Temp As Double
59870     Temp = mudtProps.FCFactor
59880     bOK = SetField_DOUBLE(mudtProps.FCFactor, CStr(val), "FCFactor", mcolStack.Count)
59890     If bOK And (Temp <> mudtProps.FCFactor) Then 'only validate if conversion is OK and the new value is other than the original
59900         bOK = ValidateObject("FCFactor")
59910     End If
59920     SetFCFactor = bOK
59930     Exit Function

59940     Exit Function
errHandler:
59950     If ErrMustStop Then Debug.Assert False: Resume
59960     ErrorIn "a_InvoiceLine.SetFCFactor(val)", val
End Function

Public Property Get FCFactor() As Double
59970     FCFactor = mudtProps.FCFactor
End Property
Public Property Get FCFactorF() As String
59980     FCFactorF = Format(FCFactor, "##0.####")
End Property
Public Property Get FCFactorInvF() As String
59990     If FCFactor <> 0 Then
60000         FCFactorInvF = Format(Round(1# / FCFactor, 3), "##0.###")
60010     End If
End Property
Public Property Get FCID() As Long
60020     FCID = mudtProps.FCID
End Property
Public Property Let FCID(val As Long)
60030     mudtProps.FCID = val
End Property





