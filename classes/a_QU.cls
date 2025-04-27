VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_QU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oSM As z_StockManager
Private oFSo As FileSystemObject

Dim mPay As Long
Dim lngFPLessDiscExt_tot As Long
Dim lngLDeposit_tot As Long
Dim lngFDeposit_tot As Long
Dim mServiceItem As Long

Dim mVAT As Long
Dim mDisc As Long
Dim mPay_Ex As Long
Dim lngFPayableLessVAT_tot As Long
Dim mPayExtras As Long
Dim lngQty_tot As Long
Dim mOriginalStatus As enStatus
Dim oPCL As z_PrintClient
Dim iExtrasCount As Integer
Private xMLDoc As ujXML
Event Valid(pMsg As String)
Event TotalChange(lngTotalExt As Long, lngDeposit_tot As Long, mVAT As Long)
Event reloadlist()
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Event RowCount(cnt As Integer)
Private mudtProps As QUProps
Private mcolStack As Collection
Private WithEvents mcolQUL As ch_QUL
Attribute mcolQUL.VB_VarHelpID = -1
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private colClassors As Collection
Private oBillToAddress As a_Address
Private oDelToAddress As a_Address
Private oBillingCompany As a_Company
Private oCurrency_Foreign As a_Currency
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Const CDOCCODE = "QU"


Dim arDocTotals() As TTotal2
Dim arDocTotalsF() As TTotal2
Dim RunningTotal As Long
Dim jTotal As Long

Public Property Get constDOCCODE() As String
    constDOCCODE = CDOCCODE
End Property

Private Sub mcolQUL_CurrRowStatus(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent CurrRowStatus(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.mcolQUL_CurrRowStatus(pMsg)", pMsg
End Sub
Public Property Get ForeignCurrency() As a_Currency
    Set ForeignCurrency = oCurrency_Foreign
End Property
Private Sub mcolQUL_RowsChange()
    On Error GoTo errHandler
'occurs for new rows and deletion of rows
    
    CalculateTotals
    RaiseEvent TotalChange(mPay, lngLDeposit_tot, mVAT)
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.mcolQUL_RowsChange"
End Sub
Private Sub mcolQUL_RowCOunt(cnt As Integer)
    RaiseEvent RowCount(cnt)
End Sub
Public Function RecalculateAllLines()
    On Error GoTo errHandler
Dim QUL As a_QUL
    For Each QUL In Me.QuoteLines
        QUL.CalculateLine
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.RecalculateAllLines"
End Function
Public Function CalculateTotals()
    On Error GoTo errHandler
    CalculateTotal_Local
    CalculateTotal_Foreign
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CalculateTotals"
End Function
Public Function CalculateTotal_Local()
    On Error GoTo errHandler
Dim QUL As a_QUL
Dim bLineDiscountApplied As Boolean
    mobjValid.RuleBroken "OVERFLOW", False
    bLineDiscountApplied = False
    mPay = 0
    lngLDeposit_tot = 0
    mVAT = 0
    mDisc = 0
    mPay_Ex = 0
    'NOTE Some items on the quotation might be ServiceItem e.g. Insurance, Postage and so on. These items are not subject to any discount
    'and they probably will not carry VAT. e.g. a handling charge would include VAT
    'VAT is calculated on the discounted price, so it is necessary to apply the discount before calculating VAT
    'Both forms of discount (by line or overall(mutually exclusive)) will need to be applied line by line, and the VAT for that line
    'calculated afterwards. Remember VAT rate can vary per line.
    'Any general discount will merely be distributed as a line by line discount on all non ServiceItem items
    'Any deposits are ignored until the last when the total deposit is subtracted from the total payable figure
    iExtrasCount = 0
    mPayExtras = 0
    lngQty_tot = 0
    mServiceItem = 0
    
    For Each QUL In mcolQUL
        If (Not QUL.IsDeleted) Then
            QUL.CalculateLine
            If QUL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            If QUL.ServiceItem = False Then
                lngQty_tot = lngQty_tot + QUL.Qty
            End If
            If QUL.ServiceItem = True Then
                mServiceItem = mServiceItem + QUL.Ext(False)
            End If
            
            mPay_Ex = mPay_Ex + QUL.Ext_Ex(False) + QUL.ExtraCharge
            mVAT = mVAT + QUL.VAT(False) + (QUL.ExtraCharge - (CDbl(QUL.ExtraCharge) * 100 / (100 + QUL.ExtraVATRate)))
            mDisc = mDisc + QUL.Disc(False)
            mPay = mPay + QUL.Ext(False) + QUL.ExtraCharge
            If QUL.BottomOfDocument Then
                mPayExtras = mPayExtras + QUL.Ext(False)
            End If
        End If
    Next
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtras = mPayExtras
    mudtProps.TotalDiscount = mDisc
    mudtProps.TotalNonVAT = mPay_Ex
    mudtProps.TotalVAT = mVAT
    mudtProps.TotalExtension = mPay
    mudtProps.TotalPayable = mPay - lngLDeposit_tot
    mudtProps.TotalServiceItem = mServiceItem
    If (Not VATable) And ShowVAT Then
        mudtProps.TotalPayable = mPay - mVAT
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
    ErrorIn "a_QU.CalculateTotal_Local"
End Function
Public Function CalculateTotal_Foreign()
    On Error GoTo errHandler
Dim QUL As a_QUL
Dim bLineDiscountApplied As Boolean
    mobjValid.RuleBroken "OVERFLOW", False
    bLineDiscountApplied = False
    lngFPLessDiscExt_tot = 0
    lngFDeposit_tot = 0
    mVAT = 0
    lngFPayableLessVAT_tot = 0
    mDisc = 0
    mPayExtras = 0
    lngQty_tot = 0
    mServiceItem = 0

    For Each QUL In mcolQUL
        If (Not QUL.IsDeleted) Then
            QUL.CalculateLine
            If QUL.BottomOfDocument Then
                iExtrasCount = iExtrasCount + 1
            End If
            If QUL.ServiceItem = False Then
                lngQty_tot = lngQty_tot + QUL.Qty
            End If
            If QUL.ServiceItem = True Then
                mServiceItem = mServiceItem + QUL.Ext_Ex(False)
            End If
            mPay_Ex = mPay_Ex + QUL.Ext_Ex(True) + QUL.ExtraCharge
            
            lngFPayableLessVAT_tot = lngFPayableLessVAT_tot + QUL.Ext_Ex(True)
            mVAT = mVAT + QUL.VAT(True) + (QUL.ExtraCharge - (CDbl(QUL.ExtraCharge) * 100 / (100 + QUL.ExtraVATRate)))
            mDisc = mDisc + QUL.Disc(True)
            mPay = mPay + QUL.Ext(True) + QUL.ExtraCharge
            If QUL.BottomOfDocument Then
                mPayExtras = mPayExtras + QUL.Ext(True)
            End If
        End If
    Next
    mudtProps.TotalQty = lngQty_tot
    mudtProps.TotalExtras_Foreign = mPayExtras
    mudtProps.TotalDiscount_Foreign = mDisc
    mudtProps.TotalVAT_Foreign = mVAT
    mudtProps.TotalNonVAT_Foreign = mPay_Ex
    mudtProps.TotalExtension_Foreign = mPay
    mudtProps.TotalPayable_Foreign = mPay - lngLDeposit_tot
    mudtProps.TotalServiceItem_Foreign = mServiceItem
    
    If (Not Me.VATable) And ShowVAT Then
        mudtProps.TotalPayable_Foreign = lngFPLessDiscExt_tot - mVAT
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
    ErrorIn "a_QU.CalculateTotal_Foreign"
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
    ErrorIn "a_QU.TotalExtWithOutExtrasF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalExtWithOutExtras(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalExtWithOutExtrasF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalLessDiscExtF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalLessDiscExt(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalNonVATF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalVATF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalPayableRaw(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalVATRaw(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalPayable(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalPayableF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalDiscount(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalDiscountF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalDepositF(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalDeposit(bForeign)", bForeign
End Property
Public Property Get DiscountRate() As Double
    On Error GoTo errHandler
    DiscountRate = mudtProps.DiscountRate
  '  RecalculateAllLines
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DiscountRate"
End Property
Public Property Get DiscountRateF() As String
    On Error GoTo errHandler
    If mudtProps.DiscountRate = 0 Then
        DiscountRateF = ""
    Else
        DiscountRateF = Format(mudtProps.DiscountRate, gPercentFormatString)
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DiscountRateF"
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
    ErrorIn "a_QU.TotalBalanceAfterDeposits(bForeign)", bForeign
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As QUData
  
  LSet udtData = mudtProps
  GetState = udtData.Buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.GetState"
End Function
Private Sub SetState(Buffer As String)
    On Error GoTo errHandler
  Dim udtData As QUData
  
  udtData.Buffer = Buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.SetState(buffer)", Buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "QULS", mcolQUL.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.GetSuperState"
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
    mcolQUL.SetSuperState .ReadProperty("QULS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.SetSuperState(buffer)", Buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
  mcolQUL.BeginEdit
  mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  mcolQUL.CancelEdit
  mudtProps.IsDeleted = False
  With mcolStack
    SetState .Item(.Count)
    .Remove .Count
  End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CancelEdit"
End Sub

Public Function ApplyEdit(Optional pMsg As String) As String
    On Error GoTo errHandler
Dim strTemp As String
Dim objPersist As a_QU_P
Dim oSM As z_StockManager
Dim OpenResult As Integer

'MsgBox "in applyedit"
    ApplyEdit = ""
    CalculateTotals
    mudtProps.TPID = Customer.ID
    mudtProps.SalesRepID = Customer.repid
    
    'Set these values fromthe objects before saving -  the objects cant be saved
    If Not oBillToAddress Is Nothing Then
        mudtProps.BillToAddressID = oBillToAddress.ID
    End If
    If Not oDelToAddress Is Nothing Then
        mudtProps.DelToAddressID = oDelToAddress.ID
    End If
    
'''''''''''''''''''''''''''
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
  '  oPC.COShort.execute "BEGIN TRANSACTION"
    If Me.DOCDate = CDate(0) Then
        Me.DOCDate = Date
        Me.CaptureDate = Now
    End If
    If DOCCode = "" Then
        DOCCode = GetQuoteNumber("GETANDWRITE")
    End If
    
    If mcolStack.Count = 0 Then Err.Raise 445
    oPC.COShort.CommandTimeout = 55
    mcolQUL.ApplyEdit
    Set objPersist = New a_QU_P
    
    If mudtProps.IsDeleted Then
        objPersist.DeleteObject mudtProps.QuoteID
        mcolStack.Remove mcolStack.Count
        mudtProps.IsNew = True
        mudtProps.IsDeleted = False
    ElseIf IsDirty Or mudtProps.IsNew Then
        If Not IsValid Then Err.Raise 445
        strTemp = objPersist.Save(GetSuperState)
    
        If UCase(Left(strTemp, 7)) <> "TIMEOUT" Then
            mcolStack.Remove mcolStack.Count
            SetSuperState strTemp
            mudtProps.IsNew = False
        Else 'save failed
            mcolQUL.BeginEdit
            If Not IsMissing(pMsg) Then
                pMsg = strTemp
            End If
            ApplyEdit = "TIMEOUT"
            Set oSM = New z_StockManager
            oSM.RollbackTransaction oPC.COShort
            Set oSM = Nothing
            Set objPersist = Nothing
        LogSaveToFile "a_QU:TimeoutExpired"
            GoTo EXIT_Handler
        End If
    
    Else
        mcolStack.Remove mcolStack.Count
    End If
    Set objPersist = Nothing
    SetDirty False
   ' oPC.COShort.execute "COMMIT TRANSACTION"
    '--------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
    '--------------
    
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
   ' oSM.RollbackTransaction oPC.COShort
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.ApplyEdit", , , oPC.COShort
End Function
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
  IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
  IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
  IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
  IsDirty = mudtProps.IsDirty Or mcolQUL.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.IsDirty"
End Property
Public Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.SetDirty(pVal)", pVal
End Sub
Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.IsValid"
End Property
Public Function GetSTatus()
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
    ErrorIn "a_QU.GetSTatus"
End Function
Private Sub Class_Initialize()
Dim oComp As a_Company
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mcolQUL = New ch_QUL
    mcolQUL.component Me
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
    mudtProps.STATUS = stInProcess
    mOriginalStatus = stInProcess
    Set oBillToAddress = New a_Address
    Set oDelToAddress = New a_Address
    mobjValid.BreakRule "TP", True
    LoadClassorsCollection
    Set oCurrency_Foreign = oPC.Configuration.DefaultCurrency
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
    If mcolStack.Count > 0 Then _
      Err.Raise vbObjectError + 1001, , "State stack is not empty"
      
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mcolQUL = Nothing
    Set mCustomer = Nothing
    Set oPCL = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(ors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(ors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.mobjValid_BrokenRule(ors)", ors
End Sub
Private Sub mobjValid_RuleUnbroken(ors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(ors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.mobjValid_RuleUnbroken(ors)", ors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(TRID As Long, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_QU_P
Dim oADd As a_Address
Dim oComp As a_Company
Dim oCurr As a_Currency
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_QU_P
    SetSuperState objPersist.Fetch(TRID)
    mOriginalStatus = Me.STATUS
    mCustomer.Load mudtProps.TPID
    Set oBillToAddress = Nothing
    Set oDelToAddress = Nothing
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
    CalculateTotals
    mudtProps.IsNew = False
    mudtProps.IsDirty = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Load(TRID,ReadOnly)", Array(TRID, ReadOnly)
End Sub
Public Property Get Customer() As a_Customer
    On Error GoTo errHandler
    Set Customer = mCustomer
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Customer"
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Delete"
End Sub



Public Property Get QuoteID() As Long
    On Error GoTo errHandler
    QuoteID = mudtProps.QuoteID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.QuoteID"
End Property
Public Property Let QuoteID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.QuoteID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.QuoteID(val)", val
End Property

Public Property Get BillToAddressID() As Long
    On Error GoTo errHandler
    BillToAddressID = mudtProps.BillToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.BillToAddressID:Get"
End Property
Public Property Let BillToAddressID(val As Long)
    On Error GoTo errHandler
    mudtProps.BillToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.BillToAddressID:Let"
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
    ErrorIn "a_QU.SetBillToAddress(val)", val
End Sub
Public Property Get DelToAddressID() As Long
    On Error GoTo errHandler
    DelToAddressID = mudtProps.DelToAddressID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DelToAddressID:Get"
End Property
Public Property Let DelToAddressID(val As Long)
    On Error GoTo errHandler
   mudtProps.DelToAddressID = val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DelToAddressID:Let"
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
    ErrorIn "a_QU.setDelToAddress(val)", val
End Sub
'Public Function SetGeneralDiscount(val As String) As Boolean
'    On Error GoTo ErrHandler
'Dim dblDiscountPercent As Double
'Dim oQUL As a_QUL
'
'    If mcolStack.Count = 0 Then Err.Raise 383
'    SetGeneralDiscount = True
'    If Trim$(val) = "" Then
'        dblDiscountPercent = 0
'    ElseIf Not ConvertToDBL(val, dblDiscountPercent) Then
'        SetGeneralDiscount = False
'        Exit Function
'    End If
'    If (dblDiscountPercent <= 0) Then
'        mobjValid.RuleBroken "DISCOUNT", True
'    Else
'        mobjValid.RuleBroken "DISCOUNT", False
'    End If
'    mudtProps.DiscountRate = dblDiscountPercent
'    SetDirty True
'    For Each oQUL In QuoteLines  'disctribute discount rate over all lines where non nonstaock items exists
'        If Not oQUL.ServiceItem Then
'            oQUL.DiscountPercent = dblDiscountPercent
'        End If
'    Next
'    mcolQUL_RowsChange  'recalulates andd raises event
'    RecalculateAllLines
'    RaiseEvent reloadlist
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_QU.SetGeneralDiscount(val)", val
'End Function

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
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.SetCustomer(pTPID)", pTPID
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
    ErrorIn "a_QU.SetCustomerFromAccNum(pAccnum)", pAccnum
End Function
Public Property Let ShowVAT(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowVAT = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.ShowVAT(val)", val
End Property
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.ShowVAT"
End Property
Public Property Get StaffID() As Long
    On Error GoTo errHandler
    StaffID = mudtProps.StaffID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.StaffID"
End Property
Public Property Let StaffID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.StaffID(val)", val
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
    ErrorIn "a_QU.StaffName"
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
    ErrorIn "a_QU.StaffNameB"
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
    ErrorIn "a_QU.CurrencyID_Foreign"
End Property
Public Property Let CurrencyID(val As Long)
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim QUL As a_QUL

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
    ErrorIn "a_QU.CurrencyID(val)", val
End Property
Public Property Get COMPID() As Long
    On Error GoTo errHandler
    COMPID = mudtProps.COMPID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.COMPID"
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
    ErrorIn "a_QU.COMPID(val)", val
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
    ErrorIn "a_QU.TPName"
End Property
Public Property Let TPName(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPName = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPName(val)", val
End Property

Public Property Get TPPhone() As String
    On Error GoTo errHandler
    TPPhone = FNS(mudtProps.TPPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPPhone"
End Property
Public Property Let TPPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPPhone(val)", val
End Property

Public Property Get TPFax() As String
    On Error GoTo errHandler
    TPFax = FNS(mudtProps.TPFax)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPFax"
End Property
Public Property Let TPFax(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPFax = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPFax(val)", val
End Property

Public Property Get TPACCNum() As String
    On Error GoTo errHandler
    TPACCNum = FNS(mudtProps.TPACCNum)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPACCNum"
End Property
Public Property Let TPACCNum(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.TPACCNum = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TPACCNum(val)", val
End Property

Public Property Get BusPhone() As String
    On Error GoTo errHandler
    BusPhone = FNS(mudtProps.BusPhone)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.BusPhone"
End Property
Public Property Let BusPhone(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.BusPhone = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.BusPhone(val)", val
End Property

Public Property Get DOCCode() As String
    On Error GoTo errHandler
    DOCCode = FNS(mudtProps.DOCCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DocCode"
End Property
Public Property Get DOCCodeF() As String
    On Error GoTo errHandler
    Select Case mudtProps.STATUS
    Case 1
        DOCCodeF = "Voided " & FNS(mudtProps.DOCCode)
    Case 2
        DOCCodeF = "Quotation in process " & FNS(mudtProps.DOCCode)
    Case 3
        DOCCodeF = "Quotation " & FNS(mudtProps.DOCCode)
    Case 4
        DOCCodeF = "Quotation " & FNS(mudtProps.DOCCode)
    End Select
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DOCCodeF"
End Property
Public Property Let DOCCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.DOCCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DocCode(val)", val
End Property

Public Property Get CurrencyFormat() As String
    On Error GoTo errHandler
    CurrencyFormat = FNS(mudtProps.CurrencyFormat)
    If CurrencyFormat <= "" Then CurrencyFormat = oPC.Configuration.Currencies(oPC.Configuration.DefaultCurrencyID & "k").FormatString
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CurrencyFormat"
End Property
Public Property Let CurrencyFormat(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CurrencyFormat = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CurrencyFormat(val)", val
End Property

Public Property Get DOCDate() As Date
    On Error GoTo errHandler
    DOCDate = CDate(ReverseDate(mudtProps.DOCDate))  'to ensure no time included
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DocDate"
End Property
Public Property Get DocDateF() As String
    On Error GoTo errHandler
    DocDateF = Format(mudtProps.DOCDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DocDateF"
End Property
Public Property Get TransDateF2() As String
    On Error GoTo errHandler
    TransDateF2 = Format(mudtProps.DOCDate, "dd mmmm yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TransDateF2"
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
    ErrorIn "a_QU.DocDate(val)", val
End Property

Public Property Get CaptureDate() As Date
    On Error GoTo errHandler
    CaptureDate = mudtProps.CaptureDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CaptureDate"
End Property
Public Property Get CaptureDateF() As Date
    On Error GoTo errHandler
    CaptureDateF = Format(mudtProps.CaptureDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CaptureDateF"
End Property
Public Property Let CaptureDate(val As Date)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.CaptureDate = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CaptureDate(val)", val
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
    ErrorIn "a_QU.CurrencyFactor"
End Property
Public Property Let CurrencyFactor(val As Double)
    On Error GoTo errHandler
    mudtProps.CurrencyFactor = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CurrencyFactor(val)", val
End Property

Public Property Get VATRate() As Double
    On Error GoTo errHandler
    VATRate = mudtProps.VATRate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.VATRate"
End Property
Public Property Get VATRateFormatted() As String
    On Error GoTo errHandler
    VATRate = Format(mudtProps.VATRate, "##0.00")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.VATRateFormatted"
End Property
Public Property Let VATRate(val As Double)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.VATRate = val
    mudtProps.IsDirty = True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.VATRate(val)", val
End Property
Public Property Get QtyNonStandardVAT() As Long
Dim i As Long
Dim QUL As a_QUL

    i = 0
    For Each QUL In Me.QuoteLines
        If QUL.VATRate <> oPC.Configuration.VATRate Then
            i = i + 1
        End If
    Next
    QtyNonStandardVAT = i
End Property

Public Property Get VATable() As Boolean
    On Error GoTo errHandler
  VATable = Me.Customer.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.VATable(val)", val
End Property


Public Property Get QuoteLines() As ch_QUL
  Set QuoteLines = mcolQUL
End Property
Public Property Get DisplayQuote() As d_QU
    On Error GoTo errHandler
    Set DisplayQuote = New d_QU
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.DisplayQuote"
End Property
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
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusForPrinting = "CANCELLED"
    End If
    
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.StatusForPrinting"
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
    ElseIf mudtProps.STATUS = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.StatusF"
End Property
Public Property Get STATUS() As enStatus
    On Error GoTo errHandler
    STATUS = mudtProps.STATUS
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Status"
End Property
Public Function SetOrderNumber(val As String) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If Len(val) > Len(mudtProps.Ref) Then
        Err.Raise 384
    End If
    If val <> mudtProps.Ref Then
        mudtProps.Ref = val
        SetDirty True
    End If
    SetOrderNumber = True
    Exit Function
End Function
Public Function SetOrderDate(val As Date) As Boolean
    If mcolStack.Count = 0 Then Err.Raise 383
    If val <> mudtProps.OrderDate Then
        mudtProps.OrderDate = val
        SetDirty True
    End If
    SetOrderDate = True
    Exit Function
End Function
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
  '  If mcolStack.Count = 0 Then Err.Raise 383
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
    oSM.SetMemo val, Me.QuoteID
    
End Function
Public Property Get Memo() As String
    On Error GoTo errHandler
    Memo = FNS(mudtProps.Memo)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Memo"
End Property
Public Property Get ForAttn() As String
    On Error GoTo errHandler
    ForAttn = FNS(mudtProps.ForAttn)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.ForAttn"
End Property
Public Property Get OrderNumber() As String
    On Error GoTo errHandler
    OrderNumber = FNS(mudtProps.Ref)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.OrderNumber"
End Property
Public Property Get OrderDate() As Date
    On Error GoTo errHandler
    OrderDate = FND(mudtProps.OrderDate)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.OrderDate"
End Property
Public Property Get OrderDateF() As String
    On Error GoTo errHandler
    OrderDateF = Format(mudtProps.OrderDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.OrderDateF"
End Property
Public Function GetQuoteNumber(strTmp As String) As String
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim strDefaultCo As String

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_GetNextQuoteNo"
    cmd.commandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
 '  Set prm = cmd.CreateParameter("@COMPID", adInteger, adParamInput, , COMPID)
 '   cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@NewCode", adVarChar, adParamOutput, 15)
    cmd.Parameters.Append prm
    cmd.execute
    GetQuoteNumber = cmd.Parameters("@NewCode").Value
    Set cmd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.GetQuoteNumber(strTmp)", strTmp
End Function

Property Get BillTOAddress() As a_Address
    Set BillTOAddress = oBillToAddress
End Property
Property Get DelToAddress() As a_Address
    Set DelToAddress = oDelToAddress
End Property

Public Function DisplayTotals(pLabels As String, pValues As String, bForeign As Boolean) 'This must be the same as PrintQuote
    On Error GoTo errHandler
Dim strLabels As String
Dim strValues As String
Dim i As Integer
    strLabels = ""
    strValues = ""
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strLabels = "subtotal"
        strValues = Me.TotalExtWithOutExtrasF(bForeign)
        For i = 1 To Me.QuoteLines.Count
            If Me.QuoteLines(i).BottomOfDocument = True Then
                With Me.QuoteLines(i)
                    If .ServiceItem Then
                        strLabels = strLabels & Chr(13) & .TitleAuthorPublisher
                        strValues = strValues & Chr(13) & .ExtF(bForeign)
                    End If
                End With
            End If
        Next i
    End If
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
    ErrorIn "a_QU.DisplayTotals(pLabels,pValues,bForeign)", Array(pLabels, pValues, bForeign)
End Function
'Private Function FillTableRows(bBySeqNo As Boolean, bForeign As Boolean, pDiscountExists As Boolean)
'    On Error GoTo ErrHandler
'Dim i As Long
'Dim iExtrasCount As Integer
'Dim QUL As a_QUL
'
'    iExtrasCount = 0
'    For i = 1 To Me.QuoteLines.Count
'  '     If bBySeqNo Then
'  '          Set QUL = Me.QuoteLines.FindLineBySeqNum(i)
'  '      Else
'            Set QUL = Me.QuoteLines.Item(i)
'  '      End If
'        If Not QUL Is Nothing Then
'            If QUL.BottomOfDocument = False Then
'                With QUL
'                    If .DiscountPercent > 0 Then pDiscountExists = True
'                    If .ServiceItem Then
'                        oPCL.ILSend "", "", "", .Title & .SpecialVATMsg, "", "", "", .PLessDiscExtF(bForeign), "", "", "", .Note
'                    Else
'                        oPCL.ILSend .CodeForExport, .Ref, .QtyF, .SpecialVATMsg & .Title, .MainAuthor, .Publisher, .PriceF(bForeign), .PLessDiscExtF(bForeign), .DepositF(bForeign), .DiscountPercentF, .PLessDiscExtNonVATF(bForeign), PackText(.Note)
'                    End If
'                End With
'            Else
'                iExtrasCount = iExtrasCount + 1
'            End If
'        End If
'    Next i
'    FillTableRows = iExtrasCount
'EXIT_Handler:
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_QU.FillTableRows(bBySeqNo,bForeign,pDiscountExists)", Array(bBySeqNo, bForeign, _
'         pDiscountExists)
'End Function
'Private Function FillExtras(pDoc As a_QU, bForeign As Boolean, pDiscountExists As Boolean) As Integer
'    On Error GoTo ErrHandler
'Dim i As Integer
'Dim iCnt As Integer
'
'    iCnt = 0
'    For i = 1 To pDoc.QuoteLines.Count
'        If pDoc.QuoteLines(i).BottomOfDocument = True Then
'            With pDoc.QuoteLines(i)
'                If .ServiceItem Then
'                    iCnt = iCnt + 1
'                    oPCL.ILSendExtra .TitleAuthorPublisher, .PLessDiscExtF(bForeign)
'                End If
'            End With
'        End If
'    Next i
'    FillExtras = iCnt
'EXIT_Handler:
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_QU.FillExtras(pDoc,bForeign,pDiscountExists)", Array(pDoc, bForeign, _
'         pDiscountExists)
'End Function



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
    ErrorIn "a_QU.TranslateErrors(pRawors)", pRawors
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
    ErrorIn "a_QU.LoadClassorsCollection"
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
    ErrorIn "a_QU.ValidateObject(pVal,pFld)", Array(pVal, pFld)
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
    Me.SetStatus stISSUED
    strResult = Me.ApplyEdit
    If strResult > "" Then  'probably a timeout or
       'Not necessary oPC.COShort.execute "ROLLBACK TRANSACTION"
        Post = strResult
        Exit Function
    End If
    
    If pNewStatus = 0 Then
        pNewStatus = Me.STATUS
    End If
    If pNewStatus > 0 Then SetStatus pNewStatus
    Post = ""
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXITH:
    Set oSM = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    RlsObjs oPC.COShort
    ErrorIn "a_QU.Post", , EA_DFTRBK, oPC.COShort
End Function
Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    mudtProps.STATUS = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.SetStatus(val)", val
End Sub
Public Function CancelDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stCANCELLED
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.CancelDocument"
End Function

Public Function VoidDocument()
    On Error GoTo errHandler
Dim strResult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.VoidDocument"
End Function
Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As Long
    lngTRID = Me.QuoteID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.Reload"
End Sub

Public Sub SetProforma()
    mudtProps.Proforma = True
End Sub

Public Property Get SendersEmail()
    SendersEmail = FNS(mudtProps.StaffEmail)
End Property

Public Function ExportToXML(bBySeqNo As Boolean, pFilename As String, Optional ByVal pRoundedUp As Boolean, _
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
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim bReferenceExists As Boolean
Dim bDiscountExists As Boolean
Dim bDepositExists As Boolean
Dim bQtyColumnNeeded As Boolean
Dim oCHKIL As a_QUL
Dim oDC As a_DocumentControl
Dim strWorkingFolder As String
Dim strDispatchMethod As String
Dim strTemplateName As String
Dim strExecutable As String
Dim lngNumberOfLines As Long

    If DispatchMode = enMail Then
        If oPC.UsesOutlookForINVEmail Then
            strWorkingFolder = oPC.LocalFolder & "Emails\"
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
    
                        p 1
    bForeign = False
    If oCurrency_Foreign.ID <> oPC.Configuration.DefaultCurrency.ID Then
        bForeign = True
        CalculateTotal_Foreign
    End If
                        p 2
    bDiscountExists = False
    bReferenceExists = False
    bDepositExists = False
    bQtyColumnNeeded = False
    For Each oCHKIL In QuoteLines
        If oCHKIL.Ref > " " Then
            bReferenceExists = True
        End If
        If oCHKIL.DiscountPercent > 0 Then
            bDiscountExists = True
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
        .docInit "QU"
            .chCreate "MessageType"
                .elText = "QUOTATION"
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
                If Me.Customer.QuotationTemplateName > "" Then
                    .elText = Me.Customer.QuotationTemplateName
                Else
                    .elText = "QU_DOC"
                End If
            strTemplateName = .elText
            
            .elCreateSibling "Columns"
            If bDiscountExists And bReferenceExists Then
                    .elText = "DISC_REF"
            ElseIf bReferenceExists Then
                    .elText = "REF"
            ElseIf bDiscountExists Then
                    .elText = "DISC"
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
            .elCreateSibling "Status"
                .elText = StatusForPrinting
            .elCreateSibling "AccompanyingMessage"
                .elText = oPC.Configuration.EmailQuoteMsg
            .elCreateSibling "LogoPath"
                If fs.FileExists(oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP") Then
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.BMP"
                Else
                    .elText = oPC.SharedFolderRoot & "\TEMPLATES\LOGO.JPG"
                End If
            .elCreateSibling "DocCode"
                .elText = DOCCodeF
           .elCreateSibling "StoreCode", True
                .elText = oPC.Configuration.DefaultStore.code
           .elCreateSibling "DocDate", True
                .elText = DocDateF
            .elCreateSibling "Sender", True
                .elText = oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName                  'oPC.Configuration.DefaultCompany.CompanyName
                                p 6
            .elCreateSibling "SenderAddress", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
            .elCreateSibling "CustomerName", True
                .elText = Customer.Name
            
            .elCreateSibling "CustomerOrderNum", True
                .elText = Me.OrderNumber
            .elCreateSibling "CustomerOrderDate", True
                .elText = Me.OrderDateF
            .elCreateSibling "OurAcnowithCustomer", True
                .elText = Customer.OurACnoWithClient
            
            
            .elCreateSibling "CustomerVATNumber", True
            If Customer.VatNumber > "" Then
                .elText = Customer.VatNumber
            Else
                .elText = ""
            End If
            .elCreateSibling "ACNO", True
                .elText = Customer.AcNo
            .elCreateSibling "BankingDetails", True
                .elText = Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).BankDetails, Chr(13) & Chr(10), Chr(10))
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
                .elText = IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
            .elCreateSibling "BillTo", True
                If Me.BillTOAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
                    If Me.ForAttn > "" Then .elText = Me.ForAttn & vbCrLf & .elText
                End If
            .elCreateSibling "DelTo", True
                If Me.DelToAddress Is Nothing Then
                    .elText = ""
                Else
                    .elText = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
                            p 8
            lngNumberOfLines = 0
            For i = 1 To QuoteLines.Count
                If QuoteLines.Item(i).BottomOfDocument = False Then
                    lngNumberOfLines = lngNumberOfLines + 1
                    .elCreateSibling "DetailLine", True
                    .chCreate "SKU"
                    If oPC.HideLocalSKUonINV Then
                        .elText = QuoteLines.Item(i).CodeForExport
                    Else
                        .elText = QuoteLines.Item(i).CodeF
                    End If
                    .elCreateSibling "Title", True
                        .elText = QuoteLines.Item(i).TitleAuthor & IIf(QuoteLines(i).ServiceItem, " (" & QuoteLines(i).Qty & " item" & IIf(QuoteLines(i).Qty = 1, ")", "s)"), "")
                    .elCreateSibling "QtyFirm", True
                        .elText = QuoteLines.Item(i).QtyF
                    .elCreateSibling "Price", True
                        .elText = QuoteLines.Item(i).PriceF(bForeign)
                    .elCreateSibling "PriceExVat", True
                        .elText = QuoteLines(i).PriceExVatF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = QuoteLines.Item(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = QuoteLines.Item(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = QuoteLines.Item(i).ExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = QuoteLines.Item(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = QuoteLines.Item(i).Ext_ExFF(bForeign)
                    .elCreateSibling "VAT", True
                        .elText = QuoteLines.Item(i).VATF(bForeign)
                    .elCreateSibling "VATRate", True
                        .elText = QuoteLines.Item(i).VATRateF

                    .elCreateSibling "DiscountValue", True
                        .elText = QuoteLines(i).DiscF(bForeign)
                    .elCreateSibling "FCName", True
                        If QuoteLines(i).FCID > 0 Then
                            .elText = oPC.Configuration.Currencies.FindCurrencyByID(QuoteLines(i).FCID).Description
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCPrice", True
                        If QuoteLines(i).FCID > 0 Then
                            .elText = QuoteLines(i).ForeignPriceF
                        Else
                            .elText = "n/a"
                        End If
                    .elCreateSibling "FCFactor", True
                        If QuoteLines(i).FCID > 0 Then
                            .elText = QuoteLines(i).FCFactorInvF
                        Else
                            .elText = "n/a"
                        End If
                    .navUP
                End If
                    If QuoteLines(i).ExtraCharge > 0 Then
                    .elCreateSibling "DetailLine", True
                        .chCreate "SKU"
                            .elText = ""
                        .elCreateSibling "Title", True
                           .elText = QuoteLines(i).ExtraChargeDescription & " (" & QuoteLines.Item(i).CodeF & ")"
                        .elCreateSibling "QtyFirm", True
                            .elText = ""
                        .elCreateSibling "Price", True
                            .elText = QuoteLines(i).ExtraChargeF
                        .elCreateSibling "PriceExVAT", True
                            .elText = QuoteLines(i).PriceExVatF(bForeign)
                        .elCreateSibling "DiscountRate", True
                            .elText = ""
                        .elCreateSibling "Reference", True
                            .elText = ""
                        .elCreateSibling "Extension", True
                            .elText = QuoteLines(i).ExtraChargeF
                        .elCreateSibling "Note", True
                            .elText = ""
                        .elCreateSibling "ExtensionExVAT", True
                            .elText = QuoteLines(i).ExtraChargeF
                        .elCreateSibling "VAT", True
                            .elText = QuoteLines(i).ExtraVATF(False)
                        .elCreateSibling "VATRate", True
                            .elText = QuoteLines(i).ExtraVATRateF
                        .navUP
                    End If
            Next i
                                p 9
            For i = 1 To QuoteLines.Count
                If QuoteLines.Item(i).BottomOfDocument = True Then
                    .elCreateSibling "BottomItems", True
                    .chCreate "SKU"
                        .elText = QuoteLines.Item(i).CodeF
                    .elCreateSibling "Title", True
                        .elText = QuoteLines.Item(i).TitleAuthorPublisher
                    .elCreateSibling "QtyFirm", True
                        .elText = QuoteLines.Item(i).QtyF
                    .elCreateSibling "Price", True
                        .elText = QuoteLines.Item(i).PriceF(bForeign)
                    .elCreateSibling "PriceExVAT", True
                        .elText = QuoteLines(i).PriceExVatF(bForeign)
                    .elCreateSibling "DiscountRate", True
                        .elText = QuoteLines.Item(i).DiscountPercentF
                    .elCreateSibling "Reference", True
                        .elText = QuoteLines.Item(i).Ref
                    .elCreateSibling "Extension", True
                        .elText = QuoteLines.Item(i).ExtF(bForeign)
                    .elCreateSibling "Note", True
                        .elText = QuoteLines.Item(i).Note
                    .elCreateSibling "ExtensionExVAT", True
                        .elText = QuoteLines.Item(i).Ext_ExFF(bForeign)
                    .elCreateSibling "VAT", True
                        .elText = QuoteLines.Item(i).VATF(bForeign)
                    .elCreateSibling "VATRate", True
                        .elText = QuoteLines.Item(i).VATRateF
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
                .elText = IIf(Me.StaffName > "", Me.StaffName, "")
            .elCreateSibling "OrderMessage", True
                .elText = oPC.Configuration.QuotationText
                
            .elCreateSibling "TotalShipping", True
                .elText = Me.TotalServiceItemF(False)
            .elCreateSibling "TotalDiscount", True
                .elText = Me.TotalDiscountF(False)
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
                                p 12

'FINALLY PRODUCE THE .XML FILE
    strXML = strWorkingFolder & "QU_" & Me.DOCCode & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

                                p 13

    If DispatchMode = enMail Then
        If oPC.EmailPOShowHTML = True Then
            'WRITE THE .HTML FILE
            objXSL.async = False
            objXSL.ValidateOnParse = False
            objXSL.resolveExternals = False
            strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_HTML.XSLT"
            Set fs = New FileSystemObject
            If fs.FileExists(strPath) Then
                objXSL.Load strPath
            End If
            If fs.FileExists(strWorkingFolder & "QU_" & Me.DOCCode & ".HTML") Then
                fs.DeleteFile strWorkingFolder & "QU_" & Me.DOCCode & ".HTML", True
            End If
            oTF.OpenTextFileToAppend strWorkingFolder & "QU_" & Me.DOCCode & ".HTML"
            pWholeMessage = xMLDoc.docObject.transformNode(objXSL)
            oTF.WriteToTextFile pWholeMessage
            oTF.CloseTextFile
        End If
    End If
    
    
'WRITE THE .PDF FILE IF NECESSARY
'Stage 1 apply the .XSLT style sheet to the .XML and produce the .FO file
    If (Me.Customer.DispatchMethod = "M" And DispatchMode = enMail) Or DispatchMode = enView Or (DispatchMode = enPrint And QuickPrint = True) Then
        Set objXSL = Nothing
        Set objXSL = New MSXML2.DOMDocument60
        objXSL.async = False
        objXSL.ValidateOnParse = False
        objXSL.resolveExternals = False
        strPath = oPC.SharedFolderRoot & "\Templates\" & strTemplateName & "_FO.XSL"
        Set fs = New FileSystemObject
        If fs.FileExists(strPath) Then
            objXSL.Load strPath
        End If
                        p 6
        Set opXMLDOC = New MSXML2.DOMDocument60
        opXMLDOC.async = False
        opXMLDOC.ValidateOnParse = False
        opXMLDOC.resolveExternals = False
        xMLDoc.docObject.transformNodeToObject objXSL, opXMLDOC
        
        strFOFile = strWorkingFolder & "QU_" & Me.DOCCode & ".FO"
        strPDFFile = strWorkingFolder & "QU_" & Me.DOCCode & ".PDF"
                        p 7
        docWriteTostream strFOFile, opXMLDOC, "UNICODE"
        
'Stage 2 Convert the .FO file to .PDF and clean up
        ChDir "\PBKS\Executables\FOP\"
       ' strCommand = "FOP.BAT" & " " & strFOFile & " " & strPDFFile
        'ShellandWait strCommand, 100
        strCommand = GetFOPCommandstring(strFOFile, strPDFFile)
        F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 600000

        If fs.FileExists(strFOFile) Then
            fs.DeleteFile strFOFile
        End If
    End If
    
'Send to Outlook
   ' CreateOutlookMail strXML
    
'Log the dispatch
    Set oSM = New z_StockManager
    Set oFSo = New FileSystemObject
    pFilename = strPDFFile
    
    If DispatchMode <> enMail And DispatchMode <> enView Then
        oSM.LogTransmission Me.QuoteID, "Dispatched: by " & strDispatchMethod & " : " & Format(Date, "dd/mm/yyyy") & vbCrLf
        Log = "Dispatched: by " & strDispatchMethod & "  created: " & Format(Date, "dd/mm/yyyy") & vbCrLf & Log
    End If
    Set oSM = Nothing
    Set oFSo = Nothing
                        p 9
    ExportToXML = True

    '
'    If DispatchMode = enView Then
'        strExecutable = GetPDFExecutable(strPDFFile)
'                If strExecutable = "" Then
'                    MsgBox "There is no application set on this computer to open the file: " & strPDFFile & ". The document cannot be displayed", vbOKOnly, "Can't do this"
'                Else
'                    Shell strExecutable & " " & strPDFFile, vbNormalFocus
'                End If
'    End If
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
    ErrorIn "a_QU.ExportToXML(bForeign)", bForeign, , , "strPos", Array(strErrPos)
End Function
Public Function ExportToSpreadsheet(bForeign As Boolean, _
        pFilename As String) As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim s As String
Dim s2 As String
Dim lngNumberOfLines As Long

    ExportToSpreadsheet = False
    If bForeign = True Then
        CalculateTotal_Foreign
    End If
    pFilename = oPC.LocalFolder & "Temp\" & DOCCode & "_" & Format(Now(), "yyyymmddHHnn") & ".xls"
    
    oTF.OpenTextFile pFilename
    oTF.WriteToTextFile "Status" & vbTab & StatusF
    oTF.WriteToTextFile "DocCode" & vbTab & DOCCodeF
    oTF.WriteToTextFile "DocDate" & vbTab & DocDateF
    oTF.WriteToTextFile "Sender" & vbTab & oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).CompanyName
    oTF.WriteToTextFile "SenderAddress" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).StreetAddress, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile "CustomerName" & vbTab & Customer.Name
    oTF.WriteToTextFile "CustomerOrderNum" & vbTab & OrderNumber
    oTF.WriteToTextFile "CustomerOrderDate" & vbTab & OrderDateF
    oTF.WriteToTextFile "OurAcnowithCustomer" & vbTab & Customer.OurACnoWithClient
    oTF.WriteToTextFile "ACNO" & vbTab & Customer.AcNo
    oTF.WriteToTextFile "BankingDetails" & vbTab & Replace(oPC.Configuration.Companies.FindCompanyByID(Me.COMPID).BankDetails, Chr(13) & Chr(10), Chr(10))
    oTF.WriteToTextFile_NoLineTerminator "CustomerWithAddress" & vbTab
                If Not BillTOAddress Is Nothing Then
                    oTF.WriteToTextFile Replace(Me.BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
                End If
    
    oTF.WriteToTextFile_NoLineTerminator "CustomerPhone" & vbTab
                If Me.BillTOAddress Is Nothing Then
                    oTF.WriteToTextFile IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "")
                End If
    oTF.WriteToTextFile_NoLineTerminator "CustomerFax" & vbTab
                If Me.BillTOAddress Is Nothing Then
                    oTF.WriteToTextFile IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
                End If
            
        If Customer.BillTOAddress.ID > 0 Then
            If Not oPC.Configuration.Stores.FindStoreByID(Customer.BillTOAddress.ID) Is Nothing Then
                s = Replace(BillTOAddress.AddressMailing, Chr(13) & Chr(10), Chr(10)) & Chr(10) & IIf(Me.BillTOAddress.Phone > "", "Phone: " & Me.BillTOAddress.Phone, "") & Chr(10) & IIf(Me.BillTOAddress.Fax > "", "Fax: " & Me.BillTOAddress.Fax, "")
            Else
                s = Replace(oPC.Configuration.DefaultStore.BillAddress, Chr(13) & Chr(10), Chr(10))
            End If
        End If
        If ForAttn > "" Then s = ForAttn & vbCrLf & s
    oTF.WriteToTextFile "BillTo" & vbTab & s
            
        If Customer.DelToAddress.ID > 0 Then
            If Not oPC.Configuration.Stores.FindStoreByID(Customer.DelToAddress.ID) Is Nothing Then
                s = Replace(Me.DelToAddress.AddressMailing, Chr(13) & Chr(10), Chr(10))
            Else
                s = ""
            End If
        End If
    oTF.WriteToTextFile "DelTo" & vbTab & s
    s = "SKU" & vbTab & "Title" & vbTab & "Qty" & vbTab & "Price" & vbTab & "PriceExVat" & vbTab & "DiscountRate" & vbTab & "Ref" & vbTab _
    & "Extension" & vbTab & "Note" & vbTab & "ExtensionExVat" & vbTab & "VAT" & vbTab & "VATRate" & vbTab & "DiscountValue" & vbTab _
    & "FCName" & vbTab & "FCPrice" & vbTab & "FCFactor"
    
    oTF.WriteToTextFile s
                           
                            
        lngNumberOfLines = 0
        For i = 1 To QuoteLines.Count
                lngNumberOfLines = lngNumberOfLines + 1
                    If oPC.HideLocalSKUonINV Then
                        s = QuoteLines.Item(i).CodeForExport
                    Else
                        s = QuoteLines.Item(i).CodeF
                    End If
                
                s = s & vbTab _
                & QuoteLines.Item(i).TitleAuthor & IIf(QuoteLines(i).ServiceItem, " (" & QuoteLines(i).Qty & " item" & IIf(QuoteLines(i).Qty = 1, ")", "s)"), "")
                
                s = s & vbTab & QuoteLines.Item(i).QtyF
                s = s & vbTab & QuoteLines.Item(i).PriceF(bForeign)
                s = s & vbTab & QuoteLines.Item(i).PriceExVatF(bForeign)
                s = s & vbTab & QuoteLines.Item(i).DiscountPercentF
                s = s & vbTab & QuoteLines.Item(i).Ref
                s = s & vbTab & QuoteLines.Item(i).ExtF(bForeign)
                s = s & vbTab & QuoteLines.Item(i).Note
                s = s & vbTab & QuoteLines.Item(i).Ext_ExFF(bForeign)
                s = s & vbTab & QuoteLines.Item(i).VATF(bForeign)
                s = s & vbTab & QuoteLines.Item(i).VATRateF
                s = s & vbTab & QuoteLines(i).DiscF(bForeign)
                        If QuoteLines(i).FCID > 0 Then
                            s2 = oPC.Configuration.Currencies.FindCurrencyByID(QuoteLines(i).FCID).Description
                        Else
                            s2 = "n/a"
                        End If

                s = s & vbTab & s2
                        If QuoteLines(i).FCID > 0 Then
                            s2 = QuoteLines(i).ForeignPriceF
                        Else
                            s2 = "n/a"
                        End If
                s = s & vbTab & s2
                        If QuoteLines(i).FCID > 0 Then
                            s2 = QuoteLines(i).FCFactorInvF
                        Else
                            s2 = "n/a"
                        End If
                
                s = s & vbTab & s2
                
                oTF.WriteToTextFile s
            Next
            
    oTF.WriteToTextFile "TotalNumberOfLines" & vbTab & CStr(lngNumberOfLines)
    oTF.WriteToTextFile "QtyItems" & vbTab & TotalQty
    oTF.WriteToTextFile "Memo" & vbTab & Memo
    oTF.WriteToTextFile "StaffMember" & vbTab & IIf(Me.StaffName > "", Me.StaffName, "")
    oTF.WriteToTextFile "TotalShipping" & vbTab & TotalServiceItemF(False)
    oTF.WriteToTextFile "TotalDiscount" & vbTab & TotalDiscountF(False)
    oTF.WriteToTextFile "TotalIncVAT" & vbTab & TotalPayableF(False)
    oTF.WriteToTextFile "TotalExVAT" & vbTab & TotalNonVATF(False)
    oTF.WriteToTextFile "TotalVAT" & vbTab & TotalVATF(False)
    oTF.WriteToTextFile "VATable" & vbTab & IIf(Me.VATable, "TRUE", "FALSE")
    oTF.WriteToTextFile "ShowVAT" & vbTab & IIf(Me.ShowVAT, "TRUE", "FALSE")
    oTF.CloseTextFile
    ExportToSpreadsheet = True
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_R.ExportToSpreadsheet(bForeign,pFilename)", Array(bForeign, pFilename)
End Function

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
    ErrorIn "a_QU.docWriteToFile(FilePath,Charset)", Array(FilePath, CharSet)
End Sub
Public Property Get Log() As String
    On Error GoTo errHandler
    Log = Trim(mudtProps.Log)
    Exit Property
errHandler:
    ErrorIn "a_QU.Log"
End Property

Public Property Let Log(val As String)
    On Error GoTo errHandler
    mudtProps.Log = Trim(val)
    Exit Property
errHandler:
    ErrorIn "a_QU.Log(val)", val
End Property

Public Property Get TotalQty() As Long
    On Error GoTo errHandler

        TotalQty = mudtProps.TotalQty

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TotalQty"
End Property
Public Property Get TotalQtyF() As String
    On Error GoTo errHandler

        TotalQtyF = CStr(TotalQty)

    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QU.TotalQtyF"
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
            For k = 1 To QuoteLines.Count
                If QuoteLines.Item(k).BottomOfDocument = True Then
                    addARow QuoteLines.Item(k).Title, QuoteLines.Item(k).PriceF(False), "", "+"
                    bBottomOfDocument = True
                End If
            Next k
            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "GRAND"
            addARow "(includes VAT at " & oPC.Configuration.VATRateF, TotalVATRaw(bForeign), TotalVATF(bForeign), "", " "
    Else
            bBottomOfDocument = False
            For k = 1 To QuoteLines.Count
                If QuoteLines.Item(k).BottomOfDocument = True And QuoteLines.Item(k).VATRate <> 0 Then
                    addARow QuoteLines.Item(k).Title, QuoteLines.Item(k).PriceF(bForeign), "", "+"
                    bBottomOfDocument = True
                End If
            Next k
            If bBottomOfDocument = True Then arDocTotals(jTotal - 1).TotalType = "SUB"
            If ShowVAT Then
                addARow "less VAT at " & oPC.Configuration.VATRateF, TotalVATRaw(bForeign), TotalVATF(bForeign), "SUB", "-"
            End If
            For k = 1 To QuoteLines.Count
                If QuoteLines(k).BottomOfDocument = True And QuoteLines(k).VATRate = 0 Then
                    addARow QuoteLines(k).Title, QuoteLines(k).Price, QuoteLines(k).PriceF(bForeign), "", "+"
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
    ErrorIn "a_QU.SetupDocumentTotals(bForeign)", bForeign
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
    ErrorIn "a_QU.addARow(pDescription,Amt,AmtF,pTotalType)", Array(pDescription, Amt, _
         AmtF, pTotalType)
End Sub


Public Sub PasteLine(PID As String, Qty As Long, Price As Long, DiscountRate As Double, VATRate As Double, Ref As String, _
                    ExtraChargePID As String, ExtraChargeValue As Long, _
                    FCPrice As Long, FCFactor As Double, FCID As Long)
Dim oSM As New z_StockManager

    oSM.PasteQuotationLine PID, Qty, Price, DiscountRate, VATRate, mudtProps.QuoteID, Ref, ExtraChargePID, ExtraChargeValue, FCPrice, FCFactor, FCID
    
End Sub
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
    ErrorIn "a_QU.TotalServiceItem(bForeign)", bForeign
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
    ErrorIn "a_QU.TotalServiceItemF(bForeign)", bForeign
End Property

